using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.ApplicationServices;
using POB;
using POB.ObjetoTransferenciaPOB;

namespace POB.Updater
{
    public class RevitUpdaterAmbienteElemento : IUpdater
    {
        static AddInId AddInId { get; set; }
        UpdaterId UpdaterId { get; set; }
        FailureDefinitionId AvisoId = null;
        FailureDefinition defincaoAviso;

        public RevitUpdaterAmbienteElemento(AddInId addInId)
        {
            AddInId = addInId;
            UpdaterId = new UpdaterId(AddInId, new Guid("e7a61e98-b671-46ed-b0ae-54460a0f7b6e"));
            AvisoId = new FailureDefinitionId(new Guid("cb25a814-4a5d-4726-9574-59d42e28f862"));
            defincaoAviso = FailureDefinition.CreateFailureDefinition(AvisoId, FailureSeverity.Warning, "Erro ao renomear");

        }

        void IUpdater.Execute(UpdaterData data)
        {

            Document uiDoc = data.GetDocument();
            var nome = new List<string>();
            UIDocument uidoc = new UIDocument(uiDoc);
            if (CriarMenu.OrgProdutoInexistente == null) CriarMenu.OrgProdutoInexistente = CriarMenu.getOrgProdutoInexistente(uiDoc);
        
            if (CriarMenu.ListaDeCategoriasDeItensContaveis.Count == 0)
                 CriarMenu.GetCategoriasItensContaveisPorQtde(uiDoc);
            /*if (CriarMenu.ListaQueTemTipoDeSistema.Count == 0)
                CriarMenu.GetQueTemTipoDeSistema(uiDoc);*/
            if (CriarMenu.ListaCategoriaItensSistemaPorMetro.Count == 0)
                CriarMenu.GetCategoriaItensSistemaPorMetro(uiDoc);


            foreach (ElementId eleId in data.GetModifiedElementIds())
            {
                if (uiDoc.GetElement(eleId) is Level)
                    foreach (Element ele in Util.GetFilterElementByParameter(uiDoc, "tocAmbienteNivelId", (uiDoc.GetElement(eleId) as Level).Id.IntegerValue))
                        Util.ObterAmbiente(uiDoc, ele.Id);
                else if (uiDoc.GetElement(eleId) is Space)
                    foreach (Element ele in Util.GetFilterElementByParameter(uiDoc, "tocAmbienteSistemaId", (uiDoc.GetElement(eleId) as Space).Id.IntegerValue)) 
                        Util.ObterAmbiente(uiDoc, ele.Id);
                else
                {
                    var element = uiDoc.GetElement(eleId);
                    GetL7InsumoVinculado(uiDoc, element);
                    //DadosDeConexoeseComprimentos(element);

                }

            }
            foreach (ElementId eleId in data.GetAddedElementIds())
            {
                if ((!(uiDoc.GetElement(eleId) is Level)) & (!(uiDoc.GetElement(eleId) is Room)))
                {
                    var element = uiDoc.GetElement(eleId);
                    Util.ObterAmbiente(uiDoc, eleId);                
                    GetL7InsumoVinculado(uiDoc, element);
                    Util.DadosDeConexoeseComprimentos(element);
                }
            }
        }

        private static void GetL7InsumoVinculado(Document uiDoc, Element element)
        {
           var L7InsumoVinculado = NovoExtrair.GetDescricao(uiDoc,element);
            if (L7InsumoVinculado == "Produto Inexistente")
                uiDoc.ActiveView.SetElementOverrides(element.Id, CriarMenu.OrgProdutoInexistente);
            else uiDoc.ActiveView.SetElementOverrides(element.Id, new OverrideGraphicSettings());
            element.LookupParameter(Properties.Settings.Default.L7InsumoVinculado).Set(L7InsumoVinculado);
        }

       

        string IUpdater.GetAdditionalInformation()
        {
            return "N/A";
        }

        ChangePriority IUpdater.GetChangePriority()
        {
            return ChangePriority.FreeStandingComponents;
        }

        UpdaterId IUpdater.GetUpdaterId()
        {
            return UpdaterId;
        }

        public UpdaterId GetUpdaterId()
        {
            return UpdaterId;
        }

        string IUpdater.GetUpdaterName()
        {
            return "Ambiente elemento";
        }

    }
}
