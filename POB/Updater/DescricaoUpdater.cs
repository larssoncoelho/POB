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
    public class RevitUpdaterDescricaoElemento : IUpdater
    {
        static AddInId AddInId { get; set; }
        UpdaterId UpdaterId { get; set; }
        FailureDefinitionId AvisoId = null;
        FailureDefinition defincaoAviso;

        public RevitUpdaterDescricaoElemento(AddInId addInId)
        {
            AddInId = addInId;
            UpdaterId = new UpdaterId(AddInId, new Guid("89A1F227-D9C3-4E40-BB90-B9384A2A22AF"));
            AvisoId = new FailureDefinitionId(new Guid("2AA8EC4D-1E72-47FC-A74A-C7C91BD3F4B0"));
            defincaoAviso = FailureDefinition.CreateFailureDefinition(AvisoId, FailureSeverity.Warning,"Erro ao renomear");

        }

        void IUpdater.Execute(UpdaterData data)
        {

            Document uiDoc = data.GetDocument();
            var nome = new List<string>();
            UIDocument uidoc = new UIDocument(uiDoc);
            var listaV = uidoc.GetOpenUIViews().ToList();
            var listaDrafthingView = new FilteredElementCollector(uiDoc).OfClass(typeof(ViewDrafting)).Cast<ViewDrafting>().ToList();
            var f = from ViewDrafting vv in listaDrafthingView
                    join UIView ui in listaV on vv.Id equals ui.ViewId
                    select ui;


            foreach (ElementId eleId in data.GetAddedElementIds())
            {
                Nomear(uiDoc, eleId);
                string am = CriarTabelasUtil.GetAmbiente(uiDoc.GetElement(eleId));
                string pv = CriarTabelasUtil.GetPavimento(uiDoc.GetElement(eleId));
                if(!nome.Contains(CriarTabelasUtil.GetNameDrafthingView(am, pv, TipoRelatorio.AF)))
                     nome.Add(CriarTabelasUtil.GetNameDrafthingView(am, pv, TipoRelatorio.AF));

                

                //string nomeVista = CriarTabelasUtil.  
            }
            foreach (ElementId eleId in data.GetModifiedElementIds())
            {
                Nomear(uiDoc, eleId);
            }
            foreach (UIView item in f.ToList())
            {
                item.Close();
            }
        }
        private static void Nomear(Document uiDoc, ElementId eleId)
        {
            Element ele = uiDoc.GetElement(eleId);
            int category = ele.Category.Id.IntegerValue;
            
            if ((ele is Autodesk.Revit.DB.Plumbing.Pipe) | (ele is Autodesk.Revit.DB.Plumbing.FlexPipe))
                ele.LookupParameter("DescricaoMaterial").Set(new Extrair4().DadosTubulacao(ele));

            else

                switch (category)
                {
                    case (int)BuiltInCategory.OST_PipeFitting:
                        if (ele is FamilyInstance)
                        {
                            var super = (ele as FamilyInstance).SuperComponent;
                            if (super != null)
                                Nomear(uiDoc, super.Id);
                        }
                        ele.LookupParameter("DescricaoMaterial").Set(new Extrair4().DadosConcexaoTubo(ele));

                        break;
                    case (int)BuiltInCategory.OST_PipeAccessory:
                    case (int)BuiltInCategory.OST_MechanicalEquipment:
                    case (int)BuiltInCategory.OST_PlumbingFixtures:
                    case (int)BuiltInCategory.OST_PlaceHolderPipes:

                        ele.LookupParameter("DescricaoMaterial").Set(new Extrair4().DadosPecaTubo(ele));
                        if (ele is FamilyInstance)
                        {
                            var super = (ele as FamilyInstance).SuperComponent;
                            if (super != null)
                                Nomear(uiDoc, super.Id);
                        }



                        break;


                    default:
                        break;
                }
        }

        private string GetNome()
        {
            return "Nome";
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
            return "Análise parede";
        }

    }
}
