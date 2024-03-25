using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Reflection;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using System.Data;
using RevitCreation = Autodesk.Revit.Creation;
using System.Runtime.InteropServices;
using Autodesk.Revit;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;

using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.ApplicationServices;
using POB.Updater;

namespace POB
{
    [TransactionAttribute(TransactionMode.Manual)]
    [RegenerationAttribute(RegenerationOption.Manual)]

    public class ExtrairDados : IExternalCommand
    {
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData revit,
           ref string message, ElementSet elements)
        {
            UIApplication uiApp = revit.Application;
            Document uiDoc = uiApp.ActiveUIDocument.Document;
            Selection sel = uiApp.ActiveUIDocument.Selection;
            XYZ P = new XYZ(0, 0, 0);
            POB.Util.uiDoc = uiDoc;
            foreach (ElementId eleId in sel.GetElementIds())
            {
               //Pipe ele = uiDoc.GetElement(eleId) as Pipe;
               // Parameter p = ele.get_Parameter(BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM);

            }
            /* Apresentacao.CriarTabelaDinamica criarTabelaDinamica = new Apresentacao.CriarTabelaDinamica();
             criarTabelaDinamica.ShowDialog();
             if (criarTabelaDinamica.Continuar)
             {

                 CriarTabelasUtil.GerarTabelaAF(uiDoc, "AMBIENTE", criarTabelaDinamica.Ambiente,
                     "PAVIMENTO", criarTabelaDinamica.Pavimento);
                 CriarTabelasUtil.GerarTabelaEsgoto(uiDoc, "AMBIENTE", criarTabelaDinamica.Ambiente,
                        "PAVIMENTO", criarTabelaDinamica.Pavimento);

             } else criarTabelaDinamica.Close();
             criarTabelaDinamica.Dispose();
             return Result.Succeeded; */
            // Funcoes.ProgressoFuncao progressoFuncao = new Funcoes.ProgressoFuncao(sel.GetElementIds().Count);
            //   progressoFuncao.Show();
            /*  if (sel.GetElementIds().Count==0)
              {
                  FilteredElementCollector fi = new FilteredElementCollector(uiDoc);
                  var lista1 = fi.WherePasses(CriarMenu.DefinirFiltroRegisterUpdateDescricaoEelemento()).ToElementIds();
                  sel.SetElementIds(lista1);
              }*/
            if (CriarMenu.OrgProdutoInexistente == null) CriarMenu.OrgProdutoInexistente = CriarMenu.getOrgProdutoInexistente(uiDoc);
            if (CriarMenu.ListaDeCategoriasDeItensContaveis.Count == 0)
                CriarMenu.GetCategoriasItensContaveisPorQtde(uiDoc);
            /*if (CriarMenu.ListaQueTemTipoDeSistema.Count == 0)
                CriarMenu.GetQueTemTipoDeSistema(uiDoc);*/
            if (CriarMenu.ListaCategoriaItensSistemaPorMetro.Count == 0)
                CriarMenu.GetCategoriaItensSistemaPorMetro(uiDoc);
            //TransactionGroup tg = new TransactionGroup(uiDoc);
            
            Transaction t = new Transaction(uiDoc);
            t.Start("Inicio");
            foreach (ElementId eleId in sel.GetElementIds())
            {
              
                try
                {
                    var ele = uiDoc.GetElement(eleId);
                    var parComposicao = ele.LookupParameter("parComposicao");
                    var descricao = ele.LookupParameter(Properties.Settings.Default.L7InsumoVinculado);
                    if (parComposicao != null)
                    {
                        if (parComposicao.HasValue)
                        {
                            if (!string.IsNullOrEmpty(parComposicao.AsString()))
                            {
                                var campos = parComposicao.AsString();
                                var descricaoFinal = "";
                                foreach (var item in campos.Split(';'))
                                {
                                    var valores = item.Split('|');
                                    var campo = valores[0];
                                    var digito = valores[1];
                                    descricaoFinal = descricaoFinal + ele.LookupParameter(campo).AsString() + digito;

                                }
                                descricao.Set(descricaoFinal);
                            }
                            else
                            {

                                if (descricao != null)
                                    descricao.Set(NovoExtrair.GetDescricao(uiDoc, ele));
                            }
                        }
                        else
                        {
                            if (descricao != null)
                                descricao.Set(NovoExtrair.GetDescricao(uiDoc, ele));
                        }
                    }
                    else
                    {
                        if (descricao != null)
                            descricao.Set(NovoExtrair.GetDescricao(uiDoc, ele));
                    }
                  
                    POB.Util.DadosDeConexoeseComprimentos(ele);                    
                    Util.ObterAmbiente(uiDoc, eleId);

                    /*if (uiDoc.GetElement(eleId) is Level)
                        foreach (Element element in Util.GetFilterElementByParameter(uiDoc, "tocAmbienteNivelId", (uiDoc.GetElement(eleId) as Level).Id.IntegerValue))
                            Util.ObterAmbiente(uiDoc, ele.Id);
                    else if (uiDoc.GetElement(eleId) is Space)
                        foreach (Element element in Util.GetFilterElementByParameter(uiDoc, "tocAmbienteSistemaId", (uiDoc.GetElement(eleId) as Space).Id.IntegerValue))
                            Util.ObterAmbiente(uiDoc, ele.Id);*/


                   
                }
                catch
                {
                   
                }
            }
            t.Commit();
         //   tg.Commit();
            return Result.Succeeded;
            /*Element ele = uiDoc.GetElement(eleId);
            int category = ele.Category.Id.IntegerValue;
            if ((ele is Autodesk.Revit.DB.Plumbing.Pipe) | (ele is Autodesk.Revit.DB.Plumbing.FlexPipe))
            {

                ele.LookupParameter("DescricaoMaterial").Set(new Extrair4().DadosTubulacao(ele));
            }
            else
                switch (category)
                {
                    case (int)BuiltInCategory.OST_PipeFitting:
                        ele.LookupParameter("DescricaoMaterial").Set(new Extrair4().DadosConcexaoTubo(ele));
                        break;
                    case (int)BuiltInCategory.OST_PipeAccessory:
                    case (int)BuiltInCategory.OST_MechanicalEquipment:
                    case (int)BuiltInCategory.OST_PlumbingFixtures:
                    case (int)BuiltInCategory.OST_PlaceHolderPipes:


                        ele.LookupParameter("DescricaoMaterial").Set(new Extrair4().DadosPecaTubo(ele));

                         break;


                    default:
                        break;
                }
       //     progressoFuncao.Incrementar();
        }
      //  t.Commit();
    //    progressoFuncao.Close();
     //   progressoFuncao.Dispose();*/

        }
    }
}
