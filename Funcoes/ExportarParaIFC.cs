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
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.ApplicationServices;
using wf = System.Windows.Forms;
using Funcoes;
using System.Windows.Forms;
using ObjetoDeTranferencia;

namespace criarMenu
{

    [TransactionAttribute(TransactionMode.Manual)]
    [RegenerationAttribute(RegenerationOption.Manual)]

    public class ExportarParaIFC : IExternalCommand
    {

        public static ElementType ee;
        public static Autodesk.Revit.DB.Plumbing.Pipe tuboLevantamento; //= new Autodesk.Revit.DB.Plumbing.Pipe();
        public static int MEDICAO_BLOCO_ID;
        public static int SERVICO_AMO_ID;
        public static double L;
        public static string resultado1;
        public static Autodesk.Revit.DB.Floor pisoRevit;

        public static ElementId preenchimentoId;
        public static double percentAvanco;
        public static DateTime diaRealizado;
        public static List<Element> listaEle = new List<Element>();

        public static FilteredElementCollector selecao;

        public DadosIntegracao dadosIntegracao { get; private set; }

        public static void BuscarAninhados(Element ele)
        {

            if (ele is FamilyInstance)
            {
                FamilyInstance aFamilyInst = ele as FamilyInstance;
                if (aFamilyInst.SuperComponent == null)
                {
                    var subElements = aFamilyInst.GetSubComponentIds();
                    if (subElements.Count != 0)
                    {
                        foreach (var aSubElemId in subElements)
                        {
                            var aSubElem = ele.Document.GetElement(aSubElemId);
                            if (aSubElem is FamilyInstance)
                            {
                                listaEle.Add(aSubElem);
                                BuscarAninhados(aSubElem);
                            }
                        }
                    }
                }
                else
                {
                    var subElements1 = aFamilyInst.GetSubComponentIds();
                    if (subElements1.Count != 0)
                    {
                        foreach (var aSubElemId in subElements1)
                        {
                            var aSubElem1 = ele.Document.GetElement(aSubElemId);
                            if (aSubElem1 is FamilyInstance)
                            {
                                listaEle.Add(aSubElem1);
                                BuscarAninhados(aSubElem1);
                            }
                        }
                    }
                }
            }
        }

        public Autodesk.Revit.UI.Result Execute(ExternalCommandData revit,
            ref string message, ElementSet elements)
        {
            UIApplication uiApp = revit.Application;
            Document uiDoc = uiApp.ActiveUIDocument.Document;
            Selection sel = uiApp.ActiveUIDocument.Selection;
            Util.uiDoc = uiDoc;
            listaEle.Clear();
            foreach (ElementId eleId in sel.GetElementIds())
            {
                Element ele = uiDoc.GetElement(eleId);
                listaEle.Add(ele);
                BuscarAninhados(ele);
            }
            List<string> comp = new List<string>();
            string composicao = "";


            foreach (Element item in listaEle)
            {
                string marca = item.LookupParameter("UAU_COMP").AsString();
                if (!string.IsNullOrEmpty(marca))
                    foreach (var item1 in marca.Split(';'))
                    {

                        string texto = item1;//.Split('|')[1];
                        comp.Add(texto);


                    }
            }
           /* CsAddPanel.MODELO_GUID_ID = "";
            DadosIntegracao dadosIntegracao = new DadosIntegracao();

            dadosIntegracao.Email = Funcoes.Properties.Settings.Default.usuario;
            dadosIntegracao.Senha = Funcoes.Properties.Settings.Default.senha;
            dadosIntegracao.ConectarNaWeb = false;
            dadosIntegracao.ConectarNaWeb = Funcoes.Properties.Settings.Default.ConectarNaWeb;
            dadosIntegracao.Diretorio = CsAddPanel.Diretorio;

            CsAddPanel.MODELO_GUID_ID = Util.RegistraDadosModelo(uiDoc, dadosIntegracao);
            foreach (var item in comp.Distinct().ToList())
            {
                var lista = Util.GetFilterElementByParameterId(uiDoc, "UAU_COMP", item);
                Util.SalvarIFC(uiApp.ActiveUIDocument, lista.ToList(), CsAddPanel.MODELO_GUID_ID, item);
            }*/

            return Result.Succeeded;


        }
    }
}