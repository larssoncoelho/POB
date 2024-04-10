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
using System.Runtime.InteropServices.ComTypes;
using wf = System.Windows.Forms;

//using LinqToExcel;

namespace POB
{



    [TransactionAttribute(TransactionMode.Manual)]
    [RegenerationAttribute(RegenerationOption.Manual)]

    public class ObterMaiorComprimentoDoSolido : IExternalCommand
    {

        FamilySymbol fs1;
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData revit,
            ref string message, ElementSet elements)
        {
            UIApplication uiApp = revit.Application;
            Document uiDoc = uiApp.ActiveUIDocument.Document;
            Selection sel = uiApp.ActiveUIDocument.Selection;
            //TransactionGroup t = new TransactionGroup(uiDoc);

            ViewSchedule _viewSchedule = uiDoc.ActiveView as ViewSchedule;
            
            TableSectionData sectionData = _viewSchedule.GetTableData().GetSectionData(Autodesk.Revit.DB.SectionType.Body);

            var numberOfRows = sectionData.NumberOfRows;
            var numberOfColumns = sectionData.NumberOfColumns;
            var firstRowNumber = sectionData.FirstRowNumber;

            return Result.Succeeded;



            List<ObjetoDeTranferencia.DadosExcel> dadosExcel = new List<ObjetoDeTranferencia.DadosExcel>();
            wf.OpenFileDialog openFileDialog = new wf.OpenFileDialog();
            openFileDialog.Title = "Escolha um arquivo";

            if (openFileDialog.ShowDialog() == wf.DialogResult.OK)
            {
                uiDoc.Save();
                // O usuário selecionou um arquivo. Faça algo com o caminho do arquivo, por exemplo:
                string caminhoArquivo = openFileDialog.FileName;


                //var excel = new ExcelQueryFactory(caminhoArquivo) { ReadOnly = true };

                // Obter todas as linhas da planilha
              /*  dadosExcel = (from linha in excel.Worksheet("depara")
                              select new ObjetoDeTranferencia.DadosExcel
                              {
                                  NomeModelo = linha["Vínculo RVT: Nome do arquivo"].ToString(),
                                  Id = linha["Código parede"].ToString(),
                                  CodComposicao = linha["CodComposicao"].ToString()
                              }).ToList();*/
                //escolher o arquivo
                //ler arquivo
                //
                dadosExcel = dadosExcel.Where(x => x.NomeModelo.Contains(System.IO.Path.GetFileNameWithoutExtension(uiDoc.PathName))).ToList();

                Autodesk.Revit.DB.ElementId idDoParametro = GetIdParametroProjeto(uiDoc, "Código parede");
                Transaction t = new Transaction(uiDoc);
                int i = 0;
                t.Start("Transacao");
                foreach (ObjetoDeTranferencia.DadosExcel d in dadosExcel)
                {
                    var elementos = FiltraElementoId(uiDoc, d.Id, idDoParametro);
                    foreach (var element in elementos)
                    {
                        element.LookupParameter("CodComposicao").Set(d.CodComposicao);
                    }
                    i++;
                    if(i == 1000)
                    {
                        i = 0;
                        t.Commit();
                        t.Start("Transacao");
                    }
                }
                if(t.GetStatus()==TransactionStatus.Started)  t.Commit();
                return Result.Succeeded;
            }
            return Result.Cancelled;

        }

        public static List<Autodesk.Revit.DB.Element> FiltraElementoId(Document uidoc, string valorProcurado, ElementId idDoParametro)
        {

            IList<ElementFilter> b = new List<ElementFilter>();

            ParameterValueProvider provider = new ParameterValueProvider(idDoParametro);
            FilterStringRuleEvaluator evaluator = new FilterStringEquals();

            FilterRule rule = new FilterStringRule(provider, evaluator, valorProcurado);

            ElementParameterFilter filter = new ElementParameterFilter(rule);
            b.Add(filter);

            var logicalOrFilter = new LogicalAndFilter(b);
            FilteredElementCollector fi = new FilteredElementCollector(uidoc);
            var consulta = fi.WherePasses(logicalOrFilter).ToElements().ToList();
            return consulta;
        }
        private static ElementId GetIdParametroProjeto(Document uidoc, string nome1)
        {
            List<ElementId> vetor = new List<ElementId>();
            BindingMap bindingMap = uidoc.ParameterBindings;
            DefinitionBindingMapIterator it = bindingMap.ForwardIterator();
            it.Reset();

            while (it.MoveNext())
            {
                if (it.Key.Name == nome1)
                {
                    return (it.Key as InternalDefinition).Id;

                }
            }
            return null;

        }
    }
}
               /* var primeiro = true;
                var modeloAtual = "";
                var modeloAnterior = "";

                //ggg.Funcoes1.Exportar(uiDoc, this.ActiveUIDocument.Selection);

                Autodesk.Revit.DB.ElementId idDoParametro = null;
                Autodesk.Revit.DB.Transaction t1 = null;
                Autodesk.Revit.DB.Document docAtivo = null;
                foreach (ObjetoDeTranferencia.DadosExcel d in dadosExcel)
                {
                    if (!primeiro)
                    {
                        modeloAtual = d.NomeModelo;
                        if (modeloAtual != modeloAnterior)
                        {
                            //mudou
                            (t1 as Autodesk.Revit.DB.Transaction).Commit();
                            docAtivo = Util.EscolherDocumentoAtivo(uiApp, d.NomeModelo);
                            idDoParametro = GetIdParametroProjeto(docAtivo, "Código parede");
                            (t1 as Autodesk.Revit.DB.Transaction).Start("teste");
                            modeloAnterior = modeloAtual;
                        }

                    }
                    else
                    {
                        docAtivo = Util.EscolherDocumentoAtivo(uiApp, d.NomeModelo);
                        modeloAtual = d.NomeModelo;
                        modeloAnterior = d.NomeModelo;
                        t1 = new Autodesk.Revit.DB.Transaction(docAtivo, "Teste");
                        t1.Start();
                        idDoParametro = GetIdParametroProjeto(docAtivo, "Código parede");
                    }
                    primeiro = false;
                    var elementos = FiltraElementoId(docAtivo, d.Id, idDoParametro);
                    foreach (var element in elementos)
                    {
                        element.LookupParameter("CodComposicao").Set(d.CodComposicao);
                        primeiro = false;
                    }
                    primeiro = false;


                }
                (t1 as Autodesk.Revit.DB.Transaction).Commit();



               
            }
            return Result.Succeeded;
        }
        
    }
}
            
            /*t.Start("Teste");
            foreach (ElementId item in sel.GetElementIds())
            {
                Element ele = uiDoc.GetElement(item);
                double comprimento = 0;
                try
                {
#if D23
                    var p1 = SpecTypeId.Area;
#else
                 var p1 =ParameterType.Area;
#endif
                    var par = ele.LookupParameter("Comprimento modelo genérico").AsDouble();// Util.GetParameter(ele, "Comprimento modelo genérico", p1, true, false);
                    foreach (Solid solid in Util.GetSolids(ele))
                    {
                        if (solid != null)
                        {
                            var facesVerticais = Util.GetFaceLista(solid);
                            var maiorFace = facesVerticais.OrderByDescending(x => x.Area).First();
                            List<Line> listaDeLinha = new List<Line>();
                            foreach (CurveLoop curveLoop in maiorFace.GetEdgesAsCurveLoops())
                            {
                                CurveLoopIterator cli = curveLoop.GetCurveLoopIterator();
                                while (cli.MoveNext())
                                {
                                    listaDeLinha.Add(cli.Current as Line);
                                }
                            }
                            comprimento = comprimento +listaDeLinha.OrderByDescending(x=>x.Length).First().Length;
                        }
                    }
                    Transaction t1 = new Transaction(ele.Document);
                    t1.Start("ttt");
                    par.Set(comprimento);
                    t1.Commit();
                    t1.Dispose();
                }
                catch
                {

                }
            }
            t.Commit();


            return Result.Succeeded;
        }
    }




}*/
