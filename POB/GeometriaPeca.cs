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
using Funcoes;
namespace POB
{
    [TransactionAttribute(TransactionMode.Manual)]
    [RegenerationAttribute(RegenerationOption.Manual)]


    public class GeometriaPeca : IExternalCommand
    {

        int qtdeReta = 0;
        int qtdeCurva = 0;
        int qtdeAresta = 0;
        List<Curve> curvasDaFace = new List<Curve>();
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData revit,
                 ref string message, ElementSet elements)
        {

            UIApplication uiApp = revit.Application;
            Document uiDoc = uiApp.ActiveUIDocument.Document;
            Selection sel = uiApp.ActiveUIDocument.Selection;
            //  Funcoes.Util.uiDoc = uiDoc;
            Funcoes.ProgressoFuncao progresso = new Funcoes.ProgressoFuncao(sel.GetElementIds().Count);
           
            progresso.Show();
            XYZ normal = new XYZ(0, 0, 1);
            // Transform offset = Transform.CreateTranslation((-.3048*.11)*normal);
            Transform offset = Transform.CreateTranslation(0 * normal);
            Transaction transaction1 = new Transaction(uiDoc, "CreateGenericModel1");
            transaction1.Start();
            foreach (ElementId eleId in uiApp.ActiveUIDocument.Selection.GetElementIds())
            {

                string DescricaoGeometriaPeca = "";
                string GeometriaPeca = "";
                Element ele = uiDoc.GetElement(eleId);
                qtdeCurva = 0;
                qtdeReta = 0;
                qtdeAresta = 0;

                if (ele is Autodesk.Revit.DB.Part)
                {
                    List<Solid> listaDeSolidos = Funcoes.Util.GetSolids(ele);

                    foreach (Solid solido in listaDeSolidos)
                    {
                       
                        if (solido.Faces.Size > 0)
                        {

                            Face faceDoSolido = Funcoes.Util.GetTopFace(solido);

                            foreach (CurveLoop curveLoop in faceDoSolido.GetEdgesAsCurveLoops())
                            {
                                CurveLoopIterator cli = curveLoop.GetCurveLoopIterator();
                                curvasDaFace = getCurvasDaFace(cli);

                                cli = curveLoop.GetCurveLoopIterator();
                                List<double> dimensoesPeca = getLadosIguais(cli);

                                if (dimensoesPeca.Count == 1)
                                {
                                    ele.LookupParameter("Largura").Set(dimensoesPeca[0]);
                                    ele.LookupParameter("Base").Set(dimensoesPeca[0]);

                                }
                                if (dimensoesPeca.Count == 2)
                                {
                                    ele.LookupParameter("Largura").Set(dimensoesPeca[0]);
                                    ele.LookupParameter("Base").Set(dimensoesPeca[1]);

                                }

                                cli = curveLoop.GetCurveLoopIterator();
                                ele.LookupParameter("Descrição Geometria da peça").Set(getFormatoPeca(curvasDaFace.Count, dimensoesPeca));
                                cli = curveLoop.GetCurveLoopIterator();
                                while (cli.MoveNext())
                                {
                                    if (cli.Current is Line)
                                    {

                                        if ((qtdeReta == 0)&(qtdeCurva==0))
                                        {

                                            GeometriaPeca = "R-" + Math.Round((cli.Current as Line).Length * .3048 * 100, 0).ToString();
                                        }
                                        else
                                        {
                                            GeometriaPeca = GeometriaPeca+ "R-" + Math.Round((cli.Current as Line).Length * .3048 * 100, 0).ToString() ;
                                        }
                                        qtdeReta = qtdeReta + 1;
                                    }
                                    if (cli.Current is Arc)
                                    {

                                        if ((qtdeCurva == 0)&(qtdeReta==0))
                                        {

                                            GeometriaPeca = "C-" + Math.Round((cli.Current as Arc).Length * .3048 * 100, 0).ToString();
                                        }
                                        else
                                        {
                                            GeometriaPeca = GeometriaPeca +  "C-" + Math.Round((cli.Current as Arc).Length * .3048 * 100, 0).ToString();
                                        }
                                        qtdeCurva = qtdeCurva + 1;
                                    }


                                }
                                ele.LookupParameter("Geometria da peça").Set(GeometriaPeca);
                            }

                        }




                     


                    }

                }
            }
            progresso.Close();
            progresso.Dispose();
            transaction1.Commit();
            return Result.Succeeded;
        }

        private List<Curve> getCurvasDaFace(CurveLoopIterator cli)
        {
            List<Curve> icurvasDaface = new List<Curve>();
            while (cli.MoveNext())
            {
                icurvasDaface.Add(cli.Current);
            }
            return icurvasDaface;
        }

        public int getQtdeArestas(CurveLoopIterator cli)
        {
            int qtdeLados = 0;
            while (cli.MoveNext())
            {
                qtdeLados = qtdeLados + 1;
            }
            return qtdeLados;
        }
        public List<double> getLadosIguais(CurveLoopIterator cli)
        {

            List<double> lista = new List<double>();
            while (cli.MoveNext())
            {

                if (!lista.Contains(Math.Round( cli.Current.Length,4) ))
                {
                    lista.Add(Math.Round(cli.Current.Length, 4));
                }
                
            }

            return lista;
        }
        public string getFormatoPeca(int iqtdeAresta,  List<Double> dimensoesPeca)
        {

            if ((iqtdeAresta == 4) & (dimensoesPeca.Count == 1))
                return "Peça quadrada";
            if ((iqtdeAresta == 4) & (dimensoesPeca.Count == 2))
                return "Peça retangular";

            return "Peça irregular";
        }
    }
}



