using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Reflection;
using wf =  System.Windows.Forms;
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


    public class RevestiPilar : IExternalCommand
    {
        public List<Curve> listaDecurvas = new List<Curve>();
    public FormRevestirPilar frmRevestirPilar = new FormRevestirPilar();
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData revit,
                  ref string message, ElementSet elements)
        {


            UIApplication uiApp = revit.Application;
            Document uiDoc = uiApp.ActiveUIDocument.Document;
            Selection sel = uiApp.ActiveUIDocument.Selection;
            Funcoes.Util.uiDoc = uiDoc;
         //   Funcoes.ProgressoFuncao progresso = new Funcoes.ProgressoFuncao();
           // progresso.pgc.Maximum = sel.GetElementIds().Count;
          //  progresso.pgc.Minimum = 0;
          //  progresso.pgc.Value = 1;
          //  progresso.pgc.Step = 1;
          //  progresso.Show();
            XYZ normal = new XYZ(0, 0, 1);
            // Transform offset = Transform.CreateTranslation((-.3048*.11)*normal);
            Transform offset = Transform.CreateTranslation(0 * normal);

            FilteredElementCollector pavimentos = Funcoes.Util.ObterLevels(uiDoc);
            Transaction transaction1 = new Transaction(uiDoc, "CreateGenericModel1");
            List<Curve> lista1 = new List<Curve>();


            FilteredElementCollector filtro = new FilteredElementCollector(uiDoc).OfClass(typeof(Autodesk.Revit.DB.WallType));
            List<string> wtLista = new List<string>();
            foreach (WallType wt1 in filtro)
            {
               
                wtLista.Add(wt1.Name);
            }

            frmRevestirPilar.PreenchecmbWallType(wtLista);
            // if ( == wf.DialogResult.Cancel) return Result.Cancelled;    
            frmRevestirPilar.ShowDialog();

            double deslocamentoDaBase = frmRevestirPilar.DeslocamentoBase;
            double alturaDesconectada = frmRevestirPilar.AlturaDesconectada;
            string wt = frmRevestirPilar.WallType;

            WallType wallType = Funcoes.Util.FindElementByName(typeof(WallType), wt ) as WallType;

            transaction1.Start();
            foreach (ElementId eleId in uiApp.ActiveUIDocument.Selection.GetElementIds())
            {

                    Element ele = uiDoc.GetElement(eleId);
                    Level pavimentoSuperior = Funcoes.Util.GetNivelMaisProximoToFace(uiDoc.GetElement(eleId), pavimentos);
                    Level pavimentoInferior = Funcoes.Util.GetNivelMaisProximo(uiDoc.GetElement(eleId), pavimentos);
                    double elevacaoSuperior = Util.GetElevacaoRelacaoAoZeroToFace(ele) / 0.3048;
                    double elevacaoInferior = Util.GetElevacaoRelacaoAoZero(ele) / 0.3048;

                    List<Solid> listaDeSolidos = Funcoes.Util.GetSolids(uiDoc.GetElement(eleId));
 
                foreach (Solid solido in listaDeSolidos)
                {
                    if (solido.Faces.Size > 0)
                    {
                        foreach (Face lateralFace in Util.GetVerticalFace(solido))
                        {
                            foreach (CurveLoop curveLoop in lateralFace.GetEdgesAsCurveLoops())
                            {
                                CurveLoopIterator cli = curveLoop.GetCurveLoopIterator();
                                listaDecurvas.Clear();
                                while (cli.MoveNext())
                                {
                                    listaDecurvas.Add(cli.Current);
                                }
                                try
                                {
                                    Wall w = Autodesk.Revit.DB.Wall.Create(uiDoc, listaDecurvas, wallType.Id, pavimentoInferior.Id, false);
                                }
                                catch (Exception e)
                                {

                                }
                            }

                        }
                    }


                }
            }

            transaction1.Commit();
            frmRevestirPilar.Dispose();
            return Result.Succeeded;
        }

        public XYZ GetPosicaoDoPilar(Face faceSuperior, Document uiDoc)
        {
            List<XYZ> listaDePontos = new List<XYZ>();
            foreach (CurveLoop curveLoop in faceSuperior.GetEdgesAsCurveLoops())
            {

                CurveLoopIterator cli = curveLoop.GetCurveLoopIterator();
                while (cli.MoveNext())
                {
                    Line l = cli.Current as Line;
                    XYZ p1 = l.GetEndPoint(0);
                    if (listaDePontos.Find(x => (x.Y == p1.Y) && (x.X == p1.X)) == null)
                    {
                        listaDePontos.Add(p1);
                    }
                    p1 = l.GetEndPoint(1);
                    if (listaDePontos.Find(x => (x.Y == p1.Y) && (x.X == p1.X)) == null)
                    {
                        listaDePontos.Add(p1);
                    }
                }
            }

            XYZ ponto1 = listaDePontos[0];
            XYZ pontoDistante = new XYZ(0, 0, 0);
            listaDePontos.Remove(ponto1);

            double maiorDistancia = 0;

            foreach (XYZ item in listaDePontos)
            {
                if (maiorDistancia == 0)
                {
                    maiorDistancia = GetDistanciaEntreDoisPontos(ponto1, item);
                    pontoDistante = item;
                }
                else
                {
                    if (maiorDistancia < GetDistanciaEntreDoisPontos(ponto1, item))
                    {
                        pontoDistante = item;
                        maiorDistancia = GetDistanciaEntreDoisPontos(ponto1, item);
                    }
                }
            }
            XYZ pm1 = GetPontoMedio(ponto1, pontoDistante);

            return pm1;

        }

        private double GetDistanciaEntreDoisPontos(XYZ ponto1, XYZ item)
        {
            double dist = 0;
            double x2 = Math.Pow((item.X - ponto1.X), 2);
            double y2 = Math.Pow((item.Y - ponto1.Y), 2);
            dist = Math.Sqrt(x2 + y2);
            return dist;

        }

        public XYZ GetPontoMedio(XYZ x1, XYZ x2)
        {
            return new XYZ((x1.X + x2.X) / 2, (x1.Y + x2.Y) / 2, x1.Z);
        }

        public FamilySymbol GetTipoDeFamilia(Face secaoTransversal, List<PilarRetangular> listai)
        {
            FamilySymbol fs;

            Line linhaX = Funcoes.Util.GetLinhaX(secaoTransversal);
            Line linhaY = Funcoes.Util.GetLinhaY(secaoTransversal);
            string secao = GetNomeSecao(linhaX.Length, linhaY.Length);
            try
            {
                fs = listai.Find(x => x.secao == secao).tipoDePilar;

                if (fs == null)
                {
                    fs = CriarNovoPilarRetangular(secao, linhaX, linhaY, listai[0].tipoDePilar);
                }
            }
            catch
            {
                fs = CriarNovoPilarRetangular(secao, linhaX, linhaY, listai[0].tipoDePilar);

            }
            return fs;


        }

        public FilteredElementCollector filtroListaDePilares(Document uiDoc)
        {

            FilteredElementCollector filtro = new FilteredElementCollector(uiDoc).OfClass(typeof(Autodesk.Revit.DB.FamilySymbol)).OfCategory(BuiltInCategory.OST_StructuralColumns);

            return filtro;
        }

        public List<PilarRetangular> GeraLista(FilteredElementCollector filtro)
        {
            List<PilarRetangular> listai = new List<PilarRetangular>();
            foreach (FamilySymbol item in filtro)
            {
                try
                {
                    if (item.LookupParameter("tipoDePilar").AsString() == "Retangular")
                    {

                        PilarRetangular v = new PilarRetangular();
                        v.tipoDePilar = item;
                        v.largura = item.LookupParameter("Largura").AsDouble();
                        v.altura = item.LookupParameter("Altura").AsDouble();
                        v.secao = GetNomeSecao(v.largura, v.altura);
                        listai.Add(v);
                    }

                }
                catch
                {

                }
            }
            return listai;
        }

        public string GetNomeSecao(double distanciaX, double distanciaY)
        {


            string secao = string.Format("{0:000}", Math.Round(distanciaX * 0.3048 * 100, 0));
            string secao1 = string.Format("{0:000}", Math.Round(distanciaY * 0.3048 * 100, 0));
            return secao + "x" + secao1;

        }

        public FamilySymbol CriarNovoPilarRetangular(string secao, Line largura, Line altura, FamilySymbol fsOriginal)
        {
            FamilySymbol fs = fsOriginal.Duplicate(secao) as FamilySymbol;
            fs.LookupParameter("Largura").Set(largura.Length);
            fs.LookupParameter("Altura").Set(altura.Length);
            PilarRetangular v = new PilarRetangular();
            v.tipoDePilar = fs;
            v.altura = altura.Length;
            v.largura = largura.Length;
            v.secao = secao;

            return fs;
        }


    }
}


