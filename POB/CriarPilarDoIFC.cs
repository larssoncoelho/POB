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

    public class PilarRetangular
    {
        public FamilySymbol tipoDePilar;
        public double largura;
        public double altura;
        public string secao;
    }

    [TransactionAttribute(TransactionMode.Manual)]
    [RegenerationAttribute(RegenerationOption.Manual)]


    public class CriarPilarDoIFC : IExternalCommand
    {
        public List<PilarRetangular> listaDePilar = new List<PilarRetangular>();
        public List<Curve> listaDecurvas = new List<Curve>();
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData revit,
                  ref string message, ElementSet elements)
        {


            UIApplication uiApp = revit.Application;
            Document uiDoc = uiApp.ActiveUIDocument.Document;
            Selection sel = uiApp.ActiveUIDocument.Selection;
            Funcoes.Util.uiDoc = uiDoc;
            Funcoes.ProgressoFuncao progresso = new Funcoes.ProgressoFuncao(sel.GetElementIds().Count);
            
            progresso.Show();
            XYZ normal = new XYZ(0, 0, 1);
            // Transform offset = Transform.CreateTranslation((-.3048*.11)*normal);
            Transform offset = Transform.CreateTranslation(0 * normal);

            FilteredElementCollector pavimentos = Funcoes.Util.ObterLevels(uiDoc);
            listaDePilar = GeraLista(filtroListaDePilares(uiDoc));
            Transaction transaction1 = new Transaction(uiDoc, "CreateGenericModel1");


            transaction1.Start();
            foreach (ElementId eleId in uiApp.ActiveUIDocument.Selection.GetElementIds())
            {
                Element ele = uiDoc.GetElement(eleId);
                
                try
                {


                    Level pavimentoSuperior = Funcoes.Util.GetNivelMaisProximoToFace(uiDoc.GetElement(eleId), pavimentos);
                    Level pavimentoInferior = Funcoes.Util.GetNivelMaisProximo(uiDoc.GetElement(eleId), pavimentos);
                    double elevacaoSuperior = Util.GetElevacaoRelacaoAoZeroToFace(ele)/0.3048;
                    double elevacaoInferior = Util.GetElevacaoRelacaoAoZero(ele)/0.3048;

                    List<Solid> listaDeSolidos = Funcoes.Util.GetSolids(uiDoc.GetElement(eleId));

                    foreach (Solid solido in listaDeSolidos)
                    {

                        if (solido.Faces.Size > 0)
                        {
                            Face faceDoSolido = Funcoes.Util.GetTopFace(solido);
                            Face secaoTransversal = Funcoes.Util.GetTopFace(solido);

                            FamilySymbol fs = GetTipoDeFamilia(secaoTransversal, listaDePilar);

                            XYZ localDeInsercao = GetPosicaoDoPilar(faceDoSolido, uiDoc);
                            double deslocamentoDaBase = ele.LookupParameter("Deslocamento da base").AsDouble();
                            double deslocamentoSuperior = ele.LookupParameter("Deslocamento superior").AsDouble();


                            Level levelBase = uiDoc.GetElement( ele.LookupParameter("Nível base").AsElementId()) as Level;
                            Level levelSuperior = uiDoc.GetElement(ele.LookupParameter("Nível superior").AsElementId()) as Level;

                            if (levelSuperior == null)
                            {
                                levelSuperior = pavimentoSuperior;
                                levelBase = pavimentoInferior;
                                deslocamentoDaBase = pavimentoInferior.Elevation - elevacaoInferior;
                                deslocamentoSuperior = pavimentoSuperior.Elevation - elevacaoSuperior;
                            }

                            FamilyInstance fi = uiDoc.Create.NewFamilyInstance(localDeInsercao, fs, levelSuperior, Autodesk.Revit.DB.Structure.StructuralType.Column);
                            fi.LookupParameter("Nível superior").Set(levelSuperior.Id);
                            fi.LookupParameter("Nível base").Set(levelBase.Id);
                            if (deslocamentoDaBase != 0)
                                fi.LookupParameter("Deslocamento da base").Set(deslocamentoDaBase);
                            else fi.LookupParameter("Deslocamento da base").Set(0);
                            if (deslocamentoSuperior != 0)
                                fi.LookupParameter("Deslocamento superior").Set(deslocamentoSuperior);
                            else fi.LookupParameter("Deslocamento superior").Set(0);


                            try
                            {
                                fi.LookupParameter(POB.Properties.Settings.Default.opbNome).Set(uiDoc.GetElement(eleId).LookupParameter("NameOverride").AsString());
                            }
                            catch
                            {

                            }
                            uiDoc.Delete(eleId);

                        }
                    }
                }
                catch (Exception e)
                {
                    transaction1.RollBack();

                }
                progresso.Incrementar();

            }
            transaction1.Commit();



            progresso.Dispose();
            return Result.Succeeded;

        }


        public XYZ GetPosicaoDoPilar(Face faceSuperior, Document uiDoc)
        {   List<XYZ> listaDePontos = new List<XYZ>();
            foreach (CurveLoop curveLoop in faceSuperior.GetEdgesAsCurveLoops())
            {

                CurveLoopIterator cli = curveLoop.GetCurveLoopIterator();
                while (cli.MoveNext())
                {
                    Line l = cli.Current as Line;
                    XYZ p1 = l.GetEndPoint(0);
                    if (listaDePontos.Find(x=> (x.Y == p1.Y) && (x.X == p1.X))==null)
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
                if (maiorDistancia==0)
                {
                    maiorDistancia = Util.GetDistanciaEntreDoisPontos(ponto1, item);
                    pontoDistante = item;
                }
                else
                {
                    if (maiorDistancia< Util.GetDistanciaEntreDoisPontos(ponto1, item))
                    {
                        pontoDistante = item;
                        maiorDistancia = Util.GetDistanciaEntreDoisPontos(ponto1, item);
                    }
                }
            }
            XYZ pm1 = GetPontoMedio(ponto1, pontoDistante);

            return pm1;

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
                        v.tipoDePilar= item;
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
            v.tipoDePilar= fs;
            v.altura = altura.Length;
            v.largura = largura.Length;
            v.secao = secao;
            listaDePilar.Add(v);
            return fs;
        }


    }
}


