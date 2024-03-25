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

namespace POB
{

    public class VigaRetangular
    {
        public FamilySymbol tipoDeViga;
        public double largura;
        public double altura;
        public string secao;
    }

    [TransactionAttribute(TransactionMode.Manual)]
    [RegenerationAttribute(RegenerationOption.Manual)]


    public class CriarVigaDoIFC: IExternalCommand
    {
        public List<VigaRetangular> listaDeVigas = new List<VigaRetangular>();
        public List<Curve> listaDecurvas = new List<Curve>();      
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

            FilteredElementCollector pavimentos = Funcoes.Util.ObterLevels(uiDoc);
            listaDeVigas = GeraLista(filtroListaDeVigasRetangulares(uiDoc));
            Transaction transaction1 = new Transaction(uiDoc, "CreateGenericModel1");
            transaction1.Start();
            foreach (ElementId eleId in uiApp.ActiveUIDocument.Selection.GetElementIds())
            {

                
                try
                {

                    
                    Level pavimento = Funcoes.Util.GetNivelMaisProximo(uiDoc.GetElement(eleId), pavimentos);
                    List<Solid> listaDeSolidos = Funcoes.Util.GetSolids(uiDoc.GetElement(eleId));

                    foreach (Solid solido in listaDeSolidos)
                    {

                        if (solido.Faces.Size > 0)
                        {

                            Face faceDoSolido = Funcoes.Util.GetTopFace(solido);
                            Face secaoTransversal = Funcoes.Util.GetSecaoTransversalViga(solido.Faces);
                            FamilySymbol fs = GetTipoDeFamilia(secaoTransversal, listaDeVigas);
                            Curve curva = GetLinhaDaViga(faceDoSolido, uiDoc);

                            FamilyInstance fi = uiDoc.Create.NewFamilyInstance(curva, fs, pavimento, Autodesk.Revit.DB.Structure.StructuralType.Beam);
                            try 
                            { 
                                fi.LookupParameter(POB.Properties.Settings.Default.opbNome).Set((uiDoc.GetElement(eleId) as FamilyInstance).Name);
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
                 //   transaction1.RollBack();

                }
                progresso.Incrementar();
            }
            transaction1.Commit();

            progresso.Dispose();  
            return Result.Succeeded;

        }


        public Curve GetLinhaDaViga(Face faceSuperior, Document uiDoc)
        {
            List<Line> listaDeLinha = new List<Line>();

            foreach (CurveLoop curveLoop in faceSuperior.GetEdgesAsCurveLoops())
            {

                CurveLoopIterator cli = curveLoop.GetCurveLoopIterator();
                while (cli.MoveNext())
                {
                    listaDeLinha.Add(cli.Current as Line);
                    // ca.Append(cli.Current.CreateTransformed(offset));
                }
                
             

            }
            var listaOrdenada = listaDeLinha.OrderBy(e => e.Length).ToList();
            Line l1 = listaOrdenada[0];
            Line l2 = listaOrdenada[1];

            XYZ p1a = l1.GetEndPoint(0);
            XYZ p1b= l1.GetEndPoint(1);

            XYZ p2a = l2.GetEndPoint(0);
            XYZ p2b = l2.GetEndPoint(1);

            XYZ pm1 = GetPontoMedio(p1a, p1b);
            XYZ pm2 = GetPontoMedio(p2a, p2b);


            //  XYZ pm1 = new XYZ()

            Curve c = Line.CreateBound(pm1, pm2);

            return c;

        }

        public XYZ GetPontoMedio(XYZ x1, XYZ x2)
        {
            return new XYZ((x1.X + x2.X) / 2, (x1.Y + x2.Y) / 2, x1.Z);
        }

        public FamilySymbol GetTipoDeFamilia(Face secaoTransversal, List<VigaRetangular> listai)
        {
            FamilySymbol fs;
            Line menorDimensao = Funcoes.Util.GetMenorDimensao(secaoTransversal);
            Line maiorDimensao = Funcoes.Util.GetMaiorDimensao(secaoTransversal);
            string secao = GetNomeSecao(menorDimensao.Length, maiorDimensao.Length);
           try
            {
                fs = listai.Find(x => x.secao == secao ).tipoDeViga;
           
                if (fs == null)
                {
                    fs = CriarNovaVigaRetangular(secao, menorDimensao, maiorDimensao, listai[0].tipoDeViga);
                }
            }
            catch
            {
                fs = CriarNovaVigaRetangular(secao, menorDimensao, maiorDimensao, listai[0].tipoDeViga);

            }
            return fs;


        }

        public FilteredElementCollector filtroListaDeVigasRetangulares(Document uiDoc)
        {

            FilteredElementCollector filtro = new FilteredElementCollector(uiDoc).OfClass(typeof(Autodesk.Revit.DB.FamilySymbol)).OfCategory(BuiltInCategory.OST_StructuralFraming);

            return filtro;
        }

        public List<VigaRetangular> GeraLista (FilteredElementCollector filtro)
        {
            List<VigaRetangular> listai = new List<VigaRetangular>();
            foreach (FamilySymbol item in filtro)
            {
                try
                {
                    if (item.LookupParameter("tipoDeViga").AsString() == "Retangular")
                    {

                        VigaRetangular v = new VigaRetangular();
                        v.tipoDeViga = item;
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

        public string GetNomeSecao(double largura, double altura)
        {


            string secao =  string.Format("{0:000}", Math.Round(largura* 0.3048*100, 0));
            string secao1 = string.Format("{0:000}", Math.Round(altura * 0.3048*100, 0));
            return secao + "x" + secao1;

        }

        public FamilySymbol CriarNovaVigaRetangular(string secao, Line largura, Line altura, FamilySymbol fsOriginal)
        {
            FamilySymbol fs = fsOriginal.Duplicate("Viga retangular de concreto "+secao)  as FamilySymbol;
            fs.LookupParameter("Largura").Set(largura.Length);
            fs.LookupParameter("Altura").Set(altura.Length);
            VigaRetangular v = new VigaRetangular();
            v.tipoDeViga = fs;
            v.altura = altura.Length;
            v.largura = largura.Length;
            v.secao = secao;
            listaDeVigas.Add(v);
            return fs;
        }


    }
}


