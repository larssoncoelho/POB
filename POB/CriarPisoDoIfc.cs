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

    public class TipoDeLaje
    {
        public FloorType tipoDeLaje;
        public string espessura;

    }

    [TransactionAttribute(TransactionMode.Manual)]
    [RegenerationAttribute(RegenerationOption.Manual)]

    public class CriarPisoDoIfc : IExternalCommand
    {
        Floor f;
        public List<TipoDeLaje> listaDeLaje = new List<TipoDeLaje>();
        public double tamanho;
        public string passo;
        public DataTable dataTable = new DataTable("Elementos");
        public double d;
        public double area;
        public List<FloorType> tiposDePiso;
        public double maiorDistancia = 0;
        public double menorDistancia = 0;
        public double espessura = 0;
        public int i1 = 0;
        public double area1;
        public double area2;
        double distanciaEmRelacaoAoZero;
        public List<Curve> listaDecurvas = new List<Curve>();
        public string nomeFamilia;
        public bool temAbertura = false;

        public Autodesk.Revit.UI.Result Execute(ExternalCommandData revit,
                  ref string message, ElementSet elements)
        {

            Face faceDoSolido;
            UIApplication uiApp = revit.Application;
            Document uiDoc = uiApp.ActiveUIDocument.Document;
            Selection sel = uiApp.ActiveUIDocument.Selection;
           Funcoes.Util.uiDoc = uiDoc;
            StringBuilder sb = new StringBuilder();
            List<ElementId> lista = new List<ElementId>();
        
        
           
            XYZ normal = new XYZ(0, 0, 1);
            // Transform offset = Transform.CreateTranslation((-.3048*.11)*normal);
            Transform offset = Transform.CreateTranslation(0 * normal);

            FilteredElementCollector pavimentos =Funcoes.Util.ObterLevels(uiDoc);
            
            
            CurveArray ca1 = new CurveArray();
            //GeraListaLaje(tiposDePiso);
            Transaction transaction1 = new Transaction(uiDoc, "CreateGenericModel1");

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
                            transaction1.Start();
                            faceDoSolido = Funcoes.Util.GetTopFace(solido);
                            //IList<CurveLoop> curveLoops  = faceDoSolido.GetEdgesAsCurveLoops();

                            maiorDistancia = Funcoes.Util.GetElevacaoRelacaoAoZeroTopFace(uiDoc.GetElement(eleId));
                            menorDistancia = Funcoes.Util.GetElevacaoRelacaoAoZero(uiDoc.GetElement(eleId));
                            espessura = maiorDistancia - menorDistancia;

                            var outerBoundary = faceDoSolido.EdgeLoops.get_Item(0);
                            ca1 = GetCurveArrayFromEdgeArary(outerBoundary);
                            if (faceDoSolido.EdgeLoops.Size > 1)
                            {
                                temAbertura = true;
                            }
                            else
                            {
                                temAbertura = false;
                            }

                            tiposDePiso = Funcoes.Util.ObterFamilyTypePorCategoria(uiDoc).Cast<FloorType>().ToList().
                                Where(x=>x.LookupParameter("Marca de tipo").AsString()=="1").ToList();
                            //FloorType tipoEscolhido = GetTipoDePiso(espessura);
                            var tipop = tiposDePiso.ToList().Cast<FloorType>().Where(x => x.Name.Contains(getNomeEspessuraConvertido(espessura))).ToList().FirstOrDefault();
                            if(tipop==null)
                            {
                               //TipoDeLaje f1 = listaDeLaje[0];
                               tipop = getNewFloorType(espessura, tiposDePiso.First());
                            }
                            FloorType tipoEscolhido = null;
#if D23 || D24

                            List<CurveLoop> curveLoops = faceDoSolido.GetEdgesAsCurveLoops().ToList();
                           


                            Floor f = Autodesk.Revit.DB.Floor.Create(uiDoc, curveLoops, tipop.Id, pavimento.Id);//uiDoc.Create.Newloor(curveArray, tipoPiso, baseLevel, false);
#else
                            f = uiDoc.Create.NewFloor(ca1, tipoEscolhido, pavimento, false);
#endif
                            var parUAU_COMP = uiDoc.GetElement(eleId).LookupParameter("UAU_COMP");
                            if ((parUAU_COMP != null)&(parUAU_COMP.HasValue)&(!string.IsNullOrEmpty(parUAU_COMP.AsString())))
                                        f.LookupParameter("UAU_COMP").Set(parUAU_COMP.AsString());


                            var parMarca = uiDoc.GetElement(eleId).get_Parameter(BuiltInParameter.ALL_MODEL_MARK);
                            if ((parMarca != null)&(parUAU_COMP.HasValue)&(!string.IsNullOrEmpty(parUAU_COMP.AsString())))
                                        f.LookupParameter("Marca").Set(parMarca.AsString());
                            try
                            {
                                f.LookupParameter(POB.Properties.Settings.Default.opbNome).Set((uiDoc.GetElement(eleId) as FamilyInstance).Name);
                            }
                            catch
                            {

                            }

                            transaction1.Commit();
                            if (!temAbertura)
                            {
                                transaction1.Start();
                                uiDoc.Delete(eleId);
                                transaction1.Commit();

                            }
                            else
                            {
                                try
                                {
                                    transaction1.Start();

                                    CreateFloorOpenings(uiDoc, faceDoSolido, f);

                                    transaction1.Commit();

                                    transaction1.Start();
                                    uiDoc.Delete(eleId);
                                    transaction1.Commit();

                                }
                                catch
                                {
                                    transaction1.RollBack();
                                    transaction1.Start();
                                    uiDoc.Delete(f.Id);
                                    transaction1.Commit();
                                }

                            }


                        }
                        

                    
                    }
                }
                catch (Exception e)
                {
                  transaction1.RollBack();

                }
            }


            return Result.Succeeded;

        }

        public FloorType GetTipoDePiso(double espessura)
        {
            string nome = "";
            FloorType tipoDePiso;
            foreach (TipoDeLaje tl in listaDeLaje)
            {
                nome = getNomeEspessuraConvertido(espessura);
                if (tl.espessura == nome)
                {
                    return tl.tipoDeLaje;
                }
            }
            TipoDeLaje f = listaDeLaje[0];
            return getNewFloorType(espessura, f.tipoDeLaje);
        }
        public FloorType getNewFloorType(double espessura, FloorType tipo)
        {
            FloorType novoTipo;

             novoTipo = (tipo.Duplicate("Laje em concreto armado " + Math.Round(espessura * 100, 4).ToString() + " cm") as FloorType);

             CompoundStructure novaEstrutura = novoTipo.GetCompoundStructure();
            foreach (CompoundStructureLayer item in novaEstrutura.GetLayers())
            {
                item.Function = MaterialFunctionAssignment.Structure;
                item.Width = Math.Round((espessura / 0.3048), 4);
            }
            novoTipo.SetCompoundStructure(novaEstrutura);
            novoTipo.LookupParameter("Marca de tipo").Set("1");
            return novoTipo;
        }

       
        public void GeraListaLaje(FilteredElementCollector filtro)
        {
            listaDeLaje.Clear();
            foreach (FloorType item in filtro)
            {

                if (item.Name.Contains("Laje em concreto armado"))
                {
                    if (item.LookupParameter("Espessura-padrão") != null)
                    {

                        TipoDeLaje tl = new TipoDeLaje();
                        tl.tipoDeLaje = item;
                        tl.espessura = getNomeEspessura(item.LookupParameter("Espessura-padrão").AsDouble());

                        listaDeLaje.Add(tl);

                    }
                }


         
            }

        }
        public string getNomeEspessura(double espessura)
        {
            return string.Format("{0:000}", Math.Round(espessura * 0.3048 * 100, 4));
        }
        public string getNomeEspessuraConvertido(double espessura)
        {
            return string.Format("{0:000}", Math.Round(espessura * 100, 4));
        }
        private CurveArray GetCurveArrayFromEdgeArary(EdgeArray edgeArray)
        {
            CurveArray curveArray =  new CurveArray();

            foreach (Edge edge in edgeArray)
            {
                var edgeCurve =       edge.AsCurve();

                curveArray.Append(edgeCurve);
            }

            return curveArray;
        }
        private void CreateFloorOpenings(Document doc, Face topFace,  Floor destFloor)
        {
            // looking if source floor has openings

                if (topFace.EdgeLoops.Size > 1)
                {
                    for (int i = 1; i < topFace.EdgeLoops.Size; i++)
                    {
                        var openingEdges =
                            topFace.EdgeLoops.get_Item(i);

                        var openingCurveArray =
                            GetCurveArrayFromEdgeArary(openingEdges);
                    
                        var opening =
                            doc
                                .Create
                                .NewOpening(destFloor,
                                            openingCurveArray,
                                            true);
                    }
                }

            
        }
    }
}

 
