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

    public class CriarForroAPartirdoPiso : IExternalCommand
    {
        Floor f;
        public List<TipoDeLaje> listaDeLaje = new List<TipoDeLaje>();
        public double tamanho;
        public string passo;
        public DataTable dataTable = new DataTable("Elementos");
        public double d;
        public double area;
        public List<CeilingType> tiposDePiso;
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

            


            CurveArray ca1 = new CurveArray();
            //GeraListaLaje(tiposDePiso);
            Transaction transaction1 = new Transaction(uiDoc, "CreateGenericModel1");

            foreach (ElementId eleId in uiApp.ActiveUIDocument.Selection.GetElementIds())
            {
                if (uiDoc.GetElement(eleId) is Floor)
                {
                    try
                    {
                        List<Solid> listaDeSolidos = Funcoes.Util.GetSolids(uiDoc.GetElement(eleId));
                        foreach (Solid solido in listaDeSolidos)
                        {
                            if (solido.Faces.Size > 0)
                            {
                                transaction1.Start();
                                faceDoSolido = Funcoes.Util.GetTopFace(solido);
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

                                tiposDePiso = new FilteredElementCollector(uiDoc).OfClass(typeof(Autodesk.Revit.DB.CeilingType)).Cast<CeilingType>().ToList();
                                //FloorType tipoEscolhido = GetTipoDePiso(espessura);
                                var tipop = tiposDePiso[0];
                                CeilingType tipoEscolhido = null;
#if D23 || D24

                                List<CurveLoop> curveLoops = faceDoSolido.GetEdgesAsCurveLoops().ToList();



                                Ceiling f = Autodesk.Revit.DB.Ceiling.Create(uiDoc, curveLoops, tipop.Id, (uiDoc.GetElement(eleId) as Floor).LevelId);//uiDoc.Create.Newloor(curveArray, tipoPiso, baseLevel, false);
                                transaction1.Commit();

#else
                            f = uiDoc.Create.NewFloor(ca1, tipoEscolhido, pavimento, false);
#endif
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        transaction1.RollBack();

                    }
                }
            }
           
            return Result.Succeeded;

        }

        private CurveArray GetCurveArrayFromEdgeArary(EdgeArray edgeArray)
        {
            CurveArray curveArray = new CurveArray();

            foreach (Edge edge in edgeArray)
            {
                var edgeCurve = edge.AsCurve();

                curveArray.Append(edgeCurve);
            }

            return curveArray;
        }
        private void CreateFloorOpenings(Document doc, Face topFace, Floor destFloor)
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



