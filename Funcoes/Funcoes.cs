using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using wf =  System.Windows.Forms;
using wDrawing =  System.Drawing;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using System.Xml;
using ObjetoDeTranferencia;


namespace Funcoes
{


   
    public class CmdRelationshipInverter

    {


        private Dictionary<ElementId, List<ElementId>> getElementIds(List<Element> elements)
        {
            Dictionary<ElementId, List<ElementId>> dict = new Dictionary<ElementId, List<ElementId>>();

            string fmt = "{0} is hosted by {1}";

            foreach (FamilyInstance fi in elements)
            {
                ElementId id = fi.Id;
                ElementId idHost = fi.Host.Id;
                if (!dict.ContainsKey(idHost))
                {
                    dict.Add(idHost, new List<ElementId>());
                }
                dict[idHost].Add(id);
            }
            return dict;
        }

        private void dumpHostedElements(Dictionary<ElementId, List<ElementId>> ids)
        {
            foreach (ElementId idHost in ids.Keys)
            {
                string s = string.Empty;

                foreach (ElementId id in ids[idHost])
                {
                    if (0 < s.Length)
                    {
                        s += ", ";
                    }
                    //         s += ElementDescription( id );
                }


            }
        }


    }



    public static class Resultado
    {
        static public List<object> vCampo = new List<object>();
        static public void Limpar()
        {
            vCampo.Clear();
        }

    }
    public static class Mestrado
    {

        public static void DitanciaWall(Document doc, UIDocument uidoc, Wall wall)
        {
            LocationCurve locCurve = wall.Location as LocationCurve;
            Curve curve = locCurve.Curve;
            // Find the direction of the tangent to the curve at the curve's midpoint
            // ComputeDerivates finds 3 vectors that describe the curve at the specified point
            // BasisX is the tangent vector
            // "true" means that the point on the curve will be described with a normalized value
            // 0 = one end of the curve, 0.5 = middle, 1.0 = other end of the curve
            XYZ curveTangent = curve.ComputeDerivatives(0.5, true).BasisX;
            // XYZ coordinate of one of the wall's endpoints
            XYZ wallEnd = curve.GetEndPoint(0);

            // Filter to find only doors
            ElementCategoryFilter filter = new ElementCategoryFilter(BuiltInCategory.OST_Doors);

            // ReferenceIntersector uses the specified filter, find elements, in a 3d view (which is the active view for this example)
            ReferenceIntersector intersector = new ReferenceIntersector(filter, FindReferenceTarget.Element, (View3D)doc.ActiveView);

            string doorDistanceInfo = "";
            // ReferenceIntersector.Find shoots a ray in the specified direction from the specified point
            // it returns a list of ReferenceWithContext objects
            foreach (ReferenceWithContext refWithContext in intersector.Find(wallEnd, curveTangent))
            {
                // Each ReferenceWithContext object specifies the distance from the ray's origin to the object (proximity)
                // and the reference of the element hit by the ray
                double proximity = refWithContext.Proximity;
                FamilyInstance door = doc.GetElement(refWithContext.GetReference()) as FamilyInstance;
                doorDistanceInfo += door.Symbol.Family.Name + " - " + door.Name + " = " + proximity + "\n";
            }
            TaskDialog.Show("ReferenceIntersector", doorDistanceInfo);

        }
    }
    

    public static class Util
    {
        public static Document uiDoc;
        public static double volumePorParametro = 0;
        public static FilteredElementCollector selecao;
        #region Unit conversion
        const double _convertFootToMm = 12 * 25.4;
        public static double area;
        public static double v;
        public static Level medicaoBlocoId;
        /// <summary>
        /// Convert a given length in millimetres to feet.
        /// </summary>


        public static SketchPlane CreateSketchPlane(Autodesk.Revit.DB.XYZ normal, Autodesk.Revit.DB.XYZ origin,
                                           Document m_familyDocument, UIApplication uiDoc)
        {
            // First create a Geometry.Plane which need in NewSketchPlane() method
            //to-do   Plane geometryPlane = uiDoc.Application.Create.NewPlane(normal, origin);
            //to-do      if (null == geometryPlane)  // assert the creation is successful
            {
                throw new Exception("Create the geometry plane failed.");
            }
            // Then create a sketch plane using the Geometry.Plane
            //to-do     SketchPlane plane = SketchPlane.Create(m_familyDocument, geometryPlane);
            // throw exception if creation failed
            //to-do      if (null == plane)
            {
                throw new Exception("Create the sketch plane failed.");
            }
            //to-do    return plane;
            return null;
        }


        public static SketchPlane CreateSketchPlane(Autodesk.Revit.DB.XYZ normal, Autodesk.Revit.DB.XYZ origin,
                                          UIApplication uiDoc)
        {
            // First create a Geometry.Plane which need in NewSketchPlane() method
            //to-do    Plane geometryPlane = uiDoc.Application.Create.NewPlane(normal, origin);
            //to-do    if (null == geometryPlane)  // assert the creation is successful
            {
                throw new Exception("Create the geometry plane failed.");
            }
            // Then create a sketch plane using the Geometry.Plane
            //SketchPlane plane = SketchPlane.Create(uiDoc.ActiveUIDocument.Document, geometryPlane);
            // throw exception if creation failed
            //to-do     if (null == plane)
            {
                throw new Exception("Create the sketch plane failed.");
            }
            //to-do     return plane;
            return null;
        }



        public static WallType NovaPateParede(Wall wall, CompoundStructureLayer camada)
        {
            WallType symbolNovo;

            try
            {
                Util.uiDoc = uiDoc;
                WallType symbol = Util.FindElementByName(typeof(WallType), wall.Name) as WallType;
                symbolNovo = symbol.Duplicate(wall.Name+ "_" + camada.Width.ToString() +"_"+camada.MaterialId.IntegerValue.ToString()) as WallType;
                CompoundStructure structure = wall.WallType.GetCompoundStructure();
                List<CompoundStructureLayer> lista = new List<CompoundStructureLayer>();
                lista.Add(camada);
                structure.SetLayers(lista);
                symbolNovo.SetCompoundStructure(structure);
                return symbolNovo;
            }
            catch (Exception e)
            {
                Util.uiDoc = uiDoc;

                return Util.FindElementByName( typeof(WallType), wall.Name + "_" + camada.Width.ToString() + "_" + camada.MaterialId.IntegerValue.ToString()) as WallType;
            }


        }

        

        public static void ProfileWall1(Wall wall, UIApplication uiapp, UIDocument uidoc, Application app)
        {


            Document doc = uidoc.Document;
            View view = doc.ActiveView;

            Autodesk.Revit.Creation.Application creapp
              = app.Create;

            Autodesk.Revit.Creation.Document credoc
              = doc.Create;

            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Wall Profile");

                // Get the external wall face for the profile
                // a little bit simpler than in the last realization

                Reference sideFaceReference = HostObjectUtils.GetSideFaces(wall, ShellLayerType.Exterior).First();

                Face face = (Face)wall.GetGeometryObjectFromReference(sideFaceReference);

                // The normal of the wall external face.

                XYZ normal = wall.Orientation;

                // Offset curve copies for visibility.

                Transform offset = Transform.CreateTranslation(
                  5 * normal);

                // If the curve loop direction is counter-
                // clockwise, change its color to RED.

                Color colorRed = new Color(255, 0, 0);

                // Get edge loops as curve loops.

                IList<CurveLoop> curveLoops = face.GetEdgesAsCurveLoops();

                foreach (var curveLoop in curveLoops)
                {
                    CurveArray curves = creapp.NewCurveArray();

                    foreach (Curve curve in curveLoop)
                        curves.Append(curve.CreateTransformed(offset));

                    // Create model lines for an curve loop.

                    Plane plane;//to-do  = creapp.NewPlane(curves);

                    //to-do    SketchPlane sketchPlane
                    //to-do    = SketchPlane.Create(doc, plane);

                /*    ModelCurveArray curveElements    = credoc.NewModelCurveArray(curves,   sketchPlane);

                    if (curveLoop.IsCounterclockwise(normal))
                        foreach (ModelCurve mcurve in curveElements)
                        {
                            OverrideGraphicSettings overrides
                                = view.GetElementOverrides(
                                    mcurve.Id);

                            overrides.SetProjectionLineColor(
                                                    colorRed);

                            view.SetElementOverrides(
                                mcurve.Id, overrides);
                        }
                        */
                }
            }
        }
        public static DataTable ServicosSelecionados(Selection sel, Document uiDoc)
        {
            DataTable dt = new DataTable();
            string texto = "";
            int contador = 0;
            dt.Columns.Add("tocServicoId", typeof(string));
            dt.Columns.Add("tocGrupoId", typeof(string));
            int servicoId;

            foreach (ElementId item in sel.GetElementIds())
            {
                
                try
                {
                    Element ele = uiDoc.GetElement(item);
                    servicoId = Convert.ToInt32(ele.LookupParameter("tocServicoId").AsString());
                    DataRow[] linhas = dt.Select("tocServicoId = '" + ele.LookupParameter("tocServicoId").AsString() + "'");
                    if (linhas.Length == 0)
                    {
                        DataRow dr = dt.NewRow();
                        dr["tocServicoId"] = ele.LookupParameter("tocServicoId").AsString();
                        dr["tocGrupoId"] = ele.LookupParameter("tocGrupoId").AsString();
                        dt.Rows.Add(dr);

                    }
                }
                catch
                {

                }
            }
            

            return dt;
        }

        public static DataTable GruposSelecionados(Selection sel, Document uiDoc)
        {
            DataTable dt = new DataTable();
            string texto = "";
            int contador = 0;
            dt.Columns.Add("tocGrupoId", typeof(string));
            int grupoId;

            foreach (ElementId item in sel.GetElementIds())
            {
                try
                {
                    Element ele = uiDoc.GetElement(item);
                    grupoId = Convert.ToInt32(ele.LookupParameter("tocGrupoId").AsString());

                    if (grupoId > 0)
                    {

                        DataRow[] linhas = dt.Select("tocGrupoId = '" + ele.LookupParameter("tocGrupoId").AsString() + "'");
                        if (linhas.Length == 0)
                        {
                            DataRow dr = dt.NewRow();
                            dr["tocGrupoId"] = ele.LookupParameter("tocGrupoId").AsString();
                            dt.Rows.Add(dr);
                        }
                    }
                }
                catch
                {

                }
            }
            return dt;
        }
        public static double MmToFoot(double length)
        {
            return length / _convertFootToMm;
        }
        #endregion // Unit conversion

        #region Formatting
        /// <summary>
        /// Return an English plural suffix for the given
        /// number of items, i.e. 's' for zero or more
        /// than one, and nothing for exactly one.
        /// </summary>

        public static List<View> GetListViewAvanco()
        {
           FilteredElementCollector selecaoView = new FilteredElementCollector(uiDoc).OfClass(typeof(Autodesk.Revit.DB.View));
            List<View> vistasDeAvanco = new List<View>();
            foreach (View view in selecaoView)
            {
                try
                {
                    if (view.AreGraphicsOverridesAllowed())
                        if (view.LookupParameter("tocVistaAvanco").AsValueString() == "Sim")
                            vistasDeAvanco.Add(view);
                }
                catch
                {

                }

            }
            return vistasDeAvanco;
        }

        public static string PluralSuffix(int n)
        {
            return 1 == n ? "" : "s";
        }

        /// <summary>
        /// Return a dot (full stop) for zero
        /// or a colon for more than zero.
        /// </summary>
        public static string DotOrColon(int n)
        {
            return 0 < n ? ":" : ".";
        }

        #region Geometrical Comparison
        const double _eps = 1.0e-9;


        public static double AreaLajeProjecao( string markPai, string nomePai, Element ele, 
            string campoMark, string campoArea)
        {
           
            double areaAcumulada = 0;
            if (ele is Autodesk.Revit.DB.Architecture.BuildingPad)
                selecao = new FilteredElementCollector(uiDoc).OfClass(typeof(Autodesk.Revit.DB.Architecture.BuildingPad));
            if (ele is Autodesk.Revit.DB.Floor)
                selecao = new FilteredElementCollector(uiDoc).OfClass(typeof(Autodesk.Revit.DB.Floor));
            if (ele is Autodesk.Revit.DB.FootPrintRoof)
                selecao = new FilteredElementCollector(uiDoc).OfClass(typeof(Autodesk.Revit.DB.FootPrintRoof));
            foreach (Element ele1 in selecao)
            {
                if (ele1.LookupParameter(campoMark).AsString() == markPai)
                {
                    if (ele1.Name == nomePai)
                    {
                        areaAcumulada = areaAcumulada + ele1.LookupParameter(campoArea).AsDouble();

                    }
                }
            }





            return areaAcumulada * 0.3048 * 0.3048;
        }
        

            public static string Retorno(string texto, string caracter, int numeroCampo)
            {
                string[] campoAlt = new string[1];
                int y = 0;

                string a = "";
                string b = "";

                for (int x = 0; x < texto.Length; x++)
                {
                    a = texto.Substring(x, 1);
                    if (a == caracter)
                    {
                        y = y + 1;
                        Array.Resize(ref campoAlt, campoAlt.Length + 1);
                        campoAlt[y - 1] = b;
                        b = "";
                    }
                    else
                    {
                        b = b + a;
                    }
                }


                return campoAlt[numeroCampo];
            }

       
        static public void AlterarGraficoElemento(List<View> selecaoView, OverrideGraphicSettings org, ElementId eleId)
        {
            foreach (View view in selecaoView)
            {

                try {
                    if (view.AreGraphicsOverridesAllowed())
                        if (view.LookupParameter("tocVistaAvanco").AsValueString() == "Sim")
                        {
                            view.SetElementOverrides(eleId, org);                       
                        }
                }
                catch (Exception e)
                {

                }
            }
        }

        static public PlanarFace GetTopFace(Solid solid)
        {
            PlanarFace topFace = null;
            FaceArray faces = solid.Faces;
            foreach (Face f in faces)
            {
                PlanarFace pf = f as PlanarFace;
                if (null != pf && Util.IsHorizontal(pf))
                {
                    if ((null == topFace) || (topFace.Origin.Z < pf.Origin.Z))
                    {
                        topFace = pf;
                    }
                }
            }
            return topFace;
        }


    

            /* double vol = 0;
             Options opt = ele.Document.Application.Create.NewGeometryOptions();
             GeometryElement ge12 = ele.get_Geometry(opt);
             try
             {

                 foreach (GeometryObject obj in ge12)
                 {
                     GeometryInstance gi = obj as GeometryInstance;
                     GeometryElement gh = gi.GetInstanceGeometry();
                     foreach (GeometryObject obj1 in gh)
                     {
                         Solid solid = obj1 as Solid;
                         if (null != solid)
                         {
                             vol = vol + solid.Volume * 0.3048 * 0.3048 * 0.3048;
                         }
                     }
                 }
             }
             catch
             {
                 Util.InfoMsg("erro interno");
                 return 0;
             }
             return vol;*/




        static public PlanarFace GetBottonFace(Solid solid)
        {
            PlanarFace bottonFace = null;
            FaceArray faces = solid.Faces;
            foreach (Face f in faces)
            {
                PlanarFace pf = f as PlanarFace;
                if (null != pf
                  && Util.IsHorizontal(pf))
                {
                    if ((null == bottonFace)
                      || (bottonFace.Origin.Z > pf.Origin.Z))
                    {

                        bottonFace = pf;
                    }
                }
            }
            return bottonFace;
        }
        public static FaceArray GetVerticalFace(Solid solid)
        {
            //List<PlanarFace> lista = new List<PlanarFace>();
            FaceArray faceArray = new FaceArray();
            // lista.Clear();
            FaceArray faces = solid.Faces;
            PlanarFace pf;
            foreach (Face f in faces)
            {

                pf = f as PlanarFace;
                if (null != pf && Util.IsVertical(pf))
                {
                    faceArray.Append(pf);
                }
            }

            return faceArray;
        }
        public static FaceArray GetHorizontalFace(Solid solid)
        {

            FaceArray faceArray = new FaceArray();

            PlanarFace horizontalFace = null;
            FaceArray faces = solid.Faces;
            PlanarFace pf;
            foreach (Face f in faces)
            {
                pf = f as PlanarFace;
                if (null != pf && Util.IsHorizontal(pf))
                {

                    faceArray.Append(pf);

                }
            }
            return faceArray;
        }
        public static double GetAreaVerical(FaceArray faceArray)
        {
            double valor = 0;
            foreach (PlanarFace pf in faceArray)
            {
                valor = valor + pf.Area;

            }
            return valor;
        }
        public static double GetVolumePorParametro(Element ele2)
        {



            foreach (Parameter param in ele2.Parameters)
            {
                try
                {
                    Util.InfoMsg(param.Definition.Name);

                    if (param.Definition.Name == "Volume")
                        volumePorParametro = param.AsDouble() * 0.3048 * 0.3048 * 0.3048;
                    continue;
                }
                catch
                {
                }

            }


            return volumePorParametro;
        }
        public static double GetAreaHorizontal(FaceArray faceArray)
        {
            double valor = 0;
            foreach (PlanarFace pf in faceArray)
            {
                valor = valor + pf.Area;
            }
            return valor;
        }
        public static string GetListaServicosId(ICollection<Autodesk.Revit.DB.ElementId> selecao, Document uiDoc)
        {
            List<string> lista = new List<string>();

            foreach (ElementId eleId in selecao)
            {

                if (!lista.Contains(uiDoc.GetElement(eleId).LookupParameter("tocServicoId").AsString()))

                    lista.Add(uiDoc.GetElement(eleId).LookupParameter("tocServicoId").AsString());




            }

            return lista.ToString();

        }
        public static double GetAreaFormaPilar(Element ele)
        {
            double area = 0;
            Options opt = ele.Document.Application.Create.NewGeometryOptions();
            GeometryElement ge12 = ele.get_Geometry(opt);

            foreach (GeometryObject obj in ge12)
            {
                GeometryInstance gi = obj as GeometryInstance;

                GeometryElement gh = gi.GetInstanceGeometry();

                foreach (GeometryObject obj1 in gh)
                {
                    Solid solid = obj1 as Solid;
                    if (null != solid)
                    {
                        area = area + GetAreaVerical(GetVerticalFace(solid));
                    }
                }
            }
            return area;
        }


        public static double GetAreaFormaLaje(Element ele)
        {
            double area = 0;
            Options opt = ele.Document.Application.Create.NewGeometryOptions();
            GeometryElement ge12 = ele.get_Geometry(opt);

            foreach (GeometryObject obj in ge12)
            {
                GeometryInstance gi = obj as GeometryInstance;

                GeometryElement gh = gi.GetInstanceGeometry();

                foreach (GeometryObject obj1 in gh)
                {
                    Solid solid = obj1 as Solid;
                    if (null != solid)
                    {
                        try
                        {

                            if (null != GetBottonFace(solid))
                                area = area + GetBottonFace(solid).Area;
                        }
                        catch (Exception e)
                        {
                            InfoMsg("GetBottonFace" + "\n" + e.Message);
                        }
                    }
                }
            }
            return area;
        }
        public static double GetElevacaoRelacaoAoZeroToFace(Element ele)
        {
            double elevacao = 0;
            foreach (Solid solid in Util.GetSolids(ele))
            {

                if (ele is Autodesk.Revit.DB.Plumbing.Pipe)
                {
                    LocationCurve lc = (ele as Autodesk.Revit.DB.Plumbing.Pipe).Location as LocationCurve;

                }
                else
                {
                    try
                    {
                        if (null != solid)
                        {

                            PlanarFace pf = Util.GetTopFace(solid);
                            foreach (CurveLoop cl in pf.GetEdgesAsCurveLoops())
                            {
                                CurveLoopIterator cli = cl.GetCurveLoopIterator();
                                Line l = (cli.Current as Line);
                                elevacao = l.GetEndPoint(0).Z * 0.3048;
                                continue;
                            }
                        }
                    }
                    catch

                    {
                    }
                }
            }
            return elevacao;
        }


        public static double GetElevacaoRelacaoAoZero(Element ele)
        {
            double elevacao = 0;
            foreach (Solid solid in Util.GetSolids(ele))
            {

                if (ele is Autodesk.Revit.DB.Plumbing.Pipe)
                {
                    LocationCurve lc = (ele as Autodesk.Revit.DB.Plumbing.Pipe).Location as LocationCurve;
                   
                }
                else
                {
                    try
                    {
                        if (null != solid)
                        {
                           
                            PlanarFace pf = Util.GetBottonFace(solid);
                            foreach (CurveLoop cl in pf.GetEdgesAsCurveLoops())
                            {
                                CurveLoopIterator cli = cl.GetCurveLoopIterator();
                                Line l = (cli.Current as Line);
                                elevacao = l.GetEndPoint(0).Z * 0.3048;
                                continue;
                            }
                        }
                    }
                    catch

                    {
                    }
                }
            }
            return elevacao;
        }


        public static double GetElevacaoRelacaoAoZeroTopFace(Element ele)
        {
            double elevacao = 0;
            foreach (Solid solid in Util.GetSolids(ele))
            {

                if (ele is Autodesk.Revit.DB.Plumbing.Pipe)
                {
                    LocationCurve lc = (ele as Autodesk.Revit.DB.Plumbing.Pipe).Location as LocationCurve;

                }
                else
                {
                    try
                    {
                        if (null != solid)
                        {

                            PlanarFace pf = Util.GetTopFace(solid);
                            foreach (CurveLoop cl in pf.GetEdgesAsCurveLoops())
                            {
                                CurveLoopIterator cli = cl.GetCurveLoopIterator();
                                Line l = (cli.Current as Line);
                                elevacao = l.GetEndPoint(0).Z * 0.3048;
                                continue;
                            }
                        }
                    }
                    catch

                    {
                    }
                }
            }
            return elevacao;
        }

        public static int GetLevelMaisProximo(Element ele, DataTable dt)

        {
            double menorLevel =-1;
            double d = Util.GetElevacaoRelacaoAoZero(ele);
            int medicaoBlocoId = 0;    

            
            foreach (DataRow dr in dt.Rows)
            {


                if ((null == -1) ||
                    (Math.Abs(d - menorLevel * 0.3048) > Math.Abs(d - Convert.ToDouble(dr["ELEVACAO"]) * 0.3048)))
                {
                    menorLevel = Convert.ToDouble(dr["ELEVACAO"]);
                    medicaoBlocoId = Convert.ToInt32(dr["medicao_bloco_id"]);
                }
            }
            return medicaoBlocoId;
        }
        public static Level GetNivelMaisProximo(Element ele, FilteredElementCollector filtro)

        {
            double menorLevel = -1;
            double d = Util.GetElevacaoRelacaoAoZero(ele);
         
            foreach (Level level in filtro)
            {
                if ((null == -1) ||
                    (Math.Abs(d - menorLevel * 0.3048) > Math.Abs(d - level.Elevation * 0.3048)))
                {
                    menorLevel = level.Elevation;
                    medicaoBlocoId = level;
                }
            }
            return medicaoBlocoId;
        }
        public static Level GetNivelMaisProximoToFace(Element ele, FilteredElementCollector filtro)

        {
            double menorLevel = -1;
            double d = Util.GetElevacaoRelacaoAoZeroToFace(ele);

            foreach (Level level in filtro)
            {
                if ((null == -1) ||
                    (Math.Abs(d - menorLevel * 0.3048) > Math.Abs(d - level.Elevation * 0.3048)))
                {
                    menorLevel = level.Elevation;
                    medicaoBlocoId = level;
                }
            }
            return medicaoBlocoId;
        }

        public static Autodesk.Revit.DB.Color GetColorRevit(wDrawing.Color colorWindows)
        {

            return new Autodesk.Revit.DB.Color(colorWindows.R, colorWindows.G, colorWindows.B);
        }

        public static List<Solid> GetSolids(Element ele)
        {
            List<Solid> lista = new List<Solid>();
            Options opt = ele.Document.Application.Create.NewGeometryOptions();
            opt.ComputeReferences = true;
            opt.DetailLevel = ViewDetailLevel.Fine;
            opt.IncludeNonVisibleObjects = true;
           // opt.View = ele.Document.ActiveView;
                
           GeometryElement ge12 = ele.get_Geometry(opt);
   
            foreach (GeometryObject obj in ge12)
            {
                try
                {

                    if (obj is Solid)
                    {
                        lista.Add((obj as Solid));
                    }
                    if (obj is GeometryInstance)
                    {

                        GeometryInstance gi = obj as GeometryInstance;
                        GeometryElement gh = gi.GetInstanceGeometry();
                        foreach (GeometryObject obj1 in gh)
                        {
                            Solid solid = obj1 as Solid;
                            if (null != solid)
                            {
                                lista.Add(solid);
                               
                            }
                        }
                    }  
                }
                catch
                {
                   
                }
            }

            return lista;
        }
        public static double GetAreaFormaEstrutura(Element ele1, string structuralColumn, string column,
                                             string structuralFraming, string floors)
        {
            area = 0;
            try
            {

                if (ele1.Category.Name == structuralColumn)
                {
                    area = GetAreaFormaPilar(ele1);
                    //InfoMsg("Categoria: " + ele1.Category.Name+"\n"+area.ToString());
                }
                if (ele1.Category.Name == column)
                {
                    area = GetAreaFormaPilar(ele1);
                    //InfoMsg("Categoria: " + ele1.Category.Name + "\n" + area.ToString());
                }
                if (ele1.Category.Name == structuralFraming)
                {
                    area = GetAreaFormaViga(ele1) + GetAreaFormaLaje(ele1);
                    //InfoMsg("Categoria: " + ele1.Category.Name + "\n" + area.ToString());
                }
                if (ele1.Category.Name == floors)
                {
                    area = GetAreaFormaLaje(ele1);
                    //InfoMsg("Categoria: " + ele1.Category.Name + "\n" + area.ToString());
                }
            }
            catch (Exception e)
            {

            }


            return area;
        }

        public static double AreaFacesFormaViga(FaceArray faceArray)
        {
            double area = 0;
            List<double> lista = new List<double>();

            foreach (PlanarFace pf in faceArray)
            {
                lista.Add(pf.Area);

            }
            lista.Sort();

            if (lista.Count > 0)
            {

                area = area + lista[lista.Count - 1];
                area = area + lista[lista.Count - 2];
            }


            return area;
        }

        public static Face GetSecaoTransversalViga(FaceArray faceArray)
        {
            Face secaoTransversalMenor = null;
            foreach (PlanarFace pf in faceArray)
            {
                if (null != pf && Util.IsVertical(pf))
                {
                    if (secaoTransversalMenor == null)
                    {
                        secaoTransversalMenor = pf;
                    }
                    else
                    {
                        if (secaoTransversalMenor.Area > pf.Area)
                        {
                            secaoTransversalMenor = pf;
                        }
                    }
                }
            }
            return secaoTransversalMenor;
        }

        public static Line GetMenorDimensao(Face face)
        {
            Line menorDimensao = null;
            foreach (CurveLoop curveLoop in face.GetEdgesAsCurveLoops())
            {

                CurveLoopIterator cli = curveLoop.GetCurveLoopIterator();
                while (cli.MoveNext())
                {
                    if (menorDimensao == null)
                    {
                        menorDimensao = (cli.Current as Line);
                    }
                    else
                    {
                        if (menorDimensao.Length > (cli.Current as Line).Length)
                            menorDimensao = (cli.Current as Line);
                    }

                }
            }

            return menorDimensao;

        }
        public static Line GetLinhaX(Face face)
        {
            foreach (CurveLoop curveLoop in face.GetEdgesAsCurveLoops())
            {
                CurveLoopIterator cli = curveLoop.GetCurveLoopIterator();
                while (cli.MoveNext())
                {
                    double m = 0; 
                    Line linhaAtual = (cli.Current as Line);
                    XYZ p1 = linhaAtual.GetEndPoint(0);
                    XYZ p2 = linhaAtual.GetEndPoint(1);
                    if ((p1.X - p2.X) > 0)
                    {
                        m = (p1.Y - p2.Y) / (p1.X - p2.X);
                        if (m > 0) return linhaAtual;
                    }
                    if(p1.Y==p2.Y)
                    {
                        return linhaAtual;
                    }
                }
            }
            return null;

        }

        public static Line GetLinhaY(Face face)
        {
            foreach (CurveLoop curveLoop in face.GetEdgesAsCurveLoops())
            {
                CurveLoopIterator cli = curveLoop.GetCurveLoopIterator();
                while (cli.MoveNext())
                {
                    double m = 0;
                    Line linhaAtual = (cli.Current as Line);
                    XYZ p1 = linhaAtual.GetEndPoint(0);
                    XYZ p2 = linhaAtual.GetEndPoint(1);
                    if ((p1.X - p2.X) > 0)
                    {
                        m = (p1.Y - p2.Y) / (p1.X - p2.X);
                        if (m <0) return linhaAtual;
                    }
                    if (p1.X == p2.X)
                    {
                        return linhaAtual;
                    }
                }
            }

            return null;

        }
        public static Line GetMaiorDimensao(Face face)
        {
            Line maiorDimensao = null;
            foreach (CurveLoop curveLoop in face.GetEdgesAsCurveLoops())
            {

                CurveLoopIterator cli = curveLoop.GetCurveLoopIterator();
                while (cli.MoveNext())
                {
                    if (maiorDimensao ==null)
                    {
                        maiorDimensao = (cli.Current as Line);
                    }
                    else
                    {
                        if (maiorDimensao.Length < (cli.Current as Line).Length)
                            maiorDimensao = (cli.Current as Line);
                    }

                }
            }

            return maiorDimensao;

        }


        public static double GetAreaFormaViga(Element ele)
        {
            
            double area = 0;
            Options opt = ele.Document.Application.Create.NewGeometryOptions();
            GeometryElement ge12 = ele.get_Geometry(opt);

            foreach (GeometryObject obj in ge12)
            {
                GeometryInstance gi = obj as GeometryInstance;

                GeometryElement gh = gi.GetInstanceGeometry();

                foreach (GeometryObject obj1 in gh)
                {
                    Solid solid = obj1 as Solid;
                    if (null != solid)
                    {
                        area = area + AreaFacesFormaViga(GetVerticalFace(solid));
                    }
                }
            }
            return area;
        }
        static public bool IsZero(double a)
        {
            return _eps > Math.Abs(a);
        }

        static public bool IsHorizontal(XYZ v)
        {
            return IsZero(v.Z);
        }

        static public bool IsVertical(XYZ v)
        {
            return IsZero(v.X) && IsZero(v.Y);
        }

        static public bool IsHorizontal(Edge e)
        {
            XYZ p = e.Evaluate(0);
            XYZ q = e.Evaluate(1);
            return IsHorizontal(q - p);
        }

        static public bool IsHorizontal(PlanarFace f)
        {
            return IsVertical(f.FaceNormal);
        }

        static public bool IsVertical(PlanarFace f)
        {
            return IsHorizontal(f.FaceNormal);
        }

        static public bool IsVertical(CylindricalFace f)
        {
            return IsVertical(f.Axis);
        }
        #endregion // Geometrical Comparison

        #region Formatting
        /// <summary>
        /// Return an English plural suffix 's' or 
        /// nothing for the given number of items.
        /// </summary>


        static public string AngleString(double a)
        {
            return RealString(a * 180 / Math.PI) + " degrees";
        }

        static public string PointString(XYZ p)
        {
            return string.Format("({0},{1},{2})",
              RealString(p.X), RealString(p.Y),
              RealString(p.Z));
        }

        static public string TransformString(Transform t)
        {
            return string.Format("({0},{1},{2},{3})", PointString(t.Origin),
              PointString(t.BasisX), PointString(t.BasisY), PointString(t.BasisZ));
        }
        #endregion // Formatting

        const string _caption = "The Building Coder";

        static public void InfoMsg(string msg)
        {
            //Debug.WriteLine( msg );
            System.Windows.Forms.MessageBox.Show(msg,
              _caption,
              System.Windows.Forms.MessageBoxButtons.OK,
              System.Windows.Forms.MessageBoxIcon.Information);
        }

        public static string ElementDescription(Element e)
        {
            // for a wall, the element name equals the 
            // wall type name, which is equivalent to the 
            // family name ...
            FamilyInstance fi = e as FamilyInstance;
            string fn = (null == fi)
              ? string.Empty
              : fi.Symbol.Family.Name + " ";

            string cn = (null == e.Category)
              ? e.GetType().Name
              : e.Category.Name;

            return string.Format("{0} {1}<{2} {3}>",
              cn, fn, e.Id.IntegerValue, e.Name);
        }

        public static Element SelectSingleElement(Document doc, string description)
        {
            //    UIApplication UIAPP = doc.Application as UIApplication;
            // Selection sel = UIAPP.ActiveUIDocument.Selection;
            Element e = null;
            //sel.Elements.Clear();
            //sel.StatusbarTip = "Please select " + description;
            /* if( sel.PickObject() )
             {
               ElementSetIterator elemSetItr
                 = sel.Elements.ForwardIterator();
               elemSetItr.MoveNext();
               e = elemSetItr.Current as Element;
             }*/
            return e;
        }

        public static bool GetSelectedElementsOrAll(
          List<Element> a,
          Document doc,
          Type t)
        {
            // Selection sel = (doc.Application as UIApplication).ActiveUIDocument.Selection;
            /* if( 0 < sel.Elements.Size )
             {
               foreach( Element e in sel.Elements )
               {
                 if( e.GetType().IsSubclassOf( t ) )
                 {
                   a.Add( e );
                 }
               }
             }
             else
             {
              // doc.get_Elements( t, a );
             }*/
            return 0 < a.Count;
        }

        /// <summary>
        /// Return a string for a real number
        /// formatted to two decimal places.
        /// </summary>



        public static TaskDialogResult PerguntarRevit()
        {

            // Creates a Revit task dialog to communicate information to the user.
            TaskDialog mainDialog = new TaskDialog("Hello, Revit!");
            mainDialog.MainInstruction = "Escolha uma opção";
            mainDialog.MainContent = "As opções abaixo determinam como os elementos serão deletados";

            // Add commmandLink options to task dialog
            mainDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1,
                     "Deletar elementos no AutoOrc");
            mainDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2,
                            "Não deletar elementos no AutoOrc");
            mainDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink3,
                               "Cancelar");


            // Set common buttons and default button. If no CommonButton or CommandLink is added,
            // task dialog will show a Close button by default
            mainDialog.CommonButtons = TaskDialogCommonButtons.Close;
            mainDialog.DefaultButton = TaskDialogResult.Close;

            // Set footer text. Footer text is usually used to link to the help document.
            mainDialog.FooterText = "Selecione uma opção";

            TaskDialogResult tResult = mainDialog.Show();
            return tResult;
        }

        public static Element GetInsulationPipeOuflexPipe(Document doc, string name)
        {
            Util.uiDoc = uiDoc;

            return Util.FindElementByName( typeof(Autodesk.Revit.DB.Plumbing.PipeInsulationType), name);
        }
        public static Element GetInsulationDuct( string name)
        {
            Util.uiDoc = uiDoc;

            return Util.FindElementByName( typeof(Autodesk.Revit.DB.Mechanical.DuctInsulationType), name);
        }

        public static void RetirarInsulation(Element ele)
        {
            if (ele is Autodesk.Revit.DB.Plumbing.Pipe)
            {
                FilteredElementCollector fec = new FilteredElementCollector(ele.Document).OfClass(typeof(Autodesk.Revit.DB.Plumbing.PipeInsulation));
                foreach (Autodesk.Revit.DB.Plumbing.PipeInsulation pi in fec)
                {
                    if (pi.HostElementId == ele.Id)
                    {
                        ele.Document.Delete(pi.Id);

                        break;
                    }
                }
            }
            if (ele is Autodesk.Revit.DB.Plumbing.FlexPipe)
            {
                FilteredElementCollector fec = new FilteredElementCollector(ele.Document).OfClass(typeof(Autodesk.Revit.DB.Plumbing.PipeInsulation));
                foreach (Autodesk.Revit.DB.Plumbing.PipeInsulation pi in fec)
                {
                    if (pi.HostElementId == ele.Id)
                    {
                        ele.Document.Delete(pi.Id);
                        break;
                    }
                }
            }
            if (ele is Autodesk.Revit.DB.Mechanical.Duct)
            {
                FilteredElementCollector fec = new FilteredElementCollector(ele.Document).OfClass(typeof(Autodesk.Revit.DB.Mechanical.DuctInsulation));
                foreach (Autodesk.Revit.DB.Mechanical.DuctInsulation pi in fec)
                {
                    if (pi.HostElementId == ele.Id)
                    {
                        ele.Document.Delete(pi.Id);
                        break;
                    }
                }
            }

        }

        public static string RealString(double a)
        {
            return a.ToString("0.##");
        }

        /// <summary>
        /// Return a string for an XYZ point
        /// or vector with its coordinates
        /// formatted to two decimal places.
        /// </summary>

        #endregion // Formatting

        #region Display a message
        public const string Caption = "Family API";



        public static void InfoMsg2(
          string instruction,
          string content)
        {

            TaskDialog d = new TaskDialog(Caption);
            d.MainInstruction = instruction;
            d.MainContent = content;
            d.Show();
        }

        public static void ErrorMsg(string msg)
        {

            TaskDialog d = new TaskDialog(Caption);
            d.MainIcon = TaskDialogIcon.TaskDialogIconWarning;
            d.MainInstruction = msg;
            d.Show();
        }

        /// <summary>
        /// Return a string describing the given element:
        /// .NET type name,
        /// category name,
        /// family and symbol name for a family instance,
        /// element id and element name.
        /// </summary>
        public static string ElementDescription1(
          Element e)
        {
            if (null == e)
            {
                return "<null>";
            }

            // For a wall, the element name equals the
            // wall type name, which is equivalent to the
            // family name ...

            FamilyInstance fi = e as FamilyInstance;

            string typeName = e.GetType().Name;

            string categoryName = (null == e.Category)
              ? string.Empty
              : e.Category.Name + " ";

            string familyName = (null == fi)
              ? string.Empty
              : fi.Symbol.Family.Name + " ";

            string symbolName = (null == fi
              || e.Name.Equals(fi.Symbol.Name))
                ? string.Empty
                : fi.Symbol.Name + " ";

            return string.Format("{0} {1}{2}{3}<{4} {5}>",
              typeName, categoryName, familyName,
              symbolName, e.Id.IntegerValue, e.Name);
        }
        #endregion // Display a message




        public static List<Family> FindFamilyByLevel(
          Document doc, Level level)
        {

            return null;
            /* new FilteredElementCollector(doc)
               .OfClass(targetType)
               .FirstOrDefault<Element>(
                 e => e.Name.Equals(targetName));*/
        }

        #region Retrieve an element by class and name
        /// <summary>
        /// Retrieve a database element 
        /// of the given type and name.
        /// </summary>
        public static Element FindElementByName(
          Type targetType,
          string targetName)
        {
            return new FilteredElementCollector(uiDoc)
              .OfClass(targetType)
              .FirstOrDefault<Element>(
                e => e.Name.Equals(targetName));
        }
        #endregion // Retrieve an element by class and name





        static public IndependentTag CreateIndependentTag(Autodesk.Revit.DB.Document document, Wall wall)
        {
            // make sure active view is not a 3D view
            Autodesk.Revit.DB.View view = document.ActiveView;
            // define tag mode and tag orientation for new tag
            TagMode tagMode = TagMode.TM_ADDBY_CATEGORY;
            TagOrientation tagorn = TagOrientation.Horizontal;

            // Add the tag to the middle of the wall
            LocationCurve wallLoc = wall.Location as LocationCurve;
            XYZ wallStart = wallLoc.Curve.GetEndPoint(0);
            XYZ wallEnd = wallLoc.Curve.GetEndPoint(1);
            XYZ wallMid = wallLoc.Curve.Evaluate(0.5, true);

         /*   IndependentTag newTag = document.Create.NewTag(view, wall, true, tagMode, tagorn, wallMid);
            if (null == newTag)
            {
                throw new Exception("Create IndependentTag Failed.");
            }*/

            // newTag.TagText is read-only, so we change the Type Mark type parameter to
            // set the tag text. The label parameter for the tag family determines
            // what type parameter is used for the tag text.

            WallType type = wall.WallType;

            Parameter foundParameter = type.LookupParameter("");
            bool result = foundParameter.Set("Hello");

            // set leader mode free
            // otherwise leader end point move with elbow point
            /*
                        newTag.LeaderEndCondition = LeaderEndCondition.Free;
                        XYZ elbowPnt = wallMid + new XYZ(5.0, 5.0, 0.0);
                        newTag.LeaderElbow = elbowPnt;
                        XYZ headerPnt = wallMid + new XYZ(10.0, 10.0, 0.0);
                        newTag.TagHeadPosition = headerPnt;

                        return newTag;*/
            return null;
        }

        public static string RegistraDadosModelo(Document uiDoc, DadosIntegracao dadosIntegracao)
        {
            /*  if (!FuncaoNBR15575.ModeloCadastrado(uiDoc))
              {

                  try
                  {

                      if (uiDoc.IsWorkshared)
                      {
                          if (wf.MessageBox.Show("Deseja Sincronizar?", "Sair", wf.MessageBoxButtons.YesNo, wf.MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes)
                          {
                              TransactWithCentralOptions tco = new TransactWithCentralOptions();

                              uiDoc.SynchronizeWithCentral(tco, new SynchronizeWithCentralOptions { SaveLocalAfter = true });
                          }
                      }
                  }
                  catch
                  {
                      TaskDialog.Show("_", "Não foi possível sincronizar");
                  }


                  string r = RegistraModelo(dadosIntegracao, uiDoc);

                  if (string.IsNullOrEmpty(r))
                      return null;


              }



              string guid = FuncaoNBR15575.GetGUIDModelo(uiDoc);
              int obraId;
              if (!int.TryParse(FuncaoNBR15575.GetObraIdModelo(uiDoc), out obraId))
                  throw new System.InvalidOperationException("Erro ao obter o ObraId");


              new ACESSO_MODELO_OBRA(dadosIntegracao.Diretorio, dadosIntegracao).
                   Sincronizar(new MODELO_OBRA
                   {
                       MODELO_GUID_ID = guid,
                       NOME_DO_ARQUIVO = uiDoc.PathName,
                       NOME_DO_PROJETO = uiDoc.ProjectInformation.Name,
                       OBRA_ID = obraId
                   }, Acao.atualizar, dadosIntegracao);
              */
            var guid = ""; 
            return guid;
        }
    

    static public void CriarMaterialFuncao(Document uiDoc)
        {
            

            try
            {

                Material defaultm = Util.FindElementByName( typeof(Material), "Default") as Material;
                Material Previsto = defaultm.Duplicate("Previsto");
                Autodesk.Revit.DB.Color cor = new Autodesk.Revit.DB.Color(255, 128, 0);
                Previsto.Color = cor;
#if DEBUG20201

#else
#endif
#if debug20211
#else
         
              /*  Previsto.SurfacePatternColor = cor;
                Previsto.CutPatternColor = cor;
                Element preenchimento = Util.FindElementByName(typeof(Element), "Preenchimento sólido") as Element;
                Previsto.SurfacePatternId = preenchimento.Id;
              */

#endif

                cor.Dispose();
            }
            catch
            {
                Material Previsto = Util.FindElementByName( typeof(Material), "Previsto") as Material;
                Autodesk.Revit.DB.Color cor = new Autodesk.Revit.DB.Color(255, 128, 0);
                Previsto.Color = cor;
#if DEBUG20201
#endif
#if DEBUG20211
#else

                //Previsto.SurfacePatternColor = cor;
               // Previsto.CutPatternColor = cor;
#endif
                cor.Dispose();
            }
            try
            {
                Material previsto = Util.FindElementByName( typeof(Material), "Default") as Material;
                Material Completo = previsto.Duplicate("Completo");
                Autodesk.Revit.DB.Color cor = new Autodesk.Revit.DB.Color(0, 0, 255);
                Completo.Color = cor;
#if DEBUG20201

#else
            
               // Completo.SurfacePatternColor = cor;
               // Completo.CutPatternColor = cor;
#endif
                cor.Dispose();
            }
            catch
            {

            }







            try
            {
                Material previsto = Util.FindElementByName( typeof(Material), "Default") as Material;
                Material Completo = previsto.Duplicate("Concluido");
                Autodesk.Revit.DB.Color cor = new Autodesk.Revit.DB.Color(0, 0, 255);
                Completo.Color = cor;

#if DEBUG20201

#else
                
          //      Completo.SurfacePatternColor = cor;
           //     Completo.CutPatternColor = cor;
#endif
                cor.Dispose();
            }
            catch
            {

            }
            try
            {
                Material previsto = Util.FindElementByName( typeof(Material), "Default") as Material;
                Material iniciado = previsto.Duplicate("Iniciado");
                Autodesk.Revit.DB.Color cor = new Autodesk.Revit.DB.Color(255, 255, 0);
                iniciado.Color = cor;
         #if DEBUG20201

#else
            
            //    iniciado.CutPatternColor = cor;
             //   iniciado.SurfacePatternColor = cor;
#endif
                cor.Dispose();

            }
            catch
            {

            }
            try
            {
                Material previsto = Util.FindElementByName( typeof(Material), "Default") as Material;
                Material iniciado = previsto.Duplicate("Projecao");
                Autodesk.Revit.DB.Color cor = new Autodesk.Revit.DB.Color(0, 128, 0);
                iniciado.Color = cor;
            #if DEBUG20201

#else
            
         //        iniciado.SurfacePatternColor = cor;
#endif
                cor.Dispose();

            }

            catch
            {

            }
            try
            {
                Material previsto = Util.FindElementByName( typeof(Material), "Default") as Material;
                Material iniciado = previsto.Duplicate("Projecao1");
                Autodesk.Revit.DB.Color cor = new Autodesk.Revit.DB.Color(255, 0, 255);
                iniciado.Color = cor;
#if DEBUG20201

#else
         //       iniciado.CutPatternColor = cor;
          //      iniciado.SurfacePatternColor = cor;
#endif
                cor.Dispose();

            }

            catch
            {

            }
            try
            {
                Material previsto = Util.FindElementByName( typeof(Material), "Default") as Material;
                Material iniciado = previsto.Duplicate("Projecao2");
                Autodesk.Revit.DB.Color cor = new Autodesk.Revit.DB.Color(0, 255, 255);
                iniciado.Color = cor;
#if DEBUG20201

#else
            
              //  iniciado.CutPatternColor = cor;
            //    iniciado.SurfacePatternColor = cor;
#endif
                cor.Dispose();

            }

            catch
            {

            }
            try
            {
                Material previsto = Util.FindElementByName( typeof(Material), "Default") as Material;
                Material iniciado = previsto.Duplicate("Projecao3");
                Autodesk.Revit.DB.Color cor = new Autodesk.Revit.DB.Color(128, 128, 0);
                iniciado.Color = cor;
#if DEBUG20201

#else
            
          //      iniciado.CutPatternColor = cor;
           //     iniciado.SurfacePatternColor = cor;
#endif
                cor.Dispose();

            }

            catch
            {

            }

            try
            {
                Material previsto = Util.FindElementByName(typeof(Material), "Default") as Material;
                Material iniciado = previsto.Duplicate("Projecao4");
                Autodesk.Revit.DB.Color cor = new Autodesk.Revit.DB.Color(0, 128, 0);
                iniciado.Color = cor;
#if DEBUG20201

#else
            
          //      iniciado.CutPatternColor = cor;
           //     iniciado.SurfacePatternColor = cor;
#endif
                cor.Dispose();

            }
            catch
            {

            }
            try
            {

                Autodesk.Revit.DB.Plumbing.PipeInsulationType insulacaoPrevisto =
                Util.FindElementByName( typeof(Autodesk.Revit.DB.Plumbing.PipeInsulationType), "previsto") as Autodesk.Revit.DB.Plumbing.PipeInsulationType;
                Autodesk.Revit.DB.Plumbing.PipeInsulationType insulacaoProjecao = insulacaoPrevisto.Duplicate("Projecao") as Autodesk.Revit.DB.Plumbing.PipeInsulationType;
                Autodesk.Revit.DB.Color cor = new Autodesk.Revit.DB.Color(128, 255, 0);
                insulacaoProjecao.LookupParameter("Material").Set((Util.FindElementByName( typeof(Material), "projecao") as Material).Id);
            }
            catch
            {

            }


            try
            {

                Autodesk.Revit.DB.Plumbing.PipeInsulationType insulacaoPrevisto =
                Util.FindElementByName( typeof(Autodesk.Revit.DB.Plumbing.PipeInsulationType), "previsto") as Autodesk.Revit.DB.Plumbing.PipeInsulationType;
                Autodesk.Revit.DB.Plumbing.PipeInsulationType insulacaoProjecao = insulacaoPrevisto.Duplicate("Completo") as Autodesk.Revit.DB.Plumbing.PipeInsulationType;
                Autodesk.Revit.DB.Color cor = new Autodesk.Revit.DB.Color(0, 0, 255);
                insulacaoProjecao.LookupParameter("Material").Set((Util.FindElementByName(typeof(Material), "Completo") as Material).Id);
            }
            catch
            {

            }
        }


        static public void AlterarLinhaReferenciaParede(Element elemento, int valorNovo)
        {
            Transaction trans = new Transaction(elemento.Document);

            trans.Start("Lab");
            elemento.LookupParameter("").Set(valorNovo);
            trans.Commit();
            trans.Dispose();



        }
        static public void CriarFiltros(UIDocument uiDoc)
        {
            //UIDocument uidoc = ActiveUIDocument;
            Document doc = uiDoc.Document;
            View view = doc.ActiveView;

            TaskDialog.Show("Boost Your BIM", "# of filters applied to this view = " + view.GetFilters().Count);

            // create list of categories that will for the filter
            IList<ElementId> categories = new List<ElementId>();
            categories.Add(new ElementId(BuiltInCategory.OST_Walls));

            // create a list of rules for the filter
            IList<FilterRule> rules = new List<FilterRule>();
            // This filter will have a single rule that the wall type width must be less than 1/2 foot
            Parameter wallWidth = new FilteredElementCollector(doc).OfClass(typeof(WallType)).FirstElement().get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM);
            rules.Add(ParameterFilterRuleFactory.CreateLessRule(wallWidth.Id, 0.5, 0.001));

            ParameterFilterElement filter = null;
            using (Transaction t = new Transaction(doc, "Create and Apply Filter"))
            {
                t.Start();
#if DEBUG20201

#else

             //   filter = ParameterFilterElement.Create(doc, "Thin Wall Filter", categories, rules);
#endif
                view.AddFilter(filter.Id);
                t.Commit();
            }

            string filterNames = "";
            foreach (ElementId id in view.GetFilters())
            {
                filterNames += doc.GetElement(id).Name + "\n";
            }
            TaskDialog.Show("Boost Your BIM", "Filters applied to this view: " + filterNames);

            // Create a new OverrideGraphicSettings object and specify cut line color and cut line weight
            OverrideGraphicSettings ogs = new OverrideGraphicSettings();
            ogs.SetCutLineColor(new Color(255, 0, 0));
            ogs.SetCutLineWeight(9);

            using (Transaction t = new Transaction(doc, "Set Override Appearance"))
            {
                t.Start();
                view.SetFilterOverrides(filter.Id, ogs);
                t.Commit();
            }
        }
        
        static public FilteredElementCollector ObterLevels(Document uiDoc)
        {

            FilteredElementCollector selecao = new FilteredElementCollector(uiDoc).OfClass(typeof(Autodesk.Revit.DB.Level));
           

            return selecao;
        }
        static public FilteredElementCollector ObterFamilyTypePorCategoria(Document uiDoc)
        {

            FilteredElementCollector selecao = new FilteredElementCollector(uiDoc).OfClass(typeof(Autodesk.Revit.DB.FloorType));
                   

            return selecao;
        }



        static public void ObterMaterialLayerParede(Document documento, Element elemento)
        {

            GetFamilyTypesInFamily(documento);

            // Element elemento;
            //elemento.
            Wall wall = elemento as Wall;
            WallType aWallType = wall.WallType;
            //somente se aperede for básica

            if (WallKind.Basic == aWallType.Kind)
            {
                //get CompoundStructure
                CompoundStructure comStruct = aWallType.GetCompoundStructure();
                Categories allCategories = documento.Settings.Categories;
                // Get the category OST_Walls default Material; 
                // use if that layer's default Material is <By Category>
                Category wallCategory = allCategories.get_Item(BuiltInCategory.OST_Walls);
                Autodesk.Revit.DB.Material wallMaterial = wallCategory.Material;

                foreach (CompoundStructureLayer structLayer in comStruct.GetLayers())
                {
                    Autodesk.Revit.DB.Material layerMaterial =
                    documento.GetElement(structLayer.MaterialId) as Material;

                    // If CompoundStructureLayer's Material is specified, use default
                    // Material of its Category
                    if (null == layerMaterial)
                    {
                        switch (structLayer.Function)
                        {
                            case MaterialFunctionAssignment.Finish1:
                                layerMaterial =
                                        allCategories.get_Item(BuiltInCategory.OST_WallsFinish1).Material;
                                break;
                            case MaterialFunctionAssignment.Finish2:
                                layerMaterial =
                                        allCategories.get_Item(BuiltInCategory.OST_WallsFinish2).Material;
                                break;
                            case MaterialFunctionAssignment.Membrane:
                                layerMaterial =
                                        allCategories.get_Item(BuiltInCategory.OST_WallsMembrane).Material;
                                break;
                            case MaterialFunctionAssignment.Structure:
                                layerMaterial =
                                        allCategories.get_Item(BuiltInCategory.OST_WallsStructure).Material;
                                break;
                            case MaterialFunctionAssignment.Substrate:
                                layerMaterial =
                                        allCategories.get_Item(BuiltInCategory.OST_WallsSubstrate).Material;
                                break;
                            case MaterialFunctionAssignment.Insulation:
                                layerMaterial =
                                        allCategories.get_Item(BuiltInCategory.OST_WallsInsulation).Material;
                                break;
                            default:
                                // It is impossible to reach here
                                break;
                        }
                        if (null == layerMaterial)
                        {
                            // CompoundStructureLayer's default Material is its SubCategory
                            layerMaterial = wallMaterial;
                        }
                    }
                    TaskDialog.Show("Revit", "Layer Material: " + layerMaterial);
                }
            }
        }
        static public void GetFamilyTypesInFamily(Document familyDoc)
        {
            if (familyDoc.IsFamilyDocument == true)
            {
                FamilyManager familyManager = familyDoc.FamilyManager;
                // get types in family
                string types = "Family Types: ";
                FamilyTypeSet familyTypes = familyManager.Types;
                FamilyTypeSetIterator familyTypesItor = familyTypes.ForwardIterator();
                familyTypesItor.Reset();
                while (familyTypesItor.MoveNext())
                {
                    FamilyType familyType = familyTypesItor.Current as FamilyType;
                    types += "\n" + familyType.Name;
                }
                wf.MessageBox.Show(types, "Revit");
            }
        }

        //inicio

        static public void UpdateLeaders(Document doc)
        {
            //UIDocument uidoc = this.ActiveUIDocument;

            // create a collection of all the Arrowhead types
            ElementId id = new ElementId(BuiltInParameter.ALL_MODEL_FAMILY_NAME);
            ParameterValueProvider provider = new ParameterValueProvider(id);
            FilterStringRuleEvaluator evaluator = new FilterStringEquals();
#if D23 || D24
            FilterRule rule = new FilterStringRule(provider, evaluator, "Basic Wall");
#else
            FilterRule rule = new FilterStringRule(provider, evaluator, "Basic Wall", false);
#endif           
            ElementParameterFilter filter = new ElementParameterFilter(rule);
            FilteredElementCollector collector2 = new FilteredElementCollector(doc).OfClass(typeof(ElementType)).WherePasses(filter);

            ElementId arrowheadId = null;

            // loop through the collection of Arrowhead types and get the Id of a specific one
            foreach (ElementId eid in collector2.ToElementIds())
            {
                if (doc.GetElement(eid).Name == "Wall 2") //change the name to the arrowhead type you want to use
                {
                    arrowheadId = eid;
                    TaskDialog.Show("Entrou", "Sim?");

                }
            }



            // create a collection of all families
            List<Family> families = new List<Family>(new FilteredElementCollector(doc).OfClass(typeof(Family)).Cast<Family>());
            using (Transaction t = new Transaction(doc, "Update Leaders"))
            {
                t.Start();
                // loop through all families and get their types       
                foreach (Family f in families)
                {
                    ISet<ElementId> symbols = f.GetFamilySymbolIds();

                    // loop through the family types and try switching the parameter Leader Arrowhead to the one chosen above
                    // if the type doesn't have this parameter it will skip to the next family type
                    foreach (ElementId eleId in symbols)
                    {
                        FamilySymbol fs = doc.GetElement(eleId) as FamilySymbol;
                        try

                        {
                            fs.LookupParameter("Name").Set(arrowheadId);
                            TaskDialog.Show("Entrou", "Sim?");
                        }
                        catch (Exception exception)
                        {
                            TaskDialog.Show("Entrou", arrowheadId.ToString());
                            TaskDialog.Show("Entrou", exception.Message);

                        }
                    }
                }
                t.Commit();
            }
        }


        [TransactionAttribute(TransactionMode.Manual)]
        [RegenerationAttribute(RegenerationOption.Manual)]
        public class VincularTubo : IExternalCommand
        {
            public static Autodesk.Revit.DB.Plumbing.Pipe tuboAvanco; //= new Autodesk.Revit.DB.Plumbing.Pipe();
            public static Autodesk.Revit.DB.Plumbing.Pipe tuboLevantamento; //= new Autodesk.Revit.DB.Plumbing.Pipe();
            public static string PLAN_SERVICO_AMO;
            public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
            {
                UIApplication uiApp = commandData.Application;
                Document uiDoc = uiApp.ActiveUIDocument.Document;
                // selecionar
                Selection sel = uiApp.ActiveUIDocument.Selection;
                int contador = 0;

                Element ele;
                //FuncoesDiversas.UpdateLeaders(uiDoc);
                foreach (ElementId eleId in sel.GetElementIds())
                {
                    ele = uiDoc.GetElement(eleId);
                    string texto = string.Format("Name: {0} | {1} | {2} \nCategoria: ", ele.Name, ele.Category.Name, ele.GetType());
                    TaskDialog.Show("Informação do elemento", texto);
                    Transaction trans = new Transaction(uiDoc);

                    switch (contador)

                    {
                        case 0:
                            tuboAvanco = ele as Autodesk.Revit.DB.Plumbing.Pipe;
                            break;
                        case 1:
                            tuboLevantamento = ele as Autodesk.Revit.DB.Plumbing.Pipe;
                            PLAN_SERVICO_AMO = ele.LookupParameter("Mark").AsString();
                            PLAN_SERVICO_AMO = ele.GetParameters("Mark")[0].AsValueString();
                            break;
                    }
                    contador = contador + 1;
                    WallType symbol = Util.FindElementByName(typeof(WallType), "Wall 2") as WallType;

                }

                if (PLAN_SERVICO_AMO == "")
                {
                    return Result.Failed;
                }

                Transaction trans1 = new Transaction(uiDoc);

                trans1.Start("Lab");
                tuboAvanco.LookupParameter("Mark").SetValueString(PLAN_SERVICO_AMO);
                trans1.Commit();
                trans1.Dispose();


                return Result.Succeeded;

            }
        }

        //termino
    }







}

