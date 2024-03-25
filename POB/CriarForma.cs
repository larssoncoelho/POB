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

using Autodesk.Revit.ApplicationServices;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Window;


namespace POB
{

    [TransactionAttribute(TransactionMode.Manual)]
    [RegenerationAttribute(RegenerationOption.Manual)]

    public class CriarForma : IExternalCommand
    {
        private ElementId baseLevelId;

        public Autodesk.Revit.UI.Result Execute(ExternalCommandData revit, ref string message, ElementSet elements)
        {
            UIApplication uiApp = revit.Application;
            Document uiDoc = uiApp.ActiveUIDocument.Document;
            Selection sel = uiApp.ActiveUIDocument.Selection;
            XYZ P = new XYZ(0, 0, 0);
            var nomeParede = "Forma";
            Util.uiDoc = uiDoc;
            var listaLevel = Util.ListaDeNiveis(uiDoc);

            WallType wallType = Util.FindElementByName(typeof(WallType), nomeParede) as WallType;
            ElementId tipoWall = wallType.Id;

            FloorType tipoPiso = Util.FindElementByName(typeof(FloorType), nomeParede) as FloorType;
            Level baseLevel;
            Autodesk.Revit.Creation.Application creapp = uiApp.Application.Create;

            List<Element> listaVigas = new List<Element>();
            List<Element> listaPilar = new List<Element>();
            List<Element> listaPiso = new List<Element>();

            Transaction transaction = new Transaction(uiDoc);
            transaction.Start("Inicio viga");
            foreach (ElementId id in sel.GetElementIds())
            {
                Element ele = uiDoc.GetElement(id);

                if (ele.Category.Id == Category.GetCategory(uiDoc, BuiltInCategory.OST_StructuralFraming).Id)
                    listaVigas.Add(ele);
                else if ((ele.Category.Id == Category.GetCategory(uiDoc, BuiltInCategory.OST_StructuralColumns).Id) ||
                           (ele.Category.Id == Category.GetCategory(uiDoc, BuiltInCategory.OST_Columns).Id))
                    listaPilar.Add(ele);
                else if (ele.Category.Id == Category.GetCategory(uiDoc, BuiltInCategory.OST_Floors).Id)
                    listaPiso.Add(ele);
            }
            transaction.Commit();
            uiDoc.Save();
            transaction.Start("Inicio laje viga");

            var listaPisoFormaLaje = new List<Floor>();

            foreach (Element item in listaPiso)
            {
                try
                {
                    foreach (Solid solid in Util.GetSolids(item))
                    {
                        if (solid != null)
                            if (solid.Faces.Size > 0)
                            {
                                Face face = Util.GetBottonFace(solid);
                                CurveArray curveArray = new CurveArray();
                                baseLevel = uiDoc.GetElement(Util.GetLevelMaisProximo(face, listaLevel)) as Level;
                                List<CurveLoop> curveLoops = face.GetEdgesAsCurveLoops().ToList();
                                foreach (CurveLoop curveLoops2 in curveLoops)
                                {
                                    foreach (Curve curve in curveLoops2)
                                    {
                                        curveArray.Append(curve);//.CreateTransformed(offset));
                                    }

#if D23 || D24
                                    Floor floor =  Autodesk.Revit.DB.Floor.Create(uiDoc, curveLoops, tipoPiso.Id, baseLevel.Id)  ;//uiDoc.Create.Newloor(curveArray, tipoPiso, baseLevel, false);
                                    
                                    Face faceBottom = Util.GetBottonFace(solid);
                                    Face faceTop = Util.GetTopFace(solid);
                                    double espessuraPiso = Math.Abs(Util.GetElevacaoRelacaoAoZeroFaceFeet(faceTop) - Util.GetElevacaoRelacaoAoZeroFaceFeet(faceBottom));
                                    floor.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM).Set(-espessuraPiso);
                                    floor.LookupParameter("Nivel4").Set(item.LookupParameter("Nivel4").AsString()) ;
#else
                                    //Floor floor = uiDoc.Create.Newloor(curveArray, tipoPiso, baseLevel, false);

#endif

                                 //   listaPisoFormaLaje.Add(floor);

                                }
                            }
                    }
                }
                catch
                {

                }
            }
            transaction.Commit();
            uiDoc.Save();
            transaction.Start("Inicio viga");
            var listaParedeVigaCriada = new List<Wall>();
            var listaParedeVigaCriadaFacesMenores = new List<Wall>();

            foreach (Element item in listaVigas)
            {

                foreach (Solid solid in Util.GetSolids(item))
                {
                    if (solid != null)
                        if (solid.Faces.Size > 0)
                        {
                            //criar fundo viga
                            Face faceBottom = Util.GetBottonFace(solid);
                            Face faceTop = Util.GetTopFace(solid);
                            CurveArray curveArray = new CurveArray();
                            baseLevel = uiDoc.GetElement(Util.GetLevelMaisProximo(faceBottom, listaLevel)) as Level;
                            List<CurveLoop> curveLoops = faceBottom.GetEdgesAsCurveLoops().ToList();
                            foreach (CurveLoop curveLoops2 in curveLoops)
                            {
                                foreach (Curve curve in curveLoops2)
                                {
                                    curveArray.Append(curve);//.CreateTransformed(offset));
                                }

#if D23 || D24
                                Floor floor = Autodesk.Revit.DB.Floor.Create(uiDoc, curveLoops, tipoPiso.Id, baseLevel.Id);//uiDoc.Create.Newloor(curveArray, tipoPiso, baseLevel, false);
                               
                                double espessuraPiso = Math.Abs(Util.GetElevacaoRelacaoAoZeroFaceFeet(faceTop) - Util.GetElevacaoRelacaoAoZeroFaceFeet(faceBottom));
                                floor.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM).Set(-espessuraPiso);
                                floor.LookupParameter("Nivel4").Set(item.LookupParameter("Nivel4").AsString());

#else
                               //     Floor floor = uiDoc.Create.Newloor(curveArray, tipoPiso, baseLevel, false);

#endif
                            }
                            baseLevelId = baseLevel.Id;
                            double espessura = wallType.Width;
                            List<Curve> perfilFundoViga = new List<Curve>();
                            foreach (CurveLoop curveLoops2 in curveLoops)
                            {
                                foreach (Curve curve in curveLoops2)
                                {
                                    perfilFundoViga.Add(curve);//.CreateTransformed(offset));                              
                                }
                            }
                            var listaMaior = perfilFundoViga.OrderByDescending(x => x.Length).ToList();
                            Curve item1 = null;
                            Wall wall = null;
                            try
                            {
                                item1 = listaMaior[0];
                                wall = Autodesk.Revit.DB.Wall.Create(uiDoc, item1, tipoWall, baseLevelId, 10, 0, false, false);
                                wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET).Set(-baseLevel.Elevation + Util.GetElevacaoRelacaoAoZeroFaceFeet(faceBottom));
                                wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).Set(Math.Abs(Util.GetElevacaoRelacaoAoZeroFaceFeet(faceTop) -
                                                                                                                                                       Util.GetElevacaoRelacaoAoZeroFaceFeet(faceBottom)));
                                wall.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(item.Id.IntegerValue.ToString());
                                wall.LookupParameter("Nivel4").Set(item.LookupParameter("Nivel4").AsString());
                                listaParedeVigaCriada.Add(wall);
                            }
                            catch
                            {

                            }
                            try
                            {
                                item1 = listaMaior[1];

                                wall = Autodesk.Revit.DB.Wall.Create(uiDoc, item1, tipoWall, baseLevelId, 10, 0, false, false);
                                wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET).Set(-baseLevel.Elevation + Util.GetElevacaoRelacaoAoZeroFaceFeet(faceBottom));
                                wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).Set(Math.Abs(Util.GetElevacaoRelacaoAoZeroFaceFeet(faceTop) -
                                                                                                                                                       Util.GetElevacaoRelacaoAoZeroFaceFeet(faceBottom)));
                                wall.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(item.Id.IntegerValue.ToString());
                                wall.LookupParameter("Nivel4").Set(item.LookupParameter("Nivel4").AsString());
                                listaParedeVigaCriada.Add(wall);
                            }
                            catch
                            {

                            }
                            try
                            {
                                item1 = listaMaior[2];
                                wall = Autodesk.Revit.DB.Wall.Create(uiDoc, item1, tipoWall, baseLevelId, 10, 0, false, false);
                                wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET).Set(-baseLevel.Elevation + Util.GetElevacaoRelacaoAoZeroFaceFeet(faceBottom));
                                wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).Set(Math.Abs(Util.GetElevacaoRelacaoAoZeroFaceFeet(faceTop) -
                                                                                                                                                       Util.GetElevacaoRelacaoAoZeroFaceFeet(faceBottom)));
                                wall.LookupParameter("Nivel4").Set(item.LookupParameter("Nivel4").AsString());
                                wall.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(item.Id.IntegerValue.ToString());
                                listaParedeVigaCriadaFacesMenores.Add(wall);
                            }
                            catch
                            {

                            }
                            try
                            {
                                item1 = listaMaior[3];
                                wall = Autodesk.Revit.DB.Wall.Create(uiDoc, item1, tipoWall, baseLevelId, 10, 0, false, false);
                                wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET).Set(-baseLevel.Elevation + Util.GetElevacaoRelacaoAoZeroFaceFeet(faceBottom));
                                wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).Set(Math.Abs(Util.GetElevacaoRelacaoAoZeroFaceFeet(faceTop) -
                                                                                                                                                       Util.GetElevacaoRelacaoAoZeroFaceFeet(faceBottom)));
                                wall.LookupParameter("Nivel4").Set(item.LookupParameter("Nivel4").AsString());
                                wall.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(item.Id.IntegerValue.ToString());
                                listaParedeVigaCriadaFacesMenores.Add(wall);
                            }
                            catch
                            {

                            }


                        }
                }

            }
            transaction.Commit();
            uiDoc.Save();
            transaction.Start("Inicio pilar");
            var listaFormaPilar = new List<Wall>();
            foreach (Element item in listaPilar)
            {
                try
                {
                    foreach (Solid solid in Util.GetSolids(item))
                    {
                        if (solid != null)
                            if (solid.Faces.Size > 0)
                            {
                                foreach (Face face in POB.Util.GetVerticalFace(solid))
                                {
                                    List<CurveLoop> curveLoops = face.GetEdgesAsCurveLoops().ToList();
                                    List<Curve> perfil = new List<Curve>();
                                    Face faceBotom = Util.GetBottonFace(solid);
                                    Face faceTop = Util.GetTopFace(solid);
                                    baseLevelId = Util.GetLevelMaisProximo(faceBotom, listaLevel);
                                    baseLevel = uiDoc.GetElement(baseLevelId) as Level;
                                    XYZ normal = face.ComputeNormal(new UV(0, 0));
                                    double espessura = wallType.Width;
                                    Transform offset = Transform.CreateTranslation(espessura / 2 * normal);

                                    foreach (CurveLoop curveLoops2 in curveLoops)
                                    {

                                        // Check if curve loop is counter-clockwise.
                                        foreach (Curve curve in curveLoops2)
                                        {
                                            perfil.Add(curve.CreateTransformed(offset));
                                        }
                                        try
                                        {
                                            Wall wall = Autodesk.Revit.DB.Wall.Create(uiDoc, perfil, tipoWall, baseLevelId, false, (face as PlanarFace).FaceNormal);
                                            wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET).Set(Math.Abs(baseLevel.Elevation - Util.GetElevacaoRelacaoAoZeroFaceFeet(faceBotom)));
                                            wall.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(item.Id.IntegerValue.ToString());
                                            //  wall.get_Parameter(BuiltInParameter.WALL_KEY_REF_PARAM).Set(3);
                                            //  parameter = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).Set(;
                                            //   parameter.Set(Util.GetElevacaoRelacaoAoZeroFaceFeet(faceBotom)- Util.GetElevacaoRelacaoAoZeroFaceFeet(faceTop));
                                            // wall.get_Parameter(BuiltInParameter.WALL_KEY_REF_PARAM).Set(3);
                                            wall.LookupParameter("Nivel4").Set(item.LookupParameter("Nivel4").AsString());
                                            listaFormaPilar.Add(wall);
                                        }
                                        catch
                                        {

                                        }
                                    }
                                }
                            }
                    }
                }
                catch
                {

                }
            }
            transaction.Commit();
            uiDoc.Save();
            transaction.Start("Inter");
          
            List<ElementId> excludedElements = new List<ElementId>();
            foreach (Wall item in listaFormaPilar)
            {
                excludedElements.Add(item.Id);
            }
            foreach (Wall item in listaParedeVigaCriada)
            {
                excludedElements.Add(item.Id);
            }
            foreach (Wall item in listaParedeVigaCriadaFacesMenores)
            {
                excludedElements.Add(item.Id);
            }
        
            foreach (Wall wall in listaFormaPilar)
            {
                foreach (Element item in Util.FindElementsInterfering(wall, excludedElements))
                {
                    if ((item.Category.Id == Category.GetCategory(uiDoc, BuiltInCategory.OST_Floors).Id) &
                        (item.Id.IntegerValue.ToString() != wall.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString()))

                    {
                        foreach (Solid solid in Util.GetSolids(item))
                        {
                            if (solid != null)
                                if (solid.Faces.Size > 0)
                                {
                                    Face faceBottom = Util.GetBottonFace(solid);
                                    Face faceTop = Util.GetTopFace(solid);
                                    double espessuraPiso = Math.Abs(Util.GetElevacaoRelacaoAoZeroFaceFeet(faceTop) - Util.GetElevacaoRelacaoAoZeroFaceFeet(faceBottom));
                                    double alturaAtual = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).AsDouble();
                                    alturaAtual = alturaAtual - espessuraPiso;
                                    wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).Set(alturaAtual);

                                }
                        }
                        continue;
                    }

                }
            }
            transaction.Commit();
            uiDoc.Save();
            transaction.Start("Inter1");
            foreach (Floor floor in listaPisoFormaLaje)
            {

            }
            foreach (Wall wall in listaParedeVigaCriada)
            {
                foreach (Element item in Util.FindElementsInterfering(wall, excludedElements))
                {
                    if ((item.Category.Id == Category.GetCategory(uiDoc, BuiltInCategory.OST_Floors).Id) &
                        (item.Id.IntegerValue.ToString() != wall.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString()))

                    {
                        try
                        {
                            foreach (Solid solid in Util.GetSolids(item))
                            {
                                if (solid != null)
                                    if (solid.Faces.Size > 0)
                                    {
                                        Face faceBottom = Util.GetBottonFace(solid);
                                        Face faceTop = Util.GetTopFace(solid);
                                        double espessuraPiso = Math.Abs(Util.GetElevacaoRelacaoAoZeroFaceFeet(faceTop) - Util.GetElevacaoRelacaoAoZeroFaceFeet(faceBottom));
                                        double alturaAtual = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).AsDouble();
                                        alturaAtual = alturaAtual - espessuraPiso;
                                        wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).Set(alturaAtual);

                                    }
                            }
                        }
                        catch
                        {

                        }
                        continue;
                    }

                }
            }
            transaction.Commit();
            uiDoc.Save();
            transaction.Start("Inter2");
            List<ElementId> listaDeletar = new List<ElementId>();

            foreach (Wall wall in listaParedeVigaCriadaFacesMenores)
            {
                foreach (Element item in Util.FindElementsInterfering(wall, excludedElements))
                {
                    try

                    {
                        if (((item.Category.Id == Category.GetCategory(uiDoc, BuiltInCategory.OST_StructuralFraming).Id) |
                             (item.Category.Id == Category.GetCategory(uiDoc, BuiltInCategory.OST_Columns).Id) |
                             (item.Category.Id == Category.GetCategory(uiDoc, BuiltInCategory.OST_StructuralColumns).Id))

                             &
                            (item.Id.IntegerValue.ToString() != wall.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString()))

                        {
                            listaDeletar.Add(wall.Id);
                            // continue;
                        }
                    }
                    catch
                    {

                    }
                }
            }
            uiDoc.Delete(listaDeletar);
            transaction.Commit();
            return Result.Succeeded;
        }
    }
}
            