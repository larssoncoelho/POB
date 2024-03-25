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

    [TransactionAttribute(TransactionMode.Manual)]
    [RegenerationAttribute(RegenerationOption.Manual)]

    public class DividirPiso : IExternalCommand
    {

        FamilySymbol fs1;
        List<XYZ> pontosInicias = new List<XYZ>();
        List<XYZ> pontosFinais = new List<XYZ>();
        List<Line> linhas = new List<Line>();
       IList<Curve> curveArray = new List<Curve>();

        public void CriarLinhas(XYZ pontoInicial)
        {
            int i = 0;
            double distanciaEmY = 1;
            double distanciaEmX = 1;

            double coordenadaX = pontoInicial.X;
            double coordenadaY = pontoInicial.Y;


            while (i < 300)
            {
                if (i == 0)
                {
                    XYZ pnti = new XYZ(-300, coordenadaY, 0);
                    XYZ pntf = new XYZ(300, coordenadaY, 0);
                    curveArray.Add(Line.CreateBound(pnti, pntf));
                    XYZ pntix = new XYZ(coordenadaX, -300, 0);
                    XYZ pntfx = new XYZ(coordenadaX, 300, 0);
                    curveArray.Add(Line.CreateBound(pntix, pntfx));

                }
                else
                {
                    coordenadaY = coordenadaY + distanciaEmY;
                    coordenadaX = coordenadaX + distanciaEmX;
                    XYZ pnti = new XYZ(-300, coordenadaY, 0);
                    XYZ pntf = new XYZ(300, coordenadaY, 0);
                    curveArray.Add(Line.CreateBound(pnti, pntf));
                    XYZ pntix = new XYZ(coordenadaX, -300, 0);
                    XYZ pntfx = new XYZ(coordenadaX, 300, 0);
                    curveArray.Add(Line.CreateBound(pntix, pntfx));

                }
                i = i + 1;

            }

        }
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData revit,
            ref string message, ElementSet elements)
        {
            UIApplication uiApp = revit.Application;
            Document uiDoc = uiApp.ActiveUIDocument.Document;
            Selection sel = uiApp.ActiveUIDocument.Selection;
            XYZ P = new XYZ(0, 0, 0);

            Transaction transaction1 = new Transaction(uiDoc, "CreateGenericModel1");
            transaction1.Start();
            foreach (ElementId id in sel.GetElementIds())
            {
                Element ele = uiDoc.GetElement(id);
                if (ele is Autodesk.Revit.DB.Part)
                {
                    XYZ ponto = sel.PickPoint("Selecione um ponto da paginação");

                    CriarLinhas(ponto);
                    List<ElementId> parts = new List<ElementId>();
                    parts.Add(id);
                    IList<ElementId> intersectionElementIds   = new List<ElementId>();
                   
             
                    Face f = Funcoes.Util.GetTopFace(Funcoes.Util.GetSolids(ele)[0]);
                    foreach (CurveLoop cl in f.GetEdgesAsCurveLoops())
                    {
                        CurveLoopIterator cli = cl.GetCurveLoopIterator();
                        while (cli.MoveNext())
                        {
                            Line l = (cli.Current as Line);
                            P = l.GetEndPoint(0);
                            continue;
                        }
                        continue;
                    }
                    
                    Autodesk.Revit.DB.XYZ normal = new XYZ(0, 0, 1);
                    //uiDoc.Application.Create.pl
                    //Plane geometryPlane = uiDoc.Application.Create.NewPlane(normal, P);
                    SketchPlane plane = SketchPlane.Create(uiDoc, ele.LevelId);

                    SketchPlane sketchPlane = Funcoes.Util.CreateSketchPlane(normal, Autodesk.Revit.DB.XYZ.Zero, uiApp);
                    PartUtils.DivideParts(uiDoc, parts,  intersectionElementIds, curveArray,
                                                                              sketchPlane.Id);
                }
            }
            uiDoc.ActiveView.PartsVisibility = PartsVisibility.ShowPartsOnly;
            transaction1.Commit();
            return Result.Succeeded;
        }
    }




}
