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
using wf = System.Windows.Forms;
using System.Reflection;

namespace POB
{

    [TransactionAttribute(TransactionMode.Manual)]
    [RegenerationAttribute(RegenerationOption.Manual)]

    public class ObterDiametro : IExternalCommand
    {

        FamilySymbol fs1;
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData revit,
            ref string message, ElementSet elements)
        {
            UIApplication uiApp = revit.Application;
            Document uiDoc = uiApp.ActiveUIDocument.Document;
            Selection sel = uiApp.ActiveUIDocument.Selection;
            Transaction t = new Transaction(uiDoc);
            t.Start("Teste");
            if (wf.MessageBox.Show("Deseja alterar os nomes?", "TocBIM", wf.MessageBoxButtons.YesNo) == wf.DialogResult.Yes)
            {
                foreach (ElementId item in sel.GetElementIds())
                {
                    Element ele = uiDoc.GetElement(item);
                    string texto = (ele as FamilyInstance).Symbol.Name.Split(' ')[0];
                    ele.LookupParameter("tocAmbiente").Set(texto);
                    Level level = uiDoc.GetElement((ele as FamilyInstance).LevelId) as Level;
                    if (level != null) ele.LookupParameter("Pavimento").Set(level.Name);
                    else
                    {
                        level = uiDoc.GetElement(new ElementId(ele.LookupParameter("NivelExtraido").AsInteger())) as Level;
                        if (level != null) ele.LookupParameter("Pavimento").Set(level.Name);

                    }
                }

                t.Commit();

                return Result.Succeeded;
            }
            foreach (ElementId item in sel.GetElementIds())
            {
                Element ele = uiDoc.GetElement(item);



                /*  StringBuilder sb = new StringBuilder();
                  List<PlanarFace> faces = new List<PlanarFace>();
                  foreach (Solid solid in Util.GetSolids(ele))
                  {
                      if (solid != null)
                      {
                          foreach (var item1 in Util.GetVerticalFaceLista(solid))
                          {
                              faces.Add(item1 as PlanarFace);
                          }
                      }
                  }
                      sb.Append("Áea"+ "\t" +"Material id"+"\t"+ "Nome");

                  foreach (var f in faces)
                  {
                      string t1 = "";
                      t1 = f.Area.ToString();
                      try
                      {
                          t1 = t1 + "\t" + f.MaterialElementId.IntegerValue.ToString();
                      }
                   catch
                      {
                          t1 = t1 + "\t" + "erro material";
                      }

                      try
                      {
                          t1 = t1 + "\t" + (uiDoc.GetElement( f.MaterialElementId)  as Material).Name;
                      }
                      catch
                      {
                          t1 = t1 + "\t" + "erro material";
                      }
                      sb.AppendLine(t1);
                  }


                  wf.Clipboard.Clear();
                  wf.Clipboard.SetText(sb.ToString());*/
             /*   ImportInstance dwg = uiDoc.GetElement(item) as ImportInstance;

                if (dwg == null)
                    return Result.Failed;
                string texto = "";
                */
              /*  foreach (GeometryObject geometryObj in dwg.get_Geometry(new Options()))
                {
                    if (geometryObj is GeometryInstance) // This will be the whole thing
                    {
                        GeometryInstance dwgInstance = geometryObj as GeometryInstance;

                        foreach (GeometryObject blockObject in dwgInstance.SymbolGeometry)
                        {
                            if (blockObject is GeometryInstance) // This could be a block
                            {
                                //get the object name and coordinates and rotation and 
                                //load into my own class
                               // clsBlockInstance blockCls = new clsBlockInstance();

                                GeometryInstance blockInstance = blockObject as GeometryInstance;

                                string name = blockInstance.Symbol.Name;

                                Transform transform = blockInstance.Transform;

                                XYZ origin = transform.Origin;

                                XYZ vectorTran = transform.OfVector(transform.BasisX.Normalize());
                                double rot = transform.BasisX.AngleOnPlaneTo(vectorTran, transform.BasisZ.Normalize()); // radians
                                rot = rot * (180 / Math.PI); // degrees

                                foreach(var p in blockInstance.Symbol.Parameters)
                                { 
                                   
                                }
                                  

                              //  blockCls.Name = name;
                               // blockCls.Origin = origin;
                               // blockCls.Rotation = rot;

                                //blockInstanceCollection.Add(blockCls);
                            }
                        }
                    }
                }
                */
                foreach (Solid solid in Util.GetSolids(ele))
                { 
                    if (solid != null)
                        if (solid.Faces.Size > 0)
                        {

                            Face faceTop = Util.GetTopFace(solid);
                            if (faceTop != null)
                            {
                                try
                                {
#if D23 || D24
                                    var p1 = SpecTypeId.Length;
                                    double diametro = Util.GetDiametro(faceTop);
                                    var p = POB.Util.GetParameter(ele, "tocDiametro", p1, true, false);
                                    p.Set(diametro);
#else
                // var p1 =ParameterType.Length;
#endif
                                    /*  double diametro = Util.GetDiametro(faceTop);
                                      var p = POB.Util.GetParameter(ele, "tocDiametro", p1, true, false);
                                      p.Set(diametro);*/
                                }
                                catch
                                {

                                }
                            }
                        }
                }
            }

            t.Commit();


            return Result.Succeeded;
        }
    }
}
