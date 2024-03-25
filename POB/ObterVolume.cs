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
    public static class FuncoesPOB
    {
        public static double GetVolumeSolid(Element ele)
        {

          double  v = 0;
            Options opt = ele.Document.Application.Create.NewGeometryOptions();
            GeometryElement ge12 = ele.get_Geometry(opt);

            foreach (GeometryObject obj in ge12)
            {
                //Util.InfoMsg("Primeira avaliação: \n" +obj.ToString());
                if (obj is Solid)
                {
                    v = v + (obj as Solid).Volume; // * 0.3048 * 0.3048 * 0.3048;
                    //Util.InfoMsg("Se já for um solid: \n"+(obj as Solid).ToString()+"\n"+ ((obj as Solid).Volume*0.3048 * 0.3048 * 0.3048).ToString());
                }


            }
            return v;
        }

            public static double GetVolumeInterno(Element ele)
        {
            double vol = 0;
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
                            vol = vol + solid.Volume;// * 0.3048 * 0.3048 * 0.3048;
                        }
                    }
                }
            }
            catch
            {

                return  GetVolumeSolid(ele);
            }
            return vol;
        }
    }


    [TransactionAttribute(TransactionMode.Manual)]
    [RegenerationAttribute(RegenerationOption.Manual)]

    public class ObterVolume : IExternalCommand
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
            foreach (ElementId item in sel.GetElementIds())
            {
                Element ele = uiDoc.GetElement(item);
                try
                {
                    double volume = FuncoesPOB.GetVolumeInterno(ele);
                    ele.LookupParameter("tocVolume").Set(volume);
                }
                catch
                {

                }
            }
            t.Commit();

        
            return Result.Succeeded;
        }
    }




}
