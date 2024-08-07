﻿using System;
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

    public class ObterMaiorFace : IExternalCommand
    {

        FamilySymbol fs1;
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData revit,
            ref string message, ElementSet elements)
        {
            UIApplication uiApp = revit.Application;
            Document uiDoc = uiApp.ActiveUIDocument.Document;
            Selection sel = uiApp.ActiveUIDocument.Selection;
            TransactionGroup t = new TransactionGroup(uiDoc);
            t.Start("Teste");
            foreach (ElementId item in sel.GetElementIds())
            {
                Element ele = uiDoc.GetElement(item);
                double area = 0;
                try
                {
#if D23 || D24
                    var p1 = SpecTypeId.Area;
#else
                 var p1 =ParameterType.Area;
#endif
                    var par = Util.GetParameter(ele, "Área de projeção", p1, true, false);
                    foreach (Solid solid in Util.GetSolids(ele))
                    {
                        if (solid != null)
                        {
                           var facesVerticais = Util.GetFaceLista(solid);
                            Autodesk.Revit.DB.Face maiorFace = null;
                            if (facesVerticais.Count>0)
                            {
                            try
                            {
                                maiorFace = facesVerticais.OrderByDescending(x => x.Area).First();

                                area = maiorFace.Area;
                            }
                            catch { 
                            }
                            }
                        }
                    }
                    Transaction t1 = new Transaction(ele.Document);
                    t1.Start("ttt");
                    par.Set(area);
                    t1.Commit();
                    t1.Dispose();
                }
                catch (Exception e10)
                {
                    TaskDialog.Show("_", e10.Message);
                }
            }
            t.Commit();


            return Result.Succeeded;
        }
    }




}
