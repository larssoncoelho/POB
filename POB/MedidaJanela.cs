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

    public class MedidaJanela : IExternalCommand
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
                    var symbol = (ele as FamilyInstance).Symbol;
                    double b = Convert.ToDouble(symbol.LookupParameter("Tamanho Nominal L x A").AsString().Split('×')[0].Trim()) / 0.3048;
                    double a = Convert.ToDouble(symbol.LookupParameter("Tamanho Nominal L x A").AsString().Split('×')[1].Trim()) / 0.3048;

                    ele.LookupParameter("Área1").Set(b * a);
                    ele.LookupParameter("Comprimento1").Set(b);
                    ele.LookupParameter("Comp verga").Set(b + 0.30 / 0.3048 * 2);
                    ele.LookupParameter("Comp contra verga").Set(b + 0.30 / 0.3048 * 2);
                    ele.LookupParameter("Comp contramarco").Set(2*b+ 2*a);

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
