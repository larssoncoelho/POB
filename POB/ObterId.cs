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

namespace POB
{

    [TransactionAttribute(TransactionMode.Manual)]
    [RegenerationAttribute(RegenerationOption.Manual)]

    public class ObterId : IExternalCommand
    {
        private ElementId baseLevelId;

        public Autodesk.Revit.UI.Result Execute(ExternalCommandData revit, ref string message, ElementSet elements)
        {
            UIApplication uiApp = revit.Application;
            Document uiDoc = uiApp.ActiveUIDocument.Document;
            Selection sel = uiApp.ActiveUIDocument.Selection;
            var itens =  sel.PickObjects(ObjectType.LinkedElement);
            string texto = "";
            foreach (var r in itens)
            {
                var link = (uiDoc.GetElement(r.ElementId) as RevitLinkInstance).Name.Split(':')[0];
                texto = texto + '\n' + link + "\t" + r.LinkedElementId.IntegerValue.ToString();
            }
            wf.Clipboard.SetText(texto);

            return Result.Succeeded;
        }
    }
}
