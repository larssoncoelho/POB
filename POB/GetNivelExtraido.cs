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


namespace POB
{



    [TransactionAttribute(TransactionMode.Manual)]
    [RegenerationAttribute(RegenerationOption.Manual)]

    public class GetNivelExtraido : IExternalCommand
    {

        FamilySymbol fs1;
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData revit,
            ref string message, ElementSet elements)
        {
            UIApplication uiApp = revit.Application;
            Document uiDoc = uiApp.ActiveUIDocument.Document;
            Selection sel = uiApp.ActiveUIDocument.Selection;
            TransactionGroup transactionGroup = new TransactionGroup(uiDoc);
            transactionGroup.Start("Teste");
            var listaLevel = Util.ListaDeNiveis(uiDoc, true).OrderBy(x=>x.Elevation).ToList();

            NegocioRevit.NivelExtraidoCommad.Execute(uiDoc, sel.GetElementIds(), true, listaLevel);

            transactionGroup.Commit();


            return Result.Succeeded;
        }
    }




}
