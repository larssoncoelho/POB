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
using Autodesk.Revit.DB.Plumbing;
using wf = System.Windows.Forms;
using Funcoes;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.ApplicationServices;
using POB.Updater;

namespace POB
{
    [TransactionAttribute(TransactionMode.Manual)]
    [RegenerationAttribute(RegenerationOption.Manual)]

    public class HerdarValoresParametros : IExternalCommand
    {
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData revit,
           ref string message, ElementSet elements)
        {
            UIApplication uiApp = revit.Application;
            Document uiDoc = uiApp.ActiveUIDocument.Document;
            Selection sel = uiApp.ActiveUIDocument.Selection;
            Transaction t = new Transaction(uiDoc);
            t.Start("Inicio");
            var campos = "";
            Perguntar.InputBox("Digite os campos", "Digite aqui", ref campos);
            foreach (ElementId eleId in sel.GetElementIds())
            {
                var ele = uiDoc.GetElement(eleId);
                if (ele is FamilyInstance)
                {
                    foreach (var itemId in  (ele as FamilyInstance).GetSubComponentIds())
                    {
                        var item = uiDoc.GetElement(itemId);
                        foreach (var campo in campos.Split(';'))
                        {
                            try
                            {
                                var parOrigem = ele.LookupParameter(campo);
                                var parDestino = item.LookupParameter(campo).Set(parOrigem.AsString());
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