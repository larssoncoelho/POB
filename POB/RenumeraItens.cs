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

    public class RenumeraItens : IExternalCommand
    {
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData revit,
            ref string message, ElementSet elements)
        {
            var uiDoc = revit.Application.ActiveUIDocument.Document;
            Transaction t = new Autodesk.Revit.DB.Transaction(uiDoc, "Teste");
            //ggg.Funcoes1.Exportar(uiDoc, this.ActiveUIDocument.Selection);
            t.Start("cddf");

            //ggg.Funcoes1.LerDadosExcel(uiDoc,"Comentários");

            var sel = revit.Application.ActiveUIDocument.Selection;


             int i = 0;
            //abrir excel
            //var listaCatMetro = Funcoes1.GetCategoriaItensSistemaPorMetro(uiDoc);
            foreach (ElementId eleId in sel.GetElementIds())
            {
                var ele = uiDoc.GetElement(eleId);                         

                (ele as Autodesk.Revit.DB.AssemblyInstance).AssemblyTypeName = "ASSEMBLY" + i.ToString().PadLeft(4, '0');
            }
            t.Commit();
              return Result.Succeeded;
        }
    }
}
