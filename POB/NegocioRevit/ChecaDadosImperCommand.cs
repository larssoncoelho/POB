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


namespace POB.NegocioRevit
{
    public static class ChecaDadosImperCommand
    {
        public static double TocEspessuraTotal { get; private set; }


        public static ResultadoExternalCommandData Execute(ExternalCommandData revit)
        {
            UIApplication uiApp = revit.Application;
            Document uiDoc = uiApp.ActiveUIDocument.Document;
            Selection sel = uiApp.ActiveUIDocument.Selection;
            XYZ P = new XYZ(0, 0, 0);
            List<ElementoResumido> listaElemento = new List<ElementoResumido>();

            FilteredElementCollector collector = new FilteredElementCollector(uiDoc);

            Transaction t = new Transaction(uiDoc);
            t.Start("t");
            foreach (ElementId id in sel.GetElementIds())
            {
                Element ele = uiDoc.GetElement(id);
                if (!(ele is Autodesk.Revit.DB.AssemblyInstance)) continue;
                Autodesk.Revit.DB.AssemblyInstance assemblyInstance = ele as AssemblyInstance;
                listaElemento.Clear();
                var s = "";
                var tocCodImper = ele.LookupParameter("tocCodigoImper").AsString().ToUpper();
                foreach (var item in assemblyInstance.GetMemberIds())
                {
                    try
                    {
                        string nomedaFamilia = uiDoc.GetElement(item).Name;

                        if (nomedaFamilia.ToUpper().Contains(tocCodImper))
                            uiDoc.GetElement(item).LookupParameter("Comentários").Set(tocCodImper + "|" + ele.Id.IntegerValue + "|" + nomedaFamilia);
                        else uiDoc.GetElement(item).LookupParameter("Comentários").Set(tocCodImper + "|" + ele.Id.IntegerValue + " | Erro: " + nomedaFamilia);
                    }
                    catch
                    {

                    }
                }
            }
            t.Commit();
            return new ResultadoExternalCommandData { Resultado = Result.Succeeded };
        }
    }
}