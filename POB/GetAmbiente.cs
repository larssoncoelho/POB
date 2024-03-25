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

    public class GetAmbiente : IExternalCommand
    {

        FamilySymbol fs1;

       

        public Autodesk.Revit.UI.Result Execute(ExternalCommandData revit,
            ref string message, ElementSet elements)
        {
            UIApplication uiApp = revit.Application;
            Document uiDoc = uiApp.ActiveUIDocument.Document;
            Selection sel = uiApp.ActiveUIDocument.Selection;
                 List<ElementId> listaEle;

        CategorySet  lista = CriarMenu.ObterCategoriasDeInstalacoesRevit(uiDoc);

#if D23 || D24
            Util.GetParameter(uiDoc, lista, "tocAmbienteSistema",SpecTypeId.String.Text, true, true);
            Util.GetParameter(uiDoc, lista, "tocAmbienteSistemaId", SpecTypeId.String.Text, true, true);
            Util.GetParameter(uiDoc, lista, "tocAmbienteNivel", SpecTypeId.String.Text, true, true);
            Util.GetParameter(uiDoc, lista, "tocAmbienteNivelId", SpecTypeId.String.Text, true, true);
            Util.GetParameter(uiDoc, lista, "tocZona", SpecTypeId.String.Text, true, true);
            Util.GetParameter(uiDoc, lista, "tocZonaId", SpecTypeId.String.Text, true, true);
#else
                Util.GetParameter(uiDoc, lista, "tocAmbienteSistema", ParameterType.Text, true, true);
            Util.GetParameter(uiDoc, lista, "tocAmbienteSistemaId", ParameterType.Integer, true, true);
            Util.GetParameter(uiDoc, lista, "tocAmbienteNivel", ParameterType.Text, true, true);
            Util.GetParameter(uiDoc, lista, "tocAmbienteNivelId", ParameterType.Text, true, true);
            Util.GetParameter(uiDoc, lista, "tocZona", ParameterType.Text, true, true);
            Util.GetParameter(uiDoc, lista, "tocZonaId", ParameterType.Text, true, true);
#endif


            Transaction t = new Transaction(uiDoc);
            t.Start("Teste");
            listaEle = new List<ElementId>();
            foreach (ElementId eleId in sel.GetElementIds())
            {
                listaEle.Add(eleId);
                Util.BuscarAninhados(eleId, uiDoc, ref listaEle);
            }

            foreach (ElementId item in listaEle)
            {
               Util.ObterAmbiente(uiDoc, item);
                uiDoc.GetElement(item).LookupParameter("L7InsumoVinculado").Set(NovoExtrair.GetDescricao(uiDoc, uiDoc.GetElement(item)));

            }
            t.Commit();

            return Result.Succeeded;
        }

      
    }
}
