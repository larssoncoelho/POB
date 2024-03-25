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

    public class ElementoResumido
    {
        public string Tipo { get; set; }
        public double? TocAlturaRodape { get; set; }
        public double? TocEspessuraMaxima { get; set; }
        public double? TocAreaRodape { get; set; }
        public double? TocPerimetro { get; set; }
        public double? TocAreaPiso { get; internal set; }
        public double? TocVolume { get; internal set; }
        public double? TocEspessuraRealMaxima { get; set; }
        public double? TocEspessuraRealMinima { get; internal set; }
        public double? TocEspessuraAcabamento { get; internal set; }
        public double? TocEspessuraTotal { get; internal set; }
        public double? TocLarguraParede { get; internal set; }
        public double? TocComprimentoJunta { get; internal set; }
        public double TocDescontarEspessuraRegularizacao { get; internal set; }
    }
    /*public class SistemaImper
    {
        public string NomeMaterial { get; set; }
        public int ordem { get; set; }
    }*/
    [TransactionAttribute(TransactionMode.Manual)]
    [RegenerationAttribute(RegenerationOption.Manual)]

    public class DadosImper : IExternalCommand
    {
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData revit,
            ref string message, ElementSet elements)
        {

            UIApplication uiApp = revit.Application;
            Document uiDoc = uiApp.ActiveUIDocument.Document;
            var vista = uiDoc.ActiveView;
            var par = vista.LookupParameter("Modelo de vista");

            Apresentacao.FrmDadosImper frmDadosImper = new Apresentacao.FrmDadosImper(revit, true);
            frmDadosImper.ShowDialog();
            if (!frmDadosImper.Continuar) return Result.Cancelled;
            else return Result.Succeeded;
        }
     
     
    }
}

