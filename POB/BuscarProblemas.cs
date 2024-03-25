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
using wd = System.Drawing;

namespace POB
{

    [TransactionAttribute(TransactionMode.Manual)]
    [RegenerationAttribute(RegenerationOption.Manual)]

    public class BuscarProblemas : IExternalCommand
    {

        FamilySymbol fs1;



        public Autodesk.Revit.UI.Result Execute(ExternalCommandData revit,
            ref string message, ElementSet elements)
        {
            UIApplication uiApp = revit.Application;
            Document uiDoc = uiApp.ActiveUIDocument.Document;
            Selection sel = uiApp.ActiveUIDocument.Selection;
            List<ElementId> listaEle;


            CategorySet lista = CriarMenu.ObterCategoriasDeInstalacoesRevit(uiDoc);

#if D23 || D24
            Util.GetParameter(uiDoc, lista, "tocAmbienteSistema", SpecTypeId.String.Text, true, true);
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
            List<Componentes> compoenetes = new List<Componentes>();
            Color c = new Color(wd.Color.LightGray.R, wd.Color.LightGray.G, wd.Color.LightGray.B);
            OverrideGraphicSettings org = new OverrideGraphicSettings();
            org.SetProjectionLineColor(c);
           
            org.SetCutLineColor(c);
#if DEBUG20201
#else
            /* org.SetProjectionFillColor(c);
            org.SetCutFillColor(c);
            org.SetProjectionFillPatternId(Util.GetSolidFill(uiDoc));
            org.SetProjectionFillPatternVisible(true);
            org.SetHalftone(false);
            org.SetSurfaceTransparency(0);*/
#endif
            foreach (ElementId item in listaEle.Distinct().ToList())
            {
                uiDoc.ActiveView.SetElementOverrides(item, org);
                var ele = uiDoc.GetElement(item);
                compoenetes.Add(new Componentes
                {
                    tocAmbienteSistema = ele.LookupParameter("tocAmbienteSistema").AsString(),
                    tocAmbienteNivel = ele.LookupParameter("tocAmbienteNivel").AsString(),
                    tocZona = ele.LookupParameter("tocZona").AsString(),
                    L7InsumoVinculado = ele.LookupParameter(Properties.Settings.Default.L7InsumoVinculado).AsString(),
                    Contador = 1
                });

            }

            var lista01 = from Componentes c1 in compoenetes
                          where c1.tocZona == "Final 01"
                          select c1;

            var lista02 = from Componentes c1 in compoenetes
                          where c1.tocZona == "Final 02"
                          select c1;
            var listaItens = (from Componentes c5 in compoenetes
                             select new
                             {
                                 L7InsumoVinculado = c5.L7InsumoVinculado,
                                 tocAmbienteSistema = c5.tocAmbienteSistema
                             }).Distinct().ToList();

            var listaResumida = (from c3 in listaItens
                                select new
                                {
                                    L7InsumoVinculado = c3.L7InsumoVinculado,
                                    tocAmbienteSistema = c3.tocAmbienteSistema,
                                    qtde1 = lista01.Where(x=>(x.tocAmbienteSistema==c3.tocAmbienteSistema) & 
                                                (x.L7InsumoVinculado== c3.L7InsumoVinculado)).Count(),
                                    qtde2 = lista02.Where(x => (x.tocAmbienteSistema == c3.tocAmbienteSistema) &
                                                (x.L7InsumoVinculado == c3.L7InsumoVinculado)).Count()                
    
                                }).ToList();

            var listaParaProcurar = (from c4 in listaResumida
                                     select new
                                     {
                                         L7InsumoVinculado = c4.L7InsumoVinculado,
                                         tocAmbienteSistema = c4.tocAmbienteSistema,
                                         qtde1 = c4.qtde1,
                                         qtde2 = c4.qtde2,
                                         diferenca = c4.qtde1 - c4.qtde2
                                     }).Where(x => x.diferenca != 0).ToList();
                                   
                                 
            foreach (var item in listaParaProcurar)
            {
                var lista1 = Util.GetFilterElementByParameterId(uiDoc, "L7InsumoVinculado", item.L7InsumoVinculado, "tocAmbienteSistema", item.tocAmbienteSistema);

                Random randomGen = new Random();
                wd.KnownColor[] names = (wd.KnownColor[])Enum.GetValues(typeof(wd.KnownColor));
                wd.KnownColor randomColorName = names[randomGen.Next(names.Length)];
                System.Drawing.Color randomColor = System.Drawing.Color.FromKnownColor(randomColorName);
                c = new Color(randomColor.R, randomColor.G, randomColor.B);

                org.SetProjectionLineColor(c);
  
                org.SetCutLineColor(c);

#if DEBUG20201
#else
                      //        org.SetProjectionFillColor(c);            
                //org.SetCutFillColor(c);
#endif
                foreach (var item2 in lista1)             
                {
                    uiDoc.ActiveView.SetElementOverrides(item2, org);
                }
            }

              t.Commit();

            return Result.Succeeded;
        }


    }
}
