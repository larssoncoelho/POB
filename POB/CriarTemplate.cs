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
using POB.ObjetoDeTranferencia;
using System.IO.Packaging;
using ex =Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Windows.Media.Media3D;


namespace POB
{

    [TransactionAttribute(TransactionMode.Manual)]
    [RegenerationAttribute(RegenerationOption.Manual)]

    public class CriarTemplate : IExternalCommand
    {

        public Autodesk.Revit.UI.Result Execute(ExternalCommandData revit, ref string message, ElementSet elements)
        {
            try
            {
                UIApplication uiApp = revit.Application;
                Document uiDoc = uiApp.ActiveUIDocument.Document;
                List<DadosExcel> materiais = new List<DadosExcel>();
                List<DadosExcel> parametros = new List<DadosExcel>();
                List<DadosExcel> dados = new List<DadosExcel>();
                CategorySet categorySet = new CategorySet();

                categorySet.Insert(Category.GetCategory(uiDoc, BuiltInCategory.OST_Materials));


                var caminho = @"d:\Parêmtros2.xlsx";
                ex.Application excelApp = new ex.Application();
                ex.Workbook workbook = excelApp.Workbooks.Open(caminho);
                ex.Worksheet dadosExcel = workbook.Sheets["dados"];
                ex.Worksheet materiaisExcel = workbook.Sheets["material"];
                ex.Worksheet worksheet = workbook.Sheets["parametros"];
                ex.Range rangeDados = dadosExcel.UsedRange;
                ex.Range rangeMateriais= materiaisExcel.UsedRange;
                ex.Range rangeParametros = worksheet.UsedRange;
                excelApp.Visible = true;

                for (int row = 2; row <= rangeParametros.Rows.Count; row++) // Assuming first row is header
                {
                    DadosExcel dado = new DadosExcel
                    {
                        Lingua = (rangeParametros.Cells[row, 1] as ex.Range).Text,
                        Parametro = (rangeParametros.Cells[row, 2] as ex.Range).Text,
                        Unidade = (rangeParametros.Cells[row, 3] as ex.Range).Text
                    };
                    parametros.Add(dado);
                }
                for (int row = 2; row <= rangeMateriais.Rows.Count; row++) // Assuming first row is header
                {
                    DadosExcel dado = new DadosExcel
                    {
                        Material = (materiaisExcel.Cells[row, 1] as ex.Range).Text,
                        R = Convert.ToByte((materiaisExcel.Cells[row, 2] as ex.Range).Text),
                        G = Convert.ToByte((materiaisExcel.Cells[row, 3] as ex.Range).Text),
                        B = Convert.ToByte((materiaisExcel.Cells[row, 4] as ex.Range).Text),

                    };
                    materiais.Add(dado);
                }
                for (int row = 2; row <= rangeDados.Rows.Count; row++) // Assuming first row is header
                {
                    DadosExcel dado = new DadosExcel
                    {
                        Lingua = (rangeDados.Cells[row, 1] as Range).Text,
                        Material = (rangeDados.Cells[row, 2] as Range).Text,
                        Parametro = (rangeDados.Cells[row, 3] as Range).Text,
                        ValorCatalogo = (rangeDados.Cells[row, 4] as Range).Text,
                        ValorConvertido = (rangeDados.Cells[row, 5] as Range).Text,
                        UnidRevit = (rangeDados.Cells[row, 6] as Range).Text,
                        ValorRevit = (rangeDados.Cells[row, 7] as Range).Text,

                    };
                    dados.Add(dado);
                }


                workbook.Close(false);
                excelApp.Quit();

                // Clean up
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

                /*
                var excel = new ExcelQueryFactory(caminho) { ReadOnly = true };
                try
                {
                    parametros = (from linha in excel.Worksheet("parametros")
                                  select new DadosExcel
                                  {
                                      Lingua = linha["Lingua"] != null ? linha["Lingua"].ToString() : null,
                                      Parametro = linha["Parametro"] != null ? linha["Parametro"].ToString() : null,
                                      Unidade = linha["Unidade"] != null ? linha["Unidade"].ToString() : null
                                  }).ToList();
                }
                catch (Exception ex) { 
                }
                materiais = (from linha in excel.Worksheet("materiais")
                             select new DadosExcel
                             {
                                 Material = linha["Material"].ToString(),
                                 R = Convert.ToByte(linha["r"].ToString()),
                                 G = Convert.ToByte(linha["g"].ToString()),
                                 B = Convert.ToByte(linha["b"].ToString()),
                             }).ToList();
                dados = (from linha in excel.Worksheet("dados")
                         select new DadosExcel
                         {
                             Lingua = linha["Lingua"].ToString(),
                             Material = linha["Material"].ToString(),
                             Parametro = linha["Parametro"].ToString(),
                             ValorCatalogo = linha["ValorCatalogo"].ToString(),
                             ValorConvertido = linha["ValorConvertido"].ToString(),
                             UnidRevit = linha["UnidRevit"].ToString(),
                             ValorRevit = linha["ValorRevit"].ToString(),
                         }).ToList();
*/
                Transaction t = new Transaction(uiDoc);
                t.Start("Teste");

                POB.Util.uiDoc = uiDoc;
                foreach (DadosExcel d in materiais)
                {
                    POB.Util.CriarMaterial(uiDoc, d.Material, d.R, d.G, d.B);
                }

                foreach (DadosExcel d in parametros)
                {
                    Autodesk.Revit.DB.ParameterType tipo = ObterTipo(d.Unidade);
                    Util.GetParameter(uiDoc, categorySet, d.Parametro, tipo, true, false);
                    //POB.Util.CriarMaterial(uiDoc, d.Material, d.R, d.G, d.B);
                }
                t.Commit();
                t.Start("Teste");
                foreach (DadosExcel d in dados)
                {
                    Autodesk.Revit.DB.Material material = Util.FindElementByName(typeof(Autodesk.Revit.DB.Material), d.Material) as Autodesk.Revit.DB.Material;
                    var p = material.LookupParameter(d.Parametro);
                    switch (p.Definition.ParameterType)
                    {
                        case ParameterType.Text: case ParameterType.URL:
                            p.Set(d.ValorRevit);
                            break;
                        default:
                              p.SetValueString(d.ValorRevit);
                            break;
                    }
                    
                    material.LookupParameter("Fabricante").Set("Revestech");
                    //material.LookupParameter("URL").Set("WWW.");
                }
               
                t.Commit();
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                return Result.Cancelled;

            }
        }
        public Autodesk.Revit.DB.ParameterType ObterTipo (string nome)
        {
            switch (nome)
            {
                case "m":
                    return Autodesk.Revit.DB.ParameterType.Length;
                case "g/m2":
                    return Autodesk.Revit.DB.ParameterType.MassPerUnitArea;
                case "mm":
                    return Autodesk.Revit.DB.ParameterType.DisplacementDeflection;
                case "Texto":
                    return Autodesk.Revit.DB.ParameterType.Text;
                case "%":
                    return Autodesk.Revit.DB.ParameterType.Slope;
                case "N/50 mm":
                    return Autodesk.Revit.DB.ParameterType.ForcePerLength;
                case "Kg":
                    return Autodesk.Revit.DB.ParameterType.Mass;
                case "°C":
                    return Autodesk.Revit.DB.ParameterType.HVACTemperature;
                default:
                    return  Autodesk.Revit.DB.ParameterType.Text;
                    break;
            }

        }
    }
}
