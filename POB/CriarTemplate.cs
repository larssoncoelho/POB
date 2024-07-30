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
using System.Windows.Controls.Primitives;
using LinqToExcel;


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
                List<Tipo> tipos = new List<Tipo>();


                List<Camada> camadas = new List<Camada>();
                CategorySet categorySet = new CategorySet();

                categorySet.Insert(Category.GetCategory(uiDoc, BuiltInCategory.OST_Materials));


                var caminho = @"d:\Parêmtros2.xlsx";
                ex.Application excelApp = new ex.Application();
                ex.Workbook workbook = excelApp.Workbooks.Open(caminho);
                ex.Worksheet dadosExcel = workbook.Sheets["dados"];
                ex.Worksheet materiaisExcel = workbook.Sheets["material"];
                ex.Worksheet worksheet = workbook.Sheets["parametros"];
                ex.Worksheet tiposExcel = workbook.Sheets["tipos"];
                ex.Worksheet camadasExcel = workbook.Sheets["camadas"];


                ex.Range rangeDados = dadosExcel.UsedRange;
                ex.Range rangeMateriais= materiaisExcel.UsedRange;
                ex.Range rangeParametros = worksheet.UsedRange;
                ex.Range rangeTipos = tiposExcel.UsedRange;
                ex.Range rangeCamadas = camadasExcel.UsedRange;
                excelApp.Visible = true;

                for (int row = 2; row <= rangeParametros.Rows.Count; row++) // Assuming first row is header
                {
                    DadosExcel dado = new DadosExcel
                    {
                        Lingua = (rangeParametros.Cells[row, 1] as ex.Range).Text,
                        Parametro = (rangeParametros.Cells[row, 2] as ex.Range).Text,
                        Unidade = (rangeParametros.Cells[row, 3] as ex.Range).Text,
                        ParametroIngles = (rangeParametros.Cells[row, 7] as ex.Range).Text,
                        ParametroEspanhol = (rangeParametros.Cells[row, 8] as ex.Range).Text,

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

                for (int row = 2; row <= rangeTipos.Rows.Count; row++) // Assuming first row is header
                {
                    Tipo dado = new Tipo
                    {
                        id = Convert.ToInt32( (rangeTipos.Cells[row, 1] as Range).Text),
                        Categoria = (rangeTipos.Cells[row, 2] as Range).Text,
                        Familia= (rangeTipos.Cells[row, 3] as Range).Text,

                        DescricaoTipo = (rangeTipos.Cells[row, 4] as Range).Text,
                        ComentariosDeTipo = (rangeTipos.Cells[row, 5] as Range).Text,
                     

                    };
                    tipos.Add(dado);
                }

                for (int row = 2; row <= rangeCamadas.Rows.Count; row++) // Assuming first row is header
                {
                    Camada camada = new Camada
                    {
                        idFamila = Convert.ToInt32( (rangeCamadas.Cells[row, 1] as Range).Text),
                        Familia = (rangeCamadas.Cells[row, 2] as Range).Text,
                        DescricaoTipo = (rangeCamadas.Cells[row, 3] as Range).Text,
                        Material = (rangeCamadas.Cells[row, 4] as Range).Text,
                        Espessura = Convert.ToDouble( (rangeCamadas.Cells[row, 5] as Range).Text),
                        TipoDeCamada = (rangeCamadas.Cells[row, 6] as Range).Text,
                        Variavel = Convert.ToInt32((rangeCamadas.Cells[row, 7] as Range).Text),
                    };
                    camadas.Add(camada);
                }



                workbook.Close(false);
                excelApp.Quit();

                // Clean up
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);


                /* var excel = new ExcelQueryFactory(caminho) { ReadOnly = true };
                 try
                 {
                     camadas = (from linha in excel.Worksheet("camadas")
                                   select new Camada
                                   {
                                      DescricaoTipo  = linha["Lingua"] .ToString(),
                                       Espessura = Convert.ToDouble( linha["Lingua"].ToString()),
                                       Familia = linha["Familia"].ToString(),
                                      idFamila = Convert.ToInt32( linha["idFamila"].ToString()),
                                      Material= linha["Material"].ToString(),
                                      TipoDeCamada = linha["TipoDeCamada"].ToString(),
                                      Variavel= linha["Variavel"].ToString()


                                   }).ToList();
                 }
                 catch (Exception ex) { 
                 }*/
                Transaction tCriarParametros = new Transaction(uiDoc);
                tCriarParametros.Start("Teste");
                POB.Util.uiDoc = uiDoc;
                foreach (DadosExcel d in parametros)
                {
                    Autodesk.Revit.DB.ParameterType tipo = ObterTipo(d.Unidade);
                    Util.GetParameter(uiDoc, categorySet, d.ParametroEspanhol, tipo, true, false);
                    Util.GetParameter(uiDoc, categorySet, d.ParametroIngles, tipo, true, false);
                }
                tCriarParametros.Commit();
                Transaction tCopiarDados = new Transaction(uiDoc);
                tCopiarDados.Start("Teste");

                
                var materiaisModelo = new Autodesk.Revit.DB.FilteredElementCollector(uiDoc).OfClass(typeof(Autodesk.Revit.DB.Material)).Cast<Autodesk.Revit.DB.Material>();

                foreach (var material in materiaisModelo)
                {

                    foreach (DadosExcel d in parametros)
                    {
                        var pOriginal = material.LookupParameter(d.Parametro);
                        var pTraduzidoIngles = material.LookupParameter(d.ParametroIngles);
                        var pTraduzidoEspanhol = material.LookupParameter(d.ParametroEspanhol);



                        switch (pOriginal.Definition.ParameterType)
                        {
                            case ParameterType.Text:
                            case ParameterType.URL:
                                pTraduzidoIngles.Set(pOriginal.AsString());
                                pTraduzidoEspanhol.Set(pOriginal.AsString());

                                break;
                            default:
                                pTraduzidoIngles.SetValueString(pOriginal.AsValueString());
                                pTraduzidoEspanhol.SetValueString(pOriginal.AsValueString());
                                break;
                        }
                    }
                }
                tCopiarDados.Commit();
                return Result.Succeeded;

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
                
                var tipoPisoBase = new Autodesk.Revit.DB.FilteredElementCollector(uiDoc).OfClass(typeof(FloorType)).Cast<FloorType>().ToList()[0];
                var tipoParedeBase = new Autodesk.Revit.DB.FilteredElementCollector(uiDoc).OfClass(typeof(WallType)).Cast<WallType>().ToList()[0];
                foreach (Tipo tipo in tipos)
                {
                    var cmd = camadas.Where(x=>x.idFamila==tipo.id).ToList();
                    if(cmd!=null)
                        if(cmd.Count()>0)
                        {
                            t.Start("sadfwe");

                            switch (tipo.Categoria)
                            {
                                case "Pisos":
                                    var v = POB.Util.DuplicarFloorType(tipoPisoBase, tipo.DescricaoTipo, cmd);
                                    v.get_Parameter(Autodesk.Revit.DB.BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).Set(tipo.ComentariosDeTipo);
                                    break;
                                case "Paredes":
                                    var p = POB.Util.DuplicarWallType(tipoParedeBase, tipo.DescricaoTipo, cmd);
                                    p.get_Parameter(Autodesk.Revit.DB.BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).Set(tipo.ComentariosDeTipo);
                                    break;

                                default:
                                    break;
                            }
                            

                            t.Commit();
                        }

                } 


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
                case "Kg/m²":
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
