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
using wf = System.Windows.Forms;
using Autodesk.Revit.ApplicationServices;
using System.Windows;
using System.Xml.Serialization;

namespace POB
{



    [TransactionAttribute(TransactionMode.Manual)]
    [RegenerationAttribute(RegenerationOption.Manual)]

    public class ExtrairParaSat : IExternalCommand
    {
        public static List<Element> ListaComSubElementos = new List<Element>();
        FamilySymbol fs1;
        private List<ExportarDados> _expotar;

        public void BuscarAninhados(Element ele)
        {

            if (ele is FamilyInstance)
            {
                FamilyInstance aFamilyInst = ele as FamilyInstance;
                if (aFamilyInst.SuperComponent == null)
                {
                    var subElements = aFamilyInst.GetSubComponentIds();
                    if (subElements.Count != 0)
                    {
                        foreach (var aSubElemId in subElements)
                        {
                            var aSubElem = ele.Document.GetElement(aSubElemId);
                            if (aSubElem is FamilyInstance)
                            {
                                ListaComSubElementos.Add(aSubElem);
                                BuscarAninhados(aSubElem);
                            }
                        }
                    }
                }
                else
                {
                    var subElements1 = aFamilyInst.GetSubComponentIds();
                    if (subElements1.Count != 0)
                    {
                        foreach (var aSubElemId in subElements1)
                        {
                            var aSubElem1 = ele.Document.GetElement(aSubElemId);
                            if (aSubElem1 is FamilyInstance)
                            {
                                ListaComSubElementos.Add(aSubElem1);
                                BuscarAninhados(aSubElem1);
                            }
                        }
                    }
                }
            }

        }
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData revit,
            ref string message, ElementSet elements)
        {
            UIApplication uiApp = revit.Application;
            Document uiDoc = uiApp.ActiveUIDocument.Document;
            Selection sel = uiApp.ActiveUIDocument.Selection;

            ListaComSubElementos.Clear();
            foreach (ElementId eleId in uiApp.ActiveUIDocument.Selection.GetElementIds())
            {
                ListaComSubElementos.Add(uiDoc.GetElement(eleId));
                BuscarAninhados(uiDoc.GetElement(eleId));
            }
            _expotar = new List<ExportarDados>();
            foreach (var item in ListaComSubElementos)
            {
                foreach (var pInstancai in item.Parameters.Cast<Parameter>().ToList())
                {
                    var valor = new ExportarDados();
                    try
                    {
                      
                        valor.Categoria = item.Category.Name;
                        valor.Familia = GetFamilia(item);
                        valor.NomeDoParametro = pInstancai.Definition.Name;
#if D23 || D24
                        valor.TipoDeDado = pInstancai.Definition.GetDataType().ToString();

#else
 valor.TipoDeDado = pInstancai.Definition.ParameterType.ToString();
#endif
                        valor.Armazenamento = pInstancai.StorageType.ToString();
                        valor.Id = item.Id.IntegerValue;
                        valor.Tipo = "Instancia";
                        if (pInstancai.HasValue)
                        {
                            valor.Valor = GetValor(pInstancai);
                            valor.ValorEmTexto = pInstancai.AsValueString();
                        }

                        else
                        {
                            valor.Valor = "";
                            valor.ValorEmTexto = "";
                        }

                }
                    catch(Exception e)
                    {
                        valor.Erro = e.Message;
                    }

                    _expotar.Add(valor);
                }
                if(item is FamilyInstance)
                {
                    try
                    {
                        var fs = (item as FamilyInstance).Symbol;
                        foreach (var pTipo in fs.Parameters.Cast<Parameter>().ToList())
                        {
                            var valor = new ExportarDados();
                            try
                            {

                                valor.Categoria = item.Category.Name;
                                valor.Familia = GetFamilia(item);
                                valor.NomeDoParametro = pTipo.Definition.Name;
#if D23 || D24
                                valor.TipoDeDado = pTipo.Definition.GetDataType().ToString();

#else
valor.TipoDeDado = pTipo.Definition.ParameterType.ToString();
#endif


                                valor.Armazenamento = pTipo.StorageType.ToString();
                                valor.Id = item.Id.IntegerValue;
                                valor.Tipo = "Tipo";
                                if (pTipo.HasValue)
                                {
                                    valor.Valor = GetValor(pTipo);
                                    valor.ValorEmTexto = pTipo.AsValueString();
                                }
                                else
                                {
                                    valor.Valor = "";
                                    valor.ValorEmTexto = "";
                                }

                            }
                            catch (Exception e)
                            {
                                valor.Erro = e.Message;
                            }

                            _expotar.Add(valor);
                        }
                    }
                    catch
                    {

                    }
                }
            }

            string caminho =@"C:\Users\Larsson\Documents\saida.xml";
            try
            {
                //cria o serializador da lista
                XmlSerializer serialiser = new XmlSerializer(typeof(List<ExportarDados>));

                //Cria o textWriter com o arquivo
                System.IO.TextWriter filestream = new System.IO.StreamWriter(caminho);

                //Gravação dos arquivos
                serialiser.Serialize(filestream, _expotar);


                //Fecha o arquivo
                filestream.Close();

            }
            catch (Exception ex)
            {
            }
            return Result.Cancelled;
            StringBuilder sb = new StringBuilder();
            string linha = "";
            foreach (var item in typeof(ExportarDados).GetProperties())
            {
                linha = linha + item.Name + "\t";
            }
            sb.AppendLine(linha);
            linha = "";
            foreach (var item in _expotar)
            {
                foreach (var prop in item.GetType().GetProperties())
                {
                    var novoValor = "";
                    try
                    {
                       novoValor = prop.GetValue(item).ToString();
                    }
                    catch
                    {
                        novoValor = "";
                    }
                    linha = linha + "\t" + novoValor;
                }
                sb.AppendLine(linha);
                linha = "";
            }
            Clipboard.SetText(sb.ToString());
            return Result.Cancelled;

        }

        private string GetValor(Parameter parametro)
        {
            /*switch (parametro.Definition.ParameterType)
            {

                case ParameterType.Integer:
                    return parametro.AsInteger().ToString();
                    break;
                case ParameterType.Number:
                    return parametro.AsDouble().ToString().Replace(".",",");
                    break;
                case ParameterType.Length:
                    return (parametro.AsDouble() * 0.3048).ToString().Replace(".", ","); 
                    break;
                case ParameterType.Area:
                    return (parametro.AsDouble() * 0.304800610 * 0.304800610).ToString().Replace(".", ",");
                    break;
                case ParameterType.Volume:
                    return (parametro.AsDouble() * 0.3048 * 0.3048 * 0.3048).ToString().Replace(".", ",");
                    break;
                case ParameterType.Angle:
                    return (parametro.AsDouble()).ToString().Replace(".", ",");
                    break;

                case ParameterType.WireSize:
                    return (parametro.AsDouble() * 0.3048).ToString().Replace(".", ",");
                    break;

                case ParameterType.Mass:
                    return (parametro.AsDouble()).ToString().Replace(".", ",");
                    break;


                case ParameterType.SurfaceArea:
                    return (parametro.AsDouble() * 0.3048 * 0.3048).ToString().Replace(".", ",");
                    break;

                case ParameterType.Weight:
                    return (parametro.AsDouble()*0.3048).ToString().Replace(".", ",");
                    break;

                    
            }*/
            
            return parametro.AsValueString();
        }

        private object GetFamilia(Element item)
        {
            if (item is FamilyInstance)
                return (item as FamilyInstance).Symbol.FamilyName;
            else return "To-do";
        }
    }




}
