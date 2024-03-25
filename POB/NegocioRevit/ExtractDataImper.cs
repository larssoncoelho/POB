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
using System.Windows.Media.Media3D;


namespace POB.NegocioRevit
{
    public static class ExtractDataImper
    {

        public static List<POB.ObjetoTransferenciaPOB.ResumoExcel> listaElemento = new List<POB.ObjetoTransferenciaPOB.ResumoExcel>();
        public static string GetViewRef(Autodesk.Revit.DB.AssemblyInstance assemblyInstance, List<Autodesk.Revit.DB.ViewPlan> lista, Document uiDoc)
        {
            var vistasReferenciadas = "";

            foreach (var vista  in lista)
            {
                var p = vista.LookupParameter("Número da folha");
                if (p != null)
                    if (p.HasValue)
                        if (!string.IsNullOrEmpty(p.AsString()))
                        {
                            var elementos = new FilteredElementCollector(uiDoc, vista.Id).Cast<Element>();
                            var elemento = elementos.Where(x => x.Id == assemblyInstance.Id).ToList();
                            if (elemento != null)
                                if (elemento.Count > 0)
                                    vistasReferenciadas = vistasReferenciadas + p.AsString() + "-".ToLower();
                        }
            }
            
            if(vistasReferenciadas!="")
                vistasReferenciadas = vistasReferenciadas.Substring(0, vistasReferenciadas.Length - 1);
            return   vistasReferenciadas;
        }
        public static ResultadoExternalCommandData Execute(ExternalCommandData revit, bool criarTrasancao)
        {
            UIApplication uiApp = revit.Application;
            Document uiDoc = uiApp.ActiveUIDocument.Document;
            Selection sel = uiApp.ActiveUIDocument.Selection;
            XYZ P = new XYZ(0, 0, 0);
            var V = uiDoc.ActiveView;


            var vistas = new FilteredElementCollector(uiDoc).OfClass(typeof(Autodesk.Revit.DB.ViewPlan)).WhereElementIsNotElementType().Cast<Autodesk.Revit.DB.ViewPlan>().ToList();

            listaElemento.Clear();
            foreach (ElementId id in sel.GetElementIds())
            {
               
                Element ele = uiDoc.GetElement(id);
                if (!(ele is Autodesk.Revit.DB.AssemblyInstance)) continue;
                Autodesk.Revit.DB.AssemblyInstance assemblyInstance = ele as AssemblyInstance;
                string tocPavimento = assemblyInstance.LookupParameter("tocPavimento").AsString();
                string tocAmbiente = assemblyInstance.LookupParameter("tocAmbiente").AsString();
                string tocFolha = GetViewRef(assemblyInstance, vistas, uiDoc);
                string tocNUmeroReferencia = assemblyInstance.LookupParameter("tocNumAmbiente").AsString();
                string tocCodImper = assemblyInstance.LookupParameter("tocCodigoImper").AsString();
                string tocNomeSistema = assemblyInstance.LookupParameter("tocNomeSistema").AsString();
                double tocAreaTotal = assemblyInstance.LookupParameter("tocAreaTotal").AsDouble() * 0.3048 * 0.3048;
                //Material materialImper = uiDoc.GetElement(new ElementId(-1)) as Material;

                foreach (var item in assemblyInstance.GetMemberIds())
                {
                    Element ele1 = uiDoc.GetElement(item);
                    if (ele1 is Wall)
                    {
                        var tocDesconsiderarElemento = ele1.LookupParameter("tocDesconsiderarElemento");
                        if (tocDesconsiderarElemento != null)
                            if (tocDesconsiderarElemento.HasValue)
                                if (tocDesconsiderarElemento.AsInteger() != 0)
                                {
                                    Autodesk.Revit.DB.WallType tipoDeParede = (ele1 as Autodesk.Revit.DB.Wall).WallType;
                                    foreach (CompoundStructureLayer camada in tipoDeParede.GetCompoundStructure().GetLayers())
                                    {
                                        Autodesk.Revit.DB.Material material = uiDoc.GetElement(camada.MaterialId) as Autodesk.Revit.DB.Material;
                                        POB.ObjetoTransferenciaPOB.ResumoExcel resumoExcel = new POB.ObjetoTransferenciaPOB.ResumoExcel();
                                        resumoExcel.Pavimento = tocPavimento;
                                        resumoExcel.Ambiente = tocAmbiente;
                                        resumoExcel.Folha = tocFolha;
                                        resumoExcel.NumFolhaReferencia = tocNUmeroReferencia;
                                        resumoExcel.CategoriaDoMaterial = material.LookupParameter("tocCategoriaMaterial").AsString();
                                        resumoExcel.NomeDoMaterial = material.Name;

                                        resumoExcel.AreaHorizontal = (ele1 as Autodesk.Revit.DB.Wall).GetMaterialArea(camada.MaterialId, false) * 0.3048 * 0.3048;
                                        resumoExcel.AreaVertical = 0.000000000;
                                        resumoExcel.tocRepeticoesNoPavimento = 1.0000000;
                                        resumoExcel.tocRepeticoesDoPavimento = 1.0000000;
                                        resumoExcel.tocRepeticoesDeTorres = 1.0000000;
                                        resumoExcel.CodigoImper = tocCodImper;
                                        resumoExcel.NomeSistema = tocNomeSistema;
                                        resumoExcel.AlturaDaParade = (ele1 as Wall).LookupParameter("Altura desconectada").AsDouble() * 0.3048;
                                        resumoExcel.IdMontagem = assemblyInstance.Id.IntegerValue;
                                        resumoExcel.IdElemento = ele1.Id.IntegerValue;
                                        resumoExcel.IdMaterial = camada.MaterialId.IntegerValue;
                                        resumoExcel.AreaTotal = tocAreaTotal;
                                        var l = listaElemento.Where(x => (x.IdMontagem == assemblyInstance.Id.IntegerValue) && (x.IdElemento == ele1.Id.IntegerValue) && (x.IdMaterial==camada.MaterialId.IntegerValue)).ToList();
                                        if(l!=null)
                                            if(l.Count==0)
                                                listaElemento.Add(resumoExcel);


                                    }

                                }
                        if ((tocDesconsiderarElemento == null) || (!tocDesconsiderarElemento.HasValue))
                        {
                            Autodesk.Revit.DB.WallType tipoDeParede = (ele1 as Autodesk.Revit.DB.Wall).WallType;
                            foreach (CompoundStructureLayer camada in tipoDeParede.GetCompoundStructure().GetLayers())
                            {
                                Autodesk.Revit.DB.Material material = uiDoc.GetElement(camada.MaterialId) as Autodesk.Revit.DB.Material;
                                POB.ObjetoTransferenciaPOB.ResumoExcel resumoExcel = new POB.ObjetoTransferenciaPOB.ResumoExcel();
                                resumoExcel.Pavimento = tocPavimento;
                                resumoExcel.Ambiente = tocAmbiente;
                                resumoExcel.Folha = tocFolha;
                                resumoExcel.NumFolhaReferencia = tocNUmeroReferencia;
                                resumoExcel.CategoriaDoMaterial = material.LookupParameter("tocCategoriaMaterial").AsString();
                                resumoExcel.NomeDoMaterial = material.Name;
                                resumoExcel.AreaHorizontal = (ele1 as Autodesk.Revit.DB.Wall).GetMaterialArea(camada.MaterialId, false) * 0.3048 * 0.3048;
                                resumoExcel.AreaVertical = 0.000000000;
                                resumoExcel.tocRepeticoesNoPavimento = 1.0000000;
                                resumoExcel.tocRepeticoesDoPavimento = 1.0000000;
                                resumoExcel.tocRepeticoesDeTorres = 1.0000000;
                                resumoExcel.CodigoImper = tocCodImper;
                                resumoExcel.NomeSistema = tocNomeSistema;
                                resumoExcel.AlturaDaParade = (ele1 as Wall).LookupParameter("Altura desconectada").AsDouble() * 0.3048;
                                resumoExcel.IdMontagem = assemblyInstance.Id.IntegerValue;
                                resumoExcel.IdElemento = ele1.Id.IntegerValue;
                                resumoExcel.IdMaterial = camada.MaterialId.IntegerValue;
                                resumoExcel.TipoDeElemento = "Parede";
                                resumoExcel.ComprimentoJunta = 0.0000;
                                resumoExcel.AreaTotal = tocAreaTotal;
                                var l = listaElemento.Where(x => (x.IdMontagem == assemblyInstance.Id.IntegerValue) && (x.IdElemento == ele1.Id.IntegerValue) && (x.IdMaterial == camada.MaterialId.IntegerValue)).ToList();
                                if (l != null)
                                    if (l.Count == 0)
                                        listaElemento.Add(resumoExcel);
                              

                            }
                        }
                    }
                    if (ele1.Category.Id == Autodesk.Revit.DB.Category.GetCategory(uiDoc, BuiltInCategory.OST_GenericModel).Id)
                    {

                        POB.ObjetoTransferenciaPOB.ResumoExcel resumoExcel = new POB.ObjetoTransferenciaPOB.ResumoExcel();
                        resumoExcel.Pavimento = tocPavimento;
                        resumoExcel.Ambiente = tocAmbiente;
                        resumoExcel.Folha = tocFolha;
                        resumoExcel.NumFolhaReferencia = tocNUmeroReferencia;
                        resumoExcel.CategoriaDoMaterial = assemblyInstance.LookupParameter("tocCodigoImper").AsString();
                        resumoExcel.NomeDoMaterial = "Junta";
                        resumoExcel.AreaHorizontal = 0.0000;
                        resumoExcel.AreaVertical = 0.000000000;
                        resumoExcel.tocRepeticoesNoPavimento = 1.0000000;
                        resumoExcel.tocRepeticoesDoPavimento = 1.0000000;
                        resumoExcel.tocRepeticoesDeTorres = 1.0000000;
                        resumoExcel.CodigoImper = tocCodImper;
                        resumoExcel.NomeSistema = tocNomeSistema;
                        resumoExcel.AlturaDaParade = 0.00;// (ele1 as Wall).LookupParameter("Altura desconectada").AsDouble() * 0.3048;
                        resumoExcel.IdMontagem = assemblyInstance.Id.IntegerValue;
                        resumoExcel.IdElemento = ele1.Id.IntegerValue;
                        resumoExcel.IdMaterial = -1;
                        resumoExcel.TipoDeElemento = "Junta";
                        resumoExcel.ComprimentoJunta = ele1.LookupParameter("tocComprimentoJunta").AsDouble() * 0.3048;
                        resumoExcel.AreaTotal = 0;
                        listaElemento.Add(resumoExcel);



                    }
                    if (ele1 is Floor)
                    {
                        Floor floor = ele1 as Floor;


                        Autodesk.Revit.DB.FloorType tipoDePiso = floor.FloorType;
                        foreach (CompoundStructureLayer camada in tipoDePiso.GetCompoundStructure().GetLayers())
                        {
                            Autodesk.Revit.DB.Material material = uiDoc.GetElement(camada.MaterialId) as Autodesk.Revit.DB.Material;
                            POB.ObjetoTransferenciaPOB.ResumoExcel resumoExcel = new POB.ObjetoTransferenciaPOB.ResumoExcel();
                            resumoExcel.Pavimento = tocPavimento;
                            resumoExcel.Ambiente = tocAmbiente;
                            resumoExcel.Folha = tocFolha;
                            resumoExcel.NumFolhaReferencia = tocNUmeroReferencia;
                            resumoExcel.CategoriaDoMaterial = material.LookupParameter("tocCategoriaMaterial").AsString();
                            resumoExcel.NomeDoMaterial = material.Name;
                            resumoExcel.AreaHorizontal = 0.000000;
                            resumoExcel.AreaVertical = floor.GetMaterialArea(camada.MaterialId, false) * 0.3048 * 0.3048;
                            resumoExcel.tocRepeticoesNoPavimento = 1.0000000;
                            resumoExcel.tocRepeticoesDoPavimento = 1.0000000;
                            resumoExcel.tocRepeticoesDeTorres = 1.0000000;
                            resumoExcel.CodigoImper = tocCodImper;
                            resumoExcel.NomeSistema = tocNomeSistema;
                            resumoExcel.AlturaDaParade = 0.000000;
                            resumoExcel.IdMontagem = assemblyInstance.Id.IntegerValue;
                            resumoExcel.IdElemento = floor.Id.IntegerValue;
                            resumoExcel.IdMaterial = camada.MaterialId.IntegerValue;
                            resumoExcel.TipoDeElemento = "Piso";
                            resumoExcel.ComprimentoJunta = 0.0000;
                            resumoExcel.AreaTotal = tocAreaTotal;
                            var l = listaElemento.Where(x => (x.IdMontagem == assemblyInstance.Id.IntegerValue) && (x.IdElemento == floor.Id.IntegerValue) && (x.IdMaterial == camada.MaterialId.IntegerValue)).ToList();
                            if (l != null)
                                if (l.Count == 0)
                                    listaElemento.Add(resumoExcel);

                        }



                    }
                    if (ele1 is Autodesk.Revit.DB.RoofBase)
                    {
                        Autodesk.Revit.DB.RoofBase roof = ele1 as Autodesk.Revit.DB.RoofBase;


                        Autodesk.Revit.DB.RoofType tipoDeCobertura = roof.RoofType;
                        foreach (CompoundStructureLayer camada in tipoDeCobertura.GetCompoundStructure().GetLayers())
                        {
                            Autodesk.Revit.DB.Material material = uiDoc.GetElement(camada.MaterialId) as Autodesk.Revit.DB.Material;
                            POB.ObjetoTransferenciaPOB.ResumoExcel resumoExcel = new POB.ObjetoTransferenciaPOB.ResumoExcel();
                            resumoExcel.Pavimento = tocPavimento;
                            resumoExcel.Ambiente = tocAmbiente;
                            resumoExcel.Folha = tocFolha;
                            resumoExcel.NumFolhaReferencia = tocNUmeroReferencia;
                            resumoExcel.CategoriaDoMaterial = material.LookupParameter("tocCategoriaMaterial").AsString();
                            resumoExcel.NomeDoMaterial = material.Name;
                            resumoExcel.AreaHorizontal = 0.000000;
                            resumoExcel.AreaVertical = roof.GetMaterialArea(camada.MaterialId, false) * 0.3048 * 0.3048;
                            resumoExcel.tocRepeticoesNoPavimento = 1.0000000;
                            resumoExcel.tocRepeticoesDoPavimento = 1.0000000;
                            resumoExcel.tocRepeticoesDeTorres = 1.0000000;
                            resumoExcel.CodigoImper = tocCodImper;
                            resumoExcel.NomeSistema = tocNomeSistema;
                            resumoExcel.AlturaDaParade = 0.000000;
                            resumoExcel.IdMontagem = assemblyInstance.Id.IntegerValue;
                            resumoExcel.IdElemento = roof.Id.IntegerValue;
                            resumoExcel.IdMaterial = camada.MaterialId.IntegerValue;
                            resumoExcel.TipoDeElemento = "Cobertura";
                            resumoExcel.ComprimentoJunta = 0.0000;
                            resumoExcel.AreaTotal = tocAreaTotal;
                            var l = listaElemento.Where(x => (x.IdMontagem == assemblyInstance.Id.IntegerValue) && (x.IdElemento == roof.Id.IntegerValue) && (x.IdMaterial == camada.MaterialId.IntegerValue)).ToList();
                            if (l != null)
                                if (l.Count == 0)
                                    listaElemento.Add(resumoExcel);

                        }


                    }
                }
            }
            return new ResultadoExternalCommandData { Resultado = Result.Succeeded };
        }
    }
}
     