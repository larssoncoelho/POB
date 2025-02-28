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
    public static class DadosImperCommand
    {
        public static double TocEspessuraTotal { get; private set; }

        /* public static string GettocCodigoImper(Document uiDoc, Element assemblyInstance)
{
ElementId idParametroGlobal = GlobalParametersManager.FindByName(uiDoc, assemblyInstance.LookupParameter("tocCodigoImper").AsString());
if ((idParametroGlobal != null) | (idParametroGlobal.IntegerValue != -1))
{
string nome = ((uiDoc.GetElement(idParametroGlobal) as GlobalParameter).GetValue() as StringParameterValue).Value;
var vetor = nome.Split('|');
string texto => vetor.Length==2 ? vetor 
if (vetor.Length == 2 )

return vetor[0];

else

return vetor[1];

}
}*/
        public static ResultadoExternalCommandData Execute(ExternalCommandData revit, bool criarTrasancao)
        {
            UIApplication uiApp = revit.Application;
            Document uiDoc = uiApp.ActiveUIDocument.Document;
            Selection sel = uiApp.ActiveUIDocument.Selection;
            XYZ P = new XYZ(0, 0, 0);
            List<ElementoResumido> listaElemento = new List<ElementoResumido>();

            FilteredElementCollector collector = new FilteredElementCollector(uiDoc);

            var listaDeImagens = collector.OfClass(typeof(ImageType)).Cast<ImageType>().ToList();

            foreach (ElementId id in sel.GetElementIds())
            {
                Transaction t = new Transaction(uiDoc);
                if (criarTrasancao) t.Start("t");
                Element ele = uiDoc.GetElement(id);
                if (!(ele is Autodesk.Revit.DB.AssemblyInstance)) continue;
                Autodesk.Revit.DB.AssemblyInstance assemblyInstance = ele as AssemblyInstance;
                //Material materialImper = uiDoc.GetElement(new ElementId(-1)) as Material;
                listaElemento.Clear();
                var tocNomeSistema = "";
                var tocVup = "";
                var tocOrdem = ""; 
                

                try
                {
                    ElementId idParametroGlobal = GlobalParametersManager.FindByName(uiDoc, assemblyInstance.LookupParameter("tocCodigoImper").AsString());
                    if ((idParametroGlobal != null) | (idParametroGlobal.IntegerValue != -1))
                    {
                        string nome = ((uiDoc.GetElement(idParametroGlobal) as GlobalParameter).GetValue() as StringParameterValue).Value;
                        var vetor = nome.Split('|');
                        if (vetor.Length >= 2)
                        {
                            tocNomeSistema = vetor[0];
                            tocVup = vetor[1];
                            tocOrdem= vetor[2];
                            assemblyInstance.LookupParameter("tocNomeSistema").Set(tocNomeSistema);
                            assemblyInstance.LookupParameter("tocVUP").Set(tocVup);
                            assemblyInstance.LookupParameter("tocOrdem").Set(Convert.ToInt32(tocOrdem));
                        }
                        else
                        {
                            tocNomeSistema = nome;
                            assemblyInstance.LookupParameter("tocNomeSistema").Set(tocNomeSistema);
                        }
                    }
                }
                catch
                {

                }

                foreach (var item in assemblyInstance.GetMemberIds())
                {
                    Element ele1 = uiDoc.GetElement(item);
                    if (/*!(ele1 is Autodesk.Revit.DB.FamilyInstance) & (!(ele1 as Wall).IsStackedWall) */
                        ele1 is Wall)
                    {
                        var tocDesconsiderarElemento = ele1.LookupParameter("tocDesconsiderarElemento");
                        if (tocDesconsiderarElemento != null)
                            if (tocDesconsiderarElemento.HasValue)
                                if (tocDesconsiderarElemento.AsInteger() != 0)
                                {
                                   /* tocNomeSistema = "";
                                    var tocVup = "";
                                    var tocOrdem = "";*/
                                    ele1.LookupParameter("tocCodigoImper").Set(assemblyInstance.LookupParameter("tocCodigoImper").AsString());
                                    ele1.LookupParameter("tocPavimento").Set(assemblyInstance.LookupParameter("tocPavimento").AsString());
                                    ele1.LookupParameter("tocAreaTotal").Set(ele1.LookupParameter("tocAreaTotal").AsDouble());

                                    ele1.LookupParameter("tocAmbiente").Set(assemblyInstance.LookupParameter("tocAmbiente").AsString());
                                    ele1.LookupParameter("tocNomeSistema").Set(tocNomeSistema);
                                    ele1.LookupParameter("tocVUP").Set(tocVup);
                                    ele1.LookupParameter("tocOrdem").Set(tocOrdem);

                                    listaElemento.Add(new ElementoResumido
                                    {
                                        Tipo = "Parede",
                                        TocAlturaRodape = (ele1 as Wall).LookupParameter("Altura desconectada").AsDouble(),
                                        TocAreaRodape = (ele1 as Wall).LookupParameter("Área").AsDouble(),
                                        TocEspessuraMaxima = 0,
                                        TocPerimetro = (ele1 as Wall).LookupParameter("Comprimento").AsDouble(),
                                        TocVolume = (ele1 as Wall).LookupParameter("Volume").AsDouble(),
                                        TocLarguraParede = (ele1 as Wall).WallType.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM).AsDouble()
                                    });
                                }
                        if((tocDesconsiderarElemento==null)||(!tocDesconsiderarElemento.HasValue))
                        {
                            var p1 = ele1.LookupParameter("tocCodigoImper");
                            if (!p1.IsReadOnly) { 
                                p1.Set(assemblyInstance.LookupParameter("tocCodigoImper").AsString());
                                ele1.LookupParameter("tocAmbiente").Set(assemblyInstance.LookupParameter("tocAmbiente").AsString());
                                ele1.LookupParameter("tocPavimento").Set(assemblyInstance.LookupParameter("tocPavimento").AsString());
                                ele1.LookupParameter("tocAreaTotal").Set(ele1.LookupParameter("Área").AsDouble());

                                ele1.LookupParameter("tocNomeSistema").Set(tocNomeSistema);
                                ele1.LookupParameter("tocVUP").Set(tocVup);
                                ele1.LookupParameter("tocOrdem").Set(tocOrdem);
                                listaElemento.Add(new ElementoResumido
                                {
                                    Tipo = "Parede",
                                    TocAlturaRodape = (ele1 as Wall).LookupParameter("Altura desconectada").AsDouble(),
                                    TocAreaRodape = (ele1 as Wall).LookupParameter("Área").AsDouble(),
                                    TocEspessuraMaxima = 0,
                                    TocPerimetro = (ele1 as Wall).LookupParameter("Comprimento").AsDouble(),
                                    TocVolume = (ele1 as Wall).LookupParameter("Volume").AsDouble(),
                                    TocLarguraParede = (ele1 as Wall).WallType.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM).AsDouble()
                                });
                            }
                        }
                    }
                  /*  if (!(ele1 is Autodesk.Revit.DB.FamilyInstance) & ((ele1 as Wall).IsStackedWall))
                    {
                        var listaParedes = (ele1 as Wall).GetStackedWallMemberIds().ToList();
                        foreach (var paredeEmpilhada in listaParedes)
                        {

                        }
                    }*/
                    if (ele1.Category.Id == Autodesk.Revit.DB.Category.GetCategory(uiDoc, BuiltInCategory.OST_GenericModel).Id)
                    {

                        ele1.LookupParameter("tocCodigoImper").Set(assemblyInstance.LookupParameter("tocCodigoImper").AsString());
                        ele1.LookupParameter("tocPavimento").Set(assemblyInstance.LookupParameter("tocPavimento").AsString());
                        ele1.LookupParameter("tocAreaTotal").Set(ele1.LookupParameter("Área").AsDouble());

                        ele1.LookupParameter("tocAmbiente").Set(assemblyInstance.LookupParameter("tocAmbiente").AsString());
                        ele1.LookupParameter("tocNomeSistema").Set(tocNomeSistema);
                        ele1.LookupParameter("tocVUP").Set(tocVup);
                        ele1.LookupParameter("tocOrdem").Set(tocOrdem);
                        listaElemento.Add(new ElementoResumido
                        {
                            Tipo = "Junta",
                            TocAlturaRodape = 0,
                            TocAreaRodape = 0,
                            TocEspessuraMaxima = 0,
                            TocPerimetro = 0,
                            TocVolume = 0,
                            TocLarguraParede = 0,
                            TocComprimentoJunta = ele1.LookupParameter("tocComprimentoJunta").AsDouble()
                        }) ;

                    }
                    if (ele1 is Floor)
                    {
                        Floor floor = ele1 as Floor;
                        double pontoMaximoReal = 0;
                        double pontoMinimoReal = 0;
                        double pontoMaisBaixoDaFaceInferior = 0;
                        double espessuraMinimaCamadaVariavel = 0;
                        double espessuraMaximaCamadaVariavel = 0;
                        double espessuraAcabamentoCamadaVariavel = 0;
                        var lista = Util.GetPointsDoPiso(floor).Distinct().ToList();
                        var listaSemOsPontosDaBase = (from a in lista  select a).ToList();

                        var p1 = floor.LookupParameter("tocCodigoImper");
                        if (!p1.IsReadOnly)
                        {
                            p1.Set(assemblyInstance.LookupParameter("tocCodigoImper").AsString());
                            ele1.LookupParameter("tocAmbiente").Set(assemblyInstance.LookupParameter("tocAmbiente").AsString());
                            ele1.LookupParameter("tocAreaTotal").Set(ele1.LookupParameter("Área").AsDouble());

                            ele1.LookupParameter("tocPavimento").Set(assemblyInstance.LookupParameter("tocPavimento").AsString());
                            ele1.LookupParameter("tocNomeSistema").Set(tocNomeSistema);
                            ele1.LookupParameter("tocVUP").Set(tocVup);
                            ele1.LookupParameter("tocOrdem").Set(tocOrdem);
                        }

                            double tocDescontarEspessuraRegularizacao = 0;

                        var parDesconto = ele1.LookupParameter("tocDescontarEspessuraRegularizacao");
                        if (parDesconto != null)
                            if (parDesconto.HasValue)
                                tocDescontarEspessuraRegularizacao = parDesconto.AsDouble();

                        pontoMaisBaixoDaFaceInferior = lista.Min(x => x.Z);
                        listaSemOsPontosDaBase.RemoveAll(x => x.Z == pontoMaisBaixoDaFaceInferior);
                        if (lista.Count > 0) pontoMaximoReal = lista.Max(x => x.Z);
                        if (listaSemOsPontosDaBase.Count > 0) pontoMinimoReal = listaSemOsPontosDaBase.Min(x => x.Z);
                        listaElemento.Add(new ElementoResumido
                        {
                            Tipo = "Piso",
                            TocAreaPiso = (ele1 as Floor).LookupParameter("Área").AsDouble(),
                            TocVolume = (ele1 as Floor).LookupParameter("Volume").AsDouble(),
                            
                            TocEspessuraRealMinima = GetEspessuraRealMinima((ele1 as Floor)),
                            TocEspessuraAcabamento = GetEspessuraAcabamento((ele1 as Floor)),
                            TocEspessuraRealMaxima = pontoMaximoReal - GetEspessuraAcabamento((ele1 as Floor)) - pontoMaisBaixoDaFaceInferior,
                            TocEspessuraTotal = pontoMaximoReal - pontoMaisBaixoDaFaceInferior,
                            TocDescontarEspessuraRegularizacao = tocDescontarEspessuraRegularizacao

                        });
                        (ele1 as Floor).LookupParameter("tocEspessuraRealMaxima").Set(Convert.ToDouble(pontoMaximoReal - GetEspessuraAcabamento((ele1 as Floor)) - pontoMaisBaixoDaFaceInferior- tocDescontarEspessuraRegularizacao));
                    }
                    if (ele1 is RoofBase)
                    {
                        RoofBase floor = ele1 as RoofBase;
                        var p1 = floor.LookupParameter("tocCodigoImper");
                        if (!p1.IsReadOnly)
                        {

                            p1.Set(assemblyInstance.LookupParameter("tocCodigoImper").AsString());
                            ele1.LookupParameter("tocAmbiente").Set(assemblyInstance.LookupParameter("tocAmbiente").AsString());
                            ele1.LookupParameter("tocPavimento").Set(assemblyInstance.LookupParameter("tocPavimento").AsString());
                            ele1.LookupParameter("tocAreaTotal").Set(ele1.LookupParameter("Área").AsDouble());

                            ele1.LookupParameter("tocNomeSistema").Set(tocNomeSistema);
                            ele1.LookupParameter("tocVup").Set(tocVup);
                            ele1.LookupParameter("tocOrdem").Set(tocOrdem);
                        }
                        listaElemento.Add(new ElementoResumido
                        {
                            Tipo = "Telhado",
                            TocAreaPiso = (ele1 as RoofBase).LookupParameter("Área").AsDouble(),
                            TocVolume = (ele1 as RoofBase).LookupParameter("Volume").AsDouble(),
                            TocEspessuraRealMinima = floor.get_Parameter(BuiltInParameter.ROOF_ATTR_THICKNESS_PARAM).AsDouble(),
                            TocEspessuraAcabamento = floor.get_Parameter(BuiltInParameter.ROOF_ATTR_THICKNESS_PARAM).AsDouble(),
                            TocEspessuraRealMaxima = floor.get_Parameter(BuiltInParameter.ROOF_ATTR_THICKNESS_PARAM).AsDouble(),
                            TocEspessuraTotal = floor.get_Parameter(BuiltInParameter.ROOF_ATTR_THICKNESS_PARAM).AsDouble()
                        });
                        (ele1 as RoofBase).LookupParameter("tocEspessuraRealMaxima").Set(floor.get_Parameter(BuiltInParameter.ROOF_ATTR_THICKNESS_PARAM).AsDouble());


                    }
                }
                

                double? tocDescontarEspessuraRegulazacaoMontagem = 0;
                try
                {
                    tocDescontarEspessuraRegulazacaoMontagem = listaElemento.Where(x => x.Tipo == "Piso").Min(x => x.TocDescontarEspessuraRegularizacao);
                }
                catch
                {
                }

                try
                {
                    var imagemSelecionada = listaDeImagens.Where(x => x.Name.ToUpper() ==
                              (assemblyInstance.LookupParameter("tocCodigoImper").AsString() + ".PNG").ToUpper()).FirstOrDefault();
                    assemblyInstance.LookupParameter("tocImagemImper").Set(imagemSelecionada.Id);
                }
                catch
                {

                }

                double? tocAreaPiso = 0;
                try
                {
                    tocAreaPiso = listaElemento.Where(x => (x.Tipo == "Piso") | (x.Tipo == "Telhado")  ).Sum(x => x.TocAreaPiso);

                    assemblyInstance.LookupParameter("tocAreaPiso").Set(Convert.ToDouble(tocAreaPiso));
                }
                catch
                {

                }
                try
                {
                    double? tocEspessuraRealMaxima = 0;
                    tocEspessuraRealMaxima = listaElemento.Where(x => (x.Tipo == "Piso") | (x.Tipo == "Telhado")).Max(x => x.TocEspessuraRealMaxima);
                    var espessuraRealMaxima = Util.GetParameter(ele, "tocEspessuraRealMaxima",
#if D23 || D24
                        SpecTypeId.Length,
#else
  ParameterType.Length, 
#endif



                        true, false);
                    espessuraRealMaxima.Set(Convert.ToDouble(tocEspessuraRealMaxima- tocDescontarEspessuraRegulazacaoMontagem));
                }
                catch
                {

                }
              
                try
                {
                    double? tocEspessuraRealMinima = 0;
                    tocEspessuraRealMinima = listaElemento.Where(x => (x.Tipo == "Piso") | (x.Tipo == "Telhado")).Min(x => x.TocEspessuraRealMinima);
                    var espessuraRealMinima = Util.GetParameter(ele, "tocEspessuraRealMinima",

#if D23 ||  D24
                        SpecTypeId.Length,
#else
 ParameterType.Length,
#endif
                        true, false);
                    espessuraRealMinima.Set(Convert.ToDouble(tocEspessuraRealMinima));
                }
                catch
                {

                }
                try
                {
                    double? tocLarguraParede = 0;
                    tocLarguraParede = listaElemento.Where(x => x.Tipo == "Parede").Max(x => x.TocLarguraParede);
                    var larguraParede = Util.GetParameter(ele, "tocLarguraParede",

#if D23 || D24
                        SpecTypeId.Length,
#else
  ParameterType.Length, 
#endif


                        true, false);
                    larguraParede.Set(Convert.ToDouble(tocLarguraParede));
                }
                catch
                {

                }
                try
                {
                    double? tocComprimentoJunta = 0;
                    tocComprimentoJunta = listaElemento.Where(x => x.Tipo == "Junta").Sum(x => x.TocComprimentoJunta);
                    var comprimentoJunta = Util.GetParameter(ele, "tocComprimentoJunta",

#if D23 || D24
                        SpecTypeId.Length,
#else
  ParameterType.Length, 
#endif


                        true, false);
                    comprimentoJunta.Set(Convert.ToDouble(tocComprimentoJunta));
                }
                catch
                {

                }
                try
                {
#if D23 || D24
                    var p1 = SpecTypeId.Length;
#else
                 var p1 =ParameterType.Length;
#endif
                    double? tocEspessuraAcabamento = 0;
                    tocEspessuraAcabamento = listaElemento.Where(x => (x.Tipo == "Piso") | (x.Tipo == "Telhado")).Max(x => x.TocEspessuraAcabamento);
                    var EspessuraAcabamento = Util.GetParameter(ele, "tocEspessuraAcabamento", p1, true, false);
                    assemblyInstance.LookupParameter("tocEspessuraAcabamento").Set(Convert.ToDouble(tocEspessuraAcabamento));
                }
                catch
                {

                }
                try
                {

#if D23 || D24
                    var p1 = SpecTypeId.Length;
#else
                 var p1 =ParameterType.Length;
#endif
                    double? tocEspessuraTotal = 0;
                    tocEspessuraTotal = listaElemento.Where(x => (x.Tipo == "Piso") | (x.Tipo == "Telhado")).Max(x => x.TocEspessuraTotal);
                    var EspessuraTotal = Util.GetParameter(ele, "tocEspessuraTotal", p1, true, false);
                    assemblyInstance.LookupParameter("tocEspessuraTotal").Set(Convert.ToDouble(tocEspessuraTotal- tocDescontarEspessuraRegulazacaoMontagem));
                }
                catch
                {

                }
                double? tocPerimetro = 0;
                try
                {
                    tocPerimetro = listaElemento.Where(x => x.Tipo == "Parede").Sum(x => x.TocPerimetro);
                    assemblyInstance.LookupParameter("tocPerimetro").Set(Convert.ToDouble(tocPerimetro));
                }
                catch
                {

                }

                double? tocAreaRodape = 0;
                try
                {
                    tocAreaRodape = listaElemento.Where(x => x.Tipo == "Parede").Sum(x => x.TocAreaRodape);
                    assemblyInstance.LookupParameter("tocAreaRodape").Set(Convert.ToDouble(tocAreaRodape));
                }
                catch
                {

                }

                double? tocAlturaRodapeMaximo = 0;
                double? tocAlturaRodapeMinimo = 0;
             


                try
                {

                    tocAlturaRodapeMaximo =  listaElemento.Where(x => x.Tipo == "Parede").Max(x => x.TocAlturaRodape);
                    tocAlturaRodapeMaximo = Math.Round(tocAlturaRodapeMaximo??0, 2);
                    assemblyInstance.LookupParameter("tocAlturaRodape").Set(Convert.ToDouble(tocAlturaRodapeMaximo));
                }
                catch
                {

                }
                try
                {

                    tocAlturaRodapeMinimo = listaElemento.Where(x => x.Tipo == "Parede").Min(x => x.TocAlturaRodape);
                    tocAlturaRodapeMinimo = Math.Round(tocAlturaRodapeMinimo ?? 0, 2);
                    assemblyInstance.LookupParameter("tocAlturaRodapeMinimo").Set(Convert.ToDouble(tocAlturaRodapeMinimo));
                }
                catch
                {

                }
                var listaTamanhoParede = new List<Alturas>();
                foreach (var item in listaElemento.Where(x=>(x.Tipo =="Parede")&&(x.TocAlturaRodape>0)).ToList())
                { var a = new Alturas();
                    a .Altura= Math.Round((item.TocAlturaRodape ?? 0) * 0.3048, 2);
                    var c = listaTamanhoParede.Where(x => x.Altura == a.Altura).Count();
                    if (c==0)
                    {
                        listaTamanhoParede.Add(a);
                    }
                }
                try
                {
                    if ((listaTamanhoParede.Count == 0))
                    {
                        assemblyInstance.LookupParameter("tocTextoAlturaRodape").Set("0,00");
                    }
                    else
                     if (listaTamanhoParede.Count > 2)
                    {
                        assemblyInstance.LookupParameter("tocTextoAlturaRodape").Set("VAR.");
                    }
                    else
                     if (listaTamanhoParede.Count == 1)
                    {
                        assemblyInstance.LookupParameter("tocTextoAlturaRodape").Set(Math.Round((tocAlturaRodapeMinimo ?? 0) * 0.3048, 3).ToString("N2").Replace(',', '.'));
                    }
                    else
                    {
                        assemblyInstance.LookupParameter("tocTextoAlturaRodape").Set(Math.Round((tocAlturaRodapeMinimo ?? 0) * 0.3048, 3).ToString("N2").Replace(',', '.') + "/" +
                           Math.Round((tocAlturaRodapeMaximo ?? 0) * 0.3048, 3).ToString("N2").Replace(',', '.'));
                    }

                }
                catch
                {

                }
                
            
                    

               
              double? tocVolume = 0;
                try
                {
                    tocVolume = listaElemento.Sum(x => x.TocVolume);

                    assemblyInstance.LookupParameter("tocVolumeImpermeabilizacao").Set(Convert.ToDouble(tocVolume));
                }
                catch
                {
                    //ReflectionContext:\
                }



                assemblyInstance.LookupParameter("tocAreaTotal").Set(Convert.ToDouble(tocAreaRodape + tocAreaPiso));



                if (criarTrasancao) t.Commit();
            }
            return new ResultadoExternalCommandData { Resultado = Result.Succeeded };
        }

        private static double? GetEspessuraAcabamento(Floor floor)
        {
            return floor.FloorType.LookupParameter("Espessura-padrão").AsDouble() - GetEspessuraRealMinima(floor);   
        }

        private static double? GetEspessuraRealMinima(Floor floor)
        {
            var compoundLayerStructura = floor.FloorType.GetCompoundStructure();
            try
            {
                var id = compoundLayerStructura.VariableLayerIndex;
                return compoundLayerStructura.GetLayerWidth(id);
            }
            catch
            {
                return null;
            }
        }
    }
}
