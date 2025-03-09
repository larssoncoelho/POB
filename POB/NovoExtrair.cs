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
using Funcoes;

using System.Windows.Forms;
using System.Windows.Media.Animation;
using Autodesk.Revit.DB.Visual;

namespace POB
{

    public class SecaoDimensoes
    {
        public double Largura { get; set; }
        public double Altura{ get; set; }

        public double Diametro { get; set; }


    }

    public static class NovoExtrair
    {

        public static ElementType ee;
        public static Autodesk.Revit.DB.Plumbing.Pipe tuboLevantamento; //= new Autodesk.Revit.DB.Plumbing.Pipe();
        public static int MEDICAO_BLOCO_ID;
        public static int SERVICO_AMO_ID;
        public static double L;
        public static string resultado1;
        public static Autodesk.Revit.DB.Floor pisoRevit;

        public static ElementId preenchimentoId;
        public static double percentAvanco;
        public static DateTime diaRealizado;
        public static List<Element> listaEle = new List<Element>();

        public static FilteredElementCollector selecao;


        public static string GetDescricao(Document uiDoc, Element item)
        {

            try
            {
                if ((item.Category.Name == Category.GetCategory(uiDoc, BuiltInCategory.OST_PipeFitting).Name) |
                    (item.Category.Name == Category.GetCategory(uiDoc, BuiltInCategory.OST_ConduitFitting).Name))
                {
                    return FuncaoExtrair.DadosConcexaoTubo(item);
                }
                if ((item.Category.Name == Category.GetCategory(uiDoc, BuiltInCategory.OST_PipeCurves).Name) |
                    (item.Category.Name == Category.GetCategory(uiDoc, BuiltInCategory.OST_Conduit).Name) |
                    (item.Category.Name == Category.GetCategory(uiDoc, BuiltInCategory.OST_FlexPipeCurves).Name))
                {
                    return FuncaoExtrair.DadosTubulacao(item);
                }
                if ((item.Category.Name == Category.GetCategory(uiDoc, BuiltInCategory.OST_DuctCurves).Name))
                {
                    return FuncaoExtrair.DadosDuto((item as Duct));
                }
                
                if ((item.Category.Name == Category.GetCategory(uiDoc, BuiltInCategory.OST_DuctFitting).Name))
                {
                    return FuncaoExtrair.DadosConexaoDuto(item);
                }

                if ((item.Category.Name == Category.GetCategory(uiDoc, BuiltInCategory.OST_CableTray).Name))
                {

                    return FuncaoExtrair.DadosBandejaDeCabos((item as Autodesk.Revit.DB.Electrical.CableTray));
                }
                if ((item.Category.Name == Category.GetCategory(uiDoc, BuiltInCategory.OST_CableTrayFitting).Name))
                {
                    return FuncaoExtrair.DadosConexaoBandejaDeCabos((item as FamilyInstance));
                }
                List<string> lista = new List<string>();

                lista.Add(Category.GetCategory(uiDoc, BuiltInCategory.OST_ElectricalFixtures).Name);
                lista.Add(Category.GetCategory(uiDoc, BuiltInCategory.OST_PlumbingFixtures).Name);
                lista.Add(Category.GetCategory(uiDoc, BuiltInCategory.OST_MechanicalEquipment).Name);


                lista.Add(Category.GetCategory(uiDoc, BuiltInCategory.OST_ElectricalEquipment).Name);
                lista.Add(Category.GetCategory(uiDoc, BuiltInCategory.OST_CommunicationDevices).Name);
                lista.Add(Category.GetCategory(uiDoc, BuiltInCategory.OST_DataDevices).Name);
                lista.Add( Category.GetCategory(uiDoc, BuiltInCategory.OST_GenericModel).Name);
                lista.Add(Category.GetCategory(uiDoc, BuiltInCategory.OST_FireAlarmDevices).Name) ;
                //lista.Add(Category.GetCategory(uiDoc, BuiltInCategory.OST_Fixtures).Name);
                lista.Add(Category.GetCategory(uiDoc, BuiltInCategory.OST_LightingDevices).Name);
                lista.Add(Category.GetCategory(uiDoc, BuiltInCategory.OST_NurseCallDevices).Name);
                lista.Add(Category.GetCategory(uiDoc, BuiltInCategory.OST_SecurityDevices).Name);
                lista.Add(Category.GetCategory(uiDoc, BuiltInCategory.OST_PipeAccessory).Name);
                lista.Add(Category.GetCategory(uiDoc, BuiltInCategory.OST_DuctAccessory).Name);
                lista.Add(Category.GetCategory(uiDoc, BuiltInCategory.OST_TelephoneDevices).Name);
                 lista.Add(Category.GetCategory(uiDoc, BuiltInCategory.OST_SecurityDevices).Name);
                if(lista.Contains(item.Category.Name))
                {
                    return FuncaoExtrair.DadosEquipamentoMecanico((item as FamilyInstance));
                }

            }
            catch (Exception e)
            {
                return e.Message;
            }
            return null;
        }
    }


    public static class FuncaoExtrair
    {
        public static double tamanho;
        public static string passo;
        public static double d;
        public static double area;
        public static double z1 = 0;
        public static double z2 = 0;
        public static int i1 = 0;
        public static double area1;
        public static double area2;



        public static string DadosTubulacao(Element ele)
        {
            Autodesk.Revit.DB.Plumbing.Pipe p;

            string nomeServico = "";
            string unidServico = "";
            string diametro = "";
            unidServico = "m";
            diametro = GetNominalDiameter(ele);
            if (ele is Autodesk.Revit.DB.Plumbing.FlexPipe)
            {
                try
                {
                    nomeServico = (ele as Autodesk.Revit.DB.Plumbing.FlexPipe).FlexPipeType.LookupParameter("Descrição").AsString();
                    if (!string.IsNullOrEmpty(nomeServico))
                        return nomeServico + " - " + diametro;
                    else return null;
                }
                catch
                {
                    nomeServico = ele.LookupParameter("Segmento de tubulação").AsString();
                    if (!string.IsNullOrEmpty(nomeServico))
                        return nomeServico + " - " + diametro;
                    else return null;

                }
            }
            if (ele is Autodesk.Revit.DB.Plumbing.Pipe)
            {



                var pardescricao = (ele as Autodesk.Revit.DB.Plumbing.Pipe).PipeType.get_Parameter(BuiltInParameter.ALL_MODEL_DESCRIPTION);
                if (pardescricao != null)
                    if (pardescricao.HasValue)
                        if (!string.IsNullOrEmpty(pardescricao.AsString()))
                        {
                            nomeServico = pardescricao.AsString();

                            if(nomeServico.ToUpper().Contains("GALVANIZADO"))
                            {
                                return nomeServico + " - " + getNominalDiameterPolegada(ele);
                            }


                            return nomeServico + " - " + diametro;

                        }

                try
                {
                    nomeServico = (ele as Autodesk.Revit.DB.Plumbing.Pipe).PipeType.LookupParameter("MOL_DESCRIÇÃO").AsString();
                    if (!string.IsNullOrEmpty(nomeServico))
                        return nomeServico + " - " + diametro;
                    else return null;
                }
                catch
                {
                    nomeServico = ele.LookupParameter("Segmento de tubulação").AsValueString();
                    if (!string.IsNullOrEmpty(nomeServico))
                        return nomeServico + " - " + diametro;
                    else return null;

                }
            }
            if (ele is Autodesk.Revit.DB.Electrical.Conduit)
            {
                string nome = (ele as Autodesk.Revit.DB.Electrical.Conduit).Name;
                Autodesk.Revit.DB.Electrical.ConduitType tipoConduite =
                     Util.FindElementByName(typeof(Autodesk.Revit.DB.Electrical.ConduitType), nome) as Autodesk.Revit.DB.Electrical.ConduitType;

                var parMolDescricao = tipoConduite.LookupParameter("MOL_DESCRIÇÃO");
                if (parMolDescricao != null)
                    if (parMolDescricao.HasValue)
                        if (!string.IsNullOrEmpty(parMolDescricao.AsString()))
                        {
                            nomeServico = parMolDescricao.AsString();
                            return nomeServico + " - " +
                                              Math.Round((ele).LookupParameter("Diâmetro (tamanho comercial)").AsDouble() * 0.3048 * 1000) + "mm";
                        }

                var parDescricao = tipoConduite.get_Parameter(BuiltInParameter.ALL_MODEL_DESCRIPTION);
                if (parDescricao != null)
                    if (parDescricao.HasValue)
                        if (!string.IsNullOrEmpty(parDescricao.AsString()))
                        {
                            var galvanizado = tipoConduite.LookupParameter("Galvanizado");
                            if (galvanizado != null)
                            {
                                if (galvanizado.HasValue)
                                {
                                    if (galvanizado.AsInteger() == 0)
                                    {
                                        nomeServico = parDescricao.AsString();
                                        return nomeServico + " - " +
                                                          Math.Round((ele).LookupParameter("Diâmetro (tamanho comercial)").AsDouble() * 0.3048 * 1000) + "mm";
                                    }
                                    else
                                    {
                                        nomeServico = parDescricao.AsString();
                                        return nomeServico + " - " + FuncaoExtrair.getNominalDiameterPolegada((ele).LookupParameter("Diâmetro (tamanho comercial)"));
                                    }

                                }
                                else
                                {
                                    nomeServico = parDescricao.AsString();
                                    return nomeServico + " - " +
                                                      Math.Round((ele).LookupParameter("Diâmetro (tamanho comercial)").AsDouble() * 0.3048 * 1000) + "mm";
                                }
                            }
                            else
                            {
                                nomeServico = parDescricao.AsString();
                                return nomeServico + " - " +
                                                  Math.Round((ele).LookupParameter("Diâmetro (tamanho comercial)").AsDouble() * 0.3048 * 1000) + "mm";
                            }
                            

                            nomeServico = parDescricao.AsString();
                            return nomeServico + " - " +
                                              Math.Round((ele).LookupParameter("Diâmetro (tamanho comercial)").AsDouble() * 0.3048 * 1000) + "mm";
                        }

                nomeServico = ele.Name;
                return nomeServico + " - " + Math.Round((ele).LookupParameter("Diâmetro (tamanho comercial)").AsDouble() * 0.3048 * 1000) + "mm";

            }



            return "erro no diametro";
        }

        private static string getNominalDiameterPolegada(Element ele)
        {
            var param = ele.LookupParameter("Diâmetro");

            if (Math.Round(param.AsDouble() * 0.3048 * 1000) == 50)
            {
                return "2\"";

            }
            if (Math.Round(param.AsDouble() * 0.3048 * 1000) == 40)
            {
                return "1 1/2\"";

            }

            if (Math.Round(param.AsDouble() * 0.3048 * 1000) == 32)
            {
                return "1 1/4\"";

            }

            if (Math.Round(param.AsDouble() * 0.3048 * 1000) == 25)
            {
                return "1\"";

            }

            if (Math.Round(param.AsDouble() * 0.3048 * 1000) == 20)
            {
                return "3/4\"";

            }


            if (Math.Round(param.AsDouble() * 0.3048 * 1000) == 65)
            {
                return "2 1/2\"";

            }


            if (Math.Round(param.AsDouble() * 0.3048 * 1000) == 80)
            {
                return "3\"";

            }
            if (Math.Round(param.AsDouble() * 0.3048 * 1000) == 100)
            {
                return "4\"";

            }
            return "erro";

        }

        private static string getNominalDiameterPolegada(double tamanho)
        {
            var param = tamanho;

            if (Math.Round(param) == 50)
            {
                return "2\"";

            }
            if (Math.Round(param ) == 40)
            {
                return "1 1/2\"";

            }

            if (Math.Round(param) == 32)
            {
                return "1 1/4\"";

            }

            if (Math.Round(param ) == 25)
            {
                return "1\"";

            }

            if (Math.Round(param ) == 20)
            {
                return "3/4\"";

            }


            if (Math.Round(param ) == 65)
            {
                return "2 1/2\"";

            }


            if (Math.Round(param) == 80)
            {
                return "3\"";

            }
            if (Math.Round(param ) == 100)
            {
                return "4\"";

            }
            return "erro";

        }

        private static string getNominalDiameterPolegada(Autodesk.Revit.DB.Parameter param)
        {
            
            
                
            if (Math.Round(param.AsDouble() * 0.3048 * 1000) == 50)
            {
                return "2\"";

            }
            if (Math.Round(param.AsDouble() * 0.3048 * 1000) == 40)
            {
                return "1 1/2\"";

            }

            if (Math.Round(param.AsDouble() * 0.3048 * 1000) == 32)
            {
                return "1 1/4\"";

            }

            if (Math.Round(param.AsDouble() * 0.3048 * 1000) == 25)
            {
                return "1\"";

            }

            if (Math.Round(param.AsDouble() * 0.3048 * 1000) == 20)
            {
                return "3/4\"";

            }


            if (Math.Round(param.AsDouble() * 0.3048 * 1000) == 65)
            {
                return "2 1/2\"";

            }


            if (Math.Round(param.AsDouble() * 0.3048 * 1000) == 80)
            {
                return "3\"";

            }
            if (Math.Round(param.AsDouble() * 0.3048 * 1000) == 100)
            {
                return "4\"";

            }
            return "erro";

        }

        public static string DadosBandejaDeCabos(Autodesk.Revit.DB.Electrical.CableTray bandeja)
        {
            string nome = bandeja.Name;
            string nomeServico = "";
            Autodesk.Revit.DB.Electrical.CableTrayType tipoBandeja = Util.FindElementByName(typeof(Autodesk.Revit.DB.Electrical.CableTrayType), nome) as Autodesk.Revit.DB.Electrical.CableTrayType;
            string altura = Math.Round(bandeja.get_Parameter(BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM).AsDouble() * 0.3048 * 1000) + "mm";
            string largura = Math.Round(bandeja.get_Parameter(BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM).AsDouble() * 0.3048 * 1000) + "mm";
            string tamanho = largura + "X" + altura;

            var parMolDescricao = tipoBandeja.LookupParameter("MOL_DESCRIÇÃO");
            if (parMolDescricao != null)
                if (parMolDescricao.HasValue)
                    if (!string.IsNullOrEmpty(parMolDescricao.AsString()))
                    {
                        nomeServico = parMolDescricao.AsString();
                        return nomeServico + " - " + tamanho;
                    }

            var parDescricao = tipoBandeja.get_Parameter(BuiltInParameter.ALL_MODEL_DESCRIPTION);
            if (parDescricao != null)
                if (parDescricao.HasValue)
                    if (!string.IsNullOrEmpty(parDescricao.AsString()))
                    {
                        nomeServico = parDescricao.AsString();
                        return nomeServico + " - " + tamanho;
                    }

            return bandeja.Name + " - " + tamanho;


        }

        public static string DadosDuto(Duct duto)
        {
            

            string nomeDoFamilia = duto.DuctType.Name;
            string altura = "";
            string largura = "";
            string chapa = "";
            double pesoDachapa = 0;
            ObjetoTransferenciaPOB.DadosChapa dadoChapa = new ObjetoTransferenciaPOB.DadosChapa();
            string tipoDeFamilia = duto.DuctType.FamilyName;

            var diametro = duto.LookupParameter("Diâmetro");
            if(diametro!=null)
            {

            }
            else
            {
              
        
                altura = (duto.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM).AsDouble()*0.3048*1000).ToString();
                largura = (duto.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM).AsDouble() * 0.3048 * 1000).ToString();
               dadoChapa = Util.GetChapa(duto, CriarMenu.ListaDeChapas);

            }

            var areaDachapa = duto.get_Parameter(BuiltInParameter.RBS_CURVE_SURFACE_AREA);

            duto.LookupParameter(Properties.Settings.Default.GrossWeigth).Set(areaDachapa.AsDouble() * dadoChapa.DensidadeArea * 0.09290304);
            duto.LookupParameter("Tipo de chapa").Set(dadoChapa.Bitola);
            duto.LookupParameter("Peso específico da chapa").Set(dadoChapa.DensidadeArea *0.09290304);

            return nomeDoFamilia + " em " + duto.DuctType.LookupParameter("Material chapa").AsString() + " " + largura + "X" + altura + "mm";


        /*    try
            {
                nomeServico = duto.LookupParameter("MOL_DESCRIÇÃO").AsString();
                if (!string.IsNullOrEmpty(nomeServico))
                    return nomeServico;
                else return null;
            }
            catch
            {
                DuctType ductType = duto.DuctType;
                nomeServico = ductType.LookupParameter("MOL DESCRIÇÃO").AsString();
                if (!string.IsNullOrEmpty(nomeServico))
                    return nomeServico;
                else return null;

            }*/
        }
        public static string DadosConexaoBandejaDeCabos(FamilyInstance conexao)
        {
            string altura;
            string largura;
            string tamanho;
            var parAltura = conexao.LookupParameter("Altura");
            if (parAltura == null)
                parAltura = conexao.LookupParameter("ALT");

            var parLargura = conexao.LookupParameter("Largura");
            if (parLargura == null)
                parLargura = conexao.LookupParameter("Largura 1");
            if (parLargura == null)
                parLargura = conexao.LookupParameter("Largura 2");

            var listaRetangula = new List<SecaoDimensoes>();
            var listaSecaoCircular = new List<SecaoDimensoes>();

            foreach (Autodesk.Revit.DB.Connector cnn in conexao.MEPModel.ConnectorManager.Connectors)
            {
                
                
                switch (cnn.Shape)
                {
                    case ConnectorProfileType.Invalid:
                        
                        break;
                    case ConnectorProfileType.Round:
                        var secao = new SecaoDimensoes();
                        secao.Diametro = cnn.Radius * .3048 * 1000 * 2;
                        if (!listaSecaoCircular.Contains(secao))
                            listaSecaoCircular.Add(secao);

                        break;
                    case ConnectorProfileType.Rectangular:
                        var secao1 = new SecaoDimensoes();
                        secao1.Altura = cnn.Height * .3048 * 1000;
                        secao1.Largura = cnn.Width * .3048 * 1000;
                        if (!listaRetangula.Contains(secao1))
                            listaRetangula.Add(secao1);
                        break;
                    case ConnectorProfileType.Oval:
                       
                        break;
                    default:
                       
                        break;
                }

            }
            var conjuntoConector = "";
            listaRetangula = listaRetangula.OrderByDescending(x=>x.Largura).ToList();
            foreach (var s in listaRetangula)
            {
                conjuntoConector = conjuntoConector + Math.Round(s.Largura) + "mm" + "X"+ Math.Round(s.Altura)+"mm" + " - ";
            }
            foreach (var s in listaSecaoCircular)
            {
                conjuntoConector = conjuntoConector + " com saída para eletroduto de " + FuncaoExtrair.getNominalDiameterPolegada(Math.Round(s.Diametro)) + "mm";
            }
            //altura = ((parAltura != null) && (parAltura.HasValue)) ? Math.Round(parAltura.AsDouble() * 0.3048 * 1000) + "mm" : "Não informado";
            //largura = ((parLargura != null) && (parLargura.HasValue)) ? Math.Round(parLargura.AsDouble() * 0.3048 * 1000) + "mm" : "Não informado";


            var parNomeParametro = conexao.LookupParameter("ParametroUilizarNome");//, ParameterType.Text,
                                                                                                                                    
            if (parNomeParametro != null)
                if (parNomeParametro.HasValue)
                    if (!string.IsNullOrEmpty(parNomeParametro.AsString()))
                    {
                        var ParametroParaUtilizar = conexao.LookupParameter(parNomeParametro.AsString());
                        if (ParametroParaUtilizar != null)
                            return ParametroParaUtilizar.AsString();
                        ParametroParaUtilizar = conexao.Symbol.LookupParameter(parNomeParametro.AsString());
                        if (ParametroParaUtilizar != null)
                            return ParametroParaUtilizar.AsString();

                    }

            var parDescricao = conexao.Symbol.get_Parameter(BuiltInParameter.ALL_MODEL_DESCRIPTION);
            if (parDescricao != null)
                if (!string.IsNullOrEmpty(parDescricao.AsString()))
                {
                    /*if ((altura == "Não informado") | (altura == "Não informado"))
                        return parDescricao.AsString();
                    return parDescricao.AsString() + " - " + largura + " x " + altura;*/
                    return parDescricao.AsString() + " - " + conjuntoConector;
                }
            return conexao.Name + " - " + conjuntoConector;

        }
        public static string DadosConcexaoTubo(Element ele)
        {
            StringBuilder sb = new StringBuilder();
            string nomeServico = "";
            string unidServico = "";
            string diametro = "";
            string diametro1 = "";
            double valorDiametro;
            double valorAngulo;
            double valorDiametro1;
            double valorAnguloGuardado;
            string angulo = "";
            string juncao = "JUNÇÃO";
            string textoDiametro = "";
            string textoAngulo = "";
            unidServico = "Unid";
           
            var parNomeParametro = ele.LookupParameter("ParametroUilizarNome");
            if (parNomeParametro != null)
                if (parNomeParametro.HasValue)
                    if (!string.IsNullOrEmpty(parNomeParametro.AsString()))
                    {
                        var ParametroParaUtilizar = ele.LookupParameter(parNomeParametro.AsString());
                        if (ParametroParaUtilizar != null)
                            return ParametroParaUtilizar.AsString();
                        ParametroParaUtilizar = (ele as FamilyInstance).Symbol.LookupParameter(parNomeParametro.AsString());
                        if (ParametroParaUtilizar != null)
                            return ParametroParaUtilizar.AsString();

                    }
            var parINCDescricao = (ele as FamilyInstance).Symbol.LookupParameter("INC Descrição geral");
            if (parINCDescricao != null)
                if (parINCDescricao.HasValue)
                    if (!string.IsNullOrEmpty(parINCDescricao.AsString()))
                        return parINCDescricao.AsString();
            var parMolDescricao = (ele as FamilyInstance).Symbol.LookupParameter("MOL_DESCRIÇÃO");
            if (parMolDescricao != null)
                if (parMolDescricao.HasValue)
                    if (!string.IsNullOrEmpty(parMolDescricao.AsString()))
                        return parMolDescricao.AsString();
            /*var parTigreDescricao = ele.LookupParameter("Tigre: Descrição");
            if (parTigreDescricao != null)
                if (parTigreDescricao.HasValue)
                    if (!string.IsNullOrEmpty(parTigreDescricao.AsString()))
                        return parTigreDescricao.AsString();
            */
            //nomeServico = (ele as FamilyInstance).Symbol.LookupParameter("Modelo").AsString();
            nomeServico = ele.Name;// (ele as FamilyInstance).naSymbol.LookupParameter("Modelo").AsString();
            var tabelaConversao = new List<(double Menor, double Maior, string Texto)>
            {
                (49, 54, "2\""),
                (55, 62, "2 1/4\""),
                (62.05, 68, "2 1/2\""),
                (69, 78, "3\""),
                (79, 105, "4\"")
            };
            var tamanhos = ele.LookupParameter("Tamanho").AsString();
            var lista = tamanhos.Split('-').ToList();
            var listaDouble = lista.Select(item => item.Replace(" mmø", "")) // Remove "mm"
                               .Distinct() // Remove duplicatas
                               .Select(double.Parse) // Converte para double
                               .OrderByDescending(x => x) // Ordena os valores
                               .ToList();
            var listaPolegadas = listaDouble.Select(valor =>tabelaConversao.FirstOrDefault(intervalo => valor >= intervalo.Menor && 
                                                                                                                valor <= intervalo.Maior).Texto ?? "todo").ToList();

            string diametros = string.Join(" X ", listaPolegadas.Select(y => y.ToString()));
            
            /*diametro = GetNominalDiameter(ele);
            diametro1 = GetNominalDiameter1(ele);
            valorDiametro = GetNominalDiameterValor(ele);
            valorDiametro1 = GetNominalDiameterValor1(ele);*/
            angulo = GetAngulo(ele);

            valorAngulo = GetAnguloValue(ele);

            if (valorAngulo > 90)
                valorAngulo = 180 - Math.Round(GetAnguloValue(ele), 0);

            else
                valorAngulo = Math.Round(GetAnguloValue(ele), 0);

            valorAnguloGuardado = valorAngulo;
            if (valorAnguloGuardado >= 75 && valorAnguloGuardado <= 91) valorAnguloGuardado = 90;
            if (valorAnguloGuardado >= 10 && valorAnguloGuardado <= 12) valorAnguloGuardado = 11;
            if (valorAnguloGuardado >= 20 && valorAnguloGuardado <= 28) valorAnguloGuardado = 22;
            if (valorAnguloGuardado >= 29 && valorAnguloGuardado <= 74) valorAnguloGuardado = 45;


            textoAngulo = Math.Round(valorAnguloGuardado, 0).ToString() + "º";


            /*if (diametro != null)
            {
                textoDiametro = diametro;
            }
            if (!string.IsNullOrEmpty(diametro1))
            {
                if (valorDiametro1 >= valorDiametro)
                {
                    textoDiametro = diametro1 + " X " + diametro;
                }
                else
                {
                    textoDiametro = diametro + " X " + diametro1;
                }
            }*/

            nomeServico = nomeServico + " - " + textoDiametro;
            if (GetAnguloValue(ele) != 0)
                nomeServico = nomeServico + " - " + textoAngulo;
            return nomeServico;
        }

        public static string DadosPecaTubo(Element ele)
        {
            string nomeServico = "";
            string unidServico = "";
            string diametro = "";
            unidServico = "Unid";

            var parMolDescricao = (ele as FamilyInstance).Symbol.LookupParameter("MOL_DESCRIÇÃO");
            if (parMolDescricao != null)
                if (parMolDescricao.HasValue)
                    if (!string.IsNullOrEmpty(parMolDescricao.AsString()))
                        return parMolDescricao.AsString();
            var parTigreDescricao = (ele as FamilyInstance).Symbol.LookupParameter("Tigre: Descrição");
            if (parTigreDescricao != null)
                if (parTigreDescricao.HasValue)
                    if (!string.IsNullOrEmpty(parTigreDescricao.AsString()))
                        return parTigreDescricao.AsString();


            var parDescricao = (ele as FamilyInstance).Symbol.get_Parameter(BuiltInParameter.ALL_MODEL_DESCRIPTION);
            if (parDescricao != null)
                if (parDescricao.HasValue)
                    if (!string.IsNullOrEmpty(parDescricao.AsString()))
                        return parDescricao.AsString();

            diametro = GetNominalDiameter(ele);
            return (ele as FamilyInstance).Symbol.FamilyName + " - " + nomeServico + " - " + diametro;


        }
        public static string DadosEquipamentoMecanico(FamilyInstance ele)
        {


            string nomeServico = "";

            var parNomeParametro = ele.LookupParameter("ParametroUilizarNome");
            if (parNomeParametro != null)
                if (parNomeParametro.HasValue)
                    if (!string.IsNullOrEmpty(parNomeParametro.AsString()))
                    {
                        var ParametroParaUtilizar = ele.LookupParameter(parNomeParametro.AsString());
                        if (ParametroParaUtilizar != null)
                            return ParametroParaUtilizar.AsString();
                        ParametroParaUtilizar = (ele as FamilyInstance).Symbol.LookupParameter(parNomeParametro.AsString());
                        if (ParametroParaUtilizar != null)
                            return ParametroParaUtilizar.AsString();

                    }


            var parMolDescricao = ele.Symbol.LookupParameter("MOL_DESCRIÇÃO");
            if (parMolDescricao != null)
                if (parMolDescricao.HasValue)
                    if (!string.IsNullOrEmpty(parMolDescricao.AsString()))
                        return parMolDescricao.AsString();
            var parTigreDescricao = ele.Symbol.LookupParameter("Tigre: Descrição");
            if (parTigreDescricao != null)
                if (parTigreDescricao.HasValue)
                    if (!string.IsNullOrEmpty(parTigreDescricao.AsString()))
                        return parTigreDescricao.AsString();
            var parDescricaoUsuario = ele.Symbol.LookupParameter("ParametroUilizarNome");
            if (parDescricaoUsuario != null)
                if (parDescricaoUsuario.HasValue)
                    if (!string.IsNullOrEmpty(parDescricaoUsuario.AsString()))
                        return parDescricaoUsuario.AsString();

            var parDescricao = ele.Symbol.get_Parameter(BuiltInParameter.ALL_MODEL_DESCRIPTION);
            if (parDescricao != null)
                if (parDescricao.HasValue)
                    if (!string.IsNullOrEmpty(parDescricao.AsString()))
                        return parDescricao.AsString();


            nomeServico = ele.Symbol.FamilyName + " - " + ele.Name;
            if (!string.IsNullOrEmpty(nomeServico))
                return nomeServico;
            else return null;
        }



        public static double GetVolume(Element ele)
        {

            try
            {

                return ele.LookupParameter("Volume"/*Properties.Settings.Default.Volume*/).AsDouble() * 0.3048 * 0.3048 * 0.3048;
            }
            catch
            {
                foreach (Parameter param in ele.Parameters)
                {

                    if (param.Definition.Name == "Volume")
                        d = param.AsDouble() * 0.3048 * 0.3048 * 0.3048;
                    continue;
                }
                return d;
            }
        }
        public static string GetNominalDiameter(Element ele)
        {

            foreach (Parameter param in ele.Parameters)
            {

                if ((param.Definition.Name == "Diâmetro nominal 1") ||
                    (param.Definition.Name == "Nominal Diameter") ||
                    (param.Definition.Name == "Nominal Diameter 1") ||
                     (param.Definition.Name == "Diâmetro nominal") ||
                       (param.Definition.Name == "Diâmetro Nominal 1") ||
                    (param.Definition.Name == "Diâmetro") ||
                    (param.Definition.Name == "Diâmetro Nominal 1"))
                {
                    try
                    {
                        if (Math.Round(param.AsDouble() * 0.3048 * 100, 5) == 0.63500)
                        {
                            return "1/4\"";

                        }
                        if (Math.Round(param.AsDouble() * 0.3048 * 100, 5) == 0.95250)
                        {
                            return "3/8\"";

                        }
                        if (Math.Round(param.AsDouble() * 0.3048 * 100, 5) == 1.27000)
                        {
                            return "1/2\"";

                        }
                        if (Math.Round(param.AsDouble() * 0.3048 * 100, 5) == 2.54000)
                        {
                            return "1\"";

                        }
                        if (Math.Round(param.AsDouble() * 0.3048 * 100, 5) == 3.17500)
                        {
                            return "1 1/4\"";

                        }
                        if (Math.Round(param.AsDouble() * 0.3048 * 100, 5) == 1.90500)
                        {
                            return "3/4\"";

                        }
                        return param.AsValueString();

                    }
                    catch
                    {
                        return "erro no diâmetro";
                    }
                }
            }
            return "";

        }
        public static double GetNominalDiameterValor(Element ele)
        {
            double valor5 = 0;
            foreach (Parameter param in ele.Parameters)
            {

                if ((param.Definition.Name == "Diâmetro nominal 1") ||
                    (param.Definition.Name == "Nominal Diameter") ||
                    (param.Definition.Name == "Nominal Diameter 1") ||
                      (param.Definition.Name == "Diâmetro Nominal 1") ||
                      (param.Definition.Name == "Diâmetro (tamanho comercial)") ||
                     (param.Definition.Name == "Diâmetro nominal") ||
                    (param.Definition.Name == "Diâmetro"))
                    valor5 = param.AsDouble();
                continue;
            }
            return valor5;

        }
        public static double GetNominalDiameterValor1(Element ele)
        {
            double valor5 = 0;
            foreach (Parameter param in ele.Parameters)
            {

                if ((param.Definition.Name == "Diâmetro nominal 2") ||
                      (param.Definition.Name == "Diâmetro nominal") ||
                      (param.Definition.Name == "Diâmetro Nominal 2") ||
                    (param.Definition.Name == "Nominal Diameter 2"))
                    valor5 = param.AsDouble();
                continue;
            }
            return valor5;

        }


        public static string GetNominalDiameter1(Element ele)
        {

            foreach (Parameter param in ele.Parameters)
            {

                if ((param.Definition.Name == "Diâmetro nominal 2") ||
                    (param.Definition.Name == "Diâmetro Nominal 2") ||
                      (param.Definition.Name == "Diâmetro nominal") ||
                    (param.Definition.Name == "Nominal Diameter 2"))
                {
                    try
                    {
                        if (Math.Round(param.AsDouble() * 0.3048 * 100, 5) == 0.63500)
                        {
                            return "1/4\"";

                        }
                        if (Math.Round(param.AsDouble() * 0.3048 * 100, 5) == 0.95250)
                        {
                            return "3/8\"";

                        }
                        if (Math.Round(param.AsDouble() * 0.3048 * 100, 5) == 1.27000)
                        {
                            return "1/2\"";

                        }
                        if (Math.Round(param.AsDouble() * 0.3048 * 100, 5) == 2.54000)
                        {
                            return "1\"";

                        }
                        if (Math.Round(param.AsDouble() * 0.3048 * 100, 5) == 3.17500)
                        {
                            return "1 1/4\"";

                        }
                        if (Math.Round(param.AsDouble() * 0.3048 * 100, 5) == 1.90500)
                        {
                            return "3/4\"";

                        }
                        return param.AsValueString();
                    }
                    catch
                    {
                        return "erro no diâmetro";
                    }
                }
            }
            return "";

        }
        public static string GetAngulo(Element ele)
        {
            string texto4 = "";
            foreach (Parameter param in ele.Parameters)
            {

                if ((param.Definition.Name == "Ângulo") ||
                    (param.Definition.Name == "Angle"))
                    texto4 = param.AsValueString();
                continue;
            }
            return texto4;

        }
        public static double GetAnguloValue(Element ele)
        {
            double valor4 = 0;
            foreach (Parameter param in ele.Parameters)
            {
                if ((param.Definition.Name == "Ângulo") ||
                    (param.Definition.Name == "Angle")||
                    (param.Definition.Name == "Angulo"))
                    valor4 = param.AsDouble() / 0.0174532925199433;
                continue;
            }
            return valor4;
        }

        internal static string DadosConexaoDuto(Element curvaDuto)
        {

            var fi = curvaDuto as FamilyInstance;
            var fs = fi.Symbol;
            var maiorDimensao = 0.000;
            Duct maiorDuto;
            List<Duct> dutosConectados = Util.GetDutoConectados(fi);
            Duct maiorLarguraDuto = dutosConectados.OrderByDescending(x => x.Width).First();
            Duct maiorAlturaDuto = dutosConectados.OrderByDescending(x => x.Height).First();
            double LarguraDuto = maiorLarguraDuto.Width;
            double AlturaDuto = maiorAlturaDuto.Height;
            if (AlturaDuto > LarguraDuto)
            {
                maiorDimensao = AlturaDuto;
                maiorDuto = maiorAlturaDuto;
            }
            else
            {
                maiorDuto = maiorAlturaDuto;
                maiorDimensao = LarguraDuto;
            }
            double areaSuperficie = Util.GetAreaTotal(fi);

            string nomeDoFamilia = fi.Name;
            string altura = "";
            string largura = "";
            string chapa = "";
            double pesoDachapa = 0;
            ObjetoTransferenciaPOB.DadosChapa dadoChapa = new ObjetoTransferenciaPOB.DadosChapa();
            string tipoDeFamilia = maiorDuto.DuctType.FamilyName;

            var diametro = maiorDuto.LookupParameter("Diâmetro");
            if (diametro != null)
            {

            }
            else
            {


                altura = (maiorDuto.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM).AsDouble() * 0.3048 * 1000).ToString();
                largura = (maiorDuto.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM).AsDouble() * 0.3048 * 1000).ToString();
                dadoChapa = Util.GetChapa(maiorDuto, CriarMenu.ListaDeChapas);

            }

            
            fi.LookupParameter(Properties.Settings.Default.GrossWeigth).Set(areaSuperficie * dadoChapa.DensidadeArea * 0.09290304);
            fi.LookupParameter("Tipo de chapa").Set(dadoChapa.Bitola);
            fi.LookupParameter("Peso específico da chapa").Set(dadoChapa.DensidadeArea * 0.09290304);

            return nomeDoFamilia + " em " + maiorDuto.DuctType.LookupParameter("Material chapa").AsString() + " " + largura + "X" + altura + "mm";

        }
    }
    public static class FuncaoExtrair2
    {
        public static double tamanho;
        public static string passo;
        public static double d;
        public static double area;
        public static double z1 = 0;
        public static double z2 = 0;
        public static int i1 = 0;
        public static double area1;
        public static double area2;



        public static string DadosTubulacao(Element ele)
        {
            Autodesk.Revit.DB.Plumbing.Pipe p;

            string nomeServico = "";
            string unidServico = "";
            string diametro = "";
            unidServico = "m";
            diametro = GetNominalDiameter(ele);
            if (ele is Autodesk.Revit.DB.Plumbing.FlexPipe)
            {
                try
                {
                    Parameter par = (ele as Autodesk.Revit.DB.Plumbing.FlexPipe).FlexPipeType.get_Parameter(BuiltInParameter.ALL_MODEL_DESCRIPTION);
                    nomeServico = par.AsString();
                    if (!string.IsNullOrEmpty(nomeServico))
                        return nomeServico + " - " + diametro;
                    else return null;
                }
                catch
                {
                    return "todo-flex";

                }
            }
            if (ele is Autodesk.Revit.DB.Plumbing.Pipe)
            {
                try
                {
                    nomeServico = (ele as Autodesk.Revit.DB.Plumbing.Pipe).PipeType.LookupParameter("MOL_DESCRIÇÃO").AsString();
                    if (!string.IsNullOrEmpty(nomeServico))
                        return nomeServico + " - " + diametro;
                    else return null;
                }
                catch
                {
                    nomeServico = ele.LookupParameter("Segmento de tubulação").AsValueString();
                    if (!string.IsNullOrEmpty(nomeServico))
                        return nomeServico + " - " + diametro;
                    else return null;

                }
            }
            if (ele is Autodesk.Revit.DB.Electrical.Conduit)
            {
                try
                {
                    string nome = (ele as Autodesk.Revit.DB.Electrical.Conduit).Name;
                    Autodesk.Revit.DB.Electrical.ConduitType tipoConduite = Util.FindElementByName(typeof(Autodesk.Revit.DB.Electrical.ConduitType), nome) as Autodesk.Revit.DB.Electrical.ConduitType;
                    nomeServico = tipoConduite.LookupParameter("MOL_DESCRIÇÃO").AsString();
                    if (!string.IsNullOrEmpty(nomeServico))
                        return nomeServico + " - " +
                                              Math.Round((ele).LookupParameter("Diâmetro (tamanho comercial)").AsDouble() * 0.3048 * 1000) + "mm";
                    else return null;
                }
                catch
                {
                    nomeServico = ele.LookupParameter("MOL_DESCRIÇÃO").AsString();
                    if (!string.IsNullOrEmpty(nomeServico))
                        return nomeServico;
                    else return null;

                }
            }
            return null;
        }

        public static string DadosBandejaDeCabos(Autodesk.Revit.DB.Electrical.CableTray bandeja)
        {
            string nome = bandeja.Name;
            string nomeServico = "";
            Autodesk.Revit.DB.Electrical.CableTrayType tipoConduite = Util.FindElementByName(typeof(Autodesk.Revit.DB.Electrical.CableTrayType), nome) as Autodesk.Revit.DB.Electrical.CableTrayType;
            try
            {
                nomeServico = tipoConduite.LookupParameter("MOL_DESCRIÇÃO").AsString();
                if (!string.IsNullOrEmpty(nomeServico))
                    return nomeServico;
                else return
                         null;
            }
            catch
            {
                try
                {
                    nomeServico = tipoConduite.LookupParameter("MOL DESCRIÇÃO").AsString();
                    if (!string.IsNullOrEmpty(nomeServico))
                        return nomeServico;
                    return null;
                }
                catch
                {
                    nomeServico = bandeja.LookupParameter("MOL_DESCRIÇÃO").AsString();
                    if (!string.IsNullOrEmpty(nomeServico))
                        return nomeServico;
                    return null;
                }
            }
        }

        public static string DadosDuto(Duct duto)
        {
            string nomeServico = "";
            try
            {
                nomeServico = duto.LookupParameter("MOL_DESCRIÇÃO").AsString();
                if (!string.IsNullOrEmpty(nomeServico))
                    return nomeServico;
                else return null;
            }
            catch
            {
                DuctType ductType = duto.DuctType;
                nomeServico = ductType.LookupParameter("MOL DESCRIÇÃO").AsString();
                if (!string.IsNullOrEmpty(nomeServico))
                    return nomeServico;
                else return null;

            }
        }
        public static string DadosConexaoBandejaDeCabos(FamilyInstance conexao)
        {
            string nomeServico = "";
            try
            {
                nomeServico = conexao.Symbol.LookupParameter("MOL_DESCRIÇÃO").AsString();
                if (!string.IsNullOrEmpty(nomeServico))
                    return nomeServico;
                else return null;
            }
            catch
            {
                nomeServico = conexao.Symbol.LookupParameter("MOL DESCRIÇÃO").AsString();
                if (!string.IsNullOrEmpty(nomeServico))
                    return nomeServico;
                else return null;

            }
        }
        public static string DadosConcexaoTubo(Element ele)
        {
            StringBuilder sb = new StringBuilder();
            string nomeServico = "";
            string unidServico = "";
            string diametro = "";
            string diametro1 = "";
            double valorDiametro;
            double valorAngulo;
            double valorDiametro1;
            double valorAnguloGuardado;
            string angulo = "";
            string juncao = "JUNÇÃO";
            string textoDiametro = "";
            string textoAngulo = "";
            unidServico = "Unid";


            /*try
            {
                nomeServico = (ele as FamilyInstance).Symbol.LookupParameter("to-do").AsString();
                if (string.IsNullOrEmpty(nomeServico))
                    return null;


            }
            catch
            {
                try
                {
                    nomeServico = (ele as FamilyInstance).Symbol.LookupParameter("to").AsString();
                    if (string.IsNullOrEmpty(nomeServico))
                        return null;
                }
                catch
                {
                    try

                    {
                        nomeServico = ele.LookupParameter("MOL DESCRIÇÃO").AsString();
                        if (!string.IsNullOrEmpty(nomeServico))
                            return nomeServico;
                        else return null;
                    }
                    catch
                    {*/

            Parameter tigreParametro = ele.LookupParameter("Tigre: Descrição");
            if (tigreParametro != null)
            {
                nomeServico = tigreParametro.AsString();
                if (!string.IsNullOrEmpty(nomeServico))
                    return nomeServico;
            }
            Parameter descricaoParametro = ele.get_Parameter(BuiltInParameter.ALL_MODEL_DESCRIPTION);
            if (descricaoParametro != null)
            {
                nomeServico = tigreParametro.AsString();
            }
            else return "todo-conexão";
            diametro = GetNominalDiameter(ele);
            diametro1 = GetNominalDiameter1(ele);
            valorDiametro = GetNominalDiameterValor(ele);
            valorDiametro1 = GetNominalDiameterValor1(ele);
            angulo = GetAngulo(ele);

            valorAngulo = GetAnguloValue(ele);

            if (valorAngulo > 90)
                valorAngulo = 180 - Math.Round(GetAnguloValue(ele), 0);

            else
                valorAngulo = Math.Round(GetAnguloValue(ele), 0);

            valorAnguloGuardado = valorAngulo;
            if (valorAnguloGuardado >= 75 && valorAnguloGuardado <= 91) valorAnguloGuardado = 90;
            if (valorAnguloGuardado >= 10 && valorAnguloGuardado <= 12) valorAnguloGuardado = 11;
            if (valorAnguloGuardado >= 20 && valorAnguloGuardado <= 28) valorAnguloGuardado = 22;
            if (valorAnguloGuardado >= 29 && valorAnguloGuardado <= 74) valorAnguloGuardado = 45;


            textoAngulo = Math.Round(valorAnguloGuardado, 0).ToString() + "º";


            if (diametro != null)
            {
                textoDiametro = diametro;
            }
            if (!string.IsNullOrEmpty(diametro1))
            {
                if (valorDiametro1 >= valorDiametro)
                {
                    textoDiametro = diametro1 + " X " + diametro;
                }
                else
                {
                    textoDiametro = diametro + " X " + diametro1;
                }
            }

            nomeServico = nomeServico + " - " + textoDiametro;
            if (GetAnguloValue(ele) != 0)
                nomeServico = nomeServico + " - " + textoAngulo;
            return nomeServico;
        }

        public static string DadosPecaTubo(Element ele)
        {
            string nomeServico = "";
            string unidServico = "";
            string diametro = "";
            unidServico = "Unid";
            try
            {
                nomeServico = ele.LookupParameter("to-do").AsString();
                if (string.IsNullOrEmpty(nomeServico))
                    return null;
            }
            catch
            {
                nomeServico = ele.LookupParameter("to-do").AsString();
                if (string.IsNullOrEmpty(nomeServico))
                    return null;
            }
            diametro = GetNominalDiameter(ele);
            Parameter par = (ele as FamilyInstance).Symbol.get_Parameter(BuiltInParameter.ALL_MODEL_DESCRIPTION);
            if (par != null)
                return par.AsString() + " - " + nomeServico + " - " + diametro;
            else return "todo-pecatubo";


        }
        public static double GetVolume(Element ele)
        {

            try
            {

                return ele.LookupParameter("Volume"/*Properties.Settings.Default.Volume*/).AsDouble() * 0.3048 * 0.3048 * 0.3048;
            }
            catch
            {
                foreach (Parameter param in ele.Parameters)
                {

                    if (param.Definition.Name == "Volume")
                        d = param.AsDouble() * 0.3048 * 0.3048 * 0.3048;
                    continue;
                }
                return d;
            }
        }
        public static string GetNominalDiameter(Element ele)
        {

            foreach (Parameter param in ele.Parameters)
            {

                if ((param.Definition.Name == "Diâmetro nominal 1") ||
                    (param.Definition.Name == "Nominal Diameter") ||
                    (param.Definition.Name == "Nominal Diameter 1") ||
                     (param.Definition.Name == "Diâmetro nominal") ||
                       (param.Definition.Name == "Diâmetro Nominal 1") ||
                    (param.Definition.Name == "Diâmetro") ||
                    (param.Definition.Name == "Diâmetro Nominal 1"))
                {
                    try
                    {
                        if (Math.Round(param.AsDouble() * 0.3048 * 100, 5) == 0.63500)
                        {
                            return "1/4\"";

                        }
                        if (Math.Round(param.AsDouble() * 0.3048 * 100, 5) == 0.95250)
                        {
                            return "3/8\"";

                        }
                        if (Math.Round(param.AsDouble() * 0.3048 * 100, 5) == 1.27000)
                        {
                            return "1/2\"";

                        }
                        if (Math.Round(param.AsDouble() * 0.3048 * 100, 5) == 2.54000)
                        {
                            return "1\"";

                        }
                        if (Math.Round(param.AsDouble() * 0.3048 * 100, 5) == 3.17500)
                        {
                            return "1 1/4\"";

                        }
                        if (Math.Round(param.AsDouble() * 0.3048 * 100, 5) == 1.90500)
                        {
                            return "3/4\"";

                        }
                        return param.AsValueString();

                    }
                    catch
                    {
                        return "erro no diâmetro";
                    }
                }
            }
            return "";

        }
        public static double GetNominalDiameterValor(Element ele)
        {
            double valor5 = 0;
            foreach (Parameter param in ele.Parameters)
            {

                if ((param.Definition.Name == "Diâmetro nominal 1") ||
                    (param.Definition.Name == "Nominal Diameter") ||
                    (param.Definition.Name == "Nominal Diameter 1") ||
                      (param.Definition.Name == "Diâmetro Nominal 1") ||
                      (param.Definition.Name == "Diâmetro (tamanho comercial)") ||
                     (param.Definition.Name == "Diâmetro nominal") ||
                    (param.Definition.Name == "Diâmetro"))
                    valor5 = param.AsDouble();
                continue;
            }
            return valor5;

        }
        public static double GetNominalDiameterValor1(Element ele)
        {
            double valor5 = 0;
            foreach (Parameter param in ele.Parameters)
            {

                if ((param.Definition.Name == "Diâmetro nominal 2") ||
                      (param.Definition.Name == "Diâmetro nominal") ||
                      (param.Definition.Name == "Diâmetro Nominal 2") ||
                    (param.Definition.Name == "Nominal Diameter 2"))
                    valor5 = param.AsDouble();
                continue;
            }
            return valor5;

        }


        public static string GetNominalDiameter1(Element ele)
        {

            foreach (Parameter param in ele.Parameters)
            {

                if ((param.Definition.Name == "Diâmetro nominal 2") ||
                    (param.Definition.Name == "Diâmetro Nominal 2") ||
                      (param.Definition.Name == "Diâmetro nominal") ||
                    (param.Definition.Name == "Nominal Diameter 2"))
                {
                    try
                    {
                        if (Math.Round(param.AsDouble() * 0.3048 * 100, 5) == 0.63500)
                        {
                            return "1/4\"";

                        }
                        if (Math.Round(param.AsDouble() * 0.3048 * 100, 5) == 0.95250)
                        {
                            return "3/8\"";

                        }
                        if (Math.Round(param.AsDouble() * 0.3048 * 100, 5) == 1.27000)
                        {
                            return "1/2\"";

                        }
                        if (Math.Round(param.AsDouble() * 0.3048 * 100, 5) == 2.54000)
                        {
                            return "1\"";

                        }
                        if (Math.Round(param.AsDouble() * 0.3048 * 100, 5) == 3.17500)
                        {
                            return "1 1/4\"";

                        }
                        if (Math.Round(param.AsDouble() * 0.3048 * 100, 5) == 1.90500)
                        {
                            return "3/4\"";

                        }
                        return param.AsValueString();
                    }
                    catch
                    {
                        return "erro no diâmetro";
                    }
                }
            }
            return "";

        }
        public static string GetAngulo(Element ele)
        {
            string texto4 = "";
            foreach (Parameter param in ele.Parameters)
            {

                if ((param.Definition.Name == "Ângulo") ||
                    (param.Definition.Name == "Angle"))
                    texto4 = param.AsValueString();
                continue;
            }
            return texto4;

        }
        public static double GetAnguloValue(Element ele)
        {
            double valor4 = 0;
            foreach (Parameter param in ele.Parameters)
            {
                if ((param.Definition.Name == "Ângulo") ||
                    (param.Definition.Name == "Angle"))
                    valor4 = param.AsDouble() / 0.0174532925199433;
                continue;
            }
            return valor4;
        }

    }
}