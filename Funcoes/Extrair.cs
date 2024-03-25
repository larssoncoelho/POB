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

using Autodesk.Revit.ApplicationServices;


using Funcoes;
using System.Windows.Forms;

namespace POB
{

  [TransactionAttribute(TransactionMode.Manual)]
    [RegenerationAttribute(RegenerationOption.Manual)]

    public class ExtrairDados : IExternalCommand
    {
     /*  public double tamanho;
        public string passo;
        public DataTable dataTable = new DataTable("Elementos");
        public double d;
        public double area;
        public double z1 = 0;
        public double z2 = 0;
        public int i1 = 0;
        public double area1;
        public double area2;
        public List<Element> ListaDeElemento = new List<Element>();
        public Element ele;

        public void DadosPilar(Element ele)
        {

        }
        public DataRow DadosTubulacao(Element ele, DataTable dt)
        {
            string nomeServico = "";
            string unidServico = "";
            string diametro = "";
            unidServico = "m";
            nomeServico = (ele as Autodesk.Revit.DB.Plumbing.Pipe).Name;
            DataRow dri = dt.NewRow();
            GetNominalDiameter(ele, ref diametro);
            if (diametro != null) nomeServico = nomeServico + " - " + diametro;
            dri["Serviço"] = nomeServico;
            try
            {
                ele.LookupParameter(Properties.Settings.Default.L7InsumoVinculado).Set(nomeServico);
            }
            catch
            {

            }
            dri["Unid"] = unidServico;
            dri["Qtde"] = ele.LookupParameter(Properties.Settings.Default.Length).AsDouble() * 0.3043;
            dri["Deslocamento_cm"] = GetDescolamento(ele) * 100;
            dri["CategoryName"] = ele.Category.Name;
            dri["ElementId"] = ele.Id.IntegerValue.ToString();
            dri["SystemType"] = ele.LookupParameter(Properties.Settings.Default.systemType).AsValueString();
            dri["Marca"] = ele.LookupParameter(Properties.Settings.Default.Mark).AsString();
            dri["Preco"] = getPreco(ele);
            dri["roofTop"] = ele.LookupParameter("roofTop").AsValueString();
            dri["tocNumeroPrumada1"] = ele.LookupParameter("tocNumeroPrumada1").AsValueString();
            dri["Nome"] = ele.LookupParameter("tocNome").AsString();

            try
            {
                dri["Comentários"] = ele.LookupParameter(Properties.Settings.Default.Comments).AsString();
            }
            catch
            {

            }
            try
            {
                dri["FamilyName"] = (ele as FamilyInstance).Symbol.FamilyName;
            }
            catch
            {

            }
            return dri;
        }
        public DataRow DadosTubulacaoFlexivel(Element ele, DataTable dt)
        {
            string nomeServico = "";
            string unidServico = "";
            string diametro = "";
            unidServico = "m";
            nomeServico = (ele as Autodesk.Revit.DB.Plumbing.FlexPipe).Name;
            DataRow dri = dt.NewRow();
            GetNominalDiameter(ele, ref diametro);
            if (diametro != null) nomeServico = nomeServico + " - " + diametro;
            dri["Serviço"] = nomeServico;
            try
            {
                ele.LookupParameter(Properties.Settings.Default.L7InsumoVinculado).Set(nomeServico);
            }
            catch
            {

            }
            dri["Unid"] = unidServico;
            dri["Qtde"] = ele.LookupParameter(Properties.Settings.Default.Length).AsDouble() * 0.3043;
            dri["Deslocamento_cm"] = GetDescolamento(ele) * 100;
            dri["CategoryName"] = ele.Category.Name;
            dri["ElementId"] = ele.Id.IntegerValue.ToString();
            dri["SystemType"] = ele.LookupParameter(Properties.Settings.Default.systemType).AsValueString();
            dri["Marca"] = ele.LookupParameter(Properties.Settings.Default.Mark).AsString();
            dri["Preco"] = getPreco(ele);
            dri["roofTop"] = ele.LookupParameter("roofTop").AsValueString();
            dri["tocNumeroPrumada1"] = ele.LookupParameter("tocNumeroPrumada1").AsValueString();
            dri["Nome"] = ele.LookupParameter("tocNome").AsString();

            try
            {
                dri["Comentários"] = ele.LookupParameter(Properties.Settings.Default.Comments).AsString();
            }
            catch
            {

            }
            try
            {
                dri["FamilyName"] = (ele as FamilyInstance).Symbol.FamilyName;
            }
            catch
            {

            }
            return dri;
        }
        public DataRow DadosElemento(Element ele, DataTable dt)
        {

            DataRow dri = dt.NewRow();

            dri["Serviço"] = ele.Name;
            dri["Unid"] = "";


            dri["CategoryName"] = ele.Category.Name;
            dri["ElementId"] = ele.Id.IntegerValue.ToString();

            dri["Marca"] = ele.LookupParameter(Properties.Settings.Default.Mark).AsString();
            dri["Volume"] = -1;
            dri["Preco"] = getPreco(ele);

           
            return dri;
        }
        public DataRow DadosConcexaoTubo(Element ele, DataTable dt)
        {
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
            nomeServico = ele.Name;
            try
            {
                nomeServico = nomeServico + " - " + ele.LookupParameter("L7Referencia").AsString();
            }
            catch
            {

            }

            DataRow dri = dt.NewRow();
            GetNominalDiameter(ele, ref diametro);
            GetNominalDiameter1(ele, ref diametro1);
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

            dri["Serviço"] = nomeServico;
            dri["Serviço"] = (ele as FamilyInstance).Symbol.FamilyName + " " + nomeServico;
            dri["Unid"] = unidServico;
            dri["qtde"] = 1;
            dri["Deslocamento_cm"] = GetDescolamento(ele) * 100;
            dri["CategoryName"] = ele.Category.Name;
            dri["ElementId"] = ele.Id.IntegerValue.ToString();
            dri["SystemType"] = ele.LookupParameter(Properties.Settings.Default.systemType).AsValueString();
            dri["Marca"] = ele.LookupParameter(Properties.Settings.Default.Mark).AsString();
            dri["Diametro"] = diametro;
            dri["Ângulo"] = angulo;
            dri["ÂnguloValor"] = (Math.Round(GetAnguloValue(ele))).ToString();
            dri["Diametro1"] = diametro1;
            dri["Preco"] = getPreco(ele);


            dri["Nome"] = ele.LookupParameter("tocNome").AsString();

            try
            {
                dri["Serviço"] = ele.LookupParameter("Tigre: Descrição").AsString();
            }
            catch
            {

            }



            try
            {
                ele.LookupParameter(Properties.Settings.Default.L7InsumoVinculado).Set(dri["Serviço"].ToString());
            }
            catch
            {

            }




            return dri;
        }

        public DataRow DadosPecaTubo(Element ele, DataTable dt)
        {
            string nomeServico = "";
            string unidServico = "";
            string diametro = "";
            unidServico = "Unid";
            nomeServico = ele.Name;
            DataRow dri = dt.NewRow();
            GetNominalDiameter(ele, ref diametro);
            dri["Serviço"] = (ele as FamilyInstance).Symbol.FamilyName + " - " + nomeServico;
            try
            {
                ele.LookupParameter(Properties.Settings.Default.L7InsumoVinculado).Set((ele as FamilyInstance).Symbol.FamilyName + " - " + dri["Serviço"].ToString());
            }
            catch
            {

            }
            dri["Unid"] = unidServico;
            dri["Qtde"] = 1;
            dri["Deslocamento_cm"] = GetDescolamento(ele) * 100;
            dri["CategoryName"] = ele.Category.Name;
            dri["ElementId"] = ele.Id.IntegerValue.ToString();
            try
            {
                dri["SystemType"] = ele.LookupParameter(Properties.Settings.Default.systemType).AsValueString();
            }
            catch
            {

            }
            dri["Nome"] = ele.LookupParameter("tocNome").AsString();
            dri["Preco"] = getPreco(ele);

            try
            {
                dri["FamilyName"] = (ele as FamilyInstance).Symbol.FamilyName;
            }
            catch
            {

            }


            return dri;

        }
        public DataRow DadosEquipamentoMecanico(Element ele, DataTable dt)
        {
            string nomeServico = "";
            string unidServico = "";
            string diametro = "";
            Double qtde = 0;
            unidServico = "Unid";
            qtde = 1;
            try
            {
                unidServico = ele.LookupParameter(Properties.Settings.Default.L7Unidade).AsString();
                if (unidServico.ToUpper() == "M")
                {
                    qtde = ele.LookupParameter(Properties.Settings.Default.L7Comprimento).AsDouble() * 0.3048;
                }

            }
            catch
            {

            }
            nomeServico = ele.Name;
            DataRow dri = dt.NewRow();
            dri["Serviço"] = (ele as FamilyInstance).Symbol.FamilyName + " - " + nomeServico;
            try
            {
                ele.LookupParameter(Properties.Settings.Default.L7InsumoVinculado).Set(dri["Serviço"].ToString());
            }
            catch
            {
            }
            dri["Unid"] = unidServico;
            dri["Qtde"] = qtde;
            dri["Deslocamento_cm"] = GetDescolamento(ele) * 100;
            dri["CategoryName"] = ele.Category.Name;
            dri["ElementId"] = ele.Id.IntegerValue.ToString();



            try
            {
                dri["SystemType"] = ele.LookupParameter(Properties.Settings.Default.systemType).AsValueString();
            }
            catch
            {
                dri["SystemType"] = "Erro";
            }
            dri["Marca"] = ele.LookupParameter(Properties.Settings.Default.Mark).AsString();
            dri["Preco"] = getPreco(ele);
            dri["Nome"] = ele.LookupParameter("tocNome").AsString();
            ele.LookupParameter(Properties.Settings.Default.L7InsumoVinculado).Set(dri["Serviço"].ToString());
            try
            {
                dri["FamilyName"] = (ele as FamilyInstance).Symbol.FamilyName;
            }
            catch
            {

            }
            return dri;
        }
        public DataRow DadosViga(Element ele, DataTable dt)
        {
            string nomeServico = "";
            string unidServico = "";
            string diametro = "";
            unidServico = "m3";
            nomeServico = ele.Name;
            DataRow dri = dt.NewRow();
            dri["Serviço"] = ele.Name;
            dri["Unid"] = "m3";
            dri["Marca"] = ele.LookupParameter(Properties.Settings.Default.Mark).AsString();
            dri["CategoryName"] = ele.Category.Name;
            dri["ElementId"] = ele.Id.IntegerValue.ToString();
            dri["Preco"] = getPreco(ele);
            dri["Nome"] = ele.LookupParameter("tocNome").AsString();
            try
            {
                dri["FamilyName"] = (ele as FamilyInstance).Symbol.FamilyName;
            }
            catch
            {

            }


            return dri;
        }
        public DataRow DadosPeca(Element ele, DataTable dt)
        {
            string nomeServico = "";
            string unidServico = "";
            ElementId id = ele.LookupParameter(Properties.Settings.Default.Material).AsElementId();

            nomeServico = (ele.Document.GetElement(id) as Material).Name;
            DataRow dri = dt.NewRow();
            dri["Serviço"] = nomeServico;
            try
            {


                ele.LookupParameter("L7InsumoVinculado").Set(nomeServico);
            }
            catch
            {

            }
            dri["Unid"] = "indefinida";
            dri["Marca"] = ele.LookupParameter(Properties.Settings.Default.Mark).AsString();
            dri["CategoryName"] = ele.Category.Name;
            dri["ElementId"] = ele.Id.IntegerValue.ToString();
            dri["Preco"] = getPreco(ele);
            dri["Nome"] = ele.LookupParameter("tocNome").AsString();
            try
            {
                dri["FamilyName"] = (ele as FamilyInstance).Symbol.FamilyName;
            }
            catch
            {

            }


            return dri;
        }
        public DataRow DadosPilar(Element ele, DataTable dt)
        {
            string nomeServico = "";
            string unidServico = "";
            string diametro = "";
            unidServico = "Unid";
            nomeServico = ele.Name;
            DataRow dri = dt.NewRow();
            GetNominalDiameter(ele, ref diametro);
            dri["Serviço"] = (ele as FamilyInstance).Symbol.FamilyName + " - " + nomeServico;
            dri["Unid"] = "m3";
            dri["Marca"] = ele.LookupParameter(Properties.Settings.Default.Mark).AsString();
            dri["CategoryName"] = ele.Category.Name;
            dri["ElementId"] = ele.Id.IntegerValue.ToString();
            dri["Preco"] = getPreco(ele);
            dri["Nome"] = ele.LookupParameter("tocNome").AsString();
            try
            {
                dri["FamilyName"] = (ele as FamilyInstance).Symbol.FamilyName;
            }
            catch
            {

            }


            return dri;
        }
        public DataRow DadosLajes(Element ele, DataTable dt)
        {
            string nomeServico = "";
            string unidServico = "";
            string diametro = "";
            unidServico = "Unid";
            nomeServico = ele.Name;
            DataRow dri = dt.NewRow();

            dri["Serviço"] = ele.Name;
            dri["Unid"] = "m3";


            dri["Marca"] = ele.LookupParameter(Properties.Settings.Default.Mark).AsString();
            dri["Marca"] = ele.LookupParameter("tocNome").AsString();
            dri["CategoryName"] = ele.Category.Name;
            dri["ElementId"] = ele.Id.IntegerValue.ToString();
            dri["Preco"] = getPreco(ele);
            dri["Nome"] = ele.LookupParameter("tocNome").AsString();
            try
            {
                dri["FamilyName"] = (ele as FamilyInstance).Symbol.FamilyName;
            }
            catch
            {

            }


            return dri;
        }
        public DataRow DadosParede(Element ele, DataTable dt)
        {
            string nomeServico = "";
            string unidServico = "";
            string diametro = "";
            unidServico = "m2";
            nomeServico = ele.Name;
            DataRow dri = dt.NewRow();

            dri["Serviço"] = ele.Name;
            dri["Unid"] = "m2";


            dri["Marca"] = ele.LookupParameter(Properties.Settings.Default.Mark).AsString();
            dri["CategoryName"] = ele.Category.Name;
            dri["ElementId"] = ele.Id.IntegerValue.ToString();

            try
            {
                dri["FamilyName"] = (ele as FamilyInstance).Symbol.FamilyName;
            }
            catch
            {

            }


            return dri;
        }
        public void BuscarDados(Element ele)
        {


            if (ele.Category.Name == Properties.Settings.Default.walls)
            {
                dataTable.Rows.Add(DadosParede(ele, dataTable));
            }
            if (ele.Category.Name == Properties.Settings.Default.parts)
            {
                dataTable.Rows.Add(DadosPeca(ele, dataTable));
            }
            if (ele.Category.Name == Properties.Settings.Default.structuralColumns)
            {
                dataTable.Rows.Add(DadosPilar(ele, dataTable));
            }
            if (ele.Category.Name == Properties.Settings.Default.Columns)
            {
                dataTable.Rows.Add(DadosPilar(ele, dataTable));
            }
            if (ele.Category.Name == Properties.Settings.Default.structuralFraming)
            {
                dataTable.Rows.Add(DadosViga(ele, dataTable));
            }
            if (ele.Category.Name == Properties.Settings.Default.floors)
            {
                dataTable.Rows.Add(DadosLajes(ele, dataTable));
            }
            if (ele.Category.Name == Properties.Settings.Default.escadas)
            {

            }
            if (ele.Category.Name == Properties.Settings.Default.walls)  //"Walls":
            {
                if (ele is Autodesk.Revit.DB.Wall)
                {


                }
                if (ele is Autodesk.Revit.DB.FamilyInstance)
                {

                }
            }
            if (ele.Category.Name == Properties.Settings.Default.pipes)  // "Pipes":
                dataTable.Rows.Add(DadosTubulacao(ele, dataTable));

            if (ele.Category.Name == Properties.Settings.Default.flexPipes)  //"Flex Pipes":
                dataTable.Rows.Add(DadosTubulacaoFlexivel(ele, dataTable));


            if (ele.Category.Name == Properties.Settings.Default.electricalFixtures)
            {
                //to-do 
            }

            if (ele.Category.Name == Properties.Settings.Default.telhados)
            {
                //to-do  
            }
            if (ele.Category.Name == Properties.Settings.Default.pipeFittings)  //"Pipe Fittings":
                dataTable.Rows.Add(DadosConcexaoTubo(ele, dataTable));

            if (ele.Category.Name == Properties.Settings.Default.plumbingFixtures)  //"Plumbing Fixtures":
                dataTable.Rows.Add(DadosEquipamentoMecanico(ele, dataTable));
            if (ele.Category.Name == Properties.Settings.Default.pipeAccessories)  //"Pipe Accessories":
                dataTable.Rows.Add(DadosEquipamentoMecanico(ele, dataTable));
            if (ele.Category.Name == Properties.Settings.Default.mechanicalEquipament)  //"Pipe Accessories":
                dataTable.Rows.Add(DadosEquipamentoMecanico(ele, dataTable));
            if (ele.Category.Name == Properties.Settings.Default.Sprinklers)  //"Pipe Accessories":
                dataTable.Rows.Add(DadosEquipamentoMecanico(ele, dataTable));
            if (ele.Category.Name == Properties.Settings.Default.genericModels)  //"Pipe Accessories":
                dataTable.Rows.Add(DadosEquipamentoMecanico(ele, dataTable));

        }
        public void CriarColunas()
        {
            dataTable.Columns.Add("Serviço", typeof(string));
            dataTable.Columns.Add("Unid", typeof(string));
            dataTable.Columns.Add("Qtde", typeof(double));
            dataTable.Columns.Add("Deslocamento_cm", typeof(double));
            dataTable.Columns.Add("ElementId", typeof(int));
            dataTable.Columns.Add("Marca", typeof(string));
            dataTable.Columns.Add("FamilyType", typeof(string));
            dataTable.Columns.Add("FamilyName", typeof(string));
            dataTable.Columns.Add("CategoryName", typeof(string));
            dataTable.Columns.Add("Nome", typeof(string));
            dataTable.Columns.Add("PsaId", typeof(string));
            dataTable.Columns.Add("ServicoId", typeof(string));
            dataTable.Columns.Add("Material", typeof(string));
            dataTable.Columns.Add("StructuralMaterial", typeof(string));
            dataTable.Columns.Add("PercentExecutado", typeof(string));
            dataTable.Columns.Add("AreaBase", typeof(double));
            dataTable.Columns.Add("AreaTopo", typeof(double));
            dataTable.Columns.Add("Volume", typeof(double));
            dataTable.Columns.Add("SystemType", typeof(string));
            dataTable.Columns.Add("Comprimento", typeof(string));
            dataTable.Columns.Add("Diametro", typeof(string));
            dataTable.Columns.Add("Diametro1", typeof(string));
            dataTable.Columns.Add("AreaForma", typeof(string));
            dataTable.Columns.Add("Ângulo", typeof(string));
            dataTable.Columns.Add("ÂnguloValor", typeof(string));
            dataTable.Columns.Add("Preco", typeof(double));
            dataTable.Columns.Add(Properties.Settings.Default.Comments, typeof(string));
            dataTable.Columns.Add("roofTop", typeof(string));
            dataTable.Columns.Add("tocNumeroPrumada1", typeof(string));

        }
        public void GetAreaTopo(Element ele, ref ExternalCommandData revit, ref double area)
        {
            area = 0;
            UIApplication uiApp = revit.Application;
            Document uiDoc = uiApp.ActiveUIDocument.Document;

            Options opt = uiApp.Application.Create.NewGeometryOptions();

            GeometryElement ge12 = ele.get_Geometry(opt);


            try
            {

                foreach (GeometryObject obj in ge12)
                {

                    GeometryInstance gi = obj as GeometryInstance;

                    GeometryElement gh = gi.GetInstanceGeometry();

                    foreach (GeometryObject obj1 in gh)
                    {


                        Solid solid = obj1 as Solid;
                        if (null != solid)
                        {
                            foreach (Face face in solid.Faces)
                            {
                                PlanarFace pf = face as PlanarFace;
                                //pf.Visibility = Visibility.Invisible;

                                if (null != pf)
                                {
                                   
                                    if (pf.FaceNormal.Z == 1)
                                    {
                                        area1 = pf.Area * 0.3048 * 0.3048;
                                        z1 = pf.Origin.Z;

                                        if (i1 == 0)
                                        {
                                            area2 = area1;
                                            z2 = z1;
                                        }
                                        else
                                        {
                                            if (z1 > z2)
                                            {
                                                area2 = area1;
                                                z2 = z1;
                                            }
                                        }
                                        i1 = i1 + 1;
                                    }
                                }
                            }
                            area = area2;

                        }
                    }
                }
            }
            catch
            {

                area = -200;
            }
        }
        public void GetArea(Element ele, ref ExternalCommandData revit, ref double area)
        {
            area = 0;
            try
            {
                UIApplication uiApp = revit.Application;
                Document uiDoc = uiApp.ActiveUIDocument.Document;

                Options opt = uiApp.Application.Create.NewGeometryOptions();

                GeometryElement ge12 = ele.get_Geometry(opt);
                Material m = Util.FindElementByName(typeof(Material), "Completo") as Material;


                foreach (GeometryObject obj in ge12)

                {
                    GeometryInstance gi = obj as GeometryInstance;

                    GeometryElement gh = gi.GetInstanceGeometry();

                    foreach (GeometryObject obj1 in gh)
                    {

                        Solid solid = obj1 as Solid;
                        if (null != solid)
                        {
                            foreach (Face face in solid.Faces)
                            {
                                PlanarFace pf = face as PlanarFace;
                                if (null != pf)
                                {

                                    area1 = pf.Area * 0.3048 * 0.3048;
                                    z1 = pf.Origin.Z;

                                    if (i1 == 0)
                                    {
                                        area2 = area1;
                                        z2 = z1;
                                    }
                                    else
                                    {
                                        if (z1 < z2)
                                        {
                                            area2 = area1;
                                            z2 = z1;
                                        }
                                    }
                                    i1 = i1 + 1;
                                }
                            }
                            area = area2;

                        }
                    }
                }
            }
            catch (Exception ex)
            {
                area = -200;
            }


        }




        public DataRow NovaLinha(ElementId eleId, ref Document uiDoc,
            /*, ref Plan_servico_amoNegocio manipulacao,*/
       /*         ref ExternalCommandData revit)
        {
            Element ele = uiDoc.GetElement(eleId);
            DataRow dr = dataTable.NewRow();
            dr["ElementId"] = eleId.IntegerValue;
            try
            {
                dr["Marca"] = ele.LookupParameter(Properties.Settings.Default.Mark).AsString();
                dr["Marca"] = ele.LookupParameter("Nome").AsString();
            }

            catch (Exception ex6)
            {
                dr["Marca"] = ex6.Message;
            }
            try
            { dr["servicoId"] = ele.LookupParameter("tocServicoId").AsString(); }
            catch { dr["servicoId"] = "Erro ao obter o nome"; }
            return dr;

        }


        public double GetVolume(Element ele)
        {

            try
            {
                return ele.LookupParameter(Properties.Settings.Default.Volume).AsDouble() * 0.3048 * 0.3048 * 0.3048;
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
        public void GetNominalDiameter(Element ele, ref string texto5)
        {

            foreach (Parameter param in ele.Parameters)
            {

                if ((param.Definition.Name == "Diâmetro nominal 1") ||
                    (param.Definition.Name == "Nominal Diameter") ||
                    (param.Definition.Name == "Nominal Diameter 1") ||
                     (param.Definition.Name == "Diâmetro nominal") ||
                       (param.Definition.Name == "Diâmetro Nominal 1") ||
                    (param.Definition.Name == "Diâmetro"))
                {
                    try
                    {
                        if (Math.Round(param.AsDouble() * 0.3048 * 100, 5) == 0.63500)
                        {
                            texto5 = "1/4\"";
                            return;
                        }
                        if (Math.Round(param.AsDouble() * 0.3048 * 100, 5) == 0.95250)
                        {
                            texto5 = "3/8\"";
                            return;
                        }
                        if (Math.Round(param.AsDouble() * 0.3048 * 100, 5) == 1.27000)
                        {
                            texto5 = "1/2\"";
                            return;
                        }
                        if (Math.Round(param.AsDouble() * 0.3048 * 100, 5) == 2.54000)
                        {
                            texto5 = "1\"";
                            return;
                        }
                        if (Math.Round(param.AsDouble() * 0.3048 * 100, 5) == 3.17500)
                        {
                            texto5 = "1 1/4\"";
                            return;
                        }
                        if (Math.Round(param.AsDouble() * 0.3048 * 100, 5) == 1.90500)
                        {
                            texto5 = "3/4\"";
                            return;
                        }
                        texto5 = param.AsValueString();
                    }
                    catch
                    {
                        texto5 = "erro no diâmetro";
                    }
                }
            }

        }
        public double GetNominalDiameterValor(Element ele)
        {
            double valor5 = 0;
            foreach (Parameter param in ele.Parameters)
            {

                if ((param.Definition.Name == "Diâmetro nominal 1") ||
                    (param.Definition.Name == "Nominal Diameter") ||
                    (param.Definition.Name == "Nominal Diameter 1") ||
                      (param.Definition.Name == "Diâmetro Nominal 1") ||
                     (param.Definition.Name == "Diâmetro nominal") ||
                    (param.Definition.Name == "Diâmetro"))
                    valor5 = param.AsDouble();
                continue;
            }
            return valor5;

        }
        public double GetNominalDiameterValor1(Element ele)
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
        public double GetDescolamento(Element ele)
        {
            double valor2 = 0;
            foreach (Parameter param in ele.Parameters)
            {

                if ((param.Definition.Name == "Deslocamento") ||
                    (param.Definition.Name == "Deslocamento") ||
                    (param.Definition.Name == "Deslocamento"))
                    valor2 = param.AsDouble() * 0.3048;
                continue;
            }

            return valor2;

        }
        public double getPreco(Element ele)
        {
            double valor2 = 0;
            foreach (Parameter param in ele.Parameters)
            {

                if ((param.Definition.Name == "tocPreco"))
                    valor2 = param.AsDouble();
                continue;
            }

            return valor2;

        }
        public void GetNominalDiameter1(Element ele, ref string texto5)
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
                            texto5 = "1/4\"";
                            return;
                        }
                        if (Math.Round(param.AsDouble() * 0.3048 * 100, 5) == 0.95250)
                        {
                            texto5 = "3/8\"";
                            return;
                        }
                        if (Math.Round(param.AsDouble() * 0.3048 * 100, 5) == 1.27000)
                        {
                            texto5 = "1/2\"";
                            return;
                        }
                        if (Math.Round(param.AsDouble() * 0.3048 * 100, 5) == 2.54000)
                        {
                            texto5 = "1\"";
                            return;
                        }
                        if (Math.Round(param.AsDouble() * 0.3048 * 100, 5) == 3.17500)
                        {
                            texto5 = "1 1/4\"";
                            return;
                        }
                        if (Math.Round(param.AsDouble() * 0.3048 * 100, 5) == 1.90500)
                        {
                            texto5 = "3/4\"";
                            return;
                        }
                        texto5 = param.AsValueString();
                    }
                    catch
                    {
                        texto5 = "erro no diâmetro";
                    }
                }
            }

        }
        public string GetAngulo(Element ele)
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
        public double GetAnguloValue(Element ele)
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
                                ListaDeElemento.Add(aSubElem);
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
                                ListaDeElemento.Add(aSubElem1);
                                BuscarAninhados(aSubElem1);
                            }
                        }
                    }
                }
            }
        }*/
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData revit,
                  ref string message, ElementSet elements)
        {


          /*  UIApplication uiApp = revit.Application;
            Document uiDoc = uiApp.ActiveUIDocument.Document;
            Selection sel = uiApp.ActiveUIDocument.Selection;
            StringBuilder sb = new StringBuilder();
            List<ElementId> lista = new List<ElementId>();
            Util.uiDoc = uiDoc;
            ListaDeElemento.Clear();

            foreach (ElementId eleId in uiApp.ActiveUIDocument.Selection.GetElementIds())
            {
                ListaDeElemento.Add(uiDoc.GetElement(eleId));
                BuscarAninhados(uiDoc.GetElement(eleId));
            }




            StringBuilder sb1 = new StringBuilder();
            string linha = "";
            CriarColunas();
            
            Level ll;



            ProgressoFuncao progresso = new ProgressoFuncao();

            progresso.pgc.Maximum = sel.GetElementIds().Count;
            progresso.pgc.Minimum = 0;
            progresso.pgc.Value = 1;
            progresso.pgc.Step = 1;
            progresso.Show();
            foreach (Element eleA in ListaDeElemento)
            {
                progresso.pgc.PerformStep();

                BuscarDados(uiDoc.GetElement(eleA.Id));

            }
            foreach (DataColumn idc in dataTable.Columns)
            {
                linha = linha + idc.ColumnName + "\t";
            }
            sb1.Append("\n" + linha);
            foreach (DataRow idr in dataTable.Rows)
            {
                linha = "";
                foreach (DataColumn idc in dataTable.Columns)
                {
                    linha = linha + idr[idc.ColumnName].ToString() + "\t";
                }
                sb1.Append("\n" + linha);
            }
            System.Windows.Forms.Clipboard.Clear();
            System.Windows.Forms.Clipboard.SetText(sb1.ToString());
            progresso.Dispose();*/
            return Result.Succeeded;

        }
    }


}
