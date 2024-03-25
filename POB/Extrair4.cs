using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Mechanical;
    using Autodesk.Revit.DB.Electrical;
using Funcoes;

namespace POB
{
    public class Extrair4
    {
        public double tamanho;
        public string passo;
        public double d;
        public double area;
        public double z1 = 0;
        public double z2 = 0;
        public int i1 = 0;
        public double area1;
        public double area2;



        public string DadosTubulacao(Element ele)
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
                    else return "";
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

                    nomeServico = (ele as Autodesk.Revit.DB.Plumbing.Pipe).PipeType.Name;
                    if (!string.IsNullOrEmpty(nomeServico))
                        return nomeServico + " - " + diametro;
                    else return null;
                }
                catch
                {
                    nomeServico = ele.LookupParameter("Segmento de tubulação").AsValueString();
                    if (!string.IsNullOrEmpty(nomeServico))
                        return nomeServico + " - " + diametro;
                    else return "";

                }
            }
            return "";
          /*  if (ele is Autodesk.Revit.DB.Electrical.Conduit)
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
            return null;*/
        }

        public string Nome()
        {
            return "nome";
        }

        public string DadosBandejaDeCabos(Autodesk.Revit.DB.Electrical.CableTray bandeja)
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

        public string DadosDuto(Duct duto)
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
        public string DadosConexaoBandejaDeCabos(FamilyInstance conexao)
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
        public string DadosConcexaoTubo(Element ele)
        {
            //StringBuilder sb = new StringBuilder();
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
                if (descricaoParametro.HasValue)
                    nomeServico = descricaoParametro.AsString();
                else return "todo para descrição";
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

        public string DadosPecaTubo(Element ele)
        {
            string nomeServico = "";
            string unidServico = "";
            string diametro = "";
            unidServico = "Unid";
        /*    try
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
            }*/
            try
            {

                //diametro = GetNominalDiameter(ele);
                Parameter par = (ele as FamilyInstance).Symbol.get_Parameter(BuiltInParameter.ALL_MODEL_DESCRIPTION);
                if (par != null)
                    return par.AsString() + " - " + nomeServico;// + " - " + diametro;
                else return "todo-pecatubo";
            }
            catch (Exception e)
            {
                return e.Message;
            }

        }
        public double GetVolume(Element ele)
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
        public string GetNominalDiameter(Element ele)
        {
         /*   var listaParametro = from Parameter  par in ele.Parameters
                                 where par.Definition.Name.Contains
                                 */
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
        public double GetNominalDiameterValor(Element ele)
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


        public string GetNominalDiameter1(Element ele)
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

    }
}