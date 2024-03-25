using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.ApplicationServices;
using POB.ObjetoTransferenciaPOB;
namespace POB
{
    public static class CriarTabelasUtil
    {
        public static ViewDrafting CriarVistas(Document uiDoc, string ambienteValor, string pavimentoValor, TipoRelatorio tipoRelatorio)
        {

            var fi = new FilteredElementCollector(uiDoc);
            ViewFamilyType viewFamilyType = fi
                                               .OfClass(typeof(ViewFamilyType))
                                               .Cast<ViewFamilyType>()
                                               .FirstOrDefault(q => q.ViewFamily == ViewFamily.Drafting);

            ViewDrafting viewExistente = GetVistaexistente(uiDoc, ambienteValor, pavimentoValor, tipoRelatorio);
            Transaction tvista;
            if (viewExistente != null)
            {
                tvista = new Transaction(uiDoc);
                tvista.Start("criar vista");

                var vv = fi.OfClass(typeof(ModelCurve)).WhereElementIsNotElementType().
                    Where(x => x.OwnerViewId == viewExistente.Id).ToList();
                var id = from Element e in vv
                         select new ElementId(e.Id.IntegerValue);

                uiDoc.Delete(id.ToList());
                id = null;
                vv = fi.OfClass(typeof(TextNote)).WhereElementIsNotElementType().
                                        Where(x => x.OwnerViewId == viewExistente.Id).ToList();
                id = from Element e in vv
                     select new ElementId(e.Id.IntegerValue);

                uiDoc.Delete(id.ToList());



                var f = tvista.Commit();

                return viewExistente;
            }

            tvista = new Transaction(uiDoc);
            tvista.Start("criar vista");

            ViewDrafting viewDrafting = ViewDrafting.Create(uiDoc, viewFamilyType.Id);
            viewDrafting.Name = GetNameDrafthingView(ambienteValor, pavimentoValor, tipoRelatorio);

            viewDrafting.Scale = 25;
            tvista.Commit();
            return viewDrafting;
        }

        public static ViewDrafting GetVistaexistente(Document uiDoc, string ambienteValor, string pavimentoValor, TipoRelatorio tipoRelatorio)
        { FilteredElementCollector fi = new FilteredElementCollector(uiDoc);
         return                 fi
                                   .OfClass(typeof(ViewDrafting))
                                   .Cast<ViewDrafting>()
                                   .FirstOrDefault(q => q.Name == GetNameDrafthingView(ambienteValor, pavimentoValor, tipoRelatorio));
        }

        public static string GetNameDrafthingView(string ambiente, string pavimento, TipoRelatorio tipoRelatorio)
        {
            return pavimento + " - " + ambiente + " - " + tipoRelatorio.ToString();
        }

        public static void GerarTabelaAF(Document uidoc, string ambiente, string valorAmbiente,
                                                                       string pavimento, string valorPavimento)
        {
            IEnumerable<ElementoPOB> vv = FiltrarDados(uidoc, ambiente, valorAmbiente, pavimento, valorPavimento);

            ViewDrafting vistaAF = CriarVistas(uidoc, valorAmbiente, valorPavimento, TipoRelatorio.AF);          
            GerarTabela(uidoc, ambiente, pavimento, TipoRelatorio.AF, vistaAF, FiltroARQRQF(vv));

        }

        public static void GerarTabelaEsgoto(Document uidoc, string ambiente, string valorAmbiente,
                                                                       string pavimento, string valorPavimento)
        {
            IEnumerable<ElementoPOB> vv = FiltrarDados(uidoc, ambiente, valorAmbiente, pavimento, valorPavimento);
           ViewDrafting vistaESG = CriarVistas(uidoc, valorAmbiente, valorPavimento, TipoRelatorio.ES);
            GerarTabela(uidoc, ambiente, pavimento, TipoRelatorio.ES, vistaESG, FiltroEsgtoVentilacao(vv));
         

        }

        private static IEnumerable<ElementoPOB> FiltrarDados(Document uidoc, string ambiente, string valorAmbiente, string pavimento, string valorPavimento)
        {
            IList<ElementFilter> b = new List<ElementFilter>();

            ElementId eleId = GetIdParametroProjeto(uidoc, ambiente);
            ParameterValueProvider provider = new ParameterValueProvider(eleId);
            FilterStringRuleEvaluator evaluator = new FilterStringEquals();
#if D23 || D24
            FilterRule rule = new FilterStringRule(provider, evaluator, valorAmbiente);
#else
FilterRule rule = new FilterStringRule(provider, evaluator, valorAmbiente, true);
#endif
            ElementParameterFilter filter = new ElementParameterFilter(rule);
            b.Add(filter);

            ElementId eleIdPavimento = GetIdParametroProjeto(uidoc, pavimento);
            provider = new ParameterValueProvider(eleIdPavimento);
            evaluator = new FilterStringEquals();

#if D23 || D24
            rule = new FilterStringRule(provider, evaluator, valorPavimento);
#else
  rule = new FilterStringRule(provider, evaluator, valorPavimento, true);
#endif          


            filter = new ElementParameterFilter(rule);
            b.Add(filter);

            /*provider = new ParameterValueProvider(new ElementId((int)classificacao));
            evaluator = new FilterStringEquals();
            rule = new FilterStringRule(provider, evaluator, valorClassificacao, true);
            filter = new ElementParameterFilter(rule);
            b.Add(filter);
            */


            var logicalOrFilter = new LogicalAndFilter(b);
            FilteredElementCollector fi = new FilteredElementCollector(uidoc);
            var consulta = fi.WherePasses(CriarMenu.DefinirFiltroRegisterUpdateDescricaoEelemento()).
                WherePasses(logicalOrFilter).ToElements();
            var vv = from Element ele in consulta
                     select new ElementoPOB
                     {
                         Item = "",
                         Descricao = GetDescicaoMaterial(ele),
                         Unid = GetUnidade(ele),
                         Qtde = GetQtde(ele),
                         TipoDeSistema = GetTipodeSistema(ele),
                         ClassificacaoSistema = GetClassificacaoSistema(ele)
                     };
            return vv;
        }

        private static string GetDescicaoMaterial(Element ele)
        { var par = ele.LookupParameter("DescricaoMaterial");
            if (par.HasValue)
            {
                if (!string.IsNullOrEmpty(par.AsString()))
                    return par.AsString();
                else return "Vazio- "+ele.Id.IntegerValue.ToString(); 
            }
            else return "Outro" + ele.Id.IntegerValue.ToString();
        }

        private static void GerarTabela(Document uiDoc, string ambiente, string pavimento,TipoRelatorio tipoRelatorio, ViewDrafting vista,
            List<ElementoPOB> lista)
        {
            
            TransactionGroup tg = new TransactionGroup(uiDoc, "gr");
            tg.Start();
        
            var alturaInicial = 0.00;
            var colunaInicial = 1;

            TextNote textNote1;
            TextNote textNote2;
            TextNote textNote3;
            double d = 0;
            // tvista.Commit();
            using (Transaction txt1 = new Transaction(uiDoc))
            {
                TextNoteType textNoteType = new FilteredElementCollector(uiDoc)
                                       .OfClass(typeof(TextNoteType))
                                       .Cast<TextNoteType>()
                                       .FirstOrDefault(x => x.Name == "3mm arial tabela");

                foreach (ElementoPOB item in lista.OrderBy(x => x.Descricao))
                {
                    txt1.Start("txt");
                    textNote1 = TextNote.Create(uiDoc, vista.Id, new XYZ(0, alturaInicial, 0), item.Descricao, textNoteType.Id);
                    textNote1.Width = 0.1243595 * 3.3;

                    textNote2 = TextNote.Create(uiDoc, vista.Id, new XYZ(3.3 / 0.3048, alturaInicial, 0), item.Unid, textNoteType.Id);
                    textNote2.Width = 0.30 / 6*0.70;
                    textNote2.HorizontalAlignment = HorizontalTextAlignment.Right;

                    textNote3 = TextNote.Create(uiDoc, vista.Id, new XYZ(3.65 / 0.3048, alturaInicial, 0), Math.Round(item.Qtde, 0).ToString(), textNoteType.Id);
                    textNote3.Width = 0.35 / 6;
                    textNote3.HorizontalAlignment = HorizontalTextAlignment.Right;
                    txt1.Commit();
                    d = textNote1.Height;
                    alturaInicial = alturaInicial - d * 7 * 3.95;
                }
                tg.Commit();

            }
            /*  txt1.Start("linha");
                       point1 = uiApp.Application.Create.NewXYZ(-0.02 / 0.3048, alturaInicial + 0.02 / 0.3048, 0);

                       point2 = uiApp.Application.Create.NewXYZ(3.5 / 0.3048, alturaInicial + 0.02 / 0.3048, 0);

                       line = Line.CreateBound(point1, point2);
                       //Create line


                       detailCurve = uiDoc.Create.NewDetailCurve(viewDrafting, line);
                       txt1.Commit();

                   }
               }*/
            //  Autodesk.Revit.Creation.Application creapp = uiApp.Application.Create;
            //Transaction tvista = new Transaction(uiDoc);
            //tvista.Start("criar vista");


            // XYZ point1 = uiApp.Application.Create.NewXYZ(x1, y1, z);

            // XYZ point2 = uiApp.Application.Create.NewXYZ(x2, y2, z);

            //Line line = Line.CreateBound(point1, point2);
            // var tipo = uiDoc.GetElement(line.GraphicsStyleId);
            //Create line


            /*DetailCurve detailCurve = uiDoc.Create.NewDetailCurve(viewDrafting, line);

            //  GraphicsStyle graphicsStyle = new FilteredElementCollector(uiDoc)
            //                      .OfClass(typeof(GraphicsStyle))
            //                    .Cast<GraphicsStyle>()
            //                  .FirstOrDefault(x=>x.Name == "Vermelho");
            //line.SetGraphicsStyleId(graphicsStyle.Id);

            TextNote textNote = TextNote.Create(uiDoc, viewDrafting.Id, new XYZ(0, 0, 0), "Deu certo", textNoteType.Id);
    */
        }

        private static List<ElementoPOB> FiltroARQRQF(IEnumerable<ElementoPOB> vv)
        {
            var vt = from ElementoPOB p1 in vv.ToList()
                     where p1.ClassificacaoSistema != "Ventilação" |
                               p1.ClassificacaoSistema != "Sanitário"
                     group p1 by new
                     {
                         p1.ClassificacaoSistema,
                         p1.Descricao,
                         p1.TipoDeSistema,
                         p1.Unid
                     } into f
                     let Total = f.Sum(x => x.Qtde)
                     select new ElementoPOB
                     {
                         ClassificacaoSistema = f.Key.ClassificacaoSistema,
                         Descricao = f.Key.Descricao,
                         Qtde = Total,
                         TipoDeSistema = f.Key.TipoDeSistema,
                         Unid = f.Key.Unid
                     };
            return vt.ToList();
        }
        private static List<ElementoPOB> FiltroEsgtoVentilacao(IEnumerable<ElementoPOB> vv)
        {
            var vt = from ElementoPOB p1 in vv.ToList()
                     where p1.ClassificacaoSistema == "Ventilação" |
                               p1.ClassificacaoSistema == "Sanitário"
                     group p1 by new
                     {
                         p1.ClassificacaoSistema,
                         p1.Descricao,
                         p1.TipoDeSistema,
                         p1.Unid
                     } into f
                     let Total = f.Sum(x => x.Qtde)
                     select new ElementoPOB
                     {
                         ClassificacaoSistema = f.Key.ClassificacaoSistema,
                         Descricao = f.Key.Descricao,
                         Qtde = Total,
                         TipoDeSistema = f.Key.TipoDeSistema,
                         Unid = f.Key.Unid
                     };
            return vt.ToList();
        }


        private static string GetClassificacaoSistema(Element ele)
        {
            Parameter par = ele.get_Parameter(BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM);
            if (par.HasValue)
            {
                if (!string.IsNullOrEmpty(par.AsString()))
                    return par.AsString();
                else return "Vazio";
            }
            else return "Outro";
        }

        private static string GetTipodeSistema(Element ele)
        {
            Parameter par = ele.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM);
            return GetStringGeral(par);
        }
        public static string GetAmbiente(Element ele)
        {
            Parameter par = ele.LookupParameter("AMBIENTE");
            return GetStringGeral(par);
        }
        public static string GetPavimento(Element ele)
        {
            Parameter par = ele.LookupParameter("PAVIMENTO");
            return GetStringGeral(par);
        }
        private static string GetStringGeral(Parameter par)
        {
            if (par.HasValue)
            {
                if (!string.IsNullOrEmpty(par.AsString()))
                    return par.AsString();
                else return "Vazio";
            }
            else return "Outro";
        }

        private static double GetQtde(Element ele)
        {
            if (ele is Pipe)
                return (ele as Pipe).get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble() * 0.3048;
            else return 1;
        }

        private static string GetUnidade(Element ele)
        {
            if (ele is Pipe)
                return "m";
            else return "unid";
        }

        private static ElementId GetIdParametroProjeto(Document uidoc, string nome1)
        {
            List<ElementId> vetor = new List<ElementId>();
            BindingMap bindingMap = uidoc.ParameterBindings;
            DefinitionBindingMapIterator it = bindingMap.ForwardIterator();
            it.Reset();
            
            while (it.MoveNext())
            {
                if (it.Key.Name == nome1)
                {
                 return (it.Key as InternalDefinition).Id;

                }
            }
            return null;
        }
    }
}