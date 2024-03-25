using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using wf = System.Windows.Forms;
using db=  Autodesk.Revit.DB;
using ui =  Autodesk.Revit.UI;
using Excel = Microsoft.Office.Interop.Excel;
using Autodesk.Revit;
using Autodesk.Revit.DB;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Autodesk.Revit.DB.Visual;

namespace POB.Apresentacao
{
    public partial class FrmDadosImper : wf.Form
    {
        ui.ExternalCommandData _revit;
        db.Document  _uidoc;
        ui.Selection.Selection _sel;
      bool _criarTransacao;
        bool _abrindo;
        private List<ParametroGlobal> grid = new List<ParametroGlobal>();

        public bool Continuar { get; private set; }

        public FrmDadosImper()
        {
            InitializeComponent();
        }
        public FrmDadosImper(ui.ExternalCommandData revit,  bool criarTransacao)
        {
            InitializeComponent();
            _abrindo = true;
            _revit = revit;
            _uidoc = revit.Application.ActiveUIDocument.Document;
            _sel = revit.Application.ActiveUIDocument.Selection;
            _criarTransacao = criarTransacao;
            dataGridView1.DataSource = Util.ObterLevelsComNomeEId(_uidoc);
          
            /*var lst = NegocioRevit.NivelExtraidoCommad.Execute(_uidoc,
                 _sel.GetElementIds(),
                 criarTransacao,
                 Util.ObterLevels(_uidoc).Cast<db.Level>().ToList()).Lista.OrderBy(x => (x.Element as db.Level).Elevation).FirstOrDefault();
                 */
            _abrindo = false;
        }
        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void FrmDadosImper_Load(object sender, EventArgs e)
        {

        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            db.TransactionGroup transactionGroup = new db.TransactionGroup(_uidoc);
            transactionGroup.Start("Teste");
            var assembly = Util.CreateAssembly(_uidoc, _sel.GetElementIds());
            if (assembly == null)
            {
                transactionGroup.RollBack();
                return;
            }
            db.Transaction transaction = new db.Transaction(_uidoc);
            transaction.Start("inicio");

#if D23 || D24

            var partocAmbiente = Util.GetParameter(assembly, "tocAmbiente", SpecTypeId.String.Text, true, false);
#else
             var partocAmbiente = Util.GetParameter(assembly, "tocAmbiente",  db.ParameterType.Text, true, false);
#endif
            if (partocAmbiente != null) partocAmbiente.Set(txtAmbiente.Text);

            var parameterComentario = assembly.get_Parameter(db.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
            if (parameterComentario != null) parameterComentario.Set(txtPavimento.Text);

            //var parInclinacao = Util.GetParameter(assembly, "tocInclinacao", db.ParameterType.Slope, true, false);
           // if (parInclinacao != null) parInclinacao.Set(Convert.ToDouble(txtInclinacao.Text)/100);

           /* var partocCodigoImper = Util.GetParameter(assembly, "tocCodigoImper", db.ParameterType.Text, true, false);
            if (partocCodigoImper != null) partocCodigoImper.Set(txtCodigoImper.Text.ToUpper());
            var v = new List<db.ElementId>();
            v.Add(assembly.Id);
            _revit.Application.ActiveUIDocument.Selection.SetElementIds(v);
            NegocioRevit.DadosImperCommand.Execute(_revit, false);
            transaction.Commit();
            transactionGroup.Commit();
            Continuar = true;*/
            //this.Close();

        }

        private void dataGridView1_CursorChanged(object sender, EventArgs e)
        {
         
        }

        private void dataGridView1_Scroll(object sender, wf.ScrollEventArgs e)
        {
            
        }

        private void dataGridView1_RowEnter(object sender, wf.DataGridViewCellEventArgs e)
        {
           
           if(!_abrindo) txtPavimento.Text = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Continuar = true;
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            NegocioRevit.DadosImperCommand.Execute(_revit, true);
            this.Continuar = true;
            this.Close();
        }

        private void btnChecar_Click(object sender, EventArgs e)
        {
            NegocioRevit.ChecaDadosImperCommand.Execute(_revit);
            this.Continuar = true;
            this.Close();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
         
        }


        private void button2_Click(object sender, EventArgs e)
        {
            db.Transaction transaction = new db.Transaction(_uidoc);
            if (_criarTransacao)


                transaction.Start("inicio");


            foreach (var item in grid)
            {
                if (!string.IsNullOrEmpty(item.Acrescentar))
                {
                    var ele = _uidoc.GetElement(new db.ElementId(item.ElementoId)) as db.GlobalParameter;
                    db.ParameterValue valor = new db.StringParameterValue(item.Valor + "|" + item.Acrescentar);
                    ele.SetValue(valor);
                }
                else
                {
                    var ele = _uidoc.GetElement(new db.ElementId(item.ElementoId)) as db.GlobalParameter;
                    db.ParameterValue valor = new db.StringParameterValue(item.Valor);
                    ele.SetValue(valor);
                }
            }
            if (_criarTransacao)
            {
                transaction.Commit();
            }
            Continuar = true;
            this.Close();
        }

        private void button1_Click_2(object sender, EventArgs e)
        {
            var lista = db.GlobalParametersManager.GetAllGlobalParameters(_uidoc);
            List<db.GlobalParameter> lista1 = new List<db.GlobalParameter>();
            foreach (var item in lista)
            {
                lista1.Add(_uidoc.GetElement(item) as db.GlobalParameter);
            }

            foreach (var a in lista1)
            {

                if (a.GetValue() is db.StringParameterValue)
                {
                    var s = (a.GetValue() as db.StringParameterValue).Value.Split('|');
                    var valor = s[0];
                    var acrescentar = "";
                    if (s.Count() == 2) acrescentar = s[1];
                    grid.Add(new ParametroGlobal
                    {
                        Nome = a.Name,
                        Valor = valor,
                        Acrescentar = acrescentar,
                        ElementoId = a.Id.IntegerValue
                    }); 
                }

            }
            wf.BindingSource bs = new wf.BindingSource();
            bs.DataSource = grid.OrderBy(x => x.Nome).ToList(); ;
            dataGridView2.DataSource = bs;
            dataGridView2.ReadOnly = false;
            dataGridView2.AutoSizeRowsMode = wf.DataGridViewAutoSizeRowsMode.AllCells;
            dataGridView2.Columns[1].DefaultCellStyle.WrapMode = wf.DataGridViewTriState.True;
           
        }
        public void CriaTituloColuna(Excel.Worksheet novaPlanilha, int linhaExcel, bool junta)
        {
            novaPlanilha.Cells[linhaExcel, 2] = "Serviço";
            novaPlanilha.Cells[linhaExcel, 3] = "Materiais";
            novaPlanilha.Cells[linhaExcel, 4] = "Consumo";
            novaPlanilha.Cells[linhaExcel, 5] = "Área Horiz. (m²) ";
            novaPlanilha.Cells[linhaExcel, 6] = "Área Vert. (m²) h = 30cm e 2,10m";
            novaPlanilha.Cells[linhaExcel, 7] = "Área Total(m²)";
            novaPlanilha.Cells[linhaExcel, 8] = "Repetições";
            novaPlanilha.Cells[linhaExcel, 9] = "Repetições";
            novaPlanilha.Cells[linhaExcel, 10] = "Repetições";
            novaPlanilha.Cells[linhaExcel, 11] = "Área Total(m²)";
            novaPlanilha.Cells[linhaExcel, 12] = "Custo";
            novaPlanilha.Cells[linhaExcel, 13] = "Custo";
            Excel.Range rangeServico = novaPlanilha.Range[novaPlanilha.Cells[linhaExcel, 2], novaPlanilha.Cells[linhaExcel + 1, 2]];
            rangeServico.Merge();
            Excel.Range rangeMaterial = novaPlanilha.Range[novaPlanilha.Cells[linhaExcel, 3], novaPlanilha.Cells[linhaExcel + 1, 3]];
            rangeMaterial.Merge();
            Excel.Range rangeConsumo = novaPlanilha.Range[novaPlanilha.Cells[linhaExcel, 4], novaPlanilha.Cells[linhaExcel + 1, 4]];
            rangeConsumo.Merge();

            linhaExcel++;

            novaPlanilha.Cells[linhaExcel, 5] = "do local";
            novaPlanilha.Cells[linhaExcel, 6] = "do local";
            novaPlanilha.Cells[linhaExcel, 7] = "do local";
            novaPlanilha.Cells[linhaExcel, 8] = "no pavimento ";

            novaPlanilha.Cells[linhaExcel, 9] = "do pavimento";
            novaPlanilha.Cells[linhaExcel, 10] = "de torres";
            novaPlanilha.Cells[linhaExcel, 11] = "do empreendimento";
            novaPlanilha.Cells[linhaExcel, 12] = "Unitário";
            novaPlanilha.Cells[linhaExcel, 13] = "Total do local";
            novaPlanilha.Range[novaPlanilha.Cells[linhaExcel - 1, 2], novaPlanilha.Cells[linhaExcel, 13]].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            novaPlanilha.Range[novaPlanilha.Cells[linhaExcel - 1, 2], novaPlanilha.Cells[linhaExcel, 13]].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            (novaPlanilha.Range[novaPlanilha.Cells[linhaExcel - 1, 2], novaPlanilha.Cells[linhaExcel, 13]]).WrapText = true;
        }
        private void btn01_Click(object sender, EventArgs e)
        {
            NegocioRevit.ExtractDataImper.Execute(_revit, true);
            var lista = NegocioRevit.ExtractDataImper.listaElemento;

            //Agrupar pavimentos
            //Agrupar por folha
            //Agrupar por Ambiente
            //Agrupar por sistema
            //dados

            var listaDePavimento = (from pavimento in lista
                                    group pavimento by new
                                    {
                                        pavimento.Pavimento
                                    } into f
                                    select new
                                    {
                                        pav = f.Key.Pavimento
                                    }).ToList();
            var listaDePavimentoFolha = (from pavimento in lista
                                    group pavimento by new
                                    {
                                        pavimento.Pavimento,
                                        pavimento.Folha
                                    } into f
                                    select new
                                    {
                                        pav = f.Key.Pavimento,
                                        folha = f.Key.Folha
                                    }).ToList().OrderBy(x=>x.pav).ThenBy(x=>x.folha).ToList();

            var listaDePavimentoFolhaSistema = (from pavimento in lista
                                                 group pavimento by new
                                                 {
                                                     pavimento.Pavimento,
                                                     pavimento.Folha,
                                                     pavimento.NomeSistema,
                                                     pavimento.CodigoImper
                                                 } into f
                                                 select new
                                                 {
                                                     pav = f.Key.Pavimento,
                                                     folha = f.Key.Folha,
                                                     nomeDoSistema = f.Key.NomeSistema,
                                                     codImper = f.Key.CodigoImper
                                                 }).ToList();

            var listaDePavimentoFolhaSistemaAmbiente = (from pavimento in lista
                                         group pavimento by new
                                         {
                                             pavimento.Pavimento,
                                             pavimento.Folha,
                                             pavimento.Ambiente,
                                             pavimento.NumFolhaReferencia,
                                             pavimento.NomeSistema,
                                             pavimento.CodigoImper
                                         } into f
                                         select new
                                         {
                                             pav = f.Key.Pavimento,
                                             folha = f.Key.Folha,
                                             ambiente  =f.Key.Ambiente,
                                             referencia = f.Key.NumFolhaReferencia,
                                             nomeDoSistema = f.Key.NomeSistema,
                                             codImper = f.Key.CodigoImper
                                         }).ToList();
            

            //cria um excel
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = true;
            Excel.Workbook workbook = excelApp.Workbooks.Add();
            foreach (var pav in listaDePavimento)
            {
                //cria uma aba
                Excel.Worksheet novaPlanilha = workbook.Sheets.Add();
                novaPlanilha.Name = pav.pav.ToUpper();
                var linhaExcel = 9;
                foreach (var folha in listaDePavimentoFolha.Where(x => x.pav == pav.pav).ToList())
                {
                    novaPlanilha.Cells[linhaExcel, 2] = folha.pav + " – " + folha.folha;
                    Excel.Range range = novaPlanilha.Range[novaPlanilha.Cells[linhaExcel, 2], novaPlanilha.Cells[linhaExcel, 13]];
                    range.Merge();
                    range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(128, 128, 128));
                    range.Font.Bold = true;
                    range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    range.WrapText = true;
                    linhaExcel++;
                    foreach (var sistema in listaDePavimentoFolhaSistema.Where(x =>(x.pav == folha.pav) && (x.folha == folha.folha)).ToList())
                    {
                        double totalizadorSistema = 0;
                        foreach (var ambiente in listaDePavimentoFolhaSistemaAmbiente.Where(x => (x.pav == folha.pav) && (x.folha == folha.folha)&&(x.nomeDoSistema==sistema.nomeDoSistema)).ToList())
                        {
                            
                            novaPlanilha.Cells[linhaExcel, 2] = ambiente.ambiente + " – REF. " + ambiente.referencia;
                            Excel.Range rangeAmbiente = novaPlanilha.Range[novaPlanilha.Cells[linhaExcel, 2], novaPlanilha.Cells[linhaExcel, 13]];
                            rangeAmbiente.Merge();
                            rangeAmbiente.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(191, 191, 191));
                            rangeAmbiente.Font.Bold = true;
                            rangeAmbiente.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                            linhaExcel++;
                            novaPlanilha.Cells[linhaExcel, 2] = "Sistema: " + ambiente.nomeDoSistema + " – código do sistema " + ambiente.codImper;
                            Excel.Range rangeSistema = novaPlanilha.Range[novaPlanilha.Cells[linhaExcel, 2], novaPlanilha.Cells[linhaExcel, 13]];
                            rangeSistema.Merge();
                            rangeSistema.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(191, 191, 191));
                            rangeSistema.Font.Bold = true;
                            rangeSistema.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                            rangeSistema.WrapText = true;
                            linhaExcel++;
                            this.CriaTituloColuna(novaPlanilha, linhaExcel, false);
                            linhaExcel++;
                            linhaExcel++;
                            var listaMaterial = lista.Where(x => (x.Pavimento == ambiente.pav) && (x.Folha == ambiente.folha) && (x.Ambiente == ambiente.ambiente) && (x.NumFolhaReferencia == ambiente.referencia)).OrderBy(x => x.CategoriaDoMaterial).ThenBy(x => x.NomeDoMaterial).ToList();
                            var listaMontagem = (from montagem in listaMaterial
                                                 group montagem by new
                                                 {
                                                     montagem.IdMontagem,
                                                     montagem.AreaTotal
                                                 } into f
                                                 select new
                                                 {
                                                     montagem = f.Key.IdMontagem,
                                                     areaTotal = f.Key.AreaTotal
                                                 }).ToList();
                            totalizadorSistema = totalizadorSistema + listaMontagem.Sum(x => x.areaTotal);
                            var qtdeDeMontagens = listaMontagem.Count();
                            double? AlturaMaximaParede = 0;
                            try
                            {
                                AlturaMaximaParede = listaMaterial.Where(x => (x.TipoDeElemento == "Parede") && (x.AlturaDaParade > 0)).Max(x => x.AlturaDaParade);
                            }
                            catch
                            {

                            }
                            double? AlturaMinimaParede = 0;
                            try
                            {
                                AlturaMinimaParede = listaMaterial.Where(x => (x.TipoDeElemento == "Parede") && (x.AlturaDaParade > 0)).Min(x => x.AlturaDaParade);
                            }
                            catch
                            {

                            }
                            var repeticoesDoPavimento = 1;
                            var repeticoesDaTorre = 1;

                            var repeticoesNoPavimento = qtdeDeMontagens;
                            var materialResumido = (from resumo in listaMaterial
                                                    group resumo by new
                                                    {
                                                        resumo.NomeDoMaterial,
                                                        resumo.CategoriaDoMaterial,
                                                        resumo.CodigoImper,
                                                        resumo.NomeSistema

                                                    }
                                                                        into f
                                                    let AreaVertical = f.Sum(x => x.AreaVertical)
                                                    let AreaHorizontal = f.Sum(x => x.AreaHorizontal)
                                                    select new
                                                    {
                                                        nomeDoMaterial = f.Key.NomeDoMaterial,
                                                        nomeDoSistema = f.Key.NomeSistema,
                                                        codigoImper = f.Key.CodigoImper,
                                                        categoriaDoMaterial = f.Key.CategoriaDoMaterial,
                                                        areaVertical = AreaVertical,
                                                        areaHorizontal = AreaHorizontal
                                                    }).ToList();
                            var linhaAntesDeIniciarOMaterial = linhaExcel;
                            foreach (var material in materialResumido)
                            {
                                novaPlanilha.Cells[linhaExcel, 2] = material.categoriaDoMaterial;
                                novaPlanilha.Cells[linhaExcel, 3] = material.nomeDoMaterial;
                                novaPlanilha.Cells[linhaExcel, 4] = "";
                                novaPlanilha.Cells[linhaExcel, 5] = (material.areaHorizontal) / repeticoesNoPavimento;
                                novaPlanilha.Cells[linhaExcel, 6] = (material.areaVertical) / repeticoesNoPavimento;
                                novaPlanilha.Cells[linhaExcel, 7] = (material.areaVertical + material.areaHorizontal) / repeticoesNoPavimento;

                                novaPlanilha.Cells[linhaExcel, 11] = material.areaVertical + material.areaHorizontal;
                                novaPlanilha.Cells[linhaExcel, 12] = 0;
                                novaPlanilha.Cells[linhaExcel, 13] = 0;
                                novaPlanilha.Cells[linhaExcel, 5].NumberFormat = "#,0.00";
                                novaPlanilha.Cells[linhaExcel, 6].NumberFormat = "#,0.00";
                                novaPlanilha.Cells[linhaExcel, 7].NumberFormat = "#,0.00";

                                novaPlanilha.Cells[linhaExcel, 11].NumberFormat = "#,0.00";
                                novaPlanilha.Cells[linhaExcel, 12].NumberFormat = "#,0.00";
                                novaPlanilha.Cells[linhaExcel, 13].NumberFormat = "#,0.00";
                                linhaExcel++;
                            }



                            novaPlanilha.Range[novaPlanilha.Cells[linhaAntesDeIniciarOMaterial, 2], novaPlanilha.Cells[linhaExcel, 13]].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                            novaPlanilha.Range[novaPlanilha.Cells[linhaAntesDeIniciarOMaterial, 2], novaPlanilha.Cells[linhaExcel, 13]].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                            (novaPlanilha.Range[novaPlanilha.Cells[linhaAntesDeIniciarOMaterial, 2], novaPlanilha.Cells[linhaExcel, 13]]).WrapText = true;
                            (novaPlanilha.Range[novaPlanilha.Cells[linhaAntesDeIniciarOMaterial, 8], novaPlanilha.Cells[linhaExcel-1, 8]]).Merge();
                            novaPlanilha.Cells[linhaAntesDeIniciarOMaterial, 8] = repeticoesNoPavimento;
                            (novaPlanilha.Range[novaPlanilha.Cells[linhaAntesDeIniciarOMaterial, 9], novaPlanilha.Cells[linhaExcel-1, 9]]).Merge();
                            novaPlanilha.Cells[linhaAntesDeIniciarOMaterial, 9] = repeticoesDoPavimento;
                            (novaPlanilha.Range[novaPlanilha.Cells[linhaAntesDeIniciarOMaterial, 10], novaPlanilha.Cells[linhaExcel-1, 10]]).Merge();
                            novaPlanilha.Cells[linhaAntesDeIniciarOMaterial, 10] = repeticoesDaTorre;

                        }

                        
                        var rangeTotalSistema = novaPlanilha.Range[novaPlanilha.Cells[linhaExcel, 2], novaPlanilha.Cells[linhaExcel, 7]];
                        rangeTotalSistema.Merge();
                        rangeTotalSistema.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(128, 128, 128));
                        rangeTotalSistema = novaPlanilha.Range[novaPlanilha.Cells[linhaExcel, 8], novaPlanilha.Cells[linhaExcel, 10]];
                        rangeTotalSistema.Merge();
                        rangeTotalSistema.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(180, 198, 231));
                        novaPlanilha.Cells[linhaExcel, 8] = "ÁREA TOTAL DO SISTEMA " + sistema.codImper;

                        novaPlanilha.Cells[linhaExcel, 11] = totalizadorSistema;
                        novaPlanilha.Cells[linhaExcel, 11].NumberFormat = "#,0.00";
                        novaPlanilha.Cells[linhaExcel, 11].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(180, 198, 231));
                        novaPlanilha.Cells[linhaExcel, 12] = "M2";
                        novaPlanilha.Cells[linhaExcel, 12].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(180, 198, 231));
                        novaPlanilha.Cells[linhaExcel, 13].Interior.Color= System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(128, 128, 128));
                        linhaExcel++;


                    }
                }
                var rangeTotal = novaPlanilha.Range[novaPlanilha.Cells[9, 2], novaPlanilha.Cells[linhaExcel, 13]];
                rangeTotal.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                rangeTotal.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                rangeTotal.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                rangeTotal.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                rangeTotal.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
                rangeTotal.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;


            }

            this.Continuar = true;
            this.Close();
        }
    }
    public class ParametroGlobal
    {
        public string Nome { get; internal set; }
        public string Valor { get; internal set; }
        public string Acrescentar { get; set; }
        public int ElementoId { get; set; }
    }
}
