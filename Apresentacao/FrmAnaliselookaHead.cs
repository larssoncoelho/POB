using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI.Selection;
using revitDB = Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using appRevit = Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Structure;
using AcessoBancoDados;
using Negocios;
using ObjetoTransferencia;
using Funcoes;

namespace Apresentacao
{
    public partial class FrmAnaliselookaHead : NForm
    {
        public Plan_servico_amoNegocio manipulacao;
        public string dir = "";
        public ExternalCommandData revit;
        private revitDB.ElementId ele1;
        private revitDB.Document uiDoc;
        private revitDB.Element ele;
        private UIApplication uiApp;
        private revitDB.FilteredElementCollector selecao;
        private List<revitDB.View> vistasDeAvanco = new List<revitDB.View>();
        private DataView dv;
        public PAR_VERIFICAR_EXECUCAO parSituacao = new PAR_VERIFICAR_EXECUCAO();
        public DateTime diaAnalise;
        public DateTime mesAnalise;

        public revitDB.OverrideGraphicSettings orgLookahead;
        public revitDB.OverrideGraphicSettings orgRestricao;
        private revitDB.ElementId preenchimentoId;
        private Boolean atualizando = false;
        List<revitDB.Element> Lista = new List<revitDB.Element>();
        DataTable dt = new DataTable();
        DataTable dtGrid = new DataTable("dtGrid");
        DataTable dtSEL_LOOKAHEAD_INSUMOS = new DataTable("dtSEL_LOOKAHEAD_INSUMOS");
        DataTable dtResumoGeral = new DataTable("dtResumoGeral");
        DataTable dtResumoGeralBloco = new DataTable("dtResumoGeralGrupo");
        public BindingSource bs = new BindingSource();
        public DataSet ds = new DataSet("Dataset");
        public PAR_PROJECAO_PSA_POR_DIA parametros = new PAR_PROJECAO_PSA_POR_DIA();
        public DataTable dtQtdePSA = new DataTable();
        public FrmAnaliselookaHead()
        {
            InitializeComponent();
        }
        public FrmAnaliselookaHead(string idir, ExternalCommandData irevit, DataTable dtDownload)
        {
            revit = irevit;
            InitializeComponent();
            ConstruirQtdePSA();
            dtpInicio.Value = DateTime.Today.AddDays(-30);
            dtpTermino.Value = DateTime.Today;
            manipulacao = new Plan_servico_amoNegocio(idir);
            dtGrid.Columns.Add("INSUMO_ID", typeof(int));
            dtGrid.Columns.Add("INSUMO", typeof(string));
            dtGrid.Columns.Add("UNID", typeof(string));
            dtGrid.Columns.Add("QTDE_TOTAL_ENTRADA", typeof(double));
            dtGrid.Columns.Add("QTDE_CONSUMIDA", typeof(double));
            dtGrid.Columns.Add("QTDE_SOLICITADA_LOOK_A_HEAD", typeof(double));
            dtGrid.Columns.Add("SALDO", typeof(double));
            dtGrid.Columns["SALDO"].Expression = "QTDE_TOTAL_ENTRADA - QTDE_CONSUMIDA - QTDE_SOLICITADA_LOOK_A_HEAD";

            dtGrid.Columns.Add("SERVICO_ID", typeof(string));

            dtGrid.Columns.Add("CONSUMO_INSUMO", typeof(double));
            dtGrid.Columns.Add("ORIGEM", typeof(int));

            DataGridViewCellStyle formatoNumerico = new DataGridViewCellStyle();
            formatoNumerico.Format = "N2";
            ds.Tables.Add(dtGrid);
            bs.DataSource = ds;
            bs.DataMember = dtGrid.TableName;
            dataGridView1.DataSource = bs;
            dataGridView1.Columns["INSUMO_ID"].Width = 90;
            dataGridView1.Columns["INSUMO_ID"].HeaderText = "Insumo";
            dataGridView1.Columns["INSUMO"].Width = 200;
            dataGridView1.Columns["INSUMO"].HeaderText = "Desc. Inusmo";

            dataGridView1.Columns["QTDE_TOTAL_ENTRADA"].Width = 120;
            dataGridView1.Columns["QTDE_TOTAL_ENTRADA"].HeaderText = "Entrada";
            dataGridView1.Columns["QTDE_TOTAL_ENTRADA"].DefaultCellStyle = formatoNumerico;

            dataGridView1.Columns["QTDE_CONSUMIDA"].Width = 120;
            dataGridView1.Columns["QTDE_CONSUMIDA"].HeaderText = "Saída";
            dataGridView1.Columns["QTDE_CONSUMIDA"].DefaultCellStyle = formatoNumerico;

            dataGridView1.Columns["QTDE_SOLICITADA_LOOK_A_HEAD"].Width = 120;
            dataGridView1.Columns["QTDE_SOLICITADA_LOOK_A_HEAD"].HeaderText = "Lookahead";
            dataGridView1.Columns["QTDE_SOLICITADA_LOOK_A_HEAD"].DefaultCellStyle = formatoNumerico;

            dataGridView1.Columns["SALDO"].Width = 120;
            dataGridView1.Columns["SALDO"].HeaderText = "Saldo";
            dataGridView1.Columns["SALDO"].DefaultCellStyle = formatoNumerico;


            FrmProcurar procurar = new FrmProcurar();
            StringBuilder sb = new StringBuilder();

            sb.Append("    select ");
            sb.Append("   obra_id,  ");
            sb.Append("   obra  ");
            sb.Append("   ");
            sb.Append(" from obra ");
            ResultadoProcura rp = new ResultadoProcura();
            rp = procurar.Pesquisar(idir, "Escolher Obra",
                                      sb.ToString(), "Obra;Desc. Obra;",
                                                    "80;250;");



            EscolherData escolherData1 = new EscolherData();
            DialogResult resultado1 = escolherData1.inputar(ref mesAnalise, "Escolha o mes de análise", ref continuar);


            EscolherData escolherData = new EscolherData();
            DialogResult resultado = escolherData.inputar(ref diaAnalise, "Escolha o dia", ref continuar);

            if (continuar)
            {
                if (procurar.resultadoProcura.fResultadoProcura)
                {
                    atualizando = true;
                    cmbServico.SelectedItem = 0;
                    foreach (DataRow dr2 in dtDownload.Rows)
                    {
                        cmbServico.Items.Add(dr2["tocServicoId"].ToString());
                    }
                    atualizando = false;
                    BuscarInsumoLookAHead(dtDownload, diaAnalise, Convert.ToInt32( rp.vCampo[0]));
                    BuscarQtdePSA(dtDownload);                  
                }
            }
            escolherData.Dispose();
            uiApp = revit.Application;
            uiDoc = uiApp.ActiveUIDocument.Document;
            Selection sel = uiApp.ActiveUIDocument.Selection;
            Util.uiDoc = uiDoc;
            selecao = new revitDB.FilteredElementCollector(uiDoc).OfClass(typeof(Autodesk.Revit.DB.View));

            foreach (revitDB.View view in selecao)
            {
                try
                {
                    if (view.AreGraphicsOverridesAllowed())
                        if (view.LookupParameter("tocVistaAvanco").AsValueString() == "Sim")
                            vistasDeAvanco.Add(view);
                }
                catch
                {

                }

            }

            preenchimentoId = (Util.FindElementByName(typeof(revitDB.Material), "Previsto") as revitDB.Material).SurfacePatternId;
            /* orgRestricao= Util.GetOverrideGraphicSettings(Util.GetColorRevit(CorLinhaResticao),
                                                          Util.GetColorRevit(CorSuperficieRestricao),
                                                          preenchimentoId, 0);*/
        }


        public void BuscarQtdePSA(DataTable dtDownload)
        {
            foreach (DataRow dr2 in dtDownload.Rows)
            {
                foreach (DataRow dr in manipulacao.PRC_EXECUTAR_DIRETO
                    ("  SELECT  *  FROM SEL_ITENS_PENDENTES(" + dr2["tocGrupoId"].ToString() + ","
                                                                                                   + dr2["tocServicoId"].ToString() + ")").Rows)
                {
                    DataRow dri = dtQtdePSA.NewRow();
                    dri["PSA_ID"] = Convert.ToInt32(dr["PSA_ID"]);
                    dri["UNIDADE_PRINCIPAL"] = dr["UNIDADE_PRINCIPAL"];
                    dri["QTDE_REALIZADA"] = dr["QTDE_REALIZADA"];
                    dri["QTDE_PROJECAO"] = dr["QTDE_PROJECAO"];
                    dri["QTDE"] = dr["QTDE"];
                    dri["QTDE_LOOKAHEAD_PENDENTE"] = dr["QTDE_LOOKAHEAD_PENDENTE"];
                    dtQtdePSA.Rows.Add(dri);
                }
            }
        }


        private void BuscarInsumoLookAHead(DataTable dtDownload, DateTime dia, int obraId)
        {
            foreach (DataRow dr2 in dtDownload.Rows)
            {
                dtSEL_LOOKAHEAD_INSUMOS = manipulacao.PRC_EXECUTAR_DIRETO
                    ("    select insumo_id, "+
                                   "  insumo, " +
                                   "   unid , " +
                                   "   obra_id, " +
                                    "  consumo_insumo, " +
                                    "  coalesce(qtde_total_entrada, 0) qtde_total_entrada, " +
                                   "   coalesce(qtde_consumida, 0) qtde_consumida, " +
                                   "   coalesce(qtde_solicitada_look_a_head, 0) qtde_solicitada_look_a_head, " +
                                   "   servico_id, " +
                                   "   coalesce(saldo, 0) saldo, " +
                                   "   origem   from SEL_LOOKAHEAD_INSUMOS('" + String.Format("{0:M/d/yyyy}", diaAnalise) + "', "
                                                                                                   + obraId.ToString() + ","
                                                                                                   + dr2["tocGrupoId"].ToString() + ","
                                                                                                   + dr2["tocServicoId"].ToString() +","
                                                                                                   +"'"+ String.Format("{0:M/d/yyyy}", mesAnalise) + "')");
                foreach (DataRow dr in dtSEL_LOOKAHEAD_INSUMOS.Rows)
                {
                    DataRow dri = dtGrid.NewRow();
                    dri["INSUMO_ID"] = dr["INSUMO_ID"];
                    dri["INSUMO"] = dr["INSUMO"];
                    dri["UNID"] = dr["UNID"];
                    dri["QTDE_TOTAL_ENTRADA"] = dr["QTDE_TOTAL_ENTRADA"];
                    dri["QTDE_CONSUMIDA"] = dr["QTDE_CONSUMIDA"];
                    dri["QTDE_SOLICITADA_LOOK_A_HEAD"] = dr["QTDE_SOLICITADA_LOOK_A_HEAD"];
                    dri["SERVICO_ID"] = dr["SERVICO_ID"];
                    dri["CONSUMO_INSUMO"] = dr["CONSUMO_INSUMO"];
                    dri["ORIGEM"] = dr["ORIGEM"];
                    dtGrid.Rows.Add(dri);
                }
            }
        }

        private void ConstruirQtdePSA()
        {
            dtQtdePSA.Columns.Add("PSA_ID", typeof(int));
            dtQtdePSA.Columns.Add("UNIDADE_PRINCIPAL", typeof(double));
            dtQtdePSA.Columns.Add("QTDE_REALIZADA", typeof(double));
            dtQtdePSA.Columns.Add("QTDE_PROJECAO", typeof(double));
            dtQtdePSA.Columns.Add("QTDE", typeof(double));
            dtQtdePSA.Columns.Add("QTDE_LOOKAHEAD_PENDENTE", typeof(double));
        }

        public string AtualizaTabela(int servicoId, double qtde, revitDB.Element ele)
        {
            string pendencias = "";
            double consumoInusmo = 0;
            string psaId = "";
            string insumoid = ""; 
            foreach (DataRow row in dtGrid.Select("SERVICO_ID like '%" + servicoId.ToString() +"%' and ORIGEM = 0" ))
            {
                if (!string.IsNullOrEmpty(ele.LookupParameter(CampoMark).AsString()))
                {
                    try
                    {
                        psaId = ele.LookupParameter(CampoMark).AsString();
                        insumoid = row["INSUMO_ID"].ToString();
                        consumoInusmo = Convert.ToDouble(
                            manipulacao.PRC_EXECUTAR_DIRETO(" SELECT  coalesce(CONSUMO_INSUMO,0)  FROM GET_CONSUMO_INSMO(" +
                                                                                          insumoid + "," + psaId + ")").Rows[0][0]);

                        row["QTDE_SOLICITADA_LOOK_A_HEAD"] = Convert.ToDouble(row["QTDE_SOLICITADA_LOOK_A_HEAD"]) + qtde * +
                                                                                                 consumoInusmo;
                    }
                    catch (Exception e)
                    {

                    }
                }
            }

            foreach (DataRow row in dtGrid.Select("SERVICO_ID like '%" + servicoId.ToString() + "%' and ORIGEM = 1"))
            {
                if (!string.IsNullOrEmpty(ele.LookupParameter(CampoMark).AsString()))
                {
                    try
                    {
                        psaId = ele.LookupParameter(CampoMark).AsString();
                        insumoid = row["INSUMO_ID"].ToString();

                        consumoInusmo = Convert.ToDouble(row["CONSUMO_INSUMO"]);

                        row["QTDE_SOLICITADA_LOOK_A_HEAD"] = Convert.ToDouble(row["QTDE_SOLICITADA_LOOK_A_HEAD"]) + qtde * +
                                                                                                 consumoInusmo;
                    }
                    catch
                    {

                    }
                }
            }


            if (!string.IsNullOrEmpty(ele.LookupParameter("L7InsumoVinculado").AsString()))
            {
             //   MessageBox.Show("SERVICO_ID like '%" + servicoId.ToString() + "%' and  origem = 3 "
              //                                                             + "   and insumo = '" + ele.LookupParameter("L7InsumoVinculado").AsString().ToString() + "'");
                foreach (DataRow row in dtGrid.Select("SERVICO_ID like '%" + servicoId.ToString() + "%' and  origem = 3 "
                                                                           + "   and insumo = '" + ele.LookupParameter("L7InsumoVinculado").AsString().ToString()+"'"))
                {
                    try
                    {
                        consumoInusmo = ele.LookupParameter("L7ConsumoInsumoVinculado").AsDouble();
                        row["QTDE_SOLICITADA_LOOK_A_HEAD"] = Convert.ToDouble(row["QTDE_SOLICITADA_LOOK_A_HEAD"]) + qtde * 
                                                                                                    consumoInusmo;
                    }
                    catch (Exception e)
                    {

                    }
                }
            }


            foreach (DataRow row in dtGrid.Select("SERVICO_ID like '%" + servicoId.ToString() + "%' and ORIGEM = 0  and  SALDO < 0"))
            {
                pendencias = pendencias + ";" + row["INSUMO"].ToString();
            }
            foreach (DataRow row in dtGrid.Select("SERVICO_ID like '%" + servicoId.ToString() + "%' and ORIGEM = 1 and  SALDO < 0"))
            {
                pendencias = pendencias + ";" + row["INSUMO"].ToString();
            }
            foreach (DataRow row in dtGrid.Select("SERVICO_ID like '%" + servicoId.ToString() + "%' and  origem = 3 "
                                                                           + "   and insumo = '" + ele.LookupParameter("L7InsumoVinculado").AsString().ToString() + "'  and  SALDO < 0"))
            {
                pendencias = pendencias + ";" + row["INSUMO"].ToString();
            }
            return pendencias;
        }
        public void AplicarLookAHead(string restricao, revitDB.Element elemento)
        {
            string psaId = "";
            psaId = elemento.LookupParameter(CampoMark).AsString();
            if (psaId == "")
            {
                return;
            }
            parametros.PSA_ID = Convert.ToInt32(psaId);
            parametros.INICIO = dtpInicio.Value;
            parametros.TERMINO = dtpTermino.Value;

            parametros.PERCENT_META = 1;
            string resultado = "";
            manipulacao.PRC_LOOKAHEAD(parametros, false);



            if (string.IsNullOrEmpty(restricao))
            {

                Util.AlterarGraficoElemento(vistasDeAvanco, orgLookahead, elemento.Id);

            }
            else
            {

                Util.AlterarGraficoElemento(vistasDeAvanco, orgRestricao, elemento.Id);
                elemento.LookupParameter("L7Restricao").Set(restricao);
            }


        }

       



        public double qtdeElemento(revitDB.Element element)
        {
            double qtde = 0;
          DataRow [] linhas =   dtQtdePSA.Select("PSA_ID = " + element.LookupParameter(CampoMark).AsString());

          if (linhas.Length >0)
            {
              qtde =   Convert.ToDouble(linhas[0]["UNIDADE_PRINCIPAL"]) -
                Convert.ToDouble(linhas[0]["QTDE_REALIZADA"]) -
                Convert.ToDouble(linhas[0]["QTDE_PROJECAO"]) -
                Convert.ToDouble(linhas[0]["QTDE"]);
            }
            else
            {
                qtde = -1000;
            }

            return qtde;

        }
        private void button2_Click(object sender, EventArgs e)
        {
            string restricao = "";
            double qtde = 0;
            foreach (revitDB.ElementId ele in uiApp.ActiveUIDocument.Selection.GetElementIds())
            {
                try
                {
                    revitDB.Element element;
                    element = uiApp.ActiveUIDocument.Document.GetElement(ele);
                    qtde = qtdeElemento(element);
                    if (qtde != -1000)
                    {
                        restricao = AtualizaTabela(Convert.ToInt32(element.LookupParameter("tocServicoId").AsString()),
                                                                          qtde, element);
                        AplicarLookAHead(restricao, element);
                    }
                    else
                    {

                    }

                }
                catch (Exception e2)
                {

                }

            }
            uiApp.ActiveUIDocument.Document.ActiveView.HideElementsTemporary(uiApp.ActiveUIDocument.Selection.GetElementIds());
            uiApp.ActiveUIDocument.Document.ActiveView.UnhideElements(uiApp.ActiveUIDocument.Selection.GetElementIds());

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void cmbServico_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (atualizando) return;
            if (cmbServico.SelectedValue.ToString() == "Todos")
            {
                bs.RemoveFilter();

            }
            else
            {
                bs.Filter = " SERVICO_ID " + " = " + cmbServico.SelectedValue.ToString();

            }
        }
    }
}

