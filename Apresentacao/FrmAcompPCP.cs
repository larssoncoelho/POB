using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using FirebirdSql.Data;
using FirebirdSql.Data.FirebirdClient;
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
    public partial class FrmAcompPCP : NForm
    {
        public Plan_servico_amoNegocio manipulacao;
        public NBindingSource bsElemento = new NBindingSource();
        public NBindingSource bsCriterio = new NBindingSource();
        public NBindingSource bsCausas = new NBindingSource();
        public NDataTable dtElemento = new NDataTable("Elemento");
        public NDataTable dtCriterio = new NDataTable("Criterio");
        public NDataTable dtCausa = new NDataTable("Causa");
        public DataSet ds = new DataSet("Tabelas");
        public DataRelation drCriterioPlanServicoAmoId;
        public DataRelation drCausa;
        public string elementosSelecionados = "";
        public FbsqlConnection sqlConnection = new FbsqlConnection();
        private UIApplication uiApp;
        public ExternalCommandData revit;
        private revitDB.ElementId ele1;
        private revitDB.Document uiDoc;
        private revitDB.Element ele;
        public string dir = "";
        public string novaCausa = "";
        public string ItemText { get; set; }
        DataView dv = new DataView();
        private revitDB.ElementId preenchimentoId;
        private revitDB.OverrideGraphicSettings orgProjecao;
        private revitDB.OverrideGraphicSettings orgCompleto;
        private revitDB.OverrideGraphicSettings orgNaoIniciado;
        private revitDB.OverrideGraphicSettings orgIniciado;
        private revitDB.OverrideGraphicSettings orgCompletoAposDataLimite;
        private revitDB.OverrideGraphicSettings orgTerminado;


        public FrmAcompPCP(string idir)
        {
            InitializeComponent();
            manipulacao = new Plan_servico_amoNegocio(idir);
        }

        public void Abrir()
        {

            bsCausas.abrindo = true;
            bsCriterio.abrindo = true;
            bsElemento.abrindo = true;
            grdElementos.PermitirFullSelect = true;

            sqlConnection.Diretorio = dir;
            sqlConnection.Ativo(true);
            DefinirElemento();
            DefinirCriterio();
            DefinirCausa();
            bsElemento.DataSource = ds;
            bsElemento.DataMember = dtElemento.TableName;
            bsElemento.icampoPrimaryKey = "";
            bsElemento.igen_id = "";
            bsElemento.ifbConnection = sqlConnection.FbConexao;
            drCriterioPlanServicoAmoId = new DataRelation("CriterioPlanServicoAmoId",
                                                                ds.Tables[dtElemento.TableName].Columns["PLAN_SERVICO_AMO_ID"],


                                                                ds.Tables[dtCriterio.TableName].Columns["plan_servico_amo_pai_id"]);

            ds.Relations.Add(drCriterioPlanServicoAmoId);
            bsCriterio.DataSource = bsElemento;
            bsCriterio.DataMember = "CriterioPlanServicoAmoId";
            bsCriterio.ifbConnection = sqlConnection.FbConexao;

            drCausa = new DataRelation("CausaCriterio",
                                                    ds.Tables[dtCriterio.TableName].Columns["criterio_plan_servico_amo_id"],
                                                    ds.Tables[dtCausa.TableName].Columns["criterio_plan_servico_amo_id"]);

            ds.Relations.Add(drCausa);
            bsCausas.DataSource = bsCriterio;
            bsCausas.DataMember = "CausaCriterio";
            bsCausas.ifbConnection = sqlConnection.FbConexao;
            bsCausas.icampoPrimaryKey = "CAUSA_CRITERIO_ID";
            bsCausas.igen_id = "GEN_CAUSA_CRITERIO_ID";

            grdElementos.DataSource = bsElemento;
            grdCriterio.DataSource = bsCriterio;
            grdCausa.DataSource = bsCausas;
            bnElemento.BindingSource = bsElemento;
            bnCausa.BindingSource = bsCausas;
            bnCriterio.BindingSource = bsCriterio;
            bsCausas.abrindo = false;
            bsCriterio.abrindo = false;
            bsElemento.abrindo = false;

            this.ShowDialog();
        }

        public void TrataGrids()
        {

        }
        public void DefinirDataSets()
        {

        }
        public void DefinirDataSourceGris()
        {

        }


        public StringBuilder sqlPLAN_SERVICO_AMO_ID(string elementos1)
        {
            StringBuilder sb1 = new StringBuilder();
            sb1.Clear();
            sb1.Append(" select psa.plan_servico_amo_id, ");
            sb1.Append(" coalesce(samo.id_bim, 0) || ' ' descricao, ");
            sb1.Append(" samo.descricao COMPLEMENTO, ");
            sb1.Append(" psa.posicao, ");
            sb1.Append("   samo.unidade_principal projeto, ");
            sb1.Append("   psa.servico_amo_id, ");
            sb1.Append("  psa.obs, ");
            sb1.Append("  samo.medicao_bloco_id, ");
            sb1.Append(" m.medicao, ");
            sb1.Append(" mb.bloco_id, ");
            sb1.Append(" b.bloco, ");
            sb1.Append(" psa.grupo_plan_servico_id, ");
            sb1.Append(" gps.grupo, psa.cad, psa.data_cad, psa.alt, psa.data_alt, samo.servico_id, s.servico, ");
            sb1.Append(" s.unid, SAMO.ID_BIM, ");
            sb1.Append(" y.projecao,");
            sb1.Append(" y.QTDE_REALIZADA_FILHO ");
            sb1.Append(" from plan_servico_amo psa ");
            sb1.Append(" inner join servico_amo samo on psa.servico_amo_id = samo.servico_amo_id ");
            sb1.Append(" inner join medicao_bloco mb on samo.medicao_bloco_id = mb.medicao_bloco_id ");
            sb1.Append(" inner join grupo_plan_servico gps on psa.grupo_plan_servico_id = gps.grupo_plan_servico_id ");
            sb1.Append(" inner join servico s on samo.servico_id = s.servico_id ");
            sb1.Append(" inner join bloco b on mb.bloco_id = b.bloco_id ");
            sb1.Append(" inner join medicao m on mb.medicao_id = m.medicao_id ");
            sb1.Append(" left join (select  sum(psa1.qtde_realizada) QTDE_REALIZADA_FILHO, sum(psa1.qtde_projecao) projecao, psaf.plan_servico_amo_pai_id ");
            sb1.Append("     from plan_servico_amo psa1 ");
            sb1.Append("   inner join plan_servico_amo_filho psaf on psa1.plan_servico_amo_id = psaf.plan_servico_amo_filho_id ");
            sb1.Append("    group by psaf.plan_servico_amo_pai_id) y on psa.plan_servico_amo_id = y.plan_servico_amo_pai_id ");
            sb1.Append("  where psa.plan_servico_amo_id in  (" + elementos1 + ")");

            return sb1;

        }

        public StringBuilder sqlCRITERIO(string elementos1)
        {
            StringBuilder sb1 = new StringBuilder();
            sb1.Clear();
            sb1.Append(" select cps.criterio_plan_servico_amo_id,  ");
            sb1.Append(" cps.plan_servico_amo_id, ");
            sb1.Append(" psaf.plan_servico_amo_pai_id, ");

            sb1.Append(" cms.criterio_descricao, ");
            sb1.Append(" cps.peso, ");
            sb1.Append(" fe.fornecedor empreiteiro, ");
            sb1.Append(" fenc.fornecedor encarregado, ");
            sb1.Append(" cps.criterio_qtde_mapa_servico_id, ");
            sb1.Append(" cps.fornecedor_empreiteiro_id, ");
            sb1.Append(" cps.fornecedor_encarregado_id, ");
            sb1.Append(" cps.latencia, ");
            sb1.Append(" cps.inicio, ");
            sb1.Append(" cps.termino, ");
            sb1.Append(" cps.qtde_mapa_servico_id, ");
            sb1.Append(" cps.produtividade, ");
            sb1.Append(" cps.tipo_ativo_id, ");
            sb1.Append(" cps.tipo_estato_id, ");
            sb1.Append(" cps.executado ");
            sb1.Append(" from criterio_plan_servico_amo cps ");
            sb1.Append("  INNER JOIN criterio_qtde_mapa_servico cms ON cps.criterio_qtde_mapa_servico_id = cms.criterio_qtde_mapa_servico_id ");
            sb1.Append("  LEFT JOIN FORNECEDOR FE ON cps.fornecedor_empreiteiro_id = FE.fornecedor_id ");
            sb1.Append("  LEFT JOIN FORNECEDOR FENC ON cps.fornecedor_encarregado_id = FENC.fornecedor_id ");
            sb1.Append(" inner join plan_servico_amo psa on cps.plan_servico_amo_id = psa.plan_servico_amo_id ");
            sb1.Append(" inner join plan_servico_amo_filho psaf on psa.plan_servico_amo_id = psaf.plan_servico_amo_filho_id ");
            sb1.Append(" where psaf.plan_servico_amo_pai_id in (" + elementos1 + ")");
            sb1.Append(" and cps.executado <> 1");

            sb1.Append(" order by cps.criterio_plan_servico_amo_id desc ");



            return sb1;

        }

        public StringBuilder sqlCAUSAS(string elementos1)
        {
            StringBuilder sb1 = new StringBuilder();
            sb1.Clear();

            sb1.Append(" select CC.CAUSA_CRITERIO_ID, CC.CRITERIO_PLAN_SERVICO_AMO_ID, CC.DESCRICAO_CAUSA ");

            sb1.Append(" FROM CAUSA_CRITERIO CC ");
            sb1.Append("  inner join criterio_plan_servico_amo cps  on cc.criterio_plan_servico_amo_id = cps.criterio_plan_servico_amo_id ");
            sb1.Append("  inner join plan_servico_amo psa on cps.plan_servico_amo_id = psa.plan_servico_amo_id ");
            sb1.Append("  inner join plan_servico_amo_filho psaf on psa.plan_servico_amo_id = psaf.plan_servico_amo_filho_id ");
            sb1.Append("    where psaf.plan_servico_amo_pai_id in (" + elementos1 + ") order by cps.criterio_plan_servico_amo_id desc ");

            return sb1;

        }

        public void DefinirElemento()
        {
            dtElemento.ifbcommandSelect.Connection = sqlConnection.FbConexao;
            dtElemento.ifbcommandSelect.CommandText = sqlPLAN_SERVICO_AMO_ID(elementosSelecionados).ToString();
            dtElemento.ida = new FbDataAdapter(dtElemento.ifbcommandSelect);
            dtElemento.ifbUpdate = new FbCommand("update plan_servico_amo " +
                                       " set prof_concretada = @prof_concretada,  " +
                                       "     prof_escavada = @prof_escavada, " +
                                       "     vol_relatorio = @vol_relatorio " +
                                       " where (plan_servico_amo_id = @plan_servico_amo_id)", sqlConnection.FbConexao);
            dtElemento.ifbUpdate.Parameters.Add(new FbParameter("plan_servico_amo_id", FbDbType.Integer, 0, "plan_servico_amo_id"));
            dtElemento.ifbUpdate.Parameters.Add(new FbParameter("prof_concretada", FbDbType.Double, 0, "prof_concretada"));
            dtElemento.ifbUpdate.Parameters.Add(new FbParameter("prof_escavada", FbDbType.Double, 0, "prof_escavada"));
            dtElemento.ifbUpdate.Parameters.Add(new FbParameter("vol_relatorio", FbDbType.Double, 0, "vol_relatorio"));
            dtElemento.ida.UpdateCommand = dtElemento.ifbUpdate;
            dtElemento.ida.Fill(dtElemento);
            ds.Tables.Add(dtElemento);
        }

        public void DefinirCriterio()
        {
            dtCriterio.ifbcommandSelect.Connection = sqlConnection.FbConexao;
            dtCriterio.ifbcommandSelect.CommandText = sqlCRITERIO(elementosSelecionados).ToString();
            dtCriterio.ida = new FbDataAdapter(dtCriterio.ifbcommandSelect);
            dtCriterio.ifbUpdate = new FbCommand("UPDATE criterio_plan_servico_amo SET " +
                          " CRITERIO_QTDE_MAPA_SERVICO_ID = @CRITERIO_QTDE_MAPA_SERVICO_ID, " +
                          " FORNECEDOR_EMPREITEIRO_ID = @FORNECEDOR_EMPREITEIRO_ID, " +
                          " FORNECEDOR_ENCARREGADO_ID = @FORNECEDOR_ENCARREGADO_ID, " +
                          " INICIO = @INICIO, " +
                         "  TERMINO = @TERMINO " +
                        " WHERE  criterio_plan_servico_amo.CRITERIO_PLAN_SERVICO_AMO_ID = @CRITERIO_PLAN_SERVICO_AMO_ID", sqlConnection.FbConexao);

            dtCriterio.ifbUpdate.Parameters.Add(new FbParameter("CRITERIO_QTDE_MAPA_SERVICO_ID", FbDbType.Integer, 0, "CRITERIO_QTDE_MAPA_SERVICO_ID"));
            dtCriterio.ifbUpdate.Parameters.Add(new FbParameter("FORNECEDOR_EMPREITEIRO_ID", FbDbType.Integer, 0, "FORNECEDOR_EMPREITEIRO_ID"));
            dtCriterio.ifbUpdate.Parameters.Add(new FbParameter("FORNECEDOR_ENCARREGADO_ID", FbDbType.Integer, 0, "FORNECEDOR_ENCARREGADO_ID"));
            dtCriterio.ifbUpdate.Parameters.Add(new FbParameter("INICIO", FbDbType.Date, 0, "INICIO"));
            dtCriterio.ifbUpdate.Parameters.Add(new FbParameter("TERMINO", FbDbType.Date, 0, "TERMINO"));
            dtCriterio.ifbUpdate.Parameters.Add(new FbParameter("CRITERIO_PLAN_SERVICO_AMO_ID", FbDbType.Integer, 0, "CRITERIO_PLAN_SERVICO_AMO_ID"));
            dtCriterio.ida.UpdateCommand = dtCriterio.ifbUpdate;
            dtCriterio.ida.Fill(dtCriterio);
            ds.Tables.Add(dtCriterio);
        }

        public void DefinirCausa()
        {
            dtCausa.ifbcommandSelect.Connection = sqlConnection.FbConexao;
            dtCausa.ifbcommandSelect.CommandText = sqlCAUSAS(elementosSelecionados).ToString();
            dtCausa.ida = new FbDataAdapter(dtCausa.ifbcommandSelect);
            dtCausa.ifbUpdate = new FbCommand("UPDATE causa_criterio SET " +
              " DESCRICAO_CAUSA = @DESCRICAO_CAUSA, " +
              " CRITERIO_PLAN_SERVICO_AMO_ID = @CRITERIO_PLAN_SERVICO_AMO_ID " +
            " WHERE  causa_criterio_ID = @causa_criterio_ID", sqlConnection.FbConexao);
            dtCausa.ifbUpdate.Parameters.Add(new FbParameter("DESCRICAO_CAUSA", FbDbType.VarChar, 250, "DESCRICAO_CAUSA"));
            dtCausa.ifbUpdate.Parameters.Add(new FbParameter("CRITERIO_PLAN_SERVICO_AMO_ID", FbDbType.Integer, 0, "CRITERIO_PLAN_SERVICO_AMO_ID"));
            dtCausa.ifbUpdate.Parameters.Add(new FbParameter("causa_criterio_ID", FbDbType.Integer, 0, "causa_criterio_ID"));
            dtCausa.ida.UpdateCommand = dtCausa.ifbUpdate;

            dtCausa.ifbDelete = new FbCommand("DELETE FROM CAUSA_CRITERIO WHERE CAUSA_CRITERIO_ID = @CAUSA_CRITERIO_ID", sqlConnection.FbConexao);
            dtCausa.ifbDelete.Parameters.Add(new FbParameter("causa_criterio_ID", FbDbType.Integer, 0, "causa_criterio_ID"));
            dtCausa.ida.DeleteCommand = dtCausa.ifbDelete;

            dtCausa.ifbInsert = new FbCommand("insert into CAUSA_CRITERIO (DESCRICAO_CAUSA, CRITERIO_PLAN_SERVICO_AMO_ID) VALUES " +
                                            " (@DESCRICAO_CAUSA, @CRITERIO_PLAN_SERVICO_AMO_ID)", sqlConnection.FbConexao);

            dtCausa.ifbInsert.Parameters.Add(new FbParameter("DESCRICAO_CAUSA", FbDbType.VarChar, 250, "DESCRICAO_CAUSA"));

            dtCausa.ifbInsert.Parameters.Add(new FbParameter("CRITERIO_PLAN_SERVICO_AMO_ID", FbDbType.Integer, 0, "CRITERIO_PLAN_SERVICO_AMO_ID"));
            dtCausa.ida.InsertCommand = dtCausa.ifbInsert;

            dtCausa.ida.Fill(dtCausa);
            ds.Tables.Add(dtCausa);
        }

        private void btnNovaCausa_Click(object sender, EventArgs e)
        {
            DataRow r;
            r = dtCausa.NewRow();
            r[bsCausas.icampoPrimaryKey] = FbSequence.Generator(bsCausas.igen_id, bsCausas.ifbConnection);
            r["insumo_id"] = resultadoProcura.vCampo[0];
            r["insumo"] = resultadoProcura.vCampo[1];
            r["Unid"] = resultadoProcura.vCampo[2];
            dtCausa.Rows.InsertAt(r, 0);
            bsCausas.EndEdit();

            bsCausas.Find("sgi_material_controlado_id", r[bsCausas.icampoPrimaryKey]);
        }

        private void grdCausa_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {

        }

        private void bindingSource1_AddingNew(object sender, AddingNewEventArgs e)
        {

        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {

            // Perguntar inputBox = new Perguntar();
            //inputBox.Inputar("Digite a unidade de medição", "Selecionar unidade de medição");
            // = inputBox.resultadoProcura.vCampo[0].ToString();
        }

        private void btnNovo_Click(object sender, EventArgs e)
        {
            Perguntar inputBox = new Perguntar();
            inputBox.Inputar("Digite a causa", "");
            novaCausa = inputBox.resultadoProcura.vCampo[0].ToString();
            PAR_INCLUIR_CAUSA par = new PAR_INCLUIR_CAUSA();
            par.DESCRICAO_CAUSA = novaCausa;
            par.CRITERIO_PLAN_SERVICO_AMO_ID = Convert.ToInt32(grdCriterio.CurrentRow.Cells["CRITERIO_PLAN_SERVICO_AMO_ID"].Value);
            manipulacao.PRC_INCLUIR_CAUSA(par);
        }

        private void btnExcluir_Click(object sender, EventArgs e)
        {
            bsCausas.RemoveCurrent();

        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            string tipoElemento = "";
            revitDB.FilteredElementCollector selecao;
        

            double percentAvanco;
            uiApp = revit.Application;
            uiDoc = uiApp.ActiveUIDocument.Document;
            Selection sel = uiApp.ActiveUIDocument.Selection;
            Util.uiDoc = uiDoc;


            selecao = new revitDB.FilteredElementCollector(uiDoc).OfClass(typeof(Autodesk.Revit.DB.View));
            List<revitDB.View> vistasDeAvanco = new List<revitDB.View>();
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
            preenchimentoId = (Util.FindElementByName( typeof(revitDB.Material), "Previsto") as revitDB.Material).SurfacePatternId;
            orgProjecao = Util.GetOverrideGraphicSettings(Util.GetColorRevit(corLinhaPCP),
                                                         Util.GetColorRevit(corSuperficiePCP),
                                                         preenchimentoId, 0);

            //parametros = new PAR_PCP_PSA();
            switch (ItemText)
            {
                case "95":
                    percentAvanco = 0.95;
                    break;
                case "50":
                    percentAvanco = 0.5;
                    break;
                default:
                    Perguntar inputBox = new Perguntar();
                    inputBox.Inputar("Digite a procentagem", "Digite a %");
                    percentAvanco = Convert.ToDouble(inputBox.resultadoProcura.vCampo[0]);
                    inputBox.Dispose();
                    break;

            }
            EscolherData escolherData = new EscolherData();
            DialogResult resultado = escolherData.inputar(ref diaProjecao, "Escolha o dia", ref continuar);

            if (!continuar)
            {
                return;
            }

            escolherData.Dispose();
            progresso = new ProgressoFuncao();
            progresso.pgc.Maximum = sel.GetElementIds().Count;
            progresso.pgc.Minimum = 0;
            progresso.pgc.Value = 1;
            progresso.pgc.Step = 1;
            progresso.Show();


            if (grdElementos.SelectedRows.Count > 0)

                foreach (DataGridViewRow row in grdElementos.SelectedRows)

                {

                    revitDB.ElementId eleId = new revitDB.ElementId(Convert.ToInt32(row.Cells["DESCRICAO"].Value));
                    ele = uiDoc.GetElement(eleId);
                    Util.RodarPCP(progresso, ele,  manipulacao,
                                   uiApp, Lista, CampoMark, diaProjecao, percentAvanco,
                                   vistasDeAvanco, orgProjecao, ele.Name, CampoArea, "onLine");

                }
            else
            {
                revitDB.ElementId eleId = new revitDB.ElementId(Convert.ToInt32(grdElementos.CurrentRow.Cells["DESCRICAO"].Value));
                ele = uiDoc.GetElement(eleId);
                Util.RodarPCP(progresso, ele,  manipulacao,
                               uiApp, Lista, CampoMark, diaProjecao, percentAvanco,
                               vistasDeAvanco, orgProjecao, ele.Name, CampoArea, "onLine");
            }

            int psaId = Convert.ToInt32(grdElementos.CurrentRow.Cells["plan_servico_amo_id"].Value);
            dtCausa.Rows.Clear();
            dtCriterio.Rows.Clear();
            dtElemento.Rows.Clear();
            dtElemento.ida.Fill(dtElemento);
            dtCriterio.ida.Fill(dtCriterio);
            dtCausa.ida.Fill(dtCausa);
            progresso.Close();
            progresso.Dispose();
            bsElemento.Position = bsElemento.Find("PLAN_SERVICO_AMO_ID", psaId);

        }

        private void toolStripButton1_Click_1(object sender, EventArgs e)
        {
            int psaId = 0;


            PAR_FECHAR_PCP par = new PAR_FECHAR_PCP();

            if (grdElementos.SelectedRows.Count > 0)
            {

                EscolherData escolherData = new EscolherData();
                DialogResult resultado = escolherData.inputar(ref diaProjecao, "Escolha o dia", ref continuar);
                foreach (DataGridViewRow row in grdElementos.SelectedRows)
                {
                    if (!continuar)
                    {
                        return;
                    }
                    par.ENOVO_TERMINO = diaProjecao;
                    par.EPSA_ID = Convert.ToInt32(row.Cells["plan_servico_amo_id"].Value);
                    manipulacao.PRC_FECHAR_PCP(par);
                }
            }
            else
            {
                EscolherData escolherData = new EscolherData();
                DialogResult resultado = escolherData.inputar(ref diaProjecao, "Escolha o dia", ref continuar);

                if (!continuar)
                {
                    return;
                }
                par.ENOVO_TERMINO = diaProjecao;
                par.EPSA_ID = Convert.ToInt32(grdElementos.CurrentRow.Cells["plan_servico_amo_id"].Value);
                manipulacao.PRC_FECHAR_PCP(par);
            }

            psaId = Convert.ToInt32(grdElementos.CurrentRow.Cells["plan_servico_amo_id"].Value);

            dtCausa.Rows.Clear();
            dtCriterio.Rows.Clear();
            dtElemento.Rows.Clear();
            dtElemento.ida.Fill(dtElemento);
            dtCriterio.ida.Fill(dtCriterio);
            dtCausa.ida.Fill(dtCausa);
            bsElemento.Position = bsElemento.Find("PLAN_SERVICO_AMO_ID", psaId);


        }

        private void toolStripButton7_Click(object sender, EventArgs e)
        {
            uiApp = revit.Application;
            uiDoc = uiApp.ActiveUIDocument.Document;
            Util.uiDoc = uiDoc;

            preenchimentoId = (Util.FindElementByName( typeof(revitDB.Material), "Previsto") as revitDB.Material).SurfacePatternId;
            orgProjecao = Util.GetOverrideGraphicSettings(Util.GetColorRevit(corLinhaPCP),
                                 Util.GetColorRevit(corSuperficiePCP),
                                 preenchimentoId, 0);
            orgCompleto= Util.GetOverrideGraphicSettings(Util.GetColorRevit(CorLinhaCompleto),
                                             Util.GetColorRevit(CorSuperficieCompleto),
                                             preenchimentoId, 50);
            orgCompleto = Util.GetOverrideGraphicSettings(Util.GetColorRevit(CorLinhaCompleto),
                                            Util.GetColorRevit(CorSuperficieCompleto),
                                            preenchimentoId, 0);
            orgNaoIniciado = Util.GetOverrideGraphicSettings(Util.GetColorRevit(CorLinhaNaoIniciado),
                                             Util.GetColorRevit(CorSuperficieNaoIniciado),
                                             preenchimentoId, 0);
            orgIniciado = Util.GetOverrideGraphicSettings(Util.GetColorRevit(CorLinhaIniciado),
                                 Util.GetColorRevit(CorSuperficieIniciado),
                                 preenchimentoId, 0);
            orgTerminado = Util.GetOverrideGraphicSettings(Util.GetColorRevit(CorLinhaCompleto),
                                             Util.GetColorRevit(CorSuperficieCompleto),
                                             preenchimentoId, 0);
            orgCompletoAposDataLimite = Util.GetOverrideGraphicSettings(Util.GetColorRevit(CorLinhaExecutadoAposLimite),
                                             Util.GetColorRevit(CorSuperficieExecutadoAposLimite),
                                             preenchimentoId, 0);

            int id = Convert.ToInt32(grdElementos.CurrentRow.Cells["PLAN_SERVICO_AMO_ID"].Value);

            if (grdElementos.SelectedRows.Count > 0)

                foreach (DataGridViewRow row in grdElementos.SelectedRows)
                {
                    manipulacao.PRC_EXECUTAR_DIRETO("EXECUTE PROCEDURE PRC_ELIMINAR_PCP_PSA(" +
                        Convert.ToString(row.Cells["PLAN_SERVICO_AMO_ID"].Value) + ")");
                    Util.RodarSubGraficoPorElemento(uiApp, new revitDB.ElementId(Convert.ToInt32(row.Cells["ID_BIM"].Value)),
                                                    dir, orgProjecao, orgIniciado, orgNaoIniciado, orgCompleto,
                                                    orgTerminado, orgCompletoAposDataLimite,  CampoMark);
                }

            else
            {
                manipulacao.PRC_EXECUTAR_DIRETO("EXECUTE PROCEDURE PRC_ELIMINAR_PCP_PSA(" +
                    Convert.ToString(grdElementos.CurrentRow.Cells["PLAN_SERVICO_AMO_ID"].Value) + ")");
                Util.RodarSubGraficoPorElemento(uiApp, new revitDB.ElementId(Convert.ToInt32(grdElementos.CurrentRow.Cells["ID_BIM"].Value)),
                                dir, orgProjecao, orgIniciado, orgNaoIniciado, orgCompleto,
                                                    orgTerminado, orgCompletoAposDataLimite, CampoMark);
            }
            dtCausa.Rows.Clear();
            dtCriterio.Rows.Clear();

            dtElemento.Rows.Clear();

            dtElemento.ida.Fill(dtElemento);
            dtCriterio.ida.Fill(dtCriterio);
            dtCausa.ida.Fill(dtCausa);
            bsElemento.Position = bsElemento.Find("PLAN_SERVICO_AMO_ID", id);
            
        }

        private void grdCriterio_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            //if ((sender as DataGridView).)
        }

        private void grdCriterio_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            string filtro = "";
            int j = 0;
            if (string.IsNullOrEmpty(textBox1.Text))
            {

                bsElemento.Filter = "";
                return;
            }
            else
            {



                for (int i = 0; i < dtElemento.Columns.Count - 1; i++)
                {


                    switch (dtElemento.Columns[i].DataType.Name)
                    {
                        case "String":
                            if (j == 0)
                            {
                                filtro = "(" + dtElemento.Columns[i].ColumnName + " like '%" + textBox1.Text + "%')";

                                j = j + 1;

                            }
                            else
                            {

                                filtro = filtro + " or (" + dtElemento.Columns[i].ColumnName + " like '%" + textBox1.Text + "%')";
                                j = j + 1;

                            }
                            break;
                        case "Int32":
                        case "DOUBLE":
                            if (j == 0)
                            {
                                try
                                {
                                    Convert.ToDouble(textBox1.Text);
                                    filtro = "(" + dtElemento.Columns[i].ColumnName + " = " + textBox1.Text + ")";
                                    j = j + 1;

                                }
                                catch
                                {

                                }

                            }
                            else
                            {
                                try
                                {
                                    Convert.ToDouble(textBox1.Text);
                                    filtro = filtro + " or (" + dtElemento.Columns[i].ColumnName + " = " + textBox1.Text + ")";
                                    j = j + 1;

                                }
                                catch
                                {

                                }


                            }

                            break;

                    }
                }

                bsElemento.Filter = filtro;
            }
        }

        private void bnElemento_RefreshItems(object sender, EventArgs e)
        {

        }
    }
}
