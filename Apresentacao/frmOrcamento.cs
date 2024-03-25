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
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using appRevit = Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Structure;


using AcessoBancoDados;
using ObjetoTransferencia;
namespace Apresentacao
{
    public partial class frmOrcamento : NForm
    {
        public string linha1;
        public string ielementos;

        public List<ElementId> lista = new List<ElementId>();
        public appRevit.Application r;
        public Element ele;
        public DataSet ds = new DataSet();
        public string internalDir;
        public StringBuilder sb = new StringBuilder();
        public UIApplication revitI;
        public ElementId ele1;
        public FbCommand fbApropriarremessa;
        public FbsqlConnection sqlConnection = new FbsqlConnection();
        public frmOrcamento()
        {
            InitializeComponent();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
                checkBox2.Checked = false;
            else checkBox2.Checked = true;

        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
                checkBox1.Checked = false;
            else checkBox1.Checked = true;
        }
        public void Abrir(string dir, string elementos, UIApplication uiApp)
        {
            revitI = uiApp;
            ielementos = elementos;
            internalDir = dir;
            /*dtgSGIConcreto.ifbcommandSelect = new FbCommand();

            sqlConnection.Diretorio = dir;
            sqlConnection.Ativo(true);
            //dados sgi


            dtgSGIConcreto.ifbcommandSelect.Connection = sqlConnection.FbConexao;

            #region SGIConcreto
            sqlSgiEstaca(elementos, "");
            dtgSGIConcreto.ifbcommandSelect.CommandText = sb.ToString();
            dtgSGIConcreto.ida = new FbDataAdapter(dtgSGIConcreto.ifbcommandSelect);
            dtgSGIConcreto.ifbUpdate = new FbCommand("update plan_servico_amo " +
                                       " set prof_concretada = @prof_concretada,  " +
                                       "     prof_escavada = @prof_escavada, " +
                                       "     vol_relatorio = @vol_relatorio " +
                                       " where (plan_servico_amo_id = @plan_servico_amo_id)", sqlConnection.FbConexao);


            dtgSGIConcreto.ifbUpdate.Parameters.Add(new FbParameter("plan_servico_amo_id", FbDbType.Integer, 0, "plan_servico_amo_id"));
            dtgSGIConcreto.ifbUpdate.Parameters.Add(new FbParameter("prof_concretada", FbDbType.Double, 0, "prof_concretada"));
            dtgSGIConcreto.ifbUpdate.Parameters.Add(new FbParameter("prof_escavada", FbDbType.Double, 0, "prof_escavada"));
            dtgSGIConcreto.ifbUpdate.Parameters.Add(new FbParameter("vol_relatorio", FbDbType.Double, 0, "vol_relatorio"));

            dtgSGIConcreto.ida.UpdateCommand = dtgSGIConcreto.ifbUpdate;
            dtgSGIConcreto.idt = new DataTable("SGI");
            dtgSGIConcreto.ida.Fill(dtgSGIConcreto.idt);
            ds.Tables.Add(dtgSGIConcreto.idt);
            #endregion

            #region remessaApropriada

            dtgRemessaApropriada.ifbcommandSelect = new FbCommand(sqlmaterial_cntrl_sgi_dado(), sqlConnection.FbConexao);



            dtgRemessaApropriada.ida = new FbDataAdapter(dtgRemessaApropriada.ifbcommandSelect);
            dtgRemessaApropriada.ifbUpdate = new FbCommand(" update material_cntrl_sgi_dado  " +
                                                      "    set posicao = @posicao,  " +
                                                        "      plan_servico_amo_id = @plan_servico_amo_id,  " +
                                                         "     sgi_material_controlado_id = @sgi_material_controlado_id,  " +
                                                         "     percent = @percent  " +
                                                        "  where (material_cntrl_sgi_dado_id =   " +
                                                    "  @material_cntrl_sgi_dado_id)", sqlConnection.FbConexao);


            dtgRemessaApropriada.ifbUpdate.Parameters.Add(new FbParameter("POSICAO", FbDbType.Integer, 0, "POSICAO"));
            dtgRemessaApropriada.ifbUpdate.Parameters.Add(new FbParameter("plan_servico_amo_id", FbDbType.Integer, 0, "plan_servico_amo_id"));
            dtgRemessaApropriada.ifbUpdate.Parameters.Add(new FbParameter("sgi_material_controlado_id", FbDbType.Integer, 0, "sgi_material_controlado_id"));
            dtgRemessaApropriada.ifbUpdate.Parameters.Add(new FbParameter("percent", FbDbType.Double, 0, "PERCENT"));
            dtgRemessaApropriada.ifbUpdate.Parameters.Add(new FbParameter("material_cntrl_sgi_dado_id", FbDbType.Integer, 0, "material_cntrl_sgi_dado_id"));
            dtgRemessaApropriada.ida.UpdateCommand = dtgRemessaApropriada.ifbUpdate;
            dtgRemessaApropriada.idt = new DataTable("REMESSA_APRORIADA");
            dtgRemessaApropriada.ida.Fill(dtgRemessaApropriada.idt);
            ds.Tables.Add(dtgRemessaApropriada.idt);
            dtgRemessaApropriada.ibs = new NBindingSource();
            #endregion

            #region Remessa

            dtgRemessa.ifbcommandSelect = new FbCommand(sqlMaterial(), sqlConnection.FbConexao);



            dtgRemessa.ida = new FbDataAdapter(dtgRemessa.ifbcommandSelect);
            dtgRemessa.ifbUpdate = new FbCommand(" update sgi_material_controlado" +
                           " set insumo_id = @insumo_id," +
                                " obra_id = @obra_id," +
                                " nf = @nf," +
                                " remessa = @remessa," +
                                " data = @data," +
                                " qtde = @qtde," +
                                " splump = @splump," +
                                " fck = @fck," +
                                " numero_cp = @numero_cp," +
                                " cp7 = @cp7," +
                                " cp28 = @cp28," +
                                " cp7c = @cp7c," +
                                " cp28c = @cp28c" +
                           "  where sgi_material_controlado_id = @sgi_material_controlado_id", sqlConnection.FbConexao);


            dtgRemessa.ifbUpdate.Parameters.Add(new FbParameter("sgi_material_controlado_id", FbDbType.Integer, 0, "sgi_material_controlado_id"));
            dtgRemessa.ifbUpdate.Parameters.Add(new FbParameter("insumo_id", FbDbType.Integer, 0, "insumo_id"));
            dtgRemessa.ifbUpdate.Parameters.Add(new FbParameter("obra_id", FbDbType.Integer, 0, "obra_id"));
            dtgRemessa.ifbUpdate.Parameters.Add(new FbParameter("NF", FbDbType.VarChar, 250, "nf"));
            dtgRemessa.ifbUpdate.Parameters.Add(new FbParameter("REMESSA", FbDbType.VarChar, 250, "remessa"));
            dtgRemessa.ifbUpdate.Parameters.Add(new FbParameter("DATA", FbDbType.Date, 0, "data"));
            dtgRemessa.ifbUpdate.Parameters.Add(new FbParameter("QTDE", FbDbType.Double, 0, "qtde"));
            dtgRemessa.ifbUpdate.Parameters.Add(new FbParameter("SPLUMP", FbDbType.Double, 0, "splump"));
            dtgRemessa.ifbUpdate.Parameters.Add(new FbParameter("FCK", FbDbType.Double, 0, "fck"));
            dtgRemessa.ifbUpdate.Parameters.Add(new FbParameter("NUMERO_CP", FbDbType.VarChar, 250, "numero_cp"));
            dtgRemessa.ifbUpdate.Parameters.Add(new FbParameter("CP7", FbDbType.Double, 0, "cp7"));
            dtgRemessa.ifbUpdate.Parameters.Add(new FbParameter("CP28", FbDbType.Double, 0, "cp28"));
            dtgRemessa.ifbUpdate.Parameters.Add(new FbParameter("CP7C", FbDbType.Double, 0, "cp7c"));
            dtgRemessa.ifbUpdate.Parameters.Add(new FbParameter("CP28C", FbDbType.Double, 0, "cp28c"));
            dtgRemessa.ida.UpdateCommand = dtgRemessa.ifbUpdate;

            dtgRemessa.ifbInsert = new FbCommand("         insert into sgi_material_controlado (sgi_material_controlado_id,  " +
                                                                                   "  insumo_id,  " +
                                                                                    "  obra_id, " +
                                                                                   "  nf,  " +
                                                                                   "  remessa,  " +
                                                                                   "  data,  " +
                                                                                   "  qtde,  " +
                                                                                   "  splump,  " +
                                                                                   "  fck,  " +
                                                                                   "  numero_cp, " +
                                                                                   "  cp7,  " +
                                                                                   "  cp28,  " +
                                                                                   "  cp7c,  " +
                                                                                   "  cp28c) " +
                                              "  values (@sgi_material_controlado_id,  " +
                                              "          @insumo_id,  " +
                                              "          @obra_id, " +
                                              "          @nf,  " +
                                              "          @remessa,  " +
                                              "          @data,  " +
                                               "         @qtde,  " +
                                               "         @splump,  " +
                                               "         @fck,  " +
                                               "         @numero_cp,  " +
                                               "         @cp7,  " +
                                               "         @cp28, " +
                                               "         @cp7c, " +
                                               "         @cp28c)", sqlConnection.FbConexao);


            dtgRemessa.ifbInsert.Parameters.Add(new FbParameter("sgi_material_controlado_id", FbDbType.Integer, 0, "sgi_material_controlado_id"));
            dtgRemessa.ifbInsert.Parameters.Add(new FbParameter("insumo_id", FbDbType.Integer, 0, "insumo_id"));
            dtgRemessa.ifbInsert.Parameters.Add(new FbParameter("obra_id", FbDbType.Integer, 0, "obra_id"));
            dtgRemessa.ifbInsert.Parameters.Add(new FbParameter("NF", FbDbType.VarChar, 250, "nf"));
            dtgRemessa.ifbInsert.Parameters.Add(new FbParameter("REMESSA", FbDbType.VarChar, 250, "remessa"));
            dtgRemessa.ifbInsert.Parameters.Add(new FbParameter("DATA", FbDbType.Date, 0, "data"));
            dtgRemessa.ifbInsert.Parameters.Add(new FbParameter("QTDE", FbDbType.Double, 0, "qtde"));
            dtgRemessa.ifbInsert.Parameters.Add(new FbParameter("SPLUMP", FbDbType.Double, 0, "splump"));
            dtgRemessa.ifbInsert.Parameters.Add(new FbParameter("FCK", FbDbType.Double, 0, "fck"));
            dtgRemessa.ifbInsert.Parameters.Add(new FbParameter("NUMERO_CP", FbDbType.VarChar, 250, "numero_cp"));
            dtgRemessa.ifbInsert.Parameters.Add(new FbParameter("CP7", FbDbType.Double, 0, "cp7"));
            dtgRemessa.ifbInsert.Parameters.Add(new FbParameter("CP28", FbDbType.Double, 0, "cp28"));
            dtgRemessa.ifbInsert.Parameters.Add(new FbParameter("CP7C", FbDbType.Double, 0, "cp7c"));
            dtgRemessa.ifbInsert.Parameters.Add(new FbParameter("CP28C", FbDbType.Double, 0, "cp28c"));




            dtgRemessa.ida.InsertCommand = dtgRemessa.ifbInsert;

            dtgRemessa.ifbDelete = new FbCommand(" delete from SGI_MATERIAL_CONTROLADO where SGI_MATERIAL_CONTROLADO_id  = @SGI_MATERIAL_CONTROLADO_id ",
                                                 sqlConnection.FbConexao);


            dtgRemessa.ifbDelete.Parameters.Add(new FbParameter("SGI_MATERIAL_CONTROLADO_id", FbDbType.Integer, 0, "SGI_MATERIAL_CONTROLADO_id"));

            dtgRemessa.ida.DeleteCommand = dtgRemessa.ifbDelete;

            #endregion

            dtgRemessa.idt = new DataTable("REMESSA");
            dtgRemessa.ida.Fill(dtgRemessa.idt);
            ds.Tables.Add(dtgRemessa.idt);
            dtgRemessa.ibs = new NBindingSource();
            dtgRemessa.ibs.DataSource = ds;
            dtgRemessa.ibs.DataMember = ds.Tables["REMESSA"].TableName;
            dtgRemessa.DataSource = dtgRemessa.ibs;
            dtgRemessa.nomePreencherTabela = "REMESSA";
            dtgRemessa.dataMenber = "REMESSA";
            dtgRemessa.dsGrid = ds;



            DataRelation drApropriacaoInsumo = new DataRelation("RemessaPlanId",
                                                                ds.Tables["SGI"].Columns["PLAN_SERVICO_AMO_ID"],
                                                                ds.Tables["REMESSA_APRORIADA"].Columns["PLAN_SERVICO_AMO_ID"]);
            ds.Relations.Add(drApropriacaoInsumo);


            dtgSGIConcreto.ibs = new NBindingSource();

            dtgSGIConcreto.ibs.DataSource = ds;
            dtgSGIConcreto.ibs.DataMember = ds.Tables["SGI"].TableName;


            dtgSGIConcreto.DataSource = dtgSGIConcreto.ibs;
            dtgSGIConcreto.nomePreencherTabela = "SGI";
            dtgSGIConcreto.dataMenber = "SGI";

            dtgSGIConcreto.dsGrid = ds;


            dtgRemessaApropriada.ibs = new NBindingSource();
            dtgRemessaApropriada.ibs.DataSource = dtgSGIConcreto.ibs;
            dtgRemessaApropriada.ibs.DataMember = "RemessaPlanId";//ds.Tables["REMESSA_APRORIADA"].TableName;
            dtgRemessaApropriada.DataSource = dtgRemessaApropriada.ibs;

            dtgRemessaApropriada.nomePreencherTabela = "REMESSA_APRORIADA";
            dtgRemessaApropriada.dsGrid = ds;

            dtgSGIConcreto.PermitirFullSelect = true;

            dtgRemessaApropriada.TrataColunasAlinhamento();
            dtgRemessa.TrataColunasAlinhamento();
            dtgSGIConcreto.TrataColunasAlinhamento();
            TrataColunas("Código;Descrição;Dia;Volume(m3);Comp. projeto;Prof. concretada; Prof. escadvada;Vol. relatório;Vol. Concreteira; Sobreconsumo; Remessa;", ";", dtgSGIConcreto);
            TrataColunas("Código;Remessa;Posição;", ";", dtgRemessaApropriada);
            TrataColunas("Código;Insumo;Unid;NF;Remessa;Qtde;Splump;fck;Nº CP; 7 dias; 28 dias; 7 dias Contra prova; 28 dias Contra prova;", ";", dtgRemessa);
            TrataColunasLargura("80;150;50;80;80;80;80;80;80;80;80;", ";", dtgRemessa);



            CriarFBInserirRemessa(sqlConnection.FbConexao);

            bindingNavigator1.BindingSource = dtgSGIConcreto.ibs;
            bindingNavigator2.BindingSource = dtgRemessaApropriada.ibs;
            bindingNavigator3.BindingSource = dtgRemessa.ibs;*/
            ShowDialog();

        }

    }
}
