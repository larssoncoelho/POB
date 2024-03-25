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
using revitDB =  Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using appRevit =  Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Structure;


using AcessoBancoDados;
using ObjetoTransferencia;
namespace Apresentacao
{

    public partial class frmSGIDadosPsa : NForm
    {
        //datagridview1 is bound to this BindingList 
        BindingList<DataGridViewRow> object_bound_list1;

        //datagridview2 is bound to this BindingList 
        BindingList<DataGridViewRow> object_bound_list2;

        List<DataGridViewRow> selected_Object_list = new List<DataGridViewRow>();
        List<int> selected_pos_list = new List<int>();

        public string linha1;
        public string ielementos;
        public NBindingSource dsRemessa = new NBindingSource();
        public NBindingSource dsApropriacaoRemessa = new NBindingSource();
        public NBindingSource dsSGIConcreto = new NBindingSource();
        public NDataTable dtRemessa = new NDataTable("REMESSA");
        public NDataTable dtRemessaApropriada = new NDataTable("REMESSA_APRORIADA");
        public NDataTable dtSGIConcreto = new NDataTable("SGI");



        public List<revitDB.ElementId> lista = new List<revitDB.ElementId>();
        public appRevit.Application r;
        public revitDB.Element ele;
        public DataSet ds = new DataSet();
        public string internalDir;
        public StringBuilder sb = new StringBuilder();
        public UIApplication revitI;
        public revitDB.ElementId ele1;
        public FbCommand fbApropriarremessa;
        public FbsqlConnection sqlConnection = new FbsqlConnection();
        public string sqlmaterial_cntrl_sgi_dado()
        {
            StringBuilder sbi = new StringBuilder();
            sbi.Append(" select md.material_cntrl_sgi_dado_id, mc.remessa, ");
            sbi.Append(" md.posicao, ");
            sbi.Append("    md.plan_servico_amo_id,       ");
            sbi.Append(" md.sgi_material_controlado_id,  ");
            sbi.Append(" md.percent  ");
            sbi.Append("  from material_cntrl_sgi_dado md ");
            sbi.Append("  inner join  sgi_material_controlado mc on md.sgi_material_controlado_id = mc.sgi_material_controlado_id");
            
           sbi.Append(" where ");
            sbi.Append("  md.plan_servico_amo_id  in (select ");
            sbi.Append("  psaf.plan_servico_amo_filho_id  ");
            sbi.Append("  from plan_servico_amo_filho psaf where ");
            sbi.Append("  psaf.plan_servico_amo_pai_id in ( ");
            sbi.Append(ielementos);
            sbi.Append("  ))");
      
            return sbi.ToString();
        }
        public string sqlMaterial()
        {
            StringBuilder sbi = new StringBuilder();
            sbi.Append("  select mc.sgi_material_controlado_id, ");
            sbi.Append("   I.INSUMO,  ");
            sbi.Append("   I.UNID,  ");
            sbi.Append("     mc.nf,  ");
            sbi.Append("     mc.remessa,  ");
            sbi.Append("     mc.qtde,  ");
            sbi.Append("     mc.splump, ");
            sbi.Append("     mc.fck,  ");
            sbi.Append("    mc.numero_cp, ");
            sbi.Append("      mc.cp7,  ");
            sbi.Append("   mc.cp28,  ");
            sbi.Append("    mc.cp7c, ");
            sbi.Append("   mc.cp28c, ");
            sbi.Append("   mc.insumo_id,  ");

            sbi.Append("   mc.obra_id,  ");
            sbi.Append("     mc.data,  ");
            sbi.Append("      mc.data_cad ");
            sbi.Append("  from sgi_material_controlado mc ");
            sbi.Append("  INNER JOIN INSUMOB I ON MC.INSUMO_ID = I.INSUMO_ID ");

            sbi.Append("    order by mc.sgi_material_controlado_id desc ");

            return sbi.ToString();
        }
        public void sqlSgiEstaca(string elementos1, string elementos2)
        {

            sb.Clear();
            sb.Append(" select ");
            sb.Append("    psa.plan_servico_amo_id, ");
            sb.Append("   AMO.descricao, ");
            sb.Append("  psa.dia_realizado, ");
            sb.Append("     samo.volume, ");
            sb.Append("    samo.comprimento, ");
            sb.Append("       psa.prof_concretada, ");
            sb.Append(" psa.prof_escavada,  ");
            sb.Append("   psa.vol_relatorio,  ");
            sb.Append(" (select sum(mc1.qtde * md1.percent) / 100 ");
            sb.Append("    from material_cntrl_sgi_dado md1 ");
            sb.Append("      inner join sgi_material_controlado mc1 on md1.sgi_material_controlado_id =  mc1.sgi_material_controlado_id ");
            sb.Append("      where md1.plan_servico_amo_id = psa.plan_servico_amo_id  ) ");

            sb.Append("       vol_concreteira, ");

            sb.Append(" iif(   (samo.area_base * psa.prof_concretada) > 0,      (  (select sum(mc1.qtde * md1.percent) / 100 ");
            sb.Append("    from material_cntrl_sgi_dado md1 ");
            sb.Append("      inner join sgi_material_controlado mc1 on md1.sgi_material_controlado_id =  mc1.sgi_material_controlado_id ");
            sb.Append("      where md1.plan_servico_amo_id = psa.plan_servico_amo_id  ) /  (samo.area_base * psa.prof_concretada) - 1) * 100,null) sobre_consumo, ");


            sb.Append("   (select list(smc.remessa ||'='|| coalesce(cast(md.percent as numeric(18,1)),0),'->')");
            sb.Append("      from material_cntrl_sgi_dado md");
            sb.Append("       inner join sgi_material_controlado smc on md.sgi_material_controlado_id = smc.sgi_material_controlado_id ");
            sb.Append("       where md.plan_servico_amo_id = psa.plan_servico_amo_id)  remessa,");
            sb.Append("     psa.obs,  ");
            sb.Append("        psa.cad,  ");
            sb.Append("  psa.data_cad, ");
            sb.Append("      psa.alt,  ");
            sb.Append("     psa.data_alt, ");
            sb.Append("      psa.tipo_ativo_id,  ");
            sb.Append("       psa.tipo_estato_id,  ");
            sb.Append("       samo.id_bim,  ");
            sb.Append("   m.medicao, ");
            sb.Append("   b.bloco, ");
            sb.Append("     psa.grupo_plan_servico_id, ");
            sb.Append("       gps.grupo,  s.servico , ");

            sb.Append("     cast (null as double precision) vol_pecas_vinculadas ");
            sb.Append("  from plan_servico_amo psa ");
            sb.Append("    inner join servico_amo samo on psa.servico_amo_id = samo.servico_amo_id ");
            sb.Append("   inner join grupo_plan_servico gps on psa.grupo_plan_servico_id = gps.grupo_plan_servico_id ");
            sb.Append(" inner join servico s on samo.servico_id = s.servico_id ");
            sb.Append("  inner join ambiente_medicao_obra amo on samo.ambiente_medicao_obra_id  = amo.ambiente_medicao_bloco_id ");
            sb.Append("  inner join medicao_bloco mb on amo.medicao_bloco_id = mb.medicao_bloco_id ");
            sb.Append("  inner join bloco b on mb.bloco_id = b.bloco_id ");
            sb.Append("    inner join medicao m on mb.medicao_id = m.medicao_id ");
            sb.Append("  where ");
            sb.Append("   psa.filho is not null and psa.tipo_projecao is null");
            sb.Append("    and PSA.plan_servico_amo_id in   (   select  psaf.plan_servico_amo_filho_id ");
            sb.Append("    from  plan_servico_amo_filho psaf  ");

            sb.Append("   where    PSAf.plan_servico_amo_pai_id in (" + elementos1 + ") )");
            sb.Append("    and psa.qtde_realizada > 0  ");

             // MessageBox.Show(sb.ToString());

        }
        public frmSGIDadosPsa()
        {
            InitializeComponent();
           
        }
        public void Abrir(string dir, string elementos, UIApplication uiApp)
        {


            revitI = uiApp;
            ielementos = elementos;
            internalDir = dir;
            dtSGIConcreto.ifbcommandSelect = new FbCommand();
            
            sqlConnection.Diretorio = dir;
            sqlConnection.Ativo(true);
            
            
            dtSGIConcreto.ifbcommandSelect.Connection = sqlConnection.FbConexao;

            #region SGIConcreto
            sqlSgiEstaca(elementos, "");
            dtSGIConcreto.ifbcommandSelect.CommandText = sb.ToString();
            dtSGIConcreto.ida = new FbDataAdapter(dtSGIConcreto.ifbcommandSelect);
            dtSGIConcreto.ifbUpdate = new FbCommand("update plan_servico_amo " +
                                       " set prof_concretada = @prof_concretada,  "+
                                       "     prof_escavada = @prof_escavada, "+
                                       "     vol_relatorio = @vol_relatorio "+
                                       " where (plan_servico_amo_id = @plan_servico_amo_id)", sqlConnection.FbConexao) ;


            dtSGIConcreto.ifbUpdate.Parameters.Add(new FbParameter("plan_servico_amo_id", FbDbType.Integer, 0, "plan_servico_amo_id"));
            dtSGIConcreto.ifbUpdate.Parameters.Add(new FbParameter("prof_concretada", FbDbType.Double, 0, "prof_concretada"));
            dtSGIConcreto.ifbUpdate.Parameters.Add(new FbParameter("prof_escavada", FbDbType.Double, 0, "prof_escavada"));
            dtSGIConcreto.ifbUpdate.Parameters.Add(new FbParameter("vol_relatorio", FbDbType.Double, 0, "vol_relatorio"));

            dtSGIConcreto.ida.UpdateCommand = dtSGIConcreto.ifbUpdate;
            dtSGIConcreto.ida.Fill(dtSGIConcreto);
            ds.Tables.Add(dtSGIConcreto);
            #endregion

            #region remessaApropriada

            dtRemessaApropriada.ifbcommandSelect = new FbCommand(sqlmaterial_cntrl_sgi_dado(), sqlConnection.FbConexao);



            dtRemessaApropriada.ida = new FbDataAdapter(dtRemessaApropriada.ifbcommandSelect);
            dtRemessaApropriada.ifbUpdate = new FbCommand(" update material_cntrl_sgi_dado  "+
                                                      "    set posicao = @posicao,  "+
                                                        "      plan_servico_amo_id = @plan_servico_amo_id,  "+
                                                         "     sgi_material_controlado_id = @sgi_material_controlado_id,  "+
                                                         "     percent = @percent  "+
                                                        "  where (material_cntrl_sgi_dado_id =   " +
                                                    "  @material_cntrl_sgi_dado_id)", sqlConnection.FbConexao);


             dtRemessaApropriada.ifbUpdate.Parameters.Add(new FbParameter("POSICAO", FbDbType.Integer, 0, "POSICAO"));
             dtRemessaApropriada.ifbUpdate.Parameters.Add(new FbParameter("plan_servico_amo_id", FbDbType.Integer, 0, "plan_servico_amo_id"));
             dtRemessaApropriada.ifbUpdate.Parameters.Add(new FbParameter("sgi_material_controlado_id", FbDbType.Integer, 0, "sgi_material_controlado_id"));
             dtRemessaApropriada.ifbUpdate.Parameters.Add(new FbParameter("percent", FbDbType.Double, 0, "PERCENT"));
             dtRemessaApropriada.ifbUpdate.Parameters.Add(new FbParameter("material_cntrl_sgi_dado_id", FbDbType.Integer, 0, "material_cntrl_sgi_dado_id"));
            dtRemessaApropriada.ida.UpdateCommand = dtRemessaApropriada.ifbUpdate;

            dtRemessaApropriada.ifbDelete = new FbCommand(" delete from      material_cntrl_sgi_dado where material_cntrl_sgi_dado_id =   " +
                                                                "  @material_cntrl_sgi_dado_id", sqlConnection.FbConexao);
            dtRemessaApropriada.ifbDelete.Parameters.Add(new FbParameter("material_cntrl_sgi_dado_id", FbDbType.Integer, 0, "material_cntrl_sgi_dado_id"));
            dtRemessaApropriada.ida.DeleteCommand = dtRemessaApropriada.ifbDelete;
            dtRemessaApropriada.ida.Fill(dtRemessaApropriada);
            ds.Tables.Add(dtRemessaApropriada);
            #endregion 

            #region Remessa

             dtRemessa.ifbcommandSelect = new FbCommand(sqlMaterial(), sqlConnection.FbConexao);



             dtRemessa.ida = new FbDataAdapter(dtRemessa.ifbcommandSelect);
             dtRemessa.ifbUpdate = new FbCommand(" update sgi_material_controlado"+
                            " set insumo_id = @insumo_id,"+
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


             dtRemessa.ifbUpdate.Parameters.Add(new FbParameter("sgi_material_controlado_id", FbDbType.Integer, 0, "sgi_material_controlado_id"));
             dtRemessa.ifbUpdate.Parameters.Add(new FbParameter("insumo_id", FbDbType.Integer, 0, "insumo_id"));
             dtRemessa.ifbUpdate.Parameters.Add(new FbParameter("obra_id", FbDbType.Integer, 0, "obra_id"));
             dtRemessa.ifbUpdate.Parameters.Add(new FbParameter("NF", FbDbType.VarChar, 250, "nf"));
             dtRemessa.ifbUpdate.Parameters.Add(new FbParameter("REMESSA", FbDbType.VarChar, 250, "remessa"));
             dtRemessa.ifbUpdate.Parameters.Add(new FbParameter("DATA", FbDbType.Date, 0, "data"));
             dtRemessa.ifbUpdate.Parameters.Add(new FbParameter("QTDE", FbDbType.Double, 0, "qtde"));
             dtRemessa.ifbUpdate.Parameters.Add(new FbParameter("SPLUMP", FbDbType.Double, 0, "splump"));
             dtRemessa.ifbUpdate.Parameters.Add(new FbParameter("FCK", FbDbType.Double, 0, "fck"));
             dtRemessa.ifbUpdate.Parameters.Add(new FbParameter("NUMERO_CP", FbDbType.VarChar, 250, "numero_cp"));
             dtRemessa.ifbUpdate.Parameters.Add(new FbParameter("CP7", FbDbType.Double, 0, "cp7"));
             dtRemessa.ifbUpdate.Parameters.Add(new FbParameter("CP28", FbDbType.Double, 0, "cp28"));
             dtRemessa.ifbUpdate.Parameters.Add(new FbParameter("CP7C", FbDbType.Double, 0, "cp7c"));
             dtRemessa.ifbUpdate.Parameters.Add(new FbParameter("CP28C", FbDbType.Double, 0, "cp28c"));
             dtRemessa.ida.UpdateCommand = dtRemessa.ifbUpdate;

             dtRemessa.ifbInsert = new FbCommand("   insert into sgi_material_controlado (sgi_material_controlado_id, "+
                                     "  insumo_id, " +
                                     "  obra_id, " +
                                     "  nf, " +
                                     "  remessa, " +
                                     "  DATA, qtde, " +
                                     "  splump, " +
                                     "  fck, " +
                                     "  numero_cp, " +
                                     "  cp7, " +
                                     "  cp28, " +
                                     "  cp7c, "+
                                     "  cp28c) "+
                                     " values(@sgi_material_controlado_id, " +
                                      "        @insumo_id, " +
                                      "        @obra_id, " +
                                      "        @nf, " +
                                      "        @remessa, " +
               
                                      "        @DATA, @qtde, " +
                                      "        @splump, " +
                                      "        @fck, " +
                                      "        @numero_cp, " +
                                      "        @cp7, " +
                                      "        @cp28, " +
                                      "        @cp7c, " +
                                      "        @cp28c) " 
                                                , sqlConnection.FbConexao);

            
             dtRemessa.ifbInsert.Parameters.Add(new FbParameter("sgi_material_controlado_id", FbDbType.Integer, 0, "sgi_material_controlado_id"));
             dtRemessa.ifbInsert.Parameters.Add(new FbParameter("insumo_id", FbDbType.Integer, 0, "insumo_id"));
             dtRemessa.ifbInsert.Parameters.Add(new FbParameter("obra_id", FbDbType.Integer, 0, "obra_id"));
             dtRemessa.ifbInsert.Parameters.Add(new FbParameter("NF", FbDbType.VarChar, 250, "nf"));
             dtRemessa.ifbInsert.Parameters.Add(new FbParameter("REMESSA", FbDbType.VarChar, 250, "remessa"));
             dtRemessa.ifbInsert.Parameters.Add(new FbParameter("DATA", FbDbType.Date, 0, "data"));
             dtRemessa.ifbInsert.Parameters.Add(new FbParameter("QTDE", FbDbType.Double, 0, "qtde"));
             dtRemessa.ifbInsert.Parameters.Add(new FbParameter("SPLUMP", FbDbType.Double, 0, "splump"));
             dtRemessa.ifbInsert.Parameters.Add(new FbParameter("FCK", FbDbType.Double, 0, "fck"));
             dtRemessa.ifbInsert.Parameters.Add(new FbParameter("NUMERO_CP", FbDbType.VarChar, 250, "numero_cp"));
             dtRemessa.ifbInsert.Parameters.Add(new FbParameter("CP7", FbDbType.Double, 0, "cp7"));
             dtRemessa.ifbInsert.Parameters.Add(new FbParameter("CP28", FbDbType.Double, 0, "cp28"));
             dtRemessa.ifbInsert.Parameters.Add(new FbParameter("CP7C", FbDbType.Double, 0, "cp7c"));
             dtRemessa.ifbInsert.Parameters.Add(new FbParameter("CP28C", FbDbType.Double, 0, "cp28c"));
              
            
             dtRemessa.ida.InsertCommand = dtRemessa.ifbInsert;
             
             dtRemessa.ifbDelete = new FbCommand(" delete from SGI_MATERIAL_CONTROLADO where SGI_MATERIAL_CONTROLADO_id  = @SGI_MATERIAL_CONTROLADO_id ", 
                                                  sqlConnection.FbConexao);


             dtRemessa.ifbDelete.Parameters.Add(new FbParameter("SGI_MATERIAL_CONTROLADO_id", FbDbType.Integer, 0, "SGI_MATERIAL_CONTROLADO_id"));

             dtRemessa.ida.DeleteCommand = dtRemessa.ifbDelete;

          

             dtRemessa.ida.Fill(dtRemessa);
             ds.Tables.Add(dtRemessa);
            #endregion

            dsRemessa.DataSource = ds;
            dsRemessa.DataMember = ds.Tables["REMESSA"].TableName;
            dtgRemessa.DataSource = dsRemessa;
            dsRemessa.icampoPrimaryKey = "sgi_material_controlado_id";
            dsRemessa.igen_id = "gen_sgi_material_controlado_id";
            dsRemessa.ifbConnection = sqlConnection.FbConexao;

            DataRelation drApropriacaoInsumo = new DataRelation("RemessaPlanId",
                                                                ds.Tables["SGI"].Columns["PLAN_SERVICO_AMO_ID"],
                                                                ds.Tables["REMESSA_APRORIADA"].Columns["PLAN_SERVICO_AMO_ID"]);
            ds.Relations.Add(drApropriacaoInsumo);
           

            
            dsSGIConcreto.DataSource = ds;
            dsSGIConcreto.DataMember = ds.Tables["SGI"].TableName;
            dtgSGIConcreto.DataSource = dsSGIConcreto;
            // dtgSGIConcreto.nomePreencherTabela = "SGI";
            // dtgSGIConcreto.dataMenber = "SGI";

            //dtgSGIConcreto.dsGrid = ds;


            dsApropriacaoRemessa.DataSource = dsSGIConcreto;
            dsApropriacaoRemessa.DataMember = "RemessaPlanId";//ds.Tables["REMESSA_APRORIADA"].TableName;
            dtgRemessaApropriada.DataSource = dsApropriacaoRemessa;
           
           /// dtgRemessaApropriada.nomePreencherTabela = "REMESSA_APRORIADA";
           // dtgRemessaApropriada.dsGrid = ds;

            dtgSGIConcreto.PermitirFullSelect = true;

            dtgSGIConcreto.Columns["vol_concreteira"].ReadOnly = true;

            dtgRemessaApropriada.TrataColunasAlinhamento();
            dtgRemessa.TrataColunasAlinhamento();
            dtgSGIConcreto.TrataColunasAlinhamento();

            TrataColunas("Código;Descrição;Dia;Volume(m3);Comp. projeto;Prof. concretada; Prof. escavada;Vol. relatório;Vol. Concreteira; Sobreconsumo (%); Remessa;", ";", dtgSGIConcreto);
            TrataColunas("Código;Remessa;Posição;CódigoPlan;Código remessa;Percentual;_;", ";", dtgRemessaApropriada);
            TrataColunas("Código;Insumo;Unid;NF;Remessa;Qtde;Splump;fck;Nº CP; 7 dias; 28 dias; 7 dias Contra prova; 28 dias Contra prova;_;", ";", dtgRemessa);

            TrataColunasLargura("80;150;50;80;80;80;80;80;80;80;80;",";", dtgRemessa);
            TrataColunasLargura("80;150;50;80;80;80;80;_;", ";", dtgRemessaApropriada);
            TrataColunasLargura("80;150;50;80;80;80;80;80;80;80;80;", ";", dtgSGIConcreto);



            CriarFBInserirRemessa(sqlConnection.FbConexao);

            bindingNavigator1.BindingSource = dsSGIConcreto;
            bindingNavigator2.BindingSource = dsApropriacaoRemessa;
            bindingNavigator3.BindingSource = dsRemessa;
            ShowDialog();

        }

        private void dtgSGIConcreto_RowEnter(object sender, DataGridViewCellEventArgs e)
        {


        }

        private void bindingNavigator1_Click(object sender, EventArgs e)
        {
            DataTable t = new DataTable();
     //   
        }

        private void bindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
       
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void bindingNavigator1_RefreshItems(object sender, EventArgs e)
        {

        }

        public void CriarFBInserirRemessa(FbConnection fb)
        {

            fbApropriarremessa = new FbCommand("insert into material_cntrl_sgi_dado (plan_servico_amo_id, " +
                                                                     "   sgi_material_controlado_id, PERCENT)   " +
                                                                           " values (@plan_servico_amo_id, @SGI_MATERIAL_CONTROLADO_ID, @PERCENT)", fb);
            fbApropriarremessa.Parameters.Add("plan_servico_amo_id", FbDbType.Integer, 0, "plan_servico_amo_id");
            fbApropriarremessa.Parameters.Add("SGI_MATERIAL_CONTROLADO_ID", FbDbType.Integer, 0, "SGI_MATERIAL_CONTROLADO_ID");
            fbApropriarremessa.Parameters.Add("PERCENT", FbDbType.Double, 0, "PERCENT");

        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            FrmProcurar procurar = new FrmProcurar();
            StringBuilder sb = new StringBuilder();

            sb.Append("    select ");
            sb.Append("   smc.sgi_material_controlado_id,  ");
            sb.Append("   smc.remessa,  ");
            sb.Append("   smc.nf,  ");
            sb.Append("   smc.qtde, ");
            sb.Append("  ( SELECT  COUNT(MD.material_cntrl_sgi_dado_id) ");
            sb.Append("     FROM material_cntrl_sgi_dado MD ");
            sb.Append("     WHERE MD.sgi_material_controlado_id = SMC.sgi_material_controlado_id)  QTDE_APROPRIADO, ");

            sb.Append("   smc.splump, ");
            sb.Append("   smc.fck, ");
            sb.Append("   smc.numero_cp , ");
            sb.Append("   smc.cp7,  ");
            sb.Append("   smc.cp28, ");
            sb.Append("   smc.cp7c, ");
            sb.Append("   smc.cp28c, ");

            sb.Append("   smc.insumo_id,  ");
            sb.Append("   i.insumo, ");
            sb.Append("   i.unid, ");
            sb.Append("   smc.obra_id,  ");
            sb.Append("   smc.data,  ");
            sb.Append("   smc.data_cad  ");
            sb.Append(" from sgi_material_controlado smc ");
            sb.Append("  inner join insumob i on smc.insumo_id = i.insumo_id ");
            //   sb.Append("  where smc.tipo_estato_id = 287");

            sb.Append(" order by ");
            sb.Append(" smc.sgi_material_controlado_id desc ");

            resultadoProcura = procurar.Pesquisar(internalDir, "Inserir remessa",
                                      sb.ToString(), "Código;Remessa;NF;Qtde(m3); Qtde aprop.;",
                                                    "80;80;80;80;80;");
            Perguntar p1 = new Perguntar();
            if (resultadoProcura.fResultadoProcura)
                foreach (DataGridViewRow dr in dtgSGIConcreto.SelectedRows)
                {
                    p1.Inputar("Digite o percentual", "% de consumno do elemento: " + dr.Cells["DESCRICAO"].Value.ToString());
                    fbApropriarremessa.Parameters["PLAN_SERVICO_AMO_ID"].Value = Convert.ToInt32(dr.Cells["PLAN_SERVICO_AMO_ID"].Value);
                    fbApropriarremessa.Parameters["SGI_MATERIAL_CONTROLADO_ID"].Value = Convert.ToInt32(resultadoProcura.vCampo[0]);
                    fbApropriarremessa.Parameters["PERCENT"].Value = Convert.ToDouble(p1.resultadoProcura.vCampo[0]);
                    //MessageBox.Show(dr.Cells["PLAN_SERVICO_AMO_ID"].Value.ToString() + "\n " + resultadoProcura.vCampo[0].ToString() + "\n" + p1.resultadoProcura.vCampo[0].ToString());
                    fbApropriarremessa.ExecuteNonQuery();
                }
            dsApropriacaoRemessa.ResetBindings(false);
        }







        private void toolStripButton4_Click(object sender, EventArgs e)
        {

       /*   DataRow dr =  dtgSGIConcreto.idt.NewRow();
            dr["PLAN_SERVICO_AMO_ID"] = 205555;
            dtgSGIConcreto.idt.Rows.Add(dr);
                
                lista.Clear();

            try
            {
                revitI.ActiveUIDocument.Selection.SetElementIds(new List<ElementId>());
                ElementId eleId = new ElementId(Convert.ToInt32(dtgSGIConcreto.CurrentRow.Cells["ID_BIM"].Value.ToString()));
                lista.Add(eleId);
                revitI.ActiveUIDocument.Selection.SetElementIds(lista);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            */
        }

        private void dtpDate_ValueChanged(object sender, EventArgs e)
        {

        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {

        }

        private void bindingNavigatorAddNewItem1_Click(object sender, EventArgs e)
        {

        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
        }

        private void toolStripButton9_Click(object sender, EventArgs e)
        {
        }

        private void dtgRemessa_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
       
          
        }

        private void toolStripButton2_Click_1(object sender, EventArgs e)
        {

        }

        private void dtgSGIConcreto_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
        }

        private void dtgSGIConcreto_NewRowNeeded(object sender, DataGridViewRowEventArgs e)
        {

        }

        private void dtgSGIConcreto_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {

        }

        private void bindingNavigatorAddNewItem_Click(object sender, EventArgs e)
        {
         }

        private void toolStripButton3_Click_1(object sender, EventArgs e)
        {

            /*dtgRemessa.ibs.AddNew();

            MessageBox.Show("teste inserir");

            dtgRemessa.ibs.CancelEdit();

            MessageBox.Show("teste remover");

            dtgRemessa.ibs.RemoveCurrent();
            */
            FrmProcurar procurar = new FrmProcurar();
            StringBuilder sb = new StringBuilder();

            sb.Append("    select ");
            sb.Append("   INSUMO_ID,  ");
            sb.Append("   INSUMO,  ");
            sb.Append("   UNID, data_cad  ");
            sb.Append(" from INSUMOB ");
            resultadoProcura = procurar.Pesquisar(internalDir, "Escolher insumo",
                                      sb.ToString(), "Insumo; Desc. Insumo;Unid;Dt. Cadastro;",
                                                    "80;250;30;60;");
  
            if (resultadoProcura.fResultadoProcura)
            {
                 dsRemessa.EndEdit();
                
                DataRow r;
                r = dtRemessa.NewRow();
                r[dsRemessa.icampoPrimaryKey] = FbSequence.Generator(dsRemessa.igen_id, dsRemessa.ifbConnection);
                r["insumo_id"] = resultadoProcura.vCampo[0];
                r["insumo"] = resultadoProcura.vCampo[1];
                r["Unid"] = resultadoProcura.vCampo[2];
                dtRemessa.Rows.InsertAt(r, 0);
                dsRemessa.EndEdit();

                dsRemessa.Find("sgi_material_controlado_id", r[dsRemessa.icampoPrimaryKey]);
            }


            //dtgRemessa.ibs.AddNew();
           
        }


        private void dtgSGIConcreto_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dtgRemessa_MouseMove(object sender, MouseEventArgs e)
        {
            if ((e.Button & MouseButtons.Left) == MouseButtons.Left)
            {
                // Proceed with the drag and drop, passing in the list item.                   
                DragDropEffects dropEffect = dtgSGIConcreto.DoDragDrop(selected_Object_list,
                DragDropEffects.Move);
            }
        }

        private void dtgRemessa_MouseDown(object sender, MouseEventArgs e)
        {
            
                // Get the index of the item the mouse is below.
                int rowIndexFromMouseDown = dtgSGIConcreto.HitTest(e.X, e.Y).RowIndex;

                //if shift key is not pressed
                if (Control.ModifierKeys != Keys.Shift && Control.ModifierKeys != Keys.Control)
                {
                    //if row under the mouse is not selected
                    if (!selected_pos_list.Contains(rowIndexFromMouseDown) && rowIndexFromMouseDown > 0)
                    {
                        //if there only one row selected
                        if (dtgSGIConcreto.SelectedRows.Count == 1)
                        {
                            //select the row below the mouse
                            dtgSGIConcreto.ClearSelection();
                            dtgSGIConcreto.Rows[rowIndexFromMouseDown].Selected = true;
                        }
                    }
                }

                //clear the selection lists
                selected_Object_list.Clear();
                selected_pos_list.Clear();

                //add the selected objects
                foreach (DataGridViewRow row in dtgSGIConcreto.SelectedRows)
                {
                    selected_Object_list.Add(object_bound_list1[row.Index]);
                    selected_pos_list.Add(row.Index);
                }
            
        }

        private void dtgSGIConcreto_DragOver(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.None;
        }

        private void dtgSGIConcreto_DragEnter(object sender, DragEventArgs e)
        {

        }

        private void dtgSGIConcreto_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Effect == DragDropEffects.None)
            {
                MessageBox.Show("sd s");
            }
        }

        private void dtgRemessa_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dtgSGIConcreto_CellEndEdit_1(object sender, DataGridViewCellEventArgs e)
        {
            dsSGIConcreto.EndEdit();
        }

        private void dtgSGIConcreto_Scroll(object sender, ScrollEventArgs e)
        {
        /*    if (e.ScrollOrientation = ScrollOrientation.VerticalScroll)
            {

            }*/
        }

        private void toolStripButton4_Click_1(object sender, EventArgs e)
        {

        }

        private void toolStripButton1_Click_1(object sender, EventArgs e)
        {

        }
    }
}
