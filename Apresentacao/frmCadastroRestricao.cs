using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using AcessoBancoDados;
using FirebirdSql.Data;
using FirebirdSql.Data.FirebirdClient;

namespace Apresentacao
{
    public partial class frmCadastroRestricao : NForm
    {
        public DataSet ds = new DataSet("Tabelas");
        public NDataTable dtRestricoes = new NDataTable("Restricoes");
        public FbsqlConnection sqlConnection = new FbsqlConnection();
        public frmCadastroRestricao()
        {
            InitializeComponent();
        }
        public void DefinirElemento()
        {
            dtRestricoes.ifbcommandSelect.Connection = sqlConnection.FbConexao;
            dtRestricoes.ifbcommandSelect.CommandText = "  select r.restricao_id," +
                                                         "   r.restricao, "+
                                                          "  r.tipo_restricao_id, "+
                                                         "   r.tipo_estato_id," +
                                                         "   r.tipo_ativo_id, "+
                                                         "   r.cad, "+
                                                         "   r.data_cad, "+
                                                        "    r.alt, "+
                                                         "   r.data_alt," +
                                                        "    r.sicronizado, "+
                                                        "    r.sycrrestricao_id"+
                                                        "from restricao r "+
                                                       " inner join tipo t on r.tipo_restricao_id = t.tipo_id "+
                                                        "order by r.restricao";
                                            dtRestricoes.ida = new FbDataAdapter(dtRestricoes.ifbcommandSelect);
                                            dtRestricoes.ifbUpdate = new FbCommand("    update restricao "+
                                  "  set restricao = @restricao, "+
                                  "      tipo_restricao_id = @tipo_restricao_id, " +
                                   "     tipo_estato_id = @tipo_estato_id, " +
                                   "     tipo_ativo_id = @tipo_ativo_id, " +
                                   "     sicronizado = @sicronizado, " +
                                   "     sycrrestricao_id = @sycrrestricao_id " +
                                  "  where(restricao_id = @restricao_id))", sqlConnection.FbConexao);
            dtRestricoes.ifbUpdate.Parameters.Add(new FbParameter("restricao", FbDbType.Integer, 0, "plan_servico_amo_id"));
            dtRestricoes.ifbUpdate.Parameters.Add(new FbParameter("tipo_restricao_id", FbDbType.Double, 0, "prof_concretada"));
            dtRestricoes.ifbUpdate.Parameters.Add(new FbParameter("tipo_restricao_id", FbDbType.Double, 0, "prof_concretada"));
            dtRestricoes.ifbUpdate.Parameters.Add(new FbParameter("tipo_restricao_id", FbDbType.Double, 0, "prof_concretada"));

            dtRestricoes.ifbUpdate.Parameters.Add(new FbParameter("sicronizado", FbDbType.Double, 0, "prof_escavada"));
            dtRestricoes.ifbUpdate.Parameters.Add(new FbParameter("sycrrestricao_id", FbDbType.Double, 0, "vol_relatorio"));
            dtRestricoes.ifbUpdate.Parameters.Add(new FbParameter("restricao_id", FbDbType.Double, 0, "vol_relatorio"));

            dtRestricoes.ida.UpdateCommand = dtRestricoes.ifbUpdate;
            dtRestricoes.ida.Fill(dtRestricoes);
            ds.Tables.Add(dtRestricoes);
        }
        public void abreDados()
        {
           
        }
    }
}
