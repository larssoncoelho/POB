using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Funcoes
{
    public partial class frmCadCausa : Form
    {
        public List<string> Causas = new List<string>();
        public DataSet ds = new DataSet("tabelas");
        public DataTable dt = new DataTable("Causas");
        public BindingSource bs = new BindingSource();
        public bool icontinuar;
        public frmCadCausa()
        {
            InitializeComponent();
            dt.Columns.Add("Causas", typeof(string));
            ds.Tables.Add(dt);
            bs.DataSource = ds;
            bs.DataMember = dt.TableName;
            dataGridView1.DataSource = bs;
            bindingNavigator1.BindingSource = bs;
            this.AcceptButton = btnOk;
            this.CancelButton = btnCancelar;
            btnOk.DialogResult = DialogResult.OK;
            btnCancelar.DialogResult = DialogResult.Cancel;

        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            icontinuar = true;
            this.Close();
        }

        private void btnCancelar_Click(object sender, EventArgs e)
        {
            icontinuar = false;
            this.Close();
        }
    }
}
