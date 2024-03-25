using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Apresentacao
{
    public partial class frmResultadoAvanco : NForm
    {
        public frmResultadoAvanco(DataTable dt)
        {
            InitializeComponent();
            this.nDataGridView1.DataSource = dt;
            DataGridViewCellStyle formatoNumerico = new DataGridViewCellStyle();
            formatoNumerico.Format = "N2";


            nDataGridView1.Columns["PSA_ID"].Width = 100;
            nDataGridView1.Columns["PSA_ID"].HeaderText = "Marca";
            nDataGridView1.Columns["ACAO"].Width = 300;
            nDataGridView1.Columns["ACAO"].HeaderText = "Ação";
            nDataGridView1.Columns["PERCENTUAL_SOLCITADO"].Width = 150;
            nDataGridView1.Columns["PERCENTUAL_SOLCITADO"].HeaderText = "% solicitado";
            nDataGridView1.Columns["PERCENTUAL_SOLCITADO"].DefaultCellStyle = formatoNumerico;

        }

        private void nDataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
