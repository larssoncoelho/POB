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
    public partial class frmOpcaoSubGrafico : Form
    {
        public bool resultado = false;

        public frmOpcaoSubGrafico()
        {
            InitializeComponent();
            dtpStatus.Value = DateTime.Today;
            dtpExecutadoMes.Value = DateTime.Today.AddDays(-30);
            dtpMostrarCriticoAte.Value = DateTime.Today.AddDays(90);

        }

        private void frmOpcaoSubGrafico_Load(object sender, EventArgs e)
        {

        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            resultado = true;
            Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            resultado = false;
            Close();
        }
    }
}
