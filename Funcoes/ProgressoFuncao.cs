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
    public partial class ProgressoFuncao : Form
    {
        public void Incrementar()
        {
            pgc.PerformStep();
        }
        public ProgressoFuncao(int total)
        {
            InitializeComponent();
            pgc.Maximum = total;
            pgc.Minimum = 1;
            pgc.Step = 1;
           
        }

        private void pgc_Click(object sender, EventArgs e)
        {

        }
    }
}
