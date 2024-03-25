using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace POB
{
    public partial class FormRevestirPilar : Form
    {
        public FormRevestirPilar()
        {
            InitializeComponent();
        }
        public void PreenchecmbWallType(List<string> lista)
        {
            cmbWallType.DataSource = lista;

        }
        public string WallType
        {
            get
            {
                return cmbWallType.SelectedItem.ToString();
            }
        }
        public double DeslocamentoBase
        {
            get
            {
                double valor;
                if (double.TryParse(txtDeslocamentoBase.Text, out valor))
                    return valor;
                else return 100;
                
            }
        }
        public double AlturaDesconectada
        {
            get
            {
                double valor;
                if (double.TryParse(txtDeslocamentoTopo.Text, out valor))
                    return valor;
                else return 100;

            }
        }


        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void FormRevestirPilar_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            this.Close();

        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();

        }
    }
}
