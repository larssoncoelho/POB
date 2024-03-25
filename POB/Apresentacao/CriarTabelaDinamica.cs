using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace POB.Apresentacao
{
    
    public partial class CriarTabelaDinamica : Form
    {
        public string Ambiente
        {
            get
            {
                return textBox1.Text;
            }
        }
        public string Pavimento
        {
            get
            {
                return textBox2.Text;
            }
        }

        public bool Continuar { get; private set; }

        public CriarTabelaDinamica()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Continuar = true;
            Close();
        }
    }
}
