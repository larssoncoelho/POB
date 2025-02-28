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
    public partial class Perguntar : Form
    {
        public Perguntar(string v)
        {
            InitializeComponent();
            Continuar = false;
            Texto = v;
        }
        public bool Continuar
        { get; set; }
        public string Texto
        {
            get
            {
                return textBox1.Text;
            }
            set
            {
                textBox1.Text = value;
            }
        }
        private void Perguntar_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            Continuar = true;
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Continuar = false;
            this.Close();
        }
    }
}
