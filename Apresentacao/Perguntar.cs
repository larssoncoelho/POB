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
    public partial class Perguntar : NForm
    {
                 
        public Perguntar()
        {
            InitializeComponent();
        }
        public void  Inputar(string prompt, string titulo  )
        {
            this.Text = titulo;
            textBox1.Text = prompt;
            ShowDialog();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            resultadoProcura.vCampo.Clear();
            resultadoProcura.fResultadoProcura = true;
            resultadoProcura.vCampo.Add(textBox1.Text);
            Close();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            resultadoProcura.fResultadoProcura = false;
            resultadoProcura.vCampo.Clear();
            Close();
        }

        private void Perguntar_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
            {
                button3_Click(textBox1, null);
            }
        }
    }
}
