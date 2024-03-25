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
    public partial class EscolherData : Form
    {
        public bool icontinuar = false;    
        public DialogResult GetDia (ref DateTime dia)
        {
       
            DialogResult dialogResult =  ShowDialog();
            dia = dtpDia.SelectionStart;
            return dialogResult;
            

        }

        public EscolherData()
        {
            InitializeComponent();
            this.AcceptButton = btnOk;
            this.CancelButton = btnCancelar;
            btnOk.DialogResult = DialogResult.OK;
            btnCancelar.DialogResult = DialogResult.Cancel;
            dtpDia.MaxDate = DateTime.Today.AddDays(90);
        }

        private void monthCalendar1_DateChanged(object sender, DateRangeEventArgs e)
        {

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
        public DialogResult inputar(ref DateTime valor,  string caption, ref bool continuar)
        {

           
            this.Text = caption;
            DialogResult dialogResult = ShowDialog();
            
            valor = dtpDia.SelectionStart;
            continuar = icontinuar;
            return dialogResult;


        }
        public DialogResult inputarComDuasDatas(ref DateTime valor, ref DateTime valor1, string caption, ref bool continuar)
        {


            this.Text = caption;
            this.Width = this.Width * 2;
            DialogResult dialogResult = ShowDialog();

            valor = dtpDia.SelectionStart;
            valor1 = dtpTermino.SelectionStart;
            continuar = icontinuar;
            
            return dialogResult;


        }

    }
}
