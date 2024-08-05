using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using rvt =  Autodesk.Revit.DB;
using rvtUI =  Autodesk.Revit.UI;
namespace POB.Apresentacao
{
    public partial class VerificarParametros : Form
    {
        public VerificarParametros(rvt.Document uidoc, List<rvt.Element> lista, List<ItemInformacaoElemento> itemInformacaoElementos)
        {
            InitializeComponent();
            UIDoc = uidoc;
            Lista = lista;
            ItemInformacaoElementos = itemInformacaoElementos;
            dataGridView1.DataSource = itemInformacaoElementos.OrderBy(x => x.Nome).ToList();
        }

        public rvt.Document UIDoc { get; private set; }
        public List<rvt.Element> Lista { get; private set; }
        public List<ItemInformacaoElemento> ItemInformacaoElementos { get; private set; }

        private void button1_Click(object sender, EventArgs e)
        {
            int id  = Convert.ToInt32( dataGridView1.CurrentRow.Cells[nameof(POB.ItemInformacaoElemento.Id)].Value.ToString());
            string valor =  dataGridView1.CurrentRow.Cells[nameof(POB.ItemInformacaoElemento.Valor)].Value.ToString();
            string novoCampo = textBox1.Text;
            var ele = UIDoc.GetElement(new rvt.ElementId(id));
            rvt.Transaction t = new rvt.Transaction(UIDoc);
            
            var par = Util.GetParameter(ele, novoCampo,
#if D24||D23
                rvt.SpecTypeId.String.Text,
#else
             rvt.ParameterType.Text,
#endif
                true, true);
            
            t.Start("t");

           
            par.Set(valor);

            t.Commit();
            t.Dispose();

        }

        private void button2_Click(object sender, EventArgs e)
        {

        }
    }
}
