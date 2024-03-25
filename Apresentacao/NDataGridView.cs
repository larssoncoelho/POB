using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.ComponentModel;

using System.Data;
using FirebirdSql.Data;
using FirebirdSql.Data.FirebirdClient;
using AcessoBancoDados;
using ObjetoTransferencia;
namespace Apresentacao
{
    
    public class NBindingSource : BindingSource
    {

        public NBindingSource() : base ()
        {

        }

        public NBindingSource(IContainer container) : base (container)
        {
           
        }

        public NBindingSource(object dataSource, string dataMember) : base(dataSource, dataMember)
        {
        
        }
        protected override void OnAddingNew(AddingNewEventArgs e)
        {
            DataTable dt;

         
            
        }
        
        
    }

    
    public class NDataGridView : DataGridView
    {
        public string dataMenber;
        public string nomePreencherTabela;

        public DataSet dsGrid;
        public FbDataReader idr;
        public FbDataAdapter ida;
        public DataTable idt;
        public NBindingSource ibs;
        public FbCommand ifbcommandSelect;
        public FbCommand ifbUpdate;
        public FbCommand ifbDelete;
        public FbCommand ifbInsert;
        private Boolean permitirFullSelect = false;

        public void Atualizar()
        {
            dsGrid.Tables[nomePreencherTabela].Clear();
            idt.Clear();
            ida.Fill(idt);
           // ibs.DataSource = dsGrid;
           // ibs.DataMember = dsGrid.Tables[dataMenber].TableName;
            DataSource = ibs;
        }



        public Boolean PermitirFullSelect
        {
            get { return permitirFullSelect; }
            set { permitirFullSelect = value; }
        }

        public void ApplyUpdates()
        {
            ida.Update(dsGrid, nomePreencherTabela);
        }
        protected void Ordenar()
        {
            DataGridViewColumn newColumn =    Columns.GetColumnCount(
            DataGridViewElementStates.Selected) == 1 ?    SelectedColumns[0] : null;

            DataGridViewColumn oldColumn = SortedColumn;
            ListSortDirection direction;

            // If oldColumn is null, then the DataGridView is not currently sorted.
            if (oldColumn != null)
            {
                // Sort the same column again, reversing the SortOrder.
                if (oldColumn == newColumn &&
                    SortOrder == SortOrder.Ascending)
                {
                    direction = ListSortDirection.Descending;
                }
                else
                {
                    // Sort a new column and remove the old SortGlyph.
                    direction = ListSortDirection.Ascending;
                    oldColumn.HeaderCell.SortGlyphDirection = SortOrder.None;
                }
            }
            else
            {
                direction = ListSortDirection.Ascending;
            }

            // If no column has been selected, display an error dialog  box.
            if (newColumn == null)
            {
                MessageBox.Show("Select a single column and try again.",
                    "Error: Invalid Selection", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            else
            {
                Sort(newColumn, direction);
                newColumn.HeaderCell.SortGlyphDirection =
                    direction == ListSortDirection.Ascending ?
                    SortOrder.Ascending : SortOrder.Descending;
            }
        }
        protected override void InitLayout()
        {
            base.InitLayout();
           
        }
        public void TrataColunasAlinhamento()
        {
            foreach (DataGridViewColumn column in Columns)
            {
                if (column.ValueType == typeof(Int32))
                {
                    column.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                }
                if (column.ValueType == typeof(Double))
                {
                    column.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    column.DefaultCellStyle.Format = "#,0.00";

                }

            }
     
        }
        protected override void OnCellClick(DataGridViewCellEventArgs e)
        {
            base.OnCellClick(e);

        }
        protected override void OnKeyDown(KeyEventArgs e)
        {
            base.OnKeyDown(e);
                    
            if (e.KeyCode == Keys.F4)
                Ordenar();
            if (e.KeyCode == Keys.F5) Atualizar(); 
 
            if((Control.ModifierKeys & Keys.Control) == Keys.Control)
            {
                if (e.KeyCode == Keys.Delete)
                {
                    this.Rows.Remove(this.CurrentRow);
                    ApplyUpdates();
                }
            }
        }


        protected override void OnMouseDown(MouseEventArgs e)
        {
                base.OnMouseDown(e);
                if((Control.ModifierKeys & Keys.Control) == Keys.Control)
                    if (PermitirFullSelect)
                    {
                        if (e.Button == MouseButtons.Left)
                        {

                            this.MultiSelect = true;
                            this.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                        }
                        else
                        {
                            this.MultiSelect = false;
                            this.SelectionMode = DataGridViewSelectionMode.CellSelect;

                        }
                    }
        }


    }



    
    public static class FbSequence
    {
        public static FbCommand fbCommand;
         public static int Generator(string genId, FbConnection fbConnection)
        {
            fbCommand = new FbCommand();
            fbCommand.Connection = fbConnection;
   
         fbCommand.CommandText  = "select gen_id("+genId+",1)  from rdb$database";
         return Convert.ToInt32(fbCommand.ExecuteScalar());
        }

    }



    public class NForm: Form
    {




        public ResultadoProcura resultadoProcura = new ResultadoProcura();
        public void TrataColunas(string texto, string caracter, NDataGridView dtg)
        {
            string[] campoAlt = new string[1];
            int y = 0;

            string a = "";
            string b = "";

            for (int x = 0; x < texto.Length; x++)
            {
                a = texto.Substring(x, 1);
                if (a == caracter)
                {
                    y = y + 1;
                    Array.Resize(ref campoAlt, campoAlt.Length + 1);
                    campoAlt[y - 1] = b;
                    b = "";
                }
                else
                {
                    b = b + a;
                }
            }

            for (int i = 0; i < y - 1; i++)
            {
                dtg.Columns[i].HeaderText = campoAlt[i];
            }

        }

        public void TrataColunasLargura(string texto, string caracter, NDataGridView dtg)
        {
            string[] campoAlt = new string[1];
            int y = 0;

            string a = "";
            string b = "";

            for (int x = 0; x < texto.Length; x++)
            {
                a = texto.Substring(x, 1);
                if (a == caracter)
                {
                    y = y + 1;
                    Array.Resize(ref campoAlt, campoAlt.Length + 1);
                    campoAlt[y - 1] = b;
                    b = "";
                }
                else
                {
                    b = b + a;
                }
            }

            for (int i = 0; i < y - 1; i++)
            {
                dtg.Columns[i].Width = Convert.ToInt32(campoAlt[i]);
            }

        }
       
    }
}
