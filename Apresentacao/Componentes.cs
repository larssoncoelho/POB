using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using FirebirdSql.Data;
using FirebirdSql.Data.FirebirdClient;
using ObjetoTransferencia;
using Funcoes;
namespace Apresentacao
{
    
    public class NBindingSource : BindingSource
    {
        public DataSet ds11;
        public string icampoPrimaryKey;
        public string igen_id;
        public bool abrindo = false;
        public FbConnection ifbConnection;
        public NBindingSource() : base ()
        {

        }

        public bool Locate(string fieldName, object valor)
        {
            int id = Find(fieldName, valor);

            if (id != null)
            {
                Position = id;
                return true;
            }
            else return false;

        }

        public NBindingSource(IContainer container) : base (container)
        {
          
        }

        public NBindingSource(object dataSource, string dataMember) : base(dataSource, dataMember)
        {

        }
        public override object AddNew()
        {
            object obj;
            obj = base.AddNew();

            return obj;


            }
        protected override void OnCurrentChanged(EventArgs e)
        {
            
            base.OnCurrentChanged(e);
         
        }
        protected override void OnCurrentItemChanged(EventArgs e)
        {
            base.OnCurrentItemChanged(e);
            this.EndEdit();
        }

        public void GetDsInterno(ref DataSet ids1, object nbs)
        {

            int contador = 0;
            while (contador <= 10)
            {
                if (nbs is NBindingSource)
                {

                    nbs = (nbs as NBindingSource).DataSource;
                }
                if (nbs is DataSet)
                {
                    nbs = (nbs as DataSet);
                    ids1 = (nbs as DataSet);
                    return;
                }
                contador = contador + 1;

            }
        }

        public DataTable GetDataTable(DataSet ds, string dataMenber)
        {
            DataRelation dr1;
            DataTable dt1;
            foreach (DataRelation dataRelation in ds11.Relations)
            {
                if (dataRelation.RelationName == dataMenber)
                {
                    dr1 = dataRelation;
                    dt1 = dr1.ChildColumns[0].Table;
                    return dt1;
                }
            }

            return null;  
        }
     



        public  void EndEdit()
        {

       

            this.DataSource.GetType().ToString();
            base.EndEdit();
            if (!abrindo)
            {
                if (DataSource is DataSet)
                {
                    DataSet ds = (this.DataSource as DataSet);
                    NDataTable dt = ds.Tables[this.DataMember] as NDataTable;
                    dt.ida.Update(ds, dt.TableName);
                }
                if (DataSource is NBindingSource)
                {
                    GetDsInterno(ref ds11, DataSource);
                    NBindingSource nbsPai = (DataSource as NBindingSource);
                    NDataTable dt = ds11.Tables[GetDataTable(ds11, this.DataMember).TableName] as NDataTable;
                    dt.ida.Update(ds11, dt.TableName);
                }

            }
        }


    }

    public class NDataTable: DataTable
    {
        public FbCommand ifbcommandSelect =new FbCommand();
        public FbCommand ifbUpdate = new FbCommand();
        public FbCommand ifbDelete = new FbCommand();
        public FbCommand ifbInsert = new FbCommand();
        public FbDataAdapter ida;
        public NDataTable() : base ()
        {

        }
        public NDataTable(string tableName) : base (tableName)
        {

        }
        public NDataTable(string tableName, string tableNamespace) : base(tableName, tableNamespace)
        {
         
            
        }
        protected override void OnColumnChanged(DataColumnChangeEventArgs e)
        {
            base.OnColumnChanged(e);

        }
        protected override void OnTableNewRow(DataTableNewRowEventArgs e)
        {
            e.Row["CAUSA_CRITERIO_ID"] = FbSequence.Generator("GEN_CAUSA_CRITERIO_ID", ifbcommandSelect.Connection);
            base.OnTableNewRow(e);
        }
        protected override void OnRowChanged(DataRowChangeEventArgs e)
        {
            base.OnRowChanged(e);
        }

        public void AcceptChanges()
        {
            base.AcceptChanges();
           
            
        }
       

    }

    public class NDataGridView : DataGridView
    {
        private Boolean permitirFullSelect = false;

        public Boolean PermitirFullSelect
        {
            get { return permitirFullSelect; }
            set { permitirFullSelect = value; }
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
                if (column.ReadOnly)
                {
                    column.DefaultCellStyle.BackColor = Color.Silver;
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

            {

            }
            if (e.KeyCode == Keys.F5)
            {
                if (DataSource is NBindingSource)
                {
                    (DataSource as NBindingSource).ResetBindings(false);
                }
            }
 
            if((Control.ModifierKeys & Keys.Control) == Keys.Control)
            {
                if (e.KeyCode == Keys.Delete)
                {

                    if (DataSource is NBindingSource)
                    {
                        if (this.SelectedRows.Count > 1)
                        {
                            foreach (DataGridViewRow rSelected in SelectedRows)
                            {
                                this.Rows.Remove(rSelected);

                            }
                        }

                        else (DataSource as NBindingSource).RemoveCurrent();
                    }

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

                          MultiSelect = true;
                          SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                          CurrentRow.Selected = true;
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
        public List<string> Lista { get; set; }
        public string CampoMark { get; set; }
        public string CampoArea { get; set; }
        public ProgressoFuncao progresso;
        public DateTime diaProjecao;
        public Color corLinhaPCP;
        public Color corSuperficiePCP;
        public Color CorLinhaNaoIniciado;
        public Color CorSuperficieNaoIniciado;
        public Color CorLinhaCompleto;
        public Color CorSuperficieCompleto;
        public Color CorLinhaIniciado;
        public Color CorSuperficieIniciado;
        public Color CorLinhaExecutadoAposLimite;
        public Color CorSuperficieExecutadoAposLimite;
        public Color CorLinhaResticao;
        public Color CorSuperficieRestricao;
        public bool continuar;
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
                if (i <= dtg.Columns.Count -1)
                dtg.Columns[i].Width = Convert.ToInt32(campoAlt[i]);
            }

        }
       
    }
}
