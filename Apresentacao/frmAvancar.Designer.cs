namespace Apresentacao
{
    partial class frmAvancar
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.panel1 = new System.Windows.Forms.Panel();
            this.splitter1 = new System.Windows.Forms.Splitter();
            this.panel2 = new System.Windows.Forms.Panel();
            this.panel3 = new System.Windows.Forms.Panel();
            this.panel4 = new System.Windows.Forms.Panel();
            this.nDataGridView1 = new Apresentacao.NDataGridView();
            this.nDataGridView2 = new Apresentacao.NDataGridView();
            this.btnAplicar = new System.Windows.Forms.Button();
            this.SERVICO_ID = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.SERVICO = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.UNID = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.DESCRICAO_CRITERIO = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PESO = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PESO_ACUMULADO = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.panel3.SuspendLayout();
            this.panel4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nDataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nDataGridView2)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.panel4);
            this.panel1.Controls.Add(this.panel3);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(602, 171);
            this.panel1.TabIndex = 0;
            // 
            // splitter1
            // 
            this.splitter1.Dock = System.Windows.Forms.DockStyle.Top;
            this.splitter1.Location = new System.Drawing.Point(0, 171);
            this.splitter1.Name = "splitter1";
            this.splitter1.Size = new System.Drawing.Size(602, 3);
            this.splitter1.TabIndex = 1;
            this.splitter1.TabStop = false;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.nDataGridView2);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(0, 174);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(602, 168);
            this.panel2.TabIndex = 2;
            // 
            // panel3
            // 
            this.panel3.Controls.Add(this.btnAplicar);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel3.Location = new System.Drawing.Point(0, 0);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(602, 39);
            this.panel3.TabIndex = 0;
            // 
            // panel4
            // 
            this.panel4.Controls.Add(this.nDataGridView1);
            this.panel4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel4.Location = new System.Drawing.Point(0, 39);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(602, 132);
            this.panel4.TabIndex = 3;
            // 
            // nDataGridView1
            // 
            this.nDataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.nDataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.SERVICO_ID,
            this.SERVICO,
            this.UNID});
            this.nDataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.nDataGridView1.Location = new System.Drawing.Point(0, 0);
            this.nDataGridView1.Name = "nDataGridView1";
            this.nDataGridView1.PermitirFullSelect = false;
            this.nDataGridView1.Size = new System.Drawing.Size(602, 132);
            this.nDataGridView1.TabIndex = 0;
            // 
            // nDataGridView2
            // 
            this.nDataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.nDataGridView2.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.DESCRICAO_CRITERIO,
            this.PESO,
            this.PESO_ACUMULADO});
            this.nDataGridView2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.nDataGridView2.Location = new System.Drawing.Point(0, 0);
            this.nDataGridView2.Name = "nDataGridView2";
            this.nDataGridView2.PermitirFullSelect = false;
            this.nDataGridView2.Size = new System.Drawing.Size(602, 168);
            this.nDataGridView2.TabIndex = 1;
            // 
            // btnAplicar
            // 
            this.btnAplicar.Location = new System.Drawing.Point(19, 7);
            this.btnAplicar.Name = "btnAplicar";
            this.btnAplicar.Size = new System.Drawing.Size(75, 23);
            this.btnAplicar.TabIndex = 0;
            this.btnAplicar.Text = "Aplicar";
            this.btnAplicar.UseVisualStyleBackColor = true;
            // 
            // SERVICO_ID
            // 
            this.SERVICO_ID.HeaderText = "Serviço";
            this.SERVICO_ID.Name = "SERVICO_ID";
            // 
            // SERVICO
            // 
            this.SERVICO.HeaderText = "Desc. serviço";
            this.SERVICO.Name = "SERVICO";
            this.SERVICO.Width = 250;
            // 
            // UNID
            // 
            this.UNID.HeaderText = "Unid";
            this.UNID.Name = "UNID";
            // 
            // DESCRICAO_CRITERIO
            // 
            this.DESCRICAO_CRITERIO.HeaderText = "Descrição";
            this.DESCRICAO_CRITERIO.Name = "DESCRICAO_CRITERIO";
            this.DESCRICAO_CRITERIO.Width = 250;
            // 
            // PESO
            // 
            this.PESO.HeaderText = "Peso";
            this.PESO.Name = "PESO";
            this.PESO.Width = 80;
            // 
            // PESO_ACUMULADO
            // 
            this.PESO_ACUMULADO.HeaderText = "Peso acumulado";
            this.PESO_ACUMULADO.Name = "PESO_ACUMULADO";
            this.PESO_ACUMULADO.Width = 80;
            // 
            // frmAvancar
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(602, 342);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.splitter1);
            this.Controls.Add(this.panel1);
            this.Name = "frmAvancar";
            this.Text = "frmAvancar";
            this.panel1.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.panel3.ResumeLayout(false);
            this.panel4.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.nDataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nDataGridView2)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Splitter splitter1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Panel panel4;
        private NDataGridView nDataGridView1;
        private System.Windows.Forms.Button btnAplicar;
        private NDataGridView nDataGridView2;
        private System.Windows.Forms.DataGridViewTextBoxColumn SERVICO_ID;
        private System.Windows.Forms.DataGridViewTextBoxColumn SERVICO;
        private System.Windows.Forms.DataGridViewTextBoxColumn UNID;
        private System.Windows.Forms.DataGridViewTextBoxColumn DESCRICAO_CRITERIO;
        private System.Windows.Forms.DataGridViewTextBoxColumn PESO;
        private System.Windows.Forms.DataGridViewTextBoxColumn PESO_ACUMULADO;
    }
}