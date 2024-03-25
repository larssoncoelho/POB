namespace Apresentacao
{
    partial class frmResultadoAvanco
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
            this.nDataGridView1 = new Apresentacao.NDataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.nDataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // nDataGridView1
            // 
            this.nDataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.nDataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.nDataGridView1.Location = new System.Drawing.Point(0, 0);
            this.nDataGridView1.Name = "nDataGridView1";
            this.nDataGridView1.PermitirFullSelect = false;
            this.nDataGridView1.Size = new System.Drawing.Size(605, 239);
            this.nDataGridView1.TabIndex = 0;
            this.nDataGridView1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.nDataGridView1_CellContentClick);
            // 
            // frmResultadoAvanco
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(605, 239);
            this.Controls.Add(this.nDataGridView1);
            this.Name = "frmResultadoAvanco";
            this.Text = "frmResultadoAvanco";
            ((System.ComponentModel.ISupportInitialize)(this.nDataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private NDataGridView nDataGridView1;
    }
}