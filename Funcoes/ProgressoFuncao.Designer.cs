namespace Funcoes
{
    partial class ProgressoFuncao
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
            this.pgc = new System.Windows.Forms.ProgressBar();
            this.SuspendLayout();
            // 
            // pgc
            // 
            this.pgc.Location = new System.Drawing.Point(22, 12);
            this.pgc.Name = "pgc";
            this.pgc.Size = new System.Drawing.Size(471, 23);
            this.pgc.TabIndex = 1;
            this.pgc.Click += new System.EventHandler(this.pgc_Click);
            // 
            // ProgressoFuncao
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(505, 54);
            this.Controls.Add(this.pgc);
            this.Name = "ProgressoFuncao";
            this.Text = "ProgressoFuncao";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ProgressBar pgc;
    }
}