namespace Funcoes
{
    partial class frmOpcaoSubGrafico
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
            this.chkDestacarMes = new System.Windows.Forms.CheckBox();
            this.chkDestacarCaminhoCritico = new System.Windows.Forms.CheckBox();
            this.dtpExecutadoMes = new System.Windows.Forms.DateTimePicker();
            this.dtpStatus = new System.Windows.Forms.DateTimePicker();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.dtpMostrarCriticoAte = new System.Windows.Forms.DateTimePicker();
            this.btnOk = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.chkDestacarPCP = new System.Windows.Forms.CheckBox();
            this.chkDestacarIniciado = new System.Windows.Forms.CheckBox();
            this.chkGradienteIniciado = new System.Windows.Forms.CheckBox();
            this.chkSomenteVistaAtiva = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // chkDestacarMes
            // 
            this.chkDestacarMes.AutoSize = true;
            this.chkDestacarMes.Location = new System.Drawing.Point(12, 12);
            this.chkDestacarMes.Name = "chkDestacarMes";
            this.chkDestacarMes.Size = new System.Drawing.Size(202, 17);
            this.chkDestacarMes.TabIndex = 0;
            this.chkDestacarMes.Text = "Destacar executado a partir da data?";
            this.chkDestacarMes.UseVisualStyleBackColor = true;
            // 
            // chkDestacarCaminhoCritico
            // 
            this.chkDestacarCaminhoCritico.AutoSize = true;
            this.chkDestacarCaminhoCritico.Location = new System.Drawing.Point(12, 35);
            this.chkDestacarCaminhoCritico.Name = "chkDestacarCaminhoCritico";
            this.chkDestacarCaminhoCritico.Size = new System.Drawing.Size(148, 17);
            this.chkDestacarCaminhoCritico.TabIndex = 1;
            this.chkDestacarCaminhoCritico.Text = "Destarcar caminho crítico";
            this.chkDestacarCaminhoCritico.UseVisualStyleBackColor = true;
            // 
            // dtpExecutadoMes
            // 
            this.dtpExecutadoMes.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpExecutadoMes.Location = new System.Drawing.Point(220, 12);
            this.dtpExecutadoMes.Name = "dtpExecutadoMes";
            this.dtpExecutadoMes.Size = new System.Drawing.Size(101, 20);
            this.dtpExecutadoMes.TabIndex = 2;
            // 
            // dtpStatus
            // 
            this.dtpStatus.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpStatus.Location = new System.Drawing.Point(12, 144);
            this.dtpStatus.Name = "dtpStatus";
            this.dtpStatus.Size = new System.Drawing.Size(101, 20);
            this.dtpStatus.TabIndex = 3;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 128);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(76, 13);
            this.label1.TabIndex = 5;
            this.label1.Text = "Data do status";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(15, 172);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(75, 13);
            this.label2.TabIndex = 7;
            this.label2.Text = "Mostrar crítica";
            // 
            // dtpMostrarCriticoAte
            // 
            this.dtpMostrarCriticoAte.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpMostrarCriticoAte.Location = new System.Drawing.Point(13, 188);
            this.dtpMostrarCriticoAte.Name = "dtpMostrarCriticoAte";
            this.dtpMostrarCriticoAte.Size = new System.Drawing.Size(101, 20);
            this.dtpMostrarCriticoAte.TabIndex = 6;
            // 
            // btnOk
            // 
            this.btnOk.Location = new System.Drawing.Point(151, 185);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(75, 23);
            this.btnOk.TabIndex = 8;
            this.btnOk.Text = "Ok";
            this.btnOk.UseVisualStyleBackColor = true;
            this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(232, 185);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 9;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // chkDestacarPCP
            // 
            this.chkDestacarPCP.AutoSize = true;
            this.chkDestacarPCP.Checked = true;
            this.chkDestacarPCP.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkDestacarPCP.Location = new System.Drawing.Point(12, 59);
            this.chkDestacarPCP.Name = "chkDestacarPCP";
            this.chkDestacarPCP.Size = new System.Drawing.Size(96, 17);
            this.chkDestacarPCP.TabIndex = 10;
            this.chkDestacarPCP.Text = "Destarcar PCP";
            this.chkDestacarPCP.UseVisualStyleBackColor = true;
            // 
            // chkDestacarIniciado
            // 
            this.chkDestacarIniciado.AutoSize = true;
            this.chkDestacarIniciado.Checked = true;
            this.chkDestacarIniciado.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkDestacarIniciado.Location = new System.Drawing.Point(12, 82);
            this.chkDestacarIniciado.Name = "chkDestacarIniciado";
            this.chkDestacarIniciado.Size = new System.Drawing.Size(111, 17);
            this.chkDestacarIniciado.TabIndex = 11;
            this.chkDestacarIniciado.Text = "Destarcar iniciado";
            this.chkDestacarIniciado.UseVisualStyleBackColor = true;
            // 
            // chkGradienteIniciado
            // 
            this.chkGradienteIniciado.AutoSize = true;
            this.chkGradienteIniciado.Location = new System.Drawing.Point(13, 106);
            this.chkGradienteIniciado.Name = "chkGradienteIniciado";
            this.chkGradienteIniciado.Size = new System.Drawing.Size(111, 17);
            this.chkGradienteIniciado.TabIndex = 12;
            this.chkGradienteIniciado.Text = "Gradiente iniciado";
            this.chkGradienteIniciado.UseVisualStyleBackColor = true;
            // 
            // chkSomenteVistaAtiva
            // 
            this.chkSomenteVistaAtiva.AutoSize = true;
            this.chkSomenteVistaAtiva.Location = new System.Drawing.Point(220, 59);
            this.chkSomenteVistaAtiva.Name = "chkSomenteVistaAtiva";
            this.chkSomenteVistaAtiva.Size = new System.Drawing.Size(96, 17);
            this.chkSomenteVistaAtiva.TabIndex = 13;
            this.chkSomenteVistaAtiva.Text = "Destarcar PCP";
            this.chkSomenteVistaAtiva.UseVisualStyleBackColor = true;
            // 
            // frmOpcaoSubGrafico
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(342, 250);
            this.Controls.Add(this.chkSomenteVistaAtiva);
            this.Controls.Add(this.chkGradienteIniciado);
            this.Controls.Add(this.chkDestacarIniciado);
            this.Controls.Add(this.chkDestacarPCP);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOk);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.dtpMostrarCriticoAte);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dtpStatus);
            this.Controls.Add(this.dtpExecutadoMes);
            this.Controls.Add(this.chkDestacarCaminhoCritico);
            this.Controls.Add(this.chkDestacarMes);
            this.Name = "frmOpcaoSubGrafico";
            this.Text = "Opções substituir gráficos";
            this.Load += new System.EventHandler(this.frmOpcaoSubGrafico_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public System.Windows.Forms.CheckBox chkDestacarMes;
        public System.Windows.Forms.CheckBox chkDestacarCaminhoCritico;
        public System.Windows.Forms.DateTimePicker dtpExecutadoMes;
        public System.Windows.Forms.DateTimePicker dtpStatus;
        public System.Windows.Forms.Label label1;
        public System.Windows.Forms.Label label2;
        public System.Windows.Forms.DateTimePicker dtpMostrarCriticoAte;
        public System.Windows.Forms.Button btnOk;
        public System.Windows.Forms.Button btnCancel;
        public System.Windows.Forms.CheckBox chkDestacarPCP;
        public System.Windows.Forms.CheckBox chkDestacarIniciado;
        public System.Windows.Forms.CheckBox chkGradienteIniciado;
        public System.Windows.Forms.CheckBox chkSomenteVistaAtiva;
    }
}