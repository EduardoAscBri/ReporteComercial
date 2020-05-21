namespace ReporteComercial.Forms
{
    partial class ReporteVentas
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ReporteVentas));
            this.dtpFechaInicial = new System.Windows.Forms.DateTimePicker();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.dtpFechaFinal = new System.Windows.Forms.DateTimePicker();
            this.bttAcept = new System.Windows.Forms.Button();
            this.bttSalir = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.cmbDoctoCargo = new System.Windows.Forms.ComboBox();
            this.cmbDoctoAbono = new System.Windows.Forms.ComboBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.label5 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // dtpFechaInicial
            // 
            this.dtpFechaInicial.CalendarFont = new System.Drawing.Font("Segoe UI Semibold", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dtpFechaInicial.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dtpFechaInicial.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpFechaInicial.Location = new System.Drawing.Point(326, 163);
            this.dtpFechaInicial.Name = "dtpFechaInicial";
            this.dtpFechaInicial.Size = new System.Drawing.Size(264, 25);
            this.dtpFechaInicial.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Segoe UI", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.LightGray;
            this.label1.Location = new System.Drawing.Point(393, 125);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(118, 25);
            this.label1.TabIndex = 1;
            this.label1.Text = "Fecha inicial";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Segoe UI", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.LightGray;
            this.label2.Location = new System.Drawing.Point(395, 213);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(106, 25);
            this.label2.TabIndex = 3;
            this.label2.Text = "Fecha final";
            // 
            // dtpFechaFinal
            // 
            this.dtpFechaFinal.CalendarFont = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dtpFechaFinal.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dtpFechaFinal.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpFechaFinal.Location = new System.Drawing.Point(326, 254);
            this.dtpFechaFinal.Name = "dtpFechaFinal";
            this.dtpFechaFinal.Size = new System.Drawing.Size(264, 25);
            this.dtpFechaFinal.TabIndex = 2;
            // 
            // bttAcept
            // 
            this.bttAcept.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.bttAcept.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.bttAcept.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.bttAcept.Location = new System.Drawing.Point(352, 300);
            this.bttAcept.Name = "bttAcept";
            this.bttAcept.Size = new System.Drawing.Size(120, 32);
            this.bttAcept.TabIndex = 5;
            this.bttAcept.Text = "Aceptar";
            this.bttAcept.UseVisualStyleBackColor = false;
            this.bttAcept.Click += new System.EventHandler(this.bttAcept_Click);
            // 
            // bttSalir
            // 
            this.bttSalir.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.bttSalir.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.bttSalir.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.bttSalir.Location = new System.Drawing.Point(485, 300);
            this.bttSalir.Name = "bttSalir";
            this.bttSalir.Size = new System.Drawing.Size(120, 32);
            this.bttSalir.TabIndex = 6;
            this.bttSalir.Text = "Salir";
            this.bttSalir.UseVisualStyleBackColor = false;
            this.bttSalir.Click += new System.EventHandler(this.bttSalir_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Segoe UI", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.Color.LightGray;
            this.label3.Location = new System.Drawing.Point(50, 125);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(200, 25);
            this.label3.TabIndex = 7;
            this.label3.Text = "Documento de cargo";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Segoe UI", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.ForeColor = System.Drawing.Color.LightGray;
            this.label4.Location = new System.Drawing.Point(50, 213);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(207, 25);
            this.label4.TabIndex = 8;
            this.label4.Text = "Documento de abono";
            // 
            // cmbDoctoCargo
            // 
            this.cmbDoctoCargo.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbDoctoCargo.FormattingEnabled = true;
            this.cmbDoctoCargo.Location = new System.Drawing.Point(31, 163);
            this.cmbDoctoCargo.Name = "cmbDoctoCargo";
            this.cmbDoctoCargo.Size = new System.Drawing.Size(241, 25);
            this.cmbDoctoCargo.TabIndex = 9;
            // 
            // cmbDoctoAbono
            // 
            this.cmbDoctoAbono.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbDoctoAbono.FormattingEnabled = true;
            this.cmbDoctoAbono.Location = new System.Drawing.Point(31, 254);
            this.cmbDoctoAbono.Name = "cmbDoctoAbono";
            this.cmbDoctoAbono.Size = new System.Drawing.Size(241, 25);
            this.cmbDoctoAbono.TabIndex = 10;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::ReporteComercial.Properties.Resources.Matriz_Logo_Color;
            this.pictureBox1.Location = new System.Drawing.Point(81, 7);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(135, 115);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox1.TabIndex = 11;
            this.pictureBox1.TabStop = false;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Segoe UI", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.ForeColor = System.Drawing.Color.LightGray;
            this.label5.Location = new System.Drawing.Point(319, 42);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(253, 37);
            this.label5.TabIndex = 12;
            this.label5.Text = "Reporte comercial";
            // 
            // ReporteVentas
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(31)))), ((int)(((byte)(31)))), ((int)(((byte)(31)))));
            this.ClientSize = new System.Drawing.Size(617, 346);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.cmbDoctoAbono);
            this.Controls.Add(this.cmbDoctoCargo);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.bttSalir);
            this.Controls.Add(this.bttAcept);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.dtpFechaFinal);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dtpFechaInicial);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "ReporteVentas";
            this.Text = "Reporte comercial";
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DateTimePicker dtpFechaInicial;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DateTimePicker dtpFechaFinal;
        private System.Windows.Forms.Button bttAcept;
        private System.Windows.Forms.Button bttSalir;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox cmbDoctoCargo;
        private System.Windows.Forms.ComboBox cmbDoctoAbono;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label label5;
    }
}