namespace Genera_Fatture
{
    partial class form
    {
        /// <summary>
        /// Variabile di progettazione necessaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Pulire le risorse in uso.
        /// </summary>
        /// <param name="disposing">ha valore true se le risorse gestite devono essere eliminate, false in caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Codice generato da Progettazione Windows Form

        /// <summary>
        /// Metodo necessario per il supporto della finestra di progettazione. Non modificare
        /// il contenuto del metodo con l'editor di codice.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(form));
            this.panel1 = new System.Windows.Forms.Panel();
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.textBoxFileAnagrafica = new System.Windows.Forms.TextBox();
            this.numericUpDownNumeroFattura = new System.Windows.Forms.NumericUpDown();
            this.label2 = new System.Windows.Forms.Label();
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.label1 = new System.Windows.Forms.Label();
            this.textBoxLog = new System.Windows.Forms.RichTextBox();
            this.textBoxFileCosti = new System.Windows.Forms.TextBox();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.buttonFileAnagrafica = new Prototipo_Denso.PersonalUI.CustomButton();
            this.buttonGeneraFatture = new Prototipo_Denso.PersonalUI.CustomButton();
            this.buttonFileCosti = new Prototipo_Denso.PersonalUI.CustomButton();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownNumeroFattura)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.White;
            this.panel1.Controls.Add(this.pictureBox2);
            this.panel1.Controls.Add(this.pictureBox1);
            this.panel1.Controls.Add(this.textBoxFileAnagrafica);
            this.panel1.Controls.Add(this.buttonFileAnagrafica);
            this.panel1.Controls.Add(this.numericUpDownNumeroFattura);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.dateTimePicker1);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.textBoxLog);
            this.panel1.Controls.Add(this.buttonGeneraFatture);
            this.panel1.Controls.Add(this.buttonFileCosti);
            this.panel1.Controls.Add(this.textBoxFileCosti);
            this.panel1.Location = new System.Drawing.Point(-2, 1);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1297, 844);
            this.panel1.TabIndex = 2;
            this.panel1.Paint += new System.Windows.Forms.PaintEventHandler(this.panel1_Paint);
            // 
            // pictureBox2
            // 
            this.pictureBox2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.pictureBox2.Enabled = false;
            this.pictureBox2.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox2.Image")));
            this.pictureBox2.Location = new System.Drawing.Point(184, 29);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(332, 144);
            this.pictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox2.TabIndex = 13;
            this.pictureBox2.TabStop = false;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Enabled = false;
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.InitialImage = null;
            this.pictureBox1.Location = new System.Drawing.Point(522, 29);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(598, 144);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox1.TabIndex = 12;
            this.pictureBox1.TabStop = false;
            // 
            // textBoxFileAnagrafica
            // 
            this.textBoxFileAnagrafica.BackColor = System.Drawing.Color.White;
            this.textBoxFileAnagrafica.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBoxFileAnagrafica.Enabled = false;
            this.textBoxFileAnagrafica.Font = new System.Drawing.Font("Times New Roman", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBoxFileAnagrafica.Location = new System.Drawing.Point(405, 316);
            this.textBoxFileAnagrafica.Name = "textBoxFileAnagrafica";
            this.textBoxFileAnagrafica.ReadOnly = true;
            this.textBoxFileAnagrafica.Size = new System.Drawing.Size(784, 30);
            this.textBoxFileAnagrafica.TabIndex = 11;
            // 
            // numericUpDownNumeroFattura
            // 
            this.numericUpDownNumeroFattura.Font = new System.Drawing.Font("Times New Roman", 12F, System.Drawing.FontStyle.Bold);
            this.numericUpDownNumeroFattura.Location = new System.Drawing.Point(1008, 409);
            this.numericUpDownNumeroFattura.Name = "numericUpDownNumeroFattura";
            this.numericUpDownNumeroFattura.Size = new System.Drawing.Size(181, 30);
            this.numericUpDownNumeroFattura.TabIndex = 9;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Times New Roman", 12F, System.Drawing.FontStyle.Bold);
            this.label2.Location = new System.Drawing.Point(834, 411);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(168, 23);
            this.label2.TabIndex = 8;
            this.label2.Text = "N° Ultima Fattura:";
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.Font = new System.Drawing.Font("Times New Roman", 12F, System.Drawing.FontStyle.Bold);
            this.dateTimePicker1.Location = new System.Drawing.Point(247, 409);
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.Size = new System.Drawing.Size(342, 30);
            this.dateTimePicker1.TabIndex = 7;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Times New Roman", 12F, System.Drawing.FontStyle.Bold);
            this.label1.Location = new System.Drawing.Point(115, 411);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(126, 23);
            this.label1.TabIndex = 6;
            this.label1.Text = "Data Fatture:";
            // 
            // textBoxLog
            // 
            this.textBoxLog.BackColor = System.Drawing.Color.WhiteSmoke;
            this.textBoxLog.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBoxLog.Font = new System.Drawing.Font("Times New Roman", 10.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBoxLog.ForeColor = System.Drawing.Color.Black;
            this.textBoxLog.Location = new System.Drawing.Point(14, 547);
            this.textBoxLog.Name = "textBoxLog";
            this.textBoxLog.ReadOnly = true;
            this.textBoxLog.Size = new System.Drawing.Size(1267, 287);
            this.textBoxLog.TabIndex = 5;
            this.textBoxLog.Text = "";
            this.textBoxLog.WordWrap = false;
            // 
            // textBoxFileCosti
            // 
            this.textBoxFileCosti.BackColor = System.Drawing.Color.White;
            this.textBoxFileCosti.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBoxFileCosti.Enabled = false;
            this.textBoxFileCosti.Font = new System.Drawing.Font("Times New Roman", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBoxFileCosti.Location = new System.Drawing.Point(405, 232);
            this.textBoxFileCosti.Name = "textBoxFileCosti";
            this.textBoxFileCosti.ReadOnly = true;
            this.textBoxFileCosti.Size = new System.Drawing.Size(784, 30);
            this.textBoxFileCosti.TabIndex = 3;
            // 
            // openFileDialog
            // 
            this.openFileDialog.FileName = "openFileDialog";
            // 
            // buttonFileAnagrafica
            // 
            this.buttonFileAnagrafica.BackColor = System.Drawing.Color.White;
            this.buttonFileAnagrafica.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.buttonFileAnagrafica.BorderColor = System.Drawing.Color.Transparent;
            this.buttonFileAnagrafica.ButtonColor = System.Drawing.Color.DodgerBlue;
            this.buttonFileAnagrafica.Cursor = System.Windows.Forms.Cursors.Hand;
            this.buttonFileAnagrafica.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.buttonFileAnagrafica.FlatAppearance.BorderSize = 0;
            this.buttonFileAnagrafica.FlatAppearance.MouseDownBackColor = System.Drawing.Color.White;
            this.buttonFileAnagrafica.FlatAppearance.MouseOverBackColor = System.Drawing.Color.White;
            this.buttonFileAnagrafica.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonFileAnagrafica.Font = new System.Drawing.Font("Times New Roman", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonFileAnagrafica.ForeColor = System.Drawing.Color.Transparent;
            this.buttonFileAnagrafica.Location = new System.Drawing.Point(110, 303);
            this.buttonFileAnagrafica.Name = "buttonFileAnagrafica";
            this.buttonFileAnagrafica.OnHoverBorderColor = System.Drawing.Color.Transparent;
            this.buttonFileAnagrafica.OnHoverButtonColor = System.Drawing.Color.LightSkyBlue;
            this.buttonFileAnagrafica.OnHoverTextColor = System.Drawing.Color.White;
            this.buttonFileAnagrafica.Size = new System.Drawing.Size(250, 50);
            this.buttonFileAnagrafica.TabIndex = 10;
            this.buttonFileAnagrafica.Text = "File Excel Anagrafica";
            this.buttonFileAnagrafica.TextColor = System.Drawing.Color.White;
            this.buttonFileAnagrafica.UseVisualStyleBackColor = false;
            this.buttonFileAnagrafica.Click += new System.EventHandler(this.buttonFileAnagrafica_Click);
            // 
            // buttonGeneraFatture
            // 
            this.buttonGeneraFatture.BackColor = System.Drawing.Color.White;
            this.buttonGeneraFatture.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.buttonGeneraFatture.BorderColor = System.Drawing.Color.Transparent;
            this.buttonGeneraFatture.ButtonColor = System.Drawing.Color.DodgerBlue;
            this.buttonGeneraFatture.Cursor = System.Windows.Forms.Cursors.Hand;
            this.buttonGeneraFatture.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.buttonGeneraFatture.FlatAppearance.BorderSize = 0;
            this.buttonGeneraFatture.FlatAppearance.MouseDownBackColor = System.Drawing.Color.White;
            this.buttonGeneraFatture.FlatAppearance.MouseOverBackColor = System.Drawing.Color.White;
            this.buttonGeneraFatture.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonGeneraFatture.Font = new System.Drawing.Font("Times New Roman", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonGeneraFatture.ForeColor = System.Drawing.Color.Transparent;
            this.buttonGeneraFatture.Location = new System.Drawing.Point(951, 477);
            this.buttonGeneraFatture.Name = "buttonGeneraFatture";
            this.buttonGeneraFatture.OnHoverBorderColor = System.Drawing.Color.Transparent;
            this.buttonGeneraFatture.OnHoverButtonColor = System.Drawing.Color.LightSkyBlue;
            this.buttonGeneraFatture.OnHoverTextColor = System.Drawing.Color.White;
            this.buttonGeneraFatture.Size = new System.Drawing.Size(250, 50);
            this.buttonGeneraFatture.TabIndex = 4;
            this.buttonGeneraFatture.Text = "Genera Fatture";
            this.buttonGeneraFatture.TextColor = System.Drawing.Color.White;
            this.buttonGeneraFatture.UseVisualStyleBackColor = false;
            this.buttonGeneraFatture.Click += new System.EventHandler(this.buttonGeneraFatture_Click);
            // 
            // buttonFileCosti
            // 
            this.buttonFileCosti.BackColor = System.Drawing.Color.White;
            this.buttonFileCosti.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.buttonFileCosti.BorderColor = System.Drawing.Color.Transparent;
            this.buttonFileCosti.ButtonColor = System.Drawing.Color.DodgerBlue;
            this.buttonFileCosti.Cursor = System.Windows.Forms.Cursors.Hand;
            this.buttonFileCosti.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.buttonFileCosti.FlatAppearance.BorderSize = 0;
            this.buttonFileCosti.FlatAppearance.MouseDownBackColor = System.Drawing.Color.White;
            this.buttonFileCosti.FlatAppearance.MouseOverBackColor = System.Drawing.Color.White;
            this.buttonFileCosti.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonFileCosti.Font = new System.Drawing.Font("Times New Roman", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonFileCosti.ForeColor = System.Drawing.Color.Transparent;
            this.buttonFileCosti.Location = new System.Drawing.Point(110, 219);
            this.buttonFileCosti.Name = "buttonFileCosti";
            this.buttonFileCosti.OnHoverBorderColor = System.Drawing.Color.Transparent;
            this.buttonFileCosti.OnHoverButtonColor = System.Drawing.Color.LightSkyBlue;
            this.buttonFileCosti.OnHoverTextColor = System.Drawing.Color.White;
            this.buttonFileCosti.Size = new System.Drawing.Size(250, 50);
            this.buttonFileCosti.TabIndex = 1;
            this.buttonFileCosti.Text = "File Excel Costi";
            this.buttonFileCosti.TextColor = System.Drawing.Color.White;
            this.buttonFileCosti.UseVisualStyleBackColor = false;
            this.buttonFileCosti.Click += new System.EventHandler(this.buttonFileCosti_Click);
            // 
            // form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1291, 846);
            this.Controls.Add(this.panel1);
            this.MaximizeBox = false;
            this.Name = "form";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Generatore Fatture";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownNumeroFattura)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private Prototipo_Denso.PersonalUI.CustomButton buttonFileCosti;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.TextBox textBoxFileCosti;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private Prototipo_Denso.PersonalUI.CustomButton buttonGeneraFatture;
        private System.Windows.Forms.RichTextBox textBoxLog;
        private System.Windows.Forms.DateTimePicker dateTimePicker1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        protected System.Windows.Forms.NumericUpDown numericUpDownNumeroFattura;
        private Prototipo_Denso.PersonalUI.CustomButton buttonFileAnagrafica;
        private System.Windows.Forms.TextBox textBoxFileAnagrafica;
        private System.Windows.Forms.PictureBox pictureBox2;
        private System.Windows.Forms.PictureBox pictureBox1;
    }
}

