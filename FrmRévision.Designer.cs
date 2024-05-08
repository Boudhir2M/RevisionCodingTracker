
namespace Révision
{
    partial class FrmRévision
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
            this.PNL = new System.Windows.Forms.Panel();
            this.GroupBox3 = new System.Windows.Forms.GroupBox();
            this.DGVP = new System.Windows.Forms.DataGridView();
            this.GroupBox2 = new System.Windows.Forms.GroupBox();
            this.DGV = new System.Windows.Forms.DataGridView();
            this.GroupBox1 = new System.Windows.Forms.GroupBox();
            this.TxtObjet = new System.Windows.Forms.TextBox();
            this.CmbDécompteM = new System.Windows.Forms.ComboBox();
            this.Label1 = new System.Windows.Forms.Label();
            this.Label6 = new System.Windows.Forms.Label();
            this.Label8 = new System.Windows.Forms.Label();
            this.CmbNumMarchéM = new System.Windows.Forms.ComboBox();
            this.BtnEditer = new System.Windows.Forms.Button();
            this.PNL.SuspendLayout();
            this.GroupBox3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DGVP)).BeginInit();
            this.GroupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DGV)).BeginInit();
            this.GroupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // PNL
            // 
            this.PNL.Controls.Add(this.GroupBox3);
            this.PNL.Controls.Add(this.GroupBox2);
            this.PNL.Controls.Add(this.GroupBox1);
            this.PNL.Location = new System.Drawing.Point(27, 12);
            this.PNL.Name = "PNL";
            this.PNL.Size = new System.Drawing.Size(1099, 478);
            this.PNL.TabIndex = 7;
            // 
            // GroupBox3
            // 
            this.GroupBox3.Controls.Add(this.DGVP);
            this.GroupBox3.Location = new System.Drawing.Point(420, 118);
            this.GroupBox3.Name = "GroupBox3";
            this.GroupBox3.Size = new System.Drawing.Size(658, 351);
            this.GroupBox3.TabIndex = 8;
            this.GroupBox3.TabStop = false;
            this.GroupBox3.Text = "Informations :";
            // 
            // DGVP
            // 
            this.DGVP.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.DGVP.BackgroundColor = System.Drawing.SystemColors.ControlLight;
            this.DGVP.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGVP.Location = new System.Drawing.Point(14, 12);
            this.DGVP.Name = "DGVP";
            this.DGVP.RowHeadersVisible = false;
            this.DGVP.Size = new System.Drawing.Size(631, 326);
            this.DGVP.TabIndex = 0;
            // 
            // GroupBox2
            // 
            this.GroupBox2.Controls.Add(this.DGV);
            this.GroupBox2.Location = new System.Drawing.Point(18, 118);
            this.GroupBox2.Name = "GroupBox2";
            this.GroupBox2.Size = new System.Drawing.Size(396, 351);
            this.GroupBox2.TabIndex = 7;
            this.GroupBox2.TabStop = false;
            this.GroupBox2.Text = "Informations :";
            // 
            // DGV
            // 
            this.DGV.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.DGV.BackgroundColor = System.Drawing.SystemColors.ControlLight;
            this.DGV.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGV.Location = new System.Drawing.Point(14, 12);
            this.DGV.MultiSelect = false;
            this.DGV.Name = "DGV";
            this.DGV.RowHeadersVisible = false;
            this.DGV.Size = new System.Drawing.Size(368, 326);
            this.DGV.TabIndex = 0;
            // 
            // GroupBox1
            // 
            this.GroupBox1.Controls.Add(this.TxtObjet);
            this.GroupBox1.Controls.Add(this.CmbDécompteM);
            this.GroupBox1.Controls.Add(this.Label1);
            this.GroupBox1.Controls.Add(this.Label6);
            this.GroupBox1.Controls.Add(this.Label8);
            this.GroupBox1.Controls.Add(this.CmbNumMarchéM);
            this.GroupBox1.Location = new System.Drawing.Point(18, 7);
            this.GroupBox1.Name = "GroupBox1";
            this.GroupBox1.Size = new System.Drawing.Size(1060, 105);
            this.GroupBox1.TabIndex = 6;
            this.GroupBox1.TabStop = false;
            // 
            // TxtObjet
            // 
            this.TxtObjet.ForeColor = System.Drawing.Color.DarkGray;
            this.TxtObjet.Location = new System.Drawing.Point(416, 27);
            this.TxtObjet.Multiline = true;
            this.TxtObjet.Name = "TxtObjet";
            this.TxtObjet.Size = new System.Drawing.Size(631, 62);
            this.TxtObjet.TabIndex = 47;
            // 
            // CmbDécompteM
            // 
            this.CmbDécompteM.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.CmbDécompteM.ForeColor = System.Drawing.Color.Black;
            this.CmbDécompteM.FormattingEnabled = true;
            this.CmbDécompteM.Location = new System.Drawing.Point(14, 68);
            this.CmbDécompteM.Name = "CmbDécompteM";
            this.CmbDécompteM.Size = new System.Drawing.Size(214, 21);
            this.CmbDécompteM.TabIndex = 46;
            this.CmbDécompteM.SelectedIndexChanged += new System.EventHandler(this.CmbDécompteM_SelectedIndexChanged);
            // 
            // Label1
            // 
            this.Label1.AutoSize = true;
            this.Label1.Location = new System.Drawing.Point(416, 12);
            this.Label1.Name = "Label1";
            this.Label1.Size = new System.Drawing.Size(38, 13);
            this.Label1.TabIndex = 45;
            this.Label1.Text = "Objet :";
            // 
            // Label6
            // 
            this.Label6.AutoSize = true;
            this.Label6.Location = new System.Drawing.Point(11, 54);
            this.Label6.Name = "Label6";
            this.Label6.Size = new System.Drawing.Size(102, 13);
            this.Label6.TabIndex = 45;
            this.Label6.Text = "Numéro Décompte :";
            // 
            // Label8
            // 
            this.Label8.AutoSize = true;
            this.Label8.Location = new System.Drawing.Point(12, 13);
            this.Label8.Name = "Label8";
            this.Label8.Size = new System.Drawing.Size(102, 13);
            this.Label8.TabIndex = 44;
            this.Label8.Text = "Réference Marché :";
            // 
            // CmbNumMarchéM
            // 
            this.CmbNumMarchéM.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.CmbNumMarchéM.ForeColor = System.Drawing.Color.Black;
            this.CmbNumMarchéM.FormattingEnabled = true;
            this.CmbNumMarchéM.Location = new System.Drawing.Point(14, 27);
            this.CmbNumMarchéM.Name = "CmbNumMarchéM";
            this.CmbNumMarchéM.Size = new System.Drawing.Size(214, 21);
            this.CmbNumMarchéM.TabIndex = 43;
            this.CmbNumMarchéM.SelectedIndexChanged += new System.EventHandler(this.CmbNumMarchéM_SelectedIndexChanged);
            // 
            // BtnEditer
            // 
            this.BtnEditer.BackColor = System.Drawing.Color.Red;
            this.BtnEditer.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BtnEditer.ForeColor = System.Drawing.Color.White;
            this.BtnEditer.Location = new System.Drawing.Point(959, 502);
            this.BtnEditer.Name = "BtnEditer";
            this.BtnEditer.Size = new System.Drawing.Size(133, 46);
            this.BtnEditer.TabIndex = 48;
            this.BtnEditer.Text = "Editer";
            this.BtnEditer.UseVisualStyleBackColor = false;
            this.BtnEditer.Click += new System.EventHandler(this.BtnEditer_Click);
            // 
            // FrmRévision
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1138, 560);
            this.Controls.Add(this.BtnEditer);
            this.Controls.Add(this.PNL);
            this.Name = "FrmRévision";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Révision d\'un Seul Décompte";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.FrmRévision_FormClosed);
            this.Load += new System.EventHandler(this.FrmRévision_Load);
            this.Resize += new System.EventHandler(this.FrmRévision_Resize);
            this.PNL.ResumeLayout(false);
            this.GroupBox3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.DGVP)).EndInit();
            this.GroupBox2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.DGV)).EndInit();
            this.GroupBox1.ResumeLayout(false);
            this.GroupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal System.Windows.Forms.Panel PNL;
        internal System.Windows.Forms.GroupBox GroupBox3;
        internal System.Windows.Forms.DataGridView DGVP;
        internal System.Windows.Forms.GroupBox GroupBox2;
        internal System.Windows.Forms.DataGridView DGV;
        internal System.Windows.Forms.GroupBox GroupBox1;
        internal System.Windows.Forms.Button BtnEditer;
        internal System.Windows.Forms.TextBox TxtObjet;
        internal System.Windows.Forms.ComboBox CmbDécompteM;
        internal System.Windows.Forms.Label Label1;
        internal System.Windows.Forms.Label Label6;
        internal System.Windows.Forms.Label Label8;
        internal System.Windows.Forms.ComboBox CmbNumMarchéM;
    }
}