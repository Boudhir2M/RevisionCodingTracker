
namespace Révision
{
    partial class frmIntervalDécompte
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
            this.Label6 = new System.Windows.Forms.Label();
            this.BtnModifier = new System.Windows.Forms.Button();
            this.CmbDécompteA = new System.Windows.Forms.ComboBox();
            this.DTPDPS = new System.Windows.Forms.DateTimePicker();
            this.CmbNumMarché = new System.Windows.Forms.ComboBox();
            this.BtnAjouter = new System.Windows.Forms.Button();
            this.Label9 = new System.Windows.Forms.Label();
            this.Label11 = new System.Windows.Forms.Label();
            this.CmbDécompte = new System.Windows.Forms.ComboBox();
            this.CmbNumMarchéM = new System.Windows.Forms.ComboBox();
            this.Label3 = new System.Windows.Forms.Label();
            this.Label1 = new System.Windows.Forms.Label();
            this.TabFin = new System.Windows.Forms.TabPage();
            this.DTPDASM = new System.Windows.Forms.DateTimePicker();
            this.Label12 = new System.Windows.Forms.Label();
            this.TabCommencement = new System.Windows.Forms.TabPage();
            this.PNL = new System.Windows.Forms.TabControl();
            this.TabFin.SuspendLayout();
            this.TabCommencement.SuspendLayout();
            this.PNL.SuspendLayout();
            this.SuspendLayout();
            // 
            // Label6
            // 
            this.Label6.AutoSize = true;
            this.Label6.Location = new System.Drawing.Point(21, 82);
            this.Label6.Name = "Label6";
            this.Label6.Size = new System.Drawing.Size(123, 13);
            this.Label6.TabIndex = 43;
            this.Label6.Text = "Date d\'Arret de Service :";
            // 
            // BtnModifier
            // 
            this.BtnModifier.Location = new System.Drawing.Point(195, 116);
            this.BtnModifier.Name = "BtnModifier";
            this.BtnModifier.Size = new System.Drawing.Size(111, 35);
            this.BtnModifier.TabIndex = 34;
            this.BtnModifier.Text = "Fin";
            this.BtnModifier.UseVisualStyleBackColor = true;
            this.BtnModifier.Click += new System.EventHandler(this.BtnModifier_Click);
            // 
            // CmbDécompteA
            // 
            this.CmbDécompteA.FormattingEnabled = true;
            this.CmbDécompteA.Location = new System.Drawing.Point(129, 46);
            this.CmbDécompteA.Name = "CmbDécompteA";
            this.CmbDécompteA.Size = new System.Drawing.Size(177, 21);
            this.CmbDécompteA.TabIndex = 48;
            this.CmbDécompteA.SelectedIndexChanged += new System.EventHandler(this.CmbDécompteA_SelectedIndexChanged);
            // 
            // DTPDPS
            // 
            this.DTPDPS.CustomFormat = "01/01/2021";
            this.DTPDPS.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.DTPDPS.Location = new System.Drawing.Point(171, 76);
            this.DTPDPS.Name = "DTPDPS";
            this.DTPDPS.Size = new System.Drawing.Size(135, 20);
            this.DTPDPS.TabIndex = 37;
            this.DTPDPS.Value = new System.DateTime(2022, 11, 21, 0, 0, 0, 0);
            // 
            // CmbNumMarché
            // 
            this.CmbNumMarché.FormattingEnabled = true;
            this.CmbNumMarché.Location = new System.Drawing.Point(129, 16);
            this.CmbNumMarché.Name = "CmbNumMarché";
            this.CmbNumMarché.Size = new System.Drawing.Size(177, 21);
            this.CmbNumMarché.TabIndex = 36;
            this.CmbNumMarché.SelectedIndexChanged += new System.EventHandler(this.CmbNumMarché_SelectedIndexChanged);
            // 
            // BtnAjouter
            // 
            this.BtnAjouter.Location = new System.Drawing.Point(195, 116);
            this.BtnAjouter.Name = "BtnAjouter";
            this.BtnAjouter.Size = new System.Drawing.Size(111, 35);
            this.BtnAjouter.TabIndex = 12;
            this.BtnAjouter.Text = "Début";
            this.BtnAjouter.UseVisualStyleBackColor = true;
            this.BtnAjouter.Click += new System.EventHandler(this.BtnAjouter_Click);
            // 
            // Label9
            // 
            this.Label9.AutoSize = true;
            this.Label9.Location = new System.Drawing.Point(21, 49);
            this.Label9.Name = "Label9";
            this.Label9.Size = new System.Drawing.Size(102, 13);
            this.Label9.TabIndex = 3;
            this.Label9.Text = "Numéro Décompte :";
            // 
            // Label11
            // 
            this.Label11.AutoSize = true;
            this.Label11.Location = new System.Drawing.Point(21, 82);
            this.Label11.Name = "Label11";
            this.Label11.Size = new System.Drawing.Size(144, 13);
            this.Label11.TabIndex = 3;
            this.Label11.Text = "Date de Reprise de Service :";
            // 
            // CmbDécompte
            // 
            this.CmbDécompte.FormattingEnabled = true;
            this.CmbDécompte.Location = new System.Drawing.Point(129, 46);
            this.CmbDécompte.Name = "CmbDécompte";
            this.CmbDécompte.Size = new System.Drawing.Size(177, 21);
            this.CmbDécompte.TabIndex = 47;
            this.CmbDécompte.SelectedIndexChanged += new System.EventHandler(this.CmbDécompte_SelectedIndexChanged);
            // 
            // CmbNumMarchéM
            // 
            this.CmbNumMarchéM.FormattingEnabled = true;
            this.CmbNumMarchéM.Location = new System.Drawing.Point(129, 16);
            this.CmbNumMarchéM.Name = "CmbNumMarchéM";
            this.CmbNumMarchéM.Size = new System.Drawing.Size(177, 21);
            this.CmbNumMarchéM.TabIndex = 47;
            this.CmbNumMarchéM.SelectedIndexChanged += new System.EventHandler(this.CmbNumMarchéM_SelectedIndexChanged);
            // 
            // Label3
            // 
            this.Label3.AutoSize = true;
            this.Label3.Location = new System.Drawing.Point(21, 49);
            this.Label3.Name = "Label3";
            this.Label3.Size = new System.Drawing.Size(102, 13);
            this.Label3.TabIndex = 41;
            this.Label3.Text = "Numéro Décompte :";
            // 
            // Label1
            // 
            this.Label1.AutoSize = true;
            this.Label1.Location = new System.Drawing.Point(21, 19);
            this.Label1.Name = "Label1";
            this.Label1.Size = new System.Drawing.Size(102, 13);
            this.Label1.TabIndex = 0;
            this.Label1.Text = "Réference Marché :";
            // 
            // TabFin
            // 
            this.TabFin.BackColor = System.Drawing.Color.Silver;
            this.TabFin.Controls.Add(this.DTPDASM);
            this.TabFin.Controls.Add(this.CmbDécompte);
            this.TabFin.Controls.Add(this.CmbNumMarchéM);
            this.TabFin.Controls.Add(this.Label3);
            this.TabFin.Controls.Add(this.Label6);
            this.TabFin.Controls.Add(this.Label12);
            this.TabFin.Controls.Add(this.BtnModifier);
            this.TabFin.Location = new System.Drawing.Point(4, 22);
            this.TabFin.Name = "TabFin";
            this.TabFin.Padding = new System.Windows.Forms.Padding(3);
            this.TabFin.Size = new System.Drawing.Size(327, 165);
            this.TabFin.TabIndex = 1;
            this.TabFin.Text = "Fin";
            // 
            // DTPDASM
            // 
            this.DTPDASM.CustomFormat = "21/11/2022";
            this.DTPDASM.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.DTPDASM.Location = new System.Drawing.Point(171, 76);
            this.DTPDASM.Name = "DTPDASM";
            this.DTPDASM.Size = new System.Drawing.Size(135, 20);
            this.DTPDASM.TabIndex = 48;
            this.DTPDASM.Value = new System.DateTime(2022, 11, 21, 0, 0, 0, 0);
            // 
            // Label12
            // 
            this.Label12.AutoSize = true;
            this.Label12.Location = new System.Drawing.Point(21, 19);
            this.Label12.Name = "Label12";
            this.Label12.Size = new System.Drawing.Size(102, 13);
            this.Label12.TabIndex = 38;
            this.Label12.Text = "Réference Marché :";
            // 
            // TabCommencement
            // 
            this.TabCommencement.BackColor = System.Drawing.Color.Silver;
            this.TabCommencement.Controls.Add(this.CmbDécompteA);
            this.TabCommencement.Controls.Add(this.DTPDPS);
            this.TabCommencement.Controls.Add(this.CmbNumMarché);
            this.TabCommencement.Controls.Add(this.BtnAjouter);
            this.TabCommencement.Controls.Add(this.Label9);
            this.TabCommencement.Controls.Add(this.Label11);
            this.TabCommencement.Controls.Add(this.Label1);
            this.TabCommencement.Location = new System.Drawing.Point(4, 22);
            this.TabCommencement.Name = "TabCommencement";
            this.TabCommencement.Padding = new System.Windows.Forms.Padding(3);
            this.TabCommencement.Size = new System.Drawing.Size(327, 165);
            this.TabCommencement.TabIndex = 0;
            this.TabCommencement.Text = "Commencement";
            // 
            // PNL
            // 
            this.PNL.Controls.Add(this.TabCommencement);
            this.PNL.Controls.Add(this.TabFin);
            this.PNL.Location = new System.Drawing.Point(11, 11);
            this.PNL.Name = "PNL";
            this.PNL.SelectedIndex = 0;
            this.PNL.Size = new System.Drawing.Size(335, 191);
            this.PNL.TabIndex = 5;
            // 
            // frmIntervalDécompte
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(357, 213);
            this.Controls.Add(this.PNL);
            this.Name = "frmIntervalDécompte";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Reprise/Arret Période de Décompte";
            this.Load += new System.EventHandler(this.frmIntervalDécompte_Load);
            this.Resize += new System.EventHandler(this.frmIntervalDécompte_Resize);
            this.TabFin.ResumeLayout(false);
            this.TabFin.PerformLayout();
            this.TabCommencement.ResumeLayout(false);
            this.TabCommencement.PerformLayout();
            this.PNL.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        internal System.Windows.Forms.Label Label6;
        internal System.Windows.Forms.Button BtnModifier;
        internal System.Windows.Forms.ComboBox CmbDécompteA;
        internal System.Windows.Forms.DateTimePicker DTPDPS;
        internal System.Windows.Forms.ComboBox CmbNumMarché;
        internal System.Windows.Forms.Button BtnAjouter;
        internal System.Windows.Forms.Label Label9;
        internal System.Windows.Forms.Label Label11;
        internal System.Windows.Forms.ComboBox CmbDécompte;
        internal System.Windows.Forms.ComboBox CmbNumMarchéM;
        internal System.Windows.Forms.Label Label3;
        internal System.Windows.Forms.Label Label1;
        internal System.Windows.Forms.TabPage TabFin;
        internal System.Windows.Forms.DateTimePicker DTPDASM;
        internal System.Windows.Forms.Label Label12;
        internal System.Windows.Forms.TabPage TabCommencement;
        internal System.Windows.Forms.TabControl PNL;
    }
}