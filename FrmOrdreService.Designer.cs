
namespace Révision
{
    partial class FrmOrdreService
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
            this.BtnModifier = new System.Windows.Forms.Button();
            this.TabReprise = new System.Windows.Forms.TabPage();
            this.CmbDécompteA = new System.Windows.Forms.ComboBox();
            this.DTPDPS = new System.Windows.Forms.DateTimePicker();
            this.CmbNumMarché = new System.Windows.Forms.ComboBox();
            this.BtnAjouter = new System.Windows.Forms.Button();
            this.Label9 = new System.Windows.Forms.Label();
            this.Label11 = new System.Windows.Forms.Label();
            this.Label1 = new System.Windows.Forms.Label();
            this.CmbDOS = new System.Windows.Forms.ComboBox();
            this.DTPDASM = new System.Windows.Forms.DateTimePicker();
            this.CmbDécompte = new System.Windows.Forms.ComboBox();
            this.CmbNumMarchéM = new System.Windows.Forms.ComboBox();
            this.Label3 = new System.Windows.Forms.Label();
            this.Label2 = new System.Windows.Forms.Label();
            this.Label6 = new System.Windows.Forms.Label();
            this.TabArret = new System.Windows.Forms.TabPage();
            this.Label12 = new System.Windows.Forms.Label();
            this.PNL = new System.Windows.Forms.TabControl();
            this.TabReprise.SuspendLayout();
            this.TabArret.SuspendLayout();
            this.PNL.SuspendLayout();
            this.SuspendLayout();
            // 
            // BtnModifier
            // 
            this.BtnModifier.Location = new System.Drawing.Point(206, 155);
            this.BtnModifier.Name = "BtnModifier";
            this.BtnModifier.Size = new System.Drawing.Size(103, 36);
            this.BtnModifier.TabIndex = 34;
            this.BtnModifier.Text = "Arreter";
            this.BtnModifier.UseVisualStyleBackColor = true;
            this.BtnModifier.Click += new System.EventHandler(this.BtnModifier_Click);
            // 
            // TabReprise
            // 
            this.TabReprise.BackColor = System.Drawing.Color.Silver;
            this.TabReprise.Controls.Add(this.CmbDécompteA);
            this.TabReprise.Controls.Add(this.DTPDPS);
            this.TabReprise.Controls.Add(this.CmbNumMarché);
            this.TabReprise.Controls.Add(this.BtnAjouter);
            this.TabReprise.Controls.Add(this.Label9);
            this.TabReprise.Controls.Add(this.Label11);
            this.TabReprise.Controls.Add(this.Label1);
            this.TabReprise.Location = new System.Drawing.Point(4, 22);
            this.TabReprise.Name = "TabReprise";
            this.TabReprise.Padding = new System.Windows.Forms.Padding(3);
            this.TabReprise.Size = new System.Drawing.Size(327, 203);
            this.TabReprise.TabIndex = 0;
            this.TabReprise.Text = "Reprise";
            // 
            // CmbDécompteA
            // 
            this.CmbDécompteA.FormattingEnabled = true;
            this.CmbDécompteA.Location = new System.Drawing.Point(132, 55);
            this.CmbDécompteA.Name = "CmbDécompteA";
            this.CmbDécompteA.Size = new System.Drawing.Size(177, 21);
            this.CmbDécompteA.TabIndex = 48;
            this.CmbDécompteA.SelectedIndexChanged += new System.EventHandler(this.CmbDécompteA_SelectedIndexChanged);
            // 
            // DTPDPS
            // 
            this.DTPDPS.CustomFormat = "01/01/2021";
            this.DTPDPS.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.DTPDPS.Location = new System.Drawing.Point(174, 91);
            this.DTPDPS.Name = "DTPDPS";
            this.DTPDPS.Size = new System.Drawing.Size(135, 20);
            this.DTPDPS.TabIndex = 37;
            this.DTPDPS.Value = new System.DateTime(2022, 11, 21, 0, 0, 0, 0);
            // 
            // CmbNumMarché
            // 
            this.CmbNumMarché.FormattingEnabled = true;
            this.CmbNumMarché.Location = new System.Drawing.Point(132, 19);
            this.CmbNumMarché.Name = "CmbNumMarché";
            this.CmbNumMarché.Size = new System.Drawing.Size(177, 21);
            this.CmbNumMarché.TabIndex = 36;
            this.CmbNumMarché.SelectedIndexChanged += new System.EventHandler(this.CmbNumMarché_SelectedIndexChanged);
            // 
            // BtnAjouter
            // 
            this.BtnAjouter.Location = new System.Drawing.Point(206, 155);
            this.BtnAjouter.Name = "BtnAjouter";
            this.BtnAjouter.Size = new System.Drawing.Size(103, 36);
            this.BtnAjouter.TabIndex = 12;
            this.BtnAjouter.Text = "Reprendre";
            this.BtnAjouter.UseVisualStyleBackColor = true;
            this.BtnAjouter.Click += new System.EventHandler(this.BtnAjouter_Click);
            // 
            // Label9
            // 
            this.Label9.AutoSize = true;
            this.Label9.Location = new System.Drawing.Point(24, 55);
            this.Label9.Name = "Label9";
            this.Label9.Size = new System.Drawing.Size(102, 13);
            this.Label9.TabIndex = 3;
            this.Label9.Text = "Numéro Décompte :";
            // 
            // Label11
            // 
            this.Label11.AutoSize = true;
            this.Label11.Location = new System.Drawing.Point(24, 91);
            this.Label11.Name = "Label11";
            this.Label11.Size = new System.Drawing.Size(144, 13);
            this.Label11.TabIndex = 3;
            this.Label11.Text = "Date de Reprise de Service :";
            // 
            // Label1
            // 
            this.Label1.AutoSize = true;
            this.Label1.Location = new System.Drawing.Point(24, 19);
            this.Label1.Name = "Label1";
            this.Label1.Size = new System.Drawing.Size(102, 13);
            this.Label1.TabIndex = 0;
            this.Label1.Text = "Réference Marché :";
            // 
            // CmbDOS
            // 
            this.CmbDOS.FormattingEnabled = true;
            this.CmbDOS.Location = new System.Drawing.Point(174, 91);
            this.CmbDOS.Name = "CmbDOS";
            this.CmbDOS.Size = new System.Drawing.Size(135, 21);
            this.CmbDOS.TabIndex = 49;
            // 
            // DTPDASM
            // 
            this.DTPDASM.CustomFormat = "01/01/2021";
            this.DTPDASM.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.DTPDASM.Location = new System.Drawing.Point(174, 125);
            this.DTPDASM.Name = "DTPDASM";
            this.DTPDASM.Size = new System.Drawing.Size(135, 20);
            this.DTPDASM.TabIndex = 48;
            this.DTPDASM.Value = new System.DateTime(2022, 11, 21, 0, 0, 0, 0);
            // 
            // CmbDécompte
            // 
            this.CmbDécompte.FormattingEnabled = true;
            this.CmbDécompte.Location = new System.Drawing.Point(132, 55);
            this.CmbDécompte.Name = "CmbDécompte";
            this.CmbDécompte.Size = new System.Drawing.Size(177, 21);
            this.CmbDécompte.TabIndex = 47;
            this.CmbDécompte.SelectedIndexChanged += new System.EventHandler(this.CmbDécompte_SelectedIndexChanged);
            // 
            // CmbNumMarchéM
            // 
            this.CmbNumMarchéM.FormattingEnabled = true;
            this.CmbNumMarchéM.Location = new System.Drawing.Point(132, 19);
            this.CmbNumMarchéM.Name = "CmbNumMarchéM";
            this.CmbNumMarchéM.Size = new System.Drawing.Size(177, 21);
            this.CmbNumMarchéM.TabIndex = 47;
            this.CmbNumMarchéM.SelectedIndexChanged += new System.EventHandler(this.CmbNumMarchéM_SelectedIndexChanged);
            // 
            // Label3
            // 
            this.Label3.AutoSize = true;
            this.Label3.Location = new System.Drawing.Point(24, 55);
            this.Label3.Name = "Label3";
            this.Label3.Size = new System.Drawing.Size(102, 13);
            this.Label3.TabIndex = 41;
            this.Label3.Text = "Numéro Décompte :";
            // 
            // Label2
            // 
            this.Label2.AutoSize = true;
            this.Label2.Location = new System.Drawing.Point(24, 91);
            this.Label2.Name = "Label2";
            this.Label2.Size = new System.Drawing.Size(127, 13);
            this.Label2.TabIndex = 43;
            this.Label2.Text = "Date d\'Ordre de Service :";
            // 
            // Label6
            // 
            this.Label6.AutoSize = true;
            this.Label6.Location = new System.Drawing.Point(17, 131);
            this.Label6.Name = "Label6";
            this.Label6.Size = new System.Drawing.Size(123, 13);
            this.Label6.TabIndex = 43;
            this.Label6.Text = "Date d\'Arret de Service :";
            // 
            // TabArret
            // 
            this.TabArret.BackColor = System.Drawing.Color.Silver;
            this.TabArret.Controls.Add(this.CmbDOS);
            this.TabArret.Controls.Add(this.DTPDASM);
            this.TabArret.Controls.Add(this.CmbDécompte);
            this.TabArret.Controls.Add(this.CmbNumMarchéM);
            this.TabArret.Controls.Add(this.Label3);
            this.TabArret.Controls.Add(this.Label2);
            this.TabArret.Controls.Add(this.Label6);
            this.TabArret.Controls.Add(this.Label12);
            this.TabArret.Controls.Add(this.BtnModifier);
            this.TabArret.Location = new System.Drawing.Point(4, 22);
            this.TabArret.Name = "TabArret";
            this.TabArret.Padding = new System.Windows.Forms.Padding(3);
            this.TabArret.Size = new System.Drawing.Size(327, 203);
            this.TabArret.TabIndex = 1;
            this.TabArret.Text = "Arret";
            // 
            // Label12
            // 
            this.Label12.AutoSize = true;
            this.Label12.Location = new System.Drawing.Point(24, 19);
            this.Label12.Name = "Label12";
            this.Label12.Size = new System.Drawing.Size(102, 13);
            this.Label12.TabIndex = 38;
            this.Label12.Text = "Réference Marché :";
            // 
            // PNL
            // 
            this.PNL.Controls.Add(this.TabReprise);
            this.PNL.Controls.Add(this.TabArret);
            this.PNL.Location = new System.Drawing.Point(11, 12);
            this.PNL.Name = "PNL";
            this.PNL.SelectedIndex = 0;
            this.PNL.Size = new System.Drawing.Size(335, 229);
            this.PNL.TabIndex = 4;
            // 
            // FrmOrdreService
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(356, 253);
            this.Controls.Add(this.PNL);
            this.Name = "FrmOrdreService";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Ordre de Service";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.FrmOrdreService_FormClosed);
            this.Load += new System.EventHandler(this.FrmOrdreService_Load);
            this.Resize += new System.EventHandler(this.FrmOrdreService_Resize);
            this.TabReprise.ResumeLayout(false);
            this.TabReprise.PerformLayout();
            this.TabArret.ResumeLayout(false);
            this.TabArret.PerformLayout();
            this.PNL.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        internal System.Windows.Forms.Button BtnModifier;
        internal System.Windows.Forms.TabPage TabReprise;
        internal System.Windows.Forms.ComboBox CmbDécompteA;
        internal System.Windows.Forms.DateTimePicker DTPDPS;
        internal System.Windows.Forms.ComboBox CmbNumMarché;
        internal System.Windows.Forms.Button BtnAjouter;
        internal System.Windows.Forms.Label Label9;
        internal System.Windows.Forms.Label Label11;
        internal System.Windows.Forms.Label Label1;
        internal System.Windows.Forms.ComboBox CmbDOS;
        internal System.Windows.Forms.DateTimePicker DTPDASM;
        internal System.Windows.Forms.ComboBox CmbDécompte;
        internal System.Windows.Forms.ComboBox CmbNumMarchéM;
        internal System.Windows.Forms.Label Label3;
        internal System.Windows.Forms.Label Label2;
        internal System.Windows.Forms.Label Label6;
        internal System.Windows.Forms.TabPage TabArret;
        internal System.Windows.Forms.Label Label12;
        internal System.Windows.Forms.TabControl PNL;
    }
}