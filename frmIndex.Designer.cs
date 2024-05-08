
namespace Révision
{
    partial class frmIndex
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
            this.OFD = new System.Windows.Forms.OpenFileDialog();
            this.DTPS = new System.Windows.Forms.DateTimePicker();
            this.Label8 = new System.Windows.Forms.Label();
            this.CmbSymboleS = new System.Windows.Forms.ComboBox();
            this.Label5 = new System.Windows.Forms.Label();
            this.Suppression = new System.Windows.Forms.TabPage();
            this.BtnSupprimer = new System.Windows.Forms.Button();
            this.Modification = new System.Windows.Forms.TabPage();
            this.DTPM = new System.Windows.Forms.DateTimePicker();
            this.CmbSymboleM = new System.Windows.Forms.ComboBox();
            this.Label7 = new System.Windows.Forms.Label();
            this.BtnModifier = new System.Windows.Forms.Button();
            this.TxtValeurM = new System.Windows.Forms.TextBox();
            this.Label4 = new System.Windows.Forms.Label();
            this.Label3 = new System.Windows.Forms.Label();
            this.Label1 = new System.Windows.Forms.Label();
            this.Label6 = new System.Windows.Forms.Label();
            this.BtnAjouter = new System.Windows.Forms.Button();
            this.TxtValeurA = new System.Windows.Forms.TextBox();
            this.PB = new System.Windows.Forms.ProgressBar();
            this.BtnExcel = new System.Windows.Forms.Button();
            this.DTPA = new System.Windows.Forms.DateTimePicker();
            this.CmbSymboleA = new System.Windows.Forms.ComboBox();
            this.Ajout = new System.Windows.Forms.TabPage();
            this.Label2 = new System.Windows.Forms.Label();
            this.PNL = new System.Windows.Forms.TabControl();
            this.Suppression.SuspendLayout();
            this.Modification.SuspendLayout();
            this.Ajout.SuspendLayout();
            this.PNL.SuspendLayout();
            this.SuspendLayout();
            // 
            // OFD
            // 
            this.OFD.FileName = "OpenFileDialog1";
            // 
            // DTPS
            // 
            this.DTPS.Location = new System.Drawing.Point(101, 52);
            this.DTPS.Name = "DTPS";
            this.DTPS.Size = new System.Drawing.Size(179, 20);
            this.DTPS.TabIndex = 13;
            this.DTPS.Value = new System.DateTime(2022, 11, 8, 0, 0, 0, 0);
            // 
            // Label8
            // 
            this.Label8.AutoSize = true;
            this.Label8.Location = new System.Drawing.Point(28, 59);
            this.Label8.Name = "Label8";
            this.Label8.Size = new System.Drawing.Size(35, 13);
            this.Label8.TabIndex = 12;
            this.Label8.Text = "Mois :";
            // 
            // CmbSymboleS
            // 
            this.CmbSymboleS.FormattingEnabled = true;
            this.CmbSymboleS.Location = new System.Drawing.Point(101, 18);
            this.CmbSymboleS.Name = "CmbSymboleS";
            this.CmbSymboleS.Size = new System.Drawing.Size(179, 21);
            this.CmbSymboleS.TabIndex = 5;
            this.CmbSymboleS.SelectedIndexChanged += new System.EventHandler(this.CmbSymboleS_SelectedIndexChanged);
            // 
            // Label5
            // 
            this.Label5.AutoSize = true;
            this.Label5.Location = new System.Drawing.Point(28, 26);
            this.Label5.Name = "Label5";
            this.Label5.Size = new System.Drawing.Size(53, 13);
            this.Label5.TabIndex = 4;
            this.Label5.Text = "Symbole :";
            // 
            // Suppression
            // 
            this.Suppression.BackColor = System.Drawing.Color.Silver;
            this.Suppression.Controls.Add(this.DTPS);
            this.Suppression.Controls.Add(this.Label8);
            this.Suppression.Controls.Add(this.CmbSymboleS);
            this.Suppression.Controls.Add(this.Label5);
            this.Suppression.Controls.Add(this.BtnSupprimer);
            this.Suppression.Location = new System.Drawing.Point(4, 22);
            this.Suppression.Name = "Suppression";
            this.Suppression.Size = new System.Drawing.Size(323, 206);
            this.Suppression.TabIndex = 2;
            this.Suppression.Text = "Suppression";
            // 
            // BtnSupprimer
            // 
            this.BtnSupprimer.Location = new System.Drawing.Point(156, 140);
            this.BtnSupprimer.Name = "BtnSupprimer";
            this.BtnSupprimer.Size = new System.Drawing.Size(113, 28);
            this.BtnSupprimer.TabIndex = 3;
            this.BtnSupprimer.Text = "Supprimer";
            this.BtnSupprimer.UseVisualStyleBackColor = true;
            this.BtnSupprimer.Click += new System.EventHandler(this.BtnSupprimer_Click);
            // 
            // Modification
            // 
            this.Modification.BackColor = System.Drawing.Color.Silver;
            this.Modification.Controls.Add(this.DTPM);
            this.Modification.Controls.Add(this.CmbSymboleM);
            this.Modification.Controls.Add(this.Label7);
            this.Modification.Controls.Add(this.BtnModifier);
            this.Modification.Controls.Add(this.TxtValeurM);
            this.Modification.Controls.Add(this.Label4);
            this.Modification.Controls.Add(this.Label3);
            this.Modification.Location = new System.Drawing.Point(4, 22);
            this.Modification.Name = "Modification";
            this.Modification.Padding = new System.Windows.Forms.Padding(3);
            this.Modification.Size = new System.Drawing.Size(323, 206);
            this.Modification.TabIndex = 1;
            this.Modification.Text = "Modification";
            // 
            // DTPM
            // 
            this.DTPM.Location = new System.Drawing.Point(101, 52);
            this.DTPM.Name = "DTPM";
            this.DTPM.Size = new System.Drawing.Size(179, 20);
            this.DTPM.TabIndex = 11;
            this.DTPM.Value = new System.DateTime(2014, 1, 1, 0, 0, 0, 0);
            // 
            // CmbSymboleM
            // 
            this.CmbSymboleM.FormattingEnabled = true;
            this.CmbSymboleM.Location = new System.Drawing.Point(101, 18);
            this.CmbSymboleM.Name = "CmbSymboleM";
            this.CmbSymboleM.Size = new System.Drawing.Size(179, 21);
            this.CmbSymboleM.TabIndex = 10;
            this.CmbSymboleM.Text = "TR1";
            this.CmbSymboleM.SelectedIndexChanged += new System.EventHandler(this.CmbSymboleM_SelectedIndexChanged);
            // 
            // Label7
            // 
            this.Label7.AutoSize = true;
            this.Label7.Location = new System.Drawing.Point(28, 96);
            this.Label7.Name = "Label7";
            this.Label7.Size = new System.Drawing.Size(43, 13);
            this.Label7.TabIndex = 9;
            this.Label7.Text = "Valeur :";
            // 
            // BtnModifier
            // 
            this.BtnModifier.Location = new System.Drawing.Point(156, 140);
            this.BtnModifier.Name = "BtnModifier";
            this.BtnModifier.Size = new System.Drawing.Size(113, 28);
            this.BtnModifier.TabIndex = 6;
            this.BtnModifier.Text = "Modifier";
            this.BtnModifier.UseVisualStyleBackColor = true;
            this.BtnModifier.Click += new System.EventHandler(this.BtnModifier_Click);
            // 
            // TxtValeurM
            // 
            this.TxtValeurM.Location = new System.Drawing.Point(101, 93);
            this.TxtValeurM.Name = "TxtValeurM";
            this.TxtValeurM.Size = new System.Drawing.Size(179, 20);
            this.TxtValeurM.TabIndex = 4;
            // 
            // Label4
            // 
            this.Label4.AutoSize = true;
            this.Label4.Location = new System.Drawing.Point(28, 59);
            this.Label4.Name = "Label4";
            this.Label4.Size = new System.Drawing.Size(35, 13);
            this.Label4.TabIndex = 2;
            this.Label4.Text = "Mois :";
            // 
            // Label3
            // 
            this.Label3.AutoSize = true;
            this.Label3.Location = new System.Drawing.Point(28, 26);
            this.Label3.Name = "Label3";
            this.Label3.Size = new System.Drawing.Size(53, 13);
            this.Label3.TabIndex = 3;
            this.Label3.Text = "Symbole :";
            // 
            // Label1
            // 
            this.Label1.AutoSize = true;
            this.Label1.Location = new System.Drawing.Point(28, 26);
            this.Label1.Name = "Label1";
            this.Label1.Size = new System.Drawing.Size(53, 13);
            this.Label1.TabIndex = 3;
            this.Label1.Text = "Symbole :";
            // 
            // Label6
            // 
            this.Label6.AutoSize = true;
            this.Label6.Location = new System.Drawing.Point(28, 96);
            this.Label6.Name = "Label6";
            this.Label6.Size = new System.Drawing.Size(43, 13);
            this.Label6.TabIndex = 7;
            this.Label6.Text = "Valeur :";
            // 
            // BtnAjouter
            // 
            this.BtnAjouter.Location = new System.Drawing.Point(156, 140);
            this.BtnAjouter.Name = "BtnAjouter";
            this.BtnAjouter.Size = new System.Drawing.Size(113, 28);
            this.BtnAjouter.TabIndex = 6;
            this.BtnAjouter.Text = "Ajouter";
            this.BtnAjouter.UseVisualStyleBackColor = true;
            this.BtnAjouter.Click += new System.EventHandler(this.BtnAjouter_Click);
            // 
            // TxtValeurA
            // 
            this.TxtValeurA.Location = new System.Drawing.Point(101, 93);
            this.TxtValeurA.Name = "TxtValeurA";
            this.TxtValeurA.Size = new System.Drawing.Size(168, 20);
            this.TxtValeurA.TabIndex = 4;
            // 
            // PB
            // 
            this.PB.Location = new System.Drawing.Point(37, 185);
            this.PB.Name = "PB";
            this.PB.Size = new System.Drawing.Size(248, 10);
            this.PB.TabIndex = 11;
            // 
            // BtnExcel
            // 
            this.BtnExcel.Location = new System.Drawing.Point(22, 140);
            this.BtnExcel.Name = "BtnExcel";
            this.BtnExcel.Size = new System.Drawing.Size(117, 27);
            this.BtnExcel.TabIndex = 10;
            this.BtnExcel.Text = "Excel >>>";
            this.BtnExcel.UseVisualStyleBackColor = true;
            this.BtnExcel.Click += new System.EventHandler(this.BtnExcel_Click);
            // 
            // DTPA
            // 
            this.DTPA.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.DTPA.Location = new System.Drawing.Point(101, 52);
            this.DTPA.Name = "DTPA";
            this.DTPA.Size = new System.Drawing.Size(168, 20);
            this.DTPA.TabIndex = 9;
            this.DTPA.Value = new System.DateTime(2014, 1, 1, 0, 0, 0, 0);
            // 
            // CmbSymboleA
            // 
            this.CmbSymboleA.FormattingEnabled = true;
            this.CmbSymboleA.Location = new System.Drawing.Point(101, 18);
            this.CmbSymboleA.Name = "CmbSymboleA";
            this.CmbSymboleA.Size = new System.Drawing.Size(168, 21);
            this.CmbSymboleA.TabIndex = 8;
            this.CmbSymboleA.SelectedIndexChanged += new System.EventHandler(this.CmbSymboleA_SelectedIndexChanged);
            // 
            // Ajout
            // 
            this.Ajout.BackColor = System.Drawing.Color.Silver;
            this.Ajout.Controls.Add(this.PB);
            this.Ajout.Controls.Add(this.BtnExcel);
            this.Ajout.Controls.Add(this.DTPA);
            this.Ajout.Controls.Add(this.CmbSymboleA);
            this.Ajout.Controls.Add(this.Label6);
            this.Ajout.Controls.Add(this.BtnAjouter);
            this.Ajout.Controls.Add(this.TxtValeurA);
            this.Ajout.Controls.Add(this.Label2);
            this.Ajout.Controls.Add(this.Label1);
            this.Ajout.Location = new System.Drawing.Point(4, 22);
            this.Ajout.Name = "Ajout";
            this.Ajout.Padding = new System.Windows.Forms.Padding(3);
            this.Ajout.Size = new System.Drawing.Size(323, 206);
            this.Ajout.TabIndex = 0;
            this.Ajout.Text = "Ajout";
            // 
            // Label2
            // 
            this.Label2.AutoSize = true;
            this.Label2.Location = new System.Drawing.Point(28, 59);
            this.Label2.Name = "Label2";
            this.Label2.Size = new System.Drawing.Size(35, 13);
            this.Label2.TabIndex = 2;
            this.Label2.Text = "Mois :";
            // 
            // PNL
            // 
            this.PNL.Controls.Add(this.Ajout);
            this.PNL.Controls.Add(this.Modification);
            this.PNL.Controls.Add(this.Suppression);
            this.PNL.Location = new System.Drawing.Point(12, 10);
            this.PNL.Name = "PNL";
            this.PNL.SelectedIndex = 0;
            this.PNL.Size = new System.Drawing.Size(331, 232);
            this.PNL.TabIndex = 5;
            // 
            // frmIndex
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(355, 252);
            this.Controls.Add(this.PNL);
            this.Name = "frmIndex";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "frmIndex";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.frmIndex_FormClosed);
            this.Load += new System.EventHandler(this.frmIndex_Load);
            this.Suppression.ResumeLayout(false);
            this.Suppression.PerformLayout();
            this.Modification.ResumeLayout(false);
            this.Modification.PerformLayout();
            this.Ajout.ResumeLayout(false);
            this.Ajout.PerformLayout();
            this.PNL.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        internal System.Windows.Forms.OpenFileDialog OFD;
        internal System.Windows.Forms.DateTimePicker DTPS;
        internal System.Windows.Forms.Label Label8;
        internal System.Windows.Forms.ComboBox CmbSymboleS;
        internal System.Windows.Forms.Label Label5;
        internal System.Windows.Forms.TabPage Suppression;
        internal System.Windows.Forms.Button BtnSupprimer;
        internal System.Windows.Forms.TabPage Modification;
        internal System.Windows.Forms.DateTimePicker DTPM;
        internal System.Windows.Forms.ComboBox CmbSymboleM;
        internal System.Windows.Forms.Label Label7;
        internal System.Windows.Forms.Button BtnModifier;
        internal System.Windows.Forms.TextBox TxtValeurM;
        internal System.Windows.Forms.Label Label4;
        internal System.Windows.Forms.Label Label3;
        internal System.Windows.Forms.Label Label1;
        internal System.Windows.Forms.Label Label6;
        internal System.Windows.Forms.Button BtnAjouter;
        internal System.Windows.Forms.TextBox TxtValeurA;
        internal System.Windows.Forms.ProgressBar PB;
        internal System.Windows.Forms.Button BtnExcel;
        internal System.Windows.Forms.DateTimePicker DTPA;
        internal System.Windows.Forms.ComboBox CmbSymboleA;
        internal System.Windows.Forms.TabPage Ajout;
        internal System.Windows.Forms.Label Label2;
        internal System.Windows.Forms.TabControl PNL;
    }
}