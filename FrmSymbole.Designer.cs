
namespace Révision
{
    partial class FrmSymbole
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
            this.CmbSymbole = new System.Windows.Forms.ComboBox();
            this.BtnSupprimer = new System.Windows.Forms.Button();
            this.Suppression = new System.Windows.Forms.TabPage();
            this.Label5 = new System.Windows.Forms.Label();
            this.Modification = new System.Windows.Forms.TabPage();
            this.CmbSymboleM = new System.Windows.Forms.ComboBox();
            this.CmbCatégorieM = new System.Windows.Forms.ComboBox();
            this.Label7 = new System.Windows.Forms.Label();
            this.BtnModifier = new System.Windows.Forms.Button();
            this.TxtDésignationM = new System.Windows.Forms.TextBox();
            this.Label4 = new System.Windows.Forms.Label();
            this.Label3 = new System.Windows.Forms.Label();
            this.Label2 = new System.Windows.Forms.Label();
            this.Label1 = new System.Windows.Forms.Label();
            this.Label6 = new System.Windows.Forms.Label();
            this.BtnAjouter = new System.Windows.Forms.Button();
            this.TxtDésignation = new System.Windows.Forms.TextBox();
            this.CmbCatégorie = new System.Windows.Forms.ComboBox();
            this.Ajout = new System.Windows.Forms.TabPage();
            this.TxtSymbole = new System.Windows.Forms.TextBox();
            this.PNL = new System.Windows.Forms.TabControl();
            this.Suppression.SuspendLayout();
            this.Modification.SuspendLayout();
            this.Ajout.SuspendLayout();
            this.PNL.SuspendLayout();
            this.SuspendLayout();
            // 
            // CmbSymbole
            // 
            this.CmbSymbole.FormattingEnabled = true;
            this.CmbSymbole.Location = new System.Drawing.Point(102, 23);
            this.CmbSymbole.Name = "CmbSymbole";
            this.CmbSymbole.Size = new System.Drawing.Size(261, 21);
            this.CmbSymbole.TabIndex = 5;
            // 
            // BtnSupprimer
            // 
            this.BtnSupprimer.Location = new System.Drawing.Point(245, 201);
            this.BtnSupprimer.Name = "BtnSupprimer";
            this.BtnSupprimer.Size = new System.Drawing.Size(113, 28);
            this.BtnSupprimer.TabIndex = 3;
            this.BtnSupprimer.Text = "Supprimer";
            this.BtnSupprimer.UseVisualStyleBackColor = true;
            this.BtnSupprimer.Click += new System.EventHandler(this.BtnSupprimer_Click);
            // 
            // Suppression
            // 
            this.Suppression.BackColor = System.Drawing.Color.Silver;
            this.Suppression.Controls.Add(this.CmbSymbole);
            this.Suppression.Controls.Add(this.Label5);
            this.Suppression.Controls.Add(this.BtnSupprimer);
            this.Suppression.Location = new System.Drawing.Point(4, 22);
            this.Suppression.Name = "Suppression";
            this.Suppression.Size = new System.Drawing.Size(402, 258);
            this.Suppression.TabIndex = 2;
            this.Suppression.Text = "Suppression";
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
            // Modification
            // 
            this.Modification.BackColor = System.Drawing.Color.Silver;
            this.Modification.Controls.Add(this.CmbSymboleM);
            this.Modification.Controls.Add(this.CmbCatégorieM);
            this.Modification.Controls.Add(this.Label7);
            this.Modification.Controls.Add(this.BtnModifier);
            this.Modification.Controls.Add(this.TxtDésignationM);
            this.Modification.Controls.Add(this.Label4);
            this.Modification.Controls.Add(this.Label3);
            this.Modification.Location = new System.Drawing.Point(4, 22);
            this.Modification.Name = "Modification";
            this.Modification.Padding = new System.Windows.Forms.Padding(3);
            this.Modification.Size = new System.Drawing.Size(402, 258);
            this.Modification.TabIndex = 1;
            this.Modification.Text = "Modification";
            // 
            // CmbSymboleM
            // 
            this.CmbSymboleM.FormattingEnabled = true;
            this.CmbSymboleM.Location = new System.Drawing.Point(102, 23);
            this.CmbSymboleM.Name = "CmbSymboleM";
            this.CmbSymboleM.Size = new System.Drawing.Size(261, 21);
            this.CmbSymboleM.TabIndex = 10;
            this.CmbSymboleM.SelectedIndexChanged += new System.EventHandler(this.CmbSymboleM_SelectedIndexChanged);
            // 
            // CmbCatégorieM
            // 
            this.CmbCatégorieM.FormattingEnabled = true;
            this.CmbCatégorieM.Location = new System.Drawing.Point(102, 149);
            this.CmbCatégorieM.Name = "CmbCatégorieM";
            this.CmbCatégorieM.Size = new System.Drawing.Size(261, 21);
            this.CmbCatégorieM.TabIndex = 10;
            this.CmbCatégorieM.SelectedIndexChanged += new System.EventHandler(this.CmbCatégorieM_SelectedIndexChanged);
            // 
            // Label7
            // 
            this.Label7.AutoSize = true;
            this.Label7.Location = new System.Drawing.Point(27, 157);
            this.Label7.Name = "Label7";
            this.Label7.Size = new System.Drawing.Size(58, 13);
            this.Label7.TabIndex = 9;
            this.Label7.Text = "Catégorie :";
            // 
            // BtnModifier
            // 
            this.BtnModifier.Location = new System.Drawing.Point(245, 201);
            this.BtnModifier.Name = "BtnModifier";
            this.BtnModifier.Size = new System.Drawing.Size(113, 28);
            this.BtnModifier.TabIndex = 6;
            this.BtnModifier.Text = "Modifier";
            this.BtnModifier.UseVisualStyleBackColor = true;
            this.BtnModifier.Click += new System.EventHandler(this.BtnModifier_Click);
            // 
            // TxtDésignationM
            // 
            this.TxtDésignationM.Location = new System.Drawing.Point(102, 56);
            this.TxtDésignationM.Multiline = true;
            this.TxtDésignationM.Name = "TxtDésignationM";
            this.TxtDésignationM.Size = new System.Drawing.Size(261, 78);
            this.TxtDésignationM.TabIndex = 4;
            // 
            // Label4
            // 
            this.Label4.AutoSize = true;
            this.Label4.Location = new System.Drawing.Point(28, 59);
            this.Label4.Name = "Label4";
            this.Label4.Size = new System.Drawing.Size(69, 13);
            this.Label4.TabIndex = 2;
            this.Label4.Text = "Désignation :";
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
            // Label2
            // 
            this.Label2.AutoSize = true;
            this.Label2.Location = new System.Drawing.Point(28, 59);
            this.Label2.Name = "Label2";
            this.Label2.Size = new System.Drawing.Size(69, 13);
            this.Label2.TabIndex = 2;
            this.Label2.Text = "Désignation :";
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
            this.Label6.Location = new System.Drawing.Point(27, 157);
            this.Label6.Name = "Label6";
            this.Label6.Size = new System.Drawing.Size(58, 13);
            this.Label6.TabIndex = 7;
            this.Label6.Text = "Catégorie :";
            // 
            // BtnAjouter
            // 
            this.BtnAjouter.Location = new System.Drawing.Point(245, 201);
            this.BtnAjouter.Name = "BtnAjouter";
            this.BtnAjouter.Size = new System.Drawing.Size(113, 28);
            this.BtnAjouter.TabIndex = 6;
            this.BtnAjouter.Text = "Ajouter";
            this.BtnAjouter.UseVisualStyleBackColor = true;
            this.BtnAjouter.Click += new System.EventHandler(this.BtnAjouter_Click);
            // 
            // TxtDésignation
            // 
            this.TxtDésignation.Location = new System.Drawing.Point(102, 56);
            this.TxtDésignation.Multiline = true;
            this.TxtDésignation.Name = "TxtDésignation";
            this.TxtDésignation.Size = new System.Drawing.Size(256, 77);
            this.TxtDésignation.TabIndex = 4;
            // 
            // CmbCatégorie
            // 
            this.CmbCatégorie.FormattingEnabled = true;
            this.CmbCatégorie.Location = new System.Drawing.Point(102, 149);
            this.CmbCatégorie.Name = "CmbCatégorie";
            this.CmbCatégorie.Size = new System.Drawing.Size(257, 21);
            this.CmbCatégorie.TabIndex = 8;
            this.CmbCatégorie.SelectedIndexChanged += new System.EventHandler(this.CmbCatégorie_SelectedIndexChanged);
            // 
            // Ajout
            // 
            this.Ajout.BackColor = System.Drawing.Color.Silver;
            this.Ajout.Controls.Add(this.CmbCatégorie);
            this.Ajout.Controls.Add(this.Label6);
            this.Ajout.Controls.Add(this.BtnAjouter);
            this.Ajout.Controls.Add(this.TxtDésignation);
            this.Ajout.Controls.Add(this.TxtSymbole);
            this.Ajout.Controls.Add(this.Label2);
            this.Ajout.Controls.Add(this.Label1);
            this.Ajout.Location = new System.Drawing.Point(4, 22);
            this.Ajout.Name = "Ajout";
            this.Ajout.Padding = new System.Windows.Forms.Padding(3);
            this.Ajout.Size = new System.Drawing.Size(402, 258);
            this.Ajout.TabIndex = 0;
            this.Ajout.Text = "Ajout";
            // 
            // TxtSymbole
            // 
            this.TxtSymbole.Location = new System.Drawing.Point(102, 23);
            this.TxtSymbole.Name = "TxtSymbole";
            this.TxtSymbole.Size = new System.Drawing.Size(256, 20);
            this.TxtSymbole.TabIndex = 5;
            // 
            // PNL
            // 
            this.PNL.Controls.Add(this.Ajout);
            this.PNL.Controls.Add(this.Modification);
            this.PNL.Controls.Add(this.Suppression);
            this.PNL.Location = new System.Drawing.Point(1, 1);
            this.PNL.Name = "PNL";
            this.PNL.SelectedIndex = 0;
            this.PNL.Size = new System.Drawing.Size(410, 284);
            this.PNL.TabIndex = 4;
            // 
            // FrmSymbole
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(413, 286);
            this.Controls.Add(this.PNL);
            this.Name = "FrmSymbole";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "FrmSymbole";
            this.Load += new System.EventHandler(this.FrmSymbole_Load);
            this.Resize += new System.EventHandler(this.FrmSymbole_Resize);
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

        internal System.Windows.Forms.ComboBox CmbSymbole;
        internal System.Windows.Forms.Button BtnSupprimer;
        internal System.Windows.Forms.TabPage Suppression;
        internal System.Windows.Forms.Label Label5;
        internal System.Windows.Forms.TabPage Modification;
        internal System.Windows.Forms.ComboBox CmbSymboleM;
        internal System.Windows.Forms.ComboBox CmbCatégorieM;
        internal System.Windows.Forms.Label Label7;
        internal System.Windows.Forms.Button BtnModifier;
        internal System.Windows.Forms.TextBox TxtDésignationM;
        internal System.Windows.Forms.Label Label4;
        internal System.Windows.Forms.Label Label3;
        internal System.Windows.Forms.Label Label2;
        internal System.Windows.Forms.Label Label1;
        internal System.Windows.Forms.Label Label6;
        internal System.Windows.Forms.Button BtnAjouter;
        internal System.Windows.Forms.TextBox TxtDésignation;
        internal System.Windows.Forms.ComboBox CmbCatégorie;
        internal System.Windows.Forms.TabPage Ajout;
        internal System.Windows.Forms.TextBox TxtSymbole;
        internal System.Windows.Forms.TabControl PNL;
    }
}