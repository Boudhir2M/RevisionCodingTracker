
namespace Révision
{
    partial class Décompte
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
            this.Label10 = new System.Windows.Forms.Label();
            this.BtnModifier = new System.Windows.Forms.Button();
            this.CBDécompte = new System.Windows.Forms.CheckBox();
            this.CmbLot = new System.Windows.Forms.ComboBox();
            this.CmbNumMarché = new System.Windows.Forms.ComboBox();
            this.BtnAjouter = new System.Windows.Forms.Button();
            this.TxtNumDécompte = new System.Windows.Forms.TextBox();
            this.Label9 = new System.Windows.Forms.Label();
            this.Label2 = new System.Windows.Forms.Label();
            this.CmbDécompte = new System.Windows.Forms.ComboBox();
            this.CmbNumMarchéM = new System.Windows.Forms.ComboBox();
            this.Label3 = new System.Windows.Forms.Label();
            this.Label1 = new System.Windows.Forms.Label();
            this.CBDécompteM = new System.Windows.Forms.CheckBox();
            this.TabModification = new System.Windows.Forms.TabPage();
            this.CmbLotM = new System.Windows.Forms.ComboBox();
            this.Label12 = new System.Windows.Forms.Label();
            this.TabAjout = new System.Windows.Forms.TabPage();
            this.PNL = new System.Windows.Forms.TabControl();
            this.TabModification.SuspendLayout();
            this.TabAjout.SuspendLayout();
            this.PNL.SuspendLayout();
            this.SuspendLayout();
            // 
            // Label10
            // 
            this.Label10.AutoSize = true;
            this.Label10.Location = new System.Drawing.Point(21, 88);
            this.Label10.Name = "Label10";
            this.Label10.Size = new System.Drawing.Size(28, 13);
            this.Label10.TabIndex = 39;
            this.Label10.Text = "Lot :";
            // 
            // BtnModifier
            // 
            this.BtnModifier.Location = new System.Drawing.Point(203, 124);
            this.BtnModifier.Name = "BtnModifier";
            this.BtnModifier.Size = new System.Drawing.Size(103, 36);
            this.BtnModifier.TabIndex = 34;
            this.BtnModifier.Text = "Modifier";
            this.BtnModifier.UseVisualStyleBackColor = true;
            this.BtnModifier.Click += new System.EventHandler(this.BtnModifier_Click);
            // 
            // CBDécompte
            // 
            this.CBDécompte.AutoSize = true;
            this.CBDécompte.Location = new System.Drawing.Point(24, 124);
            this.CBDécompte.Name = "CBDécompte";
            this.CBDécompte.Size = new System.Drawing.Size(60, 17);
            this.CBDécompte.TabIndex = 37;
            this.CBDécompte.Text = "Dérnier";
            this.CBDécompte.UseVisualStyleBackColor = true;
            // 
            // CmbLot
            // 
            this.CmbLot.FormattingEnabled = true;
            this.CmbLot.Location = new System.Drawing.Point(79, 87);
            this.CmbLot.Name = "CmbLot";
            this.CmbLot.Size = new System.Drawing.Size(227, 21);
            this.CmbLot.TabIndex = 36;
            // 
            // CmbNumMarché
            // 
            this.CmbNumMarché.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.CmbNumMarché.FormattingEnabled = true;
            this.CmbNumMarché.Location = new System.Drawing.Point(129, 16);
            this.CmbNumMarché.Name = "CmbNumMarché";
            this.CmbNumMarché.Size = new System.Drawing.Size(177, 21);
            this.CmbNumMarché.TabIndex = 36;
            this.CmbNumMarché.SelectedIndexChanged += new System.EventHandler(this.CmbNumMarché_SelectedIndexChanged);
            // 
            // BtnAjouter
            // 
            this.BtnAjouter.Location = new System.Drawing.Point(203, 124);
            this.BtnAjouter.Name = "BtnAjouter";
            this.BtnAjouter.Size = new System.Drawing.Size(103, 36);
            this.BtnAjouter.TabIndex = 12;
            this.BtnAjouter.Text = "Ajouter";
            this.BtnAjouter.UseVisualStyleBackColor = true;
            this.BtnAjouter.Click += new System.EventHandler(this.BtnAjouter_Click);
            // 
            // TxtNumDécompte
            // 
            this.TxtNumDécompte.Location = new System.Drawing.Point(129, 52);
            this.TxtNumDécompte.Name = "TxtNumDécompte";
            this.TxtNumDécompte.Size = new System.Drawing.Size(177, 20);
            this.TxtNumDécompte.TabIndex = 11;
            // 
            // Label9
            // 
            this.Label9.AutoSize = true;
            this.Label9.Location = new System.Drawing.Point(21, 54);
            this.Label9.Name = "Label9";
            this.Label9.Size = new System.Drawing.Size(102, 13);
            this.Label9.TabIndex = 3;
            this.Label9.Text = "Numéro Décompte :";
            // 
            // Label2
            // 
            this.Label2.AutoSize = true;
            this.Label2.Location = new System.Drawing.Point(21, 88);
            this.Label2.Name = "Label2";
            this.Label2.Size = new System.Drawing.Size(28, 13);
            this.Label2.TabIndex = 1;
            this.Label2.Text = "Lot :";
            // 
            // CmbDécompte
            // 
            this.CmbDécompte.FormattingEnabled = true;
            this.CmbDécompte.Location = new System.Drawing.Point(129, 52);
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
            this.Label3.Location = new System.Drawing.Point(21, 54);
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
            // CBDécompteM
            // 
            this.CBDécompteM.AutoSize = true;
            this.CBDécompteM.Location = new System.Drawing.Point(24, 124);
            this.CBDécompteM.Name = "CBDécompteM";
            this.CBDécompteM.Size = new System.Drawing.Size(60, 17);
            this.CBDécompteM.TabIndex = 48;
            this.CBDécompteM.Text = "Dérnier";
            this.CBDécompteM.UseVisualStyleBackColor = true;
            // 
            // TabModification
            // 
            this.TabModification.BackColor = System.Drawing.Color.Silver;
            this.TabModification.Controls.Add(this.CBDécompteM);
            this.TabModification.Controls.Add(this.CmbLotM);
            this.TabModification.Controls.Add(this.CmbDécompte);
            this.TabModification.Controls.Add(this.CmbNumMarchéM);
            this.TabModification.Controls.Add(this.Label3);
            this.TabModification.Controls.Add(this.Label10);
            this.TabModification.Controls.Add(this.Label12);
            this.TabModification.Controls.Add(this.BtnModifier);
            this.TabModification.Location = new System.Drawing.Point(4, 22);
            this.TabModification.Name = "TabModification";
            this.TabModification.Padding = new System.Windows.Forms.Padding(3);
            this.TabModification.Size = new System.Drawing.Size(327, 177);
            this.TabModification.TabIndex = 1;
            this.TabModification.Text = "Modification";
            // 
            // CmbLotM
            // 
            this.CmbLotM.FormattingEnabled = true;
            this.CmbLotM.Location = new System.Drawing.Point(79, 87);
            this.CmbLotM.Name = "CmbLotM";
            this.CmbLotM.Size = new System.Drawing.Size(227, 21);
            this.CmbLotM.TabIndex = 46;
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
            // TabAjout
            // 
            this.TabAjout.BackColor = System.Drawing.Color.Silver;
            this.TabAjout.Controls.Add(this.CBDécompte);
            this.TabAjout.Controls.Add(this.CmbLot);
            this.TabAjout.Controls.Add(this.CmbNumMarché);
            this.TabAjout.Controls.Add(this.BtnAjouter);
            this.TabAjout.Controls.Add(this.TxtNumDécompte);
            this.TabAjout.Controls.Add(this.Label9);
            this.TabAjout.Controls.Add(this.Label2);
            this.TabAjout.Controls.Add(this.Label1);
            this.TabAjout.Location = new System.Drawing.Point(4, 22);
            this.TabAjout.Name = "TabAjout";
            this.TabAjout.Padding = new System.Windows.Forms.Padding(3);
            this.TabAjout.Size = new System.Drawing.Size(327, 177);
            this.TabAjout.TabIndex = 0;
            this.TabAjout.Text = "Ajout";
            // 
            // PNL
            // 
            this.PNL.Controls.Add(this.TabAjout);
            this.PNL.Controls.Add(this.TabModification);
            this.PNL.Location = new System.Drawing.Point(12, 12);
            this.PNL.Name = "PNL";
            this.PNL.SelectedIndex = 0;
            this.PNL.Size = new System.Drawing.Size(335, 203);
            this.PNL.TabIndex = 3;
            // 
            // Décompte
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(358, 227);
            this.Controls.Add(this.PNL);
            this.Name = "Décompte";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Décomptes ...";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Décompte_FormClosed);
            this.Load += new System.EventHandler(this.Décompte_Load);
            this.Resize += new System.EventHandler(this.Décompte_Resize);
            this.TabModification.ResumeLayout(false);
            this.TabModification.PerformLayout();
            this.TabAjout.ResumeLayout(false);
            this.TabAjout.PerformLayout();
            this.PNL.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        internal System.Windows.Forms.Label Label10;
        internal System.Windows.Forms.Button BtnModifier;
        internal System.Windows.Forms.CheckBox CBDécompte;
        internal System.Windows.Forms.ComboBox CmbLot;
        internal System.Windows.Forms.ComboBox CmbNumMarché;
        internal System.Windows.Forms.Button BtnAjouter;
        internal System.Windows.Forms.TextBox TxtNumDécompte;
        internal System.Windows.Forms.Label Label9;
        internal System.Windows.Forms.Label Label2;
        internal System.Windows.Forms.ComboBox CmbDécompte;
        internal System.Windows.Forms.ComboBox CmbNumMarchéM;
        internal System.Windows.Forms.Label Label3;
        internal System.Windows.Forms.Label Label1;
        internal System.Windows.Forms.CheckBox CBDécompteM;
        internal System.Windows.Forms.TabPage TabModification;
        internal System.Windows.Forms.ComboBox CmbLotM;
        internal System.Windows.Forms.Label Label12;
        internal System.Windows.Forms.TabPage TabAjout;
        internal System.Windows.Forms.TabControl PNL;
    }
}