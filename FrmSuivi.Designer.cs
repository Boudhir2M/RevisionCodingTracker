
namespace Révision
{
    partial class FrmSuivi
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
            this.TxtTAF = new System.Windows.Forms.TextBox();
            this.TxtTAP = new System.Windows.Forms.TextBox();
            this.BtnEnregistrer = new System.Windows.Forms.Button();
            this.DGV = new System.Windows.Forms.DataGridView();
            this.DGVC = new System.Windows.Forms.DataGridView();
            this.TxtDQ = new System.Windows.Forms.TextBox();
            this.TxtObservations = new System.Windows.Forms.TextBox();
            this.Label3 = new System.Windows.Forms.Label();
            this.LabDelaiGarantie = new System.Windows.Forms.Label();
            this.LabDurée = new System.Windows.Forms.Label();
            this.Label7 = new System.Windows.Forms.Label();
            this.Label10 = new System.Windows.Forms.Label();
            this.Label6 = new System.Windows.Forms.Label();
            this.Label5 = new System.Windows.Forms.Label();
            this.LabDélai = new System.Windows.Forms.Label();
            this.Label8 = new System.Windows.Forms.Label();
            this.Label4 = new System.Windows.Forms.Label();
            this.Label9 = new System.Windows.Forms.Label();
            this.Label2 = new System.Windows.Forms.Label();
            this.Label1 = new System.Windows.Forms.Label();
            this.CmbStatut = new System.Windows.Forms.ComboBox();
            this.Marché = new System.Windows.Forms.Label();
            this.CmbMarché = new System.Windows.Forms.ComboBox();
            this.PNL = new System.Windows.Forms.Panel();
            ((System.ComponentModel.ISupportInitialize)(this.DGV)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.DGVC)).BeginInit();
            this.PNL.SuspendLayout();
            this.SuspendLayout();
            // 
            // TxtTAF
            // 
            this.TxtTAF.Location = new System.Drawing.Point(556, 584);
            this.TxtTAF.MaxLength = 5;
            this.TxtTAF.Name = "TxtTAF";
            this.TxtTAF.Size = new System.Drawing.Size(100, 20);
            this.TxtTAF.TabIndex = 66;
            // 
            // TxtTAP
            // 
            this.TxtTAP.Location = new System.Drawing.Point(220, 584);
            this.TxtTAP.MaxLength = 5;
            this.TxtTAP.Name = "TxtTAP";
            this.TxtTAP.Size = new System.Drawing.Size(100, 20);
            this.TxtTAP.TabIndex = 66;
            // 
            // BtnEnregistrer
            // 
            this.BtnEnregistrer.Location = new System.Drawing.Point(1068, 17);
            this.BtnEnregistrer.Name = "BtnEnregistrer";
            this.BtnEnregistrer.Size = new System.Drawing.Size(116, 21);
            this.BtnEnregistrer.TabIndex = 65;
            this.BtnEnregistrer.Text = "Enregistrer";
            this.BtnEnregistrer.UseVisualStyleBackColor = true;
            // 
            // DGV
            // 
            this.DGV.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGV.Location = new System.Drawing.Point(95, 622);
            this.DGV.Name = "DGV";
            this.DGV.Size = new System.Drawing.Size(680, 94);
            this.DGV.TabIndex = 64;
            // 
            // DGVC
            // 
            this.DGVC.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGVC.Location = new System.Drawing.Point(784, 80);
            this.DGVC.Name = "DGVC";
            this.DGVC.Size = new System.Drawing.Size(400, 524);
            this.DGVC.TabIndex = 64;
            // 
            // TxtDQ
            // 
            this.TxtDQ.Location = new System.Drawing.Point(95, 417);
            this.TxtDQ.Multiline = true;
            this.TxtDQ.Name = "TxtDQ";
            this.TxtDQ.Size = new System.Drawing.Size(680, 161);
            this.TxtDQ.TabIndex = 63;
            // 
            // TxtObservations
            // 
            this.TxtObservations.Location = new System.Drawing.Point(98, 62);
            this.TxtObservations.Multiline = true;
            this.TxtObservations.Name = "TxtObservations";
            this.TxtObservations.Size = new System.Drawing.Size(680, 160);
            this.TxtObservations.TabIndex = 63;
            // 
            // Label3
            // 
            this.Label3.AutoSize = true;
            this.Label3.Location = new System.Drawing.Point(784, 62);
            this.Label3.Name = "Label3";
            this.Label3.Size = new System.Drawing.Size(47, 13);
            this.Label3.TabIndex = 62;
            this.Label3.Text = "Journal :";
            // 
            // LabDelaiGarantie
            // 
            this.LabDelaiGarantie.BackColor = System.Drawing.Color.White;
            this.LabDelaiGarantie.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.LabDelaiGarantie.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.LabDelaiGarantie.Location = new System.Drawing.Point(95, 345);
            this.LabDelaiGarantie.Name = "LabDelaiGarantie";
            this.LabDelaiGarantie.Size = new System.Drawing.Size(683, 47);
            this.LabDelaiGarantie.TabIndex = 62;
            this.LabDelaiGarantie.Text = "Délai :";
            // 
            // LabDurée
            // 
            this.LabDurée.BackColor = System.Drawing.Color.White;
            this.LabDurée.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.LabDurée.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.LabDurée.Location = new System.Drawing.Point(95, 290);
            this.LabDurée.Name = "LabDurée";
            this.LabDurée.Size = new System.Drawing.Size(683, 45);
            this.LabDurée.TabIndex = 62;
            this.LabDurée.Text = "Délai :";
            // 
            // Label7
            // 
            this.Label7.AutoSize = true;
            this.Label7.Location = new System.Drawing.Point(360, 587);
            this.Label7.Name = "Label7";
            this.Label7.Size = new System.Drawing.Size(153, 13);
            this.Label7.TabIndex = 62;
            this.Label7.Text = "Taux d\'avancement Financier :";
            // 
            // Label10
            // 
            this.Label10.AutoSize = true;
            this.Label10.Location = new System.Drawing.Point(17, 348);
            this.Label10.Name = "Label10";
            this.Label10.Size = new System.Drawing.Size(80, 13);
            this.Label10.TabIndex = 62;
            this.Label10.Text = "Délai Garantie :";
            // 
            // Label6
            // 
            this.Label6.AutoSize = true;
            this.Label6.Location = new System.Drawing.Point(24, 587);
            this.Label6.Name = "Label6";
            this.Label6.Size = new System.Drawing.Size(153, 13);
            this.Label6.TabIndex = 62;
            this.Label6.Text = "Taux d\'avancement Physique :";
            // 
            // Label5
            // 
            this.Label5.AutoSize = true;
            this.Label5.Location = new System.Drawing.Point(17, 293);
            this.Label5.Name = "Label5";
            this.Label5.Size = new System.Drawing.Size(42, 13);
            this.Label5.TabIndex = 62;
            this.Label5.Text = "Durée :";
            // 
            // LabDélai
            // 
            this.LabDélai.BackColor = System.Drawing.Color.White;
            this.LabDélai.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.LabDélai.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.LabDélai.Location = new System.Drawing.Point(95, 235);
            this.LabDélai.Name = "LabDélai";
            this.LabDélai.Size = new System.Drawing.Size(683, 44);
            this.LabDélai.TabIndex = 62;
            this.LabDélai.Text = "Délai :";
            // 
            // Label8
            // 
            this.Label8.AutoSize = true;
            this.Label8.Location = new System.Drawing.Point(280, 133);
            this.Label8.Name = "Label8";
            this.Label8.Size = new System.Drawing.Size(75, 13);
            this.Label8.TabIndex = 62;
            this.Label8.Text = "Obsérvations :";
            // 
            // Label4
            // 
            this.Label4.AutoSize = true;
            this.Label4.Location = new System.Drawing.Point(17, 238);
            this.Label4.Name = "Label4";
            this.Label4.Size = new System.Drawing.Size(37, 13);
            this.Label4.TabIndex = 62;
            this.Label4.Text = "Délai :";
            // 
            // Label9
            // 
            this.Label9.AutoSize = true;
            this.Label9.Location = new System.Drawing.Point(17, 401);
            this.Label9.Name = "Label9";
            this.Label9.Size = new System.Drawing.Size(189, 13);
            this.Label9.TabIndex = 62;
            this.Label9.Text = "Alerte de Dépassement des quantités :";
            // 
            // Label2
            // 
            this.Label2.AutoSize = true;
            this.Label2.Location = new System.Drawing.Point(17, 62);
            this.Label2.Name = "Label2";
            this.Label2.Size = new System.Drawing.Size(75, 13);
            this.Label2.TabIndex = 62;
            this.Label2.Text = "Obsérvations :";
            // 
            // Label1
            // 
            this.Label1.AutoSize = true;
            this.Label1.Location = new System.Drawing.Point(409, 21);
            this.Label1.Name = "Label1";
            this.Label1.Size = new System.Drawing.Size(41, 13);
            this.Label1.TabIndex = 61;
            this.Label1.Text = "Statut :";
            // 
            // CmbStatut
            // 
            this.CmbStatut.FormattingEnabled = true;
            this.CmbStatut.Location = new System.Drawing.Point(469, 18);
            this.CmbStatut.Name = "CmbStatut";
            this.CmbStatut.Size = new System.Drawing.Size(309, 21);
            this.CmbStatut.TabIndex = 60;
            // 
            // Marché
            // 
            this.Marché.AutoSize = true;
            this.Marché.Location = new System.Drawing.Point(17, 21);
            this.Marché.Name = "Marché";
            this.Marché.Size = new System.Drawing.Size(49, 13);
            this.Marché.TabIndex = 61;
            this.Marché.Text = "Marché :";
            // 
            // CmbMarché
            // 
            this.CmbMarché.FormattingEnabled = true;
            this.CmbMarché.Location = new System.Drawing.Point(77, 18);
            this.CmbMarché.Name = "CmbMarché";
            this.CmbMarché.Size = new System.Drawing.Size(309, 21);
            this.CmbMarché.TabIndex = 60;
            // 
            // PNL
            // 
            this.PNL.Controls.Add(this.TxtTAF);
            this.PNL.Controls.Add(this.TxtTAP);
            this.PNL.Controls.Add(this.BtnEnregistrer);
            this.PNL.Controls.Add(this.DGV);
            this.PNL.Controls.Add(this.DGVC);
            this.PNL.Controls.Add(this.TxtDQ);
            this.PNL.Controls.Add(this.TxtObservations);
            this.PNL.Controls.Add(this.Label3);
            this.PNL.Controls.Add(this.LabDelaiGarantie);
            this.PNL.Controls.Add(this.LabDurée);
            this.PNL.Controls.Add(this.Label7);
            this.PNL.Controls.Add(this.Label10);
            this.PNL.Controls.Add(this.Label6);
            this.PNL.Controls.Add(this.Label5);
            this.PNL.Controls.Add(this.LabDélai);
            this.PNL.Controls.Add(this.Label8);
            this.PNL.Controls.Add(this.Label4);
            this.PNL.Controls.Add(this.Label9);
            this.PNL.Controls.Add(this.Label2);
            this.PNL.Controls.Add(this.Label1);
            this.PNL.Controls.Add(this.CmbStatut);
            this.PNL.Controls.Add(this.Marché);
            this.PNL.Controls.Add(this.CmbMarché);
            this.PNL.Location = new System.Drawing.Point(12, 16);
            this.PNL.Name = "PNL";
            this.PNL.Size = new System.Drawing.Size(1196, 620);
            this.PNL.TabIndex = 1;
            // 
            // FrmSuivi
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1220, 652);
            this.Controls.Add(this.PNL);
            this.Name = "FrmSuivi";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "FrmSuivi";
            ((System.ComponentModel.ISupportInitialize)(this.DGV)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.DGVC)).EndInit();
            this.PNL.ResumeLayout(false);
            this.PNL.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal System.Windows.Forms.TextBox TxtTAF;
        internal System.Windows.Forms.TextBox TxtTAP;
        internal System.Windows.Forms.Button BtnEnregistrer;
        internal System.Windows.Forms.DataGridView DGV;
        internal System.Windows.Forms.DataGridView DGVC;
        internal System.Windows.Forms.TextBox TxtDQ;
        internal System.Windows.Forms.TextBox TxtObservations;
        internal System.Windows.Forms.Label Label3;
        internal System.Windows.Forms.Label LabDelaiGarantie;
        internal System.Windows.Forms.Label LabDurée;
        internal System.Windows.Forms.Label Label7;
        internal System.Windows.Forms.Label Label10;
        internal System.Windows.Forms.Label Label6;
        internal System.Windows.Forms.Label Label5;
        internal System.Windows.Forms.Label LabDélai;
        internal System.Windows.Forms.Label Label8;
        internal System.Windows.Forms.Label Label4;
        internal System.Windows.Forms.Label Label9;
        internal System.Windows.Forms.Label Label2;
        internal System.Windows.Forms.Label Label1;
        internal System.Windows.Forms.ComboBox CmbStatut;
        internal System.Windows.Forms.Label Marché;
        internal System.Windows.Forms.ComboBox CmbMarché;
        internal System.Windows.Forms.Panel PNL;
    }
}