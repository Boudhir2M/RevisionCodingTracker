
namespace Révision
{
    partial class FrmMarchéInfos
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
            this.DGVC = new System.Windows.Forms.DataGridView();
            this.GroupBox2 = new System.Windows.Forms.GroupBox();
            this.DGVD = new System.Windows.Forms.DataGridView();
            this.GroupBox1 = new System.Windows.Forms.GroupBox();
            this.DGV = new System.Windows.Forms.DataGridView();
            this.Panel1 = new System.Windows.Forms.Panel();
            this.Label1 = new System.Windows.Forms.Label();
            this.Marché = new System.Windows.Forms.Label();
            this.CmbDécomptes = new System.Windows.Forms.ComboBox();
            this.CmbMarché = new System.Windows.Forms.ComboBox();
            this.TxtDF = new System.Windows.Forms.TextBox();
            this.TxtDD = new System.Windows.Forms.TextBox();
            this.Label19 = new System.Windows.Forms.Label();
            this.Label21 = new System.Windows.Forms.Label();
            this.PNL.SuspendLayout();
            this.GroupBox3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DGVC)).BeginInit();
            this.GroupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DGVD)).BeginInit();
            this.GroupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DGV)).BeginInit();
            this.Panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // PNL
            // 
            this.PNL.BackColor = System.Drawing.Color.Silver;
            this.PNL.Controls.Add(this.GroupBox3);
            this.PNL.Controls.Add(this.GroupBox2);
            this.PNL.Controls.Add(this.GroupBox1);
            this.PNL.Controls.Add(this.Panel1);
            this.PNL.Location = new System.Drawing.Point(12, 21);
            this.PNL.Name = "PNL";
            this.PNL.Size = new System.Drawing.Size(1272, 486);
            this.PNL.TabIndex = 9;
            // 
            // GroupBox3
            // 
            this.GroupBox3.Controls.Add(this.DGVC);
            this.GroupBox3.Location = new System.Drawing.Point(765, 191);
            this.GroupBox3.Name = "GroupBox3";
            this.GroupBox3.Size = new System.Drawing.Size(494, 280);
            this.GroupBox3.TabIndex = 9;
            this.GroupBox3.TabStop = false;
            this.GroupBox3.Text = "Ordre de Service :";
            // 
            // DGVC
            // 
            this.DGVC.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGVC.Location = new System.Drawing.Point(23, 19);
            this.DGVC.Name = "DGVC";
            this.DGVC.Size = new System.Drawing.Size(458, 246);
            this.DGVC.TabIndex = 0;
            // 
            // GroupBox2
            // 
            this.GroupBox2.Controls.Add(this.DGVD);
            this.GroupBox2.Location = new System.Drawing.Point(12, 191);
            this.GroupBox2.Name = "GroupBox2";
            this.GroupBox2.Size = new System.Drawing.Size(747, 280);
            this.GroupBox2.TabIndex = 11;
            this.GroupBox2.TabStop = false;
            this.GroupBox2.Text = "Décomptes :";
            // 
            // DGVD
            // 
            this.DGVD.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGVD.Location = new System.Drawing.Point(23, 19);
            this.DGVD.Name = "DGVD";
            this.DGVD.Size = new System.Drawing.Size(711, 246);
            this.DGVD.TabIndex = 0;
            // 
            // GroupBox1
            // 
            this.GroupBox1.Controls.Add(this.DGV);
            this.GroupBox1.Location = new System.Drawing.Point(12, 84);
            this.GroupBox1.Name = "GroupBox1";
            this.GroupBox1.Size = new System.Drawing.Size(1244, 101);
            this.GroupBox1.TabIndex = 10;
            this.GroupBox1.TabStop = false;
            this.GroupBox1.Text = "Marchés :";
            // 
            // DGV
            // 
            this.DGV.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGV.Location = new System.Drawing.Point(23, 19);
            this.DGV.Name = "DGV";
            this.DGV.Size = new System.Drawing.Size(1211, 66);
            this.DGV.TabIndex = 0;
            // 
            // Panel1
            // 
            this.Panel1.Controls.Add(this.Label1);
            this.Panel1.Controls.Add(this.Marché);
            this.Panel1.Controls.Add(this.CmbDécomptes);
            this.Panel1.Controls.Add(this.CmbMarché);
            this.Panel1.Controls.Add(this.TxtDF);
            this.Panel1.Controls.Add(this.TxtDD);
            this.Panel1.Controls.Add(this.Label19);
            this.Panel1.Controls.Add(this.Label21);
            this.Panel1.Location = new System.Drawing.Point(12, 12);
            this.Panel1.Name = "Panel1";
            this.Panel1.Size = new System.Drawing.Size(1244, 56);
            this.Panel1.TabIndex = 8;
            // 
            // Label1
            // 
            this.Label1.AutoSize = true;
            this.Label1.Location = new System.Drawing.Point(408, 21);
            this.Label1.Name = "Label1";
            this.Label1.Size = new System.Drawing.Size(62, 13);
            this.Label1.TabIndex = 59;
            this.Label1.Text = "Décompte :";
            // 
            // Marché
            // 
            this.Marché.AutoSize = true;
            this.Marché.Location = new System.Drawing.Point(17, 21);
            this.Marché.Name = "Marché";
            this.Marché.Size = new System.Drawing.Size(49, 13);
            this.Marché.TabIndex = 59;
            this.Marché.Text = "Marché :";
            // 
            // CmbDécomptes
            // 
            this.CmbDécomptes.FormattingEnabled = true;
            this.CmbDécomptes.Location = new System.Drawing.Point(488, 18);
            this.CmbDécomptes.Name = "CmbDécomptes";
            this.CmbDécomptes.Size = new System.Drawing.Size(91, 21);
            this.CmbDécomptes.TabIndex = 58;
            this.CmbDécomptes.SelectedIndexChanged += new System.EventHandler(this.CmbDécomptes_SelectedIndexChanged);
            // 
            // CmbMarché
            // 
            this.CmbMarché.FormattingEnabled = true;
            this.CmbMarché.Location = new System.Drawing.Point(77, 18);
            this.CmbMarché.Name = "CmbMarché";
            this.CmbMarché.Size = new System.Drawing.Size(309, 21);
            this.CmbMarché.TabIndex = 58;
            this.CmbMarché.SelectedIndexChanged += new System.EventHandler(this.CmbMarché_SelectedIndexChanged);
            // 
            // TxtDF
            // 
            this.TxtDF.Location = new System.Drawing.Point(1038, 24);
            this.TxtDF.Name = "TxtDF";
            this.TxtDF.Size = new System.Drawing.Size(146, 20);
            this.TxtDF.TabIndex = 54;
            // 
            // TxtDD
            // 
            this.TxtDD.Location = new System.Drawing.Point(721, 21);
            this.TxtDD.Name = "TxtDD";
            this.TxtDD.Size = new System.Drawing.Size(146, 20);
            this.TxtDD.TabIndex = 53;
            // 
            // Label19
            // 
            this.Label19.AutoSize = true;
            this.Label19.Location = new System.Drawing.Point(595, 21);
            this.Label19.Name = "Label19";
            this.Label19.Size = new System.Drawing.Size(120, 13);
            this.Label19.TabIndex = 39;
            this.Label19.Text = "Date Début Décompte :";
            // 
            // Label21
            // 
            this.Label21.AutoSize = true;
            this.Label21.Location = new System.Drawing.Point(899, 24);
            this.Label21.Name = "Label21";
            this.Label21.Size = new System.Drawing.Size(105, 13);
            this.Label21.TabIndex = 37;
            this.Label21.Text = "Date Fin Décompte :";
            // 
            // FrmMarchéInfos
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1296, 529);
            this.Controls.Add(this.PNL);
            this.Name = "FrmMarchéInfos";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "FrmMarchéInfos";
            this.Load += new System.EventHandler(this.FrmMarchéInfos_Load);
            this.Resize += new System.EventHandler(this.FrmMarchéInfos_Resize);
            this.PNL.ResumeLayout(false);
            this.GroupBox3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.DGVC)).EndInit();
            this.GroupBox2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.DGVD)).EndInit();
            this.GroupBox1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.DGV)).EndInit();
            this.Panel1.ResumeLayout(false);
            this.Panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal System.Windows.Forms.Panel PNL;
        internal System.Windows.Forms.GroupBox GroupBox3;
        internal System.Windows.Forms.DataGridView DGVC;
        internal System.Windows.Forms.GroupBox GroupBox2;
        internal System.Windows.Forms.DataGridView DGVD;
        internal System.Windows.Forms.GroupBox GroupBox1;
        internal System.Windows.Forms.DataGridView DGV;
        internal System.Windows.Forms.Panel Panel1;
        internal System.Windows.Forms.Label Label1;
        internal System.Windows.Forms.Label Marché;
        internal System.Windows.Forms.ComboBox CmbDécomptes;
        internal System.Windows.Forms.ComboBox CmbMarché;
        internal System.Windows.Forms.TextBox TxtDF;
        internal System.Windows.Forms.TextBox TxtDD;
        internal System.Windows.Forms.Label Label19;
        internal System.Windows.Forms.Label Label21;
    }
}