using System;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Windows.Forms;

namespace Révision
{
    public partial class FrmMarchéInfos : Form
    {
        private OleDbConnection connection = new OleDbConnection();
        private OleDbCommand records;
        private DataTable DataTab;
        private OleDbDataAdapter DataAdap;
        private OleDbCommandBuilder ComndBuld;
        private DataSet DataSetTab = new DataSet();
        private string SQLSTR;
        private long CodeCatégorie;
        private long Statut;
        private long NumSymbole;
        private long NumActeur;
        public FrmMarchéInfos()
        {
            InitializeComponent();
        }

        private void CmbDécomptes_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(CmbDécomptes.Text))
            {
                MessageBox.Show("Choisissez une valeur de la liste : Référence de Marché ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CmbDécomptes.Focus();
                return;
            }

            long NumDécompte;
            DataAdap = new OleDbDataAdapter("Select N° From Décompte where Désignation='" + CmbDécomptes.Text + "' and Num_Marché='" + CmbMarché.Text + "'", connection);
            DataTab = new DataTable();
            DataAdap.Fill(DataTab);
            NumDécompte = Convert.ToInt64(DataTab.Rows[0]["N°"]);

            DataAdap = new OleDbDataAdapter("Select Date_Chrono, Motif From Chrono where Num_Décompte=" + NumDécompte + " ORDER BY Date_Chrono", connection);
            DataTab = new DataTable();
            DataAdap.Fill(DataTab);
            DGVC.DataSource = DataTab;

            DataAdap = new OleDbDataAdapter("Select Date_Commencement, Date_Fin From Décompte where N°=" + NumDécompte, connection);
            DataTab = new DataTable();
            DataAdap.Fill(DataTab);
            DGVC.DataSource = DataTab;

            if (DataTab.Rows.Count > 0)
            {
                TxtDD.Text = DataTab.Rows[0]["Date_Commencement"].ToString();
                TxtDF.Text = DataTab.Rows[0]["Date_Fin"].ToString();
            }
        }
        private void CmbMarché_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (CmbMarché.Text == "")
            {
                MessageBox.Show("Choisissez une valeur de la liste : Référence de Marché ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                CmbMarché.Focus();
                return;
            }

            DataAdap = new OleDbDataAdapter("Select * From Marchér where Réference = '" + CmbMarché.Text + "'", connection);
            DataTab = new DataTable();
            DataAdap.Fill(DataTab);
            DGV.DataSource = DataTab;

            DataAdap = new OleDbDataAdapter("Select * From Décompte_Entrepreneur where Réference = '" + CmbMarché.Text + "'", connection);
            DataTab = new DataTable();
            DataAdap.Fill(DataTab);
            DGVD.DataSource = DataTab;

            DataAdap = new OleDbDataAdapter("Select DISTINCT Désignation From Décompte where Num_Marché = '" + CmbMarché.Text + "' ORDER BY Désignation", connection);
            DataTab = new DataTable();
            DataAdap.Fill(DataTab);
            CmbDécomptes.Items.Clear();
            if (DataTab.Rows.Count > 0)
            {
                for (int i = 0; i < DataTab.Rows.Count; i++)
                {
                    CmbDécomptes.Items.Add(DataTab.Rows[i][0].ToString());
                }
                CmbDécomptes.Text = DataTab.Rows[0][0].ToString();
                CmbDécomptes_SelectedIndexChanged(sender, e);
            }
            else
            {
                MessageBox.Show("Aucun décompte n'est ouvert pour le moment...", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }
        private void FrmMarchéInfos_Load(object sender, EventArgs e)
        {
            if (connection.State == ConnectionState.Closed)
            {
                connection.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data source=" + Application.StartupPath + "\\revision.accdb";
                connection.Open();
            }

            DataAdap = new OleDbDataAdapter("Select Réference From Marché", connection);
            DataTab = new DataTable();
            DataAdap.Fill(DataTab);

            for (int i = 0; i < DataTab.Rows.Count; i++)
            {
                CmbMarché.Items.Add(DataTab.Rows[i]["Réference"].ToString());
            }
            DGV.EnableHeadersVisualStyles = false;
            DGV.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            DGV.ColumnHeadersDefaultCellStyle.BackColor = Color.DarkGray;

            DGVD.EnableHeadersVisualStyles = false;
            DGVD.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            DGVD.ColumnHeadersDefaultCellStyle.BackColor = Color.DarkGray;

            DGVC.EnableHeadersVisualStyles = false;
            DGVC.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            DGVC.ColumnHeadersDefaultCellStyle.BackColor = Color.DarkGray;
        }

        private void FrmMarchéInfos_Resize(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Maximized)
            {
                PNL.Left = (this.Width - PNL.Width) / 2;
                PNL.Top = (this.Height - PNL.Height) / 2;
            }
            else
            {
                PNL.Left = 12;
                PNL.Top = 12;
            }
        }
    }
}
