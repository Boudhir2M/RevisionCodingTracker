using System;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace Révision
{
    public partial class frmInfosEntrepreneur : Form
    {
        OleDbConnection connection = new OleDbConnection();
        OleDbCommand records;
        DataTable DataTab = new DataTable();
        OleDbDataAdapter DataAdap = new OleDbDataAdapter();
        OleDbCommandBuilder ComndBuld = new OleDbCommandBuilder();
        DataSet DataSetTab = new DataSet();
        string SQLSTR;
        long CodeCatégorie;
        string Symbole;
        public frmInfosEntrepreneur()
        {
            InitializeComponent();
        }

        private void frmInfosEntrepreneur_Load(object sender, EventArgs e)
        {
            if (connection.State == ConnectionState.Closed)
            {
                connection.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Application.StartupPath + "\\revision.accdb";
                connection.Open();
            }

            DataAdap = new OleDbDataAdapter("Select Nom From Entrepreneur", connection);
            DataTab = new DataTable();
            DataAdap.Fill(DataTab);

            for (int i = 0; i < DataTab.Rows.Count; i++)
            {
                CmbNomm.Items.Add(DataTab.Rows[i][0].ToString());
            }
        }

        private void CmbNomm_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataAdap = new OleDbDataAdapter("Select * From Entrepreneur where Nom ='" + CmbNomm.Text + "'", connection);
            DataTab = new DataTable();
            DataAdap.Fill(DataTab);

            TxtAdresseM.Text = DataTab.Rows[0][2].ToString();
            TxtCNSSM.Text = DataTab.Rows[0][3].ToString();
            TxtPatenteM.Text = DataTab.Rows[0][4].ToString();
            TxtCBM.Text = DataTab.Rows[0][5].ToString();
            TxtABM.Text = DataTab.Rows[0][6].ToString();
            TxtRCM.Text = DataTab.Rows[0][7].ToString();
            TxtTelM.Text = DataTab.Rows[0][8].ToString();
            TxtEMailM.Text = DataTab.Rows[0][9].ToString();
            TxtIFM.Text = DataTab.Rows[0][10].ToString();
            TxtICEM.Text = DataTab.Rows[0][11].ToString();
            CmbMarché.Items.Clear();

            DataAdap = new OleDbDataAdapter("Select * From Marché_Entrepreneur where Entrepreneur.Nom ='" + CmbNomm.Text + "'", connection);
            DataTab = new DataTable();
            DataAdap.Fill(DataTab);

            if (DataTab.Rows.Count > 0)
            {
                for (int i = 0; i < DataTab.Rows.Count; i++)
                {
                    CmbMarché.Items.Add(DataTab.Rows[i][0].ToString());
                }

                CmbMarché.Text = DataTab.Rows[0][0].ToString();
                CmbMarché_SelectedIndexChanged(sender, e);
            }
            else
            {
                MessageBox.Show("Aucun Marché n'est attribué à la société : " + CmbNomm.Text, "Attention", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }

        private void CmbMarché_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(CmbMarché.Text))
            {
                MessageBox.Show("Choisissez une valeur de la liste : Référence de Marché ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CmbMarché.Focus();
                return;
            }

            int i;
            DataAdap = new OleDbDataAdapter("Select * From Marché_Entrepreneur where Réference = '" + CmbMarché.Text + "'", connection);
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

            for (i = 0; i < DataTab.Rows.Count; i++)
            {
                CmbDécomptes.Items.Add(DataTab.Rows[i][0].ToString());
            }

            CmbDécomptes.Text = DataTab.Rows[0][0].ToString();
            CmbDécomptes_SelectedIndexChanged(sender, e);

        }
        private void CmbDécomptes_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (CmbDécomptes.Text == "")
            {
                MessageBox.Show("Choisissez une valeur de la liste : Réference de Marché ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CmbDécomptes.Focus();
                return;
            }

            long NumDécompte;
            DataAdap = new OleDbDataAdapter("Select N° From Décompte where Désignation='" + CmbDécomptes.Text + "' and Num_Marché='" + CmbMarché.Text + "'", connection);
            DataTab = new DataTable();
            DataAdap.Fill(DataTab);
            NumDécompte = Convert.ToInt64(DataTab.Rows[0][0]);

            DataAdap = new OleDbDataAdapter("Select Date_Chrono, Motif From Chrono where Num_Décompte=" + NumDécompte + " ORDER BY Ordre", connection);
            DataTab = new DataTable();
            DataAdap.Fill(DataTab);
            DGVC.DataSource = DataTab;
        }

        private void frmInfosEntrepreneur_Resize(object sender, EventArgs e)
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
