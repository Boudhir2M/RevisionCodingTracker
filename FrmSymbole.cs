using System;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace Révision
{
    public partial class FrmSymbole : Form
    {
        OleDbConnection connection = new OleDbConnection();
        OleDbCommand records;
        DataTable DataTab = new DataTable();
        OleDbDataAdapter DataAdap = new OleDbDataAdapter();
        OleDbCommandBuilder ComndBuld;
        DataSet DataSetTab = new DataSet();
        string SQLSTR;
        long CodeCatégorie;
        string Symbole;
        public FrmSymbole()
        {
            InitializeComponent();
        }

        private void FrmSymbole_Load(object sender, EventArgs e)
        {
            if (connection.State == ConnectionState.Closed)
            {
                connection.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Application.StartupPath + "\\revision.accdb";
                connection.Open();
            }

            OleDbDataAdapter DataAdap = new OleDbDataAdapter("Select Code From symbole", connection);
            DataTable DataTab = new DataTable();
            DataAdap.Fill(DataTab);

            for (int i = 0; i < DataTab.Rows.Count; i++)
            {
                CmbSymbole.Items.Add(DataTab.Rows[i]["Code"].ToString());
                CmbSymboleM.Items.Add(DataTab.Rows[i]["Code"].ToString());
            }

            DataAdap = new OleDbDataAdapter("Select Intitulé From Catégorie", connection);
            DataTab = new DataTable();
            DataAdap.Fill(DataTab);

            for (int i = 0; i < DataTab.Rows.Count; i++)
            {
                CmbCatégorie.Items.Add(DataTab.Rows[i]["Intitulé"].ToString());
                CmbCatégorieM.Items.Add(DataTab.Rows[i]["Intitulé"].ToString());
            }
        }

        private void BtnAjouter_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(TxtSymbole.Text))
            {
                MessageBox.Show("Veuillez saisir un symbole..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                TxtSymbole.Focus();
                return;
            }
            if (string.IsNullOrEmpty(TxtDésignation.Text))
            {
                MessageBox.Show("Veuillez saisir une Désignation..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                TxtDésignation.Focus();
                return;
            }
            if (string.IsNullOrEmpty(CmbCatégorie.Text))
            {
                MessageBox.Show("Veuillez choisir une catégorie de la liste catégorie..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                CmbCatégorie.Focus();
                return;
            }

            using (OleDbDataAdapter DataAdap = new OleDbDataAdapter("Select * From Catégorie Where Intitulé='" + CmbCatégorie.Text + "'", connection))
            {
                DataTable DataTab = new DataTable();
                DataAdap.Fill(DataTab);

                if (DataTab.Rows.Count > 0)
                {
                    long CodeCatégorie = Convert.ToInt64(DataTab.Rows[0][0]);

                    // -------------- Ajouter --------------
                    string SQLSTR = "INSERT INTO symbole(Code,Intitulé,Catégorie) VALUES('";
                    SQLSTR += TxtSymbole.Text.Replace("'", "''") + "','";
                    SQLSTR += TxtDésignation.Text.Replace("'", "''") + "','";
                    SQLSTR += CodeCatégorie + "')";
                    ExecuterSQL(SQLSTR);
                    Recharger();

                    CmbCatégorie.Text = "";
                    TxtDésignation.Text = "";
                    TxtSymbole.Text = "";
                    TxtSymbole.Focus();
                }
            }
        }

        private void BtnModifier_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(TxtDésignationM.Text))
            {
                MessageBox.Show("Veuillez saisir une Désignation..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                TxtDésignationM.Focus();
                return;
            }
            if (string.IsNullOrEmpty(CmbCatégorieM.Text))
            {
                MessageBox.Show("Veuillez choisir une catégorie de la liste catégorie..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                CmbCatégorieM.Focus();
                return;
            }

            using (OleDbDataAdapter DataAdap = new OleDbDataAdapter("Select * From Catégorie Where Intitulé='" + CmbCatégorieM.Text + "'", connection))
            {
                DataTable DataTab = new DataTable();
                DataAdap.Fill(DataTab);

                if (DataTab.Rows.Count > 0)
                {
                    long CodeCatégorie = Convert.ToInt64(DataTab.Rows[0][0]);

                    // -------------- Modifier --------------
                    string SQLSTR = "UPDATE symbole SET ";
                    SQLSTR += "Intitulé = '" + TxtDésignationM.Text.Replace("'", "''") + "', ";
                    SQLSTR += "Catégorie = " + CodeCatégorie;
                    SQLSTR += " WHERE Code = '" + Symbole + "'";
                    ExecuterSQL(SQLSTR);
                }
            }
        }

        private void CmbCatégorieM_SelectedIndexChanged(object sender, EventArgs e)
        {
            using (OleDbDataAdapter DataAdap = new OleDbDataAdapter("Select * From Catégorie Where Intitulé='" + CmbCatégorieM.Text + "'", connection))
            {
                DataTable DataTab = new DataTable();
                DataAdap.Fill(DataTab);

                if (DataTab.Rows.Count > 0)
                {
                    CodeCatégorie = Convert.ToInt64(DataTab.Rows[0][0]);
                }
            }
        }

        private void CmbSymboleM_SelectedIndexChanged(object sender, EventArgs e)
        {
            using (OleDbDataAdapter DataAdap = new OleDbDataAdapter("Select * From Symboler Where Code='" + CmbSymboleM.Text + "'", connection))
            {
                DataTable DataTab = new DataTable();
                DataAdap.Fill(DataTab);

                if (DataTab.Rows.Count > 0)
                {
                    Symbole = DataTab.Rows[0][0].ToString();
                    TxtDésignationM.Text = DataTab.Rows[0][1].ToString();
                    CmbCatégorieM.Text = DataTab.Rows[0][2].ToString();
                    CodeCatégorie = Convert.ToInt64(DataTab.Rows[0][3]);
                }
            }
        }
        private void ExecuterSQL(string Strsql)
        {
            using (OleDbCommand CMND = connection.CreateCommand())
            {
                CMND.CommandText = Strsql;
                CMND.ExecuteNonQuery();
            }
        }

        private void CmbCatégorie_SelectedIndexChanged(object sender, EventArgs e)
        {
            using (OleDbDataAdapter DataAdap = new OleDbDataAdapter("Select * From Catégorie Where Intitulé='" + CmbCatégorie.Text + "'", connection))
            {
                DataTable DataTab = new DataTable();
                DataAdap.Fill(DataTab);
                if (DataTab.Rows.Count > 0)
                {
                    CodeCatégorie = Convert.ToInt64(DataTab.Rows[0][0]);
                }
            }
        }
        private void BtnSupprimer_Click(object sender, EventArgs e)
        {
            string SQLSTR = "DELETE FROM symbole WHERE Code = '" + CmbSymbole.Text + "'";
            ExecuterSQL(SQLSTR);
            Recharger();
        }
        private void Recharger()
        {
            CmbSymbole.Items.Clear();
            CmbSymboleM.Items.Clear();
            DataAdap = new OleDbDataAdapter("Select Code From symbole", connection);
            DataTab = new DataTable();
            DataAdap.Fill(DataTab);
            for (int i = 0; i < DataTab.Rows.Count; i++)
            {
                CmbSymbole.Items.Add(DataTab.Rows[i]["Code"].ToString());
                CmbSymboleM.Items.Add(DataTab.Rows[i]["Code"].ToString());
            }
        }

        private void FrmSymbole_Resize(object sender, EventArgs e)
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
