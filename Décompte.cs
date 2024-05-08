using System;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace Révision
{
    public partial class Décompte : Form
    {
        OleDbConnection connection = new OleDbConnection();
        OleDbCommand records;
        DataTable DataTab = new DataTable();
        OleDbDataAdapter DataAdap;
        OleDbCommandBuilder ComndBuld;
        DataSet DataSetTab = new DataSet();
        string SQLSTR;
        long NumSymbole;
        long NumIndexs;
        public Décompte()
        {
            InitializeComponent();
        }
        private void Recharger()
        {
            DataAdap = new OleDbDataAdapter("Select Réference From Marché", connection);
            DataTab = new DataTable();
            DataAdap.Fill(DataTab);

            for (int i = 0; i < DataTab.Rows.Count; i++)
            {
                CmbNumMarché.Items.Add(DataTab.Rows[i][0].ToString());
                CmbNumMarchéM.Items.Add(DataTab.Rows[i][0].ToString());
            }
        }
        private void ExecuterSQL(string Strsql)
        {
            OleDbCommand cmd = connection.CreateCommand();
            cmd.CommandText = Strsql;
            cmd.ExecuteNonQuery();
        }

        private void BtnAjouter_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(CmbNumMarché.Text))
                {
                    MessageBox.Show("Remplissez la zone : Référence de marché..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    CmbNumMarché.Focus();
                    return;
                }

                if (string.IsNullOrEmpty(CmbLot.Text))
                {
                    MessageBox.Show("Choisissez une valeur de la zone : Lot ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    CmbLot.Focus();
                    return;
                }

                if (string.IsNullOrEmpty(TxtNumDécompte.Text))
                {
                    MessageBox.Show("Remplissez la zone Numéro de Décompte ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    TxtNumDécompte.Focus();
                    return;
                }

                DataAdap = new OleDbDataAdapter("Select * From Décompte where Num_Marché='" + CmbNumMarchéM.Text + "' and Désignation='" + CmbDécompte.Text + "'", connection);
                DataTab = new DataTable();
                DataAdap.Fill(DataTab);
                string TD = CBDécompte.Checked ? "2" : "1";

                if (DataTab.Rows.Count == 0)
                {
                    // Ajouter
                    SQLSTR = "INSERT INTO Décompte(Désignation,Num_Marché,Type,Lot) VALUES('";
                    SQLSTR = SQLSTR + TxtNumDécompte.Text + "','";
                    SQLSTR = SQLSTR + CmbNumMarché.Text + "','";
                    SQLSTR = SQLSTR + TD + "',";
                    SQLSTR = SQLSTR + CmbLot.Text + ")";
                    ExecuterSQL(SQLSTR);
                    CmbNumMarché.Text = "";
                    CmbLot.Text = "";
                    CBDécompte.Checked = false;
                    TxtNumDécompte.Text = "";
                    TxtNumDécompte.Focus();
                }
                else
                {
                    MessageBox.Show("La valeur de Décompte suivante est déjà saisie.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void Décompte_FormClosed(object sender, FormClosedEventArgs e)
        {
            connection.Close();
        }

        private void Décompte_Load(object sender, EventArgs e)
        {
            if (connection.State == ConnectionState.Closed)
            {
                connection.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data source=" + Application.StartupPath + "\\revision.accdb";
                connection.Open();
            }

            Recharger();
        }

        private void CmbNumMarché_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(CmbNumMarché.Text))
            {
                MessageBox.Show("Choisissez une valeur de la liste : Référence de Marché ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                CmbNumMarché.Focus();
                return;
            }

            string query = "Select Lot From Marché where Référence=@Référence";
            using (OleDbCommand cmd = new OleDbCommand(query, connection))
            {
                cmd.Parameters.AddWithValue("@Référence", CmbNumMarché.Text);

                DataAdap = new OleDbDataAdapter(cmd);
                DataTab = new DataTable();
                DataAdap.Fill(DataTab);
                CmbLot.Items.Clear();
            }

            if (DataTab.Rows.Count > 0)
            {
                for (int i = 0; i < DataTab.Rows.Count; i++)
                {
                    CmbLot.Items.Add(i + 1);
                }
            }
            else
            {
                MessageBox.Show("Aucune valeur n'a été saisie dans la table détail estimatif pour le marché numéro : " + CmbNumMarchéM.Text + ".", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void BtnModifier_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(CmbNumMarchéM.Text))
                {
                    MessageBox.Show("Remplissez la zone : Référence de marché..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    CmbNumMarchéM.Focus();
                    return;
                }

                if (string.IsNullOrEmpty(CmbLotM.Text))
                {
                    MessageBox.Show("Choisissez une valeur de la zone : Lot ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    CmbLotM.Focus();
                    return;
                }

                if (string.IsNullOrEmpty(CmbDécompte.Text))
                {
                    MessageBox.Show("Remplissez la zone Numéro de Décompte ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    CmbDécompte.Focus();
                    return;
                }

                string TD = "";
                if (CBDécompte.Checked)
                {
                    TD = "2";
                }
                else
                {
                    TD = "1";
                }

                //------------ Modifier -------------
                SQLSTR = "UPDATE Décompte SET ";
                SQLSTR += "Lot = " + CmbLotM.Text + ",";
                SQLSTR += "Type = '" + TD + "' ";
                SQLSTR += "WHERE Num_Marché = '" + CmbNumMarchéM.Text + "' AND Désignation = '" + CmbDécompte.Text + "'";

                ExecuterSQL(SQLSTR);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void CmbNumMarchéM_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(CmbNumMarchéM.Text))
            {
                MessageBox.Show("Choisissez une valeur de la liste : Référence de Marché ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                CmbNumMarchéM.Focus();
                return;
            }

            DataAdap = new OleDbDataAdapter("Select Lot From Marché where Réference='" + CmbNumMarchéM.Text + "'", connection);
            DataTab = new DataTable();
            DataAdap.Fill(DataTab);

            if (DataTab.Rows.Count > 0)
            {
                CmbLotM.Items.Clear();
                for (int i = 0; i < DataTab.Rows.Count; i++)
                {
                    CmbLotM.Items.Add(i + 1);
                }
            }
            else
            {
                MessageBox.Show("Aucune valeur n'a été saisie dans la table détail éstimatif pour le marché numéro : " + CmbNumMarchéM.Text + ".", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            CmbDécompte.Items.Clear();
            DataAdap = new OleDbDataAdapter("Select Désignation From Décompte where Num_Marché='" + CmbNumMarchéM.Text + "'", connection);
            DataTab = new DataTable();
            DataAdap.Fill(DataTab);

            if (DataTab.Rows.Count > 0)
            {
                for (int i = 0; i < DataTab.Rows.Count; i++)
                {
                    CmbDécompte.Items.Add(DataTab.Rows[i][0].ToString());
                }
            }
            else
            {
                CmbDécompte.Text = "";
                CmbLotM.Text = "";
                MessageBox.Show("Aucune valeur n'a été saisie dans la table détail éstimatif pour le marché numéro : " + CmbNumMarchéM.Text + ".", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void CmbDécompte_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(CmbDécompte.Text))
            {
                MessageBox.Show("Choisissez une valeur de la liste : Référence de Marché ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                CmbDécompte.Focus();
                return;
            }

            DataAdap = new OleDbDataAdapter("Select * From Décompte where Num_Marché='" + CmbNumMarchéM.Text + "' and Désignation='" + CmbDécompte.Text + "'", connection);
            DataTab = new DataTable();
            DataAdap.Fill(DataTab);

            if (DataTab.Rows.Count > 0)
            {
                CmbLotM.Text = DataTab.Rows[0][3].ToString();
                string TD = DataTab.Rows[0][8].ToString();

                if (TD == "2")
                {
                    CBDécompteM.Checked = true;
                }
            }
        }

        private void Décompte_Resize(object sender, EventArgs e)
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
