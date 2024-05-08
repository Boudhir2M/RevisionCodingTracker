using System;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace Révision
{
    public partial class Entrepreneur : Form

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
        long NumSymbole;
        long NumIndexs;
        public Entrepreneur()
        {
            InitializeComponent();
        }
        private void Entrepreneur_Load(object sender, EventArgs e)
        {
            try
            {
                if (connection.State == ConnectionState.Closed)
                {
                    connection.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Application.StartupPath + "\\revision.accdb";
                    connection.Open();
                    MessageBox.Show("Connexion à la base de données réussie.", "Succès de la connexion", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                // Gérer les exceptions liées à la connexion.
                MessageBox.Show("Erreur de connexion : " + ex.Message, "Erreur de connexion", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            Recharger();
        }
        private void Recharger()
        {
            CmbNomm.Items.Clear();

            DataAdap = new OleDbDataAdapter("Select Nom From Entrepreneur", connection);
            DataTab = new DataTable();
            DataAdap.Fill(DataTab);

            for (int i = 0; i < DataTab.Rows.Count; i++)
            {
                CmbNomm.Items.Add(DataTab.Rows[i]["Nom"].ToString());
            }
        }
        private void ExecuterSQL(string Strsql)
        {
            OleDbCommand CMND = connection.CreateCommand();
            CMND.CommandText = Strsql;
            CMND.ExecuteNonQuery();
        }
        private void BtnAjouter_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(TxtNom.Text))
                {
                    MessageBox.Show("Remplissez la zone : Nom..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    TxtNom.Focus();
                    return;
                }

                if (string.IsNullOrEmpty(TxtAdresse.Text))
                {
                    MessageBox.Show("Remplissez la zone Adresse ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    TxtAdresse.Focus();
                    return;
                }

                if (string.IsNullOrEmpty(TxtCNSS.Text))
                {
                    MessageBox.Show("Remplissez la zone CNSS ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    TxtCNSS.Focus();
                    return;
                }

                if (string.IsNullOrEmpty(TxtICE.Text))
                {
                    MessageBox.Show("Remplissez la zone : Identifiant commun de l'entreprise ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    TxtICE.Focus();
                    return;
                }

                if (string.IsNullOrEmpty(TxtIF.Text))
                {
                    MessageBox.Show("Remplissez la zone : Identifiant Fiscal ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    TxtIF.Focus();
                    return;
                }

                if (string.IsNullOrEmpty(TxtRC.Text))
                {
                    MessageBox.Show("Remplissez la zone : Regitre de commerce ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    TxtRC.Focus();
                    return;
                }

                if (string.IsNullOrEmpty(TxtPatente.Text))
                {
                    MessageBox.Show("Remplissez la zone : Patente ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    TxtPatente.Focus();
                    return;
                }

                if (string.IsNullOrEmpty(TxtTel.Text))
                {
                    MessageBox.Show("Remplissez la zone : Téléphone ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    TxtTel.Focus();
                    return;
                }

                if (string.IsNullOrEmpty(TxtEMail.Text))
                {
                    MessageBox.Show("Remplissez la zone : EMail ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    TxtEMail.Focus();
                    return;
                }

                if (string.IsNullOrEmpty(TxtAB.Text))
                {
                    MessageBox.Show("Remplissez la zone : Agence Bancaire ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    TxtAB.Focus();
                    return;
                }

                if (string.IsNullOrEmpty(TxtCB.Text))
                {
                    MessageBox.Show("Remplissez la zone : Compte Bancaire ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    TxtCB.Focus();
                    return;
                }

                // -------------- Ajouter --------------
                SQLSTR = "INSERT INTO Entrepreneur(Nom,Adresse,CNSS,Patente,Compte_Bancaire_Num,Agence_Bancaire,Registre_Commerce,Téléphone,EMail,Identifiant_Fiscal,Identifiant_Commun_Entreprise) VALUES('";
                SQLSTR += TxtNom.Text.Replace("'", "''") + "','";
                SQLSTR += TxtAdresse.Text.Replace("'", "''") + "','";
                SQLSTR += TxtCNSS.Text.Replace("'", "''") + "','";
                SQLSTR += TxtPatente.Text.Replace("'", "''") + "','";
                SQLSTR += TxtCB.Text.Replace("'", "''") + "','";
                SQLSTR += TxtAB.Text.Replace("'", "''") + "','";
                SQLSTR += TxtRC.Text.Replace("'", "''") + "','";
                SQLSTR += TxtTel.Text.Replace("'", "''") + "','";
                SQLSTR += TxtEMail.Text.Replace("'", "''") + "','";
                SQLSTR += TxtIF.Text.Replace("'", "''") + "','";
                SQLSTR += TxtICE.Text.Replace("'", "''") + "')";

                ExecuterSQL(SQLSTR);
                Recharger();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void Entrepreneur_FormClosed(object sender, FormClosedEventArgs e)
        {
            connection.Close();
        }

        private void CmbNomm_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                DataAdap = new OleDbDataAdapter("Select * From Entrepreneur where Nom ='" + CmbNomm.Text + "'", connection);
                DataTab = new DataTable();
                DataAdap.Fill(DataTab);

                if (DataTab.Rows.Count > 0)
                {
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
                }
                else
                {
                    MessageBox.Show("Aucun enregistrement trouvé pour ce nom d'entrepreneur.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void BtnModifier_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(CmbNomm.Text))
                {
                    MessageBox.Show("Remplissez la zone : Nom..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    CmbNomm.Focus();
                    return;
                }

                if (string.IsNullOrEmpty(TxtAdresseM.Text))
                {
                    MessageBox.Show("Remplissez la zone Adresse ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    TxtAdresseM.Focus();
                    return;
                }

                if (string.IsNullOrEmpty(TxtCNSSM.Text))
                {
                    MessageBox.Show("Remplissez la zone CNSS ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    TxtCNSSM.Focus();
                    return;
                }

                if (string.IsNullOrEmpty(TxtICEM.Text))
                {
                    MessageBox.Show("Remplissez la zone : Identifiant de commerce de l'entreprise ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    TxtICEM.Focus();
                    return;
                }

                if (string.IsNullOrEmpty(TxtIFM.Text))
                {
                    MessageBox.Show("Remplissez la zone : Identifiant Fiscal ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    TxtIFM.Focus();
                    return;
                }

                if (string.IsNullOrEmpty(TxtRCM.Text))
                {
                    MessageBox.Show("Remplissez la zone : Regitre de commerce ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    TxtRCM.Focus();
                    return;
                }

                if (string.IsNullOrEmpty(TxtPatenteM.Text))
                {
                    MessageBox.Show("Remplissez la zone : Patente ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    TxtPatenteM.Focus();
                    return;
                }

                if (string.IsNullOrEmpty(TxtTelM.Text))
                {
                    MessageBox.Show("Remplissez la zone : Téléphone ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    TxtTelM.Focus();
                    return;
                }

                if (string.IsNullOrEmpty(TxtEMailM.Text))
                {
                    MessageBox.Show("Remplissez la zone : EMail ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    TxtEMailM.Focus();
                    return;
                }

                if (string.IsNullOrEmpty(TxtABM.Text))
                {
                    MessageBox.Show("Remplissez la zone : Agence Bancaire ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    TxtABM.Focus();
                    return;
                }

                if (string.IsNullOrEmpty(TxtCBM.Text))
                {
                    MessageBox.Show("Remplissez la zone : Compte Bancaire ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    TxtCBM.Focus();
                    return;
                }

                // -------------- Modifier --------------
                SQLSTR = "UPDATE Entrepreneur SET ";
                SQLSTR += "Adresse = '" + TxtAdresseM.Text.Replace("'", "''") + "',";
                SQLSTR += "CNSS = '" + TxtCNSSM.Text.Replace("'", "''") + "',";
                SQLSTR += "Patente = '" + TxtPatenteM.Text.Replace("'", "''") + "',";
                SQLSTR += "Compte_Bancaire_Num = '" + TxtCBM.Text.Replace("'", "''") + "',";
                SQLSTR += "Agence_Bancaire = '" + TxtABM.Text.Replace("'", "''") + "',";
                SQLSTR += "Registre_Commerce = '" + TxtRCM.Text.Replace("'", "''") + "',";
                SQLSTR += "Téléphone = '" + TxtTelM.Text.Replace("'", "''") + "',";
                SQLSTR += "EMail = '" + TxtEMailM.Text.Replace("'", "''") + "',";
                SQLSTR += "Identifiant_Fiscal = '" + TxtIFM.Text.Replace("'", "''") + "',";
                SQLSTR += "Identifiant_Commun_Entreprise = '" + TxtICEM.Text.Replace("'", "''") + "'";
                SQLSTR += " WHERE Nom = '" + CmbNomm.Text + "'";
                ExecuterSQL(SQLSTR);
                Recharger();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void Entrepreneur_Resize(object sender, EventArgs e)
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
