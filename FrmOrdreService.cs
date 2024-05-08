using System;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace Révision
{
    public partial class FrmOrdreService : Form
    {
        OleDbConnection connection = new OleDbConnection();
        OleDbCommand records;
        DataTable DataTab = new DataTable();
        OleDbDataAdapter DataAdap;
        OleDbCommandBuilder ComndBuld;
        DataSet DataSetTab = new DataSet();
        string SQLSTR;
        long NumDécompte;
        long NumDécompteM;
        public FrmOrdreService()
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
                CmbNumMarché.Items.Add(DataTab.Rows[i]["Réference"].ToString());
                CmbNumMarchéM.Items.Add(DataTab.Rows[i]["Réference"].ToString());
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
                    MessageBox.Show("Remplissez la zone : Réference de marché..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    CmbNumMarché.Focus();
                    return;
                }

                // Utilisez DateTimePicker.Value pour obtenir la valeur de la date
                DateTime dateRepriseService = DTPDPS.Value;

                if (string.IsNullOrEmpty(CmbDécompteA.Text))
                {
                    MessageBox.Show("Remplissez la zone Numéro de Décompte ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    CmbDécompte.Focus();
                    return;
                }

                SQLSTR = $"INSERT INTO Ordre_Service(Num_Décompte, Date_Ordre_Service) VALUES ({NumDécompte}, '{dateRepriseService:dd/MM/yyyy}')";

                ExecuterSQL(SQLSTR);

                //------------ Ajouter -------------
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Aucune valeur n'a été saisie dans la table détail estimatif pour le marché numéro : {CmbNumMarchéM.Text} .", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void FrmOrdreService_FormClosed(object sender, FormClosedEventArgs e)
        {
            connection.Close();
        }
        private void FrmOrdreService_Load(object sender, EventArgs e)
        {
            try
            {
                if (connection.State == ConnectionState.Closed)
                {
                    connection.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data source=" + Application.StartupPath + "\\revision.accdb";
                    connection.Open();
                }

                Recharger();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void CmbNumMarché_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(CmbNumMarché.Text))
            {
                MessageBox.Show("Choisissez une valeur de la liste : Réference de Marché ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                CmbNumMarché.Focus();
                return;
            }

            CmbDécompteA.Items.Clear();
            DataAdap = new OleDbDataAdapter("Select Désignation From Décompte where Num_Marché='" + CmbNumMarché.Text + "'", connection);
            DataTab = new DataTable();
            DataAdap.Fill(DataTab);

            if (DataTab.Rows.Count > 0)
            {
                for (int i = 0; i < DataTab.Rows.Count; i++)
                {
                    CmbDécompteA.Items.Add(DataTab.Rows[i].ItemArray[0].ToString());
                }
            }
            else
            {
                CmbDécompte.Text = "";
                DTPDPS.Value = DateTime.Now;
                MessageBox.Show($"Aucune valeur n'a été saisie dans la table détail estimatif pour le marché numéro : {CmbNumMarchéM.Text} .", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void BtnModifier_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(CmbNumMarchéM.Text))
                {
                    MessageBox.Show("Remplissez la zone : Référence de marché..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    CmbNumMarchéM.Focus();
                    return;
                }

                // Utilisez DateTimePicker.Value pour obtenir la valeur de la date
                DateTime dateArretService = DTPDASM.Value;

                if (string.IsNullOrEmpty(CmbDécompte.Text))
                {
                    MessageBox.Show("Remplissez la zone Numéro de Décompte ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    CmbDécompte.Focus();
                    return;
                }

                //------------ Modifier -------------
                string S;
                DataAdap = new OleDbDataAdapter("Select * From Décompte where Num_Marché='" + CmbNumMarchéM.Text + "' and Désignation='" + CmbDécompte.Text + "'", connection);
                DataTab = new DataTable();
                DataAdap.Fill(DataTab);

                if (DataTab.Rows.Count > 0)
                {
                    int numDecompteM;
                    if (int.TryParse(DataTab.Rows[0]["N°"].ToString(), out numDecompteM))
                    {
                        DataAdap = new OleDbDataAdapter("Select N° From Ordre_Service where Num_Décompte=" + numDecompteM + " and Date_Ordre_Service='" + CmbDOS.Text + "'", connection);
                        DataTab = new DataTable();
                        DataAdap.Fill(DataTab);

                        if (DataTab.Rows.Count > 0)
                        {
                            SQLSTR = "UPDATE Ordre_Service SET ";
                            SQLSTR += "Date_Arret_Service='" + dateArretService.ToString("yyyy-MM-dd") + "'";
                            SQLSTR += " where  N°=" + DataTab.Rows[0]["N°"];
                            MessageBox.Show(SQLSTR);
                            ExecuterSQL(SQLSTR);
                        }
                        else
                        {
                            MessageBox.Show("Veuillez saisir la date de reprise en premier puis revenir pour la date d'Arrêt", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Numéro de décompte invalide", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("Numéro de décompte invalide", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
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
                MessageBox.Show("Choisissez une valeur de la liste : Référence de Marché ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                CmbNumMarchéM.Focus();
                return;
            }

            CmbDécompte.Items.Clear();
            DataAdap = new OleDbDataAdapter("Select Désignation From Décompte where Num_Marché='" + CmbNumMarchéM.Text + "'", connection);
            DataTab = new DataTable();
            DataAdap.Fill(DataTab);

            if (DataTab.Rows.Count > 0)
            {
                for (int i = 0; i < DataTab.Rows.Count; i++)
                {
                    CmbDécompte.Items.Add(DataTab.Rows[i]["Désignation"].ToString());
                }
            }
            else
            {
                CmbDécompte.Text = "";
                DTPDASM.Value = DateTime.Now;
                MessageBox.Show("Aucune valeur n'a été saisie dans la table détail estimatif pour le marché numéro : " + CmbNumMarchéM.Text + " .", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void CmbDécompte_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(CmbDécompte.Text))
            {
                MessageBox.Show("Choisissez une valeur de la liste : Référence de Marché ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                CmbDécompte.Focus();
                return;
            }

            DataAdap = new OleDbDataAdapter("Select * From Décompte where Num_Marché='" + CmbNumMarchéM.Text + "' and Désignation='" + CmbDécompte.Text + "'", connection);
            DataTab = new DataTable();
            DataAdap.Fill(DataTab);

            if (DataTab.Rows.Count > 0)
            {
                NumDécompteM = Convert.ToInt32(DataTab.Rows[0]["N°"]);

                DataAdap = new OleDbDataAdapter("Select * From Ordre_Service where Num_Décompte=" + NumDécompteM, connection);
                DataTab = new DataTable();
                DataAdap.Fill(DataTab);

                CmbDOS.Items.Clear();
                for (int i = 0; i < DataTab.Rows.Count; i++)
                {
                    CmbDOS.Items.Add(DataTab.Rows[i][2]);
                }
            }
        }

        private void CmbDécompteA_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(CmbDécompteA.Text))
            {
                MessageBox.Show("Choisissez une valeur de la liste : Numéro de décompte ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                CmbDécompte.Focus();
                return;
            }

            DataAdap = new OleDbDataAdapter("Select * From Décompte where Num_Marché='" + CmbNumMarché.Text + "' and Désignation='" + CmbDécompteA.Text + "'", connection);
            DataTab = new DataTable();
            DataAdap.Fill(DataTab);

            if (DataTab.Rows.Count > 0)
            {
                NumDécompte = long.Parse(DataTab.Rows[0]["N°"].ToString());
            }
        }

        private void FrmOrdreService_Resize(object sender, EventArgs e)
        {
            try
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
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
