using System;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace Révision
{
    public partial class Marché : Form
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
        private long MaitreOuvrage;
        private long MaitreOeuvre;
        private long NumSymbole;
        private long NumActeur;
        public Marché()
        {
            InitializeComponent();
        }

        private void Marché_FormClosed(object sender, FormClosedEventArgs e)
        {
            connection.Close();
        }
        private void Marché_Load(object sender, EventArgs e)
        {
            if (connection.State == ConnectionState.Closed)
            {
                connection.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data source=" + Application.StartupPath + "\\revision.accdb";
                connection.Open();
            }

            DataAdap = new OleDbDataAdapter("Select Nom From Entrepreneur", connection);
            DataTab = new DataTable();
            DataAdap.Fill(DataTab);

            for (int i = 0; i < DataTab.Rows.Count; i++)
            {
                CmbMEM.Items.Add(DataTab.Rows[i][0].ToString());
                CmbME.Items.Add(DataTab.Rows[i][0].ToString());
            }

            CmbMOM.Items.Add("Conseil Provincial Al Haouz");
            CmbMO.Items.Add("Conseil Provincial Al Haouz");

            DataAdap = new OleDbDataAdapter("Select Nom From StatutMarché", connection);
            DataTab = new DataTable();
            DataAdap.Fill(DataTab);

            for (int i = 0; i < DataTab.Rows.Count; i++)
            {
                CmbStatutM.Items.Add(DataTab.Rows[i][0].ToString());
                CmbStatut.Items.Add(DataTab.Rows[i][0].ToString());
            }

            recharger();
        }
        private void recharger()
        {
            CmbNumMarchéM.Items.Clear();
            DataAdap = new OleDbDataAdapter("Select Réference From Marché", connection);
            DataTab = new DataTable();
            DataAdap.Fill(DataTab);

            for (int i = 0; i < DataTab.Rows.Count; i++)
            {
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
                if (string.IsNullOrEmpty(TxtNumMarché.Text))
                {
                    MessageBox.Show("Remplissez la zone : Référence..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    TxtNumMarché.Focus();
                    return;
                }
                if (string.IsNullOrEmpty(TxtObjet.Text))
                {
                    MessageBox.Show("Remplissez la zone Objet ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    TxtObjet.Focus();
                    return;
                }
                if (string.IsNullOrEmpty(TxtEstimation.Text))
                {
                    MessageBox.Show("Remplissez la zone : Estimation ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    TxtEstimation.Focus();
                    return;
                }
                if (string.IsNullOrEmpty(TxtMontant.Text))
                {
                    MessageBox.Show("Remplissez la zone Montant ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    TxtMontant.Focus();
                    return;
                }
                if (string.IsNullOrEmpty(TxtDélai.Text))
                {
                    MessageBox.Show("Remplissez la zone : Délai Marché ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    TxtDélai.Focus();
                    return;
                }
                if (string.IsNullOrEmpty(CmbMO.Text))
                {
                    MessageBox.Show("Remplissez la zone : Maitre d'ouvrage ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    CmbMO.Focus();
                    return;
                }
                if (string.IsNullOrEmpty(CmbME.Text))
                {
                    MessageBox.Show("Remplissez la zone : Maitre d'oeuvre ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    CmbME.Focus();
                    return;
                }
                if (string.IsNullOrEmpty(TxtGarantie.Text))
                {
                    MessageBox.Show("Remplissez la zone : Gagantie ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    TxtGarantie.Focus();
                    return;
                }
                if (string.IsNullOrEmpty(TxtLot.Text))
                {
                    MessageBox.Show("Remplissez la zone : Lot ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    TxtLot.Focus();
                    return;
                }
                if (DTPDRP.Value == DateTime.MinValue)
                {
                    MessageBox.Show("Remplissez la zone : Date recep. prov. ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    DTPDRP.Focus();
                    return;
                }
                if (DTPRD.Value == DateTime.MinValue)
                {
                    MessageBox.Show("Remplissez la zone : Date recep. def. ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    DTPRD.Focus();
                    return;
                }
                if (DTPDA.Value == DateTime.MinValue)
                {
                    MessageBox.Show("Choisissez une date d'approbation : --- ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    DTPDA.Focus();
                    return;
                }
                if (DTPDOS.Value == DateTime.MinValue)
                {
                    MessageBox.Show("Choisissez une date de la zone : Date d'ordre de Comm. ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    DTPDOS.Focus();
                    return;
                }
                
                if (DTPDOP.Value == DateTime.MinValue)
                {
                    MessageBox.Show("Choisissez une date de la zone : Date ouverture de plis ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    DTPDOP.Focus();
                    return;
                }
                if (string.IsNullOrEmpty(TxtDG.Text))
                {
                    MessageBox.Show("Remplissez la zone : Délai de garantie ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    TxtDG.Focus();
                    return;
                }
                if (DTPEB.Value == DateTime.MinValue)
                {
                    MessageBox.Show("Choisissez une date de la zone : Epoque de base ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    DTPEB.Focus();
                    return;
                }
                if (string.IsNullOrEmpty(TxtCautionnement.Text))
                {
                    MessageBox.Show("Remplissez la zone : Montant Cautionnement ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    TxtCautionnement.Focus();
                    return;
                }
                if (string.IsNullOrEmpty(CmbStatut.Text))
                {
                    MessageBox.Show("Remplissez la zone : Statut ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    CmbStatut.Focus();
                    return;
                }     
                if (string.IsNullOrEmpty(TxtME.Text))
                {
                    MessageBox.Show("Remplissez la zone : Montant Emis ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    TxtME.Focus();
                    return;
                }

                // Préparez la requête d'insertion paramétrée
                if (string.IsNullOrEmpty(CmbMO.Text) || string.IsNullOrEmpty(CmbME.Text) || string.IsNullOrEmpty(CmbStatut.Text))
                {
                    MessageBox.Show("Veuillez sélectionner des valeurs valides pour les comboboxes.", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Préparez la requête d'insertion paramétrée
                using (OleDbCommand cmd = new OleDbCommand("SELECT N FROM Acteur WHERE Nom = ?", connection))
                {
                    cmd.Parameters.AddWithValue("@NomActeur", CmbMO.Text);
                    MaitreOuvrage = Convert.ToInt32(cmd.ExecuteScalar());
                }

                using (OleDbCommand cmd = new OleDbCommand("SELECT N FROM Entrepreneur WHERE Nom = ?", connection))
                {
                    cmd.Parameters.AddWithValue("@NomEntrepreneur", CmbME.Text);
                    MaitreOeuvre = Convert.ToInt32(cmd.ExecuteScalar());
                }
                string insertQuery = "INSERT INTO Marché (Réference, Objet, Montant_Global, Estimation, Delai, Maitre_ouvrage, Maitre_oeuvre, Statut, Garantie, Dat_reception_Pro, Date_reception_Def, Cautionnement, Date_Approbation, Date_Ouverture_Plis, Montant_Emis, Date_Ordre_service, Delai_Garantie, Epoque_Base, Lot) " +
                    "VALUES (@Reference, @Objet, @MontantGlobal, @Estimation, @Delai, @MaitreOuvrage, @MaitreOeuvre, @Statut, @Garantie, @DateReceptionPro, @DateReceptionDef, @Cautionnement, @DateApprobation, @DateOuverturePlis, @MontantEmis, @DateOrdreService, @DelaiGarantie, @EpoqueBase, @Lot)";

                using (OleDbCommand insertCmd = new OleDbCommand(insertQuery, connection))
                {
                    insertCmd.Parameters.AddWithValue("@Reference", TxtNumMarché.Text.Replace("'", "''"));
                    insertCmd.Parameters.AddWithValue("@Objet", TxtObjet.Text.Replace("'", "''"));
                    insertCmd.Parameters.AddWithValue("@MontantGlobal", TxtMontant.Text.Replace("'", "''"));
                    insertCmd.Parameters.AddWithValue("@Estimation", TxtEstimation.Text.Replace("'", "''"));
                    insertCmd.Parameters.AddWithValue("@Delai", TxtDG.Text.Replace("'", "''"));
                    insertCmd.Parameters.AddWithValue("@MaitreOuvrage", CmbMO.Text.Replace("'", "''"));
                    insertCmd.Parameters.AddWithValue("@MaitreOeuvre", MaitreOeuvre);
                    insertCmd.Parameters.AddWithValue("@Statut", Statut);
                    insertCmd.Parameters.AddWithValue("@Garantie", TxtGarantie.Text.Replace("'", "''"));
                    insertCmd.Parameters.AddWithValue("@DateReceptionPro", DTPDRP.Value.ToString("dd/MM/yyyy"));
                    insertCmd.Parameters.AddWithValue("@DateReceptionDef", DTPRD.Value.ToString("dd/MM/yyyy"));
                    insertCmd.Parameters.AddWithValue("@Cautionnement", TxtCautionnement.Text.Replace("'", "''"));
                    insertCmd.Parameters.AddWithValue("@DateApprobation", DTPDA.Value.ToString("dd/MM/yyyy"));
                    insertCmd.Parameters.AddWithValue("@DateOuverturePlis", DTPDOP.Value.ToString("dd/MM/yyyy"));
                    insertCmd.Parameters.AddWithValue("@MontantEmis", TxtME.Text.Replace("'", "''"));
                    insertCmd.Parameters.AddWithValue("@DateOrdreService", DTPDOS.Value.ToString("dd/MM/yyyy"));
                    insertCmd.Parameters.AddWithValue("@DelaiGarantie", TxtDG.Text.Replace("'", "''"));
                    insertCmd.Parameters.AddWithValue("@EpoqueBase", DTPEB.Value.ToString("dd/MM/yyyy"));
                    insertCmd.Parameters.AddWithValue("@Lot", TxtLot.Text.Replace("'", "''"));

                    insertCmd.ExecuteNonQuery();
                }

                recharger();

                // Réinitialisez vos champs après l'insertion
                TxtObjet.Text = "";
                TxtMontant.Text = "";
                TxtNumMarché.Text = "";
                TxtGarantie.Text = "";
                TxtCautionnement.Text = "";
                TxtME.Text = "";
                TxtDG.Text = "";
                TxtLot.Text = "";
                TxtEstimation.Text = "";
                TxtDélai.Text = "";
                CmbMO.Text = "";
                CmbME.Text = "";
                TxtGarantie.Text = "";
                CmbStatut.Text = "";
                DTPDRP.Value = DateTime.Now;
                DTPRD.Value = DateTime.Now;
                DTPDA.Value = DateTime.Now;
                DTPDOP.Value = DateTime.Now;
                DTPDOS.Value = DateTime.Now;
                DTPEB.Value = DateTime.Now;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private long GetMaitreOuvrageID(string nomMaitreOuvrage)
        {
            long acteurID = -1; // Valeur par défaut si l'acteur n'est pas trouvé

            using (OleDbCommand cmd = new OleDbCommand("SELECT N FROM Acteur WHERE Nom = ?", connection))
            {
                cmd.Parameters.AddWithValue("@NomActeur", nomMaitreOuvrage);
                object result = cmd.ExecuteScalar();
                if (result != null)
                {
                    acteurID = Convert.ToInt64(result);
                }
            }

            return acteurID;
        }

        private void BtnModifier_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(CmbNumMarchéM.Text))
                {
                    MessageBox.Show("Remplissez la zone : Référence..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    CmbNumMarchéM.Focus();
                    return;
                }

                if (string.IsNullOrEmpty(TxtObjetM.Text))
                {
                    MessageBox.Show("Remplissez la zone Objet ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    TxtObjetM.Focus();
                    return;
                }

                if (string.IsNullOrEmpty(TxtMontantM.Text))
                {
                    MessageBox.Show("Remplissez la zone Montant Marché ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    TxtMontantM.Focus();
                    return;
                }

                if (string.IsNullOrEmpty(TxtEstimationM.Text))
                {
                    MessageBox.Show("Remplissez la zone Estimation ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    TxtEstimationM.Focus();
                    return;
                }

                if (string.IsNullOrEmpty(TxtDélaiM.Text))
                {
                    MessageBox.Show("Remplissez la zone Délai ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    TxtDélaiM.Focus();
                    return;
                }

                if (string.IsNullOrEmpty(TxtGarantieM.Text))
                {
                    MessageBox.Show("Remplissez la zone Garantie Marché ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    TxtGarantieM.Focus();
                    return;
                }

                if (string.IsNullOrEmpty(CmbMEM.Text))
                {
                    MessageBox.Show("Remplissez la zone Maitre d'Oeuvre ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    CmbMEM.Focus();
                    return;
                }

                if (string.IsNullOrEmpty(CmbMOM.Text))
                {
                    MessageBox.Show("Remplissez la zone Maitre d'ouvrage ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    CmbMOM.Focus();
                    return;
                }

                if (DTPDRPM.Value == DateTime.MinValue)
                {
                    MessageBox.Show("Remplissez la zone Reception provisoire ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    DTPDRPM.Focus();
                    return;
                }

                if (DTPDRDM.Value == DateTime.MinValue)
                {
                    MessageBox.Show("Remplissez la zone Reception définitive ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    DTPDRDM.Focus();
                    return;
                }

                if (DTPDAM.Value == DateTime.MinValue)
                {
                    MessageBox.Show("Choisissez une date de la zone Date d'Approbation ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    DTPDAM.Focus();
                    return;
                }

                if (DTPDOPM.Value == DateTime.MinValue)
                {
                    MessageBox.Show("Choisissez une date de la zone Date ouverture plis ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    DTPDOPM.Focus();
                    return;
                }

                if (DTPDOSM.Value == DateTime.MinValue)
                {
                    MessageBox.Show("Choisissez une date de la zone Date d'ordre commencement ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    DTPDOSM.Focus();
                    return;
                }

                if (DTPEBM.Value == DateTime.MinValue)
                {
                    MessageBox.Show("Choisissez une date de la zone Epoque de base ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    DTPEBM.Focus();
                    return;
                }

                if (string.IsNullOrEmpty(TxtDGM.Text))
                {
                    MessageBox.Show("Remplissez la zone Délai de garantie ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    TxtDGM.Focus();
                    return;
                }

                if (string.IsNullOrEmpty(TxtCautionnementM.Text))
                {
                    MessageBox.Show("Remplissez la zone Cautionnement ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    TxtCautionnementM.Focus();
                    return;
                }

                if (string.IsNullOrEmpty(CmbStatutM.Text))
                {
                    MessageBox.Show("Remplissez la zone Statut du Marché ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    CmbStatutM.Focus();
                    return;
                }

                if (string.IsNullOrEmpty(TxtMEM.Text))
                {
                    MessageBox.Show("Remplissez la zone Montant Emis ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    TxtMEM.Focus();
                    return;
                }

                // -------------- Ajouter --------------
                string SQLSTR = "UPDATE Marché SET ";
                SQLSTR += "Objet = '" + TxtObjetM.Text.Replace("'", "''") + "', ";
                SQLSTR += "Montant_Global = " + TxtMontantM.Text.Replace(",", ".") + ", ";
                SQLSTR += "Estimation = " + TxtEstimationM.Text.Replace(",", ".") + ", ";
                SQLSTR += "Delai = '" + TxtDélaiM.Text.Replace("'", "''") + "', ";
                SQLSTR += "Observation = '" + CmbStatutM.Text.Replace("'", "''") + "', ";

                using (OleDbCommand cmd = new OleDbCommand("SELECT N FROM Acteur WHERE Nom = ?", connection))
                {
                    cmd.Parameters.AddWithValue("@NomActeur", CmbMOM.Text);
                    int NumActeurMO = Convert.ToInt32(cmd.ExecuteScalar());
                    SQLSTR += "Maitre_ouvrage = " + NumActeurMO + ", ";
                }

                using (OleDbCommand cmd = new OleDbCommand("SELECT N FROM Entrepreneur WHERE Nom = ?", connection))
                {
                    cmd.Parameters.AddWithValue("@NomEntrepreneur", CmbMEM.Text);
                    int NumActeurME = Convert.ToInt32(cmd.ExecuteScalar());
                    SQLSTR += "Maitre_oeuvre = " + NumActeurME + ", ";
                }

                using (OleDbCommand cmd = new OleDbCommand("SELECT N From StatutMarché where Nom = ?", connection))
                {
                    cmd.Parameters.AddWithValue("@NomStatut", CmbStatutM.Text);
                    int Statut = Convert.ToInt32(cmd.ExecuteScalar());
                    SQLSTR += "Statut = " + Statut + ", ";
                }

                SQLSTR += "Garantie = '" + TxtGarantieM.Text.Replace("'", "''") + "', ";
                SQLSTR += "Dat_reception_Pro = #" + DTPDRPM.Value.ToString("yyyy-MM-dd") + "#, ";
                SQLSTR += "Date_reception_Def = #" + DTPDRDM.Value.ToString("yyyy-MM-dd") + "#, ";
                SQLSTR += "Cautionnement = '" + TxtCautionnementM.Text.Replace("'", "''") + "', ";
                SQLSTR += "Date_Approbation = #" + DTPDAM.Value.ToString("yyyy-MM-dd") + "#, ";
                SQLSTR += "Date_Ouverture_Plis = #" + DTPDOPM.Value.ToString("yyyy-MM-dd") + "#, ";
                SQLSTR += "Montant_Emis = " + TxtMEM.Text.Replace(",", ".") + ", ";
                SQLSTR += "Date_Ordre_service = #" + DTPDOSM.Value.ToString("yyyy-MM-dd") + "#, ";
                SQLSTR += "Delai_Garantie = " + TxtDGM.Text.Replace(",", ".") + ", ";
                SQLSTR += "Epoque_Base = #" + DTPEBM.Value.ToString("yyyy-MM-dd") + "#, ";
                SQLSTR += "Lot = " + TxtLotM.Text.Replace("'", "''");
                SQLSTR += " WHERE Réference = '" + CmbNumMarchéM.Text + "'";

                // Exécutez la requête SQL
                ExecuterSQL(SQLSTR);
                recharger();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void CmbNumMarchéM_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataAdap = new OleDbDataAdapter("Select * From Marchér where Réference = '" + CmbNumMarchéM.Text + "'", connection);
            DataTab = new DataTable();
            DataAdap.Fill(DataTab);

            if (DataTab.Rows.Count > 0)
            {
                //TxtNumMarché.Text = DataTab.Rows[0][0].ToString();
                TxtObjetM.Text = DataTab.Rows[0][1].ToString();
                TxtMontantM.Text = DataTab.Rows[0][2].ToString();
                TxtEstimationM.Text = DataTab.Rows[0][3].ToString();
                TxtDélaiM.Text = DataTab.Rows[0][4].ToString();
                CmbMOM.Text = DataTab.Rows[0][5].ToString();
                CmbMEM.Text = DataTab.Rows[0][6].ToString();
                CmbStatutM.Text = DataTab.Rows[0][7].ToString();
                TxtGarantieM.Text = DataTab.Rows[0][8].ToString();
                DTPDRPM.Value = Convert.ToDateTime(DataTab.Rows[0][9]);
                DTPDRDM.Value = Convert.ToDateTime(DataTab.Rows[0][10]);
                TxtCautionnementM.Text = DataTab.Rows[0][11].ToString();
                DTPDAM.Value = Convert.ToDateTime(DataTab.Rows[0][12]);
                DTPDOPM.Value = Convert.ToDateTime(DataTab.Rows[0][13]);
                TxtMEM.Text = DataTab.Rows[0][14].ToString();
                DTPDOSM.Value = Convert.ToDateTime(DataTab.Rows[0][15]);
                TxtDGM.Text = DataTab.Rows[0][16].ToString();
                DTPEBM.Value = Convert.ToDateTime(DataTab.Rows[0][17]);
                TxtLotM.Text = DataTab.Rows[0][18].ToString();
            }
        }
        private void Marché_Resize(object sender, EventArgs e)
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
