using Microsoft.Office.Core;
using System;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Révision
{
    public partial class FrmEditAttachement : Form
    {
        OleDbConnection connection = new OleDbConnection();
        OleDbCommand records;
        DataTable DataTab;
        OleDbDataAdapter DataAdap;
        OleDbCommandBuilder ComndBuld;
        DataSet DataSetTab = new DataSet();

        OleDbConnection connections = new OleDbConnection();
        OleDbCommand recordss;
        DataTable DataTabs;
        OleDbDataAdapter DataAdaps;
        OleDbCommandBuilder ComndBulds;
        DataSet DataSetTabs = new DataSet();

        int INDEXS;
        string SQLSTR;
        long NumDécompte, NumDécompteM, NumDécompteMA;
        long NumEntrepreneur;


        public FrmEditAttachement()
        {
            InitializeComponent();
        }

        private void CmbDécompteM_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(CmbDécompteM.Text))
            {
                MessageBox.Show("Choisissez une valeur de la liste : Référence de Marché ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                CmbDécompteM.Focus();
                return;
            }

            // Utilisation de paramètres dans la requête SQL pour éviter les attaques par injection SQL
            using (OleDbCommand command = new OleDbCommand("Select N° From Décompte where Désignation=@Désignation and Num_Marché=@Num_Marché", connection))
            {
                command.Parameters.AddWithValue("@Désignation", CmbDécompteM.Text);
                command.Parameters.AddWithValue("@Num_Marché", CmbNumMarchéM.Text);

                DataAdap = new OleDbDataAdapter(command);
                DataTab = new DataTable();
                DataAdap.Fill(DataTab);

                if (DataTab.Rows.Count > 0)
                {
                    NumDécompteM = Convert.ToInt64(DataTab.Rows[0]["N°"]);

                    using (OleDbCommand ordreServiceCommand = new OleDbCommand("Select Date_Ordre_Service, Date_Arret_Service From Ordre_Service where Num_Décompte=@Num_Décompte", connection))
                    {
                        // Supposons que NumDécompteM est de type int.
                        ordreServiceCommand.Parameters.Add("@Num_Décompte", OleDbType.Integer).Value = NumDécompteM;


                        DataAdap = new OleDbDataAdapter(ordreServiceCommand);
                        DataTab = new DataTable();
                        DataAdap.Fill(DataTab);

                        if (DataTab.Rows.Count == 0)
                        {
                            MessageBox.Show($"Remplir les dates d'arrêt et de reprise pour le décompte : {CmbDécompteM.Text}.{Environment.NewLine}Décompte > Ajouter", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            FrmOrdreService frmOrdreServiceForm = new FrmOrdreService();
                            frmOrdreServiceForm.Show();
                            CmbDécompteM.Text = string.Empty;
                            return;
                        }

                        DGV.DataSource = DataTab;
                        INDEXS = CmbDécompteM.SelectedIndex;

                        using (OleDbCommand attachementCommand = new OleDbCommand("Select Désignation,Unité,Prix_unitaire,Quantité from Attachement where Num_Décompte=@Num_Décompte", connection))
                        {
                            attachementCommand.Parameters.Add("@Num_Décompte", OleDbType.Integer).Value = NumDécompteM;

                            DataAdap = new OleDbDataAdapter(attachementCommand);
                            DataTab = new DataTable();
                            DataAdap.Fill(DataTab);

                            if (DataTab.Rows.Count == 0)
                            {
                                MessageBox.Show("Remplir l'attachement.", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                Attachement attachementForm = new Attachement();
                                attachementForm.Show();
                                CmbDécompteM.Text = string.Empty;
                                return;
                            }

                            DGVP.DataSource = DataTab;
                        }
                    }
                }
            }
        }

        private void CmbNumMarchéM_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(CmbNumMarchéM.Text))
            {
                MessageBox.Show("Choisissez une valeur de la liste : Référence de Marché ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                CmbNumMarchéM.Focus();
                return;
            }

            // Charger les désignations de décompte
            using (OleDbCommand command = new OleDbCommand("Select Désignation From Décompte where Num_Marché=@Num_Marché", connection))
            {
                command.Parameters.AddWithValue("@Num_Marché", CmbNumMarchéM.Text);

                DataAdap = new OleDbDataAdapter(command);
                DataTab = new DataTable();
                DataAdap.Fill(DataTab);

                if (DataTab.Rows.Count > 0)
                {
                    for (int i = 0; i < DataTab.Rows.Count; i++)
                    {
                        CmbDécompteM.Items.Add(DataTab.Rows[i]["Désignation"].ToString());
                    }
                }
                else
                {
                    MessageBox.Show("Aucune valeur n'a été saisie dans la table détail estimatif pour le marché numéro : " + CmbNumMarchéM.Text + ".", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
            }

            // Charger les informations du marché
            using (OleDbCommand marchéCommand = new OleDbCommand("Select * From Marché where Réference=@Réference", connection))
            {
                marchéCommand.Parameters.AddWithValue("@Réference", CmbNumMarchéM.Text);

                DataAdap = new OleDbDataAdapter(marchéCommand);
                DataTab = new DataTable();
                DataAdap.Fill(DataTab);

                if (DataTab.Rows.Count > 0)
                {
                    TxtObjet.Text = DataTab.Rows[0]["Objet"].ToString();
                    TxtLot.Text = DataTab.Rows[0]["Lot"].ToString();
                    NumEntrepreneur = Convert.ToInt64(DataTab.Rows[0]["Maitre_oeuvre"]);
                }
                else
                {
                    MessageBox.Show("Aucune valeur n'a été saisie dans la table détail estimatif pour le marché numéro : " + CmbNumMarchéM.Text + ".", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
            }

            // Charger les informations de l'entrepreneur
            using (OleDbCommand entrepreneurCommand = new OleDbCommand("SELECT * FROM Entrepreneur WHERE N° = ?", connection))
            {
                entrepreneurCommand.Parameters.Add("N°", OleDbType.Integer).Value = NumEntrepreneur;

                DataAdap = new OleDbDataAdapter(entrepreneurCommand);
                DataTab = new DataTable();
                DataAdap.Fill(DataTab);

                if (DataTab.Rows.Count > 0)
                {
                    TxtNom.Text = DataTab.Rows[0]["Nom"].ToString();
                    TxtAdresse.Text = DataTab.Rows[0]["Adresse"].ToString();
                    TxtCNSS.Text = DataTab.Rows[0]["CNSS"].ToString();
                    TxtPatente.Text = DataTab.Rows[0]["Patente"].ToString();
                    TxtNCB.Text = DataTab.Rows[0]["Compte_Bancaire_Num"].ToString();
                    TxtAB.Text = DataTab.Rows[0]["Agence_Bancaire"].ToString();
                    TxtRC.Text = DataTab.Rows[0]["Registre_Commerce"].ToString();
                    TxtTel.Text = DataTab.Rows[0]["Téléphone"].ToString();
                    TxtEMail.Text = DataTab.Rows[0]["EMail"].ToString();
                    TxtIF.Text = DataTab.Rows[0]["Identifiant_Fiscal"].ToString();
                    TxtICE.Text = DataTab.Rows[0]["Identifiant_Commun_Entreprise"].ToString();
                }
                else
                {
                    MessageBox.Show("Aucune valeur n'a été saisie dans la table détail estimatif pour le marché numéro : " + CmbNumMarchéM.Text + ".", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }
        private void FrmEditAttachement_FormClosed(object sender, FormClosedEventArgs e)
        {
            try
            {
                if (connection.State == ConnectionState.Open)
                {
                    connection.Close();
                }
            }
            catch (Exception ex)
            {
                // Gérer ou journaliser l'exception, si nécessaire.
                Console.WriteLine("Erreur lors de la fermeture de la connexion : " + ex.Message);
            }
            finally
            {
                // Assurez-vous que la connexion est fermée même en cas d'exception
                if (connection.State != ConnectionState.Closed)
                {
                    connection.Close();
                }
            }
        }

        private void FrmEditAttachement_Load(object sender, EventArgs e)
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
                // Gérer ou journaliser l'exception, si nécessaire.
                Console.WriteLine("Erreur lors de l'ouverture de la connexion : " + ex.Message);
            }
            DGV.EnableHeadersVisualStyles = false;
            DGV.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            DGV.ColumnHeadersDefaultCellStyle.BackColor = Color.DarkGray;

            DGVP.EnableHeadersVisualStyles = false;
            DGVP.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            DGVP.ColumnHeadersDefaultCellStyle.BackColor = Color.DarkGray;
        }

        private void Recharger()
        {
            DataAdap = new OleDbDataAdapter("Select Réference From Marché", connection);
            DataTab = new DataTable();
            DataAdap.Fill(DataTab);

            for (int i = 0; i < DataTab.Rows.Count; i++)
            {
                CmbNumMarchéM.Items.Add(DataTab.Rows[i]["Réference"].ToString());
            }
        }


        private void ExecuterSQL(string Strsql)
        {
            try
            {
                using (OleDbCommand cmd = connection.CreateCommand())
                {
                    cmd.CommandText = Strsql;
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                // Gérer ou journaliser l'exception, si nécessaire.
                Console.WriteLine("Erreur lors de l'exécution de la requête SQL : " + ex.Message);
            }
        }
        private void BtnEditer_Click(object sender, EventArgs e)
        {
            //ExcelApp.DecimalSeparator = ".";
            // Déclaration des variables
            Excel.Application ExcelApp;
            Excel.Workbook ExcelBook;
            Excel.Worksheet ExcelSheet;
            int i, j, l;
            string str;
            double Total, Accompte, MontantDP, MontantMarché, RetenueGarantie, MontantAvance;
            string TDécompte = "";
            int[] lh = new int[6];

            // Vérification des champs
            if (string.IsNullOrEmpty(CmbNumMarchéM.Text))
            {
                MessageBox.Show("Choisissez une valeur de la liste : Référence de Marché ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                CmbNumMarchéM.Focus();
                return;
            }

            if (string.IsNullOrEmpty(CmbDécompteM.Text))
            {
                MessageBox.Show("Choisissez une valeur de la liste : Référence de Décompte ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                CmbDécompteM.Focus();
                return;
            }

            // Création d'un objet Excel
            ExcelApp = new Excel.Application();
            ExcelBook = ExcelApp.Workbooks.Add();
            ExcelSheet = (Excel.Worksheet)ExcelBook.Worksheets[1];

            // Feuille de calcul Attachement
            ExcelSheet = (Excel.Worksheet)ExcelBook.Worksheets.Add();
            ExcelSheet = (Excel.Worksheet)ExcelBook.Worksheets[2];
            ExcelSheet.Name = "Décompte";

            ExcelSheet = (Excel.Worksheet)ExcelBook.Worksheets.Add();
            ExcelSheet = (Excel.Worksheet)ExcelBook.Worksheets[3];
            ExcelSheet.Name = "Récapitulatif";


            // ======================================="Ajout des informations dans la feuille "Attachement"=======================
            ExcelSheet = (Excel.Worksheet)ExcelBook.Worksheets[1];
            ExcelSheet.Name = "Attachement";

            ExcelSheet.Cells[1, 1] = "Royaume du Maroc";
            ExcelSheet.Rows["1:1"].RowHeight = 11;

            ExcelSheet.Cells[2, 1] = "Ministère de l'Intérieur";
            ExcelSheet.Rows["2:2"].RowHeight = 11;

            ExcelSheet.Cells[3, 1] = "Province d'Al Haouz";
            ExcelSheet.Rows["3:3"].RowHeight = 11;

            ExcelSheet.Cells[4, 1] = "Conseil Provincial Al Haouz";
            ExcelSheet.Rows["4:4"].RowHeight = 11;

            for (j = 1; j <= 4; j++)
            {
                ExcelSheet.Cells[j, 1].Font.Bold = true;
                ExcelSheet.Cells[j, 1].Font.Size = 10;
                ExcelSheet.Cells[j, 1].Font.Name = "Times New Roman";
                ExcelSheet.Cells[j, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                ExcelSheet.Cells[j, 1].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            }

            for (j = 6; j <= 9; j++)
            {
                ExcelSheet.Cells[j, 1].Font.Bold = true;
                ExcelSheet.Cells[j, 1].Font.Size = 12;
                ExcelSheet.Cells[j, 1].Font.Name = "Times New Roman";
                ExcelSheet.Cells[j, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                ExcelSheet.Cells[j, 1].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                ExcelSheet.Cells[j, 1].NumberFormat = "@";
            }

            ExcelSheet.Cells[6, 1] = " Objet : " + TxtObjet.Text;
            ExcelSheet.Rows["6:6"].RowHeight = 14;
            ExcelSheet.Range["A6:E6"].Merge();
            ExcelSheet.Cells[6, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            ExcelSheet.Cells[7, 1] = " Numéro de Marché : " + CmbNumMarchéM.Text;
            ExcelSheet.Rows["7:7"].RowHeight = 12;
            ExcelSheet.Range["A7:E7"].Merge();
            ExcelSheet.Cells[7, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            ExcelSheet.Cells[8, 1] = " Entreprise : " + TxtRC.Text;
            ExcelSheet.Rows["8:8"].RowHeight = 12;

            ExcelSheet.Cells[9, 1] = " Adresse : " + TxtAdresse.Text;
            ExcelSheet.Rows["9:9"].RowHeight = 12;

            Excel.Range attachementRange = ExcelSheet.Range["A11:E11"];
            attachementRange.Merge();

            ExcelSheet.Cells[11, 1] = "Attachement Provisoire N° :" + CmbNumMarchéM.Text.Substring(2, CmbNumMarchéM.Text.Length - 2);
            attachementRange.Interior.Color = Color.SkyBlue;
            attachementRange.Font.Bold = true;
            attachementRange.Font.Size = 13;
            attachementRange.Font.Name = "Times New Roman";
            attachementRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            attachementRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
            ExcelSheet.Rows["11:11"].RowHeight = 20;

            l = 13; // Assurez-vous que l est correctement initialisé avant cette partie du code.

            OleDbDataAdapter dataAdapter = new OleDbDataAdapter("Select Montant_Global, Montant_Avance, Cautionnement From Marché where Réference='" + CmbNumMarchéM.Text + "'", connection);
            DataTable dataTable = new DataTable();
            dataAdapter.Fill(dataTable);

            MontantMarché = Convert.ToDouble(dataTable.Rows[0]["Montant_Global"]);
            MontantAvance = Convert.ToDouble(dataTable.Rows[0]["Montant_Avance"]);
            string Cautionnement = dataTable.Rows[0]["Cautionnement"].ToString();

            dataAdapter = new OleDbDataAdapter("Select Num_Prix, Désignation, Unité, Quantité From Detail_Estimatif where Num_Marché='" + CmbNumMarchéM.Text + "'", connection);
            dataTable = new DataTable();
            dataAdapter.Fill(dataTable);

            if (dataTable.Rows.Count > 0)
            {
                ExcelSheet.Cells[13, 1] = "Prix N°";
                ExcelSheet.Cells[13, 2] = "Désignation";
                ExcelSheet.Cells[13, 3] = "Unité";
                ExcelSheet.Cells[13, 4] = "Quantité Marché";
                ExcelSheet.Cells[13, 5] = "Quantité Réalisée";

                for (i = 1; i <= dataTable.Columns.Count + 1; i++)
                {
                    ExcelSheet.Cells[l, i].Interior.Color = Color.Gold;
                    ExcelSheet.Cells[l, i].Borders.Color = Color.Black;
                    ExcelSheet.Cells[l, i].Font.Size = 12;
                    ExcelSheet.Cells[l, i].Font.Bold = true;
                    ExcelSheet.Cells[l, i].Font.Name = "Times New Roman";
                    ExcelSheet.Cells[l, i].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    ExcelSheet.Cells[l, i].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                    if (lh[i - 1] < (int)(ExcelSheet.Cells[l, i].Value?.ToString().Length * 1.3))
                        lh[i - 1] = (int)(ExcelSheet.Cells[l, i].Value?.ToString().Length * 1.3);

                    ExcelSheet.Columns[i].ColumnWidth = lh[i - 1];

                    for (j = 1; j <= dataTable.Rows.Count; j++)
                    {
                        ExcelSheet.Cells[j + l, i].Borders.Color = Color.Black;

                        if (i == dataTable.Columns.Count + 1)
                        {
                            ExcelSheet.Cells[j + l, i] = DGVP.Rows[j - 1].Cells[3].Value;
                            ExcelSheet.Cells[j + l, i].NumberFormat = "0,00";
                        }
                        else
                        {
                            if (i == dataTable.Columns.Count)
                                ExcelSheet.Cells[j + l, i].NumberFormat = "0,00";
                            else
                                ExcelSheet.Cells[j + l, i].NumberFormat = "@";

                            ExcelSheet.Cells[j + l, i] = dataTable.Rows[j - 1][i - 1]?.ToString();
                        }

                        if (lh[i - 1] < ExcelSheet.Cells[j + l, i].Value?.ToString().Length * 1.3)
                            lh[i - 1] = (int)(ExcelSheet.Cells[j + l, i].Value?.ToString().Length * 1.3);

                        ExcelSheet.Columns[i].ColumnWidth = lh[i - 1];

                        ExcelSheet.Cells[j + l, i].Font.Size = 12;
                        ExcelSheet.Cells[j + l, i].Font.Name = "Times New Roman";
                        ExcelSheet.Cells[j + l, i].HorizontalAlignment = XlHAlign.xlHAlignRight;
                        ExcelSheet.Cells[j + l, i].VerticalAlignment = XlVAlign.xlVAlignCenter;
                    }
                }
            }

            // ==========================================Feuille de calcul Décompte===========================================
            DataAdap = new OleDbDataAdapter("Select Num_Prix,Désignation,Unité,Prix,Quantité From Detail_Estimatif where Num_Marché='" + CmbNumMarchéM.Text + "'", connection);
            DataTab = new DataTable();
            DataAdap.Fill(DataTab);

            if (connections.State == ConnectionState.Closed)
            {
                connections.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data source=" + Application.StartupPath + "\\revision.accdb";
                connections.Open();
            }

            ExcelSheet = (Excel.Worksheet)ExcelBook.Worksheets[2];
            ExcelSheet.Name = "Décompte";

            if (ExcelSheet != null)
            {
                ExcelSheet.Cells[1, 1] = "Royaume du Maroc";
                ExcelSheet.Rows["1:1"].RowHeight = 12;
                ExcelSheet.Cells[2, 1] = "Ministère de l'Intérieur";
                ExcelSheet.Rows["2:2"].RowHeight = 12;
                ExcelSheet.Cells[3, 1] = "Province d'Al Haouz";
                ExcelSheet.Rows["3:3"].RowHeight = 12;
                ExcelSheet.Cells[4, 1] = "Conseil Provincial Al Haouz";
                ExcelSheet.Rows["4:4"].RowHeight = 12;

                for (j = 1; j <= 4; j++)
                {
                    ExcelSheet.Cells[j, 1].Font.Bold = false;
                    ExcelSheet.Cells[j, 1].Font.Size = 10;
                    ExcelSheet.Cells[j, 1].Font.Name = "Times New Roman";
                    ExcelSheet.Cells[j, 1].HorizontalAlignment = 1;
                    ExcelSheet.Cells[j, 1].VerticalAlignment = XlVAlign.xlVAlignCenter;
                }

                for (j = 6; j <= 17; j++)
                {
                    ExcelSheet.Cells[j, 1].Font.Bold = true;
                    ExcelSheet.Cells[j, 1].Font.Size = 12;
                    ExcelSheet.Cells[j, 1].Font.Name = "Times New Roman";
                    ExcelSheet.Cells[j, 1].HorizontalAlignment = 1;
                    ExcelSheet.Cells[j, 1].VerticalAlignment = XlVAlign.xlVAlignCenter;
                    ExcelSheet.Cells[j, 1].NumberFormat = "@";
                }

                ExcelSheet.Cells[6, 1] = " Objet : " + TxtObjet.Text;
                ExcelSheet.Rows["6:6"].RowHeight = 12;
                ExcelSheet.Cells[7, 1] = " Numéro de Marché : " + CmbNumMarchéM.Text;
                ExcelSheet.Rows["7:7"].RowHeight = 12;
                ExcelSheet.Cells[8, 1] = " Entreprise : " + TxtRC.Text;
                ExcelSheet.Rows["8:8"].RowHeight = 20;
                ExcelSheet.Cells[9, 1] = " Adresse : " + TxtAdresse.Text;
                ExcelSheet.Rows["9:9"].RowHeight = 12;
                ExcelSheet.Cells[10, 1] = " Patente : " + TxtPatente.Text;
                ExcelSheet.Rows["10:10"].RowHeight = 12;
                ExcelSheet.Cells[11, 1] = " CNSS : " + TxtCNSS.Text;
                ExcelSheet.Rows["11:11"].RowHeight = 12;
                ExcelSheet.Cells[12, 1] = " Identifiant Fiscal : " + TxtIF.Text;
                ExcelSheet.Rows["12:12"].RowHeight = 12;
                ExcelSheet.Cells[13, 1] = " Compte Bancaire : " + TxtAdresse.Text;
                ExcelSheet.Rows["13:13"].RowHeight = 12;
                ExcelSheet.Cells[14, 1] = " Agence Bancaire : " + TxtAB.Text;
                ExcelSheet.Rows["14:14"].RowHeight = 12;
                ExcelSheet.Cells[15, 1] = " Identifiant Commun Entreprise : " + TxtICE.Text;
                ExcelSheet.Rows["15:15"].RowHeight = 12;
                ExcelSheet.Cells[16, 1] = " Registre de Commerce : " + TxtRC.Text;
                ExcelSheet.Rows["16:16"].RowHeight = 12;
                ExcelSheet.Cells[17, 1] = " Tél : " + TxtTel.Text;
                ExcelSheet.Rows["17:17"].RowHeight = 12;
                ExcelSheet.Rows["18:18"].RowHeight = 20;

                ExcelSheet.Cells[18, 4] = "Montant de l'Acompte TTC (dhs) : ";
                ExcelSheet.Cells[18, 5].Font.Bold = true;
                ExcelSheet.Cells[18, 5].Font.Italic = true;
                ExcelSheet.Cells[18, 5].Font.Size = 12;
                ExcelSheet.Cells[18, 5].Font.Name = "Times New Roman";
                ExcelSheet.Cells[18, 5].HorizontalAlignment = 3;
                ExcelSheet.Cells[18, 5].VerticalAlignment = XlVAlign.xlVAlignCenter;
                ExcelSheet.Cells[18, 5].NumberFormat = "@";
                ExcelSheet.Cells[18, 4].Font.Bold = true;
                ExcelSheet.Cells[18, 4].Font.Size = 12;
                ExcelSheet.Cells[18, 4].Font.Name = "Times New Roman";
                ExcelSheet.Cells[18, 4].HorizontalAlignment = 4;
                ExcelSheet.Cells[18, 4].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                ExcelSheet.Cells[18, 4].NumberFormat = "@";
                ExcelSheet.Range["A20:E20"].Merge();

                ExcelSheet.Cells[20, 1] = "Décompte Provisoire N°:" + CmbDécompteM.Text.Substring(2, CmbDécompteM.Text.Length - 2);
                ExcelSheet.Cells[20, 1].Interior.Color = Color.SkyBlue;
                ExcelSheet.Cells[20, 1].Font.Bold = true;
                ExcelSheet.Cells[20, 1].Font.Size = 13;
                ExcelSheet.Cells[20, 1].Font.Name = "Times New Roman";
                ExcelSheet.Cells[20, 1].HorizontalAlignment = 3;
                ExcelSheet.Cells[20, 1].VerticalAlignment = XlVAlign.xlVAlignCenter;
                ExcelSheet.Rows["20:20"].RowHeight = 25;
            }

            l = 22;
            lh = new int[6];
            double MREV = 0;
            int NDEC = 0;
            int NDECT = 0;

            NDEC = int.Parse(CmbDécompteM.Text.Substring(2, CmbDécompteM.Text.Length - 2));
            DataAdap = new OleDbDataAdapter("Select Désignation From Décompte where Num_Marché='" + CmbNumMarchéM.Text + "'", connection);
            DataTab = new DataTable();
            DataAdap.Fill(DataTab);

            OleDbDataAdapter DataAdaps;
            DataTable DataTabs;

            for (i = 0; i < DataTab.Rows.Count; i++)
            {
                string numeroPrix = DataTab.Rows[i].ItemArray[0].ToString();
                if (numeroPrix.Length > 2)
                {
                    // Utilisez Substring uniquement si la longueur est suffisante
                    NDECT = int.Parse(numeroPrix.Substring(2, numeroPrix.Length - 2));
                }
                else
                {
                    // Traitez le cas où la longueur est insuffisante (éventuellement lancez une exception ou affectez une valeur par défaut)
                    // Exemple :
                    NDECT = 0; // Valeur par défaut, veuillez ajuster selon vos besoins
                }
                if (NDECT <= NDEC)
                {
                    DataAdaps = new OleDbDataAdapter("Select Montant_Révisé From Décompte where Num_Marché='" + CmbNumMarchéM.Text + "' and Désignation='DP" + NDECT.ToString() + "'", connections);
                    DataTabs = new DataTable();
                    DataAdaps.Fill(DataTabs);
                    // Vérifiez si la table contient des lignes avant d'essayer d'accéder à la position 0
                    if (DataTabs.Rows.Count > 0)
                    {
                        MREV += Convert.ToDouble(DataTabs.Rows[0].ItemArray[0]);
                    }
                    else
                    {
                        // Gérez le cas où aucune ligne n'est renvoyée (peut-être affectez une valeur par défaut)
                        // Exemple :
                        MREV += 0; // Affectez une valeur par défaut
                    }
                }
            }

            DataAdap = new OleDbDataAdapter("Select Désignation,Unité,Prix_unitaire,Quantité From Attachement where Num_Décompte=" + NumDécompteM, connection);
            DataTab = new DataTable();
            DataAdap.Fill(DataTab);

            if (INDEXS > 0)
            {
                DataAdaps = new OleDbDataAdapter("SELECT N°, Montant, Tranche_Avance, Type FROM Décompte WHERE Num_Marché='" + CmbNumMarchéM.Text + "' AND Désignation='" + CmbDécompteM.Items[INDEXS - 1].ToString() + "'", connections);
                DataTabs = new DataTable();
                DataAdaps.Fill(DataTabs);
                NumDécompteM = Convert.ToInt64(DataTabs.Rows[0][0]);
                MontantDP = Convert.ToDouble(DataTabs.Rows[0][1]);
                MontantAvance = Convert.ToDouble(DataTabs.Rows[0][2]);
                TDécompte = DataTabs.Rows[0][3].ToString();
                DataAdaps = new OleDbDataAdapter("SELECT Quantité FROM Attachement WHERE Num_Décompte=" + NumDécompteM, connections);
                DataTabs = new DataTable();
                DataAdaps.Fill(DataTabs);
            }

            if (DataTab.Rows.Count > 0)
            {
                ExcelSheet.Cells[22, 1] = "Désignation";
                ExcelSheet.Cells[22, 2] = "Unité";
                ExcelSheet.Cells[22, 3] = "Prix Unitaire";
                ExcelSheet.Cells[22, 4] = "Quantité";
                ExcelSheet.Cells[22, 5] = "Total Calculé";

                for (i = 1; i <= DataTab.Columns.Count + 1; i++)
                {
                    Excel.Range currentCell = (Excel.Range)ExcelSheet.Cells[l, i];

                    currentCell.Interior.Color = Color.Gold;
                    currentCell.Borders.Color = Color.Black;
                    currentCell.Font.Size = 12;
                    currentCell.Font.Bold = true;
                    currentCell.Font.Name = "Times New Roman";
                    currentCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    currentCell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                    int columnIndex = i - 1;

                    if (lh[columnIndex] < currentCell.Text.ToString().Length * 1.3)
                    {
                        lh[columnIndex] = Convert.ToInt32(currentCell.Text.ToString().Length * 1.3);
                    }

                    ExcelSheet.Columns[i].ColumnWidth = lh[columnIndex];

                    for (j = 1; j <= DataTab.Rows.Count; j++)
                    {
                        Excel.Range dataCell = (Excel.Range)ExcelSheet.Cells[j + l, i];
                        dataCell.Borders.Color = Color.Black;
                        object cellValue = dataCell.Value;
                        switch (i)
                        {
                            case 3:
                                dataCell.Value = DataTab.Rows[j - 1][i - 1];
                                dataCell.NumberFormat = "0.00";
                                break;
                            case 4:
                                dataCell.Value = Math.Round(Convert.ToDouble(DataTab.Rows[j - 1][i - 1]), 2);
                                dataCell.NumberFormat = "0.00";
                                break;
                            case 5:
                                cellValue = Convert.ToDouble(ExcelSheet.Cells[j + l, 3].Value) * Convert.ToDouble(ExcelSheet.Cells[j + l, 4].Value);
                                dataCell.Value = cellValue;
                                dataCell.NumberFormat = "0.00";
                                break;
                            default:
                                dataCell.Value = DataTab.Rows[j - 1][i - 1];
                                dataCell.NumberFormat = "@";
                                break;
                        }

                        
                        if (cellValue != null)
                        {
                            string cellStringValue = cellValue.ToString();

                            if (lh[columnIndex] < cellStringValue.Length * 1.3)
                            {
                                lh[columnIndex] = Convert.ToInt32(cellStringValue.Length * 1.3);
                            }
                        }

                        ExcelSheet.Columns[i].ColumnWidth = lh[columnIndex];
                        dataCell.Font.Size = 12;
                        dataCell.Font.Name = "Times New Roman";
                        dataCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        dataCell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    }
                }


                // Calcul du Total
                Total = 0;
                for (i = 1; i <= DataTab.Rows.Count; i++)
                {
                    Total += Convert.ToDouble(ExcelSheet.Cells[i + l, 3].Value) * Convert.ToDouble(ExcelSheet.Cells[i + l, 4].Value);
                }

                ExcelSheet.Range["A" + (l + 1 + DataTab.Rows.Count) + ":D" + (l + 1 + DataTab.Rows.Count)].Merge();
                ExcelSheet.Cells[l + 1 + DataTab.Rows.Count, 1] = "Total HT en Dirhams :";
                ExcelSheet.Cells[l + 1 + DataTab.Rows.Count, 1].Interior.color = Color.Gold;
                ExcelSheet.Cells[l + 1 + DataTab.Rows.Count, 1].font.Bold = true;
                ExcelSheet.Cells[l + 1 + DataTab.Rows.Count, 1].font.size = 13;
                ExcelSheet.Cells[l + 1 + DataTab.Rows.Count, 1].font.Name = "Times New Roman";
                ExcelSheet.Cells[l + 1 + DataTab.Rows.Count, 1].HorizontalAlignment = 4;
                ExcelSheet.Cells[l + 1 + DataTab.Rows.Count, 1].VerticalAlignment = HorizontalAlignment.Center;
                ExcelSheet.Cells[l + 1 + DataTab.Rows.Count, 1].Borders.Color = Color.Black;
                ExcelSheet.Cells[l + 1 + DataTab.Rows.Count, 5] = Math.Round(Total, 2);
                ExcelSheet.Cells[l + 1 + DataTab.Rows.Count, 5].NumberFormat = "0,00";
                ExcelSheet.Cells[l + 1 + DataTab.Rows.Count, 5].Interior.color = Color.Gold;
                ExcelSheet.Cells[l + 1 + DataTab.Rows.Count, 5].font.Bold = true;
                ExcelSheet.Cells[l + 1 + DataTab.Rows.Count, 5].font.size = 13;
                ExcelSheet.Cells[l + 1 + DataTab.Rows.Count, 5].font.Name = "Times New Roman";
                ExcelSheet.Cells[l + 1 + DataTab.Rows.Count, 5].HorizontalAlignment = 4;
                ExcelSheet.Cells[l + 1 + DataTab.Rows.Count, 5].VerticalAlignment = HorizontalAlignment.Center;
                ExcelSheet.Cells[l + 1 + DataTab.Rows.Count, 5].Borders.Color = Color.Black;

                ExcelSheet.Range["A" + (l + 2 + DataTab.Rows.Count) + ":D" + (l + 2 + DataTab.Rows.Count)].Merge();
                ExcelSheet.Cells[l + 2 + DataTab.Rows.Count, 1] = "Montant de la révision des prix en Dirhams HT :";
                ExcelSheet.Cells[l + 2 + DataTab.Rows.Count, 1].Interior.Color = Color.Gold;
                ExcelSheet.Cells[l + 2 + DataTab.Rows.Count, 1].Font.Bold = true;
                ExcelSheet.Cells[l + 2 + DataTab.Rows.Count, 1].Font.Size = 13;
                ExcelSheet.Cells[l + 2 + DataTab.Rows.Count, 1].Font.Name = "Times New Roman";
                ExcelSheet.Cells[l + 2 + DataTab.Rows.Count, 1].HorizontalAlignment = XlHAlign.xlHAlignRight;
                ExcelSheet.Cells[l + 2 + DataTab.Rows.Count, 1].VerticalAlignment = XlVAlign.xlVAlignCenter;
                ExcelSheet.Cells[l + 2 + DataTab.Rows.Count, 1].Borders.Color = Color.Black;
                ExcelSheet.Cells[l + 2 + DataTab.Rows.Count, 5] = Math.Round(MREV, 2);
                ExcelSheet.Cells[l + 2 + DataTab.Rows.Count, 5].NumberFormat = "0,00";
                ExcelSheet.Cells[l + 2 + DataTab.Rows.Count, 5].Interior.Color = Color.Gold;
                ExcelSheet.Cells[l + 2 + DataTab.Rows.Count, 5].Font.Bold = true;
                ExcelSheet.Cells[l + 2 + DataTab.Rows.Count, 5].Font.Size = 13;
                ExcelSheet.Cells[l + 2 + DataTab.Rows.Count, 5].Font.Name = "Times New Roman";
                ExcelSheet.Cells[l + 2 + DataTab.Rows.Count, 5].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                ExcelSheet.Cells[l + 2 + DataTab.Rows.Count, 5].VerticalAlignment = XlVAlign.xlVAlignCenter;
                ExcelSheet.Cells[l + 2 + DataTab.Rows.Count, 5].Borders.Color = Color.Black;

                ExcelSheet.Range["A" + (l + 3 + DataTab.Rows.Count) + ":D" + (l + 3 + DataTab.Rows.Count)].Merge();
                ExcelSheet.Cells[l + 3 + DataTab.Rows.Count, 1] = "Total HT en Dirhams :";
                ExcelSheet.Cells[l + 3 + DataTab.Rows.Count, 1].Interior.Color = Color.Gold;
                ExcelSheet.Cells[l + 3 + DataTab.Rows.Count, 1].Font.Bold = true;
                ExcelSheet.Cells[l + 3 + DataTab.Rows.Count, 1].Font.Size = 13;
                ExcelSheet.Cells[l + 3 + DataTab.Rows.Count, 1].Font.Name = "Times New Roman";
                ExcelSheet.Cells[l + 3 + DataTab.Rows.Count, 1].HorizontalAlignment = XlHAlign.xlHAlignRight;
                ExcelSheet.Cells[l + 3 + DataTab.Rows.Count, 1].VerticalAlignment = XlVAlign.xlVAlignCenter;
                ExcelSheet.Cells[l + 3 + DataTab.Rows.Count, 1].Borders.Color = Color.Black;
                ExcelSheet.Cells[l + 3 + DataTab.Rows.Count, 5] = Math.Round((Total + MREV), 2);
                ExcelSheet.Cells[l + 3 + DataTab.Rows.Count, 5].NumberFormat = "0,00";
                ExcelSheet.Cells[l + 3 + DataTab.Rows.Count, 5].Interior.Color = Color.Gold;
                ExcelSheet.Cells[l + 3 + DataTab.Rows.Count, 5].Font.Bold = true;
                ExcelSheet.Cells[l + 3 + DataTab.Rows.Count, 5].Font.Size = 13;
                ExcelSheet.Cells[l + 3 + DataTab.Rows.Count, 5].Font.Name = "Times New Roman";
                ExcelSheet.Cells[l + 3 + DataTab.Rows.Count, 5].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                ExcelSheet.Cells[l + 3 + DataTab.Rows.Count, 5].VerticalAlignment = XlVAlign.xlVAlignCenter;
                ExcelSheet.Cells[l + 3 + DataTab.Rows.Count, 5].Borders.Color = Color.Black;

                ExcelSheet.Range["A" + (l + 4 + DataTab.Rows.Count) + ":D" + (l + 4 + DataTab.Rows.Count)].Merge();
                ExcelSheet.Cells[l + 4 + DataTab.Rows.Count, 1] = "TVA en Dirhams (20%) :";
                ExcelSheet.Cells[l + 4 + DataTab.Rows.Count, 1].Interior.Color = Color.Gold;
                ExcelSheet.Cells[l + 4 + DataTab.Rows.Count, 1].Font.Bold = true;
                ExcelSheet.Cells[l + 4 + DataTab.Rows.Count, 1].Font.Size = 13;
                ExcelSheet.Cells[l + 4 + DataTab.Rows.Count, 1].Font.Name = "Times New Roman";
                ExcelSheet.Cells[l + 4 + DataTab.Rows.Count, 1].HorizontalAlignment = XlHAlign.xlHAlignRight;
                ExcelSheet.Cells[l + 4 + DataTab.Rows.Count, 1].VerticalAlignment = XlVAlign.xlVAlignCenter;
                ExcelSheet.Cells[l + 4 + DataTab.Rows.Count, 1].Borders.Color = Color.Black;
                ExcelSheet.Cells[l + 4 + DataTab.Rows.Count, 5] = Math.Round((Total + MREV) * 0.2, 2);
                ExcelSheet.Cells[l + 4 + DataTab.Rows.Count, 5].NumberFormat = "0,00";
                ExcelSheet.Cells[l + 4 + DataTab.Rows.Count, 5].Interior.Color = Color.Gold;
                ExcelSheet.Cells[l + 4 + DataTab.Rows.Count, 5].Font.Bold = true;
                ExcelSheet.Cells[l + 4 + DataTab.Rows.Count, 5].Font.Size = 13;
                ExcelSheet.Cells[l + 4 + DataTab.Rows.Count, 5].Font.Name = "Times New Roman";
                ExcelSheet.Cells[l + 4 + DataTab.Rows.Count, 5].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                ExcelSheet.Cells[l + 4 + DataTab.Rows.Count, 5].VerticalAlignment = XlVAlign.xlVAlignCenter;
                ExcelSheet.Cells[l + 4 + DataTab.Rows.Count, 5].Borders.Color = Color.Black;

                ExcelSheet.Range["A" + (l + 2 + DataTab.Rows.Count) + ":D" + (l + 2 + DataTab.Rows.Count)].Borders.Color = Color.Black.ToArgb();
                ExcelSheet.Range["A" + (l + 5 + DataTab.Rows.Count) + ":D" + (l + 5 + DataTab.Rows.Count)].Merge();
                ExcelSheet.Range["A" + (l + 5 + DataTab.Rows.Count) + ":D" + (l + 5 + DataTab.Rows.Count)].Borders.Color = Color.Black.ToArgb();

                ExcelSheet.Cells[l + 5 + DataTab.Rows.Count, 1] = "Montant TTC en Dirhams :";
                ExcelSheet.Cells[l + 5 + DataTab.Rows.Count, 1].Interior.Color = Color.Gold;
                ExcelSheet.Cells[l + 5 + DataTab.Rows.Count, 1].Font.Bold = true;
                ExcelSheet.Cells[l + 5 + DataTab.Rows.Count, 1].Font.Size = 13;
                ExcelSheet.Cells[l + 5 + DataTab.Rows.Count, 1].Font.Name = "Times New Roman";
                ExcelSheet.Cells[l + 5 + DataTab.Rows.Count, 1].HorizontalAlignment = XlHAlign.xlHAlignRight;
                ExcelSheet.Cells[l + 5 + DataTab.Rows.Count, 1].VerticalAlignment = XlVAlign.xlVAlignCenter;
                ExcelSheet.Cells[l + 5 + DataTab.Rows.Count, 1].Borders.Color = Color.Black;
                ExcelSheet.Cells[l + 5 + DataTab.Rows.Count, 5] = Math.Round((Total + MREV) * 1.2, 2);
                ExcelSheet.Cells[l + 5 + DataTab.Rows.Count, 5].NumberFormat = "0,00";
                ExcelSheet.Cells[l + 5 + DataTab.Rows.Count, 5].Interior.Color = Color.Gold;
                ExcelSheet.Cells[l + 5 + DataTab.Rows.Count, 5].Font.Bold = true;
                ExcelSheet.Cells[l + 5 + DataTab.Rows.Count, 5].Font.Size = 13;
                ExcelSheet.Cells[l + 5 + DataTab.Rows.Count, 5].Font.Name = "Times New Roman";
                ExcelSheet.Cells[l + 5 + DataTab.Rows.Count, 5].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                ExcelSheet.Cells[l + 5 + DataTab.Rows.Count, 5].VerticalAlignment = XlVAlign.xlVAlignCenter;
                ExcelSheet.Cells[l + 5 + DataTab.Rows.Count, 5].Borders.Color = Color.Black;
            }

            // ==================================== feuille de calcul Récapitulatif ========================================
            DataAdap = new OleDbDataAdapter("Select Num_Prix,Désignation,Unité,Prix,Quantité From Detail_Estimatif where Num_Marché='" + CmbNumMarchéM.Text + "'", connection);
            DataTab = new DataTable();
            DataAdap.Fill(DataTab);
            if (connections.State == ConnectionState.Closed)
            {
                connections.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data source=" + Application.StartupPath + "\\revision.accdb";
                connections.Open();
            }
            ExcelSheet = ExcelBook.Worksheets[3];
            ExcelSheet.Name = "Récapitulatif";

            ExcelSheet.Range["A1"].Value = " Objet : " + TxtObjet.Text;
            ExcelSheet.Rows["1:1"].RowHeight = 12;
            ExcelSheet.Range["A2"].Value = " Numéro de Marché : " + CmbNumMarchéM.Text;
            ExcelSheet.Rows["2:2"].RowHeight = 12;
            ExcelSheet.Range["A3"].Value = " Entreprise : " + TxtRC.Text;
            ExcelSheet.Rows["3:3"].RowHeight = 12;
            ExcelSheet.Range["A4"].Value = " Adresse : " + TxtAdresse.Text;
            ExcelSheet.Rows["4:4"].RowHeight = 12;
            ExcelSheet.Range["A5"].Value = " Décompte Provisoire N°:" + CmbDécompteM.Text.Substring(2, CmbDécompteM.Text.Length - 2);
            ExcelSheet.Rows["5:5"].RowHeight = 12;
            ExcelSheet.Range["A8:D8"].Merge();
            ExcelSheet.Cells[8, 1].Value = "Récapitulatif";
            ExcelSheet.Cells[8, 1].Interior.Color = Color.SkyBlue;
            ExcelSheet.Cells[8, 1].Font.Bold = true;
            ExcelSheet.Cells[8, 1].Font.Size = 13;
            ExcelSheet.Cells[8, 1].Font.Name = "Times New Roman";
            ExcelSheet.Cells[8, 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            ExcelSheet.Cells[8, 1].VerticalAlignment = XlVAlign.xlVAlignCenter;
            ExcelSheet.Rows["8:8"].RowHeight = 25;

            //***************************************
            // with ExcelSheet
            ExcelSheet.Cells[10, 1] = "Nature des Dépenses";
            ExcelSheet.Cells[10, 2] = "Montant des Dépenses Faites";
            ExcelSheet.Cells[10, 3] = "Retenue de Garantie";
            ExcelSheet.Cells[10, 4] = "Reste";

            for ( j = 1; j <= 4; j++)
            {
                ExcelSheet.Cells[10, j].Font.Bold = true;
                ExcelSheet.Cells[10, j].Font.Size = 13;
                ExcelSheet.Cells[10, j].Font.Name = "Times New Roman";
                ExcelSheet.Cells[10, j].HorizontalAlignment = XlHAlign.xlHAlignRight;
                ExcelSheet.Cells[10, j].VerticalAlignment = XlVAlign.xlVAlignCenter;
                ExcelSheet.Cells[10, j].NumberFormat = "@";
                ExcelSheet.Cells[10, j].Interior.Color = Color.Gold;
                ExcelSheet.Cells[10, j].Borders.Color = Color.Black;

                object cellValue = ExcelSheet.Cells[10, j].Value;

                if (cellValue != null)
                {
                    string cellStringValue = cellValue.ToString();

                    if (lh[j - 1] < cellStringValue.Length * 1.3)
                    {
                        lh[j - 1] = Convert.ToInt32(cellStringValue.Length * 1.3);
                    }
                }

                ExcelSheet.Columns[j].ColumnWidth = lh[j - 1];
            }

            //double RetenueGarantie;
            MontantMarché = 0;
            Total = 0;
            if (Cautionnement == "CB")
            {
                RetenueGarantie = Math.Round(0.00, 2);
            }
            else
            {
                if ((Total * 0.1 * 1.2) >= (MontantMarché * 0.07))
                {
                    RetenueGarantie = Math.Round(MontantMarché * 0.07, 2);
                }
                else
                {
                    RetenueGarantie = Math.Round((Total + MREV) * 0.1 * 1.2, 2);
                }
            }

            ExcelSheet.Cells[11, 1] = "Travaux non terminés";
            ExcelSheet.Cells[11, 2] = Math.Round((Total + MREV) * 1.2, 2);
            ExcelSheet.Cells[11, 3] = Math.Round(RetenueGarantie, 2);
            ExcelSheet.Cells[11, 4] = Math.Round((Total + MREV) * 1.2 - RetenueGarantie, 2);

            ExcelSheet.Cells[11, 2].NumberFormat = "0.00";
            ExcelSheet.Cells[11, 3].NumberFormat = "0.00";
            ExcelSheet.Cells[11, 4].NumberFormat = "0.00";

            for ( j = 1; j <= 4; j++)
            {
                ExcelSheet.Cells[11, j].Font.Bold = true;
                ExcelSheet.Cells[11, j].Font.Size = 13;
                ExcelSheet.Cells[11, j].Font.Name = "Times New Roman";

                if (j == 1)
                {
                    ExcelSheet.Cells[11, j].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                }
                else
                {
                    ExcelSheet.Cells[11, j].HorizontalAlignment = XlHAlign.xlHAlignRight;
                }

                ExcelSheet.Cells[11, j].VerticalAlignment = XlVAlign.xlVAlignCenter;
                ExcelSheet.Cells[11, j].NumberFormat = "@";
                ExcelSheet.Cells[11, j].Borders.Color = Color.Black;

                if (ExcelSheet.Cells[11, j].Value != null && ExcelSheet.Cells[11, j].Value is string)
                {
                    if (lh[j - 1] < Convert.ToInt32(ExcelSheet.Cells[11, j].Value.ToString().Length * 1.3))
                    {
                        lh[j - 1] = Convert.ToInt32(ExcelSheet.Cells[11, j].Value.ToString().Length * 1.3);
                    }
                }

                ExcelSheet.Columns[j].ColumnWidth = lh[j - 1];
            }

            // ---Totaux
            ExcelSheet.Cells[12, 1] = "Total";
            ExcelSheet.Cells[12, 2] = Math.Round((Total + MREV) * 1.2, 2);
            ExcelSheet.Cells[12, 3] = Math.Round(RetenueGarantie, 2);
            ExcelSheet.Cells[12, 4] = ExcelSheet.Cells[11, 4].Value;
            ExcelSheet.Cells[12, 4].NumberFormat = "0,00";

            for ( j = 1; j <= 4; j++)
            {
                if (ExcelSheet.Cells[12, j].Value != null)
                {
                    ExcelSheet.Cells[12, j].Font.Bold = true;
                    ExcelSheet.Cells[12, j].Font.Size = 13;
                    ExcelSheet.Cells[12, j].Font.Name = "Times New Roman";

                    if (j == 1)
                    {
                        ExcelSheet.Cells[12, j].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                    }
                    else
                    {
                        ExcelSheet.Cells[12, j].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                    }

                    ExcelSheet.Cells[12, j].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    ExcelSheet.Cells[12, j].NumberFormat = "@";
                    ExcelSheet.Cells[12, j].Interior.Color = Color.Gold;
                    ExcelSheet.Cells[12, j].Borders.Color = Color.Black;

                    if (lh[j - 1] < Convert.ToInt32(ExcelSheet.Cells[12, j].Value.ToString().Length * 1.3))
                    {
                        lh[j - 1] = Convert.ToInt32(ExcelSheet.Cells[12, j].Value.ToString().Length * 1.3);
                    }
                }

                ExcelSheet.Columns[j].ColumnWidth = lh[j - 1];
            }

            ExcelSheet.Cells[13, 1] = "Reste à payer sur l'exercice en cours ";
            ExcelSheet.Cells[13, 4] = ExcelSheet.Cells[11, 4].Value;
            ExcelSheet.Cells[13, 4].NumberFormat = "0,00";

            for ( j = 1; j <= 4; j++)
            {
                ExcelSheet.Cells[13, j].Font.Bold = true;
                ExcelSheet.Cells[13, j].Font.Size = 13;
                ExcelSheet.Cells[13, j].Font.Name = "Times New Roman";

                if (j == 1)
                {
                    ExcelSheet.Cells[13, j].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                }
                else
                {
                    ExcelSheet.Cells[13, j].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                }

                ExcelSheet.Cells[13, j].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                ExcelSheet.Cells[13, j].NumberFormat = "@";
                ExcelSheet.Cells[13, j].Borders.Color = Color.Black;

                object cellValue = ExcelSheet.Cells[13, j].Value;

                if (cellValue != null)
                {
                    string cellStringValue = cellValue.ToString();

                    int newLength = (int)(Convert.ToDouble(cellStringValue.Length * 1.3));

                    if (lh[j - 1] < newLength)
                    {
                        lh[j - 1] = newLength;
                    }
                }

                if (string.IsNullOrEmpty((string)ExcelSheet.Cells[13, 1].Value))
                {
                    // Fusionner les cellules uniquement si la cellule A13 est vide
                    ExcelSheet.Range["A13:C13"].Merge();
                }
            }
            
            MontantDP = 0; // Assurez-vous que MontantDP est initialisée
            ExcelSheet.Cells[14, 1] = "A déduire le montant des acomptes délivrés sur l'exercice en cours ";
            if (INDEXS == 0)
            {
                ExcelSheet.Cells[14, 4] = 0;
            }
            else
            {
                ExcelSheet.Cells[14, 4] = Math.Round(MontantDP, 2);
            }
            ExcelSheet.Cells[14, 4].NumberFormat = "0,00";

            for (j = 1; j <= 4; j++)
            {
                ExcelSheet.Cells[14, j].Font.Bold = true;
                ExcelSheet.Cells[14, j].Font.Size = 13;
                ExcelSheet.Cells[14, j].Font.Name = "Times New Roman";

                if (j == 1)
                {
                    ExcelSheet.Cells[14, j].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                }
                else
                {
                    ExcelSheet.Cells[14, j].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                }

                ExcelSheet.Cells[14, j].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                ExcelSheet.Cells[14, j].NumberFormat = "@";
                ExcelSheet.Cells[14, j].Borders.Color = Color.Black;

                object cellValue = ExcelSheet.Cells[14, j].Value;

                if (cellValue != null)
                {
                    string cellStringValue = cellValue.ToString();

                    int newLength = (int)(Convert.ToDouble(cellStringValue.Length * 1.3));

                    if (lh[j - 1] < newLength)
                    {
                        lh[j - 1] = newLength;
                    }
                }

                ExcelSheet.Columns[j].ColumnWidth = lh[j - 1];
            }

            // Fusion des cellules après avoir défini le style
            if (string.IsNullOrEmpty((string)ExcelSheet.Cells[13, 1].Value))
            {
                // Fusionner les cellules uniquement si la cellule A13 est vide
                ExcelSheet.Range["A13:C13"].Merge();
            }

            ExcelSheet.Cells[15, 1] = "A déduire le remboursement d'avance(10%) ";
            DataAdaps = new OleDbDataAdapter("Select Désignation, Tranche_Avance From Décompte where Num_Marché='" + CmbNumMarchéM.Text + "'", connections);
            DataTabs = new DataTable();
            DataAdaps.Fill(DataTabs);

            double MDTAvance = 0;
            for (i = 0; i < DataTab.Rows.Count; i++)
            {
                if (DataTabs.Rows[i]["Désignation"].ToString() == CmbDécompteM.Text)
                    break;

                MDTAvance += Convert.ToDouble(DataTabs.Rows[i]["Tranche_Avance"]);
            }

            if (MontantAvance - MDTAvance < MontantAvance * 0.2)
            {
                ExcelSheet.Cells[15, 4] = Math.Round(MontantAvance * 0.2, 2);
            }
            else
            {
                ExcelSheet.Cells[15, 4] = Math.Round(MontantAvance - MDTAvance, 2);
            }

            if (TDécompte == "2")
            {
                ExcelSheet.Cells[15, 4] = Math.Round(MontantAvance, 2);
            }
            else
            {
                ExcelSheet.Cells[15, 4] = Math.Round(MontantAvance * 0.2, 2);
            }

            for ( j = 1; j <= 4; j++)
            {
                ExcelSheet.Cells[15, j].Font.Bold = true;
                ExcelSheet.Cells[15, j].Font.Size = 13;
                ExcelSheet.Cells[15, j].Font.Name = "Times New Roman";

                if (j == 1)
                {
                    ExcelSheet.Cells[15, j].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                }
                else
                {
                    ExcelSheet.Cells[15, j].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                }

                ExcelSheet.Cells[15, j].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                ExcelSheet.Cells[15, j].NumberFormat = "@";
                ExcelSheet.Cells[15, j].Borders.Color = Color.Black;

                object cellValue = ExcelSheet.Cells[15, j]?.Value;

                if (cellValue != null)
                {
                    string cellStringValue = cellValue.ToString();

                    int newLength = (int)(Convert.ToDouble(cellStringValue.Length * 1.3));

                    if (lh[j - 1] < newLength)
                    {
                        lh[j - 1] = newLength;
                    }
                }

                ExcelSheet.Columns[j].ColumnWidth = lh[j - 1];
            }
            if (string.IsNullOrEmpty((string)ExcelSheet.Cells[15, 1].Value))
            {
                // Fusionner les cellules uniquement si la cellule A13 est vide
                ExcelSheet.Range["A15:C15"].Merge();
            }

            //-----MONTANT DE L'ACOMPTE A DELIVRER  T.T.C------------
            ExcelSheet.Cells[16, 1] = "MONTANT DE L'ACOMPTE A DELIVRER  T.T.C";
            ExcelSheet.Cells[16, 4] = ExcelSheet.Cells[11, 4].Value - (double)ExcelSheet.Cells[14, 4].Value - (double)ExcelSheet.Cells[15, 4].Value;
            Accompte = Math.Round(Convert.ToDouble(ExcelSheet.Cells[16, 4].Value), 2);

            for ( j = 1; j <= 4; j++)
            {
                ExcelSheet.Cells[16, j].Font.Bold = true;
                ExcelSheet.Cells[16, j].Font.Size = 13;
                ExcelSheet.Cells[16, j].Font.Name = "Times New Roman";

                if (j == 1)
                {
                    ExcelSheet.Cells[16, j].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                }
                else
                {
                    ExcelSheet.Cells[16, j].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                }

                ExcelSheet.Cells[16, j].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                ExcelSheet.Cells[16, j].NumberFormat = "@";
                ExcelSheet.Cells[16, j].Interior.Color = Color.Gold;
                ExcelSheet.Cells[16, j].Borders.Color = Color.Black;

                object cellValue = ExcelSheet.Cells[16, j]?.Value;

                if (cellValue != null)
                {
                    string cellStringValue = cellValue.ToString();

                    // Mise à jour de la longueur seulement si la valeur est un nombre
                    double doubleValue;
                    if (double.TryParse(cellStringValue, out doubleValue))
                    {
                        int newLength = (int)(cellStringValue.Length * 1.3);

                        if (lh[j - 1] < newLength)
                        {
                            lh[j - 1] = newLength;
                        }
                    }
                }

                ExcelSheet.Columns[j].ColumnWidth = lh[j - 1];
            }
            if (string.IsNullOrEmpty((string)ExcelSheet.Cells[15, 1].Value))
            {
                // Fusionner les cellules uniquement si la cellule A13 est vide
                ExcelSheet.Range["A15:C15"].Merge();
            }
            //ExcelSheet.Range["A15:C15"].Merge();


            //------------Arrêté par nous, ordonnateur, le présent décompte Provisoire N°--------
            str = " Arrêté par nous, ordonnateur, le présent décompte Provisoire N°: ";
            str += CmbDécompteM.Text.Substring(2, CmbDécompteM.Text.Length - 2);

            if (TDécompte == "2")
            {
                str += " et Dérnier ";
            }
            str += " à la somme de : " + Calcul.ChiffreToLettre(Convert.ToString(Accompte).Replace(",", ".")) + ".";
            ExcelSheet.Cells[18, 1] = str;

            for ( j = 1; j <= 4; j++)
            {
                ExcelSheet.Cells[18, j].Font.Bold = true;
                ExcelSheet.Cells[18, j].Font.Size = 13;
                ExcelSheet.Cells[18, j].Font.Name = "Times New Roman";

                if (j == 1)
                {
                    ExcelSheet.Cells[18, j].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                }
                else
                {
                    ExcelSheet.Cells[18, j].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                }

                ExcelSheet.Cells[18, j].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                ExcelSheet.Cells[18, j].NumberFormat = "@";
                ExcelSheet.Cells[18, j].Interior.Color = Color.Gold;
                ExcelSheet.Cells[18, j].Borders.Color = Color.Black;
                ExcelSheet.Cells[18, j].WrapText = true;
            }
            if (string.IsNullOrEmpty((string)ExcelSheet.Cells[18, 1].Value))
            {
                // Fusionner les cellules uniquement si la cellule A13 est vide
                ExcelSheet.Range["A18:D18"].Merge();
            }
            //ExcelSheet.Range["A18:D18"].Merge();

            object CellValue = ExcelSheet.Cells[11, 4].Value;
            string montant = CellValue != null ? Convert.ToString(CellValue).Replace(",", ".") : "0,0";

            SQLSTR = "UPDATE Décompte SET ";
            SQLSTR += "Montant=" + montant + ",";
            SQLSTR += "Retenue_Garantie=" + RetenueGarantie.ToString().Replace(",", ".");
            SQLSTR += " WHERE Num_Marché='" + CmbNumMarchéM.Text + "' AND Désignation='" + CmbDécompteM.Text + "'";
            ExecuterSQL(SQLSTR);


            ExcelSheet = (Excel.Worksheet)ExcelBook.Worksheets[2];
            ExcelSheet.Name = "Décompte";

            // Assuming Accompte is a variable holding the value to be assigned to the cell
            ExcelSheet.Cells[18, 5] = Accompte;


            if (INDEXS > 0)
                connections.Close();
            ExcelApp.Visible = true;
            ExcelSheet = null;
            ExcelBook = null;
            ExcelApp = null;

        }
        private void FrmEditAttachement_Resize(object sender, EventArgs e)
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
