using Microsoft.Office.Core;
using System;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Révision
{
    public partial class FrmDécompteDéfinitif : Form
    {
        OleDbConnection connection = new OleDbConnection();
        OleDbCommand records = new OleDbCommand();
        DataTable DataTab = new DataTable();
        OleDbDataAdapter DataAdap = new OleDbDataAdapter();
        OleDbCommandBuilder ComndBuld = new OleDbCommandBuilder();
        DataSet DataSetTab = new DataSet();

        OleDbConnection connections = new OleDbConnection();
        OleDbCommand recordss = new OleDbCommand();
        DataTable DataTabs = new DataTable();
        OleDbDataAdapter DataAdaps = new OleDbDataAdapter();
        OleDbCommandBuilder ComndBulds = new OleDbCommandBuilder();
        DataSet DataSetTabs = new DataSet();

        int INDEXS;
        string SQLSTR;
        long NumDécompte, NumDécompteM, NumDécompteMA;
        long NumEntrepreneur;
        public FrmDécompteDéfinitif()
        {
            InitializeComponent();
        }
        private void FrmDécompteDéfinitif_FormClosed(object sender, FormClosedEventArgs e)
        {
            connection.Close();
        }
        private void CmbNumMarchéM_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            if (string.IsNullOrEmpty(CmbNumMarchéM.Text))
            {
                MessageBox.Show("Choisissez une valeur de la liste : Référence de Marché ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                CmbNumMarchéM.Focus();
                return;
            }

            OleDbDataAdapter DataAdap;
            DataTable DataTab;

            DataAdap = new OleDbDataAdapter("Select Désignation From Décompte where Num_Marché='" + CmbNumMarchéM.Text + "'", connection);
            DataTab = new DataTable();
            DataAdap.Fill(DataTab);

            DataAdap = new OleDbDataAdapter("Select * From Marché where Réference='" + CmbNumMarchéM.Text + "'", connection);
            DataTab = new DataTable();
            DataAdap.Fill(DataTab);
            if (DataTab.Rows.Count > 0)
            {
                TxtObjet.Text = DataTab.Rows[0][1].ToString();
                NumEntrepreneur = Convert.ToInt64(DataTab.Rows[0][6]);
            }

            DataAdap = new OleDbDataAdapter("SELECT * FROM Entrepreneur WHERE N° = ?", connection);
            DataAdap.SelectCommand.Parameters.AddWithValue("@NumEntrepreneur", NumEntrepreneur);

            DataTab = new DataTable();
            DataAdap.Fill(DataTab);
            if (DataTab.Rows.Count > 0)
            {
                TxtNom.Text = DataTab.Rows[0][1].ToString();
                TxtAdresse.Text = DataTab.Rows[0][2].ToString();
                TxtCNSS.Text = DataTab.Rows[0][3].ToString();
                TxtPatente.Text = DataTab.Rows[0][4].ToString();
                TxtNCB.Text = DataTab.Rows[0][5].ToString();
                TxtAB.Text = DataTab.Rows[0][6].ToString();
                TxtRC.Text = DataTab.Rows[0][7].ToString();
                TxtTel.Text = DataTab.Rows[0][8].ToString();
                TxtEMail.Text = DataTab.Rows[0][9].ToString();
                TxtIF.Text = DataTab.Rows[0][10].ToString();
                TxtICE.Text = DataTab.Rows[0][11].ToString();
            }
            else
            {
                MessageBox.Show("Aucun valeur n'ai saisi dans la table détail éstimatif pour le marché numéro : " + CmbNumMarchéM.Text + " .", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void FrmDécompteDéfinitif_Load(object sender, EventArgs e)
        {
            if (connection.State == ConnectionState.Closed)
            {
                connection.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data source=" + Application.StartupPath + "\\revision.accdb";
                connection.Open();
            }

            Recharger();
        }
        private void Recharger()
        {
            using (OleDbDataAdapter dataAdapter = new OleDbDataAdapter("Select Réference From Marché", connection))
            {
                DataTable dataTable = new DataTable();
                dataAdapter.Fill(dataTable);

                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    CmbNumMarchéM.Items.Add(dataTable.Rows[i][0].ToString());
                }
            }
        }

        private void BtnEditer_Click(object sender, EventArgs e)
        {
            Excel.Application ExcelApp;
            Excel.Workbook ExcelBook;
            Excel.Worksheet ExcelSheet;
            int i;
            int j;
            int l;
            string str;
            double Total;
            double MontantTTC;
            double MontantDP;
            double MontantMarché;
            double MontantRévisé;
            double Pénalité;
            string TDécompte = "";
            int[] lh = new int[6]; // Tableau d'entiers de taille 6
            long Durée;

            if (string.IsNullOrEmpty(CmbNumMarchéM.Text))
            {
                MessageBox.Show("Choisissez une valeur de la liste : Réference de Marché ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                CmbNumMarchéM.Focus();
                return;
            }

            // Créer un objet Excel
            ExcelApp = new Excel.Application();
            ExcelBook = ExcelApp.Workbooks.Add();
            ExcelSheet = ExcelBook.Worksheets[1];

            // Feuille de calcul attachement
            ExcelSheet = ExcelBook.Worksheets.Add();
            ExcelSheet = ExcelBook.Worksheets[2];
            ExcelSheet = ExcelBook.Worksheets[1];
            ExcelSheet.Name = "Attachement Définitif";

            ExcelSheet.Cells[1, 1] = "Royaume du Maroc";
            ExcelSheet.Rows["1:1"].RowHeight = 20;
            ExcelSheet.Cells[2, 1] = "Ministère de l'Intérieur";
            ExcelSheet.Rows["2:2"].RowHeight = 20;
            ExcelSheet.Cells[3, 1] = "Province d'Al Haouz";
            ExcelSheet.Rows["3:3"].RowHeight = 20;
            ExcelSheet.Cells[4, 1] = "Conseil Provincial Al Haouz";
            ExcelSheet.Rows["4:4"].RowHeight = 20;

            for (i = 1; i <= 4; i++)
            {
                ExcelSheet.Cells[i, 1].Font.Bold = true;
                ExcelSheet.Cells[i, 1].Font.Size = 12;
                ExcelSheet.Cells[i, 1].Font.Name = "Times New Roman";
                ExcelSheet.Cells[i, 1].HorizontalAlignment = Excel.Constants.xlLeft;
                ExcelSheet.Cells[i, 1].VerticalAlignment = XlVAlign.xlVAlignCenter;
            }

            for (j = 6; j <= 9; j++)
            {
                ExcelSheet.Cells[j, 1].Font.Bold = true;
                ExcelSheet.Cells[j, 1].Font.Size = 12;
                ExcelSheet.Cells[j, 1].Font.Name = "Times New Roman";
                ExcelSheet.Cells[j, 1].HorizontalAlignment = Excel.Constants.xlLeft;
                ExcelSheet.Cells[j, 1].VerticalAlignment = XlVAlign.xlVAlignCenter;
                ExcelSheet.Cells[j, 1].NumberFormat = "@";
            }

            ExcelSheet.Cells[6, 1] = " Objet : " + TxtObjet.Text;
            ExcelSheet.Rows["6:6"].RowHeight = 12;
            ExcelSheet.Range["A11:E11"].Merge();
            ExcelSheet.Cells[11, 1] = "Attachement Définitif";
            ExcelSheet.Cells[11, 1].Interior.Color = Color.SkyBlue;
            ExcelSheet.Cells[11, 1].Font.Bold = true;
            ExcelSheet.Cells[11, 1].Font.Size = 12;
            ExcelSheet.Cells[11, 1].Font.Name = "Times New Roman";
            ExcelSheet.Cells[11, 1].HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
            ExcelSheet.Cells[11, 1].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            ExcelSheet.Rows["11:11"].RowHeight = 20;

            l = 13;
            // Récupération des données du marché
            DataAdap = new OleDbDataAdapter("Select Montant_Global, Delai From Marché where Réference='" + CmbNumMarchéM.Text + "'", connection);
            DataTab = new DataTable();
            DataAdap.Fill(DataTab);
            MontantMarché = Convert.ToDouble(DataTab.Rows[0]["Montant_Global"]);
            long Délai = Convert.ToInt64(DataTab.Rows[0][1]);

            // Récupération des données du dernier décompte
            DataAdap = new OleDbDataAdapter("Select N°, Désignation From Décompte where Num_Marché='" + CmbNumMarchéM.Text + "'", connection);
            DataTab = new DataTable();
            DataAdap.Fill(DataTab);
            NumDécompte = (int)DataTab.Rows[DataTab.Rows.Count - 1][0];
            string Désignation = DataTab.Rows[DataTab.Rows.Count - 1][1].ToString();

            // Récupération des données du détail estimatif
            DataAdap = new OleDbDataAdapter("Select Num_Prix, Désignation, Unité, Quantité From Detail_Estimatif where Num_Marché='" + CmbNumMarchéM.Text + "'", connection);
            DataTab = new DataTable();
            DataAdap.Fill(DataTab);

            // Vérification et ouverture de la connexion à la base de données "revision.accdb"
            if (connections.State == ConnectionState.Closed)
            {
                connections.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data source=" + Application.StartupPath + "\\revision.accdb";
                connections.Open();
            }

            // Récupération des données de l'attachement
            DataAdaps = new OleDbDataAdapter("Select Quantité From Attachement where Num_Décompte=" + NumDécompte, connections);
            DataTabs = new DataTable();
            DataAdaps.Fill(DataTabs);

            if (connections.State == ConnectionState.Closed)
            {
                connections.ConnectionString = ("Provider=Microsoft.ACE.OLEDB.12.0;Data source=" + Application.StartupPath + @"\revision.accdb");
                connections.Open();
            }
            DataAdaps = new OleDbDataAdapter("Select Quantité From Attachement where Num_Décompte=" + NumDécompte, connections);
            DataTabs = new DataTable();
            DataAdaps.Fill(DataTabs);

            if(DataTab.Rows.Count > 0)
{
                ExcelSheet.Cells[13, 1] = "Numéro de Prix";
                ExcelSheet.Cells[13, 2] = "Désignation";
                ExcelSheet.Cells[13, 3] = "Unité";
                ExcelSheet.Cells[13, 4] = "Quantité Marché";
                ExcelSheet.Cells[13, 5] = "Quantité Réalisée";

                for (i = 1; i <= DataTab.Columns.Count + 1; i++)
                {
                    ExcelSheet.Cells[l, i].Interior.color = Color.Gold;
                    ExcelSheet.Cells[l, i].Borders.Color = Color.Black;
                    ExcelSheet.Cells[l, i].font.size = 12;
                    ExcelSheet.Cells[l, i].font.Bold = true;
                    ExcelSheet.Cells[l, i].font.Name = "Times New Roman";
                    ExcelSheet.Cells[l, i].HorizontalAlignment = 3;
                    ExcelSheet.Cells[l, i].VerticalAlignment = HorizontalAlignment.Center;

                    if (lh[i - 1] < ExcelSheet.Cells[l, i].Value.ToString().Length * 2)
                        lh[i - 1] = Convert.ToInt32(ExcelSheet.Cells[l, i].Value.ToString().Length * 1.3);

                    ExcelSheet.Columns[i].ColumnWidth = lh[i - 1];

                    for (j = 1; j <= DataTab.Rows.Count; j++)
                    {
                        ExcelSheet.Cells[j + l, i].Borders.Color = Color.Black;

                        if (i == 4 || i == 5)
                            ExcelSheet.Cells[j + l, i].NumberFormat = "0,00";
                        else
                            ExcelSheet.Cells[j + l, i].NumberFormat = "@";

                        if (i == DataTab.Columns.Count + 1)
                            ExcelSheet.Cells[j + l, i] = DataTabs.Rows[j - 1][0];
                        else
                            ExcelSheet.Cells[j + l, i] = DataTab.Rows[j - 1][i - 1].ToString();

                        if (lh[i - 1] < ExcelSheet.Cells[j + l, i].Value.ToString().Length * 2)
                            lh[i - 1] = Convert.ToInt32(ExcelSheet.Cells[j + l, i].Value.ToString().Length * 1.3);

                        ExcelSheet.Columns[i].ColumnWidth = lh[i - 1];
                        ExcelSheet.Cells[j + l, i].font.size = 12;
                        ExcelSheet.Cells[j + l, i].font.Name = "Times New Roman";
                        ExcelSheet.Cells[j + l, i].HorizontalAlignment = 3;
                        ExcelSheet.Cells[j + l, i].VerticalAlignment = HorizontalAlignment.Center;
                    }
                }
            }

            // =================================Feuille de calcul Décompte définitif==========================================
            DataAdap = new OleDbDataAdapter("Select Num_Prix,Désignation,Unité,Prix,Quantité From Detail_Estimatif where Num_Marché='" + CmbNumMarchéM.Text + "'", connection);
            DataTab = new DataTable();
            DataAdap.Fill(DataTab);

            ExcelSheet.Name = "Décompte Définitif";
            ExcelSheet = ExcelBook.Worksheets[2];

            {
                ExcelSheet.Cells[1, 1] = "Royaume du Maroc";
                ExcelSheet.Rows["1:1"].RowHeight = 15;
                ExcelSheet.Cells[2, 1] = "Ministère de l'Intérieur";
                ExcelSheet.Rows["2:2"].RowHeight = 15;
                ExcelSheet.Cells[3, 1] = "Province d'Al Haouz";
                ExcelSheet.Rows["3:3"].RowHeight = 15;
                ExcelSheet.Cells[4, 1] = "Conseil Provincial Al Haouz";
                ExcelSheet.Rows["4:4"].RowHeight = 15;

                for (j = 1; j <= 4; j++)
                {
                    ExcelSheet.Cells[j, 1].Font.Bold = true;
                    ExcelSheet.Cells[j, 1].Font.Size = 11;
                    ExcelSheet.Cells[j, 1].Font.Name = "Times New Roman";
                    ExcelSheet.Cells[j, 1].HorizontalAlignment = 1;
                    ExcelSheet.Cells[j, 1].VerticalAlignment = HorizontalAlignment.Center;
                }

                for (j = 6; j <= 18; j++)
                {
                    ExcelSheet.Cells[j, 1].Font.Bold = true;
                    ExcelSheet.Cells[j, 1].Font.Size = 11;
                    ExcelSheet.Cells[j, 1].Font.Name = "Times New Roman";
                    ExcelSheet.Cells[j, 1].HorizontalAlignment = 1;
                    ExcelSheet.Cells[j, 1].VerticalAlignment = HorizontalAlignment.Center;
                    ExcelSheet.Cells[j, 1].NumberFormat = "@";
                }

                ExcelSheet.Cells[6, 1] = " Objet : " + TxtObjet.Text;
                ExcelSheet.Rows["6:6"].RowHeight = 12;
                ExcelSheet.Cells[7, 1] = " Numéro de Marché : " + CmbNumMarchéM.Text;
                ExcelSheet.Rows["7:7"].RowHeight = 12;
                ExcelSheet.Cells[8, 1] = " Entreprise : " + TxtRC.Text;
                ExcelSheet.Rows["8:8"].RowHeight = 12;
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
                ExcelSheet.Rows["18:18"].RowHeight = 12;

                ExcelSheet.Cells[18, 5].Font.Bold = true;
                //ExcelSheet.Cells["18, 4"] = "Montant de l'Acompte TTC (dhs) : ";
                ExcelSheet.Cells[18, 5].Font.Bold = true;
                ExcelSheet.Cells[18, 5].Font.Italic = true;
                ExcelSheet.Cells[18, 5].Font.Size = 11;
                ExcelSheet.Cells[18, 5].Font.Name = "Times New Roman";
                ExcelSheet.Cells[18, 5].HorizontalAlignment = 3;
                ExcelSheet.Cells[18, 5].VerticalAlignment = HorizontalAlignment.Center;
                ExcelSheet.Cells[18, 5].NumberFormat = "@";
                ExcelSheet.Cells[18, 4].Font.Bold = true;
                ExcelSheet.Cells[18, 4].Font.Size = 11;
                ExcelSheet.Cells[18, 4].Font.Name = "Times New Roman";
                ExcelSheet.Cells[18, 4].HorizontalAlignment = 4;
                ExcelSheet.Cells[18, 4].VerticalAlignment = HorizontalAlignment.Center;
                ExcelSheet.Cells[18, 4].NumberFormat = "@";
                ExcelSheet.Range["A20:E20"].Merge();
                ExcelSheet.Cells[20, 1] = "Décompte Définitif";
                ExcelSheet.Cells[20, 1].Interior.Color = Color.SkyBlue;
                ExcelSheet.Cells[20, 1].Font.Bold = true;
                ExcelSheet.Cells[20, 1].Font.Size = 11;
                ExcelSheet.Cells[20, 1].Font.Name = "Times New Roman";
                ExcelSheet.Cells[20, 1].HorizontalAlignment = 3;
                ExcelSheet.Cells[20, 1].VerticalAlignment = HorizontalAlignment.Center;
                ExcelSheet.Rows["20:20"].RowHeight = 15;
            }

            l = 22;
            lh = new int[6]; // Assuming lh is an integer array with a size of 6
            DataAdap = new OleDbDataAdapter("Select Désignation,Unité,Prix_unitaire,Quantité From Attachement where Num_Décompte=" + NumDécompte, connection);
            DataTab = new DataTable();
            DataAdap.Fill(DataTab);
            DataAdaps = new OleDbDataAdapter("Select N°,Montant,Tranche_Avance,Type From Décompte where Désignation='" + Désignation + "'", connections);
            DataTabs = new DataTable();
            DataAdaps.Fill(DataTabs);
            int NumDécompteM = Convert.ToInt32(DataTabs.Rows[0][0]);
            MontantDP = Convert.ToDouble(DataTabs.Rows[0][1]);
            TDécompte = DataTabs.Rows[0][3].ToString();
            
            DataAdaps = new OleDbDataAdapter("Select Quantité From Attachement where Num_Décompte=" + NumDécompteM, connections);
            DataTabs = new DataTable();
            DataAdaps.Fill(DataTabs);

            MontantTTC = 0;
            if (DataTab.Rows.Count > 0)
            {
                {
                    ExcelSheet.Cells[22, 1].Value = "Désignation";
                    ExcelSheet.Cells[22, 2].Value = "Unité";
                    ExcelSheet.Cells[22, 3].Value = "Prix Unitaire";
                    ExcelSheet.Cells[22, 4].Value = "Quantité";
                    ExcelSheet.Cells[22, 5].Value = "Total Calculé";

                    for (i = 1; i <= DataTab.Columns.Count + 1; i++)
                    {
                        ExcelSheet.Cells[l, i].Interior.Color = Color.Gold;
                        ExcelSheet.Cells[l, i].Borders.Color = Color.Black;
                        ExcelSheet.Cells[l, i].Font.Size = 12;
                        ExcelSheet.Cells[l, i].Font.Bold = true;
                        ExcelSheet.Cells[l, i].Font.Name = "Times New Roman";
                        ExcelSheet.Cells[l, i].HorizontalAlignment = XlHAlign.xlHAlignRight;
                        ExcelSheet.Cells[l, i].VerticalAlignment = XlVAlign.xlVAlignCenter;

                        if (lh[i - 1] < ExcelSheet.Cells[l, i].Value.ToString().Length * 1.3)
                        {
                            lh[i - 1] = (int)(ExcelSheet.Cells[l, i].Value.ToString().Length * 1.3);
                        }

                        ExcelSheet.Columns[i].ColumnWidth = lh[i - 1];

                        for (j = 1; j <= DataTab.Rows.Count; j++)
                        {
                            ExcelSheet.Cells[j + l, i].Borders.Color = Color.Black;

                            switch (i)
                            {
                                case 4:
                                    {
                                        double value = Convert.ToDouble(DataTab.Rows[j - 1][i - 1]);
                                        ExcelSheet.Cells[j + l, i].Value = Math.Round(value, 2);
                                        ExcelSheet.Cells[j + l, i].NumberFormat = "0,00";
                                        break;
                                    }

                                case 5:
                                    {
                                        double value3 = Convert.ToDouble(ExcelSheet.Cells[j + l, 3].Value) * Convert.ToDouble(ExcelSheet.Cells[j + l, 4].Value);
                                        ExcelSheet.Cells[j + l, i].Value = Math.Round(value3, 2);
                                        ExcelSheet.Cells[j + l, i].NumberFormat = "0,00";
                                        break;
                                    }

                                default:
                                    ExcelSheet.Cells[j + l, i].Value = DataTab.Rows[j - 1][i - 1].ToString();
                                    ExcelSheet.Cells[j + l, i].NumberFormat = "@";
                                    break;
                            }

                            if (lh[i - 1] < ExcelSheet.Cells[j + l, i].Value.ToString().Length * 1.3)
                            {
                                lh[i - 1] = (int)(ExcelSheet.Cells[j + l, i].Value.ToString().Length * 1.3);
                            }

                            ExcelSheet.Columns[i].ColumnWidth = lh[i - 1];
                            ExcelSheet.Cells[j + l, i].Font.Size = 12;
                            ExcelSheet.Cells[j + l, i].Font.Name = "Times New Roman";
                            ExcelSheet.Cells[j + l, i].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                            ExcelSheet.Cells[j + l, i].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                        }
                    }

                    Total = 0;
                    for (i = 1; i <= DataTab.Rows.Count; i++)
                    {
                        Total += Convert.ToDouble(ExcelSheet.Cells[i + l, 5].Value);
                    }

                    Microsoft.Office.Interop.Excel.Range range = ExcelSheet.Range[$"A{l + 1 + DataTab.Rows.Count}:D{l + 1 + DataTab.Rows.Count}"];

                    // Fusionnez les cellules
                    range.Merge();

                    // Affectez la valeur à la cellule
                    ExcelSheet.Cells[l + 1 + DataTab.Rows.Count, 1].Value = "Total HT en Dirhams :";

                    // Appliquez les formats et styles
                    range.Interior.Color = Color.Gold;
                    range.Font.Bold = true;
                    range.Font.Size = 13;
                    range.Font.Name = "Times New Roman";
                    range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                    range.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    range.Borders.Color = Color.Black;

                    // Affectez la valeur à la cellule suivante
                    ExcelSheet.Cells[l + 1 + DataTab.Rows.Count, 5].Value = Tronquer(Total, 2);

                    // Appliquez les formats et styles à la cellule suivante
                    range = ExcelSheet.Range[$"E{l + 1 + DataTab.Rows.Count}"];
                    range.Interior.Color = Color.Gold;
                    range.Font.Bold = true;
                    range.Font.Size = 13;
                    range.Font.Name = "Times New Roman";
                    range.HorizontalAlignment = XlHAlign.xlHAlignRight;
                    range.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    range.Borders.Color = Color.Black;

                    Excel.Range range2 = ExcelSheet.Range[$"A{l + 2 + DataTab.Rows.Count}:D{l + 2 + DataTab.Rows.Count}"];
                    range2.Merge();
                    // Appliquez les bordures à la plage
                    range2.Borders.Color = Color.Black;
                    // Affectez la valeur à la cellule
                    ExcelSheet.Cells[l + 2 + DataTab.Rows.Count, 1].Value = "MONTANT DE LA REVISION DES PRIX en Dirhams :";

                    // Appliquez les formats et styles
                    range2.Interior.Color = Color.Gold;
                    range2.Font.Bold = true;
                    range2.Font.Size = 13;
                    range2.Font.Name = "Times New Roman";
                    range2.HorizontalAlignment = XlHAlign.xlHAlignRight;
                    range2.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    range2.Borders.Color = Color.Black;


                    DataAdaps = new OleDbDataAdapter("Select Montant, Montant_Révisé From Décompte where Num_Marché='" + CmbNumMarchéM.Text + "'", connections);
                    DataTabs = new DataTable();
                    DataAdaps.Fill(DataTabs);

                    // MontantTTC = DataTabs.Rows.Item(DataTabs.Rows.Count - 1).Item(0)
                    MontantRévisé = 0;
                    for (i = 0; i < DataTabs.Rows.Count; i++)
                    {
                        MontantRévisé += Convert.ToDouble(DataTabs.Rows[i][1]);
                    }

                    ExcelSheet.Cells[l + 2 + DataTab.Rows.Count, 5].Value = Tronquer(MontantRévisé, 2);
                    ExcelSheet.Cells[l + 2 + DataTab.Rows.Count, 5].Interior.Color = Color.Gold;
                    ExcelSheet.Cells[l + 2 + DataTab.Rows.Count, 5].Font.Bold = true;
                    ExcelSheet.Cells[l + 2 + DataTab.Rows.Count, 5].Font.Size = 13;
                    ExcelSheet.Cells[l + 2 + DataTab.Rows.Count, 5].Font.Name = "Times New Roman";
                    ExcelSheet.Cells[l + 2 + DataTab.Rows.Count, 5].HorizontalAlignment = XlHAlign.xlHAlignRight;
                    ExcelSheet.Cells[l + 2 + DataTab.Rows.Count, 5].VerticalAlignment = XlVAlign.xlVAlignCenter;
                    ExcelSheet.Cells[l + 2 + DataTab.Rows.Count, 5].Borders.Color = Color.Black;

                    // Définir la couleur des bordures pour la plage spécifiée
                    Excel.Range bordersRange = ExcelSheet.Range[$"A{l + 3 + DataTab.Rows.Count}:D{l + 3 + DataTab.Rows.Count}"];
                    bordersRange.Borders.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);

                    // Fusionner les cellules de la plage spécifiée
                    bordersRange.Merge();

                    // Définir le contenu et les propriétés de la cellule fusionnée
                    ExcelSheet.Cells[l + 3 + DataTab.Rows.Count, 1].Value = "PENALITE en Dirhams (20%) :";
                    ExcelSheet.Cells[l + 3 + DataTab.Rows.Count, 1].Interior.Color = Color.Gold;
                    ExcelSheet.Cells[l + 3 + DataTab.Rows.Count, 1].Font.Bold = true;
                    ExcelSheet.Cells[l + 3 + DataTab.Rows.Count, 1].Font.Size = 13;
                    ExcelSheet.Cells[l + 3 + DataTab.Rows.Count, 1].Font.Name = "Times New Roman";
                    ExcelSheet.Cells[l + 3 + DataTab.Rows.Count, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                    ExcelSheet.Cells[l + 3 + DataTab.Rows.Count, 1].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    ExcelSheet.Cells[l + 3 + DataTab.Rows.Count, 1].Borders.Color = Color.Black;


                    // CALCUL PENALITE
                    DataAdaps = new OleDbDataAdapter("Select Date_Ordre_service, Dat_reception_Pro From Marché where Réference='" + CmbNumMarchéM.Text + "'", connections);
                    DataTabs = new DataTable();
                    DataAdaps.Fill(DataTabs);

                    Durée = (int)(DateTime.Parse(DataTabs.Rows[0][0].ToString()) - DateTime.Parse(DataTabs.Rows[0][1].ToString())).TotalDays;
                    Pénalité = 0;

                    if (Durée > Délai)
                    {
                        Pénalité = Durée * MontantMarché / 1000;

                        if (Pénalité > MontantMarché * 0.08)
                        {
                            Pénalité = MontantMarché * 0.08;
                        }
                    }

                    int rowCount = l + 3 + DataTab.Rows.Count;

                    for (i = 3; i <= 5; i++)
                    {
                        ExcelSheet.Cells[rowCount, 5].Value = Tronquer((i == 3) ? Pénalité : (Total + MontantRévisé - Pénalité) * (i == 4 ? 0.2 : 1.2), 2);

                        ExcelSheet.Cells[rowCount, 5].Interior.Color = Color.Gold;
                        ExcelSheet.Cells[rowCount, 5].Font.Bold = true;
                        ExcelSheet.Cells[rowCount, 5].Font.Size = 13;
                        ExcelSheet.Cells[rowCount, 5].Font.Name = "Times New Roman";
                        ExcelSheet.Cells[rowCount, 5].HorizontalAlignment = (i == 3) ? Excel.XlHAlign.xlHAlignCenter : Excel.XlHAlign.xlHAlignRight;
                        ExcelSheet.Cells[rowCount, 5].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                        ExcelSheet.Cells[rowCount, 5].Borders.Color = Color.Black;

                        ExcelSheet.Range["A" + rowCount + ":D" + rowCount].Borders.Color = Color.Black;

                        if (i == 4 || i == 5)
                        {
                            ExcelSheet.Range["A" + (rowCount + 1) + ":D" + (rowCount + 1)].Merge();
                            ExcelSheet.Cells[rowCount + 1, 1].Value = (i == 4) ? "TTC en Dirhams (20%)" : "Montant TTC en Dirhams";

                            ExcelSheet.Cells[rowCount + 1, 1].Interior.Color = Color.Gold;
                            ExcelSheet.Cells[rowCount + 1, 1].Font.Bold = true;
                            ExcelSheet.Cells[rowCount + 1, 1].Font.Size = 13;
                            ExcelSheet.Cells[rowCount + 1, 1].Font.Name = "Times New Roman";
                            ExcelSheet.Cells[rowCount + 1, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                            ExcelSheet.Cells[rowCount + 1, 1].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                            ExcelSheet.Cells[rowCount + 1, 1].Borders.Color = Color.Black;

                            rowCount++;
                        }

                        rowCount++;
                    }

                    // ---Arrêté par nous, ordonnateur
                    object cellValue = ExcelSheet.Cells[l + 5 + DataTab.Rows.Count, 5].Value;

                    if (cellValue != null)
                    {
                        str = "Arrêté par nous, ordonnateur, le présent décompte définitif à la somme de : " + Calcul.ChiffreToLettre(cellValue.ToString().Replace(",", "."));
                        ExcelSheet.Cells[l + 7 + DataTab.Rows.Count, 1].Value = str;
                    }
                    else
                    {
                        str = "La valeur de la cellule est null.";
                    }

                    for (int b = 1; b <= 2; b++)
                    {
                        ExcelSheet.Cells[l + 7 + DataTab.Rows.Count, b].Font.Bold = true;
                        ExcelSheet.Cells[l + 7 + DataTab.Rows.Count, b].Font.Color = Color.DarkRed;
                        ExcelSheet.Cells[l + 7 + DataTab.Rows.Count, b].Font.Size = 13;
                        ExcelSheet.Cells[l + 7 + DataTab.Rows.Count, b].Font.Name = "Times New Roman";

                        if (j == 1)
                        {
                            ExcelSheet.Cells[l + 7 + DataTab.Rows.Count, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        }
                        else
                        {
                            ExcelSheet.Cells[l + 7 + DataTab.Rows.Count, 5].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                        }

                        ExcelSheet.Cells[l + 7 + DataTab.Rows.Count, 1].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                        ExcelSheet.Cells[l + 7 + DataTab.Rows.Count, 1].NumberFormat = "@";
                        ExcelSheet.Cells[l + 7 + DataTab.Rows.Count, 1].WrapText = true;
                        ExcelSheet.Cells[l + 7 + DataTab.Rows.Count, 1].Interior.Color = Color.Gold;
                        ExcelSheet.Cells[l + 7 + DataTab.Rows.Count, 1].Borders.Color = Color.Black;
                    }

                    ExcelSheet.Range["A" + (l + 7 + DataTab.Rows.Count) + ":" + "E" + (l + 7 + DataTab.Rows.Count)].Merge();

                }
            }
            else
            {
                MessageBox.Show("Remplir l'attachement.", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Attachement form = new Attachement(); // Remplacez 'Attachement' par le nom de votre formulaire
                form.Show();
                return;
            }
            connections.Close();
            ExcelApp.Visible = true;
            ExcelSheet = null;
            ExcelBook = null;
            ExcelApp = null;
        }
        private double Tronquer(double V, int U)
        {
            string s = V.ToString();
            int i = s.IndexOf(",");
            int j = s.IndexOf(".");
            i = i + j;

            if (s.Length - i <= U || i == 0)
            {
                return Convert.ToDouble(s);
            }
            else
            {
                string D = s.Substring(0, i + U);

                if (double.TryParse(D, out double result))
                {
                    return result;
                }
                else
                {
                    // La conversion a échoué, que voulez-vous faire ici en cas d'erreur ?
                    // Vous pouvez retourner 0, lancer une exception, ou prendre une autre action.
                    // Ici, je retourne 0.
                    return 0;
                }
            }
        }
    }
}
