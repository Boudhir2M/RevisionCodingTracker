using System;
using System.Drawing;
using System.Windows.Forms;

using System.Data;
using System.Data.OleDb;
using Microsoft.Office.Interop.Excel;
using Intuit.Ipp.Utility;
using Excel = Microsoft.Office.Interop.Excel;

namespace Révision
{
    public partial class FrmRévisionTous : Form
    {
        public FrmRévisionTous()
        {
            InitializeComponent();
        }
		private OleDbConnection connection = new OleDbConnection();
		private OleDbCommand records;
		private System.Data.DataTable DataTab;
		private OleDbDataAdapter DataAdap;
		private OleDbCommandBuilder ComndBuld;
		private DataSet DataSetTab = new DataSet();

		private OleDbConnection connections = new OleDbConnection();
		private OleDbCommand recordss;
		private System.Data.DataTable DataTabs;
		private OleDbDataAdapter DataAdaps;
		private OleDbCommandBuilder ComndBulds;
		private DataSet DataSetTabs = new DataSet();

		private OleDbConnection connectionss = new OleDbConnection();
		private OleDbCommand recordsss;
		private System.Data.DataTable DataTabss;
		private OleDbDataAdapter DataAdapss;
		private OleDbCommandBuilder ComndBuldss;
		private DataSet DataSetTabss = new DataSet();

		private string SQLSTR;
		private long NumDécompte1;
		private long NumDécompte2;
		private long NumDécompte;
		private long Indexs;
		private string DateCommencement;
		private double TotalRévisé;
		private long k;
		private string STRF;

        private void CmbNumMarchéM_SelectedIndexChanged(object sender, EventArgs e)
        {
			if (string.IsNullOrEmpty(CmbNumMarchéM.Text))
			{
				MessageBox.Show("Choisissez une valeur de la liste : Réference de Marché ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
				CmbNumMarchéM.Focus();
				return;
			}
			CmbDécompteM.Items.Clear();
			CmbDécompteN.Items.Clear();
			DataAdap = new OleDbDataAdapter("Select Désignation From Décompte where Num_Marché='" + CmbNumMarchéM.Text + "'", connection);
			DataTab = new System.Data.DataTable();
			DataAdap.Fill(DataTab);
			if (DataTab.Rows.Count > 0)
			{
				for (var i = 0; i < DataTab.Rows.Count; i++)
				{
					CmbDécompteM.Items.Add(DataTab.Rows[i][0].ToString());
					CmbDécompteN.Items.Add(DataTab.Rows[i][0].ToString());
				}
			}
			else
			{
				MessageBox.Show("Aucun valeur n'ai saisi dans la table détail éstimatif pour le marché numéro : " + CmbNumMarchéM.Text + " .", "Attention!", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
			DataAdap = new OleDbDataAdapter("Select * From Marché where Réference='" + CmbNumMarchéM.Text + "'", connection);
			DataTab = new System.Data.DataTable();
			DataAdap.Fill(DataTab);
			if (DataTab.Rows.Count > 0)
			{
				TxtObjet.Text = DataTab.Rows[0][1].ToString();
				DateCommencement = DataTab.Rows[0][15].ToString();
			}
			DataAdap = new OleDbDataAdapter("Select Désignation, Date_Ordre_Service, Date_Arret_Service From Journal where Num_Marché='" + CmbNumMarchéM.Text + "'", connection);
			DataTab = new System.Data.DataTable();
			DataAdap.Fill(DataTab);
			DGV.DataSource = DataTab;
			DataAdap = new OleDbDataAdapter("Select Décompte.Désignation,Attachement.Désignation, Unité, Quantité,Prix_unitaire From AttachementParDécompte where Num_Marché='" + CmbNumMarchéM.Text + "'", connection);
			DataTab = new System.Data.DataTable();
			DataAdap.Fill(DataTab);
			DGVP.DataSource = DataTab;
		}

        private void CmbDécompteM_SelectedIndexChanged(object sender, EventArgs e)
        {
			if (string.IsNullOrEmpty(CmbDécompteM.Text))
			{
				MessageBox.Show("Choisissez une valeur de la liste : Réference de Marché ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
				CmbDécompteM.Focus();
				return;
			}
			DataAdap = new OleDbDataAdapter("Select N° From Décompte where Désignation='" + CmbDécompteM.Text + "' and Num_Marché='" + CmbNumMarchéM.Text + "'", connection);
			DataTab = new System.Data.DataTable();
			DataAdap.Fill(DataTab);
			NumDécompte1 = Convert.ToInt64(DataTab.Rows[0][0]);
		}

        private void CmbDécompteN_SelectedIndexChanged(object sender, EventArgs e)
        {
			if (string.IsNullOrEmpty(CmbDécompteM.Text))
			{
				MessageBox.Show("Choisissez une valeur de la liste : Réference de Marché ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
				CmbDécompteM.Focus();
				return;
			}
			DataAdap = new OleDbDataAdapter("Select N° From Décompte where Désignation='" + CmbDécompteM.Text + "' and Num_Marché='" + CmbNumMarchéM.Text + "'", connection);
			DataTab = new System.Data.DataTable();
			DataAdap.Fill(DataTab);
			NumDécompte2 = Convert.ToInt64(DataTab.Rows[0][0]);
		}

        private void FrmRévisionTous_Load(object sender, EventArgs e)
        {
			if (connection.State == ConnectionState.Closed)
			{
				connection.ConnectionString = ("Provider=Microsoft.ACE.OLEDB.12.0;Data source=" + System.Windows.Forms.Application.StartupPath + "\\revision.accdb");
				connection.Open();
			}
			Recharger();
		}

        private void FrmRévisionTous_FormClosed(object sender, FormClosedEventArgs e)
        {
			connection.Close();
		}
		private void Recharger()
		{
			DataAdap = new System.Data.OleDb.OleDbDataAdapter("Select Réference From Marché", connection);
			DataTab = new System.Data.DataTable();
			DataAdap.Fill(DataTab);
			int i = 0;
			for (i = 0; i < DataTab.Rows.Count; i++)
			{
				CmbNumMarchéM.Items.Add(DataTab.Rows[i][0].ToString());
			}
		}
		private void ExecuterSQL(string Strsql)
		{
			OleDbCommand CMND = (OleDbCommand)connection.CreateCommand();
			CMND.CommandText = SQLSTR;
			CMND.ExecuteNonQuery();
		}

        private void BtnEditer_Click(object sender, EventArgs e)
        {
			int i = 0;
			int j = 0;
			int mm = 0;
			int mois = 0;
			int m = 0;
			int y = 0;
			int k = 0;
			long NumSymbole = 0;
			long TM = 0;
			string MoisBase = null;
			long ln = 0;
			int NDC1 = 0;
			int NDC2 = 0;
			double MRT = 0D;
			double MRPD = 0D;
			//----------------- connection à la base des données ------------------------------------------------
			if (connections.State == ConnectionState.Closed)
			{
				connections.ConnectionString = ("Provider=Microsoft.ACE.OLEDB.12.0;Data source=" + System.Windows.Forms.Application.StartupPath + "\\revision.accdb");
				connections.Open();
			}
			if (connectionss.State == ConnectionState.Closed)
			{
				connectionss.ConnectionString = ("Provider=Microsoft.ACE.OLEDB.12.0;Data source=" + System.Windows.Forms.Application.StartupPath + "\\revision.accdb");
				connectionss.Open();
			}

			if (string.IsNullOrEmpty(CmbDécompteM.Text))
			{
				MessageBox.Show("Choisissez une valeur de la liste : Num Décompte Départ ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
				CmbDécompteM.Focus();
				return;
			}
			if (string.IsNullOrEmpty(CmbDécompteN.Text))
			{
				MessageBox.Show("Choisissez une valeur de la liste : Num Décompte Fin ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
				CmbDécompteN.Focus();
				return;
			}
			NDC1 = Convert.ToInt32(CmbDécompteM.Text.Substring(2, CmbDécompteM.Text.Length - 2));
			NDC2 = Convert.ToInt32(CmbDécompteN.Text.Substring(2, CmbDécompteN.Text.Length - 2));
			if (NDC1 > NDC2)
			{
				CmbDécompteM.Text = "DP" + NDC2.ToString();
				CmbDécompteN.Text = "DP" + NDC1.ToString();
				NDC1 = Convert.ToInt32(CmbDécompteM.Text.Substring(2, CmbDécompteM.Text.Length - 2));
				NDC2 = Convert.ToInt32(CmbDécompteN.Text.Substring(2, CmbDécompteN.Text.Length - 2));
			}
			ln = (long)NDC2 - (long)NDC1;
			long[] HL = null;

			for (k = 0; k <= ln; k++)
			{
				DataAdap = new OleDbDataAdapter("Select N° From Décompte where Désignation='DP" + (k + 1).ToString() + "' and Num_Marché='" + CmbNumMarchéM.Text + "'", connection);
				DataTab = new System.Data.DataTable();
				DataAdap.Fill(DataTab);
				NumDécompte = Convert.ToInt64(DataTab.Rows[0][0]);
				DataAdap = new OleDbDataAdapter("Select CD From CompteDate where Num_Décompte=" + NumDécompte, connection);
				DataTab = new System.Data.DataTable();
				DataAdap.Fill(DataTab);
				Array.Resize(ref HL, k + 1);
				HL[k] = Convert.ToInt64(DataTab.Rows[0][0]);
			}
			DataAdaps = new OleDbDataAdapter("Select Epoque_Base From Marché where Réference='" + CmbNumMarchéM.Text + "'", connections);
			DataTabs = new System.Data.DataTable();
			DataAdaps.Fill(DataTabs);
			MoisBase = DataTabs.Rows[0][0].ToString();
			//-------- Dresser la table des révisions pour plusieurs tables ------------------------
			dynamic ExcelApp = null;
			dynamic ExcelBook = null;
			dynamic ExcelSheet = null;
			ExcelApp = System.Activator.CreateInstance(System.Type.GetTypeFromProgID("Excel.Application"));
			ExcelBook = ExcelApp.WorkBooks.Add;
			ExcelSheet = ExcelBook.WorkSheets(1);
			ExcelSheet = ExcelBook.WorkSheets.Add;
			ExcelSheet = ExcelBook.WorkSheets(2);
			ExcelSheet.name = "Récapitulatif";
			ExcelSheet = ExcelBook.WorkSheets(1);
			ExcelSheet.name = "Révision";
			ExcelSheet.Cells[1, 1] = "Royaume du Maroc" + "\r\n" + "Ministère de l'Intérieur" + "\r\n" + "Province d'Al Haouz" + "\r\n" + "Conseil Provincial Al Haouz";
			ExcelSheet.Rows["1:1"].RowHeight = 100;
			ExcelSheet.Cells[1, 1].Font.Bold = true;
			ExcelSheet.Cells[1, 1].Font.Size = 13;
			ExcelSheet.Cells[1, 1].Font.Name = "Times New Roman";
			ExcelSheet.Cells[1, 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
			ExcelSheet.Cells[1, 1].VerticalAlignment = XlVAlign.xlVAlignCenter;
			ExcelSheet.Range["A1:B1"].Merge();

			ExcelSheet.Cells[2, 3] = "     Objet : " + TxtObjet.Text;
			ExcelSheet.Rows["2:2"].RowHeight = 50;
			ExcelSheet.Cells[2, 3].Font.Bold = true;
			ExcelSheet.Cells[2, 3].Font.Size = 14;
			ExcelSheet.Cells[2, 3].Font.Name = "Times New Roman";
			ExcelSheet.Cells[2, 3].HorizontalAlignment = XlHAlign.xlHAlignLeft;
			ExcelSheet.Cells[2, 3].VerticalAlignment = XlVAlign.xlVAlignCenter;

			ExcelSheet.Range["C2:I2"].Merge();
			ExcelSheet.Cells[1, 3] = "Marché n° :" + CmbNumMarchéM.Text;
			ExcelSheet.Cells[1, 3].Font.Bold = true;
			ExcelSheet.Cells[1, 3].Font.Size = 13;
			ExcelSheet.Cells[1, 3].Font.Name = "Times New Roman";
			ExcelSheet.Cells[1, 3].HorizontalAlignment = XlHAlign.xlHAlignRight;
			ExcelSheet.Cells[1, 3].VerticalAlignment = XlVAlign.xlVAlignCenter;
			ExcelSheet.Range["C1:I1"].Merge();

			ExcelSheet.Cells[4, 1] = "ETAT DE LA REVISION DES PRIX";
			ExcelSheet.Cells[4, 1].Interior.Color = Color.SkyBlue;
			ExcelSheet.Cells[4, 1].Font.Bold = true;
			ExcelSheet.Cells[4, 1].Font.Size = 15;
			ExcelSheet.Cells[4, 1].Font.Name = "Times New Roman";
			ExcelSheet.Cells[4, 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
			ExcelSheet.Cells[4, 1].VerticalAlignment = XlVAlign.xlVAlignCenter;
			ExcelSheet.Rows["4:4"].RowHeight = 25;
			ExcelSheet.Range["A4:I4"].Merge();

			ExcelSheet.Cells[6, 1] = "Date d'époque de base : " + MoisBase;
			ExcelSheet.Cells[6, 1].Font.Bold = true;
			ExcelSheet.Cells[6, 1].Font.Size = 13;
			ExcelSheet.Cells[6, 1].Font.Name = "Times New Roman";
			ExcelSheet.Cells[6, 1].HorizontalAlignment = XlHAlign.xlHAlignLeft;
			ExcelSheet.Cells[6, 1].VerticalAlignment = XlVAlign.xlVAlignCenter;

			ExcelSheet.Cells[7, 1] = "Ordre de service de Commencement le : " + DateCommencement;
			ExcelSheet.Cells[7, 1].Font.Bold = true;
			ExcelSheet.Cells[7, 1].Font.Size = 13;
			ExcelSheet.Cells[7, 1].Font.Name = "Times New Roman";
			ExcelSheet.Cells[7, 1].HorizontalAlignment = XlHAlign.xlHAlignLeft;
			ExcelSheet.Cells[7, 1].VerticalAlignment = XlVAlign.xlVAlignCenter;
			ExcelSheet.Rows["8:8"].RowHeight = 25;

			ExcelSheet.Cells[8, 1] = "Les règles et conditions de révision des prix sont celles fixées par l'arrêté du premier ministre n°3-302-15 du 15 safar 1437 (27 novembre 2015)";
			ExcelSheet.Cells[8, 1].Font.Bold = true;
			ExcelSheet.Cells[8, 1].Font.Underline = true;
			ExcelSheet.Cells[8, 1].Font.Size = 13;
			ExcelSheet.Cells[8, 1].Font.Name = "Times New Roman";
			ExcelSheet.Cells[8, 1].HorizontalAlignment = XlHAlign.xlHAlignLeft;
			ExcelSheet.Cells[8, 1].VerticalAlignment = XlVAlign.xlVAlignCenter;

			ExcelSheet.Cells[9, 1].Font.Bold = true;
			ExcelSheet.Cells[9, 1].Font.Size = 13;
			ExcelSheet.Cells[9, 1].Font.Name = "Times New Roman";
			ExcelSheet.Cells[9, 1].HorizontalAlignment = XlHAlign.xlHAlignLeft;
			ExcelSheet.Cells[9, 1].VerticalAlignment = XlVAlign.xlVAlignCenter;

			DataAdaps = new OleDbDataAdapter("Select DISTINCT Symbole From Detail_Estimatif where Num_Marché='" + CmbNumMarchéM.Text + "'", connections);
			DataTabs = new System.Data.DataTable();
			DataAdaps.Fill(DataTabs);

			ExcelSheet.Cells[10, 1] = "Symbole";
			ExcelSheet.Cells[10, 1].Interior.Color = Color.Gold;
			ExcelSheet.Cells[10, 1].Borders.Color = Color.Black;
			ExcelSheet.Cells[10, 1].Font.Bold = true;
			ExcelSheet.Cells[10, 1].Font.Size = 13;
			ExcelSheet.Cells[10, 1].Font.Name = "Times New Roman";
			ExcelSheet.Cells[10, 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;

			ExcelSheet.Cells[10, 2].Interior.Color = Color.Gold;
			ExcelSheet.Cells[10, 2].Borders.Color = Color.Black;
			ExcelSheet.Cells[10, 2].VerticalAlignment = XlVAlign.xlVAlignCenter;
			ExcelSheet.Cells[10, 2] = "Index de Base";
			ExcelSheet.Cells[10, 2].Font.Bold = true;
			ExcelSheet.Cells[10, 2].Font.Size = 13;
			ExcelSheet.Cells[10, 2].Font.Name = "Times New Roman";
			ExcelSheet.Cells[10, 2].HorizontalAlignment = XlHAlign.xlHAlignCenter;
			ExcelSheet.Cells[10, 2].VerticalAlignment = XlVAlign.xlVAlignCenter;

			for (i = 0; i < DataTabs.Rows.Count; i++)
			{
				DataAdap = new OleDbDataAdapter("Select Num From symbole where Code='" + DataTabs.Rows[i][0].ToString() + "'", connection);
				DataTab = new System.Data.DataTable();
				DataAdap.Fill(DataTab);
				NumSymbole = Convert.ToInt64(DataTab.Rows[0][0]);
				//------
				DataAdapss = new OleDbDataAdapter("Select Valeur From Indexs where Mois='" + MoisBase + "' and Symbole=" + NumSymbole, connectionss);
				DataTabss = new System.Data.DataTable();
				DataAdapss.Fill(DataTabss);

				ExcelSheet.Cells[11 + i, 1] = DataTabs.Rows[i][0].ToString();
				ExcelSheet.Cells[11 + i, 1].Font.Bold = true;
				ExcelSheet.Cells[11 + i, 1].Font.Size = 13;
				ExcelSheet.Cells[11 + i, 1].Font.Name = "Times New Roman";
				ExcelSheet.Cells[11 + i, 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
				ExcelSheet.Cells[11 + i, 1].Borders.Color = Color.Black;
				ExcelSheet.Cells[11 + i, 1].VerticalAlignment = XlVAlign.xlVAlignCenter;

				ExcelSheet.Cells[11 + i, 2] = DataTabss.Rows[0][0].ToString();
				ExcelSheet.Cells[11 + i, 2].Font.Size = 13;
				ExcelSheet.Cells[11 + i, 2].Font.Name = "Times New Roman";
				ExcelSheet.Cells[11 + i, 2].HorizontalAlignment = XlHAlign.xlHAlignCenter;
				ExcelSheet.Cells[11 + i, 2].VerticalAlignment = XlVAlign.xlVAlignCenter;
				ExcelSheet.Cells[11 + i, 2].Borders.Color = Color.Black;


			}
			ExcelSheet.Cells[16, 1] = "Symbole";
			ExcelSheet.Cells[16, 1].Interior.Color = Color.Gold;
			ExcelSheet.Cells[16, 1].Borders.Color = Color.Black;
			ExcelSheet.Cells[16, 1].Font.Bold = true;
			ExcelSheet.Cells[16, 1].Font.Size = 13;
			ExcelSheet.Cells[16, 1].Font.Name = "Times New Roman";
			ExcelSheet.Cells[16, 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
			TotalRévisé = 0;

            //-*-*-*-*-***--*-*-**-*-*-*-* calcul des mois *-*-*-*-*-*-*-*-*
            int p = 0;
            int o = 0;
            int g = 0;
            int s = 0;
            string MSTR = null;
            string[] MoisV = null;
            int UD = 0;
            int DD = 0;
            int DIFM = 0;
            if (CmbDécompteN.Text.Length >= 4)
            {
                UD = Convert.ToInt32(CmbDécompteN.Text.Substring(2, 2));
            }
            else
            {
                // Gérer le cas où la chaîne est trop courte
                Console.WriteLine("La chaîne est trop courte pour extraire les caractères nécessaires.");
            }

            for (p = 0; p <= DGV.Rows.Count - 2; p++)
            {
                if (DGV.Rows[p].Cells[0].Value.ToString().Length >= 4)
                {
                    DD = Convert.ToInt32(DGV.Rows[p].Cells[0].Value.ToString().Substring(2, 2));
                    if (DD > UD)
                    {
                        break;
                    }
                }
                else
                {
                    // Gérer le cas où la chaîne est trop courte
                    Console.WriteLine("La chaîne est trop courte pour extraire les caractères nécessaires.");
                }
                if (p == 0)
                {
                    g = 0;
                    Array.Resize(ref MoisV, g + 1);
                    MoisV[g] = Convert.ToString(Convert.ToDateTime(DGV.Rows[p].Cells[1].Value).Month + "/" + Convert.ToDateTime(DGV.Rows[p].Cells[1].Value).Year);
                    if (DateHelper.DateDiff(DateHelper.DateInterval.Month, Convert.ToDateTime(DGV.Rows[p].Cells[2].Value), DateTime.Parse("01/" + MoisV[g])) > 0)
                    {
                        DIFM = (int)DateHelper.DateDiff(DateHelper.DateInterval.Month, Convert.ToDateTime(DGV.Rows[p].Cells[1].Value), Convert.ToDateTime(DGV.Rows[p].Cells[2].Value));
                        s = Convert.ToDateTime(DGV.Rows[p].Cells[1].Value).Month;
                        g = g + DIFM;
                        Array.Resize(ref MoisV, g + 1);
                        for (o = 1; o <= DIFM; o++)
                        {
                            if (s == 12)
                            {
                                s = 1;
                            }
                            else
                            {
                                s = s + o;
                            }
                            MSTR = s.ToString();
                            if (MSTR.Length == 1)
                            {
                                MSTR = "0" + MSTR;
                            }
                            MoisV[g] = MSTR + "/" + Convert.ToString(Convert.ToDateTime(DGV.Rows[p].Cells[2].Value).Year);
                        }
                    }
                }
                else
                {
                    if (DateHelper.DateDiff(DateHelper.DateInterval.Month, DateTime.Parse("01/" + MoisV[g]), Convert.ToDateTime(DGV.Rows[p].Cells[1].Value)) > 0)
                    {
                        g = g + 1;
                        Array.Resize(ref MoisV, g + 1);
                        MoisV[g] = Convert.ToString(Convert.ToDateTime(DGV.Rows[p].Cells[1].Value).Month + "/" + Convert.ToDateTime(DGV.Rows[p].Cells[1].Value).Year);
                        if (DateHelper.DateDiff(DateHelper.DateInterval.Month, DateTime.Parse("01/" + MoisV[g]), Convert.ToDateTime(DGV.Rows[p].Cells[2].Value)) > 0)
                        {
                            DIFM = (int)DateHelper.DateDiff(DateHelper.DateInterval.Month, Convert.ToDateTime(DGV.Rows[p].Cells[1].Value), Convert.ToDateTime(DGV.Rows[p].Cells[2].Value));
                            s = Convert.ToDateTime(DGV.Rows[p].Cells[1].Value).Month;
                            g = g + DIFM;
                            Array.Resize(ref MoisV, g + 1);
                            for (o = 1; o <= DIFM; o++)
                            {
                                if (s == 12)
                                {
                                    s = 1;
                                }
                                else
                                {
                                    s = s + 1;
                                }
                                MSTR = s.ToString();
                                if (MSTR.Length == 1)
                                {
                                    MSTR = "0" + MSTR;
                                }
                                MoisV[g - DIFM + o] = MSTR + "/" + Convert.ToString(Convert.ToDateTime(DGV.Rows[p].Cells[2].Value).Year);
                            }
                        }
                    }
                    else
                    {
                        if (DateHelper.DateDiff(DateHelper.DateInterval.Month, DateTime.Parse("01/" + MoisV[g]), Convert.ToDateTime(DGV.Rows[p].Cells[2].Value)) > 0)
                        {
                            DIFM = (int)DateHelper.DateDiff(DateHelper.DateInterval.Month, Convert.ToDateTime(DGV.Rows[p].Cells[1].Value), Convert.ToDateTime(DGV.Rows[p].Cells[2].Value));
                            s = Convert.ToDateTime(DGV.Rows[p].Cells[1].Value).Month;
                            g = g + DIFM;
                            Array.Resize(ref MoisV, g + 1);
                            for (o = 1; o <= DIFM; o++)
                            {
                                if (s == 12)
                                {
                                    s = 1;
                                }
                                else
                                {
                                    s = s + 1;
                                }
                                MSTR = s.ToString();
                                if (MSTR.Length == 1)
                                {
                                    MSTR = "0" + MSTR;
                                }
                                MoisV[g + o - DIFM] = MSTR + "/" + Convert.ToString(Convert.ToDateTime(DGV.Rows[p].Cells[2].Value).Year);
                            }
                        }
                    }

                }
            }

            //*************************************************************
            for (i = 0; i < DataTabs.Rows.Count; i++)
            {
                DataAdap = new System.Data.OleDb.OleDbDataAdapter("Select Num From symbole where Code='" + DataTabs.Rows[i][0].ToString() + "'", connection);
                DataTab = new System.Data.DataTable();
                DataAdap.Fill(DataTab);
                NumSymbole = Convert.ToInt64(DataTab.Rows[0][0]);

                //------

                ExcelSheet.Cells[17 + i, 1] = DataTabs.Rows[i][0].ToString();
                ExcelSheet.Cells[17 + i, 1].Font.Bold = true;
                ExcelSheet.Cells[17 + i, 1].Font.Size = 13;
                ExcelSheet.Cells[17 + i, 1].Font.Name = "Times New Roman";
                ExcelSheet.Cells[17 + i, 1].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                ExcelSheet.Cells[17 + i, 1].Borders.Color = Color.Black;
                ExcelSheet.Cells[17 + i, 1].VerticalAlignment = XlVAlign.xlVAlignCenter;

                for (j = 0; j <= MoisV.GetUpperBound(0); j++)
                {
                    DataAdapss = new System.Data.OleDb.OleDbDataAdapter("Select Valeur From Indexs where Mois='" + DateTime.Parse("01/" + MoisV[j]) + "' and Symbole=" + NumSymbole, connectionss);
                    DataTabss = new System.Data.DataTable();
                    DataAdapss.Fill(DataTabss);

                    ExcelSheet.Cells[16, 2 + j] = MoisV[j];
                    ExcelSheet.Cells[16, 2 + j].Font.Bold = true;
                    ExcelSheet.Cells[16, 2 + j].Font.Size = 13;
                    ExcelSheet.Cells[16, 2 + j].Font.Name = "Times New Roman";
                    ExcelSheet.Cells[16, 2 + j].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    ExcelSheet.Cells[16, 2 + j].Borders.Color = Color.Black;
                    ExcelSheet.Cells[16, 2 + j].VerticalAlignment = XlVAlign.xlVAlignCenter;

                    if (DataTabss.Rows.Count > 0)
                    {
                        ExcelSheet.Cells[17 + i, 2 + j] = DataTabss.Rows[0][0].ToString();
                    }
                    else
                    {
                        // Gérer le cas où DataTabss n'a pas de lignes à la position 0
                        Console.WriteLine("DataTabss n'a pas de lignes à la position 0.");
                    }

                    ExcelSheet.Cells[17 + i, 2 + j].Font.Size = 13;
                    ExcelSheet.Cells[17 + i, 2 + j].Font.Name = "Times New Roman";
                    ExcelSheet.Cells[17 + i, 2 + j].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    ExcelSheet.Cells[17 + i, 2 + j].VerticalAlignment = XlVAlign.xlVAlignCenter;
                    ExcelSheet.Cells[17 + i, 2 + j].Borders.Color = Color.Black;
                }
            }


            int l = 0;
            l = l + i; // il semble y avoir une confusion ici, car i n'est pas déclaré dans le contexte actuel
            ExcelSheet.Cells[18 + l, 1] = "Décompte";
            ExcelSheet.Cells[18 + l, 1].Interior.Color = Color.Gold;
            ExcelSheet.Cells[18 + l, 1].Borders.Color = Color.Black;
            ExcelSheet.Cells[18 + l, 1].Font.Bold = true;
            ExcelSheet.Cells[18 + l, 1].Font.Size = 13;
            ExcelSheet.Cells[18 + l, 1].Font.Name = "Times New Roman";
            ExcelSheet.Cells[18 + l, 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            ExcelSheet.Cells[18 + l, 1].VerticalAlignment = XlVAlign.xlVAlignCenter;

            ExcelSheet.Cells[18 + l, 2] = "Date Reprise";
            ExcelSheet.Cells[18 + l, 2].Font.Bold = true;
            ExcelSheet.Cells[18 + l, 2].Font.Size = 13;
            ExcelSheet.Cells[18 + l, 2].Font.Name = "Times New Roman";
            ExcelSheet.Cells[18 + l, 2].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            ExcelSheet.Cells[18 + l, 2].Interior.Color = Color.Gold;
            ExcelSheet.Cells[18 + l, 2].Borders.Color = Color.Black;
            ExcelSheet.Cells[18 + l, 2].VerticalAlignment = XlVAlign.xlVAlignCenter;

            ExcelSheet.Cells[18 + l, 3] = "Date d'arrêt";
            ExcelSheet.Cells[18 + l, 3].Interior.Color = Color.Gold;
            ExcelSheet.Cells[18 + l, 3].Borders.Color = Color.Black;
            ExcelSheet.Cells[18 + l, 3].Font.Bold = true;
            ExcelSheet.Cells[18 + l, 3].Font.Size = 13;
            ExcelSheet.Cells[18 + l, 3].Font.Name = "Times New Roman";
            ExcelSheet.Cells[18 + l, 3].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            ExcelSheet.Cells[18 + l, 3].VerticalAlignment = XlVAlign.xlVAlignCenter;

            ExcelSheet.Cells[18 + l, 4] = "Formule de Calcul";
            ExcelSheet.Cells[18 + l, 4].Interior.Color = Color.Gold;
            ExcelSheet.Cells[18 + l, 4].Borders.Color = Color.Black;
            ExcelSheet.Cells[18 + l, 4].Font.Bold = true;
            ExcelSheet.Cells[18 + l, 4].Font.Size = 13;
            ExcelSheet.Cells[18 + l, 4].Font.Name = "Times New Roman";
            ExcelSheet.Cells[18 + l, 4].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            ExcelSheet.Cells[18 + l, 4].VerticalAlignment = XlVAlign.xlVAlignCenter;

            ExcelSheet.Range["D" + (18 + l).ToString() + ":I" + (18 + l).ToString()].Merge();

            for (j = 1; j <= 6; j++)
            {
                ExcelSheet.cells(18 + l, 3 + j).Borders.Color = Color.Black;
            }
            if (string.IsNullOrEmpty(CmbDécompteM.Text))
            {
                MessageBox.Show("Choisissez une valeur de la liste : Num Décompte Départ ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                CmbDécompteM.Focus();
                return;
            }
            if (string.IsNullOrEmpty(CmbDécompteN.Text))
            {
                MessageBox.Show("Choisissez une valeur de la liste : Num Décompte Fin ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                CmbDécompteN.Focus();
                return;
            }
            NDC1 = Convert.ToInt32(CmbDécompteM.Text.Substring(2, CmbDécompteM.Text.Length - 2));
            NDC2 = Convert.ToInt32(CmbDécompteN.Text.Substring(2, CmbDécompteN.Text.Length - 2));
            if (NDC1 > NDC2)
            {
                CmbDécompteM.Text = "DP" + NDC2.ToString();
                CmbDécompteN.Text = "DP" + NDC1.ToString();
                NDC1 = Convert.ToInt32(CmbDécompteM.Text.Substring(2, CmbDécompteM.Text.Length - 2));
                NDC2 = Convert.ToInt32(CmbDécompteN.Text.Substring(2, CmbDécompteN.Text.Length - 2));
            }
            int[,] MC = new int[(NDC2 - NDC1) + 1, 2];
            for (k = NDC1; k <= NDC2; k++)
            {
                DataAdap = new OleDbDataAdapter("Select N° From Décompte where Désignation='DP" + k.ToString() + "' and Num_Marché='" + CmbNumMarchéM.Text + "'", connection);
                DataTab = new System.Data.DataTable();
                DataAdap.Fill(DataTab);
                NumDécompte = Convert.ToInt64(DataTab.Rows[0][0]);
                DataAdap = new OleDbDataAdapter("Select Date_Ordre_Service,Date_Arret_Service From Ordre_Service where Num_Décompte=" + NumDécompte, connection);
                DataTab = new System.Data.DataTable();
                DataAdap.Fill(DataTab);

                for (i = 0; i < DataTab.Rows.Count; i++)
                {
                    if (i == 0)
                    {
                        ExcelSheet.Cells[19 + l + i, 1] = "DP" + k.ToString();
                    }

                    ExcelSheet.Cells[19 + l + i, 1].Font.Size = 13;
                    ExcelSheet.Cells[19 + l + i, 1].Font.Name = "Times New Roman";
                    ExcelSheet.Cells[19 + l + i, 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    ExcelSheet.Cells[19 + l + i, 1].Borders.Color = Color.Black;
                    ExcelSheet.Cells[19 + l + i, 1].VerticalAlignment = XlVAlign.xlVAlignCenter;

                    ExcelSheet.Cells[19 + l + i, 2] = DataTab.Rows[i][0].ToString();
                    ExcelSheet.Cells[19 + l + i, 2].Font.Size = 13;
                    ExcelSheet.Cells[19 + l + i, 2].Font.Name = "Times New Roman";
                    ExcelSheet.Cells[19 + l + i, 2].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    ExcelSheet.Cells[19 + l + i, 2].Borders.Color = Color.Black;
                    ExcelSheet.Cells[19 + l + i, 2].VerticalAlignment = XlVAlign.xlVAlignCenter;

                    ExcelSheet.Cells[19 + l + i, 3] = DataTab.Rows[i][1].ToString();
                    ExcelSheet.Cells[19 + l + i, 3].Font.Size = 13;
                    ExcelSheet.Cells[19 + l + i, 3].Font.Name = "Times New Roman";
                    ExcelSheet.Cells[19 + l + i, 3].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    ExcelSheet.Cells[19 + l + i, 3].VerticalAlignment = XlVAlign.xlVAlignCenter;
                    ExcelSheet.Cells[19 + l + i, 3].Borders.Color = Color.Black;

                    for (int col = 4; col <= 4; col++) // corrected the loop range for column 4
                    {
                        ExcelSheet.Cells[19 + l + i, col].Font.Size = 13;
                        ExcelSheet.Cells[19 + l + i, col].Font.Name = "Times New Roman";
                        ExcelSheet.Cells[19 + l + i, col].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        ExcelSheet.Cells[19 + l + i, col].VerticalAlignment = XlVAlign.xlVAlignCenter;
                        ExcelSheet.Cells[19 + l + i, col].Borders.Color = Color.Black;
                    }
                }

                ExcelSheet.Range["A" + (19 + l).ToString() + ":A" + (18 + l + i).ToString()].Merge();
                MC[k - NDC1, 0] = 19 + l;
                MC[k - NDC1, 1] = 18 + l + i;
                l = l + i;

            }


            int IDécompte = 0;
            int ISDécompte = 0;
            double TT = 0D;
            double PJ = 0D;
            double NJ = 0D;
            double[] TI = null; //Montant de chaque ensemble de prestations pour un même symbole
            long[] NJDP = new long[HL.GetUpperBound(0) + 1];
            int tempVar = HL.GetUpperBound(0);

            for (var h = 0; h <= tempVar; h++)
            {
                if (h == 0)
                {
                    IDécompte = (int)HL[h];
                    ISDécompte = 0;
                }
                else
                {
                    ISDécompte = IDécompte;
                    IDécompte = (int)(IDécompte + HL[h]);
                }
                int nm = (int)DateHelper.DateDiff(DateHelper.DateInterval.Month, Convert.ToDateTime(DGV.Rows[ISDécompte].Cells[1].Value), Convert.ToDateTime(DGV.Rows[IDécompte - 1].Cells[2].Value));

                // ------------------------------------------------------------------------------------------------------------
                j = 0;
                string[,] TAB = new string[nm + 2, 5];
                for (i = ISDécompte; i < IDécompte; i++)
                {
                    mm = (int)DateHelper.DateDiff(DateHelper.DateInterval.Month, Convert.ToDateTime(DGV.Rows[i].Cells[1].Value), Convert.ToDateTime(DGV.Rows[i].Cells[2].Value));
                    if (mm == 0)
                    {
                        if (i == ISDécompte)
                        {
                            TAB[0, 0] = DGV.Rows[i].Cells[1].Value.ToString().Substring(3, 7);
                            TAB[0, 1] = (DateHelper.DateDiff(DateHelper.DateInterval.Day, Convert.ToDateTime(DGV.Rows[ISDécompte].Cells[1].Value), Convert.ToDateTime(DGV.Rows[ISDécompte].Cells[2].Value)) + 1).ToString();
                        }
                        else
                        {
                            if (DateHelper.DateDiff(DateHelper.DateInterval.Month, Convert.ToDateTime(DGV.Rows[i - 1].Cells[2].Value), Convert.ToDateTime(DGV.Rows[i].Cells[1].Value)) == 0)
                            {
                                TAB[j, 1] = TAB[j, 1] + DateHelper.DateDiff(DateHelper.DateInterval.Day, Convert.ToDateTime(DGV.Rows[i].Cells[1].Value), Convert.ToDateTime(DGV.Rows[i].Cells[2].Value)) + 1;
                            }
                            else
                            {
                                j = j + 1;
                                TAB[j, 0] = DGV.Rows[i].Cells[1].Value.ToString().Substring(3, 7);
                                TAB[j, 1] = (DateHelper.DateDiff(DateHelper.DateInterval.Day, Convert.ToDateTime(DGV.Rows[i].Cells[1].Value), Convert.ToDateTime(DGV.Rows[i].Cells[2].Value)) + 1).ToString();
                            }
                        }
                    }
                    else
                    {
                        if (i == ISDécompte)
                        {
                            TAB[0, 0] = DGV.Rows[i].Cells[1].Value.ToString().Substring(3, 7);
                            mois = (int)DateHelper.DateDiff(DateHelper.DateInterval.Month, Convert.ToDateTime(DGV.Rows[i].Cells[1].Value), Convert.ToDateTime(DGV.Rows[i].Cells[2].Value));
                            for (k = 0; k <= mois; k++)
                            {
                                if (k == 0)
                                {
                                    m = Convert.ToDateTime(DGV.Rows[i].Cells[1].Value).Month;
                                    if (m == 12)
                                    {
                                        m = 1;
                                        y = Convert.ToDateTime(DGV.Rows[i].Cells[1].Value).Year + 1;
                                    }
                                    else
                                    {
                                        m = m + 1;
                                        y = Convert.ToDateTime(DGV.Rows[i].Cells[1].Value).Year;
                                    }
                                    TAB[j, 0] = DGV.Rows[i].Cells[1].Value.ToString().Substring(3, 7);
                                    TAB[j, 1] = DateHelper.DateDiff(DateHelper.DateInterval.Day, Convert.ToDateTime(DGV.Rows[i].Cells[1].Value), DateTime.Parse("01/" + m.ToString() + "/" + y.ToString())).ToString();

                                }
                                //ORIGINAL LINE: Case mois
                                else if (k == mois)
                                {
                                    j = j + 1;
                                    if (m == 12)
                                    {
                                        m = 1;
                                        y = Convert.ToDateTime(DGV.Rows[i].Cells[2].Value).Year;
                                        TAB[j, 0] = DGV.Rows[i].Cells[2].Value.ToString().Substring(3, 7);
                                        TAB[j, 1] = (DateHelper.DateDiff(DateHelper.DateInterval.Day, DateTime.Parse("01/" + 12.ToString() + "/" + y.ToString()), Convert.ToDateTime(DGV.Rows[i].Cells[2].Value)) + 1).ToString();
                                    }
                                    else
                                    {
                                        m = m + 1;
                                        TAB[j, 0] = DGV.Rows[i].Cells[2].Value.ToString().Substring(3, 7);
                                        TAB[j, 1] = (DateHelper.DateDiff(DateHelper.DateInterval.Day, DateTime.Parse("01/" + (m - 1).ToString() + "/" + y.ToString()), Convert.ToDateTime(DGV.Rows[i].Cells[2].Value)) + 1).ToString();
                                    }
                                }
                                //ORIGINAL LINE: Case Else
                                else
                                {
                                    j = j + 1;
                                    if (m == 12)
                                    {
                                        m = 1;
                                        y = y + 1;
                                        TAB[j, 0] = "12/" + y.ToString();
                                        TAB[j, 1] = DateHelper.DateDiff(DateHelper.DateInterval.Day, DateTime.Parse("01/" + 12.ToString() + "/" + (y - 1).ToString()), DateTime.Parse("01/01/" + y.ToString())).ToString();
                                    }
                                    else
                                    {
                                        m = m + 1;
                                        TAB[j, 0] = m.ToString() + "/" + y.ToString();
                                        TAB[j, 1] = DateHelper.DateDiff(DateHelper.DateInterval.Day, DateTime.Parse("01/" + (m - 1).ToString() + "/" + y.ToString()), DateTime.Parse("01/" + m.ToString() + "/" + y.ToString())).ToString();
                                    }
                                }
                            }
                        }
                        else
                        {
                            if (DateHelper.DateDiff(DateHelper.DateInterval.Month, Convert.ToDateTime(DGV.Rows[i - 1].Cells[2].Value), Convert.ToDateTime(DGV.Rows[i].Cells[1].Value)) == 0)
                            {
                                mois = (int)DateHelper.DateDiff(DateHelper.DateInterval.Month, Convert.ToDateTime(DGV.Rows[i].Cells[1].Value), Convert.ToDateTime(DGV.Rows[i].Cells[2].Value));
                                for (k = 0; k <= mois; k++)
                                {

                                    //ORIGINAL LINE: Case 0
                                    if (k == 0)
                                    {
                                        m = Convert.ToDateTime(DGV.Rows[i].Cells[1].Value).Month;
                                        if (m == 12)
                                        {
                                            m = 1;
                                            y = Convert.ToDateTime(DGV.Rows[i].Cells[1].Value).Year + 1;
                                        }
                                        else
                                        {
                                            m = m + 1;
                                            y = Convert.ToDateTime(DGV.Rows[i].Cells[1].Value).Year;
                                        }
                                        TAB[j, 1] = TAB[j, 1] + DateHelper.DateDiff(DateHelper.DateInterval.Day, Convert.ToDateTime(DGV.Rows[i].Cells[1].Value), DateTime.Parse("01/" + m.ToString() + "/" + y.ToString()));
                                    }
                                    //ORIGINAL LINE: Case mois
                                    else if (k == mois)
                                    {
                                        j = j + 1;
                                        if (m == 12)
                                        {
                                            y = Convert.ToDateTime(DGV.Rows[i].Cells[2].Value).Year;
                                            if (m != Convert.ToDateTime(DGV.Rows[i].Cells[2].Value).Month)
                                            {
                                                m = 1;
                                                y = y - 1;
                                            }
                                            TAB[j, 0] = DGV.Rows[i].Cells[2].Value.ToString().Substring(3, 7);
                                            TAB[j, 1] = (DateHelper.DateDiff(DateHelper.DateInterval.Day, DateTime.Parse("01/" + m.ToString() + "/" + y.ToString()), Convert.ToDateTime(DGV.Rows[i].Cells[2].Value)) + 1).ToString();
                                        }
                                        else
                                        {
                                            m = m + 1;
                                            TAB[j, 0] = DGV.Rows[i].Cells[2].Value.ToString().Substring(3, 7);
                                            TAB[j, 1] = (DateHelper.DateDiff(DateHelper.DateInterval.Day, DateTime.Parse("01/" + (m - 1).ToString() + "/" + y.ToString()), Convert.ToDateTime(DGV.Rows[i].Cells[2].Value)) + 1).ToString();
                                        }
                                    }
                                    //ORIGINAL LINE: Case Else
                                    else
                                    {
                                        j = j + 1;
                                        if (m == 12)
                                        {
                                            m = 1;
                                            y = y + 1;
                                            TAB[j, 0] = m.ToString() + "/" + y.ToString();
                                            TAB[j, 1] = DateHelper.DateDiff(DateHelper.DateInterval.Day, DateTime.Parse("01/" + (m - 1).ToString() + "/" + (y - 1).ToString()), DateTime.Parse("01/01/" + y.ToString())).ToString();
                                        }
                                        else
                                        {
                                            m = m + 1;
                                            TAB[j, 0] = (m - 1).ToString() + "/" + y.ToString();
                                            TAB[j, 1] = DateHelper.DateDiff(DateHelper.DateInterval.Day, DateTime.Parse("01/" + (m - 1).ToString() + "/" + y.ToString()), DateTime.Parse("01/" + m.ToString() + "/" + y.ToString())).ToString();
                                        }
                                    }
                                }
                            }
                            else
                            {
                                mois = (int)DateHelper.DateDiff(DateHelper.DateInterval.Month, Convert.ToDateTime(DGV.Rows[i].Cells[1].Value), Convert.ToDateTime(DGV.Rows[i].Cells[2].Value));
                                for (k = 0; k <= mois; k++)
                                {
                                    j = j + 1;

                                    //ORIGINAL LINE: Case 0
                                    if (k == 0)
                                    {
                                        m = Convert.ToDateTime(DGV.Rows[i].Cells[1].Value).Month;
                                        if (m == 12)
                                        {
                                            m = 1;
                                            y = Convert.ToDateTime(DGV.Rows[i].Cells[1].Value).Year + 1;
                                        }
                                        else
                                        {
                                            m = m + 1;
                                            y = Convert.ToDateTime(DGV.Rows[i].Cells[1].Value).Year;
                                        }
                                        TAB[j, 0] = DGV.Rows[i].Cells[1].Value.ToString().Substring(3, 7);
                                        TAB[j, 1] = DateHelper.DateDiff(DateHelper.DateInterval.Day, Convert.ToDateTime(DGV.Rows[i].Cells[1].Value), DateTime.Parse("01/" + m.ToString() + "/" + y.ToString())).ToString();
                                    }
                                    //ORIGINAL LINE: Case mois
                                    else if (k == mois)
                                    {
                                        if (m == 12)
                                        {
                                            m = 1;
                                            y = Convert.ToDateTime(DGV.Rows[i].Cells[2].Value).Year;
                                            TAB[j, 0] = DGV.Rows[i].Cells[2].Value.ToString().Substring(3, 7);
                                            TAB[j, 1] = (DateHelper.DateDiff(DateHelper.DateInterval.Day, DateTime.Parse("01/" + (m - 1).ToString() + "/" + (y - 1).ToString()), Convert.ToDateTime(DGV.Rows[i].Cells[2].Value)) + 1).ToString();
                                        }
                                        else
                                        {
                                            m = m + 1;
                                            TAB[j, 0] = DGV.Rows[i].Cells[2].Value.ToString().Substring(3, 7);
                                            TAB[j, 1] = (DateHelper.DateDiff(DateHelper.DateInterval.Day, DateTime.Parse("01/" + (m - 1).ToString() + "/" + y.ToString()), Convert.ToDateTime(DGV.Rows[i].Cells[2].Value)) + 1).ToString();
                                        }
                                    }
                                    //ORIGINAL LINE: Case Else
                                    else
                                    {
                                        if (m == 12)
                                        {
                                            m = 1;
                                            y = y + 1;
                                            TAB[j, 0] = "01/" + y.ToString();
                                            TAB[j, 1] = DateHelper.DateDiff(DateHelper.DateInterval.Day, DateTime.Parse("01/12" + "/" + (y - 1).ToString()), DateTime.Parse("01/01" + "/" + y.ToString())).ToString();
                                        }
                                        else
                                        {
                                            m = m + 1;
                                            TAB[j, 0] = (m - 1).ToString() + "/" + y.ToString();
                                            TAB[j, 1] = DateHelper.DateDiff(DateHelper.DateInterval.Day, DateTime.Parse("01/" + (m - 1).ToString() + "/" + y.ToString()), DateTime.Parse("01/" + m.ToString() + "/" + y.ToString())).ToString();
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                NJ = 0;
                for (i = 0; i <= nm; i++)
                {
                    // Assurez-vous que TAB[i, 1] n'est pas null ou vide avant de le convertir
                    if (!string.IsNullOrEmpty(TAB[i, 1]))
                    {
                        NJ += long.Parse(TAB[i, 1]);
                    }
                }

                NJDP[h] = Convert.ToInt64(NJ);
                double[] PRC = null;
                DataAdaps = new OleDbDataAdapter("Select DISTINCT Symbole From Detail_Estimatif where Num_Marché='" + CmbNumMarchéM.Text + "'", connections);
                DataTabs = new System.Data.DataTable();
                DataAdaps.Fill(DataTabs);
                PRC = new double[DataTabs.Rows.Count + 1];
                TI = new double[DataTabs.Rows.Count + 1];
                TT = 0;
                for (i = 0; i < DataTabs.Rows.Count; i++)
                {
                    DataAdap = new OleDbDataAdapter("Select N° From Décompte where Désignation='DP" + Convert.ToString(h + 1) + "' and Num_Marché='" + CmbNumMarchéM.Text + "'", connection);
                    DataTab = new System.Data.DataTable();
                    DataAdap.Fill(DataTab);
                    NumDécompte = Convert.ToInt64(DataTab.Rows[0][0]);
                    DataAdap = new OleDbDataAdapter("Select Quantité,Prix_Unitaire From Attachement where Num_Décompte=" + NumDécompte + " and Symbole='" + DataTabs.Rows[i][0].ToString() + "' order by Désignation", connection);
                    DataTab = new System.Data.DataTable();
                    DataAdap.Fill(DataTab);
                    if (h > 0)
                    {
                        DataAdapss = new OleDbDataAdapter("Select N° From Décompte where Désignation='DP" + Convert.ToString(h) + "' and Num_Marché='" + CmbNumMarchéM.Text + "'", connectionss);
                        DataTabss = new System.Data.DataTable();
                        DataAdapss.Fill(DataTabss);
                        NumDécompte = Convert.ToInt64(DataTabss.Rows[0][0]);
                        DataAdapss = new OleDbDataAdapter("Select Quantité,Prix_Unitaire From Attachement where Num_Décompte=" + NumDécompte + " and Symbole='" + DataTabs.Rows[i][0].ToString() + "' order by Désignation", connectionss);
                        DataTabss = new System.Data.DataTable();
                        DataAdapss.Fill(DataTabss);
                    }

                    //-**-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*

                    for (i = 0; i < NDC1; i++)
                    {
                        TI[i] = 0;

                        for (j = 0; j < DataTab.Rows.Count; j++)
                        {
                            if (h == 0)
                            {
                                TI[i] += Convert.ToDouble(DataTab.Rows[j][0]) * Convert.ToDouble(DataTab.Rows[j][1]);
                            }
                            else
                            {
                                TI[i] += (Convert.ToDouble(DataTab.Rows[j][0]) - Convert.ToDouble(DataTabss.Rows[j][0])) * Convert.ToDouble(DataTab.Rows[j][1]);
                            }
                        }

                        TT += TI[i];
                    }

                    PJ = TT / NJ;

                    //INSTANT C# NOTE: The ending condition of VB 'For' loops is tested only on entry to the loop. Instant C# has created a temporary variable in order to use the initial value of nm + 1 for every iteration:
                    int tempVar2 = nm + 1;
                    for (i = 0; i <= tempVar2; i++)
                    {
                        if (TAB[i, 1] != null)
                        {
                            TAB[i, 2] = (long.Parse(TAB[i, 1]) * PJ).ToString();
                        }
                        else
                        {
                            // Gérez le cas où TAB[i, 1] est null, par exemple en affectant une valeur par défaut
                            TAB[i, 2] = "0";
                        }
                    }
                    // Deuxième boucle
                    for (i = 0; i <= tempVar2; i++)
                    {
                        if (TAB[i, 1] != null)
                        {
                            // Utilisez TAB[i, 2] pour d'autres opérations si nécessaire
                        }
                        else
                        {
                            // Gérez le cas où TAB[i, 1] est null
                        }
                    }

                    // Calcul des coefficients a, b, c, ...
                    for (i = 0; i < DataTabs.Rows.Count; i++)
                    {
                        PRC[i] = Tronquer(TI[i] / TT, 4);

                        if (i == 0)
                        {
                            STRF = (PRC[i] == 0)
                                ? "P=P0*(0.15 + 0.85*("
                                : "P=P0*(0.15 + 0.85*(" + Tronquer(PRC[i], 4).ToString() + "*(" + DataTabs.Rows[i][0].ToString() + "/" + DataTabs.Rows[i][0].ToString() + "0)";
                        }
                        else
                        {
                            if (PRC[i] != 0)
                            {
                                STRF += " + " + Tronquer(PRC[i], 4).ToString() + "*(" + DataTabs.Rows[i][0].ToString() + "/" + DataTabs.Rows[i][0].ToString() + "0)";
                            }
                        }
                    }

                    ExcelSheet.Cells[MC[h, 0], 4].Value = STRF;
                    ExcelSheet.Cells[MC[h, 0], 4].Font.Size = 13;
                    ExcelSheet.Cells[MC[h, 0], 4].Font.Name = "Times New Roman";
                    ExcelSheet.Cells[MC[h, 0], 4].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    ExcelSheet.Cells[MC[h, 0], 4].Borders.Color = Color.Black;
                    ExcelSheet.Cells[MC[h, 0], 4].VerticalAlignment = XlVAlign.xlVAlignCenter;

                    for (i = 1; i <= 6; i++)
                    {
                        int tempVar3 = MC[h, 1];
                        for (j = MC[h, 0]; j <= tempVar3; j++)
                        {
                            ExcelSheet.cells(j, 3 + i).Borders.Color = Color.Black;
                        }
                    }
                    ExcelSheet.Range("D" + MC[h, 0].ToString() + ":I" + MC[h, 1].ToString()).Merge();

                    MRPD = 0;

                    if (h == 0)
                    {
                        l = l + i;
                        ExcelSheet.Cells[20 + l, 1] = "Décompte";
                        ExcelSheet.Columns[1].ColumnWidth = ((string)(ExcelSheet.Cells[20 + l, 1].Value)).Length * 1.3;
                        ExcelSheet.Cells[20 + l, 1].Interior.Color = Color.Gold;
                        ExcelSheet.Cells[20 + l, 1].Borders.Color = Color.Black;
                        ExcelSheet.Cells[20 + l, 1].Font.Bold = true;
                        ExcelSheet.Cells[20 + l, 1].Font.Size = 13;
                        ExcelSheet.Cells[20 + l, 1].Font.Name = "Times New Roman";
                        ExcelSheet.Cells[20 + l, 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;

                        ExcelSheet.Cells[20 + l, 2] = "Date de réalisation";
                        ExcelSheet.Columns[2].ColumnWidth = ((string)(ExcelSheet.Cells[20 + l, 2].Value)).Length * 1.3;
                        ExcelSheet.Cells[20 + l, 2].Interior.Color = Color.Gold;
                        ExcelSheet.Cells[20 + l, 2].Borders.Color = Color.Black;
                        ExcelSheet.Cells[20 + l, 2].Font.Bold = true;
                        ExcelSheet.Cells[20 + l, 2].Font.Size = 13;
                        ExcelSheet.Cells[20 + l, 2].Font.Name = "Times New Roman";
                        ExcelSheet.Cells[20 + l, 2].HorizontalAlignment = XlHAlign.xlHAlignCenter;

                        ExcelSheet.Cells[20 + l, 3] = "Montant révisable";
                        ExcelSheet.Columns[3].ColumnWidth = ((string)(ExcelSheet.Cells[20 + l, 3].Value)).Length * 1.3;
                        ExcelSheet.Cells[20 + l, 3].Font.Bold = true;
                        ExcelSheet.Cells[20 + l, 3].Font.Size = 13;
                        ExcelSheet.Cells[20 + l, 3].Font.Name = "Times New Roman";
                        ExcelSheet.Cells[20 + l, 3].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        ExcelSheet.Cells[20 + l, 3].VerticalAlignment = XlVAlign.xlVAlignCenter;


                        ExcelSheet.Cells[20 + l, 4] = "Montant à réviser";
                        ExcelSheet.Columns[4].ColumnWidth = ((string)(ExcelSheet.Cells[20 + l, 4].Value)).Length * 1.3;
                        ExcelSheet.Cells[20 + l, 4].Interior.Color = Color.Gold;
                        ExcelSheet.Cells[20 + l, 4].Borders.Color = Color.Black;
                        ExcelSheet.Cells[20 + l, 4].Font.Bold = true;
                        ExcelSheet.Cells[20 + l, 4].Font.Size = 13;
                        ExcelSheet.Cells[20 + l, 4].Font.Name = "Times New Roman";
                        ExcelSheet.Cells[20 + l, 4].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        ExcelSheet.Cells[20 + l, 4].VerticalAlignment = XlVAlign.xlVAlignCenter;

                        ExcelSheet.Cells[20 + l, 5] = "Période partielle";
                        ExcelSheet.Columns[5].ColumnWidth = ((string)(ExcelSheet.Cells[20 + l, 5].Value)).Length * 1.3;
                        ExcelSheet.Cells[20 + l, 5].Font.Bold = true;
                        ExcelSheet.Cells[20 + l, 5].Font.Size = 13;
                        ExcelSheet.Cells[20 + l, 5].Font.Name = "Times New Roman";
                        ExcelSheet.Cells[20 + l, 5].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        ExcelSheet.Cells[20 + l, 5].VerticalAlignment = XlVAlign.xlVAlignCenter;


                        ExcelSheet.Cells[20 + l, 6] = "Mois des travaux";
                        ExcelSheet.Columns[6].ColumnWidth = ((string)(ExcelSheet.Cells[20 + l, 6].Value)).Length * 1.3;
                        ExcelSheet.Cells[20 + l, 6].Interior.Color = Color.Gold;
                        ExcelSheet.Cells[20 + l, 6].Borders.Color = Color.Black;
                        ExcelSheet.Cells[20 + l, 6].Font.Bold = true;
                        ExcelSheet.Cells[20 + l, 6].Font.Size = 13;
                        ExcelSheet.Cells[20 + l, 6].Font.Name = "Times New Roman";
                        ExcelSheet.Cells[20 + l, 6].HorizontalAlignment = XlHAlign.xlHAlignCenter;

                        ExcelSheet.Cells[20 + l, 7] = "JoursMois/TotalJours";
                        ExcelSheet.Columns[7].ColumnWidth = ((string)(ExcelSheet.Cells[20 + l, 7].Value)).Length * 1.3;
                        ExcelSheet.Cells[20 + l, 7].Interior.Color = Color.Gold;
                        ExcelSheet.Cells[20 + l, 7].Borders.Color = Color.Black;
                        ExcelSheet.Cells[20 + l, 7].Font.Bold = true;
                        ExcelSheet.Cells[20 + l, 7].Font.Size = 13;
                        ExcelSheet.Cells[20 + l, 7].Font.Name = "Times New Roman";
                        ExcelSheet.Cells[20 + l, 7].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        ExcelSheet.Cells[20 + l, 7].VerticalAlignment = XlVAlign.xlVAlignCenter;


                        ExcelSheet.Cells[20 + l, 8] = "P/P0";
                        ExcelSheet.Columns[8].ColumnWidth = ((string)(ExcelSheet.Cells[20 + l, 8].Value)).Length * 1.3;
                        ExcelSheet.Cells[20 + l, 8].Interior.Color = Color.Gold;
                        ExcelSheet.Cells[20 + l, 8].Borders.Color = Color.Black;
                        ExcelSheet.Cells[20 + l, 8].Font.Bold = true;
                        ExcelSheet.Cells[20 + l, 8].Font.Size = 13;
                        ExcelSheet.Cells[20 + l, 8].Font.Name = "Times New Roman";
                        ExcelSheet.Cells[20 + l, 8].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        ExcelSheet.Cells[20 + l, 8].VerticalAlignment = XlVAlign.xlVAlignCenter;

                        ExcelSheet.Cells[20 + l, 9] = "Montant de la révision";
                        ExcelSheet.Columns[9].ColumnWidth = ((string)(ExcelSheet.Cells[20 + l, 9].Value)).Length * 1.3;
                        ExcelSheet.Cells[20 + l, 9].Interior.Color = Color.Gold;
                        ExcelSheet.Cells[20 + l, 9].Borders.Color = Color.Black;
                        ExcelSheet.Cells[20 + l, 9].Font.Bold = true;
                        ExcelSheet.Cells[20 + l, 9].Font.Size = 13;
                        ExcelSheet.Cells[20 + l, 9].Font.Name = "Times New Roman";
                        ExcelSheet.Cells[20 + l, 9].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        ExcelSheet.Cells[20 + l, 9].VerticalAlignment = XlVAlign.xlVAlignCenter;


                        i = 0;
                        TM = 0;
                        MRT = 0;
                        while (!string.IsNullOrEmpty(TAB[i, 0]))
                        {
                            TM = TM + 1;
                            MRT = MRT + double.Parse(TAB[i, 2]);
                            i += 1;
                        }

                        TT = 0;
                        for (i = 0; i < TM; i++)
                        {

                            DataAdaps = new OleDbDataAdapter("Select DISTINCT Symbole From Attachement where Num_Décompte=" + NumDécompte, connections);
                            DataTabs = new System.Data.DataTable();
                            DataAdaps.Fill(DataTabs);

                            for (j = 0; j < DataTabs.Rows.Count; j++)
                            {
                                DataAdap = new OleDbDataAdapter("Select Num From symbole where Code='" + DataTabs.Rows[j][0].ToString() + "'", connection);
                                DataTab = new System.Data.DataTable();
                                DataAdap.Fill(DataTab);
                                NumSymbole = Convert.ToInt64(DataTab.Rows[0][0]);

                                DataAdap = new OleDbDataAdapter("Select Valeur From Indexs where Mois='" + DateTime.Parse("01/" + TAB[i, 0]) + "' and Symbole=" + NumSymbole, connection);
                                DataTab = new System.Data.DataTable();
                                DataAdap.Fill(DataTab);

                                //------
                                DataAdapss = new OleDbDataAdapter("Select Valeur From Indexs where Mois='" + MoisBase + "' and Symbole=" + NumSymbole, connectionss);
                                DataTabss = new System.Data.DataTable();
                                DataAdapss.Fill(DataTabss);

                                if (DataTabss.Rows.Count > 0 && DataTabss.Rows[0][0] != null)
                                {
                                    double valeur = Convert.ToDouble(DataTabss.Rows[0][0]); // Utilisez DataTabss au lieu de DataTab
                                    double prc = PRC[j];
                                    double valeurBase = Convert.ToDouble(DataTabss.Rows[0][0]); // Utilisez DataTabss au lieu de DataTab

                                    double calcul = valeur * prc / valeurBase;
                                    TAB[i, 3] = TAB[i, 3] + Tronquer(calcul, 4);
                                }
                                else
                                {
                                    // Gérez le cas où DataTabss.Rows n'a pas de lignes ou la valeur à la position 0 est null
                                    MessageBox.Show("DataTabss n'a pas de lignes ou la valeur à la position 0 est null.");
                                }
                            }

                            double calculFinal = (double.Parse(TAB[i, 3]) * 0.85 + 0.15) * double.Parse(TAB[i, 2]);

                            TAB[i, 3] = Tronquer(calculFinal, 4).ToString();
                            TT += double.Parse(TAB[i, 3]);
                            double value;
                            if (double.TryParse(TAB[i, 3], out value))
                            {
                                TT += value;
                            }
                            else
                            {
                                // Gérez le cas où la valeur de TAB[i, 3] n'est pas un nombre valide
                                MessageBox.Show("La valeur de TAB[" + i + ", 3] n'est pas un nombre valide.");
                            }
                        }



                        for (i = 0; i < TM; i++)
                        {

                            if (i == 0)
                            {
                                DataAdap = new System.Data.OleDb.OleDbDataAdapter("Select Date_Fin From Décompte where Num_Marché='" + CmbNumMarchéM.Text + "' and Désignation='DP" + Convert.ToString(h + 1) + "'", connection);
                                DataTab = new System.Data.DataTable();
                                DataAdap.Fill(DataTab);

                                ExcelSheet.Cells[21 + l + i, 1] = "DP" + Convert.ToString(h + 1);
                                ExcelSheet.Cells[21 + l + i, 1].Borders.Color = Color.Black;
                                ExcelSheet.Cells[21 + l + i, 1].Font.Size = 13;
                                ExcelSheet.Cells[21 + l + i, 1].Font.Name = "Times New Roman";
                                ExcelSheet.Cells[21 + l + i, 1].HorizontalAlignment = 3;
                                ExcelSheet.Cells[21 + l + i, 1].Borders.Color = Color.Black;
                                ExcelSheet.Cells[21 + l + i, 1].VerticalAlignment = HorizontalAlignment.Center;

                                //----------------------------------------------------------------
                                ExcelSheet.Range["A" + (21 + l + i).ToString()].NumberFormat = "@";
                                ExcelSheet.Cells[21 + l + i, 2] = DataTab.Rows.Count > 0 ? DataTab.Rows[0][0].ToString() : string.Empty;
                                ExcelSheet.Cells[21 + l + i, 2].Borders.Color = Color.Black;
                                ExcelSheet.Cells[21 + l + i, 2].Font.Size = 13;
                                ExcelSheet.Cells[21 + l + i, 2].Font.Name = "Times New Roman";
                                ExcelSheet.Cells[21 + l + i, 2].HorizontalAlignment = 3;
                                ExcelSheet.Cells[21 + l + i, 2].Borders.Color = Color.Black;
                                ExcelSheet.Cells[21 + l + i, 2].VerticalAlignment = HorizontalAlignment.Center;
                                ExcelSheet.Cells[21 + l + i, 3].Borders.Color = Color.Black;

                                //-------------------
                                MRT = 0.0; // Remplacez cette valeur par votre propre calcul de MRT
                                ExcelSheet.Cells[21 + l + i, 3] = Tronquer(MRT, 2);
                                ExcelSheet.Cells[21 + l + i, 3].Font.Size = 13;
                                ExcelSheet.Cells[21 + l + i, 3].Font.Name = "Times New Roman";
                                ExcelSheet.Cells[21 + l + i, 3].HorizontalAlignment = 3;
                                ExcelSheet.Cells[21 + l + i, 3].VerticalAlignment = HorizontalAlignment.Center;

                            }
                            else
                            {
                                ExcelSheet.cells(21 + l + i, 1).Borders.Color = Color.Black;
                                ExcelSheet.cells(21 + l + i, 2).Borders.Color = Color.Black;
                                ExcelSheet.cells(21 + l + i, 3).Borders.Color = Color.Black;
                            }

                            ExcelSheet.Cells[21 + l + i, 4] = Tronquer(double.Parse(TAB[i, 2]), 2);

                            ExcelSheet.Cells[21 + l + i, 4].Borders.Color = Color.Black;
                            ExcelSheet.Cells[21 + l + i, 4].Font.Size = 13;
                            ExcelSheet.Cells[21 + l + i, 4].Font.Name = "Times New Roman";
                            ExcelSheet.Cells[21 + l + i, 4].HorizontalAlignment = 3;

                            ExcelSheet.Cells[21 + l + i, 5].Borders.Color = Color.Black;
                            ExcelSheet.Cells[21 + l + i, 5].VerticalAlignment = HorizontalAlignment.Center;
                            ExcelSheet.Cells[21 + l + i, 5] = TAB[i, 1];
                            ExcelSheet.Cells[21 + l + i, 5].Font.Size = 13;
                            ExcelSheet.Cells[21 + l + i, 5].Font.Name = "Times New Roman";
                            ExcelSheet.Cells[21 + l + i, 5].HorizontalAlignment = 3;
                            ExcelSheet.Cells[21 + l + i, 5].VerticalAlignment = HorizontalAlignment.Center;

                            ExcelSheet.Cells[21 + l + i, 6] = TAB[i, 0];

                            ExcelSheet.Cells[21 + l + i, 6].Borders.Color = Color.Black;
                            ExcelSheet.Cells[21 + l + i, 6].Font.Size = 13;
                            ExcelSheet.Cells[21 + l + i, 6].Font.Name = "Times New Roman";
                            ExcelSheet.Cells[21 + l + i, 6].HorizontalAlignment = 3;

                            ExcelSheet.Cells[21 + l + i, 7].Borders.Color = Color.Black;
                            ExcelSheet.Cells[21 + l + i, 7].VerticalAlignment = HorizontalAlignment.Center;
                            ExcelSheet.Cells[21 + l + i, 7] = Tronquer(long.Parse(TAB[i, 1]) / (double)NJDP[h], 4);
                            ExcelSheet.Cells[21 + l + i, 7].Font.Size = 13;
                            ExcelSheet.Cells[21 + l + i, 7].Font.Name = "Times New Roman";
                            ExcelSheet.Cells[21 + l + i, 7].HorizontalAlignment = 3;
                            ExcelSheet.Cells[21 + l + i, 7].VerticalAlignment = HorizontalAlignment.Center;

                            ExcelSheet.Cells[21 + l + i, 8] = Tronquer((double)(Convert.ToSingle(TAB[i, 3]) / double.Parse(TAB[i, 2])), 4);

                            if (ExcelSheet.Columns[8].ColumnWidth < (ExcelSheet.Cells[21 + l + i, 8].Value.ToString().Length * 1.3))
                            {
                                ExcelSheet.Columns[8].ColumnWidth = ExcelSheet.Cells[21 + l + i, 8].Value.ToString().Length * 1.3;
                            }

                            ExcelSheet.Cells[21 + l + i, 8].Borders.Color = Color.Black;
                            ExcelSheet.Cells[21 + l + i, 8].Font.Size = 13;
                            ExcelSheet.Cells[21 + l + i, 8].Font.Name = "Times New Roman";
                            ExcelSheet.Cells[21 + l + i, 8].HorizontalAlignment = 3;

                            ExcelSheet.Cells[21 + l + i, 9].Borders.Color = Color.Black;
                            ExcelSheet.Cells[21 + l + i, 9].VerticalAlignment = HorizontalAlignment.Center;

                            TotalRévisé = TotalRévisé + Tronquer(MRT * (double)ExcelSheet.Cells[21 + l + i, 7].Value * (-1 + (double)ExcelSheet.Cells[21 + l + i, 8].Value), 2);
                            MRPD = MRPD + Tronquer(MRT * (double)ExcelSheet.Cells[21 + l + i, 7].Value * (-1 + (double)ExcelSheet.Cells[21 + l + i, 8].Value), 2);
                            ExcelSheet.Cells[21 + l + i, 9] = Tronquer(MRT * (double)ExcelSheet.Cells[21 + l + i, 7].Value * (-1 + (double)ExcelSheet.Cells[21 + l + i, 8].Value), 2);

                            ExcelSheet.Cells[21 + l + i, 9].Font.Size = 13;
                            ExcelSheet.Cells[21 + l + i, 9].Font.Name = "Times New Roman";
                            ExcelSheet.Cells[21 + l + i, 9].HorizontalAlignment = 3;
                            ExcelSheet.Cells[21 + l + i, 9].VerticalAlignment = HorizontalAlignment.Center;

                        }
                        ExcelSheet.Range("A" + (21 + l).ToString() + ":A" + (21 + l + i - 1).ToString()).Merge();
                        ExcelSheet.Range("B" + (21 + l).ToString() + ":B" + (21 + l + i - 1).ToString()).Merge();
                        ExcelSheet.Range("C" + (21 + l).ToString() + ":C" + (21 + l + i - 1).ToString()).Merge();

                        //*********** save révision ********************************
                        SQLSTR = "UPDATE Décompte SET ";
                        SQLSTR = SQLSTR + "Montant_Révisé=" + MRPD.ToString().Replace(",", ".");
                        SQLSTR = SQLSTR + " where  Num_Marché='" + CmbNumMarchéM.Text + "' and N°=" + NumDécompte;
                        ExecuterSQL(SQLSTR);
                        //*******************************************************************
                        l = l + i;
                    }
                    else
                    {
                        i = 0;
                        TM = 0;
                        MRT = 0;

                        while (!string.IsNullOrEmpty(TAB[i, 0]))
                        {
                            TM = TM + 1;
                            MRT = MRT + double.Parse(TAB[i, 2]);
                            i += 1;
                        }

                        TT = 0;

                        for (i = 0; i < TM; i++)
                        {
                            DataAdaps = new System.Data.OleDb.OleDbDataAdapter("Select DISTINCT Symbole From Attachement where Num_Décompte=" + NumDécompte, connections);
                            DataTabs = new System.Data.DataTable();
                            DataAdaps.Fill(DataTabs);

                            for (j = 0; j < DataTabs.Rows.Count; j++)
                            {
                                DataAdap = new System.Data.OleDb.OleDbDataAdapter("Select Num From symbole where Code='" + DataTabs.Rows[j][0].ToString() + "'", connection);
                                DataTab = new System.Data.DataTable();
                                DataAdap.Fill(DataTab);
                                NumSymbole = Convert.ToInt64(DataTab.Rows[0][0]);

                                DataAdap = new System.Data.OleDb.OleDbDataAdapter("Select Valeur From Indexs where Mois='" + DateTime.Parse("01/" + TAB[i, 0]) + "' and Symbole=" + NumSymbole, connection);
                                DataTab = new System.Data.DataTable();
                                DataAdap.Fill(DataTab);

                                DataAdapss = new System.Data.OleDb.OleDbDataAdapter("Select Valeur From Indexs where Mois='" + MoisBase + "' and Symbole=" + NumSymbole, connectionss);
                                DataTabss = new System.Data.DataTable();
                                DataAdapss.Fill(DataTabss);

                                TAB[i, 3] = TAB[i, 3] + Tronquer(double.Parse(DataTab.Rows[0][0].ToString()) * PRC[j] / double.Parse(DataTabss.Rows[0][0].ToString()), 4);
                            }

                            TAB[i, 3] = Tronquer((double.Parse(TAB[i, 3]) * 0.85 + 0.15) * double.Parse(TAB[i, 2]), 4).ToString();
                            TT = TT + double.Parse(TAB[i, 3]);
                        }


                        for (i = 0; i < TM; i++)
                        {
                            if (i == 0)
                            {
                                DataAdap = new System.Data.OleDb.OleDbDataAdapter("Select Date_Fin From Décompte where Num_Marché='" + CmbNumMarchéM.Text + "' and Désignation='DP" + Convert.ToString(h + 1) + "'", connection);
                                DataTab = new System.Data.DataTable();
                                DataAdap.Fill(DataTab);

                                ExcelSheet.Cells[21 + l + i, 1] = "DP" + Convert.ToString(h + 1);
                                ExcelSheet.Cells[21 + l + i, 1].Borders.Color = Color.Black;
                                ExcelSheet.Cells[21 + l + i, 1].Font.Size = 13;
                                ExcelSheet.Cells[21 + l + i, 1].Font.Name = "Times New Roman";
                                ExcelSheet.Cells[21 + l + i, 1].HorizontalAlignment = 3;
                                ExcelSheet.Cells[21 + l + i, 1].Borders.Color = Color.Black;
                                ExcelSheet.Cells[21 + l + i, 1].VerticalAlignment = HorizontalAlignment.Center;

                                //----------------------------------------------------------------

                                ExcelSheet.Range["A" + (21 + l + i).ToString()].NumberFormat = "@";
                                ExcelSheet.Cells[21 + l + i, 2] = DataTab.Rows[0][0].ToString();
                                ExcelSheet.Cells[21 + l + i, 2].Borders.Color = Color.Black;
                                ExcelSheet.Cells[21 + l + i, 2].Font.Size = 13;
                                ExcelSheet.Cells[21 + l + i, 2].Font.Name = "Times New Roman";
                                ExcelSheet.Cells[21 + l + i, 2].HorizontalAlignment = 3;
                                ExcelSheet.Cells[21 + l + i, 2].Borders.Color = Color.Black;
                                ExcelSheet.Cells[21 + l + i, 2].VerticalAlignment = HorizontalAlignment.Center;

                                ExcelSheet.Cells[21 + l + i, 3].Borders.Color = Color.Black;
                                ExcelSheet.Cells[21 + l + i, 3] = Tronquer(MRT, 2);
                                ExcelSheet.Cells[21 + l + i, 3].Font.Size = 13;
                                ExcelSheet.Cells[21 + l + i, 3].Font.Name = "Times New Roman";
                                ExcelSheet.Cells[21 + l + i, 3].HorizontalAlignment = 3;
                                ExcelSheet.Cells[21 + l + i, 3].VerticalAlignment = HorizontalAlignment.Center;

                            }
                            else
                            {
                                ExcelSheet.cells(21 + l + i, 1).Borders.Color = Color.Black;
                                ExcelSheet.cells(21 + l + i, 2).Borders.Color = Color.Black;
                                ExcelSheet.cells(21 + l + i, 3).Borders.Color = Color.Black;
                            }
                            ExcelSheet.Cells[21 + l + i, 4] = Tronquer(double.Parse(TAB[i, 2]), 2);

                            ExcelSheet.Cells[21 + l + i, 4].Borders.Color = Color.Black;
                            ExcelSheet.Cells[21 + l + i, 4].Font.Size = 13;
                            ExcelSheet.Cells[21 + l + i, 4].Font.Name = "Times New Roman";
                            ExcelSheet.Cells[21 + l + i, 4].HorizontalAlignment = 3;

                            ExcelSheet.Cells[21 + l + i, 5].Borders.Color = Color.Black;
                            ExcelSheet.Cells[21 + l + i, 5].VerticalAlignment = HorizontalAlignment.Center;
                            ExcelSheet.Cells[21 + l + i, 5] = TAB[i, 1];
                            ExcelSheet.Cells[21 + l + i, 5].Font.Size = 13;
                            ExcelSheet.Cells[21 + l + i, 5].Font.Name = "Times New Roman";
                            ExcelSheet.Cells[21 + l + i, 5].HorizontalAlignment = 3;
                            ExcelSheet.Cells[21 + l + i, 5].VerticalAlignment = HorizontalAlignment.Center;

                            ExcelSheet.Cells[21 + l + i, 6] = TAB[i, 0];
                            ExcelSheet.Cells[21 + l + i, 6].Borders.Color = Color.Black;
                            ExcelSheet.Cells[21 + l + i, 6].Font.Size = 13;
                            ExcelSheet.Cells[21 + l + i, 6].Font.Name = "Times New Roman";
                            ExcelSheet.Cells[21 + l + i, 6].HorizontalAlignment = 3;

                            ExcelSheet.Cells[21 + l + i, 7].Borders.Color = Color.Black;
                            ExcelSheet.Cells[21 + l + i, 7].VerticalAlignment = HorizontalAlignment.Center;
                            ExcelSheet.Cells[21 + l + i, 7] = Tronquer(long.Parse(TAB[i, 1]) / (double)NJDP[h], 4);
                            ExcelSheet.Cells[21 + l + i, 7].Font.Size = 13;
                            ExcelSheet.Cells[21 + l + i, 7].Font.Name = "Times New Roman";
                            ExcelSheet.Cells[21 + l + i, 7].HorizontalAlignment = 3;
                            ExcelSheet.Cells[21 + l + i, 7].VerticalAlignment = HorizontalAlignment.Center;

                            ExcelSheet.Cells[21 + l + i, 8] = Tronquer((double)(Convert.ToSingle(TAB[i, 3]) / double.Parse(TAB[i, 2])), 4);

                            if (ExcelSheet.Columns(7).ColumnWidth < (((string)(ExcelSheet.cells(21 + l + i, 7).value)).Length + 4) * 1.3)
                            {
                                ExcelSheet.Columns(7).ColumnWidth = (((string)(ExcelSheet.cells(21 + l + i, 7).value)).Length + 4) * 1.3;
                            }
                            ExcelSheet.Cells[21 + l + i, 8].Borders.Color = Color.Black;
                            ExcelSheet.Cells[21 + l + i, 8].Font.Size = 13;
                            ExcelSheet.Cells[21 + l + i, 8].Font.Name = "Times New Roman";
                            ExcelSheet.Cells[21 + l + i, 8].HorizontalAlignment = 3;

                            ExcelSheet.Cells[21 + l + i, 9].Borders.Color = Color.Black;
                            ExcelSheet.Cells[21 + l + i, 9].VerticalAlignment = HorizontalAlignment.Center;

                            TotalRévisé = TotalRévisé + Tronquer(MRT * Convert.ToDouble(ExcelSheet.Cells[21 + l + i, 7].Value) * (-1 + Convert.ToDouble(ExcelSheet.Cells[21 + l + i, 8].Value)), 2);
                            MRPD = MRPD + Tronquer(MRT * Convert.ToDouble(ExcelSheet.Cells[21 + l + i, 7].Value) * (-1 + Convert.ToDouble(ExcelSheet.Cells[21 + l + i, 8].Value)), 2);

                            ExcelSheet.Cells[21 + l + i, 9] = Tronquer(MRT * Convert.ToDouble(ExcelSheet.Cells[21 + l + i, 7].Value) * (-1 + Convert.ToDouble(ExcelSheet.Cells[21 + l + i, 8].Value)), 2);

                            ExcelSheet.Cells[21 + l + i, 9].Font.Size = 13;
                            ExcelSheet.Cells[21 + l + i, 9].Font.Name = "Times New Roman";
                            ExcelSheet.Cells[21 + l + i, 9].HorizontalAlignment = 3;
                            ExcelSheet.Cells[21 + l + i, 9].VerticalAlignment = HorizontalAlignment.Center;

                        }
                        ExcelSheet.Range("A" + (21 + l).ToString() + ":A" + (21 + l + i - 1).ToString()).Merge();
                        ExcelSheet.Range("B" + (21 + l).ToString() + ":B" + (21 + l + i - 1).ToString()).Merge();
                        ExcelSheet.Range("C" + (21 + l).ToString() + ":C" + (21 + l + i - 1).ToString()).Merge();
                        //*********** save révision ********************************
                        SQLSTR = "UPDATE Décompte SET ";
                        SQLSTR = SQLSTR + "Montant_Révisé=" + MRPD.ToString().Replace(",", ".");
                        SQLSTR = SQLSTR + " where  Num_Marché='" + CmbNumMarchéM.Text + "' and N°=" + NumDécompte;
                        ExecuterSQL(SQLSTR);
                        //*******************************************************************
                        l = l + i;
                    }
                    if (h == HL.GetUpperBound(0))
                    {
                        ExcelSheet.Cells[21 + l, 1].Interior.Color = Color.Gold;
                        ExcelSheet.Cells[21 + l, 1].Borders.Color = Color.Black;
                        ExcelSheet.Cells[21 + l, 1].VerticalAlignment = XlHAlign.xlHAlignCenter;
                        ExcelSheet.Cells[21 + l, 1] = "TOTAL REVISION DES PRIX H.T";
                        ExcelSheet.Cells[21 + l, 1].Font.Bold = true;
                        ExcelSheet.Cells[21 + l, 1].Font.Size = 13;
                        ExcelSheet.Cells[21 + l, 1].Font.Name = "Times New Roman";
                        ExcelSheet.Cells[21 + l, 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        ExcelSheet.Cells[21 + l, 1].VerticalAlignment = XlVAlign.xlVAlignCenter;

                        ExcelSheet.Cells[21 + l, 9].Interior.Color = Color.Gold;
                        ExcelSheet.Cells[21 + l, 9].Borders.Color = Color.Black;
                        ExcelSheet.Cells[21 + l, 9].VerticalAlignment = XlHAlign.xlHAlignCenter;
                        ExcelSheet.Cells[21 + l, 9] = Math.Round(TotalRévisé, 2);
                        ExcelSheet.Cells[21 + l, 9].Font.Bold = true;
                        ExcelSheet.Cells[21 + l, 9].Font.Size = 13;
                        ExcelSheet.Cells[21 + l, 9].Font.Name = "Times New Roman";
                        ExcelSheet.Cells[21 + l, 9].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        ExcelSheet.Cells[21 + l, 9].VerticalAlignment = XlVAlign.xlVAlignCenter;

                        ExcelSheet.Range["A" + (21 + l).ToString() + ":H" + (21 + l).ToString()].Merge();

                        for (i = 1; i <= 8; i++)
                        {
                            ExcelSheet.cells(21 + l, i).Borders.Color = Color.Black;
                        }
                        ExcelSheet.Cells[21 + l + 1, 1].Interior.Color = Color.Gold;
                        ExcelSheet.Cells[21 + l + 1, 1].Borders.Color = Color.Black;
                        ExcelSheet.Cells[21 + l + 1, 1].VerticalAlignment = XlHAlign.xlHAlignCenter;
                        ExcelSheet.Cells[21 + l + 1, 1] = "TVA (20%)";
                        ExcelSheet.Cells[21 + l + 1, 1].Font.Bold = true;
                        ExcelSheet.Cells[21 + l + 1, 1].Font.Size = 13;
                        ExcelSheet.Cells[21 + l + 1, 1].Font.Name = "Times New Roman";
                        ExcelSheet.Cells[21 + l + 1, 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        ExcelSheet.Cells[21 + l + 1, 1].VerticalAlignment = XlVAlign.xlVAlignCenter;

                        ExcelSheet.Cells[21 + l + 1, 9].Interior.Color = Color.Gold;
                        ExcelSheet.Cells[21 + l + 1, 9].Borders.Color = Color.Black;
                        ExcelSheet.Cells[21 + l + 1, 9].VerticalAlignment = XlHAlign.xlHAlignCenter;
                        ExcelSheet.Cells[21 + l + 1, 9] = Tronquer(TotalRévisé * 0.2, 2);
                        ExcelSheet.Cells[21 + l + 1, 9].Font.Bold = true;
                        ExcelSheet.Cells[21 + l + 1, 9].Font.Size = 13;
                        ExcelSheet.Cells[21 + l + 1, 9].Font.Name = "Times New Roman";
                        ExcelSheet.Cells[21 + l + 1, 9].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        ExcelSheet.Cells[21 + l + 1, 9].VerticalAlignment = XlVAlign.xlVAlignCenter;

                        ExcelSheet.Range["A" + (21 + l + 1).ToString() + ":H" + (21 + l + 1).ToString()].Merge();

                        for (i = 1; i <= 8; i++)
                        {
                            ExcelSheet.cells(21 + l + 1, i).Borders.Color = Color.Black;
                        }
                        ExcelSheet.Cells[21 + l + 2, 1].Interior.Color = Color.Gold;
                        ExcelSheet.Cells[21 + l + 2, 1].Borders.Color = Color.Black;
                        ExcelSheet.Cells[21 + l + 2, 1].VerticalAlignment = XlHAlign.xlHAlignCenter;
                        ExcelSheet.Cells[21 + l + 2, 1] = "TOTAL REVISION DES PRIX TTC";
                        ExcelSheet.Cells[21 + l + 2, 1].Font.Bold = true;
                        ExcelSheet.Cells[21 + l + 2, 1].Font.Size = 13;
                        ExcelSheet.Cells[21 + l + 2, 1].Font.Name = "Times New Roman";
                        ExcelSheet.Cells[21 + l + 2, 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        ExcelSheet.Cells[21 + l + 2, 1].VerticalAlignment = XlVAlign.xlVAlignCenter;

                        ExcelSheet.Cells[21 + l + 2, 9].Interior.Color = Color.Gold;
                        ExcelSheet.Cells[21 + l + 2, 9].Borders.Color = Color.Black;
                        ExcelSheet.Cells[21 + l + 2, 9].VerticalAlignment = XlHAlign.xlHAlignCenter;
                        ExcelSheet.Cells[21 + l + 2, 9] = Math.Round(TotalRévisé * 1.2, 2);

                        ExcelSheet.Cells[21 + l + 2, 9].Font.Bold = true;
                        ExcelSheet.Cells[21 + l + 2, 9].Font.Size = 13;
                        ExcelSheet.Cells[21 + l + 2, 9].Font.Name = "Times New Roman";
                        ExcelSheet.Cells[21 + l + 2, 9].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        ExcelSheet.Cells[21 + l + 2, 9].VerticalAlignment = XlVAlign.xlVAlignCenter;

                        ExcelSheet.Range["A" + (21 + l + 2).ToString() + ":H" + (21 + l + 2).ToString()].Merge();

                        for (i = 1; i <= 8; i++)
                        {
                            ExcelSheet.Cells[21 + l + 2, i].Borders.Color = Color.Black;
                        }

                        ExcelSheet.Cells[21 + l + 4, 1].VerticalAlignment = XlHAlign.xlHAlignCenter;

                        if (ExcelSheet.Cells[21 + l + 2, 9].Value < 0)
                        {
                            ExcelSheet.Cells[21 + l + 4, 1] = "     Arrêté la présente note de calcul de la révision des prix à la somme TTC du : Moins(-) " + Calcul.ChiffreToLettre(Math.Abs(ExcelSheet.Cells[21 + l + 2, 9].Value));
                        }
                        else
                        {
                            ExcelSheet.Cells[21 + l + 4, 1] = "     Arrêté la présente note de calcul de la révision des prix à la somme TTC du :" + Calcul.ChiffreToLettre(ExcelSheet.Cells[21 + l + 2, 9].Value);
                        }

                        ExcelSheet.Cells[21 + l + 4, 1].Font.Bold = true;
                        ExcelSheet.Cells[21 + l + 4, 1].Font.Size = 13;
                        ExcelSheet.Cells[21 + l + 4, 1].Font.Color = Color.DarkRed;
                        ExcelSheet.Cells[21 + l + 4, 1].Font.Name = "Times New Roman";
                        ExcelSheet.Cells[21 + l + 4, 1].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                        ExcelSheet.Cells[21 + l + 4, 1].VerticalAlignment = XlVAlign.xlVAlignCenter;

                    }
                }

                connections.Close();
                connectionss.Close();
                ExcelApp.Visible = true;
                ExcelSheet = null;
                ExcelBook = null;
                ExcelApp = null;
            }
        }
        private double Tronquer(double V, int U)
        {
            string s = null;
            string D = null;
            int i = 0;
            int j = 0;
            s = V.ToString();
            i = s.IndexOf(",") + 1;
            j = s.IndexOf(".") + 1;
            i = i + j;
            if (s.Length - i <= U || i == 0)
            {
                D = s;
            }
            else
            {
                D = s.Substring(0, i + U);
            }
            return double.Parse(D);
        }

        private void BtnEditer_Resize(object sender, EventArgs e)
        {
			if (this.WindowState == FormWindowState.Maximized)
			{
				PNL.Left = Convert.ToInt32((this.Width - PNL.Width) / 2.0);
				PNL.Top = Convert.ToInt32((this.Height - PNL.Height) / 2.0);
			}
			else
			{
				PNL.Left = 12;
				PNL.Top = 12;
			}
		}
        private static FrmRévisionTous _DefaultInstance = new FrmRévisionTous();
        public static FrmRévisionTous DefaultInstance => _DefaultInstance.IsDisposed ? _DefaultInstance = new FrmRévisionTous() : _DefaultInstance;
    }
}
