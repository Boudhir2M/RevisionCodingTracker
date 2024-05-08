using System;
using System.Drawing;
using System.Windows.Forms;

using System.Data;
using System.Data.OleDb;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace Révision
{
    public partial class FrmRévision : Form
    {
        public FrmRévision()
        {
            InitializeComponent();
        }
		private OleDbConnection connection = new OleDbConnection();
		private OleDbCommand records;
		private System.Data.DataTable DataTab;
		private System.Data.OleDb.OleDbDataAdapter DataAdap;
		private System.Data.OleDb.OleDbCommandBuilder ComndBuld;
		private DataSet DataSetTab = new DataSet();

		private OleDbConnection connections = new OleDbConnection();
		private OleDbCommand recordss;
		private System.Data.DataTable DataTabs;
		private System.Data.OleDb.OleDbDataAdapter DataAdaps;
		private System.Data.OleDb.OleDbCommandBuilder ComndBulds;
		private DataSet DataSetTabs = new DataSet();

		private OleDbConnection connectionss = new OleDbConnection();
		private OleDbCommand recordsss;
		private System.Data.DataTable DataTabss;
		private System.Data.OleDb.OleDbDataAdapter DataAdapss;
		private System.Data.OleDb.OleDbCommandBuilder ComndBuldss;
		private DataSet DataSetTabss = new DataSet();

		private string SQLSTR;
		private long NumDécompte;
		private long Indexs;
		private string DateCommencement;
		private double TotalRévisé;

        private void CmbNumMarchéM_SelectedIndexChanged(object sender, EventArgs e)
        {
			if (string.IsNullOrEmpty(CmbNumMarchéM.Text))
			{
				MessageBox.Show("Choisissez une valeur de la liste : Réference de Marché ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
				CmbNumMarchéM.Focus();
				return;
			}
			CmbDécompteM.Items.Clear();
			DataAdap = new OleDbDataAdapter("Select Désignation From Décompte where Num_Marché='" + CmbNumMarchéM.Text + "'", connection);
			DataTab = new System.Data.DataTable();
			DataAdap.Fill(DataTab);
			if (DataTab.Rows.Count > 0)
			{
				for (var i = 0; i < DataTab.Rows.Count; i++)
				{
					CmbDécompteM.Items.Add(DataTab.Rows[i][0].ToString());
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
			NumDécompte = Convert.ToInt64(DataTab.Rows[0][0]);
			DataAdap = new OleDbDataAdapter("Select Date_Ordre_Service, Date_Arret_Service From Ordre_Service where Num_Décompte=" + NumDécompte, connection);
			DataTab = new System.Data.DataTable();
			DataAdap.Fill(DataTab);
			DGV.DataSource = DataTab;
			Indexs = CmbDécompteM.SelectedIndex;
			DataAdap = new OleDbDataAdapter("Select Désignation,Unité,Prix_unitaire,Quantité from Attachement where Num_Décompte=" + NumDécompte, connection);
			DataTab = new System.Data.DataTable();
			DataAdap.Fill(DataTab);
			DGVP.DataSource = DataTab;
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

			if (string.IsNullOrEmpty(CmbNumMarchéM.Text))
			{
				MessageBox.Show("Choisissez un Réference de Marché de la liste: Réference de Marché.", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				CmbNumMarchéM.Focus();
				return;
			}
			if (string.IsNullOrEmpty(CmbDécompteM.Text))
			{
				MessageBox.Show("Choisissez un numéro de décompte de la liste: Décompte.", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				CmbDécompteM.Focus();
				return;
			}
			DateTime dateDebut, dateFin;
			int nm;
			if (DateTime.TryParse(DGV.Rows[0].Cells[0].Value.ToString(), out dateDebut) &&
				DateTime.TryParse(DGV.Rows[DGV.Rows.Count - 2].Cells[1].Value.ToString(), out dateFin))
			{
				nm = (int)(dateFin - dateDebut).TotalDays / 30;
				// Utilisez nm comme nécessaire
			}
			else
			{
				// Gérer le cas où la conversion échoue
			}


			// ouverture des connections ----------------------------------------------------------------------------
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
            // ------------------------------------------------------------------------------------------------------------

            nm = 0;
            string[,] TAB = new string[nm + 2, 5];

			for (i = 0; i < DGV.RowCount - 1; i++)
			{
				mm = 0;

				if (DateTime.TryParse(DGV.Rows[i].Cells[1].Value.ToString(), out dateFin) &&
					DateTime.TryParse(DGV.Rows[i].Cells[0].Value.ToString(), out dateDebut))
				{
					mm = (int)(dateFin - dateDebut).TotalDays / 30;
				}
				else
				{
					// Gérer le cas où la conversion échoue
				}

				if (mm == 0)
				{
					if (i == 0)
					{
						TAB[0, 0] = DGV.Rows[i].Cells[0].Value.ToString().Substring(3, 7);
						// dateDebut, dateFin;

						if (DateTime.TryParse(DGV.Rows[0].Cells[0].Value?.ToString(), out dateDebut) &&
							DateTime.TryParse(DGV.Rows[0].Cells[1].Value?.ToString(), out dateFin))
						{
							TAB[0, 1] = ((dateFin - dateDebut).TotalDays + 1).ToString();
						}
						else
						{
							// Gérer le cas où la conversion échoue
						}
					}
					else
					{
						if (DGV.Rows[i - 1].Cells[1].Value is DateTime && DGV.Rows[i].Cells[0].Value is DateTime)
						{
							if (((DateTime)DGV.Rows[i - 1].Cells[1].Value - (DateTime)DGV.Rows[i].Cells[0].Value).TotalDays / 30 == 0)
							{
								if (int.TryParse(TAB[j, 1], out int tabValue))
								{
									TAB[j, 1] = (tabValue + ((DateTime)DGV.Rows[i].Cells[1].Value - (DateTime)DGV.Rows[i].Cells[0].Value).Days + 1).ToString();
								}
								else
								{
									// Gérer le cas où la conversion de TAB[j, 1] en entier échoue
								}
							}
							else
							{
								j++;
								TAB[j, 0] = DGV.Rows[i].Cells[0].Value.ToString().Substring(3, 7);
								TAB[j, 1] = (((DateTime)DGV.Rows[i].Cells[1].Value - (DateTime)DGV.Rows[i].Cells[0].Value).Days + 1).ToString();
							}
						}
						else
						{
							Console.WriteLine($"Valeur avant conversion : {DGV.Rows[i - 1].Cells[1].Value}, {DGV.Rows[i].Cells[0].Value}");
						}

					}
				}

				else
				{
					if (i == 0)
					{
						TAB[0, 0] = DGV.Rows[i].Cells[0].Value.ToString().Substring(3, 7);
						//int mois;

						if (DateTime.TryParse(DGV.Rows[i].Cells[0].Value?.ToString(), out dateDebut) &&
							DateTime.TryParse(DGV.Rows[i].Cells[1].Value?.ToString(), out dateFin))
						{
							mois = (int)(dateFin - dateDebut).TotalDays / 30;
							// Utilisez 'mois' comme nécessaire
						}
						else
						{
							// Gérer le cas où la conversion échoue
						}

						for (k = 0; k <= mois; k++)
						{
							if (k == 0)
							{
								m = 0;
								if (DateTime.TryParse(DGV.Rows[i].Cells[0].Value?.ToString(), out dateDebut))
								{
									m = dateDebut.Month;
									// Utilisez 'm' comme nécessaire
								}
								else
								{
									// Gérer le cas où la conversion échoue
									Console.WriteLine($"La conversion de DGV.Rows[{i}].Cells[0].Value en DateTime a échoué.");
								}

								if (m == 12)
								{
									m = 1;
									y = ((DateTime)dateDebut).Year + 1;
								}
								else
								{
									m++;
									y = ((DateTime)dateDebut).Year;
								}

								TAB[j, 0] = DGV.Rows[i].Cells[0].Value?.ToString().Substring(3, 7);


								if (DateTime.TryParse(DGV.Rows[i].Cells[1].Value?.ToString(), out dateFin))
								{
									TAB[j, 1] = (((DateTime)dateFin - DateTime.Parse("01/" + m.ToString() + "/" + y.ToString())).TotalDays + 1).ToString();
								}
								else
								{
									// Gérer le cas où la conversion échoue
									Console.WriteLine($"La conversion de DGV.Rows[{i}].Cells[1].Value en DateTime a échoué.");
								}

							}
							else if (k == mois)
							{
								j++;

								if (m == 12)
								{
									m = 1;
									y = ((DateTime)DGV.Rows[i].Cells[1].Value).Year;
									TAB[j, 0] = DGV.Rows[i].Cells[1].Value.ToString().Substring(3, 7);
									TAB[j, 1] = (((DateTime)DGV.Rows[i].Cells[1].Value - DateTime.Parse("01/" + (m - 1).ToString() + "/" + y.ToString())).TotalDays + 1).ToString();
								}
								else
								{
									m++;
									TAB[j, 0] = DGV.Rows[i].Cells[1].Value?.ToString().Substring(3, 7);

									//dateFin;
									if (DateTime.TryParse(DGV.Rows[i].Cells[1].Value?.ToString(), out dateFin))
									{
										TAB[j, 1] = (((DateTime)dateFin - DateTime.Parse("01/" + (m - 1).ToString() + "/" + y.ToString())).TotalDays + 1).ToString();
									}
									else
									{
										// Gérer le cas où la conversion échoue
										Console.WriteLine($"La conversion de DGV.Rows[{i}].Cells[1].Value en DateTime a échoué.");
									}

								}
							}
							else
							{
								j++;

								if (m == 12)
								{
									m = 1;
									y++;
									TAB[j, 0] = "12/" + y.ToString();
									TAB[j, 1] = ((DateTime.Parse("01/" + 12.ToString() + "/" + (y - 1).ToString()) - DateTime.Parse("01/" + m.ToString() + "/" + y.ToString())).TotalDays + 1).ToString();
								}
								else
								{
									m++;
									TAB[j, 0] = m.ToString() + "/" + y.ToString();
									TAB[j, 1] = ((DateTime.Parse("01/" + (m - 1).ToString() + "/" + y.ToString()) - DateTime.Parse("01/" + m.ToString() + "/" + y.ToString())).TotalDays + 1).ToString();
								}
							}
						}
					}
					else
					{
						for (i = 1; i < DGV.Rows.Count; i++)
						{
							// Vérifiez si les valeurs des cellules sont convertibles en DateTime
							if (DGV.Rows[i - 1].Cells[1].Value is DateTime && DGV.Rows[i].Cells[0].Value is DateTime)
							{
								// Effectuez la conversion
								DateTime date1 = (DateTime)DGV.Rows[i - 1].Cells[1].Value;
								DateTime date2 = (DateTime)DGV.Rows[i].Cells[0].Value;

								// Continuez avec votre logique
								if ((date1 - date2).TotalDays / 30 == 0)
								{
									mois = (int)(date1 - date2).TotalDays / 30;

									// Assurez-vous que j est initialisé à la bonne valeur
									j = 0;

									// Assurez-vous que m est initialisé à la bonne valeur
									m = 0;

									// Assurez-vous que y est initialisé à la bonne valeur
									y = 0;

									for (k = 0; k <= mois; k++)
									{
										if (k == 0)
										{
											m = date1.Month;

											if (m == 12)
											{
												m = 1;
												y = date1.Year + 1;
											}
											else
											{
												m++;
												y = date1.Year;
											}

											// Assurez-vous que TAB[j, 1] existe avant d'effectuer l'opération
											if (int.TryParse(TAB[j, 1], out int tabValue))
											{
												TAB[j, 1] = (tabValue + (date1 - DateTime.Parse("01/" + m.ToString() + "/" + y.ToString())).TotalDays + 1).ToString();
											}
											else
											{
												// Gestion du cas où la conversion en int échoue
												Console.WriteLine("Conversion en int échouée pour TAB[j, 1]");
											}
										}
										else if (k == mois)
										{
											j++;

											if (m == 12)
											{
												m = 1;
												y = date1.Year;
												TAB[j, 0] = date1.ToString("MM/yyyy");
												TAB[j, 1] = (((date1 - DateTime.Parse("01/" + (m - 1).ToString() + "/" + (y - 1).ToString())).TotalDays + 1).ToString());
											}
											else
											{
												m++;
												TAB[j, 0] = date1.ToString("MM/yyyy");
												TAB[j, 1] = (((date1 - DateTime.Parse("01/" + (m - 1).ToString() + "/" + y.ToString())).TotalDays + 1).ToString());
											}
										}
										else
										{
											j++;

											if (m == 12)
											{
												m = 1;
												y++;
												TAB[j, 0] = "12/" + y.ToString();
												TAB[j, 1] = (DateTime.Parse("01/" + 12.ToString() + "/" + (y - 1).ToString()).AddMonths(1) - DateTime.Parse("01/" + 12.ToString() + "/" + (y - 1).ToString())).TotalDays.ToString();
											}
											else
											{
												m++;
												TAB[j, 0] = m.ToString() + "/" + y.ToString();
												TAB[j, 1] = (DateTime.Parse("01/" + m.ToString() + "/" + y.ToString()) - DateTime.Parse("01/" + (m - 1).ToString() + "/" + y.ToString())).TotalDays.ToString();
											}
										}
									}
								}
							}
							else
							{
								// Gestion du cas où la conversion échoue
								Console.WriteLine("Conversion en DateTime échouée pour la ligne " + i);
							}
						}
					}
				}
			}

			// Vous pouvez maintenant utiliser le tableau TAB comme nécessaire.
			double[] TI = null; //Montant de chaque ensemble de prestations pour un même symbole
			double[] PRC = null;
			double PJ = 0D; // Prix journalier
			double TT = 0D; // total décompte
			double NJ = 0D; // Nombre de jours
			long NJDP = 0;
			for (i = 0; i <= nm; i++)
			{
				NJ = NJ + long.Parse(TAB[i, 1]);
			}
			NJDP = Convert.ToInt64(NJ);
			DataAdaps = new OleDbDataAdapter("Select DISTINCT Symbole From Attachement where Num_Décompte=" + NumDécompte, connections);
			DataTabs = new System.Data.DataTable();
			DataAdaps.Fill(DataTabs);
			TI = new double[DataTabs.Rows.Count + 1];
			int HH = Convert.ToInt32(CmbDécompteM.Text.Substring(2, CmbDécompteM.Text.Length - 2));

			for (i = 0; i < DataTabs.Rows.Count; i++)
			{
				DataAdap = new OleDbDataAdapter("Select Quantité,Prix_Unitaire From Attachement where Num_Décompte=" + NumDécompte + " and Symbole='" + DataTabs.Rows[i][0].ToString() + "' order by Désignation", connection);
				DataTab = new System.Data.DataTable();
				DataAdap.Fill(DataTab);

				if (HH > 1)
				{
					DataAdapss = new OleDbDataAdapter("Select N° From Décompte where Désignation='DP" + (HH - 1).ToString() + "' and Num_Marché='" + CmbNumMarchéM.Text + "'", connectionss);
					DataTabss = new System.Data.DataTable();
					DataAdapss.Fill(DataTabss);
					NumDécompte = Convert.ToInt64(DataTabss.Rows[0][0]);
					DataAdapss = new OleDbDataAdapter("Select Quantité,Prix_Unitaire From Attachement where Num_Décompte=" + NumDécompte + " and Symbole='" + DataTabs.Rows[i][0].ToString() + "' order by Désignation", connectionss);
					DataTabss = new System.Data.DataTable();
					DataAdapss.Fill(DataTabss);
				}

				for (j = 0; j < DataTab.Rows.Count; j++)
				{
					if (HH == 1)
					{
						TI[i] = TI[i] + Convert.ToDouble(DataTab.Rows[j][0]) * Convert.ToDouble(DataTab.Rows[j][1]);
					}
					else
					{
						TI[i] = TI[i] + (Convert.ToDouble(DataTab.Rows[j][0]) - Convert.ToDouble(DataTabss.Rows[j][0])) * Convert.ToDouble(DataTab.Rows[j][1]);
					}
				}

				TT = TT + TI[i];
			}

			PJ = TT / NJ;
			PRC = new double[DataTabs.Rows.Count];

			string strf = "";
			DataAdaps = new OleDbDataAdapter("Select DISTINCT Symbole From Attachement where Num_Décompte=" + NumDécompte, connections);
			DataTabs = new System.Data.DataTable();
			DataAdaps.Fill(DataTabs);

			TT = 0;

			for (i = 0; i < DataTabs.Rows.Count; i++)
			{
				DataAdap = new OleDbDataAdapter("Select Quantité,Prix_Unitaire From Attachement where Num_Décompte=" + NumDécompte + " and Symbole='" + DataTabs.Rows[i][0].ToString() + "'", connection);
				DataTab = new System.Data.DataTable();
				DataAdap.Fill(DataTab);

				// Utilisez une boucle distincte pour itérer sur les lignes de DataTab
				for (k = 0; k < DataTab.Rows.Count; k++)
				{
					if (TI.Length > i)
					{
						// Utilisez k pour accéder à l'indice de la ligne dans DataTab
						TI[i] = Convert.ToDouble(DataTab.Rows[k][0]) * Convert.ToDouble(DataTab.Rows[k][1]);
					}
				}

				TT += TI[i];
			}


			for (i = 0; i < DataTabs.Rows.Count; i++)
			{
				if (TI.Length > i && TT != 0)
				{
					PRC[i] = TI[i] / TT;
					if (i == 0)
					{
						strf = "P=P0*(0.15 + 0.85*(" + Math.Round(PRC[i], 3).ToString() + "*(" + DataTabs.Rows[i][0].ToString() + "/" + DataTabs.Rows[i][0].ToString() + "0)";
					}
					else
					{
						strf = strf + " + " + Math.Round(PRC[i], 3).ToString() + "*(" + DataTabs.Rows[i][0].ToString() + "/" + DataTabs.Rows[i][0].ToString() + "0)";
					}
					strf = strf + ")";
				}
				else
				{
					// Gérer le cas où l'indice est en dehors des limites du tableau TI
					Console.WriteLine("L'indice est en dehors des limites du tableau TI.");
				}
			}


			string MoisBase = null;
			DataAdaps = new OleDbDataAdapter("Select Epoque_Base From Marché where Réference='" + CmbNumMarchéM.Text + "'", connections);
			DataTabs = new System.Data.DataTable();
			DataAdaps.Fill(DataTabs);
			MoisBase = DataTabs.Rows[0][0].ToString();

			long NumSymbole = 0;
			long TM = 0;
			i = 0;

			// Assurez-vous que l'indice i reste dans les limites du tableau
			while (i < TAB.GetLength(0) && !string.IsNullOrEmpty(TAB[i, 0]))
			{
				TM = TM + 1;
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

					DataAdapss = new OleDbDataAdapter("Select Valeur From Indexs where Mois='" + MoisBase + "' and Symbole=" + NumSymbole, connectionss);
					DataTabss = new System.Data.DataTable();
					DataAdapss.Fill(DataTabss);

					double ratio = 0.0;

					// Assurez-vous que DataTab a au moins une ligne
					if (DataTab.Rows.Count > 0)
					{
						ratio = Convert.ToDouble(DataTab.Rows[0][0]) * PRC[j] / Convert.ToDouble(DataTabss.Rows[0][0]);
					}

					double tempResult = (Convert.ToDouble(TAB[i, 3]) * 0.85) + 0.15;
					TAB[i, 3] = tempResult.ToString();
					TT += Convert.ToDouble(TAB[i, 3]);
				}

				dynamic ExcelApp = null;
				dynamic ExcelBook = null;
				dynamic ExcelSheet = null;

				// Create object of Excel
				ExcelApp = new Excel.Application();
				ExcelBook = ExcelApp.Workbooks.Add();
				ExcelSheet = (Excel.Worksheet)ExcelBook.Worksheets[1];

				//===================================feuille de calcul attachement==========================
				// Ajouter une nouvelle feuille de calcul
				ExcelSheet = ExcelBook.Worksheets.Add();
				// Sélectionner la deuxième feuille de calcul
				ExcelSheet = ExcelBook.Worksheets[2];
				ExcelSheet.Name = "Récapitulatif";
				// Sélectionner la première feuille de calcul
				ExcelSheet = ExcelBook.Worksheets[1];
				ExcelSheet.Name = "Révision";

				// Remplir la première cellule avec le texte
				ExcelSheet.Cells[1, 1] = "Royaume du Maroc" + "\r\n" + "Ministère de l'Intérieur" + "\r\n" + "Province d'Al Haouz" + "\r\n" + "Conseil Provincial Al Haouz";
				// Ajuster la hauteur de la première ligne
				ExcelSheet.Rows["1:1"].RowHeight = 60;

				// Formater la première cellule
				ExcelSheet.Cells[1, 1].Font.Bold = true;
				ExcelSheet.Cells[1, 1].Font.Size = 10;
				ExcelSheet.Cells[1, 1].Font.Name = "Times New Roman";
				ExcelSheet.Cells[1, 1].HorizontalAlignment = XlHAlign.xlHAlignLeft;
				ExcelSheet.Cells[1, 1].VerticalAlignment = XlVAlign.xlVAlignCenter;

				// Remplir la troisième cellule avec le texte
				ExcelSheet.Cells[2, 1] = "Objet : " + TxtObjet.Text;
				//ExcelSheet.Rows["2:1"].RowHeight = 20;
				ExcelSheet.Cells[2, 1].Font.Bold = true;
				ExcelSheet.Cells[2, 1].Font.Size = 13;
				ExcelSheet.Cells[2, 1].Font.Name = "Times New Roman";
				ExcelSheet.Cells[2, 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
				ExcelSheet.Cells[2, 1].VerticalAlignment = XlVAlign.xlVAlignCenter;
				ExcelSheet.Range["A2:H2"].Merge();

				// Remplir la première cellule avec le texte
				ExcelSheet.Cells[3,1] = "Marché n° :" + CmbNumMarchéM.Text;
				ExcelSheet.Cells[3, 1].Font.Bold = true;
				ExcelSheet.Cells[3, 1].Font.Size = 13;
				ExcelSheet.Cells[3, 1].Font.Name = "Times New Roman";
				ExcelSheet.Cells[3, 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
				ExcelSheet.Cells[3, 1].VerticalAlignment = XlVAlign.xlVAlignCenter;

				// Fusionner les cellules C1 à H1
				ExcelSheet.Range["A3:H3"].Merge();

				// Remplir la quatrième cellule avec le texte
				ExcelSheet.Cells[4, 1] = "ETAT DE LA REVISION DES PRIX";
				// Formater la quatrième cellule
				ExcelSheet.Cells[4, 1].Interior.Color = ColorTranslator.ToOle(Color.SkyBlue);
				ExcelSheet.Cells[4, 1].Font.Bold = true;
				ExcelSheet.Cells[4, 1].Font.Size = 13;
				ExcelSheet.Cells[4, 1].Font.Name = "Times New Roman";
				ExcelSheet.Cells[4, 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
				ExcelSheet.Cells[4, 1].VerticalAlignment = XlVAlign.xlVAlignCenter;
				// Fusionner les cellules A4 à H4
				ExcelSheet.Range["A4:H4"].Merge();

				// Remplir la sixième cellule avec le texte
				ExcelSheet.Cells[6, 1] = "Date d'époque de base : " + MoisBase;
				// Formater la sixième cellule
				ExcelSheet.Cells[6, 1].Font.Bold = true;
				ExcelSheet.Cells[6, 1].Font.Size = 13;
				ExcelSheet.Cells[6, 1].Font.Name = "Times New Roman";
				ExcelSheet.Cells[6, 1].HorizontalAlignment = XlHAlign.xlHAlignLeft;
				ExcelSheet.Cells[6, 1].VerticalAlignment = XlVAlign.xlVAlignCenter;

				// Remplir la septième cellule avec le texte
				ExcelSheet.Cells[7, 1] = "Ordre de service de Commencement le : " + DateCommencement;
				// Formater la septième cellule
				ExcelSheet.Cells[7, 1].Font.Bold = true;
				ExcelSheet.Cells[7, 1].Font.Size = 13;
				ExcelSheet.Cells[7, 1].Font.Name = "Times New Roman";
				ExcelSheet.Cells[7, 1].HorizontalAlignment = XlHAlign.xlHAlignLeft;
				ExcelSheet.Cells[7, 1].VerticalAlignment = XlVAlign.xlVAlignCenter;

				// Remplir la huitième cellule avec le texte
				ExcelSheet.Cells[8, 1] = "Formule de la révision des prix : ";
				// Formater la huitième cellule
				ExcelSheet.Cells[8, 1].Font.Bold = true;
				ExcelSheet.Cells[8, 1].Font.Underline = true;
				ExcelSheet.Cells[8, 1].Font.Size = 13;
				ExcelSheet.Cells[8, 1].Font.Name = "Times New Roman";
				ExcelSheet.Cells[8, 1].HorizontalAlignment = XlHAlign.xlHAlignLeft;
				ExcelSheet.Cells[8, 1].VerticalAlignment = XlVAlign.xlVAlignCenter;

				// Remplir la neuvième cellule avec le texte
				ExcelSheet.Cells[9, 1] = strf;
				ExcelSheet.Cells[9, 1].Font.Bold = true;
				ExcelSheet.Cells[9, 1].Font.Size = 13;
				ExcelSheet.Cells[9, 1].Font.Name = "Times New Roman";
				ExcelSheet.Cells[9, 1].HorizontalAlignment = XlHAlign.xlHAlignLeft;
				ExcelSheet.Cells[9, 1].VerticalAlignment = XlVAlign.xlVAlignCenter;


				DataAdaps = new OleDbDataAdapter("Select DISTINCT Symbole From Attachement where Num_Décompte=" + NumDécompte, connections);
				DataTabs = new System.Data.DataTable();
				DataAdaps.Fill(DataTabs);

				// Assurez-vous que ExcelSheet est correctement défini avant d'utiliser les cellules
				if (ExcelSheet != null)
				{
					ExcelSheet.Cells[10, 1] = "Symbole";
					ExcelSheet.Cells[10, 1].Interior.Color = ColorTranslator.ToOle(Color.Gold);
					ExcelSheet.Cells[10, 1].Borders.Color = ColorTranslator.ToOle(Color.Black);
					ExcelSheet.Cells[10, 1].Font.Bold = true;
					ExcelSheet.Cells[10, 1].Font.Size = 13;
					ExcelSheet.Cells[10, 1].Font.Name = "Times New Roman";
					ExcelSheet.Cells[10, 1].HorizontalAlignment = XlHAlign.xlHAlignRight;

					ExcelSheet.Cells[10, 2] = "Index de Base";
					ExcelSheet.Cells[10, 2].Interior.Color = ColorTranslator.ToOle(Color.Gold);
					ExcelSheet.Cells[10, 2].Borders.Color = ColorTranslator.ToOle(Color.Black);
					ExcelSheet.Cells[10, 2].Font.Bold = true;
					ExcelSheet.Cells[10, 2].Font.Size = 13;
					ExcelSheet.Cells[10, 2].Font.Name = "Times New Roman";
					ExcelSheet.Cells[10, 2].HorizontalAlignment = XlHAlign.xlHAlignRight;
					ExcelSheet.Cells[10, 2].VerticalAlignment = XlVAlign.xlVAlignCenter;
				}

				for (i = 0; i < DataTabs.Rows.Count; i++)
				{
                    try
                    {
						// Récupération du Numéro de Symbole
						DataAdap = new OleDbDataAdapter("Select Num From symbole where Code='" + DataTabs.Rows[i][0].ToString() + "'", connection);
						DataTab = new System.Data.DataTable();
						DataAdap.Fill(DataTab);
						NumSymbole = Convert.ToInt64(DataTab.Rows[0][0]);

						// Log de débogage
						Console.WriteLine("Code Symbole: " + DataTabs.Rows[i][0].ToString());
						Console.WriteLine("NumSymbole: " + NumSymbole);

						// Récupération des valeurs pour chaque mois
						DataAdapss = new OleDbDataAdapter("Select Valeur From Indexs where Mois='" + MoisBase + "' and Symbole=" + NumSymbole, connectionss);
						DataTabss = new System.Data.DataTable();
						DataAdapss.Fill(DataTabss);
						ExcelSheet.Cells[11 + i, 1] = DataTabs.Rows[i][0].ToString();
						ExcelSheet.Cells[11 + i, 1].Font.Bold = true;
						ExcelSheet.Cells[11 + i, 1].Font.Size = 13;
						ExcelSheet.Cells[11 + i, 1].Font.Name = "Times New Roman";
						ExcelSheet.Cells[11 + i, 1].HorizontalAlignment = XlHAlign.xlHAlignRight;
						ExcelSheet.Cells[11 + i, 1].Borders.Color = ColorTranslator.ToOle(Color.Black);
						ExcelSheet.Cells[11 + i, 1].VerticalAlignment = XlVAlign.xlVAlignCenter;

						ExcelSheet.Cells[11 + i, 2] = DataTabss.Rows[0][0].ToString();
						ExcelSheet.Cells[11 + i, 2].Font.Size = 13;
						ExcelSheet.Cells[11 + i, 2].Font.Name = "Times New Roman";
						ExcelSheet.Cells[11 + i, 2].HorizontalAlignment = XlHAlign.xlHAlignRight;
						ExcelSheet.Cells[11 + i, 2].VerticalAlignment = XlVAlign.xlVAlignCenter;
						ExcelSheet.Cells[11 + i, 2].Borders.Color = ColorTranslator.ToOle(Color.Black);

						Console.WriteLine("Valeur récupérée : " + DataTabss.Rows[0][0].ToString());
					}
					catch (Exception ex)
					{
						// Gestion des erreurs
						Console.WriteLine("Erreur SQL : " + ex.Message);
					}
				}

				ExcelSheet.Cells[16, 1] = "Symbole";
				ExcelSheet.Cells[16, 1].Interior.Color = ColorTranslator.ToOle(Color.Gold);
				ExcelSheet.Cells[16, 1].Borders.Color = ColorTranslator.ToOle(Color.Black);
				ExcelSheet.Cells[16, 1].Font.Bold = true;
				ExcelSheet.Cells[16, 1].Font.Size = 13;
				ExcelSheet.Cells[16, 1].Font.Name = "Times New Roman";
				ExcelSheet.Cells[16, 1].HorizontalAlignment = XlHAlign.xlHAlignRight;

				for (i = 0; i < DataTabs.Rows.Count; i++)
				{
					try
					{
						DataAdap = new OleDbDataAdapter("Select Num From symbole where Code='" + DataTabs.Rows[i][0].ToString() + "'", connection);
						DataTab = new System.Data.DataTable();
						DataAdap.Fill(DataTab);

						if (DataTab.Rows.Count > 0)
						{
							NumSymbole = Convert.ToInt64(DataTab.Rows[0][0]);

							// Afficher le code Symbole dans Excel
							ExcelSheet.Cells[17 + i, 1] = DataTabs.Rows[i][0].ToString();
							ExcelSheet.Cells[17 + i, 1].Font.Bold = true;
							ExcelSheet.Cells[17 + i, 1].Font.Size = 13;
							ExcelSheet.Cells[17 + i, 1].Font.Name = "Times New Roman";
							ExcelSheet.Cells[17 + i, 1].HorizontalAlignment = XlHAlign.xlHAlignRight;
							ExcelSheet.Cells[17 + i, 1].Borders.Color = ColorTranslator.ToOle(Color.Black);
							ExcelSheet.Cells[17 + i, 1].VerticalAlignment = XlVAlign.xlVAlignCenter;

							// Utilisation d'une variable distincte pour l'index de la boucle interne
							for (j = 0; j < TM; j++)
							{
								DataAdapss = new OleDbDataAdapter("Select Valeur From Indexs where Mois='" + DateTime.Parse("01/" + TAB[j, 0]) + "' and Symbole=" + NumSymbole, connectionss);
								DataTabss = new System.Data.DataTable();
								DataAdapss.Fill(DataTabss);

								// Afficher le mois dans Excel
								ExcelSheet.Cells[16, 2 + j] = TAB[j, 0];
								ExcelSheet.Cells[16, 2 + j].Font.Bold = true;
								ExcelSheet.Cells[16, 2 + j].Font.Size = 13;
								ExcelSheet.Cells[16, 2 + j].Font.Name = "Times New Roman";
								ExcelSheet.Cells[16, 2 + j].HorizontalAlignment = XlHAlign.xlHAlignRight;
								ExcelSheet.Cells[16, 2 + j].Borders.Color = ColorTranslator.ToOle(Color.Black);
								ExcelSheet.Cells[16, 2 + j].VerticalAlignment = XlVAlign.xlVAlignCenter;

								if (DataTabss.Rows.Count > 0)
								{
									// Afficher la valeur associée dans Excel
									ExcelSheet.Cells[17 + i, 2 + j] = DataTabss.Rows[0][0].ToString();
								}
								else
								{
									Console.WriteLine("DataTabss n'a pas de lignes à la position 0.");
								}

								ExcelSheet.Cells[17 + i, 2 + j].Font.Size = 13;
								ExcelSheet.Cells[17 + i, 2 + j].Font.Name = "Times New Roman";
								ExcelSheet.Cells[17 + i, 2 + j].HorizontalAlignment = XlHAlign.xlHAlignRight;
								ExcelSheet.Cells[17 + i, 2 + j].VerticalAlignment = XlVAlign.xlVAlignCenter;
								ExcelSheet.Cells[17 + i, 2 + j].Borders.Color = ColorTranslator.ToOle(Color.Black);
							}
						}
						else
						{
							// Gérez le cas où DataTab n'a pas de lignes
							Console.WriteLine("DataTab n'a pas de lignes pour le code '" + DataTabs.Rows[i][0].ToString() + "'.");
						}
					}
					catch (Exception ex)
					{
						// Gestion des erreurs
						Console.WriteLine("Erreur SQL : " + ex.ToString());
					}
				}

				int l = 0;
				l = l + i;
				ExcelSheet.Cells[18 + l, 1] = "Date Reprise";
				ExcelSheet.Cells[18 + l, 1].Interior.Color = Color.Gold;
				ExcelSheet.Cells[18 + l, 1].Borders.Color = Color.Black;
				ExcelSheet.Cells[18 + l, 1].Font.Bold = true;
				ExcelSheet.Cells[18 + l, 1].Font.Size = 13;
				ExcelSheet.Cells[18 + l, 1].Font.Name = "Times New Roman";
				ExcelSheet.Cells[18 + l, 1].HorizontalAlignment = XlHAlign.xlHAlignRight;
				ExcelSheet.Cells[18 + l, 2].Interior.Color = Color.Gold;
				ExcelSheet.Cells[18 + l, 2].Borders.Color = Color.Black;
				ExcelSheet.Cells[18 + l, 2].VerticalAlignment = XlVAlign.xlVAlignCenter;
				ExcelSheet.Cells[18 + l, 2] = "Date d'arrêt";
				ExcelSheet.Cells[18 + l, 2].Font.Bold = true;
				ExcelSheet.Cells[18 + l, 2].Font.Size = 13;
				ExcelSheet.Cells[18 + l, 2].Font.Name = "Times New Roman";
				ExcelSheet.Cells[18 + l, 2].HorizontalAlignment = XlHAlign.xlHAlignRight;
				ExcelSheet.Cells[18 + l, 2].VerticalAlignment = XlVAlign.xlVAlignCenter;

				DataAdap = new System.Data.OleDb.OleDbDataAdapter("Select Date_Ordre_Service,Date_Arret_Service From Ordre_Service where Num_Décompte=" + NumDécompte, connection);
				DataTab = new System.Data.DataTable();
				DataAdap.Fill(DataTab);

				for (j = 0; j < DataTab.Rows.Count; j++)
				{
					ExcelSheet.Cells[19 + l + j, 1] = DataTab.Rows[j][0].ToString();
					ExcelSheet.Cells[19 + l + j, 1].Font.Size = 13;
					ExcelSheet.Cells[19 + l + j, 1].Font.Name = "Times New Roman";
					ExcelSheet.Cells[19 + l + j, 1].HorizontalAlignment = XlHAlign.xlHAlignRight;
					ExcelSheet.Cells[19 + l + j, 1].Borders.Color = Color.Black;
					ExcelSheet.Cells[19 + l + j, 1].VerticalAlignment = XlVAlign.xlVAlignCenter;

					ExcelSheet.Cells[19 + l + j, 2] = DataTab.Rows[j][1].ToString();
					ExcelSheet.Cells[19 + l + j, 2].Font.Size = 13;
					ExcelSheet.Cells[19 + l + j, 2].Font.Name = "Times New Roman";
					ExcelSheet.Cells[19 + l + j, 2].HorizontalAlignment = XlHAlign.xlHAlignRight;
					ExcelSheet.Cells[19 + l + j, 2].VerticalAlignment = XlVAlign.xlVAlignCenter;
					ExcelSheet.Cells[19 + l + j, 2].Borders.Color = Color.Black;
				}

				TotalRévisé = 0;

				//--------- dresser le tableau des révisions pour un seul décompte --------
				l = l + i;
				ExcelSheet.Cells[20 + l, 1] = "Date de réalisation";
				ExcelSheet.Columns[1].ColumnWidth = ExcelSheet.Cells[20 + l, 1].Value.ToString().Length * 1.3;
				ExcelSheet.Cells[20 + l, 1].Interior.Color = Color.Gold;
				ExcelSheet.Cells[20 + l, 1].Borders.Color = Color.Black;
				ExcelSheet.Cells[20 + l, 1].Font.Bold = true;
				ExcelSheet.Cells[20 + l, 1].Font.Size = 13;
				ExcelSheet.Cells[20 + l, 1].Font.Name = "Times New Roman";
				ExcelSheet.Cells[20 + l, 1].HorizontalAlignment = XlHAlign.xlHAlignRight;
				ExcelSheet.Cells[20 + l, 2].Interior.Color = Color.Gold;
				ExcelSheet.Cells[20 + l, 2].Borders.Color = Color.Black;
				ExcelSheet.Cells[20 + l, 2].VerticalAlignment = XlVAlign.xlVAlignCenter;
				ExcelSheet.Cells[20 + l, 2] = "Montant révisable";
				ExcelSheet.Columns[2].ColumnWidth = ExcelSheet.Cells[20 + l, 2].Value.ToString().Length * 1.3;
				ExcelSheet.Cells[20 + l, 2].Font.Bold = true;
				ExcelSheet.Cells[20 + l, 2].Font.Size = 13;
				ExcelSheet.Cells[20 + l, 2].Font.Name = "Times New Roman";
				ExcelSheet.Cells[20 + l, 2].HorizontalAlignment = XlHAlign.xlHAlignRight;
				ExcelSheet.Cells[20 + l, 2].VerticalAlignment = XlVAlign.xlVAlignCenter;
				ExcelSheet.Cells[20 + l, 3] = "Montant à réviser";
				ExcelSheet.Columns[3].ColumnWidth = ExcelSheet.Cells[20 + l, 3].Value.ToString().Length * 1.3;
				ExcelSheet.Cells[20 + l, 3].Interior.Color = Color.Gold;
				ExcelSheet.Cells[20 + l, 3].Borders.Color = Color.Black;
				ExcelSheet.Cells[20 + l, 3].Font.Bold = true;
				ExcelSheet.Cells[20 + l, 3].Font.Size = 13;
				ExcelSheet.Cells[20 + l, 3].Font.Name = "Times New Roman";
				ExcelSheet.Cells[20 + l, 3].HorizontalAlignment = XlHAlign.xlHAlignRight;
				ExcelSheet.Cells[20 + l, 4].Interior.Color = Color.Gold;
				ExcelSheet.Cells[20 + l, 4].Borders.Color = Color.Black;
				ExcelSheet.Cells[20 + l, 4].VerticalAlignment = XlVAlign.xlVAlignCenter;
				ExcelSheet.Cells[20 + l, 4] = "Période partielle";
				ExcelSheet.Columns[4].ColumnWidth = ExcelSheet.Cells[20 + l, 4].Value.ToString().Length * 1.3;
				ExcelSheet.Cells[20 + l, 4].Font.Bold = true;
				ExcelSheet.Cells[20 + l, 4].Font.Size = 13;
				ExcelSheet.Cells[20 + l, 4].Font.Name = "Times New Roman";
				ExcelSheet.Cells[20 + l, 4].HorizontalAlignment = XlHAlign.xlHAlignRight;
				ExcelSheet.Cells[20 + l, 4].VerticalAlignment = XlVAlign.xlVAlignCenter;
				ExcelSheet.Cells[20 + l, 5] = "Mois des travaux";
				ExcelSheet.Columns[5].ColumnWidth = ExcelSheet.Cells[20 + l, 5].Value.ToString().Length * 1.3;
				ExcelSheet.Cells[20 + l, 5].Interior.Color = Color.Gold;
				ExcelSheet.Cells[20 + l, 5].Borders.Color = Color.Black;
				ExcelSheet.Cells[20 + l, 5].Font.Bold = true;
				ExcelSheet.Cells[20 + l, 5].Font.Size = 13;
				ExcelSheet.Cells[20 + l, 5].Font.Name = "Times New Roman";
				ExcelSheet.Cells[20 + l, 5].HorizontalAlignment = XlHAlign.xlHAlignRight;
				ExcelSheet.Cells[20 + l, 6].Interior.Color = Color.Gold;
				ExcelSheet.Cells[20 + l, 6].Borders.Color = Color.Black;
				ExcelSheet.Cells[20 + l, 6].VerticalAlignment = XlVAlign.xlVAlignCenter;
				ExcelSheet.Cells[20 + l, 6] = "JoursMois/JoursTotal";
				ExcelSheet.Columns[6].ColumnWidth = ExcelSheet.Cells[20 + l, 6].Value.ToString().Length * 1.3;
				ExcelSheet.Cells[20 + l, 6].Font.Bold = true;
				ExcelSheet.Cells[20 + l, 6].Font.Size = 13;
				ExcelSheet.Cells[20 + l, 6].Font.Name = "Times New Roman";
				ExcelSheet.Cells[20 + l, 6].HorizontalAlignment = XlHAlign.xlHAlignRight;
				ExcelSheet.Cells[20 + l, 6].VerticalAlignment = XlVAlign.xlVAlignCenter;
				ExcelSheet.Cells[20 + l, 7] = "P/P0";
				ExcelSheet.Columns[7].ColumnWidth = ExcelSheet.Cells[20 + l, 7].Value.ToString().Length * 1.3;
				ExcelSheet.Cells[20 + l, 7].Interior.Color = Color.Gold;
				ExcelSheet.Cells[20 + l, 7].Borders.Color = Color.Black;
				ExcelSheet.Cells[20 + l, 7].Font.Bold = true;
				ExcelSheet.Cells[20 + l, 7].Font.Size = 13;
				ExcelSheet.Cells[20 + l, 7].Font.Name = "Times New Roman";
				ExcelSheet.Cells[20 + l, 7].HorizontalAlignment = XlHAlign.xlHAlignRight;
				ExcelSheet.Cells[20 + l, 8].Interior.Color = Color.Gold;
				ExcelSheet.Cells[20 + l, 8].Borders.Color = Color.Black;
				ExcelSheet.Cells[20 + l, 8].VerticalAlignment = XlVAlign.xlVAlignCenter;
				ExcelSheet.Cells[20 + l, 8] = "Montant de la révision";
				ExcelSheet.Columns[8].ColumnWidth = ExcelSheet.Cells[20 + l, 8].Value.ToString().Length * 1.3;
				ExcelSheet.Cells[20 + l, 8].Font.Bold = true;
				ExcelSheet.Cells[20 + l, 8].Font.Size = 13;
				ExcelSheet.Cells[20 + l, 8].Font.Name = "Times New Roman";
				ExcelSheet.Cells[20 + l, 8].HorizontalAlignment = XlHAlign.xlHAlignRight;
				ExcelSheet.Cells[20 + l, 8].VerticalAlignment = XlVAlign.xlVAlignCenter;

				double MRT = 0;
				double MR = 0;

				int s = 0;

				// Assurez-vous que l'indice s reste dans les limites du tableau
				while (s < TAB.GetLength(0) && !string.IsNullOrEmpty(TAB[s, 0]))
				{
					if (double.TryParse(TAB[s, 2], out double tabValue))
					{
						MRT += tabValue;
					}
					else
					{
						// Gérer le cas où la conversion échoue
						Console.WriteLine($"La conversion de TAB[{s}, 2] en double a échoué. Valeur actuelle : {TAB[s, 2]}");
					}

					s += 1;
				}

				for (i = 0; i < TM; i++)
				{
					if (i == 0)
					{
						dynamic range = ExcelSheet.Range["A" + (21 + l + i).ToString()];
						range.NumberFormat = "@";

						ExcelSheet.Cells[21 + l + i, 1] = DataTab.Rows[0][0].ToString();
						ExcelSheet.Cells[21 + l + i, 1].Borders.Color = Color.Black;
						ExcelSheet.Cells[21 + l + i, 1].Font.Size = 13;
						ExcelSheet.Cells[21 + l + i, 1].Font.Name = "Times New Roman";
						ExcelSheet.Cells[21 + l + i, 1].HorizontalAlignment = XlHAlign.xlHAlignRight;
						ExcelSheet.Cells[21 + l + i, 1].Borders.Color = Color.Black;
						ExcelSheet.Cells[21 + l + i, 1].VerticalAlignment = XlVAlign.xlVAlignCenter;

						DataAdap = new System.Data.OleDb.OleDbDataAdapter("Select Montant From Décompte where N°=" + NumDécompte, connection);
						DataTab = new System.Data.DataTable();
						DataAdap.Fill(DataTab);

						ExcelSheet.Cells[21 + l + i, 2].Borders.Color = Color.Black;
						ExcelSheet.Cells[21 + l + i, 2] = Tronquer(MRT, 2);
						ExcelSheet.Cells[21 + l + i, 2].Font.Size = 13;
						ExcelSheet.Cells[21 + l + i, 2].Font.Name = "Times New Roman";
						ExcelSheet.Cells[21 + l + i, 2].HorizontalAlignment = XlHAlign.xlHAlignRight;
						ExcelSheet.Cells[21 + l + i, 2].VerticalAlignment = XlVAlign.xlVAlignCenter;
					}
					else
					{
						ExcelSheet.Cells[21 + l + i, 1].Borders.Color = Color.Black;
						ExcelSheet.Cells[21 + l + i, 2].Borders.Color = Color.Black;
					}

					ExcelSheet.Cells[21 + l + i, 3] = Tronquer(double.Parse(TAB[i, 2] ?? "0"), 2);

					ExcelSheet.Cells[21 + l + i, 3].Borders.Color = Color.Black;
					ExcelSheet.Cells[21 + l + i, 3].Font.Size = 13;
					ExcelSheet.Cells[21 + l + i, 3].Font.Name = "Times New Roman";
					ExcelSheet.Cells[21 + l + i, 3].HorizontalAlignment = XlHAlign.xlHAlignRight;
					ExcelSheet.Cells[21 + l + i, 4].Borders.Color = Color.Black;
					ExcelSheet.Cells[21 + l + i, 4].VerticalAlignment = XlVAlign.xlVAlignCenter;
					ExcelSheet.Cells[21 + l + i, 4] = TAB[i, 1];
					ExcelSheet.Cells[21 + l + i, 4].Font.Size = 13;
					ExcelSheet.Cells[21 + l + i, 4].Font.Name = "Times New Roman";
					ExcelSheet.Cells[21 + l + i, 4].HorizontalAlignment = XlHAlign.xlHAlignRight;
					ExcelSheet.Cells[21 + l + i, 4].VerticalAlignment = XlVAlign.xlVAlignCenter;

					ExcelSheet.Cells[21 + l + i, 5] = TAB[i, 0];
					ExcelSheet.Cells[21 + l + i, 5].Borders.Color = Color.Black;
					ExcelSheet.Cells[21 + l + i, 5].Font.Size = 13;
					ExcelSheet.Cells[21 + l + i, 5].Font.Name = "Times New Roman";
					ExcelSheet.Cells[21 + l + i, 5].HorizontalAlignment = XlHAlign.xlHAlignRight;

					ExcelSheet.Cells[21 + l + i, 6].Borders.Color = Color.Black;
					ExcelSheet.Cells[21 + l + i, 6].VerticalAlignment = XlVAlign.xlVAlignCenter;
					ExcelSheet.Cells[21 + l + i, 6] = Tronquer(long.Parse(TAB[i, 1]) / (double)NJDP, 4);
					ExcelSheet.Cells[21 + l + i, 6].Font.Size = 13;
					ExcelSheet.Cells[21 + l + i, 6].Font.Name = "Times New Roman";
					ExcelSheet.Cells[21 + l + i, 6].HorizontalAlignment =XlHAlign.xlHAlignRight;
					ExcelSheet.Cells[21 + l + i, 6].VerticalAlignment = XlVAlign.xlVAlignCenter;

					double value = 0;
					object cellValue = ExcelSheet.Cells[21 + l + i, 7].Value;
					if (cellValue != null)
					{
						ExcelSheet.Cells[21 + l + i, 7] = Tronquer((double)(Convert.ToSingle(TAB[i, 3]) / double.Parse(TAB[i, 2])), 4);
						string cellStringValue = ExcelSheet.Cells[21 + l + i, 7].Value.ToString();

						if (ExcelSheet.Columns[7].ColumnWidth < cellStringValue.Length * 1.3)
						{
							ExcelSheet.Columns[7].ColumnWidth = cellStringValue.Length * 1.3;
						}

						ExcelSheet.Cells[21 + l + i, 7].Borders.Color = Color.Black;
						ExcelSheet.Cells[21 + l + i, 7].Font.Size = 13;
						ExcelSheet.Cells[21 + l + i, 7].Font.Name = "Times New Roman";
						ExcelSheet.Cells[21 + l + i, 7].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
						ExcelSheet.Cells[21 + l + i, 8].Borders.Color = Color.Black;
						ExcelSheet.Cells[21 + l + i, 8].VerticalAlignment = XlVAlign.xlVAlignCenter;

						value = Tronquer(MRT * double.Parse(ExcelSheet.Cells[21 + l + i, 6].Value.ToString()) * (double.Parse(cellStringValue) - 1), 2);
						TotalRévisé = TotalRévisé + value;
					}
					else
					{
						// Gérer le cas où la valeur de la cellule [21 + l + i, 7] est null.
						Console.WriteLine("La valeur de la cellule [21 + l + i, 7] est null.");
					}

					ExcelSheet.Cells[21 + l + i, 8] = value;
					ExcelSheet.Cells[21 + l + i, 8].Font.Size = 13;
					ExcelSheet.Cells[21 + l + i, 8].Font.Name = "Times New Roman";
					ExcelSheet.Cells[21 + l + i, 8].HorizontalAlignment = XlHAlign.xlHAlignRight;
					ExcelSheet.Cells[21 + l + i, 8].VerticalAlignment = XlVAlign.xlVAlignCenter;
				}

				ExcelSheet.Range("A" + (21 + l).ToString() + ":A" + (21 + l + i - 1).ToString()).Merge();
				ExcelSheet.Range("B" + (21 + l).ToString() + ":B" + (21 + l + i - 1).ToString()).Merge();
				l = l + i;
				ExcelSheet.Cells[21 + l, 1].Interior.Color = Color.Gold;
				ExcelSheet.Cells[21 + l, 1].Borders.Color = Color.Black;
				ExcelSheet.Cells[21 + l, 1].VerticalAlignment = XlVAlign.xlVAlignCenter;
				ExcelSheet.Cells[21 + l, 1] = "TOTAL REVISION DES PRIX H.T";
				ExcelSheet.Cells[21 + l, 1].Font.Bold = true;
				ExcelSheet.Cells[21 + l, 1].Font.Size = 13;
				ExcelSheet.Cells[21 + l, 1].Font.Name = "Times New Roman";
				ExcelSheet.Cells[21 + l, 1].HorizontalAlignment =XlHAlign.xlHAlignRight;
				ExcelSheet.Cells[21 + l, 1].VerticalAlignment = XlVAlign.xlVAlignCenter;

				ExcelSheet.Cells[21 + l, 8].Interior.Color = Color.Gold;
				ExcelSheet.Cells[21 + l, 8].Borders.Color = Color.Black;
				ExcelSheet.Cells[21 + l, 8].VerticalAlignment = XlVAlign.xlVAlignCenter;
				ExcelSheet.Cells[21 + l, 8] = Tronquer(TotalRévisé, 2);
				ExcelSheet.Cells[21 + l, 8].Font.Bold = true;
				ExcelSheet.Cells[21 + l, 8].Font.Size = 13;
				ExcelSheet.Cells[21 + l, 8].Font.Name = "Times New Roman";
				ExcelSheet.Cells[21 + l, 8].HorizontalAlignment = XlHAlign.xlHAlignRight;
				ExcelSheet.Cells[21 + l, 8].VerticalAlignment = XlVAlign.xlVAlignCenter;

				ExcelSheet.Range("A" + (21 + l).ToString() + ":G" + (21 + l).ToString()).Merge();
				for (j = 1; j <= 7; j++)
				{
					ExcelSheet.Cells[21 + l, j].Borders.Color = Color.Black;
				}

				ExcelSheet.Cells[21 + l + 1, 1].Interior.Color = Color.Gold;
				ExcelSheet.Cells[21 + l + 1, 1].Borders.Color = Color.Black;
				ExcelSheet.Cells[21 + l + 1, 1].VerticalAlignment = XlVAlign.xlVAlignCenter;
				ExcelSheet.Cells[21 + l + 1, 1] = "TVA (20%)";
				ExcelSheet.Cells[21 + l + 1, 1].Font.Bold = true;
				ExcelSheet.Cells[21 + l + 1, 1].Font.Size = 13;
				ExcelSheet.Cells[21 + l + 1, 1].Font.Name = "Times New Roman";
				ExcelSheet.Cells[21 + l + 1, 1].HorizontalAlignment = XlHAlign.xlHAlignRight;
				ExcelSheet.Cells[21 + l + 1, 1].VerticalAlignment = XlVAlign.xlVAlignCenter;

				ExcelSheet.Cells[21 + l + 1, 8].Interior.Color = Color.Gold;
				ExcelSheet.Cells[21 + l + 1, 8].Borders.Color = Color.Black;
				ExcelSheet.Cells[21 + l + 1, 8].VerticalAlignment = XlVAlign.xlVAlignCenter;
				ExcelSheet.Cells[21 + l + 1, 8] = Tronquer(TotalRévisé * 0.2, 2);
				ExcelSheet.Cells[21 + l + 1, 8].Font.Bold = true;
				ExcelSheet.Cells[21 + l + 1, 8].Font.Size = 13;
				ExcelSheet.Cells[21 + l + 1, 8].Font.Name = "Times New Roman";
				ExcelSheet.Cells[21 + l + 1, 8].HorizontalAlignment = XlHAlign.xlHAlignRight;
				ExcelSheet.Cells[21 + l + 1, 8].VerticalAlignment = XlVAlign.xlVAlignCenter;

				ExcelSheet.Range("A" + (21 + l + 1).ToString() + ":G" + (21 + l + 1).ToString()).Merge();
				for (j = 1; j <= 7; j++)
				{
					ExcelSheet.Cells[21 + l + 1, j].Borders.Color = Color.Black;
				}

				ExcelSheet.Cells[21 + l + 2, 1].Interior.Color = ColorTranslator.ToOle(Color.Gold);
				ExcelSheet.Cells[21 + l + 2, 1].Borders.Color = ColorTranslator.ToOle(Color.Black);
				ExcelSheet.Cells[21 + l + 2, 1].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
				ExcelSheet.Cells[21 + l + 2, 1].Value = "TOTAL REVISION DES PRIX TTC";
				ExcelSheet.Cells[21 + l + 2, 1].Font.Bold = true;
				ExcelSheet.Cells[21 + l + 2, 1].Font.Size = 13;
				ExcelSheet.Cells[21 + l + 2, 1].Font.Name = "Times New Roman";
				ExcelSheet.Cells[21 + l + 2, 1].HorizontalAlignment = XlHAlign.xlHAlignRight; // Remplacé HorizontalAlignment par XlHAlign
				ExcelSheet.Cells[21 + l + 2, 1].VerticalAlignment = XlVAlign.xlVAlignCenter;
				ExcelSheet.Cells[21 + l + 2, 8].Interior.Color = ColorTranslator.ToOle(Color.Gold);
				ExcelSheet.Cells[21 + l + 2, 8].Borders.Color = ColorTranslator.ToOle(Color.Black);
				ExcelSheet.Cells[21 + l + 2, 8].VerticalAlignment = XlVAlign.xlVAlignCenter;
				ExcelSheet.Cells[21 + l + 2, 8].Value = Tronquer(TotalRévisé * 1.2, 2);
				ExcelSheet.Cells[21 + l + 2, 8].Font.Bold = true;
				ExcelSheet.Cells[21 + l + 2, 8].Font.Size = 13;
				ExcelSheet.Cells[21 + l + 2, 8].Font.Name = "Times New Roman";
				ExcelSheet.Cells[21 + l + 2, 8].HorizontalAlignment = XlHAlign.xlHAlignRight; // Remplacé HorizontalAlignment par XlHAlign
				ExcelSheet.Cells[21 + l + 2, 8].VerticalAlignment = XlVAlign.xlVAlignCenter;
				ExcelSheet.Range["A" + (21 + l + 2).ToString() + ":G" + (21 + l + 2).ToString()].Merge();
				for (i = 1; i <= 7; i++)
				{
					ExcelSheet.cells(21 + l + 2, i).Borders.Color = Color.Black;
				}
				ExcelSheet.Cells[21 + l + 2, 1].Borders.Color = ColorTranslator.ToOle(Color.Black);
				ExcelSheet.Cells[21 + l + 4, 1].VerticalAlignment = XlVAlign.xlVAlignCenter;

				double valueToConvert = Convert.ToDouble(ExcelSheet.Cells[21 + l + 2, 8].Value);
				string convertedValue = valueToConvert.ToString(); // Convertir le double en string

				ExcelSheet.Cells[21 + l + 4, 1].Value = " Arrêté la présente note de calcul de la révision des prix à la somme TTC du :" + Calcul.ChiffreToLettre(convertedValue);

				ExcelSheet.Cells[21 + l + 4, 1].Font.Bold = true;
				ExcelSheet.Cells[21 + l + 4, 1].Font.Size = 13;
				ExcelSheet.Cells[21 + l + 4, 1].Font.Color = Color.DarkRed; // Remplacé font.color par Font.Color
				ExcelSheet.Cells[21 + l + 4, 1].Font.Name = "Times New Roman";
				ExcelSheet.Cells[21 + l + 4, 1].HorizontalAlignment = XlHAlign.xlHAlignLeft; // Remplacé HorizontalAlignment par XlHAlign
				ExcelSheet.Cells[21 + l + 4, 1].VerticalAlignment = XlVAlign.xlVAlignCenter;
				;

				//*********** save révision ********************************
				SQLSTR = "UPDATE Décompte SET ";
				SQLSTR = SQLSTR + "Montant_Révisé=" + (Convert.ToString(ExcelSheet.cells(21 + l + 2, 8).value)).Replace(",", ".");
				SQLSTR = SQLSTR + " where  Num_Marché='" + CmbNumMarchéM.Text + "' and N°=" + NumDécompte;
				ExecuterSQL(SQLSTR);
				//*******************************************************************
				connections.Close();
				connectionss.Close();
				ExcelApp.Visible = true;
				ExcelSheet = null;
				ExcelBook = null;
				ExcelApp = null;
			}
		}

        private void FrmRévision_FormClosed(object sender, FormClosedEventArgs e)
        {
			connection.Close();
		}

        private void FrmRévision_Load(object sender, EventArgs e)
        {
			if (connection.State == ConnectionState.Closed)
			{
				connection.ConnectionString = ("Provider=Microsoft.ACE.OLEDB.12.0;Data source=" + System.Windows.Forms.Application.StartupPath + "\\revision.accdb");
				connection.Open();
			}
			DGV.EnableHeadersVisualStyles = false;
			DGV.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
			DGV.ColumnHeadersDefaultCellStyle.BackColor = Color.DarkGray;

			DGVP.EnableHeadersVisualStyles = false;
			DGVP.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
			DGVP.ColumnHeadersDefaultCellStyle.BackColor = Color.DarkGray;
			recharger();
		}
		private void recharger()
		{
			DataAdap = new OleDbDataAdapter("Select Réference From Marché", connection);
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

        private void FrmRévision_Resize(object sender, EventArgs e)
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
		private static FrmRévision _DefaultInstance;
		public static FrmRévision DefaultInstance
		{
			get
			{
				if (_DefaultInstance == null || _DefaultInstance.IsDisposed)
					_DefaultInstance = new FrmRévision();

				return _DefaultInstance;
			}
		}
	}
}
