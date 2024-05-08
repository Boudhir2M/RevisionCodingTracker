using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;
using System;
using System.Runtime.InteropServices;

namespace Révision
{
    public partial class frmIndex : Form
    {
        OleDbConnection connection = new OleDbConnection();
        OleDbConnection connections = new OleDbConnection();
        OleDbConnection connectionXCL = new OleDbConnection();
        OleDbCommand records;
        DataTable DataTab = new DataTable();
        OleDbDataAdapter DataAdap = new OleDbDataAdapter();
        DataTable DataTabs = new DataTable();
        OleDbDataAdapter DataAdaps = new OleDbDataAdapter();
        DataTable DataTabss = new DataTable();
        OleDbDataAdapter DataAdapss = new OleDbDataAdapter();
        OleDbCommandBuilder ComndBuld;
        DataSet DataSetTab = new DataSet();
        string SQLSTR;
        long CodeCatégorie;
        string Symbole;
        long NumSymbole;
        long NumIndexs;

        public frmIndex()
        {
            InitializeComponent();
        }

        private void frmIndex_FormClosed(object sender, FormClosedEventArgs e)
        {
            connection.Close();
        }

        private void frmIndex_Load(object sender, System.EventArgs e)
        {
            if (connection.State == ConnectionState.Closed)
            {
                connection.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data source=" + Application.StartupPath + "\\revision.accdb";
                connection.Open();
            }

            Recharger();
            CmbSymboleM.Text = "TR1";
        }
        private void Recharger()
        {
            CmbSymboleA.Items.Clear();
            CmbSymboleS.Items.Clear();
            CmbSymboleM.Items.Clear();

            using (OleDbDataAdapter dataAdapter = new OleDbDataAdapter("Select Distinct Code From symbole", connection))
            {
                DataTable dataTable = new DataTable();
                dataAdapter.Fill(dataTable);

                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    CmbSymboleA.Items.Add(dataTable.Rows[i]["Code"].ToString());
                    CmbSymboleS.Items.Add(dataTable.Rows[i]["Code"].ToString());
                    CmbSymboleM.Items.Add(dataTable.Rows[i]["Code"].ToString());
                }
            }
        }
        private void ExecuterSQL(string Strsql)
        {
            using (OleDbCommand cmd = connection.CreateCommand())
            {
                cmd.CommandText = Strsql;
                cmd.ExecuteNonQuery();
            }
        }

        private void BtnAjouter_Click(object sender, System.EventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(CmbSymboleA.Text))
                {
                    MessageBox.Show("Choisissez une Catégorie de la liste : Catégorie..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    CmbSymboleA.Focus();
                    return;
                }
                if (string.IsNullOrEmpty(TxtValeurA.Text))
                {
                    MessageBox.Show("Remplir la zone nouvelle catégorie ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    TxtValeurA.Focus();
                    return;
                }

                // -------------- Ajouter --------------
                string mois = DTPA.Value.ToString("yyyy-MM");
                SQLSTR = "SELECT Num FROM symbole WHERE Code = '" + CmbSymboleA.Text + "'";
                using (OleDbCommand cmd = new OleDbCommand(SQLSTR, connection))
                {
                    int symboleNum = Convert.ToInt32(cmd.ExecuteScalar());
                    SQLSTR = "SELECT COUNT(*) FROM Indexs WHERE Mois = '" + mois + "' AND Symbole = " + symboleNum;
                    cmd.CommandText = SQLSTR;
                    int count = Convert.ToInt32(cmd.ExecuteScalar());

                    if (count <= 1)
                    {
                        SQLSTR = "INSERT INTO Indexs(Symbole, Mois, Valeur) VALUES (";
                        SQLSTR += symboleNum + ", '" + mois + "', " + TxtValeurA.Text.Replace(",", ".") + ")";
                        ExecuterSQL(SQLSTR);
                        Recharger();
                    }
                    else
                    {
                        MessageBox.Show("Le symbole suivant est déjà saisi.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void CmbSymboleA_SelectedIndexChanged(object sender, EventArgs e)
        {
            OleDbDataAdapter dataAdap = new OleDbDataAdapter("Select Num From symbole where Code ='" + CmbSymboleA.Text + "'", connection);
            DataTable dataTab = new DataTable();
            dataAdap.Fill(dataTab);
            NumSymbole = (int)dataTab.Rows[0]["Num"];
        }

        private void BtnSupprimer_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(CmbSymboleS.Text))
            {
                MessageBox.Show("Choisissez un symbole de la liste : Symbole..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                CmbSymboleS.Focus();
                return;
            }

            SQLSTR = "delete from Indexs where Symbole = " + NumSymbole + " and Mois = '" + DTPS.Value.ToString("yyyy-MM-dd") + "'";
            ExecuterSQL(SQLSTR);
            Recharger();
        }

        private void CmbSymboleS_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataAdap = new OleDbDataAdapter("Select Num From symbole where Code = '" + CmbSymboleS.Text + "'", connection);
            DataTab = new DataTable();
            DataAdap.Fill(DataTab);
            NumSymbole = (long)DataTab.Rows[0]["Num"];
        }

        private void CmbSymboleM_SelectedIndexChanged(object sender, EventArgs e)
        {
            OleDbDataAdapter DataAdap = new OleDbDataAdapter("Select Num From symbole where Code ='" + CmbSymboleM.Text + "'", connection);
            DataTable DataTab = new DataTable();
            DataAdap.Fill(DataTab);
            int NumSymbole = Convert.ToInt32(DataTab.Rows[0][0]);

            string SQLSTR = "Select * From Indexs where Symbole =" + NumSymbole + " and Mois='" + DTPM.Value.ToString("dd/MM/yyyy") + "'";
            DataAdap = new OleDbDataAdapter(SQLSTR, connection);
            DataTab = new DataTable();
            DataAdap.Fill(DataTab);

            if (DataTab.Rows.Count > 0)
            {
                TxtValeurM.Text = DataTab.Rows[0][3].ToString();
                int NumIndexs = Convert.ToInt32(DataTab.Rows[0][0]);
            }
            else
            {
                MessageBox.Show("L'Information n'existe pas dans la base de données.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void BtnModifier_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(CmbSymboleM.Text))
            {
                MessageBox.Show("Choisissez un symbole de la liste : Symbole..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                CmbSymboleM.Focus();
                return;
            }
            if (string.IsNullOrEmpty(TxtValeurM.Text))
            {
                MessageBox.Show("Remplir la nouvelle Valeur ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                TxtValeurM.Focus();
                return;
            }

            SQLSTR = "Select Num From symbole where Code = '" + CmbSymboleM.Text + "'";
            DataAdap = new OleDbDataAdapter(SQLSTR, connection);
            DataTab = new DataTable();
            DataAdap.Fill(DataTab);

            for (int i = 0; i < DataTab.Rows.Count; i++)
            {
                SQLSTR = "UPDATE Indexs SET ";
                SQLSTR += "Valeur = " + TxtValeurM.Text.Replace(",", ".");
                SQLSTR += " WHERE Num = " + (long)DataTab.Rows[i][0];
                ExecuterSQL(SQLSTR);
            }

            Recharger();
        }
        public bool IsNumeric(string value)
        {
            double result;
            return double.TryParse(value, out result);
        }
        private void BtnExcel_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook excelBook = null;
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = null;
            int i, j, n, l;

            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|Excel files 97-2000-2003 (*.xls)|*.xls|All files (*.*)|*.*";
            openFileDialog.FilterIndex = 1;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                excelBook = excelApp.Workbooks.Open(openFileDialog.FileName);
                excelSheet = excelBook.Worksheets[1];

                Microsoft.Office.Interop.Excel.Range range = excelSheet.UsedRange;
                n = range.Columns.Count;

                PB.Maximum = n - 1;
                l = 0;

                for (i = 2; i <= n; i++)
                {
                    if (!DateTime.TryParse(range.Cells[1, i].Value.ToString(), out DateTime dateValue))
                    {
                        MessageBox.Show("La valeur : " + range.Cells[1, i].Value + " n'a pas la forme d'une date.. veuillez l'écrire à cette forme: 01/mm/aaaa\nRechargez le fichier après correction", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                }

                for (i = 2; i <= n; i++)
                {
                    for (j = 2; j <= 33; j++)
                    {
                        string symbole = range.Cells[j, 1].Value.ToString();
                        DataAdap = new OleDbDataAdapter("Select Num From symbole where Code = '" + symbole + "'", connection);
                        DataTab = new DataTable();
                        DataAdap.Fill(DataTab);

                        if (DataTab.Rows.Count == 0)
                        {
                            MessageBox.Show("Le symbole : " + symbole + " n'existe pas. Veuillez le corriger et recharger le fichier à nouveau.", "Attention!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        }
                        else
                        {
                            for (int k = 0; k < DataTab.Rows.Count; k++)
                            {
                                if (connections.State == ConnectionState.Closed)
                                {
                                    connection.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data source=" + Application.StartupPath + "\\revision.accdb";
                                    connection.Open();
                                }

                                DateTime dateValue = DateTime.Now; // Remplacez cette ligne par la date que vous souhaitez utiliser
                                SQLSTR = "Select * From Indexs where Symbole = " + DataTab.Rows[k][0] + " and Mois = '" + dateValue.ToString("dd/MM/yyyy") + "'";
                                DataAdaps = new OleDbDataAdapter(SQLSTR, connections);
                                DataTabs = new DataTable();
                                DataAdaps.Fill(DataTabs);

                                if (DataTabs.Rows.Count == 0)
                                {
                                    string cellValue = range.Cells[j, i].Value.ToString();

                                    if (string.IsNullOrEmpty(cellValue))
                                    {
                                        MessageBox.Show("La valeur est vide. Veuillez la corriger et recharger le fichier à nouveau.", "Attention!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                    }
                                    else
                                    {
                                        if (!IsNumeric(cellValue))
                                        {
                                            MessageBox.Show("La valeur : " + cellValue + " n'est pas un nombre. Veuillez la corriger et recharger le fichier à nouveau.", "Attention!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                        }
                                        else
                                        {
                                            if (cellValue.Contains(",") && (cellValue.Length <= 2 || cellValue.Length > 3))
                                            {
                                                MessageBox.Show("La valeur : " + cellValue + " n'est pas valide. Veuillez la corriger et recharger le fichier à nouveau.", "Attention!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                            }
                                            else
                                            {
                                                SQLSTR = "INSERT INTO Indexs(Symbole, Mois, Valeur) VALUES (";
                                                SQLSTR += DataTab.Rows[k][0] + ", '";
                                                SQLSTR += dateValue.ToString("dd/MM/yyyy") + "', ";
                                                SQLSTR += cellValue.Replace(",", ".") + ")";
                                                ExecuterSQL(SQLSTR);
                                                l++;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                    PB.Value = i - 1;
                }

                MessageBox.Show("Transfert de : " + l + " enregistrements.\nOn ne compte pas les autres enregistrements qui sont introduits avant.", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Information);

                excelApp.Visible = false;
                Marshal.ReleaseComObject(excelSheet);
                Marshal.ReleaseComObject(excelBook);
            }

            excelApp.Quit();
            Marshal.ReleaseComObject(excelApp);
        }
    }
}
