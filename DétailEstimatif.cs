using System;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace Révision
{
    public partial class DétailEstimatif : Form
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
        private long NumSymbole;
        private long NumActeur;
        public DétailEstimatif()
        {
            InitializeComponent();
        }
        private void BtnModifier_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(CmbNumMarché.Text))
                {
                    MessageBox.Show("Remplissez la zone : Référence de marché..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    CmbNumMarché.Focus();
                    return;
                }

                if (string.IsNullOrWhiteSpace(CmbUnitéM.Text))
                {
                    MessageBox.Show("Remplissez la zone Unité ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    CmbUnité.Focus();
                    return;
                }

                if (string.IsNullOrWhiteSpace(TxtPUM.Text))
                {
                    MessageBox.Show("Remplissez la zone Prix Unitaire ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    TxtPU.Focus();
                    return;
                }

                if (string.IsNullOrWhiteSpace(TxtQM.Text))
                {
                    MessageBox.Show("Remplissez la zone Quantité ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    TxtQ.Focus();
                    return;
                }

                if (string.IsNullOrWhiteSpace(TxtNumPrixM.Text))
                {
                    MessageBox.Show("Remplissez la zone Numéro de Prix ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    TxtNumPrixM.Focus();
                    return;
                }

                if (string.IsNullOrWhiteSpace(CmbDésignation.Text))
                {
                    MessageBox.Show("Remplissez la zone Numéro de Prix ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    TxtNumPrix.Focus();
                    return;
                }

                //------------ Modifier -------------
                SQLSTR = "UPDATE Detail_Estimatif SET ";
                //SQLSTR = SQLSTR + "Réference='" + CmbNumMarchéM.Text + "',"
                SQLSTR = SQLSTR + "Num_Prix='" + TxtNumPrixM.Text + "',";
                //SQLSTR = SQLSTR + "Désignation='" + CmbDésignation.Text.Replace("'", Chr(180)) + "',"
                SQLSTR = SQLSTR + "Quantité=" + TxtQM.Text + ",";
                SQLSTR = SQLSTR + "Unité='" + CmbUnitéM.Text + "',";
                SQLSTR = SQLSTR + "Prix=" + TxtPUM.Text + ",";
                SQLSTR = SQLSTR + "Symbole='" + CmbSymboleM.Text + "',";
                SQLSTR = SQLSTR + "PrixTotal=" + (double.Parse(TxtPUM.Text) * double.Parse(TxtQM.Text));
                SQLSTR = SQLSTR + " where Num_Marché='" + CmbNumMarché.Text + "' and Désignation = '" + CmbDésignation.Text + "'";
                //MessageBox.Show(SQLSTR);
                ExecuterSQL(SQLSTR);
                Actualiser();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
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
            if (DataTab.Rows.Count > 0)
            {
                CmbNumMarché.Text = DataTab.Rows[0][0].ToString();
            }

            DataAdap = new OleDbDataAdapter("Select Désignation From Detail_Estimatif where Num_Marché='" + CmbNumMarchéM.Text + "'", connection);
            DataTab = new DataTable();
            DataAdap.Fill(DataTab);
            if (DataTab.Rows.Count > 0)
            {
                for (int i = 0; i < DataTab.Rows.Count; i++)
                {
                    CmbDésignation.Items.Add(DataTab.Rows[i][0].ToString());
                }
            }

            DataAdap = new OleDbDataAdapter("Select Code From symbole", connection);
            DataTab = new DataTable();
            DataAdap.Fill(DataTab);
            for (int i = 0; i < DataTab.Rows.Count; i++)
            {
                CmbSymbole.Items.Add(DataTab.Rows[i][0].ToString());
                CmbSymboleM.Items.Add(DataTab.Rows[i][0].ToString());
            }

            CmbUnité.Items.Add("ML");
            CmbUnité.Items.Add("M2");
            CmbUnité.Items.Add("M3");
            CmbUnité.Items.Add("KG");
            CmbUnité.Items.Add("U");
            CmbUnité.Items.Add("ENS");
            CmbUnité.Items.Add("F");
            CmbUnité.Items.Add("T");
            CmbUnitéM.Items.Add("ML");
            CmbUnitéM.Items.Add("M2");
            CmbUnitéM.Items.Add("M3");
            CmbUnitéM.Items.Add("KG");
            CmbUnitéM.Items.Add("U");
            CmbUnitéM.Items.Add("ENS");
            CmbUnitéM.Items.Add("F");
            CmbUnitéM.Items.Add("T");
        }
        private void Actualiser()
        {
            CmbDésignation.Items.Clear();
            DataAdap = new OleDbDataAdapter("Select Désignation From Detail_Estimatif where Num_Marché='" + CmbNumMarchéM.Text + "'", connection);
            DataTab = new DataTable();
            DataAdap.Fill(DataTab);
            if (DataTab.Rows.Count > 0)
            {
                for (int i = 0; i < DataTab.Rows.Count; i++)
                {
                    CmbDésignation.Items.Add(DataTab.Rows[i][0].ToString());
                }
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
                if (string.IsNullOrWhiteSpace(CmbNumMarché.Text))
                {
                    MessageBox.Show("Remplissez la zone : Référence de marché..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    CmbNumMarché.Focus();
                    return;
                }

                if (string.IsNullOrWhiteSpace(CmbUnité.Text))
                {
                    MessageBox.Show("Remplissez la zone Unité ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    CmbUnité.Focus();
                    return;
                }

                if (string.IsNullOrWhiteSpace(TxtPU.Text))
                {
                    MessageBox.Show("Remplissez la zone Prix Unitaire ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    TxtPU.Focus();
                    return;
                }

                if (string.IsNullOrWhiteSpace(TxtQ.Text))
                {
                    MessageBox.Show("Remplissez la zone Quantité ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    TxtQ.Focus();
                    return;
                }

                if (string.IsNullOrWhiteSpace(TxtNumPrix.Text))
                {
                    MessageBox.Show("Remplissez la zone Numéro de Prix ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    TxtNumPrix.Focus();
                    return;
                }

                if (string.IsNullOrWhiteSpace(TxtDésignation.Text))
                {
                    MessageBox.Show("Remplissez la zone Désignation ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    TxtDésignation.Focus();
                    return;
                }

                //------------ Ajouter -------------
                SQLSTR = "INSERT INTO Detail_Estimatif(Num_Marché, Num_Prix, Désignation, Quantité, Unité, Prix, Symbole, PrixTotal) VALUES('";
                SQLSTR += CmbNumMarché.Text + "','";
                SQLSTR += TxtNumPrix.Text + "','";
                SQLSTR += TxtDésignation.Text.Replace("'", "\u00B4") + "',";
                SQLSTR += TxtQ.Text + ",'";
                SQLSTR += CmbUnité.Text + "',";
                SQLSTR += TxtPU.Text + "','";
                SQLSTR += CmbSymbole.Text + "',";
                SQLSTR += (double.Parse(TxtPU.Text) * double.Parse(TxtQ.Text)) + ")";

                //MessageBox.Show(SQLSTR);
                ExecuterSQL(SQLSTR);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void CmbDésignation_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(CmbDésignation.Text))
                {
                    MessageBox.Show("Choisissez une valeur de la liste : Désignation ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    TxtNumPrix.Focus();
                    return;
                }

                DataAdap = new OleDbDataAdapter("Select * From Detail_Estimatif where Num_Marché ='" + CmbNumMarchéM.Text + "' and Désignation = '" + CmbDésignation.Text + "'", connection);
                DataTab = new DataTable();
                DataAdap.Fill(DataTab);
                TxtNumPrixM.Text = DataTab.Rows[0][2].ToString();
                TxtQM.Text = DataTab.Rows[0][5].ToString();
                TxtPUM.Text = DataTab.Rows[0][7].ToString();
                CmbUnitéM.Text = DataTab.Rows[0][6].ToString();
                CmbSymboleM.Text = DataTab.Rows[0][4].ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void CmbNumMarchéM_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(CmbNumMarchéM.Text))
                {
                    MessageBox.Show("Choisissez une valeur de la liste : Référence de Marché ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    CmbNumMarchéM.Focus();
                    return;
                }
                CmbDésignation.Items.Clear();
                DataAdap = new OleDbDataAdapter("Select Désignation From Detail_Estimatif where Num_Marché='" + CmbNumMarchéM.Text + "'", connection);
                DataTab = new DataTable();
                DataAdap.Fill(DataTab);

                if (DataTab.Rows.Count > 0)
                {
                    for (int i = 0; i < DataTab.Rows.Count; i++)
                    {
                        CmbDésignation.Items.Add(DataTab.Rows[i][0].ToString());
                    }
                }
                else
                {
                    MessageBox.Show("Aucune valeur n'a été saisie dans la table détail éstimatif pour le marché numéro : " + CmbNumMarchéM.Text + ".", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void DétailEstimatif_FormClosed(object sender, FormClosedEventArgs e)
        {
            connection.Close();
        }
        private void DétailEstimatif_Load(object sender, EventArgs e)
        {
            if (connection.State == ConnectionState.Closed)
            {
                connection.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data source=" + Application.StartupPath + "\\revision.accdb";
                connection.Open();
            }
            Recharger();
        }
        private void DétailEstimatif_Resize(object sender, EventArgs e)
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
