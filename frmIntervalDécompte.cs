using System;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace Révision
{
    public partial class frmIntervalDécompte : Form
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
        public frmIntervalDécompte()
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
                    MessageBox.Show("Remplissez la zone : Référence de marché..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    CmbNumMarché.Focus();
                    return;
                }

                if (DTPDPS.Value == null || string.IsNullOrEmpty(DTPDPS.Value.ToString()))
                {
                    MessageBox.Show("Choisissez une date de la zone : Date de Reprise de service ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    DTPDPS.Focus();
                    return;
                }

                if (string.IsNullOrEmpty(CmbDécompteA.Text))
                {
                    MessageBox.Show("Remplissez la zone Numéro de Décompte ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    CmbDécompte.Focus();
                    return;
                }

                DataAdap = new OleDbDataAdapter("Select * From Décompte where Num_Marché='" + CmbNumMarché.Text + "' and Désignation='" + CmbDécompteA.Text + "'", connection);
                DataTab = new DataTable();
                DataAdap.Fill(DataTab);

                if (DataTab.Rows.Count > 0)
                {
                    SQLSTR = "UPDATE Décompte SET ";
                    SQLSTR += "Date_Commencement='" + DTPDPS.Value.ToString() + "'";
                    SQLSTR += " where  N° = " + NumDécompte;
                    ExecuterSQL(SQLSTR);

                    CmbNumMarché.Text = "";
                    DTPDPS.Value = DateTime.Now;
                    CmbDécompteA.Text = "";
                }
                else
                {
                    MessageBox.Show("La valeur de Décompte suivant est saisie.", "Attention!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void frmIntervalDécompte_Load(object sender, EventArgs e)
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

            CmbDécompteA.Items.Clear();
            DataAdap = new OleDbDataAdapter("Select Désignation From Décompte where Num_Marché='" + CmbNumMarché.Text + "'", connection);
            DataTab = new DataTable();
            DataAdap.Fill(DataTab);

            if (DataTab.Rows.Count > 0)
            {
                for (int i = 0; i < DataTab.Rows.Count; i++)
                {
                    CmbDécompteA.Items.Add(DataTab.Rows[i][0].ToString());
                }
            }
            else
            {
                CmbDécompte.Text = "";
                DTPDPS.Value = DateTime.Now;
                MessageBox.Show("Aucune valeur n'a été saisie dans la table détail éstimatif pour le marché numéro : " + CmbNumMarchéM.Text + " .", "Attention!", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

                if (DTPDASM.Value == null || string.IsNullOrEmpty(DTPDASM.Value.ToString()))
                {
                    MessageBox.Show("Choisissez une date de la zone : Date d'arrêt de service ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    DTPDASM.Focus();
                    return;
                }

                if (string.IsNullOrEmpty(CmbDécompte.Text))
                {
                    MessageBox.Show("Remplissez la zone Numéro de Décompte ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    CmbDécompte.Focus();
                    return;
                }

                DataAdap = new OleDbDataAdapter("Select * From Décompte where Num_Marché='" + CmbNumMarchéM.Text + "' and Désignation='" + CmbDécompte.Text + "'", connection);
                DataTab = new DataTable();
                DataAdap.Fill(DataTab);

                if (DataTab.Rows.Count > 0)
                {
                    SQLSTR = "UPDATE Décompte SET ";
                    SQLSTR += "Date_Fin='" + DTPDASM.Value.ToString() + "'";
                    SQLSTR += " where  N° = " + NumDécompteM;
                }
                else
                {
                    MessageBox.Show("La valeur de Décompte suivant est saisie.", "Attention!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

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
                DTPDASM.Value = DateTime.Now;
                MessageBox.Show("Aucune valeur n'a été saisie dans la table détail éstimatif pour le marché numéro : " + CmbNumMarchéM.Text + ".", "Attention!", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                NumDécompteM = Convert.ToInt64(DataTab.Rows[0][0]);

                DataAdap = new OleDbDataAdapter("Select * From Ordre_Service where Num_Décompte=" + NumDécompteM, connection);
                DataTab = new DataTable();
                DataAdap.Fill(DataTab);

                if (DataTab.Rows.Count > 0)
                {
                    DTPDASM.Value = Convert.ToDateTime(DataTab.Rows[0][3].ToString());
                }
            }
        }

        private void CmbDécompteA_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(CmbDécompteA.Text))
            {
                MessageBox.Show("Choisissez une valeur de la liste : Numéro de décompte ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                CmbDécompte.Focus();
                return;
            }

            DataAdap = new OleDbDataAdapter("Select * From Décompte where Num_Marché='" + CmbNumMarché.Text + "' and Désignation='" + CmbDécompteA.Text + "'", connection);
            DataTab = new DataTable();
            DataAdap.Fill(DataTab);

            if (DataTab.Rows.Count > 0)
            {
                NumDécompte = Convert.ToInt64(DataTab.Rows[0][0]);
            }
        }

        private void frmIntervalDécompte_Resize(object sender, EventArgs e)
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
