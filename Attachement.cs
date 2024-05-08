using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Diagnostics;
using System.Windows.Forms;
using System.Linq;
using System.Xml.Linq;

using System.Data;
using System.Data.OleDb;
using static System.Data.DataTable;

namespace Révision
{
    public partial class Attachement : Form
    {
        public Attachement()
        {
            InitializeComponent();
        }
		private OleDbConnection connection = new OleDbConnection();
		private OleDbCommand records;
		private System.Data.DataTable DataTab;
		private System.Data.OleDb.OleDbDataAdapter DataAdap;
		private System.Data.OleDb.OleDbCommandBuilder ComndBuld;
		private DataSet DataSetTab = new DataSet();
		private string SQLSTR;
		private long NumDécompte;
		private long NumDécompteM;
		private long NumIndexs;

		private void recharger()
		{
			DataAdap = new System.Data.OleDb.OleDbDataAdapter("Select Réference From Marché", connection);
			DataTab = new DataTable();
			DataAdap.Fill(DataTab);
			int i = 0;
			CmbNumMarché.Items.Clear();
			CmbNumMarchéM.Items.Clear();
			for (i = 0; i < DataTab.Rows.Count; i++)
			{
				CmbNumMarché.Items.Add(DataTab.Rows[i][0].ToString());
				CmbNumMarchéM.Items.Add(DataTab.Rows[i][0].ToString());
			}

		}
		private void ExecuterSQL(string Strsql)
		{
			OleDbCommand CMND = (OleDbCommand)connection.CreateCommand();
			CMND.CommandText = SQLSTR;
			CMND.ExecuteNonQuery();
		}

        private void BtnAjouter_Click(object sender, EventArgs e)
        {
			try
			{
				if (string.IsNullOrEmpty(CmbNumMarché.Text))
				{
					MessageBox.Show("Remplissez la zone : Réference de marché..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
					CmbNumMarché.Focus();
					return;
				}
				if (string.IsNullOrEmpty(TxtUnité.Text))
				{
					MessageBox.Show("Remplissez la zone Unité ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
					TxtUnité.Focus();
					return;
				}
				if (string.IsNullOrEmpty(TxtPU.Text))
				{
					MessageBox.Show("Remplissez la zone Prix Unitaire ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
					TxtPU.Focus();
					return;
				}
				if (string.IsNullOrEmpty(TxtQ.Text))
				{
					MessageBox.Show("Remplissez la zone Quantité ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
					TxtQ.Focus();
					return;
				}
				if (string.IsNullOrEmpty(TxtNumPrix.Text))
				{
					MessageBox.Show("Remplissez la zone Numéro de Prix ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
					TxtNumPrix.Focus();
					return;
				}
				if (string.IsNullOrEmpty(CmbDésignation.Text))
				{
					MessageBox.Show("Remplissez la zone Numéro de Prix ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
					CmbDésignation.Focus();
					return;
				}
				if (string.IsNullOrEmpty(CmbDécompte.Text))
				{
					MessageBox.Show("Remplissez la zone Numéro de Prix ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
					CmbDécompte.Focus();
					return;
				}
				//------------ Ajouter -------------
				DataAdap = new System.Data.OleDb.OleDbDataAdapter("Select N° From Attachement where Num_Décompte=" + NumDécompte + " and Désignation='" + CmbDésignation.Text + "'", connection);
				DataTab = new DataTable();
				DataAdap.Fill(DataTab);
				if (DataTab.Rows.Count == 0)
				{
					SQLSTR = "INSERT INTO Attachement(Num_Décompte,Quantité,Unité,Prix_Unitaire,Symbole,Désignation) VALUES(";
					SQLSTR = SQLSTR + NumDécompte + ",";
					SQLSTR = SQLSTR + TxtQ.Text + ",'";
					SQLSTR = SQLSTR + TxtUnité.Text + "',";
					SQLSTR = SQLSTR + TxtPU.Text + ",'";
					SQLSTR = SQLSTR + CmbSymbole.Text + "','";
					SQLSTR = SQLSTR + CmbDésignation.Text.Replace("'", (Microsoft.VisualBasic.Strings.Chr(180).ToString()).ToString()) + "')";
					//MsgBox(SQLSTR)
					ExecuterSQL(SQLSTR);
				}
				else
				{
					MessageBox.Show("La valeur de l'Attachement est saisi:" + "\r\n" + "Quantité: " + TxtQ.Text + " " + TxtUnité.Text, "Attention!", MessageBoxButtons.OK, MessageBoxIcon.Information);
				}

				return;
			}
			catch
			{
				goto message;
			}
		message:
			//INSTANT C# TASK: Calls to the VB 'Err' function are not converted by Instant C#:
			MessageBox.Show(Microsoft.VisualBasic.Information.Err().Description, "Information", MessageBoxButtons.OK);
		}

        private void CmbNumMarché_SelectedIndexChanged(object sender, EventArgs e)
        {
			if (string.IsNullOrEmpty(CmbNumMarché.Text))
			{
				MessageBox.Show("Choisissez une valeur de la liste : Réference de Marché ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
				CmbNumMarché.Focus();
				return;
			}
			CmbDécompte.Items.Clear();
			CmbDésignation.Items.Clear();
			DataAdap = new System.Data.OleDb.OleDbDataAdapter("Select Désignation From Décompte where Num_Marché='" + CmbNumMarché.Text + "'", connection);
			DataTab = new DataTable();
			DataAdap.Fill(DataTab);
			if (DataTab.Rows.Count > 0)
			{
				for (var i = 0; i < DataTab.Rows.Count; i++)
				{
					CmbDécompte.Items.Add(DataTab.Rows[i][0].ToString());
				}
			}
			else
			{
				MessageBox.Show("Aucun valeur n'ai saisi dans la table détail éstimatif pour le marché numéro : " + CmbNumMarché.Text + " .", "Attention!", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
			DataAdap = new System.Data.OleDb.OleDbDataAdapter("Select Désignation From Detail_Estimatif where Num_Marché='" + CmbNumMarché.Text + "'", connection);
			DataTab = new DataTable();
			DataAdap.Fill(DataTab);
			MessageBox.Show(DataTab.Rows.Count.ToString());
			for (var i = 0; i < DataTab.Rows.Count; i++)
			{
				CmbDésignation.Items.Add(DataTab.Rows[i][0].ToString());
			}
		}

        private void Attachement_FormClosed(object sender, FormClosedEventArgs e)
        {
			connection.Close();
		}

        private void Attachement_Load(object sender, EventArgs e)
        {
			if (connection.State == ConnectionState.Closed)
			{
				connection.ConnectionString = ("Provider=Microsoft.ACE.OLEDB.12.0;Data source=" + Application.StartupPath + "\\revision.accdb");
				connection.Open();
			}
			recharger();
			DataAdap = new System.Data.OleDb.OleDbDataAdapter("Select Code From symbole", connection);
			DataTab = new DataTable();
			DataAdap.Fill(DataTab);
			int i = 0;
			CmbSymbole.Items.Clear();
			CmbSymboleM.Items.Clear();
			for (i = 0; i < DataTab.Rows.Count; i++)
			{
				CmbSymbole.Items.Add(DataTab.Rows[i][0].ToString());
				CmbSymboleM.Items.Add(DataTab.Rows[i][0].ToString());
			}
		}

        private void CmbDécompte_SelectedIndexChanged(object sender, EventArgs e)
        {
			if (string.IsNullOrEmpty(CmbDécompte.Text))
			{
				MessageBox.Show("Choisissez une valeur de la liste : Réference de Marché ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
				CmbDécompte.Focus();
				return;
			}
			DataAdap = new System.Data.OleDb.OleDbDataAdapter("Select N° From Décompte where Num_Marché='" + CmbNumMarché.Text + "' and Désignation='" + CmbDécompte.Text + "'", connection);
			DataTab = new DataTable();
			DataAdap.Fill(DataTab);
			NumDécompte = Convert.ToInt64(DataTab.Rows[0][0]);
		}

        private void CmbDésignation_SelectedIndexChanged(object sender, EventArgs e)
        {
			if (string.IsNullOrEmpty(CmbDésignation.Text))
			{
				MessageBox.Show("Choisissez une valeur de la liste : Réference de Marché ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
				CmbDésignation.Focus();
				return;
			}
			DataAdap = new System.Data.OleDb.OleDbDataAdapter("Select * From Detail_Estimatif where Num_Marché='" + CmbNumMarché.Text + "' and Désignation='" + CmbDésignation.Text + "'", connection);
			DataTab = new DataTable();
			DataAdap.Fill(DataTab);
			TxtNumPrix.Text = DataTab.Rows[0][2].ToString();
			TxtPU.Text = DataTab.Rows[0][7].ToString();
			TxtUnité.Text = DataTab.Rows[0][6].ToString();
		}

        private void CmbNumMarchéM_SelectedIndexChanged(object sender, EventArgs e)
        {
			if (string.IsNullOrEmpty(CmbNumMarchéM.Text))
			{
				MessageBox.Show("Choisissez une valeur de la liste : Réference de Marché ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
				CmbNumMarchéM.Focus();
				return;
			}
			CmbDécompteM.Items.Clear();
			CmbDésignationM.Items.Clear();
			DataAdap = new System.Data.OleDb.OleDbDataAdapter("Select distinct Désignation From Décompte where Num_Marché='" + CmbNumMarchéM.Text + "'", connection);
			DataTab = new DataTable();
			DataAdap.Fill(DataTab);
			CmbDécompteM.Items.Clear();
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
			DataAdap = new System.Data.OleDb.OleDbDataAdapter("Select Désignation From Detail_Estimatif where Num_Marché='" + CmbNumMarchéM.Text + "'", connection);
			DataTab = new DataTable();
			DataAdap.Fill(DataTab);
			MessageBox.Show(DataTab.Rows.Count.ToString());
			for (var i = 0; i < DataTab.Rows.Count; i++)
			{
				CmbDésignationM.Items.Add(DataTab.Rows[i][0].ToString());
			}
		}

        private void CmbDécompteM_SelectedIndexChanged(object sender, EventArgs e)
        {
			if (string.IsNullOrEmpty(CmbDécompteM.Text))
			{
				MessageBox.Show("Choisissez une valeur de la liste : Réference de Marché ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
				CmbDécompte.Focus();
				return;
			}
			DataAdap = new System.Data.OleDb.OleDbDataAdapter("Select N° From Décompte where Num_Marché='" + CmbNumMarchéM.Text + "' and Désignation='" + CmbDécompteM.Text + "'", connection);
			DataTab = new DataTable();
			DataAdap.Fill(DataTab);
			NumDécompteM = Convert.ToInt64(DataTab.Rows[0][0]);
		}

        private void CmbDésignationM_SelectedIndexChanged(object sender, EventArgs e)
        {
			if (string.IsNullOrEmpty(CmbDésignationM.Text))
			{
				MessageBox.Show("Choisissez une valeur de la liste : Désignation ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
				CmbDésignationM.Focus();
				return;
			}
			DataAdap = new System.Data.OleDb.OleDbDataAdapter("Select * From Attachement where Num_Décompte=" + NumDécompteM + " and Désignation='" + CmbDésignationM.Text + "'", connection);
			DataTab = new DataTable();
			DataAdap.Fill(DataTab);
			if (DataTab.Rows.Count > 0)
			{
				TxtQM.Text = DataTab.Rows[0][2].ToString();
				CmbSymboleM.Text = DataTab.Rows[0][5].ToString();
			}
			else
			{
				MessageBox.Show("Aucun attachement ni saisi pour le décompte numéro: " + CmbDécompteM.Text + " et dont la désignation: " + CmbDésignationM.Text + ".", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
		}

        private void BtnModifier_Click(object sender, EventArgs e)
        {
			try
			{
				//
				if (string.IsNullOrEmpty(CmbNumMarchéM.Text))
				{
					MessageBox.Show("Remplissez la zone : Réference de marché..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
					CmbNumMarché.Focus();
					return;
				}
				if (string.IsNullOrEmpty(CmbDécompteM.Text))
				{
					MessageBox.Show("Remplissez la zone Numéro de Prix ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
					CmbDécompteM.Focus();
					return;
				}
				if (string.IsNullOrEmpty(TxtQM.Text))
				{
					MessageBox.Show("Remplissez la zone Quantité ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
					TxtQ.Focus();
					return;
				}
				if (string.IsNullOrEmpty(CmbDésignationM.Text))
				{
					MessageBox.Show("Remplissez la zone Numéro de Prix ..", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
					CmbDésignation.Focus();
					return;
				}

				//------------ Ajouter -------------
				DataAdap = new System.Data.OleDb.OleDbDataAdapter("Select N° From Attachement where Num_Décompte=" + NumDécompteM + " and Désignation='" + CmbDésignationM.Text + "'", connection);
				DataTab = new DataTable();
				DataAdap.Fill(DataTab);
				if (DataTab.Rows.Count > 0)
				{
					SQLSTR = "UPDATE Attachement SET ";
					SQLSTR = SQLSTR + "Quantité=" + TxtQM.Text + ",";
					SQLSTR = SQLSTR + "Symbole='" + CmbSymboleM.Text + "'";
					SQLSTR = SQLSTR + " where  Num_Décompte=" + NumDécompteM + " and Désignation='" + CmbDésignationM.Text + "'";
					//MsgBox(SQLSTR)
					ExecuterSQL(SQLSTR);
				}
				else
				{
					MessageBox.Show("La valeur de l'Attachement est saisi:" + "\r\n" + "Quantité: " + DataTab.Rows[0][2].ToString() + " " + DataTab.Rows[0][3].ToString(), "Attention!", MessageBoxButtons.OK, MessageBoxIcon.Information);
				}

				return;
			}
			catch
			{
				goto message;
			}
		message:
			//INSTANT C# TASK: Calls to the VB 'Err' function are not converted by Instant C#:
			MessageBox.Show(Microsoft.VisualBasic.Information.Err().Description, "Information", MessageBoxButtons.OK);
		}
		private static Attachement _DefaultInstance;
		public static Attachement DefaultInstance
		{
			get
			{
				if (_DefaultInstance == null || _DefaultInstance.IsDisposed)
					_DefaultInstance = new Attachement();

				return _DefaultInstance;
			}
		}
	}
}
