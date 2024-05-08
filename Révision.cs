using System;
using System.Windows.Forms;

namespace Révision
{
    public partial class Révision : Form
    {
        public string Chemin { get; } = Application.StartupPath;
        private int m_ChildFormNumber = 0;
        public Révision()
        {
            InitializeComponent();
            IsMdiContainer = true; // 
        }
        private void ShowNewForm(object sender, EventArgs e)
        {
            // Créez une nouvelle instance du formulaire enfant.
            Form ChildForm = new Form();

            // Configurez-la en tant qu'enfant de ce formulaire MDI avant de l'afficher.
            ChildForm.MdiParent = this;

            m_ChildFormNumber++;
            ChildForm.Text = "Fenêtre " + m_ChildFormNumber;

            ChildForm.Show();
        }
        private void OpenFile(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            openFileDialog.Filter = "Fichiers texte (*.txt)|*.txt|Tous les fichiers (*.*)|*.*";

            if (openFileDialog.ShowDialog(this) == DialogResult.OK)
            {
                string fileName = openFileDialog.FileName;
                // TODO : ajoutez le code ici pour ouvrir le fichier.
            }
        }
        private void SaveAsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            saveFileDialog.Filter = "Fichiers texte (*.txt)|*.txt|Tous les fichiers (*.*)|*.*";

            if (saveFileDialog.ShowDialog(this) == DialogResult.OK)
            {
                string fileName = saveFileDialog.FileName;
                // TODO : ajoutez le code ici pour enregistrer le contenu actuel du formulaire dans un fichier.
            }
        }
        private void ExitToolsStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void PrintSetupToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Marché marchéForm = new Marché(); // Créez une nouvelle instance du formulaire Marché
            marchéForm.MdiParent = this; // Configurez-la en tant qu'enfant de ce formulaire MDI
            marchéForm.Show(); // Affichez le formulaire Marché
            marchéForm.PNL.SelectedTab = marchéForm.TabAjout; // Sélectionnez l'onglet TabAjout (si nécessaire)
        }

        private void ChercherInformationsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Marché marchéForm = new Marché();
            marchéForm.MdiParent = this; // "this" fait référence au formulaire principal
            marchéForm.Show();
            marchéForm.PNL.SelectedTab = marchéForm.TabModification;
        }

        private void AjouterToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            DétailEstimatif détailEstimatifForm = new DétailEstimatif();
            détailEstimatifForm.MdiParent = this; // "this" fait référence au formulaire principal
            détailEstimatifForm.Show();
            détailEstimatifForm.PNL.SelectedTab = détailEstimatifForm.TabAjout;
        }

        private void ModifierToolStripMenuItem3_Click(object sender, EventArgs e)
        {
            DétailEstimatif détailEstimatifForm = new DétailEstimatif();
            détailEstimatifForm.MdiParent = this; // "this" fait référence au formulaire principal
            détailEstimatifForm.Show();
            détailEstimatifForm.PNL.SelectedTab = détailEstimatifForm.TabModification;
        }

        private void InformationsToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            FrmMarchéInfos marchéInfosForm = new FrmMarchéInfos();
            marchéInfosForm.MdiParent = this; // "this" fait référence au formulaire principal
            marchéInfosForm.Show();
        }

        private void SuiviToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FrmSuivi suiviForm = new FrmSuivi();
            suiviForm.MdiParent = this; // "this" fait référence au formulaire principal
            suiviForm.Show();
        }

        private void SelectAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Entrepreneur entrepreneurForm = new Entrepreneur();
            entrepreneurForm.MdiParent = this; // "this" fait référence au formulaire principal
            entrepreneurForm.Show();
            entrepreneurForm.PNL.SelectedTab = entrepreneurForm.TabAjout;
        }

        private void ChercherInformationsToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Entrepreneur entrepreneurForm = new Entrepreneur();
            entrepreneurForm.MdiParent = this; // "this" fait référence au formulaire principal
            entrepreneurForm.Show();
            entrepreneurForm.PNL.SelectedTab = entrepreneurForm.TabModification;
        }

        private void InformationsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmInfosEntrepreneur infosEntrepreneurForm = new frmInfosEntrepreneur();
            infosEntrepreneurForm.MdiParent = this; // "this" fait référence au formulaire principal
            infosEntrepreneurForm.Show();
        }

        private void AjouterToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FrmSymbole frmSymboleForm = new FrmSymbole();
            frmSymboleForm.MdiParent = this; // "this" fait référence au formulaire principal
            frmSymboleForm.Show();
            frmSymboleForm.PNL.SelectedTab = frmSymboleForm.Ajout;
        }

        private void ModifierToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            FrmSymbole frmSymboleForm = new FrmSymbole();
            frmSymboleForm.MdiParent = this; // "this" fait référence au formulaire principal
            frmSymboleForm.Show();
            frmSymboleForm.PNL.SelectedTab = frmSymboleForm.Modification;
        }

        private void SupprimerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FrmSymbole frmSymboleForm = new FrmSymbole();
            frmSymboleForm.MdiParent = this; // "this" fait référence au formulaire principal
            frmSymboleForm.Show();
            frmSymboleForm.PNL.SelectedTab = frmSymboleForm.Suppression;
        }

        private void IndexsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmIndex frmIndexForm = new frmIndex();
            frmIndexForm.MdiParent = this; // "this" fait référence au formulaire principal
            frmIndexForm.Show();
        }

        private void AjouterToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            Décompte décompteForm = new Décompte();
            décompteForm.MdiParent = this; // "this" fait référence au formulaire principal
            décompteForm.Show();
            décompteForm.PNL.SelectedTab = décompteForm.TabAjout;
        }

        private void ModifierToolStripMenuItem4_Click(object sender, EventArgs e)
        {
            Décompte décompteForm = new Décompte();
            décompteForm.MdiParent = this; // "this" fait référence au formulaire principal
            décompteForm.Show();
            décompteForm.PNL.SelectedTab = décompteForm.TabModification;
        }

        private void DéfinitifToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FrmDécompteDéfinitif frmDécompteDéfinitif = new FrmDécompteDéfinitif();
            frmDécompteDéfinitif.MdiParent = this; // "this" fait référence au formulaire principal
            frmDécompteDéfinitif.Show();
        }

        private void CommencerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmIntervalDécompte frmIntervalDécompte = new frmIntervalDécompte();
            frmIntervalDécompte.MdiParent = this; // "this" fait référence au formulaire principal
            frmIntervalDécompte.Show();
            frmIntervalDécompte.PNL.SelectedTab = frmIntervalDécompte.TabCommencement;
        }

        private void ClauserToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmIntervalDécompte frmIntervalDécompte = new frmIntervalDécompte();
            frmIntervalDécompte.MdiParent = this; // "this" fait référence au formulaire principal
            frmIntervalDécompte.Show();
            frmIntervalDécompte.PNL.SelectedTab = frmIntervalDécompte.TabFin;
        }

        private void TraiterLesDonnéesToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            FrmOrdreService frmOrdreService = new FrmOrdreService();
            frmOrdreService.MdiParent = this; // "this" fait référence au formulaire principal
            frmOrdreService.Show();
            frmOrdreService.PNL.SelectedTab = frmOrdreService.TabReprise;
        }

        private void ModifierToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            FrmOrdreService frmOrdreService = new FrmOrdreService();
            frmOrdreService.MdiParent = this; // "this" fait référence au formulaire principal
            frmOrdreService.Show();
            frmOrdreService.PNL.SelectedTab = frmOrdreService.TabArret;
        }

        private void OptionsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FrmEditAttachement frmEditAttachement = new FrmEditAttachement();
            frmEditAttachement.MdiParent = this; // "this" fait référence au formulaire principal
            frmEditAttachement.Show();
        }

        private void TraiterLesDonnéesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Attachement attachement = new Attachement();
            attachement.MdiParent = this; // "this" fait référence au formulaire principal
            attachement.Show();
            attachement.PNL.SelectedTab = attachement.TabAjout;
        }

        private void ModifierToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Attachement attachement = new Attachement();
            attachement.MdiParent = this; // "this" fait référence au formulaire principal
            attachement.Show();
            attachement.PNL.SelectedTab = attachement.TabModification;
        }

        private void EditionToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FrmRévision frmRévision = new FrmRévision();
            frmRévision.MdiParent = this; // "this" fait référence au formulaire principal
            frmRévision.Show();
        }

        private void CalculToolStripMenuItem_Click(object sender, EventArgs e)
        {
             FrmRévisionTous frmRévisionTous = new FrmRévisionTous();
        frmRévisionTous.MdiParent = this; // "this" fait référence au formulaire principal
        frmRévisionTous.Show();
        }
    }
}
