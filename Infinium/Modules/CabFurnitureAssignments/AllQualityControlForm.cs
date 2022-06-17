using Infinium.Modules.CabFurnitureAssignments;

using System;
using System.Drawing;
using System.Windows.Forms;

namespace Infinium
{
    public partial class AllQualityControlForm : Form
    {
        const int eHide = 2;
        const int eShow = 1;
        const int eClose = 3;
        const int eMainMenu = 4;
        
        int FormEvent = 0;

        Form MainForm = null;

        StorePackagesManager storagePackagesManager;

        public AllQualityControlForm(Form tMainForm, StorePackagesManager SM)
        {
            MainForm = tMainForm;
            storagePackagesManager = SM;

            InitializeComponent();
            storagePackagesManager.GetAllQualityControl();
            Initialize();
        }

        private void ExitButton_Click(object sender, EventArgs e)
        {
            FormEvent = eClose;
            AnimateTimer.Enabled = true;
        }

        private void AnimateTimer_Tick(object sender, EventArgs e)
        {
            if (!DatabaseConfigsManager.Animation)
            {
                this.Opacity = 1;

                if (FormEvent == eClose || FormEvent == eHide)
                {
                    AnimateTimer.Enabled = false;

                    if (FormEvent == eClose)
                    {
                        this.Close();
                    }

                    if (FormEvent == eHide)
                    {
                        MainForm.Activate();
                        this.Hide();
                    }

                    return;
                }

                if (FormEvent == eShow)
                {
                    AnimateTimer.Enabled = false;
                    SplashForm.CloseS = true;
                    return;
                }

            }

            if (FormEvent == eClose || FormEvent == eHide)
            {
                if (Convert.ToDecimal(this.Opacity) != Convert.ToDecimal(0.00))
                    this.Opacity = Convert.ToDouble(Convert.ToDecimal(this.Opacity) - Convert.ToDecimal(0.05));
                else
                {
                    AnimateTimer.Enabled = false;

                    if (FormEvent == eClose)
                    {
                        this.Close();
                    }

                    if (FormEvent == eHide)
                    {
                        MainForm.Activate();
                        this.Hide();
                    }
                }

                return;
            }


            if (FormEvent == eShow || FormEvent == eShow)
            {
                if (this.Opacity != 1)
                    this.Opacity += 0.05;
                else
                {
                    AnimateTimer.Enabled = false;
                    SplashForm.CloseS = true;
                }

                return;
            }
        }

        private void Initialize()
        {
            dgvPackages.DataSource = storagePackagesManager.AllQualityControlBS;

            dgvPackages.AutoGenerateColumns = false;

            if (dgvPackages.Columns.Contains("PackagesCount"))
                dgvPackages.Columns["PackagesCount"].Visible = false;
            if (dgvPackages.Columns.Contains("TechStoreSubGroupID"))
                dgvPackages.Columns["TechStoreSubGroupID"].Visible = false;
            if (dgvPackages.Columns.Contains("CabFurAssignmentDetailID"))
                dgvPackages.Columns["CabFurAssignmentDetailID"].Visible = false;
            if (dgvPackages.Columns.Contains("MainOrderID"))
                dgvPackages.Columns["MainOrderID"].Visible = false;
            if (dgvPackages.Columns.Contains("QualityControlInUserID"))
                dgvPackages.Columns["QualityControlInUserID"].Visible = false;
            if (dgvPackages.Columns.Contains("QualityControlOutUserID"))
                dgvPackages.Columns["QualityControlOutUserID"].Visible = false;
            if (dgvPackages.Columns.Contains("CellID"))
                dgvPackages.Columns["CellID"].Visible = false;
            if (dgvPackages.Columns.Contains("PackNumber"))
                dgvPackages.Columns["PackNumber"].Visible = false;
            if (dgvPackages.Columns.Contains("QualityControl"))
                dgvPackages.Columns["QualityControl"].Visible = false;

            dgvPackages.Columns["CabFurniturePackageID"].HeaderText = "ID";
            dgvPackages.Columns["AddToStorageDateTime"].HeaderText = "Принято на склад";
            dgvPackages.Columns["RemoveFromStorageDateTime"].HeaderText = "Списано со склада";
            dgvPackages.Columns["Name"].HeaderText = "Ячейка";
            dgvPackages.Columns["QualityControlInDateTime"].HeaderText = "Отправлено на ОТК";
            dgvPackages.Columns["QualityControlOutDateTime"].HeaderText = "Принято с ОТК";
            dgvPackages.Columns["QualityControl"].HeaderText = "ОТК";

            foreach (DataGridViewColumn Column in dgvPackages.Columns)
            {
                Column.ReadOnly = true;
                //Column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }

            int DisplayIndex = 0;
            dgvPackages.Columns["CabFurniturePackageID"].DisplayIndex = DisplayIndex++;
            dgvPackages.Columns["AddToStorageDateTime"].DisplayIndex = DisplayIndex++;
            dgvPackages.Columns["RemoveFromStorageDateTime"].DisplayIndex = DisplayIndex++;
            dgvPackages.Columns["Name"].DisplayIndex = DisplayIndex++;
            dgvPackages.Columns["QualityControl"].DisplayIndex = DisplayIndex++;
            dgvPackages.Columns["QualityControlInDateTime"].DisplayIndex = DisplayIndex++;
            dgvPackages.Columns["QualityControlOutDateTime"].DisplayIndex = DisplayIndex++;

            dgvPackages.Columns["CabFurniturePackageID"].Width = 50;
        }
        
        private void BindPackagesToCellForm_Shown(object sender, EventArgs e)
        {

        }
        
        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (dgvPackages.Rows.Count == 0)
                return;


            bool OKCancel = Infinium.LightMessageBox.Show(ref MainForm, true,
                "Вы собираетесь принять упаковки с ОТК. Продолжить?",
                "ОТК");
            if (!OKCancel)
                return;

            int Index = dgvPackages.CurrentRow.Index;
            
            int[] packageIDs = new int[dgvPackages.SelectedRows.Count];

            for (int i = 0; i < dgvPackages.SelectedRows.Count; i++)
            {
                packageIDs[i] = Convert.ToInt32(dgvPackages.SelectedRows[i].Cells["CabFurniturePackageID"].Value);
            }
            storagePackagesManager.QualityControlOut(packageIDs);
            storagePackagesManager.GetAllQualityControl();
        }

        private void dgvPackages_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            if (e.RowIndex < 0)
                return;
        }
    }
}
