﻿using Infinium.Modules.CabFurnitureAssignments;
using Infinium.Modules.Marketing.Expedition;
using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace Infinium
{
    public partial class CabFurAssembleForm : Form
    {
        private const int eHide = 2;
        private const int eShow = 1;
        private const int eClose = 3;
        private const int eMainMenu = 4;

        private int FormEvent = 0;

        private bool CanAction = true;

        private int ClientID = 1;
        private int MegaOrderID = 1;
        private int OrderNumber = 1;
        private object CreationDateTime;
        private object PrepareDispDateTime;

        private Form MainForm;

        private AssemblePackagesManager assemblePackagesManager;
        private AssignmentsManager assignmentsManager;
        private CabFurAssembleReport cabFurAssembleReport;

        [DllImport("user32.dll")]
        private static extern IntPtr GetActiveWindow();

        public CabFurAssembleForm(Form tMainForm, CabFurAssembleReport tCabFurAssembleReport, 
            int iClientID, int iMegaOrderID, int iOrderNumber, object CreationDateTime, object PrepareDispDateTime)
        {
            InitializeComponent();

            cabFurAssembleReport = tCabFurAssembleReport;
            OrderNumber = iOrderNumber;
            ClientID = iClientID;
            MegaOrderID = iMegaOrderID;
            MainForm = tMainForm;

            assignmentsManager = new AssignmentsManager();

            assignmentsManager.Initialize();

            Initialize();
            assemblePackagesManager.GetPackagesLabels(MegaOrderID);
            assemblePackagesManager.IsPackageCompleteToClient(MegaOrderID);
            lbScaned.Text = assemblePackagesManager.ScanedPackages;
            this.MaximumSize = Screen.PrimaryScreen.WorkingArea.Size;

            while (!SplashForm.bCreated) ;
        }
        public CabFurAssembleForm(Form tMainForm, AssignmentsManager tAssignmentsManager, int iClientID, int iMegaOrderID)
        {
            InitializeComponent();

            ClientID = iClientID;
            MegaOrderID = iMegaOrderID;
            MainForm = tMainForm;
            assignmentsManager = tAssignmentsManager;
            cabFurAssembleReport = new CabFurAssembleReport();

            Initialize();

            ToolTip ToolTip1 = new ToolTip();
            ToolTip1.SetToolTip(btnPrintScanReport, "Ведомость в Excel");

            assemblePackagesManager.GetPackagesLabels(ClientID, MegaOrderID);
            assemblePackagesManager.IsPackageCompleteToClient(MegaOrderID);
            lbScaned.Text = assemblePackagesManager.ScanedPackages;

            this.MaximumSize = Screen.PrimaryScreen.WorkingArea.Size;

            while (!SplashForm.bCreated) ;
        }

        private void CabFurDispatchForm_Shown(object sender, EventArgs e)
        {
            while (!SplashForm.bCreated) ;
            FormEvent = eShow;
            AnimateTimer.Enabled = true;

            panel3.BringToFront();

            //PhantomForm PhantomForm = new Infinium.PhantomForm();
            //PhantomForm.Show();

            //ReAutorizationForm ReAutorizationForm = new ReAutorizationForm(this);

            //TopForm = ReAutorizationForm;
            //ReAutorizationForm.ShowDialog();

            //PhantomForm.Close();

            //PhantomForm.Dispose();
            //bool PressOK = ReAutorizationForm.PressOK;
            //UserID = ReAutorizationForm.UserID;
            //ReAutorizationForm.Dispose();
            //TopForm = null;

            //if (PressOK)
            //{
            //    CheckLabel.UserID = UserID;
            //    CanAction = true;
            //}
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

        private void NavigateMenuCloseButton_Click(object sender, EventArgs e)
        {
            FormEvent = eClose;
            AnimateTimer.Enabled = true;
        }

        private void Initialize()
        {
            assemblePackagesManager = new AssemblePackagesManager();

            dgvNotScanedPackagesLabels.DataSource = assemblePackagesManager.NotScanedPackageLabelsBS;
            dgvNotcanedPackagesDetails.DataSource = assemblePackagesManager.NotScanedPackageDetailsBS;

            dgvAllScanedPackagesLabels.DataSource = assemblePackagesManager.AllScanedPackageLabelsBS;
            dgvAllScanedPackagesDetails.DataSource = assemblePackagesManager.AllScanedPackageDetailsBS;

            dgvScanedStoragePackagesDetails.DataSource = assemblePackagesManager.ScanedPackageDetailsBS;

            dgvPackagesLabelsSetting(ref dgvNotScanedPackagesLabels);
            dgvPackagesDetailsSetting(ref dgvNotcanedPackagesDetails);
            dgvPackagesLabelsSetting(ref dgvAllScanedPackagesLabels);
            dgvPackagesDetailsSetting(ref dgvAllScanedPackagesDetails);

            dgvPackagesDetailsSetting(ref dgvScanedStoragePackagesDetails);
            assemblePackagesManager.Clear();
        }

        private void dgvPackagesLabelsSetting(ref PercentageDataGrid grid)
        {
            grid.AutoGenerateColumns = false;

            if (grid.Columns.Contains("CabFurnitureComplementID"))
                grid.Columns["CabFurnitureComplementID"].Visible = false;
            if (grid.Columns.Contains("TechCatalogOperationsDetailID"))
                grid.Columns["TechCatalogOperationsDetailID"].Visible = false;
            if (grid.Columns.Contains("TechStoreSubGroupID"))
                grid.Columns["TechStoreSubGroupID"].Visible = false;
            if (grid.Columns.Contains("TechStoreID"))
                grid.Columns["TechStoreID"].Visible = false;
            if (grid.Columns.Contains("CoverID"))
                grid.Columns["CoverID"].Visible = false;
            if (grid.Columns.Contains("PatinaID"))
                grid.Columns["PatinaID"].Visible = false;
            if (grid.Columns.Contains("InsetColorID"))
                grid.Columns["InsetColorID"].Visible = false;
            if (grid.Columns.Contains("MainOrderID"))
                grid.Columns["MainOrderID"].Visible = false;
            if (grid.Columns.Contains("Notes"))
                grid.Columns["Notes"].Visible = false;
            if (grid.Columns.Contains("Scan"))
                grid.Columns["Scan"].Visible = false;

            if (grid.Columns.Contains("CabFurniturePackageID"))
                grid.Columns["CabFurniturePackageID"].HeaderText = "ID упаковки";
            if (grid.Columns.Contains("PackNumber"))
                grid.Columns["PackNumber"].HeaderText = "№ упаковки";
            if (grid.Columns.Contains("Index"))
                grid.Columns["Index"].HeaderText = "№ п/п";

            foreach (DataGridViewColumn Column in grid.Columns)
            {
                Column.ReadOnly = true;
                Column.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }

            int DisplayIndex = 0;
            if (grid.Columns.Contains("Index"))
                grid.Columns["Index"].DisplayIndex = DisplayIndex++;
            if (grid.Columns.Contains("PackNumber"))
                grid.Columns["PackNumber"].DisplayIndex = DisplayIndex++;
            if (grid.Columns.Contains("CabFurniturePackageID"))
                grid.Columns["CabFurniturePackageID"].DisplayIndex = DisplayIndex++;
        }

        private void dgvPackagesDetailsSetting(ref PercentageDataGrid grid)
        {
            grid.Columns.Add(assignmentsManager.CTechStoreNameColumn);
            grid.Columns.Add(assignmentsManager.TechStoreNameColumn);
            grid.Columns.Add(assignmentsManager.CoverColumn);
            grid.Columns.Add(assignmentsManager.PatinaColumn);
            grid.Columns.Add(assignmentsManager.InsetColorColumn);
            grid.AutoGenerateColumns = false;

            if (grid.Columns.Contains("RackName"))
                grid.Columns["RackName"].Visible = false;
            if (grid.Columns.Contains("CellName"))
                grid.Columns["CellName"].Visible = false;
            if (grid.Columns.Contains("CabFurnitureComplementID"))
                grid.Columns["CabFurnitureComplementID"].Visible = false;
            if (grid.Columns.Contains("CabFurnitureComplenentDetailID"))
                grid.Columns["CabFurnitureComplenentDetailID"].Visible = false;
            if (grid.Columns.Contains("TechStoreSubGroupID"))
                grid.Columns["TechStoreSubGroupID"].Visible = false;
            if (grid.Columns.Contains("CTechStoreID"))
                grid.Columns["CTechStoreID"].Visible = false;
            if (grid.Columns.Contains("CoverID"))
                grid.Columns["CoverID"].Visible = false;
            if (grid.Columns.Contains("PatinaID"))
                grid.Columns["PatinaID"].Visible = false;
            if (grid.Columns.Contains("TechStoreID"))
                grid.Columns["TechStoreID"].Visible = false;
            if (grid.Columns.Contains("InsetColorID"))
                grid.Columns["InsetColorID"].Visible = false;
            if (grid.Columns.Contains("CabFurniturePackageDetailID"))
                grid.Columns["CabFurniturePackageDetailID"].Visible = false;
            if (grid.Columns.Contains("CabFurniturePackageID"))
                grid.Columns["CabFurniturePackageID"].Visible = false;
            if (grid.Columns.Contains("CreateUserID"))
                grid.Columns["CreateUserID"].Visible = false;
            if (grid.Columns.Contains("MainOrderID"))
                grid.Columns["MainOrderID"].Visible = false;
            if (grid.Columns.Contains("PackNumber"))
                grid.Columns["PackNumber"].Visible = false;
            if (grid.Columns.Contains("CNotes"))
                grid.Columns["CNotes"].Visible = false;
            if (grid.Columns.Contains("PackagesCount"))
                grid.Columns["PackagesCount"].Visible = false;

            grid.Columns["CreateDateTime"].HeaderText = "Дата создания";
            grid.Columns["PackNumber"].HeaderText = "№ упак.";
            grid.Columns["Notes"].HeaderText = "Примечание";
            grid.Columns["Length"].HeaderText = "Длина, мм";
            grid.Columns["Height"].HeaderText = "Высота, мм";
            grid.Columns["Width"].HeaderText = "Ширина, мм";
            grid.Columns["Count"].HeaderText = "Кол-во";

            foreach (DataGridViewColumn Column in grid.Columns)
            {
                Column.ReadOnly = true;
                //Column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }

            int DisplayIndex = 0;
            grid.Columns["CTechStoreNameColumn"].DisplayIndex = DisplayIndex++;
            grid.Columns["CoverColumn"].DisplayIndex = DisplayIndex++;
            grid.Columns["PatinaColumn"].DisplayIndex = DisplayIndex++;
            grid.Columns["InsetColorColumn"].DisplayIndex = DisplayIndex++;
            grid.Columns["Notes"].DisplayIndex = DisplayIndex++;
            grid.Columns["TechStoreNameColumn"].DisplayIndex = DisplayIndex++;
            grid.Columns["Length"].DisplayIndex = DisplayIndex++;
            grid.Columns["Height"].DisplayIndex = DisplayIndex++;
            grid.Columns["Width"].DisplayIndex = DisplayIndex++;
            grid.Columns["Count"].DisplayIndex = DisplayIndex++;
            grid.Columns["CreateDateTime"].DisplayIndex = DisplayIndex++;
        }

        private void BarcodeTextBox_Leave(object sender, EventArgs e)
        {
            CheckTimer.Enabled = true;
        }

        private void NavigateMenuCloseButton_MouseUp(object sender, MouseEventArgs e)
        {
            BarcodeTextBox.Focus();
        }

        private void CheckTimer_Tick(object sender, EventArgs e)
        {
            if (!BarcodeTextBox.Focused)
            {
                BarcodeTextBox.Focus();
            }
        }

        private void BarcodeTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (!CanAction)
                return;

            if (e.KeyCode == Keys.Enter)
            {
                System.Diagnostics.Stopwatch sw = new System.Diagnostics.Stopwatch();
                sw.Start();

                BarcodeLabel.Text = "";
                CheckPicture.Visible = false;

                BarcodeLabel.ForeColor = Color.FromArgb(240, 0, 0);

                if (BarcodeTextBox.Text.Length < 12)
                {
                    BarcodeTextBox.Clear();
                    BarcodeLabel.Text = "Неверный штрихкод";
                    CheckPicture.Visible = true;
                    CheckPicture.Image = Properties.Resources.cancel;
                    BarcodeLabel.ForeColor = Color.FromArgb(240, 0, 0);
                    return;
                }

                string Prefix = BarcodeTextBox.Text.Substring(0, 3);

                if (Prefix != "021")
                {
                    BarcodeTextBox.Clear();
                    BarcodeLabel.Text = "Это не штрихкод упаковки";
                    CheckPicture.Visible = true;
                    CheckPicture.Image = Properties.Resources.cancel;
                    BarcodeLabel.ForeColor = Color.FromArgb(240, 0, 0);
                    return;
                }

                int CabFurniturePackageID = Convert.ToInt32(BarcodeTextBox.Text.Substring(3, 9));

                if (assemblePackagesManager.IsPackageScan(CabFurniturePackageID))
                {
                    CabFurniturePackageID = -1;
                    BarcodeTextBox.Clear();
                    BarcodeLabel.Text = "Уже отсканировано";
                    CheckPicture.Visible = true;
                    CheckPicture.Image = Properties.Resources.cancel;
                    BarcodeLabel.ForeColor = Color.FromArgb(240, 0, 0);
                    return;
                }
                if (!assemblePackagesManager.IsPackageExist(CabFurniturePackageID))
                {
                    CabFurniturePackageID = -1;
                    BarcodeTextBox.Clear();
                    BarcodeLabel.Text = "Упаковки не существует";
                    CheckPicture.Visible = true;
                    CheckPicture.Image = Properties.Resources.cancel;
                    BarcodeLabel.ForeColor = Color.FromArgb(240, 0, 0);
                    return;
                }

                BarcodeLabel.Text = BarcodeTextBox.Text;
                BarcodeTextBox.Clear();

                if (assemblePackagesManager.ScanPackage(CabFurniturePackageID))
                {
                    CheckPicture.Visible = true;
                    CheckPicture.Image = Properties.Resources.OK;
                    BarcodeLabel.ForeColor = Color.FromArgb(82, 169, 24);
                }
                else
                {
                    CheckPicture.Visible = true;
                    CheckPicture.Image = Properties.Resources.cancel;
                    BarcodeLabel.ForeColor = Color.FromArgb(240, 0, 0);

                    BarcodeLabel.Text = "Упаковки не соответствует заказу";
                }
                lbScaned.Text = assemblePackagesManager.ScanedPackages;
                lbRackName.Text = assemblePackagesManager.RackName;
                lbCellName.Text = assemblePackagesManager.CellName;
            }

        }

        private void BarcodeTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!CanAction)
                return;
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
            else
            {
                if (BarcodeTextBox.Text.Length >= 12 && e.KeyChar != (char)8)
                    e.Handled = true;
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (!CanAction)
                return;
        }

        private void dgvPackagesLabels_SelectionChanged(object sender, EventArgs e)
        {
            if (assemblePackagesManager == null)
                return;
            int cabFurniturePackageID = 0;
            if (dgvNotScanedPackagesLabels.SelectedRows.Count != 0 && dgvNotScanedPackagesLabels.SelectedRows[0].Cells["CabFurnitureComplementID"].Value != DBNull.Value)
                cabFurniturePackageID = Convert.ToInt32(dgvNotScanedPackagesLabels.SelectedRows[0].Cells["CabFurnitureComplementID"].Value);
            assemblePackagesManager.FilterNotScanedPackagesDetails(cabFurniturePackageID);
        }

        private void btnComplete_Click(object sender, EventArgs e)
        {
            if (dgvAllScanedPackagesDetails.Rows.Count == 0)
            {
                InfiniumTips.ShowTip(this, 50, 85, "Упаковки не отсканированы", 2000);
                return;
            }

            Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref MainForm, "Обновление.\r\nПодождите..."); });
            T.Start();
            while (!SplashWindow.bSmallCreated) ;

            assemblePackagesManager.AssemblePackages(ClientID, MegaOrderID);
            InfiniumTips.ShowTip(this, 50, 85, "Упаковки подготовлены к отгрузке", 2000);

            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
        }

        private void kryptonButton1_Click(object sender, EventArgs e)
        {

        }

        private void dgvExcessPackagesLabels_SelectionChanged(object sender, EventArgs e)
        {

        }

        private void dgvAllScanedPackagesLabels_SelectionChanged(object sender, EventArgs e)
        {
            if (assemblePackagesManager == null)
                return;
            int cabFurniturePackageID = 0;
            if (dgvAllScanedPackagesLabels.SelectedRows.Count != 0 && dgvAllScanedPackagesLabels.SelectedRows[0].Cells["CabFurnitureComplementID"].Value != DBNull.Value)
                cabFurniturePackageID = Convert.ToInt32(dgvAllScanedPackagesLabels.SelectedRows[0].Cells["CabFurnitureComplementID"].Value);
            assemblePackagesManager.FilterAllScanedPackagesDetails(cabFurniturePackageID);
        }

        private void kryptonButton1_Click_1(object sender, EventArgs e)
        {
            assemblePackagesManager.FF();
        }

        private void btnPrintScanReport_Click(object sender, EventArgs e)
        {
            string PackagesReportName = string.Empty;
            OrderInfo orderInfo;
            orderInfo.MegaOrderID = MegaOrderID;
            orderInfo.OrderNumber = OrderNumber;
            orderInfo.ClientID = ClientID;

            Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref MainForm, "Создание ведомости.\r\nПодождите..."); });
            T.Start();
            while (!SplashWindow.bSmallCreated) ;

            cabFurAssembleReport.GetDispatchInfo(ref CreationDateTime, ref PrepareDispDateTime);
            cabFurAssembleReport.CurrentClientID = ClientID;
            cabFurAssembleReport.CurrentOrderInfo = orderInfo;
            cabFurAssembleReport.CurrentMegaOrders = new int[1] { MegaOrderID };
            cabFurAssembleReport.FillPackagesByMegaOrders();
            cabFurAssembleReport.CreateCabFurScanReport(assemblePackagesManager.ScanedPackagesToExport, true, ref PackagesReportName);

            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;

        }
    }
}
