using ComponentFactory.Krypton.Toolkit;

using Infinium.Modules.CabFurnitureAssignments;

using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Reflection.Emit;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace Infinium
{
    public partial class CabFurnitureModuleForm : Form
    {
        private const int iAgreementRole = 98;
        private const int iAdminRole = 100;
        private const int iDispatchRole = 99;

        private const int eHide = 2;
        private const int eShow = 1;
        private const int eClose = 3;
        private const int eMainMenu = 4;

        private bool NeedSplash = false;

        private decimal NonAgreementCount = 0;
        private decimal AgreedCount = 0;
        private decimal OnProdCount = 0;
        private decimal InProdCount = 0;
        private decimal OnStoreCount = 0;
        private decimal OnExpCount = 0;

        private int FormEvent = 0;

        private bool CanAction = true;
        [DllImport("user32.dll")]
        private static extern IntPtr GetActiveWindow();

        private Form TopForm = null;
        private LightStartForm LightStartForm;
        private CabFurnitureCoversForm CabFurnitureCoversForm;

        private DetailsReport detailsReport;
        private CalculateMaterial calcManager;
        private ComplementLabel complementLabel;
        private PackageLabel packageLabel;
        private CellLabel cellLabel;
        private AssignmentsManager assignmentsManager;
        private ComplementsManager complementsManager;
        private PackagesManager packagesManager;
        private StorePackagesManager storagePackagesManager;

        private CabFurStorage cabFurStorage;
        private CabFurStorageToExcel cabFurStorageToExcel;
        private Infinium.Modules.CabFurnitureAssignments.CheckLabel CheckLabel;
        private bool isDateOn = false;

        public enum RoleTypes
        {
            OrdinaryRole = 0,
            AdminRole = 1,
            ResponsibleRole = 2,
            TechnologyRole = 3,
            ControlRole = 4,
            AgreementRole = 5,
            WorkerRole = 6,
            DispatchRole = 7
        }

        private RoleTypes userRole;

        public CabFurnitureModuleForm(LightStartForm tLightStartForm)
        {
            InitializeComponent();
            LightStartForm = tLightStartForm;

            this.MaximumSize = Screen.PrimaryScreen.WorkingArea.Size;

            ToolTip ToolTip1 = new ToolTip();
            ToolTip1.SetToolTip(btnAddCell, "Добавить ячейку");
            ToolTip1.SetToolTip(btnAddRack, "Добавить стеллаж");
            ToolTip1.SetToolTip(btnAddWorkShop, "Добавить цех");
            ToolTip1.SetToolTip(btnEditCell, "Редактировать");
            ToolTip1.SetToolTip(btnEditRack, "Редактировать");
            ToolTip1.SetToolTip(btnEditWorkShop, "Редактировать");
            ToolTip1.SetToolTip(btnBindPackages, "Привязать упаковки к ячейке");
            ToolTip1.SetToolTip(btnPrintCellLabel, "Распечатать этикетку");
            ToolTip1.SetToolTip(btnQualityControlIn, "Отправить на ОТК");
            ToolTip1.SetToolTip(btnQualityControlOut, "Принять с ОТК");
            ToolTip1.SetToolTip(btnStartInventory, "Инвентаризация");

            Initialize();

            //Infinium.Classes.PictureToExcel PictureToExcel = new Classes.PictureToExcel();

            assignmentsManager.GetPermissions(Security.CurrentUserID, this.Name);

            if (assignmentsManager.PermissionGranted(iAgreementRole))
            {
                userRole = RoleTypes.AgreementRole;
                cmiSaveAssignments.Visible = false;
                agreeAssignmentContextMenuItem.Visible = true;
                setDispatchDateContextMenuItem.Visible = false;
            }
            if (assignmentsManager.PermissionGranted(iAdminRole))
            {
                userRole = RoleTypes.AdminRole;
                cmiSaveAssignments.Visible = true;
                agreeAssignmentContextMenuItem.Visible = true;
                setDispatchDateContextMenuItem.Visible = true;
            }
            if (assignmentsManager.PermissionGranted(iDispatchRole))
            {
                userRole = RoleTypes.DispatchRole;
                cmiSaveAssignments.Visible = true;
                agreeAssignmentContextMenuItem.Visible = true;
                setDispatchDateContextMenuItem.Visible = true;
            }

            while (!SplashForm.bCreated) ;
        }

        private void CabFurnitureModuleForm_Shown(object sender, EventArgs e)
        {
            MenuPanel.BringToFront();

            while (!SplashForm.bCreated) ;
            FormEvent = eShow;
            AnimateTimer.Enabled = true;
            NeedSplash = true;
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
                        LightStartForm.CloseForm(this);
                        this.Close();
                    }

                    if (FormEvent == eHide)
                    {
                        LightStartForm.HideForm(this);
                    }


                    return;
                }

                if (FormEvent == eShow)
                {
                    AnimateTimer.Enabled = false;
                    SplashForm.CloseS = true;
                    NeedSplash = true;
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
                        LightStartForm.CloseForm(this);
                        this.Close();
                    }

                    if (FormEvent == eHide)
                    {
                        LightStartForm.HideForm(this);
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
                    NeedSplash = true;
                }

                return;
            }
        }

        private void NavigateMenuCloseButton_Click(object sender, EventArgs e)
        {
            if (CabFurnitureCoversForm != null)
                CabFurnitureCoversForm.Dispose();
            FormEvent = eClose;
            AnimateTimer.Enabled = true;
        }

        private void NavigateMenuBackButton_Click(object sender, EventArgs e)
        {
            FormEvent = eHide;
            AnimateTimer.Enabled = true;
        }

        private void Initialize()
        {
            cabFurStorage = new CabFurStorage();

            storagePackagesManager = new StorePackagesManager();
            cabFurStorageToExcel = new CabFurStorageToExcel();

            DateTime FirstDay = DateTime.Now.AddDays(-100);
            DateTime Today = DateTime.Now;

            kryptonDateTimePicker10.Value = Today;
            kryptonDateTimePicker9.Value = Today;
            DateTimePicker1.Value = FirstDay;
            CreateDateTimePicker1.Value = FirstDay;

            complementLabel = new ComplementLabel();
            packageLabel = new PackageLabel();

            assignmentsManager = new AssignmentsManager();

            assignmentsManager.Initialize();

            complementsManager = new ComplementsManager();
            packagesManager = new PackagesManager();


            CheckLabel = new Modules.CabFurnitureAssignments.CheckLabel();
            dgvScan.DataSource = CheckLabel.ScanContentBS;
            dgvScanDetailsSetting(ref dgvScan);

            cmbxTechStore.DataSource = assignmentsManager.TechStoreBS;
            cmbxTechStore.DisplayMember = "TechStoreName";
            cmbxTechStore.ValueMember = "TechStoreID";
            cmbxTechStoreSubGroups.DataSource = assignmentsManager.TechStoreSubGroupsBS;
            cmbxTechStoreSubGroups.DisplayMember = "TechStoreSubGroupName";
            cmbxTechStoreSubGroups.ValueMember = "TechStoreSubGroupID";
            cmbxTechStoreGroups.DataSource = assignmentsManager.TechStoreGroupsBS;
            cmbxTechStoreGroups.DisplayMember = "TechStoreGroupName";
            cmbxTechStoreGroups.ValueMember = "TechStoreGroupID";
            cmbxCovers.DataSource = assignmentsManager.CoversBS;
            cmbxCovers.DisplayMember = "CoverName";
            cmbxCovers.ValueMember = "CoverID";
            cmbxPatina.DataSource = assignmentsManager.PatinaBS;
            cmbxPatina.DisplayMember = "PatinaName";
            cmbxPatina.ValueMember = "PatinaID";
            cmbxInsetColors.DataSource = assignmentsManager.BasicInsetColorsBS;
            cmbxInsetColors.DisplayMember = "InsetColorName";
            cmbxInsetColors.ValueMember = "InsetColorID";


            cbTStoreSearch.DataSource = assignmentsManager.TechStoreBS;
            cbTStoreSearch.DisplayMember = "TechStoreName";
            cbTStoreSearch.ValueMember = "TechStoreID";
            cbTSSubGroupsSearch.DataSource = assignmentsManager.TechStoreSubGroupsBS;
            cbTSSubGroupsSearch.DisplayMember = "TechStoreSubGroupName";
            cbTSSubGroupsSearch.ValueMember = "TechStoreSubGroupID";
            cbTSGroupsSearch.DataSource = assignmentsManager.TechStoreGroupsBS;
            cbTSGroupsSearch.DisplayMember = "TechStoreGroupName";
            cbTSGroupsSearch.ValueMember = "TechStoreGroupID";
            cbCoversSearch.DataSource = assignmentsManager.CoversBS;
            cbCoversSearch.DisplayMember = "CoverName";
            cbCoversSearch.ValueMember = "CoverID";
            cbPatinaSearch.DataSource = assignmentsManager.PatinaBS;
            cbPatinaSearch.DisplayMember = "PatinaName";
            cbPatinaSearch.ValueMember = "PatinaID";
            cbInsetColorsSearch.DataSource = assignmentsManager.BasicInsetColorsBS;
            cbInsetColorsSearch.DisplayMember = "InsetColorName";
            cbInsetColorsSearch.ValueMember = "InsetColorID";

            dgvComplements.DataSource = complementsManager.ComplementsBS;
            dgvMainOrders.DataSource = complementsManager.MainOrdersBS;
            dgvComplementLabels.DataSource = complementsManager.ComplementLabelsBS;
            dgvComplementDetails.DataSource = complementsManager.ComplementDetailsBS;

            dgvPackages.DataSource = packagesManager.PackagesBS;
            dgvPackagesLabels.DataSource = packagesManager.PackageLabelsBS;
            dgvPackagesDetails.DataSource = packagesManager.PackageDetailsBS;

            dgvDocuments.DataSource = assignmentsManager.DocumentsBS;
            dgvNewAssignment.DataSource = assignmentsManager.NewAssignmentDetailsBS;
            percentageDataGrid1.DataSource = assignmentsManager.NewAssignmentDetailsBS;
            dgvAllAssignments.DataSource = assignmentsManager.AllAssignmentsBS;
            percentageDataGrid2.DataSource = assignmentsManager.StatisticsBS;

            assignmentsManager.NonAgreementOrders(true, ref NonAgreementCount);
            assignmentsManager.AgreedOrders(true, ref AgreedCount);
            assignmentsManager.OnProductionOrders(true, ref OnProdCount);
            assignmentsManager.InProductionOrders(true, ref InProdCount);
            assignmentsManager.OnStorageOrders(true, ref OnStoreCount);
            assignmentsManager.OnExpeditionOrders(true, ref OnExpCount);

            NonAgreementDataGrid.DataSource = assignmentsManager.NonAgreementDetailBS;
            AgreedDataGrid.DataSource = assignmentsManager.AgreedDetailBS;
            OnProductionDataGrid.DataSource = assignmentsManager.OnProductionDetailBS;
            InProductionDataGrid.DataSource = assignmentsManager.InProductionDetailBS;
            OnStorageDataGrid.DataSource = assignmentsManager.OnStorageDetailBS;
            OnExpeditionDataGrid.DataSource = assignmentsManager.OnExpeditionDetailBS;

            dgvComplementsSetting(ref dgvComplements);
            dgvMainOrdersSetting(ref dgvMainOrders);
            dgvComplementLabelsSetting(ref dgvComplementLabels);
            dgvComplementDetailsSetting(ref dgvComplementDetails);

            dgvPackagesSetting(ref dgvPackages);
            dgvPackagesLabelsSetting(ref dgvPackagesLabels);
            dgvPackagesDetailsSetting(ref dgvPackagesDetails);

            dgvDocumentsSetting(ref dgvDocuments);

            dgvDetailsSetting(true, ref NonAgreementDataGrid);
            dgvDetailsSetting(true, ref AgreedDataGrid);
            dgvDetailsSetting(true, ref OnProductionDataGrid);
            dgvDetailsSetting(true, ref InProductionDataGrid);
            dgvDetailsSetting(true, ref OnStorageDataGrid);
            dgvDetailsSetting(true, ref OnExpeditionDataGrid);

            dgvNewAssignmentSetting(ref dgvNewAssignment);
            dgvNewAssignmentSetting(ref percentageDataGrid1);
            dgvAllAssignmentsSetting(ref dgvAllAssignments);
            dgvStatisticsSetting(ref percentageDataGrid2);

            assignmentsManager.UpdateNewAssignment(0);
            assignmentsManager.UpdateDocuments(FirstDay, DateTime.Now);
            assignmentsManager.UpdateAllAssignments(FirstDay, DateTime.Now);
            assignmentsManager.FilterAllAssignments(true, true, true, true, true);


            label11.Text = NonAgreementCount.ToString() + " шт.";
            label6.Text = AgreedCount.ToString() + " шт.";
            label27.Text = OnProdCount.ToString() + " шт.";
            label16.Text = InProdCount.ToString() + " шт.";
            label41.Text = OnStoreCount.ToString() + " шт.";
            label34.Text = OnExpCount.ToString() + " шт.";

            BarcodeLabel.Text = "";
            CheckPicture.Visible = false;

            BarcodeLabel.ForeColor = Color.FromArgb(240, 0, 0);

            ClearControls();

            CheckLabel.Clear();


            cellLabel = new CellLabel();
            CabStorageSetting();

            dgvStoragePackagesLabels.DataSource = storagePackagesManager.PackageLabelsBS;
            dgvStoragePackagesDetails.DataSource = storagePackagesManager.PackageDetailsBS;
            dgvCellsSetting(ref dgvCells);
            dgvStoragePackagesLabelsSetting(ref dgvStoragePackagesLabels);
            dgvPackagesDetailsSetting(ref dgvStoragePackagesDetails);
        }

        private void dgvStatisticsSetting(ref PercentageDataGrid grid)
        {
            grid.Columns.Add(assignmentsManager.CTechStoreNameColumn);
            grid.Columns.Add(assignmentsManager.CoverColumn);
            grid.Columns.Add(assignmentsManager.PatinaColumn);
            grid.Columns.Add(assignmentsManager.InsetColorColumn);
            grid.AutoGenerateColumns = false;

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
            if (grid.Columns.Contains("CreateDateTime"))
                grid.Columns["CreateDateTime"].Visible = false;

            if (grid.Columns.Contains("ACreationDateTime"))
                grid.Columns["ACreationDateTime"].Visible = false;
            if (grid.Columns.Contains("AProductionDateTime"))
                grid.Columns["AProductionDateTime"].Visible = false;







            if (grid.Columns.Contains("CounterCFA"))
                grid.Columns["CounterCFA"].Visible = false;

            if (grid.Columns.Contains("PrintDateTime"))
                grid.Columns["PrintDateTime"].Visible = false;
            //if (grid.Columns.Contains("AddToStorageDateTime"))
            //    grid.Columns["AddToStorageDateTime"].Visible = false;
            if (grid.Columns.Contains("RemoveFromStorageDateTime"))
                grid.Columns["RemoveFromStorageDateTime"].Visible = false;
            if (grid.Columns.Contains("QualityControlInDateTime"))
                grid.Columns["QualityControlInDateTime"].Visible = false;
            if (grid.Columns.Contains("QualityControlOutDateTime"))
                grid.Columns["QualityControlOutDateTime"].Visible = false;
            if (grid.Columns.Contains("ValOutProductionDateTime"))
                grid.Columns["ValOutProductionDateTime"].Visible = false;








            grid.Columns["CabFurAssignmentID"].HeaderText = "№ задания";
            grid.Columns["AccountingName"].HeaderText = "Бухгалтерское наименование детали";
            grid.Columns["Cost"].HeaderText = "Стоимость";
            grid.Columns["AddToStorageDateTime"].HeaderText = "Дата принятия на склад";
            grid.Columns["Notes"].HeaderText = "Примечание";
            grid.Columns["Count"].HeaderText = "Кол-во";
            grid.Columns["InvNumber"].HeaderText = "Инвентарный номер";



            foreach (DataGridViewColumn Column in grid.Columns)
            {
                Column.ReadOnly = true;
                //Column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }

            int DisplayIndex = 0;
            grid.Columns["CabFurAssignmentID"].DisplayIndex = DisplayIndex++;
            grid.Columns["CTechStoreNameColumn"].DisplayIndex = DisplayIndex++;
            grid.Columns["CoverColumn"].DisplayIndex = DisplayIndex++;
            grid.Columns["PatinaColumn"].DisplayIndex = DisplayIndex++;
            grid.Columns["InsetColorColumn"].DisplayIndex = DisplayIndex++;
            grid.Columns["Notes"].DisplayIndex = DisplayIndex++;
            grid.Columns["InvNumber"].DisplayIndex = DisplayIndex++;
            grid.Columns["AccountingName"].DisplayIndex = DisplayIndex++;
            grid.Columns["Count"].DisplayIndex = DisplayIndex++;
            grid.Columns["Cost"].DisplayIndex = DisplayIndex++;
            grid.Columns["ValOutProductionDateTime"].DisplayIndex = DisplayIndex++;

        }

        private void CabStorageSetting()
        {
            cmbxWorkShops.DataSource = cabFurStorage.workShopsBs;
            cmbxWorkShops.DisplayMember = "Name";
            cmbxWorkShops.ValueMember = "WorkShopID";

            cmbxRacks.DataSource = cabFurStorage.racksBs;
            cmbxRacks.DisplayMember = "Name";
            cmbxRacks.ValueMember = "RackID";

            dgvCells.DataSource = cabFurStorage.cellsBs;
        }

        private void dgvDocumentsSetting(ref PercentageDataGrid grid)
        {
            //grid.Columns.Add(AssignmentsManager.CreationUserColumn);
            grid.Columns.Add(assignmentsManager.PrintUserColumn);
            grid.Columns.Add(assignmentsManager.DocTypesColumn);
            grid.AutoGenerateColumns = false;

            foreach (DataGridViewColumn Column in grid.Columns)
            {
                Column.ReadOnly = true;
                Column.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }

            if (grid.Columns.Contains("FileSize"))
                grid.Columns["FileSize"].Visible = false;
            if (grid.Columns.Contains("CabFurDocTypeID"))
                grid.Columns["CabFurDocTypeID"].Visible = false;
            if (grid.Columns.Contains("CabFurAssignmentID"))
                grid.Columns["CabFurAssignmentID"].Visible = false;
            if (grid.Columns.Contains("CreationUserID"))
                grid.Columns["CreationUserID"].Visible = false;
            if (grid.Columns.Contains("PrintUserID"))
                grid.Columns["PrintUserID"].Visible = false;
            if (grid.Columns.Contains("CabFurDocumentID"))
                grid.Columns["CabFurDocumentID"].Visible = false;
            if (grid.Columns.Contains("SectorID"))
                grid.Columns["SectorID"].Visible = false;
            if (grid.Columns.Contains("CreationDateTime"))
                grid.Columns["CreationDateTime"].Visible = false;

            grid.Columns["CreationDateTime"].DefaultCellStyle.Format = "dd.MM.yyyy HH:mm";
            grid.Columns["PrintDateTime"].DefaultCellStyle.Format = "dd.MM.yyyy HH:mm";

            grid.Columns["CreationDateTime"].HeaderText = "Дата создания";
            grid.Columns["PrintDateTime"].HeaderText = "Дата печати";
            grid.Columns["FileName"].HeaderText = "Имя";

            int DisplayIndex = 0;
            grid.Columns["FileName"].DisplayIndex = DisplayIndex++;
            grid.Columns["DocTypesColumn"].DisplayIndex = DisplayIndex++;
            grid.Columns["PrintUserColumn"].DisplayIndex = DisplayIndex++;
            grid.Columns["PrintDateTime"].DisplayIndex = DisplayIndex++;
            grid.Columns["PrintDocColumn"].DisplayIndex = DisplayIndex++;

        }

        private void dgvDetailsSetting(bool bClient, ref PercentageDataGrid grid)
        {
            if (!grid.Columns.Contains("TechStoreNameColumn"))
                grid.Columns.Add(assignmentsManager.TechStoreNameColumn);
            if (!grid.Columns.Contains("MeasuresColumn"))
                grid.Columns.Add(assignmentsManager.MeasuresColumn);
            if (!grid.Columns.Contains("ColorColumn"))
                grid.Columns.Add(assignmentsManager.ColorColumn);
            if (!grid.Columns.Contains("PatinaColumn"))
                grid.Columns.Add(assignmentsManager.PatinaColumn);
            grid.AutoGenerateColumns = false;

            foreach (DataGridViewColumn Column in grid.Columns)
            {
                Column.ReadOnly = true;
                Column.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }

            if (grid.Columns.Contains("CabFurAssignmentDetailID"))
                grid.Columns["CabFurAssignmentDetailID"].Visible = false;
            if (grid.Columns.Contains("TechStoreID"))
                grid.Columns["TechStoreID"].Visible = false;
            if (grid.Columns.Contains("MeasureID"))
                grid.Columns["MeasureID"].Visible = false;
            if (grid.Columns.Contains("ColorID"))
                grid.Columns["ColorID"].Visible = false;
            if (grid.Columns.Contains("PatinaID"))
                grid.Columns["PatinaID"].Visible = false;

            grid.Columns["Count"].HeaderText = "Кол-во";

            int DisplayIndex = 0;
            if (bClient && grid.Columns.Contains("ClientName"))
            {
                grid.Columns["ClientName"].DisplayIndex = DisplayIndex++;
                grid.Columns["ClientName"].HeaderText = "Клиент";
            }
            grid.Columns["TechStoreNameColumn"].DisplayIndex = DisplayIndex++;
            grid.Columns["ColorColumn"].DisplayIndex = DisplayIndex++;
            grid.Columns["PatinaColumn"].DisplayIndex = DisplayIndex++;
            grid.Columns["Count"].DisplayIndex = DisplayIndex++;
            grid.Columns["MeasuresColumn"].DisplayIndex = DisplayIndex++;

        }

        private void dgvAllAssignmentsSetting(ref PercentageDataGrid grid)
        {
            grid.Columns.Add(assignmentsManager.CreationUserColumn);
            grid.Columns.Add(assignmentsManager.AgreementUserColumn);
            grid.AutoGenerateColumns = false;

            if (grid.Columns.Contains("CreationUserID"))
                grid.Columns["CreationUserID"].Visible = false;
            if (grid.Columns.Contains("AgreementUserID"))
                grid.Columns["AgreementUserID"].Visible = false;
            if (grid.Columns.Contains("PrintStatus"))
                grid.Columns["PrintStatus"].Visible = false;
            if (grid.Columns.Contains("DocsCount"))
                grid.Columns["DocsCount"].Visible = false;
            if (grid.Columns.Contains("DocsPrintedCount"))
                grid.Columns["DocsPrintedCount"].Visible = false;
            if (grid.Columns.Contains("ProductionUserID"))
                grid.Columns["ProductionUserID"].Visible = false;
            if (grid.Columns.Contains("OutProductionUserID"))
                grid.Columns["OutProductionUserID"].Visible = false;
            if (grid.Columns.Contains("PlanDispatchUserID"))
                grid.Columns["PlanDispatchUserID"].Visible = false;

            grid.Columns["CreationDateTime"].DefaultCellStyle.Format = "dd.MM.yyyy";
            grid.Columns["PlanDispatchDateTime"].DefaultCellStyle.Format = "dd.MM.yyyy";
            grid.Columns["AgreementDateTime"].DefaultCellStyle.Format = "dd.MM.yyyy";
            grid.Columns["ProductionDateTime"].DefaultCellStyle.Format = "dd.MM.yyyy";
            grid.Columns["OutProductionDateTime"].DefaultCellStyle.Format = "dd.MM.yyyy";

            grid.Columns["CreationDateTime"].HeaderText = "Дата\r\nсоздания";
            grid.Columns["PlanDispatchDateTime"].HeaderText = "Дата\r\nотгрузки";
            grid.Columns["AgreementDateTime"].HeaderText = "Дата\r\nсогласования";
            grid.Columns["ProductionDateTime"].HeaderText = "Дата\r\nзапуска";
            grid.Columns["OutProductionDateTime"].HeaderText = "Дата\r\nвыхода";
            grid.Columns["CabFurAssignmentID"].HeaderText = "№ задания";
            grid.Columns["DocsCount"].HeaderText = "Всего заданий";
            grid.Columns["ClientNotes"].HeaderText = "Клиент и № заказа";
            grid.Columns["Notes"].HeaderText = "Примечание";
            grid.Columns["PackagesCount"].HeaderText = "Упаковка, шт.";
            grid.Columns["DocsPrintedCount"].HeaderText = "Заданий распечатано";

            foreach (DataGridViewColumn Column in grid.Columns)
            {
                Column.ReadOnly = true;
                Column.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }

            grid.Columns["ClientNotes"].ReadOnly = false;
            grid.Columns["Notes"].ReadOnly = false;

            int DisplayIndex = 0;
            grid.Columns["CabFurAssignmentID"].DisplayIndex = DisplayIndex++;
            grid.Columns["CreationUserColumn"].DisplayIndex = DisplayIndex++;
            grid.Columns["CreationDateTime"].DisplayIndex = DisplayIndex++;
            grid.Columns["PlanDispatchDateTime"].DisplayIndex = DisplayIndex++;
            grid.Columns["AgreementUserColumn"].DisplayIndex = DisplayIndex++;
            grid.Columns["AgreementDateTime"].DisplayIndex = DisplayIndex++;
            grid.Columns["ProductionDateTime"].DisplayIndex = DisplayIndex++;
            grid.Columns["OutProductionDateTime"].DisplayIndex = DisplayIndex++;
            grid.Columns["PackagesCount"].DisplayIndex = DisplayIndex++;
            grid.Columns["ClientNotes"].DisplayIndex = DisplayIndex++;
            grid.Columns["Notes"].DisplayIndex = DisplayIndex++;
            grid.Columns["InProdColumn"].DisplayIndex = DisplayIndex++;
            grid.Columns["OutProdColumn"].DisplayIndex = DisplayIndex++;
        }

        private void dgvNewAssignmentSetting(ref PercentageDataGrid grid)
        {
            grid.Columns.Add(assignmentsManager.CreationUserColumn);
            grid.Columns.Add(assignmentsManager.MeasuresColumn);
            grid.Columns.Add(assignmentsManager.TechStoreNameColumn);
            grid.Columns.Add(assignmentsManager.ColorColumn);
            grid.Columns.Add(assignmentsManager.PatinaColumn);
            grid.Columns.Add(assignmentsManager.CoverColumn);
            grid.Columns.Add(assignmentsManager.InsetTypeColumn);
            grid.Columns.Add(assignmentsManager.InsetColorColumn);
            grid.AutoGenerateColumns = false;

            if (grid.Columns.Contains("CabFurAssignmentDetailID"))
                grid.Columns["CabFurAssignmentDetailID"].Visible = false;
            if (grid.Columns.Contains("CabFurAssignmentID"))
                grid.Columns["CabFurAssignmentID"].Visible = false;
            if (grid.Columns.Contains("CreationUserID"))
                grid.Columns["CreationUserID"].Visible = false;
            if (grid.Columns.Contains("TechStoreID"))
                grid.Columns["TechStoreID"].Visible = false;
            if (grid.Columns.Contains("MeasureID"))
                grid.Columns["MeasureID"].Visible = false;
            if (grid.Columns.Contains("InsetColorID"))
                grid.Columns["InsetColorID"].Visible = false;
            if (grid.Columns.Contains("InsetTypeID"))
                grid.Columns["InsetTypeID"].Visible = false;
            if (grid.Columns.Contains("PatinaID"))
                grid.Columns["PatinaID"].Visible = false;
            if (grid.Columns.Contains("CoverID"))
                grid.Columns["CoverID"].Visible = false;
            if (grid.Columns.Contains("ColorID"))
                grid.Columns["ColorID"].Visible = false;

            grid.Columns["CreationDateTime"].HeaderText = "Дата создания";
            grid.Columns["Length"].HeaderText = "Длина, мм";
            grid.Columns["Width"].HeaderText = "Ширина, мм";
            grid.Columns["Height"].HeaderText = "Высота, мм";
            grid.Columns["Thickness"].HeaderText = "Толщина, мм";
            grid.Columns["Diameter"].HeaderText = "Диаметр, мм";
            grid.Columns["Admission"].HeaderText = "Допуск, мм";
            grid.Columns["Weight"].HeaderText = "Вес, кг";
            grid.Columns["Capacity"].HeaderText = "Емкость, л";
            grid.Columns["InsetHeightAdmission"].HeaderText = "Допуск на вставку\r\n   по высоте, мм";
            grid.Columns["InsetWidthAdmission"].HeaderText = "Допуск на вставку\r\n   по ширине, мм";
            grid.Columns["WidthMin"].HeaderText = "Ширина min, мм";
            grid.Columns["WidthMax"].HeaderText = "Ширина max, мм";
            grid.Columns["HeightMin"].HeaderText = "Высота min, мм";
            grid.Columns["HeightMax"].HeaderText = "Высота max, мм";
            grid.Columns["Count"].HeaderText = "Кол-во";

            foreach (DataGridViewColumn Column in grid.Columns)
            {
                Column.ReadOnly = false;
                Column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }

            grid.Columns["CreationDateTime"].ReadOnly = true;
            grid.Columns["MeasuresColumn"].ReadOnly = true;
            grid.Columns["TechStoreNameColumn"].ReadOnly = true;
            grid.Columns["CreationUserColumn"].ReadOnly = true;

            int DisplayIndex = 0;
            grid.Columns["CabFurAssignmentDetailID"].DisplayIndex = DisplayIndex++;
            grid.Columns["TechStoreNameColumn"].DisplayIndex = DisplayIndex++;
            grid.Columns["MeasuresColumn"].DisplayIndex = DisplayIndex++;
            grid.Columns["CoverColumn"].DisplayIndex = DisplayIndex++;
            grid.Columns["ColorColumn"].DisplayIndex = DisplayIndex++;
            grid.Columns["PatinaColumn"].DisplayIndex = DisplayIndex++;
            grid.Columns["InsetTypeColumn"].DisplayIndex = DisplayIndex++;
            grid.Columns["InsetColorColumn"].DisplayIndex = DisplayIndex++;
        }

        private void dgvComplementsSetting(ref PercentageDataGrid grid)
        {
            grid.Columns.Add(assignmentsManager.ClientColumn);
            grid.AutoGenerateColumns = false;

            if (grid.Columns.Contains("ClientID"))
                grid.Columns["ClientID"].Visible = false;

            grid.Columns["OrderNumber"].HeaderText = "№ заказа";
            grid.Columns["ComplementsCount"].HeaderText = "Кол-во этикеток";
            grid.Columns["PrintedPercentage"].HeaderText = "Распечатано, %";
            grid.Columns["PackedCount"].HeaderText = "Скомплектовано, кол-во";
            grid.Columns["PackedPercentage"].HeaderText = "Скомплектовано, %";

            foreach (DataGridViewColumn Column in grid.Columns)
            {
                Column.ReadOnly = true;
                //Column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                //Column.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }

            int DisplayIndex = 0;
            grid.Columns["ClientColumn"].DisplayIndex = DisplayIndex++;
            grid.Columns["OrderNumber"].DisplayIndex = DisplayIndex++;
            grid.Columns["ComplementsCount"].DisplayIndex = DisplayIndex++;
            grid.Columns["PrintedPercentage"].DisplayIndex = DisplayIndex++;
            grid.Columns["PackedCount"].DisplayIndex = DisplayIndex++;
            grid.Columns["PackedPercentage"].DisplayIndex = DisplayIndex++;

            grid.Columns["ComplementsCount"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            grid.Columns["PackedCount"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            grid.Columns["PrintedPercentage"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            grid.AddPercentageColumn("PrintedPercentage");
            grid.Columns["PackedPercentage"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            grid.AddPercentageColumn("PackedPercentage");
        }

        private void dgvMainOrdersSetting(ref PercentageDataGrid grid)
        {
            grid.Columns.Add(complementsManager.ProductionStatusColumn);
            grid.Columns.Add(complementsManager.StorageStatusColumn);
            grid.Columns.Add(complementsManager.ExpeditionStatusColumn);
            grid.Columns.Add(complementsManager.DispatchStatusColumn);
            grid.AutoGenerateColumns = false;

            if (grid.Columns.Contains("TPSProductionStatusID"))
                grid.Columns["TPSProductionStatusID"].Visible = false;
            if (grid.Columns.Contains("TPSStorageStatusID"))
                grid.Columns["TPSStorageStatusID"].Visible = false;
            if (grid.Columns.Contains("TPSExpeditionStatusID"))
                grid.Columns["TPSExpeditionStatusID"].Visible = false;
            if (grid.Columns.Contains("TPSDispatchStatusID"))
                grid.Columns["TPSDispatchStatusID"].Visible = false;

            grid.Columns["MainOrderID"].HeaderText = "№ п\\п";
            grid.Columns["DocDateTime"].HeaderText = "Дата\r\nсоздания";
            grid.Columns["Notes"].HeaderText = "Примечание";
            grid.Columns["Weight"].HeaderText = "Вес, кг.";
            grid.Columns["PackedCount"].HeaderText = "Упаковано, кол-во";
            grid.Columns["PackedPercentage"].HeaderText = "Упаковано, %";
            grid.Columns["AllocPackDateTime"].HeaderText = "Дата\r\nраспределения";

            foreach (DataGridViewColumn Column in grid.Columns)
            {
                Column.ReadOnly = true;
                //Column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                //Column.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }

            int DisplayIndex = 0;
            grid.Columns["MainOrderID"].DisplayIndex = DisplayIndex++;
            grid.Columns["Notes"].DisplayIndex = DisplayIndex++;
            grid.Columns["PackedCount"].DisplayIndex = DisplayIndex++;
            grid.Columns["PackedPercentage"].DisplayIndex = DisplayIndex++;
            grid.Columns["AllocPackDateTime"].DisplayIndex = DisplayIndex++;
            grid.Columns["Weight"].DisplayIndex = DisplayIndex++;
            grid.Columns["DocDateTime"].DisplayIndex = DisplayIndex++;

            grid.Columns["PackedCount"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            grid.Columns["PackedPercentage"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            grid.AddPercentageColumn("PackedPercentage");
        }

        private void dgvComplementLabelsSetting(ref PercentageDataGrid grid)
        {
            grid.AutoGenerateColumns = false;

            if (grid.Columns.Contains("PackNumber"))
                grid.Columns["PackNumber"].Visible = false;
            if (grid.Columns.Contains("TechStoreSubGroupID"))
                grid.Columns["TechStoreSubGroupID"].Visible = false;
            if (grid.Columns.Contains("CabFurnitureComplementID"))
                grid.Columns["CabFurnitureComplementID"].Visible = false;
            if (grid.Columns.Contains("MainOrderID"))
                grid.Columns["MainOrderID"].Visible = false;

            grid.Columns["Index"].HeaderText = "№ этикетки";
            grid.Columns["PrintDateTime"].HeaderText = "Дата печати";
            grid.Columns["PackingDateTime"].HeaderText = "Дата упаковки";

            foreach (DataGridViewColumn Column in grid.Columns)
            {
                Column.ReadOnly = true;
                //Column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                //Column.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }

            int DisplayIndex = 0;
            grid.Columns["Index"].DisplayIndex = DisplayIndex++;
            grid.Columns["PrintDateTime"].DisplayIndex = DisplayIndex++;
        }

        private void dgvComplementDetailsSetting(ref PercentageDataGrid grid)
        {
            grid.Columns.Add(assignmentsManager.CTechStoreNameColumn);
            grid.Columns.Add(assignmentsManager.TechStoreNameColumn);
            grid.Columns.Add(assignmentsManager.CoverColumn);
            grid.Columns.Add(assignmentsManager.PatinaColumn);
            grid.Columns.Add(assignmentsManager.InsetColorColumn);
            grid.AutoGenerateColumns = false;

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
            if (grid.Columns.Contains("CabFurnitureComplenentDetailID"))
                grid.Columns["CabFurnitureComplenentDetailID"].Visible = false;
            if (grid.Columns.Contains("CabFurnitureComplementID"))
                grid.Columns["CabFurnitureComplementID"].Visible = false;
            if (grid.Columns.Contains("CreateUserID"))
                grid.Columns["CreateUserID"].Visible = false;
            if (grid.Columns.Contains("MainOrderID"))
                grid.Columns["MainOrderID"].Visible = false;
            if (grid.Columns.Contains("PackNumber"))
                grid.Columns["PackNumber"].Visible = false;
            if (grid.Columns.Contains("CNotes"))
                grid.Columns["CNotes"].Visible = false;

            grid.Columns["CreateDateTime"].HeaderText = "Дата создания";
            grid.Columns["MainOrderID"].HeaderText = "№ подзаказа";
            grid.Columns["CNotes"].HeaderText = "Прим. к подзаказу";
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
            grid.Columns["MainOrderID"].DisplayIndex = DisplayIndex++;
            grid.Columns["CNotes"].DisplayIndex = DisplayIndex++;
            grid.Columns["PackNumber"].DisplayIndex = DisplayIndex++;
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

        private void dgvPackagesSetting(ref PercentageDataGrid grid)
        {
            grid.AutoGenerateColumns = false;

            grid.Columns["ClientName"].HeaderText = "Клиент";
            grid.Columns["CabFurAssignmentID"].HeaderText = "№ задания";
            grid.Columns["PackagesCount"].HeaderText = "Кол-во этикеток";
            grid.Columns["PrintedPercentage"].HeaderText = "Распечатано, %";
            grid.Columns["PackedCount"].HeaderText = "Упаковано, кол-во";
            grid.Columns["PackedPercentage"].HeaderText = "Упаковано, %";

            foreach (DataGridViewColumn Column in grid.Columns)
            {
                Column.ReadOnly = true;
                //Column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                //Column.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }

            int DisplayIndex = 0;
            grid.Columns["ClientName"].DisplayIndex = DisplayIndex++;
            grid.Columns["CabFurAssignmentID"].DisplayIndex = DisplayIndex++;
            grid.Columns["PackagesCount"].DisplayIndex = DisplayIndex++;
            grid.Columns["PrintedPercentage"].DisplayIndex = DisplayIndex++;
            grid.Columns["PackedCount"].DisplayIndex = DisplayIndex++;
            grid.Columns["PackedPercentage"].DisplayIndex = DisplayIndex++;

            grid.Columns["PackagesCount"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            grid.Columns["PackedCount"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            grid.Columns["PrintedPercentage"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            grid.AddPercentageColumn("PrintedPercentage");
            grid.Columns["PackedPercentage"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            grid.AddPercentageColumn("PackedPercentage");
        }

        private void dgvPackagesLabelsSetting(ref PercentageDataGrid grid)
        {
            grid.AutoGenerateColumns = false;

            if (grid.Columns.Contains("PackagesCount"))
                grid.Columns["PackagesCount"].Visible = false;
            if (grid.Columns.Contains("TechStoreSubGroupID"))
                grid.Columns["TechStoreSubGroupID"].Visible = false;
            if (grid.Columns.Contains("CabFurAssignmentDetailID"))
                grid.Columns["CabFurAssignmentDetailID"].Visible = false;
            if (grid.Columns.Contains("MainOrderID"))
                grid.Columns["MainOrderID"].Visible = false;
            if (grid.Columns.Contains("QualityControlInUserID"))
                grid.Columns["QualityControlInUserID"].Visible = false;
            if (grid.Columns.Contains("QualityControlOutUserID"))
                grid.Columns["QualityControlOutUserID"].Visible = false;

            grid.Columns["CabFurniturePackageID"].HeaderText = "ID";
            grid.Columns["PackNumber"].HeaderText = "№ упаковки";
            if (grid.Columns.Contains("Index"))
                grid.Columns["Index"].HeaderText = "№ этикетки";
            grid.Columns["PrintDateTime"].HeaderText = "Дата печати";
            grid.Columns["AddToStorageDateTime"].HeaderText = "Принято на склад";
            grid.Columns["RemoveFromStorageDateTime"].HeaderText = "Списано со склада";
            grid.Columns["QualityControlInDateTime"].HeaderText = "Отправлено на ОТК";
            grid.Columns["QualityControlOutDateTime"].HeaderText = "Принято с ОТК";
            grid.Columns["QualityControl"].HeaderText = "ОТК";
            grid.Columns["Name"].HeaderText = "Ячейка";

            foreach (DataGridViewColumn Column in grid.Columns)
            {
                Column.ReadOnly = true;
                //Column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }

            int DisplayIndex = 0;
            if (grid.Columns.Contains("Index"))
                grid.Columns["Index"].DisplayIndex = DisplayIndex++;
            grid.Columns["PackNumber"].DisplayIndex = DisplayIndex++;
            grid.Columns["CabFurniturePackageID"].DisplayIndex = DisplayIndex++;
            grid.Columns["PrintDateTime"].DisplayIndex = DisplayIndex++;
            grid.Columns["AddToStorageDateTime"].DisplayIndex = DisplayIndex++;
            grid.Columns["RemoveFromStorageDateTime"].DisplayIndex = DisplayIndex++;
            grid.Columns["Name"].DisplayIndex = DisplayIndex++;
            grid.Columns["QualityControl"].DisplayIndex = DisplayIndex++;
            grid.Columns["QualityControlInDateTime"].DisplayIndex = DisplayIndex++;
            grid.Columns["QualityControlOutDateTime"].DisplayIndex = DisplayIndex++;

            grid.Columns["CabFurniturePackageID"].Width = 50;
        }

        private void dgvPackagesDetailsSetting(ref PercentageDataGrid grid)
        {
            grid.Columns.Add(assignmentsManager.CTechStoreNameColumn);
            grid.Columns.Add(assignmentsManager.TechStoreNameColumn);
            grid.Columns.Add(assignmentsManager.CoverColumn);
            grid.Columns.Add(assignmentsManager.PatinaColumn);
            grid.Columns.Add(assignmentsManager.InsetColorColumn);
            grid.AutoGenerateColumns = false;

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

        private void dgvCellsSetting(ref PercentageDataGrid grid)
        {
            grid.AutoGenerateColumns = false;

            if (grid.Columns.Contains("RackID"))
                grid.Columns["RackID"].Visible = false;

            grid.Columns["CellID"].HeaderText = "ID";
            grid.Columns["Name"].HeaderText = "Имя ячейки";
            grid.MultiSelect = true;
            grid.Columns["Name"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            grid.Columns["CellID"].Width = 50;

            int DisplayIndex = 0;
            grid.Columns["CellID"].DisplayIndex = DisplayIndex++;
            grid.Columns["Name"].DisplayIndex = DisplayIndex++;

        }

        private void dgvStoragePackagesLabelsSetting(ref PercentageDataGrid grid)
        {
            grid.AutoGenerateColumns = false;

            if (grid.Columns.Contains("PackagesCount"))
                grid.Columns["PackagesCount"].Visible = false;
            if (grid.Columns.Contains("TechStoreSubGroupID"))
                grid.Columns["TechStoreSubGroupID"].Visible = false;
            if (grid.Columns.Contains("CabFurAssignmentDetailID"))
                grid.Columns["CabFurAssignmentDetailID"].Visible = false;
            if (grid.Columns.Contains("MainOrderID"))
                grid.Columns["MainOrderID"].Visible = false;
            if (grid.Columns.Contains("QualityControlInUserID"))
                grid.Columns["QualityControlInUserID"].Visible = false;
            if (grid.Columns.Contains("QualityControlOutUserID"))
                grid.Columns["QualityControlOutUserID"].Visible = false;
            if (grid.Columns.Contains("CellID"))
                grid.Columns["CellID"].Visible = false;
            if (grid.Columns.Contains("QualityControl"))
                grid.Columns["QualityControl"].Visible = false;

            grid.Columns["CabFurniturePackageID"].HeaderText = "ID";
            grid.Columns["PackNumber"].HeaderText = "№ упаковки";
            grid.Columns["Index"].HeaderText = "№ этикетки";
            grid.Columns["AddToStorageDateTime"].HeaderText = "Принято на склад";
            grid.Columns["RemoveFromStorageDateTime"].HeaderText = "Списано со склада";
            grid.Columns["QualityControlInDateTime"].HeaderText = "Отправлено на ОТК";
            grid.Columns["QualityControlOutDateTime"].HeaderText = "Принято с ОТК";
            grid.Columns["QualityControl"].HeaderText = "ОТК";
            grid.Columns["Name"].HeaderText = "Ячейка";

            foreach (DataGridViewColumn Column in grid.Columns)
            {
                Column.ReadOnly = true;
                Column.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }

            int DisplayIndex = 0;
            grid.Columns["Index"].DisplayIndex = DisplayIndex++;
            grid.Columns["PackNumber"].DisplayIndex = DisplayIndex++;
            grid.Columns["CabFurniturePackageID"].DisplayIndex = DisplayIndex++;
            grid.Columns["AddToStorageDateTime"].DisplayIndex = DisplayIndex++;
            grid.Columns["RemoveFromStorageDateTime"].DisplayIndex = DisplayIndex++;
            grid.Columns["Name"].DisplayIndex = DisplayIndex++;
            grid.Columns["QualityControl"].DisplayIndex = DisplayIndex++;
            grid.Columns["QualityControlInDateTime"].DisplayIndex = DisplayIndex++;
            grid.Columns["QualityControlOutDateTime"].DisplayIndex = DisplayIndex++;

            grid.Columns["CabFurniturePackageID"].Width = 50;
        }

        private void dgvScanDetailsSetting(ref PercentageDataGrid grid)
        {
            grid.Columns.Add(assignmentsManager.CTechStoreNameColumn);
            grid.Columns.Add(assignmentsManager.TechStoreNameColumn);
            grid.Columns.Add(assignmentsManager.CoverColumn);
            grid.Columns.Add(assignmentsManager.PatinaColumn);
            grid.AutoGenerateColumns = false;

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
            grid.Columns["Notes"].DisplayIndex = DisplayIndex++;
            grid.Columns["TechStoreNameColumn"].DisplayIndex = DisplayIndex++;
            grid.Columns["Length"].DisplayIndex = DisplayIndex++;
            grid.Columns["Height"].DisplayIndex = DisplayIndex++;
            grid.Columns["Width"].DisplayIndex = DisplayIndex++;
            grid.Columns["Count"].DisplayIndex = DisplayIndex++;
            grid.Columns["CreateDateTime"].DisplayIndex = DisplayIndex++;
        }

        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);

            if (m.Msg == 6 && m.WParam.ToInt32() == 1)
            {
                if (TopForm != null)
                    TopForm.Activate();
            }
        }

        private void kryptonCheckSet3_CheckedButtonChanged(object sender, EventArgs e)
        {
            if (assignmentsManager == null)
                return;

            Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных.\r\nПодождите..."); });
            T.Start();
            while (!SplashWindow.bSmallCreated) ;
            NeedSplash = false;

            if (kryptonCheckSet3.CheckedButton == cbtnNewAssignment)
            {
                assignmentsManager.NewAssignment = true;
                //pnlNewAssignment.BringToFront();
                tabControl1.SelectedIndex = 1;
            }
            if (kryptonCheckSet3.CheckedButton == cbtnAllAssignments)
            {
                assignmentsManager.NewAssignment = false;
                //pnlAllAssignments.BringToFront();
                tabControl1.SelectedIndex = 0;
            }
            if (kryptonCheckSet3.CheckedButton == cbtnOrdersCondition)
            {
                //pnlOrdersCondition.BringToFront();
                tabControl1.SelectedIndex = 4;
            }
            DateTime date1 = DateTimePicker1.Value.Date;
            DateTime date2 = DateTimePicker2.Value.Date;
            if (kryptonCheckSet3.CheckedButton == cbtnComplements)
            {
                //pnlOrdersCondition.BringToFront();
                tabControl1.SelectedIndex = 2;


                complementsManager.UpdateComplements(date1, date2);
            }
            if (kryptonCheckSet3.CheckedButton == cbtnPackages)
            {
                //pnlOrdersCondition.BringToFront();
                tabControl1.SelectedIndex = 3;
                packagesManager.UpdatePackages(date1, date2);
                //PackagesManager.ClearPackges();
            }
            if (kryptonCheckSet3.CheckedButton == cbtnPackagesScan)
            {
                //pnlOrdersCondition.BringToFront();
                tabControl1.SelectedIndex = 5;
            }
            if (kryptonCheckSet3.CheckedButton == cbtnStorage)
            {
                //pnlOrdersCondition.BringToFront();
                tabControl1.SelectedIndex = 6;
            }
            if (kryptonCheckSet3.CheckedButton == cbtnStatistics)
            {
                //pnlOrdersCondition.BringToFront();
                tabControl1.SelectedIndex = 7;

                FilterStatisticsData();
            }
            if (MenuPanel.Visible == true && kryptonCheckSet3.CheckedButton == cbtnStatistics)
            {
                MenuPanel.Visible = false;
                MenuPanel2.Visible = true;
            }
            else if (MenuPanel2.Visible == true && kryptonCheckSet3.CheckedButton != cbtnStatistics)
            {
                MenuPanel.Visible = true;
                MenuPanel2.Visible = false;
            }
            NeedSplash = true;
            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
            MenuPanel.BringToFront();
        }

        private void kryptonContextMenuItem7_Click(object sender, EventArgs e)
        {
            bool OKCancel = Infinium.LightMessageBox.Show(ref TopForm, true,
                    "Вы собираетесь создать новое задание. Продолжить?",
                    "Новое задание");
            if (!OKCancel)
                return;

            Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Сохранение.\r\nПодождите..."); });
            T.Start();
            while (!SplashWindow.bSmallCreated) ;
            NeedSplash = false;

            kryptonCheckSet3.CheckedButton = cbtnNewAssignment;
            assignmentsManager.NewAssignment = true;
            assignmentsManager.CabFurAssignmentID = 0;
            assignmentsManager.UpdateNewAssignment(0);

            NeedSplash = true;
            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
        }

        private void dgvAllAssignments_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            return;

            //if (assignmentsManager == null)
            //    return;
            //int CabFurAssignmentID = 0;
            //if (dgvAllAssignments.SelectedRows.Count != 0 && dgvAllAssignments.SelectedRows[0].Cells["CabFurAssignmentID"].Value != DBNull.Value)
            //    CabFurAssignmentID = Convert.ToInt32(dgvAllAssignments.SelectedRows[0].Cells["CabFurAssignmentID"].Value);
            //assignmentsManager.CabFurAssignmentID = CabFurAssignmentID;

            //Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных.\r\nПодождите..."); });
            //T.Start();
            //while (!SplashForm.bCreated) ;
            //NeedSplash = false;

            //kryptonCheckSet3.CheckedButton = cbtnNewAssignment;
            //assignmentsManager.NewAssignment = false;
            //assignmentsManager.UpdateNewAssignment(CabFurAssignmentID);
            //tabControl1.SelectedIndex = 1;
            //CheckColumns(ref dgvNewAssignment);

            //NeedSplash = true;
            //while (SplashWindow.bSmallCreated)
            //    SmallWaitForm.CloseS = true;
        }

        private void dgvAllAssignments_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Right && e.RowIndex != -1)
            {
                kryptonContextMenu2.Show(new Point(Cursor.Position.X - 212, Cursor.Position.Y - 10));
            }
        }

        private void kryptonContextMenuItem3_Click(object sender, EventArgs e)
        {
            if (assignmentsManager == null)
                return;
            int CabFurAssignmentID = 0;
            if (dgvAllAssignments.SelectedRows.Count != 0 && dgvAllAssignments.SelectedRows[0].Cells["CabFurAssignmentID"].Value != DBNull.Value)
                CabFurAssignmentID = Convert.ToInt32(dgvAllAssignments.SelectedRows[0].Cells["CabFurAssignmentID"].Value);
            assignmentsManager.CabFurAssignmentID = CabFurAssignmentID;

            Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных.\r\nПодождите..."); });
            T.Start();
            while (!SplashForm.bCreated) ;
            NeedSplash = false;

            kryptonCheckSet3.CheckedButton = cbtnNewAssignment;
            assignmentsManager.NewAssignment = false;
            assignmentsManager.UpdateNewAssignment(CabFurAssignmentID);
            tabControl1.SelectedIndex = 1;
            CheckColumns(ref dgvNewAssignment);

            NeedSplash = true;
            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
        }

        private void SaveAllAssignments()
        {
            int CabFurAssignmentID = 0;
            if (dgvAllAssignments.SelectedRows.Count != 0 && dgvAllAssignments.SelectedRows[0].Cells["CabFurAssignmentID"].Value != DBNull.Value)
                CabFurAssignmentID = Convert.ToInt32(dgvAllAssignments.SelectedRows[0].Cells["CabFurAssignmentID"].Value);

            Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Сохранение.\r\nПодождите..."); });
            T.Start();
            while (!SplashWindow.bSmallCreated) ;
            NeedSplash = false;

            DateTime date1 = CreateDateTimePicker1.Value.Date;
            DateTime date2 = CreateDateTimePicker2.Value.Date;

            assignmentsManager.SaveAllAssignments();
            assignmentsManager.UpdateAllAssignments(date1, date2);
            assignmentsManager.MoveToAssignmentID(CabFurAssignmentID);

            NeedSplash = true;
            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
        }

        private void dgvAllAssignments_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.S)
                SaveAllAssignments();
        }

        private void cmiSaveAssignments_Click(object sender, EventArgs e)
        {
            if (userRole == RoleTypes.AdminRole)
            {
                int CabFurAssignmentID = 0;
                if (dgvAllAssignments.SelectedRows.Count != 0 &&
                    dgvAllAssignments.SelectedRows[0].Cells["CabFurAssignmentID"].Value != DBNull.Value)
                    CabFurAssignmentID =
                        Convert.ToInt32(dgvAllAssignments.SelectedRows[0].Cells["CabFurAssignmentID"].Value);

                int AllPackagesCount = 0;
                if (percentageDataGrid1.SelectedRows.Count != 0 &&
                    percentageDataGrid1.SelectedRows[0].Cells["Count"].Value != DBNull.Value)
                    AllPackagesCount = Convert.ToInt32(percentageDataGrid1.SelectedRows[0].Cells["Count"].Value);
                assignmentsManager.EditPackCount(CabFurAssignmentID, AllPackagesCount);
                assignmentsManager.SaveNewAssignment();
            }

            SaveAllAssignments();
        }

        private void MinimizeButton_Click(object sender, EventArgs e)
        {
            NeedSplash = false;
            FormEvent = eHide;
            AnimateTimer.Enabled = true;
        }

        private void kryptonContextMenuItem13_Click(object sender, EventArgs e)
        {
            int CabFurAssignmentID = 0;
            if (dgvAllAssignments.SelectedRows.Count != 0 && dgvAllAssignments.SelectedRows[0].Cells["CabFurAssignmentID"].Value != DBNull.Value)
                CabFurAssignmentID = Convert.ToInt32(dgvAllAssignments.SelectedRows[0].Cells["CabFurAssignmentID"].Value);

            Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Сохранение данных.\r\nПодождите..."); });
            T.Start();
            while (!SplashWindow.bSmallCreated) ;
            NeedSplash = false;

            assignmentsManager.AgreeAssignment(CabFurAssignmentID);
            assignmentsManager.SaveAllAssignments();
            FilterAssignments();
            assignmentsManager.MoveToAssignmentID(CabFurAssignmentID);

            NeedSplash = true;
            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
        }

        private void cmbxTechStoreGroups_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (assignmentsManager == null)
                return;
            int TechStoreGroupID = 0;
            if (cmbxTechStoreGroups.SelectedItem != null && ((DataRowView)cmbxTechStoreGroups.SelectedItem).Row["TechStoreGroupID"] != DBNull.Value)
                TechStoreGroupID = Convert.ToInt32(((DataRowView)cmbxTechStoreGroups.SelectedItem).Row["TechStoreGroupID"]);
            assignmentsManager.FilterTechStoreSubGroups(TechStoreGroupID);
            cmbxTechStoreSubGroups_SelectedIndexChanged(null, null);
        }

        private void cmbxTechStoreSubGroups_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (assignmentsManager == null)
                return;
            int TechStoreSubGroupID = 0;
            if (cmbxTechStoreSubGroups.SelectedItem != null && ((DataRowView)cmbxTechStoreSubGroups.SelectedItem).Row["TechStoreSubGroupID"] != DBNull.Value)
                TechStoreSubGroupID = Convert.ToInt32(((DataRowView)cmbxTechStoreSubGroups.SelectedItem).Row["TechStoreSubGroupID"]);
            assignmentsManager.FilterTechStore(TechStoreSubGroupID);
            cmbxTechStore_SelectedIndexChanged(null, null);
        }

        private void CheckColumns(ref PercentageDataGrid grid)
        {
            foreach (DataGridViewColumn Column in grid.Columns)
            {
                foreach (DataGridViewRow Row in grid.Rows)
                {
                    if (Row.Cells[Column.Index].FormattedValue.ToString().Length > 0)
                    {
                        grid.Columns[Column.Index].Visible = true;
                        break;
                    }
                    else
                        grid.Columns[Column.Index].Visible = false;
                }
            }

            if (grid.Columns.Contains("CabFurAssignmentDetailID"))
                grid.Columns["CabFurAssignmentDetailID"].Visible = false;
            if (grid.Columns.Contains("CabFurAssignmentID"))
                grid.Columns["CabFurAssignmentID"].Visible = false;
            if (grid.Columns.Contains("CreationUserID"))
                grid.Columns["CreationUserID"].Visible = false;
            if (grid.Columns.Contains("TechStoreID"))
                grid.Columns["TechStoreID"].Visible = false;
            if (grid.Columns.Contains("MeasureID"))
                grid.Columns["MeasureID"].Visible = false;
            if (grid.Columns.Contains("InsetColorID"))
                grid.Columns["InsetColorID"].Visible = false;
            if (grid.Columns.Contains("InsetTypeID"))
                grid.Columns["InsetTypeID"].Visible = false;
            if (grid.Columns.Contains("PatinaID"))
                grid.Columns["PatinaID"].Visible = false;
            if (grid.Columns.Contains("CoverID"))
                grid.Columns["CoverID"].Visible = false;
            if (grid.Columns.Contains("ColorID"))
                grid.Columns["ColorID"].Visible = false;
        }

        private void AddItemButton_Click(object sender, EventArgs e)
        {
            int TechStoreID = 0;
            int CoverID = 0;
            int PatinaID = 0;
            int InsetColorID = 0;
            int Count = 0;
            if (cmbxTechStore.SelectedItem != null && ((DataRowView)cmbxTechStore.SelectedItem).Row["TechStoreID"] != DBNull.Value)
                TechStoreID = Convert.ToInt32(((DataRowView)cmbxTechStore.SelectedItem).Row["TechStoreID"]);
            if (cmbxCovers.SelectedItem != null && ((DataRowView)cmbxCovers.SelectedItem).Row["CoverID"] != DBNull.Value)
                CoverID = Convert.ToInt32(((DataRowView)cmbxCovers.SelectedItem).Row["CoverID"]);
            if (cmbxPatina.SelectedItem != null && ((DataRowView)cmbxPatina.SelectedItem).Row["PatinaID"] != DBNull.Value)
                PatinaID = Convert.ToInt32(((DataRowView)cmbxPatina.SelectedItem).Row["PatinaID"]);
            if (cmbxInsetColors.SelectedItem != null && ((DataRowView)cmbxInsetColors.SelectedItem).Row["InsetColorID"] != DBNull.Value)
                InsetColorID = Convert.ToInt32(((DataRowView)cmbxInsetColors.SelectedItem).Row["InsetColorID"]);
            if (tbCount.Text.Length > 0)
                Count = Convert.ToInt32(tbCount.Text);
            assignmentsManager.AddAssignmentDetail(TechStoreID, CoverID, PatinaID, InsetColorID, Count);
            CheckColumns(ref dgvNewAssignment);
        }

        private void RemoveItemButton_Click(object sender, EventArgs e)
        {
            assignmentsManager.RemoveNewAssignmentDetail();
            CheckColumns(ref dgvNewAssignment);
        }

        private void SaveItemsButton_Click(object sender, EventArgs e)
        {
            Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Сохранение.\r\nПодождите..."); });
            T.Start();
            while (!SplashWindow.bSmallCreated) ;
            NeedSplash = false;

            DateTime date1 = CreateDateTimePicker1.Value.Date;
            DateTime date2 = CreateDateTimePicker2.Value.Date;

            if (assignmentsManager.NewAssignment)
            {
                assignmentsManager.AddAssignment();
                assignmentsManager.FillAssignmentID(assignmentsManager.CabFurAssignmentID);
                assignmentsManager.SaveNewAssignment();
                assignmentsManager.UpdateNewAssignment(assignmentsManager.CabFurAssignmentID);
            }
            else
            {
                assignmentsManager.SaveNewAssignment();
                assignmentsManager.SaveAllAssignments();
            }

            assignmentsManager.UpdateDocuments(date1, date2);
            assignmentsManager.UpdateAllAssignments(date1, date2);

            if (assignmentsManager.NewAssignment)
            {
                assignmentsManager.MoveToFirstAssignmentID();
            }
            else
            {
                assignmentsManager.MoveToAssignmentID(assignmentsManager.CabFurAssignmentID);
            }

            NeedSplash = true;
            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
        }

        private void dgvNewAssignment_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            //if (e.Button == System.Windows.Forms.MouseButtons.Right)
            //{
            //    ComponentFactory.Krypton.Toolkit.KryptonContextMenu kryptonContextMenu = new ComponentFactory.Krypton.Toolkit.KryptonContextMenu();
            //    int FrontID = Convert.ToInt32(dgvNewAssignment.SelectedRows[0].Cells["TechStoreID"].Value);

            //    if (TestAssignmentsManager == null)
            //    {
            //        TestAssignmentsManager = new TestAssignments();
            //        TestAssignmentsManager.Initialize();
            //    }
            //    DataTable DT = TestAssignmentsManager.GetOperationsGroups(FrontID);
            //    if (DT.Rows.Count > 0)
            //    {
            //        ComponentFactory.Krypton.Toolkit.KryptonContextMenuItems kryptonContextMenuItems = new ComponentFactory.Krypton.Toolkit.KryptonContextMenuItems();
            //        for (int i = 0; i < DT.Rows.Count; i++)
            //        {
            //            ComponentFactory.Krypton.Toolkit.KryptonContextMenuItem kryptonContextMenuItem = new ComponentFactory.Krypton.Toolkit.KryptonContextMenuItem();
            //            kryptonContextMenuItem.Text = DT.Rows[i]["GroupName"].ToString();
            //            kryptonContextMenuItem.Tag = Convert.ToInt32(DT.Rows[i]["TechCatalogOperationsGroupID"]);
            //            kryptonContextMenuItem.Click += kryptonContextMenuItem_Click;
            //            kryptonContextMenuItems.Items.Add(kryptonContextMenuItem);
            //        }
            //        kryptonContextMenu.Items.Add(kryptonContextMenuItems);
            //        kryptonContextMenu.Show(new Point(Cursor.Position.X - 212, Cursor.Position.Y - 10));
            //    }
            //}
        }

        private void btnCalcMaterial_Click(object sender, EventArgs e)
        {
            int CabFurAssignmentID = 0;
            if (dgvAllAssignments.SelectedRows.Count != 0 && dgvAllAssignments.SelectedRows[0].Cells["CabFurAssignmentID"].Value != DBNull.Value)
                CabFurAssignmentID = Convert.ToInt32(dgvAllAssignments.SelectedRows[0].Cells["CabFurAssignmentID"].Value);
            int TechStoreID = 0;
            int CoverID = 0;
            int PatinaID = 0;
            int Length = 0;
            int Height = 0;
            int Width = 0;
            int Count = 0;
            string TechStoreName = string.Empty;
            string CoverName = string.Empty;
            string PatinaName = string.Empty;
            if (dgvNewAssignment.SelectedRows[0].Cells["TechStoreID"].Value != DBNull.Value)
            {
                TechStoreID = Convert.ToInt32(dgvNewAssignment.SelectedRows[0].Cells["TechStoreID"].Value);
                TechStoreName = dgvNewAssignment.SelectedRows[0].Cells["TechStoreNameColumn"].FormattedValue.ToString();
            }
            if (dgvNewAssignment.SelectedRows[0].Cells["CoverID"].Value != DBNull.Value)
            {
                CoverID = Convert.ToInt32(dgvNewAssignment.SelectedRows[0].Cells["CoverID"].Value);
                CoverName = dgvNewAssignment.SelectedRows[0].Cells["CoverColumn"].FormattedValue.ToString();
            }
            if (dgvNewAssignment.SelectedRows[0].Cells["PatinaID"].Value != DBNull.Value)
            {
                PatinaID = Convert.ToInt32(dgvNewAssignment.SelectedRows[0].Cells["PatinaID"].Value);
                PatinaName = dgvNewAssignment.SelectedRows[0].Cells["PatinaColumn"].FormattedValue.ToString();
            }
            if (dgvNewAssignment.SelectedRows[0].Cells["Length"].Value != DBNull.Value)
                Length = Convert.ToInt32(dgvNewAssignment.SelectedRows[0].Cells["Length"].Value);
            if (dgvNewAssignment.SelectedRows[0].Cells["Height"].Value != DBNull.Value)
                Height = Convert.ToInt32(dgvNewAssignment.SelectedRows[0].Cells["Height"].Value);
            if (dgvNewAssignment.SelectedRows[0].Cells["Width"].Value != DBNull.Value)
                Width = Convert.ToInt32(dgvNewAssignment.SelectedRows[0].Cells["Width"].Value);
            if (dgvNewAssignment.SelectedRows[0].Cells["Count"].Value != DBNull.Value)
                Count = Convert.ToInt32(dgvNewAssignment.SelectedRows[0].Cells["Count"].Value);


            Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Сохранение данных.\r\nПодождите..."); });
            T.Start();
            while (!SplashWindow.bSmallCreated) ;
            NeedSplash = false;

            if (calcManager == null)
            {
                calcManager = new CalculateMaterial(assignmentsManager);
                calcManager.Initialize();
            }
            calcManager.MainFunction(CabFurAssignmentID, TechStoreName, TechStoreID, CoverID, PatinaID, Length, Height, Width, CoverName, PatinaName, Count);

            NeedSplash = true;
            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;

            //CabFurnitureOperationsForm CabFurnitureOperationsForm = new CabFurnitureOperationsForm(this, TechStoreID, CoverID, PatinaID, Length, Height, Width, CoverName, PatinaName,
            //    AssignmentsManager);

            //TopForm = CabFurnitureOperationsForm;

            //CabFurnitureOperationsForm.ShowDialog();

            //CabFurnitureOperationsForm.Close();
            //CabFurnitureOperationsForm.Dispose();
        }

        private void btnAddAssignment_Click(object sender, EventArgs e)
        {
            bool OKCancel = Infinium.LightMessageBox.Show(ref TopForm, true,
                    "Вы собираетесь создать новое задание. Продолжить?",
                    "Новое задание");
            if (!OKCancel)
                return;

            Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Сохранение.\r\nПодождите..."); });
            T.Start();
            while (!SplashWindow.bSmallCreated) ;
            NeedSplash = false;

            kryptonCheckSet3.CheckedButton = cbtnNewAssignment;
            assignmentsManager.NewAssignment = true;
            assignmentsManager.CabFurAssignmentID = 0;
            assignmentsManager.UpdateNewAssignment(0);

            NeedSplash = true;
            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
        }

        private void btnEditAssignment_Click(object sender, EventArgs e)
        {

        }

        private void btnDeleteAssignment_Click(object sender, EventArgs e)
        {

        }

        private void kryptonContextMenuItem5_Click(object sender, EventArgs e)
        {
            int CabFurAssignmentID = 0;
            if (dgvAllAssignments.SelectedRows.Count != 0 && dgvAllAssignments.SelectedRows[0].Cells["CabFurAssignmentID"].Value != DBNull.Value)
                CabFurAssignmentID = Convert.ToInt32(dgvAllAssignments.SelectedRows[0].Cells["CabFurAssignmentID"].Value);

            Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Сохранение данных.\r\nПодождите..."); });
            T.Start();
            while (!SplashWindow.bSmallCreated) ;
            NeedSplash = false;

            assignmentsManager.PrintAssignment(CabFurAssignmentID);
            assignmentsManager.SaveDocuments();
            FilterAssignments();
            assignmentsManager.MoveToAssignmentID(CabFurAssignmentID);

            NeedSplash = true;
            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
        }

        private void MenuButton_Click(object sender, EventArgs e)
        {
            if (kryptonCheckSet3.CheckedButton == cbtnStatistics)
            {
                MenuPanel2.Visible = !MenuPanel2.Visible;
            }
            else
            {
                MenuPanel.Visible = !MenuPanel.Visible;
            }
        }

        private void dgvDocuments_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex == -1)
                return;
            PercentageDataGrid grid = (PercentageDataGrid)sender;

            int CabFurAssignmentID = -1;
            string FileName = string.Empty;

            if (grid.SelectedRows[0].Cells["FileName"].Value != DBNull.Value)
            {
                CabFurAssignmentID = Convert.ToInt32(grid.SelectedRows[0].Cells["CabFurAssignmentID"].Value);
                FileName = grid.SelectedRows[0].Cells["FileName"].Value.ToString();
            }
            string temppath = string.Empty;
            if (CabFurAssignmentID != -1)
            {
                PhantomForm PhantomForm = new Infinium.PhantomForm();
                PhantomForm.Show();

                CabFurDocsDownloadForm TPSProdDocsDownloadForm = new CabFurDocsDownloadForm(CabFurAssignmentID, FileName, ref assignmentsManager);
                TopForm = TPSProdDocsDownloadForm;
                TPSProdDocsDownloadForm.ShowDialog();

                PhantomForm.Close();
                PhantomForm.Dispose();
                TopForm = null;
                TPSProdDocsDownloadForm.Dispose();
            }
        }

        private void dgvAllAssignments_SelectionChanged(object sender, EventArgs e)
        {
            int CabFurAssignmentID = 0;
            int PackagesCount = 0;
            if (dgvAllAssignments.SelectedRows.Count != 0 && dgvAllAssignments.SelectedRows[0].Cells["CabFurAssignmentID"].Value != DBNull.Value)
                CabFurAssignmentID = Convert.ToInt32(dgvAllAssignments.SelectedRows[0].Cells["CabFurAssignmentID"].Value);
            if (dgvAllAssignments.SelectedRows.Count != 0 && dgvAllAssignments.SelectedRows[0].Cells["PackagesCount"].Value != DBNull.Value)
                PackagesCount = Convert.ToInt32(dgvAllAssignments.SelectedRows[0].Cells["PackagesCount"].Value);
            if (NeedSplash)
            {
                Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                T.Start();
                //while (!SplashWindow.bSmallCreated) ;
                NeedSplash = false;

                assignmentsManager.FilterDocuments(CabFurAssignmentID);

                assignmentsManager.UpdateNewAssignment(CabFurAssignmentID);
                CheckColumns(ref percentageDataGrid1);

                NeedSplash = true;
                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;
            }
            else
            {
                assignmentsManager.FilterDocuments(CabFurAssignmentID);

                assignmentsManager.UpdateNewAssignment(CabFurAssignmentID);
                CheckColumns(ref percentageDataGrid1);
            }
        }

        private void FilterAssignments()
        {
            bool bAgreed = cbAgreedAssingments.Checked;
            bool bNotAgreed = cbNotAgreedAssingments.Checked;
            bool bPrinted = cbPrintedAssingments.Checked;
            bool bNotPrinted = cbNotPrintedAssingments.Checked;
            bool bPartPrinted = cbPartPrintedAssingments.Checked;

            DateTime date1 = CreateDateTimePicker1.Value.Date;
            DateTime date2 = CreateDateTimePicker2.Value.Date;

            if (NeedSplash)
            {
                Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                T.Start();
                while (!SplashWindow.bSmallCreated) ;
                NeedSplash = false;

                assignmentsManager.UpdateDocuments(date1, date2);
                assignmentsManager.UpdateAllAssignments(date1, date2);
                assignmentsManager.FilterAllAssignments(bAgreed, bNotAgreed, bPrinted, bNotPrinted, bPartPrinted);

                NeedSplash = true;
                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;
            }
            else
            {
                assignmentsManager.UpdateDocuments(date1, date2);
                assignmentsManager.UpdateAllAssignments(date1, date2);
                assignmentsManager.FilterAllAssignments(bAgreed, bNotAgreed, bPrinted, bNotPrinted, bPartPrinted);
            }
        }

        private void FilterComplements()
        {
            bool bPacked = kryptonCheckBox7.Checked;
            bool bNotPacked = kryptonCheckBox6.Checked;
            bool bPartPacked = kryptonCheckBox5.Checked;

            bool bPrinted = kryptonCheckBox4.Checked;
            bool bNotPrinted = kryptonCheckBox3.Checked;
            bool bPartPrinted = kryptonCheckBox2.Checked;

            if (NeedSplash)
            {
                Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                T.Start();
                while (!SplashWindow.bSmallCreated) ;
                NeedSplash = false;

                complementsManager.FilterComplements(bPacked, bNotPacked, bPartPacked, bPrinted, bNotPrinted, bPartPrinted);

                NeedSplash = true;
                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;
            }
            else
            {
                complementsManager.FilterComplements(bPacked, bNotPacked, bPartPacked, bPrinted, bNotPrinted, bPartPrinted);
            }
        }

        private void FilterPackages()
        {
            bool bPacked = kryptonCheckBox10.Checked;
            bool bNotPacked = kryptonCheckBox9.Checked;
            bool bPartPacked = kryptonCheckBox8.Checked;

            bool bPrinted = kryptonCheckBox13.Checked;
            bool bNotPrinted = kryptonCheckBox12.Checked;
            bool bPartPrinted = kryptonCheckBox11.Checked;

            if (NeedSplash)
            {
                Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                T.Start();
                while (!SplashWindow.bSmallCreated) ;
                NeedSplash = false;

                packagesManager.FilterPackages(bPacked, bNotPacked, bPartPacked, bPrinted, bNotPrinted, bPartPrinted);

                NeedSplash = true;
                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;
            }
            else
            {
                packagesManager.FilterPackages(bPacked, bNotPacked, bPartPacked, bPrinted, bNotPrinted, bPartPrinted);
            }
        }

        private void cbAgreedAssingments_CheckedChanged(object sender, EventArgs e)
        {
            FilterAssignments();
        }

        private void cbNotAgreedAssingments_CheckedChanged(object sender, EventArgs e)
        {
            FilterAssignments();
        }

        private void cbPrintedAssingments_CheckedChanged(object sender, EventArgs e)
        {
            FilterAssignments();
        }

        private void cbNotPrintedAssingments_CheckedChanged(object sender, EventArgs e)
        {
            FilterAssignments();
        }

        private void cbPartPrintedAssingments_CheckedChanged(object sender, EventArgs e)
        {
            FilterAssignments();
        }

        private void dgvDocuments_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Right && e.RowIndex != -1)
            {
                kryptonContextMenu1.Show(new Point(Cursor.Position.X - 212, Cursor.Position.Y - 10));
            }
        }

        private void kryptonContextMenuItem12_Click(object sender, EventArgs e)
        {
            int CabFurAssignmentID = 0;
            int CabFurDocumentID = -1;
            string FileName = string.Empty;
            if (dgvAllAssignments.SelectedRows.Count != 0 && dgvAllAssignments.SelectedRows[0].Cells["CabFurAssignmentID"].Value != DBNull.Value)
                CabFurAssignmentID = Convert.ToInt32(dgvAllAssignments.SelectedRows[0].Cells["CabFurAssignmentID"].Value);
            if (dgvDocuments.SelectedRows.Count != 0 && dgvDocuments.SelectedRows[0].Cells["CabFurDocumentID"].Value != DBNull.Value)
                CabFurDocumentID = Convert.ToInt32(dgvDocuments.SelectedRows[0].Cells["CabFurDocumentID"].Value);
            if (dgvDocuments.SelectedRows.Count != 0 && dgvDocuments.SelectedRows[0].Cells["FileName"].Value != DBNull.Value)
                FileName = dgvDocuments.SelectedRows[0].Cells["FileName"].Value.ToString();

            var T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Сохранение данных.\r\nПодождите..."); });
            T.Start();
            while (!SplashWindow.bSmallCreated) ;
            NeedSplash = false;

            //write printdate
            assignmentsManager.PrintDocument(CabFurDocumentID);
            assignmentsManager.SaveDocuments();
            FilterAssignments();
            assignmentsManager.MoveToAssignmentID(CabFurAssignmentID);

            string temppath = string.Empty;
            if (CabFurAssignmentID != -1)
            {
                PhantomForm PhantomForm = new PhantomForm();
                PhantomForm.Show();

                //open document from ftp
                CabFurDocsDownloadForm TPSProdDocsDownloadForm = new CabFurDocsDownloadForm(CabFurAssignmentID, FileName, ref assignmentsManager);
                TopForm = TPSProdDocsDownloadForm;
                TPSProdDocsDownloadForm.ShowDialog();

                PhantomForm.Close();
                PhantomForm.Dispose();
                TopForm = null;
                TPSProdDocsDownloadForm.Dispose();
            }
        }

        private void dgvDocuments_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            if (e.RowIndex < 0)
                return;
            PercentageDataGrid grid = (PercentageDataGrid)sender;
            int Printed = 0;
            if (grid.Rows[e.RowIndex].Cells["PrintUserID"].Value != DBNull.Value)
                Printed = 1;

            if (Printed == 1)
            {
                grid.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Green;
                grid.Rows[e.RowIndex].DefaultCellStyle.ForeColor = Color.White;
                grid.Rows[e.RowIndex].DefaultCellStyle.SelectionForeColor = Color.Green;
            }
            if (Printed == 0)
            {
                grid.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.White;
                grid.Rows[e.RowIndex].DefaultCellStyle.ForeColor = Color.Black;
                grid.Rows[e.RowIndex].DefaultCellStyle.SelectionForeColor = Color.White;
            }
        }

        private void dgvAllAssignments_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            if (e.RowIndex < 0)
                return;
            PercentageDataGrid grid = (PercentageDataGrid)sender;
            bool InProd = false;
            bool NotOutProd = false;
            bool OutProd = false;
            if (grid.Rows[e.RowIndex].Cells["ProductionDateTime"].Value == DBNull.Value && grid.Rows[e.RowIndex].Cells["OutProductionDateTime"].Value == DBNull.Value)
                InProd = true;
            if (grid.Rows[e.RowIndex].Cells["ProductionDateTime"].Value != DBNull.Value && grid.Rows[e.RowIndex].Cells["OutProductionDateTime"].Value == DBNull.Value)
                NotOutProd = true;
            if (grid.Rows[e.RowIndex].Cells["ProductionDateTime"].Value != DBNull.Value && grid.Rows[e.RowIndex].Cells["OutProductionDateTime"].Value != DBNull.Value)
                OutProd = true;

            if (InProd)
            {
                grid.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Red;
                grid.Rows[e.RowIndex].DefaultCellStyle.ForeColor = Color.White;
                grid.Rows[e.RowIndex].DefaultCellStyle.SelectionForeColor = Color.Red;
            }
            if (NotOutProd)
            {
                grid.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Yellow;
                grid.Rows[e.RowIndex].DefaultCellStyle.ForeColor = Color.Black;
                grid.Rows[e.RowIndex].DefaultCellStyle.SelectionForeColor = Color.Yellow;
            }
            if (OutProd)
            {
                grid.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Green;
                grid.Rows[e.RowIndex].DefaultCellStyle.ForeColor = Color.White;
                grid.Rows[e.RowIndex].DefaultCellStyle.SelectionForeColor = Color.Green;
            }
        }

        private void kryptonCheckButton1_Click(object sender, EventArgs e)
        {
        }

        private void btnCoversManager_Click(object sender, EventArgs e)
        {
            PhantomForm PhantomForm = new Infinium.PhantomForm();
            PhantomForm.Show();

            if (CabFurnitureCoversForm == null)
                CabFurnitureCoversForm = new CabFurnitureCoversForm(this, assignmentsManager);

            TopForm = CabFurnitureCoversForm;
            CabFurnitureCoversForm.ShowDialog();

            PhantomForm.Close();

            PhantomForm.Dispose();
            //CabFurnitureCoversForm.Dispose();
            TopForm = null;
        }

        private void kryptonCheckBox1_CheckedChanged(object sender, EventArgs e)
        {
            Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
            T.Start();
            while (!SplashWindow.bSmallCreated) ;
            NeedSplash = false;

            NonAgreementCount = 0;
            AgreedCount = 0;
            OnProdCount = 0;
            InProdCount = 0;
            OnStoreCount = 0;
            OnExpCount = 0;

            assignmentsManager.NonAgreementOrders(kryptonCheckBox1.Checked, ref NonAgreementCount);
            assignmentsManager.AgreedOrders(kryptonCheckBox1.Checked, ref AgreedCount);
            assignmentsManager.OnProductionOrders(kryptonCheckBox1.Checked, ref OnProdCount);
            assignmentsManager.InProductionOrders(kryptonCheckBox1.Checked, ref InProdCount);
            assignmentsManager.OnStorageOrders(kryptonCheckBox1.Checked, ref OnStoreCount);
            assignmentsManager.OnExpeditionOrders(kryptonCheckBox1.Checked, ref OnExpCount);

            label11.Text = NonAgreementCount.ToString() + " шт.";
            label6.Text = AgreedCount.ToString() + " шт.";
            label27.Text = OnProdCount.ToString() + " шт.";
            label16.Text = InProdCount.ToString() + " шт.";
            label41.Text = OnStoreCount.ToString() + " шт.";
            label34.Text = OnExpCount.ToString() + " шт.";

            NonAgreementDataGrid.DataSource = null;
            AgreedDataGrid.DataSource = null;
            OnProductionDataGrid.DataSource = null;
            InProductionDataGrid.DataSource = null;
            OnStorageDataGrid.DataSource = null;
            OnExpeditionDataGrid.DataSource = null;

            NonAgreementDataGrid.DataSource = assignmentsManager.NonAgreementDetailBS;
            AgreedDataGrid.DataSource = assignmentsManager.AgreedDetailBS;
            OnProductionDataGrid.DataSource = assignmentsManager.OnProductionDetailBS;
            InProductionDataGrid.DataSource = assignmentsManager.InProductionDetailBS;
            OnStorageDataGrid.DataSource = assignmentsManager.OnStorageDetailBS;
            OnExpeditionDataGrid.DataSource = assignmentsManager.OnExpeditionDetailBS;

            dgvDetailsSetting(kryptonCheckBox1.Checked, ref NonAgreementDataGrid);
            dgvDetailsSetting(kryptonCheckBox1.Checked, ref AgreedDataGrid);
            dgvDetailsSetting(kryptonCheckBox1.Checked, ref OnProductionDataGrid);
            dgvDetailsSetting(kryptonCheckBox1.Checked, ref InProductionDataGrid);
            dgvDetailsSetting(kryptonCheckBox1.Checked, ref OnStorageDataGrid);
            dgvDetailsSetting(kryptonCheckBox1.Checked, ref OnExpeditionDataGrid);

            NeedSplash = true;
            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
        }

        private void cmbxTechStore_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (assignmentsManager == null)
                return;
            int TechStoreID = 0;
            if (cmbxTechStore.SelectedItem != null && ((DataRowView)cmbxTechStore.SelectedItem).Row["TechStoreID"] != DBNull.Value)
                TechStoreID = Convert.ToInt32(((DataRowView)cmbxTechStore.SelectedItem).Row["TechStoreID"]);
            //AssignmentsManager.FilterCovers(TechStoreID);
            //cmbxCovers_SelectedIndexChanged(null, null);
        }

        private void cmbxCovers_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (assignmentsManager == null)
                return;
            int TechStoreID = 0;
            int CoverID = 0;
            if (cmbxTechStore.SelectedItem != null && ((DataRowView)cmbxTechStore.SelectedItem).Row["TechStoreID"] != DBNull.Value)
                TechStoreID = Convert.ToInt32(((DataRowView)cmbxTechStore.SelectedItem).Row["TechStoreID"]);
            if (cmbxCovers.SelectedItem != null && ((DataRowView)cmbxCovers.SelectedItem).Row["CoverID"] != DBNull.Value)
                CoverID = Convert.ToInt32(((DataRowView)cmbxCovers.SelectedItem).Row["CoverID"]);
            //AssignmentsManager.FilterPatina(TechStoreID, CoverID);
        }

        private void dgvDocuments_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            var senderGrid = (PercentageDataGrid)sender;

            if (senderGrid.Columns[e.ColumnIndex] is KryptonDataGridViewButtonColumn &&
                e.RowIndex >= 0)
            {
                int CabFurAssignmentID = 0;
                int CabFurDocumentID = -1;
                string FileName = string.Empty;
                if (senderGrid.Rows[e.RowIndex].Cells["CabFurAssignmentID"].Value != DBNull.Value)
                    CabFurAssignmentID = Convert.ToInt32(senderGrid.Rows[e.RowIndex].Cells["CabFurAssignmentID"].Value);
                if (senderGrid.Rows[e.RowIndex].Cells["CabFurDocumentID"].Value != DBNull.Value)
                    CabFurDocumentID = Convert.ToInt32(senderGrid.Rows[e.RowIndex].Cells["CabFurDocumentID"].Value);
                if (senderGrid.Rows[e.RowIndex].Cells["FileName"].Value != DBNull.Value)
                    FileName = senderGrid.Rows[e.RowIndex].Cells["FileName"].Value.ToString();

                var T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Сохранение данных.\r\nПодождите..."); });
                T.Start();
                while (!SplashWindow.bSmallCreated) ;
                NeedSplash = false;

                //write printdate
                assignmentsManager.PrintDocument(CabFurDocumentID);
                assignmentsManager.SaveDocuments();
                FilterAssignments();
                assignmentsManager.MoveToAssignmentID(CabFurAssignmentID);

                NeedSplash = true;
                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;

                string temppath = string.Empty;
                if (CabFurAssignmentID != -1)
                {
                    PhantomForm PhantomForm = new PhantomForm();
                    PhantomForm.Show();

                    //open document from ftp
                    CabFurDocsDownloadForm TPSProdDocsDownloadForm = new CabFurDocsDownloadForm(CabFurAssignmentID, FileName, ref assignmentsManager);
                    TopForm = TPSProdDocsDownloadForm;
                    TPSProdDocsDownloadForm.ShowDialog();

                    PhantomForm.Close();
                    PhantomForm.Dispose();
                    TopForm = null;
                    TPSProdDocsDownloadForm.Dispose();
                }

            }
        }

        private void dgvAllAssignments_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (userRole != RoleTypes.AdminRole && userRole != RoleTypes.AgreementRole)
                return;

            var senderGrid = (PercentageDataGrid)sender;

            if (senderGrid.Columns[e.ColumnIndex] is KryptonDataGridViewButtonColumn && senderGrid.Columns[e.ColumnIndex].Name == "InProdColumn" &&
                e.RowIndex >= 0)
            {
                int CabFurAssignmentID = 0;
                if (senderGrid.Rows[e.RowIndex].Cells["CabFurAssignmentID"].Value != DBNull.Value)
                    CabFurAssignmentID = Convert.ToInt32(senderGrid.Rows[e.RowIndex].Cells["CabFurAssignmentID"].Value);
                
                if (assignmentsManager.IsAssignmentInProd(CabFurAssignmentID))
                {
                    Infinium.LightMessageBox.Show(ref TopForm, false,
                    "Задание уже вошло в производство",
                    "Запуск в производство");
                    return;
                }

                bool OKCancel = Infinium.LightMessageBox.Show(ref TopForm, true,
                    "Подтвердите, что задание входит в производство",
                    "Запуск в производство", "Подтвердить", "Отмена");
                if (!OKCancel)
                    return;

                int FactoryID = 1;
                bool bTPS = false;

                PhantomForm PhantomForm = new Infinium.PhantomForm();
                PhantomForm.Show();

                CabFurSelectFactoryForm cabFurSelectFactoryForm = new CabFurSelectFactoryForm();

                TopForm = cabFurSelectFactoryForm;
                cabFurSelectFactoryForm.ShowDialog();

                PhantomForm.Close();

                bTPS = cabFurSelectFactoryForm.bTPS;
                OKCancel = cabFurSelectFactoryForm.OKCancel;

                PhantomForm.Dispose();
                cabFurSelectFactoryForm.Dispose();
                TopForm = null;

                if (!OKCancel)
                    return;
                if (bTPS)
                    FactoryID = 2;


                var T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Сохранение данных.\r\nПодождите..."); });
                T.Start();
                while (!SplashWindow.bSmallCreated) ;
                NeedSplash = false;
                
                int AllPackagesCount = assignmentsManager.CreatePackages(CabFurAssignmentID, FactoryID);
                assignmentsManager.PrintAssignment(CabFurAssignmentID);
                assignmentsManager.InProduction(CabFurAssignmentID, AllPackagesCount);
                assignmentsManager.SaveAllAssignments();
                FilterAssignments();
                assignmentsManager.MoveToAssignmentID(CabFurAssignmentID);

                NeedSplash = true;
                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;
            }

            if (senderGrid.Columns[e.ColumnIndex] is KryptonDataGridViewButtonColumn && senderGrid.Columns[e.ColumnIndex].Name == "OutProdColumn" &&
                e.RowIndex >= 0)
            {
                bool OKCancel = Infinium.LightMessageBox.Show(ref TopForm, true,
                    "Подтвердите, что задание вышло из производства",
                    "Выход из производства", "Подтвердить", "Отмена");
                if (!OKCancel)
                    return;

                int CabFurAssignmentID = 0;
                if (senderGrid.Rows[e.RowIndex].Cells["CabFurAssignmentID"].Value != DBNull.Value)
                    CabFurAssignmentID = Convert.ToInt32(senderGrid.Rows[e.RowIndex].Cells["CabFurAssignmentID"].Value);

                var T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Сохранение данных.\r\nПодождите..."); });
                T.Start();
                while (!SplashWindow.bSmallCreated) ;
                NeedSplash = false;
                
                assignmentsManager.OutProduction(CabFurAssignmentID);
                assignmentsManager.SaveAllAssignments();
                FilterAssignments();
                assignmentsManager.MoveToAssignmentID(CabFurAssignmentID);

                NeedSplash = true;
                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;
            }
        }

        private void kryptonContextMenuItem10_Click(object sender, EventArgs e)
        {
            bool OKCancel = Infinium.LightMessageBox.Show(ref TopForm, true,
                    "Вы собираетесь учистить Упаковку. Продолжить?",
                    "Удаление упаковок");
            if (!OKCancel)
                return;

            int CabFurAssignmentID = 0;
            if (dgvAllAssignments.SelectedRows.Count != 0 && dgvAllAssignments.SelectedRows[0].Cells["CabFurAssignmentID"].Value != DBNull.Value)
                CabFurAssignmentID = Convert.ToInt32(dgvAllAssignments.SelectedRows[0].Cells["CabFurAssignmentID"].Value);

            Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Сохранение данных.\r\nПодождите..."); });
            T.Start();
            while (!SplashWindow.bSmallCreated) ;
            NeedSplash = false;

            assignmentsManager.ClearCabFurniturePackages(CabFurAssignmentID);
            FilterAssignments();
            assignmentsManager.MoveToAssignmentID(CabFurAssignmentID);

            NeedSplash = true;
            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
        }

        private void kryptonContextMenuItem11_Click(object sender, EventArgs e)
        {
        }

        private void dgvComplementLabels_SelectionChanged(object sender, EventArgs e)
        {
            int CabFurnitureComplementID = 0;
            if (dgvComplementLabels.SelectedRows.Count != 0 && dgvComplementLabels.SelectedRows[0].Cells["CabFurnitureComplementID"].Value != DBNull.Value)
                CabFurnitureComplementID = Convert.ToInt32(dgvComplementLabels.SelectedRows[0].Cells["CabFurnitureComplementID"].Value);
            complementsManager.FilterComplementDetails(CabFurnitureComplementID);
        }

        private void dgvComplementLabels_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Right && e.RowIndex != -1)
            {
                kryptonContextMenu3.Show(new Point(Cursor.Position.X - 212, Cursor.Position.Y - 10));
            }
        }

        private void kryptonContextMenuItem15_Click(object sender, EventArgs e)
        {
            if (dgvComplementLabels.SelectedRows.Count == 0)
                return;
            complementLabel.ClearLabelInfo();

            int ComplementsCount = 0;
            int MainOrderID = 0;

            if (dgvComplements.SelectedRows.Count != 0 && dgvComplements.SelectedRows[0].Cells["ComplementsCount"].Value != DBNull.Value)
                ComplementsCount = Convert.ToInt32(dgvComplements.SelectedRows[0].Cells["ComplementsCount"].Value);
            if (dgvMainOrders.SelectedRows.Count != 0 && dgvMainOrders.SelectedRows[0].Cells["MainOrderID"].Value != DBNull.Value)
                MainOrderID = Convert.ToInt32(dgvMainOrders.SelectedRows[0].Cells["MainOrderID"].Value);

            int[] CabFurnitureComplementID = new int[dgvComplementLabels.SelectedRows.Count];
            int[] Index = new int[dgvComplementLabels.SelectedRows.Count];
            for (int i = 0; i < dgvComplementLabels.SelectedRows.Count; i++)
            {
                CabFurnitureComplementID[i] = Convert.ToInt32(dgvComplementLabels.SelectedRows[i].Cells["CabFurnitureComplementID"].Value);
                Index[i] = Convert.ToInt32(dgvComplementLabels.SelectedRows[i].Cells["Index"].Value);
            }
            Array.Sort(CabFurnitureComplementID);
            Array.Sort(Index);
            List<ComplementLabelInfo> Labels = assignmentsManager.CreateComplementLabels(CabFurnitureComplementID, Index, ComplementsCount);
            if (Labels.Count == 0)
                return;
            for (int i = 0; i < Labels.Count; i++)
            {
                ComplementLabelInfo LabelInfo = Labels[i];
                complementLabel.AddLabelInfo(ref LabelInfo);
            }
            PrintDialog.Document = complementLabel.PD;

            if (PrintDialog.ShowDialog() == DialogResult.OK)
            {
                complementLabel.Print();
            }

            complementsManager.PrintComplements(CabFurnitureComplementID);

            if (NeedSplash)
            {
                Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                T.Start();
                //while (!SplashWindow.bSmallCreated) ;
                NeedSplash = false;

                complementsManager.FilterComplementsLabels(MainOrderID);

                NeedSplash = true;
                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;
            }
            else
            {
                complementsManager.FilterComplementsLabels(MainOrderID);
            }
        }

        private void dgvComplements_SelectionChanged(object sender, EventArgs e)
        {
            int ClientID = 0;
            int OrderNumber = 0;
            if (dgvComplements.SelectedRows.Count != 0 && dgvComplements.SelectedRows[0].Cells["ClientID"].Value != DBNull.Value)
                ClientID = Convert.ToInt32(dgvComplements.SelectedRows[0].Cells["ClientID"].Value);
            if (dgvComplements.SelectedRows.Count != 0 && dgvComplements.SelectedRows[0].Cells["OrderNumber"].Value != DBNull.Value)
                OrderNumber = Convert.ToInt32(dgvComplements.SelectedRows[0].Cells["OrderNumber"].Value);


            if (NeedSplash)
            {
                Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Сохранение данных.\r\nПодождите..."); });
                T.Start();
                while (!SplashWindow.bSmallCreated) ;
                NeedSplash = false;

                complementsManager.FilterMainOrders(ClientID, OrderNumber);
                //ComplementsManager.ClearComplements();

                NeedSplash = true;
                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;
            }
            else
            {
                complementsManager.FilterMainOrders(ClientID, OrderNumber);
                //ComplementsManager.ClearComplements();
            }

        }

        private void dgvMainOrders_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            //int MainOrderID = 0;
            //if (dgvMainOrders.SelectedRows.Count != 0 && dgvMainOrders.SelectedRows[0].Cells["MainOrderID"].Value != DBNull.Value)
            //    MainOrderID = Convert.ToInt32(dgvMainOrders.SelectedRows[0].Cells["MainOrderID"].Value);

            //Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Сохранение данных.\r\nПодождите..."); });
            //T.Start();
            //while (!SplashWindow.bSmallCreated) ;

            //ComplementsManager.FilterComplementsLabels(MainOrderID);

            //while (SplashWindow.bSmallCreated)
            //    SmallWaitForm.CloseS = true;
        }

        private void dgvPackages_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            //int CabFurAssignmentID = 0;
            //if (dgvPackages.SelectedRows.Count != 0 && dgvPackages.SelectedRows[0].Cells["CabFurAssignmentID"].Value != DBNull.Value)
            //    CabFurAssignmentID = Convert.ToInt32(dgvPackages.SelectedRows[0].Cells["CabFurAssignmentID"].Value);

            //Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Сохранение данных.\r\nПодождите..."); });
            //T.Start();
            //while (!SplashWindow.bSmallCreated) ;

            //PackagesManager.FilterPackagesLabels(CabFurAssignmentID);

            //while (SplashWindow.bSmallCreated)
            //    SmallWaitForm.CloseS = true;
        }

        private void dgvPackagesLabels_SelectionChanged(object sender, EventArgs e)
        {
            int PackNumber = 0;
            int TechStoreSubGroupID = 0;
            int CabFurniturePackageID = 0;
            if (dgvPackagesLabels.SelectedRows.Count != 0 && dgvPackagesLabels.SelectedRows[0].Cells["CabFurniturePackageID"].Value != DBNull.Value)
                CabFurniturePackageID = Convert.ToInt32(dgvPackagesLabels.SelectedRows[0].Cells["CabFurniturePackageID"].Value);
            if (dgvPackagesLabels.SelectedRows.Count != 0 && dgvPackagesLabels.SelectedRows[0].Cells["PackNumber"].Value != DBNull.Value)
                PackNumber = Convert.ToInt32(dgvPackagesLabels.SelectedRows[0].Cells["PackNumber"].Value);
            if (dgvPackagesLabels.SelectedRows.Count != 0 && dgvPackagesLabels.SelectedRows[0].Cells["TechStoreSubGroupID"].Value != DBNull.Value)
                TechStoreSubGroupID = Convert.ToInt32(dgvPackagesLabels.SelectedRows[0].Cells["TechStoreSubGroupID"].Value);
            packagesManager.FilterPackagesDetails(CabFurniturePackageID);
        }

        private void kryptonContextMenuItem16_Click(object sender, EventArgs e)
        {
            if (dgvPackagesLabels.SelectedRows.Count == 0)
                return;
            packageLabel.ClearLabelInfo();

            int CabFurAssignmentID = 0;
            int CabFurAssignmentDetailID = 0;
            if (dgvPackages.SelectedRows.Count != 0 && dgvPackages.SelectedRows[0].Cells["CabFurAssignmentID"].Value != DBNull.Value)
                CabFurAssignmentID = Convert.ToInt32(dgvPackages.SelectedRows[0].Cells["CabFurAssignmentID"].Value);
            if (dgvPackagesLabels.SelectedRows.Count != 0 && dgvPackagesLabels.SelectedRows[0].Cells["CabFurAssignmentDetailID"].Value != DBNull.Value)
                CabFurAssignmentDetailID = Convert.ToInt32(dgvPackagesLabels.SelectedRows[0].Cells["CabFurAssignmentDetailID"].Value);

            int[] CabFurniturePackageID = new int[dgvPackagesLabels.SelectedRows.Count];
            for (int i = 0; i < dgvPackagesLabels.SelectedRows.Count; i++)
            {
                CabFurniturePackageID[i] = Convert.ToInt32(dgvPackagesLabels.SelectedRows[i].Cells["CabFurniturePackageID"].Value);
            }

            List<PackageLabelInfo> Labels = assignmentsManager.CreatePackageLabels(CabFurniturePackageID, CabFurAssignmentID);
            if (Labels.Count == 0)
                return;
            for (int i = 0; i < Labels.Count; i++)
            {
                PackageLabelInfo LabelInfo = Labels[i];
                packageLabel.AddLabelInfo(ref LabelInfo);
            }
            PrintDialog.Document = packageLabel.PD;

            if (PrintDialog.ShowDialog() == DialogResult.OK)
            {
                packageLabel.Print();
            }

            packagesManager.PrintComplements(CabFurniturePackageID);

            if (NeedSplash)
            {
                Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                T.Start();
                //while (!SplashWindow.bSmallCreated) ;
                NeedSplash = false;

                packagesManager.FilterPackagesLabels(CabFurAssignmentID);

                NeedSplash = true;
                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;
            }
            else
            {
                packagesManager.FilterPackagesLabels(CabFurAssignmentID);
            }
        }

        private void dgvPackagesLabels_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Right && e.RowIndex != -1)
            {
                kryptonContextMenu4.Show(new Point(Cursor.Position.X - 212, Cursor.Position.Y - 10));
            }
        }

        private void CheckTimer_Tick(object sender, EventArgs e)
        {
            if (!CanAction)
                return;
            if (tabControl1.SelectedIndex == 5 && !BarcodeTextBox.Focused)
            {
                BarcodeTextBox.Focus();
            }
        }

        private int GetChar(KeyEventArgs e)
        {
            int c = -1;

            if (e.KeyCode == Keys.NumPad1 || e.KeyCode == Keys.NumPad2 || e.KeyCode == Keys.NumPad3 || e.KeyCode == Keys.NumPad4 ||
                e.KeyCode == Keys.NumPad5 || e.KeyCode == Keys.NumPad6 || e.KeyCode == Keys.NumPad7 || e.KeyCode == Keys.NumPad8 ||
                e.KeyCode == Keys.NumPad9 || e.KeyCode == Keys.NumPad0 || e.KeyCode == Keys.D0 || e.KeyCode == Keys.D1 || e.KeyCode == Keys.D2 ||
                e.KeyCode == Keys.D3 || e.KeyCode == Keys.D4 || e.KeyCode == Keys.D5 || e.KeyCode == Keys.D6 || e.KeyCode == Keys.D7 ||
                e.KeyCode == Keys.D8 || e.KeyCode == Keys.D9 || e.KeyCode == Keys.D0)
            {
                switch (e.KeyCode)
                {
                    case Keys.NumPad1:
                        { c = 1; }
                        break;
                    case Keys.NumPad2:
                        { c = 2; }
                        break;
                    case Keys.NumPad3:
                        { c = 3; }
                        break;
                    case Keys.NumPad4:
                        { c = 4; }
                        break;
                    case Keys.NumPad5:
                        { c = 5; }
                        break;
                    case Keys.NumPad6:
                        { c = 6; }
                        break;
                    case Keys.NumPad7:
                        { c = 7; }
                        break;
                    case Keys.NumPad8:
                        { c = 8; }
                        break;
                    case Keys.NumPad9:
                        { c = 9; }
                        break;
                    case Keys.NumPad0:
                        { c = 0; }
                        break;


                    case Keys.D1:
                        { c = 1; }
                        break;
                    case Keys.D2:
                        { c = 2; }
                        break;
                    case Keys.D3:
                        { c = 3; }
                        break;
                    case Keys.D4:
                        { c = 4; }
                        break;
                    case Keys.D5:
                        { c = 5; }
                        break;
                    case Keys.D6:
                        { c = 6; }
                        break;
                    case Keys.D7:
                        { c = 7; }
                        break;
                    case Keys.D8:
                        { c = 8; }
                        break;
                    case Keys.D9:
                        { c = 9; }
                        break;
                    case Keys.D0:
                        { c = 0; }
                        break;
                }


            }
            return c;
        }

        private void ClearControls()
        {
            ClientNameLabel.Text = "-";
            AssignmentNumberLabel.Text = "-";
            OrderNumberLabel.Text = "-";
            MainOrderNumberLabel.Text = "-";
            CreateDateLabel.Text = "-";
            AddToStorageDateLabel.Text = "-";
            RemoveFromStorageDateLabel.Text = "-";
            PackNumberLabel.Text = "-";
            TotalLabel.Text = "-";
            TotalRemoveLabel.Text = "-";
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

                ClearControls();

                CheckLabel.Clear();

                if (BarcodeTextBox.Text.Length < 12)
                {
                    BarcodeTextBox.Clear();
                    BarcodeLabel.Text = "Неверный штрихкод";
                    SetGridColor(false);

                    return;
                }

                string Prefix = BarcodeTextBox.Text.Substring(0, 3);

                if (Prefix != "020" && Prefix != "021")
                {
                    BarcodeTextBox.Clear();
                    BarcodeLabel.Text = "Неверный штрихкод";
                    SetGridColor(false);

                    return;
                }

                BarcodeLabel.Text = BarcodeTextBox.Text;
                BarcodeTextBox.Clear();

                //комплектация
                if (Prefix == "020")
                {
                    int CabFurnitureComplementID = Convert.ToInt32(BarcodeLabel.Text.Substring(3, 9));

                    if (CheckLabel.GetComplementInfo(CabFurnitureComplementID))
                    {
                        CheckLabel.GetComplementContent(CabFurnitureComplementID);

                        ClientNameLabel.Text = CheckLabel.lInfo.ClientName;
                        OrderNumberLabel.Text = CheckLabel.lInfo.OrderNumber;
                        MainOrderNumberLabel.Text = CheckLabel.lInfo.MainOrderNumber;
                        CreateDateLabel.Text = CheckLabel.lInfo.CreateDateTime;
                        PackNumberLabel.Text = CheckLabel.lInfo.PackNumber;

                        if (CheckLabel.AlreadyComplementPack)
                        {
                            CheckPicture.Image = Properties.Resources.cancel;
                            BarcodeLabel.ForeColor = Color.FromArgb(240, 0, 0);
                            BarcodeLabel.Text = "Списано со склада ранее";
                            SetGridColor(true);
                            ClearControls();
                            return;
                        }
                        CheckPicture.Visible = true;
                        if (CheckLabel.WaitScanComplement)
                        {
                            if (CheckLabel.ArePackagesEqual())
                            {
                                if (CheckLabel.RemoveFromStorage())
                                {
                                    CheckPicture.Image = Properties.Resources.OK;
                                    BarcodeLabel.ForeColor = Color.FromArgb(82, 169, 24);
                                    CheckLabel.PackComplement(CabFurnitureComplementID);
                                    BarcodeLabel.Text = "Списано со склада";
                                    SetGridColor(true);
                                }
                                else
                                {
                                    CheckPicture.Image = Properties.Resources.cancel;
                                    BarcodeLabel.ForeColor = Color.FromArgb(240, 0, 0);

                                    BarcodeLabel.Text = "Ошибка списания";
                                    SetGridColor(false);
                                    ClearControls();
                                }
                            }
                            else
                            {
                                CheckPicture.Image = Properties.Resources.cancel;
                                BarcodeLabel.ForeColor = Color.FromArgb(240, 0, 0);

                                BarcodeLabel.Text = "Упаковка не соответствует";
                                SetGridColor(false);
                                ClearControls();
                            }
                        }
                        else
                        {
                            CheckPicture.Image = Properties.Resources.cancel;
                            BarcodeLabel.ForeColor = Color.FromArgb(240, 0, 0);

                            BarcodeLabel.Text = "Отсканируйте упаковку";
                            SetGridColor(false);
                            ClearControls();
                        }
                    }
                    else
                    {
                        CheckPicture.Visible = true;
                        CheckPicture.Image = Properties.Resources.cancel;
                        BarcodeLabel.ForeColor = Color.FromArgb(240, 0, 0);

                        BarcodeLabel.Text = "Упаковки не существует";
                        SetGridColor(false);
                        ClearControls();
                    }
                }
                //упаковка
                if (Prefix == "021")
                {
                    if (CheckLabel.WaitScanComplement)
                    {
                        CheckPicture.Visible = true;
                        CheckPicture.Image = Properties.Resources.cancel;
                        BarcodeLabel.ForeColor = Color.FromArgb(240, 0, 0);

                        BarcodeLabel.Text = "Отсканируйте комплектацию";
                        SetGridColor(false);
                        ClearControls();
                    }
                    else
                    {
                        int CabFurniturePackageID = Convert.ToInt32(BarcodeLabel.Text.Substring(3, 9));

                        if (CheckLabel.GetPackageInfo(CabFurniturePackageID))
                        {
                            CheckLabel.GetPackageContent(CabFurniturePackageID);
                            CheckPicture.Visible = true;
                            CheckPicture.Image = Properties.Resources.OK;
                            BarcodeLabel.ForeColor = Color.FromArgb(82, 169, 24);

                            ClientNameLabel.Text = CheckLabel.lInfo.ClientName;
                            AssignmentNumberLabel.Text = CheckLabel.lInfo.AssignmentNumber;
                            AddToStorageDateLabel.Text = CheckLabel.lInfo.AddToStorageDateTime;
                            RemoveFromStorageDateLabel.Text = CheckLabel.lInfo.RemoveFromStorageDateTime;
                            CreateDateLabel.Text = CheckLabel.lInfo.CreateDateTime;
                            PackNumberLabel.Text = CheckLabel.lInfo.PackNumber;
                            TotalLabel.Text = CheckLabel.lInfo.PackedToTotal;
                            TotalRemoveLabel.Text = CheckLabel.lInfo.RemoveTotal;
                            TotalLabel.ForeColor = CheckLabel.lInfo.TotalLabelColor;
                            TotalRemoveLabel.ForeColor = CheckLabel.lInfo.RemoveTotalLabelColor;
                            if (!CheckLabel.lInfo.AddToStorage)
                                BarcodeLabel.Text = "Поставлено на склад";
                            if (CheckLabel.lInfo.AddToStorage && !CheckLabel.lInfo.RemoveFromStorage)
                                BarcodeLabel.Text = "Отсканируйте комплектацию";
                            if (CheckLabel.lInfo.RemoveFromStorage)
                                BarcodeLabel.Text = "Списано со склада ранее";
                            SetGridColor(true);
                        }
                        else
                        {
                            CheckPicture.Visible = true;
                            CheckPicture.Image = Properties.Resources.cancel;
                            BarcodeLabel.ForeColor = Color.FromArgb(240, 0, 0);

                            BarcodeLabel.Text = "Упаковки не существует";
                            SetGridColor(false);
                            ClearControls();
                        }
                    }
                }
            }

        }

        public void SetGridColor(bool IsAccept)
        {
            if (IsAccept)
            {
                dgvScan.StateCommon.Background.Color1 = Color.FromArgb(82, 169, 24);
                dgvScan.StateCommon.Background.Color2 = Color.Transparent;
                dgvScan.StateCommon.Background.ColorStyle = PaletteColorStyle.Solid;
                dgvScan.StateCommon.DataCell.Back.Color1 = Color.FromArgb(82, 169, 24);
                dgvScan.StateCommon.DataCell.Content.Color1 = System.Drawing.Color.White;
            }
            else
            {
                dgvScan.StateCommon.Background.Color1 = Color.Red;
                dgvScan.StateCommon.Background.Color2 = Color.Transparent;
                dgvScan.StateCommon.Background.ColorStyle = PaletteColorStyle.Solid;
                dgvScan.StateCommon.DataCell.Back.Color1 = Color.Red;
                dgvScan.StateCommon.DataCell.Content.Color1 = Color.White;
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
            //if (!CanAction)
            //    return;
            //if (GetActiveWindow() != this.Handle && tabControl1.SelectedIndex == 5)
            //{
            //    this.Activate();
            //}
        }

        private void kryptonContextMenuItem11_Click_1(object sender, EventArgs e)
        {
            bool OKCancel = Infinium.LightMessageBox.Show(ref TopForm, true,
                    "Вы собираетесь учистить Комплектацию. Продолжить?",
                    "Удаление упаковок");
            if (!OKCancel)
                return;

            int MainOrderID = 0;
            if (dgvMainOrders.SelectedRows.Count != 0 && dgvMainOrders.SelectedRows[0].Cells["MainOrderID"].Value != DBNull.Value)
                MainOrderID = Convert.ToInt32(dgvMainOrders.SelectedRows[0].Cells["MainOrderID"].Value);

            Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Сохранение данных.\r\nПодождите..."); });
            T.Start();
            while (!SplashWindow.bSmallCreated) ;
            NeedSplash = false;

            assignmentsManager.ClearCabFurnitureComplenents(MainOrderID);


            DateTime date1 = DateTimePicker1.Value.Date;
            DateTime date2 = DateTimePicker2.Value.Date;

            complementsManager.UpdateComplements(date1, date2);

            NeedSplash = true;
            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
        }

        private void dgvMainOrders_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Right && e.RowIndex != -1)
            {
                kryptonContextMenu5.Show(new Point(Cursor.Position.X - 212, Cursor.Position.Y - 10));
            }
        }

        private void kryptonContextMenuItem9_Click_1(object sender, EventArgs e)
        {
            bool OKCancel = Infinium.LightMessageBox.Show(ref TopForm, true,
                    "Вы собираетесь удалить задание. Продолжить?",
                    "Удаление задания");
            if (!OKCancel)
                return;

            int CabFurAssignmentID = 0;
            if (dgvAllAssignments.SelectedRows.Count != 0 && dgvAllAssignments.SelectedRows[0].Cells["CabFurAssignmentID"].Value != DBNull.Value)
                CabFurAssignmentID = Convert.ToInt32(dgvAllAssignments.SelectedRows[0].Cells["CabFurAssignmentID"].Value);
            //if (!assignmentsManager.IsAssignmentDetailsEmpty(CabFurAssignmentID))
            //{
            //    Infinium.LightMessageBox.Show(ref TopForm, false,
            //        "Задание не пустое. Сначала удалите содержимое, затем повторите",
            //        "Удаление задания");
            //    return;
            //}
            assignmentsManager.ClearCabFurniturePackages(CabFurAssignmentID);
            assignmentsManager.RemoveAssignment();
            SaveAllAssignments();
        }

        private void tbCount_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                AddItemButton_Click(null, null);
            }
        }

        private void dgvMainOrders_SelectionChanged(object sender, EventArgs e)
        {
            int MainOrderID = 0;
            if (dgvMainOrders.SelectedRows.Count != 0 && dgvMainOrders.SelectedRows[0].Cells["MainOrderID"].Value != DBNull.Value)
                MainOrderID = Convert.ToInt32(dgvMainOrders.SelectedRows[0].Cells["MainOrderID"].Value);

            if (NeedSplash)
            {
                Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Сохранение данных.\r\nПодождите..."); });
                T.Start();
                while (!SplashWindow.bSmallCreated) ;

                complementsManager.FilterComplementsLabels(MainOrderID);

                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;
            }
            else
                complementsManager.FilterComplementsLabels(MainOrderID);
        }

        private void dgvPackages_SelectionChanged(object sender, EventArgs e)
        {
            int CabFurAssignmentID = 0;
            if (dgvPackages.SelectedRows.Count != 0 && dgvPackages.SelectedRows[0].Cells["CabFurAssignmentID"].Value != DBNull.Value)
                CabFurAssignmentID = Convert.ToInt32(dgvPackages.SelectedRows[0].Cells["CabFurAssignmentID"].Value);

            if (NeedSplash)
            {
                Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Сохранение данных.\r\nПодождите..."); });
                T.Start();
                while (!SplashWindow.bSmallCreated) ;

                packagesManager.FilterPackagesLabels(CabFurAssignmentID);

                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;
            }
            else
            {
                packagesManager.FilterPackagesLabels(CabFurAssignmentID);
            }
        }

        private void kryptonButton1_Click(object sender, EventArgs e)
        {
            CheckLabel.CancelPackScan();
            SetGridColor(false);
            ClearControls();
        }

        private void kryptonCheckBox13_CheckedChanged(object sender, EventArgs e)
        {
            FilterPackages();
        }

        private void kryptonCheckBox12_CheckedChanged(object sender, EventArgs e)
        {
            FilterPackages();
        }

        private void kryptonCheckBox11_CheckedChanged(object sender, EventArgs e)
        {
            FilterPackages();
        }

        private void kryptonCheckBox10_CheckedChanged(object sender, EventArgs e)
        {
            FilterPackages();
        }

        private void kryptonCheckBox9_CheckedChanged(object sender, EventArgs e)
        {
            FilterPackages();
        }

        private void kryptonCheckBox8_CheckedChanged(object sender, EventArgs e)
        {
            FilterPackages();
        }

        private void kryptonCheckBox4_CheckedChanged(object sender, EventArgs e)
        {
            FilterComplements();
        }

        private void kryptonCheckBox3_CheckedChanged(object sender, EventArgs e)
        {
            FilterComplements();
        }

        private void kryptonCheckBox7_CheckedChanged(object sender, EventArgs e)
        {
            FilterComplements();
        }

        private void kryptonCheckBox6_CheckedChanged(object sender, EventArgs e)
        {
            FilterComplements();
        }

        private void kryptonCheckBox5_CheckedChanged(object sender, EventArgs e)
        {
            FilterComplements();
        }

        private void kryptonCheckBox2_CheckedChanged(object sender, EventArgs e)
        {
            FilterComplements();
        }

        private void kryptonContextMenuItem17_Click(object sender, EventArgs e)
        {
            if (dgvComplementLabels.SelectedRows.Count == 0)
                return;
            complementLabel.ClearLabelInfo();

            int ComplementsCount = 0;
            int MainOrderID = 0;

            if (dgvComplements.SelectedRows.Count != 0 && dgvComplements.SelectedRows[0].Cells["ComplementsCount"].Value != DBNull.Value)
                ComplementsCount = Convert.ToInt32(dgvComplements.SelectedRows[0].Cells["ComplementsCount"].Value);
            if (dgvMainOrders.SelectedRows.Count != 0 && dgvMainOrders.SelectedRows[0].Cells["MainOrderID"].Value != DBNull.Value)
                MainOrderID = Convert.ToInt32(dgvMainOrders.SelectedRows[0].Cells["MainOrderID"].Value);

            int[] CabFurnitureComplementID = new int[dgvComplementLabels.SelectedRows.Count];
            int[] Index = new int[dgvComplementLabels.SelectedRows.Count];
            for (int i = 0; i < dgvComplementLabels.SelectedRows.Count; i++)
            {
                CabFurnitureComplementID[i] = Convert.ToInt32(dgvComplementLabels.SelectedRows[i].Cells["CabFurnitureComplementID"].Value);
                Index[i] = Convert.ToInt32(dgvComplementLabels.SelectedRows[i].Cells["Index"].Value);
            }
            Array.Sort(CabFurnitureComplementID);
            Array.Sort(Index);

            if (detailsReport == null)
                detailsReport = new DetailsReport();
            List<ComplementLabelInfo> Labels = assignmentsManager.CreateComplementLabels(CabFurnitureComplementID, Index, ComplementsCount);
            if (Labels.Count == 0)
                return;

            if (NeedSplash)
            {
                Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                T.Start();
                //while (!SplashWindow.bSmallCreated) ;
                NeedSplash = false;

                detailsReport.CreateReport(Labels);

                NeedSplash = true;
                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;
            }
            else
            {
                detailsReport.CreateReport(Labels);
            }
        }

        private void btnUpdateComplements_Click(object sender, EventArgs e)
        {
            Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Обновление.\r\nПодождите..."); });
            T.Start();
            while (!SplashWindow.bSmallCreated) ;
            NeedSplash = false;

            DateTime date1 = DateTimePicker1.Value.Date;
            DateTime date2 = DateTimePicker2.Value.Date;

            complementsManager.UpdateComplements(date1, date2);
            packagesManager.UpdatePackages(date1, date2);

            //AssignmentsManager.ff();

            NeedSplash = true;
            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
        }

        private void btnAddWorkShop_Click(object sender, EventArgs e)
        {
            if (cabFurStorage == null)
                return;
            PhantomForm PhantomForm = new PhantomForm();
            PhantomForm.Show();

            NewCabFurWorkShopForm newCabFurWorkShopForm = new NewCabFurWorkShopForm();

            TopForm = newCabFurWorkShopForm;

            newCabFurWorkShopForm.ShowDialog();

            PhantomForm.Close();
            PhantomForm.Dispose();

            TopForm = null;

            bool bOk = newCabFurWorkShopForm.bOk;
            string sName = newCabFurWorkShopForm.sName;

            newCabFurWorkShopForm.Dispose();

            Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Обновление.\r\nПодождите..."); });
            T.Start();
            while (!SplashWindow.bSmallCreated) ;
            NeedSplash = false;

            if (bOk)
            {
                cabFurStorage.AddWorkShop(sName);
                cabFurStorage.SaveWorkShops();
                cabFurStorage.UpdateWorkShops();
                cabFurStorage.SetWorkShopPosition(cabFurStorage.MaxWorkShopId);
            }

            NeedSplash = true;
            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
        }

        private void btnEditWorkShop_Click(object sender, EventArgs e)
        {
            if (cabFurStorage == null || !cabFurStorage.HasWorkShops || cabFurStorage.CurrentWorkShopId == -1)
                return;

            string oldName = cabFurStorage.CurrentWorkShopName;
            int workShopId = cabFurStorage.CurrentWorkShopId;

            PhantomForm PhantomForm = new PhantomForm();
            PhantomForm.Show();

            NewCabFurWorkShopForm newCabFurWorkShopForm = new NewCabFurWorkShopForm(oldName);

            TopForm = newCabFurWorkShopForm;

            newCabFurWorkShopForm.ShowDialog();

            PhantomForm.Close();
            PhantomForm.Dispose();

            TopForm = null;

            bool bOk = newCabFurWorkShopForm.DialogResult == DialogResult.OK ? true : false;
            string sName = newCabFurWorkShopForm.sName;

            newCabFurWorkShopForm.Dispose();

            Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Обновление.\r\nПодождите..."); });
            T.Start();
            while (!SplashWindow.bSmallCreated) ;
            NeedSplash = false;

            if (bOk)
            {
                cabFurStorage.EditWorkShop(cabFurStorage.CurrentWorkShopId, sName);
                cabFurStorage.SaveWorkShops();
                cabFurStorage.UpdateWorkShops();
                cabFurStorage.SetWorkShopPosition(workShopId);
            }

            NeedSplash = true;
            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
        }

        private void btnAddRack_Click(object sender, EventArgs e)
        {
            if (cabFurStorage == null || !cabFurStorage.HasWorkShops || cabFurStorage.CurrentWorkShopId == -1)
                return;
            PhantomForm PhantomForm = new PhantomForm();
            PhantomForm.Show();

            NewCabFurRackForm newCabFurRackForm = new NewCabFurRackForm();

            TopForm = newCabFurRackForm;

            newCabFurRackForm.ShowDialog();

            PhantomForm.Close();
            PhantomForm.Dispose();

            TopForm = null;

            bool bOk = newCabFurRackForm.bOk;
            string sName = newCabFurRackForm.newName;

            newCabFurRackForm.Dispose();

            Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Обновление.\r\nПодождите..."); });
            T.Start();
            while (!SplashWindow.bSmallCreated) ;
            NeedSplash = false;

            if (bOk)
            {
                cabFurStorage.AddRack(sName, cabFurStorage.CurrentWorkShopId);
                cabFurStorage.SaveRacks();
                cabFurStorage.UpdateRacks();
                cabFurStorage.SetRackPosition(cabFurStorage.MaxRackId);
            }

            NeedSplash = true;
            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
        }

        private void btnEditRack_Click(object sender, EventArgs e)
        {
            if (cabFurStorage == null || !cabFurStorage.HasRacks || cabFurStorage.CurrentRackId == -1)
                return;

            string oldName = cabFurStorage.CurrentRackName;
            int rackId = cabFurStorage.CurrentRackId;

            PhantomForm PhantomForm = new PhantomForm();
            PhantomForm.Show();

            NewCabFurRackForm newCabFurRackForm = new NewCabFurRackForm(oldName);

            TopForm = newCabFurRackForm;

            newCabFurRackForm.ShowDialog();

            PhantomForm.Close();
            PhantomForm.Dispose();

            TopForm = null;

            bool bOk = newCabFurRackForm.DialogResult == DialogResult.OK ? true : false;
            string sName = newCabFurRackForm.newName;

            newCabFurRackForm.Dispose();

            Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Обновление.\r\nПодождите..."); });
            T.Start();
            while (!SplashWindow.bSmallCreated) ;
            NeedSplash = false;

            if (bOk)
            {
                cabFurStorage.EditRack(cabFurStorage.CurrentRackId, sName);
                cabFurStorage.SaveRacks();
                cabFurStorage.UpdateRacks();
                cabFurStorage.SetRackPosition(rackId);
            }

            NeedSplash = true;
            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
        }

        private void btnAddCell_Click(object sender, EventArgs e)
        {
            if (cabFurStorage == null || !cabFurStorage.HasRacks || cabFurStorage.CurrentRackId == -1 || !cabFurStorage.HasWorkShops || cabFurStorage.CurrentWorkShopId == -1)
                return;
            PhantomForm PhantomForm = new PhantomForm();
            PhantomForm.Show();

            NewCabFurCellForm newCabFurCellForm = new NewCabFurCellForm();

            TopForm = newCabFurCellForm;

            newCabFurCellForm.ShowDialog();

            PhantomForm.Close();
            PhantomForm.Dispose();

            TopForm = null;

            bool bOk = newCabFurCellForm.bOk;
            string sName = newCabFurCellForm.sName;

            newCabFurCellForm.Dispose();

            Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Обновление.\r\nПодождите..."); });
            T.Start();
            while (!SplashWindow.bSmallCreated) ;
            NeedSplash = false;

            if (bOk)
            {
                cabFurStorage.AddCell(sName, cabFurStorage.CurrentRackId);
                cabFurStorage.SaveCells();
                cabFurStorage.UpdateCells();
                //cabFurStorage.SetRackPosition(cabFurStorage.MaxRackId);
            }

            NeedSplash = true;
            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
        }

        private void btnEditCell_Click(object sender, EventArgs e)
        {
            if (cabFurStorage == null || !cabFurStorage.HasCells || cabFurStorage.CurrentCellId == -1)
                return;

            string oldName = cabFurStorage.CurrentCellName;
            int rackId = cabFurStorage.CurrentCellId;

            PhantomForm PhantomForm = new PhantomForm();
            PhantomForm.Show();

            NewCabFurCellForm newCabFurCellForm = new NewCabFurCellForm(oldName);

            TopForm = newCabFurCellForm;

            newCabFurCellForm.ShowDialog();

            PhantomForm.Close();
            PhantomForm.Dispose();

            TopForm = null;

            bool bOk = newCabFurCellForm.DialogResult == DialogResult.OK ? true : false;
            string sName = newCabFurCellForm.sName;

            newCabFurCellForm.Dispose();

            Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Обновление.\r\nПодождите..."); });
            T.Start();
            while (!SplashWindow.bSmallCreated) ;
            NeedSplash = false;

            if (bOk)
            {
                cabFurStorage.EditCell(cabFurStorage.CurrentCellId, sName);
                cabFurStorage.SaveCells();
                cabFurStorage.UpdateCells();
                //cabFurStorage.SetRackPosition(rackId);
            }

            NeedSplash = true;
            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
        }

        private void cmbxWorkShops_SelectedValueChanged(object sender, EventArgs e)
        {
            if (cabFurStorage == null)
                return;
            int workShopID = -1;
            if (cmbxWorkShops.SelectedItem != null && ((DataRowView)cmbxWorkShops.SelectedItem).Row["WorkShopID"] != DBNull.Value)
                workShopID = Convert.ToInt32(((DataRowView)cmbxWorkShops.SelectedItem).Row["WorkShopID"]);
            //NeedSplash = true;
            cabFurStorage.FilterRacksByWorkShop(workShopID);
        }

        private void cmbxRacks_SelectedValueChanged(object sender, EventArgs e)
        {
            if (cabFurStorage == null)
                return;
            int rackId = -1;
            if (cmbxRacks.SelectedItem != null && ((DataRowView)cmbxRacks.SelectedItem).Row["RackID"] != DBNull.Value)
                rackId = Convert.ToInt32(((DataRowView)cmbxRacks.SelectedItem).Row["RackID"]);
            //NeedSplash = true;
            cabFurStorage.FilterCellsByRack(rackId);
        }

        private void dgvCells_SelectionChanged(object sender, EventArgs e)
        {
            if (storagePackagesManager == null)
                return;
            int cellId = 0;
            if (dgvCells.SelectedRows.Count != 0 && dgvCells.SelectedRows[0].Cells["CellID"].Value != DBNull.Value)
                cellId = Convert.ToInt32(dgvCells.SelectedRows[0].Cells["CellID"].Value);

            storagePackagesManager.GetPackagesLabels(cellId);

            //if (NeedSplash)
            //{
            //    Thread T = new Thread(delegate()
            //    {
            //        SplashWindow.CreateSmallSplash(ref TopForm, "Обновление.\r\nПодождите...");
            //    });
            //    T.Start();
            //    while (!SplashWindow.bSmallCreated) ;
            //    NeedSplash = false;

            //    storagePackagesManager.GetPackagesLabels(cellId);

            //    NeedSplash = true;
            //    while (SplashWindow.bSmallCreated)
            //        SmallWaitForm.CloseS = true;
            //}
            //else
            //    storagePackagesManager.GetPackagesLabels(cellId);
        }

        private void dgvStoragePackagesLabels_SelectionChanged(object sender, EventArgs e)
        {
            if (storagePackagesManager == null)
                return;
            int cabFurniturePackageID = 0;
            if (dgvStoragePackagesLabels.SelectedRows.Count != 0 && dgvStoragePackagesLabels.SelectedRows[0].Cells["CabFurniturePackageID"].Value != DBNull.Value)
                cabFurniturePackageID = Convert.ToInt32(dgvStoragePackagesLabels.SelectedRows[0].Cells["CabFurniturePackageID"].Value);
            storagePackagesManager.FilterPackagesDetails(cabFurniturePackageID);
        }

        private void btnBindPackages_Click(object sender, EventArgs e)
        {
            if (dgvCells.SelectedRows.Count == 0)
                return;
            kryptonContextMenu6.Show(new Point(Cursor.Position.X - 212, Cursor.Position.Y - 10));
        }

        private void kryptonContextMenuItem22_Click(object sender, EventArgs e)
        {
            if (dgvCells.SelectedRows.Count == 0)
                return;
            PhantomForm PhantomForm = new PhantomForm();
            PhantomForm.Show();
            int cellId = -1;

            InputCellBarCodeForm inputCellBarCodeForm = new InputCellBarCodeForm(this, storagePackagesManager);

            TopForm = inputCellBarCodeForm;
            inputCellBarCodeForm.ShowDialog();
            TopForm = null;

            if (inputCellBarCodeForm.DialogResult == DialogResult.OK)
            {
                cellId = inputCellBarCodeForm.cellId;
                inputCellBarCodeForm.Dispose();

                BindPackagesToCellForm bindPackagesToCellForm = new BindPackagesToCellForm(this, storagePackagesManager, cellId);

                TopForm = bindPackagesToCellForm;
                bindPackagesToCellForm.ShowDialog();

                bindPackagesToCellForm.Dispose();
                TopForm = null;

                Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Обновление.\r\nПодождите..."); });
                T.Start();
                while (!SplashWindow.bSmallCreated) ;
                NeedSplash = false;

                cabFurStorage.UpdateCells();

                NeedSplash = true;
                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;
            }

            PhantomForm.Close();
            PhantomForm.Dispose();

            GC.Collect();
        }

        private void kryptonContextMenuItem18_Click(object sender, EventArgs e)
        {
            if (dgvCells.SelectedRows.Count == 0)
                return;
            int cellId = Convert.ToInt32(dgvCells.SelectedRows[0].Cells["CellID"].Value);

            PhantomForm PhantomForm = new PhantomForm();
            PhantomForm.Show();

            BindPackagesToCellForm bindPackagesToCellForm = new BindPackagesToCellForm(this, storagePackagesManager, cellId);

            TopForm = bindPackagesToCellForm;
            bindPackagesToCellForm.ShowDialog();

            PhantomForm.Close();

            PhantomForm.Dispose();
            bindPackagesToCellForm.Dispose();
            TopForm = null;
            GC.Collect();

            Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Обновление.\r\nПодождите..."); });
            T.Start();
            while (!SplashWindow.bSmallCreated) ;
            NeedSplash = false;

            cabFurStorage.UpdateCells();

            NeedSplash = true;
            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
        }

        private void kryptonContextMenuItem19_Click(object sender, EventArgs e)
        {
            if (dgvCells.SelectedRows.Count == 0)
                return;
            cellLabel.ClearLabelInfo();

            int[] cellId = new int[dgvCells.SelectedRows.Count];
            for (int i = 0; i < dgvCells.SelectedRows.Count; i++)
            {
                cellId[i] = Convert.ToInt32(dgvCells.SelectedRows[i].Cells["CellID"].Value);
            }

            List<CellLabelInfo> Labels = cabFurStorage.CreateCellLabels(cellId);
            if (Labels.Count == 0)
                return;
            for (int i = 0; i < Labels.Count; i++)
            {
                CellLabelInfo LabelInfo = Labels[i];
                cellLabel.AddLabelInfo(ref LabelInfo);
            }
            PrintDialog.Document = cellLabel.PD;

            if (PrintDialog.ShowDialog() == DialogResult.OK)
            {
                cellLabel.Print();
            }
        }

        private void btnPrintCellLabel_Click(object sender, EventArgs e)
        {
            if (dgvCells.SelectedRows.Count == 0)
                return;
            cellLabel.ClearLabelInfo();

            int[] cellId = new int[dgvCells.SelectedRows.Count];
            for (int i = 0; i < dgvCells.SelectedRows.Count; i++)
            {
                cellId[i] = Convert.ToInt32(dgvCells.SelectedRows[i].Cells["CellID"].Value);
            }

            List<CellLabelInfo> Labels = cabFurStorage.CreateCellLabels(cellId);
            if (Labels.Count == 0)
                return;
            for (int i = 0; i < Labels.Count; i++)
            {
                CellLabelInfo LabelInfo = Labels[i];
                cellLabel.AddLabelInfo(ref LabelInfo);
            }
            PrintDialog.Document = cellLabel.PD;

            if (PrintDialog.ShowDialog() == DialogResult.OK)
            {
                cellLabel.Print();
            }
        }

        private void kryptonContextMenuItem20_Click(object sender, EventArgs e)
        {
            if (dgvPackagesLabels.SelectedRows.Count == 0)
                return;
            int CabFurAssignmentID = 0;
            if (dgvPackages.SelectedRows.Count != 0 && dgvPackages.SelectedRows[0].Cells["CabFurAssignmentID"].Value != DBNull.Value)
                CabFurAssignmentID = Convert.ToInt32(dgvPackages.SelectedRows[0].Cells["CabFurAssignmentID"].Value);

            int[] CabFurniturePackageIDs = new int[dgvPackagesLabels.SelectedRows.Count];
            for (int i = 0; i < dgvPackagesLabels.SelectedRows.Count; i++)
            {
                CabFurniturePackageIDs[i] = Convert.ToInt32(dgvPackagesLabels.SelectedRows[i].Cells["CabFurniturePackageID"].Value);
            }

            bool OKCancel = Infinium.LightMessageBox.Show(ref TopForm, true,
                    "Вы собираетесь отправить упаковки на ОТК. Они будут отвязаны от ячейки Продолжить?",
                    "ОТК");
            if (!OKCancel)
                return;

            if (NeedSplash)
            {
                Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                T.Start();
                while (!SplashWindow.bSmallCreated) ;
                NeedSplash = false;

                storagePackagesManager.QualityControlIn(CabFurniturePackageIDs);
                packagesManager.FilterPackagesLabels(CabFurAssignmentID);

                NeedSplash = true;
                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;
            }
            else
            {
                packagesManager.FilterPackagesLabels(CabFurAssignmentID);
            }
        }

        private void kryptonContextMenuItem21_Click(object sender, EventArgs e)
        {
            if (dgvPackagesLabels.SelectedRows.Count == 0)
                return;
            int CabFurAssignmentID = 0;
            if (dgvPackages.SelectedRows.Count != 0 && dgvPackages.SelectedRows[0].Cells["CabFurAssignmentID"].Value != DBNull.Value)
                CabFurAssignmentID = Convert.ToInt32(dgvPackages.SelectedRows[0].Cells["CabFurAssignmentID"].Value);

            int[] CabFurniturePackageID = new int[dgvPackagesLabels.SelectedRows.Count];
            for (int i = 0; i < dgvPackagesLabels.SelectedRows.Count; i++)
            {
                CabFurniturePackageID[i] = Convert.ToInt32(dgvPackagesLabels.SelectedRows[i].Cells["CabFurniturePackageID"].Value);
            }

            if (NeedSplash)
            {
                Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                T.Start();
                while (!SplashWindow.bSmallCreated) ;
                NeedSplash = false;

                storagePackagesManager.QualityControlOut(CabFurniturePackageID);
                packagesManager.FilterPackagesLabels(CabFurAssignmentID);

                NeedSplash = true;
                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;
            }
            else
            {
                packagesManager.FilterPackagesLabels(CabFurAssignmentID);
            }
        }

        private void btnStartInventory_Click(object sender, EventArgs e)
        {
            if (cabFurStorage == null || !cabFurStorage.HasWorkShops || cabFurStorage.CurrentWorkShopId == -1)
                return;
            PhantomForm PhantomForm = new PhantomForm();
            PhantomForm.Show();
            int workshopId = cabFurStorage.CurrentWorkShopId;

            CabFurInventoryForm сabFurInventoryForm = new CabFurInventoryForm(this, assignmentsManager, storagePackagesManager, workshopId);

            TopForm = сabFurInventoryForm;
            сabFurInventoryForm.ShowDialog();

            сabFurInventoryForm.Dispose();
            TopForm = null;

            Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Обновление.\r\nПодождите..."); });
            T.Start();
            while (!SplashWindow.bSmallCreated) ;
            NeedSplash = false;

            cabFurStorage.UpdateCells();

            NeedSplash = true;
            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;

            PhantomForm.Close();
            PhantomForm.Dispose();

            GC.Collect();
        }

        private void btnExportStorageToExcel_Click(object sender, EventArgs e)
        {
            Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
            T.Start();
            while (!SplashWindow.bSmallCreated) ;
            NeedSplash = false;

            if (assignmentsManager.GetStorage())
            {
                for (int i = 0; i < assignmentsManager.TotalProductsCoversDs.Tables.Count; i++)
                {
                    cabFurStorageToExcel.Form01(assignmentsManager.TotalProductsCoversDs.Tables[i]);
                }
                cabFurStorageToExcel.SaveFile("Склад корп мебели", true);
            }
            else
                InfiniumTips.ShowTip(this, 50, 85, "Данные не найдены", 2000);

            NeedSplash = true;
            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
        }

        private void dgvComplements_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Right && e.RowIndex != -1)
            {
                kryptonContextMenu8.Show(new Point(Cursor.Position.X - 212, Cursor.Position.Y - 10));
            }
        }

        private void kryptonContextMenuItem23_Click(object sender, EventArgs e)
        {
            if (dgvComplements.SelectedRows.Count == 0)
                return;

            int ClientID = 0;
            int OrderNumber = 0;
            int MegaOrderID = 0;
            if (dgvComplements.SelectedRows.Count != 0 && dgvComplements.SelectedRows[0].Cells["ClientID"].Value != DBNull.Value)
                ClientID = Convert.ToInt32(dgvComplements.SelectedRows[0].Cells["ClientID"].Value);
            if (dgvComplements.SelectedRows.Count != 0 && dgvComplements.SelectedRows[0].Cells["OrderNumber"].Value != DBNull.Value)
                OrderNumber = Convert.ToInt32(dgvComplements.SelectedRows[0].Cells["OrderNumber"].Value);
            if (dgvComplements.SelectedRows.Count != 0 && dgvComplements.SelectedRows[0].Cells["MegaOrderID"].Value != DBNull.Value)
                MegaOrderID = Convert.ToInt32(dgvComplements.SelectedRows[0].Cells["MegaOrderID"].Value);

            Thread T = new Thread(delegate () { SplashWindow.CreateSplash(); });
            T.Start();

            while (!SplashForm.bCreated) ;

            CabFurAssembleForm cabFurDispatchForm = new CabFurAssembleForm(this, assignmentsManager, ClientID, MegaOrderID);

            TopForm = cabFurDispatchForm;

            cabFurDispatchForm.ShowDialog();

            cabFurDispatchForm.Close();
            cabFurDispatchForm.Dispose();

            TopForm = null;
        }

        private void cbTSGroupsSearch_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (assignmentsManager == null)
                return;
            int TechStoreGroupID = 0;
            if (cbTSGroupsSearch.SelectedItem != null && ((DataRowView)cbTSGroupsSearch.SelectedItem).Row["TechStoreGroupID"] != DBNull.Value)
                TechStoreGroupID = Convert.ToInt32(((DataRowView)cbTSGroupsSearch.SelectedItem).Row["TechStoreGroupID"]);
            assignmentsManager.FilterTechStoreSubGroups(TechStoreGroupID);
            cbTSSubGroupsSearch_SelectedIndexChanged(null, null);
        }

        private void cbTSSubGroupsSearch_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (assignmentsManager == null)
                return;
            int TechStoreSubGroupID = 0;
            if (cbTSSubGroupsSearch.SelectedItem != null && ((DataRowView)cbTSSubGroupsSearch.SelectedItem).Row["TechStoreSubGroupID"] != DBNull.Value)
                TechStoreSubGroupID = Convert.ToInt32(((DataRowView)cbTSSubGroupsSearch.SelectedItem).Row["TechStoreSubGroupID"]);
            assignmentsManager.FilterTechStore(TechStoreSubGroupID);
        }

        private void btnSearchPackages_Click(object sender, EventArgs e)
        {
            int TechStoreID = 0;
            int CoverID = 0;
            int PatinaID = 0;
            int InsetColorID = 0;

            if (cbTStoreSearch.SelectedItem != null && ((DataRowView)cbTStoreSearch.SelectedItem).Row["TechStoreID"] != DBNull.Value)
                TechStoreID = Convert.ToInt32(((DataRowView)cbTStoreSearch.SelectedItem).Row["TechStoreID"]);
            if (cbCoversSearch.SelectedItem != null && ((DataRowView)cbCoversSearch.SelectedItem).Row["CoverID"] != DBNull.Value)
                CoverID = Convert.ToInt32(((DataRowView)cbCoversSearch.SelectedItem).Row["CoverID"]);
            if (cbPatinaSearch.SelectedItem != null && ((DataRowView)cbPatinaSearch.SelectedItem).Row["PatinaID"] != DBNull.Value)
                PatinaID = Convert.ToInt32(((DataRowView)cbPatinaSearch.SelectedItem).Row["PatinaID"]);
            if (cbInsetColorsSearch.SelectedItem != null && ((DataRowView)cbInsetColorsSearch.SelectedItem).Row["InsetColorID"] != DBNull.Value)
                InsetColorID = Convert.ToInt32(((DataRowView)cbInsetColorsSearch.SelectedItem).Row["InsetColorID"]);

            Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
            T.Start();
            while (!SplashWindow.bSmallCreated) ;
            NeedSplash = false;

            cabFurStorage.SearchCells(TechStoreID, CoverID, PatinaID, InsetColorID);
            cabFurStorage.SearchRacks(TechStoreID, CoverID, PatinaID, InsetColorID);
            cabFurStorage.SearchWorkShops(TechStoreID, CoverID, PatinaID, InsetColorID);

            NeedSplash = true;
            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
        }

        private void btnUpdateStorePackages_Click(object sender, EventArgs e)
        {
            Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Обновление.\r\nПодождите..."); });
            T.Start();
            while (!SplashWindow.bSmallCreated) ;
            NeedSplash = false;

            cabFurStorage.UpdateCells();
            cabFurStorage.UpdateRacks();
            cabFurStorage.UpdateWorkShops();

            NeedSplash = true;
            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
        }

        private void btnRemoveRack_Click(object sender, EventArgs e)
        {
            bool OKCancel = Infinium.LightMessageBox.Show(ref TopForm, true,
                "Вы собираетесь удалить стеллаж. Продолжить?",
                "Удаление");
            if (!OKCancel)
                return;

            Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Удаление.\r\nПодождите..."); });
            T.Start();
            while (!SplashWindow.bSmallCreated) ;
            NeedSplash = false;

            int rackId = cabFurStorage.CurrentRackId;

            cabFurStorage.RemoveRack(cabFurStorage.CurrentRackId);
            cabFurStorage.SaveRacks();
            cabFurStorage.SaveCells();
            cabFurStorage.UpdateRacks();
            cabFurStorage.SetRackPosition(rackId);

            NeedSplash = true;
            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
        }

        private void btnRemoveWorkShop_Click(object sender, EventArgs e)
        {
            bool OKCancel = Infinium.LightMessageBox.Show(ref TopForm, true,
                "Вы собираетесь удалить цех. Продолжить?",
                "Удаление");
            if (!OKCancel)
                return;

            Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Удаление.\r\nПодождите..."); });
            T.Start();
            while (!SplashWindow.bSmallCreated) ;
            NeedSplash = false;

            int workShopId = cabFurStorage.CurrentWorkShopId;

            cabFurStorage.RemoveWorkShop(cabFurStorage.CurrentWorkShopId);
            cabFurStorage.SaveWorkShops();
            cabFurStorage.SaveRacks();
            cabFurStorage.SaveCells();
            cabFurStorage.UpdateWorkShops();
            cabFurStorage.SetWorkShopPosition(workShopId);

            NeedSplash = true;
            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
        }

        private void btnRemoveCell_Click(object sender, EventArgs e)
        {
            bool OKCancel = Infinium.LightMessageBox.Show(ref TopForm, true,
                "Вы собираетесь удалить ячейку. Продолжить?",
                "Удаление");
            if (!OKCancel)
                return;
            if (cabFurStorage == null || !cabFurStorage.HasCells || cabFurStorage.CurrentCellId == -1)
                return;

            Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Удаление.\r\nПодождите..."); });
            T.Start();
            while (!SplashWindow.bSmallCreated) ;
            NeedSplash = false;

            cabFurStorage.RemoveCell(cabFurStorage.CurrentCellId);
            cabFurStorage.SaveCells();
            cabFurStorage.UpdateCells();

            NeedSplash = true;
            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
        }

        private void kryptonContextMenuItem24_Click(object sender, EventArgs e)
        {

        }

        private void ChangeDateButton_Click(object sender, EventArgs e)
        {
               DateTime DispatchDate = kryptonMonthCalendar2.SelectionStart;
            int CabFurAssignmentID = 0;
            if (dgvAllAssignments.SelectedRows.Count != 0 && dgvAllAssignments.SelectedRows[0].Cells["CabFurAssignmentID"].Value != DBNull.Value)
                CabFurAssignmentID = Convert.ToInt32(dgvAllAssignments.SelectedRows[0].Cells["CabFurAssignmentID"].Value);

            var T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Сохранение данных.\r\nПодождите..."); });
            T.Start();
            while (!SplashWindow.bSmallCreated) ;
            NeedSplash = false;

            assignmentsManager.SetDispatchDate(CabFurAssignmentID, DispatchDate);
            assignmentsManager.SaveAllAssignments();
            FilterAssignments();
            assignmentsManager.MoveToAssignmentID(CabFurAssignmentID);

            NeedSplash = true;
            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
        }

        private void kryptonButton2_Click(object sender, EventArgs e)
        {
            FilterAssignments();
        }

        private void btnQualityControlOut_Click(object sender, EventArgs e)
        {
            PhantomForm PhantomForm = new PhantomForm();
            PhantomForm.Show();

            QualityControlForm qualityControlForm = new QualityControlForm(this, storagePackagesManager, true);

            TopForm = qualityControlForm;
            qualityControlForm.ShowDialog();

            PhantomForm.Close();

            PhantomForm.Dispose();
            qualityControlForm.Dispose();
            TopForm = null;
            GC.Collect();

            Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Обновление.\r\nПодождите..."); });
            T.Start();
            while (!SplashWindow.bSmallCreated) ;
            NeedSplash = false;

            cabFurStorage.UpdateCells();

            NeedSplash = true;
            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;


        }

        private void btnQualityControlIn_Click(object sender, EventArgs e)
        {
            PhantomForm PhantomForm = new PhantomForm();
            PhantomForm.Show();

            QualityControlForm qualityControlForm = new QualityControlForm(this, storagePackagesManager, false);

            TopForm = qualityControlForm;
            qualityControlForm.ShowDialog();

            PhantomForm.Close();

            PhantomForm.Dispose();
            qualityControlForm.Dispose();
            TopForm = null;
            GC.Collect();

            Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Обновление.\r\nПодождите..."); });
            T.Start();
            while (!SplashWindow.bSmallCreated) ;
            NeedSplash = false;

            cabFurStorage.UpdateCells();

            NeedSplash = true;
            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;


            //if (dgvStoragePackagesLabels.SelectedRows.Count == 0)
            //    return;
            //int cellId = 0;
            //if (dgvCells.SelectedRows.Count != 0 && dgvCells.SelectedRows[0].Cells["CellID"].Value != DBNull.Value)
            //    cellId = Convert.ToInt32(dgvCells.SelectedRows[0].Cells["CellID"].Value);

            //int[] CabFurniturePackageIDs = new int[dgvStoragePackagesLabels.SelectedRows.Count];
            //for (int i = 0; i < dgvStoragePackagesLabels.SelectedRows.Count; i++)
            //{
            //    CabFurniturePackageIDs[i] = Convert.ToInt32(dgvStoragePackagesLabels.SelectedRows[i].Cells["CabFurniturePackageID"].Value);
            //}

            //bool OKCancel = Infinium.LightMessageBox.Show(ref TopForm, true,
            //        "Вы собираетесь отправить упаковки на ОТК. Они будут отвязаны от ячейки. Продолжить?",
            //        "ОТК");
            //if (!OKCancel)
            //    return;

            //if (NeedSplash)
            //{
            //    Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
            //    T.Start();
            //    while (!SplashWindow.bSmallCreated) ;
            //    NeedSplash = false;

            //    packagesManager.QualityControlIn(CabFurniturePackageIDs);
            //    storagePackagesManager.GetPackagesLabels(cellId);

            //    NeedSplash = true;
            //    while (SplashWindow.bSmallCreated)
            //        SmallWaitForm.CloseS = true;
            //}
            //else
            //{
            //    storagePackagesManager.GetPackagesLabels(cellId);
            //}
        }

        private void btnShowQualityControl_Click(object sender, EventArgs e)
        {
            PhantomForm PhantomForm = new PhantomForm();
            PhantomForm.Show();

            AllQualityControlForm qualityControlForm = new AllQualityControlForm(this, storagePackagesManager);

            TopForm = qualityControlForm;
            qualityControlForm.ShowDialog();

            PhantomForm.Close();

            PhantomForm.Dispose();
            qualityControlForm.Dispose();
            TopForm = null;
            GC.Collect();

            Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Обновление.\r\nПодождите..."); });
            T.Start();
            while (!SplashWindow.bSmallCreated) ;
            NeedSplash = false;

            cabFurStorage.UpdateCells();

            NeedSplash = true;
            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;

        }

        private void kryptonContextMenuItem5_Click_1(object sender, EventArgs e)
        {
            int[] cabFurniturePackageID = new int[dgvStoragePackagesLabels.SelectedRows.Count];
            for (int i = 0; i < dgvStoragePackagesLabels.SelectedRows.Count; i++)
                cabFurniturePackageID[i] = Convert.ToInt32(dgvStoragePackagesLabels.SelectedRows[i].Cells["cabFurniturePackageID"].Value);

            cabFurStorage.UnbindPackages(cabFurniturePackageID);
            cabFurStorage.UpdateCells();
        }

        private void dgvCells_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {

        }

        private void dgvStoragePackagesLabels_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Right && e.RowIndex != -1)
            {
                kryptonContextMenu9.Show(new Point(Cursor.Position.X - 212, Cursor.Position.Y - 10));
            }
        }


        private void FilterStatisticsData()
        {
            Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных.\r\nПодождите..."); });
            T.Start();
            while (!SplashWindow.bSmallCreated) ;
            NeedSplash = false;


            isDateOn = true;
            bool bDateCreative = kryptonRadioButton1.Checked;
            bool bDateDispatch = kryptonRadioButton2.Checked;
            bool bDateAgreement = kryptonRadioButton3.Checked;
            bool bDateProduction = kryptonRadioButton4.Checked;
            bool bDateOutProduction = kryptonRadioButton5.Checked;
            bool bDatePrintout = kryptonRadioButton6.Checked;
            bool bDatePackaging = kryptonRadioButton7.Checked;
            bool bRemoveFromStorageDateTime = kryptonRadioButton8.Checked;
            bool bQualityControlInDateTime = kryptonRadioButton9.Checked;
            bool bQualityControlOutDateTime = kryptonRadioButton10.Checked;

            DateTime DateStart = kryptonDateTimePicker10.Value;
            DateTime DataEnd = kryptonDateTimePicker9.Value;



            int typeDate = 0;
            typeDate = bDateCreative ? 1 : typeDate;
            typeDate = bDateDispatch ? 2 : typeDate;
            typeDate = bDateAgreement ? 3 : typeDate;
            typeDate = bDateProduction ? 4 : typeDate;
            typeDate = bDateOutProduction ? 5 : typeDate;
            typeDate = bDatePrintout ? 6 : typeDate;
            typeDate = bDatePackaging ? 7 : typeDate;
            typeDate = bRemoveFromStorageDateTime ? 8 : typeDate;
            typeDate = bQualityControlInDateTime ? 9 : typeDate;
            typeDate = bQualityControlOutDateTime ? 10 : typeDate;

            assignmentsManager.UpdateDateStatistics(typeDate, DateStart, DataEnd);
            FilterCheckStatistics(false);

            NeedSplash = true;
            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;

        }

        private void FilterCheckStatistics(bool IsShowMessageBox = true)
        {
            if (IsShowMessageBox)
            {
                Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных.\r\nПодождите..."); });
                T.Start();
                while (!SplashWindow.bSmallCreated) ;
                NeedSplash = false;
            }

            string filter = "";
            filter += kryptonCheckBox13t.Checked ? "" : " and PrintDateTime is NULL";
            filter += kryptonCheckBox12t.Checked ? "" : " and PrintDateTime is not NULL";
            filter += kryptonCheckBox10t.Checked ? "" : " and AddToStorageDateTime is NULL";
            filter += kryptonCheckBox9t.Checked ? "" : " and  AddToStorageDateTime is not NULL";

            filter += kryptonCheckBox4t.Checked ? "" : " and ACreationDateTime is NULL ";
            filter += kryptonCheckBox3t.Checked ? "" : " and AProductionDateTime is NULL ";
            filter += kryptonCheckBox2t.Checked ? "" : " and ValOutProductionDateTime is NULL";

            filter += kryptonCheckBox16.Checked ? "" : " and ACreationDateTime is not NULL ";
            filter += kryptonCheckBox17.Checked ? "" : " and AProductionDateTime is not NULL ";
            filter += kryptonCheckBox18.Checked ? "" : " and ValOutProductionDateTime is not NULL";

            filter += " and CabFurAssignmentID is NOT NULL";

            if (filter.Length != 0)
            {
                filter = filter.Substring(4, filter.Length - 4);
            }

            assignmentsManager.UpdateCheckStatistics(filter);

            (int Counter, float CostSum) = assignmentsManager.GetResultsStatistics();

            label56.Text = "Oбщая стоимость:" + CostSum;
            label57.Text = "Количество:" + Counter;

            if (IsShowMessageBox)
            {
                NeedSplash = true;
                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;
            }

        }

        private void kryptonButton3_Click(object sender, EventArgs e)
        {
            Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных.\r\nПодождите..."); });
            T.Start();
            while (!SplashWindow.bSmallCreated) ;
            NeedSplash = false;

            kryptonDateTimePicker10.Value = DateTime.Today;
            kryptonDateTimePicker9.Value = DateTime.Today;

            kryptonRadioButton7.Checked = true;

            kryptonCheckBox13t.Checked = true;
            kryptonCheckBox12t.Checked = true;
            kryptonCheckBox10t.Checked = true;
            kryptonCheckBox9t.Checked = true;


            kryptonCheckBox4t.Checked = true;
            kryptonCheckBox3t.Checked = true;
            kryptonCheckBox2t.Checked = true;


            kryptonCheckBox16.Checked = true;
            kryptonCheckBox17.Checked = true;
            kryptonCheckBox18.Checked = true;

            assignmentsManager.UpdateDateStatistics();

            NeedSplash = true;
            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
        }

        private void kryptonButton7_Click(object sender, EventArgs e)
        {
            FilterStatisticsData();
        }

        private void kryptonCheckBox13t_CheckedChanged(object sender, EventArgs e)
        {
            FilterCheckStatistics();
        }

        private void kryptonCheckBox12t_CheckedChanged(object sender, EventArgs e)
        {
            FilterCheckStatistics();
        }

        private void kryptonCheckBox10t_CheckedChanged(object sender, EventArgs e)
        {
            FilterCheckStatistics();
        }

        private void kryptonCheckBox9t_CheckedChanged(object sender, EventArgs e)
        {
            FilterCheckStatistics();
        }

        private void kryptonCheckBox4t_CheckedChanged(object sender, EventArgs e)
        {
            FilterCheckStatistics();
        }

        private void kryptonCheckBox3t_CheckedChanged(object sender, EventArgs e)
        {
            FilterCheckStatistics();
        }

        private void kryptonCheckBox2t_CheckedChanged(object sender, EventArgs e)
        {
            FilterCheckStatistics();
        }

        private void kryptonCheckBox16_CheckedChanged(object sender, EventArgs e)
        {
            FilterCheckStatistics();
        }

        private void kryptonCheckBox17_CheckedChanged(object sender, EventArgs e)
        {
            FilterCheckStatistics();
        }

        private void kryptonCheckBox18_CheckedChanged(object sender, EventArgs e)
        {
            FilterCheckStatistics();
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            isDateOn = true;
            string filter = "";
            filter += kryptonRadioButton1.Checked ? "дате создания" : "";
            filter += kryptonRadioButton2.Checked ? "дате отгрузки" : "";
            filter += kryptonRadioButton3.Checked ? "дате согласования" : "";
            filter += kryptonRadioButton4.Checked ? "дате запуска" : "";
            filter += kryptonRadioButton5.Checked ? "дате выхода" : "";
            filter += kryptonRadioButton6.Checked ? "дате печати" : "";
            filter += kryptonRadioButton7.Checked ? "дате принятия на склад" : "";
            filter += kryptonRadioButton8.Checked ? "дате списания со склада" : "";
            filter += kryptonRadioButton9.Checked ? "по дате отправления на ОТК" : "";
            filter += kryptonRadioButton10.Checked ? "по дате принятия с ОТК" : "";

            DateTime DateStart = kryptonDateTimePicker10.Value;
            DateTime DataEnd = kryptonDateTimePicker9.Value;
            
            assignmentsManager.CreateReport
                ("Статистика по " + filter + " c " + DateStart.ToString("yyyy-MM-dd") + " по " + DateStart.ToString("yyyy-MM-dd"));
        }

        private void percentageDataGrid2_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Right)
            {
                contextMenuStrip1.Show(new Point(Cursor.Position.X, Cursor.Position.Y));
            }
        }
    }
}
