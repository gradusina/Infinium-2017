using Infinium.Modules.Permits;

using System;
using System.Threading;
using System.Windows.Forms;

namespace Infinium
{
    public partial class RegisterManualInputForm : Form
    {
        private const int eHide = 2;
        private const int eShow = 1;
        private const int eClose = 3;

        private const int iAdminRole = 107;

        private int FormEvent = 0;

        //CheckBox headerCheckBox = new CheckBox();

        private Form TopForm = null;
        private LightStartForm LightStartForm;

        private RegisterManualInput RegisterManualInput;
        private HistoryDispatch packagesOnExp;
        private HistoryDispatch packagesOnStore;
        public enum RoleTypes
        {
            OrdinaryRole = 0,
            AdminRole = 1
        }

        public RegisterManualInputForm(LightStartForm tLightStartForm)
        {
            InitializeComponent();

            LightStartForm = tLightStartForm;

            this.MaximumSize = Screen.PrimaryScreen.WorkingArea.Size;


            while (!SplashForm.bCreated) ;
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
                }

                return;
            }
        }

        private void RegisterManualInputForm_Shown(object sender, EventArgs e)
        {
            while (!SplashForm.bCreated) ;

            FormEvent = eShow;
            AnimateTimer.Enabled = true;
        }

        private void MenuCloseButton_Click(object sender, EventArgs e)
        {
            FormEvent = eClose;
            AnimateTimer.Enabled = true;
        }

        private void RegisterManualInputForm_Load(object sender, EventArgs e)
        {

            packagesOnExp = new HistoryDispatch();
            packagesOnExp.Initialize();

            packagesOnStore = new HistoryDispatch();
            packagesOnStore.Initialize();

            dgvFrontsOrdersOnExp.DataSource = packagesOnExp.FrontsOrdersList;
            dgvDecorOrdersOnExp.DataSource = packagesOnExp.DecorOrdersList;

            dgvFrontsOrdersOnStore.DataSource = packagesOnStore.FrontsOrdersList;
            dgvDecorOrdersOnStore.DataSource = packagesOnStore.DecorOrdersList;

            dgvFrontsOrdersSetting(ref dgvFrontsOrdersOnExp);
            dgvFrontsOrdersSetting(ref dgvFrontsOrdersOnStore);
            dgvDecorOrdersSetting(ref dgvDecorOrdersOnExp);
            dgvDecorOrdersSetting(ref dgvDecorOrdersOnStore);

            RegisterManualInput = new RegisterManualInput();

            RegisterManualInput.GetPermissions(Security.CurrentUserID, this.Name);
            if (RegisterManualInput.PermissionGranted(iAdminRole))
            {
                btnAcceptStore.Enabled = true;
                toolStripMenuItem1.Enabled = true;
            }

            dgvPackagesOnExp.DataSource = RegisterManualInput.PackagesOnExpBs;
            dgvPackagesOnStore.DataSource = RegisterManualInput.PackagesOnStoreBs;

            dgvPackagesGridsSettings(ref dgvPackagesOnExp);
            dgvPackagesGridsSettings(ref dgvPackagesOnStore);

            RegisterManualInput.RefreshData();
        }

        //private void HeaderCheckBox_Clicked(object sender, EventArgs e)
        //{
        //    //Necessary to end the edit mode of the Cell.
        //    dgvPackages.EndEdit();

        //    //Loop and check and uncheck all row CheckBoxes based on Header Cell CheckBox.
        //    foreach (DataGridViewRow row in dgvPackages.Rows)
        //    {
        //        DataGridViewCheckBoxCell checkBox = (row.Cells["CheckColumn"] as DataGridViewCheckBoxCell);
        //        checkBox.Value = headerCheckBox.Checked;
        //    }
        //}

        private void dgvPackagesGridsSettings(ref DataGridView dataGrid)
        {
            ////Find the Location of Header Cell.
            //Point headerCellLocation = this.dgvPackages.GetCellDisplayRectangle(0, -1, true).Location;

            ////Place the Header CheckBox in the Location of the Header Cell.
            //headerCheckBox.Location = new Point(headerCellLocation.X + 8, headerCellLocation.Y + 8);
            //headerCheckBox.BackColor = Color.White;
            //headerCheckBox.Size = new Size(18, 18);

            ////Assign Click event to the Header CheckBox.
            //headerCheckBox.Click += new EventHandler(HeaderCheckBox_Clicked);
            //dgvPackages.Controls.Add(headerCheckBox);

            //dgvPackages.Columns.Insert(0, RegisterManualInput.CheckColumn);

            foreach (DataGridViewColumn col in dataGrid.Columns)
            {
                col.ReadOnly = true;
            }

            dataGrid.Columns["MegaOrderID"].Visible = false;
            dataGrid.Columns["ProductType"].Visible = false;

            dataGrid.Columns["OrderNumber"].HeaderText = "№ заказа";
            dataGrid.Columns["MainOrderID"].HeaderText = "№ подзаказа";
            dataGrid.Columns["UserName"].HeaderText = "Кто сканировал";
            dataGrid.Columns["BarcodeType"].HeaderText = "Префикс штрихкода";
            dataGrid.Columns["ClientName"].HeaderText = "Клиент";
            dataGrid.Columns["AddPackDateTime"].HeaderText = "Дата сканирования";
            dataGrid.Columns["PackageID"].HeaderText = "ID упаковки";
            dataGrid.Columns["TrayID"].HeaderText = "ID поддона";
        }

        private void dgvFrontsOrdersSetting(ref DataGridView dataGrid)
        {
            if (!dataGrid.Columns.Contains("FrontsColumn"))
                dataGrid.Columns.Add(packagesOnExp.FrontsColumn);
            if (!dataGrid.Columns.Contains("FrameColorsColumn"))
                dataGrid.Columns.Add(packagesOnExp.FrameColorsColumn);
            if (!dataGrid.Columns.Contains("PatinaColumn"))
                dataGrid.Columns.Add(packagesOnExp.PatinaColumn);
            if (!dataGrid.Columns.Contains("InsetTypesColumn"))
                dataGrid.Columns.Add(packagesOnExp.InsetTypesColumn);
            if (!dataGrid.Columns.Contains("InsetColorsColumn"))
                dataGrid.Columns.Add(packagesOnExp.InsetColorsColumn);
            if (!dataGrid.Columns.Contains("TechnoProfilesColumn"))
                dataGrid.Columns.Add(packagesOnExp.TechnoProfilesColumn);
            if (!dataGrid.Columns.Contains("TechnoFrameColorsColumn"))
                dataGrid.Columns.Add(packagesOnExp.TechnoFrameColorsColumn);
            if (!dataGrid.Columns.Contains("TechnoInsetTypesColumn"))
                dataGrid.Columns.Add(packagesOnExp.TechnoInsetTypesColumn);
            if (!dataGrid.Columns.Contains("TechnoInsetColorsColumn"))
                dataGrid.Columns.Add(packagesOnExp.TechnoInsetColorsColumn);

            if (dataGrid.Columns.Contains("ImpostMargin"))
                dataGrid.Columns["ImpostMargin"].Visible = false;
            if (dataGrid.Columns.Contains("CreateDateTime"))
            {
                dataGrid.Columns["CreateDateTime"].HeaderText = "Добавлено";
                dataGrid.Columns["CreateDateTime"].DefaultCellStyle.Format = "dd.MM.yyyy HH:mm";
                dataGrid.Columns["CreateDateTime"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                dataGrid.Columns["CreateDateTime"].Width = 100;
            }
            if (dataGrid.Columns.Contains("CreateUserID"))
                dataGrid.Columns["CreateUserID"].Visible = false;
            dataGrid.Columns["FrontsOrdersID"].Visible = false;
            dataGrid.Columns["MainOrderID"].Visible = false;
            dataGrid.Columns["FrontID"].Visible = false;
            dataGrid.Columns["ColorID"].Visible = false;
            dataGrid.Columns["InsetColorID"].Visible = false;
            dataGrid.Columns["PatinaID"].Visible = false;
            dataGrid.Columns["InsetTypeID"].Visible = false;
            dataGrid.Columns["TechnoProfileID"].Visible = false;
            dataGrid.Columns["TechnoColorID"].Visible = false;
            dataGrid.Columns["TechnoInsetTypeID"].Visible = false;
            dataGrid.Columns["TechnoInsetColorID"].Visible = false;

            foreach (DataGridViewColumn Column in dataGrid.Columns)
            {
                Column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column.ReadOnly = true;
            }

            dataGrid.Columns["Height"].HeaderText = "Высота";
            dataGrid.Columns["Width"].HeaderText = "Ширина";
            dataGrid.Columns["Count"].HeaderText = "Кол-во";
            dataGrid.Columns["Notes"].HeaderText = "Примечание";
            dataGrid.Columns["Square"].HeaderText = "Площадь";

            dataGrid.Columns["FrontsColumn"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGrid.Columns["FrameColorsColumn"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGrid.Columns["PatinaColumn"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGrid.Columns["InsetTypesColumn"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGrid.Columns["InsetColorsColumn"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGrid.Columns["TechnoProfilesColumn"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGrid.Columns["TechnoFrameColorsColumn"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGrid.Columns["TechnoInsetTypesColumn"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGrid.Columns["TechnoInsetColorsColumn"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGrid.Columns["Height"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            dataGrid.Columns["Height"].Width = 85;
            dataGrid.Columns["Width"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            dataGrid.Columns["Width"].Width = 85;
            dataGrid.Columns["Count"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            dataGrid.Columns["Count"].Width = 65;
            dataGrid.Columns["Square"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            dataGrid.Columns["Square"].Width = 100;
            dataGrid.Columns["Notes"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGrid.Columns["Notes"].MinimumWidth = 105;

            dataGrid.AutoGenerateColumns = false;

            int DisplayIndex = 0;
            dataGrid.Columns["FrontsColumn"].DisplayIndex = DisplayIndex++;
            dataGrid.Columns["FrameColorsColumn"].DisplayIndex = DisplayIndex++;
            dataGrid.Columns["TechnoProfilesColumn"].DisplayIndex = DisplayIndex++;
            dataGrid.Columns["TechnoFrameColorsColumn"].DisplayIndex = DisplayIndex++;
            dataGrid.Columns["PatinaColumn"].DisplayIndex = DisplayIndex++;
            dataGrid.Columns["InsetTypesColumn"].DisplayIndex = DisplayIndex++;
            dataGrid.Columns["InsetColorsColumn"].DisplayIndex = DisplayIndex++;
            dataGrid.Columns["TechnoInsetTypesColumn"].DisplayIndex = DisplayIndex++;
            dataGrid.Columns["TechnoInsetColorsColumn"].DisplayIndex = DisplayIndex++;
            dataGrid.Columns["Height"].DisplayIndex = DisplayIndex++;
            dataGrid.Columns["Width"].DisplayIndex = DisplayIndex++;
            dataGrid.Columns["Count"].DisplayIndex = DisplayIndex++;
            dataGrid.Columns["Square"].DisplayIndex = DisplayIndex++;
            dataGrid.Columns["Notes"].DisplayIndex = DisplayIndex++;
            dataGrid.CellFormatting += DgvFrontsOrders_CellFormatting;
        }

        private void DgvFrontsOrders_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            DataGridView grid = (DataGridView)sender;
            if (grid.Columns.Contains("PatinaColumn") && (e.ColumnIndex == grid.Columns["PatinaColumn"].Index)
                                                      && e.Value != null)
            {
                DataGridViewCell cell =
                    grid.Rows[e.RowIndex].Cells[e.ColumnIndex];
                int PatinaID = -1;
                string DisplayName = string.Empty;
                if (grid.Rows[e.RowIndex].Cells["PatinaID"].Value != DBNull.Value)
                {
                    PatinaID = Convert.ToInt32(grid.Rows[e.RowIndex].Cells["PatinaID"].Value);
                    DisplayName = packagesOnExp.PatinaDisplayName(PatinaID);
                }
                cell.ToolTipText = DisplayName;
            }
        }

        private void dgvDecorOrdersSetting(ref DataGridView dataGrid)
        {
            if (!dataGrid.Columns.Contains("ProductColumn"))
                dataGrid.Columns.Add(packagesOnExp.ProductColumn);
            if (!dataGrid.Columns.Contains("ItemColumn"))
                dataGrid.Columns.Add(packagesOnExp.ItemColumn);
            if (!dataGrid.Columns.Contains("ColorColumn"))
                dataGrid.Columns.Add(packagesOnExp.ColorColumn);
            if (!dataGrid.Columns.Contains("PatinaColumn"))
                dataGrid.Columns.Add(packagesOnExp.DecorPatinaColumn);
            if (dataGrid.Columns.Contains("CreateDateTime"))
            {
                dataGrid.Columns["CreateDateTime"].HeaderText = "Добавлено";
                dataGrid.Columns["CreateDateTime"].DefaultCellStyle.Format = "dd.MM.yyyy HH:mm";
                dataGrid.Columns["CreateDateTime"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                dataGrid.Columns["CreateDateTime"].Width = 100;
            }
            if (dataGrid.Columns.Contains("CreateUserID"))
                dataGrid.Columns["CreateUserID"].Visible = false;
            dataGrid.Columns["DecorOrderID"].Visible = false;
            dataGrid.Columns["ProductID"].Visible = false;
            dataGrid.Columns["DecorID"].Visible = false;
            dataGrid.Columns["ColorID"].Visible = false;
            dataGrid.Columns["PatinaID"].Visible = false;

            foreach (DataGridViewColumn Column in dataGrid.Columns)
            {
                Column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column.ReadOnly = true;
            }

            dataGrid.Columns["Length"].HeaderText = "Длина";
            dataGrid.Columns["Height"].HeaderText = "Высота";
            dataGrid.Columns["Width"].HeaderText = "Ширина";
            dataGrid.Columns["Count"].HeaderText = "Кол-во";
            dataGrid.Columns["Notes"].HeaderText = "Примечание";

            dataGrid.Columns["ProductColumn"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGrid.Columns["ProductColumn"].MinimumWidth = 110;
            dataGrid.Columns["ItemColumn"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGrid.Columns["ItemColumn"].MinimumWidth = 110;
            dataGrid.Columns["ColorsColumn"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGrid.Columns["ColorsColumn"].MinimumWidth = 110;
            dataGrid.Columns["PatinaColumn"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGrid.Columns["PatinaColumn"].MinimumWidth = 110;
            dataGrid.Columns["Height"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            dataGrid.Columns["Height"].Width = 85;
            dataGrid.Columns["Width"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            dataGrid.Columns["Width"].Width = 85;
            dataGrid.Columns["Count"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            dataGrid.Columns["Count"].Width = 85;
            dataGrid.Columns["Length"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            dataGrid.Columns["Length"].Width = 85;
            dataGrid.Columns["Notes"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGrid.Columns["Notes"].MinimumWidth = 145;

            dataGrid.AutoGenerateColumns = false;

            int DisplayIndex = 0;
            dataGrid.Columns["ProductColumn"].DisplayIndex = DisplayIndex++;
            dataGrid.Columns["ItemColumn"].DisplayIndex = DisplayIndex++;
            dataGrid.Columns["ColorsColumn"].DisplayIndex = DisplayIndex++;
            dataGrid.Columns["PatinaColumn"].DisplayIndex = DisplayIndex++;
            dataGrid.Columns["Length"].DisplayIndex = DisplayIndex++;
            dataGrid.Columns["Height"].DisplayIndex = DisplayIndex++;
            dataGrid.Columns["Width"].DisplayIndex = DisplayIndex++;
            dataGrid.Columns["Count"].DisplayIndex = DisplayIndex++;
            dataGrid.Columns["Notes"].DisplayIndex = DisplayIndex++;
        }

        private void btnAcceptExp_Click(object sender, EventArgs e)
        {
            if (dgvPackagesOnExp.SelectedRows.Count == 0)
                return;

            bool okCancel = Infinium.LightMessageBox.Show(ref TopForm, true,
                "Подтвердить?",
                "Подтверждение сканирования");

            if (!okCancel)
                return;

            Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Сохранение данных.\r\nПодождите..."); });
            T.Start();
            while (!SplashWindow.bSmallCreated) ;

            try
            {
                int id = -1;
                int packageId = -1;
                int trayId = 0;
                int mainOrderId = -1;
                int megaOrderId = -1;

                if (dgvPackagesOnExp.SelectedRows.Count > 0 && dgvPackagesOnExp.SelectedRows[0].Cells["Id"].Value != DBNull.Value)
                    id = Convert.ToInt32(dgvPackagesOnExp.SelectedRows[0].Cells["Id"].Value);
                if (dgvPackagesOnExp.SelectedRows.Count > 0 && dgvPackagesOnExp.SelectedRows[0].Cells["PackageID"].Value != DBNull.Value)
                    packageId = Convert.ToInt32(dgvPackagesOnExp.SelectedRows[0].Cells["PackageID"].Value);
                if (dgvPackagesOnExp.SelectedRows.Count > 0 && dgvPackagesOnExp.SelectedRows[0].Cells["TrayID"].Value != DBNull.Value)
                    trayId = Convert.ToInt32(dgvPackagesOnExp.SelectedRows[0].Cells["TrayID"].Value);
                if (dgvPackagesOnExp.SelectedRows.Count > 0 && dgvPackagesOnExp.SelectedRows[0].Cells["MainOrderID"].Value != DBNull.Value)
                    mainOrderId = Convert.ToInt32(dgvPackagesOnExp.SelectedRows[0].Cells["MainOrderID"].Value);
                if (dgvPackagesOnExp.SelectedRows.Count > 0 && dgvPackagesOnExp.SelectedRows[0].Cells["MegaOrderID"].Value != DBNull.Value)
                    megaOrderId = Convert.ToInt32(dgvPackagesOnExp.SelectedRows[0].Cells["MegaOrderID"].Value);

                RegisterManualInput.SetEnable(id, false);
                RegisterManualInput.SetExp(packageId);

                if (trayId != 0)
                    CheckOrdersStatus.SetStatusMarketingForTray(trayId);
                else
                    CheckOrdersStatus.SetStatusMarketingForMainOrder(megaOrderId, mainOrderId);
                RegisterManualInput.RefreshData();
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;
                InfiniumTips.ShowTip(this, 50, 85, "Сохранено", 1700);
            }
        }

        //private void dgvPackages_CellClick(object sender, DataGridViewCellEventArgs e)
        //{
        //    //Check to ensure that the row CheckBox is clicked.
        //    if (e.RowIndex >= 0 && e.ColumnIndex == 0)
        //    {
        //        //Loop to verify whether all row CheckBoxes are checked or not.
        //        bool isChecked = true;
        //        foreach (DataGridViewRow row in dgvPackages.Rows)
        //        {
        //            if (Convert.ToBoolean(row.Cells["CheckColumn"].EditedFormattedValue) == false)
        //            {
        //                isChecked = false;
        //                break;
        //            }
        //        }
        //        headerCheckBox.Checked = isChecked;
        //    }
        //}

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            btnAcceptStore_Click(null, null);
        }

        private void FilterOrdersOnExp(int PackageID, int ProductType, int MainOrderID)
        {
            packagesOnExp.ClearFrontsOrders();
            packagesOnExp.ClearDecorOrders();

            if (dgvPackagesOnExp.SelectedRows.Count == 0)
                return;

            packagesOnExp.FilterFrontsOrders(PackageID, MainOrderID);
            packagesOnExp.FilterDecorOrders(PackageID, MainOrderID);

            tabControlOnExp.SelectedIndex = ProductType;
            dgvPackagesOnExp.Focus();
        }

        private void FilterOrdersOnStore(int PackageID, int ProductType, int MainOrderID)
        {
            packagesOnStore.ClearFrontsOrders();
            packagesOnStore.ClearDecorOrders();

            if (dgvPackagesOnStore.SelectedRows.Count == 0)
                return;

            packagesOnStore.FilterFrontsOrders(PackageID, MainOrderID);
            packagesOnStore.FilterDecorOrders(PackageID, MainOrderID);

            tabControlOnStore.SelectedIndex = ProductType;
            dgvPackagesOnStore.Focus();
        }

        private void dgvPackagesOnExp_SelectionChanged(object sender, EventArgs e)
        {
            if (packagesOnExp == null)
                return;

            int mainOrderId = -1;
            int packageId = -1;
            int productType = -1;
            if (dgvPackagesOnExp.SelectedRows.Count > 0 && dgvPackagesOnExp.SelectedRows[0].Cells["PackageID"].Value != DBNull.Value)
                packageId = Convert.ToInt32(dgvPackagesOnExp.SelectedRows[0].Cells["PackageID"].Value);
            if (dgvPackagesOnExp.SelectedRows.Count > 0 && dgvPackagesOnExp.SelectedRows[0].Cells["MainOrderID"].Value != DBNull.Value)
                mainOrderId = Convert.ToInt32(dgvPackagesOnExp.SelectedRows[0].Cells["MainOrderID"].Value);
            if (dgvPackagesOnExp.SelectedRows.Count > 0 && dgvPackagesOnExp.SelectedRows[0].Cells["ProductType"].Value != DBNull.Value)
                productType = Convert.ToInt32(dgvPackagesOnExp.SelectedRows[0].Cells["ProductType"].Value);

            FilterOrdersOnExp(packageId, productType, mainOrderId);
        }

        private void btnAcceptStore_Click(object sender, EventArgs e)
        {
            if (dgvPackagesOnStore.SelectedRows.Count == 0)
                return;

            bool okCancel = Infinium.LightMessageBox.Show(ref TopForm, true,
                "Подтвердить?",
                "Подтверждение сканирования");

            if (!okCancel)
                return;

            Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Сохранение данных.\r\nПодождите..."); });
            T.Start();
            while (!SplashWindow.bSmallCreated) ;

            try
            {
                int id = -1;
                int packageId = -1;
                int trayId = 0;
                int mainOrderId = -1;
                int megaOrderId = -1;

                if (dgvPackagesOnStore.SelectedRows.Count > 0 && dgvPackagesOnStore.SelectedRows[0].Cells["Id"].Value != DBNull.Value)
                    id = Convert.ToInt32(dgvPackagesOnStore.SelectedRows[0].Cells["Id"].Value);
                if (dgvPackagesOnStore.SelectedRows.Count > 0 && dgvPackagesOnStore.SelectedRows[0].Cells["PackageID"].Value != DBNull.Value)
                    packageId = Convert.ToInt32(dgvPackagesOnStore.SelectedRows[0].Cells["PackageID"].Value);
                if (dgvPackagesOnStore.SelectedRows.Count > 0 && dgvPackagesOnStore.SelectedRows[0].Cells["TrayID"].Value != DBNull.Value)
                    trayId = Convert.ToInt32(dgvPackagesOnStore.SelectedRows[0].Cells["TrayID"].Value);
                if (dgvPackagesOnStore.SelectedRows.Count > 0 && dgvPackagesOnStore.SelectedRows[0].Cells["MainOrderID"].Value != DBNull.Value)
                    mainOrderId = Convert.ToInt32(dgvPackagesOnStore.SelectedRows[0].Cells["MainOrderID"].Value);
                if (dgvPackagesOnStore.SelectedRows.Count > 0 && dgvPackagesOnStore.SelectedRows[0].Cells["MegaOrderID"].Value != DBNull.Value)
                    megaOrderId = Convert.ToInt32(dgvPackagesOnStore.SelectedRows[0].Cells["MegaOrderID"].Value);

                RegisterManualInput.SetEnable(id, false);
                RegisterManualInput.SetStore(packageId);

                if (trayId != 0)
                    CheckOrdersStatus.SetStatusMarketingForTray(trayId);
                else
                    CheckOrdersStatus.SetStatusMarketingForMainOrder(megaOrderId, mainOrderId);
                RegisterManualInput.RefreshData();
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;
                InfiniumTips.ShowTip(this, 50, 85, "Сохранено", 1700);
            }
        }

        private void dgvPackagesOnStore_SelectionChanged(object sender, EventArgs e)
        {
            if (packagesOnStore == null)
                return;

            int mainOrderId = -1;
            int packageId = -1;
            int productType = -1;
            if (dgvPackagesOnStore.SelectedRows.Count > 0 && dgvPackagesOnStore.SelectedRows[0].Cells["PackageID"].Value != DBNull.Value)
                packageId = Convert.ToInt32(dgvPackagesOnStore.SelectedRows[0].Cells["PackageID"].Value);
            if (dgvPackagesOnStore.SelectedRows.Count > 0 && dgvPackagesOnStore.SelectedRows[0].Cells["MainOrderID"].Value != DBNull.Value)
                mainOrderId = Convert.ToInt32(dgvPackagesOnStore.SelectedRows[0].Cells["MainOrderID"].Value);
            if (dgvPackagesOnStore.SelectedRows.Count > 0 && dgvPackagesOnStore.SelectedRows[0].Cells["ProductType"].Value != DBNull.Value)
                productType = Convert.ToInt32(dgvPackagesOnStore.SelectedRows[0].Cells["ProductType"].Value);

            FilterOrdersOnStore(packageId, productType, mainOrderId);
        }
    }
}
