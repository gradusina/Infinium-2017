using Infinium.Modules.Marketing.NewOrders;
using Infinium.Modules.Marketing.WeeklyPlanning;
using Infinium.Modules.StatisticsMarketing.Reports;

using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Windows.Forms;

namespace Infinium
{
    public partial class MarketingBatchForm : Form
    {
        private const int EHide = 2;
        private const int EShow = 1;
        private const int EClose = 3;
        private const int EMainMenu = 4;

        //const int iConfirmBatch = 64;
        private const int ICreateBatch = 34;
        private const int ICreateLabel = 40;
        private const int IAddToBatch = 35;
        private const int IInProduction = 22;
        private const int ICloseBatch = 38;

        private const int IAgreedProfil = 104;
        private const int IAgreedTps = 106;

        private bool _needRefresh = false;
        private bool _needSplash = false;
        private bool _bFrontType = false;
        private bool _bFrameColor = false;
        private bool _pickUpFronts = false;

        //bool ConfirmBatch = false;
        private bool _createBatch = false;
        private bool _createLabel = false;
        private bool _addToBatch = false;
        private bool _inProduction = false;
        private bool _closeBatch = false;

        private bool _bMegaBatchClose = false;

        private int _factoryId = 1;
        private int _formEvent = 0;

        private ImageList _imageList1;
        private RoleTypes _roleType = RoleTypes.Ordinary;

        public enum RoleTypes
        {
            Ordinary = 0,
            Admin = 1,
            Production = 2,
            Marketing = 3,
            Store = 4,
            Direction = 5,
            Control = 6
        }

        private ArrayList _batchArray = null;
        private ArrayList _frontIDs;
        private ArrayList _mainOrdersArray = null;

        private Bitmap _lockBw = new Bitmap(Properties.Resources.LockSmallBlack);
        private Bitmap _unlockBw = new Bitmap(Properties.Resources.UnlockSmallBlack);

        private LightStartForm _lightStartForm;

        private Form _topForm = null;
        private MarketingPickFrontsSelectForm _pickFrontsSelectForm;
        private DataTable _rolePermissionsDataTable;

        public BatchManager BatchManager;
        public DecorCatalogOrder DecorCatalogOrder;
        public BatchReport ReportMarketing;

        public MarketingBatchForm(LightStartForm tLightStartForm)
        {
            InitializeComponent();
            _lightStartForm = tLightStartForm;


            this.MaximumSize = Screen.PrimaryScreen.WorkingArea.Size;

            _rolePermissionsDataTable = RolesAndPermissionsManager.GetPermissions(Security.CurrentUserID, this.Name);

            // Construct the ImageList.
            _imageList1 = new ImageList()
            {

                // Set the ImageSize property to a larger size 
                // (the default is 16 x 16).
                ImageSize = new Size(24, 24)
            };

            // Add two images to the list.
            _imageList1.Images.Add(_lockBw);
            _imageList1.Images.Add(_unlockBw);
            //ImageList1.Images.Add("c:\\windows\\FeatherTexture.png");
            Initialize();

            while (!SplashForm.bCreated) ;
        }

        private bool PermissionGranted(int permissionId)
        {
            var rows = _rolePermissionsDataTable.Select("PermissionID = " + permissionId);

            if (rows.Count() > 0)
            {
                return Convert.ToBoolean(rows[0]["Granted"]);
            }

            return false;
        }

        private void MarketingBatchForm_Shown(object sender, EventArgs e)
        {
            if (ZOVProfilCheckButton.Checked && !ZOVTPSCheckButton.Checked)
            {
                _factoryId = 1;
                BatchDataGrid.Columns["ProfilConfirm"].Visible = true;
                BatchDataGrid.Columns["ProfilEnabledColumn"].Visible = true;
                BatchDataGrid.Columns["ProfilName"].Visible = true;
                BatchDataGrid.Columns["ProfilConfirmDateTime"].Visible = true;
                BatchDataGrid.Columns["ProfilConfirmUserColumn"].Visible = true;
                BatchDataGrid.Columns["ProfilCloseDateTime"].Visible = true;
                BatchDataGrid.Columns["ProfilCloseUserColumn"].Visible = true;
                BatchDataGrid.Columns["TPSConfirm"].Visible = false;
                BatchDataGrid.Columns["TPSEnabledColumn"].Visible = false;
                BatchDataGrid.Columns["TPSConfirmDateTime"].Visible = false;
                BatchDataGrid.Columns["TPSConfirmUserColumn"].Visible = false;
                BatchDataGrid.Columns["TPSCloseDateTime"].Visible = false;
                BatchDataGrid.Columns["TPSCloseUserColumn"].Visible = false;
                MegaBatchDataGrid.Columns["ProfilEntryDateTime"].Visible = true;
                MegaBatchDataGrid.Columns["TPSEntryDateTime"].Visible = false;
                MegaBatchDataGrid.Columns["ProfilEnabledColumn"].Visible = true;
                MegaBatchDataGrid.Columns["TPSEnabledColumn"].Visible = false;
            }
            if (!ZOVProfilCheckButton.Checked && ZOVTPSCheckButton.Checked)
            {
                _factoryId = 2;
                BatchDataGrid.Columns["ProfilConfirm"].Visible = false;
                BatchDataGrid.Columns["ProfilEnabledColumn"].Visible = false;
                BatchDataGrid.Columns["ProfilName"].Visible = false;
                BatchDataGrid.Columns["ProfilConfirmDateTime"].Visible = false;
                BatchDataGrid.Columns["ProfilCloseDateTime"].Visible = false;
                BatchDataGrid.Columns["ProfilConfirmUserColumn"].Visible = false;
                BatchDataGrid.Columns["ProfilCloseUserColumn"].Visible = false;
                BatchDataGrid.Columns["TPSConfirm"].Visible = true;
                BatchDataGrid.Columns["TPSEnabledColumn"].Visible = true;
                BatchDataGrid.Columns["TPSName"].Visible = true;
                BatchDataGrid.Columns["TPSConfirmDateTime"].Visible = true;
                BatchDataGrid.Columns["TPSConfirmUserColumn"].Visible = true;
                BatchDataGrid.Columns["TPSCloseDateTime"].Visible = true;
                BatchDataGrid.Columns["TPSCloseUserColumn"].Visible = true;
                MegaBatchDataGrid.Columns["ProfilEntryDateTime"].Visible = false;
                MegaBatchDataGrid.Columns["TPSEntryDateTime"].Visible = true;
                MegaBatchDataGrid.Columns["ProfilEnabledColumn"].Visible = false;
                MegaBatchDataGrid.Columns["TPSEnabledColumn"].Visible = true;
            }

            while (!SplashForm.bCreated) ;

            _formEvent = EShow;
            AnimateTimer.Enabled = true;

            _needSplash = true;
        }

        private void AnimateTimer_Tick(object sender, EventArgs e)
        {
            if (!DatabaseConfigsManager.Animation)
            {
                this.Opacity = 1;

                if (_formEvent == EClose || _formEvent == EHide)
                {
                    AnimateTimer.Enabled = false;

                    if (_formEvent == EClose)
                    {

                        _lightStartForm.CloseForm(this);
                        this.Close();
                    }

                    if (_formEvent == EHide)
                    {

                        _lightStartForm.HideForm(this);
                    }


                    return;
                }

                if (_formEvent == EShow)
                {
                    AnimateTimer.Enabled = false;
                    SplashForm.CloseS = true;
                    _needSplash = true;
                    return;
                }

            }

            if (_formEvent == EClose || _formEvent == EHide)
            {
                if (Convert.ToDecimal(this.Opacity) != Convert.ToDecimal(0.00))
                    this.Opacity = Convert.ToDouble(Convert.ToDecimal(this.Opacity) - Convert.ToDecimal(0.05));
                else
                {
                    AnimateTimer.Enabled = false;

                    if (_formEvent == EClose)
                    {

                        _lightStartForm.CloseForm(this);
                        this.Close();
                    }

                    if (_formEvent == EHide)
                    {

                        _lightStartForm.HideForm(this);
                    }

                }

                return;
            }


            if (_formEvent == EShow || _formEvent == EShow)
            {
                if (this.Opacity != 1)
                    this.Opacity += 0.05;
                else
                {
                    AnimateTimer.Enabled = false;
                    SplashForm.CloseS = true;
                    _needSplash = true;
                }

                return;
            }
        }

        private void NavigateMenuCloseButton_Click(object sender, EventArgs e)
        {
            _formEvent = EClose;
            AnimateTimer.Enabled = true;
        }

        private void MenuButton_Click(object sender, EventArgs e)
        {
            if (!MenuButton.Checked)
            {
                BatchFilterPanel.Visible = false;
                OrdersFilterPanel.Visible = false;
            }
            else
            {
                if (RolesAndPermissionsTabControl.SelectedTabPageIndex == 0)
                {
                    BatchFilterPanel.Visible = true;
                    OrdersFilterPanel.Visible = false;
                }
                if (RolesAndPermissionsTabControl.SelectedTabPageIndex == 1)
                {
                    BatchFilterPanel.Visible = false;
                    OrdersFilterPanel.Visible = true;
                }
            }
        }

        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);

            if (m.Msg == 6 && m.WParam.ToInt32() == 1)
            {
                if (_topForm != null)
                    _topForm.Activate();
            }
        }

        private void Initialize()
        {
            _frontIDs = new ArrayList();
            _mainOrdersArray = new ArrayList();

            DecorCatalogOrder = new DecorCatalogOrder();
            DecorCatalogOrder.Initialize(false);

            BatchManager = new BatchManager(
                ref MegaBatchDataGrid, ref BatchDataGrid,
                ref MainOrdersDataGrid, ref MainOrdersFrontsOrdersDataGrid, ref MegaOrdersDataGrid,
                ref MainOrdersDecorTabControl, ref MainOrdersTabControl,
                ref BatchMainOrdersDataGrid, ref BatchFrontsOrdersDataGrid, ref BatchMegaOrdersDataGrid,
                ref BatchDecorTabControl, ref BatchMainOrdersTabControl,
                ref DecorCatalogOrder,
                ref FrontsDataGrid, ref FrameColorsDataGrid, ref TechnoColorsDataGrid,
                ref InsetTypesDataGrid, ref InsetColorsDataGrid, ref TechnoInsetTypesDataGrid, ref TechnoInsetColorsDataGrid,
                ref SizesDataGrid,
                ref DecorProductsDataGrid, ref DecorItemsDataGrid, ref DecorColorsDataGrid, ref DecorSizesDataGrid,
                ref PreFrontsDataGrid, ref PreFrameColorsDataGrid, ref PreTechnoColorsDataGrid,
                ref PreInsetTypesDataGrid, ref PreInsetColorsDataGrid, ref PreTechnoInsetTypesDataGrid, ref PreTechnoInsetColorsDataGrid,
                ref PreSizesDataGrid,
                ref PreDecorProductsDataGrid, ref PreDecorItemsDataGrid, ref PreDecorColorsDataGrid, ref PreDecorSizesDataGrid);

            ReportMarketing = new BatchReport(ref DecorCatalogOrder);

        }

        private void Filter()
        {
            var notConfirm = NotConfirmCheckBox.Checked;
            var confirm = ConfirmCheckBox.Checked;
            var onProduction = OnProductionCheckBox.Checked;
            var notProduction = NotProductionCheckBox.Checked;
            var inProduction = InProductionCheckBox.Checked;

            var batchOnProduction = BatchOnProductionCheckBox.Checked;
            var batchNotProduction = BatchNotProductionCheckBox.Checked;
            var batchInProduction = BatchInProductionCheckBox.Checked;
            var batchOnStorage = BatchOnStorageCheckBox.Checked;
            var batchDispatched = BatchDispatchCheckBox.Checked;

            if (_needSplash)
            {
                var T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref _topForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                T.Start();

                while (!SplashWindow.bSmallCreated) ;
                _needSplash = false;

                BatchManager.Filter(_factoryId,
                    notConfirm, confirm, onProduction, notProduction, inProduction, _pickUpFronts);
                BatchManager.Filter_MegaBatches_ByFactory(_factoryId);
                BatchManager.GetMainOrdersNotInBatch(_factoryId);

                _needSplash = true;
                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;
            }
            else
            {
                BatchManager.Filter(_factoryId,
                    notConfirm, confirm, onProduction, notProduction, inProduction, _pickUpFronts);
                BatchManager.Filter_MegaBatches_ByFactory(_factoryId);
                BatchManager.GetMainOrdersNotInBatch(_factoryId);
            }
        }

        private void BatchDataGrid_SelectionChanged(object sender, EventArgs e)
        {
            if (_pickUpFronts)
                return;

            if (BatchManager != null)
                if (BatchManager.BatchBindingSource.Count > 0)
                {
                    if (((DataRowView)BatchManager.BatchBindingSource.Current)["BatchID"] != DBNull.Value)
                    {
                        var batchOnProduction = BatchOnProductionCheckBox.Checked;
                        var batchNotProduction = BatchNotProductionCheckBox.Checked;
                        var batchInProduction = BatchInProductionCheckBox.Checked;
                        var batchOnStorage = BatchOnStorageCheckBox.Checked;
                        var batchOnExp = BatchOnExpCheckBox.Checked;
                        var batchDispatched = BatchDispatchCheckBox.Checked;

                        var array = new ArrayList();

                        if (_needSplash)
                        {
                            var T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref _topForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                            T.Start();

                            while (!SplashWindow.bSmallCreated) ;

                            _needSplash = false;
                            array = BatchManager.FilterBatchMegaOrders(Convert.ToInt32(((DataRowView)BatchManager.BatchBindingSource.Current)["BatchID"]), _factoryId,
                                batchOnProduction, batchNotProduction, batchInProduction, batchOnStorage, batchOnExp, batchDispatched, _pickUpFronts);

                            if (GroupsCheckBox.Checked)
                            {
                                var batches = BatchManager.GetSelectedBatch().OfType<int>().ToArray();

                                if (batches.Count() > 0)
                                {
                                    var clientGroups = new ArrayList();
                                    if (MarketFilterCheckBox.Checked)
                                        clientGroups.Add(0);
                                    if (ZOVFilterCheckBox.Checked)
                                        clientGroups.Add(3);
                                    if (kryptonCheckBox1.Checked)
                                        clientGroups.Add(1);
                                    if (MoscowFilterCheckBox.Checked)
                                        clientGroups.Add(2);
                                    if (ZOVProfilCheckBox.Checked)
                                        clientGroups.Add(4);
                                    if (ZOVTPSCheckBox.Checked)
                                        clientGroups.Add(5);

                                    BatchManager.GetProductInfo(_factoryId,
                                        batchOnProduction, batchNotProduction, batchInProduction, batchOnStorage, batchOnExp, batchDispatched, batches, clientGroups.OfType<int>().ToArray());
                                }
                            }

                            _needSplash = true;

                            while (SplashWindow.bSmallCreated)
                                SmallWaitForm.CloseS = true;
                        }
                        else
                        {
                            array = BatchManager.FilterBatchMegaOrders(Convert.ToInt32(((DataRowView)BatchManager.BatchBindingSource.Current)["BatchID"]), _factoryId,
                                batchOnProduction, batchNotProduction, batchInProduction, batchOnStorage, batchOnExp, batchDispatched, _pickUpFronts);

                            if (GroupsCheckBox.Checked)
                            {
                                var batches = BatchManager.GetSelectedBatch().OfType<int>().ToArray();

                                if (batches.Count() > 0)
                                {
                                    var clientGroups = new ArrayList();
                                    if (MarketFilterCheckBox.Checked)
                                        clientGroups.Add(0);
                                    if (ZOVFilterCheckBox.Checked)
                                        clientGroups.Add(3);
                                    if (kryptonCheckBox1.Checked)
                                        clientGroups.Add(1);
                                    if (MoscowFilterCheckBox.Checked)
                                        clientGroups.Add(2);
                                    if (ZOVProfilCheckBox.Checked)
                                        clientGroups.Add(4);
                                    if (ZOVTPSCheckBox.Checked)
                                        clientGroups.Add(5);

                                    BatchManager.GetProductInfo(_factoryId,
                                        batchOnProduction, batchNotProduction, batchInProduction, batchOnStorage, batchOnExp, batchDispatched, batches, clientGroups.OfType<int>().ToArray());
                                }
                            }

                        }

                        BatchManager.CurrentBatchID = Convert.ToInt32(Convert.ToInt32(((DataRowView)BatchManager.BatchBindingSource.Current)["BatchID"]));

                        var batchId = Convert.ToInt32(((DataRowView)BatchManager.BatchBindingSource.Current)["BatchID"]);
                        var megaOrders = array.OfType<int>().ToArray();

                        BatchManager.GetMegaOrdersNotInProduction(batchId, megaOrders, _factoryId, _pickUpFronts);
                    }
                }
                else
                    BatchManager.ClearMegaOrders();
        }

        #region MegaOrders

        private void MegaOrdersDataGrid_SelectionChanged(object sender, EventArgs e)
        {
            ClientInfoLabel.Text = string.Empty;
            OrderLabel.Text = string.Empty;
            ConfirmDateLabel.Text = string.Empty;

            if (BatchManager != null)
                if (BatchManager.MegaOrdersBindingSource.Count > 0)
                {
                    if (((DataRowView)BatchManager.MegaOrdersBindingSource.Current)["MegaOrderID"] != DBNull.Value)
                    {
                        var onProduction = OnProductionCheckBox.Checked;
                        var notProduction = NotProductionCheckBox.Checked;
                        var inProduction = InProductionCheckBox.Checked;

                        if (_needSplash)
                        {
                            var T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref _topForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                            T.Start();

                            while (!SplashWindow.bSmallCreated) ;

                            _needSplash = false;

                            BatchManager.FilterMainOrdersByMegaOrder(
                                    Convert.ToInt32(((DataRowView)BatchManager.MegaOrdersBindingSource.Current)["MegaOrderID"]),
                                    _factoryId, onProduction, notProduction, inProduction);
                            _needSplash = true;

                            while (SplashWindow.bSmallCreated)
                                SmallWaitForm.CloseS = true;
                        }
                        else
                        {
                            BatchManager.FilterMainOrdersByMegaOrder(
                                    Convert.ToInt32(((DataRowView)BatchManager.MegaOrdersBindingSource.Current)["MegaOrderID"]),
                                    _factoryId, onProduction, notProduction, inProduction);
                        }

                        BatchManager.GetMainOrdersNotInBatch(_factoryId);

                        var megaOrders = BatchManager.GetSelectedMegaOrders().OfType<int>().ToArray();
                        Array.Sort(megaOrders);

                        if (megaOrders.Count() > 0)
                        {
                            xtraTabControl3.TabPages[1].PageVisible = true;
                            BatchManager.GetPreProductInfo(megaOrders, _factoryId,
                                onProduction, notProduction, inProduction);
                        }
                        else
                        {
                            xtraTabControl3.TabPages[1].PageVisible = false;
                        }

                        var clientName = string.Empty;
                        var orderNumber = string.Empty;
                        var isComplaint = false;
                        var confirmDate = string.Empty;

                        if (((DataRowView)BatchManager.MegaOrdersBindingSource.Current)["ClientName"] != DBNull.Value)
                            clientName = ((DataRowView)BatchManager.MegaOrdersBindingSource.Current)["ClientName"].ToString();
                        if (((DataRowView)BatchManager.MegaOrdersBindingSource.Current)["OrderNumber"] != DBNull.Value)
                            orderNumber = ((DataRowView)BatchManager.MegaOrdersBindingSource.Current)["OrderNumber"].ToString();
                        if (((DataRowView)BatchManager.MegaOrdersBindingSource.Current)["IsComplaint"] != DBNull.Value)
                            isComplaint = Convert.ToBoolean(((DataRowView)BatchManager.MegaOrdersBindingSource.Current)["IsComplaint"]);
                        if (((DataRowView)BatchManager.MegaOrdersBindingSource.Current)["ConfirmDateTime"] != DBNull.Value)
                            confirmDate = Convert.ToDateTime(((DataRowView)BatchManager.MegaOrdersBindingSource.Current)["ConfirmDateTime"]).ToString("dd.MM.yyyy");

                        var clientInfo = clientName;
                        var orderInfo = "Заказ: №" + orderNumber;
                        var confirmDateInfo = "Согласован: " + confirmDate;

                        if (isComplaint)
                            orderInfo += ", РЕКЛАМАЦИЯ";

                        ClientInfoLabel.Text = clientInfo;
                        OrderLabel.Text = orderInfo;
                        ConfirmDateLabel.Text = confirmDateInfo;

                        var mainOrders = BatchManager.GetMainOrders().OfType<int>().ToArray();
                        Array.Sort(mainOrders);

                        //decimal Square = BatchManager.GetSelectedSquare(MainOrders);
                        //if (Square > 0)
                        //    SquareLabel.Text = "Площадь выбранных подзаказов: " + Square + " м.кв";

                    }
                }
                else
                    BatchManager.MainOrdersDataTable.Clear();
        }

        private void MegaOrdersDataGrid_MouseMove(object sender, MouseEventArgs e)
        {
            if (_needRefresh == true)
            {
                MegaOrdersDataGrid.Refresh();
                _needRefresh = false;
            }
        }

        private void MegaOrdersDataGrid_Scroll(object sender, ScrollEventArgs e)
        {
            if (e.ScrollOrientation == ScrollOrientation.HorizontalScroll)
                _needRefresh = true;
        }

        #endregion

        #region MainOrders

        private void MainOrdersDataGrid_SelectionChanged(object sender, EventArgs e)
        {
            if (BatchManager != null)
                if (BatchManager.MainOrdersBindingSource.Count > 0)
                {
                    if (((DataRowView)BatchManager.MainOrdersBindingSource.Current)["MainOrderID"] != DBNull.Value)
                    {
                        BatchManager.FilterProductsByMainOrder(Convert.ToInt32(((DataRowView)BatchManager.MainOrdersBindingSource.Current)["MainOrderID"]), _factoryId);
                        if (MainOrdersTabControl.TabPages[0].PageVisible && MainOrdersTabControl.TabPages[1].PageVisible)
                            MainOrdersTabControl.SelectedTabPage = MainOrdersTabControl.TabPages[0];

                    }
                }
                else
                {
                    BatchManager.FilterProductsByMainOrder(-1, _factoryId);
                    BatchManager.ClearPreProductsGrids();
                    //MainOrdersTabControl.SelectedTab = MainOrdersTabControl.Tabs[0];
                }
        }

        private void MainOrdersDataGrid_MouseMove(object sender, MouseEventArgs e)
        {
            if (_needRefresh == true)
            {
                MainOrdersDataGrid.Refresh();
                _needRefresh = false;
            }
        }

        private void MainOrdersDataGrid_Scroll(object sender, ScrollEventArgs e)
        {
            if (e.ScrollOrientation == ScrollOrientation.HorizontalScroll)
                _needRefresh = true;
        }

        private void MainOrdersDataGrid_Sorted(object sender, EventArgs e)
        {
            BatchManager.GetMainOrdersNotInBatch(_factoryId);
        }

        #endregion

        #region BatchMegaOrders

        private void BatchMegaOrdersDataGrid_SelectionChanged(object sender, EventArgs e)
        {
            if (BatchManager != null)
                if (BatchManager.BatchMegaOrdersBindingSource.Count > 0)
                {
                    if (((DataRowView)BatchManager.BatchBindingSource.Current)["BatchID"] != DBNull.Value)
                    {
                        var batchId = Convert.ToInt32(((DataRowView)BatchManager.BatchBindingSource.Current)["BatchID"]);
                        var batchOnProduction = BatchOnProductionCheckBox.Checked;
                        var batchNotProduction = BatchNotProductionCheckBox.Checked;
                        var batchInProduction = BatchInProductionCheckBox.Checked;
                        var batchOnStorage = BatchOnStorageCheckBox.Checked;
                        var batchOnExp = BatchOnExpCheckBox.Checked;
                        var batchDispatched = BatchDispatchCheckBox.Checked;

                        var mainOrdersArray = new ArrayList();

                        if (_needSplash)
                        {
                            var T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref _topForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                            T.Start();

                            while (!SplashWindow.bSmallCreated) ;

                            _needSplash = false;
                            if (_pickUpFronts)
                            {
                                if (_bFrontType)
                                {
                                    if (_frontIDs.Count > 0)
                                    {
                                        mainOrdersArray = BatchManager.FilterBatchMainOrdersByFront(batchId,
                                            Convert.ToInt32(((DataRowView)BatchManager.BatchMegaOrdersBindingSource.Current)["MegaOrderID"]), _factoryId,
                                            batchOnProduction, batchNotProduction, batchInProduction, batchOnStorage, batchOnExp, batchDispatched, _frontIDs);
                                    }
                                }
                                if (_bFrameColor)
                                    mainOrdersArray = BatchManager.FilterBatchMainOrdersByFront(batchId,
                                        Convert.ToInt32(((DataRowView)BatchManager.BatchMegaOrdersBindingSource.Current)["MegaOrderID"]), _factoryId,
                                        batchOnProduction, batchNotProduction, batchInProduction, batchOnStorage, batchOnExp, batchDispatched);
                            }
                            else
                                mainOrdersArray = BatchManager.FilterBatchMainOrdersByMegaOrder(
                                    Convert.ToInt32(((DataRowView)BatchManager.BatchBindingSource.Current)["BatchID"]),
                                    Convert.ToInt32(((DataRowView)BatchManager.BatchMegaOrdersBindingSource.Current)["MegaOrderID"]), _factoryId,
                                    batchOnProduction, batchNotProduction, batchInProduction, batchOnStorage, batchOnExp, batchDispatched);

                            _needSplash = true;

                            while (SplashWindow.bSmallCreated)
                                SmallWaitForm.CloseS = true;
                        }
                        else
                        {
                            if (_pickUpFronts)
                            {
                                if (_bFrontType)
                                    mainOrdersArray = BatchManager.FilterBatchMainOrdersByFront(batchId,
                                        Convert.ToInt32(((DataRowView)BatchManager.BatchMegaOrdersBindingSource.Current)["MegaOrderID"]), _factoryId,
                                        batchOnProduction, batchNotProduction, batchInProduction, batchOnStorage, batchOnExp, batchDispatched, _frontIDs);
                                if (_bFrameColor)
                                    mainOrdersArray = BatchManager.FilterBatchMainOrdersByFront(batchId,
                                        Convert.ToInt32(((DataRowView)BatchManager.BatchMegaOrdersBindingSource.Current)["MegaOrderID"]), _factoryId,
                                        batchOnProduction, batchNotProduction, batchInProduction, batchOnStorage, batchOnExp, batchDispatched);
                            }
                            else
                                mainOrdersArray = BatchManager.FilterBatchMainOrdersByMegaOrder(
                                    Convert.ToInt32(((DataRowView)BatchManager.BatchBindingSource.Current)["BatchID"]),
                                    Convert.ToInt32(((DataRowView)BatchManager.BatchMegaOrdersBindingSource.Current)["MegaOrderID"]), _factoryId,
                                    batchOnProduction, batchNotProduction, batchInProduction, batchOnStorage, batchOnExp, batchDispatched);
                        }

                        if (!GroupsCheckBox.Checked)
                        {
                            var megaOrders = BatchManager.GetSelectedBatchMegaOrders().OfType<int>().ToArray();

                            if (megaOrders.Count() > 0)
                            {
                                var mainOrders = BatchManager.GetBatchMainOrders(batchId, megaOrders, _factoryId, _pickUpFronts).OfType<int>().ToArray();

                                if (mainOrders.Count() > 0)
                                {
                                    BatchManager.GetProductInfo(mainOrders, _factoryId,
                                        batchOnProduction, batchNotProduction, batchInProduction, batchOnStorage, batchOnExp, batchDispatched);
                                }
                            }
                        }
                    }
                }
                else
                {
                    BatchManager.ClearGrids();
                    BatchManager.ClearProductsGrids();
                }
        }

        private void BatchMegaOrdersDataGrid_MouseMove(object sender, MouseEventArgs e)
        {
            if (_needRefresh == true)
            {
                BatchMegaOrdersDataGrid.Refresh();
                _needRefresh = false;
            }
        }

        private void BatchMegaOrdersDataGrid_Scroll(object sender, ScrollEventArgs e)
        {
            if (e.ScrollOrientation == ScrollOrientation.HorizontalScroll)
                _needRefresh = true;
        }

        #endregion

        #region BatchMainOrders

        private void BatchMainOrdersDataGrid_SelectionChanged(object sender, EventArgs e)
        {
            if (BatchManager != null)
                if (BatchManager.BatchMainOrdersBindingSource.Count > 0)
                {
                    if (((DataRowView)BatchManager.BatchMainOrdersBindingSource.Current)["MainOrderID"] != DBNull.Value)
                    {
                        BatchManager.FilterBatchProductsByMainOrder(Convert.ToInt32(((DataRowView)BatchManager.BatchMainOrdersBindingSource.Current)["MainOrderID"]), _factoryId);
                        if (BatchMainOrdersTabControl.TabPages[0].PageVisible && BatchMainOrdersTabControl.TabPages[1].PageVisible)
                            BatchMainOrdersTabControl.SelectedTabPage = BatchMainOrdersTabControl.TabPages[0];
                    }
                }
                else
                {
                    BatchManager.FilterBatchProductsByMainOrder(-1, _factoryId);
                }
        }

        private void BatchMainOrdersDataGrid_MouseMove(object sender, MouseEventArgs e)
        {
            if (_needRefresh == true)
            {
                BatchMainOrdersDataGrid.Refresh();
                _needRefresh = false;
            }
        }

        private void BatchMainOrdersDataGrid_Scroll(object sender, ScrollEventArgs e)
        {
            if (e.ScrollOrientation == ScrollOrientation.HorizontalScroll)
                _needRefresh = true;
        }

        #endregion

        #region Batch

        private void AddBatchButton_Click(object sender, EventArgs e)
        {
            if (BatchManager == null)
                return;

            var okCancel = Infinium.LightMessageBox.Show(ref _topForm, true,
                    "Вы собираетесь создать новую партию. Продолжить?",
                    "Новая партия");
            if (!okCancel)
                return;

            if (BatchManager.IsBatchEmpty())
            {
                Infinium.LightMessageBox.Show(ref _topForm, false,
                    "Пустая партия уже создана, добавляйте заказы в неё",
                    "Создание партии");
                return;
            }
            if ((DataRowView)BatchManager.MegaBatchBindingSource.Current != null)
            {
                var megaBatchId = Convert.ToInt32(((DataRowView)BatchManager.MegaBatchBindingSource.Current).Row["MegaBatchID"]);
                BatchManager.AddBatch(megaBatchId, _factoryId);
            }
        }

        private void RemoveBatchButton_Click(object sender, EventArgs e)
        {
            if (BatchManager == null || BatchManager.BatchBindingSource.Count < 1)
                return;

            var okCancel = Infinium.LightMessageBox.Show(ref _topForm, true,
                    "Партия будет расформирована и удалена. Продолжить?",
                    "Удаление партии");

            if (okCancel)
            {
                var batchId = Convert.ToInt32(((DataRowView)BatchManager.BatchBindingSource.Current).Row["BatchID"]);

                if (!BatchManager.IsBatchEnabled(batchId, _factoryId))
                {
                    Infinium.LightMessageBox.Show(ref _topForm, false,
                           "Партия сформирована, удаление запрещено.",
                           "Ошибка удаления");
                    return;
                }

                BatchManager.RemoveBatch(batchId, _factoryId);
            }
        }

        #endregion

        #region Batch content

        private void AddToBatchButton_Click(object sender, EventArgs e)
        {
            if (BatchManager == null || BatchManager.BatchBindingSource.Count < 1)
                return;

            var batchId = Convert.ToInt32(((DataRowView)BatchManager.BatchBindingSource.Current).Row["BatchID"]);
            var megaOrderId = Convert.ToInt32(((DataRowView)BatchManager.MegaOrdersBindingSource.Current)["MegaOrderID"]);
            var mainOrders = BatchManager.GetMainOrders().OfType<int>().ToArray();
            if (mainOrders.Count() < 1)
                return;
            Array.Sort(mainOrders);

            if (!BatchManager.IsBatchEnabled(batchId, _factoryId))
            {
                Infinium.LightMessageBox.Show(ref _topForm, false,
                       "Партия сформирована, добавление заказов в неё запрещено. Создайте новую партию",
                       "Ошибка добавления");
                return;
            }

            ////Михаил Анатольевич Кондрашов
            //if (Security.CurrentUserID == 322 && NeedCheck)
            //{
            //    NeedCheck = false;
            //    Infinium.LightMessageBox.Show(ref TopForm, false,
            //           "Михаил Анатольевич, пожалуйста, проверьте партию, в которую добавляете заказ. После этого заново нажмите 'В партию'",
            //           "Напоминание");
            //    return;
            //}

            BatchManager.AddToBatch(batchId, mainOrders, _factoryId);
            BatchManager.SetMainOrderOnProduction(mainOrders, _factoryId, true);
            BatchManager.SetMegaOrderOnProduction(megaOrderId, _factoryId);

            if (_needSplash)
            {
                var T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref _topForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                T.Start();

                while (!SplashWindow.bSmallCreated) ;
                _needSplash = false;

                var notConfirm = NotConfirmCheckBox.Checked;
                var confirm = ConfirmCheckBox.Checked;
                var onProduction = OnProductionCheckBox.Checked;
                var notProduction = NotProductionCheckBox.Checked;
                var inProduction = InProductionCheckBox.Checked;

                var batchOnProduction = BatchOnProductionCheckBox.Checked;
                var batchNotProduction = BatchNotProductionCheckBox.Checked;
                var batchInProduction = BatchInProductionCheckBox.Checked;
                var batchOnStorage = BatchOnStorageCheckBox.Checked;
                var batchOnExp = BatchOnExpCheckBox.Checked;
                var batchDispatched = BatchDispatchCheckBox.Checked;

                var batchMegaOrders = BatchManager.GetBatchMegaOrders().OfType<int>().ToArray();
                BatchManager.Filter(_factoryId,
                    notConfirm, confirm, onProduction, notProduction, inProduction, _pickUpFronts);

                BatchManager.FilterBatchMegaOrders(batchId, _factoryId,
                    batchOnProduction, batchNotProduction, batchInProduction, batchOnStorage, batchOnExp, batchDispatched,
                    _pickUpFronts);

                BatchManager.GetMainOrdersNotInBatch(_factoryId);
                BatchManager.GetMegaOrdersNotInProduction(batchId, batchMegaOrders, _factoryId, _pickUpFronts);
                _needSplash = true;
                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;
            }
        }

        private void RemoveMegaOrderButton_Click(object sender, EventArgs e)
        {
            if (BatchManager == null || BatchManager.BatchMegaOrdersBindingSource.Count < 1)
                return;

            var batchId = Convert.ToInt32(((DataRowView)BatchManager.BatchBindingSource.Current).Row["BatchID"]);
            if (!BatchManager.IsBatchEnabled(batchId, _factoryId))
            {
                Infinium.LightMessageBox.Show(ref _topForm, false,
                       "Партия сформирована, удаление заказов из неё запрещено",
                       "Ошибка удаления");
                return;
            }

            var okCancel = Infinium.LightMessageBox.Show(ref _topForm, true,
                    "Выбранный заказ будет убран из партии. Продолжить?",
                    "Редактирование партии");

            if (okCancel)
            {
                var megaOrderId = Convert.ToInt32(((DataRowView)BatchManager.BatchMegaOrdersBindingSource.Current).Row["MegaOrderID"]);
                var megaOrders = new int[1] { megaOrderId };
                var mainOrders = BatchManager.GetBatchMainOrders(batchId, megaOrders, _factoryId, false).OfType<int>().ToArray();
                Array.Sort(mainOrders);

                if (BatchManager.RemoveMegaOrderFromBatch(batchId, _factoryId))
                {
                    BatchManager.SetMainOrderOnProduction(mainOrders, _factoryId, false);
                    BatchManager.SetMegaOrderOnProduction(megaOrderId, _factoryId);
                    Filter();
                    BatchManager.GetMainOrdersNotInBatch(_factoryId);
                }
            }
        }

        #endregion

        #region Set in production

        private void PrintMegaOrderButton_Click(object sender, EventArgs e)
        {
            if (BatchManager == null || BatchManager.BatchMegaOrdersBindingSource.Count < 1)
                return;

            var okCancel = Infinium.LightMessageBox.Show(ref _topForm, true,
                    "Выбранные заказы будут распечатаны и отправлены в производство. Продолжить?",
                    "В производство");
            if (!okCancel)
                return;

            var megaOrders = BatchManager.GetSelectedBatchMegaOrders().OfType<int>().ToArray();

            if (megaOrders.Count() < 1)
                return;

            var batchId = Convert.ToInt32(((DataRowView)BatchManager.BatchBindingSource.Current)["BatchID"]);
            var megaBatchId = Convert.ToInt32(((DataRowView)BatchManager.MegaBatchBindingSource.Current)["MegaBatchID"]);

            if (!BatchManager.HasOrders(megaOrders, _factoryId))
                return;

            var T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref _topForm, "Создание документа Excel.\r\nПодождите..."); });
            T.Start();

            while (!SplashWindow.bSmallCreated) ;

            var mainOrders = BatchManager.GetBatchMainOrders(batchId, megaOrders, _factoryId, false).OfType<int>().ToArray();
            //Array.Sort(MainOrders);

            ReportMarketing.CreateReport(megaBatchId, batchId, mainOrders, _factoryId);
            BatchManager.SetMainOrderInProduction(mainOrders, _factoryId);

            for (var i = 0; i < megaOrders.Count(); i++)
            {
                CheckOrdersStatus.SetMegaOrderStatus(megaOrders[i]);
            }

            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;

            var batchOnProduction = BatchOnProductionCheckBox.Checked;
            var batchNotProduction = BatchNotProductionCheckBox.Checked;
            var batchInProduction = BatchInProductionCheckBox.Checked;
            var batchOnStorage = BatchOnStorageCheckBox.Checked;
            var batchOnExp = BatchOnExpCheckBox.Checked;
            var batchDispatched = BatchDispatchCheckBox.Checked;

            BatchManager.FilterBatchMegaOrders(batchId, _factoryId,
                batchOnProduction, batchNotProduction, batchInProduction, batchOnStorage, batchOnExp, batchDispatched,
                _pickUpFronts);


            megaOrders = BatchManager.GetBatchMegaOrders().OfType<int>().ToArray();
            BatchManager.GetMegaOrdersNotInProduction(batchId, megaOrders, _factoryId, _pickUpFronts);
        }

        private void PrintMainOrderButton_Click(object sender, EventArgs e)
        {
            if (BatchManager == null || BatchManager.BatchMainOrdersBindingSource.Count < 1)
                return;

            var okCancel = Infinium.LightMessageBox.Show(ref _topForm, true,
                    "Выбранные подзаказы будут распечатаны и отправлены в производство. Продолжить?",
                    "В производство");
            if (!okCancel)
                return;

            var megaOrders = BatchManager.GetSelectedBatchMegaOrders().OfType<int>().ToArray();

            if (megaOrders.Count() < 1)
                return;

            var batchId = Convert.ToInt32(((DataRowView)BatchManager.BatchBindingSource.Current)["BatchID"]);
            var megaBatchId = Convert.ToInt32(((DataRowView)BatchManager.MegaBatchBindingSource.Current)["MegaBatchID"]);

            var megaOrderId = Convert.ToInt32(((DataRowView)BatchManager.BatchMegaOrdersBindingSource.Current).Row["MegaOrderID"]);

            if (!BatchManager.HasOrders(megaOrders, _factoryId))
                return;

            var T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref _topForm, "Создание документа Excel.\r\nПодождите..."); });
            T.Start();

            while (!SplashWindow.bSmallCreated) ;

            var mainOrders = BatchManager.GetSelectedBatchMainOrders().OfType<int>().ToArray();
            Array.Sort(mainOrders);

            if (mainOrders.Count() < 1)
                return;

            ReportMarketing.CreateReport(megaBatchId, batchId, mainOrders, _factoryId);
            BatchManager.SetMainOrderInProduction(mainOrders, _factoryId);
            CheckOrdersStatus.SetMegaOrderStatus(megaOrderId);

            var batchOnProduction = BatchOnProductionCheckBox.Checked;
            var batchNotProduction = BatchNotProductionCheckBox.Checked;
            var batchInProduction = BatchInProductionCheckBox.Checked;
            var batchOnStorage = BatchOnStorageCheckBox.Checked;
            var batchOnExp = BatchOnExpCheckBox.Checked;
            var batchDispatched = BatchDispatchCheckBox.Checked;

            BatchManager.FilterBatchMegaOrders(batchId, _factoryId,
                batchOnProduction, batchNotProduction, batchInProduction, batchOnStorage, batchOnExp, batchDispatched,
                _pickUpFronts);

            var batchMegaOrders = BatchManager.GetBatchMegaOrders().OfType<int>().ToArray();

            BatchManager.GetMegaOrdersNotInProduction(batchId, batchMegaOrders, _factoryId, _pickUpFronts);

            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
        }

        #endregion

        #region Filter functions

        private void NotAgreedCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (BatchManager == null)
                return;

            Filter();
        }

        private void NotConfirmCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (BatchManager == null)
                return;

            Filter();
        }

        private void ConfirmCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (BatchManager == null)
                return;

            Filter();
        }

        private void NotProductionCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (BatchManager == null)
                return;

            Filter();
        }

        private void OnProductionCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (BatchManager == null)
                return;

            Filter();
        }

        private void InProductionCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (BatchManager == null)
                return;

            Filter();
        }

        private void BatchNotProductionCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (BatchManager == null)
                return;

            Filter();
        }

        private void BatchOnProductionCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (BatchManager == null)
                return;

            Filter();
        }

        private void BatchInProductionCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (BatchManager == null)
                return;

            Filter();
        }

        private void BatchOnStorageCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (BatchManager == null)
                return;

            Filter();
        }

        private void BatchOnExpCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (BatchManager == null)
                return;

            Filter();
        }

        private void BatchDispatchCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (BatchManager == null)
                return;

            Filter();
        }

        #endregion

        private void RolesAndPermissionsTabControl_SelectedPageChanged(object sender, DevExpress.XtraTab.TabPageChangedEventArgs e)
        {
            BatchManager.GetMainOrdersNotInBatch(_factoryId);
            _needSplash = true;

            if (MenuButton.Checked)
            {
                if (RolesAndPermissionsTabControl.SelectedTabPageIndex == 1)
                {
                    OrdersFilterPanel.Visible = true;
                    BatchFilterPanel.Visible = false;
                }
                if (RolesAndPermissionsTabControl.SelectedTabPageIndex == 0)
                {
                    OrdersFilterPanel.Visible = false;
                    BatchFilterPanel.Visible = true;
                }
            }

        }

        private void DecorProductsDataGrid_SelectionChanged(object sender, EventArgs e)
        {
            if (BatchManager != null)
                if (BatchManager.DecorProductsSummaryBindingSource.Count > 0)
                {
                    if (((DataRowView)BatchManager.DecorProductsSummaryBindingSource.Current)["ProductID"] != DBNull.Value)
                    {
                        BatchManager.FilterDecorProducts(
                            Convert.ToInt32(((DataRowView)BatchManager.DecorProductsSummaryBindingSource.Current).Row["ProductID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.DecorProductsSummaryBindingSource.Current).Row["MeasureID"]));
                    }
                }

            DecorPogonLabel.Text = "0";
            DecorCountLabel.Text = "0";

            if (BatchManager != null)
                if (BatchManager.DecorProductsSummaryBindingSource.Count > 0)
                {
                    decimal pogon = 0;
                    var count = 0;

                    BatchManager.GetDecorInfo(ref pogon, ref count);

                    DecorPogonLabel.Text = pogon.ToString();
                    DecorCountLabel.Text = count.ToString();
                }
        }

        private void kryptonCheckSet1_CheckedButtonChanged(object sender, EventArgs e)
        {
            if (BatchManager == null)
                return;

            if (ZOVProfilCheckButton.Checked && !ZOVTPSCheckButton.Checked)
            {
                if (PermissionGranted(IAgreedProfil))
                {
                    btnAgreedProfil.Visible = true;
                }
                if (PermissionGranted(IAgreedTps))
                {
                    btnAgreedTPS.Visible = false;
                }

                _factoryId = 1;
                BatchDataGrid.Columns["ProfilConfirm"].Visible = true;
                BatchDataGrid.Columns["ProfilEnabledColumn"].Visible = true;
                BatchDataGrid.Columns["ProfilName"].Visible = true;
                BatchDataGrid.Columns["ProfilConfirmDateTime"].Visible = true;
                BatchDataGrid.Columns["ProfilConfirmUserColumn"].Visible = true;
                BatchDataGrid.Columns["ProfilCloseDateTime"].Visible = true;
                BatchDataGrid.Columns["ProfilCloseUserColumn"].Visible = true;
                BatchDataGrid.Columns["TPSConfirm"].Visible = false;
                BatchDataGrid.Columns["TPSEnabledColumn"].Visible = false;
                BatchDataGrid.Columns["TPSConfirmDateTime"].Visible = false;
                BatchDataGrid.Columns["TPSConfirmUserColumn"].Visible = false;
                BatchDataGrid.Columns["TPSCloseDateTime"].Visible = false;
                BatchDataGrid.Columns["TPSCloseUserColumn"].Visible = false;
                MegaBatchDataGrid.Columns["ProfilEntryDateTime"].Visible = true;
                MegaBatchDataGrid.Columns["TPSEntryDateTime"].Visible = false;
                MegaBatchDataGrid.Columns["ProfilEnabledColumn"].Visible = true;
                MegaBatchDataGrid.Columns["TPSEnabledColumn"].Visible = false;
            }
            if (!ZOVProfilCheckButton.Checked && ZOVTPSCheckButton.Checked)
            {
                if (PermissionGranted(IAgreedProfil))
                {
                    btnAgreedProfil.Visible = false;
                }
                if (PermissionGranted(IAgreedTps))
                {
                    btnAgreedTPS.Visible = true;
                }

                _factoryId = 2;
                BatchDataGrid.Columns["ProfilConfirm"].Visible = false;
                BatchDataGrid.Columns["ProfilEnabledColumn"].Visible = false;
                BatchDataGrid.Columns["ProfilName"].Visible = false;
                BatchDataGrid.Columns["ProfilConfirmDateTime"].Visible = false;
                BatchDataGrid.Columns["ProfilConfirmUserColumn"].Visible = false;
                BatchDataGrid.Columns["ProfilCloseDateTime"].Visible = false;
                BatchDataGrid.Columns["ProfilCloseUserColumn"].Visible = false;
                BatchDataGrid.Columns["TPSConfirm"].Visible = true;
                BatchDataGrid.Columns["TPSEnabledColumn"].Visible = true;
                BatchDataGrid.Columns["TPSName"].Visible = true;
                BatchDataGrid.Columns["TPSConfirmDateTime"].Visible = true;
                BatchDataGrid.Columns["TPSConfirmUserColumn"].Visible = true;
                BatchDataGrid.Columns["TPSCloseDateTime"].Visible = true;
                BatchDataGrid.Columns["TPSCloseUserColumn"].Visible = true;
                MegaBatchDataGrid.Columns["ProfilEntryDateTime"].Visible = false;
                MegaBatchDataGrid.Columns["TPSEntryDateTime"].Visible = true;
                MegaBatchDataGrid.Columns["ProfilEnabledColumn"].Visible = false;
                MegaBatchDataGrid.Columns["TPSEnabledColumn"].Visible = true;
            }

            Filter();
        }

        private void MegaOrdersDataGrid_Sorted(object sender, EventArgs e)
        {
            BatchManager.GetMainOrdersNotInBatch(_factoryId);
        }

        private void BatchMegaOrdersDataGrid_Sorted(object sender, EventArgs e)
        {
            var batchId = Convert.ToInt32(((DataRowView)BatchManager.BatchBindingSource.Current)["BatchID"]);
            var megaOrders = BatchManager.GetBatchMegaOrders().OfType<int>().ToArray();

            BatchManager.GetMegaOrdersNotInProduction(batchId, megaOrders, _factoryId, _pickUpFronts);

            //BatchManager.GetMainOrdersNotInProduction(FactoryID, PickUpFronts);
        }

        private void BatchMainOrdersDataGrid_Sorted(object sender, EventArgs e)
        {
            BatchManager.GetMainOrdersNotInProduction(_factoryId, _pickUpFronts);
        }

        private void RolesAndPermissionsTabControl_SelectedPageChanging(object sender, DevExpress.XtraTab.TabPageChangingEventArgs e)
        {
            _needSplash = false;
        }

        private void OnStorageCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (BatchManager == null)
                return;

            Filter();
        }

        private void DispatchCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (BatchManager == null)
                return;

            Filter();
        }

        private void PickUpFrontsButton_Click(object sender, EventArgs e)
        {
            var batchOnProduction = BatchOnProductionCheckBox.Checked;
            var batchNotProduction = BatchNotProductionCheckBox.Checked;
            var batchInProduction = BatchInProductionCheckBox.Checked;
            var batchOnStorage = BatchOnStorageCheckBox.Checked;
            var batchOnExp = BatchOnExpCheckBox.Checked;
            var batchDispatched = BatchDispatchCheckBox.Checked;

            if (PickUpFrontsButton.Tag.ToString() == "true")
            {
                this.PickUpFrontsButton.OverrideDefault.Back.Color1 = System.Drawing.Color.Crimson;
                this.PickUpFrontsButton.StateCommon.Back.Color1 = System.Drawing.Color.Crimson;
                this.PickUpFrontsButton.StateTracking.Back.Color1 = System.Drawing.Color.Crimson;

                MegaBatchDataGrid.Enabled = false;
                BatchDataGrid.Enabled = false;
                PickUpFrontsButton.Tag = "false";
                PickUpFrontsButton.Text = "Обновить";
                _pickUpFronts = true;

                var phantomForm = new Infinium.PhantomForm();
                phantomForm.Show();

                _pickFrontsSelectForm = new MarketingPickFrontsSelectForm(_factoryId, ref _frontIDs);

                _topForm = _pickFrontsSelectForm;
                _pickFrontsSelectForm.ShowDialog();
                _topForm = null;

                phantomForm.Close();

                phantomForm.Dispose();

                if (_pickFrontsSelectForm.IsOKPress)
                {
                    _bFrontType = _pickFrontsSelectForm.bFrontType;
                    _bFrameColor = _pickFrontsSelectForm.bFrameColor;

                    _pickFrontsSelectForm.Dispose();
                    _pickFrontsSelectForm = null;
                    GC.Collect();

                    _needSplash = false;
                    var T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref _topForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                    T.Start();

                    while (!SplashWindow.bSmallCreated) ;

                    BatchManager.GetCurrentFrontID();
                    BatchManager.GetCurrentFrameColorID();

                    var array = new ArrayList();

                    var batchId = Convert.ToInt32(((DataRowView)BatchManager.BatchBindingSource.Current)["BatchID"]);
                    if (_bFrameColor)
                        array = BatchManager.PickUpFrontsTPS(_factoryId, batchId,
                            batchOnProduction, batchNotProduction, batchInProduction, batchOnStorage, batchOnExp, batchDispatched);

                    if (_bFrontType)
                    {
                        if (_frontIDs.Count > 0)
                        {
                            array = BatchManager.PickUpFrontsProfil(_factoryId, batchId,
                                batchOnProduction, batchNotProduction, batchInProduction, batchOnStorage, batchOnExp, batchDispatched, _frontIDs);
                        }
                    }

                    var megaOrders = array.OfType<int>().ToArray();

                    BatchManager.GetMegaOrdersNotInProduction(batchId, megaOrders, _factoryId, _pickUpFronts);

                    while (SplashWindow.bSmallCreated)
                        SmallWaitForm.CloseS = true;

                    _needSplash = true;
                }
                else
                {
                    _pickFrontsSelectForm.Dispose();
                    _pickFrontsSelectForm = null;
                    GC.Collect();
                }
                //Thread T = new Thread(delegate() { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                //T.Start();

                //while (!SplashWindow.bSmallCreated) ;

                //NeedSplash = false;

                //BatchManager.GetCurrentFrontID();
                //BatchManager.GetCurrentFrameColorID();

                //ArrayList array = BatchManager.PickUpFrontsTPS(FactoryID,
                //    BatchOnProduction, BatchNotProduction, BatchInProduction, BatchOnStorage, BatchDispatched);

                //int BatchID = Convert.ToInt32(((DataRowView)BatchManager.BatchBindingSource.Current)["BatchID"]);
                //int[] MegaOrders = array.OfType<int>().ToArray();

                //BatchManager.GetMegaOrdersNotInProduction(BatchID, MegaOrders, FactoryID, PickUpFronts);

                //NeedSplash = true;

                //while (SplashWindow.bSmallCreated)
                //    SmallWaitForm.CloseS = true;

                return;

            }
            else
            {
                this.PickUpFrontsButton.OverrideDefault.Back.Color1 = System.Drawing.Color.MediumPurple;
                this.PickUpFrontsButton.StateCommon.Back.Color1 = System.Drawing.Color.MediumPurple;
                this.PickUpFrontsButton.StateTracking.Back.Color1 = System.Drawing.Color.MediumPurple;

                MegaBatchDataGrid.Enabled = true;
                BatchDataGrid.Enabled = true;
                PickUpFrontsButton.Tag = "true";
                PickUpFrontsButton.Text = "Подобрать";
                _pickUpFronts = false;
                Filter();
                return;
            }
        }

        private void MegaBatchDataGrid_SelectionChanged(object sender, EventArgs e)
        {
            _bMegaBatchClose = true;
            if (_pickUpFronts)
                return;

            if (BatchManager != null)
                if (BatchManager.MegaBatchBindingSource.Count > 0)
                {
                    if (_factoryId == 1)
                        if (((DataRowView)BatchManager.MegaBatchBindingSource.Current)["ProfilBatchClose"] != DBNull.Value)
                            _bMegaBatchClose = Convert.ToBoolean(((DataRowView)BatchManager.MegaBatchBindingSource.Current)["ProfilBatchClose"]);
                    if (_factoryId == 2)
                        if (((DataRowView)BatchManager.MegaBatchBindingSource.Current)["TPSBatchClose"] != DBNull.Value)
                            _bMegaBatchClose = Convert.ToBoolean(((DataRowView)BatchManager.MegaBatchBindingSource.Current)["TPSBatchClose"]);
                    if (((DataRowView)BatchManager.MegaBatchBindingSource.Current)["MegaBatchID"] != DBNull.Value)
                    {
                        BatchManager.FilterBatch(Convert.ToInt32(((DataRowView)BatchManager.MegaBatchBindingSource.Current)["MegaBatchID"]), _factoryId);
                        BatchManager.CurrentMegaBatchID = Convert.ToInt32(Convert.ToInt32(((DataRowView)BatchManager.MegaBatchBindingSource.Current)["MegaBatchID"]));
                    }
                }
                else
                    BatchManager.ClearBatch();
            if (_bMegaBatchClose)
            {
                AddBatchButton.Enabled = false;
                EditBatchNameButton.Enabled = false;
                OKMoveOrdersButton.Enabled = false;
                MoveBatchButton.Enabled = false;
                CancelMoveOrdersButton.Enabled = false;
                PrintBatchButton.Enabled = false;
                RemoveBatchButton.Enabled = false;
                MoveMegaOrdersButton.Enabled = false;
                RemoveMegaOrderButton.Enabled = false;
                MoveMainOrdersButton.Enabled = false;
                AddToBatchButton.Enabled = false;
                PickUpFrontsButton.Enabled = false;

                cmiCloseProfilBatch.Enabled = false;
                cmiCloseTPSBatch.Enabled = false;
                cmiConfirmProfilBatch.Enabled = false;
                cmiConfirmTPSBatch.Enabled = false;
            }
            else
            {
                AddBatchButton.Enabled = true;
                EditBatchNameButton.Enabled = true;
                OKMoveOrdersButton.Enabled = true;
                MoveBatchButton.Enabled = true;
                CancelMoveOrdersButton.Enabled = true;
                PrintBatchButton.Enabled = true;
                RemoveBatchButton.Enabled = true;
                MoveMegaOrdersButton.Enabled = true;
                PickUpFrontsButton.Enabled = true;
                RemoveMegaOrderButton.Enabled = true;
                MoveMainOrdersButton.Enabled = true;
                AddToBatchButton.Enabled = true;

                cmiCloseProfilBatch.Enabled = true;
                cmiCloseTPSBatch.Enabled = true;
                cmiConfirmProfilBatch.Enabled = true;
                cmiConfirmTPSBatch.Enabled = true;
            }
        }

        private void AddMegaBatchButton_Click(object sender, EventArgs e)
        {
            if (BatchManager == null)
                return;

            var okCancel = Infinium.LightMessageBox.Show(ref _topForm, true,
                    "Вы собираетесь создать новую группу партий. Продолжить?",
                    "Группа партий");
            if (!okCancel)
                return;

            BatchManager.AddMegaBatch(_factoryId);
        }

        private void RefreshOrdersButton_Click(object sender, EventArgs e)
        {
            if (BatchManager == null)
                return;
            _pickUpFronts = false;
            Filter();
        }

        private void PrintBatchButton_Click(object sender, EventArgs e)
        {
            var okCancel = Infinium.LightMessageBox.Show(ref _topForm, true,
                    "Партия будет распечатана и отправлена в производство. Продолжить?",
                    "В производство");

            if (!okCancel)
                return;

            var T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref _topForm, "Создание документа Excel.\r\nПодождите..."); });
            T.Start();

            while (!SplashWindow.bSmallCreated) ;

            var megaBatchId = Convert.ToInt32(((DataRowView)BatchManager.MegaBatchBindingSource.Current)["MegaBatchID"]);
            var batchId = Convert.ToInt32(((DataRowView)BatchManager.BatchBindingSource.Current)["BatchID"]);
            ReportMarketing.CreateReport(megaBatchId, batchId, _factoryId);
            BatchManager.SetMainOrderInProduction(batchId, _factoryId);

            var batchOnProduction = BatchOnProductionCheckBox.Checked;
            var batchNotProduction = BatchNotProductionCheckBox.Checked;
            var batchInProduction = BatchInProductionCheckBox.Checked;
            var batchOnStorage = BatchOnStorageCheckBox.Checked;
            var batchOnExp = BatchOnExpCheckBox.Checked;
            var batchDispatched = BatchDispatchCheckBox.Checked;

            BatchManager.FilterBatchMegaOrders(batchId, _factoryId,
                batchOnProduction, batchNotProduction, batchInProduction, batchOnStorage, batchOnExp, batchDispatched,
                _pickUpFronts);

            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
        }

        //переносит заказы между партиями
        private void MoveOrdersButton_Click(object sender, EventArgs e)
        {
            var batchId = Convert.ToInt32(((DataRowView)BatchManager.BatchBindingSource.Current)["BatchID"]);
            //if (!BatchManager.IsBatchEnabled(BatchID, FactoryID))
            //{
            //    Infinium.LightMessageBox.Show(ref TopForm, false,
            //           "Партия сформирована, перенос заказов запрещен",
            //           "Ошибка добавления");
            //    return;
            //}

            var phantomForm = new Infinium.PhantomForm();
            phantomForm.Show();

            var batchOnProduction = BatchOnProductionCheckBox.Checked;
            var batchNotProduction = BatchNotProductionCheckBox.Checked;
            var batchInProduction = BatchInProductionCheckBox.Checked;
            var batchOnStorage = BatchOnStorageCheckBox.Checked;
            var batchOnExp = BatchOnExpCheckBox.Checked;
            var batchDispatched = BatchDispatchCheckBox.Checked;
            BatchManager.GetCurrentFrontID();
            BatchManager.GetCurrentFrameColorID();
            var megaOrders = new int[BatchMegaOrdersDataGrid.SelectedRows.Count];
            for (var i = 0; i < BatchMegaOrdersDataGrid.SelectedRows.Count; i++)
                megaOrders[i] = Convert.ToInt32(BatchMegaOrdersDataGrid.SelectedRows[i].Cells["MegaOrderID"].Value);

            _mainOrdersArray.Clear();

            if (_factoryId == 1)
            {
                BatchManager.GetCurrentFrontID();
                BatchManager.GetCurrentFrameColorID();

                if (_pickUpFronts)
                {
                    if (_bFrameColor)
                        _mainOrdersArray = BatchManager.GetBatchMainOrders(megaOrders, batchId, _factoryId,
                            batchOnProduction, batchNotProduction, batchInProduction, batchOnStorage, batchOnExp, batchDispatched, _pickUpFronts);

                    if (_bFrontType)
                        _mainOrdersArray = BatchManager.GetBatchMainOrders(megaOrders, batchId, _factoryId,
                            batchOnProduction, batchNotProduction, batchInProduction, batchOnStorage, batchOnExp, batchDispatched, _pickUpFronts, _frontIDs);
                }
                else
                {
                    _mainOrdersArray = BatchManager.GetBatchMainOrders(megaOrders, batchId, _factoryId,
                    batchOnProduction, batchNotProduction, batchInProduction, batchOnStorage, batchOnExp, batchDispatched, _pickUpFronts);
                }
            }
            if (_factoryId == 2)
            {
                _mainOrdersArray = BatchManager.GetBatchMainOrders(megaOrders, batchId, _factoryId,
                batchOnProduction, batchNotProduction, batchInProduction, batchOnStorage, batchOnExp, batchDispatched, _pickUpFronts);
            }

            _mainOrdersArray.Sort();
            if (_mainOrdersArray.Count < 1)
            {
                Infinium.LightMessageBox.Show(ref _topForm, false,
                          "Не выбран ни один подзаказ",
                          "Ошибка переноса");

                phantomForm.Close();

                phantomForm.Dispose();

                return;
            }

            var selectMenuForm = new MarketingBatchSelectMenu(this, BatchManager, _factoryId);

            _topForm = selectMenuForm;
            selectMenuForm.ShowDialog();

            phantomForm.Close();

            phantomForm.Dispose();
            selectMenuForm.Dispose();
            _topForm = null;

            if (BatchManager.NewBatch)
            {
                //int[] MegaOrders = MegaOrdersArray.OfType<int>().ToArray();

                //ArrayList array = new ArrayList();

                BatchManager.MoveOrdersToNewPart(_mainOrdersArray.OfType<int>().ToArray(), _factoryId);

                this.PickUpFrontsButton.OverrideDefault.Back.Color1 = System.Drawing.Color.MediumPurple;
                this.PickUpFrontsButton.StateCommon.Back.Color1 = System.Drawing.Color.MediumPurple;
                this.PickUpFrontsButton.StateTracking.Back.Color1 = System.Drawing.Color.MediumPurple;

                MegaBatchDataGrid.Enabled = true;
                BatchDataGrid.Enabled = true;
                PickUpFrontsButton.Tag = "true";
                PickUpFrontsButton.Text = "Подобрать";
                _pickUpFronts = false;

                BatchManager.FilterBatch(Convert.ToInt32(((DataRowView)BatchManager.MegaBatchBindingSource.Current)["MegaBatchID"]), _factoryId);
                Filter();
            }

            if (BatchManager.OldBatch)
            {
                MegaBatchDataGrid.Enabled = true;
                BatchDataGrid.Enabled = true;

                AddBatchButton.Visible = false;
                RemoveBatchButton.Visible = false;
                PrintBatchButton.Visible = false;
                MoveBatchButton.Visible = false;
                EditBatchNameButton.Visible = false;
                OKMoveOrdersButton.Visible = true;
                CancelMoveOrdersButton.Visible = true;
            }
        }

        //подтверждает перенос заказов
        private void OKMoveOrdersButton_Click(object sender, EventArgs e)
        {
            var batchOnProduction = BatchOnProductionCheckBox.Checked;
            var batchNotProduction = BatchNotProductionCheckBox.Checked;
            var batchInProduction = BatchInProductionCheckBox.Checked;
            var batchOnStorage = BatchOnStorageCheckBox.Checked;
            var batchOnExp = BatchOnExpCheckBox.Checked;
            var batchDispatched = BatchDispatchCheckBox.Checked;

            var mainOrders = _mainOrdersArray.OfType<int>().ToArray();

            //int[] MainOrders = BatchManager.GetBatchMainOrders(MegaOrders, FactoryID,
            //    BatchOnProduction, BatchNotProduction, BatchInProduction, BatchOnStorage,BatchOnExp, BatchDispatched, PickUpFronts).OfType<int>().ToArray();
            //if (FactoryID == 1)
            //    MainOrders = BatchManager.GetBatchMainOrders(MegaOrders, FactoryID,
            //        BatchOnProduction, BatchNotProduction, BatchInProduction, BatchOnStorage,BatchOnExp, BatchDispatched, PickUpFronts, FrontIDs).OfType<int>().ToArray();
            //Array.Sort(MainOrders);

            if (BatchManager.OldBatch)
            {
                var selectedBatchId = Convert.ToInt32(((DataRowView)BatchManager.BatchBindingSource.Current)["BatchID"]);
                BatchManager.MoveOrdersToSelectedBatch(selectedBatchId, mainOrders, _factoryId);
                BatchManager.FilterBatch(Convert.ToInt32(((DataRowView)BatchManager.MegaBatchBindingSource.Current)["MegaBatchID"]), _factoryId);
            }

            AddBatchButton.Visible = true;
            RemoveBatchButton.Visible = true;
            PrintBatchButton.Visible = true;
            MoveBatchButton.Visible = true;
            EditBatchNameButton.Visible = true;
            OKMoveOrdersButton.Visible = false;
            CancelMoveOrdersButton.Visible = false;
        }

        private void CancelMoveOrdersButton_Click(object sender, EventArgs e)
        {
            AddBatchButton.Visible = true;
            RemoveBatchButton.Visible = true;
            PrintBatchButton.Visible = true;
            MoveBatchButton.Visible = true;
            EditBatchNameButton.Visible = true;
            OKMoveOrdersButton.Visible = false;
            CancelMoveOrdersButton.Visible = false;
        }

        //переносит партии между группами
        private void MoveBatchButton_Click(object sender, EventArgs e)
        {
            if (BatchManager == null || BatchManager.BatchBindingSource.Count < 1)
                return;

            Infinium.LightMessageBox.Show(ref _topForm, false,
                    "Выберите Группу партий",
                    "Перенос партий");

            _batchArray = BatchManager.GetSelectedBatch();
            OKMoveBatchButton.Visible = true;
            CancelMoveBatchButton.Visible = true;
            AddMegaBatchButton.Visible = false;
        }

        //подтверждает перенос партий
        private void OKMoveBatchButton_Click(object sender, EventArgs e)
        {
            if (BatchManager == null || BatchManager.MegaBatchBindingSource.Count < 1)
                return;

            var batch = _batchArray.OfType<int>().ToArray();
            var destMegaBatchId = Convert.ToInt32(((DataRowView)BatchManager.MegaBatchBindingSource.Current)["MegaBatchID"]);

            BatchManager.MoveBatch(batch, destMegaBatchId, _factoryId);

            OKMoveBatchButton.Visible = false;
            CancelMoveBatchButton.Visible = false;
            AddMegaBatchButton.Visible = true;
        }

        private void CancelMoveBatchButton_Click(object sender, EventArgs e)
        {
            OKMoveBatchButton.Visible = false;
            CancelMoveBatchButton.Visible = false;
            AddMegaBatchButton.Visible = true;
        }

        private void FrontsDataGrid_SelectionChanged(object sender, EventArgs e)
        {
            if (BatchManager != null)
                if (BatchManager.FrontsSummaryBindingSource.Count > 0)
                {
                    if (((DataRowView)BatchManager.FrontsSummaryBindingSource.Current)["FrontID"] != DBNull.Value)
                    {
                        BatchManager.FilterFrameColors(
                            Convert.ToInt32(((DataRowView)BatchManager.FrontsSummaryBindingSource.Current).Row["FrontID"]));
                    }
                }

            FrontsSquareLabel.Text = "0";
            FrontsCountLabel.Text = "0";

            if (BatchManager != null)
                if (BatchManager.FrontsSummaryBindingSource.Count > 0)
                {
                    decimal frontSquare = 0;
                    var count = 0;

                    BatchManager.GetFrontsInfo(ref frontSquare, ref count);

                    FrontsSquareLabel.Text = frontSquare.ToString();
                    FrontsCountLabel.Text = count.ToString();
                }
        }

        private void FrameColorsDataGrid_SelectionChanged(object sender, EventArgs e)
        {
            if (BatchManager != null)
                if (BatchManager.FrameColorsSummaryBindingSource.Count > 0)
                {
                    if (((DataRowView)BatchManager.FrameColorsSummaryBindingSource.Current)["ColorID"] != DBNull.Value)
                    {
                        BatchManager.FilterTechnoColors(
                            Convert.ToInt32(((DataRowView)BatchManager.FrameColorsSummaryBindingSource.Current).Row["FrontID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.FrameColorsSummaryBindingSource.Current).Row["ColorID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.FrameColorsSummaryBindingSource.Current).Row["PatinaID"]));
                    }
                }
        }

        private void TechnoColorsDataGrid_SelectionChanged(object sender, EventArgs e)
        {
            if (BatchManager != null)
                if (BatchManager.TechnoColorsSummaryBindingSource.Count > 0)
                {
                    if (((DataRowView)BatchManager.TechnoColorsSummaryBindingSource.Current)["TechnoColorID"] != DBNull.Value)
                    {
                        BatchManager.FilterInsetTypes(
                            Convert.ToInt32(((DataRowView)BatchManager.TechnoColorsSummaryBindingSource.Current).Row["FrontID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.TechnoColorsSummaryBindingSource.Current).Row["ColorID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.TechnoColorsSummaryBindingSource.Current).Row["PatinaID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.TechnoColorsSummaryBindingSource.Current).Row["TechnoColorID"]));
                    }
                }
        }

        private void InsetTypesDataGrid_SelectionChanged(object sender, EventArgs e)
        {
            if (BatchManager != null)
                if (BatchManager.InsetTypesSummaryBindingSource.Count > 0)
                {
                    if (((DataRowView)BatchManager.InsetTypesSummaryBindingSource.Current)["InsetTypeID"] != DBNull.Value)
                    {
                        BatchManager.FilterInsetColors(
                            Convert.ToInt32(((DataRowView)BatchManager.InsetTypesSummaryBindingSource.Current).Row["FrontID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.InsetTypesSummaryBindingSource.Current).Row["ColorID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.InsetTypesSummaryBindingSource.Current).Row["PatinaID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.InsetTypesSummaryBindingSource.Current).Row["TechnoColorID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.InsetTypesSummaryBindingSource.Current).Row["InsetTypeID"]));
                    }
                }
        }

        private void InsetColorsDataGrid_SelectionChanged(object sender, EventArgs e)
        {
            if (BatchManager != null)
                if (BatchManager.InsetColorsSummaryBindingSource.Count > 0)
                {
                    if (((DataRowView)BatchManager.InsetColorsSummaryBindingSource.Current)["InsetColorID"] != DBNull.Value)
                    {
                        BatchManager.FilterTechnoInsetTypes(
                            Convert.ToInt32(((DataRowView)BatchManager.InsetColorsSummaryBindingSource.Current).Row["FrontID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.InsetColorsSummaryBindingSource.Current).Row["ColorID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.InsetColorsSummaryBindingSource.Current).Row["PatinaID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.InsetColorsSummaryBindingSource.Current).Row["TechnoColorID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.InsetColorsSummaryBindingSource.Current).Row["InsetTypeID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.InsetColorsSummaryBindingSource.Current).Row["InsetColorID"]));
                    }
                }
        }

        private void TechnoInsetTypesDataGrid_SelectionChanged(object sender, EventArgs e)
        {
            if (BatchManager != null)
                if (BatchManager.TechnoInsetTypesSummaryBindingSource.Count > 0)
                {
                    if (((DataRowView)BatchManager.TechnoInsetTypesSummaryBindingSource.Current)["TechnoInsetTypeID"] != DBNull.Value)
                    {
                        BatchManager.FilterTechnoInsetColors(
                            Convert.ToInt32(((DataRowView)BatchManager.TechnoInsetTypesSummaryBindingSource.Current).Row["FrontID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.TechnoInsetTypesSummaryBindingSource.Current).Row["ColorID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.TechnoInsetTypesSummaryBindingSource.Current).Row["PatinaID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.TechnoInsetTypesSummaryBindingSource.Current).Row["TechnoColorID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.TechnoInsetTypesSummaryBindingSource.Current).Row["InsetTypeID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.TechnoInsetTypesSummaryBindingSource.Current).Row["InsetColorID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.TechnoInsetTypesSummaryBindingSource.Current).Row["TechnoInsetTypeID"]));
                    }
                }
        }

        private void TechnoInsetColorsDataGrid_SelectionChanged(object sender, EventArgs e)
        {
            if (BatchManager != null)
                if (BatchManager.TechnoInsetColorsSummaryBindingSource.Count > 0)
                {
                    if (((DataRowView)BatchManager.TechnoInsetColorsSummaryBindingSource.Current)["TechnoInsetColorID"] != DBNull.Value)
                    {
                        BatchManager.FilterSizes(
                            Convert.ToInt32(((DataRowView)BatchManager.TechnoInsetColorsSummaryBindingSource.Current).Row["FrontID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.TechnoInsetColorsSummaryBindingSource.Current).Row["ColorID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.TechnoInsetColorsSummaryBindingSource.Current).Row["PatinaID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.TechnoInsetColorsSummaryBindingSource.Current).Row["TechnoColorID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.TechnoInsetColorsSummaryBindingSource.Current).Row["InsetTypeID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.TechnoInsetColorsSummaryBindingSource.Current).Row["InsetColorID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.TechnoInsetColorsSummaryBindingSource.Current).Row["TechnoInsetTypeID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.TechnoInsetColorsSummaryBindingSource.Current).Row["TechnoInsetColorID"]));
                    }
                }
        }

        private void DecorItemsDataGrid_SelectionChanged(object sender, EventArgs e)
        {
            if (BatchManager != null)
                if (BatchManager.DecorItemsSummaryBindingSource.Count > 0)
                {
                    if (((DataRowView)BatchManager.DecorItemsSummaryBindingSource.Current)["ProductID"] != DBNull.Value)
                    {
                        BatchManager.FilterDecorItems(
                            Convert.ToInt32(((DataRowView)BatchManager.DecorItemsSummaryBindingSource.Current).Row["ProductID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.DecorItemsSummaryBindingSource.Current).Row["DecorID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.DecorItemsSummaryBindingSource.Current).Row["MeasureID"]));
                    }
                }
        }

        private void DecorColorsDataGrid_SelectionChanged(object sender, EventArgs e)
        {
            if (BatchManager != null)
                if (BatchManager.DecorColorsSummaryBindingSource.Count > 0)
                {
                    if (((DataRowView)BatchManager.DecorColorsSummaryBindingSource.Current)["ProductID"] != DBNull.Value)
                    {
                        BatchManager.FilterDecorSizes(
                            Convert.ToInt32(((DataRowView)BatchManager.DecorColorsSummaryBindingSource.Current).Row["ProductID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.DecorColorsSummaryBindingSource.Current).Row["DecorID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.DecorColorsSummaryBindingSource.Current).Row["ColorID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.DecorColorsSummaryBindingSource.Current).Row["MeasureID"]));
                    }
                }
        }

        private void PreFrontsDataGrid_SelectionChanged(object sender, EventArgs e)
        {
            if (BatchManager != null)
                if (BatchManager.PreFrontsSummaryBindingSource.Count > 0)
                {
                    if (((DataRowView)BatchManager.PreFrontsSummaryBindingSource.Current)["FrontID"] != DBNull.Value)
                    {
                        BatchManager.FilterPreFrameColors(
                            Convert.ToInt32(((DataRowView)BatchManager.PreFrontsSummaryBindingSource.Current).Row["FrontID"]));
                    }
                }

            PreFrontsSquareLabel.Text = "0";
            PreFrontsCountLabel.Text = "0";

            if (BatchManager != null)
                if (BatchManager.PreFrontsSummaryBindingSource.Count > 0)
                {
                    decimal frontSquare = 0;
                    var count = 0;

                    BatchManager.GetPreFrontsInfo(ref frontSquare, ref count);

                    PreFrontsSquareLabel.Text = frontSquare.ToString();
                    PreFrontsCountLabel.Text = count.ToString();
                }
        }

        private void PreFrameColorsDataGrid_SelectionChanged(object sender, EventArgs e)
        {
            if (BatchManager != null)
                if (BatchManager.PreFrameColorsSummaryBindingSource.Count > 0)
                {
                    if (((DataRowView)BatchManager.PreFrameColorsSummaryBindingSource.Current)["ColorID"] != DBNull.Value)
                    {
                        BatchManager.FilterPreTechnoColors(
                            Convert.ToInt32(((DataRowView)BatchManager.PreFrameColorsSummaryBindingSource.Current).Row["FrontID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.PreFrameColorsSummaryBindingSource.Current).Row["ColorID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.PreFrameColorsSummaryBindingSource.Current).Row["PatinaID"]));
                    }
                }
        }

        private void PreTechnoColorsDataGrid_SelectionChanged(object sender, EventArgs e)
        {
            if (BatchManager != null)
                if (BatchManager.PreTechnoColorsSummaryBindingSource.Count > 0)
                {
                    if (((DataRowView)BatchManager.PreTechnoColorsSummaryBindingSource.Current)["TechnoColorID"] != DBNull.Value)
                    {
                        BatchManager.FilterPreInsetTypes(
                            Convert.ToInt32(((DataRowView)BatchManager.PreTechnoColorsSummaryBindingSource.Current).Row["FrontID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.PreTechnoColorsSummaryBindingSource.Current).Row["ColorID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.PreTechnoColorsSummaryBindingSource.Current).Row["PatinaID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.PreTechnoColorsSummaryBindingSource.Current).Row["TechnoColorID"]));
                    }
                }
        }

        private void PreInsetTypesDataGrid_SelectionChanged(object sender, EventArgs e)
        {
            if (BatchManager != null)
                if (BatchManager.PreInsetTypesSummaryBindingSource.Count > 0)
                {
                    if (((DataRowView)BatchManager.PreInsetTypesSummaryBindingSource.Current)["InsetTypeID"] != DBNull.Value)
                    {
                        BatchManager.FilterPreInsetColors(
                            Convert.ToInt32(((DataRowView)BatchManager.PreInsetTypesSummaryBindingSource.Current).Row["FrontID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.PreInsetTypesSummaryBindingSource.Current).Row["ColorID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.PreInsetTypesSummaryBindingSource.Current).Row["PatinaID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.PreInsetTypesSummaryBindingSource.Current).Row["TechnoColorID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.PreInsetTypesSummaryBindingSource.Current).Row["InsetTypeID"]));
                    }
                }
        }

        private void PreInsetColorsDataGrid_SelectionChanged(object sender, EventArgs e)
        {
            if (BatchManager != null)
                if (BatchManager.PreInsetColorsSummaryBindingSource.Count > 0)
                {
                    if (((DataRowView)BatchManager.PreInsetColorsSummaryBindingSource.Current)["InsetColorID"] != DBNull.Value)
                    {
                        BatchManager.FilterPreTechnoInsetTypes(
                            Convert.ToInt32(((DataRowView)BatchManager.PreInsetColorsSummaryBindingSource.Current).Row["FrontID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.PreInsetColorsSummaryBindingSource.Current).Row["ColorID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.PreInsetColorsSummaryBindingSource.Current).Row["PatinaID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.PreInsetColorsSummaryBindingSource.Current).Row["TechnoColorID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.PreInsetColorsSummaryBindingSource.Current).Row["InsetTypeID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.PreInsetColorsSummaryBindingSource.Current).Row["InsetColorID"]));
                    }
                }
        }

        private void PreTechnoInsetTypesDataGrid_SelectionChanged(object sender, EventArgs e)
        {
            if (BatchManager != null)
                if (BatchManager.PreTechnoInsetTypesSummaryBindingSource.Count > 0)
                {
                    if (((DataRowView)BatchManager.PreTechnoInsetTypesSummaryBindingSource.Current)["TechnoInsetTypeID"] != DBNull.Value)
                    {
                        BatchManager.FilterPreTechnoInsetColors(
                            Convert.ToInt32(((DataRowView)BatchManager.PreTechnoInsetTypesSummaryBindingSource.Current).Row["FrontID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.PreTechnoInsetTypesSummaryBindingSource.Current).Row["ColorID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.PreTechnoInsetTypesSummaryBindingSource.Current).Row["PatinaID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.PreTechnoInsetTypesSummaryBindingSource.Current).Row["TechnoColorID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.PreTechnoInsetTypesSummaryBindingSource.Current).Row["InsetTypeID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.PreTechnoInsetTypesSummaryBindingSource.Current).Row["InsetColorID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.PreTechnoInsetTypesSummaryBindingSource.Current).Row["TechnoInsetTypeID"]));
                    }
                }
        }

        private void PreTechnoInsetColorsDataGrid_SelectionChanged(object sender, EventArgs e)
        {
            if (BatchManager != null)
                if (BatchManager.PreTechnoInsetColorsSummaryBindingSource.Count > 0)
                {
                    if (((DataRowView)BatchManager.PreTechnoInsetColorsSummaryBindingSource.Current)["TechnoInsetColorID"] != DBNull.Value)
                    {
                        BatchManager.FilterPreSizes(
                            Convert.ToInt32(((DataRowView)BatchManager.PreTechnoInsetColorsSummaryBindingSource.Current).Row["FrontID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.PreTechnoInsetColorsSummaryBindingSource.Current).Row["ColorID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.PreTechnoInsetColorsSummaryBindingSource.Current).Row["PatinaID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.PreTechnoInsetColorsSummaryBindingSource.Current).Row["TechnoColorID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.PreTechnoInsetColorsSummaryBindingSource.Current).Row["InsetTypeID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.PreTechnoInsetColorsSummaryBindingSource.Current).Row["InsetColorID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.PreTechnoInsetColorsSummaryBindingSource.Current).Row["TechnoInsetTypeID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.PreTechnoInsetColorsSummaryBindingSource.Current).Row["TechnoInsetColorID"]));
                    }
                }
        }

        private void PreDecorProductsDataGrid_SelectionChanged(object sender, EventArgs e)
        {
            if (BatchManager != null)
                if (BatchManager.PreDecorProductsSummaryBindingSource.Count > 0)
                {
                    if (((DataRowView)BatchManager.PreDecorProductsSummaryBindingSource.Current)["ProductID"] != DBNull.Value)
                    {
                        BatchManager.FilterPreDecorProducts(
                            Convert.ToInt32(((DataRowView)BatchManager.PreDecorProductsSummaryBindingSource.Current).Row["ProductID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.PreDecorProductsSummaryBindingSource.Current).Row["MeasureID"]));
                    }
                }

            PreDecorPogonLabel.Text = "0";
            PreDecorCountLabel.Text = "0";

            if (BatchManager != null)
                if (BatchManager.PreDecorProductsSummaryBindingSource.Count > 0)
                {
                    decimal pogon = 0;
                    var count = 0;

                    BatchManager.GetPreDecorInfo(ref pogon, ref count);

                    PreDecorPogonLabel.Text = pogon.ToString();
                    PreDecorCountLabel.Text = count.ToString();
                }
        }

        private void PreDecorItemsDataGrid_SelectionChanged(object sender, EventArgs e)
        {
            if (BatchManager != null)
                if (BatchManager.PreDecorItemsSummaryBindingSource.Count > 0)
                {
                    if (((DataRowView)BatchManager.PreDecorItemsSummaryBindingSource.Current)["ProductID"] != DBNull.Value)
                    {
                        BatchManager.FilterPreDecorItems(
                            Convert.ToInt32(((DataRowView)BatchManager.PreDecorItemsSummaryBindingSource.Current).Row["ProductID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.PreDecorItemsSummaryBindingSource.Current).Row["DecorID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.PreDecorItemsSummaryBindingSource.Current).Row["MeasureID"]));
                    }
                }
        }

        private void PreDecorColorsDataGrid_SelectionChanged(object sender, EventArgs e)
        {
            if (BatchManager != null)
                if (BatchManager.PreDecorColorsSummaryBindingSource.Count > 0)
                {
                    if (((DataRowView)BatchManager.PreDecorColorsSummaryBindingSource.Current)["ProductID"] != DBNull.Value)
                    {
                        BatchManager.FilterPreDecorSizes(
                            Convert.ToInt32(((DataRowView)BatchManager.PreDecorColorsSummaryBindingSource.Current).Row["ProductID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.PreDecorColorsSummaryBindingSource.Current).Row["DecorID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.PreDecorColorsSummaryBindingSource.Current).Row["ColorID"]),
                            Convert.ToInt32(((DataRowView)BatchManager.PreDecorColorsSummaryBindingSource.Current).Row["MeasureID"]));
                    }
                }
        }

        private void EditBatchNameButton_Click(object sender, EventArgs e)
        {
            if (BatchDataGrid.SelectedRows.Count == 1)
            {
                BatchManager.SaveBatch(_factoryId);
            }
        }

        private void BatchMainOrdersDataGrid_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            if (BatchMainOrdersDataGrid.Rows.Count > 1)
                return;

            var productionStatus = "ProfilProductionStatusID";

            if (_factoryId == 2)
                productionStatus = "TPSProductionStatusID";

            if (BatchMainOrdersDataGrid.SelectedCells.Count > 0
                && BatchMainOrdersDataGrid.SelectedCells[0].RowIndex == e.RowIndex
                && Convert.ToInt32(BatchMainOrdersDataGrid.Rows[e.RowIndex].Cells[productionStatus].Value) != 3)
            {
                BatchMainOrdersDataGrid.Rows[e.RowIndex].DefaultCellStyle.SelectionBackColor = Color.DimGray;
                BatchMainOrdersDataGrid.Rows[e.RowIndex].DefaultCellStyle.SelectionForeColor = Color.White;
            }
        }

        private void SaveMegaBatch_Click(object sender, EventArgs e)
        {
            if (MegaBatchDataGrid.SelectedRows.Count == 1)
            {
                BatchManager.SaveMegaBatch(_factoryId);
            }
        }

        private void MegaBatchDataGrid_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if ((_roleType == RoleTypes.Admin || _roleType == RoleTypes.Control || _roleType == RoleTypes.Direction)
                && e.Button == System.Windows.Forms.MouseButtons.Right)
            {
                kryptonContextMenu1.Show(new Point(Cursor.Position.X - 212, Cursor.Position.Y - 10));
            }

            if ((PermissionGranted(IAgreedProfil) || PermissionGranted(IAgreedTps))
                && e.Button == System.Windows.Forms.MouseButtons.Right)
            {
                kryptonContextMenu1.Show(new Point(Cursor.Position.X - 212, Cursor.Position.Y - 10));
            }
        }

        private void GroupsCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            MarketFilterCheckBox.Enabled = GroupsCheckBox.Checked;
            ZOVFilterCheckBox.Enabled = GroupsCheckBox.Checked;
            MoscowFilterCheckBox.Enabled = GroupsCheckBox.Checked;
            kryptonCheckBox1.Enabled = GroupsCheckBox.Checked;
            ZOVProfilCheckBox.Enabled = GroupsCheckBox.Checked;
            ZOVTPSCheckBox.Enabled = GroupsCheckBox.Checked;

            MegaBatchDataGrid_SelectionChanged(null, null);
        }

        private void MarketFilterCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            MegaBatchDataGrid_SelectionChanged(null, null);
        }

        private void ZOVFilterCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            MegaBatchDataGrid_SelectionChanged(null, null);
        }

        private void MoscowFilterCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            MegaBatchDataGrid_SelectionChanged(null, null);
        }

        private void FrontsDataGrid_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Right)
            {
                kryptonContextMenu2.Show(new Point(Cursor.Position.X - 212, Cursor.Position.Y - 10));
            }
        }

        private void cmiExportToExcel_Click(object sender, EventArgs e)
        {
            var T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref _topForm, "Создание отчета.\r\nПодождите..."); });
            T.Start();

            while (!SplashWindow.bSmallCreated) ;
            _needSplash = false;

            var marketingBatchReport = new BatchExcelReport();

            marketingBatchReport.CreateReport(((DataTable)((BindingSource)FrontsDataGrid.DataSource).DataSource).DefaultView.Table,
                ((DataTable)((BindingSource)DecorProductsDataGrid.DataSource).DataSource).DefaultView.Table, "Недельное планирование. Маркетинг");

            _needSplash = true;
            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
        }

        private void DecorProductsDataGrid_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Right)
            {
                kryptonContextMenu2.Show(new Point(Cursor.Position.X - 212, Cursor.Position.Y - 10));
            }
        }

        private void BatchDataGrid_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if ((_roleType == RoleTypes.Admin || _roleType == RoleTypes.Direction || _roleType == RoleTypes.Control || _roleType == RoleTypes.Production || PermissionGranted(IAgreedProfil) || PermissionGranted(IAgreedTps))
                && e.Button == System.Windows.Forms.MouseButtons.Right)
            {
                kryptonContextMenu3.Show(new Point(Cursor.Position.X - 212, Cursor.Position.Y - 10));
            }
        }

        private void BatchTPSAccess_Click(object sender, EventArgs e)
        {
            if (BatchManager.BatchTPSAccess)
            {
                BatchManager.BatchTPSConfirm = true;
                if (BatchManager.BatchTPSConfirmDateTime == DBNull.Value)
                    BatchManager.BatchTPSConfirmDateTime = Security.GetCurrentDate();
                if (BatchManager.BatchTPSConfirmUser == DBNull.Value)
                    BatchManager.BatchTPSConfirmUser = Security.CurrentUserID;
                BatchManager.BatchTPSAccess = false;
                BatchManager.BatchTPSCloseUser = Security.CurrentUserID;
                BatchManager.BatchTPSCloseDateTime = Security.GetCurrentDate();
            }
            else
            {
                BatchManager.BatchTPSAccess = true;
                BatchManager.BatchTPSCloseUser = DBNull.Value;
                BatchManager.BatchTPSCloseDateTime = DBNull.Value;
            }

            BatchManager.SaveBatch(_factoryId);
            BatchManager.SetMegaBatchEnable();
            BatchDataGrid.Refresh();
        }

        private void BatchProfilAccess_Click(object sender, EventArgs e)
        {
            if (BatchManager.BatchProfilAccess)
            {
                BatchManager.BatchProfilConfirm = true;
                if (BatchManager.BatchProfilConfirmDateTime == DBNull.Value)
                    BatchManager.BatchProfilConfirmDateTime = Security.GetCurrentDate();
                if (BatchManager.BatchProfilConfirmUser == DBNull.Value)
                    BatchManager.BatchProfilConfirmUser = Security.CurrentUserID;
                BatchManager.BatchProfilAccess = false;
                BatchManager.BatchProfilCloseUser = Security.CurrentUserID;
                BatchManager.BatchProfilCloseDateTime = Security.GetCurrentDate();
            }
            else
            {
                BatchManager.BatchProfilAccess = true;
                BatchManager.BatchProfilCloseUser = DBNull.Value;
                BatchManager.BatchProfilCloseDateTime = DBNull.Value;
            }

            BatchManager.SaveBatch(_factoryId);
            BatchManager.SetMegaBatchEnable();
            BatchDataGrid.Refresh();
        }

        private void BatchDataGrid_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex == -1)
                return;
            //if (e.ColumnIndex == 4)
            //{
            //    if (Convert.ToBoolean(BatchDataGrid.Rows[e.RowIndex].Cells["ProfilEnabled"].Value))
            //        e.Graphics.DrawImage(Unlock_BW, e.CellBounds.Left + ((e.CellBounds.Right - e.CellBounds.Left) / 2 - Unlock_BW.Width / 2) - 1, e.CellBounds.Top + ((e.CellBounds.Bottom - e.CellBounds.Top) / 2 - Unlock_BW.Height / 2) - 1, Unlock_BW.Width, Unlock_BW.Height);
            //    else
            //        e.Graphics.DrawImage(Lock_BW, e.CellBounds.Left + ((e.CellBounds.Right - e.CellBounds.Left) / 2 - Lock_BW.Width / 2) - 1, e.CellBounds.Top + ((e.CellBounds.Bottom - e.CellBounds.Top) / 2 - Lock_BW.Height / 2) - 1, Lock_BW.Width, Lock_BW.Height);
            //}
            //if (e.ColumnIndex == 5)
            //{
            //    if (Convert.ToBoolean(BatchDataGrid.Rows[e.RowIndex].Cells["TPSEnabled"].Value))
            //        e.Graphics.DrawImage(Unlock_BW, e.CellBounds.Left + ((e.CellBounds.Right - e.CellBounds.Left) / 2 - Unlock_BW.Width / 2) - 1, e.CellBounds.Top + ((e.CellBounds.Bottom - e.CellBounds.Top) / 2 - Unlock_BW.Height / 2) - 1, Unlock_BW.Width, Unlock_BW.Height);
            //    else
            //        e.Graphics.DrawImage(Lock_BW, e.CellBounds.Left + ((e.CellBounds.Right - e.CellBounds.Left) / 2 - Lock_BW.Width / 2) - 1, e.CellBounds.Top + ((e.CellBounds.Bottom - e.CellBounds.Top) / 2 - Lock_BW.Height / 2) - 1, Lock_BW.Width, Lock_BW.Height);
            //}
        }

        private void MegaBatchProfilAccess_Click(object sender, EventArgs e)
        {
            BatchManager.MegaBatchProfilAccess = !BatchManager.MegaBatchProfilAccess;
            BatchManager.SetBatchEnabled(Convert.ToInt32(((DataRowView)BatchManager.MegaBatchBindingSource.Current)["MegaBatchID"]),
                _factoryId, BatchManager.MegaBatchProfilAccess);
            BatchManager.CloseMegaBatch(Convert.ToInt32(((DataRowView)BatchManager.MegaBatchBindingSource.Current)["MegaBatchID"]), _factoryId, BatchManager.MegaBatchProfilAccess);
            MegaBatchDataGrid.Refresh();
            BatchDataGrid.Refresh();
            MegaBatchDataGrid_SelectionChanged(null, null);
        }

        private void MegaBatchDataGrid_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex == -1)
                return;
            //if (MegaBatchDataGrid.Columns[e.ColumnIndex].Name == "ProfilEnabled")
            //{
            //    if (Convert.ToBoolean(MegaBatchDataGrid.Rows[e.RowIndex].Cells["ProfilEnabled"].Value))
            //        e.Graphics.DrawImage(Unlock_BW, e.CellBounds.Left + ((e.CellBounds.Right - e.CellBounds.Left) / 2 - Unlock_BW.Width / 2) - 1, e.CellBounds.Top + ((e.CellBounds.Bottom - e.CellBounds.Top) / 2 - Unlock_BW.Height / 2) - 1, Unlock_BW.Width, Unlock_BW.Height);
            //    else
            //        e.Graphics.DrawImage(Lock_BW, e.CellBounds.Left + ((e.CellBounds.Right - e.CellBounds.Left) / 2 - Lock_BW.Width / 2) - 1, e.CellBounds.Top + ((e.CellBounds.Bottom - e.CellBounds.Top) / 2 - Lock_BW.Height / 2) - 1, Lock_BW.Width, Lock_BW.Height);
            //}
            //if (MegaBatchDataGrid.Columns[e.ColumnIndex].Name == "TPSEnabled")
            //{
            //    if (Convert.ToBoolean(MegaBatchDataGrid.Rows[e.RowIndex].Cells["TPSEnabled"].Value))
            //        e.Graphics.DrawImage(Unlock_BW, e.CellBounds.Left + ((e.CellBounds.Right - e.CellBounds.Left) / 2 - Unlock_BW.Width / 2) - 1, e.CellBounds.Top + ((e.CellBounds.Bottom - e.CellBounds.Top) / 2 - Unlock_BW.Height / 2) - 1, Unlock_BW.Width, Unlock_BW.Height);
            //    else
            //        e.Graphics.DrawImage(Lock_BW, e.CellBounds.Left + ((e.CellBounds.Right - e.CellBounds.Left) / 2 - Lock_BW.Width / 2) - 1, e.CellBounds.Top + ((e.CellBounds.Bottom - e.CellBounds.Top) / 2 - Lock_BW.Height / 2) - 1, Lock_BW.Width, Lock_BW.Height);
            //}
        }

        private void MegaBatchTPSAccess_Click(object sender, EventArgs e)
        {
            BatchManager.MegaBatchTPSAccess = !BatchManager.MegaBatchTPSAccess;
            BatchManager.SetBatchEnabled(Convert.ToInt32(((DataRowView)BatchManager.MegaBatchBindingSource.Current)["MegaBatchID"]),
                _factoryId, BatchManager.MegaBatchTPSAccess);
            BatchManager.CloseMegaBatch(Convert.ToInt32(((DataRowView)BatchManager.MegaBatchBindingSource.Current)["MegaBatchID"]), _factoryId, BatchManager.MegaBatchTPSAccess);
            MegaBatchDataGrid.Refresh();
            BatchDataGrid.Refresh();
            MegaBatchDataGrid_SelectionChanged(null, null);
        }

        private void MarketingBatchForm_Load(object sender, EventArgs e)
        {
            if (BatchManager == null)
                return;

            Filter();

            _addToBatch = false;
            _inProduction = false;
            _createBatch = false;
            _closeBatch = false;

            if (PermissionGranted(IAddToBatch))
            {
                _addToBatch = true;
            }
            if (PermissionGranted(IInProduction))
            {
                _inProduction = true;
            }
            if (PermissionGranted(ICloseBatch))
            {
                _closeBatch = true;
            }
            if (PermissionGranted(ICreateBatch))
            {
                _createBatch = true;
            }
            if (PermissionGranted(ICreateLabel))
            {
                _createLabel = true;
            }
            //if (PermissionGranted(iConfirmBatch))
            //{
            //    ConfirmBatch = true;
            //}

            if (PermissionGranted(IAgreedProfil))
            {
                btnAgreedProfil.Visible = true;
                cmiConfirmProfilBatch.Visible = true;
            }
            if (PermissionGranted(IAgreedTps))
            {
                cmiConfirmTPSBatch.Visible = true;
            }
            //if (PermissionGranted(iAgreed2TPS))
            //{
            //    btnAgreed2TPS.Visible = true;
            //}

            //Админ
            if (_addToBatch && _inProduction && _closeBatch && _createBatch && _createLabel)
            {
                _roleType = RoleTypes.Admin;

                cmiCloseProfilBatch.Visible = true;
                cmiCloseTPSBatch.Visible = true;
                cmiCloseProfilMegaBatch.Visible = true;
                cmiCloseTPSMegaBatch.Visible = true;
                cmiConfirmProfilBatch.Visible = true;
                cmiSaveMegaBatch.Visible = true;
                cmiCreateLabel.Visible = true;
            }
            //Управление: нельзя открывать/закрывать партию, остальное - можно
            if (_addToBatch && _inProduction && !_closeBatch && _createBatch)
            {
                _roleType = RoleTypes.Control;

                cmiCloseProfilBatch.Visible = false;
                cmiCloseTPSBatch.Visible = false;
                cmiCloseProfilMegaBatch.Visible = false;
                cmiCloseTPSMegaBatch.Visible = false;
                cmiConfirmProfilBatch.Visible = true;
                cmiSaveMegaBatch.Visible = true;
                cmiCreateLabel.Visible = true;
            }
            //Вход для фа: можно открывать и закрывать партию
            if (!_addToBatch && !_inProduction && _closeBatch && !_createBatch && !_createLabel)
            {
                _roleType = RoleTypes.Direction;

                cmiCloseProfilBatch.Visible = true;
                cmiCloseTPSBatch.Visible = true;
                cmiCloseProfilMegaBatch.Visible = true;
                cmiCloseTPSMegaBatch.Visible = true;
                cmiConfirmProfilBatch.Visible = false;
                cmiSaveMegaBatch.Visible = false;
                cmiCreateLabel.Visible = false;

                CreateBatchPanel.Visible = false;
                CreateMegaBatchPanel.Visible = false;
                MegaOrderInProductionPanel.Visible = false;
                MainOrderInProductionPanel.Visible = false;
                RolesAndPermissionsTabControl.TabPages[1].PageVisible = false;

                xtraTabControl1.Height = BatchPanel.Height - 5;
                xtraTabControl4.Height = MegaBatchPanel.Height - 5;
                BatchMegaOrdersDataGrid.Height = MegaOrdersPanel.Height - 5;
                BatchMainOrdersDataGrid.Height = MainOrdersPanel.Height - 5;
            }
            //вход для пр-ва: можно формировать партию, выдавать в пр-во, нельзя создавать партию и открывать/закрывать её
            if (_addToBatch && _inProduction && !_closeBatch && !_createBatch && !_createLabel)
            {
                _roleType = RoleTypes.Production;

                cmiCloseProfilBatch.Visible = false;
                cmiCloseTPSBatch.Visible = false;
                cmiCloseProfilMegaBatch.Visible = false;
                cmiCloseTPSMegaBatch.Visible = false;
                cmiConfirmProfilBatch.Visible = true;
                cmiSaveMegaBatch.Visible = false;
                cmiCreateLabel.Visible = false;

                MoveMegaOrdersButton.Visible = false;
                PickUpFrontsButton.Visible = false;
                CreateBatchPanel.Visible = false;
                CreateMegaBatchPanel.Visible = false;

                xtraTabControl1.Height = BatchPanel.Height - 5;
                xtraTabControl4.Height = MegaBatchPanel.Height - 5;
            }
            //Маркетинг
            if (_addToBatch && !_inProduction && !_closeBatch && !_createBatch && !_createLabel)
            {
                _roleType = RoleTypes.Marketing;

                cmiCloseProfilBatch.Visible = false;
                cmiCloseTPSBatch.Visible = false;
                cmiCloseProfilMegaBatch.Visible = false;
                cmiCloseTPSMegaBatch.Visible = false;
                cmiConfirmProfilBatch.Visible = false;
                cmiSaveMegaBatch.Visible = false;
                cmiCreateLabel.Visible = false;

                CreateBatchPanel.Visible = false;
                CreateMegaBatchPanel.Visible = false;
                MegaOrderInProductionPanel.Visible = false;
                MainOrderInProductionPanel.Visible = false;

                xtraTabControl1.Height = BatchPanel.Height - 5;
                xtraTabControl4.Height = MegaBatchPanel.Height - 5;
                BatchMegaOrdersDataGrid.Height = MegaOrdersPanel.Height - 5;
                BatchMainOrdersDataGrid.Height = MainOrdersPanel.Height - 5;
            }
            //вход для обычного пользователя: нельзя ничего сделать, только посмотреть партии и их содержимое
            if (!_addToBatch && !_inProduction && !_closeBatch && !_createBatch && !_createLabel)
            {
                _roleType = RoleTypes.Ordinary;

                cmiCloseProfilBatch.Visible = false;
                cmiCloseTPSBatch.Visible = false;
                cmiCloseProfilMegaBatch.Visible = false;
                cmiCloseTPSMegaBatch.Visible = false;
                cmiConfirmProfilBatch.Visible = false;
                cmiSaveMegaBatch.Visible = false;
                cmiCreateLabel.Visible = false;

                CreateBatchPanel.Visible = false;
                CreateMegaBatchPanel.Visible = false;
                MegaOrderInProductionPanel.Visible = false;
                MainOrderInProductionPanel.Visible = false;
                RolesAndPermissionsTabControl.TabPages[1].PageVisible = false;

                xtraTabControl1.Height = BatchPanel.Height - 5;
                xtraTabControl4.Height = MegaBatchPanel.Height - 5;
                BatchMegaOrdersDataGrid.Height = MegaOrdersPanel.Height - 5;
                BatchMainOrdersDataGrid.Height = MainOrdersPanel.Height - 5;
            }
            //Склад
            if (!_addToBatch && !_inProduction && !_closeBatch && !_createBatch && _createLabel)
            {
                _roleType = RoleTypes.Store;

                cmiCloseProfilBatch.Visible = false;
                cmiCloseTPSBatch.Visible = false;
                cmiCloseProfilMegaBatch.Visible = false;
                cmiCloseTPSMegaBatch.Visible = false;
                cmiConfirmProfilBatch.Visible = false;
                cmiSaveMegaBatch.Visible = false;
                cmiCreateLabel.Visible = true;

                CreateBatchPanel.Visible = false;
                CreateMegaBatchPanel.Visible = false;
                MegaOrderInProductionPanel.Visible = false;
                MainOrderInProductionPanel.Visible = false;
                RolesAndPermissionsTabControl.TabPages[1].PageVisible = false;

                xtraTabControl1.Height = BatchPanel.Height - 5;
                xtraTabControl4.Height = MegaBatchPanel.Height - 5;
                BatchMegaOrdersDataGrid.Height = MegaOrdersPanel.Height - 5;
                BatchMainOrdersDataGrid.Height = MainOrdersPanel.Height - 5;
            }

            cmiCloseTPSBatch.Visible = false;
            cmiCloseTPSMegaBatch.Visible = false;
            cmiConfirmTPSBatch.Visible = false;

            BatchDataGrid.Columns["TPSConfirm"].Visible = false;
            BatchDataGrid.Columns["TPSBatchClose"].Visible = false;
            MegaBatchDataGrid.Columns["TPSBatchClose"].Visible = false;
            BatchDataGrid.Columns["TPSName"].Visible = false;


        }

        private void btnAgreed1TPS_Click(object sender, EventArgs e)
        {
            var megaBatchId = 0;
            var tpsAgreedUserId = -1;
            var tpsCloseUserId = -1;
            if (MegaBatchDataGrid.SelectedRows.Count != 0 && MegaBatchDataGrid.SelectedRows[0].Cells["TPSCloseUserID"].Value != DBNull.Value)
                tpsCloseUserId = Convert.ToInt32(MegaBatchDataGrid.SelectedRows[0].Cells["TPSCloseUserID"].Value);
            if (MegaBatchDataGrid.SelectedRows.Count != 0 && MegaBatchDataGrid.SelectedRows[0].Cells["MegaBatchID"].Value != DBNull.Value)
                megaBatchId = Convert.ToInt32(MegaBatchDataGrid.SelectedRows[0].Cells["MegaBatchID"].Value);
            if (MegaBatchDataGrid.SelectedRows.Count != 0 && MegaBatchDataGrid.SelectedRows[0].Cells["TPSAgreedUserID"].Value != DBNull.Value)
                tpsAgreedUserId = Convert.ToInt32(MegaBatchDataGrid.SelectedRows[0].Cells["TPSAgreedUserID"].Value);
            if (tpsCloseUserId != -1)
            {
                LightMessageBox.Show(ref _topForm, false,
                        "Партия утверждена и закрыта",
                        "Утверждение партии");
                return;
            }
            if (tpsAgreedUserId != -1)
            {
                var okCancel = LightMessageBox.Show(ref _topForm, true,
                        "Ваше утверждение уже выставлено. Переутвердить?",
                        "Утверждение партии");
                if (!okCancel)
                    return;
            }
            _needSplash = false;
            var T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref _topForm, "Загрузка данных.\r\nПодождите..."); });
            T.Start();

            while (!SplashWindow.bSmallCreated) ;
            BatchManager.SetMegaBatchAgreement(megaBatchId, _factoryId);
            BatchManager.SaveMegaBatch(_factoryId);
            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
            _needSplash = true;
        }

        private void btnAgreed1Profil_Click(object sender, EventArgs e)
        {
            var megaBatchId = 0;
            var profilAgreedUserId = -1;
            var profilCloseUserId = -1;
            if (MegaBatchDataGrid.SelectedRows.Count != 0 && MegaBatchDataGrid.SelectedRows[0].Cells["ProfilCloseUserID"].Value != DBNull.Value)
                profilCloseUserId = Convert.ToInt32(MegaBatchDataGrid.SelectedRows[0].Cells["ProfilCloseUserID"].Value);
            if (MegaBatchDataGrid.SelectedRows.Count != 0 && MegaBatchDataGrid.SelectedRows[0].Cells["MegaBatchID"].Value != DBNull.Value)
                megaBatchId = Convert.ToInt32(MegaBatchDataGrid.SelectedRows[0].Cells["MegaBatchID"].Value);
            if (MegaBatchDataGrid.SelectedRows.Count != 0 && MegaBatchDataGrid.SelectedRows[0].Cells["ProfilAgreedUserID"].Value != DBNull.Value)
                profilAgreedUserId = Convert.ToInt32(MegaBatchDataGrid.SelectedRows[0].Cells["ProfilAgreedUserID"].Value);
            if (profilCloseUserId != -1)
            {
                LightMessageBox.Show(ref _topForm, false,
                        "Партия утверждена и закрыта",
                        "Утверждение партии");
                return;
            }
            if (profilAgreedUserId != -1)
            {
                var okCancel = LightMessageBox.Show(ref _topForm, true,
                        "Ваше утверждение уже выставлено. Переутвердить?",
                        "Утверждение партии");
                if (!okCancel)
                    return;
            }
            _needSplash = false;
            var T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref _topForm, "Загрузка данных.\r\nПодождите..."); });
            T.Start();

            while (!SplashWindow.bSmallCreated) ;
            BatchManager.SetMegaBatchAgreement(megaBatchId, _factoryId);
            BatchManager.SaveMegaBatch(_factoryId);
            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
            _needSplash = true;
        }

        private void cmiCreateLabel_Click(object sender, EventArgs e)
        {
            //PhantomForm PhantomForm = new Infinium.PhantomForm();
            //PhantomForm.Show();

            //MarketingPalletsForm MarketingPalletsForm = new MarketingPalletsForm(
            //    Convert.ToInt32(((DataRowView)BatchManager.BatchMegaOrdersBindingSource.Current)["MegaOrderID"]), FactoryID,
            //    Convert.ToInt32(((DataRowView)BatchManager.MegaBatchBindingSource.Current)["MegaBatchID"]));

            //TopForm = MarketingPalletsForm;
            //MarketingPalletsForm.ShowDialog();

            //PhantomForm.Close();

            //PhantomForm.Dispose();
            //MarketingPalletsForm.Dispose();
            //TopForm = null;
        }

        private void BatchMegaOrdersDataGrid_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if ((_roleType == RoleTypes.Admin || _roleType == RoleTypes.Store || _roleType == RoleTypes.Control) && e.Button == System.Windows.Forms.MouseButtons.Right)
            {
                kryptonContextMenu4.Show(new Point(Cursor.Position.X - 212, Cursor.Position.Y - 10));
            }
        }

        private void MinimizeButton_Click(object sender, EventArgs e)
        {
            _needSplash = false;
            _formEvent = EHide;
            AnimateTimer.Enabled = true;
        }

        private void cmiConfirmBatch_Click(object sender, EventArgs e)
        {
            if (!BatchManager.BatchProfilConfirm)
            {
                BatchManager.BatchProfilConfirmUser = Security.CurrentUserID;
                BatchManager.BatchProfilConfirmDateTime = Security.GetCurrentDate();
                BatchManager.BatchProfilConfirm = true;
                BatchManager.SaveBatch(_factoryId);
            }
        }

        private void kryptonContextMenu3_Opening(object sender, CancelEventArgs e)
        {
            if (_factoryId == 1)
            {
                if (_roleType == RoleTypes.Admin || _roleType == RoleTypes.Direction)
                {
                    cmiCloseProfilBatch.Visible = true;
                    cmiCloseTPSBatch.Visible = false;
                }
                if (_roleType == RoleTypes.Admin || PermissionGranted(IAgreedProfil))
                {
                    cmiConfirmProfilBatch.Visible = true;
                    cmiConfirmTPSBatch.Visible = false;
                    if (BatchManager.BatchProfilConfirm)
                        cmiConfirmProfilBatch.Visible = false;
                }
                if (BatchManager.BatchProfilAccess)
                    cmiCloseProfilBatch.Text = "Закрыть партию";
                else
                    cmiCloseProfilBatch.Text = "Открыть партию";
            }
            if (_factoryId == 2)
            {
                if (_roleType == RoleTypes.Admin || _roleType == RoleTypes.Direction)
                {
                    cmiCloseProfilBatch.Visible = false;
                    cmiCloseTPSBatch.Visible = true;
                }
                if (_roleType == RoleTypes.Admin || PermissionGranted(IAgreedTps))
                {
                    cmiConfirmTPSBatch.Visible = true;
                    cmiConfirmProfilBatch.Visible = false;
                    if (BatchManager.BatchTPSConfirm)
                        cmiConfirmTPSBatch.Visible = false;
                }
                if (BatchManager.BatchTPSAccess)
                    cmiCloseTPSBatch.Text = "Закрыть партию";
                else
                    cmiCloseTPSBatch.Text = "Открыть партию";
            }
        }

        private void cmiConfirmTPSBatch_Click(object sender, EventArgs e)
        {
            if (!BatchManager.BatchTPSConfirm)
            {
                BatchManager.BatchTPSConfirmUser = Security.CurrentUserID;
                BatchManager.BatchTPSConfirmDateTime = Security.GetCurrentDate();
                BatchManager.BatchTPSConfirm = true;
                BatchManager.SaveBatch(_factoryId);
            }
        }

        private void kryptonContextMenu1_Opening(object sender, CancelEventArgs e)
        {
            if (_factoryId == 1)
            {
                if (_roleType == RoleTypes.Admin || _roleType == RoleTypes.Direction)
                {
                    cmiCloseProfilMegaBatch.Visible = true;
                    cmiCloseTPSMegaBatch.Visible = false;
                }
                if (MegaBatchDataGrid.SelectedRows[0].Cells["ProfilBatchClose"].Value != null)
                {
                    if (Convert.ToBoolean(MegaBatchDataGrid.SelectedRows[0].Cells["ProfilBatchClose"].Value))
                    {
                        btnAgreedProfil.Visible = false;
                        cmiCloseProfilMegaBatch.Text = "Открыть группу";
                    }
                    else
                    {
                        btnAgreedProfil.Visible = true;
                        cmiCloseProfilMegaBatch.Text = "Закрыть группу";
                    }
                }
            }
            if (_factoryId == 2)
            {
                if (_roleType == RoleTypes.Admin || _roleType == RoleTypes.Direction)
                {
                    cmiCloseProfilMegaBatch.Visible = false;
                    cmiCloseTPSMegaBatch.Visible = true;
                }
                if (MegaBatchDataGrid.SelectedRows[0].Cells["TPSBatchClose"].Value != null)
                {
                    if (Convert.ToBoolean(MegaBatchDataGrid.SelectedRows[0].Cells["TPSBatchClose"].Value))
                    {
                        btnAgreedTPS.Visible = false;
                        cmiCloseTPSMegaBatch.Text = "Открыть группу";
                    }
                    else
                    {
                        btnAgreedTPS.Visible = true;
                        cmiCloseTPSMegaBatch.Text = "Закрыть группу";
                    }
                }
            }
        }

        private void MoveMainOrdersButton_Click(object sender, EventArgs e)
        {
            //if (!BatchManager.IsBatchEnabled(Convert.ToInt32(((DataRowView)BatchManager.BatchBindingSource.Current)["BatchID"]), FactoryID))
            //{
            //    Infinium.LightMessageBox.Show(ref TopForm, false,
            //           "Партия сформирована, перенос заказов запрещен",
            //           "Ошибка добавления");
            //    return;
            //}

            var phantomForm = new Infinium.PhantomForm();
            phantomForm.Show();

            var batchOnProduction = BatchOnProductionCheckBox.Checked;
            var batchNotProduction = BatchNotProductionCheckBox.Checked;
            var batchInProduction = BatchInProductionCheckBox.Checked;
            var batchOnStorage = BatchOnStorageCheckBox.Checked;
            var batchOnExp = BatchOnExpCheckBox.Checked;
            var batchDispatched = BatchDispatchCheckBox.Checked;

            _mainOrdersArray.Clear();
            for (var i = 0; i < BatchMainOrdersDataGrid.SelectedRows.Count; i++)
                _mainOrdersArray.Add(Convert.ToInt32(BatchMainOrdersDataGrid.SelectedRows[i].Cells["MainOrderID"].Value));

            if (BatchMainOrdersDataGrid.SelectedRows.Count == 0)
            {
                Infinium.LightMessageBox.Show(ref _topForm, false,
                          "Не выбран ни один подзаказ",
                          "Ошибка переноса");

                phantomForm.Close();

                phantomForm.Dispose();

                return;
            }

            var selectMenuForm = new MarketingBatchSelectMenu(this, BatchManager, _factoryId);

            _topForm = selectMenuForm;
            selectMenuForm.ShowDialog();

            phantomForm.Close();

            phantomForm.Dispose();
            selectMenuForm.Dispose();
            _topForm = null;

            if (BatchManager.NewBatch)
            {
                //int[] MainOrders = new int[BatchMainOrdersDataGrid.SelectedRows.Count];
                //for (int i = 0; i < BatchMainOrdersDataGrid.SelectedRows.Count; i++)
                //    MainOrders[i] = Convert.ToInt32(BatchMainOrdersDataGrid.SelectedRows[i].Cells["MainOrderID"].Value);
                BatchManager.MoveOrdersToNewPart(_mainOrdersArray.OfType<int>().ToArray(), _factoryId);

                BatchManager.FilterBatch(Convert.ToInt32(((DataRowView)BatchManager.MegaBatchBindingSource.Current)["MegaBatchID"]), _factoryId);
                Filter();
            }

            if (BatchManager.OldBatch)
            {
                MegaBatchDataGrid.Enabled = true;
                BatchDataGrid.Enabled = true;

                AddBatchButton.Visible = false;
                RemoveBatchButton.Visible = false;
                PrintBatchButton.Visible = false;
                MoveBatchButton.Visible = false;
                EditBatchNameButton.Visible = false;
                OKMoveOrdersButton.Visible = true;
                CancelMoveOrdersButton.Visible = true;
            }
        }

        private void MegaBatchDataGrid_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex == -1)
                return;

            var grid = (PercentageDataGrid)sender;
            if (grid.Columns[e.ColumnIndex].Name == "ProfilEnabledColumn")
            {
                if (grid.Rows[e.RowIndex].Cells["ProfilBatchClose"].Value != null)
                {
                    if (Convert.ToBoolean(grid.Rows[e.RowIndex].Cells["ProfilBatchClose"].Value))
                    {
                        e.Value = _imageList1.Images[0];
                    }
                    else
                    {
                        e.Value = _imageList1.Images[1];
                    }
                }
            }
            if (grid.Columns[e.ColumnIndex].Name == "TPSEnabledColumn")
            {
                if (grid.Rows[e.RowIndex].Cells["TPSBatchClose"].Value != null)
                {
                    if (Convert.ToBoolean(grid.Rows[e.RowIndex].Cells["TPSBatchClose"].Value))
                    {
                        e.Value = _imageList1.Images[0];
                    }
                    else
                    {
                        e.Value = _imageList1.Images[1];
                    }
                }
            }

            if (e.Value != null)
            {
                var cell =
                    grid.Rows[e.RowIndex].Cells[e.ColumnIndex];
                int megaBatchId;
                var displayName = string.Empty;
                if (grid.Rows[e.RowIndex].Cells["MegaBatchID"].Value != DBNull.Value)
                {
                    megaBatchId = Convert.ToInt32(grid.Rows[e.RowIndex].Cells["MegaBatchID"].Value);
                    displayName = BatchManager.GetMegaBatchAgreement(megaBatchId);
                }
                cell.ToolTipText = displayName;
            }
        }

        private void BatchDataGrid_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex == -1)
                return;

            var grid = (PercentageDataGrid)sender;
            if (grid.Columns[e.ColumnIndex].Name == "ProfilEnabledColumn")
            {
                if (grid.Rows[e.RowIndex].Cells["ProfilBatchClose"].Value != null)
                {
                    if (Convert.ToBoolean(grid.Rows[e.RowIndex].Cells["ProfilBatchClose"].Value))
                    {
                        e.Value = _imageList1.Images[0];
                    }
                    else
                    {
                        e.Value = _imageList1.Images[1];
                    }
                }
            }
            if (grid.Columns[e.ColumnIndex].Name == "TPSEnabledColumn")
            {
                if (grid.Rows[e.RowIndex].Cells["TPSBatchClose"].Value != null)
                {
                    if (Convert.ToBoolean(grid.Rows[e.RowIndex].Cells["TPSBatchClose"].Value))
                    {
                        e.Value = _imageList1.Images[0];
                    }
                    else
                    {
                        e.Value = _imageList1.Images[1];
                    }
                }
            }
        }

        private void kryptonCheckBox1_CheckedChanged(object sender, EventArgs e)
        {
            MegaBatchDataGrid_SelectionChanged(null, null);
        }

        private void kryptonContextMenuItem1_Click(object sender, EventArgs e)
        {
            var T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref _topForm, "Создание отчета.\r\nПодождите..."); });
            T.Start();

            while (!SplashWindow.bSmallCreated) ;
            _needSplash = false;

            var marketingBatchReport = new BatchReport(ref DecorCatalogOrder);
            var mainOrders = BatchManager.GetBatchMainOrders().OfType<int>().ToArray();
            marketingBatchReport.CreateReportForMaketing(mainOrders);

            _needSplash = true;
            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
        }

        private void btnMegaBatchToExcel_Click(object sender, EventArgs e)
        {
            var T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref _topForm, "Создание отчета.\r\nПодождите..."); });
            T.Start();

            while (!SplashWindow.bSmallCreated) ;
            _needSplash = false;

            var megaBatchId = Convert.ToInt32(MegaBatchDataGrid.SelectedRows[0].Cells["MegaBatchID"].Value);

            decimal eurbyrCurrency = 1000000;

            var fileName = $"Отчет по группе партий №{megaBatchId} от {DateTime.Today.ToString("dd.MM.yyyy")}";

            var batches = new ArrayList();

            for (var i = 0; i < BatchDataGrid.Rows.Count; i++)
                batches.Add(Convert.ToInt32(BatchDataGrid.Rows[i].Cells["BatchID"].Value));

            var producedProducts = new BatchAccountingReport(ref DecorCatalogOrder);
            producedProducts.GetDateRates(DateTime.Today, ref eurbyrCurrency);

            producedProducts.CreateMarketingOnProdReport(fileName, eurbyrCurrency, batches, megaBatchId);

            _needSplash = true;
            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
        }
        
        private void btnBatchToExcel_Click(object sender, EventArgs e)
        {
            var T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref _topForm, "Создание отчета.\r\nПодождите..."); });
            T.Start();

            while (!SplashWindow.bSmallCreated) ;
            _needSplash = false;

            var megaBatchId = Convert.ToInt32(MegaBatchDataGrid.SelectedRows[0].Cells["MegaBatchID"].Value);

            decimal eurbyrCurrency = 1000000;

            var fileName = $"Отчет по группе партий №{megaBatchId} от {DateTime.Today.ToString("dd.MM.yyyy")}";

            var batches = new ArrayList();

            for (var i = 0; i < BatchDataGrid.Rows.Count; i++)
                batches.Add(Convert.ToInt32(BatchDataGrid.Rows[i].Cells["BatchID"].Value));

            var producedProducts = new BatchAccountingReport(ref DecorCatalogOrder);
            producedProducts.GetDateRates(DateTime.Today, ref eurbyrCurrency);

            producedProducts.CreateMarketingOnProdReport(fileName, eurbyrCurrency, batches, megaBatchId);

            _needSplash = true;
            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
        }
    }
}
