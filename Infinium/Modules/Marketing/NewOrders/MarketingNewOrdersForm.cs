using Infinium.Modules.Marketing.NewOrders;
using Infinium.Modules.Marketing.WeeklyPlanning;

using NPOI.HSSF.UserModel;

using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using Infinium.Modules.Marketing.NewOrders.ColorInvoiceReportToDbf;
using Infinium.Modules.Marketing.NewOrders.InvoiceReportToDbf;
using Infinium.Modules.Marketing.NewOrders.NotesInvoiceReportToDbf;
using ComponentFactory.Krypton.Toolkit;
using FrontsOrders = Infinium.Modules.Marketing.Orders.FrontsOrders;

namespace Infinium
{
    public partial class MarketingNewOrdersForm : Form
    {
        private const int iAdmin = 74;
        private const int iMarketing = 20;
        private const int iDirector = 75;

        private const int eHide = 2;
        private const int eShow = 1;
        private const int eClose = 3;
        private const int eMainMenu = 4;

        private bool NeedRefresh = false;
        private bool NeedSplash = false;

        private int CurrentRowIndex = -1;
        private int FormEvent = 0;

        private LightStartForm LightStartForm;

        private Form TopForm = null;
        private AddMarketingNewOrdersForm AddMainOrdersForm;
        private CurrencyForm CurrencyForm;
        private SaveDBFReportMenu SaveDBFReportMenu;
        private ClientReportMenu ClientReportMenu;
        private Report Report;
        private DetailsReport DetailsReport;
        private Modules.Marketing.NewOrders.PrepareReport.DetailsReport PrepareReport;
        private SendEmail SendEmail;
        private Infinium.Modules.TechnologyCatalog.TestTechCatalog TestTechCatalogManager;
        private CreateOrdersFromExcel CreateOrdersFromExcel;

        private DataTable RolePermissionsDataTable;

        public OrdersManager OrdersManager;
        public FrontsCatalogOrder FrontsCatalogOrder;
        public DecorCatalogOrder DecorCatalogOrder;
        public OrdersCalculate OrdersCalculate;

        private RoleTypes RoleType = RoleTypes.Ordinary;
        public enum RoleTypes
        {
            Ordinary = 0,
            Admin = 1,
            Marketing = 2,
            Director = 3
        }

        public MarketingNewOrdersForm(LightStartForm tLightStartForm)
        {
            InitializeComponent();
            LightStartForm = tLightStartForm;


            this.MaximumSize = Screen.PrimaryScreen.WorkingArea.Size;

            Initialize();
            
            while (!SplashForm.bCreated) ;
        }

        private void RolePermissions()
        {
            RolePermissionsDataTable = OrdersManager.GetPermissions(Security.CurrentUserID, this.Name);
            btnDeleteAgreedOrder.Visible = false;
            btnSetAgreementStatus.Visible = false;
            btnSetAgreedDate.Visible = false;
            btnDeleteAgreedMainOrder.Visible = false;
            btnEditAgreedMainOrder.Visible = false;
            kryptonContextMenuItem25.Visible = false;
            kryptonContextMenuItem27.Visible = false;
            if (!PermissionGranted(iMarketing) && !PermissionGranted(iDirector) && !PermissionGranted(iAdmin))
            {
                tableLayoutPanel1.Height = this.Height - NavigatePanel.Height - 10;
                ToolsPanel.Visible = false;
            }
            if (PermissionGranted(iMarketing))
            {
                RoleType = RoleTypes.Marketing;
            }
            if (PermissionGranted(iAdmin))
            {
                btnDeleteAgreedMainOrder.Visible = true;
                btnEditAgreedMainOrder.Visible = true;
                btnDeleteAgreedOrder.Visible = true;
                btnSetAgreementStatus.Visible = true;
                btnSetAgreedDate.Visible = true;
                kryptonContextMenuItem25.Visible = true;
                kryptonContextMenuItem27.Visible = true;
                RoleType = RoleTypes.Admin;
            }
            if (PermissionGranted(iDirector))
            {
                btnDeleteAgreedMainOrder.Visible = true;
                btnEditAgreedMainOrder.Visible = true;
                btnDeleteAgreedOrder.Visible = true;
                btnSetAgreementStatus.Visible = true;
                btnSetAgreedDate.Visible = true;
                kryptonContextMenuItem25.Visible = true;
                kryptonContextMenuItem27.Visible = true;
                RoleType = RoleTypes.Director;
            }
        }

        //private bool PermissionGranted(int PermissionID)
        //{
        //    DataRow[] Rows = RolePermissionsDataTable.Select("PermissionID = " + PermissionID);

        //    if (Rows.Count() > 0)
        //    {
        //        return Convert.ToBoolean(Rows[0]["Granted"]);
        //    }

        //    return false;
        //}

        private bool PermissionGranted(int RoleID)
        {
            var Rows = RolePermissionsDataTable.Select("RoleID = " + RoleID);

            return Rows.Count() > 0;
        }

        private void MarketingOrdersForm_Shown(object sender, EventArgs e)
        {
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
                    NeedSplash = true;
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

            CreateOrdersFromExcel = new CreateOrdersFromExcel();

            FrontsCatalogOrder = new FrontsCatalogOrder();
            FrontsCatalogOrder.Initialize(false);

            DecorCatalogOrder = new DecorCatalogOrder();
            DecorCatalogOrder.Initialize(false);

            OrdersManager = new OrdersManager(ref MainOrdersDataGrid, ref MainOrdersFrontsOrdersDataGrid, ref MegaOrdersDataGrid,
                ref MainOrdersDecorTabControl, ref MainOrdersTabControl, ref DecorCatalogOrder);
            RolePermissions();
            OrdersCalculate = new OrdersCalculate();
            Report = new Report(ref DecorCatalogOrder, ref OrdersCalculate.FrontsCalculate);
            DetailsReport = new DetailsReport(ref DecorCatalogOrder, ref OrdersCalculate.FrontsCalculate);
            PrepareReport = new Modules.Marketing.NewOrders.PrepareReport.DetailsReport(ref DecorCatalogOrder, ref OrdersCalculate.FrontsCalculate);
            SendEmail = new SendEmail();

            FilterClientsDataGrid.DataSource = OrdersManager.FilterClientsBindingSource;
            FilterClientsDataGrid.Columns["ClientID"].Visible = false;

            FilterManagersDataGrid.DataSource = OrdersManager.FilterManagersBindingSource;
            FilterManagersDataGrid.Columns["ManagerID"].Visible = false;
            FilterManagersDataGrid.Columns["userId"].Visible = false;
            FilterManagersDataGrid.Columns["checkCol"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            FilterManagersDataGrid.Columns["checkCol"].Width = 40;

            GetCurrency(CurrencyDateTimePicker.Value.Date);

            var bManager = false;
            var managers = new List<int>();
            for (var i = 0; i < FilterManagersDataGrid.Rows.Count; i++)
            {
                if (Convert.ToInt32(FilterManagersDataGrid.Rows[i].Cells["userId"].Value) == Security.CurrentUserID)
                {
                    if (RoleType == RoleTypes.Admin || RoleType == RoleTypes.Director)
                        continue;
                    ManagerCheckBox.Checked = true;
                    FilterManagersDataGrid.Rows[i].Cells["checkCol"].Value = true;
                    managers.Add(Convert.ToInt32(FilterManagersDataGrid.Rows[i].Cells["managerId"].Value));
                    bManager = true;
                }
            }

            OrdersManager.GetOrdersInMuttlements(
                true,
                true,
                true,
                true,
                true, true, true, true, true, true);
            OrdersManager.FilterMegaOrders(
                false, 0, bManager, managers, true, true, true, true,
                true, true, true, true, true, false, true, true, true, true, true, true);
        }

        private void MegaOrdersDataGrid_SelectionChanged(object sender, EventArgs e)
        {
            if (MegaOrdersDataGrid.Focused)
            {

            }
            if (OrdersManager != null)
                if (OrdersManager.MegaOrdersBindingSource.Count > 0)
                {
                    if (((DataRowView)OrdersManager.MegaOrdersBindingSource.Current)["MegaOrderID"] != DBNull.Value)
                    {
                        var OnProduction = OnProductionCheckBox.Checked;
                        var NotInProduction = NotProductionCheckBox.Checked;
                        var InProduction = InProductionCheckBox.Checked;
                        var OnStorage = OnStorageCheckBox.Checked;
                        var OnExpedition = cbOnExpedition.Checked;
                        var Dispatch = DispatchCheckBox.Checked;

                        if (NeedSplash)
                        {
                            var T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                            T.Start();

                            while (!SplashWindow.bSmallCreated) ;
                            NeedSplash = false;

                            OrdersManager.FilterMainOrdersByMegaOrder(
                                Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current)["MegaOrderID"]),
                                OnProduction,
                                NotInProduction,
                                InProduction,
                                OnStorage,
                                OnExpedition,
                                Dispatch);

                            while (SplashWindow.bSmallCreated)
                                SmallWaitForm.CloseS = true;
                            NeedSplash = true;
                        }
                        else
                            OrdersManager.FilterMainOrdersByMegaOrder(
                                Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current)["MegaOrderID"]),
                                OnProduction,
                                NotInProduction,
                                InProduction,
                                OnStorage,
                                OnExpedition,
                                Dispatch);

                    }
                }
                else
                    OrdersManager.MainOrdersDataTable.Clear();
        }

        private void MainOrdersDataGrid_SelectionChanged(object sender, EventArgs e)
        {
            //MainOrdersSquareLabel.Text = string.Empty;
            if (MainOrdersDataGrid.Focused)
            {

            }
            if (OrdersManager != null)
            {
                MainOrdersTabControl.SuspendLayout();
                MainOrdersDecorTabControl.SuspendLayout();
                if (OrdersManager.MainOrdersBindingSource.Count > 0)
                {
                    if (((DataRowView)OrdersManager.MainOrdersBindingSource.Current)["MainOrderID"] != DBNull.Value)
                    {
                        OrdersManager.FilterProductByMainOrder(Convert.ToInt32(((DataRowView)OrdersManager.MainOrdersBindingSource.Current)["MainOrderID"]));
                        if (MainOrdersTabControl.TabPages[0].PageVisible && MainOrdersTabControl.TabPages[1].PageVisible)
                            MainOrdersTabControl.SelectedTabPage = MainOrdersTabControl.TabPages[0];
                        OrdersManager.MainOrdersDecorOrders.ClientID = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current)["ClientID"]);
                        //decimal Square = OrdersManager.GetSelectedMainOrdersSquare();
                        //if (Square > 0)
                        //    MainOrdersSquareLabel.Text = "Площадь выбранных подзаказов: " + Square + " м.кв";
                    }
                }
                else
                {
                    OrdersManager.FilterProductByMainOrder(-1);
                }
                MainOrdersTabControl.ResumeLayout();
                MainOrdersDecorTabControl.ResumeLayout();
            }
        }

        private void MegaOrdersDataGrid_MouseMove(object sender, MouseEventArgs e)
        {
            if (NeedRefresh == true)
            {
                MegaOrdersDataGrid.Refresh();
                NeedRefresh = false;
            }
        }

        private void MegaOrdersDataGrid_Scroll(object sender, ScrollEventArgs e)
        {
            if (e.ScrollOrientation == ScrollOrientation.HorizontalScroll)
                NeedRefresh = true;
        }

        private void MainOrdersDataGrid_Scroll(object sender, ScrollEventArgs e)
        {
            if (e.ScrollOrientation == ScrollOrientation.HorizontalScroll)
                NeedRefresh = true;
        }

        private void MainOrdersDataGrid_MouseMove(object sender, MouseEventArgs e)
        {
            if (NeedRefresh == true)
            {
                MainOrdersDataGrid.Refresh();
                NeedRefresh = false;
            }
        }

        private void Filter()
        {
            var bClient = ClientCheckBox.Checked;
            var bManager = ManagerCheckBox.Checked;
            var notAgreed = NotAgreedCheckBox.Checked;
            var onAgreement = OnAgreementCheckBox.Checked;
            var notConfirm = NotConfirmCheckBox.Checked;
            var confirm = ConfirmCheckBox.Checked;
            var onProduction = OnProductionCheckBox.Checked;
            var notInProduction = NotProductionCheckBox.Checked;
            var inProduction = InProductionCheckBox.Checked;
            var onStorage = OnStorageCheckBox.Checked;
            var onExpedition = cbOnExpedition.Checked;
            var dispatch = DispatchCheckBox.Checked;
            var bsDelayOfPayment = cbDelayOfPayment.Checked;
            var bsHalfOfPayment = cbHalfOfPayment.Checked;
            var bsFullPayment = cbFullPayment.Checked;
            var bsFactoring = cbFactoring.Checked;
            var bsHalfOfPayment2 = kryptonCheckBox1.Checked;
            var bsDelayOfPayment2 = kryptonCheckBox2.Checked;

            var clientId = -1;

            var managers = new List<int>();
            for (var i = 0; i < FilterManagersDataGrid.Rows.Count; i++)
            {
                if (Convert.ToBoolean(FilterManagersDataGrid.Rows[i].Cells["checkCol"].EditedFormattedValue))
                    managers.Add(Convert.ToInt32(FilterManagersDataGrid.Rows[i].Cells["managerId"].Value));
            }

            if (NeedSplash)
            {
                var T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                T.Start();

                while (!SplashWindow.bSmallCreated) ;
                NeedSplash = false;


                if (bClient && FilterClientsDataGrid.SelectedRows.Count > 0)
                    clientId = Convert.ToInt32(FilterClientsDataGrid.SelectedRows[0].Cells["ClientID"].Value);

                OrdersManager.GetOrdersInMuttlements(
                    notAgreed,
                    onAgreement,
                    notConfirm,
                    confirm,
                    bsDelayOfPayment, bsHalfOfPayment, bsFullPayment, bsFactoring, bsHalfOfPayment2, bsDelayOfPayment2);
                OrdersManager.FilterMegaOrders(
                    bClient, clientId,
                    bManager, managers,
                    notAgreed,
                    onAgreement,
                    notConfirm,
                    confirm,
                    onProduction,
                    notInProduction,
                    inProduction,
                    onStorage,
                    onExpedition,
                    dispatch, bsDelayOfPayment, bsHalfOfPayment, bsFullPayment, bsFactoring, bsHalfOfPayment2, bsDelayOfPayment2);

                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;
                NeedSplash = true;
            }
            else
            {

                if (bClient && FilterClientsDataGrid.SelectedRows.Count > 0)
                    clientId = Convert.ToInt32(FilterClientsDataGrid.SelectedRows[0].Cells["ClientID"].Value);

                OrdersManager.GetOrdersInMuttlements(
                    notAgreed,
                    onAgreement,
                    notConfirm,
                    confirm,
                    bsDelayOfPayment, bsHalfOfPayment, bsFullPayment, bsFactoring, bsHalfOfPayment2, bsDelayOfPayment2);
                OrdersManager.FilterMegaOrders(
                    bClient, clientId,
                    bManager, managers,
                    notAgreed,
                    onAgreement,
                    notConfirm,
                    confirm,
                    onProduction,
                    notInProduction,
                    inProduction,
                    onStorage,
                    onExpedition,
                    dispatch, bsDelayOfPayment, bsHalfOfPayment, bsFullPayment, bsFactoring, bsHalfOfPayment2, bsDelayOfPayment2);
            }
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

        private void NotAgreedCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (OrdersManager == null)
                return;
            Filter();
        }

        private void NotConfirmCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (OrdersManager == null)
                return;
            Filter();
        }

        private void ConfirmCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (OrdersManager == null)
                return;
            Filter();
        }

        private void OnProductionCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (OrdersManager == null)
                return;
            Filter();
        }

        private void NotProductionCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (OrdersManager == null)
                return;
            Filter();
        }

        private void InProductionCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (OrdersManager == null)
                return;
            Filter();
        }

        private void OnStorageCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (OrdersManager == null)
                return;
            Filter();
        }

        private void DispatchCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (OrdersManager == null)
                return;
            Filter();
        }

        private void MenuButton_Click(object sender, EventArgs e)
        {
            MenuPanel.Visible = !MenuPanel.Visible;
        }

        private void ClientCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            FilterClientsDataGrid.Enabled = ClientCheckBox.Checked;

            if (OrdersManager == null)
                return; Filter();
        }

        private void FilterClientsDataGrid_SelectionChanged(object sender, EventArgs e)
        {
            if (OrdersManager != null && ClientCheckBox.Checked)
                Filter();
        }

        private void NewMainOrderButton_Click(object sender, EventArgs e)
        {
            var MegaOrderID = OrdersManager.CurrentMegaOrderID;
            if (MegaOrderID < 0)
            {
                Infinium.LightMessageBox.Show(ref TopForm, false,
                    "Для добавления подзаказа сначала нужно создать новый заказ",
                    "Добавление подзаказа");
                return;
            }

            if (Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["AgreementStatusID"]) == 2)
            {
                Infinium.LightMessageBox.Show(ref TopForm, false,
                    "Заказ был согласован, добавление подзаказов запрещено",
                    "Добавление подзаказа");
                return;
            }

            if (!OrdersManager.CanAddMainOrderToMegaOrder())
            {
                Infinium.LightMessageBox.Show(ref TopForm, false,
                    "Заказ был отдан в производство или уже отгружен, добавление подзаказов запрещено",
                    "Добавление подзаказа");
                return;
            }

            var T = new Thread(delegate () { SplashWindow.CreateSplash(); });
            T.Start();

            while (!SplashForm.bCreated) ;

            if (OrdersManager.MegaOrdersBindingSource.Count > 0)
                OrdersManager.CurrentClientID = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["ClientID"]);
            OrdersManager.CurrentDiscountPaymentConditionID = 1;
            OrdersManager.CurrencyTypeID = 1;
        
            var ClientOMC = OrdersManager.CurrentClientID == 232;

            AddMainOrdersForm = new AddMarketingNewOrdersForm(ref OrdersManager, ClientOMC, false, false, ref TopForm, ref OrdersCalculate);

            TopForm = AddMainOrdersForm;

            AddMainOrdersForm.ShowDialog();

            AddMainOrdersForm.Close();
            AddMainOrdersForm.Dispose();

            TopForm = null;

            var OnProduction = OnProductionCheckBox.Checked;
            var NotInProduction = NotProductionCheckBox.Checked;
            var InProduction = InProductionCheckBox.Checked;
            var OnStorage = OnStorageCheckBox.Checked;
            var OnExpedition = cbOnExpedition.Checked;
            var Dispatch = DispatchCheckBox.Checked;

            NeedSplash = false;

            Filter();

            OrdersManager.UpdateMainOrders(
                OnProduction,
                NotInProduction,
                InProduction,
                OnStorage,
                OnExpedition,
                Dispatch);
            NeedSplash = true;
        }

        private void MainOrdersEditOrder_Click(object sender, EventArgs e)
        {
            OrdersManager.NeedSetStatus = true;
            if (OrdersManager.MainOrdersBindingSource.Count > 0)
            {
                NeedSplash = false;

                var IsSample = false;
                var Notes = string.Empty;
                //получение значений параметров заказа, если заблокирован - выход
                if (!OrdersManager.EditMainOrder(ref Notes, ref IsSample))
                    return;

                var T = new Thread(delegate () { SplashWindow.CreateSplash(); });
                T.Start();

                while (!SplashForm.bCreated) ;

                if (OrdersManager.MegaOrdersBindingSource.Count > 0)
                {
                    OrdersManager.CurrentClientID = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["ClientID"]);
                    OrdersManager.CurrentDiscountPaymentConditionID = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["DiscountPaymentConditionID"]);
                    OrdersManager.CurrentDiscountFactoringID = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["DiscountFactoringID"]);
                    OrdersManager.CurrentProfilDiscountDirector = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["ProfilDiscountDirector"]);
                    OrdersManager.CurrentTPSDiscountDirector = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["TPSDiscountDirector"]);
                    OrdersManager.CurrentProfilTotalDiscount = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["ProfilTotalDiscount"]) - Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["ProfilDiscountDirector"]);
                    OrdersManager.CurrentTPSTotalDiscount = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["TPSTotalDiscount"]) - Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["TPSDiscountDirector"]);
                    OrdersManager.CurrencyTypeID = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["CurrencyTypeID"]);
                    OrdersManager.PaymentCurrency = Convert.ToDecimal(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["PaymentRate"]);
                    OrdersManager.ConfirmDateTime = ((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["ConfirmDateTime"];
                }
                var ClientOMC = OrdersManager.CurrentClientID == 232;

                AddMainOrdersForm = new AddMarketingNewOrdersForm(ref OrdersManager, ClientOMC, true, false, ref TopForm, ref OrdersCalculate);

                TopForm = AddMainOrdersForm;
                AddMainOrdersForm.ShowDialog();

                AddMainOrdersForm.Close();
                AddMainOrdersForm.Dispose();

                TopForm = null;

                var OnProduction = OnProductionCheckBox.Checked;
                var NotInProduction = NotProductionCheckBox.Checked;
                var InProduction = InProductionCheckBox.Checked;
                var OnStorage = OnStorageCheckBox.Checked;
                var OnExpedition = cbOnExpedition.Checked;
                var Dispatch = DispatchCheckBox.Checked;

                NeedSplash = false;
                Filter();

                OrdersManager.UpdateMainOrders(
                    OnProduction,
                    NotInProduction,
                    InProduction,
                    OnStorage,
                    OnExpedition,
                    Dispatch);
                NeedSplash = true;
            }
            OrdersManager.CurrentClientID = -1;
            OrdersManager.CurrentDiscountPaymentConditionID = 0;
            OrdersManager.CurrentDiscountFactoringID = 0;
            OrdersManager.CurrentProfilDiscountDirector = 0;
            OrdersManager.CurrentTPSDiscountDirector = 0;
            OrdersManager.CurrentProfilTotalDiscount = 0;
            OrdersManager.CurrentTPSTotalDiscount = 0;
            OrdersManager.PaymentCurrency = 1;
            OrdersManager.CurrencyTypeID = 1;
            OrdersManager.ConfirmDateTime = DBNull.Value;
        }

        private void MainOrdersRemoveMainOrder_Click(object sender, EventArgs e)
        {
            if (OrdersManager.MainOrdersBindingSource.Count == 0)
                return;

            var MainOrderID = ((DataRowView)OrdersManager.MainOrdersBindingSource.Current)["MainOrderID"].ToString();
            var MegaOrderID = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["MegaOrderID"]);
            //DialogResult result = MessageBox.Show(this, "Вы действительно хотите удалить заказ " + MainOrderID + " ?",
            //                "Удаление заказа", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            //if (OrdersManager.GetAgreementStatus() > 1)
            //{
            //    Infinium.LightMessageBox.Show(ref TopForm, false,
            //           "Невозможно удалить подзаказ, так как он уже согласован",
            //           "Удаление подзаказа");
            //    NeedSplash = true;

            //    return;
            //}

            var OKCancel = Infinium.LightMessageBox.Show(ref TopForm, true,
                    "Вы действительно хотите удалить подзаказ " + MainOrderID + " ?",
                    "Удаление подзаказа");
            if (!OKCancel)
                return;

            NeedSplash = false;

            var IsSample = false;
            var Notes = string.Empty;
            //получение значений параметров заказа, если заблокирован - выход
            if (!OrdersManager.EditMainOrder(ref Notes, ref IsSample))
                return;

            var T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
            T.Start();

            while (!SplashWindow.bSmallCreated) ;

            OrdersManager.RemoveCurrentMainOrder();

            var OnProduction = OnProductionCheckBox.Checked;
            var NotInProduction = NotProductionCheckBox.Checked;
            var InProduction = InProductionCheckBox.Checked;
            var OnStorage = OnStorageCheckBox.Checked;
            var OnExpedition = cbOnExpedition.Checked;
            var Dispatch = DispatchCheckBox.Checked;

            if (OrdersManager.MainOrdersBindingSource.Count > 0)
                OrdersManager.FilterMainOrdersByMegaOrder(
                    Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["MegaOrderID"]),
                    OnProduction,
                    NotInProduction,
                    InProduction,
                    OnStorage,
                    OnExpedition,
                    true);

            OrdersManager.SetNotAgreed(Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["MegaOrderID"]));

            OrdersManager.SummaryCost(Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["MegaOrderID"]));
            OrdersManager.FixOrderEvent(Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["MegaOrderID"]), "Удален подзаказ");
            Filter();
            OrdersManager.MoveToMegaOrder(MegaOrderID);

            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
            NeedSplash = true;
        }

        private void MainOrdersRemoveMegaOrder_Click(object sender, EventArgs e)
        {
            if (OrdersManager.MegaOrdersBindingSource.Count > 0)
            {
                var ClientID = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current)["ClientID"]);
                var OrderNumber = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current)["OrderNumber"]);
                var MegaOrderID = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current)["MegaOrderID"]);

                var OKCancel = Infinium.LightMessageBox.Show(ref TopForm, true,
                        "Вы действительно хотите удалить заказ " + MegaOrderID + " ?",
                        "Удаление заказа");
                if (!OKCancel)
                    return;

                NeedSplash = false;
                var T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                T.Start();

                while (!SplashWindow.bSmallCreated) ;

                OrdersManager.RemoveCurrentMegaOrder();
                OrdersManager.FixOrderEvent(MegaOrderID, "Заказ удален");

                if (OrdersManager.MegaOrdersCount == 0)
                {
                    OrdersManager.CurrentMegaOrderID = -1;
                }
                Filter();

                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;
                NeedSplash = true;
            }
        }

        private void MainOrdersAcceptOrder_Click(object sender, EventArgs e)
        {
            //for (int i = 0; i < MegaOrdersDataGrid.SelectedRows.Count; i++)
            //{
            //    int MegaOrderID = Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[i].Cells["MegaOrderID"].Value);
            //    int AgreementStatusID = Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[i].Cells["AgreementStatusID"].Value);
            //    if (AgreementStatusID == 2)
            //        OrdersManager.MoveOrdersTo(MegaOrderID, AgreementStatusID);
            //}
            //return;
            if (OrdersManager.MegaOrdersBindingSource.Count > 0)
            {
                OrdersManager.MainOrdersDecorOrders.AddMegaOrderToDecorAssignment(
                    Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current)["MegaOrderID"]));

                if (Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current)["AgreementStatusID"]) == 2)
                {
                    OrdersManager.MoveOrdersTo1(Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current)["MegaOrderID"]));
                    //OrdersManager.MoveOrdersTo(Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current)["MegaOrderID"]));
                    Infinium.LightMessageBox.Show(ref TopForm, false,
                           "Заказ уже согласован",
                           "Согласование заказа");
                    return;
                }

                NeedSplash = false;
                if (!OrdersManager.CanBeAccepted())
                {
                    Infinium.LightMessageBox.Show(ref TopForm, false,
                           "Проверьте статус заказа. До подтверждения статус заказа должен быть \"Не подтвержден\"",
                           "Согласование заказа");
                    NeedSplash = true;
                    return;
                }

                var MegaOrderID = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["MegaOrderID"]);

                var OKCancel = Infinium.LightMessageBox.Show(ref TopForm, true,
                        "Подтверждать заказ должен клиент. Вы уверены, что хотите подтвердить заказ?",
                        "Согласование заказа");
                var ClientID = 0;
                var ClientName = string.Empty;
                InvoiceReportToDbf DBFReport = null;
                PhantomForm PhantomForm = null;
                if (OKCancel)
                {
                    PhantomForm = new Infinium.PhantomForm();
                    PhantomForm.Show();

                    ClientID = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["ClientID"]);
                    ClientName = OrdersManager.GetClientName(ClientID);
                    var bCanDirectorDiscount = false;
                    if (RoleType == RoleTypes.Admin || RoleType == RoleTypes.Director)
                        bCanDirectorDiscount = true;

                    DBFReport = new InvoiceReportToDbf(FrontsCatalogOrder, DecorCatalogOrder);
                    CurrencyForm = new CurrencyForm(this, ref OrdersManager, ref OrdersCalculate, ref DBFReport, 
                        ClientID, ClientName, bCanDirectorDiscount, RoleType == RoleTypes.Admin);

                    OrdersManager.CurrentClientID = ClientID;
                    OrdersManager.CurrencyTypeID = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["CurrencyTypeID"]);

                    CurrencyForm.SetParameters(
                            Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["TransportType"]),
                            Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["DelayOfPayment"]),
                            ((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["DesireDate"],
                            ((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["ConfirmDateTime"],
                            Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["CurrencyTypeID"]),
                            Convert.ToDecimal(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["Rate"]),
                            Convert.ToDecimal(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["PaymentRate"]),
                            Convert.ToDecimal(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["ComplaintProfilCost"]),
                            Convert.ToDecimal(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["ComplaintTPSCost"]),
                            ((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["ComplaintNotes"].ToString(),
                            Convert.ToDecimal(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["TransportCost"]),
                            Convert.ToDecimal(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["AdditionalCost"]),
                            Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["DiscountPaymentConditionID"]),
                            Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["DiscountFactoringID"]),
                            Convert.ToDecimal(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["ProfilDiscountDirector"]),
                            Convert.ToDecimal(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["TPSDiscountDirector"]),
                            Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["Weight"]));

                    TopForm = CurrencyForm;
                    CurrencyForm.ShowDialog();

                    PhantomForm.Close();

                    PhantomForm.Dispose();
                    CurrencyForm.Dispose();
                    TopForm = null;

                    OrdersManager.AcceptOrder();
                    OrdersManager.MoveOrdersTo(Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current)["MegaOrderID"]),
                        Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current)["AgreementStatusID"]));
                    OrdersManager.FixOrderEvent(Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["MegaOrderID"]), "Заказ согласован");
                }
                else
                    return;
                if (OrdersManager.SendReport == false)
                    return;

                NeedSplash = false;
                Filter();
                OrdersManager.MoveToMegaOrder(MegaOrderID);
                NeedSplash = true;
            }
        }

        private void NewMegaOrderButton_Click(object sender, EventArgs e)
        {
            var PhantomForm = new Infinium.PhantomForm();
            PhantomForm.Show();

            var ClientID = 0;
            var FromExcel = false;

            if (OrdersManager.MegaOrdersBindingSource.Count > 0)
            {
                ClientID = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["ClientID"]);
            }

            var NewOrderSelectMenu = new NewOrderSelectClientsMenu(this, ClientID);

            TopForm = NewOrderSelectMenu;
            NewOrderSelectMenu.ShowDialog();
            ClientID = NewOrderSelectMenu.ClientID;
            FromExcel = NewOrderSelectMenu.FromExcel;

            PhantomForm.Close();

            PhantomForm.Dispose();
            NewOrderSelectMenu.Dispose();
            TopForm = null;

            if (ClientID > 0)
            {
                if (FromExcel)
                {
                    CreateOrdersFromExcel.GetDataFromExcel();
                    CreateOrdersFromExcel.CreateOrders(ClientID);
                }
                else
                    OrdersManager.CreateNewMegaOrder(ClientID);

                NeedSplash = false;
                Filter();
                NeedSplash = true;
            }
        }

        private void MainOrdersCalculateOrder_Click(object sender, EventArgs e)
        {
            //asdf();
            //Filter();
            //return;

            if (OrdersManager.MegaOrdersBindingSource.Count == 0)
                return;

            //if (Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current)["AgreementStatusID"]) == 2)
            //{
            //    Infinium.LightMessageBox.Show(ref TopForm, false,
            //           "Заказ уже согласован",
            //           "Расчет заказа");
            //    return;
            //}

            var PhantomForm = new Infinium.PhantomForm();
            PhantomForm.Show();

            var MegaOrderID = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["MegaOrderID"]);
            var ClientID = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["ClientID"]);
            var CurrencyTypeID = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["CurrencyTypeID"]);
            var ClientName = OrdersManager.GetClientName(ClientID);
            var bCanDirectorDiscount = false;
            if (RoleType == RoleTypes.Admin || RoleType == RoleTypes.Director)
                bCanDirectorDiscount = true;

            OrdersManager.ConfirmDateTime = ((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["ConfirmDateTime"];
            OrdersManager.CurrentClientID = ClientID;
            OrdersManager.CurrencyTypeID = CurrencyTypeID;

            var DBFReport = new InvoiceReportToDbf(FrontsCatalogOrder, DecorCatalogOrder);
            CurrencyForm = new CurrencyForm(this, ref OrdersManager, ref OrdersCalculate, ref DBFReport, 
                ClientID, ClientName, bCanDirectorDiscount, RoleType == RoleTypes.Admin);

            CurrencyForm.SetParameters(
                    Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["TransportType"]),
                    Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["DelayOfPayment"]),
                    ((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["DesireDate"],
                    ((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["ConfirmDateTime"],
                    Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["CurrencyTypeID"]),
                    Convert.ToDecimal(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["Rate"]),
                    Convert.ToDecimal(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["PaymentRate"]),
                    Convert.ToDecimal(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["ComplaintProfilCost"]),
                    Convert.ToDecimal(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["ComplaintTPSCost"]),
                    ((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["ComplaintNotes"].ToString(),
                    Convert.ToDecimal(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["TransportCost"]),
                    Convert.ToDecimal(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["AdditionalCost"]),
                    Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["DiscountPaymentConditionID"]),
                    Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["DiscountFactoringID"]),
                    Convert.ToDecimal(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["ProfilDiscountDirector"]),
                    Convert.ToDecimal(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["TPSDiscountDirector"]),
                            Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["Weight"]));

            TopForm = CurrencyForm;
            CurrencyForm.ShowDialog();

            PhantomForm.Close();

            PhantomForm.Dispose();
            CurrencyForm.Dispose();
            TopForm = null;

            NeedSplash = false;
            Filter();
            OrdersManager.MoveToMegaOrder(MegaOrderID);
            NeedSplash = true;

            ////REPORT
            DetailsReport.Save = false;
            DetailsReport.Send = false;

            PhantomForm = new Infinium.PhantomForm();
            PhantomForm.Show();

            ClientReportMenu = new Infinium.ClientReportMenu(this);

            TopForm = ClientReportMenu;
            ClientReportMenu.ShowDialog();

            PhantomForm.Close();
            DetailsReport.Save = ClientReportMenu.Save;
            DetailsReport.Send = ClientReportMenu.Send;

            PhantomForm.Dispose();
            ClientReportMenu.Dispose();
            TopForm = null;

            if (DetailsReport.Save || DetailsReport.Send)
            {
                var ComplaintProfilCost = Convert.ToDecimal(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["ComplaintProfilCost"]);
                var ComplaintTPSCost = Convert.ToDecimal(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["ComplaintTPSCost"]);
                var TotalWeight = Convert.ToDecimal(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["Weight"]);
                var TransportCost = Convert.ToDecimal(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["CurrencyTransportCost"]);
                var AdditionalCost = Convert.ToDecimal(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["CurrencyAdditionalCost"]);
                var TotalCost = Convert.ToDecimal(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["CurrencyTotalCost"]);
                CurrencyTypeID = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["CurrencyTypeID"]);
                var Rate = Convert.ToDecimal(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["PaymentRate"]);
                var OrderNumber = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["OrderNumber"]);

                var CheckedMegaOrders = new int[1] { Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["MegaOrderID"]) };
                var CheckedOrderNumbers = new int[1] { OrderNumber };
                var MainOrders = OrdersManager.GetMainOrders();

                NeedSplash = false;
                var T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                T.Start();

                while (!SplashWindow.bSmallCreated) ;
                if (ClientID == 145)
                {
                    var HasOrdinaryOrders = OrdersManager.HasOrders(MainOrders, false);
                    var HasSampleOrders = OrdersManager.HasOrders(MainOrders, true);
                    if (HasOrdinaryOrders)
                    {
                        var FileName = DetailsReport.Report(CheckedMegaOrders, CheckedOrderNumbers, MainOrders, ClientID, ClientName,
                            ComplaintProfilCost, ComplaintTPSCost, TransportCost, AdditionalCost,
                            OrdersManager.GetMainOrdersCost(MainOrders), CurrencyTypeID, TotalWeight, false);
                        
                        while (SplashWindow.bSmallCreated)
                            SmallWaitForm.CloseS = true;
                        NeedSplash = true;

                        if (DetailsReport.Send)
                        {
                            NeedSplash = false;

                            T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Отправка письма.\r\nПодождите..."); });
                            T.Start();

                            while (!SplashWindow.bSmallCreated) ;

                            SendEmail.Send(ClientID, string.Join(",", CheckedOrderNumbers), DetailsReport.Save, FileName);

                            while (SplashWindow.bSmallCreated)
                                SmallWaitForm.CloseS = true;
                            NeedSplash = true;

                            if (SendEmail.Success == false)
                            {
                                Infinium.LightMessageBox.Show(ref TopForm, false,
                                       "Отправка отчета невозможна: отсутствует подключение к интернету либо адрес электронной почты указан неверно",
                                       "Отправка отчета");
                            }
                            OrdersManager.FixOrderEvent(Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["MegaOrderID"]), "Отчет отправлен клиенту");
                            //SendEmail.DeleteFile(FileName);
                        }
                    }
                    if (HasSampleOrders)
                    {
                        var FileName = DetailsReport.Report(CheckedMegaOrders, CheckedOrderNumbers, MainOrders, ClientID, ClientName,
                            ComplaintProfilCost, ComplaintTPSCost, TransportCost, AdditionalCost,
                            OrdersManager.GetMainOrdersCost(MainOrders), CurrencyTypeID, TotalWeight, true);
                        
                        while (SplashWindow.bSmallCreated)
                            SmallWaitForm.CloseS = true;
                        NeedSplash = true;

                        if (DetailsReport.Send)
                        {
                            NeedSplash = false;

                            T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Отправка письма.\r\nПодождите..."); });
                            T.Start();

                            while (!SplashWindow.bSmallCreated) ;

                            SendEmail.Send(ClientID, string.Join(",", CheckedOrderNumbers), DetailsReport.Save, FileName);

                            while (SplashWindow.bSmallCreated)
                                SmallWaitForm.CloseS = true;
                            NeedSplash = true;

                            if (SendEmail.Success == false)
                            {
                                Infinium.LightMessageBox.Show(ref TopForm, false,
                                       "Отправка отчета невозможна: отсутствует подключение к интернету либо адрес электронной почты указан неверно",
                                       "Отправка отчета");
                            }
                            OrdersManager.FixOrderEvent(Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["MegaOrderID"]),
                                "Отчет отправлен клиенту");
                        }
                    }
                }
                else
                {

                    //MessageBox.Show("Выполняется функция DetailsReport.Report");
                    var FileName = DetailsReport.Report(CheckedMegaOrders, CheckedOrderNumbers, MainOrders, ClientID, ClientName,
                         ComplaintProfilCost, ComplaintTPSCost, TransportCost, AdditionalCost,
                         OrdersManager.GetMainOrdersCost(MainOrders), CurrencyTypeID, TotalWeight);

                    while (SplashWindow.bSmallCreated)
                        SmallWaitForm.CloseS = true;

                    //MessageBox.Show("Выполнена функция DetailsReport.Report");
                    NeedSplash = true;

                    if (DetailsReport.Send)
                    {
                        NeedSplash = false;

                        T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Отправка письма.\r\nПодождите..."); });
                        T.Start();

                        while (!SplashWindow.bSmallCreated) ;

                        SendEmail.Send(ClientID, string.Join(",", CheckedOrderNumbers), DetailsReport.Save, FileName);

                        while (SplashWindow.bSmallCreated)
                            SmallWaitForm.CloseS = true;
                        NeedSplash = true;

                        if (SendEmail.Success == false)
                        {
                            Infinium.LightMessageBox.Show(ref TopForm, false,
                                   "Отправка отчета невозможна: отсутствует подключение к интернету либо адрес электронной почты указан неверно",
                                   "Отправка отчета");
                        }
                        OrdersManager.FixOrderEvent(Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["MegaOrderID"]), "Отчет отправлен клиенту");
                        //SendEmail.DeleteFile(FileName);
                    }
                }
            }
        }

        private void CreateReportButton_Click(object sender, EventArgs e)
        {
            if (OrdersManager.MainOrdersBindingSource.Count == 0)
                return;

            if (MegaOrdersDataGrid.SelectedRows.Count > 0)
            {
                var d = Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[0].Cells["DiscountPaymentConditionID"].Value);
                for (var i = 0; i < MegaOrdersDataGrid.SelectedRows.Count; i++)
                {
                    if (d != Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[i].Cells["DiscountPaymentConditionID"].Value))
                    {
                        Infinium.LightMessageBox.Show(ref TopForm, false,
                               "Нельзя объединить заказы с различными условиями оплаты",
                               "Создание отчета");
                        return;
                    }
                }
            }
            if (MegaOrdersDataGrid.SelectedRows.Count > 0)
            {
                var c = Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[0].Cells["CurrencyTypeID"].Value);
                for (var i = 0; i < MegaOrdersDataGrid.SelectedRows.Count; i++)
                {
                    if (c != Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[i].Cells["CurrencyTypeID"].Value))
                    {
                        Infinium.LightMessageBox.Show(ref TopForm, false,
                               "Нельзя объединить заказы с различными валютами",
                               "Создание отчета");
                        return;
                    }
                }
            }

            if (!OrdersManager.AreSelectedMegaOrdersOneClient)
            {
                Infinium.LightMessageBox.Show(ref TopForm, false,
                       "Выбраны заказы разных клиентов",
                       "Создание отчета");
                return;
            }

            if (!OrdersManager.AreSelectedMegaOrdersAgree(OrdersManager.GetSelectedMegaOrders()))
            {
                Infinium.LightMessageBox.Show(ref TopForm, false,
                       "Выбранные заказы несогласованы",
                       "Создание отчета");
                return;
            }

            if (!OrdersManager.CanBeReported())
            {
                Infinium.LightMessageBox.Show(ref TopForm, false,
                       "Заказ не был согласован. Создание отчета невозможно.",
                       "Создание отчета");
                return;
            }

            var dbfReport = new NotesInvoiceReportToDbf();
            InvoiceReportToDBF(dbfReport);

            //using (var form = new ReportTypeForm(this))
            //{
            //    var result = form.ShowDialog();
            //    if (result == DialogResult.OK)
            //    {
            //        var invoiceReportType = form.invoiceReportType;

            //        if (invoiceReportType == InvoiceReportType.Notes)
            //        {
            //            var dbfReport = new NotesInvoiceReportToDbf();
            //            InvoiceReportToDBF(dbfReport);
            //        }
            //        if (invoiceReportType == InvoiceReportType.Standard)
            //        {
            //            var dbfReport = new InvoiceReportToDbf(FrontsCatalogOrder, DecorCatalogOrder);
            //            InvoiceReportToDBF(dbfReport);
            //        }
            //        if (invoiceReportType == InvoiceReportType.CvetPatina)
            //        {
            //            var dbfReport = new ColorInvoiceReportToDbf();
            //            InvoiceReportToDBF(dbfReport);
            //        }
            //    }
            //    else
            //    {
            //        return;
            //    }
            //}
        }

        private void InvoiceReportToDBF(Modules.Marketing.NewOrders.InvoiceReportToDbf.InvoiceReportToDbf DBFReport)
        {
            decimal ComplaintProfilCost = 0;
            decimal ComplaintTPSCost = 0;
            decimal TransportCost = 0;
            decimal AdditionalCost = 0;
            decimal TotalWeight = 0;

            var CurrencyTypeID = Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[0].Cells["CurrencyTypeID"].Value);
            var DiscountPaymentConditionID = Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[0].Cells["DiscountPaymentConditionID"].Value);
            var DelayOfPayment = Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[0].Cells["DelayOfPayment"].Value);

            var Rate = Convert.ToDecimal(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["PaymentRate"]);
            var OrderNumber = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["OrderNumber"]);

            for (var i = 0; i < MegaOrdersDataGrid.SelectedRows.Count; i++)
            {
                ComplaintProfilCost += Convert.ToDecimal(MegaOrdersDataGrid.SelectedRows[i].Cells["CurrencyComplaintProfilCost"].Value);
                ComplaintTPSCost += Convert.ToDecimal(MegaOrdersDataGrid.SelectedRows[i].Cells["CurrencyComplaintTPSCost"].Value);
                TotalWeight += Convert.ToDecimal(MegaOrdersDataGrid.SelectedRows[i].Cells["Weight"].Value);
                TransportCost += Convert.ToDecimal(MegaOrdersDataGrid.SelectedRows[i].Cells["CurrencyTransportCost"].Value);
                AdditionalCost += Convert.ToDecimal(MegaOrdersDataGrid.SelectedRows[i].Cells["CurrencyAdditionalCost"].Value);
            }

            var SelectedMegaOrders = OrdersManager.GetSelectedMegaOrders();
            var SelectedOrderNumbers = OrdersManager.GetSelectedOrderNumbers();
            var SelectedProfilOrderNumbers = new int[MegaOrdersDataGrid.SelectedRows.Count];
            var SelectedTPSOrderNumbers = new int[MegaOrdersDataGrid.SelectedRows.Count];
            var MainOrders = OrdersManager.GetMainOrders(SelectedMegaOrders);

            for (var i = 0; i < MegaOrdersDataGrid.SelectedRows.Count; i++)
            {
                if (Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[i].Cells["FactoryID"].Value) == 1 || Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[i].Cells["FactoryID"].Value) == 0)
                    SelectedProfilOrderNumbers[i] = Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[i].Cells["OrderNumber"].Value);
                if (Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[i].Cells["FactoryID"].Value) == 2 || Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[i].Cells["FactoryID"].Value) == 0)
                    SelectedTPSOrderNumbers[i] = Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[i].Cells["OrderNumber"].Value);
            }
            SelectedProfilOrderNumbers = SelectedProfilOrderNumbers.Where(val => val != 0).ToArray();
            SelectedTPSOrderNumbers = SelectedTPSOrderNumbers.Where(val => val != 0).ToArray();
            var ClientID = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["ClientID"]);
            var ClientName = OrdersManager.GetClientName(ClientID);

            NeedSplash = false;

            var MutualSettlementsManager = new Modules.Marketing.MutualSettlements.MutualSettlements();
            MutualSettlementsManager.Initialize();

            if (ClientID == 145)
            {
                var HasOrdinaryOrders = OrdersManager.HasOrders(MainOrders, false);
                var HasSampleOrders = OrdersManager.HasOrders(MainOrders, true);

                if (HasOrdinaryOrders)
                {
                    decimal TotalProfil = 0;
                    decimal TotalTPS = 0;

                    var ProfilDBFName = string.Empty;
                    var TPSDBFName = string.Empty;
                    var ReportName = string.Empty;
                    var ClientCountryID = OrdersManager.GetClientCountry(ClientID);

                    ClientName = ClientName.Replace("\"", " ");
                    ClientName = ClientName.Replace("*", " ");
                    ClientName = ClientName.Replace("|", " ");
                    ClientName = ClientName.Replace(@"\", " ");
                    ClientName = ClientName.Replace(":", " ");
                    ClientName = ClientName.Replace("<", " ");
                    ClientName = ClientName.Replace(">", " ");
                    ClientName = ClientName.Replace("?", " ");
                    ClientName = ClientName.Replace("/", " ");
                    ReportName = ClientName + " №" + string.Join(",", SelectedOrderNumbers);

                    var PhantomForm = new Infinium.PhantomForm();
                    PhantomForm.Show();
                    var InMutualSettlement = false;
                    var Result = 1;
                    var Notes = string.Empty;
                    var SaveFilePath = string.Empty;
                    SaveDBFReportMenu = new Infinium.SaveDBFReportMenu(this, ReportName);

                    TopForm = SaveDBFReportMenu;
                    SaveDBFReportMenu.ShowDialog();

                    PhantomForm.Close();
                    InMutualSettlement = SaveDBFReportMenu.InMutualSettlement;
                    Result = SaveDBFReportMenu.Result;
                    Notes = SaveDBFReportMenu.Notes;
                    SaveFilePath = SaveDBFReportMenu.SaveFilePath;
                    PhantomForm.Dispose();
                    SaveDBFReportMenu.Dispose();
                    TopForm = null;

                    if (Result != 3)
                    {
                        var T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                        T.Start();

                        while (!SplashWindow.bSmallCreated) ;

                        var hssfworkbook = new HSSFWorkbook();

                        DBFReport.CreateReport(ref hssfworkbook, SelectedMegaOrders, SelectedOrderNumbers, MainOrders,
                            ClientID, ClientName, ComplaintProfilCost, ComplaintTPSCost, TransportCost, AdditionalCost,
                            OrdersManager.GetMainOrdersCost(MainOrders), CurrencyTypeID, TotalWeight, false, ref TotalProfil, ref TotalTPS);
                        DBFReport.SaveDBF(SaveFilePath, ReportName, false, ref ProfilDBFName, ref TPSDBFName);

                        DetailsReport.Report(ref hssfworkbook, SelectedMegaOrders, SelectedOrderNumbers, MainOrders, ClientID, ClientName,
                            ComplaintProfilCost, ComplaintTPSCost, TransportCost, AdditionalCost,
                            OrdersManager.GetMainOrdersCost(MainOrders), CurrencyTypeID, TotalWeight, false);

                        if (ClientID == 609)
                        {
                            DetailsReport.ReportModuleOnline(ref hssfworkbook, MainOrders);
                        }
                        var file = new FileInfo(SaveFilePath + @"\" + ReportName + ".xls");
                        var j = 1;
                        while (file.Exists == true)
                        {
                            file = new FileInfo(SaveFilePath + @"\" + ReportName + "(" + j++ + ").xls");
                        }

                        var NewFile = new FileStream(file.FullName, FileMode.Create);
                        hssfworkbook.Write(NewFile);
                        NewFile.Close();

                        if (InMutualSettlement)
                        {
                            var MutualSettlementID = -1;
                            if (TotalProfil > 0)
                            {
                                if (ClientCountryID == 1)
                                    TotalProfil *= 1.2m;
                                TotalProfil = Decimal.Round(TotalProfil, 2, MidpointRounding.AwayFromZero);
                                MutualSettlementID = MutualSettlementsManager.AddMutualSettlement(ClientID, 1, CurrencyTypeID, DiscountPaymentConditionID, DelayOfPayment, TotalProfil, Notes);
                                var fileInfo = new System.IO.FileInfo(file.FullName);
                                MutualSettlementsManager.AttachIncomeDocument(Path.GetFileNameWithoutExtension(file.FullName),
                                    fileInfo.Extension, file.FullName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.InvoiceExcel, MutualSettlementID);
                                fileInfo = new System.IO.FileInfo(ProfilDBFName);
                                MutualSettlementsManager.AttachIncomeDocument(Path.GetFileNameWithoutExtension(ProfilDBFName),
                                    fileInfo.Extension, ProfilDBFName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.InvoiceDbf, MutualSettlementID);
                                var CurrentDate = Security.GetCurrentDate();
                                for (var i = 0; i < SelectedProfilOrderNumbers.Count(); i++)
                                    MutualSettlementsManager.AddMutualSettlementOrder(CurrentDate, ClientID, MutualSettlementID, SelectedProfilOrderNumbers[i]);
                            }
                            if (TotalTPS > 0)
                            {
                                if (ClientCountryID == 1)
                                    TotalTPS *= 1.2m;
                                TotalTPS = Decimal.Round(TotalTPS, 2, MidpointRounding.AwayFromZero);
                                MutualSettlementID = MutualSettlementsManager.AddMutualSettlement(ClientID, 2, CurrencyTypeID, DiscountPaymentConditionID, DelayOfPayment, TotalTPS, Notes);
                                var fileInfo = new System.IO.FileInfo(file.FullName);
                                MutualSettlementsManager.AttachIncomeDocument(Path.GetFileNameWithoutExtension(file.FullName),
                                    fileInfo.Extension, file.FullName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.InvoiceExcel, MutualSettlementID);
                                fileInfo = new System.IO.FileInfo(TPSDBFName);
                                MutualSettlementsManager.AttachIncomeDocument(Path.GetFileNameWithoutExtension(TPSDBFName),
                                    fileInfo.Extension, TPSDBFName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.InvoiceDbf, MutualSettlementID);
                                var CurrentDate = Security.GetCurrentDate();
                                for (var i = 0; i < SelectedTPSOrderNumbers.Count(); i++)
                                    MutualSettlementsManager.AddMutualSettlementOrder(CurrentDate, ClientID, MutualSettlementID, SelectedTPSOrderNumbers[i]);
                                MutualSettlementsManager.AddSubscribesRecord(324, MutualSettlementID);
                            }
                            MutualSettlementsManager.SaveMutualSettlements();
                            MutualSettlementsManager.SaveMutualSettlementOrders();
                            OrdersManager.FixOrderEvent(Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["MegaOrderID"]), "Счет отправлен на бухгалтерию");
                        }

                        while (SplashWindow.bSmallCreated)
                            SmallWaitForm.CloseS = true;
                        NeedSplash = true;

                        DetailsReport.Save = false;
                        DetailsReport.Send = false;

                        PhantomForm = new Infinium.PhantomForm();
                        PhantomForm.Show();

                        ClientReportMenu = new ClientReportMenu(this);
                        TopForm = ClientReportMenu;

                        ClientReportMenu.ShowDialog();

                        PhantomForm.Close();
                        DetailsReport.Save = ClientReportMenu.Save;
                        DetailsReport.Send = ClientReportMenu.Send;

                        PhantomForm.Dispose();
                        ClientReportMenu.Dispose();
                        TopForm = null;

                        System.Diagnostics.Process.Start(file.FullName);
                        ReportName = file.FullName;

                        if (DetailsReport.Save || DetailsReport.Send)
                        {
                            T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Создание отчета.\r\nПодождите..."); });
                            T.Start();

                            while (!SplashWindow.bSmallCreated) ;

                            var FileName = DetailsReport.Report(SelectedMegaOrders, SelectedOrderNumbers, MainOrders, ClientID, ClientName,
                                ComplaintProfilCost, ComplaintTPSCost, TransportCost, AdditionalCost,
                                OrdersManager.GetMainOrdersCost(MainOrders), CurrencyTypeID, TotalWeight, false);

                            while (SplashWindow.bSmallCreated)
                                SmallWaitForm.CloseS = true;
                            NeedSplash = true;

                            if (DetailsReport.Send)
                            {
                                T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Отправка письма.\r\nПодождите..."); });
                                T.Start();
                                while (!SplashWindow.bSmallCreated) ;

                                SendEmail.Send(ClientID, string.Join(",", SelectedOrderNumbers), DetailsReport.Save, FileName);

                                while (SplashWindow.bSmallCreated)
                                    SmallWaitForm.CloseS = true;
                                NeedSplash = true;

                                if (SendEmail.Success == false)
                                {
                                    Infinium.LightMessageBox.Show(ref TopForm, false,
                                           "Отправка отчета невозможна: отсутствует подключение к интернету либо адрес электронной почты указан неверно",
                                           "Отправка письма");
                                }
                                OrdersManager.FixOrderEvent(Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["MegaOrderID"]), "Отчет отправлен клиенту");
                            }
                        }
                    }
                }
                if (HasSampleOrders)
                {
                    decimal TotalProfil = 0;
                    decimal TotalTPS = 0;

                    var ProfilDBFName = string.Empty;
                    var TPSDBFName = string.Empty;
                    var ReportName = string.Empty;
                    var ClientCountryID = OrdersManager.GetClientCountry(ClientID);

                    ClientName = ClientName.Replace("\"", " ");
                    ClientName = ClientName.Replace("*", " ");
                    ClientName = ClientName.Replace("|", " ");
                    ClientName = ClientName.Replace(@"\", " ");
                    ClientName = ClientName.Replace(":", " ");
                    ClientName = ClientName.Replace("<", " ");
                    ClientName = ClientName.Replace(">", " ");
                    ClientName = ClientName.Replace("?", " ");
                    ClientName = ClientName.Replace("/", " ");
                    ReportName = ClientName + " №" + string.Join(",", SelectedOrderNumbers) + ", обр";

                    var PhantomForm = new Infinium.PhantomForm();
                    PhantomForm.Show();
                    var InMutualSettlement = false;
                    var Result = 1;
                    var Notes = string.Empty;
                    var SaveFilePath = string.Empty;
                    SaveDBFReportMenu = new Infinium.SaveDBFReportMenu(this, ReportName);

                    TopForm = SaveDBFReportMenu;
                    SaveDBFReportMenu.ShowDialog();

                    PhantomForm.Close();
                    InMutualSettlement = SaveDBFReportMenu.InMutualSettlement;
                    Result = SaveDBFReportMenu.Result;
                    Notes = SaveDBFReportMenu.Notes;
                    SaveFilePath = SaveDBFReportMenu.SaveFilePath;
                    PhantomForm.Dispose();
                    SaveDBFReportMenu.Dispose();
                    TopForm = null;

                    if (Result != 3)
                    {
                        var T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                        T.Start();

                        while (!SplashWindow.bSmallCreated) ;

                        var hssfworkbook = new HSSFWorkbook();

                        DBFReport.CreateReport(ref hssfworkbook, SelectedMegaOrders, SelectedOrderNumbers, MainOrders,
                            ClientID, ClientName, ComplaintProfilCost, ComplaintTPSCost, TransportCost, AdditionalCost,
                            OrdersManager.GetMainOrdersCost(MainOrders), CurrencyTypeID, TotalWeight, true, ref TotalProfil, ref TotalTPS);
                        DBFReport.SaveDBF(SaveFilePath, ReportName, true, ref ProfilDBFName, ref TPSDBFName);

                        DetailsReport.Report(ref hssfworkbook, SelectedMegaOrders, SelectedOrderNumbers, MainOrders, ClientID, ClientName,
                            ComplaintProfilCost, ComplaintTPSCost, TransportCost, AdditionalCost,
                            OrdersManager.GetMainOrdersCost(MainOrders), CurrencyTypeID, TotalWeight, true);

                        if (ClientID == 609)
                        {
                            DetailsReport.ReportModuleOnline(ref hssfworkbook, MainOrders);
                        }
                        var file = new FileInfo(SaveFilePath + @"\" + ReportName + ".xls");
                        var j = 1;
                        while (file.Exists == true)
                        {
                            file = new FileInfo(SaveFilePath + @"\" + ReportName + "(" + j++ + ").xls");
                        }

                        var NewFile = new FileStream(file.FullName, FileMode.Create);
                        hssfworkbook.Write(NewFile);
                        NewFile.Close();

                        if (InMutualSettlement)
                        {
                            var MutualSettlementID = -1;
                            if (TotalProfil > 0)
                            {
                                if (ClientCountryID == 1)
                                    TotalProfil *= 1.2m;
                                TotalProfil = Decimal.Round(TotalProfil, 2, MidpointRounding.AwayFromZero);
                                MutualSettlementID = MutualSettlementsManager.AddMutualSettlement(ClientID, 1, CurrencyTypeID, DiscountPaymentConditionID, DelayOfPayment, TotalProfil, Notes, true);
                                var fileInfo = new System.IO.FileInfo(file.FullName);
                                MutualSettlementsManager.AttachIncomeDocument(Path.GetFileNameWithoutExtension(file.FullName),
                                    fileInfo.Extension, file.FullName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.InvoiceExcel, MutualSettlementID);
                                fileInfo = new System.IO.FileInfo(ProfilDBFName);
                                MutualSettlementsManager.AttachIncomeDocument(Path.GetFileNameWithoutExtension(ProfilDBFName),
                                    fileInfo.Extension, ProfilDBFName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.InvoiceDbf, MutualSettlementID);
                                var CurrentDate = Security.GetCurrentDate();
                                for (var i = 0; i < SelectedProfilOrderNumbers.Count(); i++)
                                    MutualSettlementsManager.AddMutualSettlementOrder(CurrentDate, ClientID, MutualSettlementID, SelectedProfilOrderNumbers[i]);
                            }
                            if (TotalTPS > 0)
                            {
                                if (ClientCountryID == 1)
                                    TotalTPS *= 1.2m;
                                TotalTPS = Decimal.Round(TotalTPS, 2, MidpointRounding.AwayFromZero);
                                MutualSettlementID = MutualSettlementsManager.AddMutualSettlement(ClientID, 2, CurrencyTypeID, DiscountPaymentConditionID, DelayOfPayment, TotalTPS, Notes, true);
                                var fileInfo = new System.IO.FileInfo(file.FullName);
                                MutualSettlementsManager.AttachIncomeDocument(Path.GetFileNameWithoutExtension(file.FullName),
                                    fileInfo.Extension, file.FullName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.InvoiceExcel, MutualSettlementID);
                                fileInfo = new System.IO.FileInfo(TPSDBFName);
                                MutualSettlementsManager.AttachIncomeDocument(Path.GetFileNameWithoutExtension(TPSDBFName),
                                    fileInfo.Extension, TPSDBFName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.InvoiceDbf, MutualSettlementID);
                                var CurrentDate = Security.GetCurrentDate();
                                for (var i = 0; i < SelectedTPSOrderNumbers.Count(); i++)
                                    MutualSettlementsManager.AddMutualSettlementOrder(CurrentDate, ClientID, MutualSettlementID, SelectedTPSOrderNumbers[i]);
                                MutualSettlementsManager.AddSubscribesRecord(324, MutualSettlementID);
                            }
                            MutualSettlementsManager.SaveMutualSettlements();
                            MutualSettlementsManager.SaveMutualSettlementOrders();
                            OrdersManager.FixOrderEvent(Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["MegaOrderID"]), "Счет отправлен на бухгалтерию");
                        }

                        while (SplashWindow.bSmallCreated)
                            SmallWaitForm.CloseS = true;
                        NeedSplash = true;

                        DetailsReport.Save = false;
                        DetailsReport.Send = false;

                        PhantomForm = new Infinium.PhantomForm();
                        PhantomForm.Show();

                        ClientReportMenu = new ClientReportMenu(this);
                        TopForm = ClientReportMenu;

                        ClientReportMenu.ShowDialog();

                        PhantomForm.Close();
                        DetailsReport.Save = ClientReportMenu.Save;
                        DetailsReport.Send = ClientReportMenu.Send;

                        PhantomForm.Dispose();
                        ClientReportMenu.Dispose();
                        TopForm = null;

                        System.Diagnostics.Process.Start(file.FullName);
                        ReportName = file.FullName;

                        if (DetailsReport.Save || DetailsReport.Send)
                        {
                            T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Создание отчета.\r\nПодождите..."); });
                            T.Start();

                            while (!SplashWindow.bSmallCreated) ;

                            var FileName = DetailsReport.Report(SelectedMegaOrders, SelectedOrderNumbers, MainOrders, ClientID, ClientName,
                                ComplaintProfilCost, ComplaintTPSCost, TransportCost, AdditionalCost,
                                OrdersManager.GetMainOrdersCost(MainOrders), CurrencyTypeID, TotalWeight, true);

                            while (SplashWindow.bSmallCreated)
                                SmallWaitForm.CloseS = true;

                            if (DetailsReport.Send)
                            {
                                T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Отправка письма.\r\nПодождите..."); });
                                T.Start();
                                while (!SplashWindow.bSmallCreated) ;

                                SendEmail.Send(ClientID, string.Join(",", SelectedOrderNumbers), DetailsReport.Save, FileName);

                                while (SplashWindow.bSmallCreated)
                                    SmallWaitForm.CloseS = true;
                                NeedSplash = true;

                                if (SendEmail.Success == false)
                                {
                                    Infinium.LightMessageBox.Show(ref TopForm, false,
                                           "Отправка отчета невозможна: отсутствует подключение к интернету либо адрес электронной почты указан неверно",
                                           "Отправка письма");
                                }
                                OrdersManager.FixOrderEvent(Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["MegaOrderID"]), "Отчет отправлен клиенту");
                            }
                        }
                    }
                }
            }
            else
            {
                decimal TotalProfil = 0;
                decimal TotalTPS = 0;

                var ProfilDBFName = string.Empty;
                var TPSDBFName = string.Empty;
                var ReportName = string.Empty;
                var ClientCountryID = OrdersManager.GetClientCountry(ClientID);

                ClientName = ClientName.Replace("\"", " ");
                ClientName = ClientName.Replace("*", " ");
                ClientName = ClientName.Replace("|", " ");
                ClientName = ClientName.Replace(@"\", " ");
                ClientName = ClientName.Replace(":", " ");
                ClientName = ClientName.Replace("<", " ");
                ClientName = ClientName.Replace(">", " ");
                ClientName = ClientName.Replace("?", " ");
                ClientName = ClientName.Replace("/", " ");
                ReportName = ClientName + " №" + string.Join(",", SelectedOrderNumbers);

                var PhantomForm = new Infinium.PhantomForm();
                PhantomForm.Show();
                var InMutualSettlement = false;
                var Result = 1;
                var Notes = string.Empty;
                var SaveFilePath = string.Empty;
                SaveDBFReportMenu = new Infinium.SaveDBFReportMenu(this, ReportName);

                TopForm = SaveDBFReportMenu;
                SaveDBFReportMenu.ShowDialog();

                PhantomForm.Close();
                InMutualSettlement = SaveDBFReportMenu.InMutualSettlement;
                Result = SaveDBFReportMenu.Result;
                Notes = SaveDBFReportMenu.Notes;
                SaveFilePath = SaveDBFReportMenu.SaveFilePath;
                PhantomForm.Dispose();
                SaveDBFReportMenu.Dispose();
                TopForm = null;

                if (Result != 3)
                {
                    var T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                    T.Start();

                    while (!SplashWindow.bSmallCreated) ;

                    var hssfworkbook = new HSSFWorkbook();

                    DBFReport.CreateReport(ref hssfworkbook, SelectedMegaOrders, SelectedOrderNumbers, MainOrders,
                        ClientID, ClientName, ComplaintProfilCost, ComplaintTPSCost, TransportCost, AdditionalCost,
                        OrdersManager.GetMainOrdersCost(MainOrders), CurrencyTypeID, TotalWeight, ref TotalProfil, ref TotalTPS);
                    DBFReport.SaveDBF(SaveFilePath, ReportName, ref ProfilDBFName, ref TPSDBFName);

                    DetailsReport.Report(ref hssfworkbook, SelectedMegaOrders, SelectedOrderNumbers, MainOrders, ClientID, ClientName,
                        ComplaintProfilCost, ComplaintTPSCost, TransportCost, AdditionalCost,
                        OrdersManager.GetMainOrdersCost(MainOrders), CurrencyTypeID, TotalWeight);

                    if (ClientID == 609)
                    {
                        DetailsReport.ReportModuleOnline(ref hssfworkbook, MainOrders);
                    }
                    var file = new FileInfo(SaveFilePath + @"\" + ReportName + ".xls");
                    var j = 1;
                    while (file.Exists == true)
                    {
                        file = new FileInfo(SaveFilePath + @"\" + ReportName + "(" + j++ + ").xls");
                    }

                    var NewFile = new FileStream(file.FullName, FileMode.Create);
                    hssfworkbook.Write(NewFile);
                    NewFile.Close();

                    if (InMutualSettlement)
                    {
                        var MutualSettlementID = -1;

                        if (TotalProfil > 0)
                        {
                            if (ClientCountryID == 1)
                                TotalProfil *= 1.2m;
                            TotalProfil = Decimal.Round(TotalProfil, 2, MidpointRounding.AwayFromZero);
                            MutualSettlementID = MutualSettlementsManager.AddMutualSettlement(ClientID, 1, CurrencyTypeID, DiscountPaymentConditionID, DelayOfPayment, TotalProfil, Notes);
                            var fileInfo = new System.IO.FileInfo(file.FullName);
                            MutualSettlementsManager.AttachIncomeDocument(Path.GetFileNameWithoutExtension(file.FullName),
                                fileInfo.Extension, file.FullName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.InvoiceExcel, MutualSettlementID);
                            fileInfo = new System.IO.FileInfo(ProfilDBFName);
                            MutualSettlementsManager.AttachIncomeDocument(Path.GetFileNameWithoutExtension(ProfilDBFName),
                                fileInfo.Extension, ProfilDBFName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.InvoiceDbf, MutualSettlementID);
                            var CurrentDate = Security.GetCurrentDate();
                            for (var i = 0; i < SelectedProfilOrderNumbers.Count(); i++)
                                MutualSettlementsManager.AddMutualSettlementOrder(CurrentDate, ClientID, MutualSettlementID, SelectedProfilOrderNumbers[i]);
                            //MutualSettlementsManager.AddSubscribesRecord(404, MutualSettlementID);
                        }
                        if (TotalTPS > 0)
                        {
                            if (ClientCountryID == 1)
                                TotalTPS *= 1.2m;
                            TotalTPS = Decimal.Round(TotalTPS, 2, MidpointRounding.AwayFromZero);
                            MutualSettlementID = MutualSettlementsManager.AddMutualSettlement(ClientID, 2, CurrencyTypeID, DiscountPaymentConditionID, DelayOfPayment, TotalTPS, Notes);
                            var fileInfo = new System.IO.FileInfo(file.FullName);
                            MutualSettlementsManager.AttachIncomeDocument(Path.GetFileNameWithoutExtension(file.FullName),
                                fileInfo.Extension, file.FullName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.InvoiceExcel, MutualSettlementID);
                            fileInfo = new System.IO.FileInfo(TPSDBFName);
                            MutualSettlementsManager.AttachIncomeDocument(Path.GetFileNameWithoutExtension(TPSDBFName),
                                fileInfo.Extension, TPSDBFName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.InvoiceDbf, MutualSettlementID);
                            var CurrentDate = Security.GetCurrentDate();
                            for (var i = 0; i < SelectedTPSOrderNumbers.Count(); i++)
                                MutualSettlementsManager.AddMutualSettlementOrder(CurrentDate, ClientID, MutualSettlementID, SelectedTPSOrderNumbers[i]);
                            MutualSettlementsManager.AddSubscribesRecord(324, MutualSettlementID);
                        }
                        MutualSettlementsManager.SaveMutualSettlements();
                        MutualSettlementsManager.SaveMutualSettlementOrders();
                        OrdersManager.FixOrderEvent(Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["MegaOrderID"]), "Счет отправлен на бухгалтерию");
                    }

                    while (SplashWindow.bSmallCreated)
                        SmallWaitForm.CloseS = true;
                    NeedSplash = true;

                    DetailsReport.Save = false;
                    DetailsReport.Send = false;

                    PhantomForm = new Infinium.PhantomForm();
                    PhantomForm.Show();

                    ClientReportMenu = new ClientReportMenu(this);
                    TopForm = ClientReportMenu;

                    ClientReportMenu.ShowDialog();

                    PhantomForm.Close();
                    DetailsReport.Save = ClientReportMenu.Save;
                    DetailsReport.Send = ClientReportMenu.Send;

                    PhantomForm.Dispose();
                    ClientReportMenu.Dispose();
                    TopForm = null;

                    System.Diagnostics.Process.Start(file.FullName);
                    ReportName = file.FullName;

                    if (DetailsReport.Save || DetailsReport.Send)
                    {
                        //REPORT
                        T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Создание отчета.\r\nПодождите..."); });
                        T.Start();

                        while (!SplashWindow.bSmallCreated) ;

                        var FileName = DetailsReport.Report(SelectedMegaOrders, SelectedOrderNumbers, MainOrders, ClientID, ClientName,
                            ComplaintProfilCost, ComplaintTPSCost, TransportCost, AdditionalCost,
                            OrdersManager.GetMainOrdersCost(MainOrders), CurrencyTypeID, TotalWeight);

                        while (SplashWindow.bSmallCreated)
                            SmallWaitForm.CloseS = true;
                        NeedSplash = true;

                        if (DetailsReport.Send)
                        {
                            //SEND
                            T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Отправка письма.\r\nПодождите..."); });
                            T.Start();
                            while (!SplashWindow.bSmallCreated) ;

                            SendEmail.Send(ClientID, string.Join(",", SelectedOrderNumbers), DetailsReport.Save, FileName);

                            while (SplashWindow.bSmallCreated)
                                SmallWaitForm.CloseS = true;
                            NeedSplash = true;

                            if (SendEmail.Success == false)
                            {
                                Infinium.LightMessageBox.Show(ref TopForm, false,
                                       "Отправка отчета невозможна: отсутствует подключение к интернету либо адрес электронной почты указан неверно",
                                       "Отправка письма");
                            }
                            OrdersManager.FixOrderEvent(Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["MegaOrderID"]), "Отчет отправлен клиенту");
                            //SendEmail.DeleteFile(FileName);
                        }
                    }
                }
            }
        }

        private void InvoiceReportToDBF(Modules.Marketing.NewOrders.ColorInvoiceReportToDbf.ColorInvoiceReportToDbf DBFReport)
        {
            decimal ComplaintProfilCost = 0;
            decimal ComplaintTPSCost = 0;
            decimal TransportCost = 0;
            decimal AdditionalCost = 0;
            decimal TotalWeight = 0;

            var CurrencyTypeID = Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[0].Cells["CurrencyTypeID"].Value);
            var DiscountPaymentConditionID = Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[0].Cells["DiscountPaymentConditionID"].Value);
            var DelayOfPayment = Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[0].Cells["DelayOfPayment"].Value);

            var Rate = Convert.ToDecimal(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["PaymentRate"]);
            var OrderNumber = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["OrderNumber"]);

            for (var i = 0; i < MegaOrdersDataGrid.SelectedRows.Count; i++)
            {
                ComplaintProfilCost += Convert.ToDecimal(MegaOrdersDataGrid.SelectedRows[i].Cells["CurrencyComplaintProfilCost"].Value);
                ComplaintTPSCost += Convert.ToDecimal(MegaOrdersDataGrid.SelectedRows[i].Cells["CurrencyComplaintTPSCost"].Value);
                TotalWeight += Convert.ToDecimal(MegaOrdersDataGrid.SelectedRows[i].Cells["Weight"].Value);
                TransportCost += Convert.ToDecimal(MegaOrdersDataGrid.SelectedRows[i].Cells["CurrencyTransportCost"].Value);
                AdditionalCost += Convert.ToDecimal(MegaOrdersDataGrid.SelectedRows[i].Cells["CurrencyAdditionalCost"].Value);
            }

            var SelectedMegaOrders = OrdersManager.GetSelectedMegaOrders();
            var SelectedOrderNumbers = OrdersManager.GetSelectedOrderNumbers();
            var SelectedProfilOrderNumbers = new int[MegaOrdersDataGrid.SelectedRows.Count];
            var SelectedTPSOrderNumbers = new int[MegaOrdersDataGrid.SelectedRows.Count];
            var MainOrders = OrdersManager.GetMainOrders(SelectedMegaOrders);

            for (var i = 0; i < MegaOrdersDataGrid.SelectedRows.Count; i++)
            {
                if (Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[i].Cells["FactoryID"].Value) == 1 || Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[i].Cells["FactoryID"].Value) == 0)
                    SelectedProfilOrderNumbers[i] = Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[i].Cells["OrderNumber"].Value);
                if (Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[i].Cells["FactoryID"].Value) == 2 || Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[i].Cells["FactoryID"].Value) == 0)
                    SelectedTPSOrderNumbers[i] = Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[i].Cells["OrderNumber"].Value);
            }
            SelectedProfilOrderNumbers = SelectedProfilOrderNumbers.Where(val => val != 0).ToArray();
            SelectedTPSOrderNumbers = SelectedTPSOrderNumbers.Where(val => val != 0).ToArray();
            var ClientID = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["ClientID"]);
            var ClientName = OrdersManager.GetClientName(ClientID);

            NeedSplash = false;

            var MutualSettlementsManager = new Modules.Marketing.MutualSettlements.MutualSettlements();
            MutualSettlementsManager.Initialize();

            if (ClientID == 145)
            {
                var HasOrdinaryOrders = OrdersManager.HasOrders(MainOrders, false);
                var HasSampleOrders = OrdersManager.HasOrders(MainOrders, true);

                if (HasOrdinaryOrders)
                {
                    decimal TotalProfil = 0;
                    decimal TotalTPS = 0;

                    var ProfilDBFName = string.Empty;
                    var TPSDBFName = string.Empty;
                    var ReportName = string.Empty;
                    var ClientCountryID = OrdersManager.GetClientCountry(ClientID);

                    ClientName = ClientName.Replace("\"", " ");
                    ClientName = ClientName.Replace("*", " ");
                    ClientName = ClientName.Replace("|", " ");
                    ClientName = ClientName.Replace(@"\", " ");
                    ClientName = ClientName.Replace(":", " ");
                    ClientName = ClientName.Replace("<", " ");
                    ClientName = ClientName.Replace(">", " ");
                    ClientName = ClientName.Replace("?", " ");
                    ClientName = ClientName.Replace("/", " ");
                    ReportName = ClientName + " №" + string.Join(",", SelectedOrderNumbers);

                    var PhantomForm = new Infinium.PhantomForm();
                    PhantomForm.Show();
                    var InMutualSettlement = false;
                    var Result = 1;
                    var Notes = string.Empty;
                    var SaveFilePath = string.Empty;
                    SaveDBFReportMenu = new Infinium.SaveDBFReportMenu(this, ReportName);

                    TopForm = SaveDBFReportMenu;
                    SaveDBFReportMenu.ShowDialog();

                    PhantomForm.Close();
                    InMutualSettlement = SaveDBFReportMenu.InMutualSettlement;
                    Result = SaveDBFReportMenu.Result;
                    Notes = SaveDBFReportMenu.Notes;
                    SaveFilePath = SaveDBFReportMenu.SaveFilePath;
                    PhantomForm.Dispose();
                    SaveDBFReportMenu.Dispose();
                    TopForm = null;

                    if (Result != 3)
                    {
                        var T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                        T.Start();

                        while (!SplashWindow.bSmallCreated) ;

                        var hssfworkbook = new HSSFWorkbook();

                        DBFReport.CreateReport(ref hssfworkbook, SelectedMegaOrders, SelectedOrderNumbers, MainOrders,
                            ClientID, ClientName, ComplaintProfilCost, ComplaintTPSCost, TransportCost, AdditionalCost,
                            OrdersManager.GetMainOrdersCost(MainOrders), CurrencyTypeID, TotalWeight, false, ref TotalProfil, ref TotalTPS);
                        DBFReport.SaveDBF(SaveFilePath, ReportName, false, ref ProfilDBFName, ref TPSDBFName);

                        DetailsReport.Report(ref hssfworkbook, SelectedMegaOrders, SelectedOrderNumbers, MainOrders, ClientID, ClientName,
                            ComplaintProfilCost, ComplaintTPSCost, TransportCost, AdditionalCost,
                            OrdersManager.GetMainOrdersCost(MainOrders), CurrencyTypeID, TotalWeight, false);

                        if (ClientID == 609)
                        {
                            DetailsReport.ReportModuleOnline(ref hssfworkbook, MainOrders);
                        }
                        var file = new FileInfo(SaveFilePath + @"\" + ReportName + ".xls");
                        var j = 1;
                        while (file.Exists == true)
                        {
                            file = new FileInfo(SaveFilePath + @"\" + ReportName + "(" + j++ + ").xls");
                        }

                        var NewFile = new FileStream(file.FullName, FileMode.Create);
                        hssfworkbook.Write(NewFile);
                        NewFile.Close();

                        if (InMutualSettlement)
                        {
                            var MutualSettlementID = -1;
                            if (TotalProfil > 0)
                            {
                                if (ClientCountryID == 1)
                                    TotalProfil *= 1.2m;
                                TotalProfil = Decimal.Round(TotalProfil, 2, MidpointRounding.AwayFromZero);
                                MutualSettlementID = MutualSettlementsManager.AddMutualSettlement(ClientID, 1, CurrencyTypeID, DiscountPaymentConditionID, DelayOfPayment, TotalProfil, Notes);
                                var fileInfo = new System.IO.FileInfo(file.FullName);
                                MutualSettlementsManager.AttachIncomeDocument(Path.GetFileNameWithoutExtension(file.FullName),
                                    fileInfo.Extension, file.FullName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.InvoiceExcel, MutualSettlementID);
                                fileInfo = new System.IO.FileInfo(ProfilDBFName);
                                MutualSettlementsManager.AttachIncomeDocument(Path.GetFileNameWithoutExtension(ProfilDBFName),
                                    fileInfo.Extension, ProfilDBFName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.InvoiceDbf, MutualSettlementID);
                                var CurrentDate = Security.GetCurrentDate();
                                for (var i = 0; i < SelectedProfilOrderNumbers.Count(); i++)
                                    MutualSettlementsManager.AddMutualSettlementOrder(CurrentDate, ClientID, MutualSettlementID, SelectedProfilOrderNumbers[i]);
                            }
                            if (TotalTPS > 0)
                            {
                                if (ClientCountryID == 1)
                                    TotalTPS *= 1.2m;
                                TotalTPS = Decimal.Round(TotalTPS, 2, MidpointRounding.AwayFromZero);
                                MutualSettlementID = MutualSettlementsManager.AddMutualSettlement(ClientID, 2, CurrencyTypeID, DiscountPaymentConditionID, DelayOfPayment, TotalTPS, Notes);
                                var fileInfo = new System.IO.FileInfo(file.FullName);
                                MutualSettlementsManager.AttachIncomeDocument(Path.GetFileNameWithoutExtension(file.FullName),
                                    fileInfo.Extension, file.FullName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.InvoiceExcel, MutualSettlementID);
                                fileInfo = new System.IO.FileInfo(TPSDBFName);
                                MutualSettlementsManager.AttachIncomeDocument(Path.GetFileNameWithoutExtension(TPSDBFName),
                                    fileInfo.Extension, TPSDBFName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.InvoiceDbf, MutualSettlementID);
                                var CurrentDate = Security.GetCurrentDate();
                                for (var i = 0; i < SelectedTPSOrderNumbers.Count(); i++)
                                    MutualSettlementsManager.AddMutualSettlementOrder(CurrentDate, ClientID, MutualSettlementID, SelectedTPSOrderNumbers[i]);
                                MutualSettlementsManager.AddSubscribesRecord(324, MutualSettlementID);
                            }
                            MutualSettlementsManager.SaveMutualSettlements();
                            MutualSettlementsManager.SaveMutualSettlementOrders();
                            OrdersManager.FixOrderEvent(Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["MegaOrderID"]), "Счет отправлен на бухгалтерию");
                        }

                        while (SplashWindow.bSmallCreated)
                            SmallWaitForm.CloseS = true;
                        NeedSplash = true;

                        DetailsReport.Save = false;
                        DetailsReport.Send = false;

                        PhantomForm = new Infinium.PhantomForm();
                        PhantomForm.Show();

                        ClientReportMenu = new ClientReportMenu(this);
                        TopForm = ClientReportMenu;

                        ClientReportMenu.ShowDialog();

                        PhantomForm.Close();
                        DetailsReport.Save = ClientReportMenu.Save;
                        DetailsReport.Send = ClientReportMenu.Send;

                        PhantomForm.Dispose();
                        ClientReportMenu.Dispose();
                        TopForm = null;

                        System.Diagnostics.Process.Start(file.FullName);
                        ReportName = file.FullName;

                        if (DetailsReport.Save || DetailsReport.Send)
                        {
                            T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Создание отчета.\r\nПодождите..."); });
                            T.Start();

                            while (!SplashWindow.bSmallCreated) ;

                            var FileName = DetailsReport.Report(SelectedMegaOrders, SelectedOrderNumbers, MainOrders, ClientID, ClientName,
                                ComplaintProfilCost, ComplaintTPSCost, TransportCost, AdditionalCost,
                                OrdersManager.GetMainOrdersCost(MainOrders), CurrencyTypeID, TotalWeight, false);

                            while (SplashWindow.bSmallCreated)
                                SmallWaitForm.CloseS = true;
                            NeedSplash = true;

                            if (DetailsReport.Send)
                            {
                                T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Отправка письма.\r\nПодождите..."); });
                                T.Start();
                                while (!SplashWindow.bSmallCreated) ;

                                SendEmail.Send(ClientID, string.Join(",", SelectedOrderNumbers), DetailsReport.Save, FileName);

                                while (SplashWindow.bSmallCreated)
                                    SmallWaitForm.CloseS = true;
                                NeedSplash = true;

                                if (SendEmail.Success == false)
                                {
                                    Infinium.LightMessageBox.Show(ref TopForm, false,
                                           "Отправка отчета невозможна: отсутствует подключение к интернету либо адрес электронной почты указан неверно",
                                           "Отправка письма");
                                }
                                OrdersManager.FixOrderEvent(Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["MegaOrderID"]), "Отчет отправлен клиенту");
                            }
                        }
                    }
                }
                if (HasSampleOrders)
                {
                    decimal TotalProfil = 0;
                    decimal TotalTPS = 0;

                    var ProfilDBFName = string.Empty;
                    var TPSDBFName = string.Empty;
                    var ReportName = string.Empty;
                    var ClientCountryID = OrdersManager.GetClientCountry(ClientID);

                    ClientName = ClientName.Replace("\"", " ");
                    ClientName = ClientName.Replace("*", " ");
                    ClientName = ClientName.Replace("|", " ");
                    ClientName = ClientName.Replace(@"\", " ");
                    ClientName = ClientName.Replace(":", " ");
                    ClientName = ClientName.Replace("<", " ");
                    ClientName = ClientName.Replace(">", " ");
                    ClientName = ClientName.Replace("?", " ");
                    ClientName = ClientName.Replace("/", " ");
                    ReportName = ClientName + " №" + string.Join(",", SelectedOrderNumbers) + ", обр";

                    var PhantomForm = new Infinium.PhantomForm();
                    PhantomForm.Show();
                    var InMutualSettlement = false;
                    var Result = 1;
                    var Notes = string.Empty;
                    var SaveFilePath = string.Empty;
                    SaveDBFReportMenu = new Infinium.SaveDBFReportMenu(this, ReportName);

                    TopForm = SaveDBFReportMenu;
                    SaveDBFReportMenu.ShowDialog();

                    PhantomForm.Close();
                    InMutualSettlement = SaveDBFReportMenu.InMutualSettlement;
                    Result = SaveDBFReportMenu.Result;
                    Notes = SaveDBFReportMenu.Notes;
                    SaveFilePath = SaveDBFReportMenu.SaveFilePath;
                    PhantomForm.Dispose();
                    SaveDBFReportMenu.Dispose();
                    TopForm = null;

                    if (Result != 3)
                    {
                        var T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                        T.Start();

                        while (!SplashWindow.bSmallCreated) ;

                        var hssfworkbook = new HSSFWorkbook();

                        DBFReport.CreateReport(ref hssfworkbook, SelectedMegaOrders, SelectedOrderNumbers, MainOrders,
                            ClientID, ClientName, ComplaintProfilCost, ComplaintTPSCost, TransportCost, AdditionalCost,
                            OrdersManager.GetMainOrdersCost(MainOrders), CurrencyTypeID, TotalWeight, true, ref TotalProfil, ref TotalTPS);
                        DBFReport.SaveDBF(SaveFilePath, ReportName, true, ref ProfilDBFName, ref TPSDBFName);

                        DetailsReport.Report(ref hssfworkbook, SelectedMegaOrders, SelectedOrderNumbers, MainOrders, ClientID, ClientName,
                            ComplaintProfilCost, ComplaintTPSCost, TransportCost, AdditionalCost,
                            OrdersManager.GetMainOrdersCost(MainOrders), CurrencyTypeID, TotalWeight, true);

                        if (ClientID == 609)
                        {
                            DetailsReport.ReportModuleOnline(ref hssfworkbook, MainOrders);
                        }
                        var file = new FileInfo(SaveFilePath + @"\" + ReportName + ".xls");
                        var j = 1;
                        while (file.Exists == true)
                        {
                            file = new FileInfo(SaveFilePath + @"\" + ReportName + "(" + j++ + ").xls");
                        }

                        var NewFile = new FileStream(file.FullName, FileMode.Create);
                        hssfworkbook.Write(NewFile);
                        NewFile.Close();

                        if (InMutualSettlement)
                        {
                            var MutualSettlementID = -1;
                            if (TotalProfil > 0)
                            {
                                if (ClientCountryID == 1)
                                    TotalProfil *= 1.2m;
                                TotalProfil = Decimal.Round(TotalProfil, 2, MidpointRounding.AwayFromZero);
                                MutualSettlementID = MutualSettlementsManager.AddMutualSettlement(ClientID, 1, CurrencyTypeID, DiscountPaymentConditionID, DelayOfPayment, TotalProfil, Notes, true);
                                var fileInfo = new System.IO.FileInfo(file.FullName);
                                MutualSettlementsManager.AttachIncomeDocument(Path.GetFileNameWithoutExtension(file.FullName),
                                    fileInfo.Extension, file.FullName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.InvoiceExcel, MutualSettlementID);
                                fileInfo = new System.IO.FileInfo(ProfilDBFName);
                                MutualSettlementsManager.AttachIncomeDocument(Path.GetFileNameWithoutExtension(ProfilDBFName),
                                    fileInfo.Extension, ProfilDBFName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.InvoiceDbf, MutualSettlementID);
                                var CurrentDate = Security.GetCurrentDate();
                                for (var i = 0; i < SelectedProfilOrderNumbers.Count(); i++)
                                    MutualSettlementsManager.AddMutualSettlementOrder(CurrentDate, ClientID, MutualSettlementID, SelectedProfilOrderNumbers[i]);
                            }
                            if (TotalTPS > 0)
                            {
                                if (ClientCountryID == 1)
                                    TotalTPS *= 1.2m;
                                TotalTPS = Decimal.Round(TotalTPS, 2, MidpointRounding.AwayFromZero);
                                MutualSettlementID = MutualSettlementsManager.AddMutualSettlement(ClientID, 2, CurrencyTypeID, DiscountPaymentConditionID, DelayOfPayment, TotalTPS, Notes, true);
                                var fileInfo = new System.IO.FileInfo(file.FullName);
                                MutualSettlementsManager.AttachIncomeDocument(Path.GetFileNameWithoutExtension(file.FullName),
                                    fileInfo.Extension, file.FullName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.InvoiceExcel, MutualSettlementID);
                                fileInfo = new System.IO.FileInfo(TPSDBFName);
                                MutualSettlementsManager.AttachIncomeDocument(Path.GetFileNameWithoutExtension(TPSDBFName),
                                    fileInfo.Extension, TPSDBFName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.InvoiceDbf, MutualSettlementID);
                                var CurrentDate = Security.GetCurrentDate();
                                for (var i = 0; i < SelectedTPSOrderNumbers.Count(); i++)
                                    MutualSettlementsManager.AddMutualSettlementOrder(CurrentDate, ClientID, MutualSettlementID, SelectedTPSOrderNumbers[i]);
                                MutualSettlementsManager.AddSubscribesRecord(324, MutualSettlementID);
                            }
                            MutualSettlementsManager.SaveMutualSettlements();
                            MutualSettlementsManager.SaveMutualSettlementOrders();
                            OrdersManager.FixOrderEvent(Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["MegaOrderID"]), "Счет отправлен на бухгалтерию");
                        }

                        while (SplashWindow.bSmallCreated)
                            SmallWaitForm.CloseS = true;
                        NeedSplash = true;

                        DetailsReport.Save = false;
                        DetailsReport.Send = false;

                        PhantomForm = new Infinium.PhantomForm();
                        PhantomForm.Show();

                        ClientReportMenu = new ClientReportMenu(this);
                        TopForm = ClientReportMenu;

                        ClientReportMenu.ShowDialog();

                        PhantomForm.Close();
                        DetailsReport.Save = ClientReportMenu.Save;
                        DetailsReport.Send = ClientReportMenu.Send;

                        PhantomForm.Dispose();
                        ClientReportMenu.Dispose();
                        TopForm = null;

                        System.Diagnostics.Process.Start(file.FullName);
                        ReportName = file.FullName;

                        if (DetailsReport.Save || DetailsReport.Send)
                        {
                            T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Создание отчета.\r\nПодождите..."); });
                            T.Start();

                            while (!SplashWindow.bSmallCreated) ;

                            var FileName = DetailsReport.Report(SelectedMegaOrders, SelectedOrderNumbers, MainOrders, ClientID, ClientName,
                                ComplaintProfilCost, ComplaintTPSCost, TransportCost, AdditionalCost,
                                OrdersManager.GetMainOrdersCost(MainOrders), CurrencyTypeID, TotalWeight, true);

                            while (SplashWindow.bSmallCreated)
                                SmallWaitForm.CloseS = true;

                            if (DetailsReport.Send)
                            {
                                T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Отправка письма.\r\nПодождите..."); });
                                T.Start();
                                while (!SplashWindow.bSmallCreated) ;

                                SendEmail.Send(ClientID, string.Join(",", SelectedOrderNumbers), DetailsReport.Save, FileName);

                                while (SplashWindow.bSmallCreated)
                                    SmallWaitForm.CloseS = true;
                                NeedSplash = true;

                                if (SendEmail.Success == false)
                                {
                                    Infinium.LightMessageBox.Show(ref TopForm, false,
                                           "Отправка отчета невозможна: отсутствует подключение к интернету либо адрес электронной почты указан неверно",
                                           "Отправка письма");
                                }
                                OrdersManager.FixOrderEvent(Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["MegaOrderID"]), "Отчет отправлен клиенту");
                            }
                        }
                    }
                }
            }
            else
            {
                decimal TotalProfil = 0;
                decimal TotalTPS = 0;

                var ProfilDBFName = string.Empty;
                var TPSDBFName = string.Empty;
                var ReportName = string.Empty;
                var ClientCountryID = OrdersManager.GetClientCountry(ClientID);

                ClientName = ClientName.Replace("\"", " ");
                ClientName = ClientName.Replace("*", " ");
                ClientName = ClientName.Replace("|", " ");
                ClientName = ClientName.Replace(@"\", " ");
                ClientName = ClientName.Replace(":", " ");
                ClientName = ClientName.Replace("<", " ");
                ClientName = ClientName.Replace(">", " ");
                ClientName = ClientName.Replace("?", " ");
                ClientName = ClientName.Replace("/", " ");
                ReportName = ClientName + " №" + string.Join(",", SelectedOrderNumbers);

                var PhantomForm = new Infinium.PhantomForm();
                PhantomForm.Show();
                var InMutualSettlement = false;
                var Result = 1;
                var Notes = string.Empty;
                var SaveFilePath = string.Empty;
                SaveDBFReportMenu = new Infinium.SaveDBFReportMenu(this, ReportName);

                TopForm = SaveDBFReportMenu;
                SaveDBFReportMenu.ShowDialog();

                PhantomForm.Close();
                InMutualSettlement = SaveDBFReportMenu.InMutualSettlement;
                Result = SaveDBFReportMenu.Result;
                Notes = SaveDBFReportMenu.Notes;
                SaveFilePath = SaveDBFReportMenu.SaveFilePath;
                PhantomForm.Dispose();
                SaveDBFReportMenu.Dispose();
                TopForm = null;

                if (Result != 3)
                {
                    var T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                    T.Start();

                    while (!SplashWindow.bSmallCreated) ;

                    var hssfworkbook = new HSSFWorkbook();

                    DBFReport.CreateReport(ref hssfworkbook, SelectedMegaOrders, SelectedOrderNumbers, MainOrders,
                        ClientID, ClientName, ComplaintProfilCost, ComplaintTPSCost, TransportCost, AdditionalCost,
                        OrdersManager.GetMainOrdersCost(MainOrders), CurrencyTypeID, TotalWeight, ref TotalProfil, ref TotalTPS);
                    DBFReport.SaveDBF(SaveFilePath, ReportName, ref ProfilDBFName, ref TPSDBFName);

                    DetailsReport.Report(ref hssfworkbook, SelectedMegaOrders, SelectedOrderNumbers, MainOrders, ClientID, ClientName,
                        ComplaintProfilCost, ComplaintTPSCost, TransportCost, AdditionalCost,
                        OrdersManager.GetMainOrdersCost(MainOrders), CurrencyTypeID, TotalWeight);

                    if (ClientID == 609)
                    {
                        DetailsReport.ReportModuleOnline(ref hssfworkbook, MainOrders);
                    }
                    var file = new FileInfo(SaveFilePath + @"\" + ReportName + ".xls");
                    var j = 1;
                    while (file.Exists == true)
                    {
                        file = new FileInfo(SaveFilePath + @"\" + ReportName + "(" + j++ + ").xls");
                    }

                    var NewFile = new FileStream(file.FullName, FileMode.Create);
                    hssfworkbook.Write(NewFile);
                    NewFile.Close();

                    if (InMutualSettlement)
                    {
                        var MutualSettlementID = -1;

                        if (TotalProfil > 0)
                        {
                            if (ClientCountryID == 1)
                                TotalProfil *= 1.2m;
                            TotalProfil = Decimal.Round(TotalProfil, 2, MidpointRounding.AwayFromZero);
                            MutualSettlementID = MutualSettlementsManager.AddMutualSettlement(ClientID, 1, CurrencyTypeID, DiscountPaymentConditionID, DelayOfPayment, TotalProfil, Notes);
                            var fileInfo = new System.IO.FileInfo(file.FullName);
                            MutualSettlementsManager.AttachIncomeDocument(Path.GetFileNameWithoutExtension(file.FullName),
                                fileInfo.Extension, file.FullName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.InvoiceExcel, MutualSettlementID);
                            fileInfo = new System.IO.FileInfo(ProfilDBFName);
                            MutualSettlementsManager.AttachIncomeDocument(Path.GetFileNameWithoutExtension(ProfilDBFName),
                                fileInfo.Extension, ProfilDBFName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.InvoiceDbf, MutualSettlementID);
                            var CurrentDate = Security.GetCurrentDate();
                            for (var i = 0; i < SelectedProfilOrderNumbers.Count(); i++)
                                MutualSettlementsManager.AddMutualSettlementOrder(CurrentDate, ClientID, MutualSettlementID, SelectedProfilOrderNumbers[i]);
                            //MutualSettlementsManager.AddSubscribesRecord(404, MutualSettlementID);
                        }
                        if (TotalTPS > 0)
                        {
                            if (ClientCountryID == 1)
                                TotalTPS *= 1.2m;
                            TotalTPS = Decimal.Round(TotalTPS, 2, MidpointRounding.AwayFromZero);
                            MutualSettlementID = MutualSettlementsManager.AddMutualSettlement(ClientID, 2, CurrencyTypeID, DiscountPaymentConditionID, DelayOfPayment, TotalTPS, Notes);
                            var fileInfo = new System.IO.FileInfo(file.FullName);
                            MutualSettlementsManager.AttachIncomeDocument(Path.GetFileNameWithoutExtension(file.FullName),
                                fileInfo.Extension, file.FullName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.InvoiceExcel, MutualSettlementID);
                            fileInfo = new System.IO.FileInfo(TPSDBFName);
                            MutualSettlementsManager.AttachIncomeDocument(Path.GetFileNameWithoutExtension(TPSDBFName),
                                fileInfo.Extension, TPSDBFName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.InvoiceDbf, MutualSettlementID);
                            var CurrentDate = Security.GetCurrentDate();
                            for (var i = 0; i < SelectedTPSOrderNumbers.Count(); i++)
                                MutualSettlementsManager.AddMutualSettlementOrder(CurrentDate, ClientID, MutualSettlementID, SelectedTPSOrderNumbers[i]);
                            MutualSettlementsManager.AddSubscribesRecord(324, MutualSettlementID);
                        }
                        MutualSettlementsManager.SaveMutualSettlements();
                        MutualSettlementsManager.SaveMutualSettlementOrders();
                        OrdersManager.FixOrderEvent(Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["MegaOrderID"]), "Счет отправлен на бухгалтерию");
                    }

                    while (SplashWindow.bSmallCreated)
                        SmallWaitForm.CloseS = true;
                    NeedSplash = true;

                    DetailsReport.Save = false;
                    DetailsReport.Send = false;

                    PhantomForm = new Infinium.PhantomForm();
                    PhantomForm.Show();

                    ClientReportMenu = new ClientReportMenu(this);
                    TopForm = ClientReportMenu;

                    ClientReportMenu.ShowDialog();

                    PhantomForm.Close();
                    DetailsReport.Save = ClientReportMenu.Save;
                    DetailsReport.Send = ClientReportMenu.Send;

                    PhantomForm.Dispose();
                    ClientReportMenu.Dispose();
                    TopForm = null;

                    System.Diagnostics.Process.Start(file.FullName);
                    ReportName = file.FullName;

                    if (DetailsReport.Save || DetailsReport.Send)
                    {
                        //REPORT
                        T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Создание отчета.\r\nПодождите..."); });
                        T.Start();

                        while (!SplashWindow.bSmallCreated) ;

                        var FileName = DetailsReport.Report(SelectedMegaOrders, SelectedOrderNumbers, MainOrders, ClientID, ClientName,
                            ComplaintProfilCost, ComplaintTPSCost, TransportCost, AdditionalCost,
                            OrdersManager.GetMainOrdersCost(MainOrders), CurrencyTypeID, TotalWeight);

                        if (ClientID == 609)
                        {
                            DetailsReport.ReportModuleOnline(ref hssfworkbook, MainOrders);
                        }
                        while (SplashWindow.bSmallCreated)
                            SmallWaitForm.CloseS = true;
                        NeedSplash = true;

                        if (DetailsReport.Send)
                        {
                            //SEND
                            T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Отправка письма.\r\nПодождите..."); });
                            T.Start();
                            while (!SplashWindow.bSmallCreated) ;

                            SendEmail.Send(ClientID, string.Join(",", SelectedOrderNumbers), DetailsReport.Save, FileName);

                            while (SplashWindow.bSmallCreated)
                                SmallWaitForm.CloseS = true;
                            NeedSplash = true;

                            if (SendEmail.Success == false)
                            {
                                Infinium.LightMessageBox.Show(ref TopForm, false,
                                       "Отправка отчета невозможна: отсутствует подключение к интернету либо адрес электронной почты указан неверно",
                                       "Отправка письма");
                            }
                            OrdersManager.FixOrderEvent(Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["MegaOrderID"]), "Отчет отправлен клиенту");
                            //SendEmail.DeleteFile(FileName);
                        }
                    }
                }
            }
        }
        
        private void InvoiceReportToDBF(Infinium.Modules.Marketing.NewOrders.NotesInvoiceReportToDbf.NotesInvoiceReportToDbf DBFReport)
        {
            decimal ComplaintProfilCost = 0;
            decimal ComplaintTPSCost = 0;
            decimal TransportCost = 0;
            decimal AdditionalCost = 0;
            decimal TotalWeight = 0;

            var CurrencyTypeID = Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[0].Cells["CurrencyTypeID"].Value);
            var DiscountPaymentConditionID = Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[0].Cells["DiscountPaymentConditionID"].Value);
            var DelayOfPayment = Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[0].Cells["DelayOfPayment"].Value);

            var Rate = Convert.ToDecimal(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["PaymentRate"]);
            var OrderNumber = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["OrderNumber"]);

            for (var i = 0; i < MegaOrdersDataGrid.SelectedRows.Count; i++)
            {
                ComplaintProfilCost += Convert.ToDecimal(MegaOrdersDataGrid.SelectedRows[i].Cells["CurrencyComplaintProfilCost"].Value);
                ComplaintTPSCost += Convert.ToDecimal(MegaOrdersDataGrid.SelectedRows[i].Cells["CurrencyComplaintTPSCost"].Value);
                TotalWeight += Convert.ToDecimal(MegaOrdersDataGrid.SelectedRows[i].Cells["Weight"].Value);
                TransportCost += Convert.ToDecimal(MegaOrdersDataGrid.SelectedRows[i].Cells["CurrencyTransportCost"].Value);
                AdditionalCost += Convert.ToDecimal(MegaOrdersDataGrid.SelectedRows[i].Cells["CurrencyAdditionalCost"].Value);
            }

            var SelectedMegaOrders = OrdersManager.GetSelectedMegaOrders();
            var SelectedOrderNumbers = OrdersManager.GetSelectedOrderNumbers();
            var SelectedProfilOrderNumbers = new int[MegaOrdersDataGrid.SelectedRows.Count];
            var SelectedTPSOrderNumbers = new int[MegaOrdersDataGrid.SelectedRows.Count];
            var MainOrders = OrdersManager.GetMainOrders(SelectedMegaOrders);

            for (var i = 0; i < MegaOrdersDataGrid.SelectedRows.Count; i++)
            {
                if (Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[i].Cells["FactoryID"].Value) == 1 || Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[i].Cells["FactoryID"].Value) == 0)
                    SelectedProfilOrderNumbers[i] = Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[i].Cells["OrderNumber"].Value);
                if (Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[i].Cells["FactoryID"].Value) == 2 || Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[i].Cells["FactoryID"].Value) == 0)
                    SelectedTPSOrderNumbers[i] = Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[i].Cells["OrderNumber"].Value);
            }
            SelectedProfilOrderNumbers = SelectedProfilOrderNumbers.Where(val => val != 0).ToArray();
            SelectedTPSOrderNumbers = SelectedTPSOrderNumbers.Where(val => val != 0).ToArray();
            var ClientID = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["ClientID"]);
            var ClientName = OrdersManager.GetClientName(ClientID);

            NeedSplash = false;

            var MutualSettlementsManager = new Modules.Marketing.MutualSettlements.MutualSettlements();
            MutualSettlementsManager.Initialize();

            if (ClientID == 145)
            {
                var HasOrdinaryOrders = OrdersManager.HasOrders(MainOrders, false);
                var HasSampleOrders = OrdersManager.HasOrders(MainOrders, true);

                if (HasOrdinaryOrders)
                {
                    decimal TotalProfil = 0;
                    decimal TotalTPS = 0;

                    var ProfilDBFName = string.Empty;
                    var TPSDBFName = string.Empty;
                    var ReportName = string.Empty;
                    var ClientCountryID = OrdersManager.GetClientCountry(ClientID);

                    ClientName = ClientName.Replace("\"", " ");
                    ClientName = ClientName.Replace("*", " ");
                    ClientName = ClientName.Replace("|", " ");
                    ClientName = ClientName.Replace(@"\", " ");
                    ClientName = ClientName.Replace(":", " ");
                    ClientName = ClientName.Replace("<", " ");
                    ClientName = ClientName.Replace(">", " ");
                    ClientName = ClientName.Replace("?", " ");
                    ClientName = ClientName.Replace("/", " ");
                    ReportName = ClientName + " №" + string.Join(",", SelectedOrderNumbers);

                    var PhantomForm = new Infinium.PhantomForm();
                    PhantomForm.Show();
                    var InMutualSettlement = false;
                    var Result = 1;
                    var Notes = string.Empty;
                    var SaveFilePath = string.Empty;
                    SaveDBFReportMenu = new Infinium.SaveDBFReportMenu(this, ReportName);

                    TopForm = SaveDBFReportMenu;
                    SaveDBFReportMenu.ShowDialog();

                    PhantomForm.Close();
                    InMutualSettlement = SaveDBFReportMenu.InMutualSettlement;
                    Result = SaveDBFReportMenu.Result;
                    Notes = SaveDBFReportMenu.Notes;
                    SaveFilePath = SaveDBFReportMenu.SaveFilePath;
                    PhantomForm.Dispose();
                    SaveDBFReportMenu.Dispose();
                    TopForm = null;

                    if (Result != 3)
                    {
                        var T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                        T.Start();

                        while (!SplashWindow.bSmallCreated) ;

                        var hssfworkbook = new HSSFWorkbook();

                        DBFReport.CreateReport(ref hssfworkbook, SelectedMegaOrders, SelectedOrderNumbers, MainOrders,
                            ClientID, ClientName, ComplaintProfilCost, ComplaintTPSCost, TransportCost, AdditionalCost,
                            OrdersManager.GetMainOrdersCost(MainOrders), CurrencyTypeID, TotalWeight, false, ref TotalProfil, ref TotalTPS);
                        DBFReport.SaveDBF(SaveFilePath, ReportName, false, ref ProfilDBFName, ref TPSDBFName);

                        DetailsReport.Report(ref hssfworkbook, SelectedMegaOrders, SelectedOrderNumbers, MainOrders, ClientID, ClientName,
                            ComplaintProfilCost, ComplaintTPSCost, TransportCost, AdditionalCost,
                            OrdersManager.GetMainOrdersCost(MainOrders), CurrencyTypeID, TotalWeight, false);

                        if (ClientID == 609)
                        {
                            DetailsReport.ReportModuleOnline(ref hssfworkbook, MainOrders);
                        }
                        var file = new FileInfo(SaveFilePath + @"\" + ReportName + ".xls");
                        var j = 1;
                        while (file.Exists == true)
                        {
                            file = new FileInfo(SaveFilePath + @"\" + ReportName + "(" + j++ + ").xls");
                        }

                        var NewFile = new FileStream(file.FullName, FileMode.Create);
                        hssfworkbook.Write(NewFile);
                        NewFile.Close();

                        if (InMutualSettlement)
                        {
                            var MutualSettlementID = -1;
                            if (TotalProfil > 0)
                            {
                                if (ClientCountryID == 1)
                                    TotalProfil *= 1.2m;
                                TotalProfil = Decimal.Round(TotalProfil, 2, MidpointRounding.AwayFromZero);
                                MutualSettlementID = MutualSettlementsManager.AddMutualSettlement(ClientID, 1, CurrencyTypeID, DiscountPaymentConditionID, DelayOfPayment, TotalProfil, Notes);
                                var fileInfo = new System.IO.FileInfo(file.FullName);
                                MutualSettlementsManager.AttachIncomeDocument(Path.GetFileNameWithoutExtension(file.FullName),
                                    fileInfo.Extension, file.FullName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.InvoiceExcel, MutualSettlementID);
                                fileInfo = new System.IO.FileInfo(ProfilDBFName);
                                MutualSettlementsManager.AttachIncomeDocument(Path.GetFileNameWithoutExtension(ProfilDBFName),
                                    fileInfo.Extension, ProfilDBFName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.InvoiceDbf, MutualSettlementID);
                                var CurrentDate = Security.GetCurrentDate();
                                for (var i = 0; i < SelectedProfilOrderNumbers.Count(); i++)
                                    MutualSettlementsManager.AddMutualSettlementOrder(CurrentDate, ClientID, MutualSettlementID, SelectedProfilOrderNumbers[i]);
                            }
                            if (TotalTPS > 0)
                            {
                                if (ClientCountryID == 1)
                                    TotalTPS *= 1.2m;
                                TotalTPS = Decimal.Round(TotalTPS, 2, MidpointRounding.AwayFromZero);
                                MutualSettlementID = MutualSettlementsManager.AddMutualSettlement(ClientID, 2, CurrencyTypeID, DiscountPaymentConditionID, DelayOfPayment, TotalTPS, Notes);
                                var fileInfo = new System.IO.FileInfo(file.FullName);
                                MutualSettlementsManager.AttachIncomeDocument(Path.GetFileNameWithoutExtension(file.FullName),
                                    fileInfo.Extension, file.FullName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.InvoiceExcel, MutualSettlementID);
                                fileInfo = new System.IO.FileInfo(TPSDBFName);
                                MutualSettlementsManager.AttachIncomeDocument(Path.GetFileNameWithoutExtension(TPSDBFName),
                                    fileInfo.Extension, TPSDBFName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.InvoiceDbf, MutualSettlementID);
                                var CurrentDate = Security.GetCurrentDate();
                                for (var i = 0; i < SelectedTPSOrderNumbers.Count(); i++)
                                    MutualSettlementsManager.AddMutualSettlementOrder(CurrentDate, ClientID, MutualSettlementID, SelectedTPSOrderNumbers[i]);
                                MutualSettlementsManager.AddSubscribesRecord(324, MutualSettlementID);
                            }
                            MutualSettlementsManager.SaveMutualSettlements();
                            MutualSettlementsManager.SaveMutualSettlementOrders();
                            OrdersManager.FixOrderEvent(Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["MegaOrderID"]), "Счет отправлен на бухгалтерию");
                        }

                        while (SplashWindow.bSmallCreated)
                            SmallWaitForm.CloseS = true;
                        NeedSplash = true;

                        DetailsReport.Save = false;
                        DetailsReport.Send = false;

                        PhantomForm = new Infinium.PhantomForm();
                        PhantomForm.Show();

                        ClientReportMenu = new ClientReportMenu(this);
                        TopForm = ClientReportMenu;

                        ClientReportMenu.ShowDialog();

                        PhantomForm.Close();
                        DetailsReport.Save = ClientReportMenu.Save;
                        DetailsReport.Send = ClientReportMenu.Send;

                        PhantomForm.Dispose();
                        ClientReportMenu.Dispose();
                        TopForm = null;

                        System.Diagnostics.Process.Start(file.FullName);
                        ReportName = file.FullName;

                        if (DetailsReport.Save || DetailsReport.Send)
                        {
                            T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Создание отчета.\r\nПодождите..."); });
                            T.Start();

                            while (!SplashWindow.bSmallCreated) ;

                            var FileName = DetailsReport.Report(SelectedMegaOrders, SelectedOrderNumbers, MainOrders, ClientID, ClientName,
                                ComplaintProfilCost, ComplaintTPSCost, TransportCost, AdditionalCost,
                                OrdersManager.GetMainOrdersCost(MainOrders), CurrencyTypeID, TotalWeight, false);

                            if (ClientID == 609)
                            {
                                DetailsReport.ReportModuleOnline(ref hssfworkbook, MainOrders);
                            }
                            while (SplashWindow.bSmallCreated)
                                SmallWaitForm.CloseS = true;
                            NeedSplash = true;

                            if (DetailsReport.Send)
                            {
                                T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Отправка письма.\r\nПодождите..."); });
                                T.Start();
                                while (!SplashWindow.bSmallCreated) ;

                                SendEmail.Send(ClientID, string.Join(",", SelectedOrderNumbers), DetailsReport.Save, FileName);

                                while (SplashWindow.bSmallCreated)
                                    SmallWaitForm.CloseS = true;
                                NeedSplash = true;

                                if (SendEmail.Success == false)
                                {
                                    Infinium.LightMessageBox.Show(ref TopForm, false,
                                           "Отправка отчета невозможна: отсутствует подключение к интернету либо адрес электронной почты указан неверно",
                                           "Отправка письма");
                                }
                                OrdersManager.FixOrderEvent(Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["MegaOrderID"]), "Отчет отправлен клиенту");
                            }
                        }
                    }
                }
                if (HasSampleOrders)
                {
                    decimal TotalProfil = 0;
                    decimal TotalTPS = 0;

                    var ProfilDBFName = string.Empty;
                    var TPSDBFName = string.Empty;
                    var ReportName = string.Empty;
                    var ClientCountryID = OrdersManager.GetClientCountry(ClientID);

                    ClientName = ClientName.Replace("\"", " ");
                    ClientName = ClientName.Replace("*", " ");
                    ClientName = ClientName.Replace("|", " ");
                    ClientName = ClientName.Replace(@"\", " ");
                    ClientName = ClientName.Replace(":", " ");
                    ClientName = ClientName.Replace("<", " ");
                    ClientName = ClientName.Replace(">", " ");
                    ClientName = ClientName.Replace("?", " ");
                    ClientName = ClientName.Replace("/", " ");
                    ReportName = ClientName + " №" + string.Join(",", SelectedOrderNumbers) + ", обр";

                    var PhantomForm = new Infinium.PhantomForm();
                    PhantomForm.Show();
                    var InMutualSettlement = false;
                    var Result = 1;
                    var Notes = string.Empty;
                    var SaveFilePath = string.Empty;
                    SaveDBFReportMenu = new Infinium.SaveDBFReportMenu(this, ReportName);

                    TopForm = SaveDBFReportMenu;
                    SaveDBFReportMenu.ShowDialog();

                    PhantomForm.Close();
                    InMutualSettlement = SaveDBFReportMenu.InMutualSettlement;
                    Result = SaveDBFReportMenu.Result;
                    Notes = SaveDBFReportMenu.Notes;
                    SaveFilePath = SaveDBFReportMenu.SaveFilePath;
                    PhantomForm.Dispose();
                    SaveDBFReportMenu.Dispose();
                    TopForm = null;

                    if (Result != 3)
                    {
                        var T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                        T.Start();

                        while (!SplashWindow.bSmallCreated) ;

                        var hssfworkbook = new HSSFWorkbook();

                        DBFReport.CreateReport(ref hssfworkbook, SelectedMegaOrders, SelectedOrderNumbers, MainOrders,
                            ClientID, ClientName, ComplaintProfilCost, ComplaintTPSCost, TransportCost, AdditionalCost,
                            OrdersManager.GetMainOrdersCost(MainOrders), CurrencyTypeID, TotalWeight, true, ref TotalProfil, ref TotalTPS);
                        DBFReport.SaveDBF(SaveFilePath, ReportName, true, ref ProfilDBFName, ref TPSDBFName);

                        DetailsReport.Report(ref hssfworkbook, SelectedMegaOrders, SelectedOrderNumbers, MainOrders, ClientID, ClientName,
                            ComplaintProfilCost, ComplaintTPSCost, TransportCost, AdditionalCost,
                            OrdersManager.GetMainOrdersCost(MainOrders), CurrencyTypeID, TotalWeight, true);

                        if (ClientID == 609)
                        {
                            DetailsReport.ReportModuleOnline(ref hssfworkbook, MainOrders);
                        }
                        var file = new FileInfo(SaveFilePath + @"\" + ReportName + ".xls");
                        var j = 1;
                        while (file.Exists == true)
                        {
                            file = new FileInfo(SaveFilePath + @"\" + ReportName + "(" + j++ + ").xls");
                        }

                        var NewFile = new FileStream(file.FullName, FileMode.Create);
                        hssfworkbook.Write(NewFile);
                        NewFile.Close();

                        if (InMutualSettlement)
                        {
                            var MutualSettlementID = -1;
                            if (TotalProfil > 0)
                            {
                                if (ClientCountryID == 1)
                                    TotalProfil *= 1.2m;
                                TotalProfil = Decimal.Round(TotalProfil, 2, MidpointRounding.AwayFromZero);
                                MutualSettlementID = MutualSettlementsManager.AddMutualSettlement(ClientID, 1, CurrencyTypeID, DiscountPaymentConditionID, DelayOfPayment, TotalProfil, Notes, true);
                                var fileInfo = new System.IO.FileInfo(file.FullName);
                                MutualSettlementsManager.AttachIncomeDocument(Path.GetFileNameWithoutExtension(file.FullName),
                                    fileInfo.Extension, file.FullName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.InvoiceExcel, MutualSettlementID);
                                fileInfo = new System.IO.FileInfo(ProfilDBFName);
                                MutualSettlementsManager.AttachIncomeDocument(Path.GetFileNameWithoutExtension(ProfilDBFName),
                                    fileInfo.Extension, ProfilDBFName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.InvoiceDbf, MutualSettlementID);
                                var CurrentDate = Security.GetCurrentDate();
                                for (var i = 0; i < SelectedProfilOrderNumbers.Count(); i++)
                                    MutualSettlementsManager.AddMutualSettlementOrder(CurrentDate, ClientID, MutualSettlementID, SelectedProfilOrderNumbers[i]);
                            }
                            if (TotalTPS > 0)
                            {
                                if (ClientCountryID == 1)
                                    TotalTPS *= 1.2m;
                                TotalTPS = Decimal.Round(TotalTPS, 2, MidpointRounding.AwayFromZero);
                                MutualSettlementID = MutualSettlementsManager.AddMutualSettlement(ClientID, 2, CurrencyTypeID, DiscountPaymentConditionID, DelayOfPayment, TotalTPS, Notes, true);
                                var fileInfo = new System.IO.FileInfo(file.FullName);
                                MutualSettlementsManager.AttachIncomeDocument(Path.GetFileNameWithoutExtension(file.FullName),
                                    fileInfo.Extension, file.FullName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.InvoiceExcel, MutualSettlementID);
                                fileInfo = new System.IO.FileInfo(TPSDBFName);
                                MutualSettlementsManager.AttachIncomeDocument(Path.GetFileNameWithoutExtension(TPSDBFName),
                                    fileInfo.Extension, TPSDBFName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.InvoiceDbf, MutualSettlementID);
                                var CurrentDate = Security.GetCurrentDate();
                                for (var i = 0; i < SelectedTPSOrderNumbers.Count(); i++)
                                    MutualSettlementsManager.AddMutualSettlementOrder(CurrentDate, ClientID, MutualSettlementID, SelectedTPSOrderNumbers[i]);
                                MutualSettlementsManager.AddSubscribesRecord(324, MutualSettlementID);
                            }
                            MutualSettlementsManager.SaveMutualSettlements();
                            MutualSettlementsManager.SaveMutualSettlementOrders();
                            OrdersManager.FixOrderEvent(Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["MegaOrderID"]), "Счет отправлен на бухгалтерию");
                        }

                        while (SplashWindow.bSmallCreated)
                            SmallWaitForm.CloseS = true;
                        NeedSplash = true;

                        DetailsReport.Save = false;
                        DetailsReport.Send = false;

                        PhantomForm = new Infinium.PhantomForm();
                        PhantomForm.Show();

                        ClientReportMenu = new ClientReportMenu(this);
                        TopForm = ClientReportMenu;

                        ClientReportMenu.ShowDialog();

                        PhantomForm.Close();
                        DetailsReport.Save = ClientReportMenu.Save;
                        DetailsReport.Send = ClientReportMenu.Send;

                        PhantomForm.Dispose();
                        ClientReportMenu.Dispose();
                        TopForm = null;

                        System.Diagnostics.Process.Start(file.FullName);
                        ReportName = file.FullName;

                        if (DetailsReport.Save || DetailsReport.Send)
                        {
                            T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Создание отчета.\r\nПодождите..."); });
                            T.Start();

                            while (!SplashWindow.bSmallCreated) ;

                            var FileName = DetailsReport.Report(SelectedMegaOrders, SelectedOrderNumbers, MainOrders, ClientID, ClientName,
                                ComplaintProfilCost, ComplaintTPSCost, TransportCost, AdditionalCost,
                                OrdersManager.GetMainOrdersCost(MainOrders), CurrencyTypeID, TotalWeight, true);

                            if (ClientID == 609)
                            {
                                DetailsReport.ReportModuleOnline(ref hssfworkbook, MainOrders);
                            }
                            while (SplashWindow.bSmallCreated)
                                SmallWaitForm.CloseS = true;

                            if (DetailsReport.Send)
                            {
                                T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Отправка письма.\r\nПодождите..."); });
                                T.Start();
                                while (!SplashWindow.bSmallCreated) ;

                                SendEmail.Send(ClientID, string.Join(",", SelectedOrderNumbers), DetailsReport.Save, FileName);

                                while (SplashWindow.bSmallCreated)
                                    SmallWaitForm.CloseS = true;
                                NeedSplash = true;

                                if (SendEmail.Success == false)
                                {
                                    Infinium.LightMessageBox.Show(ref TopForm, false,
                                           "Отправка отчета невозможна: отсутствует подключение к интернету либо адрес электронной почты указан неверно",
                                           "Отправка письма");
                                }
                                OrdersManager.FixOrderEvent(Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["MegaOrderID"]), "Отчет отправлен клиенту");
                            }
                        }
                    }
                }
            }
            else
            {
                decimal TotalProfil = 0;
                decimal TotalTPS = 0;

                var ProfilDBFName = string.Empty;
                var TPSDBFName = string.Empty;
                var ReportName = string.Empty;
                var ClientCountryID = OrdersManager.GetClientCountry(ClientID);

                ClientName = ClientName.Replace("\"", " ");
                ClientName = ClientName.Replace("*", " ");
                ClientName = ClientName.Replace("|", " ");
                ClientName = ClientName.Replace(@"\", " ");
                ClientName = ClientName.Replace(":", " ");
                ClientName = ClientName.Replace("<", " ");
                ClientName = ClientName.Replace(">", " ");
                ClientName = ClientName.Replace("?", " ");
                ClientName = ClientName.Replace("/", " ");
                ReportName = ClientName + " №" + string.Join(",", SelectedOrderNumbers);

                var PhantomForm = new Infinium.PhantomForm();
                PhantomForm.Show();
                var InMutualSettlement = false;
                var Result = 1;
                var Notes = string.Empty;
                var SaveFilePath = string.Empty;
                SaveDBFReportMenu = new Infinium.SaveDBFReportMenu(this, ReportName);

                TopForm = SaveDBFReportMenu;
                SaveDBFReportMenu.ShowDialog();

                PhantomForm.Close();
                InMutualSettlement = SaveDBFReportMenu.InMutualSettlement;
                Result = SaveDBFReportMenu.Result;
                Notes = SaveDBFReportMenu.Notes;
                SaveFilePath = SaveDBFReportMenu.SaveFilePath;
                PhantomForm.Dispose();
                SaveDBFReportMenu.Dispose();
                TopForm = null;

                if (Result != 3)
                {
                    var T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                    T.Start();

                    while (!SplashWindow.bSmallCreated) ;

                    var hssfworkbook = new HSSFWorkbook();

                    DBFReport.CreateReport(ref hssfworkbook, SelectedMegaOrders, SelectedOrderNumbers, MainOrders,
                        ClientID, ClientName, ComplaintProfilCost, ComplaintTPSCost, TransportCost, AdditionalCost,
                        OrdersManager.GetMainOrdersCost(MainOrders), CurrencyTypeID, TotalWeight, ref TotalProfil, ref TotalTPS);
                    DBFReport.SaveDBF(SaveFilePath, ReportName, ref ProfilDBFName, ref TPSDBFName);

                    DetailsReport.Report(ref hssfworkbook, SelectedMegaOrders, SelectedOrderNumbers, MainOrders, ClientID, ClientName,
                        ComplaintProfilCost, ComplaintTPSCost, TransportCost, AdditionalCost,
                        OrdersManager.GetMainOrdersCost(MainOrders), CurrencyTypeID, TotalWeight);

                    if (ClientID == 609)
                    {
                        DetailsReport.ReportModuleOnline(ref hssfworkbook, MainOrders);
                    }
                    var file = new FileInfo(SaveFilePath + @"\" + ReportName + ".xls");
                    var j = 1;
                    while (file.Exists == true)
                    {
                        file = new FileInfo(SaveFilePath + @"\" + ReportName + "(" + j++ + ").xls");
                    }

                    var NewFile = new FileStream(file.FullName, FileMode.Create);
                    hssfworkbook.Write(NewFile);
                    NewFile.Close();

                    if (InMutualSettlement)
                    {
                        var MutualSettlementID = -1;

                        if (TotalProfil > 0)
                        {
                            if (ClientCountryID == 1)
                                TotalProfil *= 1.2m;
                            TotalProfil = Decimal.Round(TotalProfil, 2, MidpointRounding.AwayFromZero);
                            MutualSettlementID = MutualSettlementsManager.AddMutualSettlement(ClientID, 1, CurrencyTypeID, DiscountPaymentConditionID, DelayOfPayment, TotalProfil, Notes);
                            var fileInfo = new System.IO.FileInfo(file.FullName);
                            MutualSettlementsManager.AttachIncomeDocument(Path.GetFileNameWithoutExtension(file.FullName),
                                fileInfo.Extension, file.FullName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.InvoiceExcel, MutualSettlementID);
                            fileInfo = new System.IO.FileInfo(ProfilDBFName);
                            MutualSettlementsManager.AttachIncomeDocument(Path.GetFileNameWithoutExtension(ProfilDBFName),
                                fileInfo.Extension, ProfilDBFName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.InvoiceDbf, MutualSettlementID);
                            var CurrentDate = Security.GetCurrentDate();
                            for (var i = 0; i < SelectedProfilOrderNumbers.Count(); i++)
                                MutualSettlementsManager.AddMutualSettlementOrder(CurrentDate, ClientID, MutualSettlementID, SelectedProfilOrderNumbers[i]);
                            //MutualSettlementsManager.AddSubscribesRecord(404, MutualSettlementID);
                        }
                        if (TotalTPS > 0)
                        {
                            if (ClientCountryID == 1)
                                TotalTPS *= 1.2m;
                            TotalTPS = Decimal.Round(TotalTPS, 2, MidpointRounding.AwayFromZero);
                            MutualSettlementID = MutualSettlementsManager.AddMutualSettlement(ClientID, 2, CurrencyTypeID, DiscountPaymentConditionID, DelayOfPayment, TotalTPS, Notes);
                            var fileInfo = new System.IO.FileInfo(file.FullName);
                            MutualSettlementsManager.AttachIncomeDocument(Path.GetFileNameWithoutExtension(file.FullName),
                                fileInfo.Extension, file.FullName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.InvoiceExcel, MutualSettlementID);
                            fileInfo = new System.IO.FileInfo(TPSDBFName);
                            MutualSettlementsManager.AttachIncomeDocument(Path.GetFileNameWithoutExtension(TPSDBFName),
                                fileInfo.Extension, TPSDBFName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.InvoiceDbf, MutualSettlementID);
                            var CurrentDate = Security.GetCurrentDate();
                            for (var i = 0; i < SelectedTPSOrderNumbers.Count(); i++)
                                MutualSettlementsManager.AddMutualSettlementOrder(CurrentDate, ClientID, MutualSettlementID, SelectedTPSOrderNumbers[i]);
                            MutualSettlementsManager.AddSubscribesRecord(324, MutualSettlementID);
                        }
                        MutualSettlementsManager.SaveMutualSettlements();
                        MutualSettlementsManager.SaveMutualSettlementOrders();
                        OrdersManager.FixOrderEvent(Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["MegaOrderID"]), "Счет отправлен на бухгалтерию");
                    }

                    while (SplashWindow.bSmallCreated)
                        SmallWaitForm.CloseS = true;
                    NeedSplash = true;

                    DetailsReport.Save = false;
                    DetailsReport.Send = false;

                    PhantomForm = new Infinium.PhantomForm();
                    PhantomForm.Show();

                    ClientReportMenu = new ClientReportMenu(this);
                    TopForm = ClientReportMenu;

                    ClientReportMenu.ShowDialog();

                    PhantomForm.Close();
                    DetailsReport.Save = ClientReportMenu.Save;
                    DetailsReport.Send = ClientReportMenu.Send;

                    PhantomForm.Dispose();
                    ClientReportMenu.Dispose();
                    TopForm = null;

                    System.Diagnostics.Process.Start(file.FullName);
                    ReportName = file.FullName;

                    if (DetailsReport.Save || DetailsReport.Send)
                    {
                        //REPORT
                        T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Создание отчета.\r\nПодождите..."); });
                        T.Start();

                        while (!SplashWindow.bSmallCreated) ;

                        var FileName = DetailsReport.Report(SelectedMegaOrders, SelectedOrderNumbers, MainOrders, ClientID, ClientName,
                            ComplaintProfilCost, ComplaintTPSCost, TransportCost, AdditionalCost,
                            OrdersManager.GetMainOrdersCost(MainOrders), CurrencyTypeID, TotalWeight);

                        if (ClientID == 609)
                        {
                            DetailsReport.ReportModuleOnline(ref hssfworkbook, MainOrders);
                        }
                        while (SplashWindow.bSmallCreated)
                            SmallWaitForm.CloseS = true;
                        NeedSplash = true;

                        if (DetailsReport.Send)
                        {
                            //SEND
                            T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Отправка письма.\r\nПодождите..."); });
                            T.Start();
                            while (!SplashWindow.bSmallCreated) ;

                            SendEmail.Send(ClientID, string.Join(",", SelectedOrderNumbers), DetailsReport.Save, FileName);

                            while (SplashWindow.bSmallCreated)
                                SmallWaitForm.CloseS = true;
                            NeedSplash = true;

                            if (SendEmail.Success == false)
                            {
                                Infinium.LightMessageBox.Show(ref TopForm, false,
                                       "Отправка отчета невозможна: отсутствует подключение к интернету либо адрес электронной почты указан неверно",
                                       "Отправка письма");
                            }
                            OrdersManager.FixOrderEvent(Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["MegaOrderID"]), "Отчет отправлен клиенту");
                            //SendEmail.DeleteFile(FileName);
                        }
                    }
                }
            }
        }

        public void GetCurrency(DateTime Date)
        {
            var RateExist = false;

            decimal EURRUBCurrency = 1000000;
            decimal USDRUBCurrency = 1000000;
            decimal EURUSDCurrency = 1000000;
            decimal EURBYRCurrency = 1000000;
            //OrdersManager.NBRBDailyRates(Date, ref EURBYRCurrency);
            //OrdersManager.CBRDailyRates(Date, ref EURRUBCurrency, ref USDRUBCurrency);
            //if (USDRUBCurrency != 0)
            //    EURUSDCurrency = Decimal.Round(EURRUBCurrency / USDRUBCurrency, 4, MidpointRounding.AwayFromZero);

            OrdersManager.GetDateRates(Date, ref RateExist, ref EURUSDCurrency, ref EURRUBCurrency, ref EURBYRCurrency, ref USDRUBCurrency);

            if (!RateExist)
            {
                OrdersManager.CBRDailyRates(Date, ref EURRUBCurrency, ref USDRUBCurrency);
                EURBYRCurrency = CurrencyConverter.NbrbDailyRates(DateTime.Now);
                //OrdersManager.NBRBDailyRates(Date, ref EURBYRCurrency);

                if (USDRUBCurrency != 0)
                    EURUSDCurrency = Decimal.Round(EURRUBCurrency / USDRUBCurrency, 4, MidpointRounding.AwayFromZero);
                if (EURBYRCurrency == 1000000)
                {

                    Infinium.LightMessageBox.Show(ref TopForm, false,
                        "Курс не взят. Возможная причина - нет связи с банком без авторизации в kerio control. Войдите в kerio и повторите попытку",
                        "Расчет заказа");
                    return;
                }
                OrdersManager.SaveDateRates(Date, EURUSDCurrency, EURRUBCurrency, EURBYRCurrency, USDRUBCurrency);
            }

            CurrencyEURRUBTextBox.Text = EURRUBCurrency.ToString();
            CurrencyUSDRUBTextBox.Text = USDRUBCurrency.ToString();
            CurrencyEURUSDTextBox.Text = EURUSDCurrency.ToString();
            CurrencyEURBYRTextBox.Text = EURBYRCurrency.ToString();

        }

        private void CurrencyDateTimePicker_ValueChanged(object sender, EventArgs e)
        {
            GetCurrency(CurrencyDateTimePicker.Value.Date);
        }

        private void CurrencyTodayButton_Click(object sender, EventArgs e)
        {
            CurrencyDateTimePicker.Value = DateTime.Now;
        }

        private void MegaOrdersDataGrid_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Right)
            {
                CurrentRowIndex = e.RowIndex;
                //OrdersManager.MegaOrdersBindingSource.Position = CurrentRowIndex;

                kryptonContextMenu1.Show(new Point(Cursor.Position.X - 212, Cursor.Position.Y - 10));
            }
        }

        private void MainOrdersDataGrid_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Right)
            {
                CurrentRowIndex = e.RowIndex;
                //OrdersManager.MegaOrdersBindingSource.Position = CurrentRowIndex;

                kryptonContextMenu2.Show(new Point(Cursor.Position.X - 212, Cursor.Position.Y - 10));
            }
        }

        private void MegaOrderDecorInProd_Click(object sender, EventArgs e)
        {
            var T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Создание отчета.\r\nПодождите..."); });
            T.Start();

            while (!SplashWindow.bSmallCreated) ;
            NeedSplash = false;

            var MarketingBatchReport = new BatchReport(ref DecorCatalogOrder);
            var MainOrders = OrdersManager.GetMainOrders();
            MarketingBatchReport.CreateReportForMaketing(MainOrders);

            NeedSplash = true;
            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
        }

        private void MainOrderDecorInProd_Click(object sender, EventArgs e)
        {
            var T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Создание отчета.\r\nПодождите..."); });
            T.Start();

            while (!SplashWindow.bSmallCreated) ;
            NeedSplash = false;

            var MarketingBatchReport = new BatchReport(ref DecorCatalogOrder);
            var MainOrders = OrdersManager.GetSelectedMainOrders();
            MarketingBatchReport.CreateReportForMaketing(MainOrders);

            NeedSplash = true;
            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
        }

        private void MinimizeButton_Click(object sender, EventArgs e)
        {
            NeedSplash = false;
            FormEvent = eHide;
            AnimateTimer.Enabled = true;
        }

        private void cbOnExpedition_CheckedChanged(object sender, EventArgs e)
        {
            if (OrdersManager == null)
                return;
            Filter();
        }

        private void kryptonContextMenuItem1_Click(object sender, EventArgs e)
        {
        }

        private void kryptonContextMenuItem3_Click(object sender, EventArgs e)
        {

        }

        private void kryptonContextMenu2_Opening(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (RoleType == RoleTypes.Ordinary)
                e.Cancel = true;
        }

        private void kryptonContextMenu1_Opening(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (RoleType == RoleTypes.Ordinary)
                e.Cancel = true;
        }

        private void cbDelayOfPayment_CheckedChanged(object sender, EventArgs e)
        {
            if (OrdersManager == null)
                return;
            Filter();
        }

        private void cbHalfOfPayment_CheckedChanged(object sender, EventArgs e)
        {
            if (OrdersManager == null)
                return;
            Filter();
        }

        private void cbFullPayment_CheckedChanged(object sender, EventArgs e)
        {
            if (OrdersManager == null)
                return;
            Filter();
        }

        private void cbFactoring_CheckedChanged(object sender, EventArgs e)
        {
            if (OrdersManager == null)
                return;
            Filter();
        }

        private void MegaOrdersDataGrid_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            if (e.RowIndex < 0)
                return;
            var grid = (PercentageDataGrid)sender;
            var ClientID = 0;
            var OrderNumber = 0;
            if (grid.Rows[e.RowIndex].Cells["ClientID"].Value != DBNull.Value)
                ClientID = Convert.ToInt32(grid.Rows[e.RowIndex].Cells["ClientID"].Value);
            if (grid.Rows[e.RowIndex].Cells["OrderNumber"].Value != DBNull.Value)
                OrderNumber = Convert.ToInt32(grid.Rows[e.RowIndex].Cells["OrderNumber"].Value);
            var IsOrderInMuttlement = OrdersManager.IsOrderInMuttlement(ClientID, OrderNumber);
            if (IsOrderInMuttlement)
            {
                grid.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Red;
                grid.Rows[e.RowIndex].DefaultCellStyle.ForeColor = Color.White;
                grid.Rows[e.RowIndex].DefaultCellStyle.SelectionForeColor = Color.MediumBlue;
            }
            else
            {
                grid.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.White;
                grid.Rows[e.RowIndex].DefaultCellStyle.ForeColor = Color.Black;
                grid.Rows[e.RowIndex].DefaultCellStyle.SelectionForeColor = Color.White;
            }
        }

        private void kryptonContextMenuItem14_Click(object sender, EventArgs e)
        {
            var T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
            T.Start();

            while (!SplashWindow.bSmallCreated) ;

            var MainOrders = new int[MainOrdersDataGrid.SelectedRows.Count];

            decimal ComplaintProfilCost = 0;
            decimal ComplaintTPSCost = 0;
            decimal TotalWeight = 0;
            decimal TransportCost = 0;
            decimal AdditionalCost = 0;
            var CurrencyTypeID = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["CurrencyTypeID"]);
            var Rate = Convert.ToDecimal(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["PaymentRate"]);

            for (var i = 0; i < MegaOrdersDataGrid.SelectedRows.Count; i++)
            {
                ComplaintProfilCost += Convert.ToDecimal(MegaOrdersDataGrid.SelectedRows[i].Cells["CurrencyComplaintProfilCost"].Value);
                ComplaintTPSCost += Convert.ToDecimal(MegaOrdersDataGrid.SelectedRows[i].Cells["CurrencyComplaintTPSCost"].Value);
                TotalWeight += Convert.ToDecimal(MegaOrdersDataGrid.SelectedRows[i].Cells["Weight"].Value);
                TransportCost += Convert.ToDecimal(MegaOrdersDataGrid.SelectedRows[i].Cells["CurrencyTransportCost"].Value);
                AdditionalCost += Convert.ToDecimal(MegaOrdersDataGrid.SelectedRows[i].Cells["CurrencyAdditionalCost"].Value);
            }

            //ComplaintProfilCost = Convert.ToDecimal(MegaOrdersDataGrid.SelectedRows[0].Cells["CurrencyComplaintProfilCost"].Value);
            //ComplaintTPSCost = Convert.ToDecimal(MegaOrdersDataGrid.SelectedRows[0].Cells["CurrencyComplaintTPSCost"].Value);
            //TransportCost = Convert.ToDecimal(MegaOrdersDataGrid.SelectedRows[0].Cells["CurrencyTransportCost"].Value);
            //AdditionalCost = Convert.ToDecimal(MegaOrdersDataGrid.SelectedRows[0].Cells["CurrencyAdditionalCost"].Value);

            for (var i = 0; i < MainOrdersDataGrid.SelectedRows.Count; i++)
                MainOrders[i] = Convert.ToInt32(MainOrdersDataGrid.SelectedRows[i].Cells["MainOrderID"].Value);

            var SelectedMegaOrders = new int[1];
            var SelectedOrderNumbers = new int[1];
            SelectedMegaOrders[0] = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["MegaOrderID"]);
            SelectedOrderNumbers[0] = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["OrderNumber"]);

            var ClientID = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["ClientID"]);
            var ClientName = OrdersManager.GetClientName(ClientID);

            NeedSplash = false;

            var DBFReport = new ColorInvoiceReportToDbf();

            decimal TotalProfil = 0;
            decimal TotalTPS = 0;

            var ProfilDBFName = string.Empty;
            var TPSDBFName = string.Empty;
            var ReportName = string.Empty;
            var ClientCountryID = OrdersManager.GetClientCountry(ClientID);

            ClientName = ClientName.Replace("\"", " ");
            ClientName = ClientName.Replace("*", " ");
            ClientName = ClientName.Replace("|", " ");
            ClientName = ClientName.Replace(@"\", " ");
            ClientName = ClientName.Replace(":", " ");
            ClientName = ClientName.Replace("<", " ");
            ClientName = ClientName.Replace(">", " ");
            ClientName = ClientName.Replace("?", " ");
            ClientName = ClientName.Replace("/", " ");
            ReportName = ClientName + " №" + string.Join(",", SelectedOrderNumbers);

            var hssfworkbook = new HSSFWorkbook();

            DBFReport.CreateReport(ref hssfworkbook, SelectedOrderNumbers, SelectedMegaOrders, MainOrders,
                ClientID, ClientName, ComplaintProfilCost, ComplaintTPSCost, TransportCost, AdditionalCost,
                OrdersManager.GetMainOrdersCost(MainOrders), CurrencyTypeID, TotalWeight, ref TotalProfil, ref TotalTPS);
            //DBFReport.SaveDBF(Security.DBFSaveFilePath, ReportName, ref ProfilDBFName, ref TPSDBFName);

            DetailsReport.Report(ref hssfworkbook, SelectedMegaOrders, SelectedOrderNumbers, MainOrders, ClientID, ClientName,
                ComplaintProfilCost, ComplaintTPSCost, TransportCost, AdditionalCost,
                OrdersManager.GetMainOrdersCost(MainOrders), CurrencyTypeID, TotalWeight);

            if (ClientID == 609)
            {
                DetailsReport.ReportModuleOnline(ref hssfworkbook, MainOrders);
            }
            var tempFolder = System.Environment.GetEnvironmentVariable("TEMP");
            var file = new FileInfo(tempFolder + @"\" + ReportName + ".xls");
            var j = 1;
            while (file.Exists == true)
            {
                file = new FileInfo(tempFolder + @"\" + ReportName + "(" + j++ + ").xls");
            }

            var NewFile = new FileStream(file.FullName, FileMode.Create);
            hssfworkbook.Write(NewFile);
            NewFile.Close();

            System.Diagnostics.Process.Start(file.FullName);
            ReportName = file.FullName;

            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
            NeedSplash = true;

        }

        private void MainOrdersFrontsOrdersDataGrid_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Right)
            {
                var FrontsOrdersID = Convert.ToInt32(MainOrdersFrontsOrdersDataGrid.SelectedRows[0].Cells["FrontsOrdersID"].Value);
                var kryptonContextMenu = new ComponentFactory.Krypton.Toolkit.KryptonContextMenu(); 
                var kryptonContextMenuItems = new ComponentFactory.Krypton.Toolkit.KryptonContextMenuItems();
                var setOnStorage = new ComponentFactory.Krypton.Toolkit.KryptonContextMenuItem()
                {
                    Text = "На склад"
                };
                setOnStorage.Click += (o, args) =>
                {
                    OrdersManager.MainOrdersFrontsOrders.SetOnStorage(FrontsOrdersID);

                    var onStorage = Convert.ToBoolean(MainOrdersFrontsOrdersDataGrid.SelectedRows[0].Cells["onStorage"].Value);
                    MainOrdersFrontsOrdersDataGrid.SelectedRows[0].Cells["onStorage"].Value = !onStorage;
                };

                var MainOrderID = Convert.ToInt32(MainOrdersDataGrid.SelectedRows[0].Cells["MainOrderID"].Value);
                var FrontID = Convert.ToInt32(MainOrdersFrontsOrdersDataGrid.SelectedRows[0].Cells["FrontID"].Value);

                //if (TestTechCatalogManager == null)
                //{
                //    TestTechCatalogManager = new Modules.TechnologyCatalog.TestTechCatalog();
                //    TestTechCatalogManager.Initialize();
                //}
                //var DT = TestTechCatalogManager.GetOperationsGroups(FrontID);
                //if (DT.Rows.Count > 0)
                //{
                //    for (var i = 0; i < DT.Rows.Count; i++)
                //    {
                //        var kryptonContextMenuItem = new ComponentFactory.Krypton.Toolkit.KryptonContextMenuItem()
                //        {
                //            Text = DT.Rows[i]["GroupName"].ToString(),
                //            Tag = Convert.ToInt32(DT.Rows[i]["TechCatalogOperationsGroupID"])
                //        };
                //        kryptonContextMenuItem.Click += kryptonContextMenuItem_Click;
                //        kryptonContextMenuItems.Items.Add(kryptonContextMenuItem);
                //    }
                //}
                kryptonContextMenuItems.Items.Add(setOnStorage);
                kryptonContextMenu.Items.Add(kryptonContextMenuItems);
                kryptonContextMenu.Show(new Point(Cursor.Position.X - 212, Cursor.Position.Y - 10));
            }
        }

        private void kryptonContextMenuItem_Click(object sender, EventArgs e)
        {
            TestTechCatalogManager.GetFrontsOrders(
                Convert.ToInt32(MainOrdersFrontsOrdersDataGrid.SelectedRows[0].Cells["FrontsOrdersID"].Value),
                Convert.ToInt32(((ComponentFactory.Krypton.Toolkit.KryptonContextMenuItem)sender).Tag));
            TestTechCatalogManager.ReturnResultTable();

            TestTechCatalogManager.CreateFrontsExcel(MainOrdersFrontsOrdersDataGrid.SelectedRows[0].Cells["FrontsOrdersID"].Value.ToString(),
                Convert.ToInt32(((ComponentFactory.Krypton.Toolkit.KryptonContextMenuItem)sender).Tag));
        }

        private bool Func1(DataTable DT1, int TechCatalogOperationsDetailID,
            ref ComponentFactory.Krypton.Toolkit.KryptonContextMenu kryptonContextMenu,
            ref ComponentFactory.Krypton.Toolkit.KryptonContextMenuItem kryptonContextMenuItem)
        {
            var filter = "PrevTechCatalogOperationsDetailID=" + TechCatalogOperationsDetailID;
            var rows1 = DT1.Select(filter);
            if (rows1.Count() == 0)
                return false;

            var DT = DT1.Clone();
            foreach (var item in rows1)
                DT.Rows.Add(item.ItemArray);
            var kryptonContextMenuItems = new ComponentFactory.Krypton.Toolkit.KryptonContextMenuItems();
            kryptonContextMenuItem.Items.Add(kryptonContextMenuItems);
            Func2(DT1, DT, TechCatalogOperationsDetailID, ref kryptonContextMenu, ref kryptonContextMenuItems);
            return true;
        }

        private void Func2(DataTable DT1, DataTable DT2, int TechCatalogOperationsDetailID, ref ComponentFactory.Krypton.Toolkit.KryptonContextMenu kryptonContextMenu,
            ref ComponentFactory.Krypton.Toolkit.KryptonContextMenuItems kryptonContextMenuItems)
        {
            for (var j = 0; j < DT2.Rows.Count; j++)
            {
                TechCatalogOperationsDetailID = Convert.ToInt32(DT2.Rows[j]["TechCatalogOperationsDetailID"]);
                var GroupName = DT2.Rows[j]["GroupName"].ToString();
                var kryptonContextMenuItem = new ComponentFactory.Krypton.Toolkit.KryptonContextMenuItem();

                if (Func1(DT1, TechCatalogOperationsDetailID, ref kryptonContextMenu, ref kryptonContextMenuItem))
                {

                }
                kryptonContextMenuItem.Text = GroupName;
                kryptonContextMenuItem.Tag = TechCatalogOperationsDetailID;
                kryptonContextMenuItems.Items.Add(kryptonContextMenuItem);
            }
        }

        private void kryptonCheckBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (OrdersManager == null)
                return;
            Filter();
        }

        private void kryptonCheckBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (OrdersManager == null)
                return;
            Filter();
        }

        private void MegaOrdersDataGrid_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (MegaOrdersDataGrid.Columns[e.ColumnIndex].Name == "PlanDispDate")
            {
                OrdersManager.UpdatemegaOrders();
                Filter();
            }
        }

        private void btnSetAgreementStatus_Click(object sender, EventArgs e)
        {
            var PhantomForm = new Infinium.PhantomForm();
            PhantomForm.Show();

            var AgreementStatusID = 0;
            var MegaOrderID = 0;
            if (OrdersManager.MegaOrdersBindingSource.Count > 0)
                MegaOrderID = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["MegaOrderID"]);
            if (OrdersManager.MegaOrdersBindingSource.Count > 0)
                AgreementStatusID = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["AgreementStatusID"]);

            var SetOrderStatusForm = new SetOrderStatusForm(this, AgreementStatusID);

            TopForm = SetOrderStatusForm;
            SetOrderStatusForm.ShowDialog();
            AgreementStatusID = SetOrderStatusForm.AgreementStatusID;

            PhantomForm.Close();

            PhantomForm.Dispose();
            SetOrderStatusForm.Dispose();
            TopForm = null;

            if (AgreementStatusID > -1)
            {
                OrdersManager.SetAgreementStatus(AgreementStatusID);
                NeedSplash = false;
                Filter();
                OrdersManager.MoveToMegaOrder(MegaOrderID);
                NeedSplash = true;
            }
        }

        private void btnDeleteAgreedOrder_Click(object sender, EventArgs e)
        {
            if (OrdersManager.MegaOrdersBindingSource.Count > 0)
            {
                var MegaOrderID = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current)["MegaOrderID"]);

                var OKCancel = Infinium.LightMessageBox.Show(ref TopForm, true,
                        "Вы действительно хотите удалить заказ " + MegaOrderID + " ?",
                        "Удаление заказа");
                if (!OKCancel)
                    return;

                NeedSplash = false;
                var T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                T.Start();

                while (!SplashWindow.bSmallCreated) ;

                OrdersManager.RemoveCurrentAgreedMegaOrder();
                OrdersManager.FixOrderEvent(MegaOrderID, "Согласованный заказ удален");

                if (OrdersManager.MegaOrdersCount == 0)
                {
                    OrdersManager.CurrentMegaOrderID = -1;
                }
                Filter();

                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;
                NeedSplash = true;
            }
        }

        private void btnDeleteAgreedMainOrder_Click(object sender, EventArgs e)
        {
            if (OrdersManager.MainOrdersBindingSource.Count == 0)
                return;

            var MainOrderID = ((DataRowView)OrdersManager.MainOrdersBindingSource.Current)["MainOrderID"].ToString();
            var MegaOrderID = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["MegaOrderID"]);

            var OKCancel = Infinium.LightMessageBox.Show(ref TopForm, true,
                    "Вы действительно хотите удалить подзаказ " + MainOrderID + " ?",
                    "Удаление подзаказа");
            if (!OKCancel)
                return;

            NeedSplash = false;

            var T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
            T.Start();

            while (!SplashWindow.bSmallCreated) ;

            OrdersManager.RemoveCurrentAgreedMainOrder();

            var OnProduction = OnProductionCheckBox.Checked;
            var NotInProduction = NotProductionCheckBox.Checked;
            var InProduction = InProductionCheckBox.Checked;
            var OnStorage = OnStorageCheckBox.Checked;
            var OnExpedition = cbOnExpedition.Checked;
            var Dispatch = DispatchCheckBox.Checked;

            if (OrdersManager.MainOrdersBindingSource.Count > 0)
                OrdersManager.FilterMainOrdersByMegaOrder(
                    Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["MegaOrderID"]),
                    OnProduction,
                    NotInProduction,
                    InProduction,
                    OnStorage,
                    OnExpedition,
                    true);

            OrdersManager.SummaryCost(Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["MegaOrderID"]));
            OrdersManager.FixOrderEvent(Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["MegaOrderID"]), "Удален согласованный подзаказ");
            Filter();
            OrdersManager.MoveToMegaOrder(MegaOrderID);

            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
            NeedSplash = true;
        }

        private void btnEditAgreedMainOrder_Click(object sender, EventArgs e)
        {
            OrdersManager.NeedSetStatus = true;
            if (OrdersManager.MainOrdersBindingSource.Count > 0)
            {
                NeedSplash = false;

                OrdersManager.EditAgreedMainOrder();

                var T = new Thread(delegate () { SplashWindow.CreateSplash(); });
                T.Start();

                while (!SplashForm.bCreated) ;

                if (OrdersManager.MegaOrdersBindingSource.Count > 0)
                {
                    OrdersManager.CurrentClientID = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["ClientID"]);
                    OrdersManager.CurrentDiscountPaymentConditionID = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["DiscountPaymentConditionID"]);
                    OrdersManager.CurrentDiscountFactoringID = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["DiscountFactoringID"]);
                    OrdersManager.CurrentProfilDiscountDirector = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["ProfilDiscountDirector"]);
                    OrdersManager.CurrentTPSDiscountDirector = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["TPSDiscountDirector"]);
                    OrdersManager.CurrentProfilTotalDiscount = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["ProfilTotalDiscount"]) - Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["ProfilDiscountDirector"]);
                    OrdersManager.CurrentTPSTotalDiscount = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["TPSTotalDiscount"]) - Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["TPSDiscountDirector"]);
                    OrdersManager.CurrencyTypeID = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["CurrencyTypeID"]);
                    OrdersManager.PaymentCurrency = Convert.ToDecimal(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["PaymentRate"]);
                    OrdersManager.ConfirmDateTime = ((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["ConfirmDateTime"];
                }

                var ClientOMC = OrdersManager.CurrentClientID == 232;

                AddMainOrdersForm = new AddMarketingNewOrdersForm(ref OrdersManager, ClientOMC, true, true, ref TopForm, ref OrdersCalculate);

                TopForm = AddMainOrdersForm;
                AddMainOrdersForm.ShowDialog();

                AddMainOrdersForm.Close();
                AddMainOrdersForm.Dispose();

                TopForm = null;

                var OnProduction = OnProductionCheckBox.Checked;
                var NotInProduction = NotProductionCheckBox.Checked;
                var InProduction = InProductionCheckBox.Checked;
                var OnStorage = OnStorageCheckBox.Checked;
                var OnExpedition = cbOnExpedition.Checked;
                var Dispatch = DispatchCheckBox.Checked;

                NeedSplash = false;
                Filter();

                OrdersManager.UpdateMainOrders(
                    OnProduction,
                    NotInProduction,
                    InProduction,
                    OnStorage,
                    OnExpedition,
                    Dispatch);
                NeedSplash = true;
            }
            OrdersManager.CurrentClientID = -1;
            OrdersManager.CurrentDiscountPaymentConditionID = 0;
            OrdersManager.CurrentDiscountFactoringID = 0;
            OrdersManager.CurrentProfilDiscountDirector = 0;
            OrdersManager.CurrentTPSDiscountDirector = 0;
            OrdersManager.CurrentProfilTotalDiscount = 0;
            OrdersManager.CurrentTPSTotalDiscount = 0;
            OrdersManager.PaymentCurrency = 1;
            OrdersManager.CurrencyTypeID = 1;
            OrdersManager.ConfirmDateTime = DBNull.Value;
        }

        private void kryptonContextMenuItem25_Click(object sender, EventArgs e)
        {
            if (OrdersManager.MainOrdersBindingSource.Count == 0)
                return;
            if (MegaOrdersDataGrid.SelectedRows.Count > 0)
            {
                var d = Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[0].Cells["DiscountPaymentConditionID"].Value);
                for (var i = 0; i < MegaOrdersDataGrid.SelectedRows.Count; i++)
                {
                    if (d != Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[i].Cells["DiscountPaymentConditionID"].Value))
                    {
                        Infinium.LightMessageBox.Show(ref TopForm, false,
                               "Нельзя объединить заказы с различными условиями оплаты",
                               "Создание отчета");
                        return;
                    }
                }
            }
            if (MegaOrdersDataGrid.SelectedRows.Count > 0)
            {
                var c = Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[0].Cells["CurrencyTypeID"].Value);
                for (var i = 0; i < MegaOrdersDataGrid.SelectedRows.Count; i++)
                {
                    if (c != Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[i].Cells["CurrencyTypeID"].Value))
                    {
                        Infinium.LightMessageBox.Show(ref TopForm, false,
                               "Нельзя объединить заказы с различными валютами",
                               "Создание отчета");
                        return;
                    }
                }
            }

            if (!OrdersManager.AreSelectedMegaOrdersOneClient)
            {
                Infinium.LightMessageBox.Show(ref TopForm, false,
                       "Выбраны заказы разных клиентов",
                       "Создание отчета");
                return;
            }

            if (!OrdersManager.CanBeReported())
            {
                Infinium.LightMessageBox.Show(ref TopForm, false,
                       "Заказ не был согласован. Создание отчета невозможно.",
                       "Создание отчета");
                return;
            }

            var dbfReport = new Modules.Marketing.NewOrders.PrepareReport.NotesInvoiceReportToDbf.NotesInvoiceReportToDbf();
            PrepareInvoiceReportToDBF(dbfReport);

            //using (var form = new ReportTypeForm(this))
            //{
            //    var result = form.ShowDialog();
            //    if (result == DialogResult.OK)
            //    {
            //        var invoiceReportType = form.invoiceReportType;

            //        if (invoiceReportType == InvoiceReportType.Notes)
            //        {
            //            var dbfReport = new Modules.Marketing.NewOrders.PrepareReport.NotesInvoiceReportToDbf.NotesInvoiceReportToDbf();
            //            PrepareInvoiceReportToDBF(dbfReport);
            //        }
            //        if (invoiceReportType == InvoiceReportType.Standard)
            //        {
            //            var dbfReport = new Modules.Marketing.NewOrders.PrepareReport.InvoiceReportToDbf.InvoiceReportToDbf();
            //            PrepareInvoiceReportToDBF(dbfReport);
            //        }
            //        if (invoiceReportType == InvoiceReportType.CvetPatina)
            //        {
            //            var dbfReport = new Modules.Marketing.NewOrders.PrepareReport.ColorInvoiceReportToDbf.ColorInvoiceReportToDbf();
            //            PrepareInvoiceReportToDBF(dbfReport);
            //        }
            //    }
            //    else
            //    {
            //        return;
            //    }
            //}
        }

        private void PrepareInvoiceReportToDBF(Modules.Marketing.NewOrders.PrepareReport.NotesInvoiceReportToDbf.NotesInvoiceReportToDbf DBFReport)
        {
            var CurrencyTypeID = Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[0].Cells["CurrencyTypeID"].Value);
            var DiscountPaymentConditionID = Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[0].Cells["DiscountPaymentConditionID"].Value);
            var DelayOfPayment = Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[0].Cells["DelayOfPayment"].Value);

            decimal ComplaintProfilCost = 0;
            decimal ComplaintTPSCost = 0;
            decimal TransportCost = 0;
            decimal AdditionalCost = 0;
            decimal TotalWeight = 0;
            CurrencyTypeID = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["CurrencyTypeID"]);
            var Rate = Convert.ToDecimal(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["PaymentRate"]);
            var OrderNumber = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["OrderNumber"]);

            for (var i = 0; i < MegaOrdersDataGrid.SelectedRows.Count; i++)
            {
                ComplaintProfilCost += Convert.ToDecimal(MegaOrdersDataGrid.SelectedRows[i].Cells["CurrencyComplaintProfilCost"].Value);
                ComplaintTPSCost += Convert.ToDecimal(MegaOrdersDataGrid.SelectedRows[i].Cells["CurrencyComplaintTPSCost"].Value);
                TotalWeight += Convert.ToDecimal(MegaOrdersDataGrid.SelectedRows[i].Cells["Weight"].Value);
                TransportCost += Convert.ToDecimal(MegaOrdersDataGrid.SelectedRows[i].Cells["CurrencyTransportCost"].Value);
                AdditionalCost += Convert.ToDecimal(MegaOrdersDataGrid.SelectedRows[i].Cells["CurrencyAdditionalCost"].Value);
            }

            var SelectedMegaOrders = OrdersManager.GetSelectedMegaOrders();
            var SelectedOrderNumbers = OrdersManager.GetSelectedOrderNumbers();
            var SelectedProfilOrderNumbers = new int[MegaOrdersDataGrid.SelectedRows.Count];
            var SelectedTPSOrderNumbers = new int[MegaOrdersDataGrid.SelectedRows.Count];
            var MainOrders = OrdersManager.GetMainOrders(SelectedMegaOrders);

            for (var i = 0; i < MegaOrdersDataGrid.SelectedRows.Count; i++)
            {
                if (Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[i].Cells["FactoryID"].Value) == 1 || Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[i].Cells["FactoryID"].Value) == 0)
                    SelectedProfilOrderNumbers[i] = Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[i].Cells["OrderNumber"].Value);
                if (Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[i].Cells["FactoryID"].Value) == 2 || Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[i].Cells["FactoryID"].Value) == 0)
                    SelectedTPSOrderNumbers[i] = Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[i].Cells["OrderNumber"].Value);
            }
            SelectedProfilOrderNumbers = SelectedProfilOrderNumbers.Where(val => val != 0).ToArray();
            SelectedTPSOrderNumbers = SelectedTPSOrderNumbers.Where(val => val != 0).ToArray();
            var ClientID = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["ClientID"]);
            var ClientName = OrdersManager.GetClientName(ClientID);

            NeedSplash = false;

            var MutualSettlementsManager = new Modules.Marketing.MutualSettlements.MutualSettlements();
            MutualSettlementsManager.Initialize();

            if (ClientID == 145)
            {
                var HasOrdinaryOrders = OrdersManager.HasOrders(MainOrders, false);
                var HasSampleOrders = OrdersManager.HasOrders(MainOrders, true);

                if (HasOrdinaryOrders)
                {
                    decimal TotalProfil = 0;
                    decimal TotalTPS = 0;

                    var ProfilDBFName = string.Empty;
                    var TPSDBFName = string.Empty;
                    var ReportName = string.Empty;
                    var ClientCountryID = OrdersManager.GetClientCountry(ClientID);

                    ClientName = ClientName.Replace("\"", " ");
                    ClientName = ClientName.Replace("*", " ");
                    ClientName = ClientName.Replace("|", " ");
                    ClientName = ClientName.Replace(@"\", " ");
                    ClientName = ClientName.Replace(":", " ");
                    ClientName = ClientName.Replace("<", " ");
                    ClientName = ClientName.Replace(">", " ");
                    ClientName = ClientName.Replace("?", " ");
                    ClientName = ClientName.Replace("/", " ");
                    ReportName = ClientName + " №" + string.Join(",", SelectedOrderNumbers);

                    var PhantomForm = new Infinium.PhantomForm();
                    PhantomForm.Show();
                    var InMutualSettlement = false;
                    var Result = 1;
                    var Notes = string.Empty;
                    var SaveFilePath = string.Empty;
                    SaveDBFReportMenu = new Infinium.SaveDBFReportMenu(this, ReportName);

                    TopForm = SaveDBFReportMenu;
                    SaveDBFReportMenu.ShowDialog();

                    PhantomForm.Close();
                    InMutualSettlement = SaveDBFReportMenu.InMutualSettlement;
                    Result = SaveDBFReportMenu.Result;
                    Notes = SaveDBFReportMenu.Notes;
                    SaveFilePath = SaveDBFReportMenu.SaveFilePath;
                    PhantomForm.Dispose();
                    SaveDBFReportMenu.Dispose();
                    TopForm = null;

                    if (Result != 3)
                    {
                        var T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                        T.Start();

                        while (!SplashWindow.bSmallCreated) ;

                        var hssfworkbook = new HSSFWorkbook();

                        DBFReport.CreateReport(ref hssfworkbook, SelectedMegaOrders, SelectedOrderNumbers, MainOrders,
                            ClientID, ClientName, ComplaintProfilCost, ComplaintTPSCost, TransportCost, AdditionalCost,
                            OrdersManager.GetMainOrdersCost(MainOrders), CurrencyTypeID, TotalWeight, false, ref TotalProfil, ref TotalTPS);
                        DBFReport.SaveDBF(SaveFilePath, ReportName, false, ref ProfilDBFName, ref TPSDBFName);

                        PrepareReport.Report(ref hssfworkbook, SelectedMegaOrders, SelectedOrderNumbers, MainOrders, ClientID, ClientName,
                            ComplaintProfilCost, ComplaintTPSCost, TransportCost, AdditionalCost,
                            OrdersManager.GetMainOrdersCost(MainOrders), CurrencyTypeID, TotalWeight, false);

                        var file = new FileInfo(SaveFilePath + @"\" + ReportName + ".xls");
                        var j = 1;
                        while (file.Exists == true)
                        {
                            file = new FileInfo(SaveFilePath + @"\" + ReportName + "(" + j++ + ").xls");
                        }

                        var NewFile = new FileStream(file.FullName, FileMode.Create);
                        hssfworkbook.Write(NewFile);
                        NewFile.Close();

                        while (SplashWindow.bSmallCreated)
                            SmallWaitForm.CloseS = true;
                        NeedSplash = true;
                    }
                }
                if (HasSampleOrders)
                {
                    decimal TotalProfil = 0;
                    decimal TotalTPS = 0;

                    var ProfilDBFName = string.Empty;
                    var TPSDBFName = string.Empty;
                    var ReportName = string.Empty;
                    var ClientCountryID = OrdersManager.GetClientCountry(ClientID);

                    ClientName = ClientName.Replace("\"", " ");
                    ClientName = ClientName.Replace("*", " ");
                    ClientName = ClientName.Replace("|", " ");
                    ClientName = ClientName.Replace(@"\", " ");
                    ClientName = ClientName.Replace(":", " ");
                    ClientName = ClientName.Replace("<", " ");
                    ClientName = ClientName.Replace(">", " ");
                    ClientName = ClientName.Replace("?", " ");
                    ClientName = ClientName.Replace("/", " ");
                    ReportName = ClientName + " №" + string.Join(",", SelectedOrderNumbers) + ", обр";

                    var PhantomForm = new Infinium.PhantomForm();
                    PhantomForm.Show();
                    var InMutualSettlement = false;
                    var Result = 1;
                    var Notes = string.Empty;
                    var SaveFilePath = string.Empty;
                    SaveDBFReportMenu = new Infinium.SaveDBFReportMenu(this, ReportName);

                    TopForm = SaveDBFReportMenu;
                    SaveDBFReportMenu.ShowDialog();

                    PhantomForm.Close();
                    InMutualSettlement = SaveDBFReportMenu.InMutualSettlement;
                    Result = SaveDBFReportMenu.Result;
                    Notes = SaveDBFReportMenu.Notes;
                    SaveFilePath = SaveDBFReportMenu.SaveFilePath;
                    PhantomForm.Dispose();
                    SaveDBFReportMenu.Dispose();
                    TopForm = null;

                    if (Result != 3)
                    {
                        var T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                        T.Start();

                        while (!SplashWindow.bSmallCreated) ;

                        var hssfworkbook = new HSSFWorkbook();

                        DBFReport.CreateReport(ref hssfworkbook, SelectedMegaOrders, SelectedOrderNumbers, MainOrders,
                            ClientID, ClientName, ComplaintProfilCost, ComplaintTPSCost, TransportCost, AdditionalCost,
                            OrdersManager.GetMainOrdersCost(MainOrders), CurrencyTypeID, TotalWeight, true, ref TotalProfil, ref TotalTPS);
                        DBFReport.SaveDBF(SaveFilePath, ReportName, true, ref ProfilDBFName, ref TPSDBFName);

                        PrepareReport.Report(ref hssfworkbook, SelectedMegaOrders, SelectedOrderNumbers, MainOrders, ClientID, ClientName,
                            ComplaintProfilCost, ComplaintTPSCost, TransportCost, AdditionalCost,
                            OrdersManager.GetMainOrdersCost(MainOrders), CurrencyTypeID, TotalWeight, true);

                        var file = new FileInfo(SaveFilePath + @"\" + ReportName + ".xls");
                        var j = 1;
                        while (file.Exists == true)
                        {
                            file = new FileInfo(SaveFilePath + @"\" + ReportName + "(" + j++ + ").xls");
                        }

                        var NewFile = new FileStream(file.FullName, FileMode.Create);
                        hssfworkbook.Write(NewFile);
                        NewFile.Close();

                        while (SplashWindow.bSmallCreated)
                            SmallWaitForm.CloseS = true;
                        NeedSplash = true;
                    }
                }
            }
            else
            {
                decimal TotalProfil = 0;
                decimal TotalTPS = 0;

                var ProfilDBFName = string.Empty;
                var TPSDBFName = string.Empty;
                var ReportName = string.Empty;
                var ClientCountryID = OrdersManager.GetClientCountry(ClientID);

                ClientName = ClientName.Replace("\"", " ");
                ClientName = ClientName.Replace("*", " ");
                ClientName = ClientName.Replace("|", " ");
                ClientName = ClientName.Replace(@"\", " ");
                ClientName = ClientName.Replace(":", " ");
                ClientName = ClientName.Replace("<", " ");
                ClientName = ClientName.Replace(">", " ");
                ClientName = ClientName.Replace("?", " ");
                ClientName = ClientName.Replace("/", " ");
                ReportName = ClientName + " №" + string.Join(",", SelectedOrderNumbers);

                var PhantomForm = new Infinium.PhantomForm();
                PhantomForm.Show();
                var InMutualSettlement = false;
                var Result = 1;
                var Notes = string.Empty;
                var SaveFilePath = string.Empty;
                SaveDBFReportMenu = new Infinium.SaveDBFReportMenu(this, ReportName);

                TopForm = SaveDBFReportMenu;
                SaveDBFReportMenu.ShowDialog();

                PhantomForm.Close();
                InMutualSettlement = SaveDBFReportMenu.InMutualSettlement;
                Result = SaveDBFReportMenu.Result;
                Notes = SaveDBFReportMenu.Notes;
                SaveFilePath = SaveDBFReportMenu.SaveFilePath;
                PhantomForm.Dispose();
                SaveDBFReportMenu.Dispose();
                TopForm = null;

                if (Result != 3)
                {
                    var T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                    T.Start();

                    while (!SplashWindow.bSmallCreated) ;

                    var hssfworkbook = new HSSFWorkbook();

                    DBFReport.CreateReport(ref hssfworkbook, SelectedMegaOrders, SelectedOrderNumbers, MainOrders,
                        ClientID, ClientName, ComplaintProfilCost, ComplaintTPSCost, TransportCost, AdditionalCost,
                        OrdersManager.GetMainOrdersCost(MainOrders), CurrencyTypeID, TotalWeight, ref TotalProfil, ref TotalTPS);
                    DBFReport.SaveDBF(SaveFilePath, ReportName, ref ProfilDBFName, ref TPSDBFName);

                    PrepareReport.Report(ref hssfworkbook, SelectedMegaOrders, SelectedOrderNumbers, MainOrders, ClientID, ClientName,
                        ComplaintProfilCost, ComplaintTPSCost, TransportCost, AdditionalCost,
                        OrdersManager.GetMainOrdersCost(MainOrders), CurrencyTypeID, TotalWeight);

                    var file = new FileInfo(SaveFilePath + @"\" + ReportName + ".xls");
                    var j = 1;
                    while (file.Exists == true)
                    {
                        file = new FileInfo(SaveFilePath + @"\" + ReportName + "(" + j++ + ").xls");
                    }

                    var NewFile = new FileStream(file.FullName, FileMode.Create);
                    hssfworkbook.Write(NewFile);
                    NewFile.Close();
                    System.Diagnostics.Process.Start(file.FullName);

                    while (SplashWindow.bSmallCreated)
                        SmallWaitForm.CloseS = true;
                    NeedSplash = true;
                }
            }
        }
        
        private void PrepareInvoiceReportToDBF(Modules.Marketing.NewOrders.PrepareReport.ColorInvoiceReportToDbf.ColorInvoiceReportToDbf DBFReport)
        {
            var CurrencyTypeID = Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[0].Cells["CurrencyTypeID"].Value);
            var DiscountPaymentConditionID = Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[0].Cells["DiscountPaymentConditionID"].Value);
            var DelayOfPayment = Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[0].Cells["DelayOfPayment"].Value);

            decimal ComplaintProfilCost = 0;
            decimal ComplaintTPSCost = 0;
            decimal TransportCost = 0;
            decimal AdditionalCost = 0;
            decimal TotalWeight = 0;
            CurrencyTypeID = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["CurrencyTypeID"]);
            var Rate = Convert.ToDecimal(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["PaymentRate"]);
            var OrderNumber = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["OrderNumber"]);

            for (var i = 0; i < MegaOrdersDataGrid.SelectedRows.Count; i++)
            {
                ComplaintProfilCost += Convert.ToDecimal(MegaOrdersDataGrid.SelectedRows[i].Cells["CurrencyComplaintProfilCost"].Value);
                ComplaintTPSCost += Convert.ToDecimal(MegaOrdersDataGrid.SelectedRows[i].Cells["CurrencyComplaintTPSCost"].Value);
                TotalWeight += Convert.ToDecimal(MegaOrdersDataGrid.SelectedRows[i].Cells["Weight"].Value);
                TransportCost += Convert.ToDecimal(MegaOrdersDataGrid.SelectedRows[i].Cells["CurrencyTransportCost"].Value);
                AdditionalCost += Convert.ToDecimal(MegaOrdersDataGrid.SelectedRows[i].Cells["CurrencyAdditionalCost"].Value);
            }

            var SelectedMegaOrders = OrdersManager.GetSelectedMegaOrders();
            var SelectedOrderNumbers = OrdersManager.GetSelectedOrderNumbers();
            var SelectedProfilOrderNumbers = new int[MegaOrdersDataGrid.SelectedRows.Count];
            var SelectedTPSOrderNumbers = new int[MegaOrdersDataGrid.SelectedRows.Count];
            var MainOrders = OrdersManager.GetMainOrders(SelectedMegaOrders);

            for (var i = 0; i < MegaOrdersDataGrid.SelectedRows.Count; i++)
            {
                if (Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[i].Cells["FactoryID"].Value) == 1 || Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[i].Cells["FactoryID"].Value) == 0)
                    SelectedProfilOrderNumbers[i] = Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[i].Cells["OrderNumber"].Value);
                if (Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[i].Cells["FactoryID"].Value) == 2 || Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[i].Cells["FactoryID"].Value) == 0)
                    SelectedTPSOrderNumbers[i] = Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[i].Cells["OrderNumber"].Value);
            }
            SelectedProfilOrderNumbers = SelectedProfilOrderNumbers.Where(val => val != 0).ToArray();
            SelectedTPSOrderNumbers = SelectedTPSOrderNumbers.Where(val => val != 0).ToArray();
            var ClientID = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["ClientID"]);
            var ClientName = OrdersManager.GetClientName(ClientID);

            NeedSplash = false;

            var MutualSettlementsManager = new Modules.Marketing.MutualSettlements.MutualSettlements();
            MutualSettlementsManager.Initialize();

            if (ClientID == 145)
            {
                var HasOrdinaryOrders = OrdersManager.HasOrders(MainOrders, false);
                var HasSampleOrders = OrdersManager.HasOrders(MainOrders, true);

                if (HasOrdinaryOrders)
                {
                    decimal TotalProfil = 0;
                    decimal TotalTPS = 0;

                    var ProfilDBFName = string.Empty;
                    var TPSDBFName = string.Empty;
                    var ReportName = string.Empty;
                    var ClientCountryID = OrdersManager.GetClientCountry(ClientID);

                    ClientName = ClientName.Replace("\"", " ");
                    ClientName = ClientName.Replace("*", " ");
                    ClientName = ClientName.Replace("|", " ");
                    ClientName = ClientName.Replace(@"\", " ");
                    ClientName = ClientName.Replace(":", " ");
                    ClientName = ClientName.Replace("<", " ");
                    ClientName = ClientName.Replace(">", " ");
                    ClientName = ClientName.Replace("?", " ");
                    ClientName = ClientName.Replace("/", " ");
                    ReportName = ClientName + " №" + string.Join(",", SelectedOrderNumbers);

                    var PhantomForm = new Infinium.PhantomForm();
                    PhantomForm.Show();
                    var InMutualSettlement = false;
                    var Result = 1;
                    var Notes = string.Empty;
                    var SaveFilePath = string.Empty;
                    SaveDBFReportMenu = new Infinium.SaveDBFReportMenu(this, ReportName);

                    TopForm = SaveDBFReportMenu;
                    SaveDBFReportMenu.ShowDialog();

                    PhantomForm.Close();
                    InMutualSettlement = SaveDBFReportMenu.InMutualSettlement;
                    Result = SaveDBFReportMenu.Result;
                    Notes = SaveDBFReportMenu.Notes;
                    SaveFilePath = SaveDBFReportMenu.SaveFilePath;
                    PhantomForm.Dispose();
                    SaveDBFReportMenu.Dispose();
                    TopForm = null;

                    if (Result != 3)
                    {
                        var T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                        T.Start();

                        while (!SplashWindow.bSmallCreated) ;

                        var hssfworkbook = new HSSFWorkbook();

                        DBFReport.CreateReport(ref hssfworkbook, SelectedMegaOrders, SelectedOrderNumbers, MainOrders,
                            ClientID, ClientName, ComplaintProfilCost, ComplaintTPSCost, TransportCost, AdditionalCost,
                            OrdersManager.GetMainOrdersCost(MainOrders), CurrencyTypeID, TotalWeight, false, ref TotalProfil, ref TotalTPS);
                        DBFReport.SaveDBF(SaveFilePath, ReportName, false, ref ProfilDBFName, ref TPSDBFName);

                        PrepareReport.Report(ref hssfworkbook, SelectedMegaOrders, SelectedOrderNumbers, MainOrders, ClientID, ClientName,
                            ComplaintProfilCost, ComplaintTPSCost, TransportCost, AdditionalCost,
                            OrdersManager.GetMainOrdersCost(MainOrders), CurrencyTypeID, TotalWeight, false);

                        var file = new FileInfo(SaveFilePath + @"\" + ReportName + ".xls");
                        var j = 1;
                        while (file.Exists == true)
                        {
                            file = new FileInfo(SaveFilePath + @"\" + ReportName + "(" + j++ + ").xls");
                        }

                        var NewFile = new FileStream(file.FullName, FileMode.Create);
                        hssfworkbook.Write(NewFile);
                        NewFile.Close();

                        while (SplashWindow.bSmallCreated)
                            SmallWaitForm.CloseS = true;
                        NeedSplash = true;
                    }
                }
                if (HasSampleOrders)
                {
                    decimal TotalProfil = 0;
                    decimal TotalTPS = 0;

                    var ProfilDBFName = string.Empty;
                    var TPSDBFName = string.Empty;
                    var ReportName = string.Empty;
                    var ClientCountryID = OrdersManager.GetClientCountry(ClientID);

                    ClientName = ClientName.Replace("\"", " ");
                    ClientName = ClientName.Replace("*", " ");
                    ClientName = ClientName.Replace("|", " ");
                    ClientName = ClientName.Replace(@"\", " ");
                    ClientName = ClientName.Replace(":", " ");
                    ClientName = ClientName.Replace("<", " ");
                    ClientName = ClientName.Replace(">", " ");
                    ClientName = ClientName.Replace("?", " ");
                    ClientName = ClientName.Replace("/", " ");
                    ReportName = ClientName + " №" + string.Join(",", SelectedOrderNumbers) + ", обр";

                    var PhantomForm = new Infinium.PhantomForm();
                    PhantomForm.Show();
                    var InMutualSettlement = false;
                    var Result = 1;
                    var Notes = string.Empty;
                    var SaveFilePath = string.Empty;
                    SaveDBFReportMenu = new Infinium.SaveDBFReportMenu(this, ReportName);

                    TopForm = SaveDBFReportMenu;
                    SaveDBFReportMenu.ShowDialog();

                    PhantomForm.Close();
                    InMutualSettlement = SaveDBFReportMenu.InMutualSettlement;
                    Result = SaveDBFReportMenu.Result;
                    Notes = SaveDBFReportMenu.Notes;
                    SaveFilePath = SaveDBFReportMenu.SaveFilePath;
                    PhantomForm.Dispose();
                    SaveDBFReportMenu.Dispose();
                    TopForm = null;

                    if (Result != 3)
                    {
                        var T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                        T.Start();

                        while (!SplashWindow.bSmallCreated) ;

                        var hssfworkbook = new HSSFWorkbook();

                        DBFReport.CreateReport(ref hssfworkbook, SelectedMegaOrders, SelectedOrderNumbers, MainOrders,
                            ClientID, ClientName, ComplaintProfilCost, ComplaintTPSCost, TransportCost, AdditionalCost,
                            OrdersManager.GetMainOrdersCost(MainOrders), CurrencyTypeID, TotalWeight, true, ref TotalProfil, ref TotalTPS);
                        DBFReport.SaveDBF(SaveFilePath, ReportName, true, ref ProfilDBFName, ref TPSDBFName);

                        PrepareReport.Report(ref hssfworkbook, SelectedMegaOrders, SelectedOrderNumbers, MainOrders, ClientID, ClientName,
                            ComplaintProfilCost, ComplaintTPSCost, TransportCost, AdditionalCost,
                            OrdersManager.GetMainOrdersCost(MainOrders), CurrencyTypeID, TotalWeight, true);

                        var file = new FileInfo(SaveFilePath + @"\" + ReportName + ".xls");
                        var j = 1;
                        while (file.Exists == true)
                        {
                            file = new FileInfo(SaveFilePath + @"\" + ReportName + "(" + j++ + ").xls");
                        }

                        var NewFile = new FileStream(file.FullName, FileMode.Create);
                        hssfworkbook.Write(NewFile);
                        NewFile.Close();

                        while (SplashWindow.bSmallCreated)
                            SmallWaitForm.CloseS = true;
                        NeedSplash = true;
                    }
                }
            }
            else
            {
                decimal TotalProfil = 0;
                decimal TotalTPS = 0;

                var ProfilDBFName = string.Empty;
                var TPSDBFName = string.Empty;
                var ReportName = string.Empty;
                var ClientCountryID = OrdersManager.GetClientCountry(ClientID);

                ClientName = ClientName.Replace("\"", " ");
                ClientName = ClientName.Replace("*", " ");
                ClientName = ClientName.Replace("|", " ");
                ClientName = ClientName.Replace(@"\", " ");
                ClientName = ClientName.Replace(":", " ");
                ClientName = ClientName.Replace("<", " ");
                ClientName = ClientName.Replace(">", " ");
                ClientName = ClientName.Replace("?", " ");
                ClientName = ClientName.Replace("/", " ");
                ReportName = ClientName + " №" + string.Join(",", SelectedOrderNumbers);

                var PhantomForm = new Infinium.PhantomForm();
                PhantomForm.Show();
                var InMutualSettlement = false;
                var Result = 1;
                var Notes = string.Empty;
                var SaveFilePath = string.Empty;
                SaveDBFReportMenu = new Infinium.SaveDBFReportMenu(this, ReportName);

                TopForm = SaveDBFReportMenu;
                SaveDBFReportMenu.ShowDialog();

                PhantomForm.Close();
                InMutualSettlement = SaveDBFReportMenu.InMutualSettlement;
                Result = SaveDBFReportMenu.Result;
                Notes = SaveDBFReportMenu.Notes;
                SaveFilePath = SaveDBFReportMenu.SaveFilePath;
                PhantomForm.Dispose();
                SaveDBFReportMenu.Dispose();
                TopForm = null;

                if (Result != 3)
                {
                    var T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                    T.Start();

                    while (!SplashWindow.bSmallCreated) ;

                    var hssfworkbook = new HSSFWorkbook();

                    DBFReport.CreateReport(ref hssfworkbook, SelectedMegaOrders, SelectedOrderNumbers, MainOrders,
                        ClientID, ClientName, ComplaintProfilCost, ComplaintTPSCost, TransportCost, AdditionalCost,
                        OrdersManager.GetMainOrdersCost(MainOrders), CurrencyTypeID, TotalWeight, ref TotalProfil, ref TotalTPS);
                    DBFReport.SaveDBF(SaveFilePath, ReportName, ref ProfilDBFName, ref TPSDBFName);

                    PrepareReport.Report(ref hssfworkbook, SelectedMegaOrders, SelectedOrderNumbers, MainOrders, ClientID, ClientName,
                        ComplaintProfilCost, ComplaintTPSCost, TransportCost, AdditionalCost,
                        OrdersManager.GetMainOrdersCost(MainOrders), CurrencyTypeID, TotalWeight);

                    var file = new FileInfo(SaveFilePath + @"\" + ReportName + ".xls");
                    var j = 1;
                    while (file.Exists == true)
                    {
                        file = new FileInfo(SaveFilePath + @"\" + ReportName + "(" + j++ + ").xls");
                    }

                    var NewFile = new FileStream(file.FullName, FileMode.Create);
                    hssfworkbook.Write(NewFile);
                    NewFile.Close();
                    System.Diagnostics.Process.Start(file.FullName);

                    while (SplashWindow.bSmallCreated)
                        SmallWaitForm.CloseS = true;
                    NeedSplash = true;
                }
            }
        }

        private void PrepareInvoiceReportToDBF(Modules.Marketing.NewOrders.PrepareReport.InvoiceReportToDbf.InvoiceReportToDbf DBFReport)
        {
            var CurrencyTypeID = Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[0].Cells["CurrencyTypeID"].Value);
            var DiscountPaymentConditionID = Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[0].Cells["DiscountPaymentConditionID"].Value);
            var DelayOfPayment = Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[0].Cells["DelayOfPayment"].Value);

            decimal ComplaintProfilCost = 0;
            decimal ComplaintTPSCost = 0;
            decimal TransportCost = 0;
            decimal AdditionalCost = 0;
            decimal TotalWeight = 0;
            CurrencyTypeID = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["CurrencyTypeID"]);
            var Rate = Convert.ToDecimal(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["PaymentRate"]);
            var OrderNumber = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["OrderNumber"]);

            for (var i = 0; i < MegaOrdersDataGrid.SelectedRows.Count; i++)
            {
                ComplaintProfilCost += Convert.ToDecimal(MegaOrdersDataGrid.SelectedRows[i].Cells["CurrencyComplaintProfilCost"].Value);
                ComplaintTPSCost += Convert.ToDecimal(MegaOrdersDataGrid.SelectedRows[i].Cells["CurrencyComplaintTPSCost"].Value);
                TotalWeight += Convert.ToDecimal(MegaOrdersDataGrid.SelectedRows[i].Cells["Weight"].Value);
                TransportCost += Convert.ToDecimal(MegaOrdersDataGrid.SelectedRows[i].Cells["CurrencyTransportCost"].Value);
                AdditionalCost += Convert.ToDecimal(MegaOrdersDataGrid.SelectedRows[i].Cells["CurrencyAdditionalCost"].Value);
            }

            var SelectedMegaOrders = OrdersManager.GetSelectedMegaOrders();
            var SelectedOrderNumbers = OrdersManager.GetSelectedOrderNumbers();
            var SelectedProfilOrderNumbers = new int[MegaOrdersDataGrid.SelectedRows.Count];
            var SelectedTPSOrderNumbers = new int[MegaOrdersDataGrid.SelectedRows.Count];
            var MainOrders = OrdersManager.GetMainOrders(SelectedMegaOrders);

            for (var i = 0; i < MegaOrdersDataGrid.SelectedRows.Count; i++)
            {
                if (Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[i].Cells["FactoryID"].Value) == 1 || Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[i].Cells["FactoryID"].Value) == 0)
                    SelectedProfilOrderNumbers[i] = Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[i].Cells["OrderNumber"].Value);
                if (Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[i].Cells["FactoryID"].Value) == 2 || Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[i].Cells["FactoryID"].Value) == 0)
                    SelectedTPSOrderNumbers[i] = Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[i].Cells["OrderNumber"].Value);
            }
            SelectedProfilOrderNumbers = SelectedProfilOrderNumbers.Where(val => val != 0).ToArray();
            SelectedTPSOrderNumbers = SelectedTPSOrderNumbers.Where(val => val != 0).ToArray();
            var ClientID = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["ClientID"]);
            var ClientName = OrdersManager.GetClientName(ClientID);

            NeedSplash = false;

            var MutualSettlementsManager = new Modules.Marketing.MutualSettlements.MutualSettlements();
            MutualSettlementsManager.Initialize();

            if (ClientID == 145)
            {
                var HasOrdinaryOrders = OrdersManager.HasOrders(MainOrders, false);
                var HasSampleOrders = OrdersManager.HasOrders(MainOrders, true);

                if (HasOrdinaryOrders)
                {
                    decimal TotalProfil = 0;
                    decimal TotalTPS = 0;

                    var ProfilDBFName = string.Empty;
                    var TPSDBFName = string.Empty;
                    var ReportName = string.Empty;
                    var ClientCountryID = OrdersManager.GetClientCountry(ClientID);

                    ClientName = ClientName.Replace("\"", " ");
                    ClientName = ClientName.Replace("*", " ");
                    ClientName = ClientName.Replace("|", " ");
                    ClientName = ClientName.Replace(@"\", " ");
                    ClientName = ClientName.Replace(":", " ");
                    ClientName = ClientName.Replace("<", " ");
                    ClientName = ClientName.Replace(">", " ");
                    ClientName = ClientName.Replace("?", " ");
                    ClientName = ClientName.Replace("/", " ");
                    ReportName = ClientName + " №" + string.Join(",", SelectedOrderNumbers);

                    var PhantomForm = new Infinium.PhantomForm();
                    PhantomForm.Show();
                    var InMutualSettlement = false;
                    var Result = 1;
                    var Notes = string.Empty;
                    var SaveFilePath = string.Empty;
                    SaveDBFReportMenu = new Infinium.SaveDBFReportMenu(this, ReportName);

                    TopForm = SaveDBFReportMenu;
                    SaveDBFReportMenu.ShowDialog();

                    PhantomForm.Close();
                    InMutualSettlement = SaveDBFReportMenu.InMutualSettlement;
                    Result = SaveDBFReportMenu.Result;
                    Notes = SaveDBFReportMenu.Notes;
                    SaveFilePath = SaveDBFReportMenu.SaveFilePath;
                    PhantomForm.Dispose();
                    SaveDBFReportMenu.Dispose();
                    TopForm = null;

                    if (Result != 3)
                    {
                        var T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                        T.Start();

                        while (!SplashWindow.bSmallCreated) ;

                        var hssfworkbook = new HSSFWorkbook();

                        DBFReport.CreateReport(ref hssfworkbook, SelectedMegaOrders, SelectedOrderNumbers, MainOrders,
                            ClientID, ClientName, ComplaintProfilCost, ComplaintTPSCost, TransportCost, AdditionalCost,
                            OrdersManager.GetMainOrdersCost(MainOrders), CurrencyTypeID, TotalWeight, false, ref TotalProfil, ref TotalTPS);
                        DBFReport.SaveDBF(SaveFilePath, ReportName, false, ref ProfilDBFName, ref TPSDBFName);

                        PrepareReport.Report(ref hssfworkbook, SelectedMegaOrders, SelectedOrderNumbers, MainOrders, ClientID, ClientName,
                            ComplaintProfilCost, ComplaintTPSCost, TransportCost, AdditionalCost,
                            OrdersManager.GetMainOrdersCost(MainOrders), CurrencyTypeID, TotalWeight, false);

                        var file = new FileInfo(SaveFilePath + @"\" + ReportName + ".xls");
                        var j = 1;
                        while (file.Exists == true)
                        {
                            file = new FileInfo(SaveFilePath + @"\" + ReportName + "(" + j++ + ").xls");
                        }

                        var NewFile = new FileStream(file.FullName, FileMode.Create);
                        hssfworkbook.Write(NewFile);
                        NewFile.Close();

                        while (SplashWindow.bSmallCreated)
                            SmallWaitForm.CloseS = true;
                        NeedSplash = true;
                    }
                }
                if (HasSampleOrders)
                {
                    decimal TotalProfil = 0;
                    decimal TotalTPS = 0;

                    var ProfilDBFName = string.Empty;
                    var TPSDBFName = string.Empty;
                    var ReportName = string.Empty;
                    var ClientCountryID = OrdersManager.GetClientCountry(ClientID);

                    ClientName = ClientName.Replace("\"", " ");
                    ClientName = ClientName.Replace("*", " ");
                    ClientName = ClientName.Replace("|", " ");
                    ClientName = ClientName.Replace(@"\", " ");
                    ClientName = ClientName.Replace(":", " ");
                    ClientName = ClientName.Replace("<", " ");
                    ClientName = ClientName.Replace(">", " ");
                    ClientName = ClientName.Replace("?", " ");
                    ClientName = ClientName.Replace("/", " ");
                    ReportName = ClientName + " №" + string.Join(",", SelectedOrderNumbers) + ", обр";

                    var PhantomForm = new Infinium.PhantomForm();
                    PhantomForm.Show();
                    var InMutualSettlement = false;
                    var Result = 1;
                    var Notes = string.Empty;
                    var SaveFilePath = string.Empty;
                    SaveDBFReportMenu = new Infinium.SaveDBFReportMenu(this, ReportName);

                    TopForm = SaveDBFReportMenu;
                    SaveDBFReportMenu.ShowDialog();

                    PhantomForm.Close();
                    InMutualSettlement = SaveDBFReportMenu.InMutualSettlement;
                    Result = SaveDBFReportMenu.Result;
                    Notes = SaveDBFReportMenu.Notes;
                    SaveFilePath = SaveDBFReportMenu.SaveFilePath;
                    PhantomForm.Dispose();
                    SaveDBFReportMenu.Dispose();
                    TopForm = null;

                    if (Result != 3)
                    {
                        var T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                        T.Start();

                        while (!SplashWindow.bSmallCreated) ;

                        var hssfworkbook = new HSSFWorkbook();

                        DBFReport.CreateReport(ref hssfworkbook, SelectedMegaOrders, SelectedOrderNumbers, MainOrders,
                            ClientID, ClientName, ComplaintProfilCost, ComplaintTPSCost, TransportCost, AdditionalCost,
                            OrdersManager.GetMainOrdersCost(MainOrders), CurrencyTypeID, TotalWeight, true, ref TotalProfil, ref TotalTPS);
                        DBFReport.SaveDBF(SaveFilePath, ReportName, true, ref ProfilDBFName, ref TPSDBFName);

                        PrepareReport.Report(ref hssfworkbook, SelectedMegaOrders, SelectedOrderNumbers, MainOrders, ClientID, ClientName,
                            ComplaintProfilCost, ComplaintTPSCost, TransportCost, AdditionalCost,
                            OrdersManager.GetMainOrdersCost(MainOrders), CurrencyTypeID, TotalWeight, true);

                        var file = new FileInfo(SaveFilePath + @"\" + ReportName + ".xls");
                        var j = 1;
                        while (file.Exists == true)
                        {
                            file = new FileInfo(SaveFilePath + @"\" + ReportName + "(" + j++ + ").xls");
                        }

                        var NewFile = new FileStream(file.FullName, FileMode.Create);
                        hssfworkbook.Write(NewFile);
                        NewFile.Close();

                        while (SplashWindow.bSmallCreated)
                            SmallWaitForm.CloseS = true;
                        NeedSplash = true;
                    }
                }
            }
            else
            {
                decimal TotalProfil = 0;
                decimal TotalTPS = 0;

                var ProfilDBFName = string.Empty;
                var TPSDBFName = string.Empty;
                var ReportName = string.Empty;
                var ClientCountryID = OrdersManager.GetClientCountry(ClientID);

                ClientName = ClientName.Replace("\"", " ");
                ClientName = ClientName.Replace("*", " ");
                ClientName = ClientName.Replace("|", " ");
                ClientName = ClientName.Replace(@"\", " ");
                ClientName = ClientName.Replace(":", " ");
                ClientName = ClientName.Replace("<", " ");
                ClientName = ClientName.Replace(">", " ");
                ClientName = ClientName.Replace("?", " ");
                ClientName = ClientName.Replace("/", " ");
                ReportName = ClientName + " №" + string.Join(",", SelectedOrderNumbers);

                var PhantomForm = new Infinium.PhantomForm();
                PhantomForm.Show();
                var InMutualSettlement = false;
                var Result = 1;
                var Notes = string.Empty;
                var SaveFilePath = string.Empty;
                SaveDBFReportMenu = new Infinium.SaveDBFReportMenu(this, ReportName);

                TopForm = SaveDBFReportMenu;
                SaveDBFReportMenu.ShowDialog();

                PhantomForm.Close();
                InMutualSettlement = SaveDBFReportMenu.InMutualSettlement;
                Result = SaveDBFReportMenu.Result;
                Notes = SaveDBFReportMenu.Notes;
                SaveFilePath = SaveDBFReportMenu.SaveFilePath;
                PhantomForm.Dispose();
                SaveDBFReportMenu.Dispose();
                TopForm = null;

                if (Result != 3)
                {
                    var T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                    T.Start();

                    while (!SplashWindow.bSmallCreated) ;

                    var hssfworkbook = new HSSFWorkbook();

                    DBFReport.CreateReport(ref hssfworkbook, SelectedMegaOrders, SelectedOrderNumbers, MainOrders,
                        ClientID, ClientName, ComplaintProfilCost, ComplaintTPSCost, TransportCost, AdditionalCost,
                        OrdersManager.GetMainOrdersCost(MainOrders), CurrencyTypeID, TotalWeight, ref TotalProfil, ref TotalTPS);
                    DBFReport.SaveDBF(SaveFilePath, ReportName, ref ProfilDBFName, ref TPSDBFName);

                    PrepareReport.Report(ref hssfworkbook, SelectedMegaOrders, SelectedOrderNumbers, MainOrders, ClientID, ClientName,
                        ComplaintProfilCost, ComplaintTPSCost, TransportCost, AdditionalCost,
                        OrdersManager.GetMainOrdersCost(MainOrders), CurrencyTypeID, TotalWeight);

                    var file = new FileInfo(SaveFilePath + @"\" + ReportName + ".xls");
                    var j = 1;
                    while (file.Exists == true)
                    {
                        file = new FileInfo(SaveFilePath + @"\" + ReportName + "(" + j++ + ").xls");
                    }

                    var NewFile = new FileStream(file.FullName, FileMode.Create);
                    hssfworkbook.Write(NewFile);
                    NewFile.Close();
                    System.Diagnostics.Process.Start(file.FullName);

                    while (SplashWindow.bSmallCreated)
                        SmallWaitForm.CloseS = true;
                    NeedSplash = true;
                }
            }
        }

        private void kryptonContextMenuItem26_Click(object sender, EventArgs e)
        {
            var PhantomForm = new Infinium.PhantomForm();
            PhantomForm.Show();

            var ClientID = 0;
            var MegaOrderID = 0;
            var MainOrderID = 0;

            if (OrdersManager.MegaOrdersBindingSource.Count > 0)
            {
                ClientID = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["ClientID"]);
                MegaOrderID = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["MegaOrderID"]);
            }
            if (OrdersManager.MainOrdersBindingSource.Count > 0)
            {
                MainOrderID = Convert.ToInt32(((DataRowView)OrdersManager.MainOrdersBindingSource.Current).Row["MainOrderID"]);
            }

            var CopyMarketingOrderForm = new CopyMarketingOrderForm(this, false, ClientID, MegaOrderID, MainOrderID);

            TopForm = CopyMarketingOrderForm;
            CopyMarketingOrderForm.ShowDialog();
            ClientID = CopyMarketingOrderForm.ClientID;

            PhantomForm.Close();

            PhantomForm.Dispose();
            CopyMarketingOrderForm.Dispose();
            TopForm = null;

            if (ClientID > 0)
            {
                NeedSplash = false;
                Filter();
                NeedSplash = true;
            }
        }

        private void kryptonContextMenuItem27_Click(object sender, EventArgs e)
        {
            var PhantomForm = new Infinium.PhantomForm();
            PhantomForm.Show();

            var ClientID = 0;
            var MegaOrderID = 0;
            var MainOrderID = 0;

            if (OrdersManager.MegaOrdersBindingSource.Count > 0)
            {
                ClientID = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["ClientID"]);
                MegaOrderID = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["MegaOrderID"]);
            }
            if (OrdersManager.MainOrdersBindingSource.Count > 0)
            {
                MainOrderID = Convert.ToInt32(((DataRowView)OrdersManager.MainOrdersBindingSource.Current).Row["MainOrderID"]);
            }

            var CopyMarketingOrderForm = new CopyMarketingOrderForm(this, true, ClientID, MegaOrderID, MainOrderID);

            TopForm = CopyMarketingOrderForm;
            CopyMarketingOrderForm.ShowDialog();
            ClientID = CopyMarketingOrderForm.ClientID;

            PhantomForm.Close();

            PhantomForm.Dispose();
            CopyMarketingOrderForm.Dispose();
            TopForm = null;

            if (ClientID > 0)
            {
                NeedSplash = false;
                Filter();
                NeedSplash = true;
            }
        }

        private void kryptonContextMenuItem28_Click(object sender, EventArgs e)
        {
            for (var i = 0; i < MegaOrdersDataGrid.SelectedRows.Count; i++)
            {
                CheckOrdersStatus.GG(
                    Convert.ToInt32(Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[i].Cells["MegaOrderID"].Value)));
            }
        }

        private void MainOrdersFrontsOrdersDataGrid_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            //var dataGridViewRow = ((PercentageDataGrid)sender).CurrentRow;
            //if (dataGridViewRow != null)
            //    dataGridViewRow.Cells[e.ColumnIndex].Value = "Наименование удалено";
        }

        private void OnAgreementCheckBox_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void btnExportToCabModule_Click(object sender, EventArgs e)
        {
            if (OrdersManager.MegaOrdersBindingSource.Count == 0 
                || OrdersManager.MegaOrdersBindingSource.Current == null)
                return;

            var megaOrderId = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["MegaOrderID"]);
            
            var OKCancel = Infinium.LightMessageBox.Show(ref TopForm, true,
                    $"Вы собираетесь создать задания в модуле Корпусная мебель. Продолжить?",
                    "Задания корпусной мебели");
            if (!OKCancel)
                return;

            NeedSplash = false;

            var T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
            T.Start();

            while (!SplashWindow.bSmallCreated)
            {
            }

            OrdersManager.CreateCabFurAssignments(megaOrderId);
            InfiniumTips.ShowTip(this, 50, 85, "Задания созданы", 2000);

            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
            NeedSplash = true;
        }

        private void btnSetAgreedDate_Click(object sender, EventArgs e)
        {
            DateTime agreedDate;
            if (((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["ConfirmDateTime"] != null)
                agreedDate =
                    Convert.ToDateTime(
                        ((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["ConfirmDateTime"]);
            else
                return;

            var phantomForm = new Infinium.PhantomForm();
            phantomForm.Show();

            var megaOrderId = 0;
            if (OrdersManager.MegaOrdersBindingSource.Count > 0)
                megaOrderId = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["MegaOrderID"]);

            var agreedDateForm = new SetOrderAgreedDateForm(this, agreedDate);

            TopForm = agreedDateForm;
            agreedDateForm.ShowDialog();

            if (agreedDateForm.DialogResult == DialogResult.OK)
            {
                agreedDate = agreedDateForm.AgreedDate;

                OrdersManager.SetAgredDate(agreedDate);
                NeedSplash = false;
                Filter();
                OrdersManager.MoveToMegaOrder(megaOrderId);
                NeedSplash = true;
            }

            phantomForm.Close();

            phantomForm.Dispose();
            agreedDateForm.Dispose();

            TopForm = null;
        }

        private void ManagerCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            FilterManagersDataGrid.Enabled = ManagerCheckBox.Checked;

            if (OrdersManager == null)
                return; 
            Filter();
        }

        private void FilterManagersDataGrid_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            var senderGrid = (DataGridView)sender;
            
            if (senderGrid.Columns[e.ColumnIndex].Name == "checkCol" && e.RowIndex >= 0)
            {
                Filter();
            }
        }
    }
}
