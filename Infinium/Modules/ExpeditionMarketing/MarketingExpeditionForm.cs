﻿using Infinium.Modules.ExpeditionMarketing.ColorDispatchReportToDbf;
using Infinium.Modules.ExpeditionMarketing.NotesDispatchReportToDbf;
using Infinium.Modules.ExpeditionMarketing.DispatchReportToDbf;
using Infinium.Modules.Marketing.Expedition;
using Infinium.Modules.Marketing.Orders;

using NPOI.HPSF;
using NPOI.HSSF.UserModel;

using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;

namespace Infinium
{
    public partial class MarketingExpeditionForm : Form
    {
        private const int iLogisticRole = 41;
        private const int iAdminRole = 48;
        private const int iApprovedRole = 49;
        private const int iDispatchRole = 50;
        private const int iProduction = 80;

        private const int eHide = 2;
        private const int eShow = 1;
        private const int eClose = 3;
        private const int eMainMenu = 4;

        private int FormEvent = 0;
        private int CurrentFactoryID = 0;
        private int CurrentMainOrder = 0;
        private int CurrentPackNumber = 0;
        private readonly int MDispatchID = 0;

        private bool CanEditDispatch = false;
        //bool bCreateDispatch = false;
        //bool bSetDispatchDate = false;
        //bool bConfirmDispatch = false;
        //bool bPrintDispReport = false;

        private readonly bool MoveFromPermits = false;
        private bool NeedRefresh = false;
        private bool NeedSplash = false;

        private readonly Form MainForm;
        private SaveDBFReportMenu SaveDBFReportMenu;
        private InputNewDecorCountForm InputNewDecorCountForm;
        private readonly RoleTypes RoleType = RoleTypes.OrdinaryRole;

        public enum RoleTypes
        {
            OrdinaryRole = 0,
            AdminRole = 1,
            LogisticsRole = 2,
            ApprovedRole = 3,
            DispatchRole = 4,
            Production = 5
        }

        private readonly DateTime PrepareDispatchDateTime;

        private DataTable MonthsDT;
        private DataTable YearsDT;
        private readonly DataTable RolePermissionsDataTable;

        private Form TopForm = null;
        private readonly LightStartForm LightStartForm;

        private FrontsCatalogOrder FrontsCatalogOrder;
        private DecorCatalogOrder DecorCatalogOrder;
        private DispatchReport DispatchReport;
        private CabFurAssembleReport cabFurAssembleReport;
        private Modules.CabFurnitureAssignments.CabFurAssemble cabFurAssembleManager;
        private MarketingDispatch MarketingDispatchManager;
        private DispatchReportToDbf _dispatchReportToDbf;
        private ColorInvoiceReportToDbf _colorInvoiceReportToDbf;
        private NotesInvoiceReportToDbf _notesInvoiceReportToDbf;
        private Infinium.Modules.Marketing.Dispatch.DetailsReport DetailsReport;
        private Infinium.Modules.Marketing.Orders.InvoiceReportToDbf.InvoiceReportToDbf iDBFReport;
        private MarketingExpeditionManager MarketingExpeditionManager;
        private OrdersCalculate OrdersCalculate;

        public MarketingExpeditionForm(LightStartForm tLightStartForm)
        {
            InitializeComponent();
            LightStartForm = tLightStartForm;


            this.MaximumSize = Screen.PrimaryScreen.WorkingArea.Size;

            //RolePermissionsDataTable = RolesAndPermissionsManager.GetPermissions(Security.CurrentUserID, this.Name);
            Initialize();
            MarketingExpeditionManager.GetPermissions(Security.CurrentUserID, this.Name);
            if (MarketingExpeditionManager.PermissionGranted(iAdminRole))
            {
                kryptonButton5.Visible = true;
                RoleType = RoleTypes.AdminRole;
            }
            if (MarketingExpeditionManager.PermissionGranted(iLogisticRole))
            {
                RoleType = RoleTypes.LogisticsRole;
            }
            if (MarketingExpeditionManager.PermissionGranted(iApprovedRole))
            {
                RoleType = RoleTypes.ApprovedRole;
            }
            if (MarketingExpeditionManager.PermissionGranted(iDispatchRole))
            {
                RoleType = RoleTypes.DispatchRole;
            }
            if (MarketingExpeditionManager.PermissionGranted(iProduction))
            {
                RoleType = RoleTypes.Production;
            }

            //if (PermissionGranted(iCreateDispatch))
            //{
            //    bCreateDispatch = true;
            //}
            //if (PermissionGranted(iSetDispatchDate))
            //{
            //    bSetDispatchDate = true;
            //}
            //if (PermissionGranted(iApprovedRole))
            //{
            //    RoleType = RoleTypes.ApprovedRole;
            //}
            //if (PermissionGranted(iPrintDispReport))
            //{
            //    bPrintDispReport = true;
            //}

            //if (bCreateDispatch && bSetDispatchDate && bPrintDispReport && bConfirmDispatch)
            //    RoleType = RoleTypes.AdminRole;
            //if (bCreateDispatch && bSetDispatchDate && bPrintDispReport && !bConfirmDispatch)
            //    RoleType = RoleTypes.LogisticsRole;
            //if (!bCreateDispatch && !bSetDispatchDate && bPrintDispReport && !bConfirmDispatch)
            //    RoleType = RoleTypes.DispatchRole;
            //if (!bCreateDispatch && !bSetDispatchDate && !bPrintDispReport && bConfirmDispatch)
            //    RoleType = RoleTypes.ApprovedRole;

            if (RoleType == RoleTypes.OrdinaryRole)
            {
                btnAddDispatch.Visible = false;
                btnEditDispatch.Visible = false;
                btnChangeDispatchDate.Visible = false;
                cmiChangeDispatchDate.Visible = false;
                btnConfirmExpedition.Visible = false;
                ConfirmExpContextMenuItem.Visible = false;
                btnConfirmDispatch.Visible = false;
                ConfirmDispatchContextMenuItem.Visible = false;
                PrintDispatchListMenuItem.Visible = false;
                AttachContextMenuItem.Visible = false;
                cmiBindToPermit.Visible = false;
                cmiDispatchPackages.Visible = false;
                kryptonContextMenuItem5.Visible = false;
                //kryptonContextMenuItem6.Visible = false;
                kryptonContextMenuItem14.Visible = false;
            }
            if (RoleType == RoleTypes.AdminRole)
            {
                btnAddDispatch.Visible = true;
                btnEditDispatch.Visible = true;
                btnChangeDispatchDate.Visible = true;
                cmiChangeDispatchDate.Visible = true;
                btnConfirmExpedition.Visible = true;
                ConfirmExpContextMenuItem.Visible = true;
                btnConfirmDispatch.Visible = true;
                ConfirmDispatchContextMenuItem.Visible = true;
                PrintDispatchListMenuItem.Visible = true;
                AttachContextMenuItem.Visible = true;
                cmiBindToPermit.Visible = true;
                cmiDispatchPackages.Visible = true;
                kryptonContextMenuItem5.Visible = true;
                //kryptonContextMenuItem6.Visible = true;
                kryptonContextMenuItem14.Visible = true;
                this.MenuPanel.Size = new System.Drawing.Size(1080, 507);
            }
            if (RoleType == RoleTypes.ApprovedRole)
            {
                btnConfirmExpedition.Visible = false;
                btnChangeDispatchDate.Visible = false;
                cmiChangeDispatchDate.Visible = false;
                ConfirmExpContextMenuItem.Visible = false;
                btnConfirmDispatch.Visible = true;
                ConfirmDispatchContextMenuItem.Visible = true;
                PrintDispatchListMenuItem.Visible = false;
                AttachContextMenuItem.Visible = false;
                cmiBindToPermit.Visible = false;
                cmiDispatchPackages.Visible = false;
                kryptonContextMenuItem5.Visible = true;
                //kryptonContextMenuItem6.Visible = true;
                kryptonContextMenuItem14.Visible = true;
            }
            if (RoleType == RoleTypes.LogisticsRole)
            {
                btnAddDispatch.Visible = true;
                btnEditDispatch.Visible = true;
                btnChangeDispatchDate.Visible = true;
                cmiChangeDispatchDate.Visible = true;
                btnConfirmExpedition.Visible = true;
                ConfirmExpContextMenuItem.Visible = true;
                btnConfirmDispatch.Visible = false;
                ConfirmDispatchContextMenuItem.Visible = false;
                cmiBindToPermit.Visible = true;
                cmiDispatchPackages.Visible = true;
                kryptonContextMenuItem5.Visible = true;
                //kryptonContextMenuItem6.Visible = true;
                kryptonContextMenuItem14.Visible = true;
            }
            if (RoleType == RoleTypes.DispatchRole)
            {
                btnAddDispatch.Visible = false;
                btnEditDispatch.Visible = false;
                btnChangeDispatchDate.Visible = false;
                cmiChangeDispatchDate.Visible = false;
                btnConfirmExpedition.Visible = false;
                ConfirmExpContextMenuItem.Visible = false;
                btnConfirmDispatch.Visible = false;
                ConfirmDispatchContextMenuItem.Visible = false;
                PrintDispatchListMenuItem.Visible = true;
                AttachContextMenuItem.Visible = true;
                cmiBindToPermit.Visible = false;
                cmiDispatchPackages.Visible = true;
                kryptonContextMenuItem5.Visible = false;
                //kryptonContextMenuItem6.Visible = false;
                kryptonContextMenuItem14.Visible = false;
            }

            while (!SplashForm.bCreated) ;
        }

        public MarketingExpeditionForm(Form tMainForm, DateTime dPrepareDispatchDateTime, int iDispatchID)
        {
            InitializeComponent();
            MainForm = tMainForm;
            PrepareDispatchDateTime = dPrepareDispatchDateTime;
            MDispatchID = iDispatchID;
            this.MaximumSize = Screen.PrimaryScreen.WorkingArea.Size;

            RolePermissionsDataTable = RolesAndPermissionsManager.GetPermissions(Security.CurrentUserID);
            MoveFromPermits = true;
            Initialize();
            pnlExpedition.BringToFront();

            while (!SplashForm.bCreated) ;
        }

        private void MarketingExpeditionForm_Shown(object sender, EventArgs e)
        {
            while (!SplashForm.bCreated) ;
            NeedSplash = true;
            FormEvent = eShow;
            AnimateTimer.Enabled = true;
            MegaOrdersDataGrid.SelectionChanged += MegaOrdersDataGrid_SelectionChanged;
            if ((DataRowView)MarketingExpeditionManager.MegaOrdersBindingSource.Current != null)
            {
                MarketingExpeditionManager.FilterMainOrdersByMegaOrder(false,
                    false, false, true, true, true, false,
                    Convert.ToInt32(((DataRowView)MarketingExpeditionManager.MegaOrdersBindingSource.Current)["MegaOrderID"]), -1, 0);
                MarketingExpeditionManager.PackedMainOrdersFrontsOrders.SetColor(CurrentMainOrder, CurrentPackNumber);
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
                        if (!MoveFromPermits)
                            LightStartForm.CloseForm(this);
                        this.Close();
                    }

                    if (FormEvent == eHide)
                    {
                        if (!MoveFromPermits)
                            LightStartForm.HideForm(this);
                        else
                        {
                            MainForm.Activate();
                            this.Hide();
                        }
                    }


                    return;
                }

                if (FormEvent == eShow)
                {
                    NeedSplash = true;
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
                        if (!MoveFromPermits)
                            LightStartForm.CloseForm(this);
                        this.Close();
                    }

                    if (FormEvent == eHide)
                    {
                        if (!MoveFromPermits)
                            LightStartForm.HideForm(this);
                        else
                        {
                            MainForm.Activate();
                            this.Hide();
                        }
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

        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);

            if (m.Msg == 6 && m.WParam.ToInt32() == 1)
            {
                if (TopForm != null)
                    TopForm.Activate();
            }
        }

        private void NavigateMenuCloseButton_Click(object sender, EventArgs e)
        {
            FormEvent = eClose;
            AnimateTimer.Enabled = true;
        }

        private void NavigateMenuBackButton_Click(object sender, EventArgs e)
        {
            NeedSplash = false;
            FormEvent = eHide;
            AnimateTimer.Enabled = true;
        }

        private void NavigateMenuHerculesButton_Click(object sender, EventArgs e)
        {
            FormEvent = eMainMenu;
            AnimateTimer.Enabled = true;
        }

        private bool PermissionGranted(int PermissionID)
        {
            DataRow[] Rows = RolePermissionsDataTable.Select("PermissionID = " + PermissionID);

            if (Rows.Count() > 0)
            {
                return Convert.ToBoolean(Rows[0]["Granted"]);
            }

            return false;
        }

        private DataGridViewComboBoxColumn MegaFactoryTypeColumn = null;
        private DataGridViewComboBoxColumn ProfilOrderStatusColumn = null;
        private DataGridViewComboBoxColumn TPSOrderStatusColumn = null;
        private DataGridViewComboBoxColumn CurrencyTypeColumn = null;

        private void CreateColumns()
        {
            ProfilOrderStatusColumn = new DataGridViewComboBoxColumn()
            {
                Name = "ProfilOrderStatusColumn",
                HeaderText = "Статус заказа\n\rПрофиль",
                DataPropertyName = "ProfilOrderStatusID",
                DataSource = new DataView(MarketingExpeditionManager.FirmOrderStatusesDataTable),
                ValueMember = "FirmOrderStatusID",
                DisplayMember = "FirmOrderStatus",
                DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing,
                SortMode = DataGridViewColumnSortMode.Automatic
            };
            TPSOrderStatusColumn = new DataGridViewComboBoxColumn()
            {
                Name = "TPSOrderStatusColumn",
                HeaderText = "Статус\n\rзаказа ТПС",
                DataPropertyName = "TPSOrderStatusID",
                DataSource = new DataView(MarketingExpeditionManager.FirmOrderStatusesDataTable),
                ValueMember = "FirmOrderStatusID",
                DisplayMember = "FirmOrderStatus",
                DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing,
                SortMode = DataGridViewColumnSortMode.Automatic
            };
            MegaFactoryTypeColumn = new DataGridViewComboBoxColumn()
            {
                Name = "MegaFactoryTypeColumn",
                HeaderText = "  Тип\n\rпр-ва",
                DataPropertyName = "FactoryID",
                DataSource = new DataView(MarketingExpeditionManager.FactoryTypesDataTable),
                ValueMember = "FactoryID",
                DisplayMember = "FactoryName",
                DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing,
                SortMode = DataGridViewColumnSortMode.Automatic
            };
            CurrencyTypeColumn = new DataGridViewComboBoxColumn()
            {
                Name = "CurrencyTypeColumn",
                HeaderText = "Валюта",
                DataPropertyName = "CurrencyTypeID",
                DataSource = MarketingExpeditionManager.CurrencyTypesBindingSource,
                ValueMember = "CurrencyTypeID",
                DisplayMember = "CurrencyType",
                DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing,
                SortMode = DataGridViewColumnSortMode.Automatic
            };
        }

        private void MegaGridSetting()
        {
            foreach (DataGridViewColumn Column in MegaOrdersDataGrid.Columns)
            {
                Column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }

            if (MegaOrdersDataGrid.Columns.Contains("DelayOfPayment"))
                MegaOrdersDataGrid.Columns["DelayOfPayment"].Visible = false;
            if (MegaOrdersDataGrid.Columns.Contains("CreatedByClient"))
                MegaOrdersDataGrid.Columns["CreatedByClient"].Visible = false;
            if (MegaOrdersDataGrid.Columns.Contains("FixedPaymentRate"))
                MegaOrdersDataGrid.Columns["FixedPaymentRate"].Visible = false;
            if (MegaOrdersDataGrid.Columns.Contains("TransportType"))
                MegaOrdersDataGrid.Columns["TransportType"].Visible = false;
            MegaOrdersDataGrid.Columns["ClientID"].Visible = false;
            MegaOrdersDataGrid.Columns["OrderStatusID"].Visible = false;
            MegaOrdersDataGrid.Columns["ProfilOrderStatusID"].Visible = false;
            MegaOrdersDataGrid.Columns["TPSOrderStatusID"].Visible = false;
            MegaOrdersDataGrid.Columns["AgreementStatusID"].Visible = false;
            MegaOrdersDataGrid.Columns["CurrencyTypeID"].Visible = false;
            MegaOrdersDataGrid.Columns["FactoryID"].Visible = false;
            MegaOrdersDataGrid.Columns["DesireDate"].Visible = false;
            MegaOrdersDataGrid.Columns["LastCalcDate"].Visible = false;
            if (MegaOrdersDataGrid.Columns.Contains("LastCalcUserID"))
                MegaOrdersDataGrid.Columns["LastCalcUserID"].Visible = false;
            MegaOrdersDataGrid.Columns["OrderDate"].Visible = false;
            MegaOrdersDataGrid.Columns["ConfirmDateTime"].Visible = false;
            MegaOrdersDataGrid.Columns["TPSPackAllocStatusID"].Visible = false;
            MegaOrdersDataGrid.Columns["TPSPackCount"].Visible = false;
            MegaOrdersDataGrid.Columns["ProfilPackAllocStatusID"].Visible = false;
            MegaOrdersDataGrid.Columns["ProfilPackCount"].Visible = false;
            if (MegaOrdersDataGrid.Columns.Contains("IsComplaint"))
                MegaOrdersDataGrid.Columns["IsComplaint"].Visible = false;

            if (!Security.PriceAccess)
            {
                if (MegaOrdersDataGrid.Columns.Contains("ComplaintProfilCost"))
                    MegaOrdersDataGrid.Columns["ComplaintProfilCost"].Visible = false;
                if (MegaOrdersDataGrid.Columns.Contains("ComplaintTPSCost"))
                    MegaOrdersDataGrid.Columns["ComplaintTPSCost"].Visible = false;
                if (MegaOrdersDataGrid.Columns.Contains("ComplaintNotes"))
                    MegaOrdersDataGrid.Columns["ComplaintNotes"].Visible = false;
                if (MegaOrdersDataGrid.Columns.Contains("CurrencyComplaintProfilCost"))
                    MegaOrdersDataGrid.Columns["CurrencyComplaintProfilCost"].Visible = false;
                if (MegaOrdersDataGrid.Columns.Contains("CurrencyComplaintTPSCost"))
                    MegaOrdersDataGrid.Columns["CurrencyComplaintTPSCost"].Visible = false;
                if (MegaOrdersDataGrid.Columns.Contains("DelayOfPayment"))
                    MegaOrdersDataGrid.Columns["DelayOfPayment"].Visible = false;
                MegaOrdersDataGrid.Columns["OrderCost"].Visible = false;
                MegaOrdersDataGrid.Columns["TransportCost"].Visible = false;
                MegaOrdersDataGrid.Columns["AdditionalCost"].Visible = false;
                MegaOrdersDataGrid.Columns["CurrencyOrderCost"].Visible = false;
                MegaOrdersDataGrid.Columns["CurrencyTotalCost"].Visible = false;
                MegaOrdersDataGrid.Columns["CurrencyTransportCost"].Visible = false;
                MegaOrdersDataGrid.Columns["CurrencyAdditionalCost"].Visible = false;
                MegaOrdersDataGrid.Columns["TotalCost"].Visible = false;
                MegaOrdersDataGrid.Columns["Rate"].Visible = false;
                MegaOrdersDataGrid.Columns["CurrencyTypeColumn"].Visible = false;
            }

            if (MegaOrdersDataGrid.Columns.Contains("ProfilConfirmProduction"))
                MegaOrdersDataGrid.Columns["ProfilConfirmProduction"].Visible = false;
            if (MegaOrdersDataGrid.Columns.Contains("TPSConfirmProduction"))
                MegaOrdersDataGrid.Columns["TPSConfirmProduction"].Visible = false;
            if (MegaOrdersDataGrid.Columns.Contains("ProfilAllowDispatch"))
                MegaOrdersDataGrid.Columns["ProfilAllowDispatch"].Visible = false;
            if (MegaOrdersDataGrid.Columns.Contains("TPSAllowDispatch"))
                MegaOrdersDataGrid.Columns["TPSAllowDispatch"].Visible = false;
            if (MegaOrdersDataGrid.Columns.Contains("ProfilConfirmDispatch"))
                MegaOrdersDataGrid.Columns["ProfilConfirmDispatch"].Visible = false;
            if (MegaOrdersDataGrid.Columns.Contains("TPSConfirmDispatch"))
                MegaOrdersDataGrid.Columns["TPSConfirmDispatch"].Visible = false;

            if (MegaOrdersDataGrid.Columns.Contains("ProfilProductionDate"))
                MegaOrdersDataGrid.Columns["ProfilProductionDate"].Visible = false;
            if (MegaOrdersDataGrid.Columns.Contains("TPSProductionDate"))
                MegaOrdersDataGrid.Columns["TPSProductionDate"].Visible = false;

            MegaOrdersDataGrid.Columns["OrderCost"].HeaderText = "Стоимость\r\nзаказа, евро";
            MegaOrdersDataGrid.Columns["TransportCost"].HeaderText = "Стоимость\r\nтранспорта, евро";
            MegaOrdersDataGrid.Columns["AdditionalCost"].HeaderText = "Дополнительная\r\nстоимость, евро";
            MegaOrdersDataGrid.Columns["ConfirmDateTime"].HeaderText = "Дата\r\nсогласования";
            MegaOrdersDataGrid.Columns["CurrencyOrderCost"].HeaderText = "Стоимость\r\nзаказа, расчет";
            MegaOrdersDataGrid.Columns["CurrencyTotalCost"].HeaderText = "Итого, расчет";
            MegaOrdersDataGrid.Columns["CurrencyTransportCost"].HeaderText = "Стоимость \r\nтранспорта, расчет";
            MegaOrdersDataGrid.Columns["CurrencyAdditionalCost"].HeaderText = "Дополнительная\r\nстоимость, расчет";
            MegaOrdersDataGrid.Columns["TotalCost"].HeaderText = "Итого, евро";
            MegaOrdersDataGrid.Columns["Rate"].HeaderText = "Курс";

            MegaOrdersDataGrid.Columns["MegaOrderID"].HeaderText = "№ п\\п";
            MegaOrdersDataGrid.Columns["ClientName"].HeaderText = "Клиент";
            MegaOrdersDataGrid.Columns["OrderNumber"].HeaderText = "№\r\nзаказа";
            MegaOrdersDataGrid.Columns["OrderDate"].HeaderText = "Дата\r\nсоздания";
            MegaOrdersDataGrid.Columns["ProfilDispatchDate"].HeaderText = "Дата отгрузки\r\nПрофиль";
            MegaOrdersDataGrid.Columns["TPSDispatchDate"].HeaderText = "Дата отгрузки\r\nТПС";
            MegaOrdersDataGrid.Columns["AdditionalCost"].HeaderText = "Дополнительная\r\nстоимость, евро";
            MegaOrdersDataGrid.Columns["ConfirmDateTime"].HeaderText = "Дата\r\nсогласования";
            MegaOrdersDataGrid.Columns["Weight"].HeaderText = "Вес, кг.";
            MegaOrdersDataGrid.Columns["Square"].HeaderText = "Площадь\r\nфасадов, м.кв.";
            MegaOrdersDataGrid.Columns["DesireDate"].HeaderText = "Предварит.\r\nдата отгрузки";

            MegaOrdersDataGrid.Columns["ProfilPackPercentage"].HeaderText = "Упаковано\r\nПрофиль, %";
            MegaOrdersDataGrid.Columns["ProfilPackedCount"].HeaderText = "Упаковано\r\n Профиль, кол-во";
            MegaOrdersDataGrid.Columns["ProfilStoreCount"].HeaderText = "Склад,\r\nПрофиль, кол-во";
            MegaOrdersDataGrid.Columns["ProfilStorePercentage"].HeaderText = "Склад,\r\nПрофиль, %";
            MegaOrdersDataGrid.Columns["ProfilExpCount"].HeaderText = "Экспедиция,\r\nПрофиль, кол-во";
            MegaOrdersDataGrid.Columns["ProfilExpPercentage"].HeaderText = "Экспедиция,\r\nПрофиль, %";
            MegaOrdersDataGrid.Columns["ProfilDispPercentage"].HeaderText = "Отгружено\r\nПрофиль, %";
            MegaOrdersDataGrid.Columns["ProfilDispatchedCount"].HeaderText = "Отгружено\r\n Профиль, кол-во";
            MegaOrdersDataGrid.Columns["TPSPackPercentage"].HeaderText = "Упаковано\r\nТПС, %";
            MegaOrdersDataGrid.Columns["TPSPackedCount"].HeaderText = "Упаковано\r\nТПС, кол-во";
            MegaOrdersDataGrid.Columns["TPSStoreCount"].HeaderText = "Склад,\r\nТПС, кол-во";
            MegaOrdersDataGrid.Columns["TPSStorePercentage"].HeaderText = "Склад,\r\nТПС, %";
            MegaOrdersDataGrid.Columns["TPSExpCount"].HeaderText = "Экспедиция,\r\nТПС, кол-во";
            MegaOrdersDataGrid.Columns["TPSExpPercentage"].HeaderText = "Экспедиция,\r\nТПС, %";
            MegaOrdersDataGrid.Columns["TPSDispPercentage"].HeaderText = "Отгружено\r\nТПС, %";
            MegaOrdersDataGrid.Columns["TPSDispatchedCount"].HeaderText = " Отгружено\r\nТПС, кол-во";
            MegaOrdersDataGrid.Columns["IsComplaint"].Visible = false;
            MegaOrdersDataGrid.Columns["MegaBatchNumber"].HeaderText = "Группа\n\rпартий";

            MegaOrdersDataGrid.Columns["PlanDispDate"].HeaderText = "Плановая\r\nдата отгрузки";
            MegaOrdersDataGrid.Columns["PlanDispDate"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MegaOrdersDataGrid.Columns["PlanDispDate"].Width = 110;
            NumberFormatInfo nfi1 = new NumberFormatInfo()
            {
                CurrencyGroupSeparator = " ",
                CurrencySymbol = "",
                CurrencyDecimalDigits = 1
            };
            NumberFormatInfo nfi2 = new NumberFormatInfo()
            {
                CurrencyGroupSeparator = " ",
                CurrencySymbol = "",
                CurrencyDecimalDigits = 4
            };
            NumberFormatInfo nfi3 = new NumberFormatInfo()
            {
                CurrencyGroupSeparator = " ",
                CurrencySymbol = "",
                CurrencyDecimalDigits = 2
            };
            MegaOrdersDataGrid.Columns["OrderCost"].DefaultCellStyle.Format = "C";
            MegaOrdersDataGrid.Columns["OrderCost"].DefaultCellStyle.FormatProvider = nfi1;

            MegaOrdersDataGrid.Columns["TransportCost"].DefaultCellStyle.Format = "C";
            MegaOrdersDataGrid.Columns["TransportCost"].DefaultCellStyle.FormatProvider = nfi1;

            MegaOrdersDataGrid.Columns["AdditionalCost"].DefaultCellStyle.Format = "C";
            MegaOrdersDataGrid.Columns["AdditionalCost"].DefaultCellStyle.FormatProvider = nfi1;

            MegaOrdersDataGrid.Columns["TotalCost"].DefaultCellStyle.Format = "C";
            MegaOrdersDataGrid.Columns["TotalCost"].DefaultCellStyle.FormatProvider = nfi1;

            MegaOrdersDataGrid.Columns["CurrencyOrderCost"].DefaultCellStyle.Format = "C";
            MegaOrdersDataGrid.Columns["CurrencyOrderCost"].DefaultCellStyle.FormatProvider = nfi1;

            MegaOrdersDataGrid.Columns["CurrencyTransportCost"].DefaultCellStyle.Format = "C";
            MegaOrdersDataGrid.Columns["CurrencyTransportCost"].DefaultCellStyle.FormatProvider = nfi1;

            MegaOrdersDataGrid.Columns["CurrencyAdditionalCost"].DefaultCellStyle.Format = "C";
            MegaOrdersDataGrid.Columns["CurrencyAdditionalCost"].DefaultCellStyle.FormatProvider = nfi1;

            MegaOrdersDataGrid.Columns["CurrencyTotalCost"].DefaultCellStyle.Format = "C";
            MegaOrdersDataGrid.Columns["CurrencyTotalCost"].DefaultCellStyle.FormatProvider = nfi1;

            MegaOrdersDataGrid.Columns["Rate"].DefaultCellStyle.Format = "C";
            MegaOrdersDataGrid.Columns["Rate"].DefaultCellStyle.FormatProvider = nfi2;

            MegaOrdersDataGrid.Columns["Weight"].DefaultCellStyle.Format = "C";
            MegaOrdersDataGrid.Columns["Weight"].DefaultCellStyle.FormatProvider = nfi1;

            MegaOrdersDataGrid.Columns["Square"].DefaultCellStyle.Format = "C";
            MegaOrdersDataGrid.Columns["Square"].DefaultCellStyle.FormatProvider = nfi3;

            MegaOrdersDataGrid.Columns["ClientName"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            MegaOrdersDataGrid.Columns["ClientName"].MinimumWidth = 150;
            MegaOrdersDataGrid.Columns["MegaOrderID"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MegaOrdersDataGrid.Columns["MegaOrderID"].Width = 70;
            MegaOrdersDataGrid.Columns["OrderNumber"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MegaOrdersDataGrid.Columns["OrderNumber"].Width = 70;
            MegaOrdersDataGrid.Columns["OrderDate"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MegaOrdersDataGrid.Columns["OrderDate"].Width = 150;
            MegaOrdersDataGrid.Columns["ProfilDispatchDate"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MegaOrdersDataGrid.Columns["ProfilDispatchDate"].Width = 140;
            MegaOrdersDataGrid.Columns["TPSDispatchDate"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MegaOrdersDataGrid.Columns["TPSDispatchDate"].Width = 140;
            MegaOrdersDataGrid.Columns["DesireDate"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MegaOrdersDataGrid.Columns["DesireDate"].Width = 130;
            MegaOrdersDataGrid.Columns["ConfirmDateTime"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MegaOrdersDataGrid.Columns["ConfirmDateTime"].Width = 130;
            MegaOrdersDataGrid.Columns["ProfilOrderStatusColumn"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MegaOrdersDataGrid.Columns["ProfilOrderStatusColumn"].Width = 160;
            MegaOrdersDataGrid.Columns["TPSOrderStatusColumn"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MegaOrdersDataGrid.Columns["TPSOrderStatusColumn"].Width = 160;
            MegaOrdersDataGrid.Columns["Weight"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MegaOrdersDataGrid.Columns["Weight"].Width = 110;
            MegaOrdersDataGrid.Columns["Square"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MegaOrdersDataGrid.Columns["Square"].Width = 140;

            MegaOrdersDataGrid.Columns["OrderCost"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MegaOrdersDataGrid.Columns["OrderCost"].Width = 130;
            MegaOrdersDataGrid.Columns["TotalCost"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MegaOrdersDataGrid.Columns["TotalCost"].Width = 130;
            MegaOrdersDataGrid.Columns["TransportCost"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MegaOrdersDataGrid.Columns["TransportCost"].Width = 150;
            MegaOrdersDataGrid.Columns["AdditionalCost"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MegaOrdersDataGrid.Columns["AdditionalCost"].Width = 150;
            MegaOrdersDataGrid.Columns["CurrencyOrderCost"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MegaOrdersDataGrid.Columns["CurrencyOrderCost"].Width = 130;
            MegaOrdersDataGrid.Columns["CurrencyTotalCost"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MegaOrdersDataGrid.Columns["CurrencyTotalCost"].Width = 130;
            MegaOrdersDataGrid.Columns["CurrencyTransportCost"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MegaOrdersDataGrid.Columns["CurrencyTransportCost"].Width = 190;
            MegaOrdersDataGrid.Columns["CurrencyAdditionalCost"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MegaOrdersDataGrid.Columns["CurrencyAdditionalCost"].Width = 170;
            MegaOrdersDataGrid.Columns["CurrencyTypeColumn"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MegaOrdersDataGrid.Columns["CurrencyTypeColumn"].Width = 90;
            MegaOrdersDataGrid.Columns["Rate"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MegaOrdersDataGrid.Columns["Rate"].Width = 100;

            MegaOrdersDataGrid.Columns["ProfilPackPercentage"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MegaOrdersDataGrid.Columns["ProfilPackPercentage"].Width = 155;
            MegaOrdersDataGrid.Columns["ProfilPackedCount"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MegaOrdersDataGrid.Columns["ProfilPackedCount"].Width = 155;
            MegaOrdersDataGrid.Columns["TPSPackPercentage"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MegaOrdersDataGrid.Columns["TPSPackPercentage"].Width = 125;
            MegaOrdersDataGrid.Columns["TPSPackedCount"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MegaOrdersDataGrid.Columns["TPSPackedCount"].Width = 125;

            MegaOrdersDataGrid.Columns["ProfilStorePercentage"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MegaOrdersDataGrid.Columns["ProfilStorePercentage"].Width = 155;
            MegaOrdersDataGrid.Columns["ProfilStoreCount"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MegaOrdersDataGrid.Columns["ProfilStoreCount"].Width = 155;
            MegaOrdersDataGrid.Columns["TPSStorePercentage"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MegaOrdersDataGrid.Columns["TPSStorePercentage"].Width = 125;
            MegaOrdersDataGrid.Columns["TPSStoreCount"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MegaOrdersDataGrid.Columns["TPSStoreCount"].Width = 125;

            MegaOrdersDataGrid.Columns["ProfilExpPercentage"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MegaOrdersDataGrid.Columns["ProfilExpPercentage"].Width = 155;
            MegaOrdersDataGrid.Columns["ProfilExpCount"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MegaOrdersDataGrid.Columns["ProfilExpCount"].Width = 155;
            MegaOrdersDataGrid.Columns["TPSExpPercentage"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MegaOrdersDataGrid.Columns["TPSExpPercentage"].Width = 125;
            MegaOrdersDataGrid.Columns["TPSExpCount"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MegaOrdersDataGrid.Columns["TPSExpCount"].Width = 125;

            MegaOrdersDataGrid.Columns["ProfilDispPercentage"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MegaOrdersDataGrid.Columns["ProfilDispPercentage"].Width = 155;
            MegaOrdersDataGrid.Columns["ProfilDispatchedCount"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MegaOrdersDataGrid.Columns["ProfilDispatchedCount"].Width = 155;
            MegaOrdersDataGrid.Columns["TPSDispPercentage"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MegaOrdersDataGrid.Columns["TPSDispPercentage"].Width = 125;
            MegaOrdersDataGrid.Columns["TPSDispatchedCount"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MegaOrdersDataGrid.Columns["TPSDispatchedCount"].Width = 125;

            MegaOrdersDataGrid.Columns["MegaFactoryTypeColumn"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MegaOrdersDataGrid.Columns["MegaFactoryTypeColumn"].Width = 130;

            MegaOrdersDataGrid.Columns["MegaBatchNumber"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MegaOrdersDataGrid.Columns["MegaBatchNumber"].Width = 80;

            MegaOrdersDataGrid.AutoGenerateColumns = false;

            MegaOrdersDataGrid.Columns["ClientName"].DisplayIndex = 0;
            MegaOrdersDataGrid.Columns["OrderNumber"].DisplayIndex = 1;
            MegaOrdersDataGrid.Columns["MegaOrderID"].DisplayIndex = 2;

            MegaOrdersDataGrid.Columns["ProfilPackPercentage"].DisplayIndex = 3;
            MegaOrdersDataGrid.Columns["ProfilPackedCount"].DisplayIndex = 4;
            MegaOrdersDataGrid.Columns["ProfilStorePercentage"].DisplayIndex = 5;
            MegaOrdersDataGrid.Columns["ProfilStoreCount"].DisplayIndex = 6;
            MegaOrdersDataGrid.Columns["ProfilDispPercentage"].DisplayIndex = 7;
            MegaOrdersDataGrid.Columns["ProfilDispatchedCount"].DisplayIndex = 8;

            MegaOrdersDataGrid.Columns["TPSPackPercentage"].DisplayIndex = 9;
            MegaOrdersDataGrid.Columns["TPSPackedCount"].DisplayIndex = 10;
            MegaOrdersDataGrid.Columns["TPSStorePercentage"].DisplayIndex = 11;
            MegaOrdersDataGrid.Columns["TPSStoreCount"].DisplayIndex = 12;
            MegaOrdersDataGrid.Columns["TPSDispPercentage"].DisplayIndex = 13;
            MegaOrdersDataGrid.Columns["TPSDispatchedCount"].DisplayIndex = 14;

            MegaOrdersDataGrid.Columns["ProfilDispatchDate"].DisplayIndex = 15;
            MegaOrdersDataGrid.Columns["TPSDispatchDate"].DisplayIndex = 16;

            MegaOrdersDataGrid.Columns["MegaFactoryTypeColumn"].DisplayIndex = 17;
            MegaOrdersDataGrid.Columns["ProfilOrderStatusColumn"].DisplayIndex = 18;
            MegaOrdersDataGrid.Columns["TPSOrderStatusColumn"].DisplayIndex = 19;
            MegaOrdersDataGrid.Columns["Square"].DisplayIndex = 20;
            MegaOrdersDataGrid.Columns["Weight"].DisplayIndex = 21;
            MegaOrdersDataGrid.Columns["OrderCost"].DisplayIndex = 22;
            MegaOrdersDataGrid.Columns["TransportCost"].DisplayIndex = 23;
            MegaOrdersDataGrid.Columns["AdditionalCost"].DisplayIndex = 24;
            MegaOrdersDataGrid.Columns["TotalCost"].DisplayIndex = 25;
            MegaOrdersDataGrid.Columns["CurrencyOrderCost"].DisplayIndex = 26;
            MegaOrdersDataGrid.Columns["CurrencyTransportCost"].DisplayIndex = 27;
            MegaOrdersDataGrid.Columns["CurrencyAdditionalCost"].DisplayIndex = 28;
            MegaOrdersDataGrid.Columns["CurrencyTotalCost"].DisplayIndex = 29;
            MegaOrdersDataGrid.Columns["CurrencyTypeColumn"].DisplayIndex = 30;
            MegaOrdersDataGrid.Columns["Rate"].DisplayIndex = 31;

            MegaOrdersDataGrid.Columns["MegaBatchNumber"].DisplayIndex = 3;

            MegaOrdersDataGrid.Columns["ProfilExpPercentage"].DisplayIndex = 8;
            MegaOrdersDataGrid.Columns["ProfilExpCount"].DisplayIndex = 8;
            MegaOrdersDataGrid.Columns["TPSExpCount"].DisplayIndex = 16;
            MegaOrdersDataGrid.Columns["TPSExpPercentage"].DisplayIndex = 16;

            MegaOrdersDataGrid.Columns["ClientName"].Frozen = true;
            MegaOrdersDataGrid.Columns["OrderNumber"].Frozen = true;
            MegaOrdersDataGrid.RightToLeft = RightToLeft.No;

            MegaOrdersDataGrid.Columns["Weight"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            MegaOrdersDataGrid.Columns["Square"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            MegaOrdersDataGrid.Columns["ProfilPackedCount"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            MegaOrdersDataGrid.Columns["ProfilPackPercentage"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            MegaOrdersDataGrid.AddPercentageColumn("ProfilPackPercentage");
            MegaOrdersDataGrid.Columns["ProfilStoreCount"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            MegaOrdersDataGrid.Columns["ProfilStorePercentage"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            MegaOrdersDataGrid.AddPercentageColumn("ProfilStorePercentage");
            MegaOrdersDataGrid.Columns["ProfilExpCount"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            MegaOrdersDataGrid.Columns["ProfilExpPercentage"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            MegaOrdersDataGrid.AddPercentageColumn("ProfilExpPercentage");
            MegaOrdersDataGrid.Columns["ProfilDispatchedCount"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            MegaOrdersDataGrid.Columns["ProfilDispPercentage"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            MegaOrdersDataGrid.AddPercentageColumn("ProfilDispPercentage");

            MegaOrdersDataGrid.Columns["TPSPackedCount"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            MegaOrdersDataGrid.Columns["TPSPackPercentage"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            MegaOrdersDataGrid.AddPercentageColumn("TPSPackPercentage");
            MegaOrdersDataGrid.Columns["TPSStoreCount"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            MegaOrdersDataGrid.Columns["TPSStorePercentage"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            MegaOrdersDataGrid.AddPercentageColumn("TPSStorePercentage");
            MegaOrdersDataGrid.Columns["TPSExpCount"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            MegaOrdersDataGrid.Columns["TPSExpPercentage"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            MegaOrdersDataGrid.AddPercentageColumn("TPSExpPercentage");
            MegaOrdersDataGrid.Columns["TPSDispatchedCount"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            MegaOrdersDataGrid.Columns["TPSDispPercentage"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            MegaOrdersDataGrid.AddPercentageColumn("TPSDispPercentage");
        }

        private void ShowColumns(int FactoryID)
        {
            if (FactoryID == 0)
            {
                MainOrdersDataGrid.Columns["ProfilProductionStatusColumn"].Visible = true;
                MainOrdersDataGrid.Columns["ProfilStorageStatusColumn"].Visible = true;
                MainOrdersDataGrid.Columns["ProfilDispatchStatusColumn"].Visible = true;
                MainOrdersDataGrid.Columns["ProfilPackPercentage"].Visible = true;
                MainOrdersDataGrid.Columns["ProfilPackedCount"].Visible = true;
                MainOrdersDataGrid.Columns["ProfilStorePercentage"].Visible = true;
                MainOrdersDataGrid.Columns["ProfilStoreCount"].Visible = true;
                MainOrdersDataGrid.Columns["ProfilExpPercentage"].Visible = true;
                MainOrdersDataGrid.Columns["ProfilExpCount"].Visible = true;
                MainOrdersDataGrid.Columns["ProfilDispPercentage"].Visible = true;
                MainOrdersDataGrid.Columns["ProfilDispatchedCount"].Visible = true;
                MainOrdersDataGrid.Columns["ProfilProductionDate"].Visible = true;
                MainOrdersDataGrid.Columns["ProfilOnProductionDate"].Visible = true;

                MainOrdersDataGrid.Columns["TPSProductionStatusColumn"].Visible = true;
                MainOrdersDataGrid.Columns["TPSStorageStatusColumn"].Visible = true;
                MainOrdersDataGrid.Columns["TPSDispatchStatusColumn"].Visible = true;
                MainOrdersDataGrid.Columns["TPSPackPercentage"].Visible = true;
                MainOrdersDataGrid.Columns["TPSPackedCount"].Visible = true;
                MainOrdersDataGrid.Columns["TPSStorePercentage"].Visible = true;
                MainOrdersDataGrid.Columns["TPSStoreCount"].Visible = true;
                MainOrdersDataGrid.Columns["TPSExpPercentage"].Visible = true;
                MainOrdersDataGrid.Columns["TPSExpCount"].Visible = true;
                MainOrdersDataGrid.Columns["TPSDispPercentage"].Visible = true;
                MainOrdersDataGrid.Columns["TPSDispatchedCount"].Visible = true;
                MainOrdersDataGrid.Columns["TPSProductionDate"].Visible = true;
                MainOrdersDataGrid.Columns["TPSOnProductionDate"].Visible = true;


                MegaOrdersDataGrid.Columns["ProfilPackPercentage"].Visible = true;
                MegaOrdersDataGrid.Columns["ProfilPackedCount"].Visible = true;
                MegaOrdersDataGrid.Columns["ProfilStorePercentage"].Visible = true;
                MegaOrdersDataGrid.Columns["ProfilStoreCount"].Visible = true;
                MegaOrdersDataGrid.Columns["ProfilExpPercentage"].Visible = true;
                MegaOrdersDataGrid.Columns["ProfilExpCount"].Visible = true;
                MegaOrdersDataGrid.Columns["ProfilDispPercentage"].Visible = true;
                MegaOrdersDataGrid.Columns["ProfilDispatchedCount"].Visible = true;
                MegaOrdersDataGrid.Columns["ProfilDispatchDate"].Visible = true;
                MegaOrdersDataGrid.Columns["ProfilOrderStatusColumn"].Visible = true;

                MegaOrdersDataGrid.Columns["TPSPackPercentage"].Visible = true;
                MegaOrdersDataGrid.Columns["TPSPackedCount"].Visible = true;
                MegaOrdersDataGrid.Columns["TPSStorePercentage"].Visible = true;
                MegaOrdersDataGrid.Columns["TPSStoreCount"].Visible = true;
                MegaOrdersDataGrid.Columns["TPSExpPercentage"].Visible = true;
                MegaOrdersDataGrid.Columns["TPSExpCount"].Visible = true;
                MegaOrdersDataGrid.Columns["TPSDispPercentage"].Visible = true;
                MegaOrdersDataGrid.Columns["TPSDispatchedCount"].Visible = true;
                MegaOrdersDataGrid.Columns["TPSDispatchDate"].Visible = true;
                MegaOrdersDataGrid.Columns["TPSOrderStatusColumn"].Visible = true;
            }

            if (FactoryID == 1)
            {
                MainOrdersDataGrid.Columns["ProfilProductionStatusColumn"].Visible = true;
                MainOrdersDataGrid.Columns["ProfilStorageStatusColumn"].Visible = true;
                MainOrdersDataGrid.Columns["ProfilDispatchStatusColumn"].Visible = true;
                MainOrdersDataGrid.Columns["ProfilPackPercentage"].Visible = true;
                MainOrdersDataGrid.Columns["ProfilStorePercentage"].Visible = true;
                MainOrdersDataGrid.Columns["ProfilStoreCount"].Visible = true;
                MainOrdersDataGrid.Columns["ProfilExpPercentage"].Visible = true;
                MainOrdersDataGrid.Columns["ProfilExpCount"].Visible = true;
                MainOrdersDataGrid.Columns["ProfilPackedCount"].Visible = true;
                MainOrdersDataGrid.Columns["ProfilDispPercentage"].Visible = true;
                MainOrdersDataGrid.Columns["ProfilDispatchedCount"].Visible = true;
                MainOrdersDataGrid.Columns["ProfilProductionDate"].Visible = true;
                MainOrdersDataGrid.Columns["ProfilOnProductionDate"].Visible = true;

                MainOrdersDataGrid.Columns["TPSProductionStatusColumn"].Visible = false;
                MainOrdersDataGrid.Columns["TPSStorageStatusColumn"].Visible = false;
                MainOrdersDataGrid.Columns["TPSDispatchStatusColumn"].Visible = false;
                MainOrdersDataGrid.Columns["TPSPackPercentage"].Visible = false;
                MainOrdersDataGrid.Columns["TPSPackedCount"].Visible = false;
                MainOrdersDataGrid.Columns["TPSStorePercentage"].Visible = false;
                MainOrdersDataGrid.Columns["TPSStoreCount"].Visible = false;
                MainOrdersDataGrid.Columns["TPSExpPercentage"].Visible = false;
                MainOrdersDataGrid.Columns["TPSExpCount"].Visible = false;
                MainOrdersDataGrid.Columns["TPSDispPercentage"].Visible = false;
                MainOrdersDataGrid.Columns["TPSDispatchedCount"].Visible = false;
                MainOrdersDataGrid.Columns["TPSProductionDate"].Visible = false;
                MainOrdersDataGrid.Columns["TPSOnProductionDate"].Visible = false;


                MegaOrdersDataGrid.Columns["ProfilPackPercentage"].Visible = true;
                MegaOrdersDataGrid.Columns["ProfilPackedCount"].Visible = true;
                MegaOrdersDataGrid.Columns["ProfilStorePercentage"].Visible = true;
                MegaOrdersDataGrid.Columns["ProfilStoreCount"].Visible = true;
                MegaOrdersDataGrid.Columns["ProfilExpPercentage"].Visible = true;
                MegaOrdersDataGrid.Columns["ProfilExpCount"].Visible = true;
                MegaOrdersDataGrid.Columns["ProfilDispPercentage"].Visible = true;
                MegaOrdersDataGrid.Columns["ProfilDispatchedCount"].Visible = true;
                MegaOrdersDataGrid.Columns["ProfilDispatchDate"].Visible = true;
                MegaOrdersDataGrid.Columns["ProfilOrderStatusColumn"].Visible = true;

                MegaOrdersDataGrid.Columns["TPSPackPercentage"].Visible = false;
                MegaOrdersDataGrid.Columns["TPSPackedCount"].Visible = false;
                MegaOrdersDataGrid.Columns["TPSStorePercentage"].Visible = false;
                MegaOrdersDataGrid.Columns["TPSStoreCount"].Visible = false;
                MegaOrdersDataGrid.Columns["TPSExpPercentage"].Visible = false;
                MegaOrdersDataGrid.Columns["TPSExpCount"].Visible = false;
                MegaOrdersDataGrid.Columns["TPSDispPercentage"].Visible = false;
                MegaOrdersDataGrid.Columns["TPSDispatchedCount"].Visible = false;
                MegaOrdersDataGrid.Columns["TPSDispatchDate"].Visible = false;
                MegaOrdersDataGrid.Columns["TPSOrderStatusColumn"].Visible = false;
            }

            if (FactoryID == 2)
            {

                MegaOrdersDataGrid.Columns["ProfilPackPercentage"].Visible = false;
                MegaOrdersDataGrid.Columns["ProfilPackedCount"].Visible = false;
                MegaOrdersDataGrid.Columns["ProfilStorePercentage"].Visible = false;
                MegaOrdersDataGrid.Columns["ProfilStoreCount"].Visible = false;
                MegaOrdersDataGrid.Columns["ProfilExpPercentage"].Visible = false;
                MegaOrdersDataGrid.Columns["ProfilExpCount"].Visible = false;
                MegaOrdersDataGrid.Columns["ProfilDispPercentage"].Visible = false;
                MegaOrdersDataGrid.Columns["ProfilDispatchedCount"].Visible = false;
                MegaOrdersDataGrid.Columns["ProfilDispatchDate"].Visible = false;
                MegaOrdersDataGrid.Columns["ProfilOrderStatusColumn"].Visible = false;

                MegaOrdersDataGrid.Columns["TPSPackPercentage"].Visible = true;
                MegaOrdersDataGrid.Columns["TPSPackedCount"].Visible = true;
                MegaOrdersDataGrid.Columns["TPSStorePercentage"].Visible = true;
                MegaOrdersDataGrid.Columns["TPSStoreCount"].Visible = true;
                MegaOrdersDataGrid.Columns["TPSExpPercentage"].Visible = true;
                MegaOrdersDataGrid.Columns["TPSExpCount"].Visible = true;
                MegaOrdersDataGrid.Columns["TPSDispPercentage"].Visible = true;
                MegaOrdersDataGrid.Columns["TPSDispatchedCount"].Visible = true;
                MegaOrdersDataGrid.Columns["TPSDispatchDate"].Visible = true;
                MegaOrdersDataGrid.Columns["TPSOrderStatusColumn"].Visible = true;
            }
        }

        private void Initialize()
        {
            monthCalendar1.SelectionStart = DateTime.Today - TimeSpan.FromDays(90);
            monthCalendar2.SelectionStart = DateTime.Today;
            monthCalendar1.TodayDate= DateTime.Today;
            monthCalendar2.TodayDate= DateTime.Today;

            MegaOrdersDataGrid.SelectionChanged -= MegaOrdersDataGrid_SelectionChanged;
            MarketingExpeditionManager = new MarketingExpeditionManager(
                ref MainOrdersDataGrid, ref PackagesDataGrid,
                ref MainOrdersFrontsOrdersDataGrid, ref MainOrdersDecorOrdersDataGrid, ref MainOrdersTabControl);

            MegaOrdersDataGrid.DataSource = MarketingExpeditionManager.MegaOrdersBindingSource;

            CreateColumns();

            MegaOrdersDataGrid.Columns.Add(ProfilOrderStatusColumn);
            MegaOrdersDataGrid.Columns.Add(TPSOrderStatusColumn);
            MegaOrdersDataGrid.Columns.Add(MegaFactoryTypeColumn);
            MegaOrdersDataGrid.Columns.Add(CurrencyTypeColumn);

            MegaGridSetting();

            FilterClientsDataGrid.DataSource = MarketingExpeditionManager.FilterClientsBindingSource;
            FilterClientsDataGrid.Columns["ClientID"].Visible = false;

            BatchDataGrid.DataSource = MarketingExpeditionManager.BatchDetailsBindingSource;
            FrontsCatalogOrder = new FrontsCatalogOrder();
            FrontsCatalogOrder.Initialize();
            DecorCatalogOrder = new DecorCatalogOrder();
            DecorCatalogOrder.Initialize();
            DispatchReport = new DispatchReport();
            cabFurAssembleReport = new CabFurAssembleReport();

            MonthsDT = new DataTable();
            MonthsDT.Columns.Add(new DataColumn("MonthID", Type.GetType("System.Int32")));
            MonthsDT.Columns.Add(new DataColumn("MonthName", Type.GetType("System.String")));

            YearsDT = new DataTable();
            YearsDT.Columns.Add(new DataColumn("YearID", Type.GetType("System.Int32")));
            YearsDT.Columns.Add(new DataColumn("YearName", Type.GetType("System.String")));

            for (int i = 1; i <= 12; i++)
            {
                DataRow NewRow = MonthsDT.NewRow();
                NewRow["MonthID"] = i;
                NewRow["MonthName"] = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(i).ToString();
                MonthsDT.Rows.Add(NewRow);
            }
            cbxMonths.DataSource = MonthsDT.DefaultView;
            cbxMonths.ValueMember = "MonthID";
            cbxMonths.DisplayMember = "MonthName";

            cbxCabFurMonths.DataSource = MonthsDT.DefaultView;
            cbxCabFurMonths.ValueMember = "MonthID";
            cbxCabFurMonths.DisplayMember = "MonthName";

            DateTime LastDay = new System.DateTime(DateTime.Now.Year + 1, 12, 31);
            System.Collections.ArrayList Years = new System.Collections.ArrayList();
            for (int i = 2013; i <= LastDay.Year; i++)
            {
                DataRow NewRow = YearsDT.NewRow();
                NewRow["YearID"] = i;
                NewRow["YearName"] = i;
                YearsDT.Rows.Add(NewRow);
                Years.Add(i);
            }
            cbxYears.DataSource = YearsDT.DefaultView;
            cbxYears.ValueMember = "YearID";
            cbxYears.DisplayMember = "YearName";

            cbxCabFurYears.DataSource = YearsDT.DefaultView;
            cbxCabFurYears.ValueMember = "YearID";
            cbxCabFurYears.DisplayMember = "YearName";

            cbxMonths.SelectedValue = DateTime.Now.Month;
            cbxYears.SelectedValue = DateTime.Now.Year;

            cbxCabFurMonths.SelectedValue = DateTime.Now.Month;
            cbxCabFurYears.SelectedValue = DateTime.Now.Year;

            monthCalendar3.SelectionStart = DateTime.Now;

            _dispatchReportToDbf = new DispatchReportToDbf(FrontsCatalogOrder, DecorCatalogOrder);
            _colorInvoiceReportToDbf = new ColorInvoiceReportToDbf(FrontsCatalogOrder, DecorCatalogOrder);
            _notesInvoiceReportToDbf = new NotesInvoiceReportToDbf(FrontsCatalogOrder, DecorCatalogOrder);
            OrdersCalculate = new OrdersCalculate();
            DetailsReport = new Modules.Marketing.Dispatch.DetailsReport(FrontsCatalogOrder, DecorCatalogOrder, ref OrdersCalculate.FrontsCalculate);
            MarketingDispatchManager = new MarketingDispatch();
            iDBFReport = new Modules.Marketing.Orders.InvoiceReportToDbf.InvoiceReportToDbf(FrontsCatalogOrder, DecorCatalogOrder);

            MarketingDispatchManager.Initialize();

            dgvDispatchSetting();
            dgvDispatchDatesSetting();
            dgvMegaOrdersSetting();
            kryptonCheckSet1_CheckedButtonChanged(null, null);

        }

        private void MegaOrdersDataGrid_SelectionChanged(object sender, EventArgs e)
        {
            if (MarketingExpeditionManager != null)
                if (MarketingExpeditionManager.MegaOrdersBindingSource.Count > 0)
                {
                    if (((DataRowView)MarketingExpeditionManager.MegaOrdersBindingSource.Current)["MegaOrderID"] != DBNull.Value)
                    {
                        int MegaOrderID = Convert.ToInt32(((DataRowView)MarketingExpeditionManager.MegaOrdersBindingSource.Current)["MegaOrderID"]);
                        int FactoryID = -1;

                        if (ProfilCheckBox.Checked && TPSCheckBox.Checked)
                            FactoryID = 0;
                        if (ProfilCheckBox.Checked && !TPSCheckBox.Checked)
                            FactoryID = 1;
                        if (!ProfilCheckBox.Checked && TPSCheckBox.Checked)
                            FactoryID = 2;

                        bool NotInProduction = NotProductionCheckBox.Checked;
                        bool OnProduction = OnProductionCheckBox.Checked;
                        bool InProduction = InProductionCheckBox.Checked;
                        bool OnStorage = OnStorageCheckBox.Checked;
                        bool OnExpedition = cbOnExpedition.Checked;
                        bool Dispatch = DispatchCheckBox.Checked;

                        bool bMegaBatch = MegaBatchCheckBox.Checked;

                        int MegaBatchID = -1;
                        if (MegaBatchCheckBox.Checked)
                            MegaBatchID = Convert.ToInt32(BatchDataGrid.SelectedRows[0].Cells["MegaBatchID"].Value);

                        if (NeedSplash)
                        {
                            Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                            T.Start();

                            while (!SplashWindow.bSmallCreated) ;

                            if (AllPackagesCheckBox.Checked)

                                MarketingExpeditionManager.FilterPackagesByMegaOrder(
                                    Convert.ToInt32(((DataRowView)MarketingExpeditionManager.MegaOrdersBindingSource.Current)["MegaOrderID"]),

                                    FactoryID);
                            else
                                MarketingExpeditionManager.FilterMainOrdersByMegaOrder(bMegaBatch,
                                    NotInProduction, OnProduction, InProduction, OnStorage, OnExpedition, Dispatch,
                                    Convert.ToInt32(((DataRowView)MarketingExpeditionManager.MegaOrdersBindingSource.Current)["MegaOrderID"]),
                                    MegaBatchID,
                                    FactoryID);

                            while (SplashWindow.bSmallCreated)
                                SmallWaitForm.CloseS = true;
                        }
                        else
                        {
                            if (AllPackagesCheckBox.Checked)
                                MarketingExpeditionManager.FilterPackagesByMegaOrder(
                                    Convert.ToInt32(((DataRowView)MarketingExpeditionManager.MegaOrdersBindingSource.Current)["MegaOrderID"]),
                                    FactoryID);
                            else
                                MarketingExpeditionManager.FilterMainOrdersByMegaOrder(bMegaBatch,
                                    NotInProduction, OnProduction, InProduction, OnStorage, OnExpedition, Dispatch,
                                    Convert.ToInt32(((DataRowView)MarketingExpeditionManager.MegaOrdersBindingSource.Current)["MegaOrderID"]),
                                    MegaBatchID,
                                    FactoryID);
                        }
                    }
                }
                else
                    MarketingExpeditionManager.MainOrdersDataTable.Clear();
        }

        private void MainOrdersDataGrid_SelectionChanged(object sender, EventArgs e)
        {
            if (MarketingExpeditionManager != null)
                if (MarketingExpeditionManager.MainOrdersBindingSource.Count > 0)
                {
                    if (((DataRowView)MarketingExpeditionManager.MainOrdersBindingSource.Current)["MainOrderID"] != DBNull.Value)
                    {
                        int FactoryID = -1;

                        if (ProfilCheckBox.Checked && TPSCheckBox.Checked)
                            FactoryID = 0;
                        if (ProfilCheckBox.Checked && !TPSCheckBox.Checked)
                            FactoryID = 1;
                        if (!ProfilCheckBox.Checked && TPSCheckBox.Checked)
                            FactoryID = 2;

                        MarketingExpeditionManager.FilterProductsByPackage(Convert.ToInt32(((DataRowView)MarketingExpeditionManager.MainOrdersBindingSource.Current)["MainOrderID"]), FactoryID);
                        MarketingExpeditionManager.FilterPackagesByMainOrder(Convert.ToInt32(((DataRowView)MarketingExpeditionManager.MainOrdersBindingSource.Current)["MainOrderID"]), FactoryID);
                    }
                }
                else
                {
                    MarketingExpeditionManager.FilterProductsByPackage(-1, 0);
                    MarketingExpeditionManager.PackagesDataTable.Clear();
                    MarketingExpeditionManager.PackedMainOrdersFrontsOrders.FrontsOrdersDataTable.Clear();
                    MarketingExpeditionManager.PackedMainOrdersDecorOrders.DecorOrdersDataTable.Clear();
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

        private void PackagesDataGrid_SelectionChanged(object sender, EventArgs e)
        {
            if (MarketingExpeditionManager != null)
                if (MarketingExpeditionManager.PackagesBindingSource.Count > 0)
                {
                    if (((DataRowView)MarketingExpeditionManager.PackagesBindingSource.Current)["PackNumber"] != DBNull.Value)
                    {
                        int MainOrderID = Convert.ToInt32(((DataRowView)MarketingExpeditionManager.PackagesBindingSource.Current)["MainOrderID"]);
                        int PackNumber = Convert.ToInt32(((DataRowView)MarketingExpeditionManager.PackagesBindingSource.Current)["PackNumber"]);
                        int ProductType = Convert.ToInt32(((DataRowView)MarketingExpeditionManager.PackagesBindingSource.Current)["ProductType"]);
                        CurrentMainOrder = MainOrderID;
                        CurrentPackNumber = PackNumber;

                        if (ProductType == 0)
                        {
                            MainOrdersTabControl.SelectedTabPage = MainOrdersTabControl.TabPages[0];
                            MarketingExpeditionManager.PackedMainOrdersFrontsOrders.MoveToFrontOrder(MainOrderID, PackNumber);
                            MarketingExpeditionManager.PackedMainOrdersFrontsOrders.SetColor(MainOrderID, PackNumber);
                        }
                        if (ProductType == 1)
                        {
                            MainOrdersTabControl.SelectedTabPage = MainOrdersTabControl.TabPages[1];

                            MarketingExpeditionManager.PackedMainOrdersDecorOrders.MoveToDecorOrder(MainOrderID, PackNumber);
                            MarketingExpeditionManager.PackedMainOrdersDecorOrders.SetColor(MainOrderID, PackNumber);
                        }
                    }
                }
        }

        private void MainOrdersDataGrid_ColumnAdded(object sender, DataGridViewColumnEventArgs e)
        {
            if (e.Column.Name == "PackPercentage")
            {
                e.Column.SortMode = DataGridViewColumnSortMode.Programmatic;
            }
            if (e.Column.Name == "DispPercentage")
            {
                e.Column.SortMode = DataGridViewColumnSortMode.Programmatic;
            }
        }

        private void MegaOrdersDataGrid_ColumnAdded(object sender, DataGridViewColumnEventArgs e)
        {
            if (e.Column.Name == "PackPercentage")
            {
                e.Column.SortMode = DataGridViewColumnSortMode.Programmatic;
            }
            if (e.Column.Name == "DispPercentage")
            {
                e.Column.SortMode = DataGridViewColumnSortMode.Programmatic;
            }
        }

        private void kryptonButton1_Click(object sender, EventArgs e)
        {
            MenuPanel.Visible = !MenuPanel.Visible;
        }

        private void ClientCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (MarketingExpeditionManager != null)
            {
                //bool bClient = ClientCheckBox.Checked;
                //bool bNoDispatched = NoDispatchedCheckBox.Checked;

                //int FactoryID = -1;

                //if (ProfilCheckBox.Checked && TPSCheckBox.Checked)
                //    FactoryID = 0;
                //if (ProfilCheckBox.Checked && !TPSCheckBox.Checked)
                //    FactoryID = 1;
                //if (!ProfilCheckBox.Checked && TPSCheckBox.Checked)
                //    FactoryID = 2;

                //int ClientID = -1;

                //if (ClientCheckBox.Checked)
                //{
                //    ClientID = Convert.ToInt32(FilterClientsDataGrid.SelectedRows[0].Cells["ClientID"].Value);
                //    MarketingExpeditionManager.FilterMegaBatch(bClient, false, true, false, true, !bNoDispatched, ClientID, FactoryID);
                //}
                //else
                //{
                //    MarketingExpeditionManager.FilterMegaBatch(bClient, false, true, false, true, !bNoDispatched, ClientID, FactoryID);
                //}

                Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                T.Start();

                while (!SplashWindow.bSmallCreated) ;
                NeedSplash = false;

                FilterMegaBatch();
                FilterOrders();

                NeedSplash = true;
                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;
            }
        }

        private void FilterClientsDataGrid_SelectionChanged(object sender, EventArgs e)
        {
            if (MarketingExpeditionManager != null && ClientCheckBox.Checked)
            {
                //bool bClient = ClientCheckBox.Checked;
                //bool bNoDispatched = NoDispatchedCheckBox.Checked;

                //int FactoryID = -1;

                //if (ProfilCheckBox.Checked && TPSCheckBox.Checked)
                //    FactoryID = 0;
                //if (ProfilCheckBox.Checked && !TPSCheckBox.Checked)
                //    FactoryID = 1;
                //if (!ProfilCheckBox.Checked && TPSCheckBox.Checked)
                //    FactoryID = 2;

                //int ClientID = Convert.ToInt32(FilterClientsDataGrid.SelectedRows[0].Cells["ClientID"].Value);
                //MarketingExpeditionManager.FilterMegaBatch(bClient, false, true, false, true, !bNoDispatched, ClientID, FactoryID);

                Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                T.Start();

                while (!SplashWindow.bSmallCreated) ;
                NeedSplash = false;

                FilterMegaBatch();
                FilterOrders();

                NeedSplash = true;
                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;
            }
        }

        private void FilterMegaBatch()
        {
            bool bClient = ClientCheckBox.Checked;
            bool NotInProduction = NotProductionCheckBox.Checked;
            bool OnProduction = OnProductionCheckBox.Checked;
            bool InProduction = InProductionCheckBox.Checked;
            bool OnStorage = OnStorageCheckBox.Checked;
            bool OnExpedition = cbOnExpedition.Checked;
            bool Dispatch = DispatchCheckBox.Checked;

            int FactoryID = -1;

            int ClientID = -1;

            if (ClientCheckBox.Checked && FilterClientsDataGrid.SelectedRows.Count > 0)
                ClientID = Convert.ToInt32(FilterClientsDataGrid.SelectedRows[0].Cells["ClientID"].Value);

            if (ProfilCheckBox.Checked && TPSCheckBox.Checked)
                FactoryID = 0;
            if (ProfilCheckBox.Checked && !TPSCheckBox.Checked)
                FactoryID = 1;
            if (!ProfilCheckBox.Checked && TPSCheckBox.Checked)
                FactoryID = 2;

            if (NeedSplash)
            {
                Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                T.Start();

                while (!SplashWindow.bSmallCreated) ;
                NeedSplash = false;

                MarketingExpeditionManager.FilterMegaBatch(bClient, NotInProduction, OnProduction, InProduction, OnStorage, OnExpedition, Dispatch, ClientID, FactoryID);
                NeedSplash = true;
                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;
            }
            else
            {
                MarketingExpeditionManager.FilterMegaBatch(bClient, NotInProduction, OnProduction, InProduction, OnStorage, OnExpedition, Dispatch, ClientID, FactoryID);
            }
        }

        private void FilterOrders()
        {
            DateTime firstDate = monthCalendar1.SelectionStart;
            DateTime secondDate = monthCalendar2.SelectionStart;

            bool bClient = ClientCheckBox.Checked;
            bool bMegaBatch = MegaBatchCheckBox.Checked;

            bool NotInProduction = NotProductionCheckBox.Checked;
            bool OnProduction = OnProductionCheckBox.Checked;
            bool InProduction = InProductionCheckBox.Checked;
            bool OnStorage = OnStorageCheckBox.Checked;
            bool OnExpedition = cbOnExpedition.Checked;
            bool Dispatch = DispatchCheckBox.Checked;

            int ClientID = -1;
            int FactoryID = -1;

            int MegaBatchID = -1;

            if (ClientCheckBox.Checked && FilterClientsDataGrid.SelectedRows.Count > 0)
                ClientID = Convert.ToInt32(FilterClientsDataGrid.SelectedRows[0].Cells["ClientID"].Value);
            if (MegaBatchCheckBox.Checked && BatchDataGrid.SelectedRows.Count > 0)
                MegaBatchID = Convert.ToInt32(BatchDataGrid.SelectedRows[0].Cells["MegaBatchID"].Value);

            if (ProfilCheckBox.Checked && TPSCheckBox.Checked)
                FactoryID = 0;
            if (ProfilCheckBox.Checked && !TPSCheckBox.Checked)
                FactoryID = 1;
            if (!ProfilCheckBox.Checked && TPSCheckBox.Checked)
                FactoryID = 2;

            if (NeedSplash)
            {
                Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                T.Start();

                while (!SplashWindow.bSmallCreated) ;
                NeedSplash = false;

                MarketingExpeditionManager.Filter(firstDate, secondDate, bClient, bMegaBatch, 
                    NotInProduction, OnProduction, InProduction, OnStorage, OnExpedition, Dispatch, ClientID, MegaBatchID, FactoryID);
                ShowColumns(FactoryID);

                NeedSplash = true;
                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;
            }
            else
            {
                MarketingExpeditionManager.Filter(firstDate, secondDate, bClient, bMegaBatch, 
                    NotInProduction, OnProduction, InProduction, OnStorage, OnExpedition, Dispatch, ClientID, MegaBatchID, FactoryID);
                ShowColumns(FactoryID);
            }
        }

        private void NoDispatchedCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (MarketingExpeditionManager != null)
            {
                Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                T.Start();

                while (!SplashWindow.bSmallCreated) ;
                NeedSplash = false;

                FilterMegaBatch();
                FilterOrders();

                NeedSplash = true;
                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;
            }
        }

        private void ProfilCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (MarketingExpeditionManager != null)
            {
                Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                T.Start();

                while (!SplashWindow.bSmallCreated) ;
                NeedSplash = false;

                FilterMegaBatch();
                FilterOrders();

                NeedSplash = true;
                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;
            }
        }

        private void TPSCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (MarketingExpeditionManager != null)
            {
                Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                T.Start();

                while (!SplashWindow.bSmallCreated) ;
                NeedSplash = false;

                FilterMegaBatch();
                FilterOrders();

                NeedSplash = true;
                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;
            }
        }

        private void AllPackagesCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (AllPackagesCheckBox.Checked)
            {
                MainOrdersDataGrid.SelectionChanged -= MainOrdersDataGrid_SelectionChanged;

                if (NeedSplash)
                {
                    NeedSplash = false;
                    Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                    T.Start();

                    while (!SplashWindow.bSmallCreated) ;

                    int FactoryID = -1;

                    if (ProfilCheckBox.Checked && TPSCheckBox.Checked)
                        FactoryID = 0;
                    if (ProfilCheckBox.Checked && !TPSCheckBox.Checked)
                        FactoryID = 1;
                    if (!ProfilCheckBox.Checked && TPSCheckBox.Checked)
                        FactoryID = 2;

                    MarketingExpeditionManager.FilterPackagesByMegaOrder(Convert.ToInt32(((DataRowView)MarketingExpeditionManager.MegaOrdersBindingSource.Current)["MegaOrderID"]),
                        FactoryID);

                    while (SplashWindow.bSmallCreated)
                        SmallWaitForm.CloseS = true;

                    NeedSplash = true;
                }
            }
            else
            {
                MainOrdersDataGrid.SelectionChanged += MainOrdersDataGrid_SelectionChanged;
                MegaOrdersDataGrid_SelectionChanged(null, null);
            }
        }

        private void kryptonButton1_Click_1(object sender, EventArgs e)
        {
            if (MainOrdersDataGrid.SelectedRows.Count == 0)
                return;
            CheckOrdersStatus.SetToDispatch(monthCalendar3.SelectionStart, GetSelectedPackages(),
                Convert.ToInt32(MainOrdersDataGrid.SelectedRows[0].Cells["MainOrderID"].Value));
        }

        public int[] GetSelectedPackages()
        {
            System.Collections.ArrayList array = new System.Collections.ArrayList();

            for (int i = 0; i < PackagesDataGrid.SelectedRows.Count; i++)
                array.Add(Convert.ToInt32(PackagesDataGrid.SelectedRows[i].Cells["PackageID"].Value));

            int[] rows = array.OfType<int>().ToArray();
            Array.Sort(rows);

            return rows;
        }

        private void kryptonButton2_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < MegaOrdersDataGrid.Rows.Count; i++)
            {
                //CheckOrdersStatus.GG(
                //    Convert.ToInt32(Convert.ToInt32(MegaOrdersDataGrid.Rows[i].Cells["MegaOrderID"].Value)));
            }

            for (int i = 0; i < MegaOrdersDataGrid.SelectedRows.Count; i++)
            {
                CheckOrdersStatus.GG(
                    Convert.ToInt32(Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[i].Cells["MegaOrderID"].Value)));
            }
            if (PackagesDataGrid.SelectedRows.Count == 0)
                CheckOrdersStatus.DispatchMegaOrder(Convert.ToInt32(((DataRowView)MarketingExpeditionManager.MegaOrdersBindingSource.Current)["MegaOrderID"]));
        }

        private void BatchDataGrid_SelectionChanged(object sender, EventArgs e)
        {
            if (MegaBatchCheckBox.Checked)
            {
                if (MarketingExpeditionManager != null)
                    FilterOrders();
            }
        }

        private void MegaBatchCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (MarketingExpeditionManager != null)
                FilterOrders();
        }

        private void ClientCheckBox_Click(object sender, EventArgs e)
        {
            //if (BatchCheckBox.Checked)
            //    BatchCheckBox.Checked = false;
        }

        private void NoDispatchedCheckBox_Click(object sender, EventArgs e)
        {
            //if (BatchCheckBox.Checked)
            //    BatchCheckBox.Checked = false;
        }

        private void ProfilCheckBox_Click(object sender, EventArgs e)
        {
            //if (BatchCheckBox.Checked)
            //    BatchCheckBox.Checked = false;
        }

        private void TPSCheckBox_Click(object sender, EventArgs e)
        {
            //if (BatchCheckBox.Checked)
            //    BatchCheckBox.Checked = false;
        }

        private void AllPackagesCheckBox_Click(object sender, EventArgs e)
        {
            //if (BatchCheckBox.Checked)
            //    BatchCheckBox.Checked = false;
        }

        private void kryptonButton3_Click(object sender, EventArgs e)
        {
            CheckOrdersStatus.SetMainOrderStatus(true, Convert.ToInt32(MainOrdersDataGrid.SelectedRows[0].Cells["MainOrderID"].Value), false);
            CheckOrdersStatus.SetMegaOrderStatus(Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[0].Cells["MegaOrderID"].Value));
            //int ClientID = Convert.ToInt32(FilterClientsDataGrid.SelectedRows[0].Cells["ClientID"].Value);

            //MarketingExpeditionManager.SearchPackedOrders(ClientID);
        }

        private void NotProductionCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (MarketingExpeditionManager != null)
            {
                Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                T.Start();

                while (!SplashWindow.bSmallCreated) ;
                NeedSplash = false;

                FilterMegaBatch();
                FilterOrders();

                NeedSplash = true;
                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;
            }
        }

        private void OnProductionCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (MarketingExpeditionManager != null)
            {
                Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                T.Start();

                while (!SplashWindow.bSmallCreated) ;
                NeedSplash = false;

                FilterMegaBatch();
                FilterOrders();

                NeedSplash = true;
                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;
            }
        }

        private void InProductionCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (MarketingExpeditionManager != null)
            {
                Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                T.Start();

                while (!SplashWindow.bSmallCreated) ;
                NeedSplash = false;

                FilterMegaBatch();
                FilterOrders();

                NeedSplash = true;
                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;
            }
        }

        private void OnStorageCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (MarketingExpeditionManager != null)
            {
                Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                T.Start();

                while (!SplashWindow.bSmallCreated) ;
                NeedSplash = false;

                FilterMegaBatch();
                FilterOrders();

                NeedSplash = true;
                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;
            }
        }

        private void DispatchCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (MarketingExpeditionManager != null)
            {
                Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                T.Start();

                while (!SplashWindow.bSmallCreated) ;
                NeedSplash = false;

                FilterMegaBatch();
                FilterOrders();

                NeedSplash = true;
                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;
            }
        }

        private void MinimizeButton_Click(object sender, EventArgs e)
        {
            NeedSplash = false;
            FormEvent = eHide;
            AnimateTimer.Enabled = true;
        }

        private void MegaOrdersDataGrid_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if ((RoleType == RoleTypes.AdminRole || RoleType == RoleTypes.LogisticsRole)
                && e.Button == System.Windows.Forms.MouseButtons.Right)
            {
                kryptonContextMenu1.Show(new Point(Cursor.Position.X - 212, Cursor.Position.Y - 10));
            }
        }

        private void SetProfilDispatchDate_Click(object sender, EventArgs e)
        {
            CurrentFactoryID = 1;
            popupControlContainer1.ShowPopup(barManager1, this.PointToScreen(new System.Drawing.Point(
                   MousePosition.X, MousePosition.Y)));
        }

        private void SetTPSDispatchDate_Click(object sender, EventArgs e)
        {
            CurrentFactoryID = 2;
            popupControlContainer1.ShowPopup(barManager1, this.PointToScreen(new System.Drawing.Point(
                   MousePosition.X, MousePosition.Y)));
        }

        private void ChangeDateButton_Click(object sender, EventArgs e)
        {
            if (MarketingExpeditionManager != null)
            {
                int MegaOrderID = Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[0].Cells["MegaOrderID"].Value);
                DateTime DispatchDate = kryptonMonthCalendar2.SelectionStart;

                MarketingExpeditionManager.SetDispatchDate(CurrentFactoryID, MegaOrderID, DispatchDate);
                popupControlContainer1.HidePopup();
                FilterOrders();
                MarketingExpeditionManager.MoveToMegaOrder(MegaOrderID);
            }
        }

        private void MarketingExpeditionForm_Load(object sender, EventArgs e)
        {
            UpdateDispatchDate();
            UpdateCabFurDispatchDate();
            MenuPanel.BringToFront();
            if (MoveFromPermits)
            {
                MenuButton.Location = MinimizeButton.Location;
                MinimizeButton.Visible = false;
                cbtnDispatch.Checked = true;
                MarketingDispatchManager.MoveToDispatchDate(PrepareDispatchDateTime);
                MarketingDispatchManager.MoveToDispatch(MDispatchID);
            }
        }

        private void cbOnExpedition_CheckedChanged(object sender, EventArgs e)
        {
            if (MarketingExpeditionManager != null)
            {
                Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                T.Start();

                while (!SplashWindow.bSmallCreated) ;
                NeedSplash = false;

                FilterMegaBatch();
                FilterOrders();

                NeedSplash = true;
                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;
            }
        }

        private void kryptonCheckSet1_CheckedButtonChanged(object sender, EventArgs e)
        {
            if (MarketingDispatchManager == null)
                return;
            if (kryptonCheckSet1.CheckedButton == cbtnExpedition)
            {
                flowLayoutPanel3.Visible = false;
                pnlExpedition.BringToFront();
            }
            if (kryptonCheckSet1.CheckedButton == cbtnCabFur)
            {
                flowLayoutPanel3.Visible = false;
                pnlCabFur.BringToFront();
            }
            if (kryptonCheckSet1.CheckedButton == cbtnDispatch)
            {
                flowLayoutPanel3.Visible = true;
                pnlDispatch.BringToFront();
            }
            MenuPanel.BringToFront();
        }

        private void dgvCabFurMegaOrdersSetting()
        {
            dgvCabFurMegaOrders.DataSource = cabFurAssembleManager.MegaOrdersList;

            foreach (DataGridViewColumn Column in dgvCabFurMegaOrders.Columns)
            {
                Column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }

            dgvCabFurMegaOrders.AutoGenerateColumns = false;

            if (dgvCabFurMegaOrders.Columns.Contains("ClientID"))
                dgvCabFurMegaOrders.Columns["ClientID"].Visible = false;
            if (dgvCabFurMegaOrders.Columns.Contains("PrepareDispatchDateTime"))
                dgvCabFurMegaOrders.Columns["PrepareDispatchDateTime"].Visible = false;


            dgvCabFurMegaOrders.Columns["OrderDate"].DefaultCellStyle.Format = "dd.MM.yyyy HH:mm";

            dgvCabFurMegaOrders.Columns["OrderDate"].HeaderText = "Дата\r\nсоздания";
            dgvCabFurMegaOrders.Columns["ClientName"].HeaderText = "Клиент";
            dgvCabFurMegaOrders.Columns["OrderNumber"].HeaderText = "№\r\nзаказа";
            dgvCabFurMegaOrders.Columns["Weight"].HeaderText = "Вес, кг";
            dgvCabFurMegaOrders.Columns["PackagesCount"].HeaderText = "Кол-во\r\nупаковок";
            dgvCabFurMegaOrders.Columns["Status"].HeaderText = "Статус";
            dgvCabFurMegaOrders.Columns["MegaOrderID"].HeaderText = "ID заказа";

            dgvCabFurMegaOrders.Columns["MegaOrderID"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            dgvCabFurMegaOrders.Columns["MegaOrderID"].Width = 80;
            dgvCabFurMegaOrders.Columns["Weight"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            dgvCabFurMegaOrders.Columns["Weight"].Width = 80;
            dgvCabFurMegaOrders.Columns["ClientName"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dgvCabFurMegaOrders.Columns["ClientName"].MinimumWidth = 200;
            dgvCabFurMegaOrders.Columns["PackagesCount"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            dgvCabFurMegaOrders.Columns["PackagesCount"].Width = 90;
            dgvCabFurMegaOrders.Columns["Status"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            dgvCabFurMegaOrders.Columns["Status"].Width = 200;
            dgvCabFurMegaOrders.Columns["OrderDate"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            dgvCabFurMegaOrders.Columns["OrderDate"].Width = 130;
            dgvCabFurMegaOrders.Columns["OrderNumber"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            dgvCabFurMegaOrders.Columns["OrderNumber"].Width = 70;

            int DisplayIndex = 0;
            dgvCabFurMegaOrders.Columns["ClientName"].DisplayIndex = DisplayIndex++;
            dgvCabFurMegaOrders.Columns["OrderNumber"].DisplayIndex = DisplayIndex++;
            dgvCabFurMegaOrders.Columns["OrderDate"].DisplayIndex = DisplayIndex++;
            dgvCabFurMegaOrders.Columns["PackagesCount"].DisplayIndex = DisplayIndex++;
            dgvCabFurMegaOrders.Columns["Weight"].DisplayIndex = DisplayIndex++;
            dgvCabFurMegaOrders.Columns["Status"].DisplayIndex = DisplayIndex++;
            dgvCabFurMegaOrders.Columns["MegaOrderID"].DisplayIndex = DisplayIndex++;

            dgvCabFurMegaOrders.Columns["PackagesCount"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
        }

        private void dgvCabFurDatesSetting()
        {
            dgvCabFurAssembleDates.DataSource = cabFurAssembleManager.AssembleDatesList;

            dgvCabFurAssembleDates.AutoGenerateColumns = false;

            if (dgvCabFurAssembleDates.Columns.Contains("PrepareDateTime"))
            {
                dgvCabFurAssembleDates.Columns["PrepareDateTime"].DefaultCellStyle.Format = "dd MMMM dddd";
                dgvCabFurAssembleDates.Columns["PrepareDateTime"].MinimumWidth = 150;
                dgvCabFurAssembleDates.Columns["PrepareDateTime"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                dgvCabFurAssembleDates.Columns["PrepareDateTime"].DisplayIndex = 0;
            }
            if (dgvCabFurAssembleDates.Columns.Contains("WeekNumber"))
            {
                dgvCabFurAssembleDates.Columns["WeekNumber"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                dgvCabFurAssembleDates.Columns["WeekNumber"].Width = 70;
                dgvCabFurAssembleDates.Columns["WeekNumber"].DisplayIndex = 1;
            }
            if (dgvCabFurAssembleDates.Columns.Contains("DateName"))
            {
                dgvCabFurAssembleDates.Columns["DateName"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                dgvCabFurAssembleDates.Columns["DateName"].Width = 100;
                dgvCabFurAssembleDates.Columns["DateName"].DisplayIndex = 2;
            }
        }

        private void dgvDispatchSetting()
        {
            dgvDispatch.DataSource = MarketingDispatchManager.DispatchList;

            foreach (DataGridViewColumn Column in dgvDispatch.Columns)
            {
                Column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }

            dgvDispatch.AutoGenerateColumns = false;

            //dgvDispatch.Columns["DispatchID"].Visible = false;
            dgvDispatch.Columns["ClientID"].Visible = false;
            dgvDispatch.Columns["NewClientID"].Visible = false;
            dgvDispatch.Columns["ConfirmExpUserID"].Visible = false;
            dgvDispatch.Columns["ConfirmDispUserID"].Visible = false;
            dgvDispatch.Columns["PrepareDispatchDateTime"].Visible = false;

            dgvDispatch.Columns["CreationDateTime"].DefaultCellStyle.Format = "dd.MM.yyyy HH:mm";
            dgvDispatch.Columns["ConfirmDispDateTime"].DefaultCellStyle.Format = "dd.MM.yyyy HH:mm";
            dgvDispatch.Columns["RealDispDateTime"].DefaultCellStyle.Format = "dd.MM.yyyy HH:mm";

            dgvDispatch.Columns["ClientName"].HeaderText = "Клиент";
            dgvDispatch.Columns["NewClientName"].HeaderText = "Переадресовка";
            dgvDispatch.Columns["Weight"].HeaderText = "Вес, кг";
            dgvDispatch.Columns["CreationDateTime"].HeaderText = "Дата\r\nсоздания";
            dgvDispatch.Columns["DispPackagesCount"].HeaderText = "Кол-во\r\nупаковок";
            dgvDispatch.Columns["DispatchStatus"].HeaderText = "Статус";
            dgvDispatch.Columns["ConfirmExpDateTime"].HeaderText = "Эксп-ция\r\nутверждена";
            dgvDispatch.Columns["ConfirmDispDateTime"].HeaderText = "Отгрузка\r\nутверждена";
            dgvDispatch.Columns["ConfirmExpUser"].HeaderText = "Утвердил\r\nэксп-цию";
            dgvDispatch.Columns["ConfirmDispUser"].HeaderText = "Утвердил\r\nотгрузку";
            dgvDispatch.Columns["InMutualSettlement"].HeaderText = "Включена\r\nв счет";
            dgvDispatch.Columns["RealDispDateTime"].HeaderText = "Дата отгрузки";
            dgvDispatch.Columns["DispatchID"].HeaderText = "№ отгр.";

            dgvDispatch.Columns["InMutualSettlement"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            dgvDispatch.Columns["InMutualSettlement"].Width = 80;
            dgvDispatch.Columns["DispatchID"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            dgvDispatch.Columns["DispatchID"].Width = 80;
            dgvDispatch.Columns["Weight"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            dgvDispatch.Columns["Weight"].Width = 80;
            dgvDispatch.Columns["ClientName"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dgvDispatch.Columns["ClientName"].MinimumWidth = 200;
            dgvDispatch.Columns["NewClientName"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dgvDispatch.Columns["NewClientName"].MinimumWidth = 200;
            dgvDispatch.Columns["RealDispDateTime"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            dgvDispatch.Columns["RealDispDateTime"].Width = 130;
            dgvDispatch.Columns["CreationDateTime"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            dgvDispatch.Columns["CreationDateTime"].Width = 130;
            dgvDispatch.Columns["DispPackagesCount"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            dgvDispatch.Columns["DispPackagesCount"].Width = 90;
            dgvDispatch.Columns["DispatchStatus"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dgvDispatch.Columns["DispatchStatus"].MinimumWidth = 200;
            dgvDispatch.Columns["ConfirmExpDateTime"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            dgvDispatch.Columns["ConfirmExpDateTime"].Width = 130;
            dgvDispatch.Columns["ConfirmExpUser"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dgvDispatch.Columns["ConfirmExpUser"].MinimumWidth = 150;
            dgvDispatch.Columns["ConfirmDispDateTime"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            dgvDispatch.Columns["ConfirmDispDateTime"].Width = 130;
            dgvDispatch.Columns["ConfirmDispUser"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dgvDispatch.Columns["ConfirmDispUser"].MinimumWidth = 150;

            int DisplayIndex = 0;
            dgvDispatch.Columns["ClientName"].DisplayIndex = DisplayIndex++;
            dgvDispatch.Columns["CreationDateTime"].DisplayIndex = DisplayIndex++;
            dgvDispatch.Columns["DispPackagesCount"].DisplayIndex = DisplayIndex++;
            dgvDispatch.Columns["Weight"].DisplayIndex = DisplayIndex++;
            dgvDispatch.Columns["DispatchStatus"].DisplayIndex = DisplayIndex++;
            dgvDispatch.Columns["ConfirmExpDateTime"].DisplayIndex = DisplayIndex++;
            dgvDispatch.Columns["ConfirmExpUser"].DisplayIndex = DisplayIndex++;
            dgvDispatch.Columns["ConfirmDispDateTime"].DisplayIndex = DisplayIndex++;
            dgvDispatch.Columns["ConfirmDispUser"].DisplayIndex = DisplayIndex++;
            dgvDispatch.Columns["RealDispDateTime"].DisplayIndex = DisplayIndex++;
            dgvDispatch.Columns["InMutualSettlement"].DisplayIndex = DisplayIndex++;
            dgvDispatch.Columns["DispatchID"].DisplayIndex = DisplayIndex++;
            dgvDispatch.Columns["NewClientName"].DisplayIndex = DisplayIndex++;

            dgvDispatch.Columns["DispPackagesCount"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
        }

        private void dgvDispatchDatesSetting()
        {
            dgvDispatchDates.DataSource = MarketingDispatchManager.DispatchDatesList;

            dgvDispatchDates.AutoGenerateColumns = false;

            if (dgvDispatchDates.Columns.Contains("PrepareDispatchDateTime"))
            {
                dgvDispatchDates.Columns["PrepareDispatchDateTime"].DefaultCellStyle.Format = "dd MMMM dddd";
                dgvDispatchDates.Columns["PrepareDispatchDateTime"].MinimumWidth = 150;
                dgvDispatchDates.Columns["PrepareDispatchDateTime"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                dgvDispatchDates.Columns["PrepareDispatchDateTime"].DisplayIndex = 0;
            }
            if (dgvDispatchDates.Columns.Contains("WeekNumber"))
            {
                dgvDispatchDates.Columns["WeekNumber"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                dgvDispatchDates.Columns["WeekNumber"].Width = 70;
                dgvDispatchDates.Columns["WeekNumber"].DisplayIndex = 1;
            }
        }

        private void dgvCabFurMainOrdersSetting()
        {
            dgvCabFurMainOrders.DataSource = cabFurAssembleManager.MainOrdersList;

            foreach (DataGridViewColumn Column in dgvCabFurMainOrders.Columns)
            {
                Column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }

            dgvCabFurMainOrders.AutoGenerateColumns = false;

            if (dgvCabFurMainOrders.Columns.Contains("ClientID"))
                dgvCabFurMainOrders.Columns["ClientID"].Visible = false;
            if (dgvCabFurMainOrders.Columns.Contains("FactoryID"))
                dgvCabFurMainOrders.Columns["FactoryID"].Visible = false;
            if (dgvCabFurMainOrders.Columns.Contains("MegaOrderID"))
                dgvCabFurMainOrders.Columns["MegaOrderID"].Visible = false;

            if (dgvCabFurMainOrders.Columns.Contains("OrderNumber"))
            {
                dgvCabFurMainOrders.Columns["OrderNumber"].HeaderText = "№ заказа";
                dgvCabFurMainOrders.Columns["OrderNumber"].Width = 100;
                dgvCabFurMainOrders.Columns["OrderNumber"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                dgvCabFurMainOrders.Columns["OrderNumber"].DisplayIndex = 1;
            }
            if (dgvCabFurMainOrders.Columns.Contains("MainOrderID"))
            {
                dgvCabFurMainOrders.Columns["MainOrderID"].HeaderText = "№ подзаказа";
                dgvCabFurMainOrders.Columns["MainOrderID"].Width = 100;
                dgvCabFurMainOrders.Columns["MainOrderID"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                dgvCabFurMainOrders.Columns["MainOrderID"].DisplayIndex = 2;
            }

            dgvCabFurMainOrders.Columns["Weight"].HeaderText = "Вес";
            dgvCabFurMainOrders.Columns["AllPackCount"].HeaderText = "  Кол-во\r\nупаковок";
            dgvCabFurMainOrders.Columns["PackPercentage"].HeaderText = "Скомплектовано, %";

            dgvCabFurMainOrders.Columns["AllPackCount"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            dgvCabFurMainOrders.Columns["AllPackCount"].Width = 85;
            dgvCabFurMainOrders.Columns["Weight"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            dgvCabFurMainOrders.Columns["Weight"].Width = 105;

            dgvCabFurMainOrders.Columns["PackPercentage"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvCabFurMainOrders.AddPercentageColumn("PackPercentage");
            dgvCabFurMainOrders.Columns["PackPercentage"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            dgvCabFurMainOrders.Columns["PackPercentage"].Width = 205;

        }

        private void dgvMegaOrdersSetting()
        {
            dgvMainOrders.DataSource = MarketingDispatchManager.DispatchContentList;

            foreach (DataGridViewColumn Column in dgvMainOrders.Columns)
            {
                Column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }

            dgvMainOrders.AutoGenerateColumns = false;

            if (dgvMainOrders.Columns.Contains("ClientID"))
                dgvMainOrders.Columns["ClientID"].Visible = false;
            if (dgvMainOrders.Columns.Contains("FactoryID"))
                dgvMainOrders.Columns["FactoryID"].Visible = false;
            if (dgvMainOrders.Columns.Contains("ProfilPackAllocStatusID"))
                dgvMainOrders.Columns["ProfilPackAllocStatusID"].Visible = false;
            if (dgvMainOrders.Columns.Contains("TPSPackAllocStatusID"))
                dgvMainOrders.Columns["TPSPackAllocStatusID"].Visible = false;
            if (dgvMainOrders.Columns.Contains("MegaOrderID"))
            {
                dgvMainOrders.Columns["MegaOrderID"].Visible = false;
            }
            if (dgvMainOrders.Columns.Contains("OrderNumber"))
            {
                dgvMainOrders.Columns["OrderNumber"].HeaderText = "№ заказа";
                dgvMainOrders.Columns["OrderNumber"].Width = 100;
                dgvMainOrders.Columns["OrderNumber"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                dgvMainOrders.Columns["OrderNumber"].DisplayIndex = 1;
            }
            if (dgvMainOrders.Columns.Contains("MainOrderID"))
            {
                dgvMainOrders.Columns["MainOrderID"].HeaderText = "№ подзаказа";
                dgvMainOrders.Columns["MainOrderID"].Width = 100;
                dgvMainOrders.Columns["MainOrderID"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                dgvMainOrders.Columns["MainOrderID"].DisplayIndex = 2;
            }
            dgvMainOrders.Columns["MegaBatchID"].HeaderText = "Группа партий";
            dgvMainOrders.Columns["Square"].HeaderText = "Квадратура";
            dgvMainOrders.Columns["Weight"].HeaderText = "Вес";
            dgvMainOrders.Columns["AllPackCount"].HeaderText = "  Кол-во\r\nупаковок";
            dgvMainOrders.Columns["PackPercentage"].HeaderText = "Упаковано, %";
            dgvMainOrders.Columns["StorePercentage"].HeaderText = "Склад, %";
            dgvMainOrders.Columns["ExpPercentage"].HeaderText = "Экспедиция, %";
            dgvMainOrders.Columns["DispPercentage"].HeaderText = "Отгружено, %";

            dgvMainOrders.Columns["MegaBatchID"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            dgvMainOrders.Columns["MegaBatchID"].Width = 125;
            dgvMainOrders.Columns["AllPackCount"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            dgvMainOrders.Columns["AllPackCount"].Width = 85;
            dgvMainOrders.Columns["Weight"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            dgvMainOrders.Columns["Weight"].Width = 105;
            dgvMainOrders.Columns["Square"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            dgvMainOrders.Columns["Square"].Width = 105;

            dgvMainOrders.Columns["PackPercentage"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvMainOrders.AddPercentageColumn("PackPercentage");
            dgvMainOrders.Columns["StorePercentage"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvMainOrders.AddPercentageColumn("StorePercentage");
            dgvMainOrders.Columns["ExpPercentage"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvMainOrders.AddPercentageColumn("ExpPercentage");
            dgvMainOrders.Columns["DispPercentage"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvMainOrders.AddPercentageColumn("DispPercentage");

        }

        private void dgvDispatchDates_SelectionChanged(object sender, EventArgs e)
        {
            if (MarketingDispatchManager == null)
                return;
            MarketingDispatchManager.ClearDispatch();
            if (dgvDispatchDates.SelectedRows.Count == 0)
            {
                return;
            }
            object Date = MarketingDispatchManager.CurrentDispatchDate;

            Infinium.Modules.Marketing.MutualSettlements.MutualSettlements MutualSettlementsManager = new Modules.Marketing.MutualSettlements.MutualSettlements();
            MutualSettlementsManager.Initialize();

            if (Date != DBNull.Value)
            {
                if (NeedSplash)
                {
                    Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                    T.Start();

                    while (!SplashWindow.bSmallCreated) ;
                    NeedSplash = false;
                    MarketingDispatchManager.GetMegaBatchNumbers(Convert.ToDateTime(Date));
                    MarketingDispatchManager.GetMainOrdersSquareAndWeight(Convert.ToDateTime(Date));
                    MarketingDispatchManager.FilterDispatchByDate(Convert.ToDateTime(Date));

                    List<int> Clients = new List<int>();
                    for (int i = 0; i < dgvDispatch.Rows.Count; i++)
                        Clients.Add(Convert.ToInt32(dgvDispatch.Rows[i].Cells["ClientID"].Value));

                    if (Clients.Count > 0)
                        MarketingDispatchManager.DelayOfPayment(Clients);
                    for (int i = 0; i < dgvDispatch.Rows.Count; i++)
                    {
                        int ClientID = Convert.ToInt32(dgvDispatch.Rows[i].Cells["ClientID"].Value);
                        int DispatchID = Convert.ToInt32(dgvDispatch.Rows[i].Cells["DispatchID"].Value);
                        int ProfilMutualSettlementID = 0;
                        int TPSMutualSettlementID = 0;
                        int PaidStatus = -1;

                        DataTable MegaOrdersInDispatchDT = MarketingDispatchManager.GetMegaOrdersInDispatch(DispatchID);
                        ArrayList OrderNumbers = new ArrayList();
                        for (int j = 0; j < MegaOrdersInDispatchDT.Rows.Count; j++)
                            OrderNumbers.Add(Convert.ToInt32(MegaOrdersInDispatchDT.Rows[j]["OrderNumber"]));

                        ProfilMutualSettlementID = MutualSettlementsManager.FindMutualSettlementIDByOrderNumber(ClientID, 1, OrderNumbers.OfType<Int32>().ToArray());
                        TPSMutualSettlementID = MutualSettlementsManager.FindMutualSettlementIDByOrderNumber(ClientID, 2, OrderNumbers.OfType<Int32>().ToArray());
                        if (PaidStatus == -1)
                        {
                            PaidStatus = MarketingDispatchManager.GetPaidStatus(ClientID, ProfilMutualSettlementID);
                            dgvDispatch.Rows[i].Cells["PaidStatus"].Value = PaidStatus;
                        }
                        if (PaidStatus == -1)
                        {
                            PaidStatus = MarketingDispatchManager.GetPaidStatus(ClientID, TPSMutualSettlementID);
                            dgvDispatch.Rows[i].Cells["PaidStatus"].Value = PaidStatus;
                        }
                    }
                    NeedSplash = true;
                    while (SplashWindow.bSmallCreated)
                        SmallWaitForm.CloseS = true;
                }
                else
                {
                    MarketingDispatchManager.GetMegaBatchNumbers(Convert.ToDateTime(Date));
                    MarketingDispatchManager.GetMainOrdersSquareAndWeight(Convert.ToDateTime(Date));
                    MarketingDispatchManager.FilterDispatchByDate(Convert.ToDateTime(Date));

                    List<int> Clients = new List<int>();
                    for (int i = 0; i < dgvDispatch.Rows.Count; i++)
                        Clients.Add(Convert.ToInt32(dgvDispatch.Rows[i].Cells["ClientID"].Value));
                    if (Clients.Count > 0)
                        MarketingDispatchManager.DelayOfPayment(Clients);
                    for (int i = 0; i < dgvDispatch.Rows.Count; i++)
                    {
                        int ClientID = Convert.ToInt32(dgvDispatch.Rows[i].Cells["ClientID"].Value);
                        int DispatchID = Convert.ToInt32(dgvDispatch.Rows[i].Cells["DispatchID"].Value);
                        int ProfilMutualSettlementID = 0;
                        int TPSMutualSettlementID = 0;
                        int PaidStatus = -1;

                        DataTable MegaOrdersInDispatchDT = MarketingDispatchManager.GetMegaOrdersInDispatch(DispatchID);
                        ArrayList OrderNumbers = new ArrayList();
                        for (int j = 0; j < MegaOrdersInDispatchDT.Rows.Count; j++)
                            OrderNumbers.Add(Convert.ToInt32(MegaOrdersInDispatchDT.Rows[j]["OrderNumber"]));

                        ProfilMutualSettlementID = MutualSettlementsManager.FindMutualSettlementIDByOrderNumber(ClientID, 1, OrderNumbers.OfType<Int32>().ToArray());
                        TPSMutualSettlementID = MutualSettlementsManager.FindMutualSettlementIDByOrderNumber(ClientID, 2, OrderNumbers.OfType<Int32>().ToArray());
                        if (PaidStatus == -1)
                        {
                            PaidStatus = MarketingDispatchManager.GetPaidStatus(ClientID, ProfilMutualSettlementID);
                            dgvDispatch.Rows[i].Cells["PaidStatus"].Value = PaidStatus;
                        }
                        if (PaidStatus == -1)
                        {
                            PaidStatus = MarketingDispatchManager.GetPaidStatus(ClientID, TPSMutualSettlementID);
                            dgvDispatch.Rows[i].Cells["PaidStatus"].Value = PaidStatus;
                        }
                    }
                }
            }
            if (dgvDispatch.SelectedRows[0].Cells["DispatchStatus"].Value.ToString() == "Отгружена")
            {
                btnConfirmDispatch.Visible = false;
                ConfirmDispatchContextMenuItem.Visible = false;
            }
            if (RoleType == RoleTypes.AdminRole || RoleType == RoleTypes.LogisticsRole)
            {
                btnChangeDispatchDate.Visible = true;
                cmiChangeDispatchDate.Visible = true;
                btnEditDispatch.Visible = true;
            }
            if (dgvDispatch.SelectedRows[0].Cells["DispatchStatus"].Value.ToString() == "Отгружена")
            {
                btnConfirmDispatch.Visible = false;
                ConfirmDispatchContextMenuItem.Visible = false;
                btnChangeDispatchDate.Visible = false;
                cmiChangeDispatchDate.Visible = false;
                btnEditDispatch.Visible = false;
            }
        }

        private void dgvCabFurDates_SelectionChanged(object sender, EventArgs e)
        {
            if (cabFurAssembleManager == null)
                return;
            cabFurAssembleManager.ClearMegaOrders();
            if (dgvCabFurAssembleDates.SelectedRows.Count == 0)
            {
                return;
            }
            object Date = cabFurAssembleManager.CurrentDate;

            if (Date != DBNull.Value)
            {
                if (NeedSplash)
                {
                    Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                    T.Start();

                    while (!SplashWindow.bSmallCreated) ;
                    NeedSplash = false;
                    //CabFurDispatchManager.GetMegaBatchNumbers(Convert.ToDateTime(Date));
                    cabFurAssembleManager.GetMainOrdersSquareAndWeight(Convert.ToDateTime(Date));
                    cabFurAssembleManager.FilterAssembleByDate(Convert.ToDateTime(Date));

                    NeedSplash = true;
                    while (SplashWindow.bSmallCreated)
                        SmallWaitForm.CloseS = true;
                }
                else
                {
                    //CabFurDispatchManager.GetMegaBatchNumbers(Convert.ToDateTime(Date));
                    cabFurAssembleManager.GetMainOrdersSquareAndWeight(Convert.ToDateTime(Date));
                    cabFurAssembleManager.FilterAssembleByDate(Convert.ToDateTime(Date));
                }
            }
            else
            {
                if (NeedSplash)
                {
                    Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                    T.Start();

                    while (!SplashWindow.bSmallCreated) ;
                    NeedSplash = false;
                    //CabFurDispatchManager.GetMegaBatchNumbers(Convert.ToDateTime(Date));
                    cabFurAssembleManager.GetMainOrdersSquareAndWeight();
                    cabFurAssembleManager.FilterAssembleByID(-1);

                    NeedSplash = true;
                    while (SplashWindow.bSmallCreated)
                        SmallWaitForm.CloseS = true;
                }
                else
                {
                    //CabFurDispatchManager.GetMegaBatchNumbers(Convert.ToDateTime(Date));
                    cabFurAssembleManager.GetMainOrdersSquareAndWeight();
                    cabFurAssembleManager.FilterAssembleByID(-1);
                }
            }
        }

        private void UpdateDispatchDate()
        {
            DateTime FilterDate = new DateTime(Convert.ToInt32(cbxYears.SelectedValue), Convert.ToInt32(cbxMonths.SelectedValue), 1);
            MarketingDispatchManager.ClearDispatchDates();
            MarketingDispatchManager.UpdateDispatchDates(FilterDate);
        }

        private void UpdateCabFurDispatchDate()
        {
            if (cabFurAssembleManager == null)
                return;
            DateTime FilterDate = new DateTime(Convert.ToInt32(cbxCabFurYears.SelectedValue), Convert.ToInt32(cbxCabFurMonths.SelectedValue), 1);
            cabFurAssembleManager.ClearAssembleDates();
            cabFurAssembleManager.UpdateAssembleDates(FilterDate);
        }

        private void cbxMonths_SelectionChangeCommitted(object sender, EventArgs e)
        {
            UpdateDispatchDate();
        }

        private void cbxYears_SelectionChangeCommitted(object sender, EventArgs e)
        {
            UpdateDispatchDate();
        }

        private void cbxCabFurMonths_SelectionChangeCommitted(object sender, EventArgs e)
        {
            UpdateCabFurDispatchDate();
        }

        private void cbxCabFurYears_SelectionChangeCommitted(object sender, EventArgs e)
        {
            UpdateCabFurDispatchDate();
        }

        private void dgvDispatch_SelectionChanged(object sender, EventArgs e)
        {
            if (MarketingDispatchManager == null)
                return;
            MarketingDispatchManager.ClearDispatchContent();
            if (dgvDispatch.SelectedRows.Count == 0)
            {
                return;
            }
            int DispatchID = Convert.ToInt32(dgvDispatch.SelectedRows[0].Cells["DispatchID"].Value);

            if (dgvDispatch.SelectedRows[0].Cells["ConfirmDispDateTime"].Value != DBNull.Value
                && (RoleType == RoleTypes.AdminRole || RoleType == RoleTypes.LogisticsRole))
            {
                CanEditDispatch = false;
            }
            else
            {
                CanEditDispatch = true;
            }

            if (RoleType != RoleTypes.OrdinaryRole && RoleType != RoleTypes.DispatchRole)
            {
                if (dgvDispatch.SelectedRows[0].Cells["ConfirmDispDateTime"].Value == DBNull.Value)
                {
                    if (RoleType == RoleTypes.AdminRole || RoleType == RoleTypes.LogisticsRole)
                    {
                        btnConfirmExpedition.Visible = true;
                        ConfirmExpContextMenuItem.Visible = true;
                    }
                    if (dgvDispatch.SelectedRows[0].Cells["ConfirmExpDateTime"].Value == DBNull.Value)
                    {
                        btnConfirmDispatch.Visible = false;
                        ConfirmDispatchContextMenuItem.Visible = false;
                        btnConfirmExpedition.Text = "Разрешить\r\nэксп-цию";
                        ConfirmExpContextMenuItem.Text = "Разрешить\r\nэксп-цию";
                    }
                    else
                    {
                        if (RoleType == RoleTypes.AdminRole || RoleType == RoleTypes.ApprovedRole)
                        {
                            btnConfirmDispatch.Visible = true;
                            ConfirmDispatchContextMenuItem.Visible = true;
                        }
                        btnConfirmExpedition.Text = "Запретить\r\nэксп-цию";
                        ConfirmExpContextMenuItem.Text = "Запретить\r\nэксп-цию";
                    }
                    btnConfirmDispatch.Text = "Разрешить\r\nотгрузку";
                    ConfirmDispatchContextMenuItem.Text = "Разрешить\r\nотгрузку";
                }
                else
                {
                    if (RoleType == RoleTypes.AdminRole || RoleType == RoleTypes.ApprovedRole)
                    {
                        btnConfirmDispatch.Visible = true;
                        ConfirmDispatchContextMenuItem.Visible = true;
                    }
                    btnConfirmExpedition.Visible = false;
                    ConfirmExpContextMenuItem.Visible = false;
                    if (dgvDispatch.SelectedRows[0].Cells["ConfirmExpDateTime"].Value == DBNull.Value)
                    {
                        btnConfirmExpedition.Text = "Разрешить\r\nэксп-цию";
                        ConfirmExpContextMenuItem.Text = "Разрешить\r\nэксп-цию";
                    }
                    else
                    {
                        btnConfirmExpedition.Text = "Запретить\r\nэксп-цию";
                        ConfirmExpContextMenuItem.Text = "Запретить\r\nэксп-цию";
                    }
                    btnConfirmDispatch.Text = "Запретить\r\nотгрузку";
                    ConfirmDispatchContextMenuItem.Text = "Запретить\r\nотгрузку";
                }
            }

            if (NeedSplash)
            {
                Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                T.Start();

                while (!SplashWindow.bSmallCreated) ;
                NeedSplash = false;
                MarketingDispatchManager.FilterDispatchContent(DispatchID);
                MarketingDispatchManager.FillPercColumns(DispatchID);
                NeedSplash = true;
                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;
            }
            else
            {
                MarketingDispatchManager.FilterDispatchContent(DispatchID);
                MarketingDispatchManager.FillPercColumns(DispatchID);
            }
            if (dgvDispatch.SelectedRows.Count > 0 && dgvDispatch.SelectedRows[0].Cells["DispatchStatus"].Value.ToString() == "Отгружена")
            {
                btnConfirmDispatch.Visible = false;
                ConfirmDispatchContextMenuItem.Visible = false;
            }
            if (RoleType == RoleTypes.AdminRole || RoleType == RoleTypes.LogisticsRole)
            {
                btnChangeDispatchDate.Visible = true;
                cmiChangeDispatchDate.Visible = true;
                btnEditDispatch.Visible = true;
            }
            if (dgvDispatch.SelectedRows.Count > 0 && dgvDispatch.SelectedRows[0].Cells["DispatchStatus"].Value.ToString() == "Отгружена")
            {
                btnConfirmDispatch.Visible = false;
                ConfirmDispatchContextMenuItem.Visible = false;
                btnChangeDispatchDate.Visible = false;
                cmiChangeDispatchDate.Visible = false;
                btnEditDispatch.Visible = false;
            }
        }

        private void dgvCabFur_SelectionChanged(object sender, EventArgs e)
        {
            if (cabFurAssembleManager == null)
                return;
            cabFurAssembleManager.ClearMainOrders();
            if (dgvCabFurMegaOrders.SelectedRows.Count == 0)
            {
                return;
            }
            int ClientID = Convert.ToInt32(dgvCabFurMegaOrders.SelectedRows[0].Cells["ClientID"].Value);
            int MegaOrderID = Convert.ToInt32(dgvCabFurMegaOrders.SelectedRows[0].Cells["MegaOrderID"].Value);

            if (NeedSplash)
            {
                Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                T.Start();

                while (!SplashWindow.bSmallCreated) ;
                NeedSplash = false;
                cabFurAssembleManager.FilterMainOrders(MegaOrderID);
                cabFurAssembleManager.FillPercColumns(ClientID, MegaOrderID);
                NeedSplash = true;
                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;
            }
            else
            {
                cabFurAssembleManager.FilterMainOrders(MegaOrderID);
                cabFurAssembleManager.FillPercColumns(ClientID, MegaOrderID);
            }
        }

        private void dgvMegaOrders_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (RoleType != RoleTypes.AdminRole && RoleType != RoleTypes.LogisticsRole)
            {
                CanEditDispatch = false;
            }

            Thread T = new Thread(delegate () { SplashWindow.CreateSplash(); });
            T.Start();

            while (!SplashForm.bCreated) ;
            int ClientID = Convert.ToInt32(dgvDispatch.SelectedRows[0].Cells["ClientID"].Value);
            int DispatchID = Convert.ToInt32(dgvDispatch.SelectedRows[0].Cells["DispatchID"].Value);
            int MainOrderID = Convert.ToInt32(dgvMainOrders.SelectedRows[0].Cells["MainOrderID"].Value);
            int MegaOrderID = Convert.ToInt32(dgvMainOrders.SelectedRows[0].Cells["MegaOrderID"].Value);
            DateTime DispatchDate = Convert.ToDateTime(dgvDispatchDates.SelectedRows[0].Cells["PrepareDispatchDateTime"].Value);

            MarketingNewDispatchForm MarketingNewDispatchForm = new MarketingNewDispatchForm(CanEditDispatch, DispatchDate, ClientID, DispatchID, MegaOrderID, MainOrderID);

            TopForm = MarketingNewDispatchForm;

            MarketingNewDispatchForm.ShowDialog();

            MarketingNewDispatchForm.Close();
            MarketingNewDispatchForm.Dispose();

            Thread T1 = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
            T1.Start();

            while (!SplashWindow.bSmallCreated) ;

            NeedSplash = false;
            UpdateDispatchDate();
            MarketingDispatchManager.MoveToDispatchDate(DispatchDate);
            MarketingDispatchManager.MoveToDispatch(DispatchID);
            NeedSplash = true;
            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;

            TopForm = null;
        }

        private void MegaOrdersDataGrid_KeyPress(object sender, KeyPressEventArgs e)
        {
            //if (Char.IsLetter(e.KeyChar))
            //{
            //    for (int i = 0; i < (MegaOrdersDataGrid.Rows.Count); i++)
            //    {
            //        if (MegaOrdersDataGrid.Rows[i].Cells["ClientName"].Value.ToString().StartsWith(e.KeyChar.ToString(), true, CultureInfo.InvariantCulture))
            //        {
            //            MegaOrdersDataGrid.Rows[i].Cells[0].Selected = true;
            //            return; // stop looping
            //        }
            //    }
            //}
            //if (Char.IsLetter(e.KeyChar))
            //{
            //    for (int i = 0; i < (MegaOrdersDataGrid.Rows.Count); i++)
            //    {
            //        if (MegaOrdersDataGrid.Rows[i].Cells["ClientName"].Value.ToString().StartsWith(e.KeyChar.ToString(), true, CultureInfo.InvariantCulture))
            //        {
            //            if (lastKey == e.KeyChar && lastIndex < i)
            //            {
            //                continue;
            //            }

            //            lastKey = e.KeyChar;
            //            lastIndex = i;
            //            MegaOrdersDataGrid.Rows[i].Cells[0].Selected = true;
            //            return;
            //        }
            //    }
            //}
        }

        public bool ContainsText(string source, string toCheck, StringComparison comp)
        {
            return source.IndexOf(toCheck, comp) >= 0;
        }

        private void kryptonTextBox1_TextChanged(object sender, EventArgs e)
        {
            //string searchText = kryptonTextBox1.Text;

            //FilterClientsDataGrid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            //FilterClientsDataGrid.ClearSelection();
            //try
            //{
            //    for (int i = 0; i < (FilterClientsDataGrid.Rows.Count); i++)
            //    {
            //        if (ContainsText(FilterClientsDataGrid.Rows[i].Cells["ClientName"].Value.ToString(), searchText, StringComparison.OrdinalIgnoreCase))
            //        {
            //            FilterClientsDataGrid.Rows[i].Cells[0].Selected = true;
            //            return; // stop looping
            //        }
            //    }
            //    foreach (DataGridViewRow row in FilterClientsDataGrid.Rows)
            //    {
            //        if (ContainsText(row.Cells["ClientName"].Value.ToString(), searchText, StringComparison.OrdinalIgnoreCase))
            //        {
            //            row.Selected = true;
            //            FilterClientsDataGrid.FirstDisplayedScrollingRowIndex = row.Index;
            //            break;
            //        }
            //    }
            //}
            //catch (Exception exc)
            //{
            //    MessageBox.Show(exc.Message);
            //}
        }

        private void btnEditDispatch_Click(object sender, EventArgs e)
        {
            if (dgvDispatch.Rows.Count == 0)
                return;

            if ((RoleType != RoleTypes.AdminRole && RoleType != RoleTypes.LogisticsRole)
                || dgvDispatch.SelectedRows[0].Cells["DispatchStatus"].Value.ToString() == "Отгружена")
            {
                CanEditDispatch = false;
            }

            int ClientID = Convert.ToInt32(dgvDispatch.SelectedRows[0].Cells["ClientID"].Value);
            int DispatchID = Convert.ToInt32(dgvDispatch.SelectedRows[0].Cells["DispatchID"].Value);
            DateTime DispatchDate = Convert.ToDateTime(dgvDispatchDates.SelectedRows[0].Cells["PrepareDispatchDateTime"].Value);

            Thread T = new Thread(delegate () { SplashWindow.CreateSplash(); });
            T.Start();

            while (!SplashForm.bCreated) ;
            MarketingNewDispatchForm MarketingNewDispatchForm = new MarketingNewDispatchForm(CanEditDispatch, false, DispatchDate, ClientID, DispatchID);

            TopForm = MarketingNewDispatchForm;

            MarketingNewDispatchForm.ShowDialog();

            MarketingNewDispatchForm.Close();
            MarketingNewDispatchForm.Dispose();

            Thread T1 = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
            T1.Start();

            while (!SplashWindow.bSmallCreated) ;

            NeedSplash = false;
            UpdateDispatchDate();
            MarketingDispatchManager.MoveToDispatchDate(DispatchDate);
            MarketingDispatchManager.MoveToDispatch(DispatchID);
            NeedSplash = true;
            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;

            TopForm = null;
        }

        private void btnAddDispatch_Click(object sender, EventArgs e)
        {
            bool PressOK = false;
            int ClientID = 0;
            object DispatchDate = null;

            PhantomForm PhantomForm = new Infinium.PhantomForm();
            PhantomForm.Show();

            MarketingNewDispatchMenu MarketingNewDispatchMenu = new MarketingNewDispatchMenu(this, MarketingExpeditionManager.ClientsDT, false);
            TopForm = MarketingNewDispatchMenu;
            MarketingNewDispatchMenu.ShowDialog();

            PressOK = MarketingNewDispatchMenu.PressOK;
            ClientID = MarketingNewDispatchMenu.ClientID;
            DispatchDate = MarketingNewDispatchMenu.DispatchDate;

            PhantomForm.Close();
            PhantomForm.Dispose();
            MarketingNewDispatchMenu.Dispose();
            TopForm = null;

            if (!PressOK)
                return;

            Thread T = new Thread(delegate () { SplashWindow.CreateSplash(); });
            T.Start();

            while (!SplashForm.bCreated) ;
            int DispatchID = -1;

            MarketingNewDispatchForm MarketingNewDispatchForm = new MarketingNewDispatchForm(CanEditDispatch, true, DispatchDate, ClientID, DispatchID);

            TopForm = MarketingNewDispatchForm;

            MarketingNewDispatchForm.ShowDialog();

            MarketingNewDispatchForm.Close();
            MarketingNewDispatchForm.Dispose();

            Thread T1 = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
            T1.Start();

            while (!SplashWindow.bSmallCreated) ;

            NeedSplash = false;
            UpdateDispatchDate();
            NeedSplash = true;
            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;

            TopForm = null;
        }

        private void ProfilPrintContextMenuItem_Click(object sender, EventArgs e)
        {
            bool NeedProfilList = true;
            bool NeedTPSList = false;
            int ClientID = Convert.ToInt32(dgvDispatch.SelectedRows[0].Cells["ClientID"].Value);
            int DispatchID = Convert.ToInt32(dgvDispatch.SelectedRows[0].Cells["DispatchID"].Value);
            DateTime DispatchDate = Convert.ToDateTime(dgvDispatchDates.SelectedRows[0].Cells["PrepareDispatchDateTime"].Value);

            int[] Dispatches = new int[dgvDispatch.SelectedRows.Count];
            for (int i = 0; i < dgvDispatch.SelectedRows.Count; i++)
                Dispatches[i] = Convert.ToInt32(dgvDispatch.SelectedRows[i].Cells["DispatchID"].Value);

            if (!MarketingDispatchManager.HasPackages(DispatchID))
            {
                InfiniumTips.ShowTip(this, 50, 85, "Отгрузка пуста", 1700);
                return;
            }

            bool PressOK = false;
            bool ColorFullName = false;
            object MachineName = DBNull.Value;
            object PermitNumber = DBNull.Value;
            object SealNumber = DBNull.Value;

            PhantomForm PhantomForm = new Infinium.PhantomForm();
            PhantomForm.Show();

            MarketingDispatchInfoMenu MarketingDispatchInfoMenu = new MarketingDispatchInfoMenu(this);
            TopForm = MarketingDispatchInfoMenu;
            MarketingDispatchInfoMenu.ShowDialog();

            PressOK = MarketingDispatchInfoMenu.PressOK;
            ColorFullName = MarketingDispatchInfoMenu.ColorFullName;
            MachineName = MarketingDispatchInfoMenu.MachineName;
            PermitNumber = MarketingDispatchInfoMenu.PermitNumber;
            SealNumber = MarketingDispatchInfoMenu.SealNumber;

            PhantomForm.Close();
            PhantomForm.Dispose();
            MarketingDispatchInfoMenu.Dispose();
            TopForm = null;

            Thread T1 = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Создание документа Excel.\r\nПодождите..."); });
            T1.Start();

            while (!SplashWindow.bSmallCreated) ;

            object CreationDateTime = dgvDispatch.SelectedRows[0].Cells["CreationDateTime"].Value;
            object ConfirmExpDateTime = dgvDispatch.SelectedRows[0].Cells["ConfirmExpDateTime"].Value;
            object ConfirmDispDateTime = dgvDispatch.SelectedRows[0].Cells["ConfirmDispDateTime"].Value;
            object PrepareDispDateTime = dgvDispatchDates.SelectedRows[0].Cells["PrepareDispatchDateTime"].Value;
            object ConfirmExpUserID = dgvDispatch.SelectedRows[0].Cells["ConfirmExpUserID"].Value;
            object ConfirmDispUserID = dgvDispatch.SelectedRows[0].Cells["ConfirmDispUserID"].Value;
            object RealDispDateTime = DBNull.Value;
            object DispUserID = DBNull.Value;
            string PackagesReportName = string.Empty;
            MarketingDispatchManager.GetRealDispDateTime(DispatchID, ref RealDispDateTime, ref DispUserID);

            DispatchReport.GetDispatchInfo(ref CreationDateTime, ref ConfirmExpDateTime, ref ConfirmDispDateTime, ref RealDispDateTime, ref PrepareDispDateTime,
                ref ConfirmExpUserID, ref ConfirmDispUserID, ref DispUserID, ref MachineName, ref PermitNumber, ref SealNumber);
            DispatchReport.CurrentClient = ClientID;
            DispatchReport.CurrentDispatches = Dispatches;
            DispatchReport.Initialize();
            DispatchReport.CreateReport(NeedProfilList, NeedTPSList, false, ColorFullName, true, ref PackagesReportName);

            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
        }

        private void TPSPrintContextMenuItem_Click(object sender, EventArgs e)
        {
            bool NeedProfilList = false;
            bool NeedTPSList = true;
            int ClientID = Convert.ToInt32(dgvDispatch.SelectedRows[0].Cells["ClientID"].Value);
            int DispatchID = Convert.ToInt32(dgvDispatch.SelectedRows[0].Cells["DispatchID"].Value);

            int[] Dispatches = new int[dgvDispatch.SelectedRows.Count];
            for (int i = 0; i < dgvDispatch.SelectedRows.Count; i++)
                Dispatches[i] = Convert.ToInt32(dgvDispatch.SelectedRows[i].Cells["DispatchID"].Value);

            if (!MarketingDispatchManager.HasPackages(DispatchID))
            {
                InfiniumTips.ShowTip(this, 50, 85, "Отгрузка пуста", 1700);
                return;
            }

            bool PressOK = false;
            bool ColorFullName = false;
            object MachineName = DBNull.Value;
            object PermitNumber = DBNull.Value;
            object SealNumber = DBNull.Value;

            PhantomForm PhantomForm = new Infinium.PhantomForm();
            PhantomForm.Show();

            MarketingDispatchInfoMenu MarketingDispatchInfoMenu = new MarketingDispatchInfoMenu(this);
            TopForm = MarketingDispatchInfoMenu;
            MarketingDispatchInfoMenu.ShowDialog();

            PressOK = MarketingDispatchInfoMenu.PressOK;
            ColorFullName = MarketingDispatchInfoMenu.ColorFullName;
            MachineName = MarketingDispatchInfoMenu.MachineName;
            PermitNumber = MarketingDispatchInfoMenu.PermitNumber;
            SealNumber = MarketingDispatchInfoMenu.SealNumber;

            PhantomForm.Close();
            PhantomForm.Dispose();
            MarketingDispatchInfoMenu.Dispose();
            TopForm = null;

            Thread T1 = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Создание документа Excel.\r\nПодождите..."); });
            T1.Start();

            while (!SplashWindow.bSmallCreated) ;

            object CreationDateTime = dgvDispatch.SelectedRows[0].Cells["CreationDateTime"].Value;
            object ConfirmExpDateTime = dgvDispatch.SelectedRows[0].Cells["ConfirmExpDateTime"].Value;
            object ConfirmDispDateTime = dgvDispatch.SelectedRows[0].Cells["ConfirmDispDateTime"].Value;
            object PrepareDispDateTime = dgvDispatchDates.SelectedRows[0].Cells["PrepareDispatchDateTime"].Value;
            object ConfirmExpUserID = dgvDispatch.SelectedRows[0].Cells["ConfirmExpUserID"].Value;
            object ConfirmDispUserID = dgvDispatch.SelectedRows[0].Cells["ConfirmDispUserID"].Value;
            object RealDispDateTime = DBNull.Value;
            object DispUserID = DBNull.Value;
            string PackagesReportName = string.Empty;
            MarketingDispatchManager.GetRealDispDateTime(DispatchID, ref RealDispDateTime, ref DispUserID);

            DispatchReport.GetDispatchInfo(ref CreationDateTime, ref ConfirmExpDateTime, ref ConfirmDispDateTime, ref RealDispDateTime, ref PrepareDispDateTime,
                ref ConfirmExpUserID, ref ConfirmDispUserID, ref DispUserID, ref MachineName, ref PermitNumber, ref SealNumber);
            DispatchReport.CurrentClient = ClientID;
            DispatchReport.CurrentDispatches = Dispatches;
            DispatchReport.Initialize();
            DispatchReport.CreateReport(NeedProfilList, NeedTPSList, false, ColorFullName, true, ref PackagesReportName);
            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
        }

        private void AllPrintContextMenuItem_Click(object sender, EventArgs e)
        {
            bool NeedProfilList = true;
            bool NeedTPSList = true;
            int ClientID = Convert.ToInt32(dgvDispatch.SelectedRows[0].Cells["ClientID"].Value);
            int DispatchID = Convert.ToInt32(dgvDispatch.SelectedRows[0].Cells["DispatchID"].Value);
            DateTime DispatchDate = Convert.ToDateTime(dgvDispatchDates.SelectedRows[0].Cells["PrepareDispatchDateTime"].Value);

            int[] Dispatches = new int[dgvDispatch.SelectedRows.Count];
            for (int i = 0; i < dgvDispatch.SelectedRows.Count; i++)
                Dispatches[i] = Convert.ToInt32(dgvDispatch.SelectedRows[i].Cells["DispatchID"].Value);
            if (!MarketingDispatchManager.HasPackages(DispatchID))
            {
                InfiniumTips.ShowTip(this, 50, 85, "Отгрузка пуста", 1700);
                return;
            }

            bool PressOK = false;
            bool ColorFullName = false;
            object MachineName = DBNull.Value;
            object PermitNumber = DBNull.Value;
            object SealNumber = DBNull.Value;

            PhantomForm PhantomForm = new Infinium.PhantomForm();
            PhantomForm.Show();

            MarketingDispatchInfoMenu MarketingDispatchInfoMenu = new MarketingDispatchInfoMenu(this);
            TopForm = MarketingDispatchInfoMenu;
            MarketingDispatchInfoMenu.ShowDialog();

            PressOK = MarketingDispatchInfoMenu.PressOK;
            ColorFullName = MarketingDispatchInfoMenu.ColorFullName;
            MachineName = MarketingDispatchInfoMenu.MachineName;
            PermitNumber = MarketingDispatchInfoMenu.PermitNumber;
            SealNumber = MarketingDispatchInfoMenu.SealNumber;

            PhantomForm.Close();
            PhantomForm.Dispose();
            MarketingDispatchInfoMenu.Dispose();
            TopForm = null;

            Thread T1 = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Создание документа Excel.\r\nПодождите..."); });
            T1.Start();

            while (!SplashWindow.bSmallCreated) ;

            object CreationDateTime = dgvDispatch.SelectedRows[0].Cells["CreationDateTime"].Value;
            object ConfirmExpDateTime = dgvDispatch.SelectedRows[0].Cells["ConfirmExpDateTime"].Value;
            object ConfirmDispDateTime = dgvDispatch.SelectedRows[0].Cells["ConfirmDispDateTime"].Value;
            object PrepareDispDateTime = dgvDispatchDates.SelectedRows[0].Cells["PrepareDispatchDateTime"].Value;
            object ConfirmExpUserID = dgvDispatch.SelectedRows[0].Cells["ConfirmExpUserID"].Value;
            object ConfirmDispUserID = dgvDispatch.SelectedRows[0].Cells["ConfirmDispUserID"].Value;
            object RealDispDateTime = DBNull.Value;
            object DispUserID = DBNull.Value;
            string PackagesReportName = string.Empty;
            MarketingDispatchManager.GetRealDispDateTime(DispatchID, ref RealDispDateTime, ref DispUserID);

            DispatchReport.GetDispatchInfo(ref CreationDateTime, ref ConfirmExpDateTime, ref ConfirmDispDateTime, ref RealDispDateTime, ref PrepareDispDateTime,
                ref ConfirmExpUserID, ref ConfirmDispUserID, ref DispUserID, ref MachineName, ref PermitNumber, ref SealNumber);
            DispatchReport.CurrentClient = ClientID;
            DispatchReport.CurrentDispatches = Dispatches;
            DispatchReport.Initialize();
            DispatchReport.CreateReport(NeedProfilList, NeedTPSList, false, ColorFullName, true, ref PackagesReportName);

            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
        }

        private void dgvDispatch_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if ((RoleType != RoleTypes.OrdinaryRole) && e.Button == System.Windows.Forms.MouseButtons.Right)
            {
                //if (dgvDispatch.Rows[e.RowIndex].Cells["ConfirmDispDateTime"].Value == DBNull.Value)
                //{
                //    if (dgvDispatch.Rows[e.RowIndex].Cells["ConfirmExpDateTime"].Value == DBNull.Value)
                //    {
                //        btnConfirmExpedition.Text = "Разрешить эксп-цию";
                //        ConfirmExpContextMenuItem.Text = "Разрешить эксп-цию";
                //    }
                //    else
                //    {
                //        btnConfirmExpedition.Text = "Запретить эксп-цию";
                //        ConfirmExpContextMenuItem.Text = "Запретить эксп-цию";
                //    }
                //    btnConfirmDispatch.Text = "Разрешить отгрузку";
                //    ConfirmDispatchContextMenuItem.Text = "Разрешить отгрузку";
                //}
                //else
                //{
                //    btnConfirmDispatch.Text = "Запретить отгрузку";
                //    ConfirmDispatchContextMenuItem.Text = "Запретить отгрузку";
                //}
                dgvDispatch.Rows[e.RowIndex].Selected = true;
                kryptonContextMenu2.Show(new Point(Cursor.Position.X - 212, Cursor.Position.Y - 10));
            }
        }

        private void ProfilAttach_Click(object sender, EventArgs e)
        {
            bool NeedProfilList = true;
            bool NeedTPSList = false;
            int ClientID = Convert.ToInt32(dgvDispatch.SelectedRows[0].Cells["ClientID"].Value);
            int DispatchID = Convert.ToInt32(dgvDispatch.SelectedRows[0].Cells["DispatchID"].Value);
            DateTime DispatchDate = Convert.ToDateTime(dgvDispatchDates.SelectedRows[0].Cells["PrepareDispatchDateTime"].Value);

            int[] Dispatches = new int[dgvDispatch.SelectedRows.Count];
            for (int i = 0; i < dgvDispatch.SelectedRows.Count; i++)
                Dispatches[i] = Convert.ToInt32(dgvDispatch.SelectedRows[i].Cells["DispatchID"].Value);
            if (!MarketingDispatchManager.HasPackages(DispatchID))
            {
                InfiniumTips.ShowTip(this, 50, 85, "Отгрузка пуста", 1700);
                return;
            }

            bool PressOK = false;
            bool ColorFullName = false;
            object MachineName = DBNull.Value;
            object PermitNumber = DBNull.Value;
            object SealNumber = DBNull.Value;

            PhantomForm PhantomForm = new Infinium.PhantomForm();
            PhantomForm.Show();

            MarketingDispatchInfoMenu MarketingDispatchInfoMenu = new MarketingDispatchInfoMenu(this);
            TopForm = MarketingDispatchInfoMenu;
            MarketingDispatchInfoMenu.ShowDialog();

            PressOK = MarketingDispatchInfoMenu.PressOK;
            ColorFullName = MarketingDispatchInfoMenu.ColorFullName;
            MachineName = MarketingDispatchInfoMenu.MachineName;
            PermitNumber = MarketingDispatchInfoMenu.PermitNumber;
            SealNumber = MarketingDispatchInfoMenu.SealNumber;

            PhantomForm.Close();
            PhantomForm.Dispose();
            MarketingDispatchInfoMenu.Dispose();
            TopForm = null;

            Thread T1 = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Создание документа Excel.\r\nПодождите..."); });
            T1.Start();

            while (!SplashWindow.bSmallCreated) ;

            object CreationDateTime = dgvDispatch.SelectedRows[0].Cells["CreationDateTime"].Value;
            object ConfirmExpDateTime = dgvDispatch.SelectedRows[0].Cells["ConfirmExpDateTime"].Value;
            object ConfirmDispDateTime = dgvDispatch.SelectedRows[0].Cells["ConfirmDispDateTime"].Value;
            object PrepareDispDateTime = dgvDispatchDates.SelectedRows[0].Cells["PrepareDispatchDateTime"].Value;
            object ConfirmExpUserID = dgvDispatch.SelectedRows[0].Cells["ConfirmExpUserID"].Value;
            object ConfirmDispUserID = dgvDispatch.SelectedRows[0].Cells["ConfirmDispUserID"].Value;
            object RealDispDateTime = DBNull.Value;
            object DispUserID = DBNull.Value;
            string PackagesReportName = string.Empty;
            MarketingDispatchManager.GetRealDispDateTime(DispatchID, ref RealDispDateTime, ref DispUserID);

            DispatchReport.GetDispatchInfo(ref CreationDateTime, ref ConfirmExpDateTime, ref ConfirmDispDateTime, ref RealDispDateTime, ref PrepareDispDateTime,
                ref ConfirmExpUserID, ref ConfirmDispUserID, ref DispUserID, ref MachineName, ref PermitNumber, ref SealNumber);
            DispatchReport.CurrentClient = ClientID;
            DispatchReport.CurrentDispatches = Dispatches;
            DispatchReport.Initialize();
            DispatchReport.CreateReport(NeedProfilList, NeedTPSList, true, ColorFullName, true, ref PackagesReportName);

            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
        }

        private void TPSAttach_Click(object sender, EventArgs e)
        {
            bool NeedProfilList = false;
            bool NeedTPSList = true;
            int ClientID = Convert.ToInt32(dgvDispatch.SelectedRows[0].Cells["ClientID"].Value);
            int DispatchID = Convert.ToInt32(dgvDispatch.SelectedRows[0].Cells["DispatchID"].Value);
            DateTime DispatchDate = Convert.ToDateTime(dgvDispatchDates.SelectedRows[0].Cells["PrepareDispatchDateTime"].Value);

            int[] Dispatches = new int[dgvDispatch.SelectedRows.Count];
            for (int i = 0; i < dgvDispatch.SelectedRows.Count; i++)
                Dispatches[i] = Convert.ToInt32(dgvDispatch.SelectedRows[i].Cells["DispatchID"].Value);
            if (!MarketingDispatchManager.HasPackages(DispatchID))
            {
                InfiniumTips.ShowTip(this, 50, 85, "Отгрузка пуста", 1700);
                return;
            }

            bool PressOK = false;
            bool ColorFullName = false;
            object MachineName = DBNull.Value;
            object PermitNumber = DBNull.Value;
            object SealNumber = DBNull.Value;

            PhantomForm PhantomForm = new Infinium.PhantomForm();
            PhantomForm.Show();

            MarketingDispatchInfoMenu MarketingDispatchInfoMenu = new MarketingDispatchInfoMenu(this);
            TopForm = MarketingDispatchInfoMenu;
            MarketingDispatchInfoMenu.ShowDialog();

            PressOK = MarketingDispatchInfoMenu.PressOK;
            ColorFullName = MarketingDispatchInfoMenu.ColorFullName;
            MachineName = MarketingDispatchInfoMenu.MachineName;
            PermitNumber = MarketingDispatchInfoMenu.PermitNumber;
            SealNumber = MarketingDispatchInfoMenu.SealNumber;

            PhantomForm.Close();
            PhantomForm.Dispose();
            MarketingDispatchInfoMenu.Dispose();
            TopForm = null;

            Thread T1 = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Создание документа Excel.\r\nПодождите..."); });
            T1.Start();

            while (!SplashWindow.bSmallCreated) ;

            object CreationDateTime = dgvDispatch.SelectedRows[0].Cells["CreationDateTime"].Value;
            object ConfirmExpDateTime = dgvDispatch.SelectedRows[0].Cells["ConfirmExpDateTime"].Value;
            object ConfirmDispDateTime = dgvDispatch.SelectedRows[0].Cells["ConfirmDispDateTime"].Value;
            object PrepareDispDateTime = dgvDispatchDates.SelectedRows[0].Cells["PrepareDispatchDateTime"].Value;
            object ConfirmExpUserID = dgvDispatch.SelectedRows[0].Cells["ConfirmExpUserID"].Value;
            object ConfirmDispUserID = dgvDispatch.SelectedRows[0].Cells["ConfirmDispUserID"].Value;
            object RealDispDateTime = DBNull.Value;
            object DispUserID = DBNull.Value;
            string PackagesReportName = string.Empty;
            MarketingDispatchManager.GetRealDispDateTime(DispatchID, ref RealDispDateTime, ref DispUserID);

            DispatchReport.GetDispatchInfo(ref CreationDateTime, ref ConfirmExpDateTime, ref ConfirmDispDateTime, ref RealDispDateTime, ref PrepareDispDateTime,
                ref ConfirmExpUserID, ref ConfirmDispUserID, ref DispUserID, ref MachineName, ref PermitNumber, ref SealNumber);
            DispatchReport.CurrentClient = ClientID;
            DispatchReport.CurrentDispatches = Dispatches;
            DispatchReport.Initialize();
            DispatchReport.CreateReport(NeedProfilList, NeedTPSList, true, ColorFullName, true, ref PackagesReportName);

            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
        }

        private void AllAttach_Click(object sender, EventArgs e)
        {
            bool NeedProfilList = true;
            bool NeedTPSList = true;
            int ClientID = Convert.ToInt32(dgvDispatch.SelectedRows[0].Cells["ClientID"].Value);
            int DispatchID = Convert.ToInt32(dgvDispatch.SelectedRows[0].Cells["DispatchID"].Value);
            DateTime DispatchDate = Convert.ToDateTime(dgvDispatchDates.SelectedRows[0].Cells["PrepareDispatchDateTime"].Value);

            int[] Dispatches = new int[dgvDispatch.SelectedRows.Count];
            for (int i = 0; i < dgvDispatch.SelectedRows.Count; i++)
                Dispatches[i] = Convert.ToInt32(dgvDispatch.SelectedRows[i].Cells["DispatchID"].Value);
            if (!MarketingDispatchManager.HasPackages(DispatchID))
            {
                InfiniumTips.ShowTip(this, 50, 85, "Отгрузка пуста", 1700);
                return;
            }

            bool PressOK = false;
            bool ColorFullName = false;
            object MachineName = DBNull.Value;
            object PermitNumber = DBNull.Value;
            object SealNumber = DBNull.Value;

            PhantomForm PhantomForm = new Infinium.PhantomForm();
            PhantomForm.Show();

            MarketingDispatchInfoMenu MarketingDispatchInfoMenu = new MarketingDispatchInfoMenu(this);
            TopForm = MarketingDispatchInfoMenu;
            MarketingDispatchInfoMenu.ShowDialog();

            PressOK = MarketingDispatchInfoMenu.PressOK;
            ColorFullName = MarketingDispatchInfoMenu.ColorFullName;
            MachineName = MarketingDispatchInfoMenu.MachineName;
            PermitNumber = MarketingDispatchInfoMenu.PermitNumber;
            SealNumber = MarketingDispatchInfoMenu.SealNumber;

            PhantomForm.Close();
            PhantomForm.Dispose();
            MarketingDispatchInfoMenu.Dispose();
            TopForm = null;

            Thread T1 = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Создание документа Excel.\r\nПодождите..."); });
            T1.Start();

            while (!SplashWindow.bSmallCreated) ;

            object CreationDateTime = dgvDispatch.SelectedRows[0].Cells["CreationDateTime"].Value;
            object ConfirmExpDateTime = dgvDispatch.SelectedRows[0].Cells["ConfirmExpDateTime"].Value;
            object ConfirmDispDateTime = dgvDispatch.SelectedRows[0].Cells["ConfirmDispDateTime"].Value;
            object PrepareDispDateTime = dgvDispatchDates.SelectedRows[0].Cells["PrepareDispatchDateTime"].Value;
            object ConfirmExpUserID = dgvDispatch.SelectedRows[0].Cells["ConfirmExpUserID"].Value;
            object ConfirmDispUserID = dgvDispatch.SelectedRows[0].Cells["ConfirmDispUserID"].Value;
            object RealDispDateTime = DBNull.Value;
            object DispUserID = DBNull.Value;
            string PackagesReportName = string.Empty;
            MarketingDispatchManager.GetRealDispDateTime(DispatchID, ref RealDispDateTime, ref DispUserID);

            DispatchReport.GetDispatchInfo(ref CreationDateTime, ref ConfirmExpDateTime, ref ConfirmDispDateTime, ref RealDispDateTime, ref PrepareDispDateTime,
                ref ConfirmExpUserID, ref ConfirmDispUserID, ref DispUserID, ref MachineName, ref PermitNumber, ref SealNumber);
            DispatchReport.CurrentClient = ClientID;
            DispatchReport.CurrentDispatches = Dispatches;
            DispatchReport.Initialize();
            DispatchReport.CreateReport(NeedProfilList, NeedTPSList, true, ColorFullName, true, ref PackagesReportName);

            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
        }

        private void dgvDispatch_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dgvDispatch.Rows.Count == 0)
                return;

            if ((RoleType != RoleTypes.AdminRole && RoleType != RoleTypes.LogisticsRole)
                || dgvDispatch.SelectedRows[0].Cells["DispatchStatus"].Value.ToString() == "Отгружена")
            {
                CanEditDispatch = false;
            }

            Thread T = new Thread(delegate () { SplashWindow.CreateSplash(); });
            T.Start();

            while (!SplashForm.bCreated) ;
            int ClientID = Convert.ToInt32(dgvDispatch.SelectedRows[0].Cells["ClientID"].Value);
            int DispatchID = Convert.ToInt32(dgvDispatch.SelectedRows[0].Cells["DispatchID"].Value);
            DateTime DispatchDate = Convert.ToDateTime(dgvDispatchDates.SelectedRows[0].Cells["PrepareDispatchDateTime"].Value);

            MarketingNewDispatchForm MarketingNewDispatchForm = new MarketingNewDispatchForm(CanEditDispatch, false, DispatchDate, ClientID, DispatchID);

            TopForm = MarketingNewDispatchForm;

            MarketingNewDispatchForm.ShowDialog();

            MarketingNewDispatchForm.Close();
            MarketingNewDispatchForm.Dispose();

            Thread T1 = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
            T1.Start();

            while (!SplashWindow.bSmallCreated) ;

            NeedSplash = false;
            UpdateDispatchDate();
            MarketingDispatchManager.MoveToDispatchDate(DispatchDate);
            MarketingDispatchManager.MoveToDispatch(DispatchID);
            NeedSplash = true;
            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;

            TopForm = null;
        }

        private void ConfirmDispatchContextMenuItem_Click(object sender, EventArgs e)
        {
            int DispatchID = Convert.ToInt32(dgvDispatch.SelectedRows[0].Cells["DispatchID"].Value);
            DateTime DispatchDate = Convert.ToDateTime(dgvDispatchDates.SelectedRows[0].Cells["PrepareDispatchDateTime"].Value);
            for (int i = 0; i < dgvDispatch.SelectedRows.Count; i++)
            {
                int id = Convert.ToInt32(dgvDispatch.SelectedRows[i].Cells["DispatchID"].Value);

                if (!MarketingDispatchManager.HasPackages(id))
                {
                    LightMessageBox.Show(ref TopForm, false,
                            "Отгрузка №" + id.ToString() + " пуста",
                            "Разрешить отгрузку");
                    return;
                }

                bool Confirm = false;
                if (dgvDispatch.SelectedRows[i].Cells["ConfirmDispDateTime"].Value == DBNull.Value)
                    Confirm = true;
                if (Confirm && dgvDispatch.SelectedRows[i].Cells["ConfirmExpDateTime"].Value == DBNull.Value)
                {
                    LightMessageBox.Show(ref TopForm, false,
                            "Отгрузка №" + id.ToString() + " не утверждена к экспедиции",
                            "Разрешить отгрузку");
                    return;
                }
                MarketingDispatchManager.SaveConfirmDispInfo(id, Confirm);
            }

            Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Обновление данных.\r\nПодождите..."); });
            T.Start();

            while (!SplashWindow.bSmallCreated) ;

            NeedSplash = false;

            UpdateDispatchDate();
            MarketingDispatchManager.MoveToDispatchDate(DispatchDate);
            MarketingDispatchManager.MoveToDispatch(DispatchID);
            NeedSplash = true;

            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
        }

        private void ConfirmExpContextMenuItem_Click(object sender, EventArgs e)
        {
            int DispatchID = Convert.ToInt32(dgvDispatch.SelectedRows[0].Cells["DispatchID"].Value);
            DateTime DispatchDate = Convert.ToDateTime(dgvDispatchDates.SelectedRows[0].Cells["PrepareDispatchDateTime"].Value);

            for (int i = 0; i < dgvDispatch.SelectedRows.Count; i++)
            {
                int id = Convert.ToInt32(dgvDispatch.SelectedRows[i].Cells["DispatchID"].Value);
                if (!MarketingDispatchManager.HasPackages(id))
                {
                    LightMessageBox.Show(ref TopForm, false,
                            "Отгрузка №" + id.ToString() + " пуста",
                            "Разрешить экспедицию");
                    return;
                }

                bool Confirm = false;
                if (dgvDispatch.SelectedRows[i].Cells["ConfirmExpDateTime"].Value == DBNull.Value)
                    Confirm = true;

                if (Confirm && !MarketingDispatchManager.IsDispatchCanExp(id))
                {
                    Infinium.LightMessageBox.Show(ref TopForm, false,
                            "Запрещено: не вся продукция принята на склад.",
                            "Разрешить экспедицию");
                    return;
                }
                MarketingDispatchManager.SaveConfirmExpInfo(id, Confirm);
            }

            Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
            T.Start();

            while (!SplashWindow.bSmallCreated) ;

            NeedSplash = false;

            UpdateDispatchDate();
            MarketingDispatchManager.MoveToDispatchDate(DispatchDate);
            MarketingDispatchManager.MoveToDispatch(DispatchID);
            NeedSplash = true;

            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
        }

        private void ChangeCabFurDate_Click(object sender, EventArgs e)
        {
            if (dgvCabFurMegaOrders.Rows.Count == 0)
                return;

            int MegaOrderID = Convert.ToInt32(dgvCabFurMegaOrders.SelectedRows[0].Cells["MegaOrderID"].Value);
            int ClientID = Convert.ToInt32(dgvCabFurMegaOrders.SelectedRows[0].Cells["ClientID"].Value);
            object date = null;

            PhantomForm PhantomForm = new Infinium.PhantomForm();
            PhantomForm.Show();

            MarketingNewCabFurMenu marketingNewCabFurMenu = new MarketingNewCabFurMenu(this);
            TopForm = marketingNewCabFurMenu;
            marketingNewCabFurMenu.ShowDialog();

            bool bOk = marketingNewCabFurMenu.DialogResult == DialogResult.OK ? true : false;
            date = marketingNewCabFurMenu.DispatchDate;

            PhantomForm.Close();
            PhantomForm.Dispose();
            marketingNewCabFurMenu.Dispose();
            TopForm = null;

            if (bOk && date != null)
            {
                Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                T.Start();

                while (!SplashWindow.bSmallCreated) ;

                NeedSplash = false;
                cabFurAssembleManager.ChangeAssembleDate(MegaOrderID, ClientID, date);
                UpdateCabFurDispatchDate();
                cabFurAssembleManager.MoveToAssembleDate(Convert.ToDateTime(date));
                cabFurAssembleManager.MoveToMegaOrder(MegaOrderID);
                NeedSplash = true;

                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;
            }
        }

        private void cmiChangeDispatchDate_Click(object sender, EventArgs e)
        {
            if (dgvDispatch.Rows.Count == 0)
                return;

            int ClientID = Convert.ToInt32(dgvDispatch.SelectedRows[0].Cells["ClientID"].Value);
            int DispatchID = Convert.ToInt32(dgvDispatch.SelectedRows[0].Cells["DispatchID"].Value);
            object DispatchDate = null;

            bool PressOK = false;

            PhantomForm PhantomForm = new Infinium.PhantomForm();
            PhantomForm.Show();

            MarketingNewDispatchMenu MarketingNewDispatchMenu = new MarketingNewDispatchMenu(this, MarketingExpeditionManager.ClientsDT, true);
            TopForm = MarketingNewDispatchMenu;
            MarketingNewDispatchMenu.ShowDialog();

            PressOK = MarketingNewDispatchMenu.PressOK;
            DispatchDate = MarketingNewDispatchMenu.DispatchDate;

            PhantomForm.Close();
            PhantomForm.Dispose();
            MarketingNewDispatchMenu.Dispose();
            TopForm = null;

            if (PressOK && DispatchDate != null)
            {
                Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                T.Start();

                while (!SplashWindow.bSmallCreated) ;

                NeedSplash = false;
                MarketingDispatchManager.ChangeDispatchDate(DispatchID, DispatchDate);
                MarketingDispatchManager.SetDispatchDate(DispatchID, Convert.ToDateTime(DispatchDate));
                UpdateDispatchDate();
                MarketingDispatchManager.MoveToDispatchDate(Convert.ToDateTime(DispatchDate));
                MarketingDispatchManager.MoveToDispatch(DispatchID);
                NeedSplash = true;

                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;
            }
        }

        private void cmiBindToPermit_Click(object sender, EventArgs e)
        {
            if (dgvDispatch.SelectedRows.Count == 0)
                return;

            Thread T = new Thread(delegate () { SplashWindow.CreateSplash(); });
            T.Start();

            while (!SplashForm.bCreated) ;

            bool BindingOk = false;
            int[] Dispatchs = new int[dgvDispatch.SelectedRows.Count];
            for (int i = 0; i < dgvDispatch.SelectedRows.Count; i++)
            {
                Dispatchs[i] = Convert.ToInt32(dgvDispatch.SelectedRows[i].Cells["DispatchID"].Value);
            }

            PermitsForm PermitsForm = new PermitsForm(this, Dispatchs);

            TopForm = PermitsForm;

            PermitsForm.ShowDialog();
            BindingOk = PermitsForm.BindingOk;
            PermitsForm.Close();
            PermitsForm.Dispose();

            TopForm = null;
            if (BindingOk)
                InfiniumTips.ShowTip(this, 50, 85, "Пропуск привязан", 1700);
        }

        private void kryptonContextMenuItem1_Click(object sender, EventArgs e)
        {
            if (dgvDispatch.SelectedRows.Count == 0)
                return;

            int[] Dispatchs = new int[dgvDispatch.SelectedRows.Count];
            for (int i = 0; i < dgvDispatch.SelectedRows.Count; i++)
            {
                Dispatchs[i] = Convert.ToInt32(dgvDispatch.SelectedRows[i].Cells["DispatchID"].Value);
            }

            if (Dispatchs.Count() == 0)
            {
                Infinium.LightMessageBox.Show(ref TopForm, false,
                    "Не выбрана отгрузка.",
                    "Ошибка");
                return;
            }

            if (!MarketingDispatchManager.IsDispatchBindToPermit(Convert.ToInt32(dgvDispatch.SelectedRows[0].Cells["DispatchID"].Value)))
            {
                Infinium.LightMessageBox.Show(ref TopForm, false,
                    "Необходимо привязать отгрузку к пропуску.",
                    "Привязка к пропуску");
                Thread T = new Thread(delegate () { SplashWindow.CreateSplash(); });
                T.Start();

                while (!SplashForm.bCreated) ;

                bool BindingOk = false;

                PermitsForm PermitsForm = new PermitsForm(this, Dispatchs);

                TopForm = PermitsForm;

                PermitsForm.ShowDialog();
                BindingOk = PermitsForm.BindingOk;
                PermitsForm.Close();
                PermitsForm.Dispose();

                TopForm = null;
                //if (BindingOk)
                //    InfiniumTips.ShowTip(this, 50, 85, "Пропуск привязан", 1700);
            }

            ArrayList DispatchIDs = new ArrayList();
            string s = string.Empty;
            for (int i = 0; i < dgvDispatch.SelectedRows.Count; i++)
            {
                if (dgvDispatch.SelectedRows[i].Cells["ConfirmDispDateTime"].Value == DBNull.Value)
                {
                    s += Convert.ToInt32(dgvDispatch.SelectedRows[i].Cells["DispatchID"].Value) + ",";
                    continue;
                }
                DispatchIDs.Add(Convert.ToInt32(dgvDispatch.SelectedRows[i].Cells["DispatchID"].Value));
            }
            if (DispatchIDs.Count == 0)
            {
                Infinium.LightMessageBox.Show(ref TopForm, false,
                    "Ни одна отгрузка не может быть отгружена.",
                    "Ошибка");
                return;
            }
            if (s.Length > 0)
            {
                s = "Следующие отгрузки не могут быть отгружены, так как не утверждены: " + s.Substring(0, s.Length - 1);
                Infinium.LightMessageBox.Show(ref TopForm, false, s, "Ошибка");
            }
            Thread T1 = new Thread(delegate () { SplashWindow.CreateSplash(); });
            T1.Start();

            while (!SplashForm.bCreated) ;

            MarketingDispatchLabelCheckForm DispatchLabelCheckForm = new MarketingDispatchLabelCheckForm(this, DispatchIDs);

            TopForm = DispatchLabelCheckForm;

            DispatchLabelCheckForm.ShowDialog();

            DispatchLabelCheckForm.Close();
            DispatchLabelCheckForm.Dispose();

            TopForm = null;
        }

        private void dgvDispatchDates_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if ((RoleType != RoleTypes.OrdinaryRole) && e.Button == System.Windows.Forms.MouseButtons.Right)
            {
                kryptonContextMenu3.Show(new Point(Cursor.Position.X - 212, Cursor.Position.Y - 10));
            }
        }

        private void kryptonContextMenuItem1_Click_1(object sender, EventArgs e)
        {
            Thread T1 = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Создание документа Excel.\r\nПодождите..."); });
            T1.Start();

            while (!SplashWindow.bSmallCreated) ;

            object PrepareDispDateTime = dgvDispatchDates.SelectedRows[0].Cells["PrepareDispatchDateTime"].Value;
            System.Collections.ArrayList Dispatches = new System.Collections.ArrayList();
            for (int i = 0; i < dgvDispatch.Rows.Count; i++)
                Dispatches.Add(Convert.ToInt32(dgvDispatch.Rows[i].Cells["DispatchID"].Value));
            Dispatches.Sort();


            NotReadyProducts.GetDispatchInfo(ref PrepareDispDateTime);
            NotReadyProducts.CurrentDispatches = Dispatches;
            NotReadyProducts.FileName = "Не готово к отгрузке " + Convert.ToDateTime(PrepareDispDateTime).ToString("dd.MM.yyyy");
            NotReadyProducts.Initialize1();
            NotReadyProducts.CreateReport();

            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
        }

        private void kryptonContextMenuItem2_Click(object sender, EventArgs e)
        {
            if (dgvDispatch.SelectedRows.Count == 0)
                return;

            Thread T1 = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Создание документа Excel.\r\nПодождите..."); });
            T1.Start();

            while (!SplashWindow.bSmallCreated) ;

            object PrepareDispDateTime = dgvDispatchDates.SelectedRows[0].Cells["PrepareDispatchDateTime"].Value;
            System.Collections.ArrayList Dispatches = new System.Collections.ArrayList();
            for (int i = 0; i < dgvDispatch.SelectedRows.Count; i++)
                Dispatches.Add(Convert.ToInt32(dgvDispatch.SelectedRows[i].Cells["DispatchID"].Value));

            NotReadyProducts.GetDispatchInfo(ref PrepareDispDateTime);
            NotReadyProducts.CurrentDispatches = Dispatches;
            NotReadyProducts.FileName = "Не готово к отгрузке " + Convert.ToDateTime(PrepareDispDateTime).ToString("dd.MM.yyyy");
            NotReadyProducts.Initialize1();
            NotReadyProducts.CreateReport();

            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
        }

        private void kryptonContextMenuItem4_Click(object sender, EventArgs e)
        {
            Thread T1 = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Создание документа Excel.\r\nПодождите..."); });
            T1.Start();

            while (!SplashWindow.bSmallCreated) ;

            object PrepareDispDateTime = dgvDispatchDates.SelectedRows[0].Cells["PrepareDispatchDateTime"].Value;
            System.Collections.ArrayList Dispatches = new System.Collections.ArrayList();
            for (int i = 0; i < dgvDispatch.SelectedRows.Count; i++)
                Dispatches.Add(Convert.ToInt32(dgvDispatch.SelectedRows[i].Cells["DispatchID"].Value));
            Dispatches.Sort();

            NotReadyProducts.CurrentDispatches = Dispatches;
            NotReadyProducts.FileName = "Номера заказов " + Convert.ToDateTime(PrepareDispDateTime).ToString("dd.MM.yyyy");
            NotReadyProducts.Initialize2();
            NotReadyProducts.CreateOrdersNumbersReport();

            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
        }

        private void kryptonContextMenuItem3_Click(object sender, EventArgs e)
        {
            Thread T1 = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Создание документа Excel.\r\nПодождите..."); });
            T1.Start();

            while (!SplashWindow.bSmallCreated) ;

            object PrepareDispDateTime = dgvDispatchDates.SelectedRows[0].Cells["PrepareDispatchDateTime"].Value;
            System.Collections.ArrayList Dispatches = new System.Collections.ArrayList();
            for (int i = 0; i < dgvDispatch.Rows.Count; i++)
                Dispatches.Add(Convert.ToInt32(dgvDispatch.Rows[i].Cells["DispatchID"].Value));
            Dispatches.Sort();

            NotReadyProducts.CurrentDispatches = Dispatches;
            NotReadyProducts.FileName = "Номера заказов " + Convert.ToDateTime(PrepareDispDateTime).ToString("dd.MM.yyyy");
            NotReadyProducts.Initialize2();
            NotReadyProducts.CreateOrdersNumbersReport();

            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
        }

        private void kryptonContextMenuItem5_Click(object sender, EventArgs e)
        {
            bool InMutualSettlement = Convert.ToBoolean(dgvDispatch.SelectedRows[0].Cells["InMutualSettlement"].Value);
            //int DispatchID = Convert.ToInt32(dgvDispatch.SelectedRows[0].Cells["DispatchID"].Value);
            DateTime DispatchDate = Convert.ToDateTime(dgvDispatchDates.SelectedRows[0].Cells["PrepareDispatchDateTime"].Value);
            int ClientID = Convert.ToInt32(dgvDispatch.SelectedRows[0].Cells["ClientID"].Value);
            string ClientName = dgvDispatch.SelectedRows[0].Cells["ClientName"].Value.ToString();
            int[] Dispatches = new int[dgvDispatch.SelectedRows.Count];
            for (int i = 0; i < dgvDispatch.SelectedRows.Count; i++)
                Dispatches[i] = Convert.ToInt32(dgvDispatch.SelectedRows[i].Cells["DispatchID"].Value);

            int ClientCountryID = OrdersManager.GetClientCountry(ClientID);
            int iDiscountPaymentConditionID = 0;
            string ClientLogin = OrdersManager.GetClientLogin(ClientID);

            decimal DiscountPaymentCondition = 0;
            decimal TotalProfil = 0;
            decimal TotalTPS = 0;

            string ProfilDBFName = string.Empty;
            string TPSDBFName = string.Empty;
            string ReportName = string.Empty;
            InMutualSettlement = false;
            _dispatchReportToDbf.ClearTempPackages();
            if (!InMutualSettlement)
            {
                Modules.Marketing.MutualSettlements.MutualSettlements MutualSettlementsManager = new Modules.Marketing.MutualSettlements.MutualSettlements();
                MutualSettlementsManager.Initialize();

                if (ClientID == 145)
                {
                    bool HasOrdinaryOrders = MarketingDispatchManager.HasOrders(Dispatches, false);
                    bool HasSampleOrders = MarketingDispatchManager.HasOrders(Dispatches, true);

                    if (HasSampleOrders)
                    {
                        PhantomForm PhantomForm = new Infinium.PhantomForm();
                        PhantomForm.Show();

                        int Result = 1;
                        string SaveFilePath = string.Empty;
                        SaveDBFReportMenu = new Infinium.SaveDBFReportMenu(this, ReportName);

                        TopForm = SaveDBFReportMenu;
                        SaveDBFReportMenu.ShowDialog();

                        PhantomForm.Close();
                        Result = SaveDBFReportMenu.Result;
                        SaveFilePath = SaveDBFReportMenu.SaveFilePath;
                        PhantomForm.Dispose();
                        SaveDBFReportMenu.Dispose();
                        TopForm = null;

                        if (Result == 3)
                            return;

                        int ClientDispatchID = -1;
                        int MutualSettlementID1 = -1;
                        int MutualSettlementID2 = -1;
                        bool ProfilVerify = true;
                        bool TPSVerify = true;

                        ClientName = ClientName.Replace("\"", " ");
                        ClientName = ClientName.Replace("*", " ");
                        ClientName = ClientName.Replace("|", " ");
                        ClientName = ClientName.Replace(@"\", " ");
                        ClientName = ClientName.Replace(":", " ");
                        ClientName = ClientName.Replace("<", " ");
                        ClientName = ClientName.Replace(">", " ");
                        ClientName = ClientName.Replace("?", " ");
                        ClientName = ClientName.Replace("/", " ");

                        string CurrentMonthName = DateTime.Now.ToString("dd-MM-yyyy");
                        string FilePath = Security.DBFSaveFilePath;
                        if (Security.DBFSaveFilePath.Length == 0)
                            FilePath = Environment.GetFolderPath(Environment.SpecialFolder.Personal) + "/Infinium/1C data/" + CurrentMonthName;

                        Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Создание документа Excel.\r\nПодождите..."); });
                        T.Start();

                        while (!SplashWindow.bSmallCreated) ;
                        NeedSplash = false;

                        for (int j = 0; j < Dispatches.Count(); j++)
                        {
                            int[] MainOrders = MarketingDispatchManager.GetMainOrdersInDispatch(Dispatches[j]);

                            DataTable ReturnDT = MarketingDispatchManager.GetMegaOrdersInDispatch(Dispatches[j]);
                            ArrayList OrderNumbers = new ArrayList();
                            for (int i = 0; i < ReturnDT.Rows.Count; i++)
                                OrderNumbers.Add(Convert.ToInt32(ReturnDT.Rows[i]["OrderNumber"]));

                            int CurrencyTypeID = 1;
                            if (ReturnDT.Rows.Count > 0)
                                CurrencyTypeID = Convert.ToInt32(ReturnDT.Rows[0]["CurrencyTypeID"]);

                            MutualSettlementID1 = MutualSettlementsManager.FindMutualSettlementIDByOrderNumber(ClientID, 1, OrderNumbers.OfType<Int32>().ToArray(), true);
                            if (MutualSettlementID1 != 0)
                            {
                                decimal DispatchSum = MutualSettlementsManager.GetDispatchSum(MutualSettlementID1);
                                ProfilVerify = MutualSettlementsManager.VerifyTransaction(MutualSettlementID1, DispatchSum, ref DiscountPaymentCondition);
                                for (int i = 0; i < ReturnDT.Rows.Count; i++)
                                {
                                    int MegaOrderID = Convert.ToInt32(ReturnDT.Rows[i]["MegaOrderID"]);

                                    MegaOrderInfo m = MarketingDispatchManager.GetMegaOrders(MegaOrderID);
                                    iDiscountPaymentConditionID = ProfilReCalcMegaOrder(ClientID, CurrencyTypeID, DiscountPaymentCondition, ProfilVerify, true, MegaOrderID, m);
                                    OrdersManager.FixOrderEvent(MegaOrderID, "Выписать накладную. Пересчитан заказ Профиль");
                                }
                            }
                            MutualSettlementID2 = 0;
                            MutualSettlementID2 = MutualSettlementsManager.FindMutualSettlementIDByOrderNumber(ClientID, 2, OrderNumbers.OfType<Int32>().ToArray(), true);
                            if (MutualSettlementID2 != 0)
                            {
                                decimal DispatchSum = MutualSettlementsManager.GetDispatchSum(MutualSettlementID2);
                                TPSVerify = MutualSettlementsManager.VerifyTransaction(MutualSettlementID2, DispatchSum, ref DiscountPaymentCondition);
                                for (int i = 0; i < ReturnDT.Rows.Count; i++)
                                {
                                    int MegaOrderID = Convert.ToInt32(ReturnDT.Rows[i]["MegaOrderID"]);
                                    MegaOrderInfo m = MarketingDispatchManager.GetMegaOrders(MegaOrderID);
                                    iDiscountPaymentConditionID = TPSReCalcMegaOrder(ClientID, CurrencyTypeID, DiscountPaymentCondition, TPSVerify, true, MegaOrderID, m);
                                    OrdersManager.FixOrderEvent(MegaOrderID, "Выписать накладную. Пересчитан заказ ТПС");
                                }
                            }

                            HSSFWorkbook hssfworkbook = new HSSFWorkbook();

                            ////create a entry of DocumentSummaryInformation
                            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
                            dsi.Company = "NPOI Team";
                            hssfworkbook.DocumentSummaryInformation = dsi;

                            ////create a entry of SummaryInformation
                            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
                            si.Subject = "NPOI SDK Example";
                            hssfworkbook.SummaryInformation = si;

                            ReportName = ClientLogin + " №" + Dispatches[j].ToString();
                            _dispatchReportToDbf.CreateReport(ref hssfworkbook, new int[1] { Dispatches[j] }, MainOrders, ClientID, ClientName, ref TotalProfil, ref TotalTPS, ProfilVerify, TPSVerify, iDiscountPaymentConditionID, true);
                            _dispatchReportToDbf.SaveDBF(FilePath, ReportName, true, ref ProfilDBFName, ref TPSDBFName);
                            DetailsReport.CreateReport(ref hssfworkbook, new int[1] { Dispatches[j] }, OrderNumbers.OfType<Int32>().ToArray(), MainOrders, ClientID, ClientName, ProfilVerify, TPSVerify, iDiscountPaymentConditionID, true);

                            if (ClientID == 609)
                            {
                                DetailsReport.ReportModuleOnline(ref hssfworkbook, Dispatches);
                            }
                            FileInfo file = new FileInfo(FilePath + @"\" + ReportName + ".xls");
                            int x = 1;
                            while (file.Exists == true)
                            {
                                file = new FileInfo(FilePath + @"\" + ReportName + "(" + x++ + ").xls");
                            }

                            FileStream NewFile = new FileStream(file.FullName, FileMode.Create);
                            hssfworkbook.Write(NewFile);
                            NewFile.Close();

                            ReportName = file.FullName;

                            if (Result == 1)
                            {
                                if (TotalProfil > 0)
                                {
                                    if (ClientCountryID == 1)
                                        TotalProfil *= 1.2m;
                                    ClientDispatchID = MutualSettlementsManager.AddClientDispatch(MutualSettlementID1, CurrencyTypeID, TotalProfil);
                                    var fileInfo = new System.IO.FileInfo(ReportName);
                                    MutualSettlementsManager.AttachDispatchDocument(Path.GetFileNameWithoutExtension(ReportName),
                                        fileInfo.Extension, ReportName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.DispatchExcel, ClientDispatchID);
                                    fileInfo = new System.IO.FileInfo(ProfilDBFName);
                                    MutualSettlementsManager.AttachDispatchDocument(Path.GetFileNameWithoutExtension(ProfilDBFName),
                                        fileInfo.Extension, ProfilDBFName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.DispatchDbf, ClientDispatchID);
                                }
                                if (TotalTPS > 0)
                                {
                                    if (ClientCountryID == 1)
                                        TotalTPS *= 1.2m;
                                    ClientDispatchID = MutualSettlementsManager.AddClientDispatch(MutualSettlementID2, CurrencyTypeID, TotalTPS);
                                    var fileInfo = new System.IO.FileInfo(ReportName);
                                    MutualSettlementsManager.AttachDispatchDocument(Path.GetFileNameWithoutExtension(ReportName),
                                        fileInfo.Extension, ReportName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.DispatchExcel, ClientDispatchID);
                                    fileInfo = new System.IO.FileInfo(TPSDBFName);
                                    MutualSettlementsManager.AttachDispatchDocument(Path.GetFileNameWithoutExtension(TPSDBFName),
                                        fileInfo.Extension, TPSDBFName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.DispatchDbf, ClientDispatchID);
                                }
                            }
                            MarketingDispatchManager.IncludeInMutualSettlement(Dispatches[j], MutualSettlementID1, MutualSettlementID2);
                        }

                        MutualSettlementsManager.SaveMutualSettlements();
                        MutualSettlementsManager.SaveMutualSettlementOrders();

                        UpdateDispatchDate();
                        MarketingDispatchManager.MoveToDispatchDate(DispatchDate);
                        MarketingDispatchManager.MoveToDispatch(Dispatches[0]);

                        DispatchInvoiceClient(ClientID, ClientName, Dispatches, iDiscountPaymentConditionID, ClientLogin, ProfilVerify, TPSVerify, FilePath);

                        NeedSplash = true;
                        while (SplashWindow.bSmallCreated)
                            SmallWaitForm.CloseS = true;
                    }
                    if (HasOrdinaryOrders)
                    {
                        PhantomForm PhantomForm = new Infinium.PhantomForm();
                        PhantomForm.Show();

                        int Result = 1;
                        string SaveFilePath = string.Empty;
                        SaveDBFReportMenu = new Infinium.SaveDBFReportMenu(this, ReportName);

                        TopForm = SaveDBFReportMenu;
                        SaveDBFReportMenu.ShowDialog();

                        PhantomForm.Close();
                        Result = SaveDBFReportMenu.Result;
                        SaveFilePath = SaveDBFReportMenu.SaveFilePath;
                        PhantomForm.Dispose();
                        SaveDBFReportMenu.Dispose();
                        TopForm = null;

                        if (Result == 3)
                            return;

                        int ClientDispatchID = -1;
                        int MutualSettlementID1 = -1;
                        int MutualSettlementID2 = -1;
                        bool ProfilVerify = true;
                        bool TPSVerify = true;

                        ClientName = ClientName.Replace("\"", " ");
                        ClientName = ClientName.Replace("*", " ");
                        ClientName = ClientName.Replace("|", " ");
                        ClientName = ClientName.Replace(@"\", " ");
                        ClientName = ClientName.Replace(":", " ");
                        ClientName = ClientName.Replace("<", " ");
                        ClientName = ClientName.Replace(">", " ");
                        ClientName = ClientName.Replace("?", " ");
                        ClientName = ClientName.Replace("/", " ");

                        string CurrentMonthName = DateTime.Now.ToString("dd-MM-yyyy");
                        string FilePath = Security.DBFSaveFilePath;
                        if (Security.DBFSaveFilePath.Length == 0)
                            FilePath = Environment.GetFolderPath(Environment.SpecialFolder.Personal) + "/Infinium/1C data/" + CurrentMonthName;

                        Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Создание документа Excel.\r\nПодождите..."); });
                        T.Start();

                        while (!SplashWindow.bSmallCreated) ;
                        NeedSplash = false;

                        for (int j = 0; j < Dispatches.Count(); j++)
                        {
                            int[] MainOrders = MarketingDispatchManager.GetMainOrdersInDispatch(Dispatches[j]);

                            DataTable ReturnDT = MarketingDispatchManager.GetMegaOrdersInDispatch(Dispatches[j]);
                            ArrayList OrderNumbers = new ArrayList();
                            for (int i = 0; i < ReturnDT.Rows.Count; i++)
                                OrderNumbers.Add(Convert.ToInt32(ReturnDT.Rows[i]["OrderNumber"]));

                            int CurrencyTypeID = 1;
                            if (ReturnDT.Rows.Count > 0)
                                CurrencyTypeID = Convert.ToInt32(ReturnDT.Rows[0]["CurrencyTypeID"]);

                            MutualSettlementID1 = MutualSettlementsManager.FindMutualSettlementIDByOrderNumber(ClientID, 1, OrderNumbers.OfType<Int32>().ToArray(), false);
                            if (MutualSettlementID1 != 0)
                            {
                                decimal DispatchSum = MutualSettlementsManager.GetDispatchSum(MutualSettlementID1);
                                ProfilVerify = MutualSettlementsManager.VerifyTransaction(MutualSettlementID1, DispatchSum, ref DiscountPaymentCondition);
                                for (int i = 0; i < ReturnDT.Rows.Count; i++)
                                {
                                    int MegaOrderID = Convert.ToInt32(ReturnDT.Rows[i]["MegaOrderID"]);

                                    MegaOrderInfo m = MarketingDispatchManager.GetMegaOrders(MegaOrderID);
                                    iDiscountPaymentConditionID = ProfilReCalcMegaOrder(ClientID, CurrencyTypeID, DiscountPaymentCondition, ProfilVerify, false, MegaOrderID, m);
                                    OrdersManager.FixOrderEvent(MegaOrderID, "Выписать накладную. Пересчитан заказ Профиль");
                                }
                            }
                            MutualSettlementID2 = 0;
                            MutualSettlementID2 = MutualSettlementsManager.FindMutualSettlementIDByOrderNumber(ClientID, 2, OrderNumbers.OfType<Int32>().ToArray(), false);
                            if (MutualSettlementID2 != 0)
                            {
                                decimal DispatchSum = MutualSettlementsManager.GetDispatchSum(MutualSettlementID2);
                                TPSVerify = MutualSettlementsManager.VerifyTransaction(MutualSettlementID2, DispatchSum, ref DiscountPaymentCondition);
                                for (int i = 0; i < ReturnDT.Rows.Count; i++)
                                {
                                    int MegaOrderID = Convert.ToInt32(ReturnDT.Rows[i]["MegaOrderID"]);
                                    MegaOrderInfo m = MarketingDispatchManager.GetMegaOrders(MegaOrderID);
                                    iDiscountPaymentConditionID = TPSReCalcMegaOrder(ClientID, CurrencyTypeID, DiscountPaymentCondition, TPSVerify, false, MegaOrderID, m);
                                    OrdersManager.FixOrderEvent(MegaOrderID, "Выписать накладную. Пересчитан заказ ТПС");
                                }
                            }

                            HSSFWorkbook hssfworkbook = new HSSFWorkbook();

                            ////create a entry of DocumentSummaryInformation
                            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
                            dsi.Company = "NPOI Team";
                            hssfworkbook.DocumentSummaryInformation = dsi;

                            ////create a entry of SummaryInformation
                            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
                            si.Subject = "NPOI SDK Example";
                            hssfworkbook.SummaryInformation = si;

                            ReportName = ClientLogin + " №" + Dispatches[j].ToString();
                            _dispatchReportToDbf.CreateReport(ref hssfworkbook, new int[1] { Dispatches[j] }, MainOrders, ClientID, ClientName, ref TotalProfil, ref TotalTPS, ProfilVerify, TPSVerify, iDiscountPaymentConditionID, false);
                            _dispatchReportToDbf.SaveDBF(FilePath, ReportName, false, ref ProfilDBFName, ref TPSDBFName);
                            DetailsReport.CreateReport(ref hssfworkbook, new int[1] { Dispatches[j] }, OrderNumbers.OfType<Int32>().ToArray(), MainOrders, ClientID, ClientName, ProfilVerify, TPSVerify, iDiscountPaymentConditionID, false);

                            if (ClientID == 609)
                            {
                                DetailsReport.ReportModuleOnline(ref hssfworkbook, Dispatches);
                            }
                            FileInfo file = new FileInfo(FilePath + @"\" + ReportName + ".xls");
                            int x = 1;
                            while (file.Exists == true)
                            {
                                file = new FileInfo(FilePath + @"\" + ReportName + "(" + x++ + ").xls");
                            }

                            FileStream NewFile = new FileStream(file.FullName, FileMode.Create);
                            hssfworkbook.Write(NewFile);
                            NewFile.Close();

                            ReportName = file.FullName;

                            if (Result == 1)
                            {
                                if (TotalProfil > 0)
                                {
                                    if (ClientCountryID == 1)
                                        TotalProfil *= 1.2m;
                                    ClientDispatchID = MutualSettlementsManager.AddClientDispatch(MutualSettlementID1, CurrencyTypeID, TotalProfil);
                                    var fileInfo = new System.IO.FileInfo(ReportName);
                                    MutualSettlementsManager.AttachDispatchDocument(Path.GetFileNameWithoutExtension(ReportName),
                                        fileInfo.Extension, ReportName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.DispatchExcel, ClientDispatchID);
                                    fileInfo = new System.IO.FileInfo(ProfilDBFName);
                                    MutualSettlementsManager.AttachDispatchDocument(Path.GetFileNameWithoutExtension(ProfilDBFName),
                                        fileInfo.Extension, ProfilDBFName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.DispatchDbf, ClientDispatchID);
                                }
                                if (TotalTPS > 0)
                                {
                                    if (ClientCountryID == 1)
                                        TotalTPS *= 1.2m;
                                    ClientDispatchID = MutualSettlementsManager.AddClientDispatch(MutualSettlementID2, CurrencyTypeID, TotalTPS);
                                    var fileInfo = new System.IO.FileInfo(ReportName);
                                    MutualSettlementsManager.AttachDispatchDocument(Path.GetFileNameWithoutExtension(ReportName),
                                        fileInfo.Extension, ReportName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.DispatchExcel, ClientDispatchID);
                                    fileInfo = new System.IO.FileInfo(TPSDBFName);
                                    MutualSettlementsManager.AttachDispatchDocument(Path.GetFileNameWithoutExtension(TPSDBFName),
                                        fileInfo.Extension, TPSDBFName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.DispatchDbf, ClientDispatchID);
                                }
                            }
                            MarketingDispatchManager.IncludeInMutualSettlement(Dispatches[j], MutualSettlementID1, MutualSettlementID2);
                        }

                        MutualSettlementsManager.SaveMutualSettlements();
                        MutualSettlementsManager.SaveMutualSettlementOrders();

                        UpdateDispatchDate();
                        MarketingDispatchManager.MoveToDispatchDate(DispatchDate);
                        MarketingDispatchManager.MoveToDispatch(Dispatches[0]);

                        DispatchInvoiceClient(ClientID, ClientName, Dispatches, iDiscountPaymentConditionID, ClientLogin, ProfilVerify, TPSVerify, FilePath);

                        NeedSplash = true;
                        while (SplashWindow.bSmallCreated)
                            SmallWaitForm.CloseS = true;
                    }
                }
                else
                {
                    PhantomForm PhantomForm = new Infinium.PhantomForm();
                    PhantomForm.Show();

                    int Result = 1;
                    string SaveFilePath = string.Empty;
                    SaveDBFReportMenu = new Infinium.SaveDBFReportMenu(this, ReportName);

                    TopForm = SaveDBFReportMenu;
                    SaveDBFReportMenu.ShowDialog();

                    PhantomForm.Close();
                    Result = SaveDBFReportMenu.Result;
                    SaveFilePath = SaveDBFReportMenu.SaveFilePath;
                    PhantomForm.Dispose();
                    SaveDBFReportMenu.Dispose();
                    TopForm = null;

                    if (Result == 3)
                        return;

                    int ClientDispatchID = -1;
                    int MutualSettlementID1 = -1;
                    int MutualSettlementID2 = -1;
                    bool ProfilVerify = true;
                    bool TPSVerify = true;

                    ClientName = ClientName.Replace("\"", " ");
                    ClientName = ClientName.Replace("*", " ");
                    ClientName = ClientName.Replace("|", " ");
                    ClientName = ClientName.Replace(@"\", " ");
                    ClientName = ClientName.Replace(":", " ");
                    ClientName = ClientName.Replace("<", " ");
                    ClientName = ClientName.Replace(">", " ");
                    ClientName = ClientName.Replace("?", " ");
                    ClientName = ClientName.Replace("/", " ");

                    string CurrentMonthName = DateTime.Now.ToString("dd-MM-yyyy");
                    string FilePath = Security.DBFSaveFilePath;
                    if (Security.DBFSaveFilePath.Length == 0)
                        FilePath = Environment.GetFolderPath(Environment.SpecialFolder.Personal) + "/Infinium/1C data/" + CurrentMonthName;

                    Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Создание документа Excel.\r\nПодождите..."); });
                    T.Start();

                    while (!SplashWindow.bSmallCreated) ;
                    NeedSplash = false;

                    for (int j = 0; j < Dispatches.Count(); j++)
                    {
                        int[] MainOrders = MarketingDispatchManager.GetMainOrdersInDispatch(Dispatches[j]);

                        DataTable ReturnDT = MarketingDispatchManager.GetMegaOrdersInDispatch(Dispatches[j]);
                        ArrayList OrderNumbers = new ArrayList();
                        for (int i = 0; i < ReturnDT.Rows.Count; i++)
                            OrderNumbers.Add(Convert.ToInt32(ReturnDT.Rows[i]["OrderNumber"]));

                        int CurrencyTypeID = 1;
                        int FactoryID = 1;

                        if (ReturnDT.Rows.Count > 0)
                        {
                            CurrencyTypeID = Convert.ToInt32(ReturnDT.Rows[0]["CurrencyTypeID"]);
                            FactoryID = Convert.ToInt32(ReturnDT.Rows[0]["FactoryID"]);
                        }

                        MutualSettlementID1 = MutualSettlementsManager.FindMutualSettlementIDByOrderNumber(ClientID, 1, OrderNumbers.OfType<Int32>().ToArray());
                        if (MutualSettlementID1 != 0)
                        {
                            decimal DispatchSum = MutualSettlementsManager.GetDispatchSum(MutualSettlementID1);
                            ProfilVerify = MutualSettlementsManager.VerifyTransaction(MutualSettlementID1, DispatchSum, ref DiscountPaymentCondition);
                            for (int i = 0; i < ReturnDT.Rows.Count; i++)
                            {
                                int MegaOrderID = Convert.ToInt32(ReturnDT.Rows[i]["MegaOrderID"]);

                                MegaOrderInfo m = MarketingDispatchManager.GetMegaOrders(MegaOrderID);
                                iDiscountPaymentConditionID = ProfilReCalcMegaOrder(ClientID, CurrencyTypeID, DiscountPaymentCondition, ProfilVerify, MegaOrderID, m);
                                OrdersManager.FixOrderEvent(MegaOrderID, "Выписать накладную. Пересчитан заказ Профиль");
                            }
                        }
                        MutualSettlementID2 = 0;
                        MutualSettlementID2 = MutualSettlementsManager.FindMutualSettlementIDByOrderNumber(ClientID, 2, OrderNumbers.OfType<Int32>().ToArray());
                        if (MutualSettlementID2 != 0)
                        {
                            decimal DispatchSum = MutualSettlementsManager.GetDispatchSum(MutualSettlementID2);
                            TPSVerify = MutualSettlementsManager.VerifyTransaction(MutualSettlementID2, DispatchSum, ref DiscountPaymentCondition);
                            for (int i = 0; i < ReturnDT.Rows.Count; i++)
                            {
                                int MegaOrderID = Convert.ToInt32(ReturnDT.Rows[i]["MegaOrderID"]);
                                MegaOrderInfo m = MarketingDispatchManager.GetMegaOrders(MegaOrderID);
                                iDiscountPaymentConditionID = TPSReCalcMegaOrder(ClientID, CurrencyTypeID, DiscountPaymentCondition, TPSVerify, MegaOrderID, m);
                                OrdersManager.FixOrderEvent(MegaOrderID, "Выписать накладную. Пересчитан заказ ТПС");
                            }
                        }

                        HSSFWorkbook hssfworkbook = new HSSFWorkbook();

                        ////create a entry of DocumentSummaryInformation
                        DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
                        dsi.Company = "NPOI Team";
                        hssfworkbook.DocumentSummaryInformation = dsi;

                        ////create a entry of SummaryInformation
                        SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
                        si.Subject = "NPOI SDK Example";
                        hssfworkbook.SummaryInformation = si;

                        ReportName = ClientLogin + " №" + Dispatches[j].ToString();

                        _dispatchReportToDbf.CreateReport(ref hssfworkbook, new int[1] { Dispatches[j] }, MainOrders, ClientID, ClientName, ref TotalProfil, ref TotalTPS, ProfilVerify, TPSVerify, iDiscountPaymentConditionID);
                        _dispatchReportToDbf.SaveDBF(FilePath, ReportName, ref ProfilDBFName, ref TPSDBFName);

                        DetailsReport.CreateReport(ref hssfworkbook, new int[1] { Dispatches[j] }, OrderNumbers.OfType<Int32>().ToArray(), MainOrders, ClientID, ClientName, ProfilVerify, TPSVerify, iDiscountPaymentConditionID);

                        if (ClientID == 609)
                        {
                            DetailsReport.ReportModuleOnline(ref hssfworkbook, Dispatches);
                        }
                        FileInfo file = new FileInfo(FilePath + @"\" + ReportName + ".xls");
                        int x = 1;
                        while (file.Exists == true)
                        {
                            file = new FileInfo(FilePath + @"\" + ReportName + "(" + x++ + ").xls");
                        }

                        FileStream NewFile = new FileStream(file.FullName, FileMode.Create);
                        hssfworkbook.Write(NewFile);
                        NewFile.Close();

                        ReportName = file.FullName;

                        if (Result == 1)
                        {
                            if (TotalProfil > 0)
                            {
                                if (ClientCountryID == 1)
                                    TotalProfil *= 1.2m;
                                ClientDispatchID = MutualSettlementsManager.AddClientDispatch(MutualSettlementID1, CurrencyTypeID, TotalProfil);
                                var fileInfo = new System.IO.FileInfo(ReportName);
                                MutualSettlementsManager.AttachDispatchDocument(Path.GetFileNameWithoutExtension(ReportName),
                                    fileInfo.Extension, ReportName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.DispatchExcel, ClientDispatchID);
                                fileInfo = new System.IO.FileInfo(ProfilDBFName);
                                MutualSettlementsManager.AttachDispatchDocument(Path.GetFileNameWithoutExtension(ProfilDBFName),
                                    fileInfo.Extension, ProfilDBFName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.DispatchDbf, ClientDispatchID);
                            }
                            if (TotalTPS > 0)
                            {
                                if (ClientCountryID == 1)
                                    TotalTPS *= 1.2m;
                                ClientDispatchID = MutualSettlementsManager.AddClientDispatch(MutualSettlementID2, CurrencyTypeID, TotalTPS);
                                var fileInfo = new System.IO.FileInfo(ReportName);
                                MutualSettlementsManager.AttachDispatchDocument(Path.GetFileNameWithoutExtension(ReportName),
                                    fileInfo.Extension, ReportName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.DispatchExcel, ClientDispatchID);
                                fileInfo = new System.IO.FileInfo(TPSDBFName);
                                MutualSettlementsManager.AttachDispatchDocument(Path.GetFileNameWithoutExtension(TPSDBFName),
                                    fileInfo.Extension, TPSDBFName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.DispatchDbf, ClientDispatchID);
                            }
                        }
                        MarketingDispatchManager.IncludeInMutualSettlement(Dispatches[j], MutualSettlementID1, MutualSettlementID2);
                    }

                    MutualSettlementsManager.SaveMutualSettlements();
                    MutualSettlementsManager.SaveMutualSettlementOrders();

                    UpdateDispatchDate();
                    MarketingDispatchManager.MoveToDispatchDate(DispatchDate);
                    MarketingDispatchManager.MoveToDispatch(Dispatches[0]);

                    DispatchInvoiceClient(ClientID, ClientName, Dispatches, iDiscountPaymentConditionID, ClientLogin, ProfilVerify, TPSVerify, FilePath);

                    NeedSplash = true;
                    while (SplashWindow.bSmallCreated)
                        SmallWaitForm.CloseS = true;
                }
            }
        }

        private void ColorInvoiceClient(int ClientID, string ClientName, int[] Dispatches, int iDiscountPaymentConditionID, string ClientLogin,
            bool ProfilVerify, bool TPSVerify, string FilePath)
        {
            decimal TotalProfil = 0;
            decimal TotalTPS = 0;

            HSSFWorkbook hssfworkbook = new HSSFWorkbook();

            ////create a entry of DocumentSummaryInformation
            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = "NPOI Team";
            hssfworkbook.DocumentSummaryInformation = dsi;

            ////create a entry of SummaryInformation
            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
            si.Subject = "NPOI SDK Example";
            hssfworkbook.SummaryInformation = si;

            DataTable ReturnDT = MarketingDispatchManager.GetMegaOrders(Dispatches);
            ArrayList OrderNumbers = new ArrayList();
            for (int i = 0; i < ReturnDT.Rows.Count; i++)
                OrderNumbers.Add(Convert.ToInt32(ReturnDT.Rows[i]["OrderNumber"]));
            int[] MainOrders = MarketingDispatchManager.GetMainOrders(Dispatches);

            string ReportName = ClientLogin + " №" + string.Join(",", Dispatches);
            _colorInvoiceReportToDbf.CreateReport(ref hssfworkbook, Dispatches, MainOrders, ClientID, ClientName, ref TotalProfil, ref TotalTPS, ProfilVerify, TPSVerify, iDiscountPaymentConditionID);
            DetailsReport.CreateReport(ref hssfworkbook, Dispatches, OrderNumbers.OfType<Int32>().ToArray(), MainOrders, ClientID, ClientName, ProfilVerify, TPSVerify, iDiscountPaymentConditionID);

            if (ClientID == 609)
            {
                DetailsReport.ReportModuleOnline(ref hssfworkbook, Dispatches);
            }
            object CreationDateTime = dgvDispatch.SelectedRows[0].Cells["CreationDateTime"].Value;
            object ConfirmExpDateTime = dgvDispatch.SelectedRows[0].Cells["ConfirmExpDateTime"].Value;
            object ConfirmDispDateTime = dgvDispatch.SelectedRows[0].Cells["ConfirmDispDateTime"].Value;
            object PrepareDispDateTime = dgvDispatchDates.SelectedRows[0].Cells["PrepareDispatchDateTime"].Value;
            object ConfirmExpUserID = dgvDispatch.SelectedRows[0].Cells["ConfirmExpUserID"].Value;
            object ConfirmDispUserID = dgvDispatch.SelectedRows[0].Cells["ConfirmDispUserID"].Value;
            object RealDispDateTime = DBNull.Value;
            object DispUserID = DBNull.Value;
            MarketingDispatchManager.GetRealDispDateTime(Dispatches[0], ref RealDispDateTime, ref DispUserID);

            DispatchReport.GetDispatchInfo(ref CreationDateTime, ref ConfirmExpDateTime, ref ConfirmDispDateTime, ref RealDispDateTime, ref PrepareDispDateTime,
                ref ConfirmExpUserID, ref ConfirmDispUserID, ref DispUserID);
            DispatchReport.CurrentClient = ClientID;
            DispatchReport.CurrentDispatches = Dispatches;
            DispatchReport.Initialize();
            DispatchReport.CreateReport(ref hssfworkbook, true, true, true, false);

            FileInfo file = new FileInfo(FilePath + @"\" + ReportName + ".xls");
            int j = 1;
            while (file.Exists == true)
            {
                file = new FileInfo(FilePath + @"\" + ReportName + "(" + j++ + ").xls");
            }

            FileStream NewFile = new FileStream(file.FullName, FileMode.Create);
            hssfworkbook.Write(NewFile);
            NewFile.Close();

            ReportName = file.FullName;

            SendEmail SendEmail = new SendEmail();
            SendEmail.Send(ClientID, string.Join(",", OrderNumbers.OfType<Int32>().ToArray()), DetailsReport.Save, ReportName);
        }

        private void NotesInvoiceClient(int ClientID, string ClientName, int[] Dispatches, int iDiscountPaymentConditionID, string ClientLogin,
            bool ProfilVerify, bool TPSVerify, string FilePath)
        {
            decimal TotalProfil = 0;
            decimal TotalTPS = 0;

            HSSFWorkbook hssfworkbook = new HSSFWorkbook();

            ////create a entry of DocumentSummaryInformation
            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = "NPOI Team";
            hssfworkbook.DocumentSummaryInformation = dsi;

            ////create a entry of SummaryInformation
            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
            si.Subject = "NPOI SDK Example";
            hssfworkbook.SummaryInformation = si;

            DataTable ReturnDT = MarketingDispatchManager.GetMegaOrders(Dispatches);
            ArrayList OrderNumbers = new ArrayList();
            for (int i = 0; i < ReturnDT.Rows.Count; i++)
                OrderNumbers.Add(Convert.ToInt32(ReturnDT.Rows[i]["OrderNumber"]));
            int[] MainOrders = MarketingDispatchManager.GetMainOrders(Dispatches);

            string ReportName = ClientLogin + " №" + string.Join(",", Dispatches);
            _notesInvoiceReportToDbf.CreateReport(ref hssfworkbook, Dispatches, MainOrders, ClientID, ClientName, ref TotalProfil, ref TotalTPS, ProfilVerify, TPSVerify, iDiscountPaymentConditionID);
            DetailsReport.CreateReport(ref hssfworkbook, Dispatches, OrderNumbers.OfType<Int32>().ToArray(), MainOrders, ClientID, ClientName, ProfilVerify, TPSVerify, iDiscountPaymentConditionID);

            if (ClientID == 609)
            {
                DetailsReport.ReportModuleOnline(ref hssfworkbook, Dispatches);
            }
            object CreationDateTime = dgvDispatch.SelectedRows[0].Cells["CreationDateTime"].Value;
            object ConfirmExpDateTime = dgvDispatch.SelectedRows[0].Cells["ConfirmExpDateTime"].Value;
            object ConfirmDispDateTime = dgvDispatch.SelectedRows[0].Cells["ConfirmDispDateTime"].Value;
            object PrepareDispDateTime = dgvDispatchDates.SelectedRows[0].Cells["PrepareDispatchDateTime"].Value;
            object ConfirmExpUserID = dgvDispatch.SelectedRows[0].Cells["ConfirmExpUserID"].Value;
            object ConfirmDispUserID = dgvDispatch.SelectedRows[0].Cells["ConfirmDispUserID"].Value;
            object RealDispDateTime = DBNull.Value;
            object DispUserID = DBNull.Value;
            MarketingDispatchManager.GetRealDispDateTime(Dispatches[0], ref RealDispDateTime, ref DispUserID);

            DispatchReport.GetDispatchInfo(ref CreationDateTime, ref ConfirmExpDateTime, ref ConfirmDispDateTime, ref RealDispDateTime, ref PrepareDispDateTime,
                ref ConfirmExpUserID, ref ConfirmDispUserID, ref DispUserID);
            DispatchReport.CurrentClient = ClientID;
            DispatchReport.CurrentDispatches = Dispatches;
            DispatchReport.Initialize();
            DispatchReport.CreateReport(ref hssfworkbook, true, true, true, false);

            FileInfo file = new FileInfo(FilePath + @"\" + ReportName + ".xls");
            int j = 1;
            while (file.Exists == true)
            {
                file = new FileInfo(FilePath + @"\" + ReportName + "(" + j++ + ").xls");
            }

            FileStream NewFile = new FileStream(file.FullName, FileMode.Create);
            hssfworkbook.Write(NewFile);
            NewFile.Close();

            ReportName = file.FullName;

            SendEmail SendEmail = new SendEmail();
            SendEmail.Send(ClientID, string.Join(",", OrderNumbers.OfType<Int32>().ToArray()), DetailsReport.Save, ReportName);
        }
        
        private void DispatchInvoiceClient(int ClientID, string ClientName, int[] Dispatches, int iDiscountPaymentConditionID, string ClientLogin,
            bool ProfilVerify, bool TPSVerify, string FilePath)
        {
            decimal TotalProfil = 0;
            decimal TotalTPS = 0;

            HSSFWorkbook hssfworkbook = new HSSFWorkbook();

            ////create a entry of DocumentSummaryInformation
            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = "NPOI Team";
            hssfworkbook.DocumentSummaryInformation = dsi;

            ////create a entry of SummaryInformation
            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
            si.Subject = "NPOI SDK Example";
            hssfworkbook.SummaryInformation = si;

            DataTable ReturnDT = MarketingDispatchManager.GetMegaOrders(Dispatches);
            ArrayList OrderNumbers = new ArrayList();
            for (int i = 0; i < ReturnDT.Rows.Count; i++)
                OrderNumbers.Add(Convert.ToInt32(ReturnDT.Rows[i]["OrderNumber"]));
            int[] MainOrders = MarketingDispatchManager.GetMainOrders(Dispatches);

            string ReportName = ClientLogin + " №" + string.Join(",", Dispatches);
            _dispatchReportToDbf.CreateReport(ref hssfworkbook, Dispatches, MainOrders, ClientID, ClientName, ref TotalProfil, ref TotalTPS, ProfilVerify, TPSVerify, iDiscountPaymentConditionID);
            DetailsReport.CreateReport(ref hssfworkbook, Dispatches, OrderNumbers.OfType<Int32>().ToArray(), MainOrders, ClientID, ClientName, ProfilVerify, TPSVerify, iDiscountPaymentConditionID);

            if (ClientID == 609)
            {
                DetailsReport.ReportModuleOnline(ref hssfworkbook, Dispatches);
            }
            object CreationDateTime = dgvDispatch.SelectedRows[0].Cells["CreationDateTime"].Value;
            object ConfirmExpDateTime = dgvDispatch.SelectedRows[0].Cells["ConfirmExpDateTime"].Value;
            object ConfirmDispDateTime = dgvDispatch.SelectedRows[0].Cells["ConfirmDispDateTime"].Value;
            object PrepareDispDateTime = dgvDispatchDates.SelectedRows[0].Cells["PrepareDispatchDateTime"].Value;
            object ConfirmExpUserID = dgvDispatch.SelectedRows[0].Cells["ConfirmExpUserID"].Value;
            object ConfirmDispUserID = dgvDispatch.SelectedRows[0].Cells["ConfirmDispUserID"].Value;
            object RealDispDateTime = DBNull.Value;
            object DispUserID = DBNull.Value;
            MarketingDispatchManager.GetRealDispDateTime(Dispatches[0], ref RealDispDateTime, ref DispUserID);

            DispatchReport.GetDispatchInfo(ref CreationDateTime, ref ConfirmExpDateTime, ref ConfirmDispDateTime, ref RealDispDateTime, ref PrepareDispDateTime,
                ref ConfirmExpUserID, ref ConfirmDispUserID, ref DispUserID);
            DispatchReport.CurrentClient = ClientID;
            DispatchReport.CurrentDispatches = Dispatches;
            DispatchReport.Initialize();
            DispatchReport.CreateReport(ref hssfworkbook, true, true, true, false);

            FileInfo file = new FileInfo(FilePath + @"\" + ReportName + ".xls");
            int j = 1;
            while (file.Exists == true)
            {
                file = new FileInfo(FilePath + @"\" + ReportName + "(" + j++ + ").xls");
            }

            FileStream NewFile = new FileStream(file.FullName, FileMode.Create);
            hssfworkbook.Write(NewFile);
            NewFile.Close();

            ReportName = file.FullName;

            SendEmail SendEmail = new SendEmail();
            SendEmail.Send(ClientID, string.Join(",", OrderNumbers.OfType<Int32>().ToArray()), DetailsReport.Save, ReportName);
        }

        private int ProfilReCalcMegaOrder(int ClientID, int CurrencyTypeID, decimal DiscountPaymentCondition, bool Verify, int MegaOrderID, MegaOrderInfo m)
        {
            int iDiscountPaymentConditionID;
            decimal D = DiscountPaymentCondition;

            if (m.DiscountPaymentConditionID != 6)
            {
                if (m.DiscountPaymentConditionID == 1)
                    D = 0;
                if (DiscountPaymentCondition == 0)
                    m.DiscountPaymentConditionID = 1;
                if (ClientID != 145 && m.DiscountPaymentConditionID == 1 && (CurrencyTypeID == 3 || CurrencyTypeID == 5))
                {
                    //m.PaymentRate = m.OriginalRate * 1.05m;
                }

                if (Verify)
                {
                    m.ProfilTotalDiscount = D + m.ProfilDiscountOrderSum + m.ProfilDiscountDirector;
                }
                else
                {
                    m.ProfilDiscountOrderSum = 0;
                    m.ProfilTotalDiscount = m.ProfilDiscountDirector;
                }
            }
            iDiscountPaymentConditionID = m.DiscountPaymentConditionID;
            decimal CurrencyTotalCost = iDBFReport.CalcCurrencyCost(MegaOrderID, ClientID, m.PaymentRate);
            OrdersCalculate.Recalculate(MegaOrderID, m.ProfilDiscountDirector, m.TPSDiscountDirector,
                (m.ProfilTotalDiscount - m.ProfilDiscountDirector), (m.TPSTotalDiscount - m.TPSDiscountDirector), m.DiscountPaymentCondition,
                CurrencyTypeID, m.PaymentRate, m.ConfirmDateTime, CurrencyTotalCost, !Verify);
            OrdersCalculate.SetProfilCurrencyCost(MegaOrderID, m.ProfilTotalDiscount, m.TPSTotalDiscount,
                m.ProfilDiscountOrderSum, m.TPSDiscountOrderSum, m.PaymentRate);
            return iDiscountPaymentConditionID;
        }

        private int TPSReCalcMegaOrder(int ClientID, int CurrencyTypeID, decimal DiscountPaymentCondition, bool Verify, int MegaOrderID, MegaOrderInfo m)
        {
            int iDiscountPaymentConditionID;
            decimal D = DiscountPaymentCondition;

            if (m.DiscountPaymentConditionID != 6)
            {
                if (m.DiscountPaymentConditionID == 1)
                    D = 0;
                if (DiscountPaymentCondition == 0)
                    m.DiscountPaymentConditionID = 1;
                if (ClientID != 145 && m.DiscountPaymentConditionID == 1 && (CurrencyTypeID == 3 || CurrencyTypeID == 5))
                {
                    //m.PaymentRate = m.OriginalRate * 1.05m;
                }

                if (Verify)
                {
                    m.ProfilTotalDiscount = D + m.ProfilDiscountOrderSum + m.ProfilDiscountDirector;
                }
                else
                {
                    m.ProfilDiscountOrderSum = 0;
                    m.ProfilTotalDiscount = m.ProfilDiscountDirector;
                }
            }
            iDiscountPaymentConditionID = m.DiscountPaymentConditionID;
            decimal CurrencyTotalCost = iDBFReport.CalcCurrencyCost(MegaOrderID, ClientID, m.PaymentRate);
            OrdersCalculate.Recalculate(MegaOrderID, m.ProfilDiscountDirector, m.TPSDiscountDirector,
                (m.ProfilTotalDiscount - m.ProfilDiscountDirector), (m.TPSTotalDiscount - m.TPSDiscountDirector), m.DiscountPaymentCondition,
                CurrencyTypeID, m.PaymentRate, m.ConfirmDateTime, CurrencyTotalCost, !Verify);
            OrdersCalculate.SetTPSCurrencyCost(MegaOrderID, m.ProfilTotalDiscount, m.TPSTotalDiscount,
                m.ProfilDiscountOrderSum, m.TPSDiscountOrderSum, m.PaymentRate);
            return iDiscountPaymentConditionID;
        }

        private int ProfilReCalcMegaOrder(int ClientID, int CurrencyTypeID, decimal DiscountPaymentCondition, bool Verify, bool IsSample, int MegaOrderID, MegaOrderInfo m)
        {
            int iDiscountPaymentConditionID;
            decimal D = DiscountPaymentCondition;

            if (m.DiscountPaymentConditionID != 6)
            {
                if (m.DiscountPaymentConditionID == 1)
                    D = 0;
                if (DiscountPaymentCondition == 0)
                    m.DiscountPaymentConditionID = 1;
                if (ClientID != 145 && m.DiscountPaymentConditionID == 1 && (CurrencyTypeID == 3 || CurrencyTypeID == 5))
                {
                    //m.PaymentRate = m.OriginalRate * 1.05m;
                }

                if (Verify)
                {
                    m.ProfilTotalDiscount = D + m.ProfilDiscountOrderSum + m.ProfilDiscountDirector;
                }
                else
                {
                    m.ProfilDiscountOrderSum = 0;
                    m.ProfilTotalDiscount = m.ProfilDiscountDirector;
                }
            }
            iDiscountPaymentConditionID = m.DiscountPaymentConditionID;
            decimal CurrencyTotalCost = iDBFReport.CalcCurrencyCost(MegaOrderID, ClientID, m.PaymentRate, IsSample);
            OrdersCalculate.Recalculate(MegaOrderID, m.ProfilDiscountDirector, m.TPSDiscountDirector,
                (m.ProfilTotalDiscount - m.ProfilDiscountDirector), (m.TPSTotalDiscount - m.TPSDiscountDirector), m.DiscountPaymentCondition,
                CurrencyTypeID, m.PaymentRate, m.ConfirmDateTime, CurrencyTotalCost, !Verify);
            OrdersCalculate.SetProfilCurrencyCost(MegaOrderID, m.ProfilTotalDiscount, m.TPSTotalDiscount,
                m.ProfilDiscountOrderSum, m.TPSDiscountOrderSum, m.PaymentRate);
            return iDiscountPaymentConditionID;
        }

        private int TPSReCalcMegaOrder(int ClientID, int CurrencyTypeID, decimal DiscountPaymentCondition, bool Verify, bool IsSample, int MegaOrderID, MegaOrderInfo m)
        {
            int iDiscountPaymentConditionID;
            decimal D = DiscountPaymentCondition;

            if (m.DiscountPaymentConditionID != 6)
            {
                if (m.DiscountPaymentConditionID == 1)
                    D = 0;
                if (DiscountPaymentCondition == 0)
                    m.DiscountPaymentConditionID = 1;
                if (ClientID != 145 && m.DiscountPaymentConditionID == 1 && (CurrencyTypeID == 3 || CurrencyTypeID == 5))
                {
                    //m.PaymentRate = m.OriginalRate * 1.05m;
                }

                if (Verify)
                {
                    m.ProfilTotalDiscount = D + m.ProfilDiscountOrderSum + m.ProfilDiscountDirector;
                }
                else
                {
                    m.ProfilDiscountOrderSum = 0;
                    m.ProfilTotalDiscount = m.ProfilDiscountDirector;
                }
            }
            iDiscountPaymentConditionID = m.DiscountPaymentConditionID;
            decimal CurrencyTotalCost = iDBFReport.CalcCurrencyCost(MegaOrderID, ClientID, m.PaymentRate, IsSample);
            OrdersCalculate.Recalculate(MegaOrderID, m.ProfilDiscountDirector, m.TPSDiscountDirector,
                (m.ProfilTotalDiscount - m.ProfilDiscountDirector), (m.TPSTotalDiscount - m.TPSDiscountDirector), m.DiscountPaymentCondition,
                CurrencyTypeID, m.PaymentRate, m.ConfirmDateTime, CurrencyTotalCost, !Verify);
            OrdersCalculate.SetTPSCurrencyCost(MegaOrderID, m.ProfilTotalDiscount, m.TPSTotalDiscount,
                m.ProfilDiscountOrderSum, m.TPSDiscountOrderSum, m.PaymentRate);
            return iDiscountPaymentConditionID;
        }

        private void kryptonContextMenuItem6_Click(object sender, EventArgs e)
        {
            //Thread T1 = new Thread(delegate() { SplashWindow.CreateSmallSplash(ref TopForm, "Создание документа Excel.\r\nПодождите..."); });
            //T1.Start();

            //while (!SplashWindow.bSmallCreated) ;
            //NeedSplash = false;

            //bool InMutualSettlement = Convert.ToBoolean(dgvDispatch.SelectedRows[0].Cells["InMutualSettlement"].Value);
            //int DispatchID = Convert.ToInt32(dgvDispatch.SelectedRows[0].Cells["DispatchID"].Value);
            //DateTime DispatchDate = Convert.ToDateTime(dgvDispatchDates.SelectedRows[0].Cells["PrepareDispatchDateTime"].Value);

            //string ClientName = dgvDispatch.SelectedRows[0].Cells["ClientName"].Value.ToString();
            //int[] Dispatches = new int[dgvDispatch.SelectedRows.Count];
            //for (int i = 0; i < dgvDispatch.SelectedRows.Count; i++)
            //    Dispatches[i] = Convert.ToInt32(dgvDispatch.SelectedRows[i].Cells["DispatchID"].Value);
            //int[] MainOrders = MarketingDispatchManager.GetMainOrders(Dispatches);

            //decimal DiscountPaymentCondition = 0;
            //decimal TotalProfil = 0;
            //decimal TotalTPS = 0;

            //string ProfilDBFName = string.Empty;
            //string TPSDBFName = string.Empty;
            //string ReportName = string.Empty;
            //DataTable ReturnDT = MarketingDispatchManager.GetMegaOrders(Dispatches);
            //ArrayList OrderNumbers = new ArrayList();
            //for (int i = 0; i < ReturnDT.Rows.Count; i++)
            //    OrderNumbers.Add(Convert.ToInt32(ReturnDT.Rows[i]["OrderNumber"]));

            //int CurrencyTypeID = 1;
            //if (ReturnDT.Rows.Count > 0)
            //    CurrencyTypeID = Convert.ToInt32(ReturnDT.Rows[0]["CurrencyTypeID"]);

            //int ClientID = Convert.ToInt32(dgvDispatch.SelectedRows[0].Cells["ClientID"].Value);
            //int ClientCountryID = OrdersManager.GetClientCountry(ClientID);
            //int DelayOfPayment = OrdersManager.GetDelayOfPayment(ClientID);
            //int DiscountPaymentConditionID = 1;
            //int MutualSettlementID = -1;
            //string ClientLogin = OrdersManager.GetClientLogin(ClientID);
            //Infinium.Modules.Marketing.MutualSettlements.MutualSettlements MutualSettlementsManager = new Modules.Marketing.MutualSettlements.MutualSettlements();
            //MutualSettlementsManager.Initialize();

            //ClientName = ClientName.Replace("\"", " ");
            //ClientName = ClientName.Replace("*", " ");
            //ClientName = ClientName.Replace("|", " ");
            //ClientName = ClientName.Replace(@"\", " ");
            //ClientName = ClientName.Replace(":", " ");
            //ClientName = ClientName.Replace("<", " ");
            //ClientName = ClientName.Replace(">", " ");
            //ClientName = ClientName.Replace("?", " ");
            //ClientName = ClientName.Replace("/", " ");
            //ReportName = ClientLogin + " №" + string.Join(",", Dispatches);

            //string CurrentMonthName = DateTime.Now.ToString("dd-MM-yyyy");
            //string FilePath = Security.DBFSaveFilePath;
            //if (Security.DBFSaveFilePath.Length == 0)
            //    FilePath = Environment.GetFolderPath(Environment.SpecialFolder.Personal) + "/Infinium/1C data/" + CurrentMonthName;

            //HSSFWorkbook hssfworkbook = new HSSFWorkbook();

            //DBFReport.CreateReport(ref hssfworkbook, Dispatches, MainOrders, ClientID, ClientName, ref TotalProfil, ref TotalTPS, false, false, 1);
            //DBFReport.SaveDBF(FilePath, ReportName, ref ProfilDBFName, ref TPSDBFName);
            //DetailsReport.CreateReport(ref hssfworkbook, Dispatches, OrderNumbers.OfType<Int32>().ToArray(), MainOrders, ClientID, ClientName, false, false, 1);

            //FileInfo file = new FileInfo(FilePath + @"\" + ReportName + ".xls");
            //int j = 1;
            //while (file.Exists == true)
            //{
            //    file = new FileInfo(FilePath + @"\" + ReportName + "(" + j++ + ").xls");
            //}

            //FileStream NewFile = new FileStream(file.FullName, FileMode.Create);
            //hssfworkbook.Write(NewFile);
            //NewFile.Close();

            //ReportName = file.FullName;

            //NeedSplash = true;
            //while (SplashWindow.bSmallCreated)
            //    SmallWaitForm.CloseS = true;

            //PhantomForm PhantomForm = new Infinium.PhantomForm();
            //PhantomForm.Show();

            //int Result = 1;
            //string Notes = string.Empty;
            //string SaveFilePath = string.Empty;
            //SaveDBFReportMenu = new Infinium.SaveDBFReportMenu(this, ReportName);

            //TopForm = SaveDBFReportMenu;
            //SaveDBFReportMenu.ShowDialog();

            //PhantomForm.Close();
            //Result = SaveDBFReportMenu.Result;
            //Notes = SaveDBFReportMenu.Notes;
            //SaveFilePath = SaveDBFReportMenu.SaveFilePath;
            //PhantomForm.Dispose();
            //SaveDBFReportMenu.Dispose();
            //TopForm = null;

            //if (Result != 1)
            //{
            //    return;
            //}

            //T1 = new Thread(delegate() { SplashWindow.CreateSmallSplash(ref TopForm, "Создание документа Excel.\r\nПодождите..."); });
            //T1.Start();

            //while (!SplashWindow.bSmallCreated) ;
            //NeedSplash = false;

            //if (TotalProfil > 0)
            //{
            //    if (ClientCountryID == 1)
            //        TotalProfil *= 1.2m;
            //    MutualSettlementID = MutualSettlementsManager.AddMutualSettlement(ClientID, 1, CurrencyTypeID, DiscountPaymentConditionID, DelayOfPayment, TotalProfil, Notes);
            //    var fileInfo = new System.IO.FileInfo(ReportName);
            //    MutualSettlementsManager.AttachIncomeDocument(Path.GetFileNameWithoutExtension(ReportName),
            //        fileInfo.Extension, ReportName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.InvoiceExcel, MutualSettlementID);
            //    fileInfo = new System.IO.FileInfo(ProfilDBFName);
            //    MutualSettlementsManager.AttachIncomeDocument(Path.GetFileNameWithoutExtension(ProfilDBFName),
            //        fileInfo.Extension, ProfilDBFName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.InvoiceDbf, MutualSettlementID);
            //    DateTime CurrentDate = Security.GetCurrentDate();
            //    for (int i = 0; i < OrderNumbers.OfType<Int32>().ToArray().Count(); i++)
            //        MutualSettlementsManager.AddMutualSettlementOrder(CurrentDate, ClientID, MutualSettlementID, OrderNumbers.OfType<Int32>().ToArray()[i]);
            //}
            //if (TotalTPS > 0)
            //{
            //    if (ClientCountryID == 1)
            //        TotalTPS *= 1.2m;
            //    MutualSettlementID = MutualSettlementsManager.AddMutualSettlement(ClientID, 2, CurrencyTypeID, DiscountPaymentConditionID, DelayOfPayment, TotalTPS, Notes);
            //    var fileInfo = new System.IO.FileInfo(ReportName);
            //    MutualSettlementsManager.AttachIncomeDocument(Path.GetFileNameWithoutExtension(ReportName),
            //        fileInfo.Extension, ReportName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.InvoiceExcel, MutualSettlementID);
            //    fileInfo = new System.IO.FileInfo(TPSDBFName);
            //    MutualSettlementsManager.AttachIncomeDocument(Path.GetFileNameWithoutExtension(TPSDBFName),
            //        fileInfo.Extension, TPSDBFName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.InvoiceDbf, MutualSettlementID);
            //    DateTime CurrentDate = Security.GetCurrentDate();
            //    for (int i = 0; i < OrderNumbers.OfType<Int32>().ToArray().Count(); i++)
            //        MutualSettlementsManager.AddMutualSettlementOrder(CurrentDate, ClientID, MutualSettlementID, OrderNumbers.OfType<Int32>().ToArray()[i]);
            //}
            //MutualSettlementsManager.SaveMutualSettlements();
            //MutualSettlementsManager.SaveMutualSettlementOrders();


            //InMutualSettlement = false;
            //if (!InMutualSettlement)
            //{
            //    int ClientDispatchID = -1;
            //    int MutualSettlementID1 = -1;
            //    int MutualSettlementID2 = -1;

            //    MutualSettlementsManager.Initialize();

            //    MutualSettlementID1 = MutualSettlementsManager.FindMutualSettlementIDByOrderNumber(ClientID, 1, OrderNumbers.OfType<Int32>().ToArray());
            //    if (MutualSettlementID1 != 0)
            //    {
            //        decimal DispatchSum = MutualSettlementsManager.GetDispatchSum(MutualSettlementID1);
            //        bool bVerify = MutualSettlementsManager.VerifyTransaction(MutualSettlementID1, DispatchSum, ref DiscountPaymentCondition);
            //        Infinium.Modules.Marketing.Orders.OrdersCalculate OrdersCalculate = new Modules.Marketing.Orders.OrdersCalculate();
            //        for (int i = 0; i < ReturnDT.Rows.Count; i++)
            //        {
            //            int MegaOrderID = Convert.ToInt32(ReturnDT.Rows[i]["MegaOrderID"]);

            //            DataTable FFFDT = MarketingDispatchManager.GetMegaOrders(MegaOrderID);
            //            decimal D = DiscountPaymentCondition;
            //            //int DiscountPaymentConditionID = Convert.ToInt32(FFFDT.Rows[0]["DiscountPaymentConditionID"]);
            //            decimal ProfilDiscountOrderSum = Convert.ToDecimal(FFFDT.Rows[0]["ProfilDiscountOrderSum"]);
            //            decimal TPSDiscountOrderSum = Convert.ToDecimal(FFFDT.Rows[0]["TPSDiscountOrderSum"]);
            //            decimal ProfilDiscountDirector = Convert.ToDecimal(FFFDT.Rows[0]["ProfilDiscountDirector"]);
            //            decimal TPSDiscountDirector = Convert.ToDecimal(FFFDT.Rows[0]["TPSDiscountDirector"]);
            //            decimal ProfilTotalDiscount = Convert.ToDecimal(FFFDT.Rows[0]["ProfilTotalDiscount"]);
            //            decimal TPSTotalDiscount = Convert.ToDecimal(FFFDT.Rows[0]["TPSTotalDiscount"]);
            //            decimal OriginalRate = Convert.ToDecimal(FFFDT.Rows[0]["Rate"]);
            //            decimal PaymentRate = OriginalRate;
            //            object ConfirmDateTime = FFFDT.Rows[0]["ConfirmDateTime"];
            //            if (DiscountPaymentConditionID == 1)
            //                D = 0;
            //            if (DiscountPaymentCondition == 0)
            //                DiscountPaymentConditionID = 1;
            //            if (DiscountPaymentConditionID == 1 && (CurrencyTypeID == 3 || CurrencyTypeID == 5))
            //            {
            //                PaymentRate = OriginalRate * 1.05m;
            //            }

            //            if (bVerify)
            //            {
            //                ProfilTotalDiscount = D + ProfilDiscountOrderSum + ProfilDiscountDirector;
            //                //TPSTotalDiscount = DiscountPaymentCondition + TPSDiscountOrderSum + TPSDiscountDirector;
            //            }
            //            else
            //            {
            //                ProfilDiscountOrderSum = 0;
            //                //TPSDiscountOrderSum = 0;
            //                ProfilTotalDiscount = ProfilDiscountDirector;
            //                //TPSTotalDiscount = TPSDiscountDirector;
            //            }

            //            InvoiceReportToDBF iDBFReport = new InvoiceReportToDBF(ref DecorCatalogOrder, ref OrdersCalculate.FrontsCalculate);
            //            decimal CurrencyTotalCost = iDBFReport.CalcCurrencyCost(MegaOrderID, ClientID, PaymentRate);
            //            OrdersCalculate.Recalculate(MegaOrderID, ProfilDiscountDirector, TPSDiscountDirector,
            //                (ProfilTotalDiscount - ProfilDiscountDirector), (TPSTotalDiscount - TPSDiscountDirector), CurrencyTypeID, PaymentRate, ConfirmDateTime, CurrencyTotalCost, !bVerify);
            //            OrdersCalculate.SetProfilCurrencyCost(MegaOrderID, ProfilTotalDiscount, TPSTotalDiscount,
            //                ProfilDiscountOrderSum, TPSDiscountOrderSum, PaymentRate);
            //            OrdersManager.FixOrderEvent(MegaOrderID, "Выписать с отсрочкой. Пересчитан заказ Профиль");
            //        }
            //    }
            //    MutualSettlementID2 = 0;
            //    MutualSettlementID2 = MutualSettlementsManager.FindMutualSettlementIDByOrderNumber(ClientID, 2, OrderNumbers.OfType<Int32>().ToArray());
            //    if (MutualSettlementID2 != 0)
            //    {
            //        decimal DispatchSum = MutualSettlementsManager.GetDispatchSum(MutualSettlementID2);
            //        bool bVerify = MutualSettlementsManager.VerifyTransaction(MutualSettlementID2, DispatchSum, ref DiscountPaymentCondition);
            //        Infinium.Modules.Marketing.Orders.OrdersCalculate OrdersCalculate = new Modules.Marketing.Orders.OrdersCalculate();
            //        for (int i = 0; i < ReturnDT.Rows.Count; i++)
            //        {
            //            int MegaOrderID = Convert.ToInt32(ReturnDT.Rows[i]["MegaOrderID"]);
            //            DataTable FFFDT = MarketingDispatchManager.GetMegaOrders(MegaOrderID);
            //            decimal D = DiscountPaymentCondition;
            //            //int DiscountPaymentConditionID = Convert.ToInt32(FFFDT.Rows[0]["DiscountPaymentConditionID"]);
            //            decimal ProfilDiscountOrderSum = Convert.ToDecimal(FFFDT.Rows[0]["ProfilDiscountOrderSum"]);
            //            decimal TPSDiscountOrderSum = Convert.ToDecimal(FFFDT.Rows[0]["TPSDiscountOrderSum"]);
            //            decimal ProfilDiscountDirector = Convert.ToDecimal(FFFDT.Rows[0]["ProfilDiscountDirector"]);
            //            decimal TPSDiscountDirector = Convert.ToDecimal(FFFDT.Rows[0]["TPSDiscountDirector"]);
            //            decimal ProfilTotalDiscount = Convert.ToDecimal(FFFDT.Rows[0]["ProfilTotalDiscount"]);
            //            decimal TPSTotalDiscount = Convert.ToDecimal(FFFDT.Rows[0]["TPSTotalDiscount"]);
            //            decimal OriginalRate = Convert.ToDecimal(FFFDT.Rows[0]["Rate"]);
            //            decimal PaymentRate = OriginalRate;
            //            object ConfirmDateTime = FFFDT.Rows[0]["ConfirmDateTime"];
            //            if (DiscountPaymentConditionID == 1)
            //                D = 0;
            //            if (DiscountPaymentCondition == 0)
            //                DiscountPaymentConditionID = 1;
            //            if (DiscountPaymentConditionID == 1 && (CurrencyTypeID == 3 || CurrencyTypeID == 5))
            //            {
            //                PaymentRate = OriginalRate * 1.05m;
            //            }
            //            if (bVerify)
            //            {
            //                //ProfilTotalDiscount = DiscountPaymentCondition + ProfilDiscountOrderSum + ProfilDiscountDirector;
            //                TPSTotalDiscount = D + TPSDiscountOrderSum + TPSDiscountDirector;
            //            }
            //            else
            //            {
            //                //ProfilDiscountOrderSum = 0;
            //                TPSDiscountOrderSum = 0;
            //                //ProfilTotalDiscount = ProfilDiscountDirector;
            //                TPSTotalDiscount = TPSDiscountDirector;
            //            }
            //            InvoiceReportToDBF iDBFReport = new InvoiceReportToDBF(ref DecorCatalogOrder, ref OrdersCalculate.FrontsCalculate);
            //            decimal CurrencyTotalCost = iDBFReport.CalcCurrencyCost(MegaOrderID, ClientID, PaymentRate);
            //            OrdersCalculate.Recalculate(MegaOrderID, ProfilDiscountDirector, TPSDiscountDirector,
            //                (ProfilTotalDiscount - ProfilDiscountDirector), (TPSTotalDiscount - TPSDiscountDirector), CurrencyTypeID, PaymentRate, ConfirmDateTime, CurrencyTotalCost, !bVerify);
            //            OrdersCalculate.SetTPSCurrencyCost(MegaOrderID, ProfilTotalDiscount, TPSTotalDiscount,
            //                ProfilDiscountOrderSum, TPSDiscountOrderSum, PaymentRate);
            //            OrdersManager.FixOrderEvent(MegaOrderID, "Выписать с отсрочкой. Пересчитан заказ ТПС");
            //        }
            //    }

            //    if (TotalProfil > 0)
            //    {
            //        //if (ClientCountryID == 1)
            //        //    TotalProfil *= 1.2m;
            //        ClientDispatchID = MutualSettlementsManager.AddClientDispatch(MutualSettlementID1, CurrencyTypeID, TotalProfil);
            //        var fileInfo = new System.IO.FileInfo(ReportName);
            //        MutualSettlementsManager.AttachDispatchDocument(Path.GetFileNameWithoutExtension(ReportName),
            //            fileInfo.Extension, ReportName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.DispatchExcel, ClientDispatchID);
            //        fileInfo = new System.IO.FileInfo(ProfilDBFName);
            //        MutualSettlementsManager.AttachDispatchDocument(Path.GetFileNameWithoutExtension(ProfilDBFName),
            //            fileInfo.Extension, ProfilDBFName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.DispatchDbf, ClientDispatchID);
            //    }
            //    if (TotalTPS > 0)
            //    {
            //        //if (ClientCountryID == 1)
            //        //    TotalTPS *= 1.2m;
            //        ClientDispatchID = MutualSettlementsManager.AddClientDispatch(MutualSettlementID2, CurrencyTypeID, TotalTPS);
            //        var fileInfo = new System.IO.FileInfo(ReportName);
            //        MutualSettlementsManager.AttachDispatchDocument(Path.GetFileNameWithoutExtension(ReportName),
            //            fileInfo.Extension, ReportName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.DispatchExcel, ClientDispatchID);
            //        fileInfo = new System.IO.FileInfo(TPSDBFName);
            //        MutualSettlementsManager.AttachDispatchDocument(Path.GetFileNameWithoutExtension(TPSDBFName),
            //            fileInfo.Extension, TPSDBFName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.DispatchDbf, ClientDispatchID);
            //    }
            //    MutualSettlementsManager.SaveMutualSettlements();
            //    MutualSettlementsManager.SaveMutualSettlementOrders();

            //    MarketingDispatchManager.IncludeInMutualSettlement(DispatchID, MutualSettlementID1, MutualSettlementID2);
            //    UpdateDispatchDate();
            //    MarketingDispatchManager.MoveToDispatchDate(DispatchDate);
            //    MarketingDispatchManager.MoveToDispatch(DispatchID);
            //}
            //SendEmail SendEmail = new SendEmail();
            //SendEmail.Send(ClientID, string.Join(",", OrderNumbers.OfType<Int32>().ToArray()), DetailsReport.Save, ReportName);

            //if (SendEmail.Success == false)
            //{
            //    Infinium.LightMessageBox.Show(ref TopForm, false,
            //           "Отправка отчета невозможна: отсутствует подключение к интернету либо адрес электронной почты указан неверно",
            //           "Отправка письма");
            //}

            //NeedSplash = true;
            //while (SplashWindow.bSmallCreated)
            //    SmallWaitForm.CloseS = true;
        }

        private void kryptonContextMenuItem14_Click(object sender, EventArgs e)
        {
            bool PressOK = false;
            int NewClientID = 0;

            PhantomForm PhantomForm = new Infinium.PhantomForm();
            PhantomForm.Show();

            ReaddressMarketingDispatchMenu ReaddressMarketingDispatchMenu = new ReaddressMarketingDispatchMenu(this, MarketingExpeditionManager.ClientsDT);
            TopForm = ReaddressMarketingDispatchMenu;
            ReaddressMarketingDispatchMenu.ShowDialog();

            PressOK = ReaddressMarketingDispatchMenu.PressOK;
            NewClientID = ReaddressMarketingDispatchMenu.ClientID;

            PhantomForm.Close();
            PhantomForm.Dispose();
            ReaddressMarketingDispatchMenu.Dispose();
            TopForm = null;

            if (!PressOK)
                return;

            Thread T1 = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Создание документа Excel.\r\nПодождите..."); });
            T1.Start();

            while (!SplashWindow.bSmallCreated) ;
            NeedSplash = false;

            bool InMutualSettlement = Convert.ToBoolean(dgvDispatch.SelectedRows[0].Cells["InMutualSettlement"].Value);
            int DispatchID = Convert.ToInt32(dgvDispatch.SelectedRows[0].Cells["DispatchID"].Value);
            DateTime DispatchDate = Convert.ToDateTime(dgvDispatchDates.SelectedRows[0].Cells["PrepareDispatchDateTime"].Value);

            MarketingDispatchManager.ReaddressDispatch(DispatchID, NewClientID);
            string ClientName = dgvDispatch.SelectedRows[0].Cells["NewClientName"].Value.ToString();
            int[] Dispatches = new int[dgvDispatch.SelectedRows.Count];
            for (int i = 0; i < dgvDispatch.SelectedRows.Count; i++)
                Dispatches[i] = Convert.ToInt32(dgvDispatch.SelectedRows[i].Cells["DispatchID"].Value);
            int[] MainOrders = MarketingDispatchManager.GetMainOrders(Dispatches);

            decimal DiscountPaymentCondition = 0;
            decimal TotalProfil = 0;
            decimal TotalTPS = 0;

            string ProfilDBFName = string.Empty;
            string TPSDBFName = string.Empty;
            string ReportName = string.Empty;
            DataTable ReturnDT = MarketingDispatchManager.GetMegaOrders(Dispatches);
            ArrayList OrderNumbers = new ArrayList();
            for (int i = 0; i < ReturnDT.Rows.Count; i++)
                OrderNumbers.Add(Convert.ToInt32(ReturnDT.Rows[i]["OrderNumber"]));

            int CurrencyTypeID = 1;
            int DiscountPaymentConditionID = 1;
            if (ReturnDT.Rows.Count > 0)
            {
                CurrencyTypeID = Convert.ToInt32(ReturnDT.Rows[0]["CurrencyTypeID"]);
                DiscountPaymentConditionID = Convert.ToInt32(ReturnDT.Rows[0]["DiscountPaymentConditionID"]);
            }

            int ClientID = Convert.ToInt32(dgvDispatch.SelectedRows[0].Cells["ClientID"].Value);
            int ClientCountryID = OrdersManager.GetClientCountry(ClientID);
            int DelayOfPayment = OrdersManager.GetDelayOfPayment(ClientID);
            //int DiscountPaymentConditionID = 1;
            int MutualSettlementID = -1;
            string ClientLogin = OrdersManager.GetClientLogin(NewClientID);
            Infinium.Modules.Marketing.MutualSettlements.MutualSettlements MutualSettlementsManager = new Modules.Marketing.MutualSettlements.MutualSettlements();
            MutualSettlementsManager.Initialize();

            ClientName = ClientName.Replace("\"", " ");
            ClientName = ClientName.Replace("*", " ");
            ClientName = ClientName.Replace("|", " ");
            ClientName = ClientName.Replace(@"\", " ");
            ClientName = ClientName.Replace(":", " ");
            ClientName = ClientName.Replace("<", " ");
            ClientName = ClientName.Replace(">", " ");
            ClientName = ClientName.Replace("?", " ");
            ClientName = ClientName.Replace("/", " ");
            ReportName = ClientLogin + " №" + string.Join(",", Dispatches);

            string CurrentMonthName = DateTime.Now.ToString("dd-MM-yyyy");
            string FilePath = Security.DBFSaveFilePath;
            if (Security.DBFSaveFilePath.Length == 0)
                FilePath = Environment.GetFolderPath(Environment.SpecialFolder.Personal) + "/Infinium/1C data/" + CurrentMonthName;

            HSSFWorkbook hssfworkbook = new HSSFWorkbook();

            int MutualSettlementID1 = -1;
            int MutualSettlementID2 = -1;
            bool IsSample = false;
            bool bProfilVerify = false;
            MutualSettlementID1 = MutualSettlementsManager.FindMutualSettlementIDByOrderNumber(ClientID, 1, OrderNumbers.OfType<Int32>().ToArray());
            decimal DispatchSum = MutualSettlementsManager.GetDispatchSum(MutualSettlementID1);
            if (MutualSettlementID1 != 0)
                bProfilVerify = MutualSettlementsManager.VerifyTransaction(MutualSettlementID1, DispatchSum, ref DiscountPaymentCondition);

            bool bTPSVerify = false;
            MutualSettlementID2 = MutualSettlementsManager.FindMutualSettlementIDByOrderNumber(ClientID, 2, OrderNumbers.OfType<Int32>().ToArray());
            DispatchSum = MutualSettlementsManager.GetDispatchSum(MutualSettlementID2);
            if (MutualSettlementID2 != 0)
                bTPSVerify = MutualSettlementsManager.VerifyTransaction(MutualSettlementID2, DispatchSum, ref DiscountPaymentCondition);

            _dispatchReportToDbf.CreateReport(ref hssfworkbook, Dispatches, MainOrders, NewClientID, ClientName,
                ref TotalProfil, ref TotalTPS, bProfilVerify, bTPSVerify, DiscountPaymentConditionID, IsSample);
            _dispatchReportToDbf.SaveDBF(FilePath, ReportName, ref ProfilDBFName, ref TPSDBFName);

            DetailsReport.CreateReport(ref hssfworkbook, Dispatches, OrderNumbers.OfType<Int32>().ToArray(), MainOrders, NewClientID, ClientName,
                bProfilVerify, bTPSVerify, DiscountPaymentConditionID, false);

            if (ClientID == 609)
            {
                DetailsReport.ReportModuleOnline(ref hssfworkbook, Dispatches);
            }
            FileInfo file = new FileInfo(FilePath + @"\" + ReportName + ".xls");
            int j = 1;
            while (file.Exists == true)
            {
                file = new FileInfo(FilePath + @"\" + ReportName + "(" + j++ + ").xls");
            }

            FileStream NewFile = new FileStream(file.FullName, FileMode.Create);
            hssfworkbook.Write(NewFile);
            NewFile.Close();

            ReportName = file.FullName;

            NeedSplash = true;
            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;

            PhantomForm = new Infinium.PhantomForm();
            PhantomForm.Show();

            int Result = 1;
            string Notes = string.Empty;
            string SaveFilePath = string.Empty;
            SaveDBFReportMenu = new Infinium.SaveDBFReportMenu(this, ReportName);

            TopForm = SaveDBFReportMenu;
            SaveDBFReportMenu.ShowDialog();

            PhantomForm.Close();
            Result = SaveDBFReportMenu.Result;
            Notes = SaveDBFReportMenu.Notes;
            SaveFilePath = SaveDBFReportMenu.SaveFilePath;
            PhantomForm.Dispose();
            SaveDBFReportMenu.Dispose();
            TopForm = null;

            if (Result != 1)
            {
                return;
            }

            T1 = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Создание документа Excel.\r\nПодождите..."); });
            T1.Start();

            while (!SplashWindow.bSmallCreated) ;
            NeedSplash = false;

            if (TotalProfil > 0)
            {
                if (ClientCountryID == 1)
                    TotalProfil *= 1.2m;
                TotalProfil = Decimal.Round(TotalProfil, 2, MidpointRounding.AwayFromZero);
                MutualSettlementID = MutualSettlementsManager.AddMutualSettlement(NewClientID, 1, CurrencyTypeID, DiscountPaymentConditionID, DelayOfPayment, TotalProfil, Notes);
                var fileInfo = new System.IO.FileInfo(ReportName);
                MutualSettlementsManager.AttachIncomeDocument(Path.GetFileNameWithoutExtension(ReportName),
                    fileInfo.Extension, ReportName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.InvoiceExcel, MutualSettlementID);
                fileInfo = new System.IO.FileInfo(ProfilDBFName);
                MutualSettlementsManager.AttachIncomeDocument(Path.GetFileNameWithoutExtension(ProfilDBFName),
                    fileInfo.Extension, ProfilDBFName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.InvoiceDbf, MutualSettlementID);
                DateTime CurrentDate = Security.GetCurrentDate();
                for (int i = 0; i < OrderNumbers.OfType<Int32>().ToArray().Count(); i++)
                    MutualSettlementsManager.AddMutualSettlementOrder(CurrentDate, NewClientID, MutualSettlementID, OrderNumbers.OfType<Int32>().ToArray()[i]);
            }
            if (TotalTPS > 0)
            {
                if (ClientCountryID == 1)
                    TotalTPS *= 1.2m;
                TotalTPS = Decimal.Round(TotalTPS, 2, MidpointRounding.AwayFromZero);
                MutualSettlementID = MutualSettlementsManager.AddMutualSettlement(NewClientID, 2, CurrencyTypeID, DiscountPaymentConditionID, DelayOfPayment, TotalTPS, Notes);
                var fileInfo = new System.IO.FileInfo(ReportName);
                MutualSettlementsManager.AttachIncomeDocument(Path.GetFileNameWithoutExtension(ReportName),
                    fileInfo.Extension, ReportName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.InvoiceExcel, MutualSettlementID);
                fileInfo = new System.IO.FileInfo(TPSDBFName);
                MutualSettlementsManager.AttachIncomeDocument(Path.GetFileNameWithoutExtension(TPSDBFName),
                    fileInfo.Extension, TPSDBFName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.InvoiceDbf, MutualSettlementID);
                DateTime CurrentDate = Security.GetCurrentDate();
                for (int i = 0; i < OrderNumbers.OfType<Int32>().ToArray().Count(); i++)
                    MutualSettlementsManager.AddMutualSettlementOrder(CurrentDate, NewClientID, MutualSettlementID, OrderNumbers.OfType<Int32>().ToArray()[i]);
            }
            MutualSettlementsManager.SaveMutualSettlements();
            MutualSettlementsManager.SaveMutualSettlementOrders();


            InMutualSettlement = false;
            if (!InMutualSettlement)
            {
                int ClientDispatchID = -1;

                MutualSettlementsManager.Initialize();

                MutualSettlementID1 = MutualSettlementsManager.FindMutualSettlementID(NewClientID, 1);
                if (MutualSettlementID1 != 0)
                {
                    for (int i = 0; i < ReturnDT.Rows.Count; i++)
                    {
                        int MegaOrderID = Convert.ToInt32(ReturnDT.Rows[i]["MegaOrderID"]);

                        MegaOrderInfo m = MarketingDispatchManager.GetMegaOrders(MegaOrderID);
                        decimal D = DiscountPaymentCondition;
                        DiscountPaymentConditionID = m.DiscountPaymentConditionID;

                        if (DiscountPaymentConditionID == 1)
                            D = 0;
                        if (DiscountPaymentCondition == 0)
                            DiscountPaymentConditionID = 1;
                        if (ClientID != 145 && DiscountPaymentConditionID == 1 && (CurrencyTypeID == 3 || CurrencyTypeID == 5))
                        {
                            //m.PaymentRate = m.OriginalRate * 1.05m;
                        }

                        if (bProfilVerify)
                        {
                            m.ProfilTotalDiscount = D + m.ProfilDiscountOrderSum + m.ProfilDiscountDirector;
                            //TPSTotalDiscount = DiscountPaymentCondition + TPSDiscountOrderSum + TPSDiscountDirector;
                        }
                        else
                        {
                            m.ProfilDiscountOrderSum = 0;
                            //TPSDiscountOrderSum = 0;
                            m.ProfilTotalDiscount = m.ProfilDiscountDirector;
                            //TPSTotalDiscount = TPSDiscountDirector;
                        }

                        decimal CurrencyTotalCost = iDBFReport.CalcCurrencyCost(MegaOrderID, ClientID, m.PaymentRate);
                        OrdersCalculate.Recalculate(MegaOrderID, m.ProfilDiscountDirector, m.TPSDiscountDirector,
                            (m.ProfilTotalDiscount - m.ProfilDiscountDirector), (m.TPSTotalDiscount - m.TPSDiscountDirector), m.DiscountPaymentCondition, CurrencyTypeID, m.PaymentRate, m.ConfirmDateTime, CurrencyTotalCost, !bProfilVerify);
                        OrdersCalculate.SetProfilCurrencyCost(MegaOrderID, m.ProfilTotalDiscount, m.TPSTotalDiscount,
                            m.ProfilDiscountOrderSum, m.TPSDiscountOrderSum, m.PaymentRate);
                        OrdersManager.FixOrderEvent(MegaOrderID, "Выписать с отсрочкой. Пересчитан заказ Профиль");
                    }
                }
                MutualSettlementID2 = 0;
                MutualSettlementID2 = MutualSettlementsManager.FindMutualSettlementID(NewClientID, 2);
                if (MutualSettlementID2 != 0)
                {
                    for (int i = 0; i < ReturnDT.Rows.Count; i++)
                    {
                        int MegaOrderID = Convert.ToInt32(ReturnDT.Rows[i]["MegaOrderID"]);
                        MegaOrderInfo m = MarketingDispatchManager.GetMegaOrders(MegaOrderID);
                        decimal D = DiscountPaymentCondition;
                        DiscountPaymentConditionID = m.DiscountPaymentConditionID;

                        if (DiscountPaymentConditionID == 1)
                            D = 0;
                        if (DiscountPaymentCondition == 0)
                            DiscountPaymentConditionID = 1;
                        if (ClientID != 145 && DiscountPaymentConditionID == 1 && (CurrencyTypeID == 3 || CurrencyTypeID == 5))
                        {
                            //m.PaymentRate = m.OriginalRate * 1.05m;
                        }
                        if (bTPSVerify)
                        {
                            //ProfilTotalDiscount = DiscountPaymentCondition + ProfilDiscountOrderSum + ProfilDiscountDirector;
                            m.TPSTotalDiscount = D + m.TPSDiscountOrderSum + m.TPSDiscountDirector;
                        }
                        else
                        {
                            //ProfilDiscountOrderSum = 0;
                            m.TPSDiscountOrderSum = 0;
                            //ProfilTotalDiscount = ProfilDiscountDirector;
                            m.TPSTotalDiscount = m.TPSDiscountDirector;
                        }
                        decimal CurrencyTotalCost = iDBFReport.CalcCurrencyCost(MegaOrderID, ClientID, m.PaymentRate);
                        OrdersCalculate.Recalculate(MegaOrderID, m.ProfilDiscountDirector, m.TPSDiscountDirector,
                            (m.ProfilTotalDiscount - m.ProfilDiscountDirector), (m.TPSTotalDiscount - m.TPSDiscountDirector), m.DiscountPaymentCondition, CurrencyTypeID, m.PaymentRate, m.ConfirmDateTime, CurrencyTotalCost, !bTPSVerify);
                        OrdersCalculate.SetTPSCurrencyCost(MegaOrderID, m.ProfilTotalDiscount, m.TPSTotalDiscount,
                            m.ProfilDiscountOrderSum, m.TPSDiscountOrderSum, m.PaymentRate);
                        OrdersManager.FixOrderEvent(MegaOrderID, "Выписать с отсрочкой. Пересчитан заказ ТПС");
                    }
                }

                if (TotalProfil > 0)
                {
                    //if (ClientCountryID == 1)
                    //    TotalProfil *= 1.2m;
                    ClientDispatchID = MutualSettlementsManager.AddClientDispatch(MutualSettlementID1, CurrencyTypeID, TotalProfil);
                    var fileInfo = new System.IO.FileInfo(ReportName);
                    MutualSettlementsManager.AttachDispatchDocument(Path.GetFileNameWithoutExtension(ReportName),
                        fileInfo.Extension, ReportName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.DispatchExcel, ClientDispatchID);
                    fileInfo = new System.IO.FileInfo(ProfilDBFName);
                    MutualSettlementsManager.AttachDispatchDocument(Path.GetFileNameWithoutExtension(ProfilDBFName),
                        fileInfo.Extension, ProfilDBFName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.DispatchDbf, ClientDispatchID);
                }
                if (TotalTPS > 0)
                {
                    //if (ClientCountryID == 1)
                    //    TotalTPS *= 1.2m;
                    ClientDispatchID = MutualSettlementsManager.AddClientDispatch(MutualSettlementID2, CurrencyTypeID, TotalTPS);
                    var fileInfo = new System.IO.FileInfo(ReportName);
                    MutualSettlementsManager.AttachDispatchDocument(Path.GetFileNameWithoutExtension(ReportName),
                        fileInfo.Extension, ReportName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.DispatchExcel, ClientDispatchID);
                    fileInfo = new System.IO.FileInfo(TPSDBFName);
                    MutualSettlementsManager.AttachDispatchDocument(Path.GetFileNameWithoutExtension(TPSDBFName),
                        fileInfo.Extension, TPSDBFName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.DispatchDbf, ClientDispatchID);
                }
                MutualSettlementsManager.SaveMutualSettlements();
                MutualSettlementsManager.SaveMutualSettlementOrders();

                MarketingDispatchManager.IncludeInMutualSettlement(DispatchID, MutualSettlementID1, MutualSettlementID2);
                UpdateDispatchDate();
                MarketingDispatchManager.MoveToDispatchDate(DispatchDate);
                MarketingDispatchManager.MoveToDispatch(DispatchID);
            }
            SendEmail SendEmail = new SendEmail();
            SendEmail.Send(ClientID, string.Join(",", OrderNumbers.OfType<Int32>().ToArray()), DetailsReport.Save, ReportName);

            if (SendEmail.Success == false)
            {
                Infinium.LightMessageBox.Show(ref TopForm, false,
                       "Отправка отчета невозможна: отсутствует подключение к интернету либо адрес электронной почты указан неверно",
                       "Отправка письма");
            }

            NeedSplash = true;
            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
        }

        private void MainOrdersDecorOrdersDataGrid_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if ((RoleType == RoleTypes.AdminRole || RoleType == RoleTypes.Production)
                && e.Button == System.Windows.Forms.MouseButtons.Right)
            {
                PercentageDataGrid grid = (PercentageDataGrid)sender;
                grid.CurrentCell = grid.Rows[e.RowIndex].Cells[e.ColumnIndex];
                kryptonContextMenu5.Show(new Point(Cursor.Position.X - 212, Cursor.Position.Y - 10));
            }
        }

        private void kryptonContextMenuItem8_Click(object sender, EventArgs e)
        {
            if (MainOrdersDecorOrdersDataGrid.SelectedRows.Count == 0 || MainOrdersDecorOrdersDataGrid.SelectedRows[0].Cells["DecorOrderID"].Value == DBNull.Value)
                return;

            int DecorOrderID = Convert.ToInt32(MainOrdersDecorOrdersDataGrid.SelectedRows[0].Cells["DecorOrderID"].Value);
            int PackageID = 0;
            int OldCount = Convert.ToInt32(MainOrdersDecorOrdersDataGrid.SelectedRows[0].Cells["Count"].Value);
            int NewCount = 0;
            if (PackagesDataGrid.SelectedRows.Count > 0 && PackagesDataGrid.SelectedRows[0].Cells["PackageID"].Value != DBNull.Value)
                PackageID = Convert.ToInt32(PackagesDataGrid.SelectedRows[0].Cells["PackageID"].Value);
            PhantomForm PhantomForm = new Infinium.PhantomForm();
            PhantomForm.Show();

            InputNewDecorCountForm = new InputNewDecorCountForm(this, OldCount);

            TopForm = InputNewDecorCountForm;
            InputNewDecorCountForm.ShowDialog();
            TopForm = null;

            PhantomForm.Close();

            PhantomForm.Dispose();
            if (InputNewDecorCountForm.Result == 1)
            {
                NewCount = InputNewDecorCountForm.NewCount;
                MarketingExpeditionManager.EditDecorCount(PackageID, DecorOrderID, NewCount);
                OrdersManager.FixOrderEvent(Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[0].Cells["MegaOrderID"].Value),
                    "Изменено кол-во. PackageID=" + PackageID.ToString() + ". DecorOrderID=" + DecorOrderID.ToString() + ". Было=" + OldCount.ToString() + ". Стало=" + NewCount.ToString());
                int FactoryID = -1;

                if (ProfilCheckBox.Checked && TPSCheckBox.Checked)
                    FactoryID = 0;
                if (ProfilCheckBox.Checked && !TPSCheckBox.Checked)
                    FactoryID = 1;
                if (!ProfilCheckBox.Checked && TPSCheckBox.Checked)
                    FactoryID = 2;
                MarketingExpeditionManager.FilterProductsByPackage();
                MarketingExpeditionManager.FilterProductsByPackage(Convert.ToInt32(MainOrdersDataGrid.SelectedRows[0].Cells["MainOrderID"].Value), FactoryID);
                MarketingExpeditionManager.FilterPackagesByMainOrder(Convert.ToInt32(MainOrdersDataGrid.SelectedRows[0].Cells["MainOrderID"].Value), FactoryID);
                MarketingExpeditionManager.MoveToPackage(PackageID);
            }
            InputNewDecorCountForm.Dispose();
            InputNewDecorCountForm = null;
            GC.Collect();
        }

        private void kryptonContextMenuItem10_Click(object sender, EventArgs e)
        {
            if (MainOrdersDecorOrdersDataGrid.SelectedRows.Count == 0 || MainOrdersDecorOrdersDataGrid.SelectedRows[0].Cells["DecorOrderID"].Value == DBNull.Value)
                return;

            int PackageID = 0;
            if (PackagesDataGrid.SelectedRows.Count > 0 && PackagesDataGrid.SelectedRows[0].Cells["PackageID"].Value != DBNull.Value)
                PackageID = Convert.ToInt32(PackagesDataGrid.SelectedRows[0].Cells["PackageID"].Value);
            int DecorOrderID = Convert.ToInt32(MainOrdersDecorOrdersDataGrid.SelectedRows[0].Cells["DecorOrderID"].Value);
            int OldCount = Convert.ToInt32(MainOrdersDecorOrdersDataGrid.SelectedRows[0].Cells["Length"].Value);
            int NewCount = 0;

            PhantomForm PhantomForm = new Infinium.PhantomForm();
            PhantomForm.Show();

            InputNewDecorCountForm = new InputNewDecorCountForm(this, OldCount);

            TopForm = InputNewDecorCountForm;
            InputNewDecorCountForm.ShowDialog();
            TopForm = null;

            PhantomForm.Close();

            PhantomForm.Dispose();
            if (InputNewDecorCountForm.Result == 1)
            {
                NewCount = InputNewDecorCountForm.NewCount;
                MarketingExpeditionManager.EditDecorLength(DecorOrderID, NewCount);
                OrdersManager.FixOrderEvent(Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[0].Cells["MegaOrderID"].Value),
                    "Изменена длина. PackageID=" + PackageID + ". DecorOrderID=" + DecorOrderID + ". Было=" + OldCount + ". Стало=" + NewCount);
                int FactoryID = -1;

                if (ProfilCheckBox.Checked && TPSCheckBox.Checked)
                    FactoryID = 0;
                if (ProfilCheckBox.Checked && !TPSCheckBox.Checked)
                    FactoryID = 1;
                if (!ProfilCheckBox.Checked && TPSCheckBox.Checked)
                    FactoryID = 2;
                MarketingExpeditionManager.FilterProductsByPackage();
                MarketingExpeditionManager.FilterProductsByPackage(Convert.ToInt32(MainOrdersDataGrid.SelectedRows[0].Cells["MainOrderID"].Value), FactoryID);
                MarketingExpeditionManager.FilterPackagesByMainOrder(Convert.ToInt32(MainOrdersDataGrid.SelectedRows[0].Cells["MainOrderID"].Value), FactoryID);
            }
            InputNewDecorCountForm.Dispose();
            InputNewDecorCountForm = null;
            GC.Collect();
        }

        private void dgvMainOrders_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if ((RoleType != RoleTypes.OrdinaryRole) && e.Button == System.Windows.Forms.MouseButtons.Right)
            {
                kryptonContextMenu4.Show(new Point(Cursor.Position.X - 212, Cursor.Position.Y - 10));
            }
        }

        private void kryptonContextMenuItem7_Click(object sender, EventArgs e)
        {
            DateTime DispatchDate = Convert.ToDateTime(dgvDispatchDates.SelectedRows[0].Cells["PrepareDispatchDateTime"].Value);
            int DispatchID = 0;
            int ClientID = 0;
            int[] MainOrders = new int[dgvMainOrders.SelectedRows.Count];
            for (int i = 0; i < dgvMainOrders.SelectedRows.Count; i++)
                MainOrders[i] = Convert.ToInt32(dgvMainOrders.SelectedRows[i].Cells["MainOrderID"].Value);
            if (dgvDispatch.SelectedRows.Count > 0)
            {
                DispatchID = Convert.ToInt32(dgvDispatch.SelectedRows[0].Cells["DispatchID"].Value);
                ClientID = Convert.ToInt32(dgvDispatch.SelectedRows[0].Cells["ClientID"].Value);
            }
            if (DispatchID != 0 && MainOrders.Count() > 0)
            {
                int NewDispatchID = MarketingExpeditionManager.MoveMainOrdersToAnotherDispatch(MainOrders, DispatchID);
                UpdateDispatchDate();
                MarketingDispatchManager.MoveToDispatchDate(DispatchDate);
                MarketingDispatchManager.MoveToDispatch(DispatchID);
            }
        }

        private void kryptonButton4_Click(object sender, EventArgs e)
        {
            CheckOrdersStatus.SetToStorage(monthCalendar3.SelectionStart, GetSelectedPackages());
        }

        private void PackagesDataGrid_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Right)
            {
                kryptonContextMenu6.Show(new Point(Cursor.Position.X - 212, Cursor.Position.Y - 10));
            }
        }

        private void kryptonContextMenuItem9_Click(object sender, EventArgs e)
        {
            bool ColVisible = false;
            if (PackagesDataGrid.SelectedRows.Count > 0 && PackagesDataGrid.SelectedRows[0].Cells["PackUsersColumn"].Value != DBNull.Value)
                ColVisible = Convert.ToBoolean(PackagesDataGrid.SelectedRows[0].Cells["PackUsersColumn"].Visible);
            MarketingExpeditionManager.ShowPackUsersColumn(!ColVisible);
        }

        private void kryptonContextMenuItem11_Click(object sender, EventArgs e)
        {
            bool ColVisible = false;
            if (PackagesDataGrid.SelectedRows.Count > 0 && PackagesDataGrid.SelectedRows[0].Cells["StoreUsersColumn"].Value != DBNull.Value)
                ColVisible = Convert.ToBoolean(PackagesDataGrid.SelectedRows[0].Cells["StoreUsersColumn"].Visible);
            MarketingExpeditionManager.ShowStoreUsersColumn(!ColVisible);
        }

        private void kryptonContextMenuItem12_Click(object sender, EventArgs e)
        {
            bool ColVisible = false;
            if (PackagesDataGrid.SelectedRows.Count > 0 && PackagesDataGrid.SelectedRows[0].Cells["ExpUsersColumn"].Value != DBNull.Value)
                ColVisible = Convert.ToBoolean(PackagesDataGrid.SelectedRows[0].Cells["ExpUsersColumn"].Visible);
            MarketingExpeditionManager.ShowExpUsersColumn(!ColVisible);
        }

        private void kryptonContextMenuItem13_Click(object sender, EventArgs e)
        {
            bool ColVisible = false;
            if (PackagesDataGrid.SelectedRows.Count > 0 && PackagesDataGrid.SelectedRows[0].Cells["DispUsersColumn"].Value != DBNull.Value)
                ColVisible = Convert.ToBoolean(PackagesDataGrid.SelectedRows[0].Cells["DispUsersColumn"].Visible);
            MarketingExpeditionManager.ShowDispUsersColumn(!ColVisible);
        }

        private void kryptonMonthCalendar1_DateChanged(object sender, DateRangeEventArgs e)
        {

        }

        private void dgvDispatch_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            if (e.RowIndex < 0)
                return;
            PercentageDataGrid grid = (PercentageDataGrid)sender;
            int PaidStatus = -1;
            if (grid.Rows[e.RowIndex].Cells["PaidStatus"].Value != DBNull.Value)
                PaidStatus = Convert.ToInt32(grid.Rows[e.RowIndex].Cells["PaidStatus"].Value);

            if (PaidStatus == 0)
            {
                grid.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Red;
                grid.Rows[e.RowIndex].DefaultCellStyle.ForeColor = Color.White;
                grid.Rows[e.RowIndex].DefaultCellStyle.SelectionForeColor = Color.Red;
            }
            if (PaidStatus == 1)
            {
                grid.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Yellow;
                grid.Rows[e.RowIndex].DefaultCellStyle.ForeColor = Color.Black;
                grid.Rows[e.RowIndex].DefaultCellStyle.SelectionForeColor = Color.Yellow;
            }
            if (PaidStatus == 2)
            {
                grid.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Orange;
                grid.Rows[e.RowIndex].DefaultCellStyle.ForeColor = Color.White;
                grid.Rows[e.RowIndex].DefaultCellStyle.SelectionForeColor = Color.Orange;
            }
            if (PaidStatus == -1)
            {
                grid.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.White;
                grid.Rows[e.RowIndex].DefaultCellStyle.ForeColor = Color.Black;
                grid.Rows[e.RowIndex].DefaultCellStyle.SelectionForeColor = Color.White;
            }
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
            T.Start();

            while (!SplashWindow.bSmallCreated) ;
            NeedSplash = false;
            UpdateDispatchDate();
            NeedSplash = true;
            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
        }

        private void kryptonContextMenuItem16_Click(object sender, EventArgs e)
        {
            if (PackagesDataGrid.SelectedRows.Count == 0)
                return;
            int PackageID = Convert.ToInt32(PackagesDataGrid.SelectedRows[0].Cells["PackageID"].Value);
            int DispatchID = 0;

            if (MarketingDispatchManager.IsPackageInDispatch(PackageID, ref DispatchID))
            {
                object PrepareDispatchDateTime = MarketingDispatchManager.GetPrepareDispatchDateTime(DispatchID);
                if (PrepareDispatchDateTime != DBNull.Value)
                {
                    Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                    T.Start();

                    while (!SplashWindow.bSmallCreated) ;

                    NeedSplash = false;

                    cbxYears.SelectedValue = Convert.ToDateTime(PrepareDispatchDateTime).Year;
                    cbxMonths.SelectedValue = Convert.ToDateTime(PrepareDispatchDateTime).Month;
                    UpdateDispatchDate();
                    MarketingDispatchManager.MoveToDispatchDate(Convert.ToDateTime(PrepareDispatchDateTime));
                    MarketingDispatchManager.MoveToDispatch(DispatchID);
                    cbtnDispatch.Checked = true;

                    NeedSplash = true;

                    while (SplashWindow.bSmallCreated)
                        SmallWaitForm.CloseS = true;
                }
            }
            else
            {
                Infinium.LightMessageBox.Show(ref TopForm, false,
                        "Упаковка не включена в отгрузку",
                        "Поиск отгрузки");
                return;
            }
        }

        private void kryptonContextMenuItem15_Click(object sender, EventArgs e)
        {
            if (PackagesDataGrid.SelectedRows.Count == 0)
                return;
            int PackageID = Convert.ToInt32(PackagesDataGrid.SelectedRows[0].Cells["PackageID"].Value);
            int DispatchID = 0;

            if (MarketingDispatchManager.IsPackageInDispatch(PackageID, ref DispatchID))
            {
                object PrepareDispatchDateTime = MarketingDispatchManager.GetPrepareDispatchDateTime(DispatchID);
                if (PrepareDispatchDateTime != DBNull.Value)
                {
                    Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                    T.Start();

                    while (!SplashWindow.bSmallCreated) ;

                    NeedSplash = false;

                    cbxYears.SelectedValue = Convert.ToDateTime(PrepareDispatchDateTime).Year;
                    cbxMonths.SelectedValue = Convert.ToDateTime(PrepareDispatchDateTime).Month;
                    UpdateDispatchDate();
                    MarketingDispatchManager.MoveToDispatchDate(Convert.ToDateTime(PrepareDispatchDateTime));
                    MarketingDispatchManager.MoveToDispatch(DispatchID);
                    cbtnDispatch.Checked = true;

                    NeedSplash = true;

                    while (SplashWindow.bSmallCreated)
                        SmallWaitForm.CloseS = true;
                }
            }
            else
            {
                Infinium.LightMessageBox.Show(ref TopForm, false,
                        "Заказ не включен в отгрузку",
                        "Поиск отгрузки");
                return;
            }
        }

        private void MainOrdersDataGrid_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if ((RoleType == RoleTypes.AdminRole || RoleType == RoleTypes.LogisticsRole)
                && e.Button == System.Windows.Forms.MouseButtons.Right)
            {
                kryptonContextMenu7.Show(new Point(Cursor.Position.X - 212, Cursor.Position.Y - 10));
            }
        }

        private void kryptonContextMenuItem19_Click(object sender, EventArgs e)
        {
            if (PackagesDataGrid.SelectedRows.Count == 0)
                return;
            int PackageID = Convert.ToInt32(PackagesDataGrid.SelectedRows[0].Cells["PackageID"].Value);
            int DispatchID = 0;

            if (MarketingDispatchManager.IsPackageInDispatch(PackageID, ref DispatchID))
            {
                object PrepareDispatchDateTime = MarketingDispatchManager.GetPrepareDispatchDateTime(DispatchID);
                if (PrepareDispatchDateTime != DBNull.Value)
                {
                    Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                    T.Start();

                    while (!SplashWindow.bSmallCreated) ;

                    NeedSplash = false;

                    cbxYears.SelectedValue = Convert.ToDateTime(PrepareDispatchDateTime).Year;
                    cbxMonths.SelectedValue = Convert.ToDateTime(PrepareDispatchDateTime).Month;
                    UpdateDispatchDate();
                    MarketingDispatchManager.MoveToDispatchDate(Convert.ToDateTime(PrepareDispatchDateTime));
                    MarketingDispatchManager.MoveToDispatch(DispatchID);
                    cbtnDispatch.Checked = true;

                    NeedSplash = true;

                    while (SplashWindow.bSmallCreated)
                        SmallWaitForm.CloseS = true;
                }
            }
            else
            {
                Infinium.LightMessageBox.Show(ref TopForm, false,
                        "Подзаказ не включен в отгрузку",
                        "Поиск отгрузки");
                return;
            }
        }

        private void kryptonButton5_Click(object sender, EventArgs e)
        {
            if (Convert.ToInt32(kryptonButton5.Tag) == 1)
            {
                MenuPanel.Height = MenuPanel.Height + 210;
                kryptonButton5.Tag = 0;
            }
            else
            {
                MenuPanel.Height = MenuPanel.Height - 210;
                kryptonButton5.Tag = 1;
            }
        }

        private void kryptonButton3_Click_1(object sender, EventArgs e)
        {
            CheckOrdersStatus.SetToExpedition(monthCalendar3.SelectionStart, GetSelectedPackages());
        }

        private void kryptonButton6_Click(object sender, EventArgs e)
        {
            CheckOrdersStatus.SetToPack(monthCalendar3.SelectionStart, GetSelectedPackages());
        }

        private void kryptonButton7_Click(object sender, EventArgs e)
        {
            CheckOrdersStatus.SetToNotPack(GetSelectedPackages());
        }

        private void dgvCabFur_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {

        }

        private void dgvCabFur_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if ((RoleType != RoleTypes.OrdinaryRole) && e.Button == System.Windows.Forms.MouseButtons.Right)
            {
                dgvCabFurMegaOrders.Rows[e.RowIndex].Selected = true;
                kryptonContextMenu8.Show(new Point(Cursor.Position.X - 212, Cursor.Position.Y - 10));
            }
        }

        private void kryptonContextMenuItem18_Click(object sender, EventArgs e)
        {
            if (dgvCabFurMegaOrders.SelectedRows.Count == 0)
                return;
            int MegaOrderID = Convert.ToInt32(dgvCabFurMegaOrders.SelectedRows[0].Cells["MegaOrderID"].Value);

            object PrepareDispatchDateTime = MarketingDispatchManager.GetPrepareDispatchDateTimeByMegaOrder(MegaOrderID);
            if (PrepareDispatchDateTime != DBNull.Value)
            {
                Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                T.Start();

                while (!SplashWindow.bSmallCreated) ;

                NeedSplash = false;

                cbxYears.SelectedValue = Convert.ToDateTime(PrepareDispatchDateTime).Year;
                cbxMonths.SelectedValue = Convert.ToDateTime(PrepareDispatchDateTime).Month;
                UpdateDispatchDate();
                MarketingDispatchManager.MoveToDispatchDate(Convert.ToDateTime(PrepareDispatchDateTime));
                //MarketingDispatchManager.MoveToDispatch(MegaOrderID);
                cbtnDispatch.Checked = true;
                kryptonCheckSet1_CheckedButtonChanged(null, null);

                NeedSplash = true;

                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;
            }
            else
            {
                Infinium.LightMessageBox.Show(ref TopForm, false,
                        "Не найдено в отгрузке",
                        "Поиск отгрузки");
                return;
            }
        }

        /// <summary>
        /// Приложения к отгрузке для корпусной мебели
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void kryptonContextMenuItem20_Click(object sender, EventArgs e)
        {
            bool NeedProfilList = true;
            bool NeedTPSList = true;
            int ClientID = Convert.ToInt32(dgvCabFurMegaOrders.SelectedRows[0].Cells["ClientID"].Value);
            int MegaOrderID = Convert.ToInt32(dgvCabFurMegaOrders.SelectedRows[0].Cells["MegaOrderID"].Value);

            int[] MegaOrders = new int[dgvCabFurMegaOrders.SelectedRows.Count];
            for (int i = 0; i < dgvCabFurMegaOrders.SelectedRows.Count; i++)
                MegaOrders[i] = Convert.ToInt32(dgvCabFurMegaOrders.SelectedRows[i].Cells["MegaOrderID"].Value);
            if (!cabFurAssembleManager.HasPackages(MegaOrderID))
            {
                InfiniumTips.ShowTip(this, 50, 85, "Заказ пуст", 1700);
                return;
            }

            Thread T1 = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Создание документа Excel.\r\nПодождите..."); });
            T1.Start();

            while (!SplashWindow.bSmallCreated) ;

            object CreationDateTime = dgvCabFurMegaOrders.SelectedRows[0].Cells["OrderDate"].Value;
            object PrepareDispDateTime = dgvCabFurAssembleDates.SelectedRows[0].Cells["PrepareDateTime"].Value;
            string PackagesReportName = string.Empty;

            cabFurAssembleReport.GetDispatchInfo(ref CreationDateTime, ref PrepareDispDateTime);
            cabFurAssembleReport.CurrentClientID = ClientID;
            cabFurAssembleReport.CurrentMegaOrders = MegaOrders;
            cabFurAssembleReport.FillPackagesByMegaOrders();
            cabFurAssembleReport.CreateCabFurReport(NeedProfilList, NeedTPSList, true, ref PackagesReportName);

            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
        }

        private void kryptonContextMenuItem21_Click(object sender, EventArgs e)
        {
            if (dgvCabFurMegaOrders.SelectedRows.Count == 0)
                return;

            int MegaOrderID = 0;
            int ClientID = 0;
            int OrderNumber = 0;
            if (dgvCabFurMegaOrders.SelectedRows.Count != 0 && dgvCabFurMegaOrders.SelectedRows[0].Cells["ClientID"].Value != DBNull.Value)
                ClientID = Convert.ToInt32(dgvCabFurMegaOrders.SelectedRows[0].Cells["ClientID"].Value);
            if (dgvCabFurMegaOrders.SelectedRows.Count != 0 && dgvCabFurMegaOrders.SelectedRows[0].Cells["MegaOrderID"].Value != DBNull.Value)
                MegaOrderID = Convert.ToInt32(dgvCabFurMegaOrders.SelectedRows[0].Cells["MegaOrderID"].Value);
            if (dgvCabFurMegaOrders.SelectedRows.Count != 0 && dgvCabFurMegaOrders.SelectedRows[0].Cells["OrderNumber"].Value != DBNull.Value)
                OrderNumber = Convert.ToInt32(dgvCabFurMegaOrders.SelectedRows[0].Cells["OrderNumber"].Value);

            object CreationDateTime = dgvCabFurMegaOrders.SelectedRows[0].Cells["OrderDate"].Value;
            object PrepareDispDateTime = dgvCabFurAssembleDates.SelectedRows[0].Cells["PrepareDateTime"].Value;

            Thread T = new Thread(delegate () { SplashWindow.CreateSplash(); });
            T.Start();

            while (!SplashForm.bCreated) ;

            CabFurAssembleForm cabFurAssembleForm = new CabFurAssembleForm(this, cabFurAssembleReport, ClientID, MegaOrderID, OrderNumber, CreationDateTime, PrepareDispDateTime);

            TopForm = cabFurAssembleForm;

            cabFurAssembleForm.ShowDialog();

            cabFurAssembleForm.Close();
            cabFurAssembleForm.Dispose();
            cabFurAssembleManager.UpdateAllCabFurniturePackages();
            UpdateCabFurDispatchDate();
            dgvCabFurDates_SelectionChanged(null, null);

            TopForm = null;
        }

        private void kryptonContextMenuItem23_Click(object sender, EventArgs e)
        {
            if (dgvCabFurMegaOrders.SelectedRows.Count == 0)
                return;
            int MegaOrderID = Convert.ToInt32(dgvCabFurMegaOrders.SelectedRows[0].Cells["MegaOrderID"].Value);

            MarketingExpeditionManager.MoveToMegaOrder(MegaOrderID);
            cbtnExpedition.Checked = true;
            kryptonCheckSet1_CheckedButtonChanged(null, null);
        }

        private void btnShowCabFurAssemble_Click(object sender, EventArgs e)
        {
            Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
            T.Start();

            while (!SplashWindow.bSmallCreated) ;

            NeedSplash = false;
            MarketingExpeditionManager.SearchCabFurAssembled(cabFurAssembleManager);

            NeedSplash = true;

            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
        }

        private void btnShowAllOrders_Click(object sender, EventArgs e)
        {
            if (MarketingExpeditionManager != null)
                FilterOrders();
        }
        
        private void kryptonContextMenuItem22_Click(object sender, EventArgs e)
        {
            bool InMutualSettlement = Convert.ToBoolean(dgvDispatch.SelectedRows[0].Cells["InMutualSettlement"].Value);
            //int DispatchID = Convert.ToInt32(dgvDispatch.SelectedRows[0].Cells["DispatchID"].Value);
            DateTime DispatchDate = Convert.ToDateTime(dgvDispatchDates.SelectedRows[0].Cells["PrepareDispatchDateTime"].Value);
            int ClientID = Convert.ToInt32(dgvDispatch.SelectedRows[0].Cells["ClientID"].Value);
            string ClientName = dgvDispatch.SelectedRows[0].Cells["ClientName"].Value.ToString();
            int[] Dispatches = new int[dgvDispatch.SelectedRows.Count];
            for (int i = 0; i < dgvDispatch.SelectedRows.Count; i++)
                Dispatches[i] = Convert.ToInt32(dgvDispatch.SelectedRows[i].Cells["DispatchID"].Value);

            int ClientCountryID = OrdersManager.GetClientCountry(ClientID);
            int iDiscountPaymentConditionID = 0;
            string ClientLogin = OrdersManager.GetClientLogin(ClientID);

            decimal DiscountPaymentCondition = 0;
            decimal TotalProfil = 0;
            decimal TotalTPS = 0;

            string ProfilDBFName = string.Empty;
            string TPSDBFName = string.Empty;
            string ReportName = string.Empty;
            InMutualSettlement = false;
            _colorInvoiceReportToDbf.ClearTempPackages();
            if (!InMutualSettlement)
            {
                Modules.Marketing.MutualSettlements.MutualSettlements MutualSettlementsManager = new Modules.Marketing.MutualSettlements.MutualSettlements();
                MutualSettlementsManager.Initialize();

                if (ClientID == 145)
                {
                    bool HasOrdinaryOrders = MarketingDispatchManager.HasOrders(Dispatches, false);
                    bool HasSampleOrders = MarketingDispatchManager.HasOrders(Dispatches, true);

                    if (HasSampleOrders)
                    {
                        PhantomForm PhantomForm = new Infinium.PhantomForm();
                        PhantomForm.Show();

                        int Result = 1;
                        string SaveFilePath = string.Empty;
                        SaveDBFReportMenu = new Infinium.SaveDBFReportMenu(this, ReportName);

                        TopForm = SaveDBFReportMenu;
                        SaveDBFReportMenu.ShowDialog();

                        PhantomForm.Close();
                        Result = SaveDBFReportMenu.Result;
                        SaveFilePath = SaveDBFReportMenu.SaveFilePath;
                        PhantomForm.Dispose();
                        SaveDBFReportMenu.Dispose();
                        TopForm = null;

                        if (Result == 3)
                            return;

                        int ClientDispatchID = -1;
                        int MutualSettlementID1 = -1;
                        int MutualSettlementID2 = -1;
                        bool ProfilVerify = true;
                        bool TPSVerify = true;

                        ClientName = ClientName.Replace("\"", " ");
                        ClientName = ClientName.Replace("*", " ");
                        ClientName = ClientName.Replace("|", " ");
                        ClientName = ClientName.Replace(@"\", " ");
                        ClientName = ClientName.Replace(":", " ");
                        ClientName = ClientName.Replace("<", " ");
                        ClientName = ClientName.Replace(">", " ");
                        ClientName = ClientName.Replace("?", " ");
                        ClientName = ClientName.Replace("/", " ");

                        string CurrentMonthName = DateTime.Now.ToString("dd-MM-yyyy");
                        string FilePath = Security.DBFSaveFilePath;
                        if (Security.DBFSaveFilePath.Length == 0)
                            FilePath = Environment.GetFolderPath(Environment.SpecialFolder.Personal) + "/Infinium/1C data/" + CurrentMonthName;

                        Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Создание документа Excel.\r\nПодождите..."); });
                        T.Start();

                        while (!SplashWindow.bSmallCreated) ;
                        NeedSplash = false;

                        for (int j = 0; j < Dispatches.Count(); j++)
                        {
                            int[] MainOrders = MarketingDispatchManager.GetMainOrdersInDispatch(Dispatches[j]);

                            DataTable ReturnDT = MarketingDispatchManager.GetMegaOrdersInDispatch(Dispatches[j]);
                            ArrayList OrderNumbers = new ArrayList();
                            for (int i = 0; i < ReturnDT.Rows.Count; i++)
                                OrderNumbers.Add(Convert.ToInt32(ReturnDT.Rows[i]["OrderNumber"]));

                            int CurrencyTypeID = 1;
                            if (ReturnDT.Rows.Count > 0)
                                CurrencyTypeID = Convert.ToInt32(ReturnDT.Rows[0]["CurrencyTypeID"]);

                            MutualSettlementID1 = MutualSettlementsManager.FindMutualSettlementIDByOrderNumber(ClientID, 1, OrderNumbers.OfType<Int32>().ToArray(), true);
                            if (MutualSettlementID1 != 0)
                            {
                                decimal DispatchSum = MutualSettlementsManager.GetDispatchSum(MutualSettlementID1);
                                ProfilVerify = MutualSettlementsManager.VerifyTransaction(MutualSettlementID1, DispatchSum, ref DiscountPaymentCondition);
                                for (int i = 0; i < ReturnDT.Rows.Count; i++)
                                {
                                    int MegaOrderID = Convert.ToInt32(ReturnDT.Rows[i]["MegaOrderID"]);

                                    MegaOrderInfo m = MarketingDispatchManager.GetMegaOrders(MegaOrderID);
                                    iDiscountPaymentConditionID = ProfilReCalcMegaOrder(ClientID, CurrencyTypeID, DiscountPaymentCondition, ProfilVerify, true, MegaOrderID, m);
                                    OrdersManager.FixOrderEvent(MegaOrderID, "Выписать накладную. Пересчитан заказ Профиль");
                                }
                            }
                            MutualSettlementID2 = 0;
                            MutualSettlementID2 = MutualSettlementsManager.FindMutualSettlementIDByOrderNumber(ClientID, 2, OrderNumbers.OfType<Int32>().ToArray(), true);
                            if (MutualSettlementID2 != 0)
                            {
                                decimal DispatchSum = MutualSettlementsManager.GetDispatchSum(MutualSettlementID2);
                                TPSVerify = MutualSettlementsManager.VerifyTransaction(MutualSettlementID2, DispatchSum, ref DiscountPaymentCondition);
                                for (int i = 0; i < ReturnDT.Rows.Count; i++)
                                {
                                    int MegaOrderID = Convert.ToInt32(ReturnDT.Rows[i]["MegaOrderID"]);
                                    MegaOrderInfo m = MarketingDispatchManager.GetMegaOrders(MegaOrderID);
                                    iDiscountPaymentConditionID = TPSReCalcMegaOrder(ClientID, CurrencyTypeID, DiscountPaymentCondition, TPSVerify, true, MegaOrderID, m);
                                    OrdersManager.FixOrderEvent(MegaOrderID, "Выписать накладную. Пересчитан заказ ТПС");
                                }
                            }

                            HSSFWorkbook hssfworkbook = new HSSFWorkbook();

                            ////create a entry of DocumentSummaryInformation
                            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
                            dsi.Company = "NPOI Team";
                            hssfworkbook.DocumentSummaryInformation = dsi;

                            ////create a entry of SummaryInformation
                            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
                            si.Subject = "NPOI SDK Example";
                            hssfworkbook.SummaryInformation = si;

                            ReportName = ClientLogin + " №" + Dispatches[j].ToString();
                            _colorInvoiceReportToDbf.CreateReport(ref hssfworkbook, new int[1] { Dispatches[j] }, MainOrders, ClientID, ClientName, ref TotalProfil, ref TotalTPS, ProfilVerify, TPSVerify, iDiscountPaymentConditionID, true);
                            _colorInvoiceReportToDbf.SaveDBF(FilePath, ReportName, true, ref ProfilDBFName, ref TPSDBFName);
                            DetailsReport.CreateReport(ref hssfworkbook, new int[1] { Dispatches[j] }, OrderNumbers.OfType<Int32>().ToArray(), MainOrders, ClientID, ClientName, ProfilVerify, TPSVerify, iDiscountPaymentConditionID, true);

                            if (ClientID == 609)
                            {
                                DetailsReport.ReportModuleOnline(ref hssfworkbook, Dispatches);
                            }
                            FileInfo file = new FileInfo(FilePath + @"\" + ReportName + ".xls");
                            int x = 1;
                            while (file.Exists == true)
                            {
                                file = new FileInfo(FilePath + @"\" + ReportName + "(" + x++ + ").xls");
                            }

                            FileStream NewFile = new FileStream(file.FullName, FileMode.Create);
                            hssfworkbook.Write(NewFile);
                            NewFile.Close();

                            ReportName = file.FullName;

                            if (Result == 1)
                            {
                                if (TotalProfil > 0)
                                {
                                    if (ClientCountryID == 1)
                                        TotalProfil *= 1.2m;
                                    ClientDispatchID = MutualSettlementsManager.AddClientDispatch(MutualSettlementID1, CurrencyTypeID, TotalProfil);
                                    var fileInfo = new System.IO.FileInfo(ReportName);
                                    MutualSettlementsManager.AttachDispatchDocument(Path.GetFileNameWithoutExtension(ReportName),
                                        fileInfo.Extension, ReportName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.DispatchExcel, ClientDispatchID);
                                    fileInfo = new System.IO.FileInfo(ProfilDBFName);
                                    MutualSettlementsManager.AttachDispatchDocument(Path.GetFileNameWithoutExtension(ProfilDBFName),
                                        fileInfo.Extension, ProfilDBFName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.DispatchDbf, ClientDispatchID);
                                }
                                if (TotalTPS > 0)
                                {
                                    if (ClientCountryID == 1)
                                        TotalTPS *= 1.2m;
                                    ClientDispatchID = MutualSettlementsManager.AddClientDispatch(MutualSettlementID2, CurrencyTypeID, TotalTPS);
                                    var fileInfo = new System.IO.FileInfo(ReportName);
                                    MutualSettlementsManager.AttachDispatchDocument(Path.GetFileNameWithoutExtension(ReportName),
                                        fileInfo.Extension, ReportName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.DispatchExcel, ClientDispatchID);
                                    fileInfo = new System.IO.FileInfo(TPSDBFName);
                                    MutualSettlementsManager.AttachDispatchDocument(Path.GetFileNameWithoutExtension(TPSDBFName),
                                        fileInfo.Extension, TPSDBFName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.DispatchDbf, ClientDispatchID);
                                }
                            }
                            MarketingDispatchManager.IncludeInMutualSettlement(Dispatches[j], MutualSettlementID1, MutualSettlementID2);
                        }

                        MutualSettlementsManager.SaveMutualSettlements();
                        MutualSettlementsManager.SaveMutualSettlementOrders();

                        UpdateDispatchDate();
                        MarketingDispatchManager.MoveToDispatchDate(DispatchDate);
                        MarketingDispatchManager.MoveToDispatch(Dispatches[0]);

                        ColorInvoiceClient(ClientID, ClientName, Dispatches, iDiscountPaymentConditionID, ClientLogin, ProfilVerify, TPSVerify, FilePath);

                        NeedSplash = true;
                        while (SplashWindow.bSmallCreated)
                            SmallWaitForm.CloseS = true;
                    }
                    if (HasOrdinaryOrders)
                    {
                        PhantomForm PhantomForm = new Infinium.PhantomForm();
                        PhantomForm.Show();

                        int Result = 1;
                        string SaveFilePath = string.Empty;
                        SaveDBFReportMenu = new Infinium.SaveDBFReportMenu(this, ReportName);

                        TopForm = SaveDBFReportMenu;
                        SaveDBFReportMenu.ShowDialog();

                        PhantomForm.Close();
                        Result = SaveDBFReportMenu.Result;
                        SaveFilePath = SaveDBFReportMenu.SaveFilePath;
                        PhantomForm.Dispose();
                        SaveDBFReportMenu.Dispose();
                        TopForm = null;

                        if (Result == 3)
                            return;

                        int ClientDispatchID = -1;
                        int MutualSettlementID1 = -1;
                        int MutualSettlementID2 = -1;
                        bool ProfilVerify = true;
                        bool TPSVerify = true;

                        ClientName = ClientName.Replace("\"", " ");
                        ClientName = ClientName.Replace("*", " ");
                        ClientName = ClientName.Replace("|", " ");
                        ClientName = ClientName.Replace(@"\", " ");
                        ClientName = ClientName.Replace(":", " ");
                        ClientName = ClientName.Replace("<", " ");
                        ClientName = ClientName.Replace(">", " ");
                        ClientName = ClientName.Replace("?", " ");
                        ClientName = ClientName.Replace("/", " ");

                        string CurrentMonthName = DateTime.Now.ToString("dd-MM-yyyy");
                        string FilePath = Security.DBFSaveFilePath;
                        if (Security.DBFSaveFilePath.Length == 0)
                            FilePath = Environment.GetFolderPath(Environment.SpecialFolder.Personal) + "/Infinium/1C data/" + CurrentMonthName;

                        Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Создание документа Excel.\r\nПодождите..."); });
                        T.Start();

                        while (!SplashWindow.bSmallCreated) ;
                        NeedSplash = false;

                        for (int j = 0; j < Dispatches.Count(); j++)
                        {
                            int[] MainOrders = MarketingDispatchManager.GetMainOrdersInDispatch(Dispatches[j]);

                            DataTable ReturnDT = MarketingDispatchManager.GetMegaOrdersInDispatch(Dispatches[j]);
                            ArrayList OrderNumbers = new ArrayList();
                            for (int i = 0; i < ReturnDT.Rows.Count; i++)
                                OrderNumbers.Add(Convert.ToInt32(ReturnDT.Rows[i]["OrderNumber"]));

                            int CurrencyTypeID = 1;
                            if (ReturnDT.Rows.Count > 0)
                                CurrencyTypeID = Convert.ToInt32(ReturnDT.Rows[0]["CurrencyTypeID"]);

                            MutualSettlementID1 = MutualSettlementsManager.FindMutualSettlementIDByOrderNumber(ClientID, 1, OrderNumbers.OfType<Int32>().ToArray(), false);
                            if (MutualSettlementID1 != 0)
                            {
                                decimal DispatchSum = MutualSettlementsManager.GetDispatchSum(MutualSettlementID1);
                                ProfilVerify = MutualSettlementsManager.VerifyTransaction(MutualSettlementID1, DispatchSum, ref DiscountPaymentCondition);
                                for (int i = 0; i < ReturnDT.Rows.Count; i++)
                                {
                                    int MegaOrderID = Convert.ToInt32(ReturnDT.Rows[i]["MegaOrderID"]);

                                    MegaOrderInfo m = MarketingDispatchManager.GetMegaOrders(MegaOrderID);
                                    iDiscountPaymentConditionID = ProfilReCalcMegaOrder(ClientID, CurrencyTypeID, DiscountPaymentCondition, ProfilVerify, false, MegaOrderID, m);
                                    OrdersManager.FixOrderEvent(MegaOrderID, "Выписать накладную. Пересчитан заказ Профиль");
                                }
                            }
                            MutualSettlementID2 = 0;
                            MutualSettlementID2 = MutualSettlementsManager.FindMutualSettlementIDByOrderNumber(ClientID, 2, OrderNumbers.OfType<Int32>().ToArray(), false);
                            if (MutualSettlementID2 != 0)
                            {
                                decimal DispatchSum = MutualSettlementsManager.GetDispatchSum(MutualSettlementID2);
                                TPSVerify = MutualSettlementsManager.VerifyTransaction(MutualSettlementID2, DispatchSum, ref DiscountPaymentCondition);
                                for (int i = 0; i < ReturnDT.Rows.Count; i++)
                                {
                                    int MegaOrderID = Convert.ToInt32(ReturnDT.Rows[i]["MegaOrderID"]);
                                    MegaOrderInfo m = MarketingDispatchManager.GetMegaOrders(MegaOrderID);
                                    iDiscountPaymentConditionID = TPSReCalcMegaOrder(ClientID, CurrencyTypeID, DiscountPaymentCondition, TPSVerify, false, MegaOrderID, m);
                                    OrdersManager.FixOrderEvent(MegaOrderID, "Выписать накладную. Пересчитан заказ ТПС");
                                }
                            }

                            HSSFWorkbook hssfworkbook = new HSSFWorkbook();

                            ////create a entry of DocumentSummaryInformation
                            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
                            dsi.Company = "NPOI Team";
                            hssfworkbook.DocumentSummaryInformation = dsi;

                            ////create a entry of SummaryInformation
                            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
                            si.Subject = "NPOI SDK Example";
                            hssfworkbook.SummaryInformation = si;

                            ReportName = ClientLogin + " №" + Dispatches[j].ToString();
                            _colorInvoiceReportToDbf.CreateReport(ref hssfworkbook, new int[1] { Dispatches[j] }, MainOrders, ClientID, ClientName, ref TotalProfil, ref TotalTPS, ProfilVerify, TPSVerify, iDiscountPaymentConditionID, false);
                            _colorInvoiceReportToDbf.SaveDBF(FilePath, ReportName, false, ref ProfilDBFName, ref TPSDBFName);
                            DetailsReport.CreateReport(ref hssfworkbook, new int[1] { Dispatches[j] }, OrderNumbers.OfType<Int32>().ToArray(), MainOrders, ClientID, ClientName, ProfilVerify, TPSVerify, iDiscountPaymentConditionID, false);

                            if (ClientID == 609)
                            {
                                DetailsReport.ReportModuleOnline(ref hssfworkbook, Dispatches);
                            }
                            FileInfo file = new FileInfo(FilePath + @"\" + ReportName + ".xls");
                            int x = 1;
                            while (file.Exists == true)
                            {
                                file = new FileInfo(FilePath + @"\" + ReportName + "(" + x++ + ").xls");
                            }

                            FileStream NewFile = new FileStream(file.FullName, FileMode.Create);
                            hssfworkbook.Write(NewFile);
                            NewFile.Close();

                            ReportName = file.FullName;

                            if (Result == 1)
                            {
                                if (TotalProfil > 0)
                                {
                                    if (ClientCountryID == 1)
                                        TotalProfil *= 1.2m;
                                    ClientDispatchID = MutualSettlementsManager.AddClientDispatch(MutualSettlementID1, CurrencyTypeID, TotalProfil);
                                    var fileInfo = new System.IO.FileInfo(ReportName);
                                    MutualSettlementsManager.AttachDispatchDocument(Path.GetFileNameWithoutExtension(ReportName),
                                        fileInfo.Extension, ReportName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.DispatchExcel, ClientDispatchID);
                                    fileInfo = new System.IO.FileInfo(ProfilDBFName);
                                    MutualSettlementsManager.AttachDispatchDocument(Path.GetFileNameWithoutExtension(ProfilDBFName),
                                        fileInfo.Extension, ProfilDBFName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.DispatchDbf, ClientDispatchID);
                                }
                                if (TotalTPS > 0)
                                {
                                    if (ClientCountryID == 1)
                                        TotalTPS *= 1.2m;
                                    ClientDispatchID = MutualSettlementsManager.AddClientDispatch(MutualSettlementID2, CurrencyTypeID, TotalTPS);
                                    var fileInfo = new System.IO.FileInfo(ReportName);
                                    MutualSettlementsManager.AttachDispatchDocument(Path.GetFileNameWithoutExtension(ReportName),
                                        fileInfo.Extension, ReportName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.DispatchExcel, ClientDispatchID);
                                    fileInfo = new System.IO.FileInfo(TPSDBFName);
                                    MutualSettlementsManager.AttachDispatchDocument(Path.GetFileNameWithoutExtension(TPSDBFName),
                                        fileInfo.Extension, TPSDBFName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.DispatchDbf, ClientDispatchID);
                                }
                            }
                            MarketingDispatchManager.IncludeInMutualSettlement(Dispatches[j], MutualSettlementID1, MutualSettlementID2);
                        }

                        MutualSettlementsManager.SaveMutualSettlements();
                        MutualSettlementsManager.SaveMutualSettlementOrders();

                        UpdateDispatchDate();
                        MarketingDispatchManager.MoveToDispatchDate(DispatchDate);
                        MarketingDispatchManager.MoveToDispatch(Dispatches[0]);

                        ColorInvoiceClient(ClientID, ClientName, Dispatches, iDiscountPaymentConditionID, ClientLogin, ProfilVerify, TPSVerify, FilePath);

                        NeedSplash = true;
                        while (SplashWindow.bSmallCreated)
                            SmallWaitForm.CloseS = true;
                    }
                }
                else
                {
                    PhantomForm PhantomForm = new Infinium.PhantomForm();
                    PhantomForm.Show();

                    int Result = 1;
                    string SaveFilePath = string.Empty;
                    SaveDBFReportMenu = new Infinium.SaveDBFReportMenu(this, ReportName);

                    TopForm = SaveDBFReportMenu;
                    SaveDBFReportMenu.ShowDialog();

                    PhantomForm.Close();
                    Result = SaveDBFReportMenu.Result;
                    SaveFilePath = SaveDBFReportMenu.SaveFilePath;
                    PhantomForm.Dispose();
                    SaveDBFReportMenu.Dispose();
                    TopForm = null;

                    if (Result == 3)
                        return;

                    int ClientDispatchID = -1;
                    int MutualSettlementID1 = -1;
                    int MutualSettlementID2 = -1;
                    bool ProfilVerify = true;
                    bool TPSVerify = true;

                    ClientName = ClientName.Replace("\"", " ");
                    ClientName = ClientName.Replace("*", " ");
                    ClientName = ClientName.Replace("|", " ");
                    ClientName = ClientName.Replace(@"\", " ");
                    ClientName = ClientName.Replace(":", " ");
                    ClientName = ClientName.Replace("<", " ");
                    ClientName = ClientName.Replace(">", " ");
                    ClientName = ClientName.Replace("?", " ");
                    ClientName = ClientName.Replace("/", " ");

                    string CurrentMonthName = DateTime.Now.ToString("dd-MM-yyyy");
                    string FilePath = Security.DBFSaveFilePath;
                    if (Security.DBFSaveFilePath.Length == 0)
                        FilePath = Environment.GetFolderPath(Environment.SpecialFolder.Personal) + "/Infinium/1C data/" + CurrentMonthName;

                    Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Создание документа Excel.\r\nПодождите..."); });
                    T.Start();

                    while (!SplashWindow.bSmallCreated) ;
                    NeedSplash = false;

                    for (int j = 0; j < Dispatches.Count(); j++)
                    {
                        int[] MainOrders = MarketingDispatchManager.GetMainOrdersInDispatch(Dispatches[j]);

                        DataTable ReturnDT = MarketingDispatchManager.GetMegaOrdersInDispatch(Dispatches[j]);
                        ArrayList OrderNumbers = new ArrayList();
                        for (int i = 0; i < ReturnDT.Rows.Count; i++)
                            OrderNumbers.Add(Convert.ToInt32(ReturnDT.Rows[i]["OrderNumber"]));

                        int CurrencyTypeID = 1;
                        int FactoryID = 1;

                        if (ReturnDT.Rows.Count > 0)
                        {
                            CurrencyTypeID = Convert.ToInt32(ReturnDT.Rows[0]["CurrencyTypeID"]);
                            FactoryID = Convert.ToInt32(ReturnDT.Rows[0]["FactoryID"]);
                        }

                        MutualSettlementID1 = MutualSettlementsManager.FindMutualSettlementIDByOrderNumber(ClientID, 1, OrderNumbers.OfType<Int32>().ToArray());
                        if (MutualSettlementID1 != 0)
                        {
                            decimal DispatchSum = MutualSettlementsManager.GetDispatchSum(MutualSettlementID1);
                            ProfilVerify = MutualSettlementsManager.VerifyTransaction(MutualSettlementID1, DispatchSum, ref DiscountPaymentCondition);
                            for (int i = 0; i < ReturnDT.Rows.Count; i++)
                            {
                                int MegaOrderID = Convert.ToInt32(ReturnDT.Rows[i]["MegaOrderID"]);

                                MegaOrderInfo m = MarketingDispatchManager.GetMegaOrders(MegaOrderID);
                                iDiscountPaymentConditionID = ProfilReCalcMegaOrder(ClientID, CurrencyTypeID, DiscountPaymentCondition, ProfilVerify, MegaOrderID, m);
                                OrdersManager.FixOrderEvent(MegaOrderID, "Выписать накладную. Пересчитан заказ Профиль");
                            }
                        }
                        MutualSettlementID2 = 0;
                        MutualSettlementID2 = MutualSettlementsManager.FindMutualSettlementIDByOrderNumber(ClientID, 2, OrderNumbers.OfType<Int32>().ToArray());
                        if (MutualSettlementID2 != 0)
                        {
                            decimal DispatchSum = MutualSettlementsManager.GetDispatchSum(MutualSettlementID2);
                            TPSVerify = MutualSettlementsManager.VerifyTransaction(MutualSettlementID2, DispatchSum, ref DiscountPaymentCondition);
                            for (int i = 0; i < ReturnDT.Rows.Count; i++)
                            {
                                int MegaOrderID = Convert.ToInt32(ReturnDT.Rows[i]["MegaOrderID"]);
                                MegaOrderInfo m = MarketingDispatchManager.GetMegaOrders(MegaOrderID);
                                iDiscountPaymentConditionID = TPSReCalcMegaOrder(ClientID, CurrencyTypeID, DiscountPaymentCondition, TPSVerify, MegaOrderID, m);
                                OrdersManager.FixOrderEvent(MegaOrderID, "Выписать накладную. Пересчитан заказ ТПС");
                            }
                        }

                        HSSFWorkbook hssfworkbook = new HSSFWorkbook();

                        ////create a entry of DocumentSummaryInformation
                        DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
                        dsi.Company = "NPOI Team";
                        hssfworkbook.DocumentSummaryInformation = dsi;

                        ////create a entry of SummaryInformation
                        SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
                        si.Subject = "NPOI SDK Example";
                        hssfworkbook.SummaryInformation = si;

                        ReportName = ClientLogin + " №" + Dispatches[j].ToString();

                        _colorInvoiceReportToDbf.CreateReport(ref hssfworkbook, new int[1] { Dispatches[j] }, MainOrders, ClientID, ClientName, ref TotalProfil, ref TotalTPS, ProfilVerify, TPSVerify, iDiscountPaymentConditionID);
                        _colorInvoiceReportToDbf.SaveDBF(FilePath, ReportName, ref ProfilDBFName, ref TPSDBFName);

                        DetailsReport.CreateReport(ref hssfworkbook, new int[1] { Dispatches[j] }, OrderNumbers.OfType<Int32>().ToArray(), MainOrders, ClientID, ClientName, ProfilVerify, TPSVerify, iDiscountPaymentConditionID);

                        if (ClientID == 609)
                        {
                            DetailsReport.ReportModuleOnline(ref hssfworkbook, Dispatches);
                        }
                        FileInfo file = new FileInfo(FilePath + @"\" + ReportName + ".xls");
                        int x = 1;
                        while (file.Exists == true)
                        {
                            file = new FileInfo(FilePath + @"\" + ReportName + "(" + x++ + ").xls");
                        }

                        FileStream NewFile = new FileStream(file.FullName, FileMode.Create);
                        hssfworkbook.Write(NewFile);
                        NewFile.Close();

                        //System.Diagnostics.Process.Start(file.FullName);

                        ReportName = file.FullName;

                        if (Result == 1)
                        {
                            if (TotalProfil > 0)
                            {
                                if (ClientCountryID == 1)
                                    TotalProfil *= 1.2m;
                                ClientDispatchID = MutualSettlementsManager.AddClientDispatch(MutualSettlementID1, CurrencyTypeID, TotalProfil);
                                var fileInfo = new System.IO.FileInfo(ReportName);
                                MutualSettlementsManager.AttachDispatchDocument(Path.GetFileNameWithoutExtension(ReportName),
                                    fileInfo.Extension, ReportName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.DispatchExcel, ClientDispatchID);
                                fileInfo = new System.IO.FileInfo(ProfilDBFName);
                                MutualSettlementsManager.AttachDispatchDocument(Path.GetFileNameWithoutExtension(ProfilDBFName),
                                    fileInfo.Extension, ProfilDBFName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.DispatchDbf, ClientDispatchID);
                            }
                            if (TotalTPS > 0)
                            {
                                if (ClientCountryID == 1)
                                    TotalTPS *= 1.2m;
                                ClientDispatchID = MutualSettlementsManager.AddClientDispatch(MutualSettlementID2, CurrencyTypeID, TotalTPS);
                                var fileInfo = new System.IO.FileInfo(ReportName);
                                MutualSettlementsManager.AttachDispatchDocument(Path.GetFileNameWithoutExtension(ReportName),
                                    fileInfo.Extension, ReportName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.DispatchExcel, ClientDispatchID);
                                fileInfo = new System.IO.FileInfo(TPSDBFName);
                                MutualSettlementsManager.AttachDispatchDocument(Path.GetFileNameWithoutExtension(TPSDBFName),
                                    fileInfo.Extension, TPSDBFName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.DispatchDbf, ClientDispatchID);
                            }
                        }
                        MarketingDispatchManager.IncludeInMutualSettlement(Dispatches[j], MutualSettlementID1, MutualSettlementID2);
                    }

                    MutualSettlementsManager.SaveMutualSettlements();
                    MutualSettlementsManager.SaveMutualSettlementOrders();

                    UpdateDispatchDate();
                    MarketingDispatchManager.MoveToDispatchDate(DispatchDate);
                    MarketingDispatchManager.MoveToDispatch(Dispatches[0]);

                    ColorInvoiceClient(ClientID, ClientName, Dispatches, iDiscountPaymentConditionID, ClientLogin, ProfilVerify, TPSVerify, FilePath);

                    NeedSplash = true;
                    while (SplashWindow.bSmallCreated)
                        SmallWaitForm.CloseS = true;
                }
            }
        }

        private void kryptonContextMenuItem24_Click(object sender, EventArgs e)
        {
            bool InMutualSettlement = Convert.ToBoolean(dgvDispatch.SelectedRows[0].Cells["InMutualSettlement"].Value);
            //int DispatchID = Convert.ToInt32(dgvDispatch.SelectedRows[0].Cells["DispatchID"].Value);
            DateTime DispatchDate = Convert.ToDateTime(dgvDispatchDates.SelectedRows[0].Cells["PrepareDispatchDateTime"].Value);
            int ClientID = Convert.ToInt32(dgvDispatch.SelectedRows[0].Cells["ClientID"].Value);
            string ClientName = dgvDispatch.SelectedRows[0].Cells["ClientName"].Value.ToString();
            int[] Dispatches = new int[dgvDispatch.SelectedRows.Count];
            for (int i = 0; i < dgvDispatch.SelectedRows.Count; i++)
                Dispatches[i] = Convert.ToInt32(dgvDispatch.SelectedRows[i].Cells["DispatchID"].Value);

            int ClientCountryID = OrdersManager.GetClientCountry(ClientID);
            int iDiscountPaymentConditionID = 0;
            string ClientLogin = OrdersManager.GetClientLogin(ClientID);

            decimal DiscountPaymentCondition = 0;
            decimal TotalProfil = 0;
            decimal TotalTPS = 0;

            string ProfilDBFName = string.Empty;
            string TPSDBFName = string.Empty;
            string ReportName = string.Empty;
            InMutualSettlement = false;
            _notesInvoiceReportToDbf.ClearTempPackages();
            if (!InMutualSettlement)
            {
                Modules.Marketing.MutualSettlements.MutualSettlements MutualSettlementsManager = new Modules.Marketing.MutualSettlements.MutualSettlements();
                MutualSettlementsManager.Initialize();

                if (ClientID == 145)
                {
                    bool HasOrdinaryOrders = MarketingDispatchManager.HasOrders(Dispatches, false);
                    bool HasSampleOrders = MarketingDispatchManager.HasOrders(Dispatches, true);

                    if (HasSampleOrders)
                    {
                        PhantomForm PhantomForm = new Infinium.PhantomForm();
                        PhantomForm.Show();

                        int Result = 1;
                        string SaveFilePath = string.Empty;
                        SaveDBFReportMenu = new Infinium.SaveDBFReportMenu(this, ReportName);

                        TopForm = SaveDBFReportMenu;
                        SaveDBFReportMenu.ShowDialog();

                        PhantomForm.Close();
                        Result = SaveDBFReportMenu.Result;
                        SaveFilePath = SaveDBFReportMenu.SaveFilePath;
                        PhantomForm.Dispose();
                        SaveDBFReportMenu.Dispose();
                        TopForm = null;

                        if (Result == 3)
                            return;

                        int ClientDispatchID = -1;
                        int MutualSettlementID1 = -1;
                        int MutualSettlementID2 = -1;
                        bool ProfilVerify = true;
                        bool TPSVerify = true;

                        ClientName = ClientName.Replace("\"", " ");
                        ClientName = ClientName.Replace("*", " ");
                        ClientName = ClientName.Replace("|", " ");
                        ClientName = ClientName.Replace(@"\", " ");
                        ClientName = ClientName.Replace(":", " ");
                        ClientName = ClientName.Replace("<", " ");
                        ClientName = ClientName.Replace(">", " ");
                        ClientName = ClientName.Replace("?", " ");
                        ClientName = ClientName.Replace("/", " ");

                        string CurrentMonthName = DateTime.Now.ToString("dd-MM-yyyy");
                        string FilePath = Security.DBFSaveFilePath;
                        if (Security.DBFSaveFilePath.Length == 0)
                            FilePath = Environment.GetFolderPath(Environment.SpecialFolder.Personal) + "/Infinium/1C data/" + CurrentMonthName;

                        Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Создание документа Excel.\r\nПодождите..."); });
                        T.Start();

                        while (!SplashWindow.bSmallCreated) ;
                        NeedSplash = false;

                        for (int j = 0; j < Dispatches.Count(); j++)
                        {
                            int[] MainOrders = MarketingDispatchManager.GetMainOrdersInDispatch(Dispatches[j]);

                            DataTable ReturnDT = MarketingDispatchManager.GetMegaOrdersInDispatch(Dispatches[j]);
                            ArrayList OrderNumbers = new ArrayList();
                            for (int i = 0; i < ReturnDT.Rows.Count; i++)
                                OrderNumbers.Add(Convert.ToInt32(ReturnDT.Rows[i]["OrderNumber"]));

                            int CurrencyTypeID = 1;
                            if (ReturnDT.Rows.Count > 0)
                                CurrencyTypeID = Convert.ToInt32(ReturnDT.Rows[0]["CurrencyTypeID"]);

                            MutualSettlementID1 = MutualSettlementsManager.FindMutualSettlementIDByOrderNumber(ClientID, 1, OrderNumbers.OfType<Int32>().ToArray(), true);
                            if (MutualSettlementID1 != 0)
                            {
                                decimal DispatchSum = MutualSettlementsManager.GetDispatchSum(MutualSettlementID1);
                                ProfilVerify = MutualSettlementsManager.VerifyTransaction(MutualSettlementID1, DispatchSum, ref DiscountPaymentCondition);
                                for (int i = 0; i < ReturnDT.Rows.Count; i++)
                                {
                                    int MegaOrderID = Convert.ToInt32(ReturnDT.Rows[i]["MegaOrderID"]);

                                    MegaOrderInfo m = MarketingDispatchManager.GetMegaOrders(MegaOrderID);
                                    iDiscountPaymentConditionID = ProfilReCalcMegaOrder(ClientID, CurrencyTypeID, DiscountPaymentCondition, ProfilVerify, true, MegaOrderID, m);
                                    OrdersManager.FixOrderEvent(MegaOrderID, "Выписать накладную. Пересчитан заказ Профиль");
                                }
                            }
                            MutualSettlementID2 = 0;
                            MutualSettlementID2 = MutualSettlementsManager.FindMutualSettlementIDByOrderNumber(ClientID, 2, OrderNumbers.OfType<Int32>().ToArray(), true);
                            if (MutualSettlementID2 != 0)
                            {
                                decimal DispatchSum = MutualSettlementsManager.GetDispatchSum(MutualSettlementID2);
                                TPSVerify = MutualSettlementsManager.VerifyTransaction(MutualSettlementID2, DispatchSum, ref DiscountPaymentCondition);
                                for (int i = 0; i < ReturnDT.Rows.Count; i++)
                                {
                                    int MegaOrderID = Convert.ToInt32(ReturnDT.Rows[i]["MegaOrderID"]);
                                    MegaOrderInfo m = MarketingDispatchManager.GetMegaOrders(MegaOrderID);
                                    iDiscountPaymentConditionID = TPSReCalcMegaOrder(ClientID, CurrencyTypeID, DiscountPaymentCondition, TPSVerify, true, MegaOrderID, m);
                                    OrdersManager.FixOrderEvent(MegaOrderID, "Выписать накладную. Пересчитан заказ ТПС");
                                }
                            }

                            HSSFWorkbook hssfworkbook = new HSSFWorkbook();

                            ////create a entry of DocumentSummaryInformation
                            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
                            dsi.Company = "NPOI Team";
                            hssfworkbook.DocumentSummaryInformation = dsi;

                            ////create a entry of SummaryInformation
                            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
                            si.Subject = "NPOI SDK Example";
                            hssfworkbook.SummaryInformation = si;

                            ReportName = ClientLogin + " №" + Dispatches[j].ToString();
                            _notesInvoiceReportToDbf.CreateReport(ref hssfworkbook, new int[1] { Dispatches[j] }, MainOrders, ClientID, ClientName, ref TotalProfil, ref TotalTPS, ProfilVerify, TPSVerify, iDiscountPaymentConditionID, true);
                            _notesInvoiceReportToDbf.SaveDBF(FilePath, ReportName, true, ref ProfilDBFName, ref TPSDBFName);
                            DetailsReport.CreateReport(ref hssfworkbook, new int[1] { Dispatches[j] }, OrderNumbers.OfType<Int32>().ToArray(), MainOrders, ClientID, ClientName, ProfilVerify, TPSVerify, iDiscountPaymentConditionID, true);

                            if (ClientID == 609)
                            {
                                DetailsReport.ReportModuleOnline(ref hssfworkbook, Dispatches);
                            }
                            FileInfo file = new FileInfo(FilePath + @"\" + ReportName + ".xls");
                            int x = 1;
                            while (file.Exists == true)
                            {
                                file = new FileInfo(FilePath + @"\" + ReportName + "(" + x++ + ").xls");
                            }

                            FileStream NewFile = new FileStream(file.FullName, FileMode.Create);
                            hssfworkbook.Write(NewFile);
                            NewFile.Close();

                            ReportName = file.FullName;

                            if (Result == 1)
                            {
                                if (TotalProfil > 0)
                                {
                                    if (ClientCountryID == 1)
                                        TotalProfil *= 1.2m;
                                    ClientDispatchID = MutualSettlementsManager.AddClientDispatch(MutualSettlementID1, CurrencyTypeID, TotalProfil);
                                    var fileInfo = new System.IO.FileInfo(ReportName);
                                    MutualSettlementsManager.AttachDispatchDocument(Path.GetFileNameWithoutExtension(ReportName),
                                        fileInfo.Extension, ReportName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.DispatchExcel, ClientDispatchID);
                                    fileInfo = new System.IO.FileInfo(ProfilDBFName);
                                    MutualSettlementsManager.AttachDispatchDocument(Path.GetFileNameWithoutExtension(ProfilDBFName),
                                        fileInfo.Extension, ProfilDBFName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.DispatchDbf, ClientDispatchID);
                                }
                                if (TotalTPS > 0)
                                {
                                    if (ClientCountryID == 1)
                                        TotalTPS *= 1.2m;
                                    ClientDispatchID = MutualSettlementsManager.AddClientDispatch(MutualSettlementID2, CurrencyTypeID, TotalTPS);
                                    var fileInfo = new System.IO.FileInfo(ReportName);
                                    MutualSettlementsManager.AttachDispatchDocument(Path.GetFileNameWithoutExtension(ReportName),
                                        fileInfo.Extension, ReportName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.DispatchExcel, ClientDispatchID);
                                    fileInfo = new System.IO.FileInfo(TPSDBFName);
                                    MutualSettlementsManager.AttachDispatchDocument(Path.GetFileNameWithoutExtension(TPSDBFName),
                                        fileInfo.Extension, TPSDBFName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.DispatchDbf, ClientDispatchID);
                                }
                            }
                            MarketingDispatchManager.IncludeInMutualSettlement(Dispatches[j], MutualSettlementID1, MutualSettlementID2);
                        }

                        MutualSettlementsManager.SaveMutualSettlements();
                        MutualSettlementsManager.SaveMutualSettlementOrders();

                        UpdateDispatchDate();
                        MarketingDispatchManager.MoveToDispatchDate(DispatchDate);
                        MarketingDispatchManager.MoveToDispatch(Dispatches[0]);

                        NotesInvoiceClient(ClientID, ClientName, Dispatches, iDiscountPaymentConditionID, ClientLogin, ProfilVerify, TPSVerify, FilePath);

                        NeedSplash = true;
                        while (SplashWindow.bSmallCreated)
                            SmallWaitForm.CloseS = true;
                    }
                    if (HasOrdinaryOrders)
                    {
                        PhantomForm PhantomForm = new Infinium.PhantomForm();
                        PhantomForm.Show();

                        int Result = 1;
                        string SaveFilePath = string.Empty;
                        SaveDBFReportMenu = new Infinium.SaveDBFReportMenu(this, ReportName);

                        TopForm = SaveDBFReportMenu;
                        SaveDBFReportMenu.ShowDialog();

                        PhantomForm.Close();
                        Result = SaveDBFReportMenu.Result;
                        SaveFilePath = SaveDBFReportMenu.SaveFilePath;
                        PhantomForm.Dispose();
                        SaveDBFReportMenu.Dispose();
                        TopForm = null;

                        if (Result == 3)
                            return;

                        int ClientDispatchID = -1;
                        int MutualSettlementID1 = -1;
                        int MutualSettlementID2 = -1;
                        bool ProfilVerify = true;
                        bool TPSVerify = true;

                        ClientName = ClientName.Replace("\"", " ");
                        ClientName = ClientName.Replace("*", " ");
                        ClientName = ClientName.Replace("|", " ");
                        ClientName = ClientName.Replace(@"\", " ");
                        ClientName = ClientName.Replace(":", " ");
                        ClientName = ClientName.Replace("<", " ");
                        ClientName = ClientName.Replace(">", " ");
                        ClientName = ClientName.Replace("?", " ");
                        ClientName = ClientName.Replace("/", " ");

                        string CurrentMonthName = DateTime.Now.ToString("dd-MM-yyyy");
                        string FilePath = Security.DBFSaveFilePath;
                        if (Security.DBFSaveFilePath.Length == 0)
                            FilePath = Environment.GetFolderPath(Environment.SpecialFolder.Personal) + "/Infinium/1C data/" + CurrentMonthName;

                        Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Создание документа Excel.\r\nПодождите..."); });
                        T.Start();

                        while (!SplashWindow.bSmallCreated) ;
                        NeedSplash = false;

                        for (int j = 0; j < Dispatches.Count(); j++)
                        {
                            int[] MainOrders = MarketingDispatchManager.GetMainOrdersInDispatch(Dispatches[j]);

                            DataTable ReturnDT = MarketingDispatchManager.GetMegaOrdersInDispatch(Dispatches[j]);
                            ArrayList OrderNumbers = new ArrayList();
                            for (int i = 0; i < ReturnDT.Rows.Count; i++)
                                OrderNumbers.Add(Convert.ToInt32(ReturnDT.Rows[i]["OrderNumber"]));

                            int CurrencyTypeID = 1;
                            if (ReturnDT.Rows.Count > 0)
                                CurrencyTypeID = Convert.ToInt32(ReturnDT.Rows[0]["CurrencyTypeID"]);

                            MutualSettlementID1 = MutualSettlementsManager.FindMutualSettlementIDByOrderNumber(ClientID, 1, OrderNumbers.OfType<Int32>().ToArray(), false);
                            if (MutualSettlementID1 != 0)
                            {
                                decimal DispatchSum = MutualSettlementsManager.GetDispatchSum(MutualSettlementID1);
                                ProfilVerify = MutualSettlementsManager.VerifyTransaction(MutualSettlementID1, DispatchSum, ref DiscountPaymentCondition);
                                for (int i = 0; i < ReturnDT.Rows.Count; i++)
                                {
                                    int MegaOrderID = Convert.ToInt32(ReturnDT.Rows[i]["MegaOrderID"]);

                                    MegaOrderInfo m = MarketingDispatchManager.GetMegaOrders(MegaOrderID);
                                    iDiscountPaymentConditionID = ProfilReCalcMegaOrder(ClientID, CurrencyTypeID, DiscountPaymentCondition, ProfilVerify, false, MegaOrderID, m);
                                    OrdersManager.FixOrderEvent(MegaOrderID, "Выписать накладную. Пересчитан заказ Профиль");
                                }
                            }
                            MutualSettlementID2 = 0;
                            MutualSettlementID2 = MutualSettlementsManager.FindMutualSettlementIDByOrderNumber(ClientID, 2, OrderNumbers.OfType<Int32>().ToArray(), false);
                            if (MutualSettlementID2 != 0)
                            {
                                decimal DispatchSum = MutualSettlementsManager.GetDispatchSum(MutualSettlementID2);
                                TPSVerify = MutualSettlementsManager.VerifyTransaction(MutualSettlementID2, DispatchSum, ref DiscountPaymentCondition);
                                for (int i = 0; i < ReturnDT.Rows.Count; i++)
                                {
                                    int MegaOrderID = Convert.ToInt32(ReturnDT.Rows[i]["MegaOrderID"]);
                                    MegaOrderInfo m = MarketingDispatchManager.GetMegaOrders(MegaOrderID);
                                    iDiscountPaymentConditionID = TPSReCalcMegaOrder(ClientID, CurrencyTypeID, DiscountPaymentCondition, TPSVerify, false, MegaOrderID, m);
                                    OrdersManager.FixOrderEvent(MegaOrderID, "Выписать накладную. Пересчитан заказ ТПС");
                                }
                            }

                            HSSFWorkbook hssfworkbook = new HSSFWorkbook();

                            ////create a entry of DocumentSummaryInformation
                            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
                            dsi.Company = "NPOI Team";
                            hssfworkbook.DocumentSummaryInformation = dsi;

                            ////create a entry of SummaryInformation
                            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
                            si.Subject = "NPOI SDK Example";
                            hssfworkbook.SummaryInformation = si;

                            ReportName = ClientLogin + " №" + Dispatches[j].ToString();
                            _notesInvoiceReportToDbf.CreateReport(ref hssfworkbook, new int[1] { Dispatches[j] }, MainOrders, ClientID, ClientName, ref TotalProfil, ref TotalTPS, ProfilVerify, TPSVerify, iDiscountPaymentConditionID, false);
                            _notesInvoiceReportToDbf.SaveDBF(FilePath, ReportName, false, ref ProfilDBFName, ref TPSDBFName);
                            DetailsReport.CreateReport(ref hssfworkbook, new int[1] { Dispatches[j] }, OrderNumbers.OfType<Int32>().ToArray(), MainOrders, ClientID, ClientName, ProfilVerify, TPSVerify, iDiscountPaymentConditionID, false);

                            if (ClientID == 609)
                            {
                                DetailsReport.ReportModuleOnline(ref hssfworkbook, Dispatches);
                            }
                            FileInfo file = new FileInfo(FilePath + @"\" + ReportName + ".xls");
                            int x = 1;
                            while (file.Exists == true)
                            {
                                file = new FileInfo(FilePath + @"\" + ReportName + "(" + x++ + ").xls");
                            }

                            FileStream NewFile = new FileStream(file.FullName, FileMode.Create);
                            hssfworkbook.Write(NewFile);
                            NewFile.Close();

                            ReportName = file.FullName;

                            if (Result == 1)
                            {
                                if (TotalProfil > 0)
                                {
                                    if (ClientCountryID == 1)
                                        TotalProfil *= 1.2m;
                                    ClientDispatchID = MutualSettlementsManager.AddClientDispatch(MutualSettlementID1, CurrencyTypeID, TotalProfil);
                                    var fileInfo = new System.IO.FileInfo(ReportName);
                                    MutualSettlementsManager.AttachDispatchDocument(Path.GetFileNameWithoutExtension(ReportName),
                                        fileInfo.Extension, ReportName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.DispatchExcel, ClientDispatchID);
                                    fileInfo = new System.IO.FileInfo(ProfilDBFName);
                                    MutualSettlementsManager.AttachDispatchDocument(Path.GetFileNameWithoutExtension(ProfilDBFName),
                                        fileInfo.Extension, ProfilDBFName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.DispatchDbf, ClientDispatchID);
                                }
                                if (TotalTPS > 0)
                                {
                                    if (ClientCountryID == 1)
                                        TotalTPS *= 1.2m;
                                    ClientDispatchID = MutualSettlementsManager.AddClientDispatch(MutualSettlementID2, CurrencyTypeID, TotalTPS);
                                    var fileInfo = new System.IO.FileInfo(ReportName);
                                    MutualSettlementsManager.AttachDispatchDocument(Path.GetFileNameWithoutExtension(ReportName),
                                        fileInfo.Extension, ReportName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.DispatchExcel, ClientDispatchID);
                                    fileInfo = new System.IO.FileInfo(TPSDBFName);
                                    MutualSettlementsManager.AttachDispatchDocument(Path.GetFileNameWithoutExtension(TPSDBFName),
                                        fileInfo.Extension, TPSDBFName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.DispatchDbf, ClientDispatchID);
                                }
                            }
                            MarketingDispatchManager.IncludeInMutualSettlement(Dispatches[j], MutualSettlementID1, MutualSettlementID2);
                        }

                        MutualSettlementsManager.SaveMutualSettlements();
                        MutualSettlementsManager.SaveMutualSettlementOrders();

                        UpdateDispatchDate();
                        MarketingDispatchManager.MoveToDispatchDate(DispatchDate);
                        MarketingDispatchManager.MoveToDispatch(Dispatches[0]);

                        NotesInvoiceClient(ClientID, ClientName, Dispatches, iDiscountPaymentConditionID, ClientLogin, ProfilVerify, TPSVerify, FilePath);

                        NeedSplash = true;
                        while (SplashWindow.bSmallCreated)
                            SmallWaitForm.CloseS = true;
                    }
                }
                else
                {
                    PhantomForm PhantomForm = new Infinium.PhantomForm();
                    PhantomForm.Show();

                    int Result = 1;
                    string SaveFilePath = string.Empty;
                    SaveDBFReportMenu = new Infinium.SaveDBFReportMenu(this, ReportName);

                    TopForm = SaveDBFReportMenu;
                    SaveDBFReportMenu.ShowDialog();

                    PhantomForm.Close();
                    Result = SaveDBFReportMenu.Result;
                    SaveFilePath = SaveDBFReportMenu.SaveFilePath;
                    PhantomForm.Dispose();
                    SaveDBFReportMenu.Dispose();
                    TopForm = null;

                    if (Result == 3)
                        return;

                    int ClientDispatchID = -1;
                    int MutualSettlementID1 = -1;
                    int MutualSettlementID2 = -1;
                    bool ProfilVerify = true;
                    bool TPSVerify = true;

                    ClientName = ClientName.Replace("\"", " ");
                    ClientName = ClientName.Replace("*", " ");
                    ClientName = ClientName.Replace("|", " ");
                    ClientName = ClientName.Replace(@"\", " ");
                    ClientName = ClientName.Replace(":", " ");
                    ClientName = ClientName.Replace("<", " ");
                    ClientName = ClientName.Replace(">", " ");
                    ClientName = ClientName.Replace("?", " ");
                    ClientName = ClientName.Replace("/", " ");

                    string CurrentMonthName = DateTime.Now.ToString("dd-MM-yyyy");
                    string FilePath = Security.DBFSaveFilePath;
                    if (Security.DBFSaveFilePath.Length == 0)
                        FilePath = Environment.GetFolderPath(Environment.SpecialFolder.Personal) + "/Infinium/1C data/" + CurrentMonthName;

                    Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Создание документа Excel.\r\nПодождите..."); });
                    T.Start();

                    while (!SplashWindow.bSmallCreated) ;
                    NeedSplash = false;

                    for (int j = 0; j < Dispatches.Count(); j++)
                    {
                        int[] MainOrders = MarketingDispatchManager.GetMainOrdersInDispatch(Dispatches[j]);

                        DataTable ReturnDT = MarketingDispatchManager.GetMegaOrdersInDispatch(Dispatches[j]);
                        ArrayList OrderNumbers = new ArrayList();
                        for (int i = 0; i < ReturnDT.Rows.Count; i++)
                            OrderNumbers.Add(Convert.ToInt32(ReturnDT.Rows[i]["OrderNumber"]));

                        int CurrencyTypeID = 1;
                        int FactoryID = 1;

                        if (ReturnDT.Rows.Count > 0)
                        {
                            CurrencyTypeID = Convert.ToInt32(ReturnDT.Rows[0]["CurrencyTypeID"]);
                            FactoryID = Convert.ToInt32(ReturnDT.Rows[0]["FactoryID"]);
                        }

                        MutualSettlementID1 = MutualSettlementsManager.FindMutualSettlementIDByOrderNumber(ClientID, 1, OrderNumbers.OfType<Int32>().ToArray());
                        if (MutualSettlementID1 != 0)
                        {
                            decimal DispatchSum = MutualSettlementsManager.GetDispatchSum(MutualSettlementID1);
                            ProfilVerify = MutualSettlementsManager.VerifyTransaction(MutualSettlementID1, DispatchSum, ref DiscountPaymentCondition);
                            for (int i = 0; i < ReturnDT.Rows.Count; i++)
                            {
                                int MegaOrderID = Convert.ToInt32(ReturnDT.Rows[i]["MegaOrderID"]);

                                MegaOrderInfo m = MarketingDispatchManager.GetMegaOrders(MegaOrderID);
                                iDiscountPaymentConditionID = ProfilReCalcMegaOrder(ClientID, CurrencyTypeID, DiscountPaymentCondition, ProfilVerify, MegaOrderID, m);
                                OrdersManager.FixOrderEvent(MegaOrderID, "Выписать накладную. Пересчитан заказ Профиль");
                            }
                        }
                        MutualSettlementID2 = 0;
                        MutualSettlementID2 = MutualSettlementsManager.FindMutualSettlementIDByOrderNumber(ClientID, 2, OrderNumbers.OfType<Int32>().ToArray());
                        if (MutualSettlementID2 != 0)
                        {
                            decimal DispatchSum = MutualSettlementsManager.GetDispatchSum(MutualSettlementID2);
                            TPSVerify = MutualSettlementsManager.VerifyTransaction(MutualSettlementID2, DispatchSum, ref DiscountPaymentCondition);
                            for (int i = 0; i < ReturnDT.Rows.Count; i++)
                            {
                                int MegaOrderID = Convert.ToInt32(ReturnDT.Rows[i]["MegaOrderID"]);
                                MegaOrderInfo m = MarketingDispatchManager.GetMegaOrders(MegaOrderID);
                                iDiscountPaymentConditionID = TPSReCalcMegaOrder(ClientID, CurrencyTypeID, DiscountPaymentCondition, TPSVerify, MegaOrderID, m);
                                OrdersManager.FixOrderEvent(MegaOrderID, "Выписать накладную. Пересчитан заказ ТПС");
                            }
                        }

                        HSSFWorkbook hssfworkbook = new HSSFWorkbook();

                        ////create a entry of DocumentSummaryInformation
                        DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
                        dsi.Company = "NPOI Team";
                        hssfworkbook.DocumentSummaryInformation = dsi;

                        ////create a entry of SummaryInformation
                        SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
                        si.Subject = "NPOI SDK Example";
                        hssfworkbook.SummaryInformation = si;

                        ReportName = ClientLogin + " №" + Dispatches[j].ToString();

                        _notesInvoiceReportToDbf.CreateReport(ref hssfworkbook, new int[1] { Dispatches[j] }, MainOrders, ClientID, ClientName, ref TotalProfil, ref TotalTPS, ProfilVerify, TPSVerify, iDiscountPaymentConditionID);
                        _notesInvoiceReportToDbf.SaveDBF(FilePath, ReportName, ref ProfilDBFName, ref TPSDBFName);

                        DetailsReport.CreateReport(ref hssfworkbook, new int[1] { Dispatches[j] }, OrderNumbers.OfType<Int32>().ToArray(), MainOrders, ClientID, ClientName, ProfilVerify, TPSVerify, iDiscountPaymentConditionID);

                        if (ClientID == 609)
                        {
                            DetailsReport.ReportModuleOnline(ref hssfworkbook, Dispatches);
                        }
                        FileInfo file = new FileInfo(FilePath + @"\" + ReportName + ".xls");
                        int x = 1;
                        while (file.Exists == true)
                        {
                            file = new FileInfo(FilePath + @"\" + ReportName + "(" + x++ + ").xls");
                        }

                        FileStream NewFile = new FileStream(file.FullName, FileMode.Create);
                        hssfworkbook.Write(NewFile);
                        NewFile.Close();

                        ReportName = file.FullName;

                        if (Result == 1)
                        {
                            if (TotalProfil > 0)
                            {
                                if (ClientCountryID == 1)
                                    TotalProfil *= 1.2m;
                                ClientDispatchID = MutualSettlementsManager.AddClientDispatch(MutualSettlementID1, CurrencyTypeID, TotalProfil);
                                var fileInfo = new System.IO.FileInfo(ReportName);
                                MutualSettlementsManager.AttachDispatchDocument(Path.GetFileNameWithoutExtension(ReportName),
                                    fileInfo.Extension, ReportName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.DispatchExcel, ClientDispatchID);
                                fileInfo = new System.IO.FileInfo(ProfilDBFName);
                                MutualSettlementsManager.AttachDispatchDocument(Path.GetFileNameWithoutExtension(ProfilDBFName),
                                    fileInfo.Extension, ProfilDBFName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.DispatchDbf, ClientDispatchID);
                            }
                            if (TotalTPS > 0)
                            {
                                if (ClientCountryID == 1)
                                    TotalTPS *= 1.2m;
                                ClientDispatchID = MutualSettlementsManager.AddClientDispatch(MutualSettlementID2, CurrencyTypeID, TotalTPS);
                                var fileInfo = new System.IO.FileInfo(ReportName);
                                MutualSettlementsManager.AttachDispatchDocument(Path.GetFileNameWithoutExtension(ReportName),
                                    fileInfo.Extension, ReportName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.DispatchExcel, ClientDispatchID);
                                fileInfo = new System.IO.FileInfo(TPSDBFName);
                                MutualSettlementsManager.AttachDispatchDocument(Path.GetFileNameWithoutExtension(TPSDBFName),
                                    fileInfo.Extension, TPSDBFName, Infinium.Modules.Marketing.MutualSettlements.DocumentTypes.DispatchDbf, ClientDispatchID);
                            }
                        }
                        MarketingDispatchManager.IncludeInMutualSettlement(Dispatches[j], MutualSettlementID1, MutualSettlementID2);
                    }

                    MutualSettlementsManager.SaveMutualSettlements();
                    MutualSettlementsManager.SaveMutualSettlementOrders();

                    UpdateDispatchDate();
                    MarketingDispatchManager.MoveToDispatchDate(DispatchDate);
                    MarketingDispatchManager.MoveToDispatch(Dispatches[0]);

                    NotesInvoiceClient(ClientID, ClientName, Dispatches, iDiscountPaymentConditionID, ClientLogin, ProfilVerify, TPSVerify, FilePath);

                    NeedSplash = true;
                    while (SplashWindow.bSmallCreated)
                        SmallWaitForm.CloseS = true;
                }
            }
        }

        private void kryptonButton8_Click(object sender, EventArgs e)
        {
            Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
            T.Start();

            while (!SplashWindow.bSmallCreated) ;

            NeedSplash = false;
            cabFurAssembleManager = new Modules.CabFurnitureAssignments.CabFurAssemble();
            cabFurAssembleManager.Initialize();
            dgvCabFurMegaOrdersSetting();
            dgvCabFurDatesSetting();
            dgvCabFurMainOrdersSetting();
            UpdateCabFurDispatchDate();

            NeedSplash = true;

            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;

        }

        private void btnFilterByDate_Click(object sender, EventArgs e)
        {
            if (MarketingExpeditionManager != null)
            {
                Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                T.Start();

                while (!SplashWindow.bSmallCreated) ;
                NeedSplash = false;

                FilterMegaBatch();
                FilterOrders();

                NeedSplash = true;
                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;
            }
        }
    }
}