﻿using Infinium.Modules.Marketing.NewOrders;
using Infinium.Modules.Marketing.NewOrders.InvoiceReportToDbf;
using System;
using System.Data;
using System.Windows.Forms;

namespace Infinium
{
    public partial class AddMarketingNewOrdersForm : Form
    {
        private const int eHide = 2;
        private const int eShow = 1;
        private const int eClose = 3;
        private const int eMainMenu = 4;

        private int FormEvent = 0;
        private int MainOrderID = -1;

        private Form TopForm = null;

        private bool EditAgreedMainOrder = false;
        private bool EditMainOrder = false;
        private bool ClientOMC = false;

        public OrdersManager _ordersManager;

        private FrontsCatalogOrder FrontsCatalogOrder = null;
        private DecorCatalogOrder DecorCatalogOrder = null;
        private FrontsOrders FrontsOrders;
        private DecorOrders DecorOrders;
        private OrdersCalculate _ordersCalculate;



        public AddMarketingNewOrdersForm(ref OrdersManager tOrdersManager, bool bClientOMC,
            bool bEditMainOrder, bool bEditAgreedMainOrder, ref Form tTopForm, ref OrdersCalculate tOrdersCalculate)
        {
            TopForm = tTopForm;

            InitializeComponent();

            this.MaximumSize = Screen.PrimaryScreen.WorkingArea.Size;

            _ordersManager = tOrdersManager;
            EditMainOrder = bEditMainOrder;
            ClientOMC = bClientOMC;
            EditAgreedMainOrder = bEditAgreedMainOrder;
            _ordersCalculate = tOrdersCalculate;

            Initialize();

            while (!SplashForm.bCreated) ;
        }

        private void AddMarketMainOrdersForm_Shown(object sender, EventArgs e)
        {
            FormEvent = eShow;
            AnimateTimer.Enabled = true;

            FrontsComboBox.SelectionLength = 0;
            FrameColorComboBox.SelectionLength = 0;
            TechnoProfilesComboBox.SelectionLength = 0;
            TechnoFrameColorComboBox.SelectionLength = 0;
            PatinaComboBox.SelectionLength = 0;
            InsetTypesComboBox.SelectionLength = 0;
            InsetColorComboBox.SelectionLength = 0;
            TechnoInsetTypesComboBox.SelectionLength = 0;
            TechnoInsetColorsComboBox.SelectionLength = 0;

            FrontsComboBox.Focus();
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

        private void NavigateMenuBackButton_Click(object sender, EventArgs e)
        {
            FormEvent = eHide;
            AnimateTimer.Enabled = true;
        }

        private void NavigateMenuHerculesButton_Click(object sender, EventArgs e)
        {
            FormEvent = eMainMenu;
            AnimateTimer.Enabled = true;
        }

        private void Initialize()
        {
            FrontsCatalogOrder = new FrontsCatalogOrder(ref FrontsHeightComboBox, ref FrontsWidthComboBox);
            FrontsCatalogOrder.Initialize(true);

            DecorCatalogOrder = new DecorCatalogOrder();
            DecorCatalogOrder.Initialize(true);

            //ExportOrdersFromExcel T = new ExportOrdersFromExcel(ref DecorCatalogOrder, ref FrontsCatalogOrder);
            //T.FF();

            System.Diagnostics.Stopwatch sw = new System.Diagnostics.Stopwatch();
            sw.Start();

            FrontsOrders = new FrontsOrders(ref FrontsOrdersDataGrid,
                ref FrontsCatalogOrder);

            sw.Stop();
            TimeSpan ts1 = sw.Elapsed;
            sw.Restart();

            if (!ClientOMC)
                FrontsOrders.HasClientExcluzive(_ordersManager.CurrentClientID);

            sw.Stop();
            TimeSpan ts2 = sw.Elapsed;
            sw.Restart();
            FrontsOrders.Initialize();

            sw.Stop();
            TimeSpan ts3 = sw.Elapsed;

            DecorOrders = new DecorOrders(ref MainOrdersDecorTabControl,
                ref DecorCatalogOrder, ref FrontsOrdersDataGrid);

            sw.Stop();
            TimeSpan ts4 = sw.Elapsed;
            sw.Restart();

            if (!ClientOMC)
                DecorOrders.HasClientExcluzive(_ordersManager.CurrentClientID);

            sw.Stop();
            TimeSpan ts5 = sw.Elapsed;
            sw.Restart();

            DecorOrders.Initialize();

            sw.Stop();
            TimeSpan ts6 = sw.Elapsed;

            FrontsCatalogOrderBinding();

            DecorCatalogOrderBinding();

            
            //cbOnlyExcluzive.Visible = false;
            if (EditMainOrder)
            {
                DateTime DispatchDate = DateTime.Now;
                bool IsSample = false;
                string Notes = null;
                MainOrderID = _ordersManager.CurrentMainOrderID;
                //получение значений параметров заказа, если заблокирован - выход
                //if (!EditAgreedMainOrder)
                //{
                //        return;
                //}

                if (!_ordersManager.EditMainOrder(ref Notes, ref IsSample))
                    return;

                //редактирование фасадов
                FrontsOrdersDataGrid.Visible = FrontsOrders.EditOrder(MainOrderID);

                //редактирование декора
                if (DecorOrders.EditDecorOrder(MainOrderID))
                    NoDecorLabel.Visible = false;
                else
                    NoDecorLabel.Visible = true;

                //установка параметров в контролы
                MainOrderNotes.Text = Notes;
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

        private void kryptonButton1_Click(object sender, EventArgs e)
        {
            if (FrontsOrders.HasRows() == false && DecorOrders.HasRows() == false)
            {
                DialogResult result = MessageBox.Show(this, "В заказе нет позиций. Удалить заказ?",
                                "Сохранение заказа", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                if (result == System.Windows.Forms.DialogResult.Yes)
                {
                    _ordersManager.RemoveCurrentMainOrder();
                }

                if (result == System.Windows.Forms.DialogResult.No)
                {
                    return;
                }
            }

            FormEvent = eClose;
            AnimateTimer.Enabled = true;
        }

        //ORDER FRONTS CATALOG
        public void NGetFrameColors()
        {
            bool bExcluzive = cbOnlyExcluzive.Checked;
            if (!FrontsOrders.HasExcluzive)
                bExcluzive = false;
            if (FrontsComboBox.SelectedItem == null)
                return;
            if (FrontsCatalogOrder.FrontsBindingSource.Current != null)
            {
                FrontsCatalogOrder.FilterCatalogFrameColors(((DataRowView)FrontsComboBox.SelectedItem).Row["FrontName"].ToString(), bExcluzive);
            }
            else
            {
                FrontsCatalogOrder.FilterCatalogFrameColors(string.Empty, bExcluzive);
            }

        }

        public void NGetTechnoProfiles()
        {
            bool bExcluzive = cbOnlyExcluzive.Checked;
            if (!FrontsOrders.HasExcluzive)
                bExcluzive = false;
            if (FrontsComboBox.SelectedItem == null)
                return;
            if (FrontsCatalogOrder.FrameColorsBindingSource.Current != null)
            {
                FrontsCatalogOrder.FilterCatalogTechnoProfiles(((DataRowView)FrontsComboBox.SelectedItem).Row["FrontName"].ToString(),
                    Convert.ToInt32(FrameColorComboBox.SelectedValue), bExcluzive);
            }
            else
            {
                FrontsCatalogOrder.FilterCatalogTechnoProfiles(string.Empty, -1, bExcluzive);
            }
        }

        public void NGetTechnoFrameColors()
        {
            bool bExcluzive = cbOnlyExcluzive.Checked;
            if (!FrontsOrders.HasExcluzive)
                bExcluzive = false;
            if (FrontsComboBox.SelectedItem == null)
                return;
            if (FrontsCatalogOrder.FrameColorsBindingSource.Current != null)
            {
                FrontsCatalogOrder.FilterCatalogTechnoFrameColors(((DataRowView)FrontsComboBox.SelectedItem).Row["FrontName"].ToString(),
                    Convert.ToInt32(FrameColorComboBox.SelectedValue), Convert.ToInt32(TechnoProfilesComboBox.SelectedValue), bExcluzive);
            }
            else
            {
                FrontsCatalogOrder.FilterCatalogTechnoFrameColors(string.Empty, -1, -1, bExcluzive);
            }
        }

        public void NGetPatina()
        {
            bool bExcluzive = cbOnlyExcluzive.Checked;
            if (!FrontsOrders.HasExcluzive)
                bExcluzive = false;
            if (FrontsComboBox.SelectedItem == null)
                return;
            if (FrontsCatalogOrder.TechnoFrameColorsBindingSource.Current != null)
            {
                FrontsCatalogOrder.FilterCatalogPatina(((DataRowView)FrontsComboBox.SelectedItem).Row["FrontName"].ToString(),
                    Convert.ToInt32(FrameColorComboBox.SelectedValue), Convert.ToInt32(TechnoProfilesComboBox.SelectedValue),
                    Convert.ToInt32(TechnoFrameColorComboBox.SelectedValue), bExcluzive);
            }
            else
            {
                FrontsCatalogOrder.FilterCatalogPatina(string.Empty, -1, -1, -1, bExcluzive);
            }
        }

        public void NGetInsetTypes()
        {
            bool bExcluzive = cbOnlyExcluzive.Checked;
            if (!FrontsOrders.HasExcluzive)
                bExcluzive = false;
            if (FrontsComboBox.SelectedItem == null)
                return;
            if (FrontsCatalogOrder.PatinaBindingSource.Current != null)
            {
                int PatinaID = Convert.ToInt32(PatinaComboBox.SelectedValue);

                FrontsCatalogOrder.FilterCatalogInsetTypes(((DataRowView)FrontsComboBox.SelectedItem).Row["FrontName"].ToString(),
                    Convert.ToInt32(FrameColorComboBox.SelectedValue), Convert.ToInt32(TechnoProfilesComboBox.SelectedValue),
                    Convert.ToInt32(TechnoFrameColorComboBox.SelectedValue),
                    PatinaID, bExcluzive);
            }
            else
            {
                FrontsCatalogOrder.FilterCatalogInsetTypes(string.Empty, -1, -1, -1, -1, bExcluzive);
            }
        }

        public void NGetInsetColors()
        {
            bool bExcluzive = cbOnlyExcluzive.Checked;
            if (!FrontsOrders.HasExcluzive)
                bExcluzive = false;
            if (FrontsComboBox.SelectedItem == null)
                return;
            if (FrontsCatalogOrder.InsetTypesBindingSource.Current != null)
            {
                int PatinaID = Convert.ToInt32(PatinaComboBox.SelectedValue);

                FrontsCatalogOrder.FilterCatalogInsetColors(((DataRowView)FrontsComboBox.SelectedItem).Row["FrontName"].ToString(),
                    Convert.ToInt32(FrameColorComboBox.SelectedValue), Convert.ToInt32(TechnoProfilesComboBox.SelectedValue),
                    Convert.ToInt32(TechnoFrameColorComboBox.SelectedValue),
                    PatinaID,
                    Convert.ToInt32(InsetTypesComboBox.SelectedValue), bExcluzive);
            }
            else
            {
                FrontsCatalogOrder.FilterCatalogInsetColors(string.Empty, -1, -1, -1, -1, -1, bExcluzive);
            }
        }

        public void NGetTechnoInsetTypes()
        {
            bool bExcluzive = cbOnlyExcluzive.Checked;
            if (!FrontsOrders.HasExcluzive)
                bExcluzive = false;
            if (FrontsComboBox.SelectedItem == null)
                return;
            if (FrontsCatalogOrder.InsetColorsBindingSource.Current != null)
            {
                int PatinaID = Convert.ToInt32(PatinaComboBox.SelectedValue);

                FrontsCatalogOrder.FilterCatalogTechnoInsetTypes(((DataRowView)FrontsComboBox.SelectedItem).Row["FrontName"].ToString(),
                    Convert.ToInt32(FrameColorComboBox.SelectedValue), Convert.ToInt32(TechnoProfilesComboBox.SelectedValue),
                    Convert.ToInt32(TechnoFrameColorComboBox.SelectedValue),
                    PatinaID,
                    Convert.ToInt32(InsetTypesComboBox.SelectedValue),
                    Convert.ToInt32(InsetColorComboBox.SelectedValue), bExcluzive);
            }
            else
            {
                FrontsCatalogOrder.FilterCatalogTechnoInsetTypes(string.Empty, -1, -1, -1, -1, -1, -1, bExcluzive);
            }
        }

        public void NGetTechnoInsetColors()
        {
            bool bExcluzive = cbOnlyExcluzive.Checked;
            if (!FrontsOrders.HasExcluzive)
                bExcluzive = false;
            if (FrontsComboBox.SelectedItem == null)
                return;
            if (FrontsCatalogOrder.TechnoInsetTypesBindingSource.Current != null)
            {
                int PatinaID = Convert.ToInt32(PatinaComboBox.SelectedValue);

                FrontsCatalogOrder.FilterCatalogTechnoInsetColors(((DataRowView)FrontsComboBox.SelectedItem).Row["FrontName"].ToString(),
                                                       Convert.ToInt32(FrameColorComboBox.SelectedValue), Convert.ToInt32(TechnoProfilesComboBox.SelectedValue),
                                                       Convert.ToInt32(TechnoFrameColorComboBox.SelectedValue),
                                                       PatinaID,
                                                       Convert.ToInt32(InsetTypesComboBox.SelectedValue),
                                                       Convert.ToInt32(InsetColorComboBox.SelectedValue),
                                                       Convert.ToInt32(TechnoInsetTypesComboBox.SelectedValue), bExcluzive);
            }
            else
            {
                FrontsCatalogOrder.FilterCatalogTechnoInsetColors(string.Empty, -1, -1, -1, -1, -1, -1, -1, bExcluzive);
            }
        }

        public void NGetHeight()
        {
            if (FrontsComboBox.SelectedItem == null)
                return;
            int PatinaID = Convert.ToInt32(PatinaComboBox.SelectedValue);

            FrontsCatalogOrder.FilterCatalogHeight(((DataRowView)FrontsComboBox.SelectedItem).Row["FrontName"].ToString(),
                                                            Convert.ToInt32(FrameColorComboBox.SelectedValue), Convert.ToInt32(TechnoProfilesComboBox.SelectedValue),
                                                            Convert.ToInt32(TechnoFrameColorComboBox.SelectedValue),
                                                            PatinaID,
                                                            Convert.ToInt32(InsetTypesComboBox.SelectedValue),
                                                            Convert.ToInt32(InsetColorComboBox.SelectedValue),
                                                            Convert.ToInt32(TechnoInsetTypesComboBox.SelectedValue),
                                                            Convert.ToInt32(TechnoInsetColorsComboBox.SelectedValue));
        }

        public void NGetWidth()
        {
            if (FrontsComboBox.SelectedItem == null)
                return;
            int Height = 0;

            if (FrontsHeightComboBox.Text != "")
                Height = Convert.ToInt32(FrontsHeightComboBox.Text);
            int PatinaID = Convert.ToInt32(PatinaComboBox.SelectedValue);

            FrontsCatalogOrder.FilterCatalogWidth(((DataRowView)FrontsComboBox.SelectedItem).Row["FrontName"].ToString(),
                                                            Convert.ToInt32(FrameColorComboBox.SelectedValue), Convert.ToInt32(TechnoProfilesComboBox.SelectedValue),
                                                            Convert.ToInt32(TechnoFrameColorComboBox.SelectedValue),
                                                            PatinaID,
                                                            Convert.ToInt32(InsetTypesComboBox.SelectedValue),
                                                            Convert.ToInt32(InsetColorComboBox.SelectedValue),
                                                            Convert.ToInt32(TechnoInsetTypesComboBox.SelectedValue),
                                                            Convert.ToInt32(TechnoInsetColorsComboBox.SelectedValue),
                                                            Height);
        }

        public void FrontsCatalogOrderBinding()
        {
            FrontsComboBox.DisplayMember = FrontsCatalogOrder.FrontsBindingSourceDisplayMember;
            FrontsComboBox.ValueMember = FrontsCatalogOrder.FrontsBindingSourceValueMember;
            FrontsComboBox.DataSource = FrontsCatalogOrder.FrontsBindingSource;

            TechnoProfilesComboBox.DisplayMember = FrontsCatalogOrder.TechnoProfilesBindingSourceDisplayMember;
            TechnoProfilesComboBox.ValueMember = FrontsCatalogOrder.TechnoProfilesBindingSourceValueMember;
            TechnoProfilesComboBox.DataSource = FrontsCatalogOrder.TechnoProfilesBindingSource;

            FrameColorComboBox.DisplayMember = FrontsCatalogOrder.FrameColorsBindingSourceDisplayMember;
            FrameColorComboBox.ValueMember = FrontsCatalogOrder.FrameColorsBindingSourceValueMember;
            FrameColorComboBox.DataSource = FrontsCatalogOrder.FrameColorsBindingSource;

            PatinaComboBox.DisplayMember = FrontsCatalogOrder.PatinaBindingSourceDisplayMember;
            PatinaComboBox.ValueMember = FrontsCatalogOrder.PatinaBindingSourceValueMember;
            PatinaComboBox.DataSource = FrontsCatalogOrder.PatinaBindingSource;

            InsetTypesComboBox.DisplayMember = FrontsCatalogOrder.InsetTypesBindingSourceDisplayMember;
            InsetTypesComboBox.ValueMember = FrontsCatalogOrder.InsetTypesBindingSourceValueMember;
            InsetTypesComboBox.DataSource = FrontsCatalogOrder.InsetTypesBindingSource;

            InsetColorComboBox.DisplayMember = FrontsCatalogOrder.InsetColorsBindingSourceDisplayMember;
            InsetColorComboBox.ValueMember = FrontsCatalogOrder.InsetColorsBindingSourceValueMember;
            InsetColorComboBox.DataSource = FrontsCatalogOrder.InsetColorsBindingSource;

            TechnoFrameColorComboBox.DisplayMember = FrontsCatalogOrder.FrameColorsBindingSourceDisplayMember;
            TechnoFrameColorComboBox.ValueMember = FrontsCatalogOrder.FrameColorsBindingSourceValueMember;
            TechnoFrameColorComboBox.DataSource = FrontsCatalogOrder.TechnoFrameColorsBindingSource;

            TechnoInsetTypesComboBox.DisplayMember = FrontsCatalogOrder.InsetTypesBindingSourceDisplayMember;
            TechnoInsetTypesComboBox.ValueMember = FrontsCatalogOrder.InsetTypesBindingSourceValueMember;
            TechnoInsetTypesComboBox.DataSource = FrontsCatalogOrder.TechnoInsetTypesBindingSource;

            TechnoInsetColorsComboBox.DisplayMember = FrontsCatalogOrder.InsetColorsBindingSourceDisplayMember;
            TechnoInsetColorsComboBox.ValueMember = FrontsCatalogOrder.InsetColorsBindingSourceValueMember;
            TechnoInsetColorsComboBox.DataSource = FrontsCatalogOrder.TechnoInsetColorsBindingSource;

            bool bExcluzive = cbOnlyExcluzive.Checked;
            if (!FrontsOrders.HasExcluzive)
                bExcluzive = false;
            FrontsCatalogOrder.FilterFronts(bExcluzive);
            NGetFrameColors();
            NGetTechnoProfiles();
            NGetTechnoFrameColors();
            NGetPatina();
            NGetInsetTypes();
            NGetInsetColors();
            NGetTechnoInsetTypes();
            NGetTechnoInsetColors();
            NGetHeight();
            NGetWidth();
        }

        private void FrontsComboBox_SelectionChangeCommitted(object sender, EventArgs e)
        {
            NGetFrameColors();
            NGetTechnoProfiles();
            NGetTechnoFrameColors();
            NGetPatina();
            NGetInsetTypes();
            NGetInsetColors();
            NGetTechnoInsetTypes();
            NGetTechnoInsetColors();
            NGetHeight();
            NGetWidth();
        }
        private void FrameColorComboBox_SelectionChangeCommitted(object sender, EventArgs e)
        {
            NGetTechnoProfiles();
            NGetTechnoFrameColors();
            NGetPatina();
            NGetInsetTypes();
            NGetInsetColors();
            NGetTechnoInsetTypes();
            NGetTechnoInsetColors();
            NGetHeight();
            NGetWidth();
        }

        private void TechnoProfilesComboBox_SelectionChangeCommitted(object sender, EventArgs e)
        {
            NGetTechnoFrameColors();
            NGetPatina();
            NGetInsetTypes();
            NGetInsetColors();
            NGetTechnoInsetTypes();
            NGetTechnoInsetColors();
            NGetHeight();
            NGetWidth();
        }

        private void TechnoFrameColorComboBox_SelectionChangeCommitted(object sender, EventArgs e)
        {
            NGetPatina();
            NGetInsetTypes();
            NGetInsetColors();
            NGetTechnoInsetTypes();
            NGetTechnoInsetColors();
            NGetHeight();
            NGetWidth();
        }

        private void PatinaComboBox_SelectionChangeCommitted(object sender, EventArgs e)
        {
            NGetInsetTypes();
            NGetInsetColors();
            NGetTechnoInsetTypes();
            NGetTechnoInsetColors();
            NGetHeight();
            NGetWidth();
        }

        private void InsetTypesComboBox_SelectionChangeCommitted(object sender, EventArgs e)
        {
            NGetInsetColors();
            NGetTechnoInsetTypes();
            NGetTechnoInsetColors();
            NGetHeight();
            NGetWidth();
        }

        private void InsetColorComboBox_SelectionChangeCommitted(object sender, EventArgs e)
        {
            NGetTechnoInsetTypes();
            NGetTechnoInsetColors();
            NGetHeight();
            NGetWidth();
        }

        private void TechnoInsetTypesComboBox_SelectionChangeCommitted(object sender, EventArgs e)
        {
            NGetTechnoInsetColors();
            NGetHeight();
            NGetWidth();
        }

        private void TechnoInsetColorsComboBox_SelectionChangeCommitted(object sender, EventArgs e)
        {
            NGetHeight();
            NGetWidth();
        }
        private void FrontsHeightComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            NGetWidth();
        }



        //ORDER DECOR CATALOG
        public void NGetDecorItems()
        {
            bool bExcluzive = cbOnlyExcluzive.Checked;
            if (!DecorOrders.HasExcluzive)
                bExcluzive = false;
            DecorCatalogOrder.FilterItems(Convert.ToInt32(DecorProductsComboBox.SelectedValue), bExcluzive);
        }

        public void NGetDecorColors()
        {
            bool bExcluzive = cbOnlyExcluzive.Checked;
            if (!DecorOrders.HasExcluzive)
                bExcluzive = false;
            if (DecorItemsComboBox.SelectedItem == null)
                return;
            if (DecorItemsComboBox.SelectedItem != null &&
                DecorCatalogOrder.FilterColors(((DataRowView)DecorItemsComboBox.SelectedItem).Row["Name"].ToString(), bExcluzive) == true)
            {
                panel31.Visible = true;
            }
            else
            {
                panel31.Visible = false;
            }
        }

        public void NGetDecorPatina()
        {
            bool bExcluzive = cbOnlyExcluzive.Checked;
            if (!DecorOrders.HasExcluzive)
                bExcluzive = false;
            if (DecorItemsComboBox.SelectedItem == null)
                return;
            if (DecorItemsComboBox.SelectedItem != null &&
                DecorCatalogOrder.FilterPatina(((DataRowView)DecorItemsComboBox.SelectedItem).Row["Name"].ToString(),
                Convert.ToInt32(DecorColorsComboBox.SelectedValue), bExcluzive) == true)
            {
                panel29.Visible = true;
            }
            else
            {
                panel29.Visible = false;
            }
        }

        public void NGetDecorInsetTypes()
        {
            if (DecorItemsComboBox.SelectedItem == null)
                return;
            bool bExcluzive = cbOnlyExcluzive.Checked;
            if (!DecorOrders.HasExcluzive)
                bExcluzive = false;
            DecorCatalogOrder.FilterInsetType(((DataRowView)DecorItemsComboBox.SelectedItem).Row["Name"].ToString(),
                Convert.ToInt32(DecorColorsComboBox.SelectedValue),
                Convert.ToInt32(DecorPatinaComboBox.SelectedValue), bExcluzive);
        }

        public void NGetDecorInsetColors()
        {
            if (DecorItemsComboBox.SelectedItem == null)
                return;
            bool bExcluzive = cbOnlyExcluzive.Checked;
            if (!DecorOrders.HasExcluzive)
                bExcluzive = false;
            DecorCatalogOrder.FilterInsetColor(((DataRowView)DecorItemsComboBox.SelectedItem).Row["Name"].ToString(),
                Convert.ToInt32(DecorColorsComboBox.SelectedValue),
                Convert.ToInt32(DecorPatinaComboBox.SelectedValue),
                Convert.ToInt32(DecorInsetTypesComboBox.SelectedValue), bExcluzive);
        }

        public void NGetDecorLength()
        {
            if (DecorItemsComboBox.SelectedItem == null)
                return;
            int L = DecorCatalogOrder.FilterLength(((DataRowView)DecorItemsComboBox.SelectedItem).Row["Name"].ToString(),
                Convert.ToInt32(DecorColorsComboBox.SelectedValue),
                Convert.ToInt32(DecorPatinaComboBox.SelectedValue),
                Convert.ToInt32(DecorInsetTypesComboBox.SelectedValue),
                Convert.ToInt32(DecorInsetColorsComboBox.SelectedValue));

            DecorLengthComboBox.DropDownStyle = ComboBoxStyle.DropDown;
            if (L == -1)
            {
                panel4.Visible = false;
            }
            else if (L == 0)
            {
                panel4.Visible = true;
                DecorLengthComboBox.DropDownStyle = ComboBoxStyle.DropDown;
            }
            else if (L > 0)
            {
                panel4.Visible = true;
                DecorLengthComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            }
        }

        public void NGetDecorHeight()
        {
            if (DecorItemsComboBox.SelectedItem == null)
                return;
            int Length = -1;
            //if (DecorCatalogOrder.ItemLengthBindingSource.Count > 0)
            //    Length = Convert.ToInt32(((DataRowView)DecorCatalogOrder.ItemLengthBindingSource.Current).Row["Length"]);
            if (DecorLengthComboBox.SelectedValue != null)
                Length = Convert.ToInt32(DecorLengthComboBox.SelectedValue);

            int H = DecorCatalogOrder.FilterHeight(((DataRowView)DecorItemsComboBox.SelectedItem).Row["Name"].ToString(),
                Convert.ToInt32(DecorColorsComboBox.SelectedValue),
                Convert.ToInt32(DecorPatinaComboBox.SelectedValue),
                Convert.ToInt32(DecorInsetTypesComboBox.SelectedValue),
                Convert.ToInt32(DecorInsetColorsComboBox.SelectedValue), Length);

            DecorHeightComboBox.DropDownStyle = ComboBoxStyle.DropDown;
            if (H == -1)
            {
                panel1.Visible = false;
            }
            else if (H == 0)
            {
                panel1.Visible = true;
                DecorHeightComboBox.DropDownStyle = ComboBoxStyle.DropDown;
            }
            else if (H > 0)
            {
                panel1.Visible = true;
                DecorHeightComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            }
        }

        public void NGetDecorWidth()
        {
            if (DecorItemsComboBox.SelectedItem == null)
                return;
            int Length = -1;
            int Height = -1;
            //if (DecorCatalogOrder.ItemLengthBindingSource.Count > 0)
            //    Length = Convert.ToInt32(((DataRowView)DecorCatalogOrder.ItemLengthBindingSource.Current).Row["Length"]);
            //if (DecorCatalogOrder.ItemHeightBindingSource.Count > 0)
            //    Height = Convert.ToInt32(((DataRowView)DecorCatalogOrder.ItemHeightBindingSource.Current).Row["Height"]);
            if (DecorLengthComboBox.SelectedValue != null)
                Length = Convert.ToInt32(DecorLengthComboBox.SelectedValue);
            if (DecorHeightComboBox.SelectedValue != null)
                Height = Convert.ToInt32(DecorHeightComboBox.SelectedValue);
            int W = DecorCatalogOrder.FilterWidth(((DataRowView)DecorItemsComboBox.SelectedItem).Row["Name"].ToString(),
                Convert.ToInt32(DecorColorsComboBox.SelectedValue),
                Convert.ToInt32(DecorPatinaComboBox.SelectedValue),
                Convert.ToInt32(DecorInsetTypesComboBox.SelectedValue),
                Convert.ToInt32(DecorInsetColorsComboBox.SelectedValue), Length, Height);

            DecorWidthComboBox.DropDownStyle = ComboBoxStyle.DropDown;
            if (W == -1)
            {
                panel2.Visible = false;
            }
            else if (W == 0)
            {
                panel2.Visible = true;
                DecorWidthComboBox.DropDownStyle = ComboBoxStyle.DropDown;
            }
            else if (W > 0)
            {
                panel2.Visible = true;
                DecorWidthComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            }
        }

        public void DecorCatalogOrderBinding()
        {
            DecorProductsComboBox.DataSource = DecorCatalogOrder.DecorProductsBindingSource;
            DecorProductsComboBox.DisplayMember = DecorCatalogOrder.DecorProductsBindingSourceDisplayMember;
            DecorProductsComboBox.ValueMember = DecorCatalogOrder.DecorProductsBindingSourceValueMember;

            DecorItemsComboBox.DataSource = DecorCatalogOrder.ItemsBindingSource;
            DecorItemsComboBox.DisplayMember = DecorCatalogOrder.ItemsBindingSourceDisplayMember;
            DecorItemsComboBox.ValueMember = DecorCatalogOrder.ItemsBindingSourceValueMember;

            DecorColorsComboBox.DataSource = DecorCatalogOrder.ItemColorsBindingSource;
            DecorColorsComboBox.DisplayMember = DecorCatalogOrder.ItemColorsBindingSourceDisplayMember;
            DecorColorsComboBox.ValueMember = DecorCatalogOrder.ItemColorsBindingSourceValueMember;

            DecorPatinaComboBox.DataSource = DecorCatalogOrder.ItemPatinaBindingSource;
            DecorPatinaComboBox.DisplayMember = DecorCatalogOrder.ItemPatinaBindingSourceDisplayMember;
            DecorPatinaComboBox.ValueMember = DecorCatalogOrder.ItemPatinaBindingSourceValueMember;

            DecorInsetTypesComboBox.DataSource = DecorCatalogOrder.ItemInsetTypesBindingSource;
            DecorInsetTypesComboBox.DisplayMember = "InsetTypeName";
            DecorInsetTypesComboBox.ValueMember = "InsetTypeID";

            DecorInsetColorsComboBox.DataSource = DecorCatalogOrder.ItemInsetColorsBindingSource;
            DecorInsetColorsComboBox.DisplayMember = "InsetColorName";
            DecorInsetColorsComboBox.ValueMember = "InsetColorID";

            DecorLengthComboBox.DataSource = DecorCatalogOrder.ItemLengthBindingSource;
            DecorLengthComboBox.DisplayMember = "Length";
            DecorLengthComboBox.ValueMember = "Length";

            DecorHeightComboBox.DataSource = DecorCatalogOrder.ItemHeightBindingSource;
            DecorHeightComboBox.DisplayMember = "Height";
            DecorHeightComboBox.ValueMember = "Height";

            DecorWidthComboBox.DataSource = DecorCatalogOrder.ItemWidthBindingSource;
            DecorWidthComboBox.DisplayMember = "Width";
            DecorWidthComboBox.ValueMember = "Width";

            bool bExcluzive = cbOnlyExcluzive.Checked;
            if (!DecorOrders.HasExcluzive)
                bExcluzive = false;
            DecorCatalogOrder.FilterProducts(bExcluzive);
            NGetDecorItems();
            NGetDecorColors();
            NGetDecorPatina();
            NGetDecorInsetTypes();
            NGetDecorInsetColors();
            NGetDecorLength();
            NGetDecorHeight();
            NGetDecorWidth();
        }

        private void DecorProductsComboBox_SelectionChangeCommitted(object sender, EventArgs e)
        {
            NGetDecorItems();
            NGetDecorColors();
            NGetDecorPatina();
            NGetDecorInsetTypes();
            NGetDecorInsetColors();
            NGetDecorLength();
            NGetDecorHeight();
            NGetDecorWidth();
        }

        private void DecorItemsComboBox_SelectionChangeCommitted(object sender, EventArgs e)
        {
            NGetDecorColors();
            NGetDecorPatina();
            NGetDecorInsetTypes();
            NGetDecorInsetColors();
            NGetDecorLength();
            NGetDecorHeight();
            NGetDecorWidth();
        }

        private void DecorColorsComboBox_SelectionChangeCommitted(object sender, EventArgs e)
        {
            NGetDecorPatina();
            NGetDecorInsetTypes();
            NGetDecorInsetColors();
            NGetDecorLength();
            NGetDecorHeight();
            NGetDecorWidth();
        }

        private void DecorHeightComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            NGetDecorWidth();
        }

        private void DecorProductsComboBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                DecorProductsComboBox_SelectionChangeCommitted(sender, e);
                SetAutoComplete(DecorProductsComboBox, true);
            }
        }

        private void DecorItemsComboBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                DecorItemsComboBox_SelectionChangeCommitted(sender, e);
                SetAutoComplete(DecorItemsComboBox, true);
            }
        }

        private void DecorColorsComboBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                DecorColorsComboBox_SelectionChangeCommitted(sender, e);
                SetAutoComplete(DecorColorsComboBox, true);
            }
        }

        private void FrontsComboBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                FrontsComboBox_SelectionChangeCommitted(sender, e);
                SetAutoComplete(FrontsComboBox, true);
            }

        }

        private void FrameColorComboBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                FrameColorComboBox_SelectionChangeCommitted(sender, e);
                SetAutoComplete(FrameColorComboBox, true);
            }
        }

        private void PatinaComboBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                PatinaComboBox_SelectionChangeCommitted(sender, e);
                SetAutoComplete(PatinaComboBox, true);
            }
        }

        private void InsetColorComboBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                InsetColorComboBox_SelectionChangeCommitted(sender, e);
                SetAutoComplete(InsetColorComboBox, true);
            }
        }

        private void InsetTypesComboBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                InsetTypesComboBox_SelectionChangeCommitted(sender, e);
                SetAutoComplete(InsetTypesComboBox, true);
            }
        }

        private void TechnoInsetTypesComboBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                TechnoInsetTypesComboBox_SelectionChangeCommitted(sender, e);
                SetAutoComplete(TechnoInsetTypesComboBox, true);
            }
        }

        private void TechnoInsetColorsComboBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                TechnoInsetColorsComboBox_SelectionChangeCommitted(sender, e);
                SetAutoComplete(TechnoInsetColorsComboBox, true);
            }
        }

        private void SetAutoComplete(ComponentFactory.Krypton.Toolkit.KryptonComboBox ComboBox, bool Autocomplete)
        {
            if (Autocomplete)
            {
                bool focused = ComboBox.Focused;
                if (ComboBox.Focused)
                {
                    ComboBox.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
                    ComboBox.AutoCompleteSource = AutoCompleteSource.ListItems;
                }
            }
            else
            {
                ComboBox.AutoCompleteMode = AutoCompleteMode.None;
                ComboBox.AutoCompleteSource = AutoCompleteSource.None;
            }
        }

        private void FrontAddButton_Click(object sender, EventArgs e)
        {
            int FrontID = -1;
            int ColorID = -1;
            int PatinaID = -1;
            int InsetTypeID = -1;
            int InsetColorID = -1;
            int TechnoProfileID = -1;
            int TechnoColorID = -1;
            int TechnoInsetTypeID = -1;
            int TechnoInsetColorID = -1;
            int Height = -1;
            int Width = -1;
            int Count = -1;
            int ImpostMargin = 0;

            //фасад
            if (FrontsComboBox.SelectedIndex == -1)
            {
                MessageBox.Show(this, "Не выбран фасад или нет в базе", "Ошибка добавления фасада",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            //FrontID = Convert.ToInt32(FrontsComboBox.SelectedValue);

            //цвет профиля
            if (FrameColorComboBox.SelectedIndex == -1)
            {
                MessageBox.Show(this, "Не выбран цвет профиля или нет в базе", "Ошибка добавления фасада",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            ColorID = Convert.ToInt32(FrameColorComboBox.SelectedValue);

            //тип патины
            if (PatinaComboBox.SelectedIndex == -1)
            {
                MessageBox.Show(this, "Не выбран тип патины или нет в базе", "Ошибка добавления фасада",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            PatinaID = Convert.ToInt32(PatinaComboBox.SelectedValue);

            //тип наполнителя
            if (InsetTypesComboBox.SelectedIndex == -1)
            {
                MessageBox.Show(this, "Не выбран тип наполнителя или нет в базе", "Ошибка добавления фасада",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            InsetTypeID = Convert.ToInt32(InsetTypesComboBox.SelectedValue);

            //цвет наполнителя
            if (InsetColorComboBox.SelectedIndex == -1)
            {
                MessageBox.Show(this, "Не выбран цвет наполнителя или нет в базе", "Ошибка добавления фасада",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            InsetColorID = Convert.ToInt32(InsetColorComboBox.SelectedValue);

            //тип профиля-2
            if (TechnoProfilesComboBox.SelectedIndex == -1)
            {
                MessageBox.Show(this, "Не выбран Тип профиля-2 или нет в базе", "Ошибка добавления фасада",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            TechnoProfileID = Convert.ToInt32(TechnoProfilesComboBox.SelectedValue);

            //Цвет профиля-2
            if (TechnoFrameColorComboBox.SelectedIndex == -1)
            {
                MessageBox.Show(this, "Не выбран Цвет профиля-2 или нет в базе", "Ошибка добавления фасада",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            TechnoColorID = Convert.ToInt32(TechnoFrameColorComboBox.SelectedValue);

            //Тип наполнителя-2
            if (TechnoInsetTypesComboBox.SelectedIndex == -1)
            {
                MessageBox.Show(this, "Не выбран Тип наполнителя-2 или нет в базе", "Ошибка добавления фасада",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            TechnoInsetTypeID = Convert.ToInt32(TechnoInsetTypesComboBox.SelectedValue);

            //Цвет наполнителя-2
            if (TechnoInsetColorsComboBox.SelectedIndex == -1)
            {
                MessageBox.Show(this, "Не выбран Цвет наполнителя-2 или нет в базе", "Ошибка добавления фасада",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            TechnoInsetColorID = Convert.ToInt32(TechnoInsetColorsComboBox.SelectedValue);

            //высота
            if (FrontsHeightComboBox.Text.Length < 1)
            {
                MessageBox.Show(this, "Не указана высота фасада", "Ошибка добавления фасада", MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            Height = Convert.ToInt32(FrontsHeightComboBox.Text);

            //ширина
            if (FrontsWidthComboBox.Text.Length < 1)
            {
                MessageBox.Show(this, "Не указана ширина фасада", "Ошибка добавления фасада", MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            Width = Convert.ToInt32(FrontsWidthComboBox.Text);

            //количество
            if (string.IsNullOrWhiteSpace(FrontsCountComboBox.Text) || Convert.ToInt32(FrontsCountComboBox.Text) < 1)
            {
                MessageBox.Show(this, "Неверное количество фасадов", "Ошибка добавления фасада", MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            Count = Convert.ToInt32(FrontsCountComboBox.Text);
            //смещение для импоста
            if (!string.IsNullOrWhiteSpace(cbImpostMargin.Text) && Convert.ToInt32(cbImpostMargin.Text) > 0)
            {
                ImpostMargin = Convert.ToInt32(cbImpostMargin.Text);
            }

            int FactoryID = 0;
            int AreaID = 0;
            FrontsOrders.FrontsCatalogOrder.GetFrontConfigID(
                ((DataRowView)FrontsComboBox.SelectedItem).Row["FrontName"].ToString(), ColorID, PatinaID, InsetTypeID,
                InsetColorID, TechnoProfileID, TechnoColorID, TechnoInsetTypeID, TechnoInsetColorID, Height, Width,
                ref FrontID, ref FactoryID, ref AreaID);
            if (FrontID == -1)
                return;
            
            var clientNonStandard = false;
            var oversize = FrontsOrders.IsOversize(FrontID, Height);
            var clientConfirm = false;
            
            if (oversize)
            {
                clientNonStandard = FrontsOrders.GetClientNonStandard(_ordersManager.CurrentClientID);
                
                string heightText = "Подтвердите добавление размера";

                if (clientNonStandard)
                {
                    var FrontConfigID = FrontsCatalogOrder.GetFrontConfigID(FrontID, ColorID, PatinaID,
                        InsetTypeID, InsetColorID, TechnoProfileID, TechnoColorID, TechnoInsetTypeID, TechnoInsetColorID, Height, Width);
                    decimal extraPrice = FrontsCatalogOrder.GetNonStandardMargin(FrontConfigID);

                    string heightCaption2 = @$"За качество и эксплуатацию указанного размера предприятие ответственности не несет! Всё равно добавить?
За изготовление нестандартного размера взимается наценка {extraPrice}%";

                    DialogResult result = MessageBox.Show(this, heightCaption2,
                        heightText, MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                    if (result == DialogResult.No)
                    {
                        return;
                    }

                    clientConfirm = true;
                }
                else
                {
                    string heightCaption1 = "За качество и эксплуатацию указанного размера предприятие ответственности не несет! Всё равно добавить?";

                    DialogResult result = MessageBox.Show(this, heightCaption1,
                        heightText, MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                    if (result == DialogResult.No)
                    {
                        return;
                    }

                    clientConfirm = true;
                }
            }

            FrontsOrders.AddFrontsOrder(clientConfirm, MainOrderID, FrontID, ColorID, PatinaID, InsetTypeID, InsetColorID,
                TechnoProfileID, TechnoColorID, TechnoInsetTypeID, TechnoInsetColorID, Height, Width, Count,
                FrontsOrdersNotesTextBox.Text, clientNonStandard, ImpostMargin);
            cbImpostMargin.ResetText();
        }

        private void FrontRemoveButton_Click(object sender, EventArgs e)
        {
            FrontsOrders.RemoveOrder();
        }

        private void SaveOrdersSet(int MegaOrderID, int MainOrderID, int OrdersSetCount)
        {
            int[] MainOrdersID = new int[OrdersSetCount];
            MainOrdersID[0] = MainOrderID;

            int FactoryIDF = -1;
            int FactoryIDD = -1;

            for (int i = 1; i < OrdersSetCount; i++)
            {
                if (FrontsOrders.CanSaveFrontsOrder(ref FactoryIDF) == false)
                    return;
                if (DecorOrders.CanSaveDecorOrderSet(ref FactoryIDD) == false)
                    return;
            }

            for (int i = 1; i < OrdersSetCount; i++)
            {
                _ordersManager.CreateNewMainOrder();
                FrontsOrders.FrontsOrdersSet(MainOrderID, _ordersManager.CurrentMainOrderID);
                DecorOrders.DecorOrdersSet(MainOrderID, _ordersManager.CurrentMainOrderID);
                MainOrdersID[i] = _ordersManager.CurrentMainOrderID;
                FrontsOrders.SaveFrontsOrder(_ordersManager.CurrentMainOrderID);
                DecorOrders.SaveDecorOrderSet(_ordersManager.CurrentMainOrderID);
            }

            bool IsSample = false;
            bool isFrontsSample = FrontsOrders.HasSample();
            bool isDecorSample = DecorOrders.HasSample();

            if (isFrontsSample || isDecorSample)
                IsSample = true;

            int F = 0;

            if (FactoryIDF == 1)
                if (FactoryIDD == 1)
                    F = 1;
                else if (FactoryIDD == -1)
                    F = 1;
                else
                    F = 0;

            if (FactoryIDF == 2)
                if (FactoryIDD == 2)
                    F = 2;
                else if (FactoryIDD == -1)
                    F = 2;
                else
                    F = 0;

            if (FactoryIDF == -1)
                F = FactoryIDD;

            _ordersManager.SaveOrder(MainOrderID, MainOrderNotes.Text, IsSample, F);
            decimal DiscountPaymentCondition = _ordersManager.DiscountPaymentCondition(_ordersManager.CurrentDiscountPaymentConditionID);
            _ordersCalculate.CalculateOrder(MegaOrderID, MainOrderID, _ordersManager.CurrentProfilDiscountDirector, _ordersManager.CurrentTPSDiscountDirector,
                _ordersManager.CurrentProfilTotalDiscount, _ordersManager.CurrentTPSTotalDiscount, DiscountPaymentCondition,
                _ordersManager.CurrencyTypeID, _ordersManager.PaymentCurrency, _ordersManager.ConfirmDateTime);

            _ordersManager.OrdersSet(MainOrdersID);
        }

        private void SaveButton_Click(object sender, EventArgs e)
        {
            int FactoryIDF = -1;
            int FactoryIDD = -1;

            if (FrontsOrders.CanSaveFrontsOrder(ref FactoryIDF) == false)
                return;
            if (DecorOrders.CanSaveDecorOrder(ref FactoryIDD) == false)
                return;
            if (!EditMainOrder)
            {
                _ordersManager.CreateNewMainOrder();
                MainOrderID = _ordersManager.CurrentMainOrderID;
            }

            int MegaOrderID = _ordersManager.CurrentMegaOrderID;

            int F = 0;

            if (FrontsOrders.HasZeroCount || DecorOrders.HasZeroCount)
            {
                MessageBox.Show("Кол-во не может быть равно 0");
                return;
            }

            if (FrontsOrders.HasRows() == false && DecorOrders.HasRows() == false)
            {
                return;
            }

            FrontsOrders.SaveFrontsOrder(MainOrderID);
            DecorOrders.SaveDecorOrder(MainOrderID);
            DecorOrders.DeleteDecorAssignmentByDecorOrders(MainOrderID);

            bool IsSample = false;
            bool isFrontsSample = FrontsOrders.HasSample();
            bool isDecorSample = DecorOrders.HasSample();

            if (isFrontsSample || isDecorSample)
                IsSample = true;

            if (FactoryIDF == 1)
                if (FactoryIDD == 1)
                    F = 1;
                else if (FactoryIDD == -1)
                    F = 1;
                else
                    F = 0;

            if (FactoryIDF == 2)
                if (FactoryIDD == 2)
                    F = 2;
                else if (FactoryIDD == -1)
                    F = 2;
                else
                    F = 0;

            if (FactoryIDF == -1)
                F = FactoryIDD;

            //Если комплект подзаказов
            if (OrdersSetCheckBox.Checked)
            {
                if (string.IsNullOrWhiteSpace(OrdersSetTextBox.Text))
                {
                    MessageBox.Show("Введите размер комплекта");
                    return;
                }
                else
                {
                    SaveOrdersSet(MegaOrderID, MainOrderID, Convert.ToInt32(OrdersSetTextBox.Text));
                }
            }
            else
            {
                //if (OrdersManager.NeedSetStatus)
                _ordersManager.SaveOrder(MainOrderNotes.Text, IsSample, F, _ordersManager.NeedSetStatus);
                decimal DiscountPaymentCondition = _ordersManager.DiscountPaymentCondition(_ordersManager.CurrentDiscountPaymentConditionID);
                if (_ordersManager.CurrentDiscountPaymentConditionID == 4)
                {
                    DiscountPaymentCondition = _ordersManager.DiscountFactoring(_ordersManager.CurrentDiscountFactoringID);
                }
                
                _ordersCalculate.ConfirmDateTime = _ordersManager.ConfirmDateTime == DBNull.Value ? DateTime.Now : Convert.ToDateTime(_ordersManager.ConfirmDateTime);


                Tuple<bool, decimal, decimal, decimal> clientRates = 
                    _ordersCalculate.GetFixedPaymentRate(_ordersManager.CurrentClientID, _ordersCalculate.ConfirmDateTime);

                Tuple<bool, decimal, decimal, decimal, decimal> dateRates =
                    _ordersManager.GetDateRates(_ordersCalculate.ConfirmDateTime);

                bool fixedClientRate = clientRates.Item1;

                if (fixedClientRate)
                {
                    switch (_ordersManager.CurrencyTypeID)
                    {
                        case 1:
                            if (dateRates.Item1)
                                try
                                {
                                    _ordersCalculate.Rate = clientRates.Item4 / dateRates.Item4;
                                }
                                catch (DivideByZeroException exception)
                                {
                                    MessageBox.Show($"Параметр BYN равен 0. Исправьте в Редакторе курсов валют. Расчет не выполнен");
                                    Console.WriteLine(exception);
                                }
                                catch (Exception exception)
                                {
                                    MessageBox.Show($"Ошибка. Расчет не выполнен");
                                    Console.WriteLine(exception);
                                }

                            break;
                        case 2:
                            _ordersCalculate.Rate = clientRates.Item2;
                            break;
                        case 3:
                            _ordersCalculate.Rate = clientRates.Item3;
                            break;
                        case 5:
                            _ordersCalculate.Rate = clientRates.Item4;
                            break;
                        default:
                            break;
                    }
                }
                else
                {
                    if (dateRates.Item1)
                    {
                        switch (_ordersManager.CurrencyTypeID)
                        {
                            case 1:
                                _ordersCalculate.Rate = 1;
                                break;
                            case 2:
                                _ordersCalculate.Rate = dateRates.Item2;
                                break;
                            case 3:
                                _ordersCalculate.Rate = dateRates.Item3;
                                break;
                            case 5:
                                _ordersCalculate.Rate = dateRates.Item4;
                                break;
                            default:
                                break;
                        }
                    }

                }

                _ordersCalculate.Recalculate(MegaOrderID, _ordersManager.CurrentProfilDiscountDirector, _ordersManager.CurrentTPSDiscountDirector,
                    _ordersManager.CurrentProfilTotalDiscount, _ordersManager.CurrentTPSTotalDiscount, DiscountPaymentCondition, _ordersManager.CurrencyTypeID,
                    _ordersCalculate.Rate, _ordersCalculate.ConfirmDateTime);
                InvoiceReportToDbf DBFReport = new InvoiceReportToDbf(FrontsCatalogOrder, DecorCatalogOrder);
                decimal CurrencyTotalCost = DBFReport.CalcCurrencyCost(
                    Convert.ToInt32(((DataRowView)_ordersManager.MegaOrdersBindingSource.Current).Row["MegaOrderID"]),
                    Convert.ToInt32(((DataRowView)_ordersManager.MegaOrdersBindingSource.Current).Row["ClientID"]), _ordersCalculate.Rate);
                _ordersManager.SetCurrencyCost(Convert.ToInt32(((DataRowView)_ordersManager.MegaOrdersBindingSource.Current).Row["MegaOrderID"]), CurrencyTotalCost);
            }

            if (_ordersManager.NeedSetStatus)
            {
                //OrdersManager.SetCurrentMegaOrderStatus();
                CheckOrdersStatus.SetNewMegaOrderStatus(MegaOrderID);
            }

            if (!EditMainOrder)
                _ordersManager.FixOrderEvent(MegaOrderID, "Добавлен позаказ");
            else
                _ordersManager.FixOrderEvent(MegaOrderID, "Отредактирован подзаказ");
            //OrdersManager.SummaryCost(OrdersManager.CurrentMegaOrderID);
            InfiniumTips.ShowTip(this, 50, 85, "Заказ сохранён", 1700);
            FormEvent = eClose;
            AnimateTimer.Enabled = true;
        }

        public void AddFrontsCupboards(ref ComponentFactory.Krypton.Toolkit.KryptonListBox CupboardsExportListBox)
        {
            int FrontID = -1;
            int PatinaID = -1;
            int InsetTypeID = -1;
            int ColorID = -1;
            int InsetColorID = -1;
            int TechnoProfileID = -1;
            int TechnoColorID = -1;
            int TechnoInsetTypeID = -1;
            int TechnoInsetColorID = -1;
            int Height = 0;
            int Width = 0;
            int Count = 1;

            //фасад
            if (FrontsComboBox.SelectedIndex == -1)
            {
                Infinium.LightMessageBox.Show(ref TopForm, false,
                    "Не выбран фасад или нет в базе", "Ошибка добавления фасада");

                return;
            }

            //FrontID = Convert.ToInt32(FrontsComboBox.SelectedValue);


            //цвет профиля
            if (FrameColorComboBox.SelectedIndex == -1)
            {
                Infinium.LightMessageBox.Show(ref TopForm, false,
                    "Не выбран цвет профиля или нет в базе", "Ошибка добавления фасада");

                return;
            }

            ColorID = Convert.ToInt32(FrameColorComboBox.SelectedValue);

            //тип патины
            if (PatinaComboBox.SelectedIndex == -1)
            {
                Infinium.LightMessageBox.Show(ref TopForm, false,
                    "Не выбран тип фасада или нет в базе", "Ошибка добавления фасада");

                return;
            }

            PatinaID = Convert.ToInt32(PatinaComboBox.SelectedValue);


            //тип наполнителя
            if (InsetTypesComboBox.SelectedIndex == -1)
            {
                Infinium.LightMessageBox.Show(ref TopForm, false,
                    "Не выбран тип наполнителя или нет в базе", "Ошибка добавления фасада");

                return;
            }

            InsetTypeID = Convert.ToInt32(InsetTypesComboBox.SelectedValue);


            //цвет наполнителя
            if (InsetColorComboBox.SelectedIndex == -1)
            {
                Infinium.LightMessageBox.Show(ref TopForm, false,
                    "Не выбран цвет наполнителя или нет в базе", "Ошибка добавления фасада");

                return;
            }

            InsetColorID = Convert.ToInt32(InsetColorComboBox.SelectedValue);

            //тип профиля-2
            if (TechnoProfilesComboBox.SelectedIndex == -1)
            {
                MessageBox.Show(this, "Не выбран Тип профиля-2 или нет в базе", "Ошибка добавления фасада", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            TechnoProfileID = Convert.ToInt32(TechnoProfilesComboBox.SelectedValue);

            //Цвет профиля-2
            if (TechnoFrameColorComboBox.SelectedIndex == -1)
            {
                MessageBox.Show(this, "Не выбран Цвет профиля-2 или нет в базе", "Ошибка добавления фасада", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            TechnoColorID = Convert.ToInt32(TechnoFrameColorComboBox.SelectedValue);

            //Тип наполнителя-2
            if (TechnoInsetTypesComboBox.SelectedIndex == -1)
            {
                MessageBox.Show(this, "Не выбран Тип наполнителя-2 или нет в базе", "Ошибка добавления фасада", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            TechnoInsetTypeID = Convert.ToInt32(TechnoInsetTypesComboBox.SelectedValue);

            //цвет наполнителя техно
            if (TechnoInsetColorsComboBox.SelectedIndex == -1)
            {
                MessageBox.Show(this, "Не выбран цвет наполнителя техно или нет в базе", "Ошибка добавления фасада", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            TechnoInsetColorID = Convert.ToInt32(TechnoInsetColorsComboBox.SelectedValue);

            if (FrontsCountComboBox.Text.Length > 0)
                Count = Convert.ToInt32(FrontsCountComboBox.Text);

            int FactoryID = 0;
            int AreaID = 0;
            FrontsOrders.FrontsCatalogOrder.GetFrontConfigID(((DataRowView)FrontsComboBox.SelectedItem).Row["FrontName"].ToString(), ColorID, PatinaID, InsetTypeID,
                   InsetColorID, TechnoProfileID, TechnoColorID, TechnoInsetTypeID, TechnoInsetColorID, Height, Width, ref FrontID, ref FactoryID, ref AreaID);
            if (FrontID == -1)
                return;
            FrontsOrders.AddFrontsOrderCupboards(ref CupboardsExportListBox, MainOrderID, FrontID, ColorID, PatinaID, InsetTypeID,
                InsetColorID, TechnoProfileID, TechnoColorID, TechnoInsetTypeID, TechnoInsetColorID, Height, Width, Count);
        }

        private void DecorAddButton_Click(object sender, EventArgs e)
        {
            int ProductID = Convert.ToInt32(DecorProductsComboBox.SelectedValue);
            int ItemID = -1;
            int ColorID = -1;
            int PatinaID = -1;
            int InsetTypeID = -1;
            int InsetColorID = -1;
            int Height = -1;
            int Length = -1;
            int Width = -1;
            int Count = 0;

            if (panel31.Visible)
                if (DecorColorsComboBox.Text != "")
                    ColorID = Convert.ToInt32(DecorColorsComboBox.SelectedValue);
                else
                {
                    MessageBox.Show("Введены не все параметры продукта - цвет");
                    return;
                }
            if (panel29.Visible)
                if (DecorPatinaComboBox.Text != "")
                    PatinaID = Convert.ToInt32(DecorPatinaComboBox.SelectedValue);
                else
                {
                    MessageBox.Show("Введены не все параметры продукта - патина");
                    return;
                }
            if (panel5.Visible)
                if (DecorInsetTypesComboBox.Text != "")
                    InsetTypeID = Convert.ToInt32(DecorInsetTypesComboBox.SelectedValue);
                else
                {
                    MessageBox.Show("Введены не все параметры продукта - тип наполнителя");
                    return;
                }
            if (panel6.Visible)
                if (DecorInsetColorsComboBox.Text != "")
                    InsetColorID = Convert.ToInt32(DecorInsetColorsComboBox.SelectedValue);
                else
                {
                    MessageBox.Show("Введены не все параметры продукта - цвет наполнителя");
                    return;
                }

            if (panel4.Visible)
                if (DecorLengthComboBox.Text != "")
                    Length = Convert.ToInt32(DecorLengthComboBox.Text);
                else
                {
                    MessageBox.Show("Введены не все параметры продукта - длина");
                    return;
                }

            if (panel1.Visible)
                if (DecorHeightComboBox.Text != "")
                    Height = Convert.ToInt32(DecorHeightComboBox.Text);
                else
                {
                    MessageBox.Show("Введены не все параметры продукта - высота");
                    return;
                }

            if (panel2.Visible)
                if (DecorWidthComboBox.Text != "")
                    Width = Convert.ToInt32(DecorWidthComboBox.Text);
                else
                {
                    MessageBox.Show("Введены не все параметры продукта - ширина");
                    return;
                }


            //количество
            if (DecorCountNumUpDown.Text.Length < 1 || Convert.ToInt32(DecorCountNumUpDown.Text) < 1)
            {
                MessageBox.Show(this, "Неверное количество", "Ошибка добавления декора", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Count = Convert.ToInt32(DecorCountNumUpDown.Text);

            int FactoryID = 0;
            int AreaID = 0;
            DecorOrders.DecorCatalogOrder.GetDecorConfigID(ProductID,
                ((DataRowView)DecorItemsComboBox.SelectedItem).Row["Name"].ToString(), ColorID, PatinaID, InsetTypeID, InsetColorID,
                Length, Height, Width, ref ItemID, ref FactoryID, ref AreaID);
            if (ItemID == -1)
                return;
            DecorOrders.AddDecorOrder(MainOrderID, ProductID, ItemID, ColorID, PatinaID, InsetTypeID, InsetColorID, Length, Height, Width, Count, DecorOrdersNotesTextBox.Text);
            NoDecorLabel.Visible = false;
        }

        private void DecorRemoveButton_Click(object sender, EventArgs e)
        {
            if (MainOrdersDecorTabControl.SelectedTabPageIndex == -1)
                return;

            DecorOrders.DeleteCurrentDecorItem(MainOrdersDecorTabControl.SelectedTabPageIndex);

            if (MainOrdersDecorTabControl.SelectedTabPage == null)
                NoDecorLabel.Visible = true;
        }

        private void kryptonButton2_Click(object sender, EventArgs e)
        {
            DecorOrders.DecorItemOrdersDataGrids[0].Size.Width.ToString();
        }

        private void kryptonCheckSet1_CheckedButtonChanged(object sender, EventArgs e)
        {
            if (FrontsCheckButton.Checked)
            {
                FrontsPanel1.BringToFront();
                MainOrdersTabControl.SelectedTabPageIndex = 0;

                FrontsComboBox.SelectionLength = 0;
                FrameColorComboBox.SelectionLength = 0;
                TechnoProfilesComboBox.SelectionLength = 0;
                TechnoFrameColorComboBox.SelectionLength = 0;
                PatinaComboBox.SelectionLength = 0;
                InsetTypesComboBox.SelectionLength = 0;
                InsetColorComboBox.SelectionLength = 0;
                TechnoInsetTypesComboBox.SelectionLength = 0;
                TechnoInsetColorsComboBox.SelectionLength = 0;
                FrontsComboBox.Focus();
            }
            else
            {
                MainOrdersTabControl.SelectedTabPageIndex = 1;
                DecorPanel1.BringToFront();
                DecorProductsComboBox.SelectionLength = 0;
                DecorItemsComboBox.SelectionLength = 0;
                DecorColorsComboBox.SelectionLength = 0;
                DecorPatinaComboBox.SelectionLength = 0;
                DecorInsetTypesComboBox.SelectionLength = 0;
                DecorInsetColorsComboBox.SelectionLength = 0;
                DecorProductsComboBox.Focus();
            }
        }

        private void MainOrdersFrontsOrdersDataGrid_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            if (FrontsOrdersDataGrid.RowCount > 0)
                FrontsOrdersDataGrid.Visible = true;
        }

        private void FrontsOrdersDataGrid_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            if (FrontsOrders.FrontsOrdersBindingSource.Count < 1)
                FrontsOrdersDataGrid.Visible = false;
        }

        private void FrontsHeightComboBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Left)
            {
                InsetColorComboBox.Focus();

                e.SuppressKeyPress = true;
            }

            if (e.KeyCode == Keys.Right)
            {
                FrontsWidthComboBox.Focus();
            }
        }

        private void FrontsWidthComboBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Left)
            {
                FrontsHeightComboBox.Focus();
            }

            if (e.KeyCode == Keys.Right)
            {
                FrontsCountComboBox.Focus();
            }
        }

        private void FrontsCountComboBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                FrontAddButton_Click(null, null);
                FrontsHeightComboBox.Focus();
            }

            //if (e.KeyCode == Keys.Left)
            //{
            //    if (FrontsWidthComboBox.Enabled)
            //        FrontsWidthComboBox.Focus();
            //    else
            //        FrontsHeightComboBox.Focus();
            //}
        }

        private void FrontsComboBox_Leave(object sender, EventArgs e)
        {
            SetAutoComplete(FrontsComboBox, false);
        }

        private void PatinaComboBox_Leave(object sender, EventArgs e)
        {
            SetAutoComplete(PatinaComboBox, false);
        }

        private void InsetTypesComboBox_Leave(object sender, EventArgs e)
        {
            SetAutoComplete(InsetTypesComboBox, false);
        }

        private void FrameColorComboBox_Leave(object sender, EventArgs e)
        {
            SetAutoComplete(FrameColorComboBox, false);
        }

        private void InsetColorComboBox_Leave(object sender, EventArgs e)
        {
            SetAutoComplete(InsetColorComboBox, false);
        }

        private void TechnoInsetTypesComboBox_Leave(object sender, EventArgs e)
        {
            SetAutoComplete(TechnoInsetTypesComboBox, false);
        }

        private void TechnoInsetColorsComboBox_Leave(object sender, EventArgs e)
        {
            SetAutoComplete(TechnoInsetColorsComboBox, false);
        }

        private void FrontsComboBox_Enter(object sender, EventArgs e)
        {
            FrontsComboBox.Focus();
            SetAutoComplete(FrontsComboBox, true);
        }

        private void PatinaComboBox_Enter(object sender, EventArgs e)
        {
            PatinaComboBox.Focus();
            SetAutoComplete(PatinaComboBox, true);
        }

        private void InsetTypesComboBox_Enter(object sender, EventArgs e)
        {
            InsetTypesComboBox.Focus();
            SetAutoComplete(InsetTypesComboBox, true);
        }

        private void FrameColorComboBox_Enter(object sender, EventArgs e)
        {
            FrameColorComboBox.Focus();
            SetAutoComplete(FrameColorComboBox, true);
        }

        private void InsetColorComboBox_Enter(object sender, EventArgs e)
        {
            InsetColorComboBox.Focus();
            SetAutoComplete(InsetColorComboBox, true);
        }

        private void TechnoInsetTypesComboBox_Enter(object sender, EventArgs e)
        {
            TechnoInsetTypesComboBox.Focus();
            SetAutoComplete(TechnoInsetTypesComboBox, true);
        }

        private void TechnoInsetColorsComboBox_Enter(object sender, EventArgs e)
        {
            TechnoInsetColorsComboBox.Focus();
            SetAutoComplete(TechnoInsetColorsComboBox, true);
        }

        private void FrontsHeightComboBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar) && !char.IsPunctuation(e.KeyChar))
                e.Handled = true;
        }

        private void FrontsWidthComboBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar) && !char.IsPunctuation(e.KeyChar))
                e.Handled = true;
        }

        private void FrontsCountComboBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar) && !char.IsPunctuation(e.KeyChar))
                e.Handled = true;
        }

        private void DecorHeightComboBox_KeyDown(object sender, KeyEventArgs e)
        {

        }

        private void DecorWidthComboBox_KeyDown(object sender, KeyEventArgs e)
        {
        }

        private void DecorCountNumUpDown_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                DecorAddButton_Click(null, null);
                DecorProductsComboBox.Focus();
            }
        }

        private void DecorProductsComboBox_Enter(object sender, EventArgs e)
        {
            DecorProductsComboBox.Focus();
            SetAutoComplete(DecorProductsComboBox, true);
        }

        private void DecorItemsComboBox_Enter(object sender, EventArgs e)
        {
            DecorItemsComboBox.Focus();
            SetAutoComplete(DecorItemsComboBox, true);
        }

        private void DecorColorsComboBox_Enter(object sender, EventArgs e)
        {
            DecorColorsComboBox.Focus();

            SetAutoComplete(DecorColorsComboBox, true);
        }

        private void DecorProductsComboBox_Leave(object sender, EventArgs e)
        {
            SetAutoComplete(DecorProductsComboBox, false);
        }

        private void DecorItemsComboBox_Leave(object sender, EventArgs e)
        {
            SetAutoComplete(DecorItemsComboBox, false);
        }

        private void DecorColorsComboBox_Leave(object sender, EventArgs e)
        {
            SetAutoComplete(DecorColorsComboBox, false);
        }

        private void DecorHeightComboBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar) && !char.IsPunctuation(e.KeyChar))
                e.Handled = true;
        }

        private void DecorWidthComboBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar) && !char.IsPunctuation(e.KeyChar))
                e.Handled = true;
        }
        //e03d86502ae9876afa62fbd370a83f5d
        private void DecorCountNumUpDown_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar) && !char.IsPunctuation(e.KeyChar))
                e.Handled = true;
        }

        private void OrdersSetCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (OrdersSetCheckBox.Checked)
            {
                OrdersSetTextBox.Enabled = true;
                OrdersSetTextBox.Focus();
            }
            else
            {
                OrdersSetTextBox.Clear();
                OrdersSetTextBox.Enabled = false;
            }
        }

        private void OrdersSetTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar) && !char.IsPunctuation(e.KeyChar))
                e.Handled = true;
        }

        private void SampleDiscountTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar) && !char.IsPunctuation(e.KeyChar))
                e.Handled = true;
        }

        private void FrontsDropButton_DropDown(object sender, ComponentFactory.Krypton.Toolkit.ContextPositionMenuArgs e)
        {
            FrontsComboBox.Focus();
            FrontsComboBox.DroppedDown = true;
        }

        private void PatinaDropButton_DropDown(object sender, ComponentFactory.Krypton.Toolkit.ContextPositionMenuArgs e)
        {
            PatinaComboBox.Focus();
            PatinaComboBox.DroppedDown = true;
        }

        private void InsetTypesDropButton_DropDown(object sender, ComponentFactory.Krypton.Toolkit.ContextPositionMenuArgs e)
        {
            InsetTypesComboBox.Focus();
            InsetTypesComboBox.DroppedDown = true;
        }

        private void FrameColorsDropButton_DropDown(object sender, ComponentFactory.Krypton.Toolkit.ContextPositionMenuArgs e)
        {
            FrameColorComboBox.Focus();
            FrameColorComboBox.DroppedDown = true;
        }

        private void InsetColorsDropButton_DropDown(object sender, ComponentFactory.Krypton.Toolkit.ContextPositionMenuArgs e)
        {
            InsetColorComboBox.Focus();
            InsetColorComboBox.DroppedDown = true;
        }

        private void TechnoInsetTypesDropButton_DropDown(object sender, ComponentFactory.Krypton.Toolkit.ContextPositionMenuArgs e)
        {
            TechnoInsetTypesComboBox.Focus();
            TechnoInsetTypesComboBox.DroppedDown = true;
        }

        private void DecorProductDropButton_DropDown(object sender, ComponentFactory.Krypton.Toolkit.ContextPositionMenuArgs e)
        {
            DecorProductsComboBox.Focus();
            DecorProductsComboBox.DroppedDown = true;
        }

        private void DecorItemDropButton_DropDown(object sender, ComponentFactory.Krypton.Toolkit.ContextPositionMenuArgs e)
        {
            DecorItemsComboBox.Focus();
            DecorItemsComboBox.DroppedDown = true;
        }

        private void DecorColorDropButton_DropDown(object sender, ComponentFactory.Krypton.Toolkit.ContextPositionMenuArgs e)
        {
            DecorColorsComboBox.Focus();
            DecorColorsComboBox.DroppedDown = true;

        }

        private void CupboardsImportButton_Click(object sender, EventArgs e)
        {
            int ImportType = -1;

            if (CupboardsRadioButton.Checked)
                ImportType = 0;
            else
                ImportType = 1;

            PhantomForm PhantomForm = new Infinium.PhantomForm();
            PhantomForm.Show();

            MCupboardsExportForm CupboardsExportForm = new MCupboardsExportForm(this, ref FrontsOrders);

            CupboardsExportForm.SetImportType(ImportType);

            TopForm = CupboardsExportForm;
            CupboardsExportForm.ShowDialog();

            PhantomForm.Close();

            PhantomForm.Dispose();
            CupboardsExportForm.Dispose();
            TopForm = null;
        }

        public void AddFrontsFromSizeTable(ref PercentageDataGrid DataGrid)
        {
            int MainOrderID = -1;
            //int FrontID = -1;
            int ColorID = -1;
            int PatinaID = -1;
            int InsetTypeID = -1;
            int InsetColorID = -1;
            int TechnoProfileID = -1;
            int TechnoColorID = -1;
            int TechnoInsetTypeID = -1;
            int TechnoInsetColorID = -1;
            //int Height = 0;
            //int Width = 0;

            //фасад
            //if (FrontsComboBox.SelectedIndex == -1)
            //{
            //    MessageBox.Show(this, "Не выбран фасад или нет в базе", "Ошибка добавления фасада", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    return;
            //}
            //FrontID = Convert.ToInt32(FrontsComboBox.SelectedValue);

            //цвет профиля
            if (FrameColorComboBox.SelectedIndex == -1)
            {
                MessageBox.Show(this, "Не выбран цвет профиля или нет в базе", "Ошибка добавления фасада", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            ColorID = Convert.ToInt32(FrameColorComboBox.SelectedValue);

            //тип патины
            if (PatinaComboBox.SelectedIndex == -1)
            {
                MessageBox.Show(this, "Не выбран тип патины или нет в базе", "Ошибка добавления фасада", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            PatinaID = Convert.ToInt32(PatinaComboBox.SelectedValue);

            //тип наполнителя
            if (InsetTypesComboBox.SelectedIndex == -1)
            {
                MessageBox.Show(this, "Не выбран тип наполнителя или нет в базе", "Ошибка добавления фасада", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            InsetTypeID = Convert.ToInt32(InsetTypesComboBox.SelectedValue);

            //цвет наполнителя
            if (InsetColorComboBox.SelectedIndex == -1)
            {
                MessageBox.Show(this, "Не выбран цвет наполнителя или нет в базе", "Ошибка добавления фасада", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            InsetColorID = Convert.ToInt32(InsetColorComboBox.SelectedValue);

            //тип профиля-2
            if (TechnoProfilesComboBox.SelectedIndex == -1)
            {
                MessageBox.Show(this, "Не выбран Тип профиля-2 или нет в базе", "Ошибка добавления фасада", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            TechnoProfileID = Convert.ToInt32(TechnoProfilesComboBox.SelectedValue);

            //Цвет профиля-2
            if (TechnoFrameColorComboBox.SelectedIndex == -1)
            {
                MessageBox.Show(this, "Не выбран Цвет профиля-2 или нет в базе", "Ошибка добавления фасада", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            TechnoColorID = Convert.ToInt32(TechnoFrameColorComboBox.SelectedValue);

            //Тип наполнителя-2
            if (TechnoInsetTypesComboBox.SelectedIndex == -1)
            {
                MessageBox.Show(this, "Не выбран Тип наполнителя-2 или нет в базе", "Ошибка добавления фасада", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            TechnoInsetTypeID = Convert.ToInt32(TechnoInsetTypesComboBox.SelectedValue);

            //Цвет наполнителя-2
            if (TechnoInsetColorsComboBox.SelectedIndex == -1)
            {
                MessageBox.Show(this, "Не выбран Цвет наполнителя-2 или нет в базе", "Ошибка добавления фасада", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            TechnoInsetColorID = Convert.ToInt32(TechnoInsetColorsComboBox.SelectedValue);
            //int FactoryID = 0;
            FrontsOrders.AddFrontsOrderFromSizeTable(ref DataGrid, MainOrderID, ((DataRowView)FrontsComboBox.SelectedItem).Row["FrontName"].ToString(),
                ColorID, PatinaID, InsetTypeID, InsetColorID, TechnoProfileID, TechnoColorID, TechnoInsetTypeID, TechnoInsetColorID);
        }

        public void AddDecorFromSizeTable(ref PercentageDataGrid DataGrid)
        {
            int MainOrderID = -1;
            int ProductID = Convert.ToInt32(DecorProductsComboBox.SelectedValue);
            int ItemID = -1;
            int ColorID = 0;
            int PatinaID = 0;
            int InsetTypeID = 0;
            int InsetColorID = 0;
            int Length = 0;
            int Height = 0;
            int Width = 0;
            int FactoryID = 0;

            if (panel31.Visible)
                if (DecorColorsComboBox.Text != "")
                    ColorID = Convert.ToInt32(DecorColorsComboBox.SelectedValue);
                else
                {
                    MessageBox.Show("Введены не все параметры продукта - цвет");
                    return;
                }
            if (panel29.Visible)
                if (DecorPatinaComboBox.Text != "")
                    PatinaID = Convert.ToInt32(DecorPatinaComboBox.SelectedValue);
                else
                {
                    MessageBox.Show("Введены не все параметры продукта - патина");
                    return;
                }
            if (panel5.Visible)
                if (DecorInsetTypesComboBox.Text != "")
                    InsetTypeID = Convert.ToInt32(DecorInsetTypesComboBox.SelectedValue);
                else
                {
                    MessageBox.Show("Введены не все параметры продукта - тип наполнителя");
                    return;
                }
            if (panel6.Visible)
                if (DecorInsetColorsComboBox.Text != "")
                    InsetColorID = Convert.ToInt32(DecorInsetColorsComboBox.SelectedValue);
                else
                {
                    MessageBox.Show("Введены не все параметры продукта - цвет наполнителя");
                    return;
                }

            int AreaID = 0;
            DecorOrders.DecorCatalogOrder.GetDecorConfigID(ProductID,
                ((DataRowView)DecorItemsComboBox.SelectedItem).Row["Name"].ToString(), ColorID, PatinaID, InsetTypeID, InsetColorID,
                Length, Height, Width, ref ItemID, ref FactoryID, ref AreaID);
            if (ItemID == -1)
                return;
            DecorOrders.AddDecorOrderFromSizeTable(
                ref DataGrid, MainOrderID, ProductID, ItemID, ColorID, PatinaID, InsetTypeID, InsetColorID, DecorOrdersNotesTextBox.Text);
            NoDecorLabel.Visible = false;
        }

        private void MainOrdersTabControl_SelectedPageChanged(object sender, DevExpress.XtraTab.TabPageChangedEventArgs e)
        {
            if (MainOrdersTabControl.SelectedTabPageIndex == 0)
                kryptonCheckSet1.CheckedButton = FrontsCheckButton;
            if (MainOrdersTabControl.SelectedTabPageIndex == 1)
                kryptonCheckSet1.CheckedButton = DecorCheckButton;
        }

        private void DecorImportButton_Click(object sender, EventArgs e)
        {
            PhantomForm PhantomForm = new Infinium.PhantomForm();
            PhantomForm.Show();

            DecorSizesExportForm DecorSizesExportForm = new DecorSizesExportForm(this, ref DecorOrders);

            //CupboardsExportForm.SetImportType(ImportType);

            TopForm = DecorSizesExportForm;
            DecorSizesExportForm.ShowDialog();

            PhantomForm.Close();

            PhantomForm.Dispose();
            DecorSizesExportForm.Dispose();
            TopForm = null;
        }

        private void TechnoInsetColorsDropButton_DropDown(object sender, ComponentFactory.Krypton.Toolkit.ContextPositionMenuArgs e)
        {
            TechnoInsetColorsComboBox.Focus();
            TechnoInsetColorsComboBox.DroppedDown = true;
        }

        private void TechnoProfilesDropButton_DropDown(object sender, ComponentFactory.Krypton.Toolkit.ContextPositionMenuArgs e)
        {
            TechnoProfilesComboBox.Focus();
            TechnoProfilesComboBox.DroppedDown = true;
        }

        private void TechnoFrameColorsDropButton_DropDown(object sender, ComponentFactory.Krypton.Toolkit.ContextPositionMenuArgs e)
        {
            TechnoFrameColorComboBox.Focus();
            TechnoFrameColorComboBox.DroppedDown = true;
        }

        private void TechnoProfilesComboBox_Enter(object sender, EventArgs e)
        {
            TechnoProfilesComboBox.Focus();
            SetAutoComplete(TechnoProfilesComboBox, true);
        }

        private void TechnoFrameColorComboBox_Enter(object sender, EventArgs e)
        {
            TechnoFrameColorComboBox.Focus();
            SetAutoComplete(TechnoFrameColorComboBox, true);
        }

        private void TechnoProfilesComboBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                TechnoProfilesComboBox_SelectionChangeCommitted(sender, e);
                SetAutoComplete(TechnoProfilesComboBox, true);
            }
        }

        private void TechnoFrameColorComboBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                TechnoFrameColorComboBox_SelectionChangeCommitted(sender, e);
                SetAutoComplete(TechnoFrameColorComboBox, true);
            }
        }

        private void TechnoProfilesComboBox_Leave(object sender, EventArgs e)
        {
            SetAutoComplete(TechnoProfilesComboBox, false);
        }

        private void TechnoFrameColorComboBox_Leave(object sender, EventArgs e)
        {
            SetAutoComplete(TechnoFrameColorComboBox, false);
        }

        private void DecorPatinaDropButton_DropDown(object sender, ComponentFactory.Krypton.Toolkit.ContextPositionMenuArgs e)
        {
            DecorPatinaComboBox.Focus();
            DecorPatinaComboBox.DroppedDown = true;
        }

        private void DecorPatinaComboBox_Enter(object sender, EventArgs e)
        {
            DecorPatinaComboBox.Focus();
            SetAutoComplete(DecorPatinaComboBox, true);
        }

        private void DecorPatinaComboBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                DecorPatinaComboBox_SelectionChangeCommitted(sender, e);
                SetAutoComplete(DecorPatinaComboBox, true);
            }
        }

        private void DecorPatinaComboBox_Leave(object sender, EventArgs e)
        {
            SetAutoComplete(DecorPatinaComboBox, false);
        }

        private void DecorPatinaComboBox_SelectionChangeCommitted(object sender, EventArgs e)
        {
            NGetDecorInsetTypes();
            NGetDecorInsetColors();
            NGetDecorLength();
            NGetDecorHeight();
            NGetDecorWidth();
        }

        private void DecorInsetTypesComboBox_SelectionChangeCommitted(object sender, EventArgs e)
        {
            NGetDecorInsetColors();
            NGetDecorLength();
            NGetDecorHeight();
            NGetDecorWidth();
        }

        private void DecorInsetColorsComboBox_SelectionChangeCommitted(object sender, EventArgs e)
        {
            NGetDecorLength();
            NGetDecorHeight();
            NGetDecorWidth();
        }

        private void DecorLengthComboBox_KeyDown(object sender, KeyEventArgs e)
        {

        }

        private void DecorLengthComboBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar) && !char.IsPunctuation(e.KeyChar))
                e.Handled = true;
        }

        private void DecorLengthComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            NGetDecorHeight();
            NGetDecorWidth();
        }

        private void cbOnlyExcluzive_CheckedChanged(object sender, EventArgs e)
        {
            if (FrontsOrders.HasExcluzive)
            {
                FrontsOrders.UpdateExcluziveCatalog();
                FrontsCatalogOrder.FilterFronts(cbOnlyExcluzive.Checked);

                NGetFrameColors();
                NGetTechnoProfiles();
                NGetTechnoFrameColors();
                NGetPatina();
                NGetInsetTypes();
                NGetInsetColors();
                NGetTechnoInsetTypes();
                NGetTechnoInsetColors();
                NGetHeight();
                NGetWidth();
            }
            if (DecorOrders.HasExcluzive)
            {
                DecorOrders.UpdateExcluziveCatalog();
                DecorCatalogOrder.FilterProducts(cbOnlyExcluzive.Checked);

                NGetDecorItems();
                NGetDecorColors();
                NGetDecorPatina();
                NGetDecorInsetTypes();
                NGetDecorInsetColors();
                NGetDecorLength();
                NGetDecorHeight();
                NGetDecorWidth();
            }
        }

        private void FrontsOrdersDataGrid_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (e.Control is ComboBox)
            {
                ((ComboBox)e.Control).DropDownStyle = ComboBoxStyle.DropDown;
                ((ComboBox)e.Control).AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            }
        }

        private void ExportContextMenuButton_DropDown(object sender, ComponentFactory.Krypton.Toolkit.ContextPositionMenuArgs e)
        {
            ExportContextMenu.Show(ExportContextMenuButton);
        }

        private void CupboardsFormShowButton_Click(object sender, EventArgs e)
        {
            PhantomForm PhantomForm = new Infinium.PhantomForm();
            PhantomForm.Show();

            MCupboardsForm CupboardsForm = new MCupboardsForm(this, ref FrontsOrders);

            TopForm = CupboardsForm;
            CupboardsForm.ShowDialog();

            PhantomForm.Close();

            PhantomForm.Dispose();
            CupboardsForm.Dispose();
            TopForm = null;
        }

        private void FrontsAddSizeButton_Click(object sender, EventArgs e)
        {
            FrontsOrders.AddSize();
        }

        private void DecorInsetTypesComboBox_Leave(object sender, EventArgs e)
        {
            SetAutoComplete(DecorInsetTypesComboBox, false);
        }

        private void DecorInsetColorsComboBox_Leave(object sender, EventArgs e)
        {
            SetAutoComplete(DecorInsetColorsComboBox, false);
        }

        private void DecorInsetTypeDropButton_DropDown(object sender, ComponentFactory.Krypton.Toolkit.ContextPositionMenuArgs e)
        {
            DecorInsetTypesComboBox.Focus();
            DecorInsetTypesComboBox.DroppedDown = true;
        }

        private void DecorInsetColorDropButton_DropDown(object sender, ComponentFactory.Krypton.Toolkit.ContextPositionMenuArgs e)
        {
            DecorInsetColorsComboBox.Focus();
            DecorInsetColorsComboBox.DroppedDown = true;
        }

        private void DecorInsetTypesComboBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                DecorInsetTypesComboBox_SelectionChangeCommitted(sender, e);
                SetAutoComplete(DecorInsetTypesComboBox, true);
            }
        }

        private void DecorInsetColorsComboBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                DecorInsetColorsComboBox_SelectionChangeCommitted(sender, e);
                SetAutoComplete(DecorInsetColorsComboBox, true);
            }
        }

        private void DecorInsetTypesComboBox_Enter(object sender, EventArgs e)
        {
            DecorInsetTypesComboBox.Focus();
            SetAutoComplete(DecorInsetTypesComboBox, true);
        }

        private void DecorInsetColorsComboBox_Enter(object sender, EventArgs e)
        {
            DecorInsetColorsComboBox.Focus();
            SetAutoComplete(DecorInsetColorsComboBox, true);
        }
    }
}
