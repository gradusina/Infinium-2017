﻿using System;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Windows.Forms;

namespace Infinium
{
    public partial class MarketingPackagesAllocProfilForm : Form
    {
        private const int eHide = 2;
        private const int eShow = 1;
        private const int eClose = 3;
        private const int eMainMenu = 4;
        private const int FactoryID = 1;

        private const string PackCount = "ProfilPackCount";

        private int FormEvent = 0;
        private int PackType = 0;

        private bool NeedRefresh = false;
        private bool NeedSplash = false;

        private MarketingSplitOrdersForm MarketingSplitOrdersForm;
        private MarketingSplitPackagesForm SplitPackagesForm;
        private Form TopForm = null;
        private LightStartForm LightStartForm;


        private Modules.Packages.Marketing.MarketingPackagesAllocManager PackagesOrdersManager;

        public MarketingPackagesAllocProfilForm(LightStartForm tLightStartForm)
        {
            InitializeComponent();
            LightStartForm = tLightStartForm;


            this.MaximumSize = Screen.PrimaryScreen.WorkingArea.Size;

            Initialize();
            PackagesOrdersManager.CheckPackages();
            PackedCheckButton.Checked = false;
            Filter();
            while (!SplashForm.bCreated) ;
        }

        private void MarketingPackagesAllocProfilForm_Shown(object sender, EventArgs e)
        {
            while (!SplashForm.bCreated) ;
            NeedSplash = true;
            FormEvent = eShow;
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

        private void Initialize()
        {
            PackagesOrdersManager = new Modules.Packages.Marketing.MarketingPackagesAllocManager(
                ref MainOrdersDataGrid, ref MainOrdersFrontsOrdersDataGrid, ref MegaOrdersDataGrid,
                ref MainOrdersDecorOrdersDataGrid, ref MainOrdersTabControl, FactoryID);
        }

        private void MegaOrdersDataGrid_SelectionChanged(object sender, EventArgs e)
        {
            if (PackagesOrdersManager != null)
                if (PackagesOrdersManager.MegaOrdersBindingSource.Count > 0)
                {
                    if (((DataRowView)PackagesOrdersManager.MegaOrdersBindingSource.Current)["MegaOrderID"] != DBNull.Value)
                    {
                        if (NeedSplash)
                        {
                            NeedSplash = false;
                            Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                            T.Start();

                            while (!SplashWindow.bSmallCreated) ;

                            System.Diagnostics.Stopwatch sw = new System.Diagnostics.Stopwatch();
                            sw.Start();
                            PackagesOrdersManager.FilterMainOrders(Convert.ToInt32(((DataRowView)PackagesOrdersManager.MegaOrdersBindingSource.Current)["MegaOrderID"]));
                            sw.Stop();
                            double G = sw.Elapsed.TotalSeconds;

                            while (SplashWindow.bSmallCreated)
                                SmallWaitForm.CloseS = true;
                            NeedSplash = true;
                        }
                        else
                            PackagesOrdersManager.FilterMainOrders(Convert.ToInt32(((DataRowView)PackagesOrdersManager.MegaOrdersBindingSource.Current)["MegaOrderID"]));
                    }
                }
        }

        private void MainOrdersDataGrid_SelectionChanged(object sender, EventArgs e)
        {
            if (PackagesOrdersManager != null)
                if (PackagesOrdersManager.MainOrdersBindingSource.Count > 0)
                {
                    if (((DataRowView)PackagesOrdersManager.MainOrdersBindingSource.Current)["MainOrderID"] != DBNull.Value)
                    {
                        PackagesOrdersManager.Filter(Convert.ToInt32(((DataRowView)PackagesOrdersManager.MainOrdersBindingSource.Current)["MainOrderID"]));

                        if (MainOrdersTabControl.TabPages[0].Visible && MainOrdersTabControl.TabPages[1].Visible)
                            MainOrdersTabControl.SelectedTabPage = MainOrdersTabControl.TabPages[0];

                        NextPackNumberLabel.Text = (PackagesOrdersManager.MainOrderPackCount + PackagesOrdersManager.GetPackCount + 1).ToString();
                    }
                }
                else
                {
                    PackagesOrdersManager.Filter(-1);
                    NextPackNumberLabel.Text = "";
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
            PackagesOrdersManager.FilterMegaOrders(PackedCheckButton.Checked,
                NonPackedCheckButton.Checked);
            PackagesOrdersManager.FilterMainOrders(PackedCheckButton.Checked,
                NonPackedCheckButton.Checked);
        }

        private void MainOrdersFrontsOrdersDataGrid_KeyDown(object sender, KeyEventArgs e)
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

                string r = MainOrdersFrontsOrdersDataGrid.SelectedRows[0].Cells["PackNumber"].FormattedValue.ToString();

                MainOrdersFrontsOrdersDataGrid.SelectedRows[0].Cells["PackNumber"].Value = r + c.ToString();
            }

            if (e.KeyCode == Keys.Delete)
                MainOrdersFrontsOrdersDataGrid.SelectedRows[0].Cells["PackNumber"].Value = DBNull.Value;

            if (e.KeyCode == Keys.Back)
            {
                string r = MainOrdersFrontsOrdersDataGrid.SelectedRows[0].Cells["PackNumber"].FormattedValue.ToString();
                if (r.Length > 1)
                {
                    r = r.Substring(0, r.Length - 1);
                    MainOrdersFrontsOrdersDataGrid.SelectedRows[0].Cells["PackNumber"].Value = r;
                }
                else
                {
                    if (r.Length == 1)
                    {
                        MainOrdersFrontsOrdersDataGrid.SelectedRows[0].Cells["PackNumber"].Value = DBNull.Value;
                    }
                }
            }
        }

        private void MainOrdersDecorOrdersDataGrid_KeyDown(object sender, KeyEventArgs e)
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

                string r = MainOrdersDecorOrdersDataGrid.SelectedRows[0].Cells["PackNumber"].FormattedValue.ToString();

                MainOrdersDecorOrdersDataGrid.SelectedRows[0].Cells["PackNumber"].Value = r + c.ToString();
            }

            if (e.KeyCode == Keys.Delete)
                MainOrdersDecorOrdersDataGrid.SelectedRows[0].Cells["PackNumber"].Value = DBNull.Value;

            if (e.KeyCode == Keys.Back)
            {
                string r = MainOrdersDecorOrdersDataGrid.SelectedRows[0].Cells["PackNumber"].FormattedValue.ToString();
                if (r.Length > 1)
                {
                    r = r.Substring(0, r.Length - 1);
                    MainOrdersDecorOrdersDataGrid.SelectedRows[0].Cells["PackNumber"].Value = r;
                }
                else
                {
                    if (r.Length == 1)
                    {
                        MainOrdersDecorOrdersDataGrid.SelectedRows[0].Cells["PackNumber"].Value = DBNull.Value;
                    }
                }
            }
        }

        private void SavePackagesButton_Click(object sender, EventArgs e)
        {
            NeedSplash = false;
            Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Сохранение упаковок.\r\nПодождите..."); });
            T.Start();

            while (!SplashWindow.bSmallCreated) ;

            //if (!PackagesOrdersManager.IsSimpleAndCurved)
            //{
            //    while (SplashWindow.bSmallCreated)
            //        SmallWaitForm.CloseS = true;
            //    NeedSplash = true;

            //    Infinium.LightMessageBox.Show(ref TopForm, false,
            //       "В одной упаковке не могут лежать гнутые и негнутые фасады. Перепакуйте",
            //       "Ошибка сохранения упаковок");

            //    return;
            //}

            if (PackagesOrdersManager.IsEmptyPack)
            {
                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;
                NeedSplash = true;

                Infinium.LightMessageBox.Show(ref TopForm, false,
                   "Не все позиции распределены",
                   "Ошибка сохранения упаковок");

                return;
            }

            if (PackagesOrdersManager.IsOverflow == false)
            {
                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;
                NeedSplash = true;

                Infinium.LightMessageBox.Show(ref TopForm, false,
                   "В упаковке больше 6 позиций. Перепакуйте",
                   "Ошибка сохранения упаковок");

                return;
            }

            if (PackagesOrdersManager.IsConsequence == false)
            {
                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;
                NeedSplash = true;

                Infinium.LightMessageBox.Show(ref TopForm, false,
                    "Упаковки в заказе непоследовательны!", "Ошибка");
                return;

            }

            if (PackagesOrdersManager.IsDifferentPackNumbers == false)
            {
                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;
                NeedSplash = true;

                Infinium.LightMessageBox.Show(ref TopForm, false,
                    "Фасады и декор находятся в одной упаковке",
                    "Ошибка сохранения упаковок");
                return;
            }

            if (MainOrdersDataGrid.SelectedRows.Count > 0 &&
                Convert.ToInt32(MainOrdersDataGrid.SelectedRows[0].Cells["FactoryID"].Value) == 0)
            {
                if (PackagesOrdersManager.PackagesMainOrdersFrontsOrders.FrontsOrdersBindingSource.Count > 0)
                {
                    if (!PackagesOrdersManager.PackagesMainOrdersFrontsOrders.IsPacked)
                    {
                        while (SplashWindow.bSmallCreated)
                            SmallWaitForm.CloseS = true;
                        NeedSplash = true;

                        Infinium.LightMessageBox.Show(ref TopForm, false,
                            "Заказ на обеих фирмах. Не все фасады запакованы",
                            "Ошибка сохранения упаковок");
                        return;
                    }
                }

                if (PackagesOrdersManager.PackagesMainOrdersDecorOrders.DecorOrdersBindingSource.Count > 0)
                {
                    if (!PackagesOrdersManager.PackagesMainOrdersDecorOrders.IsPacked)
                    {
                        while (SplashWindow.bSmallCreated)
                            SmallWaitForm.CloseS = true;
                        NeedSplash = true;

                        Infinium.LightMessageBox.Show(ref TopForm, false,
                            "Заказ на обеих фирмах. Не весь декор запакован",
                            "Ошибка сохранения упаковок");
                        return;

                    }
                }
            }

            int ClientID = 0;
            int OrderNumber = 0;
            int MainOrderID = 0;
            int MegaOrderID = 0;
            string Notes = string.Empty;
            if (MegaOrdersDataGrid.SelectedRows.Count > 0 && MegaOrdersDataGrid.SelectedRows[0].Cells["ClientID"].Value != DBNull.Value)
                ClientID = Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[0].Cells["ClientID"].Value);
            if (MegaOrdersDataGrid.SelectedRows.Count > 0 && MegaOrdersDataGrid.SelectedRows[0].Cells["OrderNumber"].Value != DBNull.Value)
                OrderNumber = Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[0].Cells["OrderNumber"].Value);
            if (MegaOrdersDataGrid.SelectedRows.Count > 0 && MegaOrdersDataGrid.SelectedRows[0].Cells["MegaOrderID"].Value != DBNull.Value)
                MegaOrderID = Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[0].Cells["MegaOrderID"].Value);
            if (MainOrdersDataGrid.SelectedRows.Count > 0 && MainOrdersDataGrid.SelectedRows[0].Cells["MainOrderID"].Value != DBNull.Value)
                MainOrderID = Convert.ToInt32(MainOrdersDataGrid.SelectedRows[0].Cells["MainOrderID"].Value);
            if (MainOrdersDataGrid.SelectedRows.Count > 0 && MainOrdersDataGrid.SelectedRows[0].Cells["Notes"].Value != DBNull.Value)
                Notes = MainOrdersDataGrid.SelectedRows[0].Cells["Notes"].Value.ToString();
            PackagesOrdersManager.SetCabFurParameters(ClientID, OrderNumber, MegaOrderID, MainOrderID, Notes);

            PackagesOrdersManager.SavePackageDetails();
            PackagesOrdersManager.SetPackStatus();

            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
            NeedSplash = true;
        }

        private void PackedCheckButton_CheckedChanged(object sender, EventArgs e)
        {
            NeedSplash = false;
            Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
            T.Start();

            while (!SplashWindow.bSmallCreated) ;
            Filter();

            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;

            NeedSplash = true;
        }

        private void NonPackedCheckButton_CheckedChanged(object sender, EventArgs e)
        {
            NeedSplash = false;
            Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
            T.Start();

            while (!SplashWindow.bSmallCreated) ;
            Filter();

            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
            NeedSplash = true;
        }

        private void SplitPackagesButton_Click(object sender, EventArgs e)
        {
            if (PackType == 0)//fronts
            {
                if (MainOrdersFrontsOrdersDataGrid.Rows.Count < 1 || Convert.ToInt32(MainOrdersFrontsOrdersDataGrid.SelectedRows[0].Cells["Count"].Value) < 2)
                    return;

                Infinium.Modules.Packages.Marketing.SplitStruct SS = new Modules.Packages.Marketing.SplitStruct();

                SplitPackagesForm = new MarketingSplitPackagesForm();
                SplitPackagesForm.SetSplitPositions(Convert.ToInt32(MainOrdersFrontsOrdersDataGrid.SelectedRows[0].Cells["Count"].Value),
                                                    ref SS);
                TopForm = SplitPackagesForm;
                SplitPackagesForm.ShowDialog();
                TopForm = null;

                if (SplitPackagesForm.AccessibleName == "true")
                {
                    SplitPackagesForm.Dispose();
                    SplitPackagesForm = null;
                    GC.Collect();
                }

                if (SS.IsSplit)
                    PackagesOrdersManager.PackagesMainOrdersFrontsOrders.SplitCurrentItem(SS);
            }

            if (PackType == 1)//decor
            {
                if (MainOrdersDecorOrdersDataGrid.Rows.Count < 1 || Convert.ToInt32(MainOrdersDecorOrdersDataGrid.SelectedRows[0].Cells["Count"].Value) < 2)
                    return;

                Infinium.Modules.Packages.Marketing.SplitStruct SS = new Modules.Packages.Marketing.SplitStruct();

                SplitPackagesForm = new MarketingSplitPackagesForm();
                SplitPackagesForm.SetSplitPositions(
                            Convert.ToInt32(MainOrdersDecorOrdersDataGrid.SelectedRows[0].Cells["Count"].Value),
                            ref SS);

                TopForm = SplitPackagesForm;
                SplitPackagesForm.ShowDialog();
                TopForm = null;

                if (SplitPackagesForm.AccessibleName == "true")
                {
                    SplitPackagesForm.Dispose();
                    SplitPackagesForm = null;
                    GC.Collect();
                }

                if (SS.IsSplit)
                    PackagesOrdersManager.PackagesMainOrdersDecorOrders.SplitCurrentItem(SS);
            }
        }

        private void MainOrdersFrontsOrdersDataGrid_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            PackType = 0;
        }

        private void MainOrdersDecorOrdersDataGrid_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            PackType = 1;
        }

        private void ClearPackageButton_Click(object sender, EventArgs e)
        {
            if (PackagesOrdersManager != null)
                if (PackagesOrdersManager.MainOrdersBindingSource.Count > 0)
                {
                    if (((DataRowView)PackagesOrdersManager.MainOrdersBindingSource.Current)["MainOrderID"] != DBNull.Value)
                    {
                        //if (PackagesOrdersManager.CheckPackStatus())
                        //{
                        //    Infinium.LightMessageBox.Show(ref TopForm, false,
                        //        "Заказ уже запакован. Очистить упаковки нельзя",
                        //        "Ошибка");

                        //    return;
                        //}

                        bool OKCancel = Infinium.LightMessageBox.Show(ref TopForm, true,
                            "В данном подзаказе очистятся ВСЕ упаковки. Распечатанные этикетки будут недействительны. Продолжить?",
                            "Внимание");
                        if (!OKCancel)
                            return;

                        NeedSplash = false;
                        Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Очистка упаковок.\r\nПодождите..."); });
                        T.Start();

                        while (!SplashWindow.bSmallCreated) ;

                        //PackagesOrdersManager.FixOrderEvent(Convert.ToInt32(((DataRowView)PackagesOrdersManager.MainOrdersBindingSource.Current)["MainOrderID"]),
                        //    "Упаковки очищены, UserID=" + Security.CurrentUserID);
                        PackagesOrdersManager.ClearPackage();
                        PackagesOrdersManager.SetPackStatus();

                        while (SplashWindow.bSmallCreated)
                            SmallWaitForm.CloseS = true;
                        NeedSplash = true;
                    }
                }
        }

        private void PackOrdersSetButton_Click(object sender, EventArgs e)
        {
            int[] SelectedMainOrders = PackagesOrdersManager.GetSelectedMainOrders();

            if (SelectedMainOrders.Count() < 2)
                return;

            NeedSplash = false;
            Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Сохранение упаковок.\r\nПодождите..."); });
            T.Start();

            while (!SplashWindow.bSmallCreated) ;

            if (PackagesOrdersManager.AreMainOrdersInSet(SelectedMainOrders))
            {
                int StartingMainOrderID = 0;
                if (PackagesOrdersManager.AreSelectedMainOrdersPacked(ref SelectedMainOrders, ref StartingMainOrderID))
                {
                    if (SelectedMainOrders.Count() > 0)
                        PackagesOrdersManager.SaveOrdersSet(SelectedMainOrders, StartingMainOrderID,
                            Convert.ToInt32(MegaOrdersDataGrid.SelectedRows[0].Cells["MegaOrderID"].Value));
                    Filter();
                    while (SplashWindow.bSmallCreated)
                        SmallWaitForm.CloseS = true;
                    NeedSplash = true;
                }
                else
                {
                    while (SplashWindow.bSmallCreated)
                        SmallWaitForm.CloseS = true;
                    NeedSplash = true;
                    Infinium.LightMessageBox.Show(ref TopForm, false,
                        "Среди выбранных подзаказов ни один не упакован",
                        "Ошибка");

                    return;
                }
            }
            else
            {
                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;
                NeedSplash = true;
                Infinium.LightMessageBox.Show(ref TopForm, false,
                    "Выбранные подзаказы разные", "Ошибка");
                return;
            }
            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
            NeedSplash = true;
        }

        private void MainOrdersFrontsOrdersDataGrid_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            if (e.ColumnIndex == 24 && !string.IsNullOrEmpty(e.FormattedValue.ToString()))
            {

                if (int.TryParse(e.FormattedValue.ToString(), out int NewPackNumber))
                {
                    if (NewPackNumber <= PackagesOrdersManager.MainOrderPackCount)
                        MainOrdersFrontsOrdersDataGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = DBNull.Value;
                }
            }
            PackagesOrdersManager.PackagesMainOrdersFrontsOrders.FrontsOrdersBindingSource.EndEdit();
            NextPackNumberLabel.Text = (PackagesOrdersManager.MainOrderPackCount + PackagesOrdersManager.GetPackCount + 1).ToString();
        }

        private void MainOrdersDecorOrdersDataGrid_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            if (e.ColumnIndex == 19 && !string.IsNullOrEmpty(e.FormattedValue.ToString()))
            {

                if (int.TryParse(e.FormattedValue.ToString(), out int NewPackNumber))
                {
                    if (NewPackNumber <= PackagesOrdersManager.MainOrderPackCount)
                        MainOrdersDecorOrdersDataGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = DBNull.Value;
                }
            }
            PackagesOrdersManager.PackagesMainOrdersDecorOrders.DecorOrdersBindingSource.EndEdit();
            NextPackNumberLabel.Text = (PackagesOrdersManager.MainOrderPackCount + PackagesOrdersManager.GetPackCount + 1).ToString();
        }

        private void SplitMainOrderMenuItem_Click(object sender, EventArgs e)
        {
            int MegaOrderID = Convert.ToInt32(((DataRowView)PackagesOrdersManager.MegaOrdersBindingSource.Current).Row["MegaOrderID"]);
            int MainOrderID = Convert.ToInt32(((DataRowView)PackagesOrdersManager.MainOrdersBindingSource.Current).Row["MainOrderID"]);

            Thread T = new Thread(delegate () { SplashWindow.CreateSplash(); });
            T.Start();

            while (!SplashForm.bCreated) ;
            MarketingSplitOrdersForm = new MarketingSplitOrdersForm(this, MegaOrderID, MainOrderID, FactoryID);

            TopForm = MarketingSplitOrdersForm;
            MarketingSplitOrdersForm.ShowDialog();
            TopForm = null;

            MarketingSplitOrdersForm.Dispose();
            MarketingSplitOrdersForm = null;
            GC.Collect();
        }

        private void MainOrdersDataGrid_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Right && PackagesOrdersManager.IsMainOrderNotPacked)
            {
                MainOrdersDataGrid.CurrentCell = MainOrdersDataGrid[e.ColumnIndex, e.RowIndex];
                kryptonContextMenu1.Show(new Point(Cursor.Position.X - 212, Cursor.Position.Y - 10));
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

        private void MinimizeButton_Click(object sender, EventArgs e)
        {
            NeedSplash = false;
            FormEvent = eHide;
            AnimateTimer.Enabled = true;
        }

        private void btnDeletePackage_Click(object sender, EventArgs e)
        {
            if (PackagesOrdersManager != null)
                if (PackagesOrdersManager.MainOrdersBindingSource.Count > 0)
                {
                    if (((DataRowView)PackagesOrdersManager.MainOrdersBindingSource.Current)["MainOrderID"] != DBNull.Value)
                    {

                        bool OKCancel = Infinium.LightMessageBox.Show(ref TopForm, true,
                            "Выбранная позиция будет очищена. Продолжить?",
                            "Внимание");
                        if (!OKCancel)
                            return;

                        NeedSplash = false;
                        Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref TopForm, "Очистка упаковок.\r\nПодождите..."); });
                        T.Start();

                        while (!SplashWindow.bSmallCreated) ;

                        //PackagesOrdersManager.FixOrderEvent(Convert.ToInt32(((DataRowView)PackagesOrdersManager.MainOrdersBindingSource.Current)["MainOrderID"]),
                        //    "Упаковки очищены, UserID=" + Security.CurrentUserID);
                        if (PackType == 0)
                        {
                            int MainOrderID = 0;
                            int PackageID = 0;
                            if (MainOrdersDataGrid.SelectedRows.Count > 0 && MainOrdersDataGrid.SelectedRows[0].Cells["MainOrderID"].Value != DBNull.Value)
                                MainOrderID = Convert.ToInt32(MainOrdersDataGrid.SelectedRows[0].Cells["MainOrderID"].Value);
                            if (MainOrdersFrontsOrdersDataGrid.SelectedRows.Count > 0 && MainOrdersFrontsOrdersDataGrid.SelectedRows[0].Cells["PackNumber"].Value != DBNull.Value)
                                PackageID = Convert.ToInt32(MainOrdersFrontsOrdersDataGrid.SelectedRows[0].Cells["PackNumber"].Value);
                            PackagesOrdersManager.DeletePackage(MainOrderID, PackageID);
                        }
                        if (PackType == 1)
                        {
                            int MainOrderID = 0;
                            int PackageID = 0;
                            if (MainOrdersDataGrid.SelectedRows.Count > 0 && MainOrdersDataGrid.SelectedRows[0].Cells["MainOrderID"].Value != DBNull.Value)
                                MainOrderID = Convert.ToInt32(MainOrdersDataGrid.SelectedRows[0].Cells["MainOrderID"].Value);
                            if (MainOrdersDecorOrdersDataGrid.SelectedRows.Count > 0 && MainOrdersDecorOrdersDataGrid.SelectedRows[0].Cells["PackNumber"].Value != DBNull.Value)
                                PackageID = Convert.ToInt32(MainOrdersDecorOrdersDataGrid.SelectedRows[0].Cells["PackNumber"].Value);
                            PackagesOrdersManager.DeletePackage(MainOrderID, PackageID);
                        }
                        PackagesOrdersManager.SetPackStatus();

                        while (SplashWindow.bSmallCreated)
                            SmallWaitForm.CloseS = true;
                        NeedSplash = true;
                    }
                }
        }
    }
}