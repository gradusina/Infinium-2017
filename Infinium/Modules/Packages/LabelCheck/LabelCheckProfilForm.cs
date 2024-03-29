﻿using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Infinium
{
    public partial class LabelCheckProfilForm : Form
    {
        private const int eHide = 2;
        private const int eShow = 1;
        private const int eClose = 3;
        private const int eMainMenu = 4;

        private int FormEvent = 0;

        private LightStartForm LightStartForm;

        private bool CanAction = false;
        private int UserID = 0;
        public CheckLabel CheckLabel;
        private Form TopForm = null;
        [DllImport("user32.dll")]
        private static extern IntPtr GetActiveWindow();

        public LabelCheckProfilForm(LightStartForm tLightStartForm)
        {
            InitializeComponent();

            LightStartForm = tLightStartForm;


            this.MaximumSize = Screen.PrimaryScreen.WorkingArea.Size;

            Initialize();

            while (!SplashForm.bCreated) ;
        }

        private void LabelCheckProfilForm_Shown(object sender, EventArgs e)
        {
            while (!SplashForm.bCreated) ;
            FormEvent = eShow;
            AnimateTimer.Enabled = true;

            PhantomForm PhantomForm = new Infinium.PhantomForm();
            PhantomForm.Show();

            ReAutorizationForm ReAutorizationForm = new ReAutorizationForm(this);

            TopForm = ReAutorizationForm;
            ReAutorizationForm.ShowDialog();

            PhantomForm.Close();

            PhantomForm.Dispose();
            bool PressOK = ReAutorizationForm.PressOK;
            UserID = ReAutorizationForm.UserID;
            ReAutorizationForm.Dispose();
            TopForm = null;

            if (PressOK)
            {
                CheckLabel.UserID = UserID;
                CanAction = true;
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

        private void NavigateMenuCloseButton_Click(object sender, EventArgs e)
        {
            FormEvent = eClose;
            AnimateTimer.Enabled = true;
        }

        private void NavigateMenuBackButton_Click(object sender, EventArgs e)
        {

        }

        private void NavigateMenuHerculesButton_Click(object sender, EventArgs e)
        {

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


        private void Initialize()
        {
            CheckLabel = new CheckLabel(ref FrontsPackContentDataGrid, ref DecorPackContentDataGrid);

            ClientLabel.Text = "";
            MegaOrderNumberLabel.Text = "";
            MainOrderNumberLabel.Text = "";
            DispatchDateLabel.Text = "";
            OrderDateLabel.Text = "";
            ProductTypeLabel.Text = "";
            PackNumberLabel.Text = "";
            TotalLabel.Text = "";
            GroupLabel.Text = "";
        }

        private void BarcodeTextBox_Leave(object sender, EventArgs e)
        {
            if (!CanAction)
                return;
            CheckTimer.Enabled = true;
        }

        private void NavigateMenuCloseButton_MouseUp(object sender, MouseEventArgs e)
        {
            if (!CanAction)
                return;
            BarcodeTextBox.Focus();
        }

        private void CheckTimer_Tick(object sender, EventArgs e)
        {
            if (!CanAction)
                return;
            if (!BarcodeTextBox.Focused)
            {
                BarcodeTextBox.Focus();
            }
            //CheckTimer.Enabled = false;
        }

        private void BarcodeTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (!CanAction)
                return;
            double G = 0;

            if (e.KeyCode == Keys.Enter)
            {
                System.Diagnostics.Stopwatch sw = new System.Diagnostics.Stopwatch();
                sw.Start();

                BarcodeLabel.Text = "";
                CheckPicture.Visible = false;

                CheckLabel.Clear();

                if (BarcodeTextBox.Text.Length < 12)
                {
                    BarcodeTextBox.Clear();

                    ClientLabel.Text = "";
                    MegaOrderNumberLabel.Text = "";
                    MainOrderNumberLabel.Text = "";
                    DispatchDateLabel.Text = "";
                    OrderDateLabel.Text = "";
                    ProductTypeLabel.Text = "";
                    PackNumberLabel.Text = "";
                    TotalLabel.Text = "";
                    GroupLabel.Text = "";

                    return;
                }

                BarcodeLabel.Text = BarcodeTextBox.Text;

                BarcodeTextBox.Clear();


                if (CheckLabel.CheckBarcode(BarcodeLabel.Text))
                {
                    CheckPicture.Visible = true;
                    CheckPicture.Image = Properties.Resources.OK;
                    BarcodeLabel.ForeColor = Color.FromArgb(82, 169, 24);
                    CheckLabel.GetLabelInfo(BarcodeLabel.Text);

                    ClientLabel.Text = CheckLabel.LabelInfo.ClientName;
                    MegaOrderNumberLabel.Text = CheckLabel.LabelInfo.MegaOrderNumber;
                    MainOrderNumberLabel.Text = CheckLabel.LabelInfo.MainOrderNumber;
                    DispatchDateLabel.Text = CheckLabel.LabelInfo.DispatchDate;
                    OrderDateLabel.Text = CheckLabel.LabelInfo.OrderDate;
                    ProductTypeLabel.Text = CheckLabel.LabelInfo.ProductType;
                    PackNumberLabel.Text = CheckLabel.LabelInfo.CurrentPackNumber;
                    TotalLabel.Text = CheckLabel.LabelInfo.PackedToTotal;
                    DispatchDateLabel.ForeColor = CheckLabel.LabelInfo.DispatchDateColor;
                    TotalLabel.ForeColor = CheckLabel.LabelInfo.TotalLabelColor;
                    GroupLabel.Text = CheckLabel.LabelInfo.Group;

                    CheckLabel.SetGridColor(CheckLabel.LabelInfo.ProductType, true);

                    if (CheckLabel.LabelInfo.Group == "Маркетинг")
                    {
                        //CheckLabel.SetMainOrderStatus(BarcodeLabel.Text, Convert.ToInt32(CheckLabel.LabelInfo.MainOrderNumber));
                        CheckOrdersStatus.SetStatusMarketingForMainOrder(Convert.ToInt32(CheckLabel.LabelInfo.MegaOrderID),
                            Convert.ToInt32(CheckLabel.LabelInfo.MainOrderID));

                    }
                    if (CheckLabel.LabelInfo.Group == "ЗОВ")
                    {
                        //CheckLabel.SetMainOrderStatus(BarcodeLabel.Text, Convert.ToInt32(CheckLabel.LabelInfo.MainOrderNumber));
                        CheckOrdersStatus.SetStatusZOV(CheckLabel.LabelInfo.MainOrderID, false);
                    }
                }
                else
                {
                    CheckPicture.Visible = true;
                    CheckPicture.Image = Properties.Resources.cancel;
                    BarcodeLabel.ForeColor = Color.FromArgb(240, 0, 0);

                    ClientLabel.Text = "";
                    MegaOrderNumberLabel.Text = "";
                    MainOrderNumberLabel.Text = "";
                    DispatchDateLabel.Text = "";
                    OrderDateLabel.Text = "";
                    ProductTypeLabel.Text = "";
                    PackNumberLabel.Text = "";
                    TotalLabel.Text = "";
                    GroupLabel.Text = "";

                    CheckLabel.SetGridColor(CheckLabel.LabelInfo.ProductType, false);
                }

                sw.Stop();
                G = sw.Elapsed.TotalMilliseconds;
                //MessageBox.Show(G.ToString());
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
            if (GetActiveWindow() != this.Handle)
            {
                this.Activate();
            }
        }

        private void LabelCheckProfilForm_Load(object sender, EventArgs e)
        {
        }

    }
}
