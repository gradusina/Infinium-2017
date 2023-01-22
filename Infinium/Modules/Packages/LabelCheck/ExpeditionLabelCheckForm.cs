using System;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Infinium
{
    public partial class ExpeditionLabelCheckForm : Form
    {
        private const int eHide = 2;
        private const int eShow = 1;
        private const int eClose = 3;
        private const int eMainMenu = 4;

        private int FormEvent = 0;

        private DataTable EventsDataTable;

        private LightStartForm LightStartForm;
        private bool CanAction = false;
        private int UserID = 0;
        private Form TopForm = null;
        private bool bClearNextPackage = false;
        public ExpeditionCheckLabel CheckLabel;

        [DllImport("user32.dll")]
        private static extern IntPtr GetActiveWindow();


        public ExpeditionLabelCheckForm(LightStartForm tLightStartForm)
        {
            InitializeComponent();

            LightStartForm = tLightStartForm;

            this.MaximumSize = Screen.PrimaryScreen.WorkingArea.Size;

            EventsDataTable = new DataTable();

            //ScanEvents.SetEventsDataTable(EventsDataTable);

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

                    //ScanEvents.AddEvent(EventsDataTable, "Закрыл форму сканирования на экспедицию", 0);

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
                    //ScanEvents.AddEvent(EventsDataTable, "Закрыл форму сканирования на экспедицию", 0);


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
            //ScanEvents.SaveEvents(EventsDataTable);

            FormEvent = eClose;
            AnimateTimer.Enabled = true;
        }

        private void Initialize()
        {
            CheckLabel = new ExpeditionCheckLabel(ref FrontsPackContentDataGrid, ref DecorPackContentDataGrid, ref PackagesDataGrid);

            OrderDateLabel.Text = "";
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
            if (e.KeyCode == Keys.Enter)
            {
                ErrorPackLabel.Visible = false;
                BarcodeLabel.Text = "";
                CheckPicture.Visible = false;

                CheckLabel.Clear();

                if (BarcodeTextBox.Text.Length < 12)
                {
                    //ScanEvents.AddEvent(EventsDataTable, "Ошибка: неверный штрихкод " + BarcodeLabel.Text, 0);
                    BarcodeTextBox.Clear();
                    OrderDateLabel.Text = "";
                    GroupLabel.Text = "";
                    bClearNextPackage = false;
                    return;
                }

                if (bClearNextPackage)
                {
                    int TrayID = 0;
                    CheckLabel.RemovePackageFromTray(Convert.ToInt32(BarcodeTextBox.Text.Substring(3, 9)), ref TrayID);
                    ErrorPackLabel.Visible = true;
                    ErrorPackLabel.Text = "Упаковка №" + Convert.ToInt32(BarcodeTextBox.Text.Substring(3, 9)) + " убрана с поддона №" + TrayID;
                    bClearNextPackage = false;
                    BarcodeTextBox.Clear();

                    return;
                }

                if (string.Equals(BarcodeTextBox.Text, "000000000000"))
                {
                    ErrorPackLabel.Visible = true;
                    ErrorPackLabel.Text = "Следующая отсканированная упаковка будет убрана с поддона";
                    bClearNextPackage = true;
                    BarcodeTextBox.Clear();

                    return;
                }
                else
                    bClearNextPackage = false;

                BarcodeLabel.Text = BarcodeTextBox.Text;

                BarcodeTextBox.Clear();

                string Prefix = BarcodeLabel.Text.Substring(0, 3);

                //ScanEvents.AddEvent(EventsDataTable, "Сканируется этикетка " + BarcodeLabel.Text, 0);
                

                if (CheckLabel.IsPackageLabel(BarcodeLabel.Text))
                {
                    int packageId = Convert.ToInt32(BarcodeLabel.Text.Substring(3, 9));
                    if (CheckLabel.HasPackages(BarcodeLabel.Text))
                    {
                        string message = string.Empty;
                        CheckLabel.GetPackageInfo(packageId);
                        switch (Prefix)
                        {
                            case "001":
                            case "002":
                            {
                                if (CheckLabel.CanPackageExp(packageId, ref message))
                                {
                                    if (bManualInput)
                                    {
                                        CheckLabel.RegisterManualInputOnExp(packageId, CheckLabel.CurrentMainOrderID, CheckLabel.CurrentProductType, Prefix);
                                        message = $"Этикетка отсканирована вручную. Требуется подтверждение в Регистраторе";
                                        ErrorPackLabel.Visible = true;
                                        ErrorPackLabel.Text = message;
                                    }
                                    else
                                    {
                                        if (!CheckLabel.SetExp(ref EventsDataTable, false, packageId, ref message))
                                        {
                                            ErrorPackLabel.Visible = true;
                                            ErrorPackLabel.Text = message;
                                            CheckPicture.Visible = true;
                                            CheckPicture.Image = Properties.Resources.cancel;
                                            BarcodeLabel.ForeColor = Color.FromArgb(240, 0, 0);
                                            GroupLabel.Text = "";
                                            OrderDateLabel.Text = "";
                                            return;
                                        }

                                        if (message.Length > 0)
                                        {
                                            ErrorPackLabel.Visible = true;
                                            ErrorPackLabel.Text = message;
                                        }

                                        CheckOrdersStatus.SetStatusZOV(CheckLabel.CurrentMainOrderID, false);
                                    }
                                }
                                else
                                {
                                    ErrorPackLabel.Visible = true;
                                    ErrorPackLabel.Text = message;
                                    CheckPicture.Visible = true;
                                    CheckPicture.Image = Properties.Resources.cancel;
                                    BarcodeLabel.ForeColor = Color.FromArgb(240, 0, 0);
                                    GroupLabel.Text = "";
                                    OrderDateLabel.Text = "";
                                    return;
                                }
                                break;
                            }
                            case "003":
                            case "004":
                            {
                                if (CheckLabel.CanPackageExp(packageId, ref message))
                                {
                                    if (bManualInput)
                                    {
                                        if (CheckLabel.ExistPackageInRegisterManualInput(packageId))
                                        {
                                            message = "Упаковка уже в Регистраторе";
                                            ErrorPackLabel.Visible = true;
                                            ErrorPackLabel.Text = message;
                                        }
                                        else
                                        {
                                            CheckLabel.RegisterManualInputOnExp(packageId,
                                                CheckLabel.CurrentMainOrderID, CheckLabel.CurrentProductType, Prefix);
                                            message =
                                                $"Этикетка отсканирована вручную. Требуется подтверждение в Регистраторе";
                                            ErrorPackLabel.Visible = true;
                                            ErrorPackLabel.Text = message;
                                        }
                                    }
                                    else
                                    {
                                        if (!CheckLabel.SetExp(ref EventsDataTable, true, packageId, ref message))
                                        {
                                            ErrorPackLabel.Visible = true;
                                            ErrorPackLabel.Text = message;
                                            CheckPicture.Visible = true;
                                            CheckPicture.Image = Properties.Resources.cancel;
                                            BarcodeLabel.ForeColor = Color.FromArgb(240, 0, 0);
                                            GroupLabel.Text = "";
                                            OrderDateLabel.Text = "";
                                            return;
                                        }

                                        if (message.Length > 0)
                                        {
                                            ErrorPackLabel.Visible = true;
                                            ErrorPackLabel.Text = message;
                                        }

                                        CheckLabel.WriteOffFromStore(packageId);
                                        CheckOrdersStatus.SetStatusMarketingForMainOrder(CheckLabel.CurrentMegaOrderID,
                                            CheckLabel.CurrentMainOrderID);
                                    }
                                }
                                else
                                {
                                    ErrorPackLabel.Visible = true;
                                    ErrorPackLabel.Text = message;
                                    CheckPicture.Visible = true;
                                    CheckPicture.Image = Properties.Resources.cancel;
                                    BarcodeLabel.ForeColor = Color.FromArgb(240, 0, 0);
                                    GroupLabel.Text = "";
                                    OrderDateLabel.Text = "";
                                    //ScanEvents.AddEvent(EventsDataTable,
                                    //    "Упаковка №" + Convert.ToInt32(BarcodeLabel.Text.Substring(3, 9)) + " не принята на эксп-цию: запрет на эксп-цию, экспедиция Маркетинг", 2);
                                    return;
                                }

                                break;
                            }
                        }

                        CheckLabel.FilterByPackageID(BarcodeLabel.Text);

                        CheckPicture.Visible = true;
                        CheckPicture.Image = Properties.Resources.OK;
                        BarcodeLabel.ForeColor = Color.FromArgb(82, 169, 24);
                        GroupLabel.Text = CheckLabel.LabelInfo.Group;
                        OrderDateLabel.Text = CheckLabel.LabelInfo.OrderDate;
                    }
                    else
                    {
                        CheckPicture.Visible = true;
                        CheckPicture.Image = Properties.Resources.cancel;
                        BarcodeLabel.ForeColor = Color.FromArgb(240, 0, 0);
                        GroupLabel.Text = "";
                        OrderDateLabel.Text = "";
                    }
                }
                else
                {
                    if (CheckLabel.IsTrayLabel(BarcodeLabel.Text))
                    {
                        if (CheckLabel.HasTray(BarcodeLabel.Text))
                        {
                            string message = string.Empty;
                            int trayId = Convert.ToInt32(BarcodeLabel.Text.Substring(3, 9));
                            switch (Prefix)
                            {
                                //ZOV
                                case "005":

                                    if (bManualInput)
                                    {
                                        CheckLabel.AddTrayToRegisterManualInput(trayId, Prefix);
                                        message = $"Этикетка отсканирована вручную. Требуется подтверждение в Регистраторе";
                                    }
                                    else
                                    {
                                        CheckLabel.SetExp(BarcodeLabel.Text);
                                        CheckOrdersStatus.SetStatusZOV(Convert.ToInt32(BarcodeLabel.Text.Substring(3, 9)), true);
                                    }

                                    break;
                                //Marketing
                                case "006":
                                {
                                    if (CheckLabel.CanTrayExp(Convert.ToInt32(BarcodeLabel.Text.Substring(3, 9)), ref message))
                                    {
                                        if (bManualInput)
                                        {
                                            CheckLabel.AddTrayToRegisterManualInput(trayId, Prefix);
                                            message = "Этикетка отсканирована вручную. Требуется подтверждение в Регистраторе";
                                        }
                                        else
                                        {
                                            CheckLabel.SetExp(BarcodeLabel.Text);
                                            if (message.Length > 0)
                                            {
                                                ErrorPackLabel.Visible = true;
                                                ErrorPackLabel.Text = message;
                                            }

                                            CheckOrdersStatus.SetStatusMarketingForTray(
                                                Convert.ToInt32(BarcodeLabel.Text.Substring(3, 9)));
                                        }
                                    }
                                    else
                                    {
                                        ErrorPackLabel.Visible = true;
                                        ErrorPackLabel.Text = message;
                                        CheckPicture.Visible = true;
                                        CheckPicture.Image = Properties.Resources.cancel;
                                        BarcodeLabel.ForeColor = Color.FromArgb(240, 0, 0);
                                        OrderDateLabel.Text = "";
                                        GroupLabel.Text = "";
                                        CheckLabel.Clear();
                                        return;
                                    }

                                    break;
                                }
                            }

                            CheckLabel.FillPackages(BarcodeLabel.Text);

                            CheckPicture.Visible = true;
                            CheckPicture.Image = Properties.Resources.OK;
                            BarcodeLabel.ForeColor = Color.FromArgb(82, 169, 24);
                            GroupLabel.Text = CheckLabel.LabelInfo.Group;
                            OrderDateLabel.Text = CheckLabel.LabelInfo.OrderDate;

                            if (bManualInput)
                            {
                                ErrorPackLabel.Visible = true;
                                ErrorPackLabel.Text = message;
                            }
                        }
                        else
                        {
                            //ScanEvents.AddEvent(EventsDataTable, "Ошибка: в таблице Packages нет упаковок с номером поддона №" + BarcodeLabel.Text, 0);
                            CheckPicture.Visible = true;
                            CheckPicture.Image = Properties.Resources.cancel;
                            BarcodeLabel.ForeColor = Color.FromArgb(240, 0, 0);
                            GroupLabel.Text = "";
                            OrderDateLabel.Text = "";
                        }
                    }
                    else
                    {
                        //ScanEvents.AddEvent(EventsDataTable, "Ошибка: в таблице Trays нет записей с номером поддона №" + BarcodeLabel.Text, 0);
                        CheckPicture.Visible = true;
                        CheckPicture.Image = Properties.Resources.cancel;
                        BarcodeLabel.ForeColor = Color.FromArgb(240, 0, 0);
                        OrderDateLabel.Text = "";
                        GroupLabel.Text = "";
                        CheckLabel.Clear();
                        return;
                    }
                }
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

        private void PackagesDataGrid_SelectionChanged(object sender, EventArgs e)
        {
            if (CheckLabel == null || CheckLabel.IsPackagesTableEmpty)
                return;

            if (PackagesDataGrid.SelectedRows.Count == 0)
                return;

            int PackageID = Convert.ToInt32(PackagesDataGrid.SelectedRows[0].Cells["PackageID"].Value);
            MainOrdersTabControl.TabPages[0].PageVisible = CheckLabel.FillFrontsPackContent(PackageID);
            MainOrdersTabControl.TabPages[1].PageVisible = CheckLabel.FillDecorPackContent(PackageID);
        }

        private bool bManualInput;
        private Stopwatch sw;
        private int BarCodeSerialLength = 12;
        private int TimeAHumanWouldTakeToType = 1500;
        private void FirstCharacterEntered()
        {
            if (sw == null)
                sw = new Stopwatch();
            sw.Start();
        }

        private void BarcodeTextBox_TextChanged(object sender, EventArgs e)
        {
            if (BarcodeTextBox.Text.Length == 1)
                FirstCharacterEntered();

            if (BarcodeTextBox.Text.Length != BarCodeSerialLength) return;

            sw.Stop();
            if (sw.ElapsedMilliseconds < TimeAHumanWouldTakeToType)
            {
                //Input is from the BarCode Scanner
                bManualInput = false;
            }
            else
            {
                bManualInput = true;
            }
        }
    }
}
