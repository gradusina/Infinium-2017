using Infinium.Modules.Entrance;

using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Infinium
{
    public partial class EntranceForm : Form
    {
        private const int eHide = 2;
        private const int eShow = 1;
        private const int eClose = 3;

        public EntranceCheck Entrance;
        public PrintUsersBarcode PrintUsersBarcode;

        public static bool BSmallCreated;

        private int FormEvent = 0;
        private bool NeedSplash = false;

        private Form TopForm = null;
        private LightStartForm LightStartForm;

        [DllImport("user32.dll")]
        private static extern IntPtr GetActiveWindow();

        public EntranceForm(LightStartForm tLightStartForm)
        {
            InitializeComponent();

            LightStartForm = tLightStartForm;

            MaximumSize = Screen.PrimaryScreen.WorkingArea.Size;

            while (!SplashForm.bCreated) ;
        }

        private void MenuCloseButton_Click(object sender, EventArgs e)
        {
            FormEvent = eClose;
            AnimateTimer.Enabled = true;
        }

        private void AnimateTimer_Tick(object sender, EventArgs e)
        {
            if (!DatabaseConfigsManager.Animation)
            {
                Opacity = 1;

                if (FormEvent == eClose || FormEvent == eHide)
                {
                    AnimateTimer.Enabled = false;

                    if (FormEvent == eClose)
                    {

                        LightStartForm.CloseForm(this);
                        Close();
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
                if (Convert.ToDecimal(Opacity) != Convert.ToDecimal(0.00))
                    Opacity = Convert.ToDouble(Convert.ToDecimal(Opacity) - Convert.ToDecimal(0.05));
                else
                {
                    AnimateTimer.Enabled = false;

                    if (FormEvent == eClose)
                    {

                        LightStartForm.CloseForm(this);
                        Close();
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
                if (Opacity != 1)
                    Opacity += 0.05;
                else
                {
                    NeedSplash = true;
                    AnimateTimer.Enabled = false;
                    SplashForm.CloseS = true;
                }
            }
        }

        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);

            if (m.Msg == 6 &&
                m.WParam.ToInt32() == 1)
            {
                if (TopForm != null)
                    TopForm.Activate();
            }
        }

        private void LoadCalculationsForm_Load(object sender, EventArgs e)
        {
            Entrance = new EntranceCheck();
            PrintUsersBarcode = new PrintUsersBarcode();

            dataGridView1.DataSource = Entrance.EntranceBs;
        }

        private void LoadCalculationsForm_FormClosing(object sender, FormClosingEventArgs e)
        {

            //_loginForm.Close();
        }

        private void dataGridSettings(ref DataGridView grid)
        {
            int displayIndex = 0;

            if (grid.Columns.Contains("id"))
            {
                grid.Columns["id"].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
                grid.Columns["id"].HeaderText = "№ п/п";
            }
            if (grid.Columns.Contains("accountName"))
            {
                grid.Columns["accountName"].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
                grid.Columns["accountName"].HeaderText = "ТМЦ";
            }
            if (grid.Columns.Contains("sumtotal"))
            {
                grid.Columns["sumtotal"].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
                grid.Columns["sumtotal"].HeaderText = "ИТОГО";
            }
            if (grid.Columns.Contains("invnumber"))
            {
                grid.Columns["invnumber"].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
                grid.Columns["invnumber"].HeaderText = "Артикул";
            }
            if (grid.Columns.Contains("count"))
            {
                grid.Columns["count"].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
                grid.Columns["count"].HeaderText = "Кол-во";
            }

            if (grid.Columns.Contains("cost"))
                grid.Columns["cost"].Visible = false;
            if (grid.Columns.Contains("configid"))
                grid.Columns["configid"].Visible = false;
            if (grid.Columns.Contains("itemid"))
                grid.Columns["itemid"].Visible = false;
        }

        private void LoadCalculationsForm_Shown(object sender, EventArgs e)
        {
            while (!SplashForm.bCreated) ;
            FormEvent = eShow;
            AnimateTimer.Enabled = true;
            BarcodeTextBox.Focus();
        }

        private void MinimizeButton_Click(object sender, EventArgs e)
        {
            NeedSplash = false;
            FormEvent = eHide;
            AnimateTimer.Enabled = true;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            //if (GetActiveWindow() != this.Handle)
            //{
            //    this.Activate();
            //}
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
            if (e.KeyCode == Keys.Enter)
            {
                ErrorPackLabel.Visible = false;
                BarcodeLabel.Text = "";
                CheckPicture.Visible = false;
                CheckPicture.Image = Properties.Resources.cancel;
                BarcodeLabel.ForeColor = Color.FromArgb(240, 0, 0);

                lbEntry.Text = "";
                lbUser.Text = "";

                Entrance.Clear();

                if (BarcodeTextBox.Text.Length < 12)
                {
                    return;
                }

                if (string.Equals(BarcodeTextBox.Text, "000000000000"))
                {
                    ErrorPackLabel.Visible = true;
                    ErrorPackLabel.Text = "отсканируйте еще раз";
                    BarcodeTextBox.Clear();

                    return;
                }

                BarcodeLabel.Text = BarcodeTextBox.Text;
                BarcodeTextBox.Clear();

                var prefix = BarcodeLabel.Text.Substring(0, 3);
                var userId = Convert.ToInt32(BarcodeLabel.Text.Substring(3, 9));

                //проверка - баркод пользователя?
                if (prefix != "015") return;

                //это вход или выход?
                if (Entrance.IsUserEnter(userId))
                {
                    Entrance.UserQuit(userId);
                }
                else
                {
                    Entrance.UserEnter(userId);
                }

                Entrance.FilterEntrance(DateTime.Now);

                CheckPicture.Visible = true;
                CheckPicture.Image = Properties.Resources.OK;
                BarcodeLabel.ForeColor = Color.FromArgb(82, 169, 24);
                lbUser.Text = Entrance.Entry.UserName;
                lbEntry.Text = Entrance.Entry.EntryDate.ToString();
                lbExit.Text = Entrance.Entry.QuitDate == null ? "-" : Entrance.Entry.QuitDate.ToString();

                return;
            }
        }

        private void BarcodeTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
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

        private void BarcodeTextBox_Leave(object sender, EventArgs e)
        {
            CheckTimer.Enabled = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            PrintUsersBarcode.ClearLabelInfo();
            if (printDialog1.ShowDialog() == DialogResult.OK)
            {
                var labelInfo = new UsersBarcodeInfo()
                {
                    BarcodeNumber = string.Empty
                };
                var userId = 321;

                labelInfo.BarcodeNumber = PrintUsersBarcode.GetBarcodeNumber(15, userId);

                PrintUsersBarcode.AddLabelInfo(ref labelInfo);
                printDialog1.Document = PrintUsersBarcode.PD;

                PrintUsersBarcode.Print();
            }
        }
    }
}
