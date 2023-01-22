using System;
using System.Data;
using System.Windows.Forms;

namespace Infinium
{
    public partial class ReaddressMarketingDispatchMenu : Form
    {
        private const int eHide = 2;
        private const int eShow = 1;
        private const int eClose = 3;
        private const int eMainMenu = 4;

        public bool PressOK = false;
        public int ClientID = 0;

        private int FormEvent = 0;

        private DataTable ClientsDT;
        private Form MainForm = null;
        private Form TopForm = null;

        public ReaddressMarketingDispatchMenu(Form tMainForm, DataTable tClientsDT)
        {
            MainForm = tMainForm;
            ClientsDT = tClientsDT;
            InitializeComponent();
            Initialize();
        }

        private void btnPressOK_Click(object sender, EventArgs e)
        {
            ClientID = Convert.ToInt32(((DataRowView)cbClients.SelectedItem).Row["ClientID"]);
            PressOK = true;
            FormEvent = eClose;
            AnimateTimer.Enabled = true;
        }

        private void btnPressCancel_Click(object sender, EventArgs e)
        {
            FormEvent = eClose;
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
                        this.Close();
                    }

                    if (FormEvent == eHide)
                    {
                        MainForm.Activate();
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
                        MainForm.Activate();
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

        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);

            if (m.Msg == 6 && m.WParam.ToInt32() == 1)
            {
                if (TopForm != null)
                    TopForm.Activate();
            }
        }

        private void Initialize()
        {
            cbClients.DataSource = new DataView(ClientsDT);
            cbClients.DisplayMember = "ClientName";
            cbClients.ValueMember = "ClientID";
        }
    }
}
