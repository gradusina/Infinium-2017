using Infinium.Modules.ZOV;

using System;
using System.Windows.Forms;

namespace Infinium
{
    public partial class SelectManagerMenu : Form
    {
        private const int eHide = 2;
        private const int eShow = 1;
        private const int eClose = 3;
        private const int eMainMenu = 4;

        private int FormEvent = 0;

        private Form MainForm = null;

        private OrdersManager OrdersManager;

        public SelectManagerMenu(Form tMainForm, OrdersManager tOrdersManager)
        {
            MainForm = tMainForm;
            OrdersManager = tOrdersManager;
            InitializeComponent();
            Initialize();
        }

        private void NewOrderButton_Click(object sender, EventArgs e)
        {
            OrdersManager.ManagerID = Convert.ToInt32(ManagersComboBox.SelectedValue);
            FormEvent = eClose;
            AnimateTimer.Enabled = true;
        }

        private void CancelOrderButton_Click(object sender, EventArgs e)
        {
            OrdersManager.ManagerID = 0;
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

        private void Initialize()
        {
            ManagersComboBox.DataSource = OrdersManager.ManagersBindingSource;
            ManagersComboBox.DisplayMember = "Name";
            ManagersComboBox.ValueMember = "ManagerID";

            if (OrdersManager.ManagerID > 0)
                ManagersComboBox.SelectedValue = OrdersManager.ManagerID;

        }
    }
}
