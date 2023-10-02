using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using Infinium.Modules.Packages.Marketing;

namespace Infinium
{
    public partial class MarketingPackagesClientForm : Form
    {
        public static bool NeedFilter = true;

        private const int eHide = 2;
        private const int eShow = 1;
        private const int eClose = 3;
        private const int eMainMenu = 4;

        private int FormEvent = 0;
        private PrintBarCode PrintBarCode;

        public string firm = "";

        private Form MainForm = null;

        public MarketingPackagesClientForm(PrintBarCode tPrintBarCode)
        {
            InitializeComponent();

            PrintBarCode = tPrintBarCode;

            Initialize();

            while (!SplashForm.bCreated) ;
        }

        private void Initialize()
        {

            cmbClients.DataSource = PrintBarCode.ClientsDataTable;
            cmbClients.DisplayMember = "ClientName";
            cmbClients.ValueMember = "ClientID";
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

        private void SplitPackagesForm_Shown(object sender, EventArgs e)
        {
            FormEvent = eShow;
            AnimateTimer.Enabled = true;
        }

        private void MarketingSplitPackagesForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            FormEvent = eClose;
            AnimateTimer.Enabled = true;
        }
        
        private void OKButton_Click(object sender, EventArgs e)
        {
            if (cbProfil.Checked)
                firm = "";
            else
                firm = ((DataRowView)cmbClients.SelectedItem)[1].ToString();

            FormEvent = eClose;
            AnimateTimer.Enabled = true;
            NeedFilter = true;
        }

        private void CancelPackButton_Click(object sender, EventArgs e)
        {
            FormEvent = eClose;
            AnimateTimer.Enabled = true;
            NeedFilter = false;
        }

        private void cbProfil_CheckedChanged(object sender, EventArgs e)
        {
            cmbClients.Enabled = !cbProfil.Checked;
        }
    }
}
