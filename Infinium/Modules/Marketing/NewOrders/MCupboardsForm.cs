﻿using Infinium.Modules.Marketing.NewOrders;

using System;
using System.Windows.Forms;

namespace Infinium
{
    public partial class MCupboardsForm : Form
    {
        private const int eHide = 2;
        private const int eShow = 1;
        private const int eClose = 3;
        private const int eMainMenu = 4;

        private int FormEvent = 0;

        private Form MainForm;
        private FrontsOrders FrontsOrders;

        public MCupboardsForm(Form tMainForm, ref FrontsOrders cFrontsOrders)
        {
            MainForm = tMainForm;
            FrontsOrders = cFrontsOrders;
            InitializeComponent();
            Initialize();
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
            if (CupboardsComboBox.DataSource == null)
            {
                CupboardsComboBox.DataSource = FrontsOrders.CupboardsBindingSource;
                CupboardsComboBox.DisplayMember = "CupboardName";
                CupboardsComboBox.ValueMember = "CupboardsID";
            }
        }

        private void CupboardsPanelSaveButton_Click(object sender, EventArgs e)
        {
            FrontsOrders.SaveCupboards();
        }

        private void CupboardsPanelRemoveButton_Click(object sender, EventArgs e)
        {
            FrontsOrders.RemoveCupboard();
        }

        private void CupboardsPanelAddButton_Click(object sender, EventArgs e)
        {
            FrontsOrders.AddCupboard(CupboardsComboBox.Text);
        }

        private void CupboardsForm_Load(object sender, EventArgs e)
        {
            if (CupboardsDataGrid.DataSource == null)
                FrontsOrders.SetGrid(ref CupboardsDataGrid);
        }

        private void CupboardsPanelOKButton_Click(object sender, EventArgs e)
        {
            FormEvent = eClose;
            AnimateTimer.Enabled = true;
        }
    }
}
