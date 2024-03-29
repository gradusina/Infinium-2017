﻿using System;
using System.Data;
using System.IO;
using System.Windows.Forms;

namespace Infinium
{
    public partial class NewToolsSellerForm : Form
    {
        private const int eHide = 2;
        private const int eShow = 1;
        private const int eClose = 3;

        private int FormEvent;

        private Form TopForm = null;
        private ToolsSellersManager ToolsSellersManager;
        private DataTable ToolsSellerInfoDataTable;
        private int Index;

        public NewToolsSellerForm(ref ToolsSellersManager tToolsSellersManager, int tIndex)
        {
            InitializeComponent();

            MaximumSize = Screen.PrimaryScreen.WorkingArea.Size;

            Index = tIndex;
            ToolsSellersManager = tToolsSellersManager;

            if (Index == -1)
            {
                ToolsSellerInfoDataTable = ToolsSellersManager.InfoDataTableClone;
            }
            else
            {
                NameTextBox.Text = ToolsSellersManager.CurrentToolsSellerName;
                CountryTextBox.Text = ToolsSellersManager.CurrentToolsSellerCountry;
                AddressTextBox.Text = ToolsSellersManager.CurrentToolsSellerAddress;
                ContractDocNumberTextBox.Text = ToolsSellersManager.CurrentContractDocNumber;
                EmailTextBox.Text = ToolsSellersManager.CurrentToolsSellerEmail;
                SiteTextBox.Text = ToolsSellersManager.CurrentToolsSellerSite;
                NotesTextBox.Text = ToolsSellersManager.CurrentToolsSellerNotes;

                ToolsSellerInfoDataTable = ToolsSellersManager.InfoDataTableCopy;
            }

            ToolsSellerInfoDataGrid.DataSource = ToolsSellerInfoDataTable;

            Initialize();

            while (!SplashForm.bCreated) ;
        }

        private void Initialize()
        {
            ToolsSellerInfoDataGrid.Columns["Name"].MinimumWidth = 150;
            ToolsSellerInfoDataGrid.Columns["Name"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            ToolsSellerInfoDataGrid.Columns["Position"].MinimumWidth = 150;
            ToolsSellerInfoDataGrid.Columns["Position"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            ToolsSellerInfoDataGrid.Columns["Phone"].MinimumWidth = 150;
            ToolsSellerInfoDataGrid.Columns["Phone"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            ToolsSellerInfoDataGrid.Columns["Email"].MinimumWidth = 150;
            ToolsSellerInfoDataGrid.Columns["Email"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            ToolsSellerInfoDataGrid.Columns["ICQ"].MinimumWidth = 150;
            ToolsSellerInfoDataGrid.Columns["ICQ"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            ToolsSellerInfoDataGrid.Columns["Skype"].MinimumWidth = 150;
            ToolsSellerInfoDataGrid.Columns["Skype"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            ToolsSellerInfoDataGrid.Columns["Name"].HeaderText = "Имя";
            ToolsSellerInfoDataGrid.Columns["Position"].HeaderText = "Должность";
            ToolsSellerInfoDataGrid.Columns["Phone"].HeaderText = "Телефон";
            ToolsSellerInfoDataGrid.Columns["Email"].HeaderText = "E-mail";
            ToolsSellerInfoDataGrid.Columns["ICQ"].HeaderText = "ICQ";
            ToolsSellerInfoDataGrid.Columns["Skype"].HeaderText = "Skype";
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
                Opacity = 1;

                if (FormEvent == eClose || FormEvent == eHide)
                {
                    AnimateTimer.Enabled = false;

                    if (FormEvent == eClose)
                    {
                        Close();
                    }

                    if (FormEvent == eHide)
                    {
                        Hide();
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
                if (Convert.ToDecimal(Opacity) != Convert.ToDecimal(0.00))
                    Opacity = Convert.ToDouble(Convert.ToDecimal(Opacity) - Convert.ToDecimal(0.05));
                else
                {
                    AnimateTimer.Enabled = false;

                    if (FormEvent == eClose)
                    {
                        Close();
                    }

                    if (FormEvent == eHide)
                    {
                        Hide();
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
                    AnimateTimer.Enabled = false;
                    SplashForm.CloseS = true;
                }
            }
        }

        private void NewSellersForm_Shown(object sender, EventArgs e)
        {
            FormEvent = eShow;
            AnimateTimer.Enabled = true;
        }

        private void MenuCloseButton_Click(object sender, EventArgs e)
        {
            FormEvent = eClose;
            AnimateTimer.Enabled = true;
        }

        private void SaveButton_Click(object sender, EventArgs e)
        {
            if (NameTextBox.Text != string.Empty)
            {
                using (StringWriter SW = new StringWriter())
                {
                    ToolsSellerInfoDataTable.WriteXml(SW);

                    if (Index == -1)
                    {
                        ToolsSellersManager.AddSeller(NameTextBox.Text, CountryTextBox.Text, AddressTextBox.Text,
                              ContractDocNumberTextBox.Text, EmailTextBox.Text, SiteTextBox.Text, SW.ToString(), NotesTextBox.Text);
                    }
                    else
                    {
                        ToolsSellersManager.EditSeller(NameTextBox.Text, CountryTextBox.Text, AddressTextBox.Text,
                           ContractDocNumberTextBox.Text, EmailTextBox.Text, SiteTextBox.Text, SW.ToString(), NotesTextBox.Text);
                    }
                }
            }
            FormEvent = eClose;
            AnimateTimer.Enabled = true;
        }
    }
}
