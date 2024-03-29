﻿using Infinium.Modules.Marketing.NewOrders;

using System;
using System.Windows.Forms;

namespace Infinium
{
    public partial class FrontsSizesExportForm : Form
    {
        private const int eHide = 2;
        private const int eShow = 1;
        private const int eClose = 3;
        private const int eMainMenu = 4;

        private int FormEvent = 0;

        private AddMarketingNewOrdersForm MainForm = null;

        private int ImportType = 0;

        /// <summary>
        /// 0 - импорт шкафов, 1 - импорт размеров
        /// </summary>
        /// <param name="ImportType"></param>
        public void SetImportType(int iImportType)
        {
            ImportType = iImportType;
        }

        private FrontsOrders FrontsOrders = null;

        public FrontsSizesExportForm(AddMarketingNewOrdersForm tMainForm, ref FrontsOrders cFrontsOrders)
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

        }

        private void CupboardsExportForm_Load(object sender, EventArgs e)
        {
            FrontsOrders.ImportFromSizeTable(ref SizeTableDataGrid);
        }

        private void SizeTableOKButton_Click(object sender, EventArgs e)
        {
            MainForm.AddFrontsFromSizeTable(ref SizeTableDataGrid);
            SizeTableDataGrid.Rows.Clear();
            FormEvent = eClose;
            AnimateTimer.Enabled = true;
        }

        private void SizeTableDeleteButton_Click(object sender, EventArgs e)
        {
            if (SizeTableDataGrid.Rows.Count > 1)
                if (SizeTableDataGrid.CurrentRow.Index != SizeTableDataGrid.Rows.Count - 1)
                    SizeTableDataGrid.Rows.Remove(SizeTableDataGrid.CurrentRow);
        }

        private void SizeTableCancelButton_Click(object sender, EventArgs e)
        {
            SizeTableDataGrid.Rows.Clear();
            FormEvent = eClose;
            AnimateTimer.Enabled = true;
        }
    }
}
