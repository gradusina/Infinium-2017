
using ComponentFactory.Krypton.Toolkit;

using Infinium.Catalog.UserControls;
using Infinium.Modules.LoadCalculations;

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;

namespace Infinium
{
    public partial class ExcelLoadCalculationsForm : Form
    {
        private const int eHide = 2;
        private const int eShow = 1;
        private const int eClose = 3;

        private ExcelLoadCalculations _excelLoadCalculations;

        public static bool BSmallCreated;
        
        private int FormEvent = 0;
        private bool NeedSplash = false;

        private Form TopForm = null;
        private LightStartForm LightStartForm;

        public ExcelLoadCalculationsForm(LightStartForm tLightStartForm)
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
        }

        private void LoadCalculationsForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (_excelLoadCalculations == null)
            {
                //_loginForm.Close();
            }

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

            if (grid.Columns.Contains("configid"))
                grid.Columns["configid"].Visible = false;
            if (grid.Columns.Contains("itemid"))
                grid.Columns["itemid"].Visible = false;
        }

        private void LoadCalculationsForm_Shown(object sender, EventArgs e)
        {
            while (!SplashForm.bCreated) ;

            NeedSplash = true;
            FormEvent = eShow;
            AnimateTimer.Enabled = true;
        }

        private void MinimizeButton_Click(object sender, EventArgs e)
        {
            NeedSplash = false;
            FormEvent = eHide;
            AnimateTimer.Enabled = true;
        }

        private void openFileDialog1_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {
            var fileInfo = new System.IO.FileInfo(openFileDialog1.FileName);

            LoadCalculationsReport _loadCalculationsReport = new LoadCalculationsReport
            {
                NeedStartFile = true
            };
            //_loadCalculationsReport.CreateReport();

            _excelLoadCalculations.ExportFile(openFileDialog1.FileName);
            _excelLoadCalculations.CreateSectors();
            _excelLoadCalculations.CalculateLoad();
            _excelLoadCalculations.CreateReportTable();

            dataGridView1.DataSource = _excelLoadCalculations.ReportDt;
            dataGridSettings(ref dataGridView1);

            //_excelLoadCalculations.GetOrdersFromExcel();
            //_excelLoadCalculations.GetOrders();


        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
        }
    }
}
