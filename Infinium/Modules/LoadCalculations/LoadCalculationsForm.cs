
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
    public partial class LoadCalculationsForm : Form
    {
        private const int eHide = 2;
        private const int eShow = 1;
        private const int eClose = 3;


        private NumberFormatInfo nfi;
        private const string DigitalFormat = "#0,0.000";

        private LoadCalculations _loadCalculations;

        private const int SectorControlWidth = 841;
        private const int SectorControlMargin = 40;
        private const int PnlForAgreedHeightMax = 170;
        private const int PnlForAgreedHeightMin = 0;

        public static bool BSmallCreated;

        private List<SectorControl> _sectorControls;

        private int FormEvent = 0;
        private bool NeedSplash = false;

        private Form TopForm = null;
        private LightStartForm LightStartForm;

        public LoadCalculationsForm(LightStartForm tLightStartForm)
        {
            InitializeComponent();
            coverPanel.BringToFront();

            this.panel3.MouseWheel += Panel3_MouseWheel;
            LightStartForm = tLightStartForm;

            this.MaximumSize = Screen.PrimaryScreen.WorkingArea.Size;


            while (!SplashForm.bCreated) ;
        }

        private void Panel3_MouseWheel(object sender, MouseEventArgs e)
        {
            int newValue = panel3.VerticalScroll.Value;

            foreach (SectorControl control in panel2.Controls)
                control.ScrollMachinePanel(newValue);
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
                this.Opacity = 1;

                if (FormEvent == eClose || FormEvent == eHide)
                {
                    AnimateTimer.Enabled = false;

                    if (FormEvent == eClose)
                    {

                        LightStartForm.CloseForm(this);
                        this.Close();
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
                if (Convert.ToDecimal(this.Opacity) != Convert.ToDecimal(0.00))
                    this.Opacity = Convert.ToDouble(Convert.ToDecimal(this.Opacity) - Convert.ToDecimal(0.05));
                else
                {
                    AnimateTimer.Enabled = false;

                    if (FormEvent == eClose)
                    {

                        LightStartForm.CloseForm(this);
                        this.Close();
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
                    NeedSplash = true;
                    AnimateTimer.Enabled = false;
                    SplashForm.CloseS = true;
                }

                return;
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
            nfi = (NumberFormatInfo)CultureInfo.InvariantCulture.NumberFormat.Clone();
            nfi.NumberDecimalSeparator = ",";
            nfi.NumberGroupSeparator = " ";

            _loadCalculations = new LoadCalculations();

            _sectorControls = new List<SectorControl>();

            pnlNotConfirmed.Height = PnlForAgreedHeightMin;
            pnlAgreed.Height = PnlForAgreedHeightMin;
            pnlForAgreed.Height = PnlForAgreedHeightMin;
            pnlOnProduction.Height = PnlForAgreedHeightMin;
            pnlInProduction.Height = PnlForAgreedHeightMin;

            lbNotConfirmedSumRank.Text = $@"Σ {0}";
            lbAgreedSumRank.Text = $@"Σ {0}";
            lbForAgreedSumRank.Text = $@"Σ {0}";
            lbOnProductionSumRank.Text = $@"Σ {0}";
            lbInProductionSumRank.Text = $@"Σ {0}";

            var sectorsList = _loadCalculations.GetAllSectors();

            const int xPos = 6;
            var yPos = 17;

            if (groupBox3.Controls.Count == 0)
                foreach (var sec in sectorsList)
                {
                    var cb = new System.Windows.Forms.CheckBox
                    {
                        AutoSize = true,
                        Checked = true,
                        Location = new System.Drawing.Point(xPos, yPos),
                        Name = "cb" + sec.Id,
                        Size = new System.Drawing.Size(114, 17),
                        Tag = sec.Id,
                        Text = sec.Name,
                        UseVisualStyleBackColor = true
                    };
                    groupBox3.Controls.Add(cb);
                    groupBox3.Height += 23;
                    yPos += 23;
                }

            LoadCalculate();

            dgvClientsNotConfirmed.DataSource = _loadCalculations.ClientsNotConfirmedList;
            dgvClientsForAgreed.DataSource = _loadCalculations.ClientsForAgreedList;
            dgvClientsAgreed.DataSource = _loadCalculations.ClientsAgreedList;
            dgvClientsOnProduction.DataSource = _loadCalculations.ClientsOnProductionList;
            dgvClientsInProduction.DataSource = _loadCalculations.ClientsInProductionList;

            if (dgvClientsNotConfirmed.Columns["orderNumber"] != null)
            {
                dgvClientsNotConfirmed.Columns["orderNumber"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                dgvClientsNotConfirmed.Columns["orderNumber"].Width = 45;
            }
            if (dgvClientsNotConfirmed.Columns["OrderDate"] != null)
            {
                dgvClientsNotConfirmed.Columns["OrderDate"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                dgvClientsNotConfirmed.Columns["OrderDate"].Width = 85;
            }
            if (dgvClientsForAgreed.Columns["orderNumber"] != null)
            {
                dgvClientsForAgreed.Columns["orderNumber"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                dgvClientsForAgreed.Columns["orderNumber"].Width = 45;
            }
            if (dgvClientsForAgreed.Columns["OrderDate"] != null)
            {
                dgvClientsForAgreed.Columns["OrderDate"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                dgvClientsForAgreed.Columns["OrderDate"].Width = 85;
            }
            if (dgvClientsAgreed.Columns["orderNumber"] != null)
            {
                dgvClientsAgreed.Columns["orderNumber"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                dgvClientsAgreed.Columns["orderNumber"].Width = 45;
            }
            if (dgvClientsAgreed.Columns["OrderDate"] != null)
            {
                dgvClientsAgreed.Columns["OrderDate"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                dgvClientsAgreed.Columns["OrderDate"].Width = 85;
            }
            if (dgvClientsOnProduction.Columns["orderNumber"] != null)
            {
                dgvClientsOnProduction.Columns["orderNumber"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                dgvClientsOnProduction.Columns["orderNumber"].Width = 45;
            }
            if (dgvClientsOnProduction.Columns["OrderDate"] != null)
            {
                dgvClientsOnProduction.Columns["OrderDate"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                dgvClientsOnProduction.Columns["OrderDate"].Width = 85;
            }
            if (dgvClientsInProduction.Columns["orderNumber"] != null)
            {
                dgvClientsInProduction.Columns["orderNumber"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                dgvClientsInProduction.Columns["orderNumber"].Width = 45;
            }
            if (dgvClientsInProduction.Columns["OrderDate"] != null)
            {
                dgvClientsInProduction.Columns["OrderDate"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                dgvClientsInProduction.Columns["OrderDate"].Width = 85;
            }

            if (dgvClientsNotConfirmed.Columns["sumTotal"] != null)
            {
                dgvClientsNotConfirmed.Columns["sumTotal"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                dgvClientsNotConfirmed.Columns["sumTotal"].Width = 65;
            }
            if (dgvClientsForAgreed.Columns["sumTotal"] != null)
            {
                dgvClientsForAgreed.Columns["sumTotal"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                dgvClientsForAgreed.Columns["sumTotal"].Width = 65;
            }
            if (dgvClientsAgreed.Columns["sumTotal"] != null)
            {
                dgvClientsAgreed.Columns["sumTotal"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                dgvClientsAgreed.Columns["sumTotal"].Width = 65;
            }
            if (dgvClientsOnProduction.Columns["sumTotal"] != null)
            {
                dgvClientsOnProduction.Columns["sumTotal"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                dgvClientsOnProduction.Columns["sumTotal"].Width = 65;
            }
            if (dgvClientsInProduction.Columns["sumTotal"] != null)
            {
                dgvClientsInProduction.Columns["sumTotal"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                dgvClientsInProduction.Columns["sumTotal"].Width = 65;
            }

        }

        private void LoadCalculationsForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (_loadCalculations == null)
            {
                //_loginForm.Close();
                return;
            }

            //_loginForm.Close();
        }

        private void LoadCalculate()
        {
            if (NeedSplash)
            {
                var T = new Thread(delegate ()
                {
                    SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка .\r\nПодождите...");
                });
                T.Start();
            }

            var sectorsId = new List<int>();
            foreach (var control in groupBox3.Controls)
            {
                var cb = (CheckBox)control;
                if (cb.Checked)
                    sectorsId.Add(Convert.ToInt32(cb.Tag));
            }

            if (rbRankCoef8.Checked)
                _loadCalculations.RankCoef = 37;
            if (rbRankCoef12.Checked)
                _loadCalculations.RankCoef = 51;

            _loadCalculations.ClearCalculations();
            _loadCalculations.CreateSectors(sectorsId);

            if (cbNotConfirmed.Checked)
            {
                _loadCalculations.GetDecorOrdersNotConfirmed(cbFronts.Checked, cbDecor.Checked);
                _loadCalculations.CalculateLoad(OrderStatus.NotConfirmed);
            }
            if (cbForAgreed.Checked)
            {
                _loadCalculations.GetDecorOrdersForAgreed(cbFronts.Checked, cbDecor.Checked);
                _loadCalculations.CalculateLoad(OrderStatus.ForAgreed);
            }
            if (cbAgreed.Checked)
            {
                _loadCalculations.GetDecorOrdersAgreed(cbFronts.Checked, cbDecor.Checked);
                _loadCalculations.CalculateLoad(OrderStatus.Agreed);
            }
            if (cbOnProduction.Checked)
            {
                _loadCalculations.GetDecorOrdersOnProduction(cbFronts.Checked, cbDecor.Checked);
                _loadCalculations.CalculateLoad(OrderStatus.OnProduction);
            }
            if (cbInProduction.Checked)
            {
                _loadCalculations.GetDecorOrdersInProduction(cbFronts.Checked, cbDecor.Checked);
                _loadCalculations.CalculateLoad(OrderStatus.InProduction);
            }

            _loadCalculations.GroupBySectors(sectorsId);
            _loadCalculations.DeleteEmptySectors();

            lbNotConfirmedSumRank.Text = $@"{_loadCalculations.NotConfirmedSumRank.ToString(DigitalFormat, nfi)}";
            lbForAgreedSumRank.Text = $@"{_loadCalculations.ForAgreedSumRank.ToString(DigitalFormat, nfi)}";
            lbAgreedSumRank.Text = $@"{_loadCalculations.AgreedSumRank.ToString(DigitalFormat, nfi)}";
            lbOnProductionSumRank.Text = $@"{_loadCalculations.OnProductionSumRank.ToString(DigitalFormat, nfi)}";
            lbInProductionSumRank.Text = $@"{_loadCalculations.InProductionSumRank.ToString(DigitalFormat, nfi)}";

            panel2.Visible = false;
            panel2.Controls.Clear();
            _sectorControls.Clear();

            //var sectorsList = _loadCalculations.GroupBySectors(sectorsId);
            var sectorsList = _loadCalculations.SectorsList;
            for (var i = 0; i < sectorsList.Count; i++)
            {
                //panel2.Controls.ContainsKey("SectorControl" + i);
                var sectorControl = new SectorControl
                {
                    Percent = sectorsList[i].Percent,
                    SumRank = sectorsList[i].SumRank,
                    Rank1 = sectorsList[i].Rank1,
                    Rank2 = sectorsList[i].Rank2,
                    Rank3 = sectorsList[i].Rank3,
                    Rank4 = sectorsList[i].Rank4,
                    SumTotal = sectorsList[i].SumTotal,
                    Total1 = sectorsList[i].Total1,
                    Total2 = sectorsList[i].Total2,
                    Total3 = sectorsList[i].Total3,
                    Total4 = sectorsList[i].Total4,
                    SectorId = sectorsList[i].Id,
                    Name = "SectorControl" + i,
                    SectorName = sectorsList[i].Name
                };

                var sectorId = sectorsList[i].Id;
                var machinesList = _loadCalculations.GroupByMachines(sectorId);
                sectorControl.AddMachines(machinesList);

                sectorControl.PForm = this;

                _sectorControls.Add(sectorControl);
                sectorControl.Dock = DockStyle.Left;
                sectorControl.Location = new Point(i * SectorControlWidth + SectorControlMargin, 5);

                panel2.Controls.Add(sectorControl);
                //sectorControl.Height = panelSectors.Height - 24;
            }
            GC.Collect();
            panel2.Visible = true;

            if (NeedSplash)
                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            CollapseNotConfirmed(VisualOrientation.Right);
            CollapseForAgreed(VisualOrientation.Right);
            CollapseAgreed(VisualOrientation.Right);
            CollapseOnProduction(VisualOrientation.Right);
            CollapseInProduction(VisualOrientation.Right);

            LoadCalculate();

            foreach (var item in _sectorControls)
                item.ExplandMachines(true);
        }

        private void btnExplandSectors_Click(object sender, EventArgs e)
        {
            var button = (KryptonButton)sender;

            var explandMachines = false;

            switch (button.Orientation)
            {
                case VisualOrientation.Right:
                    button.Orientation = VisualOrientation.Top;
                    explandMachines = false;
                    pnlsSectors.Height -= 160;
                    break;
                case VisualOrientation.Top:
                    button.Orientation = VisualOrientation.Right;
                    explandMachines = true;
                    pnlsSectors.Height += 160;
                    break;
                case VisualOrientation.Bottom:
                    break;
                case VisualOrientation.Left:
                    break;
                default:
                    throw new ArgumentOutOfRangeException();
            }

            foreach (var item in _sectorControls)
                item.ExplandMachines(explandMachines);
        }

        private void CollapseNotConfirmed(VisualOrientation orientation)
        {
            switch (orientation)
            {
                case VisualOrientation.Right:
                    foreach (SectorControl control in panel2.Controls)
                        control.ShowNotConfirmed(false);
                    btnExplandNotConfirmed.Orientation = VisualOrientation.Top;
                    pnlNotConfirmed.Height = PnlForAgreedHeightMin;
                    break;
                case VisualOrientation.Top:
                    foreach (SectorControl control in panel2.Controls)
                        control.ShowNotConfirmed(true);
                    btnExplandNotConfirmed.Orientation = VisualOrientation.Right;
                    pnlNotConfirmed.Height = PnlForAgreedHeightMax;
                    break;
                case VisualOrientation.Bottom:
                    break;
                case VisualOrientation.Left:
                    break;
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }

        private void CollapseForAgreed(VisualOrientation orientation)
        {
            switch (orientation)
            {
                case VisualOrientation.Right:
                    foreach (SectorControl control in panel2.Controls)
                        control.ShowForAgreed(false);
                    btnExplandForAgreed.Orientation = VisualOrientation.Top;
                    pnlForAgreed.Height = PnlForAgreedHeightMin;
                    break;
                case VisualOrientation.Top:
                    foreach (SectorControl control in panel2.Controls)
                        control.ShowForAgreed(true);
                    btnExplandForAgreed.Orientation = VisualOrientation.Right;
                    pnlForAgreed.Height = PnlForAgreedHeightMax;
                    break;
                case VisualOrientation.Bottom:
                    break;
                case VisualOrientation.Left:
                    break;
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }

        private void CollapseAgreed(VisualOrientation orientation)
        {
            switch (orientation)
            {
                case VisualOrientation.Right:
                    foreach (SectorControl control in panel2.Controls)
                        control.ShowAgreed(false);
                    btnExplandAgreed.Orientation = VisualOrientation.Top;
                    pnlAgreed.Height = PnlForAgreedHeightMin;
                    break;
                case VisualOrientation.Top:
                    foreach (SectorControl control in panel2.Controls)
                        control.ShowAgreed(true);
                    btnExplandAgreed.Orientation = VisualOrientation.Right;
                    pnlAgreed.Height = PnlForAgreedHeightMax;
                    break;
                case VisualOrientation.Bottom:
                    break;
                case VisualOrientation.Left:
                    break;
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }

        private void CollapseOnProduction(VisualOrientation orientation)
        {
            switch (orientation)
            {
                case VisualOrientation.Right:
                    foreach (SectorControl control in panel2.Controls)
                        control.ShowOnProduction(false);
                    btnExplandOnProduction.Orientation = VisualOrientation.Top;
                    pnlOnProduction.Height = PnlForAgreedHeightMin;
                    break;
                case VisualOrientation.Top:
                    foreach (SectorControl control in panel2.Controls)
                        control.ShowOnProduction(true);
                    btnExplandOnProduction.Orientation = VisualOrientation.Right;
                    pnlOnProduction.Height = PnlForAgreedHeightMax;
                    break;
                case VisualOrientation.Bottom:
                    break;
                case VisualOrientation.Left:
                    break;
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }

        private void CollapseInProduction(VisualOrientation orientation)
        {
            switch (orientation)
            {
                case VisualOrientation.Right:
                    foreach (SectorControl control in panel2.Controls)
                        control.ShowInProduction(false);
                    btnExplandInProduction.Orientation = VisualOrientation.Top;
                    pnlInProduction.Height = PnlForAgreedHeightMin;
                    break;
                case VisualOrientation.Top:
                    foreach (SectorControl control in panel2.Controls)
                        control.ShowInProduction(true);
                    btnExplandInProduction.Orientation = VisualOrientation.Right;
                    pnlInProduction.Height = PnlForAgreedHeightMax;
                    break;
                case VisualOrientation.Bottom:
                    break;
                case VisualOrientation.Left:
                    break;
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }

        private void btnExplandForAgreed_Click(object sender, EventArgs e)
        {
            var button = (KryptonButton)sender;
            CollapseForAgreed(button.Orientation);
        }

        private void btnExplandAgreed_Click(object sender, EventArgs e)
        {
            var button = (KryptonButton)sender;
            CollapseAgreed(button.Orientation);
        }

        private void btnExplandNotConfirmed_Click(object sender, EventArgs e)
        {
            var button = (KryptonButton)sender;
            CollapseNotConfirmed(button.Orientation);
        }

        private void btnExplandOnProduction_Click(object sender, EventArgs e)
        {
            var button = (KryptonButton)sender;
            CollapseOnProduction(button.Orientation);
        }

        private void btnExplandInProduction_Click(object sender, EventArgs e)
        {
            var button = (KryptonButton)sender;
            CollapseInProduction(button.Orientation);
        }

        public static T Clone<T>(T controlToClone) where T : Control
        {
            T instance = Activator.CreateInstance<T>();

            Type control = controlToClone.GetType();
            PropertyInfo[] info = control.GetProperties();
            object p = control.InvokeMember("", System.Reflection.BindingFlags.CreateInstance, null, controlToClone, null);
            foreach (PropertyInfo pi in info)
            {
                if ((pi.CanWrite) && !(pi.Name == "WindowTarget") && !(pi.Name == "Capture"))
                {
                    pi.SetValue(instance, pi.GetValue(controlToClone, null), null);
                }
            }
            return instance;
        }

        public void dgvNotConfirmedFirstDisplayedScrollingRowIndex(int firstDisplayedScrollingRowIndex)
        {
            dgvClientsNotConfirmed.FirstDisplayedScrollingRowIndex = firstDisplayedScrollingRowIndex;
        }

        public void dgvForAgreedFirstDisplayedScrollingRowIndex(int firstDisplayedScrollingRowIndex)
        {
            dgvClientsForAgreed.FirstDisplayedScrollingRowIndex = firstDisplayedScrollingRowIndex;
        }

        public void dgvAgreedFirstDisplayedScrollingRowIndex(int firstDisplayedScrollingRowIndex)
        {
            dgvClientsAgreed.FirstDisplayedScrollingRowIndex = firstDisplayedScrollingRowIndex;
        }

        public void dgvOnProductionFirstDisplayedScrollingRowIndex(int firstDisplayedScrollingRowIndex)
        {
            dgvClientsOnProduction.FirstDisplayedScrollingRowIndex = firstDisplayedScrollingRowIndex;
        }

        public void dgvInProductionFirstDisplayedScrollingRowIndex(int firstDisplayedScrollingRowIndex)
        {
            dgvClientsInProduction.FirstDisplayedScrollingRowIndex = firstDisplayedScrollingRowIndex;
        }

        private void dgvClientsAgreed_Scroll(object sender, ScrollEventArgs e)
        {
            foreach (SectorControl control in panel2.Controls)
                control.dgvAgreedScroll(dgvClientsAgreed.FirstDisplayedScrollingRowIndex);
        }

        private void dgvClientsOnProduction_Scroll(object sender, ScrollEventArgs e)
        {
            foreach (SectorControl control in panel2.Controls)
                control.dgvOnProductionScroll(dgvClientsOnProduction.FirstDisplayedScrollingRowIndex);
        }

        private void dgvClientsInProduction_Scroll(object sender, ScrollEventArgs e)
        {
            foreach (SectorControl control in panel2.Controls)
                control.dgvInProductionScroll(dgvClientsInProduction.FirstDisplayedScrollingRowIndex);
        }

        private void dgvClientsNotConfirmed_Scroll(object sender, ScrollEventArgs e)
        {
            foreach (SectorControl control in panel2.Controls)
                control.dgvNotConfirmedScroll(dgvClientsNotConfirmed.FirstDisplayedScrollingRowIndex);
        }

        private void dgvClientsForAgreed_Scroll(object sender, ScrollEventArgs e)
        {
            foreach (SectorControl control in panel2.Controls)
                control.dgvForAgreedScroll(dgvClientsForAgreed.FirstDisplayedScrollingRowIndex);
        }

        public void DgvNotConfirmedSelectRow(int selectedRowIndex)
        {
            try
            {
                dgvClientsNotConfirmed.Rows[selectedRowIndex].Selected = true;
            }
            catch (ArgumentOutOfRangeException argumentOutOfRangeException)
            {
                Console.Write($"Error: {argumentOutOfRangeException.Message}");
            }
        }

        public void DgvForAgreedSelectRow(int selectedRowIndex)
        {
            try
            {
                dgvClientsForAgreed.Rows[selectedRowIndex].Selected = true;
            }
            catch (ArgumentOutOfRangeException argumentOutOfRangeException)
            {
                Console.Write($"Error: {argumentOutOfRangeException.Message}");
            }
        }

        public void DgvAgreedSelectRow(int selectedRowIndex)
        {
            try
            {
                dgvClientsAgreed.Rows[selectedRowIndex].Selected = true;
            }
            catch (ArgumentOutOfRangeException argumentOutOfRangeException)
            {
                Console.Write($"Error: {argumentOutOfRangeException.Message}");
            }
        }

        public void DgvOnProductionSelectRow(int selectedRowIndex)
        {
            try
            {
                dgvClientsOnProduction.Rows[selectedRowIndex].Selected = true;
            }
            catch (ArgumentOutOfRangeException argumentOutOfRangeException)
            {
                Console.Write($"Error: {argumentOutOfRangeException.Message}");
            }
        }

        public void DgvInProductionSelectRow(int selectedRowIndex)
        {
            try
            {
                dgvClientsInProduction.Rows[selectedRowIndex].Selected = true;
            }
            catch (ArgumentOutOfRangeException argumentOutOfRangeException)
            {
                Console.Write($"Error: {argumentOutOfRangeException.Message}");
            }
        }

        private void dgvClientsOnProduction_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvClientsOnProduction.SelectedRows.Count == 0)
                return;

            foreach (SectorControl control in panel2.Controls)
                control.dgvOnProductionSelectRow(dgvClientsOnProduction.SelectedRows[0].Index);
        }

        private void dgvClientsInProduction_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvClientsInProduction.SelectedRows.Count == 0)
                return;

            foreach (SectorControl control in panel2.Controls)
                control.dgvInProductionSelectRow(dgvClientsInProduction.SelectedRows[0].Index);
        }

        private void dgvClientsAgreed_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvClientsAgreed.SelectedRows.Count == 0)
                return;

            foreach (SectorControl control in panel2.Controls)
                control.dgvAgreedSelectRow(dgvClientsAgreed.SelectedRows[0].Index);
        }

        private void dgvClientsNotConfirmed_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvClientsNotConfirmed.SelectedRows.Count == 0)
                return;

            foreach (SectorControl control in panel2.Controls)
                control.dgvNotConfirmedSelectRow(dgvClientsNotConfirmed.SelectedRows[0].Index);
        }

        private void dgvClientsForAgreed_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvClientsForAgreed.SelectedRows.Count == 0)
                return;

            foreach (SectorControl control in panel2.Controls)
                control.dgvForAgreedSelectRow(dgvClientsForAgreed.SelectedRows[0].Index);
        }

        public void ScrollClientsPanel(int newValue)
        {
            panel3.VerticalScroll.Value = newValue;
            foreach (SectorControl control in panel2.Controls)
                control.ScrollMachinePanel(newValue);
        }

        private void panel3_Scroll(object sender, ScrollEventArgs e)
        {
            int newValue = panel3.VerticalScroll.Value;

            foreach (SectorControl control in panel2.Controls)
                control.ScrollMachinePanel(newValue);
        }

        private void LoadCalculationsForm_Shown(object sender, EventArgs e)
        {
            while (!SplashForm.bCreated) ;

            NeedSplash = true;
            coverPanel.SendToBack();
            FormEvent = eShow;
            AnimateTimer.Enabled = true;
        }

        private void MinimizeButton_Click(object sender, EventArgs e)
        {
            NeedSplash = false;
            FormEvent = eHide;
            AnimateTimer.Enabled = true;
        }
    }
}
