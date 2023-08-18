using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace Infinium.Catalog.UserControls
{
    public partial class MachineControl : UserControl
    {
        private Dictionary<string, string> _groupByMachinesColumnNames;

        public MachineControl()
        {
            InitializeComponent();
            this.panel2.MouseWheel += Panel2_MouseWheel;
            CreateDictionary();
        }
        
        public void ShowNotConfirmed(bool visible)
        {
            dgvNotConfirmed.Visible = visible;
        }
        public void ShowForAgreed(bool visible)
        {
            dgvForAgreed.Visible = visible;
        }
        public void ShowAgreed(bool visible)
        {
            dgvAgreed.Visible = visible;
        }
        public void ShowOnProduction(bool visible)
        {
            dgvOnProduction.Visible = visible;
        }
        public void ShowInProduction(bool visible)
        {
            dgvInProduction.Visible = visible;
        }

        public int FirstDisplayedScrollingRowIndex
        {
            get;
            set;
        } = 0;

        public int SelectedRowIndex
        {
            get;
            set;
        } = 0;

        public object DataSource1
        {
            set {
                dgvCalculations.DataSource = value;

                SettingGrid(dgvCalculations, _groupByMachinesColumnNames);
            }
        }

        public object DataSourceAgreed
        {
            set {
                dgvAgreed.DataSource = value;

                SettingGrid(dgvAgreed, _groupByMachinesColumnNames);
            }
        }
        public object DataSourceNotConfirmed
        {
            set {
                dgvNotConfirmed.DataSource = value;

                SettingGrid(dgvNotConfirmed, _groupByMachinesColumnNames);
            }
        }
        public object DataSourceForAgreed
        {
            set {
                dgvForAgreed.DataSource = value;

                SettingGrid(dgvForAgreed, _groupByMachinesColumnNames);
            }
        }
        public object DataSourceOnProduction
        {
            set {
                dgvOnProduction.DataSource = value;

                SettingGrid(dgvOnProduction, _groupByMachinesColumnNames);
            }
        }
        public object DataSourceInProduction
        {
            set {
                dgvInProduction.DataSource = value;

                SettingGrid(dgvInProduction, _groupByMachinesColumnNames);
            }
        }

        public string MachineName
        {
            set =>
                lbName.Text = $@"{value}";
        }

        public string SumTotal
        {
            set =>
                lbSumRank.Text = $@"Σ {value}";
        }

        private void SettingGrid(DataGridView dgv, Dictionary<string, string> columnNames)
        {
            foreach (DataGridViewColumn column in dgv.Columns)
                if (columnNames.ContainsKey(column.Name.ToLower()))
                    column.HeaderText = columnNames[column.Name.ToLower()];
                else
                    column.Visible = false;
        }

        private void CreateDictionary()
        {
            _groupByMachinesColumnNames = new Dictionary<string, string> {
                { "total1", "0" }, { "total2", "2" }, { "total3", "3" }, { "total4", "4" }
            };
            _groupByMachinesColumnNames = _groupByMachinesColumnNames.ToDictionary(k => k.Key.ToLower(), k => k.Value);
        }

        public void dgvNotConfirmedScroll(int firstDisplayedScrollingRowIndex)
        {
            try
            {
                dgvNotConfirmed.FirstDisplayedScrollingRowIndex = firstDisplayedScrollingRowIndex;
            }
            catch (ArgumentOutOfRangeException argumentOutOfRangeException)
            {
                Console.Write($"Error: {argumentOutOfRangeException.Message}");
            }
        }

        public void dgvForAgreedScroll(int firstDisplayedScrollingRowIndex)
        {
            try
            {
                dgvForAgreed.FirstDisplayedScrollingRowIndex = firstDisplayedScrollingRowIndex;
            }
            catch (ArgumentOutOfRangeException argumentOutOfRangeException)
            {
                Console.Write($"Error: {argumentOutOfRangeException.Message}");
            }
        }

        public void dgvAgreedScroll(int firstDisplayedScrollingRowIndex)
        {
            try
            {
                dgvAgreed.FirstDisplayedScrollingRowIndex = firstDisplayedScrollingRowIndex;
            }
            catch (ArgumentOutOfRangeException argumentOutOfRangeException)
            {
                Console.Write($"Error: {argumentOutOfRangeException.Message}");
            }
        }

        public void dgvOnProductionScroll(int firstDisplayedScrollingRowIndex)
        {
            try
            {
                dgvOnProduction.FirstDisplayedScrollingRowIndex = firstDisplayedScrollingRowIndex;
            }
            catch (ArgumentOutOfRangeException argumentOutOfRangeException)
            {
                Console.Write($"Error: {argumentOutOfRangeException.Message}");
            }
        }

        public void dgvInProductionScroll(int firstDisplayedScrollingRowIndex)
        {
            try
            {
                dgvInProduction.FirstDisplayedScrollingRowIndex = firstDisplayedScrollingRowIndex;
            }
            catch (ArgumentOutOfRangeException argumentOutOfRangeException)
            {
                Console.Write($"Error: {argumentOutOfRangeException.Message}");
            }
        }

        public SectorControl PForm
        {
            get; set;
        }

        private void dgvAgreed_Scroll(object sender, ScrollEventArgs e)
        {
            if (e.ScrollOrientation == ScrollOrientation.VerticalScroll)
            {
                FirstDisplayedScrollingRowIndex = dgvAgreed.FirstDisplayedScrollingRowIndex;
                PForm.DgvAgreedFirstDisplayedScrollingRowIndex(dgvAgreed.FirstDisplayedScrollingRowIndex);
            }
        }

        private void dgvNotConfirmed_Scroll(object sender, ScrollEventArgs e)
        {
            if (e.ScrollOrientation == ScrollOrientation.VerticalScroll)
            {
                FirstDisplayedScrollingRowIndex = dgvNotConfirmed.FirstDisplayedScrollingRowIndex;
                PForm.DgvNotConfirmedFirstDisplayedScrollingRowIndex(dgvNotConfirmed.FirstDisplayedScrollingRowIndex);
            }
        }

        private void dgvForAgreed_Scroll(object sender, ScrollEventArgs e)
        {
            if (e.ScrollOrientation == ScrollOrientation.VerticalScroll)
            {
                FirstDisplayedScrollingRowIndex = dgvForAgreed.FirstDisplayedScrollingRowIndex;
                PForm.DgvForAgreedFirstDisplayedScrollingRowIndex(dgvForAgreed.FirstDisplayedScrollingRowIndex);
            }
        }

        private void dgvOnProduction_Scroll(object sender, ScrollEventArgs e)
        {
            if (e.ScrollOrientation == ScrollOrientation.VerticalScroll)
            {
                FirstDisplayedScrollingRowIndex = dgvOnProduction.FirstDisplayedScrollingRowIndex;
                PForm.DgvOnProductionFirstDisplayedScrollingRowIndex(dgvOnProduction.FirstDisplayedScrollingRowIndex);
            }
        }

        private void dgvInProduction_Scroll(object sender, ScrollEventArgs e)
        {
            if (e.ScrollOrientation == ScrollOrientation.VerticalScroll)
            {
                FirstDisplayedScrollingRowIndex = dgvInProduction.FirstDisplayedScrollingRowIndex;
                PForm.DgvInProductionFirstDisplayedScrollingRowIndex(dgvInProduction.FirstDisplayedScrollingRowIndex);
            }
        }

        public void dgvNotConfirmedSelectRow(int selectedRowIndex)
        {
            try
            {
                dgvNotConfirmed.Rows[selectedRowIndex].Selected = true;
            }
            catch (ArgumentOutOfRangeException argumentOutOfRangeException)
            {
                Console.Write($"Error: {argumentOutOfRangeException.Message}");
            }
        }

        public void dgvForAgreedSelectRow(int selectedRowIndex)
        {
            try
            {
                dgvForAgreed.Rows[selectedRowIndex].Selected = true;
            }
            catch (ArgumentOutOfRangeException argumentOutOfRangeException)
            {
                Console.Write($"Error: {argumentOutOfRangeException.Message}");
            }
        }

        public void dgvAgreedSelectRow(int selectedRowIndex)
        {
            try
            {
                dgvAgreed.Rows[selectedRowIndex].Selected = true;
            }
            catch (ArgumentOutOfRangeException argumentOutOfRangeException)
            {
                Console.Write($"Error: {argumentOutOfRangeException.Message}");
            }
        }

        public void dgvOnProductionSelectRow(int selectedRowIndex)
        {
            try
            {
                dgvOnProduction.Rows[selectedRowIndex].Selected = true;
            }
            catch (ArgumentOutOfRangeException argumentOutOfRangeException)
            {
                Console.Write($"Error: {argumentOutOfRangeException.Message}");
            }
        }

        public void dgvInProductionSelectRow(int selectedRowIndex)
        {
            try
            {
                dgvInProduction.Rows[selectedRowIndex].Selected = true;
            }
            catch (ArgumentOutOfRangeException argumentOutOfRangeException)
            {
                Console.Write($"Error: {argumentOutOfRangeException.Message}");
            }
        }

        private void dgvNotConfirmed_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvNotConfirmed.SelectedRows.Count > 0)
                SelectedRowIndex = dgvNotConfirmed.SelectedRows[0].Index;
            PForm.DgvNotConfirmedSelectRowIndex(SelectedRowIndex);
        }

        private void dgvForAgreed_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvForAgreed.SelectedRows.Count > 0)
                SelectedRowIndex = dgvForAgreed.SelectedRows[0].Index;
            PForm.DgvForAgreedSelectRowIndex(SelectedRowIndex);
        }

        private void dgvAgreed_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvAgreed.SelectedRows.Count > 0)
                SelectedRowIndex = dgvAgreed.SelectedRows[0].Index;
            PForm.DgvAgreedSelectRowIndex(SelectedRowIndex);
        }

        private void dgvOnProduction_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvOnProduction.SelectedRows.Count > 0)
                SelectedRowIndex = dgvOnProduction.SelectedRows[0].Index;
            PForm.DgvOnProductionSelectRowIndex(SelectedRowIndex);
        }

        private void dgvInProduction_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvInProduction.SelectedRows.Count > 0)
                SelectedRowIndex = dgvInProduction.SelectedRows[0].Index;
            PForm.DgvInProductionSelectRowIndex(SelectedRowIndex);
        }

        public void ScrollPanel(int newValue)
        {
            panel2.VerticalScroll.Value = newValue;
        }

        private void Panel2_MouseWheel(object sender, MouseEventArgs e)
        {
            PForm.ScrollClientsPanel(panel2.VerticalScroll.Value);
        }

        private void panel2_Scroll(object sender, ScrollEventArgs e)
        {
            PForm.ScrollClientsPanel(panel2.VerticalScroll.Value);
        }
    }
}