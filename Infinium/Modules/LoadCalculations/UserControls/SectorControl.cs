using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;
using Infinium.Modules.LoadCalculations;

namespace Infinium.Catalog.UserControls
{
    public partial class SectorControl : UserControl
    {
        private NumberFormatInfo nfi;
        private const string DigitalFormat = "#0.000";

        private const int MachineControlWidth = 400;
        private const int MachineControlMargin = 40;

        public SectorControl()
        {
            InitializeComponent();
            SetStyle(ControlStyles.UserPaint
                     | ControlStyles.OptimizedDoubleBuffer
                     | ControlStyles.AllPaintingInWmPaint
                     | ControlStyles.SupportsTransparentBackColor, true);
            nfi = (NumberFormatInfo)CultureInfo.InvariantCulture.NumberFormat.Clone();

            nfi.NumberDecimalSeparator = ",";
            nfi.NumberGroupSeparator = " ";
        }

        public string SectorName
        {
            set =>
                lbName.Text = $@"{value}";
        }

        public decimal Percent
        {
            set =>
                lbPercent.Text = $@"{value.ToString(DigitalFormat, nfi)} %";
        }

        public decimal SumRank
        {
            set =>
                lbSumRank.Text = $@"Σ {value.ToString(DigitalFormat, nfi)}";
        }

        public decimal Rank1
        {
            set =>
                lbRank1.Text = $@"#0: {value.ToString(DigitalFormat, nfi)}";
        }

        public decimal Rank2
        {
            set =>
                lbRank2.Text = $@"#2: {value.ToString(DigitalFormat, nfi)}";
        }

        public decimal Rank3
        {
            set =>
                lbRank3.Text = $@"#3: {value.ToString(DigitalFormat, nfi)}";
        }

        public decimal Rank4
        {
            set =>
                lbRank4.Text = $@"#4: {value.ToString(DigitalFormat, nfi)}";
        }

        public decimal SumTotal
        {
            set =>
                lbSumTotal.Text = $@"Σ {value.ToString(DigitalFormat, nfi)}";
        }

        public decimal Total1
        {
            set => lbTotal1.Text = $@"#0: {value.ToString(DigitalFormat, nfi)}";
        }

        public decimal Total2
        {
            set =>
                lbTotal2.Text = $@"#2: {value.ToString(DigitalFormat, nfi)}";
        }

        public decimal Total3
        {
            set =>
                lbTotal3.Text = $@"#3: {value.ToString(DigitalFormat, nfi)}";
        }

        public decimal Total4
        {
            set =>
                lbTotal4.Text = $@"#4: {value.ToString(DigitalFormat, nfi)}";
        }

        public int SectorId
        {
            get;
            set;
        } = 0;

        public LoadCalculationsForm PForm
        {
            get; set;
        }

        public void AddMachines(List<LoadCalculations.Machine> machinesList)
        {
            panel4.Controls.Clear();

            for (var i = 0; i < machinesList.Count; i++)
            {
                var control1 = new MachineControl
                {
                    DataSource1 = machinesList[i].DataSource1,
                    DataSourceNotConfirmed = machinesList[i].DataSourceNotConfirmed,
                    DataSourceForAgreed = machinesList[i].DataSourceForAgreed,
                    DataSourceAgreed = machinesList[i].DataSourceAgreed,
                    DataSourceOnProduction = machinesList[i].DataSourceOnProduction,
                    DataSourceInProduction = machinesList[i].DataSourceInProduction,
                    Name = "MachineControl" + machinesList[i].Id,
                    SumTotal = machinesList[i].SumTotal.ToString(),
                    MachineName = machinesList[i].Name
                };
                control1.PForm = this;
                control1.Dock = DockStyle.Left;
                control1.Location = new Point(i * MachineControlWidth + MachineControlMargin, 5);
                control1.Margin = new Padding(0, 0, 15, 0);
                panel4.Controls.Add(control1);
            }
        }

        public void DgvNotConfirmedFirstDisplayedScrollingRowIndex(int firstDisplayedScrollingRowIndex)
        {
            PForm.dgvNotConfirmedFirstDisplayedScrollingRowIndex(firstDisplayedScrollingRowIndex);
        }
        public void DgvForAgreedFirstDisplayedScrollingRowIndex(int firstDisplayedScrollingRowIndex)
        {
            PForm.dgvForAgreedFirstDisplayedScrollingRowIndex(firstDisplayedScrollingRowIndex);
        }
        public void DgvAgreedFirstDisplayedScrollingRowIndex(int firstDisplayedScrollingRowIndex)
        {
            PForm.dgvAgreedFirstDisplayedScrollingRowIndex(firstDisplayedScrollingRowIndex);
        }
        public void DgvOnProductionFirstDisplayedScrollingRowIndex(int firstDisplayedScrollingRowIndex)
        {
            PForm.dgvOnProductionFirstDisplayedScrollingRowIndex(firstDisplayedScrollingRowIndex);
        }

        public void DgvInProductionFirstDisplayedScrollingRowIndex(int firstDisplayedScrollingRowIndex)
        {
            PForm.dgvInProductionFirstDisplayedScrollingRowIndex(firstDisplayedScrollingRowIndex);
        }

        public void DgvNotConfirmedSelectRowIndex(int selectedRowIndex)
        {
            PForm.DgvNotConfirmedSelectRow(selectedRowIndex);
        }

        public void DgvForAgreedSelectRowIndex(int selectedRowIndex)
        {
            PForm.DgvForAgreedSelectRow(selectedRowIndex);
        }

        public void DgvAgreedSelectRowIndex(int selectedRowIndex)
        {
            PForm.DgvAgreedSelectRow(selectedRowIndex);
        }

        public void DgvOnProductionSelectRowIndex(int selectedRowIndex)
        {
            PForm.DgvOnProductionSelectRow(selectedRowIndex);
        }

        public void DgvInProductionSelectRowIndex(int selectedRowIndex)
        {
            PForm.DgvInProductionSelectRow(selectedRowIndex);
        }

        private void CollapsePanel()
        {
            Height -= 150;
            panel4.Visible = false;
        }

        private void ExplandPanel()
        {
            Height += 150;
            panel4.Visible = true;
        }

        public void dgvNotConfirmedScroll(int firstDisplayedScrollingRowIndex)
        {
            foreach (MachineControl control in panel4.Controls)
            {
                control.dgvNotConfirmedScroll(firstDisplayedScrollingRowIndex);
            }
        }

        public void dgvForAgreedScroll(int firstDisplayedScrollingRowIndex)
        {
            foreach (MachineControl control in panel4.Controls)
            {
                control.dgvForAgreedScroll(firstDisplayedScrollingRowIndex);
            }
        }

        public void dgvAgreedScroll(int firstDisplayedScrollingRowIndex)
        {
            foreach (MachineControl control in panel4.Controls)
            {
                control.dgvAgreedScroll(firstDisplayedScrollingRowIndex);
            }
        }

        public void dgvOnProductionScroll(int firstDisplayedScrollingRowIndex)
        {
            foreach (MachineControl control in panel4.Controls)
            {
                control.dgvOnProductionScroll(firstDisplayedScrollingRowIndex);
            }
        }

        public void dgvInProductionScroll(int firstDisplayedScrollingRowIndex)
        {
            foreach (MachineControl control in panel4.Controls)
            {
                control.dgvInProductionScroll(firstDisplayedScrollingRowIndex);
            }
        }

        public void dgvNotConfirmedSelectRow(int selectedRowIndex)
        {
            foreach (MachineControl control in panel4.Controls)
            {
                control.dgvNotConfirmedSelectRow(selectedRowIndex);
            }
        }

        public void dgvForAgreedSelectRow(int selectedRowIndex)
        {
            foreach (MachineControl control in panel4.Controls)
            {
                control.dgvForAgreedSelectRow(selectedRowIndex);
            }
        }

        public void dgvAgreedSelectRow(int selectedRowIndex)
        {
            foreach (MachineControl control in panel4.Controls)
            {
                control.dgvAgreedSelectRow(selectedRowIndex);
            }
        }

        public void dgvOnProductionSelectRow(int selectedRowIndex)
        {
            foreach (MachineControl control in panel4.Controls)
            {
                control.dgvOnProductionSelectRow(selectedRowIndex);
            }
        }

        public void dgvInProductionSelectRow(int selectedRowIndex)
        {
            foreach (MachineControl control in panel4.Controls)
            {
                control.dgvInProductionSelectRow(selectedRowIndex);
            }
        }

        public void ShowAgreed(bool visible)
        {
            foreach (MachineControl control in panel4.Controls)
                control.ShowAgreed(visible);
        }

        public void ShowNotConfirmed(bool visible)
        {
            foreach (MachineControl control in panel4.Controls)
                control.ShowNotConfirmed(visible);
        }

        public void ShowForAgreed(bool visible)
        {
            foreach (MachineControl control in panel4.Controls)
                control.ShowForAgreed(visible);
        }

        public void ShowOnProduction(bool visible)
        {
            foreach (MachineControl control in panel4.Controls)
                control.ShowOnProduction(visible);
        }

        public void ShowInProduction(bool visible)
        {
            foreach (MachineControl control in panel4.Controls)
                control.ShowInProduction(visible);
        }

        public void ScrollClientsPanel(int newValue)
        {
            PForm.ScrollClientsPanel(newValue);
            foreach (MachineControl control in panel4.Controls)
                control.ScrollPanel(newValue);
        }

        public void ScrollMachinePanel(int newValue)
        {
            foreach (MachineControl control in panel4.Controls)
                control.ScrollPanel(newValue);
        }

        public void ExplandMachines(bool Expland)
        {
            switch (Expland)
            {
                case false:
                    CollapsePanel();

                    break;
                case true:
                    ExplandPanel();

                    break;
                default:
                    throw new ArgumentOutOfRangeException();
            }

            //if (this.CollapseClick != null)
            //    CollapseClick(sender, e);
        }
    }
}