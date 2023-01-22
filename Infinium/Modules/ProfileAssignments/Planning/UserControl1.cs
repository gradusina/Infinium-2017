using ComponentFactory.Krypton.Toolkit;

using System;
using System.Windows.Forms;

namespace Infinium.Modules.ProfileAssignments.Planning
{
    public partial class UserControl1 : UserControl
    {
        public UserControl1()
        {
            InitializeComponent();
        }

        public object DataSource
        {
            set
            {
                dgvOrders.DataSource = value;
                dgvOrdersGridSettings();
            }
        }

        public int ProduceCount
        {
            set =>
                lbTotalProduceCount.Text = $@"Общее кол-во штук к производству: {value}";
        }

        public int SheetThickness
        {
            set =>
                lbSheetThickness.Text = $@"Толщина листа: {value}";
        }

        public int CollapseHeight { get; set; } = 25;
        public int ExplandHeight { get; set; } = 312;

        public int btnCollapsePanelHeight
        {
            get => btnCollapsePanel.Height;
            set => btnCollapsePanel.Height = value;
        }

        public int PnlTopHeight
        {
            get => pnlTop.Height;
            set => pnlTop.Height = value;
        }

        public void CollapseControl()
        {
            btnCollapsePanel.Orientation = VisualOrientation.Top;
            this.Height = CollapseHeight;
        }
        public void ExplandControl()
        {
            btnCollapsePanel.Orientation = VisualOrientation.Right;
            this.Height = ExplandHeight;
        }

        private void dgvOrdersGridSettings()
        {
            foreach (DataGridViewColumn col in dgvOrders.Columns)
            {
                col.ReadOnly = true;
            }
            
            if (dgvOrders.Columns.Contains("ToProduce"))
                dgvOrders.Columns["ToProduce"].ReadOnly = false;

            //dataGrid.Columns["MegaOrderID"].Visible = false;
            //dataGrid.Columns["ProductType"].Visible = false;

            //dataGrid.Columns["OrderNumber"].HeaderText = "№ заказа";
            //dataGrid.Columns["MainOrderID"].HeaderText = "№ подзаказа";
            //dataGrid.Columns["UserName"].HeaderText = "Кто сканировал";
            //dataGrid.Columns["BarcodeType"].HeaderText = "Префикс штрихкода";
            //dataGrid.Columns["ClientName"].HeaderText = "Клиент";
            //dataGrid.Columns["AddPackDateTime"].HeaderText = "Дата сканирования";
            //dataGrid.Columns["PackageID"].HeaderText = "ID упаковки";
            //dataGrid.Columns["TrayID"].HeaderText = "ID поддона";
        }
        private void btnCollapsePanel_Click(object sender, System.EventArgs e)
        {
            switch (btnCollapsePanel.Orientation)
            {
                case VisualOrientation.Right:
                    btnCollapsePanel.Orientation = VisualOrientation.Top;
                    this.Height = CollapseHeight;
                    break;
                case VisualOrientation.Top:
                    btnCollapsePanel.Orientation = VisualOrientation.Right;
                    this.Height = ExplandHeight;
                    break;
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }
    }
}
