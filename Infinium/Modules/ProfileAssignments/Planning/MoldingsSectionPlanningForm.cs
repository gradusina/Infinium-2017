
using Infinium.Modules.ProfileAssignments.Planning;

using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace Infinium
{
    public partial class MoldingsSectionPlanningForm : Form
    {
        const int eHide = 2;
        const int eShow = 1;
        const int eClose = 3;

        //const int iAdminRole = 107;

        int FormEvent = 0;

        Form TopForm = null;
        LightStartForm LightStartForm;

        private ClientsManager ClientsManager;

        //private PermissionsManager PermissionsManager;
        MoldingsSectionPlanning moldingsSectionPlanning;

        //RoleTypes RoleType = RoleTypes.OrdinaryRole;

        //public enum RoleTypes
        //{
        //    OrdinaryRole = 0,
        //    AdminRole = 1
        //}

        public MoldingsSectionPlanningForm(LightStartForm tLightStartForm)
        {
            InitializeComponent();

            LightStartForm = tLightStartForm;

            this.MaximumSize = Screen.PrimaryScreen.WorkingArea.Size;


            while (!SplashForm.bCreated) ;
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

        private void AnimateTimer_Tick(object sender, EventArgs e)
        {
            if (!DatabaseConfigsManager.Animation)
            {
                this.Opacity = 1;

                if (FormEvent == eClose ||
                    FormEvent == eHide)
                {
                    AnimateTimer.Enabled = false;

                    if (FormEvent == eClose)
                    {

                        LightStartForm.CloseForm(this);
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
                    return;
                }

            }

            if (FormEvent == eClose ||
                FormEvent == eHide)
            {
                if (Convert.ToDecimal(this.Opacity) != Convert.ToDecimal(0.00))
                    this.Opacity = Convert.ToDouble(Convert.ToDecimal(this.Opacity) - Convert.ToDecimal(0.05));
                else
                {
                    AnimateTimer.Enabled = false;

                    if (FormEvent == eClose)
                    {

                        LightStartForm.CloseForm(this);
                    }

                    if (FormEvent == eHide)
                    {

                        LightStartForm.HideForm(this);
                    }

                }

                return;
            }


            if (FormEvent == eShow ||
                FormEvent == eShow)
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

        private void RegisterManualInputForm_Shown(object sender, EventArgs e)
        {
            while (!SplashForm.bCreated) ;

            FormEvent = eShow;
            AnimateTimer.Enabled = true;
        }

        private void MenuCloseButton_Click(object sender, EventArgs e)
        {
            FormEvent = eClose;
            AnimateTimer.Enabled = true;
        }

        private void RegisterManualInputForm_Load(object sender, EventArgs e)
        {
            ClientsManager = new ClientsManager();
            //PermissionsManager = new PermissionsManager();

            lbClients.DisplayMember = "ClientName";
            lbClients.ValueMember = "ClientId";
            lbClients.DataSource = ClientsManager.clientsBs;

            lbOrderNumbers.DisplayMember = "OrderNumber";
            lbOrderNumbers.ValueMember = "MegaOrderID";
            lbOrderNumbers.DataSource = ClientsManager.orderNumbersBs;

            moldingsSectionPlanning = new MoldingsSectionPlanning();

            dgvProdOrders.DataSource = moldingsSectionPlanning.CurrentProdOrdersTable();
            dgvProdOrdersGridSettings();

            //dgvFrontsOrdersOnStore.DataSource = packagesOnStore.FrontsOrdersList;
            //dgvDecorOrdersOnStore.DataSource = packagesOnStore.DecorOrdersList;

            //dgvFrontsOrdersSetting(ref dgvFrontsOrdersOnStore);
            //dgvDecorOrdersSetting(ref dgvDecorOrdersOnStore);

            //RegisterManualInput = new RegisterManualInput();

            //PermissionsManager.GetPermissions(Security.CurrentUserID, this.Name);
            //if (PermissionsManager.PermissionGranted(iAdminRole))
            //{
            //    RoleType = RoleTypes.AdminRole;
            //    toolStripMenuItem1.Enabled = true;
            //}

            //dgvOrders.DataSource = RegisterManualInput.PackagesOnStoreBs;

            //dgvPackagesGridsSettings(ref dgvOrders);

            //RegisterManualInput.RefreshData();
        }

        private void dgvProdOrdersGridSettings()
        {
            DataGridViewButtonColumn downPriorityColumn =
                new DataGridViewButtonColumn
                {
                    HeaderText = "",
                    Name = "DownPriority",
                    Text = "Вниз",
                    UseColumnTextForButtonValue = true
                };
            DataGridViewButtonColumn upPriorityColumn =
                new DataGridViewButtonColumn
                {
                    HeaderText = "",
                    Name = "UpPriority",
                    Text = "Вверх",
                    UseColumnTextForButtonValue = true
                };

            dgvProdOrders.Columns.Add(moldingsSectionPlanning.MillingOptionColumn);
            dgvProdOrders.Columns.Add(moldingsSectionPlanning.FacingOptionColumn);
            dgvProdOrders.Columns.Add(downPriorityColumn);
            dgvProdOrders.Columns.Add(upPriorityColumn);

            foreach (DataGridViewColumn col in dgvProdOrders.Columns)
            {
                col.ReadOnly = true;
            }

            dgvProdOrders.Columns["MillingOptionColumn"].ReadOnly = false;
            dgvProdOrders.Columns["FacingOptionColumn"].ReadOnly = false;

            foreach (string colName in moldingsSectionPlanning.ColNamesHide)
                dgvProdOrders.Columns[colName].Visible = false;

        }

        private void dgvOrders_SelectionChanged(object sender, EventArgs e)
        {
            //int mainOrderId = -1;
            //int packageId = -1;
            //int productType = -1;
            //if (dgvOrders.SelectedRows.Count > 0 && dgvOrders.SelectedRows[0].Cells["PackageID"].Value != DBNull.Value)
            //    packageId = Convert.ToInt32(dgvOrders.SelectedRows[0].Cells["PackageID"].Value);
            //if (dgvOrders.SelectedRows.Count > 0 && dgvOrders.SelectedRows[0].Cells["MainOrderID"].Value != DBNull.Value)
            //    mainOrderId = Convert.ToInt32(dgvOrders.SelectedRows[0].Cells["MainOrderID"].Value);
            //if (dgvOrders.SelectedRows.Count > 0 && dgvOrders.SelectedRows[0].Cells["ProductType"].Value != DBNull.Value)
            //    productType = Convert.ToInt32(dgvOrders.SelectedRows[0].Cells["ProductType"].Value);

            //FilterOrdersOnStore(packageId, productType, mainOrderId);
        }

        private void lbClients_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lbClients.SelectedIndex < 0 ||
                lbClients.SelectedItem == null)
                return;

            int id = Convert.ToInt32(lbClients.SelectedValue);

            string name = lbClients.SelectedItem.ToString();

            ClientsManager.FillOrderNumberDt(id);
        }

        private void SearchOrders()
        {
            bool groupByThickness = cbGroupByThickness.Checked;

            if (groupByThickness)
            {
                bool allClients = cbAllClients.Checked;
                bool showZero = cbShowZero.Checked;
                int clientId = -1;

                if (lbClients.SelectedIndex > -1 &&
                    lbClients.SelectedItem != null)
                    clientId = Convert.ToInt32(lbClients.SelectedValue);

                moldingsSectionPlanning.ClearOrders();

                moldingsSectionPlanning.UpdateOrders(ClientsManager, groupByThickness,
                    rbOnProd.Checked, rbInProd.Checked, rbOnAgreement.Checked, allClients, showZero, clientId);

                flowLayoutPanel1.SuspendLayout();
                flowLayoutPanel1.Controls.Clear();

                for (int i = 0; i < moldingsSectionPlanning.DistinctThicknessTables; i++)
                {
                    UserControl1 control1 = new UserControl1()
                    {
                        Width = flowLayoutPanel1.Width,
                        Location = new Point(0, i * 312),
                        Name = "ProfilUserFunctionsContainer" + i,
                        DataSource = moldingsSectionPlanning.CurrentTable(i),
                        ProduceCount = moldingsSectionPlanning.TotalProduceCount(i),
                        SheetThickness = moldingsSectionPlanning.CurrentTableThickness(i)
                    };
                    flowLayoutPanel1.Controls.Add(control1);

                    control1.Width = flowLayoutPanel1.Width - 25;
                    control1.CollapseHeight = control1.PnlTopHeight;
                    control1.ExplandHeight = control1.Height;
                    control1.CollapseControl();
                }

                flowLayoutPanel1.ResumeLayout();
            }
            else
            {
                bool allClients = cbAllClients.Checked;
                bool showZero = cbShowZero.Checked;
                int clientId = -1;

                if (lbClients.SelectedIndex > -1 &&
                    lbClients.SelectedItem != null)
                    clientId = Convert.ToInt32(lbClients.SelectedValue);

                moldingsSectionPlanning.ClearOrders();

                moldingsSectionPlanning.UpdateOrders(ClientsManager, groupByThickness,
                    rbOnProd.Checked, rbInProd.Checked, rbOnAgreement.Checked, allClients, showZero, clientId);

                flowLayoutPanel1.SuspendLayout();
                flowLayoutPanel1.Controls.Clear();

                for (int i = 0; i < 1; i++)
                {
                    UserControl1 control1 = new UserControl1()
                    {
                        Width = flowLayoutPanel1.Width,
                        Location = new Point(0, i * 312),
                        Name = "ProfilUserFunctionsContainer" + i,
                        DataSource = moldingsSectionPlanning.CurrentTable(i),
                        ProduceCount = moldingsSectionPlanning.TotalProduceCount(i),
                        SheetThickness = moldingsSectionPlanning.CurrentTableThickness(i)
                    };
                    flowLayoutPanel1.Controls.Add(control1);

                    control1.Width = flowLayoutPanel1.Width - 25;
                    control1.CollapseHeight = control1.PnlTopHeight;
                    control1.ExplandHeight = control1.Height;
                    control1.ExplandControl();
                }

                flowLayoutPanel1.ResumeLayout();
            }
        }

        private void SearchProdOrders()
        {
            moldingsSectionPlanning.ClearProdOrders();

            moldingsSectionPlanning.UpdateProdOrders();

        }

        private void btnGetOrders_Click(object sender, EventArgs e)
        {
            SearchOrders();
        }

        private void cbAllClients_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cb = (CheckBox)sender;

            panelClient.Enabled = !cb.Checked;
        }

        private void btnSaveProduceOrders_Click(object sender, EventArgs e)
        {
            moldingsSectionPlanning.SaveOrdersToDb();
            SearchOrders();
        }

        private void btnGetProdOrders_Click(object sender, EventArgs e)
        {
            SearchProdOrders();

        }

        private void button2_Click(object sender, EventArgs e)
        {
            UpPriority();
        }

        private void UpPriority()
        {
            int id = -1;
            int priority = -1;

            if (dgvProdOrders.SelectedCells.Count != 0 &&
                dgvProdOrders.CurrentRow?.Cells["Id"].Value != null)
                id = Convert.ToInt32(dgvProdOrders.CurrentRow?.Cells["Id"].Value);
            if (dgvProdOrders.SelectedCells.Count != 0 &&
                dgvProdOrders.CurrentRow?.Cells["priority"].Value != null)
                priority = Convert.ToInt32(dgvProdOrders.CurrentRow?.Cells["priority"].Value);

            moldingsSectionPlanning.UpPriority(id, priority);
            dgvProdOrders.Sort(dgvProdOrders.Columns["Priority"], ListSortDirection.Ascending);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DownPriority();
        }

        private void DownPriority()
        {
            int id = -1;
            int priority = -1;

            if (dgvProdOrders.SelectedCells.Count != 0 &&
                dgvProdOrders.CurrentRow?.Cells["Id"].Value != null)
                id = Convert.ToInt32(dgvProdOrders.CurrentRow?.Cells["Id"].Value);
            if (dgvProdOrders.SelectedCells.Count != 0 &&
                dgvProdOrders.CurrentRow?.Cells["priority"].Value != null)
                priority = Convert.ToInt32(dgvProdOrders.CurrentRow?.Cells["priority"].Value);

            moldingsSectionPlanning.DownPriority(id, priority);
            dgvProdOrders.Sort(dgvProdOrders.Columns["Priority"], ListSortDirection.Ascending);
        }

        private void dgvProdOrders_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            var senderGrid = (DataGridView)sender;

            if (senderGrid.Columns[e.ColumnIndex] is DataGridViewButtonColumn &&
                senderGrid.Columns[e.ColumnIndex].Name == "DownPriority" &&
                e.RowIndex >= 0)
            {
                DownPriority();
            }

            if (senderGrid.Columns[e.ColumnIndex] is DataGridViewButtonColumn &&
                senderGrid.Columns[e.ColumnIndex].Name == "UpPriority" &&
                e.RowIndex >= 0)
            {
                UpPriority();
            }
        }

        private void dgvProdOrders_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
        }

        private void dgvProdOrders_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            var senderGrid = (DataGridView)sender;

            if (senderGrid.Columns[e.ColumnIndex] is DataGridViewComboBoxColumn &&
                senderGrid.Columns[e.ColumnIndex].Name == "MillingOptionColumn" &&
                e.RowIndex >= 0)
            {

                int decorId = -1;

                if (dgvProdOrders.SelectedCells.Count != 0 &&
                    dgvProdOrders.CurrentRow?.Cells["decorId"].Value != null)
                    decorId = Convert.ToInt32(dgvProdOrders.CurrentRow?.Cells["decorId"].Value);
                //moldingsSectionPlanning.FilterMillingMachines(decorId);
            }
            if (senderGrid.Columns[e.ColumnIndex] is DataGridViewComboBoxColumn &&
                senderGrid.Columns[e.ColumnIndex].Name == "FacingOptionColumn" &&
                e.RowIndex >= 0)
            {

                int decorId = -1;

                if (dgvProdOrders.SelectedCells.Count != 0 &&
                    dgvProdOrders.CurrentRow?.Cells["decorId"].Value != null)
                    decorId = Convert.ToInt32(dgvProdOrders.CurrentRow?.Cells["decorId"].Value);
                //moldingsSectionPlanning.FilterFacingMachines(decorId);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            moldingsSectionPlanning.SaveProdOrdersToDb();
            SearchProdOrders();
        }
    }
}
