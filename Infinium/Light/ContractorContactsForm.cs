using System;
using System.Data;
using System.Windows.Forms;

namespace Infinium
{
    public partial class ContractorContactsForm : Form
    {
        private const int eHide = 2;
        private const int eShow = 1;
        private const int eClose = 3;

        private int FormEvent;

        private Form TopForm;

        public bool bStopTransfer = false;

        public bool bCanceled;

        public DataTable ContactsDataTable;

        private int ContractorID = -1;

        public ContractorContactsForm(ref Form tTopForm, DataTable tContactsDataTable, int iContractorID)
        {
            InitializeComponent();

            ContactsDataTable = tContactsDataTable;

            ContractorID = iContractorID;

            ContactsDataGrid.DataSource = ContactsDataTable;
            ContactsDataGrid.Columns["ContractorContactID"].Visible = false;
            ContactsDataGrid.Columns["ContractorID"].Visible = false;
            ContactsDataGrid.Columns["Name"].HeaderText = "Имя";
            ContactsDataGrid.Columns["Position"].HeaderText = "Должность";
            ContactsDataGrid.Columns["Phone1"].HeaderText = "Телефон1";
            ContactsDataGrid.Columns["Phone2"].HeaderText = "Телефон2";
            ContactsDataGrid.Columns["Phone3"].HeaderText = "Телефон3";
            ContactsDataGrid.Columns["Email"].HeaderText = "E-mail";
            ContactsDataGrid.Columns["Website"].HeaderText = "Веб-сайт";
            ContactsDataGrid.Columns["Skype"].HeaderText = "Skype";
            ContactsDataGrid.Columns["Description"].HeaderText = "Описание";

            TopForm = tTopForm;
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

        private void AddNewsForm_Shown(object sender, EventArgs e)
        {
            FormEvent = eShow;
            AnimateTimer.Enabled = true;
        }

        private void CreateButton_Click(object sender, EventArgs e)
        {
            if (ContractorID > -1)
            {
                foreach (DataRow Row in ContactsDataTable.Rows)
                {
                    if (Row.RowState == DataRowState.Deleted)
                        continue;

                    Row["ContractorID"] = ContractorID;
                }
            }

            FormEvent = eClose;
            AnimateTimer.Enabled = true;
        }

        private void CancelsButton_Click(object sender, EventArgs e)
        {
            bCanceled = true;
            ContactsDataTable.Clear();

            FormEvent = eClose;
            AnimateTimer.Enabled = true;
        }

    }
}
