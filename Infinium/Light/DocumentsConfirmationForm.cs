using System;
using System.Data;
using System.Windows.Forms;

namespace Infinium
{
    public partial class DocumentsConfirmationForm : Form
    {
        private const int eHide = 2;
        private const int eShow = 1;
        private const int eClose = 3;

        private int FormEvent;

        private Form TopForm;

        private InfiniumDocuments InfiniumDocuments;

        public bool bStopTransfer = false;

        public bool bCanceled;

        public int DocumentID = -1;
        public int DocumentCategoryID = -1;
        public int DocumentConfirmationID = -1;

        private bool bEdit;

        public DocumentsConfirmationForm(ref Form tTopForm, ref InfiniumDocuments tInfiniumDocuments, int iDocumentID, int iDocumentCategoryID)
        {
            InitializeComponent();

            InfiniumDocuments = tInfiniumDocuments;

            DocumentCategoryID = iDocumentCategoryID;
            DocumentID = iDocumentID;

            RecipientsList.ItemsDataTable = InfiniumDocuments.UsersDataTable;
            RecipientsList.InitializeItems();

            TopForm = tTopForm;
        }

        public DocumentsConfirmationForm(ref Form tTopForm, ref InfiniumDocuments tInfiniumDocuments, int iDocumentConfirmationID)
        {
            InitializeComponent();

            InfiniumDocuments = tInfiniumDocuments;

            DocumentConfirmationID = iDocumentConfirmationID;

            DataTable DT = InfiniumDocuments.GetConfirmationRecipients(iDocumentConfirmationID);

            RecipientsList.ItemsDataTable = InfiniumDocuments.UsersDataTable;
            RecipientsList.InitializeItems();

            bEdit = true;

            foreach (DataRow Row in DT.Rows)
            {
                RecipientsList.SelectItem(Convert.ToInt32(Row["UserID"]));
            }

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

        private void kryptonButton1_Click(object sender, EventArgs e)
        {
            bCanceled = true;

            FormEvent = eClose;
            AnimateTimer.Enabled = true;
        }

        private void CreateButton_Click(object sender, EventArgs e)
        {
            if (RecipientsList.GetSelectedDataTable().Rows.Count == 0)
            {
                bCanceled = true;

                FormEvent = eClose;
                AnimateTimer.Enabled = true;
            }

            if (!bEdit)
                InfiniumDocuments.AddConfirm(RecipientsList.GetSelectedDataTable(), Security.CurrentUserID, DocumentID, DocumentCategoryID);
            else
                InfiniumDocuments.EditConfirm(DocumentConfirmationID, RecipientsList.GetSelectedDataTable());

            FormEvent = eClose;
            AnimateTimer.Enabled = true;
        }

    }
}
