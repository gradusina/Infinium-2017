﻿using System;
using System.Data;
using System.Linq;
using System.Windows.Forms;

namespace Infinium
{
    public partial class AddDocumentRecipientsForm : Form
    {
        private const int eHide = 2;
        private const int eShow = 1;
        private const int eClose = 3;

        private int FormEvent;

        private Form TopForm;

        private InfiniumDocuments InfiniumDocuments;

        public bool bCanceled;

        public int DocumentID = -1;
        public int DocumentCategoryID = -1;

        private DataTable UsersDT;

        public AddDocumentRecipientsForm(ref Form tTopForm, ref InfiniumDocuments tInfiniumDocuments, int iDocumentID, int iDocumentCategoryID)
        {
            InitializeComponent();

            InfiniumDocuments = tInfiniumDocuments;

            DocumentCategoryID = iDocumentCategoryID;
            DocumentID = iDocumentID;

            UsersDT = InfiniumDocuments.UsersDataTable.Clone();

            DataTable RecDT = InfiniumDocuments.GetDocumentsRecipients(DocumentCategoryID, DocumentID);

            foreach (DataRow Row in InfiniumDocuments.UsersDataTable.Rows)
            {
                if (RecDT.Select("UserID = " + Row["UserID"]).Count() > 0)
                    continue;

                if (Convert.ToInt32(Row["UserID"]) == Security.CurrentUserID)
                    continue;

                DataRow NewRow = UsersDT.NewRow();
                NewRow["UserID"] = Row["UserID"];
                NewRow["Name"] = Row["Name"];
                UsersDT.Rows.Add(NewRow);
            }

            RecipientsList.ItemsDataTable = UsersDT;
            RecipientsList.InitializeItems();

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

            InfiniumDocuments.AddRecipients(DocumentCategoryID, DocumentID, RecipientsList.GetSelectedDataTable());


            FormEvent = eClose;
            AnimateTimer.Enabled = true;
        }

        private void AddDocumentRecipientsForm_Load(object sender, EventArgs e)
        {

        }

    }
}
