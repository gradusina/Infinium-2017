﻿using System;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace Infinium
{
    public partial class CreateOuterDocumentForm : Form
    {
        private const int eHide = 2;
        private const int eShow = 1;
        private const int eClose = 3;

        private int FormEvent;

        private Form TopForm;

        private InfiniumDocuments InfiniumDocuments;

        private DataTable AttachmentsDataTable;
        private DataTable RecipientsDataTable;

        private bool bEdit;
        public bool bCanceled;

        public bool bStopTransfer;

        public int OuterDocumentID = -1;

        public CreateOuterDocumentForm(ref Form tTopForm, ref InfiniumDocuments tInfiniumDocuments)
        {
            InitializeComponent();

            InfiniumDocuments = tInfiniumDocuments;

            TopForm = tTopForm;

            DocumentsTypeComboBox.DataSource = InfiniumDocuments.DocumentTypesDataTable;
            DocumentsTypeComboBox.DisplayMember = "DocumentType";
            DocumentsTypeComboBox.ValueMember = "DocumentTypeID";

            RecipientComboBox.DataSource = InfiniumDocuments.CorrespondentsDataTable;
            RecipientComboBox.DisplayMember = "CorrespondentName";
            RecipientComboBox.ValueMember = "CorrespondentID";

            DocumentsStatesComboBox.DataSource = InfiniumDocuments.DocumentsStatesDataTable;
            DocumentsStatesComboBox.DisplayMember = "DocumentState";
            DocumentsStatesComboBox.ValueMember = "DocumentsStateID";

            FactoryTypesComboBox.DataSource = InfiniumDocuments.FactoryTypesDataTable;
            FactoryTypesComboBox.DisplayMember = "Factory";
            FactoryTypesComboBox.ValueMember = "FactoryID";

            RecipientsList.ItemsDataTable = InfiniumDocuments.UsersDataTable;
            RecipientsList.InitializeItems();

            AttachmentsDataTable = new DataTable();
            AttachmentsDataTable.Columns.Add(new DataColumn("FileName", Type.GetType("System.String")));
            AttachmentsDataTable.Columns.Add(new DataColumn("Path", Type.GetType("System.String")));
            AttachmentsDataTable.Columns.Add(new DataColumn("IsNew", Type.GetType("System.Boolean")));

            RecipientsDataTable = new DataTable();
            RecipientsDataTable.Columns.Add(new DataColumn("UserID", Type.GetType("System.Int32")));

            AttachmentsGrid.DataSource = AttachmentsDataTable;
            AttachmentsGrid.Columns["IsNew"].Visible = false;
            AttachmentsGrid.Columns["Path"].Visible = false;
        }

        public CreateOuterDocumentForm(ref Form tTopForm, ref InfiniumDocuments tInfiniumDocuments, int iOuterDocumentID)
        {
            InitializeComponent();

            InfiniumDocuments = tInfiniumDocuments;

            TopForm = tTopForm;

            DocumentsTypeComboBox.DataSource = InfiniumDocuments.DocumentTypesDataTable;
            DocumentsTypeComboBox.DisplayMember = "DocumentType";
            DocumentsTypeComboBox.ValueMember = "DocumentTypeID";

            RecipientComboBox.DataSource = InfiniumDocuments.CorrespondentsDataTable;
            RecipientComboBox.DisplayMember = "CorrespondentName";
            RecipientComboBox.ValueMember = "CorrespondentID";

            DocumentsStatesComboBox.DataSource = InfiniumDocuments.DocumentsStatesDataTable;
            DocumentsStatesComboBox.DisplayMember = "DocumentState";
            DocumentsStatesComboBox.ValueMember = "DocumentsStateID";

            FactoryTypesComboBox.DataSource = InfiniumDocuments.FactoryTypesDataTable;
            FactoryTypesComboBox.DisplayMember = "Factory";
            FactoryTypesComboBox.ValueMember = "FactoryID";

            RecipientsList.ItemsDataTable = InfiniumDocuments.UsersDataTable;
            RecipientsList.InitializeItems();

            AttachmentsDataTable = new DataTable();
            AttachmentsDataTable.Columns.Add(new DataColumn("FileName", Type.GetType("System.String")));
            AttachmentsDataTable.Columns.Add(new DataColumn("Path", Type.GetType("System.String")));
            AttachmentsDataTable.Columns.Add(new DataColumn("IsNew", Type.GetType("System.Boolean")));

            RecipientsDataTable = new DataTable();
            RecipientsDataTable.Columns.Add(new DataColumn("UserID", Type.GetType("System.Int32")));

            AttachmentsGrid.DataSource = AttachmentsDataTable;
            AttachmentsGrid.Columns["IsNew"].Visible = false;
            AttachmentsGrid.Columns["Path"].Visible = false;

            int DocumentTypeID = -1;
            DateTime DateTime = DateTime.Now;
            int UserID = -1;
            int CorrespondentID = -1;
            string Description = "";
            int DocumentStateID = -1;
            int FactoryID = -1;
            string RegNumber = "";

            OuterDocumentID = iOuterDocumentID;

            InfiniumDocuments.GetEditOuterDocument(OuterDocumentID, ref DocumentTypeID, ref DateTime, ref UserID, ref CorrespondentID,
                                               RecipientsDataTable, ref Description, ref RegNumber, ref DocumentStateID, AttachmentsDataTable, ref FactoryID);

            DocumentsTypeComboBox.SelectedValue = DocumentTypeID;

            if (CorrespondentID > 0)
            {
                RecipientCheckBox.Checked = true;
                RecipientComboBox.SelectedValue = CorrespondentID;
            }

            DocumentsStatesComboBox.SelectedValue = DocumentStateID;
            FactoryTypesComboBox.SelectedValue = FactoryID;
            DescriptionTextBox.Text = Description;
            AttachmentsGrid.DataSource = AttachmentsDataTable;

            foreach (DataRow Row in RecipientsDataTable.Rows)
            {
                RecipientsList.SelectItem(Convert.ToInt32(Row["UserID"]));
            }

            bEdit = true;

            CreateButton.Text = "Сохранить";
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

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            foreach (string FileName in openFileDialog1.FileNames)
            {
                var fileInfo = new FileInfo(FileName);

                DataRow NewRow = AttachmentsDataTable.NewRow();
                NewRow["FileName"] = Path.GetFileName(FileName);
                NewRow["Path"] = FileName;
                NewRow["IsNew"] = true;
                AttachmentsDataTable.Rows.Add(NewRow);
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

        private void LoadTimer_Tick(object sender, EventArgs e)
        {
            if (InfiniumDocuments.FM.TotalFileSize == 0)
                return;

            if (InfiniumDocuments.FM.Position == InfiniumDocuments.FM.TotalFileSize || InfiniumDocuments.FM.Position > InfiniumDocuments.FM.TotalFileSize)
            {
                ProgressBar.Value = 100;
                return;
            }

            DownloadLabel.Text = FileManager.GetIntegerWithThousands(Convert.ToInt32(InfiniumDocuments.FM.Position / 1024)) + " / " +
                                   FileManager.GetIntegerWithThousands(Convert.ToInt32((InfiniumDocuments.FM.TotalFileSize / 1024))) + " КБайт";

            SpeedLabel.Text = FileManager.GetIntegerWithThousands(Convert.ToInt32(InfiniumDocuments.FM.CurrentSpeed)) + " КБайт/c";

            ProgressBar.Value = Convert.ToInt32(100 * InfiniumDocuments.FM.Position / InfiniumDocuments.FM.TotalFileSize);
            PercentsLabel.Text = ProgressBar.Value + " %";
        }

        private void CancelMessagesButton_Click(object sender, EventArgs e)
        {
            bStopTransfer = true;
            bCanceled = true;

            FormEvent = eClose;
            AnimateTimer.Enabled = true;
        }

        private void RecipientCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (RecipientCheckBox.Checked)
                RecipientComboBox.Enabled = true;
            else
                RecipientComboBox.Enabled = false;
        }


        private void CreateButton_Click(object sender, EventArgs e)
        {
            int iCorrespondentID = -1;

            if (RecipientCheckBox.Checked)
                iCorrespondentID = Convert.ToInt32(RecipientComboBox.SelectedValue);

            if (iCorrespondentID == 0)
            {
                InfiniumTips.ShowTip(this, 50, 85, "Выберите получателя", 2000);
                return;
            }

            if (AttachmentsDataTable.Rows.Count == 0)
            {
                InfiniumTips.ShowTip(this, 50, 85, "Не выбраны файлы", 2000);
                return;
            }

            int iCurrentUploadedFile = 0;

            bool Ok = false;

            InfiniumDocuments.FM.bStopTransfer = false;

            LoadLabel.Visible = true;
            ProgressBar.Visible = true;
            LoadLabel.Text = "Загрузка прикрепленных файлов...";
            LoadTimer.Enabled = true;
            CancelLoadingFilesButton.Visible = true;

            Application.DoEvents();

            int LastUploadedFile = 0;
            int CurrentUploadedFile = 0;
            int TotalFilesCount = AttachmentsDataTable.Rows.Count;

            int DocumentType = Convert.ToInt32(DocumentsTypeComboBox.SelectedValue);
            string Description = DescriptionTextBox.Text;
            int DocumentState = Convert.ToInt32(DocumentsStatesComboBox.SelectedValue);
            int FactoryID = Convert.ToInt32(FactoryTypesComboBox.SelectedValue);

            if (DocumentType == 0)
            {
                InfiniumTips.ShowTip(this, 50, 85, "Выберите тип документа", 2000);
                return;
            }

            RecipientsDataTable.Clear();

            foreach (InfiniumDocumentsSelectUserItem Item in RecipientsList.Items)
            {
                if (Item.Checked)
                {
                    DataRow NewRow = RecipientsDataTable.NewRow();
                    NewRow["UserID"] = Item.UserID;
                    RecipientsDataTable.Rows.Add(NewRow);
                }
            }

            Thread T;

            if (!bEdit)
            {
                T = new Thread(delegate ()
                {
                    Ok = InfiniumDocuments.AddOuterDocument(DocumentType,
                                                        Security.GetCurrentDate(), Security.CurrentUserID, iCorrespondentID, RecipientsDataTable,
                                                        Description, RegNumberTextBox.Text, DocumentState, AttachmentsDataTable,
                                                        ref iCurrentUploadedFile, FactoryID);
                });
                T.Start();
            }
            else
            {
                T = new Thread(delegate ()
                {
                    Ok = InfiniumDocuments.EditOuterDocument(OuterDocumentID, DocumentType, Security.CurrentUserID, iCorrespondentID, RecipientsDataTable,
                                                        Description, RegNumberTextBox.Text, DocumentState, AttachmentsDataTable,
                                                        ref iCurrentUploadedFile, FactoryID);
                });
                T.Start();
            }

            Activate();
            Application.DoEvents();

            while (T.IsAlive)
            {
                T.Join(50);
                Application.DoEvents();

                if (CurrentUploadedFile != LastUploadedFile)
                {
                    LoadLabel.Text = "Загрузка прикрепленных файлов(" + CurrentUploadedFile + " из " + TotalFilesCount + ")";
                    LastUploadedFile = CurrentUploadedFile;
                }

                if (bStopTransfer)
                {
                    InfiniumDocuments.FM.bStopTransfer = true;
                    bStopTransfer = false;
                    LoadTimer.Enabled = false;
                    ProgressBar.Visible = false;
                    LoadLabel.Text = "Отмена загрузки файлов...";
                    DownloadLabel.Text = "";
                    SpeedLabel.Text = "";
                    PercentsLabel.Text = "";
                    CancelLoadingFilesButton.Visible = false;

                    Application.DoEvents();

                    while (T.IsAlive)
                        Thread.Sleep(50);

                    FormEvent = eClose;
                    AnimateTimer.Enabled = true;
                }
            }

            FormEvent = eClose;
            AnimateTimer.Enabled = true;
        }

        private void AddCorrespondentButton_Click(object sender, EventArgs e)
        {
            if (RecipientComboBox.Text.Length == 0)
                return;

            string CorName = RecipientComboBox.Text;

            if (InfiniumDocuments.AddCorrespondent(RecipientComboBox.Text) == false)
            {
                InfiniumTips.ShowTip(this, 50, 85, "Корреспондент уже существует в базе", 3800);
                return;
            }

            InfiniumTips.ShowTip(this, 50, 85, "Корреспондент добавлен в базу данных", 3800);

            InfiniumDocuments.RefillCorrespondents();
            RecipientComboBox.SelectedValue = Convert.ToInt32(InfiniumDocuments.CorrespondentsDataTable.Select("CorrespondentName = '" + CorName + "'")[0]["CorrespondentID"]);
        }

        private void AttachButton_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
        }

        private void CancelLoadingFilesButton_Click(object sender, EventArgs e)
        {
            bStopTransfer = true;
        }

        private void DetachButton_Click(object sender, EventArgs e)
        {
            AttachmentsDataTable.Select("FileName = '" + AttachmentsGrid.SelectedRows[0].Cells["FileName"].FormattedValue + "'")[0].Delete();
            AttachmentsDataTable.AcceptChanges();
        }

        private void CheckAll_CheckedChanged(object sender, EventArgs e)
        {
            foreach (InfiniumDocumentsSelectUserItem Item in RecipientsList.Items)
            {
                Item.Checked = CheckAll.Checked;
            }
        }
    }
}
