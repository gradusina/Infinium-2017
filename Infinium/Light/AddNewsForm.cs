﻿using System;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace Infinium
{
    public partial class AddNewsForm : Form
    {
        private LightNews LightNews;

        private const int eHide = 2;
        private const int eShow = 1;
        private const int eClose = 3;

        private bool bStopTransfer;

        private int FormEvent;

        public bool Canceled;

        public int iNewsIDEdit = -1;

        public int AttachsCount;

        private Form TopForm;

        public DateTime DateTime;

        public DataTable AttachmentsDataTable;
        public BindingSource AttachmentsBindingSource;

        private void CreateAttachments()
        {
            AttachmentsDataTable = new DataTable();
            AttachmentsDataTable.Columns.Add(new DataColumn("FileName", Type.GetType("System.String")));
            AttachmentsDataTable.Columns.Add(new DataColumn("Path", Type.GetType("System.String")));

            AttachmentsBindingSource = new BindingSource
            {
                DataSource = AttachmentsDataTable
            };
            AttachmentsGrid.DataSource = AttachmentsBindingSource;

            AttachmentsGrid.Columns["FileName"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            AttachmentsGrid.Columns["FileName"].Width = 250;
        }

        private void CopyAttachs(int NewsID)
        {
            using (DataTable DT = LightNews.GetAttachments(NewsID))
            {
                if (DT.Rows.Count == 0)
                    return;

                AttachsCount = DT.Rows.Count;

                foreach (DataRow Row in DT.Rows)
                {
                    DataRow NewRow = AttachmentsDataTable.NewRow();
                    NewRow["FileName"] = Row["FileName"];
                    NewRow["Path"] = "server";
                    AttachmentsDataTable.Rows.Add(NewRow);
                }
            }
        }

        public AddNewsForm(ref LightNews tLightNews, ref Form tTopForm)
        {
            InitializeComponent();

            TopForm = tTopForm;
            LightNews = tLightNews;

            CreateAttachments();
        }

        public AddNewsForm(ref LightNews tLightNews, int SenderTypeID, string HeaderText, string BodyText, int iNewsID, DateTime dDateTime, ref Form tTopForm)
        {
            InitializeComponent();

            LightNews = tLightNews;

            TopForm = tTopForm;

            HeaderTextEdit.Text = HeaderText;
            BodyTextEdit.Text = BodyText;

            iNewsIDEdit = iNewsID;
            DateTime = dDateTime;

            CreateAttachments();
            CopyAttachs(iNewsID);
        }

        private void OKDateButton_Click(object sender, EventArgs e)
        {
            //if (HeaderTextEdit.Text.Length < 1 || HeaderTextEdit.Text.Length > 1024)
            //    return;

            if (BodyTextEdit.Text.Length < 1)
                return;

            int DepartmentID = Convert.ToInt32(LightNews.UsersDataTable.Select("UserID = " + Security.CurrentUserID)[0]["DepartmentID"]);

            if (iNewsIDEdit == -1)
            {
                DateTime Date;

                LoadLabel.Text = "Создание новости...";
                OKNewsButton.Enabled = false;
                CancelNewsButton.Enabled = false;
                AttachmentsGrid.ReadOnly = true;
                BodyTextEdit.Enabled = false;
                AttachButton.Enabled = false;
                DetachButton.Enabled = false;
                HeaderTextEdit.Enabled = false;

                Application.DoEvents();

                Date = LightNews.AddNews(Security.CurrentUserID, 0, HeaderTextEdit.Text, BodyTextEdit.Text, -1);

                bool Ok = false;

                bStopTransfer = false;
                LightNews.FM.bStopTransfer = false;

                if (AttachmentsBindingSource.Count > 0)
                {
                    LoadLabel.Visible = true;
                    ProgressBar.Visible = true;
                    //LoadLabel.Text = "Загрузка прикрепленых файлов...";
                    LoadTimer.Enabled = true;
                    CancelLoadingFilesButton.Visible = true;

                    Application.DoEvents();

                    int LastUploadedFile = 0;
                    int CurrentUploadedFile = 0;
                    int TotalFilesCount = AttachmentsDataTable.Rows.Count;

                    Thread T = new Thread(delegate () { Ok = LightNews.Attach(AttachmentsDataTable, Date, ref CurrentUploadedFile); });
                    T.Start();

                    Activate();
                    Application.DoEvents();

                    while (T.IsAlive)
                    {
                        T.Join(50);
                        Application.DoEvents();

                        if (CurrentUploadedFile != LastUploadedFile)
                        {
                            LoadLabel.Text = "Загрузка прикрепленых файлов(" + CurrentUploadedFile + " из " + TotalFilesCount + ")";
                            LastUploadedFile = CurrentUploadedFile;
                        }

                        if (bStopTransfer)
                        {
                            LightNews.FM.bStopTransfer = true;
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

                            LightNews.RemoveAttachments(LightNews.GetNewsIDByDateTime(Date));

                            FormEvent = eClose;
                            AnimateTimer.Enabled = true;
                        }
                    }

                    LoadTimer.Enabled = false;

                    if (!Ok)
                    {
                        MessageBox.Show("Один или несколько файлов не были найдены или отсутствовал доступ");
                        return;
                    }
                }

                LightNews.ClearPending(Date);
                LightNews.AddSubscribeForNews(Date);
            }
            else
            {
                LoadLabel.Text = "Обновление новости...";
                OKNewsButton.Enabled = false;
                CancelNewsButton.Enabled = false;
                AttachmentsGrid.ReadOnly = true;
                BodyTextEdit.Enabled = false;
                AttachButton.Enabled = false;
                DetachButton.Enabled = false;
                HeaderTextEdit.Enabled = false;

                Application.DoEvents();

                LightNews.EditNews(iNewsIDEdit, Security.CurrentUserID, 0, HeaderTextEdit.Text, BodyTextEdit.Text, -1, DateTime);


                if (AttachsCount > 0 || AttachmentsBindingSource.Count > 0)
                {
                    bool Ok = false;


                    LoadLabel.Visible = true;
                    ProgressBar.Visible = true;
                    //LoadLabel.Text = "Загрузка прикрепленых файлов...";
                    LoadTimer.Enabled = true;
                    CancelLoadingFilesButton.Visible = true;

                    Application.DoEvents();

                    int LastUploadedFile = 0;
                    int CurrentUploadedFile = 0;
                    int TotalFilesCount = 0;

                    Thread T = new Thread(delegate () { Ok = LightNews.EditAttachments(iNewsIDEdit, AttachmentsDataTable, ref CurrentUploadedFile, ref TotalFilesCount); });
                    T.Start();

                    Activate();
                    Application.DoEvents();


                    while (T.IsAlive)
                    {
                        T.Join(50);
                        Application.DoEvents();

                        if (CurrentUploadedFile != LastUploadedFile)
                        {
                            LoadLabel.Text = "Загрузка прикрепленых файлов(" + CurrentUploadedFile + " из " + TotalFilesCount + ")";
                            LastUploadedFile = CurrentUploadedFile;
                        }

                        if (bStopTransfer)
                        {
                            LightNews.FM.bStopTransfer = true;
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

                            LightNews.RemoveCurrentAttachments(LightNews.GetNewsIDByDateTime(DateTime), AttachmentsDataTable);

                            FormEvent = eClose;
                            AnimateTimer.Enabled = true;
                        }
                    }

                    LoadTimer.Enabled = false;



                    if (!Ok)
                    {
                        MessageBox.Show("Один или несколько файлов не были найдены или отсутствовал доступ");
                    }
                }
            }

            FormEvent = eClose;
            AnimateTimer.Enabled = true;
        }

        private void CancelDateButton_Click(object sender, EventArgs e)
        {
            bStopTransfer = true;

            Canceled = true;

            FormEvent = eClose;
            AnimateTimer.Enabled = true;
        }

        private void AttachButton_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
        }

        private void DetachButton_Click(object sender, EventArgs e)
        {
            if (AttachmentsBindingSource.Count == 0)
                return;

            AttachmentsBindingSource.RemoveCurrent();
            AttachmentsDataTable.AcceptChanges();
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
            if (LightNews.FM.TotalFileSize == 0)
                return;

            if (LightNews.FM.Position == LightNews.FM.TotalFileSize || LightNews.FM.Position > LightNews.FM.TotalFileSize)
            {
                ProgressBar.Value = 100;
                return;
            }



            DownloadLabel.Text = FileManager.GetIntegerWithThousands(Convert.ToInt32(LightNews.FM.Position / 1024)) + " / " +
                                   FileManager.GetIntegerWithThousands(Convert.ToInt32((LightNews.FM.TotalFileSize / 1024))) + " КБайт";

            SpeedLabel.Text = FileManager.GetIntegerWithThousands(Convert.ToInt32(LightNews.FM.CurrentSpeed)) + " КБайт/c";

            ProgressBar.Value = Convert.ToInt32(100 * LightNews.FM.Position / LightNews.FM.TotalFileSize);
            PercentsLabel.Text = ProgressBar.Value + " %";
        }

        private void CancelLoadingFilesButton_Click(object sender, EventArgs e)
        {
            bStopTransfer = true;
        }
    }
}
