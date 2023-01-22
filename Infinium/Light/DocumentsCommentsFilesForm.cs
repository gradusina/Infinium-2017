using System;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Windows.Forms;

namespace Infinium
{
    public partial class DocumentsCommentsFilesForm : Form
    {
        private const int eHide = 2;
        private const int eShow = 1;
        private const int eClose = 3;

        private int FormEvent;

        private Form TopForm;

        private InfiniumDocuments InfiniumDocuments;

        public bool bStopTransfer = false;

        public int InnerDocumentID = -1;

        public DataTable CurrentFilesDataTable;



        public DocumentsCommentsFilesForm(ref Form tTopForm, ref InfiniumDocuments tInfiniumDocuments, DataTable tFilesDataTable)
        {
            InitializeComponent();

            InfiniumDocuments = tInfiniumDocuments;

            CurrentFilesDataTable = tFilesDataTable;

            FilesDataGrid.DataSource = CurrentFilesDataTable;
            FilesDataGrid.Columns["IsNew"].Visible = false;
            FilesDataGrid.Columns["FileSize"].Visible = false;
            FilesDataGrid.Columns["FileSizeText"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            FilesDataGrid.Columns["FileSizeText"].Width = 100;


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
            FormEvent = eClose;
            AnimateTimer.Enabled = true;
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            int Size = 0;

            foreach (string FileName in openFileDialog1.FileNames)
            {
                var fileInfo = new FileInfo(FileName);

                DataRow NewRow = CurrentFilesDataTable.NewRow();
                NewRow["FileName"] = Path.GetFileName(FileName);

                Size = (int)(fileInfo.Length / 1024);
                if (Size == 0)
                    Size = 1;

                NewRow["FileSizeText"] = Size + " КБ";
                NewRow["FileSize"] = fileInfo.Length;
                NewRow["FilePath"] = FileName;
                NewRow["IsNew"] = true;
                CurrentFilesDataTable.Rows.Add(NewRow);
            }
        }

        private void AttachButton_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
        }

        private void DetachButton_Click(object sender, EventArgs e)
        {
            CurrentFilesDataTable.Select("FileName = '" + FilesDataGrid.SelectedRows[0].Cells["FileName"].FormattedValue + "'")[0].Delete();
            CurrentFilesDataTable.AcceptChanges();
        }

        private void CreateButton_Click(object sender, EventArgs e)
        {
            Close();
        }

    }
}
