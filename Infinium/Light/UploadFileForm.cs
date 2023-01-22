using System;
using System.Data;
using System.Linq;
using System.Threading;
using System.Windows.Forms;

namespace Infinium
{
    public partial class UploadFileForm : Form
    {
        private const int EHide = 2;
        private const int EShow = 1;
        private const int EClose = 3;

        private int _formEvent;
        
        private readonly Form _topForm;

        private InfiniumFiles _infiniumFiles;

        private readonly FileManager _fm;

        //InfiniumDocuments InfiniumDocuments;

        private int _lastUploadedFile;
        private int _currentUploadedFile;
        private readonly int _totalFilesCount;

        private bool _bStopTransfer;

        private readonly string[] _fileNames;
        private readonly string _fileName = "";

        private readonly int _folderId = -1;
        private readonly int _fileId = -1;

        private readonly bool _bUpload;
        private readonly bool _bSave;
        private readonly bool _bDownload;
        private readonly bool _bMultiple;
        private readonly bool _bReplace;

        private readonly string _saveToPath = "";

        public int BOk = 5;

        private readonly DataTable _itemsDataTable;



        public UploadFileForm(ref FileManager tFm, ref InfiniumFiles tInfiniumDocuments, string[] sFileNames, int iFolderId, ref Form tTopForm)//upload
        {
            InitializeComponent();

            _fm = tFm;

            _fileNames = sFileNames;

            _totalFilesCount = sFileNames.Count();

            _folderId = iFolderId;

            _topForm = tTopForm;
            _infiniumFiles = tInfiniumDocuments;

            _bUpload = true;
        }

        public UploadFileForm(ref FileManager tFm, ref InfiniumFiles tInfiniumDocuments, string sFileName, int iFolderId, int iFileId,
                              bool save, string sSaveToPath, ref Form tTopForm)//download
        {
            InitializeComponent();

            _fm = tFm;

            _fileName = sFileName;

            _totalFilesCount = 1;

            _folderId = iFolderId;
            _fileId = iFileId;

            _topForm = tTopForm;
            _infiniumFiles = tInfiniumDocuments;

            _bDownload = true;

            _bSave = save;
            _saveToPath = sSaveToPath;
            _bUpload = false;
        }

        public UploadFileForm(ref FileManager tFm, ref InfiniumFiles tInfiniumDocuments, string sSaveToPath, DataTable dtItemsDataTable, ref Form tTopForm)//download multiple
        {
            InitializeComponent();

            _fm = tFm;

            _totalFilesCount = 1;

            _topForm = tTopForm;

            _infiniumFiles = tInfiniumDocuments;

            _saveToPath = sSaveToPath;

            _itemsDataTable = dtItemsDataTable;

            _bUpload = false;
            _bMultiple = true;
        }

        public UploadFileForm(ref FileManager tFm, ref InfiniumFiles tInfiniumDocuments, int iFileId, string sFileName, ref Form tTopForm)//replace upload
        {
            InitializeComponent();

            _fm = tFm;

            _fileId = iFileId;

            _totalFilesCount = 1;

            _fileName = sFileName;

            _topForm = tTopForm;
            _infiniumFiles = tInfiniumDocuments;

            _bReplace = true;
        }


        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);

            if (m.Msg == 6 && m.WParam.ToInt32() == 1)
            {
                if (_topForm != null)
                    _topForm.Activate();
            }
        }

        private void AnimateTimer_Tick(object sender, EventArgs e)
        {
            if (!DatabaseConfigsManager.Animation)
            {
                Opacity = 1;

                if (_formEvent == EClose || _formEvent == EHide)
                {
                    AnimateTimer.Enabled = false;

                    if (_formEvent == EClose)
                    {
                        Close();
                    }

                    if (_formEvent == EHide)
                    {
                        Hide();
                    }

                    return;
                }

                if (_formEvent == EShow)
                {
                    AnimateTimer.Enabled = false;
                    SplashForm.CloseS = true;

                    if (_bUpload)
                        BOk = StartLoading();

                    if (_bMultiple)
                        StartDownloadMulti();

                    if (_bDownload)
                        BOk = StartDownload();

                    if (_bReplace)
                        StartReplace();

                    return;
                }

            }

            if (_formEvent == EClose || _formEvent == EHide)
            {
                if (Convert.ToDecimal(Opacity) != Convert.ToDecimal(0.00))
                    Opacity = Convert.ToDouble(Convert.ToDecimal(Opacity) - Convert.ToDecimal(0.05));
                else
                {
                    AnimateTimer.Enabled = false;

                    if (_formEvent == EClose)
                    {
                        Close();
                    }

                    if (_formEvent == EHide)
                    {
                        Hide();
                    }
                }

                return;
            }


            if (_formEvent == EShow)
            {
                if (Opacity != 1)
                    Opacity += 0.05;
                else
                {
                    AnimateTimer.Enabled = false;
                    SplashForm.CloseS = true;

                    if (_bUpload)
                        BOk = StartLoading();

                    if (_bMultiple)
                        StartDownloadMulti();

                    if (_bDownload)
                        BOk = StartDownload();

                    if (_bReplace)
                        StartReplace();
                }
            }
        }

        private void CreateFolderForm_Shown(object sender, EventArgs e)
        {
            _formEvent = EShow;
            AnimateTimer.Enabled = true;
        }

        private void CancelFolderButton_Click(object sender, EventArgs e)
        {
            _bStopTransfer = true;

            _formEvent = EClose;
            AnimateTimer.Enabled = true;
        }

        private void LoadTimer_Tick(object sender, EventArgs e)
        {
            if (_fm.TotalFileSize == 0)
                return;

            if (_fm.Position == _fm.TotalFileSize ||
                _fm.Position > _fm.TotalFileSize)
            {
                ProgressBar.Value = 100;
                return;
            }


            DownloadLabel.Text = FileManager.GetIntegerWithThousands(Convert.ToInt32(_fm.Position / 1024)) + " / " +
                                   FileManager.GetIntegerWithThousands(Convert.ToInt32((_fm.TotalFileSize / 1024))) + " КБайт";

            ProgressBar.Value = Convert.ToInt32(100 * _fm.Position / _fm.TotalFileSize);
            PercentsLabel.Text = ProgressBar.Value + " %";
        }

        private int StartLoading()
        {
            _bStopTransfer = false;
            _fm.bStopTransfer = false;

            LoadLabel.Visible = true;
            ProgressBar.Visible = true;
            LoadTimer.Enabled = true;

            bool bOk = false;

            Thread T = new Thread(delegate () { bOk = _infiniumFiles.UploadFile(_fileNames, _folderId, ref _currentUploadedFile); });

            T.Start();

            Activate();
            Application.DoEvents();

            while (T.IsAlive)
            {
                T.Join(50);
                Application.DoEvents();

                if (_currentUploadedFile != _lastUploadedFile)
                {
                    LoadLabel.Text = "Загрузка файлов (" + _currentUploadedFile + " из " + _totalFilesCount + ")";
                    _lastUploadedFile = _currentUploadedFile;
                }

                if (_bStopTransfer)
                {
                    _fm.bStopTransfer = true;
                    _bStopTransfer = false;
                    LoadTimer.Enabled = false;
                    ProgressBar.Visible = false;
                    LoadLabel.Text = "Отмена загрузки файлов...";
                    DownloadLabel.Text = "";
                    SpeedLabel.Text = "";
                    PercentsLabel.Text = "";
                    CancelFilesButton.Visible = false;

                    Application.DoEvents();

                    while (T.IsAlive)
                        Thread.Sleep(50);

                    _formEvent = EClose;
                    AnimateTimer.Enabled = true;

                    return -1;
                }
            }

            LoadTimer.Enabled = false;

            _formEvent = EClose;
            AnimateTimer.Enabled = true;

            return Convert.ToInt32(bOk);
        }

        private int StartDownload()
        {
            _bStopTransfer = false;
            _fm.bStopTransfer = false;

            LoadLabel.Visible = true;
            ProgressBar.Visible = true;
            LoadTimer.Enabled = true;

            Thread T;

            bool bOk = false;

            if (_bSave)
                T = new Thread(delegate () { bOk = _infiniumFiles.SaveFile(_fileId, _saveToPath); });
            else
                T = new Thread(delegate () { bOk = _infiniumFiles.OpenFile(_fileId); });

            T.Start();

            Activate();
            Application.DoEvents();

            while (T.IsAlive)
            {
                T.Join(50);
                Application.DoEvents();

                if (_currentUploadedFile != _lastUploadedFile)
                {
                    LoadLabel.Text = "Загрузка файлов (" + _currentUploadedFile + " из " + _totalFilesCount + ")";
                    _lastUploadedFile = _currentUploadedFile;
                }

                if (_bStopTransfer)
                {
                    _fm.bStopTransfer = true;
                    _bStopTransfer = false;
                    LoadTimer.Enabled = false;
                    ProgressBar.Visible = false;
                    LoadLabel.Text = "Отмена загрузки файлов...";
                    DownloadLabel.Text = "";
                    SpeedLabel.Text = "";
                    PercentsLabel.Text = "";
                    CancelFilesButton.Visible = false;

                    Application.DoEvents();

                    while (T.IsAlive)
                        Thread.Sleep(50);

                    _formEvent = EClose;
                    AnimateTimer.Enabled = true;

                    return -1;
                }
            }

            LoadTimer.Enabled = false;

            _formEvent = EClose;
            AnimateTimer.Enabled = true;

            return Convert.ToInt32(bOk);
        }

        private void StartDownloadMulti()
        {
            _bStopTransfer = false;
            _fm.bStopTransfer = false;

            LoadLabel.Visible = true;
            ProgressBar.Visible = true;
            LoadTimer.Enabled = true;

            Thread T;

            T = new Thread(delegate () { _infiniumFiles.SaveFiles(_itemsDataTable, _saveToPath, ref _currentUploadedFile); });

            T.Start();

            Activate();
            Application.DoEvents();

            while (T.IsAlive)
            {
                T.Join(50);
                Application.DoEvents();

                if (_currentUploadedFile != _lastUploadedFile)
                {
                    LoadLabel.Text = "Загрузка файлов (" + _currentUploadedFile + " из " + _totalFilesCount + ")";
                    _lastUploadedFile = _currentUploadedFile;
                }

                if (_bStopTransfer)
                {
                    _fm.bStopTransfer = true;
                    _bStopTransfer = false;
                    LoadTimer.Enabled = false;
                    ProgressBar.Visible = false;
                    LoadLabel.Text = "Отмена загрузки файлов...";
                    DownloadLabel.Text = "";
                    SpeedLabel.Text = "";
                    PercentsLabel.Text = "";
                    CancelFilesButton.Visible = false;

                    Application.DoEvents();

                    while (T.IsAlive)
                        Thread.Sleep(50);

                    _formEvent = EClose;
                    AnimateTimer.Enabled = true;

                    return;
                }
            }

            LoadTimer.Enabled = false;

            _formEvent = EClose;
            AnimateTimer.Enabled = true;
        }

        private void StartReplace()
        {
            _bStopTransfer = false;
            _fm.bStopTransfer = false;

            LoadLabel.Visible = true;
            ProgressBar.Visible = true;
            LoadTimer.Enabled = true;

            bool ok = false;

            Thread T = new Thread(delegate () { ok = _infiniumFiles.ReplaceFile(_fileId, _fileName); });

            T.Start();

            Activate();
            Application.DoEvents();

            while (T.IsAlive)
            {
                T.Join(50);
                Application.DoEvents();

                if (_currentUploadedFile != _lastUploadedFile)
                {
                    LoadLabel.Text = "Загрузка файлов (" + _currentUploadedFile + " из " + _totalFilesCount + ")";
                    _lastUploadedFile = _currentUploadedFile;
                }

                if (_bStopTransfer)
                {
                    _fm.bStopTransfer = true;
                    _bStopTransfer = false;
                    LoadTimer.Enabled = false;
                    ProgressBar.Visible = false;
                    LoadLabel.Text = "Отмена загрузки файлов...";
                    DownloadLabel.Text = "";
                    SpeedLabel.Text = "";
                    PercentsLabel.Text = "";
                    CancelFilesButton.Visible = false;

                    Application.DoEvents();

                    while (T.IsAlive)
                        Thread.Sleep(50);

                    _formEvent = EClose;
                    AnimateTimer.Enabled = true;

                    return;
                }
            }

            LoadTimer.Enabled = false;

            if (!ok)
            {
                MessageBox.Show("Один или несколько файлов не были найдены или отсутствовал доступ");
                return;
            }

            _formEvent = EClose;
            AnimateTimer.Enabled = true;
        }

    }
}
