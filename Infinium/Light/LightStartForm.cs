using Infinium.Properties;

using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Media;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace Infinium
{
    public partial class LightStartForm : Form
    {
        //NotifyForm NotifyForm = null;

        public Form TopForm;
        private LoginForm _loginForm;
        private const int EHide = 2;
        private const int EShow = 1;
        private const int EClose = 3;
        private const int EMainMenu = 4;

        public static bool NotifyShowed = false;
        private int _formEvent;
        private bool _bC;
        private bool _bNeedSplash;
        private ActiveNotifySystem _activeNotifySystem;
        private InfiniumStart _infiniumStart;
        private bool _logout;
        private bool _firstLoad = true;
        private CultureInfo _ci = new CultureInfo("ru-RU");

        [DllImport("user32.dll")]
        private static extern IntPtr GetActiveWindow();

        private Thread _notifyThread;
        private int _iRefreshTime = 4000;
        private NotifyForm _notifyForm;

        public LightStartForm(LoginForm tLoginForm)
        {
            InitializeComponent();

            _loginForm = tLoginForm;

            MaximumSize = Screen.PrimaryScreen.WorkingArea.Size;

            //CurrentTimeLabel.Text = DateTime.Now.ToString("HH:mm");
            //CurrentDayOfWeekLabel.Text = DateTime.Now.ToString("dddd");
            //CurrentDayMonthLabel.Text = DateTime.Now.ToString("dd MMMM");

            AnimateTimer.Enabled = true;

            UserLabel.Text = Security.CurrentUserRName;

            //UpdatesPanel.Width = UserPanel.Left - UpdatesPanel.Left - 5;

            PhotoBox.Image = UserProfile.GetUserPhoto();

            _infiniumStart = new InfiniumStart();
            InfiniumTilesContainer.ItemsDataTable = _infiniumStart.ModulesDataTable;
            InfiniumTilesContainer.MenuItemID = 0;
            InfiniumTilesContainer.InitializeItems();

            _activeNotifySystem = new ActiveNotifySystem();
            InfiniumNotifyList.ModulesDataTable = _infiniumStart.FullModulesDataTable;

            InfiniumStartMenu.ItemsDataTable = _infiniumStart.MenuItemsDataTable;
            InfiniumStartMenu.InitializeItems();
            InfiniumStartMenu.Selected = 0;

            InfiniumMinimizeList.ModulesDataTable = _infiniumStart.FullModulesDataTable;

            if (_infiniumStart.ModulesDataTable.Select("MenuItemID = 0").Count() == 0)
                InfiniumStartMenu.Selected = 1;

            TopForm = null;
            if (TopForm != null)
                OnLineControl.IamOnline(_activeNotifySystem.GetModuleIDByForm(TopForm.Name), true);
            else
                OnLineControl.IamOnline(0, true);

            OnLineControl.SetOffline();

            _notifyRefreshT = FuckingNotify;
            _onlineFuck = GetTopMostAndModuleName;

            _notifyThread = new Thread(delegate () { NotifyCheck(); });
            _notifyThread.Start();

            while (!SplashForm.bCreated) ;
        }

        private OnlineStruct _os;

        private delegate void NotifyRefresh();

        //delegate void ANSNotifyContainerRefresh();
        private delegate OnlineStruct OnlineFuckingDelegate();

        private OnlineFuckingDelegate _onlineFuck;
        private NotifyRefresh _notifyRefreshT;

        private void FuckingNotify()
        {
            //try
            //{
            InfiniumNotifyList.ItemsDataTable = _activeNotifySystem.ModulesUpdatesDataTable;
            InfiniumNotifyList.InitializeItems();

            if (TopForm != null)
            {
                if (TopForm.Name == "MessagesForm")
                {
                    using (Stream str = Resources.MESSAGENOTIFY)
                    {
                        using (var snd = new SoundPlayer(str))
                        {
                            snd.Play();
                        }
                    }
                }
                else
                {
                    using (Stream str = Resources._01_01)
                    {
                        using (var snd = new SoundPlayer(str))
                        {
                            snd.Play();
                        }
                    }
                }
            }
            else
                using (Stream str = Resources._01_01)
                {
                    using (var snd = new SoundPlayer(str))
                    {
                        snd.Play();
                    }
                }



            //если открытый модуль типа InfiniumForm
            if (TopForm != null)
            {
                if (TopForm.GetType().BaseType.Name == "InfiniumForm")
                    ((InfiniumForm)TopForm).OnANSUpdate();
            }

            if (TopForm != null)
                if (_activeNotifySystem.ModulesDataTable.Select("ModuleID = " + _activeNotifySystem.iCurrentModuleID)[0]["FormName"].ToString() == TopForm.Name)
                    if (GetActiveWindow() == TopForm.Handle)
                        return;

            if (GetActiveWindow() == Handle)//только звук
            {
                return;
            }

            //показываем всплывающее окно
            if (NotifyForm.bShowed)
            {
                if (_notifyForm != null)
                    _notifyForm.Close();

                var moduleId = 0;
                var moreCount = 0;
                var count = _activeNotifySystem.GetLastModuleUpdate(ref moduleId, ref moreCount);
                _notifyForm = new NotifyForm(this, ref _activeNotifySystem, moduleId, count, moreCount);
                _notifyForm.Show();
            }
            else
            {
                var moduleId = 0;
                var moreCount = 0;
                var count = _activeNotifySystem.GetLastModuleUpdate(ref moduleId, ref moreCount);
                _notifyForm = new NotifyForm(this, ref _activeNotifySystem, moduleId, count, moreCount);
                _notifyForm.Show();
            }
            //}
            //catch (Exception e)
            //{
            //    MessageBox.Show(e.Message);
            //}
        }

        public void StartModuleFromNotify(int moduleId)
        {
            if (TopForm != null)
                if (TopForm.Name == _infiniumStart.FullModulesDataTable.Select("ModuleID = " + moduleId)[0]["FormName"].ToString())
                {
                    return;
                }

            if (moduleId != 80)
            {
                var T = new Thread(delegate () { SplashWindow.CreateSplash(); });
                T.Start();

                while (!SplashForm.bCreated) ;
            }

            if (TopForm != null)
                if (TopForm.Name != _infiniumStart.FullModulesDataTable.Select("ModuleID = " + moduleId)[0]["FormName"].ToString())
                {
                    HideForm(TopForm);
                }

            Form moduleForm = null;

            if (moduleId == 80)//messages
            {
                MessagesButton_Click(null, null);
            }
            else
            {
                //check if running
                var form = InfiniumMinimizeList.GetForm(_infiniumStart.FullModulesDataTable.Select("ModuleID = " + moduleId)[0]["FormName"].ToString());
                if (form != null)
                {
                    TopForm = form;
                    form.Show();
                }
                else
                {
                    var caType = Type.GetType("Infinium." + _infiniumStart.FullModulesDataTable.Select("ModuleID = " + moduleId)[0]["FormName"]);
                    moduleForm = (Form)Activator.CreateInstance(caType, this);
                    TopForm = moduleForm;
                    moduleForm.ShowDialog();
                }
            }

            _activeNotifySystem.FillUpdates();
            InfiniumNotifyList.ItemsDataTable = _activeNotifySystem.ModulesUpdatesDataTable;
            InfiniumNotifyList.InitializeItems();
        }

        private struct OnlineStruct
        {
            public int ModuleId;
            public bool TopMost;
        }

        private OnlineStruct GetTopMostAndModuleName()
        {
            if (TopForm == null)
            {
                if (GetActiveWindow() == Handle)
                {
                    _os.TopMost = true;
                    _os.ModuleId = 0;
                }
                else
                {
                    _os.TopMost = false;
                    _os.ModuleId = 0;
                }
            }
            else
            {
                if (TopForm == this && GetActiveWindow() != Handle)
                {
                    _os.TopMost = false;
                    _os.ModuleId = 0;
                }
                else
                {
                    if (GetActiveWindow() == Handle)
                    {
                        _os.TopMost = true;
                        _os.ModuleId = 0;
                    }
                    else
                        if ((int)GetActiveWindow() == 0)
                    {
                        _os.TopMost = false;
                        _os.ModuleId = _activeNotifySystem.GetModuleIDByForm(TopForm.Name);
                    }
                    else
                    {
                        _os.TopMost = true;
                        _os.ModuleId = _activeNotifySystem.GetModuleIDByForm(TopForm.Name);
                    }
                }
            }

            return _os;
        }

        private void NotifyCheck()
        {
            while (true)
            {
                if (IsHandleCreated)
                {
                    //OnLineControl.SetOffline();
                    //OnLineControl.SetOfflineClient();
                    //OnLineControl.SetOfflineManager();

                    if (OnLineControl.GetForciblyOffline())
                    {
                        OnLineControl.SetOffline();
                        OnLineControl.SetForciblyOffline(false);
                        OnLineControl.KillProcess();
                    }

                    Invoke(_onlineFuck);
                    OnLineControl.IamOnline(_os.ModuleId, _os.TopMost);

                    if (ActiveNotifySystem.IsNewUpdates(Security.CurrentUserID) > 0)
                    {
                        if (_activeNotifySystem.CheckLastUpdate())
                        {
                            _activeNotifySystem.FillUpdates();

                            Invoke(_notifyRefreshT);
                        }
                    }
                }

                Thread.Sleep(_iRefreshTime);
            }

        }

        public void HideForm(Form form)
        {
            Activate();
            TopMost = true;
            form.Hide();

            _loginForm.Activate();
            TopMost = false;
            Activate();

            TopForm = null;

            InfiniumMinimizeList.AddModule(ref form);
        }


        public void CloseForm(Form form)
        {
            Activate();
            TopMost = true;
            form.Hide();

            _loginForm.Activate();
            TopMost = false;
            Activate();

            TopForm = null;

            var formName = form.Name;

            InfiniumMinimizeList.RemoveModule(form.Name);

            var T = new Thread(delegate ()
            {
                Security.ExitFromModule(formName);
            });
            T.Start();

            form.Dispose();

            GC.Collect();
        }


        private void AnimateTimer_Tick(object sender, EventArgs e)
        {
            //without animation
            if (!DatabaseConfigsManager.Animation)
            {
                Opacity = 1;
                AnimateTimer.Enabled = false;

                if (_formEvent == EClose)
                {
                    _loginForm.Show();
                    _loginForm.ShowInTaskbar = true;
                    //LoginForm.CloseJournalRec();
                    Close();
                    return;
                }

                SplashForm.CloseS = true;

                if (_firstLoad)
                {
                    _loginForm.ShowInTaskbar = false;
                    _loginForm.Visible = false;
                    AnimateTimer.Enabled = false;
                    _firstLoad = false;
                }
                else
                {
                    Activate();
                    TopMost = true;
                    AnimateTimer.Enabled = false;
                    TopMost = false;
                }

                _bNeedSplash = true;

                //InfiniumTilesContainer.InitializeItems();

                return;
            }


            //with animation
            if (_formEvent == EClose)
            {
                if (Convert.ToDecimal(Opacity) != Convert.ToDecimal(0.00))
                {
                    Opacity = Convert.ToDouble(Convert.ToDecimal(Opacity) - Convert.ToDecimal(0.05));
                }
                else
                {
                    AnimateTimer.Enabled = false;
                    _loginForm.Show();
                    _loginForm.ShowInTaskbar = true;

                    Close();
                }

                return;
            }

            if (Opacity != 1)
                Opacity += 0.05;
            else
            {
                AnimateTimer.Enabled = false;

                SplashForm.CloseS = true;

                if (_firstLoad)
                {
                    _loginForm.ShowInTaskbar = false;
                    _loginForm.Visible = false;
                    AnimateTimer.Enabled = false;
                    _firstLoad = false;
                }
                else
                {
                    Activate();
                    TopMost = true;
                    AnimateTimer.Enabled = false;
                    TopMost = false;
                }

                _bNeedSplash = true;
            }
        }


        private void MenuCloseButton_Click_1(object sender, EventArgs e)
        {
            _notifyThread.Abort();

            _loginForm.CloseJournalRec();
            _loginForm.Close();
        }

        private void LightStartForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            _loginForm.CloseJournalRec();

            if (!_logout)
                _loginForm.Close();
        }

        private void LogoutButton_Click(object sender, EventArgs e)
        {
            _notifyThread.Abort();
            _logout = true;

            _formEvent = EClose;

            AnimateTimer.Enabled = true;
        }

        private void MinimizeButton_Click(object sender, EventArgs e)
        {
            WindowState = FormWindowState.Minimized;
        }


        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);

            if (m.Msg == 6 && m.WParam.ToInt32() == 1)
            {
                if (FromHandle(m.LParam) == null)
                {
                    if (TopForm != null)
                        TopForm.Activate();
                }
            }
        }

        private void notifyIcon1_Click(object sender, EventArgs e)
        {
            if (WindowState == FormWindowState.Minimized)
                WindowState = FormWindowState.Maximized;

            Activate();
        }

        private void LightStartForm_Activated(object sender, EventArgs e)
        {
            if (TopForm != null)
                TopForm.Activate();
        }

        private void LightStartForm_Load(object sender, EventArgs e)
        {
            InfiniumTilesContainer.InitializeItems();
        }

        private void CurrentTimer_Tick(object sender, EventArgs e)
        {
            InfiniumClock.iMinutes = (DateTime.Now).Minute;
            InfiniumClock.iSeconds = (DateTime.Now).Second;
            InfiniumClock.iHours = (DateTime.Now).Hour;
            InfiniumClock.Refresh();

            TimeLabel.Text = (DateTime.Now).ToString("HH:mm:ss");
            DateLabel.Text = (DateTime.Now).ToString("dd MMMM yyyy");
            DayLabel.Text = (DateTime.Now).ToString("dddd");
        }

        private void button1_Click_1(object sender, EventArgs e)
        {

        }


        private void UpdatesTimer_Tick(object sender, EventArgs e)
        {
            //NotifyCheck();
        }

        private void InfiniumStartMenu_ItemClicked(object sender, string menuItemName, int menuItemId)
        {
            if (_bNeedSplash)
            {
                var T = new Thread(delegate ()
                {
                    SplashWindow.CreateCoverSplash(InfiniumTilesContainer.Top + MainPanel.Top, InfiniumTilesContainer.Left,
                                                   InfiniumTilesContainer.Height, InfiniumTilesContainer.Width);
                });
                T.Start();

                while (!SplashWindow.bSmallCreated) ;
            }

            InfiniumTilesContainer.MenuItemID = menuItemId;
            InfiniumTilesContainer.InitializeItems();

            _bC = true;
        }

        private void SplashTimer_Tick(object sender, EventArgs e)
        {
            if (_bC)
            {
                while (SplashWindow.bSmallCreated)
                    CoverWaitForm.CloseS = true;

                _bC = false;
            }
        }

        private void button1_MouseMove(object sender, MouseEventArgs e)
        {

        }

        private void InfiniumTilesContainer_RightMouseClicked(object sender, string formName)
        {
            _curFormName = formName;

            if (InfiniumStartMenu.Selected == 0)
            {
                MenuRemoveFromFavorite.Visible = true;
                MenuAddToFavorite.Visible = false;
            }
            else
            {
                MenuRemoveFromFavorite.Visible = false;
                MenuAddToFavorite.Visible = true;
            }

            TileContextMenu.Show(InfiniumTilesContainer);
        }

        private string _curFormName = "";

        private void MenuAddToFavorite_Click(object sender, EventArgs e)
        {
            if (_infiniumStart.AddModuleToFavorite(_curFormName) == false)
            {
                InfiniumTips.ShowTip(this, 50, 85, "Модуль уже был добавлен ранее", 4000);
                return;
            }

            InfiniumTips.ShowTip(this, 50, 85, "Модуль добавлен", 2500);
        }

        private void MenuRemoveFromFavorite_Click(object sender, EventArgs e)
        {
            _infiniumStart.RemoveModuleFromFavorite(_curFormName);

            if (_bNeedSplash)
            {
                var T = new Thread(delegate ()
                {
                    SplashWindow.CreateCoverSplash(InfiniumTilesContainer.Top + MainPanel.Top, InfiniumTilesContainer.Left,
                                                   InfiniumTilesContainer.Height, InfiniumTilesContainer.Width);
                });
                T.Start();

                while (!SplashWindow.bSmallCreated) ;
            }

            InfiniumTilesContainer.MenuItemID = 0;
            InfiniumTilesContainer.InitializeItems();

            _bC = true;
        }

        private void InfiniumTilesContainer_ItemClicked(object sender, string formName)
        {
            var T = new Thread(delegate () { SplashWindow.CreateSplash(); });
            T.Start();

            while (!SplashForm.bCreated) ;

            Form moduleForm = null;

            //check if running
            var form = InfiniumMinimizeList.GetForm(formName);
            if (form != null)
            {
                TopForm = form;
                form.ShowDialog();
            }
            else
            {
                var caType = Type.GetType("Infinium." + formName);
                moduleForm = (Form)Activator.CreateInstance(caType, this);
                TopForm = moduleForm;
                Security.EnterInModule(formName);
                moduleForm.ShowDialog();
            }

            if (InfiniumNotifyList.Items != null)
                if (InfiniumNotifyList.Items.Count() > 0)
                {
                    _activeNotifySystem.FillUpdates();
                    InfiniumNotifyList.ItemsDataTable = _activeNotifySystem.ModulesUpdatesDataTable;
                    InfiniumNotifyList.InitializeItems();
                }
        }

        private void InfiniumNotifyList_ItemClicked(object sender, string formName)
        {
            if (formName == "MessagesForm")
            {
                MessagesButton_Click(null, null);

                if (_notifyForm != null)
                {
                    _notifyForm.Close();
                    _notifyForm.Dispose();
                    _notifyForm = null;
                }

                _activeNotifySystem.FillUpdates();
                InfiniumNotifyList.ItemsDataTable = _activeNotifySystem.ModulesUpdatesDataTable;
                InfiniumNotifyList.InitializeItems();

                return;
            }

            var T = new Thread(delegate () { SplashWindow.CreateSplash(); });
            T.Start();

            while (!SplashForm.bCreated) ;


            if (_notifyForm != null)
            {
                _notifyForm.Close();
                _notifyForm.Dispose();
                _notifyForm = null;
            }

            Form moduleForm = null;

            //check if running
            var form = InfiniumMinimizeList.GetForm(formName);
            if (form != null)
            {
                TopForm = form;
                form.ShowDialog();
            }
            else
            {
                var caType = Type.GetType("Infinium." + formName);
                moduleForm = (Form)Activator.CreateInstance(caType, this);
                TopForm = moduleForm;
                moduleForm.ShowDialog();
            }

            _activeNotifySystem.FillUpdates();
            InfiniumNotifyList.ItemsDataTable = _activeNotifySystem.ModulesUpdatesDataTable;
            InfiniumNotifyList.InitializeItems();
        }

        private void InfiniumMinimizeList_ItemClicked(object sender, string formName)
        {
            if (formName == "MessagesForm")
            {
                MessagesButton_Click(null, null);
                InfiniumMinimizeList.RemoveModule(formName);
            }
            else
            {
                var T = new Thread(delegate () { SplashWindow.CreateSplash(); });
                T.Start();

                while (!SplashForm.bCreated) ;

                var moduleForm = InfiniumMinimizeList.GetForm(formName);
                TopForm = moduleForm;
                moduleForm.ShowDialog();
            }

            if (InfiniumNotifyList.Items != null)
                if (InfiniumNotifyList.Items.Count() > 0)
                {
                    _activeNotifySystem.FillUpdates();
                    InfiniumNotifyList.ItemsDataTable = _activeNotifySystem.ModulesUpdatesDataTable;
                    InfiniumNotifyList.InitializeItems();
                }
        }

        private void InfiniumMinimizeList_CloseClicked(object sender, string formName)
        {
            var moduleForm = InfiniumMinimizeList.GetForm(formName);

            InfiniumMinimizeList.RemoveModule(formName);
            Security.ExitFromModule(formName);
            moduleForm.Dispose();

            GC.Collect();
        }

        private void InfiniumClock_Click(object sender, EventArgs e)
        {
            var T = new Thread(delegate () { SplashWindow.CreateSplash(); });
            T.Start();

            while (!SplashForm.bCreated) ;

            Form moduleForm = null;

            //check if running
            var form = InfiniumMinimizeList.GetForm("DayPlannerForm");
            if (form != null)
            {
                TopForm = form;
                form.ShowDialog();
            }
            else
            {
                var caType = Type.GetType("Infinium." + "DayPlannerForm");
                moduleForm = (Form)Activator.CreateInstance(caType, this);
                TopForm = moduleForm;
                Security.EnterInModule("DayPlannerForm");
                moduleForm.ShowDialog();
            }

            if (InfiniumNotifyList.Items != null)
                if (InfiniumNotifyList.Items.Count() > 0)
                {
                    _activeNotifySystem.FillUpdates();
                    InfiniumNotifyList.ItemsDataTable = _activeNotifySystem.ModulesUpdatesDataTable;
                    InfiniumNotifyList.InitializeItems();
                }
        }

        private void MessagesButton_Click(object sender, EventArgs e)
        {
            var phantomForm = new PhantomForm();
            phantomForm.Show();

            var messagesForm = new MessagesForm(ref TopForm);

            TopForm = messagesForm;

            Security.EnterInModule("MessagesForm");

            messagesForm.ShowDialog();

            phantomForm.Close();
            phantomForm.Dispose();

            TopForm = null;

            if (InfiniumNotifyList.Items != null)
                if (InfiniumNotifyList.Items.Count() > 0)
                {
                    _activeNotifySystem.FillUpdates();
                    InfiniumNotifyList.ItemsDataTable = _activeNotifySystem.ModulesUpdatesDataTable;
                    InfiniumNotifyList.InitializeItems();
                }
        }

        private void PhotoBox_Click(object sender, EventArgs e)
        {
            var T = new Thread(delegate () { SplashWindow.CreateSplash(); });
            T.Start();

            while (!SplashForm.bCreated) ;

            Form moduleForm = null;

            //check if running
            var form = InfiniumMinimizeList.GetForm("PersonalSettingsForm");
            if (form != null)
            {
                TopForm = form;
                form.ShowDialog();
            }
            else
            {
                var caType = Type.GetType("Infinium." + "PersonalSettingsForm");
                moduleForm = (Form)Activator.CreateInstance(caType, this);
                TopForm = moduleForm;
                Security.EnterInModule("PersonalSettingsForm");
                moduleForm.ShowDialog();
            }

            if (InfiniumNotifyList.Items != null)
                if (InfiniumNotifyList.Items.Count() > 0)
                {
                    _activeNotifySystem.FillUpdates();
                    InfiniumNotifyList.ItemsDataTable = _activeNotifySystem.ModulesUpdatesDataTable;
                    InfiniumNotifyList.InitializeItems();
                }
        }

        private void kryptonButton1_Click(object sender, EventArgs e)
        {
        }
    }
}
