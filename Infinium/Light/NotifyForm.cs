﻿using System;
using System.Windows.Forms;

namespace Infinium
{
    public partial class NotifyForm : Form
    {
        private const int eHide = 2;
        private const int eShow = 1;
        private const int eClose = 3;

        private const int nCloseNotifyOnly = 0;
        private const int nActivateInfinium = 1;
        private const int nActivateModule = 2;

        private int FormEvent;
        private int ClickEvent = -1;

        public static bool bShowed;
        public static bool bCloseNotify;

        public bool bCloseOnly;

        private LightStartForm LightStartForm;

        private int ModuleID = -1;

        private ActiveNotifySystem ActiveNotifySystem;

        public NotifyForm(LightStartForm tLightStartForm, ref ActiveNotifySystem tActiveNotifySystem, int tModuleID, int Count, int MoreCount)
        {
            InitializeComponent();

            ActiveNotifySystem = tActiveNotifySystem;

            LightStartForm = tLightStartForm;

            if (MoreCount > 0)
            {
                MoreUpdatesLabel.Visible = true;
                MoreUpdatesLabel.Text = "и еще (" + MoreCount + ") обновлений";
            }

            ModuleID = tModuleID;

            NewUpdatesLabel.Text = "Новых: " + Count;
            PictureBox.Image = ActiveNotifySystem.GetModuleImage(tModuleID);
            ModuleNameLabel.Text = ActiveNotifySystem.GetModuleName(tModuleID);
        }

        private void AnimateTimer_Tick(object sender, EventArgs e)
        {
            if (FormEvent == eClose || FormEvent == eHide)
            {
                if (Convert.ToDecimal(Opacity) != Convert.ToDecimal(0.00))
                    Opacity = Convert.ToDouble(Convert.ToDecimal(Opacity) - Convert.ToDecimal(0.05));
                else
                {
                    AnimateTimer.Enabled = false;

                    if (FormEvent == eClose)
                    {
                        bShowed = false;

                        if (ClickEvent == nActivateModule)
                        {
                            LightStartForm.Activate();
                            //ActiveNotifySystem.StartModule(ModuleBtnName);

                            LightStartForm.StartModuleFromNotify(ModuleID);

                            Close();
                            return;
                        }

                        //if (ClickEvent == nActivateInfinium)
                        //{
                        //    LightStartForm.Activate();

                        //    this.Close();
                        //    return;
                        //}

                        if (ClickEvent == nCloseNotifyOnly)
                        {

                            Close();
                            return;
                        }
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
                    bShowed = true;
                }
            }
        }

        private void NewsCommentsForm_Shown(object sender, EventArgs e)
        {
            FormEvent = eShow;
            AnimateTimer.Enabled = true;

            StartPosition = FormStartPosition.Manual;
            Left = Screen.PrimaryScreen.WorkingArea.Width - Width - 5;
            Top = Screen.PrimaryScreen.WorkingArea.Height - Height - 5;

            //Stream str = Properties.Resources._01_01;

            //using (SoundPlayer snd = new SoundPlayer(str))
            //{
            //    snd.Play();
            //}

            //str.Dispose();
        }

        private void MenuCloseButton_Click(object sender, EventArgs e)
        {
            FormEvent = eClose;
            AnimateTimer.Enabled = true;
        }

        private void LifeTimer_Tick(object sender, EventArgs e)
        {
            //if (bCloseNotify)
            //{
            //FormEvent = eClose;
            //AnimateTimer.Enabled = true;
            //LifeTimer.Enabled = false;
            //}
        }

        private void NotifyForm_Click(object sender, EventArgs e)
        {
            ClickEvent = nActivateInfinium;
            FormEvent = eClose;
            AnimateTimer.Enabled = true;
            WatchTimer.Enabled = false;
        }

        //private const int CS_DROPSHADOW = 0x20000;

        //protected override CreateParams CreateParams
        //{
        //    get
        //    {
        //        CreateParams cp = base.CreateParams;
        //        cp.ClassStyle |= CS_DROPSHADOW;
        //        return cp;
        //    }
        //}

        private void WatchTimer_Tick(object sender, EventArgs e)
        {
            if (bCloseNotify)
            {
                bCloseNotify = false;
                FormEvent = eClose;
                AnimateTimer.Enabled = true;
                WatchTimer.Enabled = false;
            }
        }

        private void ansPopupNotify1_Click(object sender, EventArgs e)
        {
            //bCloseNotify = false;
            //FormEvent = eClose;
            //AnimateTimer.Enabled = true;
            //WatchTimer.Enabled = false;           
        }

        private void label2_Click(object sender, EventArgs e)
        {
            ClickEvent = nActivateInfinium;
            FormEvent = eClose;
            AnimateTimer.Enabled = true;
            WatchTimer.Enabled = false;
        }

        private void panel1_Click(object sender, EventArgs e)
        {
            ClickEvent = nActivateInfinium;
            FormEvent = eClose;
            AnimateTimer.Enabled = true;
            WatchTimer.Enabled = false;
        }

        private void ansPopupNotify1_ClickUpdate(object sender, bool Open, string ModuleButtonName)
        {
            if (Open)
            {
                ClickEvent = nActivateModule;
                WatchTimer.Enabled = false;
                FormEvent = eClose;
                AnimateTimer.Enabled = true;
            }
            else
            {
                ClickEvent = nActivateInfinium;
                FormEvent = eClose;
                AnimateTimer.Enabled = true;
                WatchTimer.Enabled = false;
            }
        }

        private void PictureBox_Click(object sender, EventArgs e)
        {
            ClickEvent = nActivateModule;
            FormEvent = eClose;
            AnimateTimer.Enabled = true;
            WatchTimer.Enabled = false;
            bCloseOnly = true;
        }
    }
}
