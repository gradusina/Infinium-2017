﻿using System;
using System.Windows.Forms;

namespace Infinium
{
    public partial class AuthorizationForm : Form
    {
        private const int eHide = 2;
        private const int eShow = 1;
        private const int eClose = 3;

        private int FormEvent;

        public bool bCanceled;

        private Form TopForm;

        public AuthorizationForm(ref Form tTopForm)
        {
            InitializeComponent();

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

        private void CreateFolderForm_Shown(object sender, EventArgs e)
        {
            FormEvent = eShow;
            AnimateTimer.Enabled = true;
        }

        private void CancelFolderButton_Click(object sender, EventArgs e)
        {
            bCanceled = true;

            FormEvent = eClose;
            AnimateTimer.Enabled = true;
        }

        private void OKFolderButton_Click(object sender, EventArgs e)
        {
            if (CodeTextBox.Text.Length == 0)
                return;

            if (Security.CheckAuth(CodeTextBox.Text))
            {
                FormEvent = eClose;
                AnimateTimer.Enabled = true;
            }
            else
            {
                CodeTextBox.Clear();
            }
        }

        private void FolderNameTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                OKFolderButton_Click(null, null);
        }
    }
}
