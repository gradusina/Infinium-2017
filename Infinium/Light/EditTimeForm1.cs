﻿using System;
using System.Windows.Forms;

namespace Infinium
{
    public partial class EditTimeForm1 : Form
    {
        private const int eHide = 2;
        private const int eShow = 1;
        private const int eClose = 3;

        private int FormEvent;
        private int DayEventType;
        private int WorkDayID;

        public const int dStartDay = 0;
        public const int dBreakDay = 1;
        public const int dContinueDay = 2;
        public const int dEndDay = 3;

        private Form TopForm;

        public EditTimeForm1(ref Form tTopForm, int tDayEventType, int iWorkDayID, object NewDateTime, object StartDateTime)
        {
            InitializeComponent();

            DayEventType = tDayEventType;
            WorkDayID = iWorkDayID;

            TopForm = tTopForm;

            timeEdit1.EditValue = DateTime.Now;
            if (StartDateTime != DBNull.Value)
                timeEdit1.EditValue = Convert.ToDateTime(StartDateTime);
            if (DayEventType == dBreakDay)
            {
                label2.Text = "Изменить время начала перерыва";

                if (NewDateTime != DBNull.Value)
                {
                    timeEdit1.EditValue = Convert.ToDateTime(NewDateTime);
                    FactTimeLabel.Text = "с " + Convert.ToDateTime(NewDateTime).ToString("HH:mm") + " на";
                }
            }
            if (DayEventType == dContinueDay)
            {
                label2.Text = "Изменить время окончания перерыва";

                if (NewDateTime != DBNull.Value)
                {
                    timeEdit1.EditValue = Convert.ToDateTime(NewDateTime);
                    FactTimeLabel.Text = "с " + Convert.ToDateTime(NewDateTime).ToString("HH:mm") + " на";
                }
            }
            if (DayEventType == dStartDay)
            {
                label2.Text = "Изменить время начала рабочего дня";

                if (NewDateTime != DBNull.Value)
                {
                    timeEdit1.EditValue = Convert.ToDateTime(NewDateTime);
                    FactTimeLabel.Text = "с " + Convert.ToDateTime(NewDateTime).ToString("HH:mm") + " на";
                }
            }
            if (DayEventType == dEndDay)
            {
                label2.Text = "Изменить время завершения рабочего дня";

                if (NewDateTime != DBNull.Value)
                {
                    timeEdit1.EditValue = Convert.ToDateTime(NewDateTime);
                    FactTimeLabel.Text = "с " + Convert.ToDateTime(NewDateTime).ToString("HH:mm") + " на";
                }
            }

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

        private void EditTimeForm_Shown(object sender, EventArgs e)
        {
            FormEvent = eShow;
            AnimateTimer.Enabled = true;
        }

        private void MenuCloseButton_Click(object sender, EventArgs e)
        {
            FormEvent = eClose;
            AnimateTimer.Enabled = true;
        }

        private void AcceptButton_Click(object sender, EventArgs e)
        {
            DateTime D = Convert.ToDateTime(timeEdit1.EditValue);

            if (DayEventType == dStartDay)
            {
                WorkTimeRegister.EditStartWorkDay(WorkDayID, D);
            }
            if (DayEventType == dEndDay)
            {
                WorkTimeRegister.EditEndWorkDay(WorkDayID, D);
            }
            if (DayEventType == dBreakDay)
            {
                WorkTimeRegister.EditBreakStartWorkDay(WorkDayID, D);
            }
            if (DayEventType == dContinueDay)
            {
                WorkTimeRegister.EditBreakEndWorkDay(WorkDayID, D);
            }

            FormEvent = eClose;
            AnimateTimer.Enabled = true;
        }
    }
}
