using System;
using System.Windows.Forms;

namespace Infinium
{
    public partial class SampleDecorLabelInfoMenu : Form
    {
        private const int eHide = 2;
        private const int eShow = 1;
        private const int eClose = 3;
        private const int eMainMenu = 4;

        private bool NeedLength;
        private bool NeedHeight;
        private bool NeedWidth;
        public bool PressOK;
        public int LabelsCount = 1;
        public int LabelsLength;
        public int LabelsHeight;
        public int LabelsWidth;
        public int PositionsCount = 1;
        private int FormEvent;

        private Form MainForm;

        public SampleDecorLabelInfoMenu(Form tMainForm, bool bNeedLength, bool bNeedHeight, bool bNeedWidth)
        {
            MainForm = tMainForm;
            NeedLength = bNeedLength;
            NeedHeight = bNeedHeight;
            NeedWidth = bNeedWidth;
            InitializeComponent();
            if (!bNeedLength)
                tbLength.Enabled = false;
            if (!bNeedHeight)
                tbHeight.Enabled = false;
            if (!bNeedWidth)
                tbWidth.Enabled = false;
        }

        private void AnimateTimer_Tick(object sender, EventArgs e)
        {
            if (!DatabaseConfigsManager.Animation)
            {
                this.Opacity = 1;

                if (FormEvent == eClose || FormEvent == eHide)
                {
                    AnimateTimer.Enabled = false;

                    if (FormEvent == eClose)
                    {
                        this.Close();
                    }

                    if (FormEvent == eHide)
                    {
                        MainForm.Activate();
                        this.Hide();
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
                if (Convert.ToDecimal(this.Opacity) != Convert.ToDecimal(0.00))
                    this.Opacity = Convert.ToDouble(Convert.ToDecimal(this.Opacity) - Convert.ToDecimal(0.05));
                else
                {
                    AnimateTimer.Enabled = false;

                    if (FormEvent == eClose)
                    {
                        this.Close();
                    }

                    if (FormEvent == eHide)
                    {
                        MainForm.Activate();
                        this.Hide();
                    }
                }

                return;
            }


            if (FormEvent == eShow || FormEvent == eShow)
            {
                if (this.Opacity != 1)
                    this.Opacity += 0.05;
                else
                {
                    AnimateTimer.Enabled = false;
                    SplashForm.CloseS = true;
                }

                return;
            }
        }

        private void btnOKInput_Click(object sender, EventArgs e)
        {
            PressOK = true;
            if (NeedLength)
                int.TryParse(tbLength.Text, out LabelsLength);
            if (NeedHeight)
                int.TryParse(tbHeight.Text, out LabelsHeight);
            if (NeedWidth)
                int.TryParse(tbWidth.Text, out LabelsWidth);
            int.TryParse(tbLabelsCount.Text, out LabelsCount);
            int.TryParse(tbPositionsCount.Text, out PositionsCount);
            FormEvent = eClose;
            AnimateTimer.Enabled = true;
        }

        private void btnCancelInput_Click(object sender, EventArgs e)
        {
            FormEvent = eClose;
            AnimateTimer.Enabled = true;
        }
    }
}
