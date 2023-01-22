using System;
using System.Windows.Forms;

namespace Infinium
{
    public partial class SampleDecorManually : Form
    {
        private const int eHide = 2;
        private const int eShow = 1;
        private const int eClose = 3;
        private const int eMainMenu = 4;

        public bool PressOK;
        public string Product = string.Empty;
        public string Decor = string.Empty;
        public string Color = string.Empty;
        public string Patina = string.Empty;
        public string LabelsCount = "1";
        public string LabelsLength = "0";
        public string LabelsHeight = "0";
        public string LabelsWidth = "0";
        public string PositionsCount = "1";
        private int FormEvent;

        private Form MainForm;

        public SampleDecorManually(Form tMainForm)
        {
            MainForm = tMainForm;

            InitializeComponent();
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
            Product = textBox1.Text;
            Decor = textBox2.Text;
            Color = textBox3.Text;
            Patina = textBox4.Text;
            LabelsHeight = textBox5.Text;
            LabelsWidth = textBox6.Text;
            LabelsLength = textBox7.Text;
            PositionsCount = textBox8.Text;
            LabelsCount = textBox9.Text;

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
