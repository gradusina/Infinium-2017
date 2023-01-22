using System;
using System.Windows.Forms;

namespace Infinium
{
    public partial class SamplesLabelsForm : Form
    {
        private const int eHide = 2;
        private const int eShow = 1;
        private const int eClose = 3;
        private const int eMainMenu = 4;

        public bool PressOK;
        public string DocDateTime = string.Empty;
        public string Batch = string.Empty;
        public string Pallet = string.Empty;
        public string Profile = string.Empty;
        public string LabelsCount = "1";
        public string Serviceman = "0";
        public string Milling = "0";
        public string DocDateTime1 = string.Empty;
        public string Batch1 = string.Empty;
        public string Pallet1 = string.Empty;
        public string Profile1 = string.Empty;
        public string LabelsCount1 = "1";
        public string Serviceman1 = "0";
        public string Milling1 = "0";
        private int FormEvent;

        private Form MainForm;

        public SamplesLabelsForm(Form tMainForm)
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
            if (dateTimePicker1.Enabled)
                DocDateTime = dateTimePicker1.Value.ToString("dd.MM.yyyy");
            Batch = textBox2.Text;
            Pallet = textBox3.Text;
            Profile = textBox4.Text;
            Serviceman = textBox5.Text;
            Milling = textBox6.Text;
            LabelsCount = textBox9.Text;

            if (dateTimePicker2.Enabled)
                DocDateTime1 = dateTimePicker2.Value.ToString("dd.MM.yyyy");
            Batch1 = textBox8.Text;
            Pallet1 = textBox7.Text;
            Profile1 = textBox1.Text;
            Serviceman1 = textBox10.Text;
            Milling1 = textBox11.Text;
            LabelsCount1 = textBox12.Text;

            FormEvent = eClose;
            AnimateTimer.Enabled = true;
        }

        private void btnCancelInput_Click(object sender, EventArgs e)
        {
            FormEvent = eClose;
            AnimateTimer.Enabled = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            dateTimePicker1.Enabled = !dateTimePicker1.Enabled;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            dateTimePicker2.Enabled = !dateTimePicker2.Enabled;
        }
    }
}
