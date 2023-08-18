using System;
using System.Windows.Forms;

using static Infinium.FrontsCatalog;

namespace Infinium
{
    public partial class ProductDescriptionForm : Form
    {
        private const int eHide = 2;
        private const int eShow = 1;
        private const int eClose = 3;
        private const int eMainMenu = 4;

        public bool PressOK;

        public ConfigImageInfo ConfigImageInfo;

        //public bool ToSite;
        //public bool Latest;
        //public bool Basic;
        //public string Category = string.Empty;
        //public string NameProd = string.Empty;
        //public string Description = string.Empty;
        //public string Sizes = string.Empty;
        //public string Material = string.Empty;

        private int FormEvent;

        private Form MainForm;

        public ProductDescriptionForm(Form tMainForm, ConfigImageInfo configImageInfo)
        {
            this.ConfigImageInfo = configImageInfo;

            MainForm = tMainForm;

            InitializeComponent();

            kryptonRichTextBox4.Text = configImageInfo.Category;
            kryptonRichTextBox5.Text = configImageInfo.Name;
            kryptonRichTextBox1.Text = configImageInfo.Description;
            kryptonRichTextBox2.Text = configImageInfo.Sizes;
            kryptonRichTextBox3.Text = configImageInfo.Material;
            cbToSite.Checked = configImageInfo.ToSite;
            cbLatest.Checked = configImageInfo.Latest;
            cbBasic.Checked = configImageInfo.Basic;
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
            ConfigImageInfo.Category = kryptonRichTextBox4.Text;
            ConfigImageInfo.Name = kryptonRichTextBox5.Text;
            ConfigImageInfo.Description = kryptonRichTextBox1.Text;
            ConfigImageInfo.Sizes = kryptonRichTextBox2.Text;
            ConfigImageInfo.Material = kryptonRichTextBox3.Text;
            ConfigImageInfo.ToSite = cbToSite.Checked;
            ConfigImageInfo.Latest = cbLatest.Checked;
            ConfigImageInfo.Basic = cbBasic.Checked;

            FormEvent = eClose;
            AnimateTimer.Enabled = true;
        }

        private void btnCancelInput_Click(object sender, EventArgs e)
        {
            FormEvent = eClose;
            AnimateTimer.Enabled = true;
        }

        private void ProductDescriptionForm_Load(object sender, EventArgs e)
        {

        }
    }
}
