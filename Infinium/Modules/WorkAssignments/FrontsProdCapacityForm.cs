using Infinium.Modules.WorkAssignments;

using System;
using System.Windows.Forms;

namespace Infinium
{
    public partial class FrontsProdCapacityForm : Form
    {
        private const int eHide = 2;
        private const int eShow = 1;
        private const int eClose = 3;

        private int FormEvent = 0;
        private int TechStoreID, InsetTypeID, PatinaID, TechHeight, TechWidth = 0;

        private Form MainForm = null;
        private Form TopForm = null;

        private FrontsProdCapacity ClassManager;

        public FrontsProdCapacityForm(Form tMainForm, int iTechStoreID, int iInsetTypeID, int iPatinaID, int iHeight, int iWidth)
        {
            TechStoreID = iTechStoreID;
            InsetTypeID = iInsetTypeID;
            PatinaID = iPatinaID;
            TechHeight = iHeight;
            TechWidth = iWidth;
            MainForm = tMainForm;
            InitializeComponent();

            this.MaximumSize = Screen.PrimaryScreen.WorkingArea.Size;

            while (!SplashForm.bCreated) ;
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

        private void FrontsProdCapacityForm_Shown(object sender, EventArgs e)
        {
            FormEvent = eShow;
            AnimateTimer.Enabled = true;
        }

        private void MenuCloseButton_Click(object sender, EventArgs e)
        {
            FormEvent = eClose;
            AnimateTimer.Enabled = true;
        }

        private void FrontsProdCapacityForm_Load(object sender, EventArgs e)
        {
            ClassManager = new FrontsProdCapacity(TechStoreID, InsetTypeID, PatinaID, TechHeight, TechWidth);
            ClassManager.Test();
        }
    }
}
