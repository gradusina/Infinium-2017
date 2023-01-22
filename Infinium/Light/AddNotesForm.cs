using System;
using System.Windows.Forms;

namespace Infinium
{
    public partial class AddNotesForm : Form
    {
        private const int eHide = 2;
        private const int eShow = 1;
        private const int eClose = 3;

        private int FormEvent;

        private Form TopForm;

        public static bool bCanceled;

        private LightNotes LightNotes;

        public AddNotesForm(ref LightNotes tLightNotes, ref Form tTopForm)
        {
            InitializeComponent();

            LightNotes = tLightNotes;

            TopForm = tTopForm;
        }

        //public AddNewsForm(ref Infinium.LightNews tLightNews, int SenderTypeID, string HeaderText, string BodyText, int iNewsID, DateTime dDateTime, ref Form tTopForm)
        //{
        //    InitializeComponent();

        //    LightNews = tLightNews;

        //    Medium.Text = LightNews.UsersDataTable.Select("UserID = " + Security.CurrentUserID)[0]["Name"].ToString();
        //    Hight.Text = LightNews.DepartmentsDataTable.Select("DepartmentID = " + LightNews.UsersDataTable.Select("UserID = " + Security.CurrentUserID)[0]["DepartmentID"])[0]["DepartmentName"].ToString();

        //    TopForm = tTopForm;

        //    if (SenderTypeID == 0)
        //        Medium.Checked = true;
        //    else
        //        Hight.Checked = true;

        //    HeaderTextEdit.Text = HeaderText;
        //    BodyTextEdit.Text = BodyText;

        //    iNewsIDEdit = iNewsID;
        //    DateTime = dDateTime;

        //    CreateAttachments();
        //    CopyAttachs(iNewsID);
        //}

        private void OKDateButton_Click(object sender, EventArgs e)
        {
            if (HeaderTextEdit.Text.Length == 0)
                return;

            int P = 0;

            if (High.Checked)
                P = 1;
            if (Medium.Checked)
                P = 2;
            if (Low.Checked)
                P = 3;

            LightNotes.AddNotes(HeaderTextEdit.Text, P);

            FormEvent = eClose;
            AnimateTimer.Enabled = true;
        }

        private void CancelDateButton_Click(object sender, EventArgs e)
        {
            bCanceled = true;

            FormEvent = eClose;
            AnimateTimer.Enabled = true;
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

        private void AddNewsForm_Shown(object sender, EventArgs e)
        {
            FormEvent = eShow;
            AnimateTimer.Enabled = true;
        }
    }
}
