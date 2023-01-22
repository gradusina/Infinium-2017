using System;
using System.Data;
using System.Threading;
using System.Windows.Forms;

namespace Infinium
{
    public partial class CoderBlogForm : InfiniumForm
    {
        private const int eHide = 2;
        private const int eShow = 1;
        private const int eClose = 3;

        private int FormEvent;

        private LightStartForm LightStartForm;
        private Form TopForm;

        private DataTable SDT = new DataTable();

        private CoderBlog CoderBlog;

        public CoderBlogForm(LightStartForm tLightStartForm)
        {
            InitializeComponent();


            LightStartForm = tLightStartForm;

            MaximumSize = Screen.PrimaryScreen.WorkingArea.Size;

            CoderBlog = new CoderBlog();

            //ActiveNotifySystem.ClearSubscribesRecords(Security.CurrentUserID, LightTile.Name);

            Initialize();

            //LightNewsContainer.PageChanged(null);

            ActiveNotifySystem.ClearSubscribesRecords(Security.CurrentUserID, Name);
            //ActiveNotifySystem.ClearModuleUpdates(LightTile.Name);
            //ActiveNotifySystem.ClearCurrentOpenModuleUpdates(this.Name);

            //OnANSUpdate();

            while (!SplashForm.bCreated) ;
        }

        private void LightNewsForm_Shown(object sender, EventArgs e)
        {
            while (!SplashForm.bCreated) ;
            CoderBlogContainer.Visible = true;

            FormEvent = eShow;
            AnimateTimer.Enabled = true;
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

                        LightStartForm.CloseForm(this);
                    }

                    if (FormEvent == eHide)
                    {
                        LightStartForm.HideForm(this);
                    }

                    return;
                }

                if (FormEvent == eShow)
                {
                    CoderBlogContainer.CreateNews();
                    CoderBlogContainer.Focus();
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

                        LightStartForm.CloseForm(this);
                    }

                    if (FormEvent == eHide)
                    {
                        LightStartForm.HideForm(this);
                    }

                }

                return;
            }


            if (FormEvent == eShow)
            {
                if (Opacity != 1)
                    Opacity += 0.05;
                else
                {
                    CoderBlogContainer.CreateNews();
                    CoderBlogContainer.Focus();
                    AnimateTimer.Enabled = false;
                    SplashForm.CloseS = true;
                }
            }
        }

        private void NavigateMenuCloseButton_Click(object sender, EventArgs e)
        {
            CoderBlogContainer.Visible = false;

            FormEvent = eClose;
            AnimateTimer.Enabled = true;
        }

        private void NavigateMenuBackButton_Click(object sender, EventArgs e)
        {
            FormEvent = eHide;
            AnimateTimer.Enabled = true;
        }


        private void Initialize()
        {
            CoderBlogContainer.CurrentUserID = Security.CurrentUserID;
            CoderBlogContainer.NewsDataTable = CoderBlog.BlogNewsDataTable.Copy();
            CoderBlogContainer.UsersDataTable = CoderBlog.UsersDataTable.Copy();
            CoderBlogContainer.AttachsDT = CoderBlog.AttachmentsDataTable.Copy();
            CoderBlogContainer.CommentsDT = CoderBlog.CommentsDataTable.Copy();
            CoderBlogContainer.NewsLikesDT = CoderBlog.BlogNewsLikesDataTable.Copy();
            CoderBlogContainer.CommentsLikesDT = CoderBlog.CommentsLikesDataTable.Copy();
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

        private void LightNewsForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            //LightNews.ClearCommentsSubs(Security.CurrentUserID);
        }

        //private void GetNewNews()
        //{
        //    int c = LightNews.GetNewsUpdatesCount();

        //    if (c > 0)
        //    {
        //        UpdateNewsButton.Text = "Обновления " + c.ToString();
        //    }


        //    //LightNews.FillNews();
        //    ////LightNews.RefillSubs();
        //    ////LightNews.RefillComments();
        //    ////LightNews.RefillCurrentComments();
        //    //ActiveNotifySystem.ClearSubscribesRecords(Security.CurrentUserID, LightTile.Name);
        //    //ActiveNotifySystem.ClearModuleUpdates(LightTile.Name);
        //    //ActiveNotifySystem.ClearCurrentOpenModuleUpdates(this.Name);
        //}

        private void LightNewsForm_ANSUpdate(object sender)
        {
            int c = CoderBlog.GetNewsUpdatesCount();

            if (c > 0)
            {
                UpdateBlogButton.Text = "Обновления: " + c;
                UpdateBlogButton.Visible = true;
            }

            //ActiveNotifySystem.ClearNewUpdates("CoderBlogButton");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //LightNewsContainer.Refresh();
            //LightNewsContainer.Visible = true;
        }

        private bool bC;

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (bC)
            {
                while (SplashWindow.bSmallCreated)
                    CoverWaitForm.CloseS = true;

                bC = false;
            }
        }

        private void LightNewsContainer_EditCommentClicked(object sender, int NewsID, int NewsCommentID)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            int O = CoderBlogContainer.Offset;
            int SC = CoderBlogContainer.ScrollContainer.Height;
            int H = CoderBlogContainer.Height;
        }

        private void LightNewsContainer_VerticalScrollBar_ScrollPositionChanged(object sender, int tOffset)
        {

        }

        private void button1_Click_2(object sender, EventArgs e)
        {
            CoderBlogContainer.SetNoClip();
            //LightNewsContainer.Refresh();
        }

        private void LightNewsContainer_CommentsClicked(object sender, EventArgs e)
        {
            //this.Refresh();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            CoderBlogContainer.SetClipStandard();
        }

        private void MoreBlogNewsButton_Click(object sender, EventArgs e)
        {
            int iNewsCount = CoderBlogContainer.NewsCount;

            CoderBlog.FillMoreNews(CoderBlogContainer.NewsCount + 20);
            CoderBlogContainer.NewsDataTable = CoderBlog.BlogNewsDataTable.Copy();
            CoderBlogContainer.CreateNews();
            CoderBlogContainer.ScrollToNews(iNewsCount);
            CoderBlogContainer.Focus();
            MoreBlogNewsButton.Visible = false;
            CoderBlogContainer.bNeedMoreNews = false;
        }

        private void CoderBlogContainer_AttachClicked(object sender, int NewsAttachID)
        {
            PhantomForm PhantomForm = new PhantomForm();
            PhantomForm.Show();

            AttachDownloadForm AttachDownloadForm = new AttachDownloadForm(NewsAttachID, ref CoderBlog.FM, ref CoderBlog);

            TopForm = AttachDownloadForm;

            AttachDownloadForm.ShowDialog();

            PhantomForm.Close();

            PhantomForm.Dispose();

            TopForm = null;

            AttachDownloadForm.Dispose();
        }

        private void CoderBlogContainer_CommentLikeClicked(object sender, int NewsID, int NewsCommentID)
        {
            CoderBlog.LikeComments(Security.CurrentUserID, NewsID, NewsCommentID);

            CoderBlog.FillLikes();
            CoderBlogContainer.CommentsLikesDT = CoderBlog.CommentsLikesDataTable;
            CoderBlogContainer.ReloadLikes(NewsID);
            CoderBlogContainer.Focus();
        }

        private void CoderBlogContainer_CommentSendButtonClicked(object sender, string Text, int NewsID, int SenderID, bool bEdit, int NewsCommentID)
        {
            Thread T = new Thread(delegate ()
            {
                SplashWindow.CreateCoverSplash(CoderBlogContainer.Top, CoderBlogContainer.Left,
                                     CoderBlogContainer.Height, CoderBlogContainer.Width);
            });
            T.Start();

            while (!SplashWindow.bSmallCreated) ;

            if (bEdit)
                CoderBlog.EditComments(SenderID, NewsCommentID, Text);
            else
                CoderBlog.AddComments(SenderID, NewsID, Text);

            CoderBlog.ReloadNews(CoderBlogContainer.NewsCount);
            CoderBlog.ReloadComments();
            CoderBlogContainer.NewsDataTable = CoderBlog.BlogNewsDataTable.Copy();
            CoderBlogContainer.CommentsDT = CoderBlog.CommentsDataTable.Copy();
            CoderBlogContainer.CreateNews();
            CoderBlogContainer.ScrollToTop();
            CoderBlogContainer.ReloadNewsItem(NewsID, true);

            bC = true;
        }

        private void CoderBlogContainer_CommentsQuoteLabelClicked(object sender)
        {
            InfiniumTips.ShowTip(this, 50, 85, "Комментарий скопирован в буфер обмена...", 1700);
        }

        private void CoderBlogContainer_EditNewsClicked(object sender, int NewsID)
        {
            PhantomForm PhantomForm = new PhantomForm();
            PhantomForm.Show();

            AddBlogNewsForm AddBlogNewsForm = new AddBlogNewsForm(ref CoderBlog, CoderBlog.GetThisNewsSenderTypeID(NewsID),
                                                      CoderBlog.GetThisNewsHeaderText(NewsID),
                                                      CoderBlog.GetThisNewsBodyText(NewsID),
                                                      NewsID, CoderBlog.GetThisNewsDateTime(NewsID), ref TopForm);

            TopForm = AddBlogNewsForm;

            AddBlogNewsForm.ShowDialog();

            PhantomForm.Close();
            PhantomForm.Dispose();

            TopForm = null;

            if (AddBlogNewsForm.Canceled)
                return;

            Thread T = new Thread(delegate ()
            {
                SplashWindow.CreateCoverSplash(CoderBlogContainer.Top, CoderBlogContainer.Left,
                    CoderBlogContainer.Height, CoderBlogContainer.Width);
            });
            T.Start();

            while (!SplashWindow.bSmallCreated) ;


            CoderBlog.ReloadNews(CoderBlogContainer.NewsCount);
            CoderBlog.ReloadComments();
            CoderBlog.ReloadAttachments();
            CoderBlogContainer.NewsDataTable = CoderBlog.BlogNewsDataTable.Copy();
            CoderBlogContainer.CommentsDT = CoderBlog.CommentsDataTable.Copy();
            CoderBlogContainer.AttachsDT = CoderBlog.AttachmentsDataTable.Copy();
            CoderBlogContainer.CreateNews();

            bC = true;
        }

        private void CoderBlogContainer_LikeClicked(object sender, int NewsID)
        {
            CoderBlog.LikeNews(Security.CurrentUserID, NewsID);

            CoderBlog.FillLikes();
            CoderBlogContainer.NewsLikesDT = CoderBlog.BlogNewsLikesDataTable;
            CoderBlogContainer.ReloadLikes(NewsID);
            CoderBlogContainer.Focus();
        }

        private void CoderBlogContainer_RemoveNewsClicked(object sender, int NewsID)
        {
            PhantomForm PhantomForm = new PhantomForm();
            PhantomForm.Show();

            LightMessageBoxForm LightMessageBoxForm = new LightMessageBoxForm(true, "Сообщение будет удалено безвозвратно.\nПродолжить?",
                                                                                    "Удаление сообщения");

            TopForm = LightMessageBoxForm;

            LightMessageBoxForm.ShowDialog();

            TopForm = null;

            PhantomForm.Close();
            PhantomForm.Dispose();

            if (LightMessageBoxForm.OKCancel)
            {
                Thread T = new Thread(delegate ()
                {
                    SplashWindow.CreateCoverSplash(CoderBlogContainer.Top, CoderBlogContainer.Left,
                    CoderBlogContainer.Height, CoderBlogContainer.Width);
                });
                T.Start();

                while (!SplashWindow.bSmallCreated) ;

                CoderBlog.RemoveNews(NewsID);
                CoderBlog.ReloadNews(CoderBlogContainer.NewsCount);
                CoderBlog.ReloadComments();
                CoderBlogContainer.NewsDataTable = CoderBlog.BlogNewsDataTable.Copy();
                CoderBlogContainer.CommentsDT = CoderBlog.CommentsDataTable.Copy();
                CoderBlogContainer.CreateNews();
            }

            LightMessageBoxForm.Dispose();

            bC = true;
        }

        private void CoderBlogContainer_RemoveCommentClicked(object sender, int NewsID, int NewsCommentID)
        {
            PhantomForm PhantomForm = new PhantomForm();
            PhantomForm.Show();

            LightMessageBoxForm LightMessageBoxForm = new LightMessageBoxForm(true, "Комментарий будет удален.\nПродолжить?",
                                                                                    "Удаление комментария");

            TopForm = LightMessageBoxForm;

            LightMessageBoxForm.ShowDialog();

            TopForm = null;

            PhantomForm.Close();
            PhantomForm.Dispose();


            if (LightMessageBoxForm.OKCancel)
            {
                Thread T = new Thread(delegate ()
                {
                    SplashWindow.CreateCoverSplash(CoderBlogContainer.Top, CoderBlogContainer.Left,
                    CoderBlogContainer.Height, CoderBlogContainer.Width);
                });
                T.Start();

                while (!SplashWindow.bSmallCreated) ;

                CoderBlog.RemoveComment(NewsCommentID);

                CoderBlog.ReloadNews(CoderBlogContainer.NewsCount);
                CoderBlog.ReloadComments();
                CoderBlogContainer.NewsDataTable = CoderBlog.BlogNewsDataTable.Copy();
                CoderBlogContainer.CommentsDT = CoderBlog.CommentsDataTable.Copy();
                CoderBlogContainer.CreateNews();
                CoderBlogContainer.ReloadNewsItem(NewsID, true);

                bC = true;
            }
        }

        private void CoderBlogContainer_Refreshed(object sender, EventArgs e)
        {
            Thread.Sleep(100);

            while (SplashWindow.bSmallCreated)
                CoverWaitForm.CloseS = true;
        }

        private void CoderBlogContainer_NoNeedMoreNews(object sender, EventArgs e)
        {
            if (MoreBlogNewsButton.Visible)
                MoreBlogNewsButton.Visible = false;
        }

        private void CoderBlogContainer_NewsQuoteLabelClicked(object sender)
        {
            InfiniumTips.ShowTip(this, 50, 85, "Текст скопирован в буфер обмена...", 1700);
        }

        private void CoderBlogContainer_NeedMoreNews(object sender, EventArgs e)
        {
            if (CoderBlog.IsMoreNews(CoderBlogContainer.NewsCount))
                MoreBlogNewsButton.Visible = true;
            else
                MoreBlogNewsButton.Visible = false;
        }

        private void UpdateBlogButton_Click(object sender, EventArgs e)
        {
            Thread T = new Thread(delegate ()
            {
                SplashWindow.CreateCoverSplash(CoderBlogContainer.Top, CoderBlogContainer.Left,
                                     CoderBlogContainer.Height, CoderBlogContainer.Width);
            });
            T.Start();

            while (!SplashWindow.bSmallCreated) ;

            CoderBlog.ReloadSubscribes();
            CoderBlog.ReloadNews(20);//default count
            CoderBlog.ReloadComments();
            CoderBlog.ReloadAttachments();
            CoderBlog.FillLikes();
            CoderBlogContainer.NewsDataTable = CoderBlog.BlogNewsDataTable.Copy();
            CoderBlogContainer.CommentsDT = CoderBlog.CommentsDataTable.Copy();
            CoderBlogContainer.AttachsDT = CoderBlog.AttachmentsDataTable.Copy();
            CoderBlogContainer.NewsLikesDT = CoderBlog.BlogNewsLikesDataTable;
            CoderBlogContainer.CommentsLikesDT = CoderBlog.CommentsLikesDataTable;
            CoderBlogContainer.CreateNews();
            CoderBlogContainer.ScrollToTop();
            CoderBlogContainer.Focus();

            ActiveNotifySystem.ClearSubscribesRecords(Security.CurrentUserID, Name);
            //ActiveNotifySystem.ClearModuleUpdates(LightTile.Name);
            //ActiveNotifySystem.ClearCurrentOpenModuleUpdates(this.Name);

            UpdateBlogButton.Visible = false;

            bC = true;
        }

        private void AddBlogNewsButton_Click(object sender, EventArgs e)
        {
            PhantomForm PhantomForm = new PhantomForm();
            PhantomForm.Show();

            AddBlogNewsForm AddBlogNewsForm = new AddBlogNewsForm(ref CoderBlog, ref TopForm);

            TopForm = AddBlogNewsForm;

            AddBlogNewsForm.ShowDialog();

            PhantomForm.Close();
            PhantomForm.Dispose();

            TopForm = null;

            if (AddBlogNewsForm.Canceled)
                return;

            Thread T = new Thread(delegate ()
            {
                SplashWindow.CreateCoverSplash(CoderBlogContainer.Top, CoderBlogContainer.Left,
                                     CoderBlogContainer.Height, CoderBlogContainer.Width);
            });
            T.Start();

            while (!SplashWindow.bSmallCreated) ;


            CoderBlog.ReloadNews(20);//default
            CoderBlog.ReloadComments();
            CoderBlog.ReloadAttachments();
            CoderBlogContainer.NewsDataTable = CoderBlog.BlogNewsDataTable.Copy();
            CoderBlogContainer.CommentsDT = CoderBlog.CommentsDataTable.Copy();
            CoderBlogContainer.AttachsDT = CoderBlog.AttachmentsDataTable.Copy();
            CoderBlogContainer.CreateNews();
            CoderBlogContainer.ScrollToTop();
            CoderBlogContainer.Focus();

            bC = true;

        }

        private void MinimizeButton_Click(object sender, EventArgs e)
        {
            CoderBlogContainer.Visible = false;

            FormEvent = eHide;
            AnimateTimer.Enabled = true;
        }
    }
}
