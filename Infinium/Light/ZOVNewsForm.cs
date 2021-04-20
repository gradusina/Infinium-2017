﻿using System;
using System.Data;
using System.Threading;
using System.Windows.Forms;

namespace Infinium
{
    public partial class ZOVNewsForm : InfiniumForm
    {
        const int eHide = 2;
        const int eShow = 1;
        const int eClose = 3;

        int FormEvent = 0;

        LightStartForm LightStartForm;

        Form TopForm;

        DataTable SDT = new DataTable();

        ZOVNews ZOVNews;


        public ZOVNewsForm(LightStartForm tLightStartForm)
        {
            InitializeComponent();

            LightStartForm = tLightStartForm;

            this.MaximumSize = Screen.PrimaryScreen.WorkingArea.Size;

            ZOVNews = new Infinium.ZOVNews();

            //ActiveNotifySystem.ClearSubscribesRecords(Security.CurrentUserID, this.Name);

            Initialize();



            //LightNewsContainer.PageChanged(null);

            ActiveNotifySystem.ClearSubscribesRecords(Security.CurrentUserID, this.Name);

            //OnANSUpdate();



            while (!SplashForm.bCreated) ;
        }


        private void LightNewsForm_Shown(object sender, EventArgs e)
        {
            while (!SplashForm.bCreated) ;

            LightNewsContainer.Visible = true;
            FormEvent = eShow;
            AnimateTimer.Enabled = true;
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
                    LightNewsContainer.CreateNews();
                    LightNewsContainer.Focus();
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
                if (this.Opacity != 1)
                    this.Opacity += 0.05;
                else
                {
                    LightNewsContainer.CreateNews();
                    LightNewsContainer.Focus();
                    AnimateTimer.Enabled = false;
                    SplashForm.CloseS = true;
                }

                return;
            }
        }

        private void NavigateMenuCloseButton_Click(object sender, EventArgs e)
        {
            LightNewsContainer.Visible = false;

            FormEvent = eClose;
            AnimateTimer.Enabled = true;
        }

        private void Initialize()
        {
            LightNewsContainer.ManagersDT = ZOVNews.ManagersDataTable;
            LightNewsContainer.CurrentUserID = Security.CurrentUserID;
            LightNewsContainer.NewsDataTable = ZOVNews.NewsDataTable.Copy();
            LightNewsContainer.UsersDataTable = ZOVNews.UsersDataTable.Copy();
            LightNewsContainer.AttachsDT = ZOVNews.AttachmentsDataTable.Copy();
            LightNewsContainer.CommentsDT = ZOVNews.CommentsDataTable.Copy();
            LightNewsContainer.NewsLikesDT = ZOVNews.NewsLikesDataTable.Copy();
            LightNewsContainer.CommentsLikesDT = ZOVNews.CommentsLikesDataTable.Copy();
            LightNewsContainer.CreateNews();
            LightNewsContainer.Visible = true;
            LightNewsContainer.Focus();
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
        //    //ActiveNotifySystem.ClearSubscribesRecords(Security.CurrentUserID, this.Name);
        //    //ActiveNotifySystem.ClearModuleUpdates(LightTile.Name);
        //    //ActiveNotifySystem.ClearCurrentOpenModuleUpdates(this.Name);
        //}

        private void LightNewsForm_ANSUpdate(object sender)
        {
            int c = ZOVNews.GetNewsUpdatesCount();

            if (c > 0)
            {
                UpdateNewsButton.Text = "Обновления: " + c.ToString();
                UpdateNewsButton.Visible = true;
            }

            //ActiveNotifySystem.ClearNewUpdates("LightNewsButton");
        }


        private void button1_Click(object sender, EventArgs e)
        {
            //LightNewsContainer.Refresh();
            //LightNewsContainer.Visible = true;
        }

        bool bC = false;


        private void timer1_Tick(object sender, EventArgs e)
        {
            if (bC)
            {
                while (SplashWindow.bSmallCreated)
                    CoverWaitForm.CloseS = true;

                bC = false;
            }
        }


        private void AddNewsButton_Click(object sender, EventArgs e)
        {
            PhantomForm PhantomForm = new PhantomForm();
            PhantomForm.Show();

            AddZOVNewsForm AddNewsForm = new AddZOVNewsForm(ref ZOVNews, ref TopForm);

            TopForm = AddNewsForm;

            AddNewsForm.ShowDialog();

            PhantomForm.Close();
            PhantomForm.Dispose();

            TopForm = null;

            if (AddNewsForm.Canceled)
                return;

            Thread T = new Thread(delegate ()
            {
                SplashWindow.CreateCoverSplash(LightNewsContainer.Top, LightNewsContainer.Left,
                                     LightNewsContainer.Height, LightNewsContainer.Width);
            });
            T.Start();

            while (!SplashWindow.bSmallCreated) ;


            ZOVNews.ReloadNews(20);//default
            ZOVNews.ReloadComments();
            ZOVNews.ReloadAttachments();
            LightNewsContainer.NewsDataTable = ZOVNews.NewsDataTable.Copy();
            LightNewsContainer.CommentsDT = ZOVNews.CommentsDataTable.Copy();
            LightNewsContainer.AttachsDT = ZOVNews.AttachmentsDataTable.Copy();
            LightNewsContainer.CreateNews();
            LightNewsContainer.ScrollToTop();
            LightNewsContainer.Focus();

            bC = true;

        }

        private void LightNewsContainer_RemoveNewsClicked(object sender, int NewsID)
        {
            PhantomForm PhantomForm = new Infinium.PhantomForm();
            PhantomForm.Show();

            LightMessageBoxForm LightMessageBoxForm = new Infinium.LightMessageBoxForm(true, "Сообщение будет удалено безвозвратно.\nПродолжить?",
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
                    SplashWindow.CreateCoverSplash(LightNewsContainer.Top, LightNewsContainer.Left,
                    LightNewsContainer.Height, LightNewsContainer.Width);
                });
                T.Start();

                while (!SplashWindow.bSmallCreated) ;


                ZOVNews.RemoveNews(NewsID);
                ZOVNews.ReloadNews(LightNewsContainer.NewsCount);
                ZOVNews.ReloadComments();
                LightNewsContainer.NewsDataTable = ZOVNews.NewsDataTable.Copy();
                LightNewsContainer.CommentsDT = ZOVNews.CommentsDataTable.Copy();
                LightNewsContainer.CreateNews();
            }

            LightMessageBoxForm.Dispose();


            bC = true;

        }

        private void LightNewsContainer_CommentSendButtonClicked(object sender, string Text, int NewsID, int SenderID, bool bEdit, int NewsCommentID)
        {
            Thread T = new Thread(delegate ()
            {
                SplashWindow.CreateCoverSplash(LightNewsContainer.Top, LightNewsContainer.Left,
                                     LightNewsContainer.Height, LightNewsContainer.Width);
            });
            T.Start();

            while (!SplashWindow.bSmallCreated) ;

            if (bEdit)
                ZOVNews.EditComments(SenderID, NewsCommentID, Text);
            else
                ZOVNews.AddComments(SenderID, 0, NewsID, Text);

            ZOVNews.ReloadNews(LightNewsContainer.NewsCount);
            ZOVNews.ReloadComments();
            LightNewsContainer.NewsDataTable = ZOVNews.NewsDataTable.Copy();
            LightNewsContainer.CommentsDT = ZOVNews.CommentsDataTable.Copy();
            LightNewsContainer.CreateNews();
            LightNewsContainer.ScrollToTop();
            LightNewsContainer.ReloadNewsItem(NewsID, true);

            bC = true;
        }

        private void LightNewsContainer_RemoveCommentClicked(object sender, int NewsID, int NewsCommentID)
        {
            PhantomForm PhantomForm = new Infinium.PhantomForm();
            PhantomForm.Show();

            LightMessageBoxForm LightMessageBoxForm = new Infinium.LightMessageBoxForm(true, "Комментарий будет удален.\nПродолжить?",
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
                    SplashWindow.CreateCoverSplash(LightNewsContainer.Top, LightNewsContainer.Left,
                    LightNewsContainer.Height, LightNewsContainer.Width);
                });
                T.Start();

                while (!SplashWindow.bSmallCreated) ;

                ZOVNews.RemoveComment(NewsCommentID);

                ZOVNews.ReloadNews(LightNewsContainer.NewsCount);
                ZOVNews.ReloadComments();
                LightNewsContainer.NewsDataTable = ZOVNews.NewsDataTable.Copy();
                LightNewsContainer.CommentsDT = ZOVNews.CommentsDataTable.Copy();
                LightNewsContainer.CreateNews();
                LightNewsContainer.ReloadNewsItem(NewsID, true);

                bC = true;
            }
        }

        private void LightNewsContainer_EditNewsClicked(object sender, int NewsID)
        {
            PhantomForm PhantomForm = new PhantomForm();
            PhantomForm.Show();

            AddZOVNewsForm AddNewsForm = new AddZOVNewsForm(ref ZOVNews, ZOVNews.GetThisNewsSenderTypeID(NewsID),
                                                      ZOVNews.GetThisNewsHeaderText(NewsID),
                                                      ZOVNews.GetThisNewsBodyText(NewsID),
                                                      NewsID, ZOVNews.GetThisNewsDateTime(NewsID), ref TopForm);

            TopForm = AddNewsForm;

            AddNewsForm.ShowDialog();

            PhantomForm.Close();
            PhantomForm.Dispose();

            TopForm = null;

            if (AddNewsForm.Canceled)
                return;

            Thread T = new Thread(delegate ()
            {
                SplashWindow.CreateCoverSplash(LightNewsContainer.Top, LightNewsContainer.Left,
      LightNewsContainer.Height, LightNewsContainer.Width);
            });
            T.Start();

            while (!SplashWindow.bSmallCreated) ;


            ZOVNews.ReloadNews(LightNewsContainer.NewsCount);
            ZOVNews.ReloadComments();
            ZOVNews.ReloadAttachments();
            LightNewsContainer.NewsDataTable = ZOVNews.NewsDataTable.Copy();
            LightNewsContainer.CommentsDT = ZOVNews.CommentsDataTable.Copy();
            LightNewsContainer.AttachsDT = ZOVNews.AttachmentsDataTable.Copy();
            LightNewsContainer.CreateNews();

            bC = true;
        }

        private void LightNewsContainer_Refreshed(object sender, EventArgs e)
        {
            System.Threading.Thread.Sleep(100);

            while (SplashWindow.bSmallCreated)
                CoverWaitForm.CloseS = true;
        }

        private void LightNewsContainer_AttachClicked(object sender, int NewsAttachID)
        {
            PhantomForm PhantomForm = new Infinium.PhantomForm();
            PhantomForm.Show();

            ZOVNewsAttachDownloadForm AttachDownloadForm = new ZOVNewsAttachDownloadForm(NewsAttachID, ref ZOVNews.FM, ref ZOVNews);

            TopForm = AttachDownloadForm;

            AttachDownloadForm.ShowDialog();

            PhantomForm.Close();

            PhantomForm.Dispose();

            TopForm = null;

            AttachDownloadForm.Dispose();
        }

        private void LightNewsContainer_EditCommentClicked(object sender, int NewsID, int NewsCommentID)
        {

        }

        private void LightNewsContainer_NeedMoreNews(object sender, EventArgs e)
        {
            if (ZOVNews.IsMoreNews(LightNewsContainer.NewsCount))
                MoreNewsButton.Visible = true;
            else
                MoreNewsButton.Visible = false;
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            int O = LightNewsContainer.Offset;
            int SC = LightNewsContainer.ScrollContainer.Height;
            int H = LightNewsContainer.Height;
        }

        private void LightNewsContainer_VerticalScrollBar_ScrollPositionChanged(object sender, int tOffset)
        {

        }

        private void LightNewsContainer_NoNeedMoreNews(object sender, EventArgs e)
        {
            if (MoreNewsButton.Visible)
                MoreNewsButton.Visible = false;
        }

        private void MoreNewsButton_Click(object sender, EventArgs e)
        {
            int iNewsCount = LightNewsContainer.NewsCount;

            ZOVNews.FillMoreNews(LightNewsContainer.NewsCount + 20);
            LightNewsContainer.NewsDataTable = ZOVNews.NewsDataTable.Copy();
            LightNewsContainer.CreateNews();
            LightNewsContainer.ScrollToNews(iNewsCount);
            LightNewsContainer.Focus();
            MoreNewsButton.Visible = false;
            LightNewsContainer.bNeedMoreNews = false;
        }


        private void LightNewsContainer_CommentsClicked(object sender, EventArgs e)
        {
            //this.Refresh();
        }

        private void LightNewsContainer_LikeClicked(object sender, int NewsID)
        {
            ZOVNews.LikeNews(Security.CurrentUserID, NewsID);

            ZOVNews.FillLikes();
            LightNewsContainer.NewsLikesDT = ZOVNews.NewsLikesDataTable;
            LightNewsContainer.ReloadLikes(NewsID);
            LightNewsContainer.Focus();
        }

        private void LightNewsContainer_CommentLikeClicked(object sender, int NewsID, int NewsCommentID)
        {
            ZOVNews.LikeComments(Security.CurrentUserID, NewsID, NewsCommentID);

            ZOVNews.FillLikes();
            LightNewsContainer.CommentsLikesDT = ZOVNews.CommentsLikesDataTable;
            LightNewsContainer.ReloadLikes(NewsID);
            LightNewsContainer.Focus();
        }


        private void LightNewsContainer_CommentsQuoteLabelClicked(object sender)
        {
            InfiniumTips.ShowTip(this, 50, 85, "Комментарий скопирован в буфер обмена...", 1700);
        }

        private void LightNewsContainer_NewsQuoteLabelClicked(object sender)
        {
            InfiniumTips.ShowTip(this, 50, 85, "Текст скопирован в буфер обмена...", 1700);
        }

        private void UpdateNewsButton_Click(object sender, EventArgs e)
        {
            Thread T = new Thread(delegate ()
            {
                SplashWindow.CreateCoverSplash(LightNewsContainer.Top, LightNewsContainer.Left,
                                     LightNewsContainer.Height, LightNewsContainer.Width);
            });
            T.Start();

            while (!SplashWindow.bSmallCreated) ;

            ZOVNews.ReloadSubscribes();
            ZOVNews.ReloadNews(20);//default count
            ZOVNews.ReloadComments();
            ZOVNews.ReloadAttachments();
            ZOVNews.FillLikes();
            LightNewsContainer.NewsDataTable = ZOVNews.NewsDataTable.Copy();
            LightNewsContainer.CommentsDT = ZOVNews.CommentsDataTable.Copy();
            LightNewsContainer.AttachsDT = ZOVNews.AttachmentsDataTable.Copy();
            LightNewsContainer.NewsLikesDT = ZOVNews.NewsLikesDataTable;
            LightNewsContainer.CommentsLikesDT = ZOVNews.CommentsLikesDataTable;
            LightNewsContainer.CreateNews();
            LightNewsContainer.ScrollToTop();
            LightNewsContainer.Focus();

            ActiveNotifySystem.ClearSubscribesRecords(Security.CurrentUserID, this.Name);
            //ActiveNotifySystem.ClearModuleUpdates(LightTile.Name);
            //ActiveNotifySystem.ClearCurrentOpenModuleUpdates(this.Name);
            //LightNews.ClearCommentsSubs(Security.CurrentUserID);

            UpdateNewsButton.Visible = false;

            bC = true;
        }

        private void MinimizeButton_Click(object sender, EventArgs e)
        {
            LightNewsContainer.Visible = false;

            FormEvent = eHide;
            AnimateTimer.Enabled = true;
        }

    }
}
