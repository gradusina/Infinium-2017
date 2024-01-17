﻿
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Text;
using System.Windows.Forms;
using Infinium.Properties;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using Encoder = System.Drawing.Imaging.Encoder;
using Region = NPOI.HSSF.Util.Region;

namespace Infinium
{
    public class InfiniumStart
    {
        public DataTable MenuItemsDataTable;
        public DataTable ModulesDataTable;
        public DataTable FullModulesDataTable;

        public static void RotateFlipEXIF(ref Image img)
        {
            if (Array.IndexOf(img.PropertyIdList, 274) > -1)
            {
                int orientation = img.GetPropertyItem(274).Value[0];
                switch (orientation)
                {
                    case 1:
                        // No rotation required.
                        break;
                    case 2:
                        img.RotateFlip(RotateFlipType.RotateNoneFlipX);
                        break;
                    case 3:
                        img.RotateFlip(RotateFlipType.Rotate180FlipNone);
                        break;
                    case 4:
                        img.RotateFlip(RotateFlipType.Rotate180FlipX);
                        break;
                    case 5:
                        img.RotateFlip(RotateFlipType.Rotate90FlipX);
                        break;
                    case 6:
                        img.RotateFlip(RotateFlipType.Rotate90FlipNone);
                        break;
                    case 7:
                        img.RotateFlip(RotateFlipType.Rotate270FlipX);
                        break;
                    case 8:
                        img.RotateFlip(RotateFlipType.Rotate270FlipNone);
                        break;
                }
            }
        }

        public InfiniumStart()
        {
            Fill();
        }

        public void Fill()
        {
            MenuItemsDataTable = new DataTable();

            using (var DA = new SqlDataAdapter("SELECT * FROM MenuItems WHERE MenuItemID IN " +
                                               "(SELECT MenuItemID FROM Modules WHERE SecurityAccess = 0 OR (SecurityAccess = 1 AND ModuleID IN " +
                                               "(SELECT ModuleID FROM ModulesAccess WHERE UserID = " + Security.CurrentUserID + "))) OR MenuItemID = 0",
                                                         ConnectionStrings.UsersConnectionString))
            {
                DA.Fill(MenuItemsDataTable);
            }


            FullModulesDataTable = TablesManager.ModulesDataTable;

            //using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM Modules",
            //                                            ConnectionStrings.UsersConnectionString))
            //{
            //    DA.Fill(FullModulesDataTable);
            //}


            ModulesDataTable = new DataTable();

            using (var DA = new SqlDataAdapter("SELECT ModuleID, ModuleName, Description, FormName, SecurityAccess, PriceView, MenuItemID, ModuleDesigner FROM Modules WHERE SecurityAccess = 0 OR (SecurityAccess = 1 AND ModuleID IN " +
                                               "(SELECT ModuleID FROM ModulesAccess WHERE UserID = " + Security.CurrentUserID + "))",
                                                         ConnectionStrings.UsersConnectionString))
            {
                DA.Fill(ModulesDataTable);

                ModulesDataTable.Columns.Add(new DataColumn("Picture", Type.GetType("System.Byte[]")));

                foreach (DataRow Row in ModulesDataTable.Rows)
                {
                    Row["Picture"] = FullModulesDataTable.Select("ModuleID = " + Row["ModuleID"])[0]["Picture"];
                }
            }


            using (var DA = new SqlDataAdapter("SELECT ModuleID, ModuleName, Description, FormName, SecurityAccess, PriceView, MenuItemID, ModuleDesigner FROM Modules WHERE (ModuleID IN (SELECT ModuleID FROM MyModules WHERE UserID = " + Security.CurrentUserID +
                                               ") AND (SecurityAccess = 0 OR (SecurityAccess = 1 AND ModuleID IN " +
                                               "(SELECT ModuleID FROM ModulesAccess WHERE UserID = " + Security.CurrentUserID + "))))",
                                                         ConnectionStrings.UsersConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                    DT.Columns.Add(new DataColumn("Picture", Type.GetType("System.Byte[]")));

                    foreach (DataRow Row in DT.Rows)
                    {
                        Row["Picture"] = FullModulesDataTable.Select("ModuleID = " + Row["ModuleID"])[0]["Picture"];
                    }
                    foreach (DataRow Row in DT.Rows)
                    {
                        var NewRow = ModulesDataTable.NewRow();

                        foreach (DataColumn Column in DT.Columns)
                        {
                            if (Column.ColumnName != "MenuItemID")
                            {
                                NewRow[Column.ColumnName] = Row[Column.ColumnName];
                            }
                            else
                                NewRow["MenuItemID"] = 0;
                        }

                        ModulesDataTable.Rows.Add(NewRow);
                    }
                }
            }
        }

        public bool AddModuleToFavorite(string FormName)
        {
            using (var DA = new SqlDataAdapter("SELECT TOP 1 * FROM MyModules WHERE ModuleID = " +
                                               ModulesDataTable.Select("FormName = '" + FormName + "'")[0]["ModuleID"] + " AND UserID = " + Security.CurrentUserID, ConnectionStrings.UsersConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        if (DT.Rows.Count > 0)
                            return false;//exists

                        var NewRow = DT.NewRow();
                        NewRow["ModuleID"] = ModulesDataTable.Select("FormName = '" + FormName + "'")[0]["ModuleID"];
                        NewRow["UserID"] = Security.CurrentUserID;
                        DT.Rows.Add(NewRow);

                        var mNewRow = ModulesDataTable.NewRow();
                        var Row = ModulesDataTable.Select("FormName = '" + FormName + "'")[0];

                        foreach (DataColumn Column in ModulesDataTable.Columns)
                        {
                            if (Column.ColumnName != "MenuItemID")
                                mNewRow[Column.ColumnName] = Row[Column.ColumnName];
                            else
                                mNewRow["MenuItemID"] = 0;
                        }

                        ModulesDataTable.Rows.Add(mNewRow);

                        DA.Update(DT);

                        return true;
                    }
                }
            }

        }

        public void RemoveModuleFromFavorite(string FormName)
        {
            using (var DA = new SqlDataAdapter("DELETE FROM MyModules WHERE ModuleID = " +
                                               ModulesDataTable.Select("FormName = '" + FormName + "'")[0]["ModuleID"] + " AND UserID = " + Security.CurrentUserID, ConnectionStrings.UsersConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    ModulesDataTable.Select("FormName = '" + FormName + "' AND MenuItemID = 0")[0].Delete();
                }
            }

        }
    }





    public class LightNews
    {
        public DataTable UsersDataTable;
        public DataTable NewsDataTable;
        public DataTable NewsLikesDataTable;
        public DataTable DepartmentsDataTable;
        public DataTable AttachmentsDataTable;
        public DataTable CommentsDataTable;
        public DataTable CommentsLikesDataTable;
        public DataTable CommentsSubsRecordsDataTable;
        public DataTable NewsSubsRecordsDataTable;
        public DataTable CurrentCommentsDataTable;



        public FileManager FM = new FileManager();

        public BindingSource DepartmentsBindingSource;

        public LightNews()
        {
            NewsDataTable = new DataTable();
            NewsLikesDataTable = new DataTable();
            CommentsLikesDataTable = new DataTable();
            UsersDataTable = new DataTable();
            DepartmentsDataTable = new DataTable();

            Fill();
            FillNews();
            FillLikes();
        }

        public void Fill()
        {
            UsersDataTable.Clear();

            DepartmentsBindingSource = new BindingSource();

            DepartmentsDataTable = new DataView(TablesManager.DepartmentsDataTable).ToTable(true, "DepartmentID", "DepartmentName", "Photo");

            DepartmentsBindingSource.DataSource = DepartmentsDataTable;

            UsersDataTable = new DataView(TablesManager.UsersDataTable).ToTable(true, "UserID", "Name", "DepartmentID", "Photo");

            if (UsersDataTable.Columns["SenderTypeID"] == null)
                UsersDataTable.Columns.Add(new DataColumn("SenderTypeID", Type.GetType("System.Int32")));

            foreach (DataRow Row in UsersDataTable.Rows)
                Row["SenderTypeID"] = 0;

            foreach (DataRow Row in DepartmentsDataTable.Rows)
            {
                var NewRow = UsersDataTable.NewRow();
                NewRow["UserID"] = Row["DepartmentID"];
                NewRow["Name"] = Row["DepartmentName"];
                NewRow["SenderTypeID"] = 1;
                NewRow["Photo"] = Row["Photo"];
                UsersDataTable.Rows.Add(NewRow);
            }

            {
                var NewRow = DepartmentsDataTable.NewRow();
                NewRow["DepartmentID"] = -1;
                NewRow["DepartmentName"] = "Все отделы";
                DepartmentsDataTable.Rows.Add(NewRow);
            }

            DepartmentsBindingSource.Sort = "DepartmentID ASC";

            AttachmentsDataTable = new DataTable();

            using (var DA = new SqlDataAdapter("SELECT NewsAttachID, NewsID, FileName, FileSize FROM NewsAttachs", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(AttachmentsDataTable);
            }

            CommentsDataTable = new DataTable();

            using (var DA = new SqlDataAdapter("SELECT * FROM NewsComments", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(CommentsDataTable);
            }

            CommentsSubsRecordsDataTable = new DataTable();

            using (var DA = new SqlDataAdapter("SELECT SubscribesRecords.*, NewsComments.NewsID FROM SubscribesRecords INNER JOIN NewsComments ON NewsComments.NewsCommentID = SubscribesRecords.TableItemID WHERE SubscribesItemID = 2 AND UserTypeID = 0 AND SubscribesRecords.UserID = " + Security.CurrentUserID, ConnectionStrings.LightConnectionString))
            {
                DA.Fill(CommentsSubsRecordsDataTable);
            }

            NewsSubsRecordsDataTable = new DataTable();

            using (var sDA = new SqlDataAdapter("SELECT * FROM SubscribesRecords WHERE SubscribesItemID = 1 " +
                                                " AND UserTypeID = 0 AND UserID = " + Security.CurrentUserID, ConnectionStrings.LightConnectionString))
            {
                sDA.Fill(NewsSubsRecordsDataTable);
            }
        }

        public void ReloadNews(int Count)
        {
            NewsDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT TOP " + Count + " * FROM News WHERE Pending <> 1 ORDER BY LastCommentDateTime DESC", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(NewsDataTable);
            }

            //subscribes///////////////////////
            if (NewsDataTable.Columns["New"] == null)
                NewsDataTable.Columns.Add(new DataColumn("New", Type.GetType("System.Int32")));

            if (NewsDataTable.Columns["NewComments"] == null)
                NewsDataTable.Columns.Add(new DataColumn("NewComments", Type.GetType("System.Int32")));

            foreach (DataRow Row in NewsSubsRecordsDataTable.Rows)
            {
                if (NewsDataTable.Select("NewsID = " + Row["TableItemID"]).Count() > 0)
                {
                    NewsDataTable.Select("NewsID = " + Row["TableItemID"])[0]["New"] = 1;
                }
            }

            foreach (DataRow Row in CommentsSubsRecordsDataTable.Rows)
            {
                if (NewsDataTable.Select("NewsID = " + Row["NewsID"]).Count() > 0)
                {
                    NewsDataTable.Select("NewsID = " + Row["NewsID"])[0]["NewComments"] =
                            CommentsSubsRecordsDataTable.Select("NewsID = " + Row["NewsID"]).Count();
                }
            }
            //////////////////////////////////
        }

        public void ReloadComments()
        {
            CommentsDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT * FROM NewsComments", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(CommentsDataTable);
            }
        }

        public void ReloadSubscribes()
        {
            CommentsSubsRecordsDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT SubscribesRecords.*, NewsComments.NewsID FROM SubscribesRecords INNER JOIN NewsComments ON NewsComments.NewsCommentID = SubscribesRecords.TableItemID WHERE SubscribesItemID = 2 AND UserTypeID = 0 AND SubscribesRecords.UserID = " + Security.CurrentUserID, ConnectionStrings.LightConnectionString))
            {
                DA.Fill(CommentsSubsRecordsDataTable);
            }

            NewsSubsRecordsDataTable.Clear();

            using (var sDA = new SqlDataAdapter("SELECT * FROM SubscribesRecords WHERE SubscribesItemID = 1 " +
                                                " AND UserTypeID = 0 AND UserID = " + Security.CurrentUserID, ConnectionStrings.LightConnectionString))
            {
                sDA.Fill(NewsSubsRecordsDataTable);
            }
        }

        public void FillNews()
        {
            NewsDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT TOP 20 * FROM News WHERE Pending <> 1 ORDER BY LastCommentDateTime DESC", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(NewsDataTable);
            }

            if (NewsDataTable.Columns["New"] == null)
                NewsDataTable.Columns.Add(new DataColumn("New", Type.GetType("System.Int32")));

            if (NewsDataTable.Columns["NewComments"] == null)
                NewsDataTable.Columns.Add(new DataColumn("NewComments", Type.GetType("System.Int32")));

            foreach (DataRow Row in NewsSubsRecordsDataTable.Rows)
            {
                if (NewsDataTable.Select("NewsID = " + Row["TableItemID"]).Count() > 0)
                {
                    NewsDataTable.Select("NewsID = " + Row["TableItemID"])[0]["New"] = 1;
                }
            }

            foreach (DataRow Row in CommentsSubsRecordsDataTable.Rows)
            {
                if (NewsDataTable.Select("NewsID = " + Row["NewsID"]).Count() > 0)
                {
                    NewsDataTable.Select("NewsID = " + Row["NewsID"])[0]["NewComments"] =
                            CommentsSubsRecordsDataTable.Select("NewsID = " + Row["NewsID"]).Count();
                }
            }
        }

        public void FillLikes()
        {
            NewsLikesDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT * FROM NewsLikes", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(NewsLikesDataTable);
            }


            CommentsLikesDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT * FROM NewsCommentsLikes", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(CommentsLikesDataTable);
            }
        }

        public void FillMoreNews(int Count)
        {
            NewsDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT TOP " + Count + " * FROM News WHERE Pending <> 1 ORDER BY LastCommentDateTime DESC", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(NewsDataTable);
            }
        }

        public bool IsMoreNews(int Count)
        {
            using (var DA = new SqlDataAdapter("SELECT Count(NewsID) FROM News", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    return Convert.ToInt32(DT.Rows[0][0]) > Count;
                }
            }
        }

        public DateTime AddNews(int SenderID, int SenderTypeID, string HeaderText, string BodyText, int NewsCategoryID)
        {
            DateTime DateTime;

            using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM News", ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        var NewRow = DT.NewRow();
                        NewRow["SenderID"] = SenderID;
                        NewRow["SenderTypeID"] = SenderTypeID;
                        NewRow["HeaderText"] = HeaderText;
                        NewRow["BodyText"] = BodyText;
                        NewRow["Pending"] = true;

                        DateTime = Security.GetCurrentDate();

                        NewRow["DateTime"] = DateTime;
                        NewRow["LastCommentDateTime"] = DateTime;
                        DT.Rows.Add(NewRow);

                        DA.Update(DT);

                        FillNews();
                    }
                }
            }

            return DateTime;
        }

        public void AddSubscribeForNews(DateTime DateTime)
        {
            using (var sDA = new SqlDataAdapter("SELECT * FROM News WHERE DateTime = '" + DateTime.ToString("yyyy-MM-dd HH:mm:ss.fff") + "'",
                                                                       ConnectionStrings.LightConnectionString))
            {
                using (var sDT = new DataTable())
                {
                    sDA.Fill(sDT);

                    ActiveNotifySystem.CreateSubscribeRecord(1, Convert.ToInt32(sDT.Rows[0][0]), Security.CurrentUserID);
                }
            }
        }

        public void ClearPending(DateTime DateTime)
        {
            using (var sDA = new SqlDataAdapter("SELECT TOP 1 * FROM News WHERE DateTime = '" + DateTime.ToString("yyyy-MM-dd HH:mm:ss.fff") + "'",
                                                                      ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(sDA))
                {
                    using (var sDT = new DataTable())
                    {
                        sDA.Fill(sDT);

                        sDT.Rows[0]["Pending"] = false;

                        sDA.Update(sDT);
                    }

                }

            }
        }

        public void AddComments(int UserID, int NewsID, string Text)
        {
            DateTime Date;

            using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM NewsComments", ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        var NewRow = DT.NewRow();
                        NewRow["NewsID"] = NewsID;
                        NewRow["NewsComment"] = Text;
                        NewRow["UserID"] = UserID;

                        Date = Security.GetCurrentDate();

                        NewRow["DateTime"] = Date;
                        DT.Rows.Add(NewRow);

                        DA.Update(DT);
                    }
                }
            }

            using (var DA = new SqlDataAdapter("SELECT NewsID, LastCommentDateTime FROM News WHERE NewsID =" + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        DT.Rows[0]["LastCommentDateTime"] = Date;

                        DA.Update(DT);
                    }
                }
            }

            AddNewsCommentsSubs(NewsID, UserID, Date);
        }


        public void LikeNews(int UserID, int NewsID)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM NewsLikes WHERE NewsID = " + NewsID + " AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        if (DT.Rows.Count > 0)//i like
                        {
                            using (var dDA = new SqlDataAdapter("DELETE FROM NewsLikes WHERE NewsID = " + NewsID + " AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
                            {
                                using (var dDT = new DataTable())
                                {
                                    dDA.Fill(dDT);
                                }
                            }

                            return;
                        }

                        var NewRow = DT.NewRow();
                        NewRow["NewsID"] = NewsID;
                        NewRow["UserID"] = UserID;
                        DT.Rows.Add(NewRow);

                        DA.Update(DT);
                    }
                }
            }
        }

        public void LikeComments(int UserID, int NewsID, int NewsCommentID)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM NewsCommentsLikes WHERE NewsCommentID = " + NewsCommentID + " AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        if (DT.Rows.Count > 0)
                        {
                            using (var dDA = new SqlDataAdapter("DELETE FROM NewsCommentsLikes WHERE NewsCommentID = " + NewsCommentID + " AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
                            {
                                using (var dDT = new DataTable())
                                {
                                    dDA.Fill(dDT);
                                }
                            }

                            return;
                        }

                        var NewRow = DT.NewRow();
                        NewRow["NewsID"] = NewsID;
                        NewRow["NewsCommentID"] = NewsCommentID;
                        NewRow["UserID"] = UserID;
                        DT.Rows.Add(NewRow);

                        DA.Update(DT);
                    }
                }
            }
        }


        public void EditComments(int UserID, int NewsCommentID, string Text)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM NewsComments WHERE NewsCommentID = " + NewsCommentID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        DT.Rows[0]["NewsComment"] = Text;

                        DA.Update(DT);
                    }
                }
            }
        }


        public bool Attach(DataTable AttachmentsDataTable, DateTime NewsDateTime, ref int CurrentUploadedFile)
        {
            if (AttachmentsDataTable.Rows.Count == 0)
                return true;

            var Ok = true;

            int NewsID;

            using (var DA = new SqlDataAdapter("SELECT NewsID FROM News WHERE DateTime = '" + NewsDateTime.ToString("yyyy-MM-dd HH:mm:ss.fff") + "'", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    NewsID = Convert.ToInt32(DT.Rows[0]["NewsID"]);
                }
            }

            using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM NewsAttachs", ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        foreach (DataRow Row in AttachmentsDataTable.Rows)
                        {
                            FileInfo fi;

                            try
                            {
                                fi = new FileInfo(Row["Path"].ToString());
                                //WriteLog("UploadFile fileinfosize: " + Row["Path"].ToString(), Security.CurrentUserID, -1);
                            }
                            catch
                            {
                                Ok = false;
                                continue;
                            }

                            var NewRow = DT.NewRow();
                            NewRow["NewsID"] = NewsID;
                            NewRow["FileName"] = Row["FileName"];
                            NewRow["FileSize"] = fi.Length;
                            DT.Rows.Add(NewRow);
                        }

                        DA.Update(DT);
                    }
                }
            }

            //write to ftp
            using (var DA = new SqlDataAdapter("SELECT * FROM NewsAttachs WHERE NewsID = " + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    foreach (DataRow Row in AttachmentsDataTable.Rows)
                    {
                        CurrentUploadedFile++;

                        try
                        {
                            if (FM.UploadFile(Row["Path"].ToString(), Configs.DocumentsZOVTPSPath +
                                          FileManager.GetPath("LightNews") + "/" + Row["FileName"], Configs.FTPType) == false)
                            {
                                break;
                            }
                        }
                        catch
                        {
                            Ok = false;
                            break;
                        }
                    }

                    DA.Update(DT);
                }
            }

            ReloadAttachments();

            return Ok;
        }

        public bool EditAttachments(int NewsID, DataTable AttachmentsDataTable, ref int CurrentUploadedFile, ref int TotalFilesCount)
        {
            var Ok = false;

            using (var DA = new SqlDataAdapter("SELECT NewsAttachID, NewsID, FileName FROM NewsAttachs WHERE NewsID = " + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        foreach (DataRow Row in DT.Rows)
                        {
                            if (AttachmentsDataTable.Select("FileName = '" + Row["FileName"] + "'").Count() == 0)
                            {
                                FM.DeleteFile(Configs.DocumentsZOVTPSPath + FileManager.GetPath("LightNews") + "/" + Row["FileName"], Configs.FTPType);
                                Row.Delete();
                            }
                        }

                        DA.Update(DT);
                    }
                }
            }

            using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM NewsAttachs WHERE NewsID = " + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);
                        Ok = true;

                        //add new
                        foreach (DataRow Row in AttachmentsDataTable.Rows)
                        {
                            if (Row["Path"].ToString() == "server")
                                continue;

                            TotalFilesCount++;

                            FileInfo fi;

                            try
                            {
                                fi = new FileInfo(Row["Path"].ToString());

                            }
                            catch
                            {
                                Ok = false;
                                continue;
                            }

                            var NewRow = DT.NewRow();
                            NewRow["NewsID"] = NewsID;
                            NewRow["FileName"] = Row["FileName"];
                            NewRow["FileSize"] = fi.Length;
                            DT.Rows.Add(NewRow);
                        }

                        DA.Update(DT);
                    }
                }
            }


            //write to ftp
            using (var DA = new SqlDataAdapter("SELECT * FROM NewsAttachs WHERE NewsID = " + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    foreach (DataRow Row in AttachmentsDataTable.Rows)
                    {
                        if (Row["Path"].ToString() == "server")
                            continue;

                        CurrentUploadedFile++;

                        try
                        {
                            if (FM.UploadFile(Row["Path"].ToString(), Configs.DocumentsZOVTPSPath +
                                         FileManager.GetPath("LightNews") + "/" + Row["FileName"], Configs.FTPType) == false)
                                break;
                        }
                        catch
                        {
                            Ok = false;
                            break;
                        }
                    }

                    DA.Update(DT);
                }
            }


            ReloadAttachments();

            return Ok;
        }

        public void ReloadAttachments()
        {
            AttachmentsDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT NewsAttachID, NewsID, FileName, FileSize FROM NewsAttachs", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(AttachmentsDataTable);
            }
        }

        public void EditNews(int NewsID, int SenderID, int SenderTypeID, string HeaderText, string BodyText, int NewsCategoryID, DateTime DateTime)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM News WHERE NewsID = " + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        if (DA.Fill(DT) == 0)
                            return;

                        DT.Rows[0]["SenderID"] = SenderID;
                        DT.Rows[0]["SenderTypeID"] = SenderTypeID;
                        DT.Rows[0]["HeaderText"] = HeaderText;
                        DT.Rows[0]["BodyText"] = BodyText;
                        DT.Rows[0]["DateTime"] = DateTime;

                        DA.Update(DT);

                        FillNews();
                    }
                }
            }
        }

        public static string TempPath()
        {
            return Environment.GetEnvironmentVariable("TEMP");
        }

        public int GetNewsUpdatesCount()
        {
            using (var DA = new SqlDataAdapter("SELECT Count (SubscribesRecordID) FROM SubscribesRecords WHERE SubscribesItemID = 1 OR SubscribesItemID = 2 AND UserTypeID = 0 AND UserID = " + Security.CurrentUserID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    if (DA.Fill(DT) > 0)
                        return Convert.ToInt32(DT.Rows[0][0]);
                }
            }

            return 0;
        }

        public int GetNewsIDByDateTime(DateTime DateTime)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM News WHERE DateTime = '" + DateTime.ToString("yyyy-MM-dd HH:mm:ss.fff") + "'",
                                                                         ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    return Convert.ToInt32(DT.Rows[0]["NewsID"]);
                }
            }
        }

        public void RemoveNews(int NewsID)
        {
            ActiveNotifySystem.DeleteSubscribesRecord(1);

            using (var DA = new SqlDataAdapter("DELETE FROM NewsLikes WHERE NewsID = " + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }

            using (var DA = new SqlDataAdapter("DELETE FROM News WHERE NewsID = " + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);
                    }
                }
            }

            RemoveComments(NewsID);

            RemoveAttachments(NewsID);

            FillNews();

            ReloadAttachments();
        }

        public void RemoveComments(int NewsID)
        {
            DeleteCommentsSubs(NewsID);

            using (var DA = new SqlDataAdapter("DELETE FROM NewsCommentsLikes WHERE NewsID = " + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }

            using (var DA = new SqlDataAdapter("DELETE FROM NewsComments WHERE NewsID = " + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);
                    }
                }
            }


        }

        public void RemoveComment(int NewsCommentID)
        {
            using (var DA = new SqlDataAdapter("DELETE FROM NewsCommentsLikes WHERE NewsCommentID = " + NewsCommentID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }

            using (var DA = new SqlDataAdapter("DELETE FROM NewsComments WHERE NewsCommentID = " + NewsCommentID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }

            DeleteCommentsSub(NewsCommentID);
        }

        public void RemoveAttachments(int NewsID)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM NewsAttachs WHERE NewsID = " + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    try
                    {
                        foreach (DataRow Row in DT.Rows)
                        {
                            FM.DeleteFile(Configs.DocumentsZOVTPSPath + FileManager.GetPath("LightNews") + "/" + Row["FileName"], Configs.FTPType);
                        }
                    }
                    catch
                    {
                        using (var fDA = new SqlDataAdapter("DELETE FROM NewsAttachs WHERE NewsID = " + NewsID, ConnectionStrings.LightConnectionString))
                        {
                            using (var CB = new SqlCommandBuilder(fDA))
                            {
                                using (var fDT = new DataTable())
                                {
                                    fDA.Fill(fDT);
                                }
                            }
                        }

                        ReloadAttachments();

                        return;
                    }
                }
            }



            using (var DA = new SqlDataAdapter("DELETE FROM NewsAttachs WHERE NewsID = " + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);
                    }
                }
            }

            ReloadAttachments();
        }

        public void RemoveCurrentAttachments(int NewsID, DataTable AttachmentsDT)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM NewsAttachs WHERE NewsID = " + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        foreach (DataRow Row in AttachmentsDT.Rows)
                        {
                            if (Row["Path"].ToString() == "server")
                                continue;

                            FM.DeleteFile(Configs.DocumentsZOVTPSPath + FileManager.GetPath("LightNews") + "/" + Row["FileName"], Configs.FTPType);

                            DT.Select("FileName = '" + Row["FileName"] + "'")[0].Delete();
                        }

                        DA.Update(DT);
                    }
                }
            }

            ReloadAttachments();
        }

        public DataTable GetAttachments(int NewsID)
        {
            using (var DA = new SqlDataAdapter("SELECT NewsAttachID, FileName FROM NewsAttachs WHERE NewsID = " + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    return DT;
                }
            }
        }

        public int GetThisNewsSenderTypeID(int NewsID)
        {
            return Convert.ToInt32(NewsDataTable.Select("NewsID = " + NewsID)[0]["SenderTypeID"]);
        }

        public string GetThisNewsBodyText(int NewsID)
        {
            return NewsDataTable.Select("NewsID = " + NewsID)[0]["BodyText"].ToString();
        }

        public string GetThisNewsHeaderText(int NewsID)
        {
            return NewsDataTable.Select("NewsID = " + NewsID)[0]["HeaderText"].ToString();
        }

        public DateTime GetThisNewsDateTime(int NewsID)
        {
            return Convert.ToDateTime(NewsDataTable.Select("NewsID = " + NewsID)[0]["DateTime"]);
        }

        public int ShowAttachDownloadMenu(string FileName)
        {
            //0 cancel, 1 open, 2 save


            return AttachDownloadForm.Result;
        }

        public string GetAttachmentName(int NewsAttachID)
        {
            return AttachmentsDataTable.Select("NewsAttachID = " + NewsAttachID)[0]["FileName"].ToString();
        }

        public string SaveFile(int NewsAttachID)//temp folder
        {
            var tempFolder = Environment.GetEnvironmentVariable("TEMP");

            var FileName = "";

            using (var DA = new SqlDataAdapter("SELECT FileName, FileSize FROM NewsAttachs WHERE NewsAttachID = " + NewsAttachID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    FileName = DT.Rows[0]["FileName"].ToString();

                    try
                    {
                        FM.DownloadFile(Configs.DocumentsZOVTPSPath + FileManager.GetPath("LightNews") + "/" + FileName,
                                            tempFolder + "\\" + FileName, Convert.ToInt64(DT.Rows[0]["FileSize"]), Configs.FTPType);
                    }
                    catch
                    {
                        return null;
                    }

                    //byte[] b = (byte[])DT.Rows[0]["FileBytes"];
                    //MemoryStream ms = new MemoryStream(b);

                    //FileName = DT.Rows[0]["FileName"].ToString();

                    //FileStream s2;

                    //try
                    //{
                    //    s2 = new FileStream(tempFolder + "\\" + FileName, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite);
                    //}
                    //catch 
                    //{
                    //    FileName = GetNewFileName(tempFolder, DT.Rows[0]["FileName"].ToString());
                    //    s2 = new FileStream(tempFolder + "\\" + FileName, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite);
                    //}

                    //s2.Write(ms.ToArray(), 0, ms.Capacity);

                    //s2.Close();
                }
            }

            return tempFolder + "\\" + FileName;
        }

        public void SaveFile(int NewsAttachID, string sDestFileName)
        {
            using (var DA = new SqlDataAdapter("SELECT FileName, FileSize FROM NewsAttachs WHERE NewsAttachID = " + NewsAttachID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    var FileName = DT.Rows[0]["FileName"].ToString();

                    try
                    {
                        FM.DownloadFile(Configs.DocumentsZOVTPSPath + FileManager.GetPath("LightNews") + "/" + FileName,
                                           sDestFileName, Convert.ToInt64(DT.Rows[0]["FileSize"]), Configs.FTPType);
                    }
                    catch
                    {
                    }
                }
            }
        }

        private string SetNumber(string FileName, int Number)
        {
            var Ext = "";
            var Name = "";

            for (var i = FileName.Length - 1; i > 0; i--)
            {
                if (FileName[i] == '.')
                {
                    Ext = FileName.Substring(i + 1, FileName.Length - i - 1);
                    Name = FileName.Substring(0, i) + " (" + Number + ")." + Ext;
                    break;
                }
            }

            return Name;
        }

        private string GetNewFileName(string path, string FileName)
        {
            var fileInfo = new FileInfo(path + "\\" + FileName);

            if (!fileInfo.Exists)
                return FileName;

            var Ok = false;
            var n = 1;

            while (!Ok)
            {
                fileInfo = new FileInfo(path + "\\" + SetNumber(FileName, n));

                if (!fileInfo.Exists)
                    return SetNumber(FileName, n);

                n++;
            }

            return "";
        }

        public void DeleteCommentsSub(int NewsCommentID)
        {
            using (var DA = new SqlDataAdapter("DELETE FROM SubscribesRecords WHERE SubscribesItemID = 2 AND TableItemID = " + NewsCommentID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }
        }

        public void DeleteCommentsSubs(int NewsID)
        {
            using (var DA = new SqlDataAdapter("DELETE FROM SubscribesRecords WHERE SubscribesItemID = 2 AND TableItemID IN (SELECT NewsCommentID FROM NewsComments WHERE NewsID = " + NewsID + ")", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }
        }

        public void AddNewsCommentsSubs(int NewsID, int UserID, DateTime DateTime)
        {
            using (var DA = new SqlDataAdapter("SELECT NewsCommentID FROM NewsComments WHERE DateTime = '" + DateTime.ToString("yyyy-MM-dd HH:mm:ss.fff") + "'", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    ActiveNotifySystem.CreateSubscribeRecord(2, Convert.ToInt32(DT.Rows[0][0]), Security.CurrentUserID);
                }
            }
        }
    }





    public class CoderBlog
    {
        public DataTable UsersDataTable;
        public DataTable BlogNewsDataTable;
        public DataTable BlogNewsLikesDataTable;
        public DataTable DepartmentsDataTable;
        public DataTable AttachmentsDataTable;
        public DataTable CommentsDataTable;
        public DataTable CommentsLikesDataTable;
        public DataTable CommentsSubsRecordsDataTable;
        public DataTable NewsSubsRecordsDataTable;
        public DataTable CurrentCommentsDataTable;

        public FileManager FM = new FileManager();

        public BindingSource DepartmentsBindingSource;

        public CoderBlog()
        {
            BlogNewsDataTable = new DataTable();
            BlogNewsLikesDataTable = new DataTable();
            CommentsLikesDataTable = new DataTable();
            UsersDataTable = new DataTable();
            DepartmentsDataTable = new DataTable();

            Fill();
            FillNews();
            FillLikes();
        }

        public void Fill()
        {
            UsersDataTable.Clear();

            DepartmentsBindingSource = new BindingSource();

            DepartmentsDataTable = new DataView(TablesManager.DepartmentsDataTable).ToTable(true, "DepartmentID", "DepartmentName", "Photo");

            DepartmentsBindingSource.DataSource = DepartmentsDataTable;

            UsersDataTable = new DataView(TablesManager.UsersDataTable).ToTable(true, "UserID", "Name", "DepartmentID", "Photo");

            if (UsersDataTable.Columns["SenderTypeID"] == null)
                UsersDataTable.Columns.Add(new DataColumn("SenderTypeID", Type.GetType("System.Int32")));

            foreach (DataRow Row in UsersDataTable.Rows)
                Row["SenderTypeID"] = 0;

            foreach (DataRow Row in DepartmentsDataTable.Rows)
            {
                var NewRow = UsersDataTable.NewRow();
                NewRow["UserID"] = Row["DepartmentID"];
                NewRow["Name"] = Row["DepartmentName"];
                NewRow["SenderTypeID"] = 1;
                NewRow["Photo"] = Row["Photo"];
                UsersDataTable.Rows.Add(NewRow);
            }

            {
                var NewRow = DepartmentsDataTable.NewRow();
                NewRow["DepartmentID"] = -1;
                NewRow["DepartmentName"] = "Все отделы";
                DepartmentsDataTable.Rows.Add(NewRow);
            }

            DepartmentsBindingSource.Sort = "DepartmentID ASC";

            AttachmentsDataTable = new DataTable();

            using (var DA = new SqlDataAdapter("SELECT NewsAttachID, NewsID, FileName, FileSize FROM BlogAttachs", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(AttachmentsDataTable);
            }

            CommentsDataTable = new DataTable();

            using (var DA = new SqlDataAdapter("SELECT * FROM BlogComments", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(CommentsDataTable);
            }

            CommentsSubsRecordsDataTable = new DataTable();

            using (var DA = new SqlDataAdapter("SELECT SubscribesRecords.*, BlogComments.NewsID FROM SubscribesRecords INNER JOIN BlogComments ON BlogComments.NewsCommentID = SubscribesRecords.TableItemID WHERE SubscribesItemID = 4 AND UserTypeID = 0 AND SubscribesRecords.UserID = " + Security.CurrentUserID, ConnectionStrings.LightConnectionString))
            {
                DA.Fill(CommentsSubsRecordsDataTable);
            }

            NewsSubsRecordsDataTable = new DataTable();

            using (var sDA = new SqlDataAdapter("SELECT * FROM SubscribesRecords WHERE SubscribesItemID = 3 " +
                                                " AND UserTypeID = 0 AND UserID = " + Security.CurrentUserID, ConnectionStrings.LightConnectionString))
            {
                sDA.Fill(NewsSubsRecordsDataTable);
            }
        }

        public void ReloadNews(int Count)
        {
            BlogNewsDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT TOP " + Count + " * FROM Blog WHERE Pending <> 1  ORDER BY LastCommentDateTime DESC", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(BlogNewsDataTable);
            }

            //subscribes///////////////////////
            if (BlogNewsDataTable.Columns["New"] == null)
                BlogNewsDataTable.Columns.Add(new DataColumn("New", Type.GetType("System.Int32")));

            if (BlogNewsDataTable.Columns["NewComments"] == null)
                BlogNewsDataTable.Columns.Add(new DataColumn("NewComments", Type.GetType("System.Int32")));

            foreach (DataRow Row in NewsSubsRecordsDataTable.Rows)
            {
                if (BlogNewsDataTable.Select("NewsID = " + Row["TableItemID"]).Count() > 0)
                {
                    BlogNewsDataTable.Select("NewsID = " + Row["TableItemID"])[0]["New"] = 1;
                }
            }

            foreach (DataRow Row in CommentsSubsRecordsDataTable.Rows)
            {
                if (BlogNewsDataTable.Select("NewsID = " + Row["NewsID"]).Count() > 0)
                {
                    BlogNewsDataTable.Select("NewsID = " + Row["NewsID"])[0]["NewComments"] =
                            CommentsSubsRecordsDataTable.Select("NewsID = " + Row["NewsID"]).Count();
                }
            }
            //////////////////////////////////
        }

        public void ReloadComments()
        {
            CommentsDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT * FROM BlogComments", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(CommentsDataTable);
            }
        }

        public void ReloadSubscribes()
        {
            CommentsSubsRecordsDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT SubscribesRecords.*, NewsComments.NewsID FROM SubscribesRecords INNER JOIN NewsComments ON NewsComments.NewsCommentID = SubscribesRecords.TableItemID WHERE SubscribesItemID = 4 AND UserTypeID = 0 AND SubscribesRecords.UserID = " + Security.CurrentUserID, ConnectionStrings.LightConnectionString))
            {
                DA.Fill(CommentsSubsRecordsDataTable);
            }

            NewsSubsRecordsDataTable.Clear();

            using (var sDA = new SqlDataAdapter("SELECT * FROM SubscribesRecords WHERE SubscribesItemID = 3 " +
                                                " AND UserTypeID = 0 AND UserID = " + Security.CurrentUserID, ConnectionStrings.LightConnectionString))
            {
                sDA.Fill(NewsSubsRecordsDataTable);
            }
        }

        //public void RefillComments()
        //{
        //    CommentsDataTable.Clear();

        //    using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM NewsComments", ConnectionStrings.LightConnectionString))
        //    {
        //        DA.Fill(CommentsDataTable);
        //    }
        //}

        //public void RefillCurrentComments()
        //{
        //    CommentsSubsRecordsDataTable.Clear();

        //    using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM NewsSubsRecords WHERE UserID = " + Security.CurrentUserID, ConnectionStrings.LightConnectionString))
        //    {
        //        DA.Fill(CommentsSubsRecordsDataTable);
        //    }
        //}

        //public void RefillSubs()
        //{
        //    SubsRecordsDataTable.Clear();

        //    using (SqlDataAdapter DA = new SqlDataAdapter("SELECT ModuleID FROM Modules WHERE ModuleButtonName = '" + "LightNewsButton" + "'", ConnectionStrings.UsersConnectionString))
        //    {
        //        using (DataTable DT = new DataTable())
        //        {
        //            DA.Fill(DT);

        //            using (SqlDataAdapter sDA = new SqlDataAdapter("SELECT * FROM SubscribesRecords WHERE ModuleID = " + DT.Rows[0]["ModuleID"] +
        //                                                           " AND UserID = " + Security.CurrentUserID, ConnectionStrings.LightConnectionString))
        //            {
        //                sDA.Fill(SubsRecordsDataTable);
        //            }
        //        }
        //    }
        //}

        public void FillNews()
        {
            BlogNewsDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT TOP 20 * FROM Blog  WHERE Pending <> 1 ORDER BY LastCommentDateTime DESC", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(BlogNewsDataTable);
            }

            if (BlogNewsDataTable.Columns["New"] == null)
                BlogNewsDataTable.Columns.Add(new DataColumn("New", Type.GetType("System.Int32")));

            if (BlogNewsDataTable.Columns["NewComments"] == null)
                BlogNewsDataTable.Columns.Add(new DataColumn("NewComments", Type.GetType("System.Int32")));

            foreach (DataRow Row in NewsSubsRecordsDataTable.Rows)
            {
                if (BlogNewsDataTable.Select("NewsID = " + Row["TableItemID"]).Count() > 0)
                {
                    BlogNewsDataTable.Select("NewsID = " + Row["TableItemID"])[0]["New"] = 1;
                }
            }

            foreach (DataRow Row in CommentsSubsRecordsDataTable.Rows)
            {
                if (BlogNewsDataTable.Select("NewsID = " + Row["NewsID"]).Count() > 0)
                {
                    BlogNewsDataTable.Select("NewsID = " + Row["NewsID"])[0]["NewComments"] =
                            CommentsSubsRecordsDataTable.Select("NewsID = " + Row["NewsID"]).Count();
                }
            }

            //if (NewsDataTable.Columns["CanEdit"] == null)
            //    NewsDataTable.Columns.Add(new DataColumn("CanEdit", Type.GetType("System.Boolean")));

            //string CurrentUserSender = UsersDataTable.Select("UserID = " + Security.CurrentUserID)[0]["Name"].ToString();

            //foreach (DataRow Row in NewsDataTable.Rows)
            //{
            //    Row["Sender"] = UsersDataTable.Select("UserID = " + Row["SenderID"] + " AND SenderTypeID = " + Row["SenderTypeID"])[0]["Name"].ToString();

            //    if (Row["SenderTypeID"].ToString() == "0")
            //    {
            //        if (CurrentUserSender == Row["Sender"].ToString())
            //            Row["CanEdit"] = true;
            //        else
            //            Row["CanEdit"] = false;
            //    }
            //    else
            //    {
            //        if (Convert.ToInt32(UsersDataTable.Select("UserID = " + Security.CurrentUserID)[0]["DepartmentID"]) ==
            //            Convert.ToInt32(Row["SenderID"]))
            //            Row["CanEdit"] = true;
            //        else
            //            Row["CanEdit"] = false;
            //    }

            //}
        }

        public void FillLikes()
        {
            BlogNewsLikesDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT * FROM BlogLikes", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(BlogNewsLikesDataTable);
            }


            CommentsLikesDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT * FROM BlogCommentsLikes", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(CommentsLikesDataTable);
            }
        }

        public void FillMoreNews(int Count)
        {
            BlogNewsDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT TOP " + Count + " * FROM Blog WHERE Pending <> 1  ORDER BY LastCommentDateTime DESC", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(BlogNewsDataTable);
            }
        }

        public bool IsMoreNews(int Count)
        {
            using (var DA = new SqlDataAdapter("SELECT Count(NewsID) FROM Blog", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    return Convert.ToInt32(DT.Rows[0][0]) > Count;
                }
            }
        }

        //public void FillCurrentComments(int NewsID)
        //{
        //    if (CurrentCommentsDataTable == null)
        //    {
        //        CurrentCommentsDataTable = new DataTable();
        //        CurrentCommentsDataTable.Columns.Add(new DataColumn("User", Type.GetType("System.String")));
        //    }
        //    else
        //        CurrentCommentsDataTable.Clear();

        //    using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM NewsComments WHERE NewsID = " + NewsID + " ORDER BY DateTime ASC", ConnectionStrings.LightConnectionString))
        //    {
        //        DA.Fill(CurrentCommentsDataTable);

        //        foreach (DataRow Row in CurrentCommentsDataTable.Rows)
        //        {
        //            Row["User"] = UsersDataTable.Select("SenderTypeID = 0 AND UserID = " + Row["UserID"])[0]["Name"];
        //        }
        //    }
        //}

        //public void GetCurrentNewsData(int NewsID, ref string Header, ref string SenderAndDate, ref string Text, ref Image Photo)
        //{ 
        //    DataRow[] Row = NewsDataTable.Select("NewsID = "+NewsID);

        //    Header = Row[0]["HeaderText"].ToString();
        //    SenderAndDate = Row[0]["Sender"].ToString() + ", " + Row[0]["DateTime"].ToString();
        //    Text = Row[0]["BodyText"].ToString();

        //    DataRow[] URow = UsersDataTable.Select("UserID = " + Row[0]["SenderID"] + " AND SenderTypeID = " + Row[0]["SenderTypeID"]);

        //    if (Row.Count() > 0 && URow[0]["Photo"] != DBNull.Value)
        //    {

        //        byte[] b = (byte[])URow[0]["Photo"];
        //        MemoryStream ms = new MemoryStream(b);

        //        Photo = Image.FromStream(ms);

        //        ms.Dispose();
        //    }
        //}

        //private int GetNewNews()
        //{
        //    using (SqlDataAdapter DA = new SqlDataAdapter("SELECT SubscribesRecordID, TableItemID FROM SubscribesRecords WHERE UserID = " + Security.CurrentUserID,
        //                                                            ConnectionStrings.LightConnectionString))
        //    {
        //        using (DataTable DT = new DataTable())
        //        {
        //            return DA.Fill(DT);
        //        }
        //    }
        //}

        //private int GetNewCommentsCount(int NewsID)
        //{
        //    using (SqlDataAdapter DA = new SqlDataAdapter("SELECT NewsSubsRecordID, NewsCommentID FROM NewsSubsRecords WHERE UserID = " + Security.CurrentUserID +
        //                                                 " AND NewsID = " + NewsID, ConnectionStrings.LightConnectionString))
        //    {
        //        using (DataTable DT = new DataTable())
        //        {
        //            return DA.Fill(DT);
        //        }
        //    }
        //}

        public DateTime AddNews(int SenderID, int SenderTypeID, string HeaderText, string BodyText, int NewsCategoryID)
        {
            DateTime DateTime;

            using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM Blog", ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        var NewRow = DT.NewRow();
                        NewRow["SenderID"] = SenderID;
                        NewRow["SenderTypeID"] = SenderTypeID;
                        NewRow["HeaderText"] = HeaderText;
                        NewRow["BodyText"] = BodyText;
                        NewRow["Pending"] = true;

                        DateTime = Security.GetCurrentDate();

                        NewRow["DateTime"] = DateTime;
                        NewRow["LastCommentDateTime"] = DateTime;
                        DT.Rows.Add(NewRow);

                        DA.Update(DT);

                        FillNews();
                    }
                }
            }

            return DateTime;
        }

        public void AddSubscribeForNews(DateTime DateTime)
        {
            using (var sDA = new SqlDataAdapter("SELECT * FROM Blog WHERE DateTime = '" + DateTime.ToString("yyyy-MM-dd HH:mm:ss.fff") + "'",
                                                                       ConnectionStrings.LightConnectionString))
            {
                using (var sDT = new DataTable())
                {
                    sDA.Fill(sDT);

                    ActiveNotifySystem.CreateSubscribeRecord(3, Convert.ToInt32(sDT.Rows[0][0]), Security.CurrentUserID);
                }
            }
        }

        public void ClearPending(DateTime DateTime)
        {
            using (var sDA = new SqlDataAdapter("SELECT TOP 1 * FROM Blog WHERE DateTime = '" + DateTime.ToString("yyyy-MM-dd HH:mm:ss.fff") + "'",
                                                                      ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(sDA))
                {
                    using (var sDT = new DataTable())
                    {
                        sDA.Fill(sDT);

                        sDT.Rows[0]["Pending"] = false;

                        sDA.Update(sDT);
                    }

                }

            }
        }


        public void AddComments(int UserID, int NewsID, string Text)
        {
            DateTime Date;

            using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM BlogComments", ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        var NewRow = DT.NewRow();
                        NewRow["NewsID"] = NewsID;
                        NewRow["NewsComment"] = Text;
                        NewRow["UserID"] = UserID;

                        Date = Security.GetCurrentDate();

                        NewRow["DateTime"] = Date;
                        DT.Rows.Add(NewRow);

                        DA.Update(DT);
                    }
                }
            }

            using (var DA = new SqlDataAdapter("SELECT NewsID, LastCommentDateTime FROM Blog WHERE NewsID =" + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        DT.Rows[0]["LastCommentDateTime"] = Date;

                        DA.Update(DT);
                    }
                }
            }

            AddNewsCommentsSubs(NewsID, UserID, Date);
        }

        public void LikeNews(int UserID, int NewsID)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM BlogLikes WHERE NewsID = " + NewsID + " AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        if (DT.Rows.Count > 0)//i like
                        {
                            using (var dDA = new SqlDataAdapter("DELETE FROM BlogLikes WHERE NewsID = " + NewsID + " AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
                            {
                                using (var dDT = new DataTable())
                                {
                                    dDA.Fill(dDT);
                                }
                            }

                            return;
                        }

                        var NewRow = DT.NewRow();
                        NewRow["NewsID"] = NewsID;
                        NewRow["UserID"] = UserID;
                        DT.Rows.Add(NewRow);

                        DA.Update(DT);
                    }
                }
            }
        }

        public void LikeComments(int UserID, int NewsID, int NewsCommentID)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM BlogCommentsLikes WHERE NewsCommentID = " + NewsCommentID + " AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        if (DT.Rows.Count > 0)
                        {
                            using (var dDA = new SqlDataAdapter("DELETE FROM BlogCommentsLikes WHERE NewsCommentID = " + NewsCommentID + " AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
                            {
                                using (var dDT = new DataTable())
                                {
                                    dDA.Fill(dDT);
                                }
                            }

                            return;
                        }

                        var NewRow = DT.NewRow();
                        NewRow["NewsID"] = NewsID;
                        NewRow["NewsCommentID"] = NewsCommentID;
                        NewRow["UserID"] = UserID;
                        DT.Rows.Add(NewRow);

                        DA.Update(DT);
                    }
                }
            }
        }

        public void EditComments(int UserID, int NewsCommentID, string Text)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM BlogComments WHERE NewsCommentID = " + NewsCommentID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        DT.Rows[0]["NewsComment"] = Text;

                        DA.Update(DT);
                    }
                }
            }
        }

        public bool Attach(DataTable AttachmentsDataTable, DateTime NewsDateTime, ref int CurrentUploadedFile)
        {
            if (AttachmentsDataTable.Rows.Count == 0)
                return true;

            var Ok = true;

            int NewsID;

            using (var DA = new SqlDataAdapter("SELECT NewsID FROM Blog WHERE DateTime = '" + NewsDateTime.ToString("yyyy-MM-dd HH:mm:ss.fff") + "'", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    NewsID = Convert.ToInt32(DT.Rows[0]["NewsID"]);
                }
            }

            using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM BlogAttachs", ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        foreach (DataRow Row in AttachmentsDataTable.Rows)
                        {
                            FileInfo fi;

                            try
                            {
                                fi = new FileInfo(Row["Path"].ToString());

                            }
                            catch
                            {
                                Ok = false;
                                continue;
                            }

                            var NewRow = DT.NewRow();
                            NewRow["NewsID"] = NewsID;
                            NewRow["FileName"] = Row["FileName"];
                            NewRow["FileSize"] = fi.Length;
                            DT.Rows.Add(NewRow);
                        }

                        DA.Update(DT);
                    }
                }
            }

            //write to ftp
            using (var DA = new SqlDataAdapter("SELECT * FROM BlogAttachs WHERE NewsID = " + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    foreach (DataRow Row in AttachmentsDataTable.Rows)
                    {
                        CurrentUploadedFile++;

                        try
                        {
                            if (FM.UploadFile(Row["Path"].ToString(), Configs.DocumentsPath +
                                          FileManager.GetPath("CoderBlog") + "/" +
                                          DT.Select("FileName = '" + Row["FileName"] + "'")[0]["NewsAttachID"] + ".idf", Configs.FTPType) == false)
                                break;
                        }
                        catch
                        {
                            Ok = false;
                            break;
                        }
                    }

                    DA.Update(DT);
                }
            }

            ReloadAttachments();

            return Ok;
        }

        public bool EditAttachments(int NewsID, DataTable AttachmentsDataTable, ref int CurrentUploadedFile, ref int TotalFilesCount)
        {
            var Ok = false;

            using (var DA = new SqlDataAdapter("SELECT NewsAttachID, NewsID, FileName FROM BlogAttachs WHERE NewsID = " + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        foreach (DataRow Row in DT.Rows)
                        {
                            if (AttachmentsDataTable.Select("FileName = '" + Row["FileName"] + "'").Count() == 0)
                            {
                                FM.DeleteFile(Configs.DocumentsPath + FileManager.GetPath("CoderBlog") + "/" +
                                              DT.Select("FileName = '" + Row["FileName"] + "'")[0]["NewsAttachID"] + ".idf", Configs.FTPType);
                                Row.Delete();
                            }
                        }

                        DA.Update(DT);
                    }
                }
            }

            using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM BlogAttachs WHERE NewsID = " + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);
                        Ok = true;

                        //add new
                        foreach (DataRow Row in AttachmentsDataTable.Rows)
                        {
                            if (Row["Path"].ToString() == "server")
                                continue;

                            TotalFilesCount++;

                            FileInfo fi;

                            try
                            {
                                fi = new FileInfo(Row["Path"].ToString());

                            }
                            catch
                            {
                                Ok = false;
                                continue;
                            }

                            var NewRow = DT.NewRow();
                            NewRow["NewsID"] = NewsID;
                            NewRow["FileName"] = Row["FileName"];
                            NewRow["FileSize"] = fi.Length;
                            DT.Rows.Add(NewRow);
                        }

                        DA.Update(DT);
                    }
                }
            }


            //write to ftp
            using (var DA = new SqlDataAdapter("SELECT * FROM BlogAttachs WHERE NewsID = " + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    foreach (DataRow Row in AttachmentsDataTable.Rows)
                    {
                        if (Row["Path"].ToString() == "server")
                            continue;

                        CurrentUploadedFile++;

                        try
                        {
                            if (FM.UploadFile(Row["Path"].ToString(), Configs.DocumentsPath +
                                         FileManager.GetPath("CoderBlog") + "/" +
                                         DT.Select("FileName = '" + Row["FileName"] + "'")[0]["NewsAttachID"] + ".idf", Configs.FTPType) == false)
                                break;
                        }
                        catch
                        {
                            Ok = false;
                            break;
                        }
                    }

                    DA.Update(DT);
                }
            }


            ReloadAttachments();

            return Ok;
        }

        public void ReloadAttachments()
        {
            AttachmentsDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT NewsAttachID, NewsID, FileName, FileSize FROM BlogAttachs", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(AttachmentsDataTable);
            }
        }

        public void EditNews(int NewsID, int SenderID, int SenderTypeID, string HeaderText, string BodyText, int NewsCategoryID, DateTime DateTime)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM Blog WHERE NewsID = " + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        if (DA.Fill(DT) == 0)
                            return;

                        DT.Rows[0]["SenderID"] = SenderID;
                        DT.Rows[0]["SenderTypeID"] = SenderTypeID;
                        DT.Rows[0]["HeaderText"] = HeaderText;
                        DT.Rows[0]["BodyText"] = BodyText;
                        DT.Rows[0]["DateTime"] = DateTime;

                        DA.Update(DT);

                        FillNews();
                    }
                }
            }
        }

        public static string TempPath()
        {
            return Environment.GetEnvironmentVariable("TEMP");
        }

        public int GetNewsUpdatesCount()
        {
            using (var DA = new SqlDataAdapter("SELECT Count (SubscribesRecordID) FROM SubscribesRecords WHERE SubscribesItemID = 1 OR SubscribesItemID = 2 AND UserTypeID = 0 AND UserID = " + Security.CurrentUserID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    if (DA.Fill(DT) > 0)
                        return Convert.ToInt32(DT.Rows[0][0]);
                }
            }

            return 0;
        }

        public int GetNewsIDByDateTime(DateTime DateTime)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM Blog WHERE DateTime = '" + DateTime.ToString("yyyy-MM-dd HH:mm:ss.fff") + "'",
                                                                         ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    return Convert.ToInt32(DT.Rows[0]["NewsID"]);
                }
            }
        }

        public void RemoveNews(int NewsID)
        {
            ActiveNotifySystem.DeleteSubscribesRecord(3);

            using (var DA = new SqlDataAdapter("DELETE FROM BlogLikes WHERE NewsID = " + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }

            using (var DA = new SqlDataAdapter("DELETE FROM Blog WHERE NewsID = " + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);
                    }
                }
            }

            RemoveComments(NewsID);

            RemoveAttachments(NewsID);

            FillNews();

            ReloadAttachments();
        }

        public void RemoveComments(int NewsID)
        {
            DeleteCommentsSubs(NewsID);

            using (var DA = new SqlDataAdapter("DELETE FROM BlogCommentsLikes WHERE NewsID = " + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }

            using (var DA = new SqlDataAdapter("DELETE FROM BlogComments WHERE NewsID = " + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);
                    }
                }
            }


        }

        public void RemoveComment(int NewsCommentID)
        {
            using (var DA = new SqlDataAdapter("DELETE FROM BlogCommentsLikes WHERE NewsCommentID = " + NewsCommentID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }

            using (var DA = new SqlDataAdapter("DELETE FROM BlogComments WHERE NewsCommentID = " + NewsCommentID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }

            DeleteCommentsSub(NewsCommentID);
        }

        public void RemoveAttachments(int NewsID)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM BlogAttachs WHERE NewsID = " + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    foreach (DataRow Row in DT.Rows)
                    {
                        FM.DeleteFile(Configs.DocumentsPath + FileManager.GetPath("CoderBlog") + "/" +
                                      DT.Select("FileName = '" + Row["FileName"] + "'")[0]["NewsAttachID"] + ".idf", Configs.FTPType);
                    }
                }
            }



            using (var DA = new SqlDataAdapter("DELETE FROM BlogAttachs WHERE NewsID = " + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);
                    }
                }
            }

            ReloadAttachments();
        }

        public void RemoveCurrentAttachments(int NewsID, DataTable AttachmentsDT)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM BlogAttachs WHERE NewsID = " + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        foreach (DataRow Row in AttachmentsDT.Rows)
                        {
                            if (Row["Path"].ToString() == "server")
                                continue;

                            FM.DeleteFile(Configs.DocumentsPath + FileManager.GetPath("CoderBlog") + "/" +
                                      DT.Select("FileName = '" + Row["FileName"] + "'")[0]["NewsAttachID"] + ".idf", Configs.FTPType);

                            DT.Select("FileName = '" + Row["FileName"] + "'")[0].Delete();
                        }

                        DA.Update(DT);
                    }
                }
            }

            ReloadAttachments();
        }

        public DataTable GetAttachments(int NewsID)
        {
            using (var DA = new SqlDataAdapter("SELECT NewsAttachID, FileName FROM BlogAttachs WHERE NewsID = " + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    return DT;
                }
            }
        }

        public int GetThisNewsSenderTypeID(int NewsID)
        {
            return Convert.ToInt32(BlogNewsDataTable.Select("NewsID = " + NewsID)[0]["SenderTypeID"]);
        }

        public string GetThisNewsBodyText(int NewsID)
        {
            return BlogNewsDataTable.Select("NewsID = " + NewsID)[0]["BodyText"].ToString();
        }

        public string GetThisNewsHeaderText(int NewsID)
        {
            return BlogNewsDataTable.Select("NewsID = " + NewsID)[0]["HeaderText"].ToString();
        }

        public DateTime GetThisNewsDateTime(int NewsID)
        {
            return Convert.ToDateTime(BlogNewsDataTable.Select("NewsID = " + NewsID)[0]["DateTime"]);
        }

        public int ShowAttachDownloadMenu(string FileName)
        {
            //0 cancel, 1 open, 2 save


            return AttachDownloadForm.Result;
        }

        public string GetAttachmentName(int NewsAttachID)
        {
            return AttachmentsDataTable.Select("NewsAttachID = " + NewsAttachID)[0]["FileName"].ToString();
        }

        public string SaveFile(int NewsAttachID)//temp folder
        {
            var tempFolder = Environment.GetEnvironmentVariable("TEMP");

            var FileName = "";

            using (var DA = new SqlDataAdapter("SELECT FileName, FileSize FROM BlogAttachs WHERE NewsAttachID = " + NewsAttachID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    FileName = DT.Rows[0]["FileName"].ToString();

                    try
                    {
                        FM.DownloadFile(Configs.DocumentsPath + FileManager.GetPath("CoderBlog") + "/" + NewsAttachID + ".idf",
                                            tempFolder + "\\" + FileName, Convert.ToInt64(DT.Rows[0]["FileSize"]), Configs.FTPType);
                    }
                    catch
                    {
                        return null;
                    }

                    //byte[] b = (byte[])DT.Rows[0]["FileBytes"];
                    //MemoryStream ms = new MemoryStream(b);

                    //FileName = DT.Rows[0]["FileName"].ToString();

                    //FileStream s2;

                    //try
                    //{
                    //    s2 = new FileStream(tempFolder + "\\" + FileName, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite);
                    //}
                    //catch 
                    //{
                    //    FileName = GetNewFileName(tempFolder, DT.Rows[0]["FileName"].ToString());
                    //    s2 = new FileStream(tempFolder + "\\" + FileName, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite);
                    //}

                    //s2.Write(ms.ToArray(), 0, ms.Capacity);

                    //s2.Close();
                }
            }

            return tempFolder + "\\" + FileName;
        }

        public void SaveFile(int NewsAttachID, string sDestFileName)
        {
            using (var DA = new SqlDataAdapter("SELECT FileName, FileSize FROM BlogAttachs WHERE NewsAttachID = " + NewsAttachID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    var FileName = DT.Rows[0]["FileName"].ToString();

                    try
                    {
                        FM.DownloadFile(Configs.DocumentsPath + FileManager.GetPath("CoderBlog") + "/" + NewsAttachID + ".idf",
                                           sDestFileName, Convert.ToInt64(DT.Rows[0]["FileSize"]), Configs.FTPType);
                    }
                    catch
                    {
                    }
                }
            }
        }

        private string SetNumber(string FileName, int Number)
        {
            var Ext = "";
            var Name = "";

            for (var i = FileName.Length - 1; i > 0; i--)
            {
                if (FileName[i] == '.')
                {
                    Ext = FileName.Substring(i + 1, FileName.Length - i - 1);
                    Name = FileName.Substring(0, i) + " (" + Number + ")." + Ext;
                    break;
                }
            }

            return Name;
        }

        private string GetNewFileName(string path, string FileName)
        {
            var fileInfo = new FileInfo(path + "\\" + FileName);

            if (!fileInfo.Exists)
                return FileName;

            var Ok = false;
            var n = 1;

            while (!Ok)
            {
                fileInfo = new FileInfo(path + "\\" + SetNumber(FileName, n));

                if (!fileInfo.Exists)
                    return SetNumber(FileName, n);

                n++;
            }

            return "";
        }

        public void DeleteCommentsSub(int NewsCommentID)
        {
            using (var DA = new SqlDataAdapter("DELETE FROM SubscribesRecords WHERE SubscribesItemID = 4 AND TableItemID = " + NewsCommentID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }
        }

        public void DeleteCommentsSubs(int NewsID)
        {
            using (var DA = new SqlDataAdapter("DELETE FROM SubscribesRecords WHERE SubscribesItemID = 4 AND TableItemID IN (SELECT NewsCommentID FROM NewsComments WHERE NewsID = " + NewsID + ")", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }
        }

        public void AddNewsCommentsSubs(int NewsID, int UserID, DateTime DateTime)
        {
            using (var DA = new SqlDataAdapter("SELECT NewsCommentID FROM BlogComments WHERE DateTime = '" + DateTime.ToString("yyyy-MM-dd HH:mm:ss.fff") + "'", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    ActiveNotifySystem.CreateSubscribeRecord(4, Convert.ToInt32(DT.Rows[0][0]), Security.CurrentUserID);
                }
            }
        }
    }





    public class InfiniumProjects
    {
        public DataTable ProjectsDataTable;
        public DataTable DepartmentsDataTable;
        public DataTable ProjectUsersMembersDataTable;
        public DataTable ProjectStatusesDataTable;
        public DataTable UsersDataTable;
        public DataTable ProjectDepsMembersDataTable;

        public DataTable ProjectNewsDataTable;
        public DataTable ProjectNewsCommentsDataTable;
        public DataTable ProjectNewsLikesDataTable;
        public DataTable ProjectNewsCommentsLikesDataTable;
        public DataTable ProjectNewsCommentsSubsRecordsDataTable;
        public DataTable ProjectNewsSubsRecordsDataTable;
        public DataTable ProjectNewsAttachmentsDataTable;
        public DataTable ProjectSubsRecordsDataTable;

        public DataTable CurrentProjectUsersDataTable;
        public DataTable CurrentProjectDepartmentsDataTable;

        public FileManager FM = new FileManager();

        public InfiniumProjects()
        {
            ProjectsDataTable = new DataTable();
            ProjectUsersMembersDataTable = new DataTable();
            ProjectStatusesDataTable = new DataTable();
            UsersDataTable = new DataTable();
            DepartmentsDataTable = new DataTable();
            ProjectDepsMembersDataTable = new DataTable();

            ProjectNewsDataTable = new DataTable();
            ProjectNewsCommentsDataTable = new DataTable();
            ProjectNewsLikesDataTable = new DataTable();
            ProjectNewsCommentsLikesDataTable = new DataTable();
            ProjectNewsCommentsSubsRecordsDataTable = new DataTable();
            ProjectNewsSubsRecordsDataTable = new DataTable();
            ProjectNewsAttachmentsDataTable = new DataTable();
            ProjectSubsRecordsDataTable = new DataTable();

            CurrentProjectUsersDataTable = new DataTable();
            CurrentProjectDepartmentsDataTable = new DataTable();
        }

        public void Fill()
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM ProjectStatuses", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(ProjectStatusesDataTable);
            }

            UsersDataTable = new DataTable();
            using (var DA = new SqlDataAdapter("SELECT * FROM Users ORDER BY Name", ConnectionStrings.UsersConnectionString))
            {
                UsersDataTable.Clear();
                DA.Fill(UsersDataTable);
            }

            using (var DA = new SqlDataAdapter("SELECT DepartmentID, DepartmentName FROM Departments", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(DepartmentsDataTable);
            }
        }

        public void FillProjects(int StateID, int GroupID, int UserID, int DepartmentID)
        {
            ProjectsDataTable.Clear();

            var FillExpr = "";

            if (StateID == 4)//all states
                FillExpr = "";
            else
                FillExpr = "WHERE ProjectStatusID = " + StateID;

            if (GroupID == 0)//my projects
            {
                if (FillExpr.Length > 0)
                    FillExpr += "AND ProjectID IN (SELECT ProjectID FROM ProjectMembers WHERE (ProjectUserTypeID = 0 AND UserID = " + Security.CurrentUserID +
                        ") OR (ProjectUserTypeID = 1 AND UserID = " + UsersDataTable.Select("UserID = " + Security.CurrentUserID)[0]["DepartmentID"] + " ))";
                else
                    FillExpr = "WHERE ProjectID IN (SELECT ProjectID FROM ProjectMembers WHERE (ProjectUserTypeID = 0 AND UserID = " + Security.CurrentUserID +
                                " OR (ProjectUserTypeID = 1 AND UserID = " + UsersDataTable.Select("UserID = " + Security.CurrentUserID)[0]["DepartmentID"] + " )))";
            }

            if (GroupID == 1)//users
            {
                if (FillExpr.Length > 0)
                    FillExpr += "AND ProjectID IN (SELECT ProjectID FROM ProjectMembers WHERE ProjectUserTypeID = 0 AND UserID = " + UserID + " )";
                else
                    FillExpr = "WHERE ProjectID IN (SELECT ProjectID FROM ProjectMembers WHERE ProjectUserTypeID = 0 AND UserID = " + UserID + " )";
            }

            if (GroupID == 2)//departments
            {
                if (FillExpr.Length > 0)
                    FillExpr += "AND ProjectID IN (SELECT ProjectID FROM ProjectMembers WHERE ProjectUserTypeID = 1 AND UserID = " + DepartmentID + " )";
                else
                    FillExpr = "WHERE ProjectID IN (SELECT ProjectID FROM ProjectMembers WHERE ProjectUserTypeID = 1 AND UserID = " + DepartmentID + " )";
            }

            if (GroupID == 3)//all groups
            {
                if (FillExpr.Length > 0)
                    FillExpr += "";
                else
                    FillExpr = "";
            }

            if (FillExpr.Length > 0)//pending
                FillExpr += " AND Pending = 0";
            else
                FillExpr = "WHERE Pending = 0";

            if (FillExpr.Length > 0)
                FillExpr += " AND IsProposition = 0";
            else
                FillExpr = " WHERE IsProposition = 0";

            using (var DA = new SqlDataAdapter("SELECT * FROM Projects " + FillExpr, ConnectionStrings.LightConnectionString))
            {
                DA.Fill(ProjectsDataTable);
            }
        }

        public void FillProjectsNew(int UserID)
        {
            ProjectsDataTable.Clear();

            var FillExpr = "";

            if (ProjectSubsRecordsDataTable.Rows.Count == 0)
                return;

            foreach (DataRow Row in ProjectSubsRecordsDataTable.Rows)
            {
                if (FillExpr.Length == 0)
                    FillExpr += " ProjectID = " + Row["TableItemID"];
                else
                    FillExpr += " OR ProjectID = " + Row["TableItemID"];
            }

            using (var DA = new SqlDataAdapter("SELECT * FROM Projects WHERE " + FillExpr + " AND Pending = 0 AND IsProposition = 0", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(ProjectsDataTable);
            }
        }

        public void FillProject(int ProjectID)
        {
            ProjectsDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT * FROM Projects WHERE ProjectID = " + ProjectID + " AND Pending = 0", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(ProjectsDataTable);
            }
        }

        public void FillProjectsNewsUpdates(int UserID)
        {
            ProjectsDataTable.Clear();

            var FillExpr = "";

            foreach (DataRow Row in ProjectNewsSubsRecordsDataTable.Rows)
            {
                if (FillExpr.Length == 0)
                    FillExpr += " ProjectID = " + Row["ProjectID"];
                else
                    FillExpr += " OR ProjectID = " + Row["ProjectID"];
            }

            foreach (DataRow Row in ProjectNewsCommentsSubsRecordsDataTable.Rows)
            {
                if (FillExpr.Length == 0)
                    FillExpr += " ProjectID = " + Row["ProjectID"];
                else
                    FillExpr += " OR ProjectID = " + Row["ProjectID"];
            }



            using (var DA = new SqlDataAdapter("SELECT * FROM Projects WHERE " + FillExpr + " AND Pending = 0", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(ProjectsDataTable);
            }
        }

        public void FillMembers(int StateID)
        {
            ProjectUsersMembersDataTable.Clear();

            if (StateID < 4)//not all
            {
                using (var DA = new SqlDataAdapter("SELECT infiniu2_users.dbo.Users.UserID, infiniu2_users.dbo.Users.Name" +
                                                   " FROM infiniu2_users.dbo.Users WHERE Fired<> 1 AND  infiniu2_users.dbo.Users.UserID IN" +
                                                   " (SELECT DISTINCT UserID FROM infiniu2_light.dbo.ProjectMembers WHERE ProjectUserTypeID = 0 AND " +
                                                   " ProjectID IN (SELECT ProjectID FROM Projects WHERE ProjectStatusID = " + StateID +
                                                   ")) ORDER BY infiniu2_users.dbo.Users.Name", ConnectionStrings.LightConnectionString))
                {
                    DA.Fill(ProjectUsersMembersDataTable);
                }
            }
            else
            {
                using (var DA = new SqlDataAdapter("SELECT infiniu2_users.dbo.Users.UserID, infiniu2_users.dbo.Users.Name" +
                                                   " FROM infiniu2_users.dbo.Users WHERE Fired<> 1 AND  infiniu2_users.dbo.Users.UserID IN" +
                                                   " (SELECT DISTINCT UserID FROM infiniu2_light.dbo.ProjectMembers)" +
                                                   " ORDER BY infiniu2_users.dbo.Users.Name",
                                                              ConnectionStrings.LightConnectionString))
                {
                    DA.Fill(ProjectUsersMembersDataTable);
                }
            }



            ProjectDepsMembersDataTable.Clear();

            if (StateID < 4)//not all
            {
                using (var DA = new SqlDataAdapter("SELECT infiniu2_light.dbo.Departments.DepartmentID, infiniu2_light.dbo.Departments.DepartmentName " +
                                                   " FROM infiniu2_light.dbo.Departments WHERE infiniu2_light.dbo.Departments.DepartmentID IN " +
                                                   " (SELECT DISTINCT UserID FROM infiniu2_light.dbo.ProjectMembers WHERE ProjectUserTypeID = 1 AND " +
                                                   " ProjectID IN (SELECT ProjectID FROM Projects WHERE ProjectStatusID = " + StateID +
                                                   ")) ORDER BY infiniu2_light.dbo.Departments.DepartmentName", ConnectionStrings.LightConnectionString))
                {
                    DA.Fill(ProjectDepsMembersDataTable);
                }
            }
            else
            {
                using (var DA = new SqlDataAdapter("SELECT infiniu2_light.dbo.Departments.DepartmentID, infiniu2_light.dbo.Departments.DepartmentName " +
                                                   " FROM infiniu2_light.dbo.Departments WHERE infiniu2_light.dbo.Departments.DepartmentID IN " +
                                                   " (SELECT DISTINCT UserID FROM infiniu2_light.dbo.ProjectMembers)" +
                                                   " ORDER BY infiniu2_light.dbo.Departments.DepartmentName",
                                                               ConnectionStrings.LightConnectionString))
                {
                    DA.Fill(ProjectDepsMembersDataTable);
                }
            }
        }

        public void AddSubscribeForNewProject(int ProjectID)
        {
            var NewRow = ProjectSubsRecordsDataTable.NewRow();
            NewRow["SubscribesItemID"] = 7;
            NewRow["TableItemID"] = ProjectID;
            NewRow["UserID"] = Security.CurrentUserID;
            ProjectSubsRecordsDataTable.Rows.Add(NewRow);
        }

        public void FillProjectNews(int ProjectID)
        {
            ProjectNewsDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT TOP 20 * FROM ProjectNews WHERE Pending <> 1 AND ProjectID = " + ProjectID + " ORDER BY LastCommentDateTime DESC", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(ProjectNewsDataTable);
            }

            if (ProjectNewsDataTable.Columns["New"] == null)
                ProjectNewsDataTable.Columns.Add(new DataColumn("New", Type.GetType("System.Int32")));

            if (ProjectNewsDataTable.Columns["NewComments"] == null)
                ProjectNewsDataTable.Columns.Add(new DataColumn("NewComments", Type.GetType("System.Int32")));

            foreach (DataRow Row in ProjectNewsSubsRecordsDataTable.Rows)
            {
                if (ProjectNewsDataTable.Select("NewsID = " + Row["TableItemID"]).Count() > 0)
                {
                    ProjectNewsDataTable.Select("NewsID = " + Row["TableItemID"])[0]["New"] = 1;
                }
            }

            foreach (DataRow Row in ProjectNewsCommentsSubsRecordsDataTable.Rows)
            {
                var Rows = ProjectNewsDataTable.Select("NewsID = " + Row["NewsID"]);

                if (Rows.Count() > 0)
                    Rows[0]["New"] = 1;
            }
        }

        public void FillMoreNews(int Count, int ProjectID)
        {
            ProjectNewsDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT TOP " + Count + " * FROM ProjectNews WHERE Pending <> 1 AND ProjectID = " + ProjectID + " ORDER BY LastCommentDateTime DESC", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(ProjectNewsDataTable);
            }

            if (ProjectNewsDataTable.Columns["New"] == null)
                ProjectNewsDataTable.Columns.Add(new DataColumn("New", Type.GetType("System.Int32")));

            if (ProjectNewsDataTable.Columns["NewComments"] == null)
                ProjectNewsDataTable.Columns.Add(new DataColumn("NewComments", Type.GetType("System.Int32")));

            foreach (DataRow Row in ProjectNewsSubsRecordsDataTable.Rows)
            {
                if (ProjectNewsDataTable.Select("NewsID = " + Row["NewsID"]).Count() > 0)
                {
                    ProjectNewsDataTable.Select("NewsID = " + Row["NewsID"])[0]["New"] = 1;
                }
            }

            foreach (DataRow Row in ProjectNewsCommentsSubsRecordsDataTable.Rows)
            {
                if (ProjectNewsDataTable.Select("NewsID = " + Row["NewsID"]).Count() > 0)
                {
                    ProjectNewsDataTable.Select("NewsID = " + Row["NewsID"])[0]["NewComments"] =
                            ProjectNewsCommentsSubsRecordsDataTable.Select("NewsID = " + Row["NewsID"]).Count();
                }
            }
        }

        public void FillPropositions()
        {
            ProjectsDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT * FROM Projects WHERE IsProposition = 1", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(ProjectsDataTable);
            }
        }

        public void FillComments()
        {
            ProjectNewsCommentsDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT * FROM ProjectNewsComments", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(ProjectNewsCommentsDataTable);
            }
        }

        public void FillAttachments()
        {
            ProjectNewsAttachmentsDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT NewsAttachID, NewsID, FileName, FileSize FROM ProjectNewsAttachs", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(ProjectNewsAttachmentsDataTable);
            }
        }

        public void FillProjectSubs()
        {
            ProjectNewsCommentsSubsRecordsDataTable.Clear();

            //comments
            using (var DA = new SqlDataAdapter("SELECT SubscribesRecords.*, ProjectNewsComments.ProjectID, ProjectNewsComments.NewsID FROM SubscribesRecords" +
                                               " INNER JOIN ProjectNewsComments ON ProjectNewsComments.NewsCommentID = SubscribesRecords.TableItemID WHERE SubscribesItemID = 9 AND UserTypeID = 0 AND SubscribesRecords.UserID = " + Security.CurrentUserID, ConnectionStrings.LightConnectionString))
            {
                DA.Fill(ProjectNewsCommentsSubsRecordsDataTable);
            }


            ProjectNewsSubsRecordsDataTable.Clear();

            //news
            using (var sDA = new SqlDataAdapter("SELECT * FROM SubscribesRecords " +
                                                " INNER JOIN ProjectNews ON ProjectNews.NewsID = SubscribesRecords.TableItemID WHERE SubscribesItemID = 8 AND UserTypeID = 0 AND SubscribesRecords.UserID = " + Security.CurrentUserID, ConnectionStrings.LightConnectionString))
            {
                sDA.Fill(ProjectNewsSubsRecordsDataTable);
            }


            ProjectSubsRecordsDataTable.Clear();

            //projects
            using (var sDA = new SqlDataAdapter("SELECT * FROM SubscribesRecords WHERE SubscribesItemID = 7 AND UserTypeID = 0 AND UserID = " + Security.CurrentUserID, ConnectionStrings.LightConnectionString))
            {
                sDA.Fill(ProjectSubsRecordsDataTable);
            }
        }

        public void FillLikes()
        {
            ProjectNewsLikesDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT * FROM ProjectNewsLikes", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(ProjectNewsLikesDataTable);
            }


            ProjectNewsCommentsLikesDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT * FROM ProjectNewsCommentsLikes", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(ProjectNewsCommentsLikesDataTable);
            }
        }


        public void FillProjectMembers(int ProjectID)
        {
            CurrentProjectUsersDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT infiniu2_users.dbo.Users.UserID, infiniu2_users.dbo.Users.Name" +
                                               " FROM infiniu2_users.dbo.Users WHERE Fired<> 1 AND infiniu2_users.dbo.Users.UserID IN" +
                                               " (SELECT DISTINCT UserID FROM infiniu2_light.dbo.ProjectMembers WHERE ProjectID = " + ProjectID + ")" +
                                               " ORDER BY infiniu2_users.dbo.Users.Name",
                                                              ConnectionStrings.LightConnectionString))
            {
                DA.Fill(CurrentProjectUsersDataTable);
            }



            CurrentProjectDepartmentsDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT infiniu2_light.dbo.Departments.DepartmentID, infiniu2_light.dbo.Departments.DepartmentName " +
                                               " FROM infiniu2_light.dbo.Departments WHERE infiniu2_light.dbo.Departments.DepartmentID IN " +
                                               " (SELECT DISTINCT UserID FROM infiniu2_light.dbo.ProjectMembers WHERE ProjectID = " + ProjectID + ")" +
                                               " ORDER BY infiniu2_light.dbo.Departments.DepartmentName",
                                                               ConnectionStrings.LightConnectionString))
            {
                DA.Fill(CurrentProjectDepartmentsDataTable);
            }
        }

        public void ReloadAttachments()
        {
            ProjectNewsAttachmentsDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT NewsAttachID, NewsID, FileName, FileSize FROM ProjectNewsAttachs", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(ProjectNewsAttachmentsDataTable);
            }
        }


        public void ReloadNews(int Count, int ProjectID)
        {
            ProjectNewsDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT TOP " + Count + " * FROM ProjectNews WHERE Pending <> 1 AND ProjectID = " + ProjectID + " ORDER BY LastCommentDateTime DESC", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(ProjectNewsDataTable);
            }

            if (ProjectNewsDataTable.Columns["New"] == null)
                ProjectNewsDataTable.Columns.Add(new DataColumn("New", Type.GetType("System.Int32")));

            if (ProjectNewsDataTable.Columns["NewComments"] == null)
                ProjectNewsDataTable.Columns.Add(new DataColumn("NewComments", Type.GetType("System.Int32")));

            foreach (DataRow Row in ProjectNewsSubsRecordsDataTable.Rows)
            {
                if (ProjectNewsDataTable.Select("NewsID = " + Row["ProjectID"]).Count() > 0)
                {
                    ProjectNewsDataTable.Select("NewsID = " + Row["ProjectID"])[0]["New"] = 1;
                }
            }

            foreach (DataRow Row in ProjectNewsCommentsSubsRecordsDataTable.Rows)
            {
                ProjectNewsDataTable.Select("NewsID = " + Row["NewsID"])[0]["New"] = 1;
            }
        }

        public void ReloadComments()
        {
            ProjectNewsCommentsDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT * FROM ProjectNewsComments", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(ProjectNewsCommentsDataTable);
            }
        }

        public void ReloadSubscribes()
        {
            ProjectNewsCommentsDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT * FROM ProjectNewsComments", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(ProjectNewsCommentsDataTable);
            }


            ProjectNewsCommentsSubsRecordsDataTable.Clear();

            //comments
            using (var DA = new SqlDataAdapter("SELECT SubscribesRecords.*, ProjectNewsComments.ProjectID, ProjectNewsComments.NewsID FROM SubscribesRecords" +
                                               " INNER JOIN ProjectNewsComments ON ProjectNewsComments.NewsCommentID = SubscribesRecords.TableItemID WHERE SubscribesItemID = 9 AND UserTypeID = 0 AND SubscribesRecords.UserID = " + Security.CurrentUserID, ConnectionStrings.LightConnectionString))
            {
                DA.Fill(ProjectNewsCommentsSubsRecordsDataTable);
            }


            ProjectNewsSubsRecordsDataTable.Clear();

            //news
            using (var sDA = new SqlDataAdapter("SELECT * FROM SubscribesRecords " +
                                                " INNER JOIN ProjectNews ON ProjectNews.NewsID = SubscribesRecords.TableItemID WHERE SubscribesItemID = 8 AND UserTypeID = 0 AND SubscribesRecords.UserID = " + Security.CurrentUserID, ConnectionStrings.LightConnectionString))
            {
                sDA.Fill(ProjectNewsSubsRecordsDataTable);
            }


            ProjectSubsRecordsDataTable.Clear();

            //projects
            using (var sDA = new SqlDataAdapter("SELECT * FROM SubscribesRecords WHERE SubscribesItemID = 7 AND UserID = " + Security.CurrentUserID, ConnectionStrings.LightConnectionString))
            {
                sDA.Fill(ProjectSubsRecordsDataTable);
            }
        }

        public void AddSubscribeToUpdates(int ModuleID, int ItemID, int UserID)
        {
            using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM SubscribesToUpdates", ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        var NewRow = DT.NewRow();
                        NewRow["ModuleID"] = ModuleID;
                        NewRow["TableItemID"] = ItemID;
                        NewRow["UserID"] = UserID;
                        DT.Rows.Add(NewRow);

                        DA.Update(DT);
                    }
                }
            }
        }

        public void RemoveSubscribeToUpdates(int ModuleID, int ItemID, int UserID)
        {
            using (var DA = new SqlDataAdapter("DELETE FROM SubscribesToUpdates WHERE ModuleID = " + ModuleID + "AND UserID = " + UserID + " AND TableItemID = " + ItemID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }
        }


        public string GetUserName(int UserID)
        {
            return UsersDataTable.Select("UserID = " + UserID)[0]["Name"].ToString();
        }

        public Image GetUserPhoto(int UserID)
        {
            var Row = UsersDataTable.Select("UserID = " + UserID)[0];

            if (Row["Photo"] == DBNull.Value)
                return null;

            Image Image = null;

            var b = (byte[])Row["Photo"];

            using (var ms = new MemoryStream(b))
            {
                Image = Image.FromStream(ms);
            }

            return Image;
        }

        public int AddProject(TreeView TreeView, string ProjectName, string ProjectDescription, int AuthorID, bool bProposition)
        {
            DateTime DateTime;

            using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM Projects", ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        var NewRow = DT.NewRow();
                        NewRow["ProjectName"] = ProjectName;
                        NewRow["ProjectDescription"] = ProjectDescription;
                        NewRow["AuthorID"] = AuthorID;
                        NewRow["IsProposition"] = bProposition;
                        NewRow["Pending"] = true;

                        DateTime = Security.GetCurrentDate();

                        if (bProposition)
                            NewRow["ProjectStatusID"] = 1;
                        else
                            NewRow["ProjectStatusID"] = 0;

                        if (!bProposition)
                            NewRow["StartDate"] = DateTime;

                        NewRow["CreationDate"] = DateTime;
                        DT.Rows.Add(NewRow);

                        DA.Update(DT);
                    }
                }
            }

            var ProjectID = -1;

            //get project id AND add members
            using (var DA = new SqlDataAdapter("SELECT ProjectID FROM Projects WHERE CreationDate = '" + DateTime.ToString("yyyy-MM-dd HH:mm:ss.fff") + "'", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    ProjectID = Convert.ToInt32(DT.Rows[0]["ProjectID"]);

                    AddProjectMembers(TreeView, Convert.ToInt32(DT.Rows[0]["ProjectID"]));
                }
            }

            return ProjectID;
        }


        public int EditProject(TreeView TreeView, bool bNewMembers, string ProjectName, string ProjectDescription, int ProjectID, bool bProposition)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM Projects WHERE ProjectID = " + ProjectID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        DT.Rows[0]["ProjectName"] = ProjectName;
                        DT.Rows[0]["ProjectDescription"] = ProjectDescription;
                        DT.Rows[0]["Pending"] = true;
                        DT.Rows[0]["IsProposition"] = bProposition;


                        DA.Update(DT);
                    }
                }
            }

            if (bNewMembers)
                EditProjectMembers(TreeView, ProjectID, bProposition);

            return ProjectID;
        }


        public void AddProjectMembers(TreeView MembersTree, int ProjectID)
        {
            using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM ProjectMembers", ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        for (var i = 0; i < MembersTree.Nodes.Count; i++)
                        {
                            var bAll = true;

                            var UDT = new DataTable();
                            UDT.Columns.Add(new DataColumn("UserName", Type.GetType("System.String")));

                            for (var j = 0; j < MembersTree.Nodes[i].Nodes.Count; j++)
                            {
                                if (MembersTree.Nodes[i].Nodes[j].Checked)
                                {
                                    var NewRow = UDT.NewRow();
                                    NewRow["UserName"] = MembersTree.Nodes[i].Nodes[j].Text;
                                    UDT.Rows.Add(NewRow);
                                }
                                else
                                    bAll = false;
                            }

                            if (bAll)
                            {
                                UDT.Clear();

                                var NewRow = DT.NewRow();
                                NewRow["ProjectID"] = ProjectID;
                                NewRow["UserID"] = DepartmentsDataTable.Select("DepartmentName = '" + MembersTree.Nodes[i].Text + "'")[0]["DepartmentID"];
                                NewRow["ProjectUserTypeID"] = 1;
                                DT.Rows.Add(NewRow);
                            }
                            else
                            {
                                foreach (DataRow Row in UDT.Rows)
                                {
                                    var NewRow = DT.NewRow();
                                    NewRow["ProjectID"] = ProjectID;
                                    NewRow["UserID"] = UsersDataTable.Select("Name = '" + Row["UserName"] + "'")[0]["UserID"];
                                    NewRow["ProjectUserTypeID"] = 0;
                                    DT.Rows.Add(NewRow);
                                }
                            }
                        }

                        DA.Update(DT);

                    }
                }
            }
        }

        public void EditProjectMembers(TreeView MembersTree, int ProjectID, bool bProposition)
        {
            var MembersDataTable = new DataTable();
            MembersDataTable.Columns.Add(new DataColumn("UserID", Type.GetType("System.Int32")));

            //get old members list
            using (var DA = new SqlDataAdapter("SELECT UserID, ProjectUserTypeID FROM ProjectMembers WHERE ProjectID = " + ProjectID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    foreach (DataRow Row in DT.Rows)
                    {
                        if (Row["ProjectUserTypeID"].ToString() == "0")//user
                        {
                            var NewRow = MembersDataTable.NewRow();
                            NewRow["UserID"] = Row["UserID"];
                            MembersDataTable.Rows.Add(NewRow);
                        }
                        else//department
                        {
                            foreach (var uRow in UsersDataTable.Select("DepartmentID = " + Row["UserID"]))
                            {
                                var NewRow = MembersDataTable.NewRow();
                                NewRow["UserID"] = uRow["UserID"];
                                MembersDataTable.Rows.Add(NewRow);
                            }
                        }
                    }
                }
            }



            using (var DA = new SqlDataAdapter("DELETE FROM ProjectMembers WHERE ProjectID = " + ProjectID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);
                    }
                }
            }


            using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM ProjectMembers", ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        for (var i = 0; i < MembersTree.Nodes.Count; i++)
                        {
                            var bAll = true;

                            var UDT = new DataTable();
                            UDT.Columns.Add(new DataColumn("UserName", Type.GetType("System.String")));

                            for (var j = 0; j < MembersTree.Nodes[i].Nodes.Count; j++)
                            {
                                if (MembersTree.Nodes[i].Nodes[j].Checked)
                                {
                                    var NewRow = UDT.NewRow();
                                    NewRow["UserName"] = MembersTree.Nodes[i].Nodes[j].Text;
                                    UDT.Rows.Add(NewRow);
                                }
                                else
                                    bAll = false;
                            }

                            if (bAll)
                            {
                                UDT.Clear();

                                var NewRow = DT.NewRow();
                                NewRow["ProjectID"] = ProjectID;
                                NewRow["UserID"] = DepartmentsDataTable.Select("DepartmentName = '" + MembersTree.Nodes[i].Text + "'")[0]["DepartmentID"];
                                NewRow["ProjectUserTypeID"] = 1;
                                DT.Rows.Add(NewRow);


                            }
                            else
                            {
                                foreach (DataRow Row in UDT.Rows)
                                {
                                    var NewRow = DT.NewRow();
                                    NewRow["ProjectID"] = ProjectID;
                                    NewRow["UserID"] = UsersDataTable.Select("Name = '" + Row["UserName"] + "'")[0]["UserID"];
                                    NewRow["ProjectUserTypeID"] = 0;
                                    DT.Rows.Add(NewRow);
                                }
                            }
                        }

                        DA.Update(DT);


                        //set new members datatable where department expands to users
                        var NewMembersDT = new DataTable();
                        NewMembersDT = MembersDataTable.Clone();

                        foreach (DataRow mRow in DT.Rows)
                        {
                            if (mRow["ProjectUserTypeID"].ToString() == "0")//user
                            {
                                var NewRow = NewMembersDT.NewRow();
                                NewRow["UserID"] = mRow["UserID"];
                                NewMembersDT.Rows.Add(NewRow);
                            }
                            else//department
                            {
                                foreach (var uRow in UsersDataTable.Select("DepartmentID = " + mRow["UserID"]))
                                {
                                    var NewRow = NewMembersDT.NewRow();
                                    NewRow["UserID"] = uRow["UserID"];
                                    NewMembersDT.Rows.Add(NewRow);
                                }
                            }
                        }

                        for (var i = 0; i < MembersDataTable.Rows.Count; i++)
                        {
                            if (NewMembersDT.Select("UserID = " + MembersDataTable.Rows[i]["UserID"]).Count() == 0)//was deleted
                            {
                                ClearProjectSubscribe(ProjectID, Convert.ToInt32(MembersDataTable.Rows[i]["UserID"]));
                                MembersDataTable.Rows[i].Delete();
                                i--;
                            }
                        }

                        foreach (DataRow uRow in NewMembersDT.Rows)
                        {
                            var mRow = MembersDataTable.Select("UserID = " + uRow["UserID"]);

                            if (mRow.Count() == 0)//new user
                            {
                                var NewRow = MembersDataTable.NewRow();
                                NewRow["UserID"] = uRow["UserID"];
                                MembersDataTable.Rows.Add(NewRow);
                            }
                            else
                                mRow[0].Delete();

                        }
                    }
                }
            }


            AddProjectSubscribe(ProjectID, MembersDataTable, bProposition);
        }

        public void EditNews(int NewsID, string BodyText)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM ProjectNews WHERE NewsID = " + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        if (DA.Fill(DT) == 0)
                            return;

                        DT.Rows[0]["BodyText"] = BodyText;

                        DA.Update(DT);
                    }
                }
            }
        }

        public bool CanRemove(int ProjectID)
        {
            using (var DA = new SqlDataAdapter("SELECT ProjectID FROM Projects WHERE ProjectID = " + ProjectID + " AND AuthorID = " + Security.CurrentUserID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    return (DA.Fill(DT) > 0);//only author can delete project
                }
            }
        }

        public bool CanEdit(int ProjectID)
        {
            using (var DA = new SqlDataAdapter("SELECT ProjectID FROM Projects WHERE ProjectID = " + ProjectID + " AND AuthorID = " + Security.CurrentUserID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    return (DA.Fill(DT) > 0);//only author can delete project
                }
            }
        }

        public void RemoveProject(int ProjectID)
        {
            using (var DA = new SqlDataAdapter("DELETE FROM Projects WHERE ProjectID = " + ProjectID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }

            using (var DA = new SqlDataAdapter("DELETE FROM SubscribesRecords WHERE (SubscribesItemID = 12 OR SubscribesItemID = 7 OR SubscribesItemID = 8 OR SubscribesItemID = 9) AND TableItemID = " + ProjectID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }

            using (var DA = new SqlDataAdapter("DELETE FROM ProjectNews WHERE ProjectID = " + ProjectID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }

            using (var DA = new SqlDataAdapter("DELETE FROM ProjectMembers WHERE ProjectID = " + ProjectID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }

            using (var DA = new SqlDataAdapter("DELETE FROM ProjectNewsAttachs WHERE ProjectID = " + ProjectID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }

            using (var DA = new SqlDataAdapter("DELETE FROM ProjectNewsComments WHERE ProjectID = " + ProjectID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }

            using (var DA = new SqlDataAdapter("DELETE FROM ProjectNewsCommentsLikes WHERE ProjectID = " + ProjectID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }

            using (var DA = new SqlDataAdapter("DELETE FROM ProjectNewsLikes WHERE ProjectID = " + ProjectID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }
        }


        public bool IsProposition(int ProjectID)
        {
            using (var DA = new SqlDataAdapter("SELECT ProjectID, IsProposition FROM Projects WHERE ProjectID = " + ProjectID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    return Convert.ToBoolean(DT.Rows[0]["IsProposition"]);
                }
            }
        }


        public bool EditAttachments(int NewsID, int ProjectID, DataTable AttachmentsDataTable, ref int CurrentUploadedFile, ref int TotalFilesCount)
        {
            var Ok = false;

            using (var DA = new SqlDataAdapter("SELECT NewsAttachID, NewsID, FileName FROM ProjectNewsAttachs WHERE NewsID = " + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        foreach (DataRow Row in DT.Rows)
                        {
                            if (AttachmentsDataTable.Select("FileName = '" + Row["FileName"] + "'").Count() == 0)
                            {
                                FM.DeleteFile(Configs.DocumentsPath + FileManager.GetPath("ProjectNews") + "/" +
                                              DT.Select("FileName = '" + Row["FileName"] + "'")[0]["NewsAttachID"] + ".idf", Configs.FTPType);
                                Row.Delete();
                            }
                        }

                        DA.Update(DT);
                    }
                }
            }

            using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM ProjectNewsAttachs WHERE NewsID = " + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);
                        Ok = true;

                        //add new
                        foreach (DataRow Row in AttachmentsDataTable.Rows)
                        {
                            if (Row["Path"].ToString() == "server")
                                continue;

                            TotalFilesCount++;

                            FileInfo fi;

                            try
                            {
                                fi = new FileInfo(Row["Path"].ToString());

                            }
                            catch
                            {
                                Ok = false;
                                continue;
                            }

                            var NewRow = DT.NewRow();
                            NewRow["NewsID"] = NewsID;
                            NewRow["ProjectID"] = ProjectID;
                            NewRow["FileName"] = Row["FileName"];
                            NewRow["FileSize"] = fi.Length;
                            DT.Rows.Add(NewRow);
                        }

                        DA.Update(DT);
                    }
                }
            }


            //write to ftp
            using (var DA = new SqlDataAdapter("SELECT * FROM ProjectNewsAttachs WHERE NewsID = " + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    foreach (DataRow Row in AttachmentsDataTable.Rows)
                    {
                        if (Row["Path"].ToString() == "server")
                            continue;

                        CurrentUploadedFile++;

                        try
                        {
                            if (FM.UploadFile(Row["Path"].ToString(), Configs.DocumentsPath +
                                         FileManager.GetPath("ProjectNews") + "/" +
                                         DT.Select("FileName = '" + Row["FileName"] + "'")[0]["NewsAttachID"] + ".idf", Configs.FTPType) == false)
                                break;
                        }
                        catch
                        {
                            Ok = false;
                            break;
                        }
                    }

                    DA.Update(DT);
                }
            }


            ReloadAttachments();

            return Ok;
        }


        public DataTable GetAttachments(int NewsID)
        {
            using (var DA = new SqlDataAdapter("SELECT NewsAttachID, FileName FROM ProjectNewsAttachs WHERE NewsID = " + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    return DT;
                }
            }
        }

        public int GetPropositionsUpdatesCount()
        {
            var c = 0;

            using (var DA = new SqlDataAdapter("SELECT Count (SubscribesRecordID) FROM SubscribesRecords WHERE SubscribesItemID = 12 AND UserTypeID = 0 AND UserID = " + Security.CurrentUserID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    if (DA.Fill(DT) > 0)
                        c += Convert.ToInt32(DT.Rows[0][0]);
                }
            }

            return c;
        }

        public int GetNewsUpdatesCount()
        {
            var c = 0;

            using (var DA = new SqlDataAdapter("SELECT Count (SubscribesRecordID) FROM SubscribesRecords WHERE SubscribesItemID = 8 OR SubscribesItemID = 9 AND UserTypeID = 0 AND UserID = " + Security.CurrentUserID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    if (DA.Fill(DT) > 0)
                        c += Convert.ToInt32(DT.Rows[0][0]);
                }
            }

            return c;
        }

        public int GetProjectsUpdatesCount()
        {
            var c = 0;

            using (var DA = new SqlDataAdapter("SELECT Count (SubscribesRecordID) FROM SubscribesRecords WHERE SubscribesItemID = 7 AND UserID = " + Security.CurrentUserID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    if (DA.Fill(DT) > 0)
                        c += Convert.ToInt32(DT.Rows[0][0]);
                }
            }

            return c;
        }

        public void GetProjectsUpdates()
        {
            ProjectSubsRecordsDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT * FROM SubscribesRecords WHERE SubscribesItemID = 7 AND UserTypeID = 0 AND UserID = " + Security.CurrentUserID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(ProjectSubsRecordsDataTable);
                }
            }
        }

        public void GetNewsUpdates()
        {
            ProjectNewsCommentsSubsRecordsDataTable.Clear();

            //comments
            using (var DA = new SqlDataAdapter("SELECT SubscribesRecords.*, ProjectNewsComments.ProjectID, ProjectNewsComments.NewsID FROM SubscribesRecords" +
                                               " INNER JOIN ProjectNewsComments ON ProjectNewsComments.NewsCommentID = SubscribesRecords.TableItemID WHERE SubscribesItemID = 9 AND UserTypeID = 0 AND SubscribesRecords.UserID = " + Security.CurrentUserID, ConnectionStrings.LightConnectionString))
            {
                DA.Fill(ProjectNewsCommentsSubsRecordsDataTable);
            }


            ProjectNewsSubsRecordsDataTable.Clear();

            //news
            using (var sDA = new SqlDataAdapter("SELECT * FROM SubscribesRecords " +
                                                " INNER JOIN ProjectNews ON ProjectNews.NewsID = SubscribesRecords.TableItemID WHERE SubscribesItemID = 8 AND UserTypeID = 0 AND SubscribesRecords.UserID = " + Security.CurrentUserID, ConnectionStrings.LightConnectionString))
            {
                sDA.Fill(ProjectNewsSubsRecordsDataTable);
            }
        }


        public void ClearAllSubs()
        {
            using (var DA = new SqlDataAdapter("DELETE FROM SubscribesRecords WHERE (SubscribesItemID = 7 OR SubscribesItemID = 8 OR SubscribesItemID = 9) AND UserTypeID = 0 AND UserID = " + Security.CurrentUserID,
                                                          ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);
                    }
                }
            }
        }


        public void RemoveCurrentAttachments(int NewsID, DataTable AttachmentsDT)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM ProjectNewsAttachs WHERE NewsID = " + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        foreach (DataRow Row in AttachmentsDT.Rows)
                        {
                            if (Row["Path"].ToString() == "server")
                                continue;

                            FM.DeleteFile(Configs.DocumentsPath + FileManager.GetPath("ProjectNews") + "/" +
                                      DT.Select("FileName = '" + Row["FileName"] + "'")[0]["NewsAttachID"] + ".idf", Configs.FTPType);

                            DT.Select("FileName = '" + Row["FileName"] + "'")[0].Delete();
                        }

                        DA.Update(DT);
                    }
                }
            }

            ReloadAttachments();
        }

        public void RemoveNews(int NewsID)
        {
            using (var DA = new SqlDataAdapter("DELETE FROM ProjectNewsLikes WHERE NewsID = " + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }

            using (var DA = new SqlDataAdapter("DELETE FROM SubscribesRecords WHERE SubscribesItemID = 8 AND TableItemID = " + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }

            using (var DA = new SqlDataAdapter("DELETE FROM ProjectNews WHERE NewsID = " + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }

            RemoveComments(NewsID);

            RemoveAttachments(NewsID);
        }

        public void RemoveComments(int NewsID)
        {
            DeleteCommentsSubs(NewsID);

            using (var DA = new SqlDataAdapter("DELETE FROM ProjectNewsCommentsLikes WHERE NewsID = " + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }

            using (var DA = new SqlDataAdapter("DELETE FROM ProjectNewsComments WHERE NewsID = " + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);
                    }
                }
            }

        }


        public void DeleteCommentsSub(int NewsCommentID)
        {
            using (var DA = new SqlDataAdapter("DELETE FROM SubscribesRecords WHERE SubscribesItemID = 9 AND TableItemID = " + NewsCommentID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }
        }

        public void DeleteCommentsSubs(int NewsID)
        {
            using (var DA = new SqlDataAdapter("DELETE FROM SubscribesRecords WHERE SubscribesItemID = 9 AND TableItemID IN" +
                                               "(SELECT NewsCommentID FROM ProjectNewsComments WHERE NewsID = " + NewsID + ")", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }
        }


        public void RemoveAttachments(int NewsID)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM ProjectNewsAttachs WHERE NewsID = " + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    foreach (DataRow Row in DT.Rows)
                    {
                        FM.DeleteFile(Configs.DocumentsPath + FileManager.GetPath("ProjectNews") + "/" +
                                      DT.Select("FileName = '" + Row["FileName"] + "'")[0]["NewsAttachID"] + ".idf", Configs.FTPType);
                    }
                }
            }



            using (var DA = new SqlDataAdapter("DELETE FROM ProjectNewsAttachs WHERE NewsID = " + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);
                    }
                }
            }

            ReloadAttachments();
        }


        public void LikeNews(int UserID, int NewsID, int ProjectID)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM ProjectNewsLikes WHERE NewsID = " + NewsID + " AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        if (DT.Rows.Count > 0)//i like
                        {
                            using (var dDA = new SqlDataAdapter("DELETE FROM ProjectNewsLikes WHERE NewsID = " + NewsID + " AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
                            {
                                using (var dDT = new DataTable())
                                {
                                    dDA.Fill(dDT);
                                }
                            }

                            return;
                        }

                        var NewRow = DT.NewRow();
                        NewRow["NewsID"] = NewsID;
                        NewRow["UserID"] = UserID;
                        NewRow["ProjectID"] = ProjectID;
                        DT.Rows.Add(NewRow);

                        DA.Update(DT);
                    }
                }
            }
        }

        public void LikeComments(int UserID, int NewsID, int NewsCommentID, int ProjectID)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM ProjectNewsCommentsLikes WHERE NewsCommentID = " + NewsCommentID + " AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        if (DT.Rows.Count > 0)
                        {
                            using (var dDA = new SqlDataAdapter("DELETE FROM ProjectNewsCommentsLikes WHERE NewsCommentID = " + NewsCommentID + " AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
                            {
                                using (var dDT = new DataTable())
                                {
                                    dDA.Fill(dDT);
                                }
                            }

                            return;
                        }

                        var NewRow = DT.NewRow();
                        NewRow["NewsID"] = NewsID;
                        NewRow["NewsCommentID"] = NewsCommentID;
                        NewRow["UserID"] = UserID;
                        NewRow["ProjectID"] = ProjectID;
                        DT.Rows.Add(NewRow);

                        DA.Update(DT);
                    }
                }
            }
        }



        public int GetNewsIDByDateTime(DateTime DateTime)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM ProjectNews WHERE DateTime = '" + DateTime.ToString("yyyy-MM-dd HH:mm:ss.fff") + "'",
                                                                         ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    return Convert.ToInt32(DT.Rows[0]["NewsID"]);
                }
            }
        }


        public void ClearNewsPending(DateTime DateTime)
        {
            using (var sDA = new SqlDataAdapter("SELECT TOP 1 * FROM ProjectNews WHERE DateTime = '" + DateTime.ToString("yyyy-MM-dd HH:mm:ss.fff") + "'",
                                                                      ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(sDA))
                {
                    using (var sDT = new DataTable())
                    {
                        sDA.Fill(sDT);

                        sDT.Rows[0]["Pending"] = false;

                        sDA.Update(sDT);
                    }

                }

            }
        }

        public void ClearProjectsPending(int ProjectID)
        {
            using (var sDA = new SqlDataAdapter("SELECT TOP 1 * FROM Projects WHERE ProjectID = " + ProjectID,
                                                                      ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(sDA))
                {
                    using (var sDT = new DataTable())
                    {
                        sDA.Fill(sDT);

                        sDT.Rows[0]["Pending"] = false;

                        sDA.Update(sDT);
                    }

                }

            }
        }


        public bool Attach(DataTable AttachmentsDataTable, DateTime NewsDateTime, ref int CurrentUploadedFile, int ProjectID)
        {
            if (AttachmentsDataTable.Rows.Count == 0)
                return true;

            var Ok = true;

            int NewsID;

            using (var DA = new SqlDataAdapter("SELECT NewsID FROM ProjectNews WHERE DateTime = '" + NewsDateTime.ToString("yyyy-MM-dd HH:mm:ss.fff") + "'", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    NewsID = Convert.ToInt32(DT.Rows[0]["NewsID"]);
                }
            }

            using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM ProjectNewsAttachs", ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        foreach (DataRow Row in AttachmentsDataTable.Rows)
                        {
                            FileInfo fi;

                            try
                            {
                                fi = new FileInfo(Row["Path"].ToString());
                            }
                            catch
                            {
                                Ok = false;
                                continue;
                            }

                            var NewRow = DT.NewRow();
                            NewRow["NewsID"] = NewsID;
                            NewRow["FileName"] = Row["FileName"];
                            NewRow["FileSize"] = fi.Length;
                            NewRow["ProjectID"] = ProjectID;
                            DT.Rows.Add(NewRow);
                        }

                        DA.Update(DT);
                    }
                }
            }

            //write to ftp
            using (var DA = new SqlDataAdapter("SELECT * FROM ProjectNewsAttachs WHERE NewsID = " + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    foreach (DataRow Row in AttachmentsDataTable.Rows)
                    {
                        CurrentUploadedFile++;

                        try
                        {
                            if (FM.UploadFile(Row["Path"].ToString(), Configs.DocumentsPath +
                                          FileManager.GetPath("ProjectNews") + "/" +
                                          DT.Select("FileName = '" + Row["FileName"] + "'")[0]["NewsAttachID"] + ".idf", Configs.FTPType) == false)
                            {
                                break;
                            }
                        }
                        catch
                        {
                            Ok = false;
                            break;
                        }
                    }

                    DA.Update(DT);
                }
            }

            ReloadAttachments();

            return Ok;
        }

        public bool CanChangeStatusProject(int ProjectID)
        {
            using (var DA = new SqlDataAdapter("SELECT ProjectMemberID FROM ProjectMembers WHERE ProjectID = " + ProjectID +
                                               " AND ProjectUserTypeID = 0 AND UserID = " + Security.CurrentUserID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    if (DA.Fill(DT) > 0)
                    {
                        return true;
                    }
                }
            }

            using (var DA = new SqlDataAdapter("SELECT ProjectMemberID FROM ProjectMembers WHERE ProjectID = " + ProjectID +
                                               " AND ProjectUserTypeID = 1 AND UserID = " +
                                               UsersDataTable.Select("UserID = " + Security.CurrentUserID)[0]["DepartmentID"], ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    if (DA.Fill(DT) > 0)
                    {
                        return true;
                    }
                }
            }

            using (var DA = new SqlDataAdapter("SELECT ProjectID FROM Projects WHERE ProjectID = " + ProjectID +
                                               " AND AuthorID = " + Security.CurrentUserID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    if (DA.Fill(DT) > 0)
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        public int GetProjectStatus(int ProjectID)
        {
            using (var DA = new SqlDataAdapter("SELECT ProjectID, ProjectStatusID FROM Projects WHERE ProjectID = " + ProjectID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    return Convert.ToInt32(DT.Rows[0]["ProjectStatusID"]);
                }
            }
        }

        public void StartProject(int ProjectID)
        {
            using (var DA = new SqlDataAdapter("SELECT ProjectID, StartDate, ProjectStatusID FROM Projects WHERE ProjectID = " + ProjectID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        DT.Rows[0]["StartDate"] = Security.GetCurrentDate();
                        DT.Rows[0]["ProjectStatusID"] = 0;

                        DA.Update(DT);
                    }
                }
            }
        }


        public void StartProject(int ProjectID, bool bProposition)
        {
            using (var DA = new SqlDataAdapter("SELECT ProjectID, StartDate, IsProposition, ProjectStatusID FROM Projects WHERE ProjectID = " + ProjectID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        DT.Rows[0]["StartDate"] = Security.GetCurrentDate();
                        DT.Rows[0]["ProjectStatusID"] = 0;

                        if (bProposition)
                            DT.Rows[0]["IsProposition"] = false;

                        DA.Update(DT);
                    }
                }
            }
        }


        public void PauseProject(int ProjectID)
        {
            using (var DA = new SqlDataAdapter("SELECT ProjectID, SuspendedDate, ProjectStatusID FROM Projects WHERE ProjectID = " + ProjectID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        DT.Rows[0]["SuspendedDate"] = Security.GetCurrentDate();
                        DT.Rows[0]["ProjectStatusID"] = 1;

                        DA.Update(DT);
                    }
                }
            }
        }

        public void CancelProject(int ProjectID)
        {
            using (var DA = new SqlDataAdapter("SELECT ProjectID, CanceledDate, ProjectStatusID FROM Projects WHERE ProjectID = " + ProjectID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        DT.Rows[0]["CanceledDate"] = Security.GetCurrentDate();
                        DT.Rows[0]["ProjectStatusID"] = 2;

                        DA.Update(DT);
                    }
                }
            }
        }

        public void EndProject(int ProjectID)
        {
            using (var DA = new SqlDataAdapter("SELECT ProjectID, CompletedDate, ProjectStatusID FROM Projects WHERE ProjectID = " + ProjectID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        DT.Rows[0]["CompletedDate"] = Security.GetCurrentDate();
                        DT.Rows[0]["ProjectStatusID"] = 3;

                        DA.Update(DT);
                    }
                }
            }
        }

        public void AddNewsSubscribe(DateTime DateTime, int ProjectID, bool bAllUsers)
        {
            using (var sDA = new SqlDataAdapter("SELECT * FROM ProjectNews WHERE DateTime = '" + DateTime.ToString("yyyy-MM-dd HH:mm:ss.fff") + "'",
                                                                       ConnectionStrings.LightConnectionString))
            {
                using (var sDT = new DataTable())
                {
                    sDA.Fill(sDT);

                    using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM SubscribesRecords", ConnectionStrings.LightConnectionString))
                    {
                        using (var CB = new SqlCommandBuilder(DA))
                        {
                            using (var DT = new DataTable())
                            {
                                DA.Fill(DT);

                                if (bAllUsers)
                                {
                                    foreach (DataRow Row in UsersDataTable.Rows)
                                    {
                                        if (Convert.ToInt32(Row["UserID"]) == Security.CurrentUserID)
                                            continue;

                                        var NewRow = DT.NewRow();
                                        NewRow["SubscribesItemID"] = 8;
                                        NewRow["TableItemID"] = sDT.Rows[0]["NewsID"];
                                        NewRow["UserID"] = Row["UserID"];
                                        NewRow["UserTypeID"] = 0;
                                        DT.Rows.Add(NewRow);
                                    }
                                }
                                else//members only
                                {
                                    using (var mDA = new SqlDataAdapter("SELECT * FROM ProjectMembers WHERE ProjectID = " + ProjectID,
                                                                                 ConnectionStrings.LightConnectionString))
                                    {
                                        using (var mDT = new DataTable())
                                        {
                                            mDA.Fill(mDT);

                                            foreach (DataRow Row in mDT.Rows)
                                            {
                                                if (Row["ProjectUserTypeID"].ToString() == "0")//user
                                                {
                                                    if (Convert.ToInt32(Row["UserID"]) == Security.CurrentUserID)
                                                        continue;

                                                    var NewRow = DT.NewRow();
                                                    NewRow["SubscribesItemID"] = 8;
                                                    NewRow["TableItemID"] = sDT.Rows[0]["NewsID"];
                                                    NewRow["UserID"] = Row["UserID"];
                                                    NewRow["UserTypeID"] = 0;
                                                    DT.Rows.Add(NewRow);
                                                }
                                                else//department
                                                {
                                                    foreach (var uRow in UsersDataTable.Select("DepartmentID = " + Row["UserID"]))
                                                    {
                                                        if (Convert.ToInt32(uRow["UserID"]) == Security.CurrentUserID)
                                                            continue;

                                                        var NewRow = DT.NewRow();
                                                        NewRow["SubscribesItemID"] = 8;
                                                        NewRow["TableItemID"] = sDT.Rows[0]["NewsID"];
                                                        NewRow["UserID"] = uRow["UserID"];
                                                        NewRow["UserTypeID"] = 0;
                                                        DT.Rows.Add(NewRow);
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }

                                DA.Update(DT);
                            }
                        }
                    }
                }
            }
        }

        public void AddProjectSubscribe(int ProjectID, TreeView MembersTree, bool bAll, bool bProposition)
        {
            using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM SubscribesRecords", ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        if (bAll)
                        {
                            foreach (DataRow uRow in UsersDataTable.Rows)
                            {
                                if (Convert.ToInt32(uRow["UserID"]) == Security.CurrentUserID)
                                    continue;

                                var NewRow = DT.NewRow();
                                NewRow["UserID"] = uRow["UserID"];
                                NewRow["UserTypeID"] = 0;

                                if (bProposition)
                                    NewRow["SubscribesItemID"] = 12;
                                else
                                    NewRow["SubscribesItemID"] = 7;

                                NewRow["TableItemID"] = ProjectID;
                                DT.Rows.Add(NewRow);
                            }
                        }
                        else
                        {
                            for (var i = 0; i < MembersTree.Nodes.Count; i++)
                            {
                                for (var j = 0; j < MembersTree.Nodes[i].Nodes.Count; j++)
                                {
                                    if (MembersTree.Nodes[i].Nodes[j].Checked)
                                    {
                                        var UserID = Convert.ToInt32(UsersDataTable.Select(
                                                                      "Name = '" + MembersTree.Nodes[i].Nodes[j].Text + "'")[0]["UserID"]);

                                        if (UserID == Security.CurrentUserID)
                                            continue;

                                        var NewRow = DT.NewRow();
                                        NewRow["UserID"] = UserID;
                                        NewRow["UserTypeID"] = 0;

                                        if (bProposition)
                                            NewRow["SubscribesItemID"] = 12;
                                        else
                                            NewRow["SubscribesItemID"] = 7;

                                        NewRow["TableItemID"] = ProjectID;
                                        DT.Rows.Add(NewRow);
                                    }
                                }
                            }
                        }

                        DA.Update(DT);
                    }
                }
            }
        }

        public void AddProjectSubscribe(int ProjectID, DataTable NewMembersDT, bool bProposition)//new members
        {
            using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM SubscribesRecords", ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        foreach (DataRow Row in NewMembersDT.Rows)
                        {
                            if (Row.RowState == DataRowState.Deleted)
                                continue;

                            if (Convert.ToInt32(Row["UserID"]) == Security.CurrentUserID)
                                continue;

                            var NewRow = DT.NewRow();

                            if (bProposition)
                                NewRow["SubscribesItemID"] = 12;
                            else
                                NewRow["SubscribesItemID"] = 7;

                            NewRow["UserID"] = Row["UserID"];
                            NewRow["UserTypeID"] = 0;
                            NewRow["TableItemID"] = ProjectID;

                            DT.Rows.Add(NewRow);
                        }

                        DA.Update(DT);
                    }
                }
            }
        }

        public void ClearProjectSubscribe(int ProjectID, int UserID)
        {
            using (var DA = new SqlDataAdapter("DELETE FROM SubscribesRecords WHERE SubscribesItemID = 7 AND UserTypeID = 0 AND UserID = "
                                               + UserID + " AND TableItemID = " + ProjectID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }
        }

        public bool IsMoreNews(int Count, int ProjectID)
        {
            using (var DA = new SqlDataAdapter("SELECT Count(NewsID) FROM ProjectNews WHERE ProjectID = " + ProjectID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    return Convert.ToInt32(DT.Rows[0][0]) > Count;
                }
            }
        }

        public DateTime AddNews(int SenderID, int ProjectID, string BodyText)
        {
            DateTime DateTime;

            using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM ProjectNews", ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        var NewRow = DT.NewRow();
                        NewRow["SenderID"] = SenderID;
                        NewRow["ProjectID"] = ProjectID;
                        NewRow["BodyText"] = BodyText;
                        NewRow["Pending"] = true;

                        DateTime = Security.GetCurrentDate();

                        NewRow["DateTime"] = DateTime;
                        NewRow["LastCommentDateTime"] = DateTime;
                        DT.Rows.Add(NewRow);

                        DA.Update(DT);
                    }
                }
            }

            return DateTime;
        }


        public string GetAttachmentName(int NewsAttachID)
        {
            return ProjectNewsAttachmentsDataTable.Select("NewsAttachID = " + NewsAttachID)[0]["FileName"].ToString();
        }

        public string SaveFile(int NewsAttachID)//temp folder
        {
            var tempFolder = Environment.GetEnvironmentVariable("TEMP");

            var FileName = "";

            using (var DA = new SqlDataAdapter("SELECT FileName, FileSize FROM ProjectNewsAttachs WHERE NewsAttachID = " + NewsAttachID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    FileName = DT.Rows[0]["FileName"].ToString();

                    try
                    {
                        FM.DownloadFile(Configs.DocumentsPath + FileManager.GetPath("ProjectNews") + "/" + NewsAttachID + ".idf",
                                            tempFolder + "\\" + FileName, Convert.ToInt64(DT.Rows[0]["FileSize"]), Configs.FTPType);
                    }
                    catch
                    {
                        return null;
                    }

                }
            }

            return tempFolder + "\\" + FileName;
        }

        public void SaveFile(int NewsAttachID, string sDestFileName)
        {
            using (var DA = new SqlDataAdapter("SELECT FileName, FileSize FROM ProjectNewsAttachs WHERE NewsAttachID = " + NewsAttachID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    var FileName = DT.Rows[0]["FileName"].ToString();

                    try
                    {
                        FM.DownloadFile(Configs.DocumentsPath + FileManager.GetPath("ProjectNews") + "/" + NewsAttachID + ".idf",
                                           sDestFileName, Convert.ToInt64(DT.Rows[0]["FileSize"]), Configs.FTPType);
                    }
                    catch
                    {
                    }
                }
            }
        }



        public void AddComments(int UserID, int ProjectID, int NewsID, string Text, bool bNoNotify)
        {
            DateTime Date;

            using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM ProjectNewsComments", ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        var NewRow = DT.NewRow();
                        NewRow["NewsID"] = NewsID;
                        NewRow["NewsComment"] = Text;
                        NewRow["UserID"] = UserID;
                        NewRow["ProjectID"] = ProjectID;

                        Date = Security.GetCurrentDate();

                        NewRow["DateTime"] = Date;
                        DT.Rows.Add(NewRow);

                        DA.Update(DT);
                    }
                }
            }

            using (var DA = new SqlDataAdapter("SELECT NewsID, LastCommentDateTime FROM ProjectNews WHERE NewsID =" + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        DT.Rows[0]["LastCommentDateTime"] = Date;

                        DA.Update(DT);
                    }
                }
            }

            if (!bNoNotify)
                AddNewsCommentsSubs(NewsID, ProjectID, UserID, Date);
        }


        public void AddNewsCommentsSubs(int NewsID, int ProjectID, int UserID, DateTime DateTime)
        {
            using (var DA = new SqlDataAdapter("SELECT NewsCommentID FROM ProjectNewsComments WHERE DateTime = '" + DateTime.ToString("yyyy-MM-dd HH:mm:ss.fff") + "'", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);


                    using (var cDA = new SqlDataAdapter("SELECT TOP 0 * FROM SubscribesRecords", ConnectionStrings.LightConnectionString))
                    {
                        using (var CB = new SqlCommandBuilder(cDA))
                        {
                            using (var cDT = new DataTable())
                            {
                                cDA.Fill(cDT);

                                {
                                    using (var mDA = new SqlDataAdapter("SELECT * FROM ProjectMembers WHERE ProjectID = " + ProjectID,
                                                                                 ConnectionStrings.LightConnectionString))
                                    {
                                        using (var mDT = new DataTable())
                                        {
                                            mDA.Fill(mDT);

                                            foreach (DataRow Row in mDT.Rows)
                                            {
                                                if (Row["ProjectUserTypeID"].ToString() == "0")//user
                                                {
                                                    if (Convert.ToInt32(Row["UserID"]) == Security.CurrentUserID)
                                                        continue;

                                                    var NewRow = cDT.NewRow();
                                                    NewRow["SubscribesItemID"] = 9;
                                                    NewRow["TableItemID"] = DT.Rows[0]["NewsCommentID"];
                                                    NewRow["UserID"] = Row["UserID"];
                                                    NewRow["UserTypeID"] = 0;
                                                    cDT.Rows.Add(NewRow);
                                                }
                                                else//department
                                                {
                                                    foreach (var uRow in UsersDataTable.Select("DepartmentID = " + Row["UserID"]))
                                                    {
                                                        if (Convert.ToInt32(uRow["UserID"]) == Security.CurrentUserID)
                                                            continue;

                                                        var NewRow = cDT.NewRow();
                                                        NewRow["SubscribesItemID"] = 9;
                                                        NewRow["TableItemID"] = DT.Rows[0]["NewsCommentID"];
                                                        NewRow["UserID"] = uRow["UserID"];
                                                        NewRow["UserTypeID"] = 0;
                                                        cDT.Rows.Add(NewRow);
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }

                                cDA.Update(cDT);
                            }
                        }
                    }
                }
            }
        }

        public void EditComments(int UserID, int NewsCommentID, string Text)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM ProjectNewsComments WHERE NewsCommentID = " + NewsCommentID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        DT.Rows[0]["NewsComment"] = Text;

                        DA.Update(DT);
                    }
                }
            }
        }


        public string GetThisNewsBodyText(int NewsID)
        {
            return ProjectNewsDataTable.Select("NewsID = " + NewsID)[0]["BodyText"].ToString();
        }

        public DateTime GetThisNewsDateTime(int NewsID)
        {
            return Convert.ToDateTime(ProjectNewsDataTable.Select("NewsID = " + NewsID)[0]["DateTime"]);
        }

        public int CheckSubscribeToUpdates(int ModuleID, int ItemID, int UserID)
        {
            using (var DA = new SqlDataAdapter("SELECT UserID FROM ProjectMembers WHERE ProjectID = " + ItemID + " AND ((ProjectUserTypeID = 0 AND UserID = " + UserID +
                                               ") OR (ProjectUserTypeID = 1 AND UserID = " + UsersDataTable.Select("UserID = " + UserID)[0]["DepartmentID"] + "))", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    if (DA.Fill(DT) > 0)
                        return -1;
                }
            }


            using (var DA = new SqlDataAdapter("SELECT * FROM SubscribesToUpdates WHERE UserID = " + UserID + " AND TableItemID = " + ItemID + " AND ModuleID = " + ModuleID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        if (DA.Fill(DT) > 0)
                            return 1;
                    }
                }
            }

            return 0;
        }


        public void RemoveComment(int NewsCommentID)
        {
            using (var DA = new SqlDataAdapter("DELETE FROM ProjectNewsComments WHERE NewsCommentID = " + NewsCommentID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }

            DeleteCommentsSub(NewsCommentID);
        }


    }





    public class LightUsers
    {
        public DataTable UsersDataTable;
        public DataTable DepartmentsDataTable;
        public DataTable PositionsDataTable;

        public LightUsers()
        {
            UsersDataTable = new DataTable();
            DepartmentsDataTable = new DataTable();

            Fill();
        }

        private void Fill()
        {
            PositionsDataTable = new DataTable();

            using (var DA = new SqlDataAdapter(@"SELECT dbo.StaffList.UserID,  dbo.Positions.Position, dbo.StaffList.FactoryID,
                dbo.StaffList.Rate FROM dbo.StaffList 
                INNER JOIN dbo.Positions ON dbo.StaffList.PositionID = dbo.Positions.PositionID", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(PositionsDataTable);
            }

            using (var DA = new SqlDataAdapter(@"SELECT StaffList.DepartmentID, Users.Name, Positions.Position, Factory.FactoryName FROM StaffList 
                INNER JOIN Positions ON StaffList.PositionID = Positions.PositionID 
                INNER JOIN infiniu2_users.dbo.Users AS Users ON StaffList.UserID = Users.UserID
                INNER JOIN infiniu2_catalog.dbo.Factory AS Factory ON StaffList.FactoryID = Factory.FactoryID
                ORDER BY Name, Position, FactoryName", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(UsersDataTable);
            }
            UsersDataTable = TablesManager.UsersDataTable.Copy();

            /*  using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM Users ORDER BY Name", ConnectionStrings.UsersConnectionString))
              {
                  DA.Fill(UsersDataTable);
              }*/

            UsersDataTable.Columns.Add(new DataColumn("ProfilPosition", Type.GetType("System.String")));
            UsersDataTable.Columns.Add(new DataColumn("TPSPosition", Type.GetType("System.String")));

            foreach (DataRow Row in UsersDataTable.Rows)
            {
                var rows = PositionsDataTable.Select("FactoryID=1 AND UserID = " + Row["UserID"]);
                var ProfilPosition = string.Empty;
                var TPSPosition = string.Empty;
                for (var i = 0; i < rows.Count(); i++)
                {
                    var Rate = string.Empty;
                    if (rows[i]["Rate"].ToString().Length > 0)
                        Rate = Convert.ToDecimal(rows[i]["Rate"]).ToString("G29");
                    ProfilPosition += rows[i]["Position"] +
                        " (" + Rate + " ОМЦ-ПРОФИЛЬ)";
                    if (i != rows.Count() - 1)
                        ProfilPosition = ProfilPosition.Insert(ProfilPosition.Length, "\n");
                }

                if (rows.Count() > 0)
                {
                    Row["ProfilPosition"] = ProfilPosition;
                }
                rows = PositionsDataTable.Select("FactoryID=2 AND UserID = " + Row["UserID"]);
                for (var i = 0; i < rows.Count(); i++)
                {
                    var Rate = string.Empty;
                    if (rows[i]["Rate"].ToString().Length > 0)
                        Rate = Convert.ToDecimal(rows[i]["Rate"]).ToString("G29");
                    TPSPosition += rows[i]["Position"] +
                        " (" + Rate + " ЗОВ-ТПС)";
                    if (i != rows.Count() - 1)
                        TPSPosition = TPSPosition.Insert(TPSPosition.Length, "\n");
                }
                if (rows.Count() > 0)
                {
                    Row["TPSPosition"] = TPSPosition;
                }
            }

            /*  using (SqlDataAdapter DA = new SqlDataAdapter("SELECT TOP 0 * FROM Departments ORDER BY DepartmentName", ConnectionStrings.LightConnectionString))
              {
                  DA.Fill(DepartmentsDataTable);
              }*/

            DepartmentsDataTable = TablesManager.DepartmentsDataTable.Copy();

            if (DepartmentsDataTable.Columns["Count"] == null)
                DepartmentsDataTable.Columns.Add(new DataColumn("Count", Type.GetType("System.String")));

            var NewRow = DepartmentsDataTable.NewRow();
            NewRow["DepartmentName"] = "Все группы";
            NewRow["Count"] = DepartmentsDataTable.Rows.Count;
            DepartmentsDataTable.Rows.InsertAt(NewRow, 0);

            /*  using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM Departments ORDER BY DepartmentName", ConnectionStrings.LightConnectionString))
              {
                  DA.Fill(DepartmentsDataTable);
              }*/

            DepartmentsDataTable.Columns["DepartmentName"].ColumnName = "Name";

            foreach (DataRow Row in DepartmentsDataTable.Rows)
            {
                if (Row["Name"].ToString() == "Все группы")
                {
                    Row["Count"] = UsersDataTable.Rows.Count;
                    continue;
                }

                var GU = UsersDataTable.Select("DepartmentID = " + Row["DepartmentID"]);

                Row["Count"] = GU.Count();
            }


            if (UsersDataTable.Columns["Department"] == null)
                UsersDataTable.Columns.Add(new DataColumn("Department", Type.GetType("System.String")));

            foreach (DataRow Row in UsersDataTable.Rows)
            {
                var GR = DepartmentsDataTable.Select("DepartmentID = " + Row["DepartmentID"]);

                if (GR.Count() > 0)
                    Row["Department"] = GR[0]["Name"];
                else
                    Row["Department"] = "не указана";

                if (Row["PersonalMobilePhone"] == DBNull.Value || Row["PersonalMobilePhone"].ToString().Length == 0)
                    Row["PersonalMobilePhone"] = "не указан";
                else
                    Row["PersonalMobilePhone"] = GetPhoneFormat(Row["PersonalMobilePhone"].ToString());

                if (Row["WorkMobilePhone"] == DBNull.Value || Row["WorkMobilePhone"].ToString().Length == 0)
                    Row["WorkMobilePhone"] = "не указан";
                else
                    Row["WorkMobilePhone"] = GetPhoneFormat(Row["WorkMobilePhone"].ToString());

                if (Row["WorkExtPhone"] == DBNull.Value || Row["WorkExtPhone"].ToString().Length == 0)
                    Row["WorkExtPhone"] = "не указан";

                if (Row["ICQ"] == DBNull.Value || Row["ICQ"].ToString().Length == 0)
                    Row["ICQ"] = "не указан";
            }
        }

        public static string GetPhoneFormat(string Phone)
        {
            if (Phone.Length == 0)
                return "не указан";

            return "+" + string.Format("{0:### ## ### ## ##}", Convert.ToUInt64(Phone));
        }


    }





    public class MailManager
    {
        public void MainManager()
        {

        }

        public static bool SendMessage(string Address, string Subject, string Body)
        {
            var AccountPassword = "allthebestondick1029{}1q";
            var SenderEmail = "zovprofilinfinium@gmail.com";

            var to = Address;
            var from = SenderEmail;

            using (var message = new MailMessage(from, to))
            {
                message.Subject = Subject;
                message.Body = Body;

                var client = new SmtpClient("smtp.gmail.com", 587)
                {
                    UseDefaultCredentials = false,
                    Credentials = new NetworkCredential(SenderEmail, AccountPassword),
                    EnableSsl = true
                };
                try
                {
                    client.Send(message);
                }

                catch
                {
                    return false;
                }

                client.Dispose();
            }

            return true;
        }
    }





    public class ActiveNotifySystem
    {
        public DataTable SubscribesItemsDataTable;
        public DataTable ModulesDataTable;
        public DataTable ModulesUpdatesDataTable;
        private int iCurrentUpdateID = -1;
        public int iCurrentModuleID = -1;

        public ActiveNotifySystem()
        {
            ModulesUpdatesDataTable = new DataTable();
            ModulesUpdatesDataTable.Columns.Add(new DataColumn("ModuleID", Type.GetType("System.Int32")));
            ModulesUpdatesDataTable.Columns.Add(new DataColumn("ModuleName", Type.GetType("System.String")));
            ModulesUpdatesDataTable.Columns.Add(new DataColumn("FormName", Type.GetType("System.String")));
            ModulesUpdatesDataTable.Columns.Add(new DataColumn("Count", Type.GetType("System.Int32")));

            Fill();
        }

        private void Fill()
        {
            SubscribesItemsDataTable = new DataTable();

            using (var DA = new SqlDataAdapter("SELECT * FROM SubscribesItems", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(SubscribesItemsDataTable);
            }

            ModulesDataTable = TablesManager.ModulesDataTable;
        }

        public static int GetModuleID(string FormName)
        {
            using (var DA = new SqlDataAdapter("SELECT ModuleID FROM Modules WHERE FormName = '" + FormName + "'", ConnectionStrings.UsersConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    return Convert.ToInt32(DT.Rows[0]["ModuleID"]);
                }
            }
        }

        public int GetModuleIDByForm(string FormName)
        {
            if (FormName == "LightStartForm")
                return 0;

            if (FormName == "UnderConstructionForm")
                return -2;

            if (ModulesDataTable.Select("FormName = '" + FormName + "'").Count() == 0)
                return -1;

            return Convert.ToInt32(ModulesDataTable.Select("FormName = '" + FormName + "'")[0]["ModuleID"]);
        }

        public static int IsNewUpdates(int UserID)
        {
            var iUpdates = 0;

            using (var DA = new SqlDataAdapter("SELECT Count(SubscribesRecordID) FROM SubscribesRecords WHERE UserTypeID = 0 AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    try
                    {
                        if (DA.Fill(DT) > 0)
                            iUpdates += Convert.ToInt32(DT.Rows[0][0]);
                    }
                    catch { return 0; }
                }
            }

            return iUpdates;
        }

        public Image GetModuleImage(int ModuleID)
        {
            var Row = ModulesDataTable.Select("ModuleID = " + ModuleID)[0];

            if (Row["Picture"] == DBNull.Value)
                return null;

            var b = (byte[])Row["Picture"];
            var ms = new MemoryStream(b);

            return Image.FromStream(ms);
        }

        public string GetModuleName(int ModuleID)
        {
            return ModulesDataTable.Select("ModuleID = " + ModuleID)[0]["ModuleName"].ToString();
        }

        public int GetLastModuleUpdate(ref int ModuleID, ref int MoreCount)
        {
            ModuleID = Convert.ToInt32(ModulesUpdatesDataTable.Rows[0]["ModuleID"]);

            if (ModulesUpdatesDataTable.Rows.Count > 1)
            {
                for (var i = 1; i < ModulesUpdatesDataTable.Rows.Count; i++)
                {
                    MoreCount += Convert.ToInt32(ModulesUpdatesDataTable.Rows[i]["Count"]);
                }
            }

            return Convert.ToInt32(ModulesUpdatesDataTable.Rows[0]["Count"]);
        }

        public bool CheckLastUpdate()
        {
            using (var DA = new SqlDataAdapter("SELECT TOP 1 SubscribesRecordID, SubscribesItemID FROM SubscribesRecords WHERE UserTypeID = 0 AND UserID = " +
                                               Security.CurrentUserID + " ORDER BY SubscribesRecordID DESC", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    if (DT.Rows.Count == 0)
                    {
                        iCurrentUpdateID = -1;
                        return false;
                    }

                    iCurrentModuleID = Convert.ToInt32(SubscribesItemsDataTable.Select("SubscribesItemID = " + DT.Rows[0]["SubscribesItemID"])[0]["ModuleID"]);

                    if (Convert.ToInt32(DT.Rows[0]["SubscribesRecordID"]) == iCurrentUpdateID)
                        return false;

                    iCurrentUpdateID = Convert.ToInt32(DT.Rows[0]["SubscribesRecordID"]);
                    return true;
                }
            }
        }

        public void FillUpdates()
        {
            ModulesUpdatesDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT SubscribesRecordID, SubscribesRecords.SubscribesItemID, SubscribesItems.ModuleID FROM SubscribesRecords " +
                                               " INNER JOIN SubscribesItems ON SubscribesItems.SubscribesItemID = SubscribesRecords.SubscribesItemID WHERE UserTypeID = 0 AND UserID = " +
                                               Security.CurrentUserID + " ORDER BY SubscribesRecordID DESC", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    foreach (DataRow Row in DT.Rows)
                    {
                        var Rows = ModulesUpdatesDataTable.Select("ModuleID = " + Row["ModuleID"]);

                        if (Rows.Count() > 0)
                        {
                            Rows[0]["Count"] = Convert.ToInt32(Rows[0]["Count"]) + 1;
                        }
                        else
                        {
                            var NewRow = ModulesUpdatesDataTable.NewRow();
                            NewRow["ModuleID"] = Row["ModuleID"];
                            NewRow["ModuleName"] = ModulesDataTable.Select("ModuleID = " + Row["ModuleID"])[0]["ModuleName"];
                            NewRow["FormName"] = ModulesDataTable.Select("ModuleID = " + Row["ModuleID"])[0]["FormName"];
                            NewRow["Count"] = 1;
                            ModulesUpdatesDataTable.Rows.Add(NewRow);
                        }
                    }
                }
            }
        }

        public static void ClearSubscribesRecords(int UserID, string FormName)
        {
            using (var sDA = new SqlDataAdapter("DELETE FROM SubscribesRecords WHERE SubscribesItemID IN (SELECT SubscribesItemID FROM SubscribesItems WHERE ModuleID IN (SELECT ModuleID FROM infiniu2_users.dbo.Modules WHERE FormName = '" + FormName + "')) AND UserTypeID IN (0,3) AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
            {
                using (var sDT = new DataTable())
                {
                    sDA.Fill(sDT);
                }
            }
        }

        public static void DeleteSubscribesRecord(int SubscribesItemID)
        {
            using (var sDA = new SqlDataAdapter("DELETE FROM SubscribesRecords WHERE SubscribesItemID = " + SubscribesItemID, ConnectionStrings.LightConnectionString))
            {
                using (var sDT = new DataTable())
                {
                    sDA.Fill(sDT);
                }
            }
        }

        public static void CreateSubscribeRecord(string ButtonName, int TableItemID)
        {
            using (var DA = new SqlDataAdapter("SELECT ModuleID FROM Modules WHERE ModuleButtonName = '" + ButtonName + "'", ConnectionStrings.UsersConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    using (var sDA = new SqlDataAdapter("SELECT TOP 0 * FROM SubscribesRecords", ConnectionStrings.LightConnectionString))
                    {
                        using (var CB = new SqlCommandBuilder(sDA))
                        {
                            using (var sDT = new DataTable())
                            {
                                sDA.Fill(sDT);

                                using (var uDA = new SqlDataAdapter("SELECT UserID FROM Subscribers WHERE ModuleID = " + DT.Rows[0]["ModuleID"], ConnectionStrings.LightConnectionString))
                                {
                                    using (var uDT = new DataTable())
                                    {
                                        uDA.Fill(uDT);

                                        foreach (DataRow Row in uDT.Rows)
                                        {
                                            var NewRow = sDT.NewRow();
                                            NewRow["ModuleID"] = DT.Rows[0]["ModuleID"];
                                            NewRow["TableItemID"] = TableItemID;
                                            NewRow["UserID"] = Row["UserID"];
                                            NewRow["UserTypeID"] = 0;
                                            sDT.Rows.Add(NewRow);
                                        }

                                        sDA.Update(sDT);
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }


        public static void CreateSubscribeRecord(int SubscribesItemID, int TableItemID, int CurrentUserID)
        {
            using (var sDA = new SqlDataAdapter("SELECT TOP 0 * FROM SubscribesRecords", ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(sDA))
                {
                    using (var sDT = new DataTable())
                    {
                        sDA.Fill(sDT);

                        using (var uDA = new SqlDataAdapter("SELECT UserID FROM Users WHERE Fired <> 1", ConnectionStrings.UsersConnectionString))
                        {
                            using (var uDT = new DataTable())
                            {
                                uDA.Fill(uDT);

                                foreach (DataRow Row in uDT.Rows)
                                {
                                    if (Convert.ToInt32(Row["UserID"]) == CurrentUserID)
                                        continue;

                                    var NewRow = sDT.NewRow();
                                    NewRow["SubscribesItemID"] = SubscribesItemID;
                                    NewRow["TableItemID"] = TableItemID;
                                    NewRow["UserID"] = Row["UserID"];
                                    NewRow["UserTypeID"] = 0;
                                    sDT.Rows.Add(NewRow);
                                }
                            }
                        }

                        sDA.Update(sDT);
                    }
                }
            }
        }


        public static void CreateSubscribeRecord(string ButtonName, int TableItemID, int CurrentUserID)
        {
            using (var DA = new SqlDataAdapter("SELECT ModuleID FROM Modules WHERE ModuleButtonName = '" + ButtonName + "'", ConnectionStrings.UsersConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    using (var sDA = new SqlDataAdapter("SELECT TOP 0 * FROM SubscribesRecords", ConnectionStrings.LightConnectionString))
                    {
                        using (var CB = new SqlCommandBuilder(sDA))
                        {
                            using (var sDT = new DataTable())
                            {
                                sDA.Fill(sDT);

                                using (var uDA = new SqlDataAdapter("SELECT UserID FROM Subscribers WHERE ModuleID = " + DT.Rows[0]["ModuleID"], ConnectionStrings.LightConnectionString))
                                {
                                    using (var uDT = new DataTable())
                                    {
                                        uDA.Fill(uDT);

                                        foreach (DataRow Row in uDT.Rows)
                                        {
                                            if (Convert.ToInt32(Row["UserID"]) == CurrentUserID)
                                                continue;

                                            var NewRow = sDT.NewRow();
                                            NewRow["ModuleID"] = DT.Rows[0]["ModuleID"];
                                            NewRow["TableItemID"] = TableItemID;
                                            NewRow["UserID"] = Row["UserID"];
                                            NewRow["UserTypeID"] = 0;
                                            sDT.Rows.Add(NewRow);
                                        }

                                        sDA.Update(sDT);
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        public static void CreateSubscribeRecordForOneUser(string ButtonName, int TableItemID, int RecipientUserID)
        {
            using (var DA = new SqlDataAdapter("SELECT ModuleID FROM Modules WHERE ModuleButtonName = '" + ButtonName + "'", ConnectionStrings.UsersConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    using (var sDA = new SqlDataAdapter("SELECT TOP 0 * FROM SubscribesRecords", ConnectionStrings.LightConnectionString))
                    {
                        using (var CB = new SqlCommandBuilder(sDA))
                        {
                            using (var sDT = new DataTable())
                            {
                                sDA.Fill(sDT);

                                var NewRow = sDT.NewRow();
                                NewRow["ModuleID"] = DT.Rows[0]["ModuleID"];
                                NewRow["TableItemID"] = TableItemID;
                                NewRow["UserID"] = RecipientUserID;
                                NewRow["UserTypeID"] = 0;
                                sDT.Rows.Add(NewRow);

                                sDA.Update(sDT);
                            }
                        }
                    }
                }
            }
        }
        public static void CreateSubscribeRecordForOneUser(int SubscribesItemID, int TableItemID, int RecipientUserID)
        {
            using (var sDA = new SqlDataAdapter("SELECT TOP 0 * FROM SubscribesRecords", ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(sDA))
                {
                    using (var sDT = new DataTable())
                    {
                        sDA.Fill(sDT);

                        var NewRow = sDT.NewRow();
                        NewRow["SubscribesItemID"] = SubscribesItemID;
                        NewRow["TableItemID"] = TableItemID;
                        NewRow["UserID"] = RecipientUserID;
                        NewRow["UserTypeID"] = 0;
                        sDT.Rows.Add(NewRow);

                        sDA.Update(sDT);
                    }
                }
            }
        }
    }





    //public class LightWorkDay
    //{
    //    InfiniumFunctionsContainer InfiniumFunctionsContainer;

    //    public DataTable MyFunctionDataTable;
    //    DataTable WorkDayDetailsDT;
    //    DataTable FunctionsDT;
    //    DataTable UserFunctionsDT;
    //    DataTable FunctionsExecTypesDT;
    //    DataTable DepartmentsDT;


    //    public DataTable WorkDaysDataTable;

    //    public int sDayNotStarted = 0;
    //    public int sDayStarted = 1;
    //    public int sBreakStarted = 2;
    //    public int sDayContinued = 3;
    //    public int sDayEnded = 4;
    //    public int sDayNotSaved = 5;

    //    private int CurrentWorkDayID;

    //    public string Comments;

    //    public LightWorkDay()
    //    {
    //        CreateDataTable();
    //        FillDataTable();
    //    }

    //    public void RefInfiniumFunctionsContainer(ref InfiniumFunctionsContainer tInfiniumFunctionsContainer)
    //    {
    //        InfiniumFunctionsContainer = tInfiniumFunctionsContainer;
    //    }

    //    private void CreateDataTable()
    //    {
    //        MyFunctionDataTable = new DataTable();
    //        MyFunctionDataTable.Columns.Add(new DataColumn("FunctionID", Type.GetType("System.Int32")));
    //        MyFunctionDataTable.Columns.Add(new DataColumn("FunctionName", Type.GetType("System.String")));
    //        MyFunctionDataTable.Columns.Add(new DataColumn("DepartmentName", Type.GetType("System.String")));
    //        MyFunctionDataTable.Columns.Add(new DataColumn("ExecType", Type.GetType("System.String")));
    //        MyFunctionDataTable.Columns.Add(new DataColumn("FunctionExecTypeID", Type.GetType("System.Int64")));
    //        MyFunctionDataTable.Columns.Add(new DataColumn("Hours", Type.GetType("System.Int32")));
    //        MyFunctionDataTable.Columns.Add(new DataColumn("Minutes", Type.GetType("System.Int32")));
    //        MyFunctionDataTable.Columns.Add(new DataColumn("IsComplete", Type.GetType("System.Boolean")));
    //    }

    //    private void FillDataTable()
    //    {
    //        WorkDayDetailsDT = new DataTable();
    //        using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM WorkDayDetails WHERE UserID = " + Security.CurrentUserID + " AND WorkDayID = " + CurrentWorkDayID, ConnectionStrings.LightConnectionString))
    //        {
    //            DA.Fill(WorkDayDetailsDT);
    //        }

    //        FunctionsDT = new DataTable();
    //        using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM Functions WHERE FunctionID IN (SELECT FunctionID FROM UserFunctions WHERE UserID = " + Security.CurrentUserID + ")", ConnectionStrings.LightConnectionString))
    //        {
    //            DA.Fill(FunctionsDT);
    //        }

    //        UserFunctionsDT = new DataTable();
    //        using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM UserFunctions WHERE UserID = " + Security.CurrentUserID, ConnectionStrings.LightConnectionString))
    //        {
    //            DA.Fill(UserFunctionsDT);
    //        }

    //        FunctionsExecTypesDT = new DataTable();
    //        using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM FunctionsExecTypes", ConnectionStrings.LightConnectionString))
    //        {
    //            DA.Fill(FunctionsExecTypesDT);
    //        }

    //        DepartmentsDT = new DataTable();
    //        using (SqlDataAdapter DA = new SqlDataAdapter("SELECT DepartmentID, DepartmentName FROM Departments", ConnectionStrings.LightConnectionString))
    //        {
    //            DA.Fill(DepartmentsDT);
    //        }
    //    }

    //    public void UpdateFunctions()
    //    {
    //        using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM Functions WHERE FunctionID IN (SELECT FunctionID FROM UserFunctions WHERE UserID = " + Security.CurrentUserID + ")", ConnectionStrings.LightConnectionString))
    //        {
    //            FunctionsDT.Clear();
    //            DA.Fill(FunctionsDT);
    //        }

    //        using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM UserFunctions WHERE UserID = " + Security.CurrentUserID, ConnectionStrings.LightConnectionString))
    //        {
    //            UserFunctionsDT.Clear();
    //            DA.Fill(UserFunctionsDT);
    //        }
    //    }

    //    public void FillMyFunctionDataTable()
    //    {
    //        using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM WorkDayDetails WHERE UserID = " + Security.CurrentUserID + " AND WorkDayID = " + CurrentWorkDayID, ConnectionStrings.LightConnectionString))
    //        {
    //            WorkDayDetailsDT.Clear();
    //            DA.Fill(WorkDayDetailsDT);

    //            DataRow[] rows = WorkDayDetailsDT.Select("FunctionID = 0");
    //            if (rows.Count() > 0)
    //                Comments = rows[0]["Comments"].ToString();
    //        }

    //        MyFunctionDataTable.Clear();

    //        foreach (DataRow Row in FunctionsDT.Rows)
    //        {
    //            DataRow NewRow = MyFunctionDataTable.NewRow();
    //            NewRow["FunctionID"] = Row["FunctionID"];
    //            NewRow["FunctionName"] = Row["FunctionName"];
    //            NewRow["DepartmentName"] = DepartmentsDT.Select("DepartmentID = " + Row["DepartmentID"])[0]["DepartmentName"];
    //            NewRow["FunctionExecTypeID"] = UserFunctionsDT.Select("FunctionID = " + Row["FunctionID"])[0]["FunctionExecTypeID"];
    //            NewRow["ExecType"] = FunctionsExecTypesDT.Select("FunctionExecTypeID = " + NewRow["FunctionExecTypeID"])[0]["FunctionExecType"];

    //            if (WorkDayDetailsDT.Select("FunctionID = " + Row["FunctionID"]).Count() == 0)
    //            {
    //                NewRow["Hours"] = 0;
    //                NewRow["Minutes"] = 0;
    //                NewRow["IsComplete"] = false;
    //            }
    //            else
    //            {
    //                NewRow["Hours"] = Convert.ToInt32(Convert.ToInt32(WorkDayDetailsDT.Select("FunctionID = " + Row["FunctionID"])[0]["Minutes"]) / 60);
    //                NewRow["Minutes"] = Convert.ToInt32(WorkDayDetailsDT.Select("FunctionID = " + Row["FunctionID"])[0]["Minutes"])
    //                                    - Convert.ToInt32(NewRow["Hours"]) * 60;
    //                NewRow["IsComplete"] = WorkDayDetailsDT.Select("FunctionID = " + Row["FunctionID"])[0]["IsComplete"];
    //            }                

    //            MyFunctionDataTable.Rows.Add(NewRow);
    //        }

    //        ////другие работы
    //        //if (MyFunctionDataTable.Select("FunctionID = 0").Count() == 0)
    //        //{
    //        //    DataRow NewRow = MyFunctionDataTable.NewRow();
    //        //    NewRow["FunctionID"] = 0;
    //        //    NewRow["FunctionName"] = "Другие работы";
    //        //    NewRow["DepartmentName"] = Security.CurrentUserName;
    //        //    NewRow["FunctionExecTypeID"] = 1;
    //        //    NewRow["ExecType"] = "Ежедневно";

    //        //    if (WorkDayDetailsDT.Select("FunctionID = 0").Count() == 0)
    //        //    {
    //        //        NewRow["Hours"] = 0;
    //        //        NewRow["Minutes"] = 0;
    //        //        NewRow["IsComplete"] = false;
    //        //    }
    //        //    else
    //        //    {
    //        //        NewRow["Hours"] = Convert.ToInt32(Convert.ToInt32(WorkDayDetailsDT.Select("FunctionID = 0")[0]["Minutes"]) / 60);
    //        //        NewRow["Minutes"] = Convert.ToInt32(WorkDayDetailsDT.Select("FunctionID = 0")[0]["Minutes"])
    //        //                            - Convert.ToInt32(NewRow["Hours"]) * 60;
    //        //        NewRow["IsComplete"] = WorkDayDetailsDT.Select("FunctionID = 0")[0]["IsComplete"];
    //        //    }                

    //        //    MyFunctionDataTable.Rows.Add(NewRow);
    //        //}

    //        InfiniumFunctionsContainer.FunctionsDataTable = MyFunctionDataTable.Copy();
    //    }

    //    public static bool StartWorkDay(int UserID)
    //    {
    //        using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM WorkDays WHERE CAST(DayStartDateTime AS DATE) = CAST(GetDate() AS DATE) AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
    //        {
    //            using (SqlCommandBuilder CB = new SqlCommandBuilder(DA))
    //            {
    //                using (DataTable DT = new DataTable())
    //                {
    //                    if (DA.Fill(DT) > 0)
    //                        return false;

    //                    DataRow NewRow = DT.NewRow();
    //                    NewRow["UserID"] = UserID;
    //                    NewRow["DayStartDateTime"] = Security.GetCurrentDate();
    //                    NewRow["DayStartFactDateTime"] = NewRow["DayStartDateTime"];
    //                    DT.Rows.Add(NewRow);


    //                    DA.Update(DT);

    //                    return true;
    //                }
    //            }
    //        }
    //    }

    //    public static bool StartWorkDay(int UserID, DateTime DateTime, string Notes)
    //    {
    //        using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM WorkDays WHERE CAST(DayStartDateTime AS DATE) = CAST(GetDate() AS DATE) AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
    //        {
    //            using (SqlCommandBuilder CB = new SqlCommandBuilder(DA))
    //            {
    //                using (DataTable DT = new DataTable())
    //                {
    //                    if (DA.Fill(DT) > 0)
    //                        return false;

    //                    DataRow NewRow = DT.NewRow();
    //                    NewRow["UserID"] = UserID;
    //                    NewRow["DayStartDateTime"] = DateTime;
    //                    NewRow["DayStartFactDateTime"] = Security.GetCurrentDate();
    //                    NewRow["DayStartNotes"] = Notes;
    //                    DT.Rows.Add(NewRow);

    //                    DA.Update(DT);

    //                    return true;
    //                }
    //            }
    //        }
    //    }


    //    public static bool BreakStartWorkDay(int UserID)
    //    {
    //        using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM WorkDays WHERE CAST(DayStartDateTime AS DATE) = CAST(GetDate() AS DATE) AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
    //        {
    //            using (SqlCommandBuilder CB = new SqlCommandBuilder(DA))
    //            {
    //                using (DataTable DT = new DataTable())
    //                {
    //                    if (DA.Fill(DT) < 0)
    //                        return false;

    //                    DT.Rows[0]["DayBreakStartDateTime"] = Security.GetCurrentDate();
    //                    DT.Rows[0]["DayBreakStartFactDateTime"] = DT.Rows[0]["DayBreakStartDateTime"];

    //                    DA.Update(DT);

    //                    return true;
    //                }
    //            }
    //        }
    //    }

    //    public static bool BreakStartWorkDay(int UserID, DateTime DateTime, string Notes)
    //    {
    //        using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM WorkDays WHERE CAST(DayStartDateTime AS DATE) = CAST(GetDate() AS DATE) AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
    //        {
    //            using (SqlCommandBuilder CB = new SqlCommandBuilder(DA))
    //            {
    //                using (DataTable DT = new DataTable())
    //                {
    //                    if (DA.Fill(DT) < 0)
    //                        return false;

    //                    DT.Rows[0]["DayBreakStartDateTime"] = DateTime;
    //                    DT.Rows[0]["DayBreakStartFactDateTime"] = Security.GetCurrentDate();
    //                    DT.Rows[0]["DayBreakStartNotes"] = Notes;

    //                    DA.Update(DT);

    //                    return true;
    //                }
    //            }
    //        }
    //    }

    //    public static bool BreakStartWorkDay(int UserID, DateTime DateTime, string Notes, bool bOverdued)
    //    {
    //        using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM WorkDays WHERE CAST(DayStartDateTime AS DATE) = '" + DateTime.ToString("yyyy-MM-dd") + "' AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
    //        {
    //            using (SqlCommandBuilder CB = new SqlCommandBuilder(DA))
    //            {
    //                using (DataTable DT = new DataTable())
    //                {
    //                    if (DA.Fill(DT) < 0)
    //                        return false;

    //                    DT.Rows[0]["DayBreakStartDateTime"] = DateTime;
    //                    DT.Rows[0]["DayBreakStartFactDateTime"] = Security.GetCurrentDate();
    //                    DT.Rows[0]["DayBreakStartNotes"] = Notes;

    //                    DA.Update(DT);

    //                    return true;
    //                }
    //            }
    //        }
    //    }


    //    public static bool IsDayOverdued(int UserID, ref DateTime ActDateTime)
    //    {
    //        using (SqlDataAdapter DA = new SqlDataAdapter("SELECT TOP 1 * FROM WorkDays WHERE CAST(DayStartDateTime AS DATE) < CAST(GetDATE() AS DATE) AND DayEndDateTime IS NULL AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
    //        {
    //            using (SqlCommandBuilder CB = new SqlCommandBuilder(DA))
    //            {
    //                using (DataTable DT = new DataTable())
    //                {
    //                    DA.Fill(DT);

    //                    if (DT.Rows.Count > 0)
    //                    {
    //                        DT.Rows[0]["DayEndDateTime"] = Security.GetCurrentDate();
    //                        DT.Rows[0]["DayEndFactDateTime"] = DT.Rows[0]["DayEndDateTime"];

    //                        ActDateTime = Convert.ToDateTime(DT.Rows[0]["DayStartDateTime"]);
    //                        return true;
    //                    }
    //                    else
    //                        return false;
    //                }
    //            }
    //        }
    //    }


    //    public static void EndWorkDay(int UserID)
    //    {
    //        using (SqlDataAdapter DA = new SqlDataAdapter("SELECT TOP 1 * FROM WorkDays WHERE CAST(DayStartDateTime AS DATE) < CAST(GetDATE() AS DATE) AND DayEndDateTime IS NULL AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
    //        {
    //            using (SqlCommandBuilder CB = new SqlCommandBuilder(DA))
    //            {
    //                using (DataTable DT = new DataTable())
    //                {
    //                    DA.Fill(DT);

    //                    if (DT.Rows.Count > 0)
    //                    {
    //                        DT.Rows[0]["DayEndDateTime"] = Security.GetCurrentDate();
    //                        DT.Rows[0]["DayEndFactDateTime"] = DT.Rows[0]["DayEndDateTime"];

    //                        DA.Update(DT);

    //                        return;
    //                    }
    //                }
    //            }
    //        }


    //        using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM WorkDays WHERE CAST(DayStartDateTime AS DATE) = CAST(GetDate() AS DATE) AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
    //        {
    //            using (SqlCommandBuilder CB = new SqlCommandBuilder(DA))
    //            {
    //                using (DataTable DT = new DataTable())
    //                {
    //                    if (DA.Fill(DT) < 0)
    //                        return;

    //                    DT.Rows[0]["UserID"] = UserID;
    //                    DT.Rows[0]["DayEndDateTime"] = Security.GetCurrentDate();
    //                    DT.Rows[0]["DayEndFactDateTime"] = DT.Rows[0]["DayEndDateTime"];

    //                    DA.Update(DT);

    //                    return;
    //                }
    //            }
    //        }

    //    }

    //    public static void EndWorkDay(int UserID, DateTime DateTime, string Notes)
    //    {
    //        using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM WorkDays WHERE CAST(DayStartDateTime AS DATE) = CAST(GetDate() AS DATE) AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
    //        {
    //            using (SqlCommandBuilder CB = new SqlCommandBuilder(DA))
    //            {
    //                using (DataTable DT = new DataTable())
    //                {
    //                    if (DA.Fill(DT) < 0)
    //                        return;

    //                    DT.Rows[0]["DayEndDateTime"] = DateTime;
    //                    DT.Rows[0]["DayEndFactDateTime"] = Security.GetCurrentDate();
    //                    DT.Rows[0]["DayEndNotes"] = Notes;
    //                    DA.Update(DT);

    //                    return;
    //                }
    //            }
    //        }

    //    }

    //    public static void EndWorkDay(int UserID, DateTime DateTime, string Notes, bool bOverdued)
    //    {
    //        using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM WorkDays WHERE CAST(DayStartDateTime AS DATE) = '" + DateTime.ToString("yyyy-MM-dd") + "'AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
    //        {
    //            using (SqlCommandBuilder CB = new SqlCommandBuilder(DA))
    //            {
    //                using (DataTable DT = new DataTable())
    //                {
    //                    if (DA.Fill(DT) < 0)
    //                        return;

    //                    DT.Rows[0]["DayEndDateTime"] = DateTime;
    //                    DT.Rows[0]["DayEndFactDateTime"] = Security.GetCurrentDate();
    //                    DT.Rows[0]["DayEndNotes"] = Notes;
    //                    DA.Update(DT);

    //                    return;
    //                }
    //            }
    //        }

    //    }


    //    public static void ContinueWorkDay(int UserID)
    //    {
    //        using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM WorkDays WHERE CAST(DayStartDateTime AS DATE) = CAST(GetDate() AS DATE) AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
    //        {
    //            using (SqlCommandBuilder CB = new SqlCommandBuilder(DA))
    //            {
    //                using (DataTable DT = new DataTable())
    //                {
    //                    if (DA.Fill(DT) < 0)
    //                        return;

    //                    DT.Rows[0]["DayBreakEndDateTime"] = Security.GetCurrentDate();
    //                    DT.Rows[0]["DayBreakEndFactDateTime"] = DT.Rows[0]["DayBreakEndDateTime"];

    //                    DA.Update(DT);

    //                    return;
    //                }
    //            }
    //        }
    //    }

    //    public static void ContinueWorkDay(int UserID, DateTime DateTime, string Notes)
    //    {
    //        using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM WorkDays WHERE CAST(DayStartDateTime AS DATE) = CAST(GetDate() AS DATE) AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
    //        {
    //            using (SqlCommandBuilder CB = new SqlCommandBuilder(DA))
    //            {
    //                using (DataTable DT = new DataTable())
    //                {
    //                    if (DA.Fill(DT) < 0)
    //                        return;

    //                    DT.Rows[0]["DayBreakEndDateTime"] = DateTime;
    //                    DT.Rows[0]["DayBreakEndFactDateTime"] = Security.GetCurrentDate();
    //                    DT.Rows[0]["DayContinueNotes"] = Notes;

    //                    DA.Update(DT);

    //                    return;
    //                }
    //            }
    //        }
    //    }

    //    public static void ContinueWorkDay(int UserID, DateTime DateTime, string Notes, bool bOverdued)
    //    {
    //        using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM WorkDays WHERE CAST(DayStartDateTime AS DATE) = '" + DateTime.ToString("yyyy-MM-dd") + "'AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
    //        {
    //            using (SqlCommandBuilder CB = new SqlCommandBuilder(DA))
    //            {
    //                using (DataTable DT = new DataTable())
    //                {
    //                    if (DA.Fill(DT) < 0)
    //                        return;

    //                    DT.Rows[0]["DayBreakEndDateTime"] = DateTime;
    //                    DT.Rows[0]["DayBreakEndFactDateTime"] = Security.GetCurrentDate();
    //                    DT.Rows[0]["DayContinueNotes"] = Notes;

    //                    DA.Update(DT);

    //                    return;
    //                }
    //            }
    //        }
    //    }

    //    public void SaveCurrentWorkDay()
    //    {
    //        using (SqlDataAdapter DA = new SqlDataAdapter("SELECT TOP 1 * FROM WorkDays WHERE WorkDayID = " + CurrentWorkDayID, ConnectionStrings.LightConnectionString))
    //        {
    //            using (SqlCommandBuilder CB = new SqlCommandBuilder(DA))
    //            {
    //                using (DataTable DT = new DataTable())
    //                {
    //                    DA.Fill(DT);
    //                    if (DT.Rows[0]["DayEndDateTime"] != DBNull.Value)
    //                    {
    //                        DT.Rows[0]["Saved"] = true;
    //                        DA.Update(DT);
    //                        CurrentWorkDayID = 0;
    //                    }
    //                }
    //            }
    //        }
    //    }

    //    public string SaveMyFunction(bool Saved)
    //    {
    //        using (DataTable newWorkDayDetailsDT = WorkDayDetailsDT.Clone())
    //        {
    //            foreach (DataRow Row in InfiniumFunctionsContainer.FunctionsDataTable.Rows)
    //            {
    //                if (Convert.ToBoolean(Row["IsComplete"]) && Convert.ToInt32(Row["Minutes"]) == 0 && Convert.ToInt32(Row["Hours"]) == 0)
    //                    return "Обязанность не может быть завершенной,\nесли на неё не затрачено время.";

    //                if (Saved && !Convert.ToBoolean(Row["IsComplete"]) && (Convert.ToInt32(Row["Minutes"]) != 0 || Convert.ToInt32(Row["Hours"]) != 0))
    //                    return "После завершения рабочего дня,\nвсе обязанности, на которые было затрачено\nвремя, должны быть завершены.";

    //                if (Saved && !Convert.ToBoolean(Row["IsComplete"]))
    //                    continue;

    //                DataRow NewRow = newWorkDayDetailsDT.NewRow();
    //                NewRow["WorkDayID"] = CurrentWorkDayID;
    //                NewRow["UserID"] = Security.CurrentUserID;
    //                NewRow["FunctionID"] = Row["FunctionID"];
    //                NewRow["Minutes"] = Convert.ToInt32(Row["Hours"]) * 60 + Convert.ToInt32(Row["Minutes"]);
    //                NewRow["IsComplete"] = Row["IsComplete"];

    //                if (Convert.ToInt32(Row["FunctionID"]) == 0)
    //                {
    //                    NewRow["Comments"] = Comments;
    //                }

    //                newWorkDayDetailsDT.Rows.Add(NewRow);
    //            }

    //            using (SqlDataAdapter DA = new SqlDataAdapter("DELETE FROM WorkDayDetails WHERE UserID = " + Security.CurrentUserID + " AND WorkDayID = " + CurrentWorkDayID, ConnectionStrings.LightConnectionString))
    //            {
    //                DA.Fill(WorkDayDetailsDT);
    //            }

    //            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM WorkDayDetails", ConnectionStrings.LightConnectionString))
    //            {
    //                using (SqlCommandBuilder CB = new SqlCommandBuilder(DA))
    //                {
    //                    DA.Update(newWorkDayDetailsDT);
    //                }
    //            }

    //            return "";
    //        }
    //    }

    //    public DayStatus GetDayStatus(int UserID)
    //    {
    //        DayStatus DS = new DayStatus();

    //        using (SqlDataAdapter DA = new SqlDataAdapter("SELECT TOP 1 * FROM WorkDays WHERE CAST(DayStartDateTime AS DATE) = CAST(GETDATE() AS DATE) AND DayEndDateTime IS NOT NULL AND  UserID = " + UserID + " ORDER BY WorkDayID DESC", ConnectionStrings.LightConnectionString))
    //        {
    //            using (DataTable DT = new DataTable())
    //            {
    //                if (DA.Fill(DT) > 0)
    //                {
    //                    DS.DayStarted = Convert.ToDateTime(DT.Rows[0]["DayStartDateTime"]);

    //                    if (DT.Rows[0]["DayBreakStartDateTime"] != DBNull.Value)
    //                    {
    //                        DS.BreakStarted = Convert.ToDateTime(DT.Rows[0]["DayBreakStartDateTime"]);
    //                        DS.BreakEnded = Convert.ToDateTime(DT.Rows[0]["DayBreakEndDateTime"]);
    //                    }

    //                    DS.DayEnded = Convert.ToDateTime(DT.Rows[0]["DayEndDateTime"]);

    //                    DS.iDayStatus = sDayEnded;

    //                    if (DT.Rows[0]["DayBreakStartDateTime"] != DBNull.Value)
    //                        DS.bBreak = true;
    //                    else
    //                        DS.bBreak = false;

    //                    if (!Convert.ToBoolean(DT.Rows[0]["Saved"]))
    //                        DS.iDayStatus = sDayNotSaved;

    //                    CurrentWorkDayID = Convert.ToInt32(DT.Rows[0]["WorkDayID"]);

    //                    return DS;
    //                }
    //            }
    //        }

    //        using (SqlDataAdapter DA = new SqlDataAdapter("SELECT TOP 1 * FROM WorkDays WHERE Saved = 0 AND UserID = " + UserID + " ORDER BY WorkDayID DESC", ConnectionStrings.LightConnectionString))
    //        {
    //            using (DataTable DT = new DataTable())
    //            {
    //                DA.Fill(DT);

    //                if (DT.Rows.Count != 0)
    //                {
    //                    DS.iDayStatus = sDayStarted;

    //                    DS.DayStarted = Convert.ToDateTime(DT.Rows[0]["DayStartDateTime"]);

    //                    if (DT.Rows[0]["DayBreakStartDateTime"] != DBNull.Value)
    //                    {
    //                        DS.iDayStatus = sBreakStarted;
    //                        DS.BreakStarted = Convert.ToDateTime(DT.Rows[0]["DayBreakStartDateTime"]);
    //                        DS.bBreak = true;
    //                    }
    //                    else
    //                        DS.bBreak = false;

    //                    if (DT.Rows[0]["DayBreakEndDateTime"] != DBNull.Value)
    //                    {
    //                        DS.iDayStatus = sDayContinued;
    //                        DS.BreakEnded = Convert.ToDateTime(DT.Rows[0]["DayBreakEndDateTime"]);
    //                    }

    //                    if (DT.Rows[0]["DayEndDateTime"] != DBNull.Value)
    //                    {
    //                        DS.iDayStatus = sDayNotSaved;
    //                        DS.DayEnded = Convert.ToDateTime(DT.Rows[0]["DayEndDateTime"]);
    //                    }

    //                    CurrentWorkDayID = Convert.ToInt32(DT.Rows[0]["WorkDayID"]);
    //                }
    //                else
    //                    DS.iDayStatus = sDayNotStarted;
    //            }
    //        }

    //        return DS;
    //    }

    //    public DayFactStatus GetStatusFactTime(ref InfiniumTimeLabel ChangeDayStartLabel, ref InfiniumTimeLabel ChangeBreakStartLabel, ref InfiniumTimeLabel ChangeBreakEndLabel, ref InfiniumTimeLabel ChangeDayEndLabel)
    //    {
    //        ChangeDayStartLabel.ForeColor = Color.Gray;
    //        ChangeDayStartLabel.Text = "без изменений";

    //        ChangeBreakStartLabel.ForeColor = Color.Gray;
    //        ChangeBreakStartLabel.Text = "без изменений";

    //        ChangeBreakEndLabel.ForeColor = Color.Gray;
    //        ChangeBreakEndLabel.Text = "без изменений";

    //        ChangeDayEndLabel.ForeColor = Color.Gray;
    //        ChangeDayEndLabel.Text = "без изменений";

    //        DayFactStatus DFS = new DayFactStatus();


    //        using (SqlDataAdapter DA = new SqlDataAdapter("SELECT TOP 1 * FROM WorkDays WHERE (CAST(DayStartDateTime AS DATE) = CAST(GetDate() AS DATE) OR Saved = 0) AND UserID = " + Security.CurrentUserID + " ORDER BY WorkDayID DESC", ConnectionStrings.LightConnectionString))
    //        {
    //            using (DataTable DT = new DataTable())
    //            {
    //                if (DA.Fill(DT) > 0)
    //                {
    //                    DateTime Time1, Time2;

    //                    if (DT.Rows[0]["DayStartDateTime"] != DBNull.Value)
    //                    {
    //                        Time1 = Convert.ToDateTime(DT.Rows[0]["DayStartDateTime"]);
    //                        Time2 = Convert.ToDateTime(DT.Rows[0]["DayStartFactDateTime"]);
    //                        DFS.DayFactStarted = Time2;

    //                        if (Time1 != Time2)
    //                        {
    //                            ChangeDayStartLabel.ForeColor = Color.Red;
    //                            ChangeDayStartLabel.Text = "с " + Time2.ToString("HH:mm") + " на " + Time1.ToString("HH:mm");
    //                        }
    //                        DFS.DayStartNotes = DT.Rows[0]["DayStartNotes"].ToString();
    //                    }

    //                    if (DT.Rows[0]["DayBreakStartDateTime"] != DBNull.Value)
    //                    {
    //                        Time1 = Convert.ToDateTime(DT.Rows[0]["DayBreakStartDateTime"]);
    //                        Time2 = Convert.ToDateTime(DT.Rows[0]["DayBreakStartFactDateTime"]);
    //                        DFS.BreakFactStarted = Time2;

    //                        if (Time1 != Time2)
    //                        {
    //                            ChangeBreakStartLabel.ForeColor = Color.Red;
    //                            ChangeBreakStartLabel.Text = "с " + Time2.ToString("HH:mm") + " на " + Time1.ToString("HH:mm");
    //                        }
    //                        DFS.DayBreakStartNotes = DT.Rows[0]["DayBreakStartNotes"].ToString();
    //                    }

    //                    if (DT.Rows[0]["DayBreakEndDateTime"] != DBNull.Value)
    //                    {
    //                        Time1 = Convert.ToDateTime(DT.Rows[0]["DayBreakEndDateTime"]);
    //                        Time2 = Convert.ToDateTime(DT.Rows[0]["DayBreakEndFactDateTime"]);
    //                        DFS.BreakFactEnded = Time2;

    //                        if (Time1 != Time2)
    //                        {
    //                            ChangeBreakEndLabel.ForeColor = Color.Red;
    //                            ChangeBreakEndLabel.Text = "с " + Time2.ToString("HH:mm") + " на " + Time1.ToString("HH:mm");
    //                        }
    //                        DFS.DayContinueNotes = DT.Rows[0]["DayContinueNotes"].ToString();
    //                    }

    //                    if (DT.Rows[0]["DayEndDateTime"] != DBNull.Value)
    //                    {
    //                        Time1 = Convert.ToDateTime(DT.Rows[0]["DayEndDateTime"]);
    //                        Time2 = Convert.ToDateTime(DT.Rows[0]["DayEndFactDateTime"]);
    //                        DFS.DayFactEnded = Time2;

    //                        if (Time1 != Time2)
    //                        {
    //                            ChangeDayEndLabel.ForeColor = Color.Red;
    //                            ChangeDayEndLabel.Text = "с " + Time2.ToString("HH:mm") + " на " + Time1.ToString("HH:mm");
    //                        }
    //                        DFS.DayEndNotes = DT.Rows[0]["DayEndNotes"].ToString();
    //                    }
    //                }
    //            }
    //        }
    //        return DFS;
    //    }

    //    public static void GetWeekTimeSheet(ref InfiniumTimeLabel MonTimeSheetlabel, ref InfiniumTimeLabel TueTimeSheetlabel, ref InfiniumTimeLabel WedTimeSheetlabel, ref InfiniumTimeLabel ThuTimeSheetlabel, ref InfiniumTimeLabel FriTimeSheetlabel, ref InfiniumTimeLabel SatTimeSheetlabel, ref InfiniumTimeLabel SunTimeSheetlabel)
    //    {
    //        MonTimeSheetlabel.ForeColor = Color.Gray;
    //        MonTimeSheetlabel.Text = "-- : --";
    //        TueTimeSheetlabel.ForeColor = Color.Gray;
    //        TueTimeSheetlabel.Text = "-- : --";
    //        WedTimeSheetlabel.ForeColor = Color.Gray;
    //        WedTimeSheetlabel.Text = "-- : --";
    //        ThuTimeSheetlabel.ForeColor = Color.Gray;
    //        ThuTimeSheetlabel.Text = "-- : --";
    //        FriTimeSheetlabel.ForeColor = Color.Gray;
    //        FriTimeSheetlabel.Text = "-- : --";
    //        SatTimeSheetlabel.ForeColor = Color.Gray;
    //        SatTimeSheetlabel.Text = "-- : --";
    //        SunTimeSheetlabel.ForeColor = Color.Gray;
    //        SunTimeSheetlabel.Text = "-- : --";

    //        DateTime DayStartDateTime;
    //        DateTime DayEndDateTime;
    //        DateTime DayBreakStartDateTime;
    //        DateTime DayBreakEndDateTime;
    //        TimeSpan TimeWork;

    //        DayStartDateTime = DateTime.Now;
    //        DayStartDateTime = DayStartDateTime.Subtract(new TimeSpan(Convert.ToInt32(DateTime.Now.DayOfWeek - 1), 0, 0, 0));

    //        using (SqlDataAdapter DA = new SqlDataAdapter("SELECT DayStartDateTime, DayEndDateTime, DayBreakStartDateTime, DayBreakEndDateTime FROM WorkDays where DayStartDateTime > '" + DayStartDateTime.ToString("MM-dd-yyyy") + "' and UserId = " + Security.CurrentUserID, ConnectionStrings.LightConnectionString))
    //        {
    //            using (DataTable DT = new DataTable())
    //            {
    //                DA.Fill(DT);
    //                for (int i = 0; i < DT.Rows.Count; i++)
    //                {
    //                    if (DT.Rows[i]["DayStartDateTime"] != DBNull.Value && DT.Rows[i]["DayEndDateTime"] != DBNull.Value)
    //                    {
    //                        DayStartDateTime = (DateTime)DT.Rows[i]["DayStartDateTime"];
    //                        DayEndDateTime = (DateTime)DT.Rows[i]["DayEndDateTime"];
    //                        TimeWork = DayEndDateTime.TimeOfDay - DayStartDateTime.TimeOfDay;
    //                        if (DT.Rows[i]["DayBreakStartDateTime"].ToString() != "" && DT.Rows[i]["DayBreakEndDateTime"].ToString() != "")
    //                        {
    //                            DayBreakStartDateTime = (DateTime)DT.Rows[i]["DayBreakStartDateTime"];
    //                            DayBreakEndDateTime = (DateTime)DT.Rows[i]["DayBreakEndDateTime"];
    //                            TimeWork -= DayBreakEndDateTime.TimeOfDay - DayBreakStartDateTime.TimeOfDay;
    //                        }

    //                        if (TimeWork.Seconds > 29)
    //                        {
    //                            TimeWork.Add(new TimeSpan(0, 1, 0));
    //                        }

    //                        switch ((int)DayStartDateTime.DayOfWeek)
    //                        {
    //                            case 1:
    //                                MonTimeSheetlabel.ForeColor = Color.ForestGreen;
    //                                MonTimeSheetlabel.Text = TimeWork.Hours.ToString() + " ч : " + TimeWork.Minutes.ToString("D2") + " м";
    //                                break;
    //                            case 2:
    //                                TueTimeSheetlabel.ForeColor = Color.ForestGreen;
    //                                TueTimeSheetlabel.Text = TimeWork.Hours.ToString() + " ч : " + TimeWork.Minutes.ToString("D2") + " м";
    //                                break;
    //                            case 3:
    //                                WedTimeSheetlabel.ForeColor = Color.ForestGreen;
    //                                WedTimeSheetlabel.Text = TimeWork.Hours.ToString() + " ч : " + TimeWork.Minutes.ToString("D2") + " м";
    //                                break;
    //                            case 4:
    //                                ThuTimeSheetlabel.ForeColor = Color.ForestGreen;
    //                                ThuTimeSheetlabel.Text = TimeWork.Hours.ToString() + " ч : " + TimeWork.Minutes.ToString("D2") + " м";
    //                                break;
    //                            case 5:
    //                                FriTimeSheetlabel.ForeColor = Color.ForestGreen;
    //                                FriTimeSheetlabel.Text = TimeWork.Hours.ToString() + " ч : " + TimeWork.Minutes.ToString("D2") + " м";
    //                                break;
    //                            case 6:
    //                                SatTimeSheetlabel.ForeColor = Color.ForestGreen;
    //                                SatTimeSheetlabel.Text = TimeWork.Hours.ToString() + " ч : " + TimeWork.Minutes.ToString("D2") + " м";
    //                                break;
    //                            case 7:
    //                                SunTimeSheetlabel.ForeColor = Color.ForestGreen;
    //                                SunTimeSheetlabel.Text = TimeWork.Hours.ToString() + " ч : " + TimeWork.Minutes.ToString("D2") + " м";
    //                                break;
    //                        }
    //                    }
    //                }
    //            }
    //        }
    //    }

    //    public static void ChangeStartTime(int UserID, DateTime NewDateTime, string Notes)
    //    {
    //        using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM WorkDays WHERE CAST(DayStartDateTime AS DATE) = '" + NewDateTime.ToString("yyyy-MM-dd") + "' AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
    //        {
    //            using (SqlCommandBuilder CB = new SqlCommandBuilder(DA))
    //            {
    //                using (DataTable DT = new DataTable())
    //                {
    //                    if (DA.Fill(DT) > 0)
    //                    {
    //                        DT.Rows[0]["DayStartDateTime"] = NewDateTime;

    //                        if (DT.Rows[0]["DayStartDateTime"].ToString() == DT.Rows[0]["DayStartFactDateTime"].ToString())
    //                        {
    //                            DT.Rows[0]["DayStartNotes"] = "";
    //                        }
    //                        else
    //                            DT.Rows[0]["DayStartNotes"] = Notes;

    //                        DA.Update(DT);
    //                    }
    //                }
    //            }
    //        }
    //    }

    //    public static void ChangeEndTime(int UserID, DateTime NewDateTime, string Notes)
    //    {
    //        using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM WorkDays WHERE CAST(DayStartDateTime AS DATE) = '" + NewDateTime.ToString("yyyy-MM-dd") + "' AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
    //        {
    //            using (SqlCommandBuilder CB = new SqlCommandBuilder(DA))
    //            {
    //                using (DataTable DT = new DataTable())
    //                {
    //                    if (DA.Fill(DT) > 0)
    //                    {
    //                        DT.Rows[0]["DayEndDateTime"] = NewDateTime;

    //                        if (DT.Rows[0]["DayEndDateTime"].ToString() == DT.Rows[0]["DayEndFactDateTime"].ToString())
    //                        {
    //                            DT.Rows[0]["DayEndNotes"] = "";
    //                        }
    //                        else
    //                            DT.Rows[0]["DayEndNotes"] = Notes;

    //                        DA.Update(DT);
    //                    }
    //                }
    //            }
    //        }
    //    }

    //    public static void ChangeBreakStartTime(int UserID, DateTime NewDateTime, string Notes)
    //    {
    //        using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM WorkDays WHERE CAST(DayStartDateTime AS DATE) = '" + NewDateTime.ToString("yyyy-MM-dd") + "' AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
    //        {
    //            using (SqlCommandBuilder CB = new SqlCommandBuilder(DA))
    //            {
    //                using (DataTable DT = new DataTable())
    //                {
    //                    if (DA.Fill(DT) > 0)
    //                    {
    //                        DT.Rows[0]["DayBreakStartDateTime"] = NewDateTime;

    //                        if (DT.Rows[0]["DayBreakStartDateTime"].ToString() == DT.Rows[0]["DayBreakStartFactDateTime"].ToString())
    //                        {
    //                            DT.Rows[0]["DayBreakStartNotes"] = "";
    //                        }
    //                        else
    //                            DT.Rows[0]["DayBreakStartNotes"] = Notes;

    //                        DA.Update(DT);
    //                    }
    //                }
    //            }
    //        }
    //    }

    //    public static void ChangeBreakEndTime(int UserID, DateTime NewDateTime, string Notes)
    //    {
    //        using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM WorkDays WHERE CAST(DayStartDateTime AS DATE) = '" + NewDateTime.ToString("yyyy-MM-dd") + "' AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
    //        {
    //            using (SqlCommandBuilder CB = new SqlCommandBuilder(DA))
    //            {
    //                using (DataTable DT = new DataTable())
    //                {
    //                    if (DA.Fill(DT) > 0)
    //                    {
    //                        DT.Rows[0]["DayBreakEndDateTime"] = NewDateTime;

    //                        if (DT.Rows[0]["DayBreakEndDateTime"].ToString() == DT.Rows[0]["DayBreakEndFactDateTime"].ToString())
    //                        {
    //                            DT.Rows[0]["DayContinueNotes"] = "";
    //                        }
    //                        else
    //                            DT.Rows[0]["DayContinueNotes"] = Notes;

    //                        DA.Update(DT);
    //                    }
    //                }
    //            }
    //        }
    //    }
    //}


    public struct DayStatus
    {
        public int iDayStatus;
        public int iWorkDayID;
        public decimal dTimesheetHours;

        public DateTime DayStarted;
        public DateTime BreakStarted;
        public DateTime BreakEnded;
        public DateTime DayEnded;
        public bool bDayStarted;
        public bool bBreakStarted;
        public bool bBreakEnded;
        public bool bDayEnded;

        public bool bBreak;
        public bool bOverdued;
    }

    public struct DayFactStatus
    {
        public DateTime DayFactStarted;
        public DateTime BreakFactStarted;
        public DateTime BreakFactEnded;
        public DateTime DayFactEnded;
        public string DayStartNotes;
        public string DayBreakStartNotes;
        public string DayContinueNotes;
        public string DayEndNotes;
    }



    public class LightWorkDay
    {
        public DataTable MyFunctionDataTable;
        private DataTable WorkDayDetailsDT;
        private DataTable newWorkDayDetailsDT;
        private DataTable FunctionsDT;
        private DataTable UserFunctionsDT;
        private DataTable FunctionsExecTypesDT;
        private DataTable DepartmentsDT;


        public DataTable WorkDaysDataTable;

        public int TotalMin = 0;

        public int sDayNotStarted = 0;
        public int sDayStarted = 1;
        public int sBreakStarted = 2;
        public int sDayContinued = 3;
        public int sDayEnded = 4;
        public int sDayNotSaved = 5;
        public int sDaySaved = 6;

        private int CurrentWorkDayID;

        public string ProfilComments;
        public string TPSComments;

        public LightWorkDay()
        {
            CreateDataTable();
            FillDataTable();
        }

        private void CreateDataTable()
        {
            MyFunctionDataTable = new DataTable();
            MyFunctionDataTable.Columns.Add(new DataColumn("FunctionID", Type.GetType("System.Int32")));
            MyFunctionDataTable.Columns.Add(new DataColumn("FunctionID1", Type.GetType("System.Int32")));
            MyFunctionDataTable.Columns.Add(new DataColumn("FactoryID", Type.GetType("System.Int32")));
            MyFunctionDataTable.Columns.Add(new DataColumn("FunctionName", Type.GetType("System.String")));
            MyFunctionDataTable.Columns.Add(new DataColumn("DepartmentName", Type.GetType("System.String")));
            MyFunctionDataTable.Columns.Add(new DataColumn("ExecType", Type.GetType("System.String")));
            MyFunctionDataTable.Columns.Add(new DataColumn("FunctionExecTypeID", Type.GetType("System.Int64")));
            MyFunctionDataTable.Columns.Add(new DataColumn("Hours", Type.GetType("System.Int32")));
            MyFunctionDataTable.Columns.Add(new DataColumn("Minutes", Type.GetType("System.Int32")));
            MyFunctionDataTable.Columns.Add(new DataColumn("IsComplete", Type.GetType("System.Boolean")));
        }

        private void FillDataTable()
        {
            WorkDayDetailsDT = new DataTable();
            using (var DA = new SqlDataAdapter(@"SELECT WorkDayDetails.*, UserFunctions.FunctionID as FunctionID1 FROM WorkDayDetails 
                INNER JOIN UserFunctions ON WorkDayDetails.FunctionID=UserFunctions.UserFunctionID WHERE WorkDayDetails.UserID = " + Security.CurrentUserID + " AND WorkDayDetails.WorkDayID = " + CurrentWorkDayID, ConnectionStrings.LightConnectionString))
            {
                DA.Fill(WorkDayDetailsDT);
            }
            newWorkDayDetailsDT = WorkDayDetailsDT.Clone();

            //using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM Functions WHERE FunctionID IN (SELECT FunctionID FROM UserFunctions WHERE UserID = " + Security.CurrentUserID + ")", ConnectionStrings.LightConnectionString))
            //{
            //    DA.Fill(FunctionsDT);
            //}

            UserFunctionsDT = new DataTable();

            var SelectCommand = @"SELECT TOP 0 UserFunctions.*, StaffList.FactoryID, StaffList.UserID, Users.Name, StaffList.DepartmentID, StaffList.PositionID, Positions.Position, StaffList.Rate, Factory.FactoryName, Functions.FunctionName, Functions.FunctionDescription FROM UserFunctions
                INNER JOIN StaffList ON UserFunctions.StaffListID=StaffList.StaffListID AND (StaffList.UserID=" + Security.CurrentUserID + @")
                INNER JOIN Functions ON UserFunctions.FunctionID=Functions.FunctionID
                INNER JOIN Positions ON StaffList.PositionID=Positions.PositionID
                INNER JOIN infiniu2_catalog.dbo.Factory AS Factory ON StaffList.FactoryID=Factory.FactoryID
                INNER JOIN infiniu2_users.dbo.Users AS Users ON StaffList.UserID=Users.UserID ORDER BY FunctionName";
            using (var DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.LightConnectionString))
            {
                UserFunctionsDT.Clear();
                DA.Fill(UserFunctionsDT);
            }

            FunctionsDT = new DataTable();
            FunctionsDT = UserFunctionsDT.Clone();

            FunctionsExecTypesDT = new DataTable();
            using (var DA = new SqlDataAdapter("SELECT * FROM FunctionsExecTypes", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(FunctionsExecTypesDT);
            }

            DepartmentsDT = new DataTable();
            using (var DA = new SqlDataAdapter("SELECT DepartmentID, DepartmentName FROM Departments", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(DepartmentsDT);
            }
        }

        public void UpdateFunctions(int FactoryID, int PositionID)
        {

            var SelectCommand = @"SELECT UserFunctions.*, StaffList.FactoryID, StaffList.UserID, Users.Name, StaffList.DepartmentID, StaffList.PositionID, Positions.Position, StaffList.Rate, Factory.FactoryName, Functions.FunctionName, Functions.FunctionDescription FROM UserFunctions
                INNER JOIN StaffList ON UserFunctions.StaffListID=StaffList.StaffListID AND (StaffList.UserID=" + Security.CurrentUserID + @")
                INNER JOIN Functions ON UserFunctions.FunctionID=Functions.FunctionID
                INNER JOIN Positions ON StaffList.PositionID=Positions.PositionID
                INNER JOIN infiniu2_catalog.dbo.Factory AS Factory ON StaffList.FactoryID=Factory.FactoryID
                INNER JOIN infiniu2_users.dbo.Users AS Users ON StaffList.UserID=Users.UserID ORDER BY FunctionName";

            using (var DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.LightConnectionString))
            {
                UserFunctionsDT.Clear();
                DA.Fill(UserFunctionsDT);
            }
            FunctionsDT.Clear();
            var Rows = UserFunctionsDT.Select("FactoryID=" + FactoryID + " AND PositionID=" + PositionID);
            for (var j = 0; j < Rows.Count(); j++)
                FunctionsDT.Rows.Add(Rows[j].ItemArray);
        }

        public void FillMyFunctionDataTable(int FactoryID)
        {
            using (var DA = new SqlDataAdapter(@"SELECT WorkDayDetails.*, UserFunctions.FunctionID as FunctionID1 FROM WorkDayDetails 
                INNER JOIN UserFunctions ON WorkDayDetails.FunctionID=UserFunctions.UserFunctionID WHERE WorkDayDetails.UserID = " + Security.CurrentUserID + " AND WorkDayDetails.WorkDayID = " + CurrentWorkDayID, ConnectionStrings.LightConnectionString))
            {
                WorkDayDetailsDT.Clear();
                DA.Fill(WorkDayDetailsDT);

                if (FactoryID == 1)
                {
                    var rows = WorkDayDetailsDT.Select("FunctionID1 = 0 AND FactoryID=1");
                    if (rows.Count() > 0)
                    {
                        ProfilComments = rows[0]["Comments"].ToString();
                    }
                }
                if (FactoryID == 2)
                {
                    var rows = WorkDayDetailsDT.Select("FunctionID1 = 0 AND FactoryID=2");
                    if (rows.Count() > 0)
                    {
                        TPSComments = rows[0]["Comments"].ToString();
                    }
                }
            }

            MyFunctionDataTable.Clear();

            foreach (DataRow Row in FunctionsDT.Rows)
            {
                var NewRow = MyFunctionDataTable.NewRow();
                NewRow["FunctionID"] = Row["UserFunctionID"];
                NewRow["FunctionID1"] = Row["FunctionID"];
                NewRow["FactoryID"] = FactoryID;
                NewRow["FunctionName"] = Row["FunctionName"];
                NewRow["DepartmentName"] = DepartmentsDT.Select("DepartmentID = " + Row["DepartmentID"])[0]["DepartmentName"];
                NewRow["FunctionExecTypeID"] = UserFunctionsDT.Select("FunctionID = " + Row["FunctionID"])[0]["FunctionExecTypeID"];
                NewRow["ExecType"] = FunctionsExecTypesDT.Select("FunctionExecTypeID = " + NewRow["FunctionExecTypeID"])[0]["FunctionExecType"];

                if (!WorkDayDetailsDT.Select("FunctionID = " + Row["UserFunctionID"]).Any())
                {
                    NewRow["Hours"] = 0;
                    NewRow["Minutes"] = 0;
                    NewRow["IsComplete"] = false;
                }
                else
                {
                    NewRow["Hours"] = Convert.ToInt32(Convert.ToInt32(WorkDayDetailsDT.Select("FunctionID = " + Row["UserFunctionID"])[0]["Minutes"]) / 60);
                    NewRow["Minutes"] = Convert.ToInt32(WorkDayDetailsDT.Select("FunctionID = " + Row["UserFunctionID"])[0]["Minutes"])
                                        - Convert.ToInt32(NewRow["Hours"]) * 60;
                    NewRow["IsComplete"] = WorkDayDetailsDT.Select("FunctionID = " + Row["UserFunctionID"])[0]["IsComplete"];
                }

                MyFunctionDataTable.Rows.Add(NewRow);
            }

            ////другие работы
            //if (MyFunctionDataTable.Select("FunctionID = 0").Count() == 0)
            //{
            //    DataRow NewRow = MyFunctionDataTable.NewRow();
            //    NewRow["FunctionID"] = 0;
            //    NewRow["FunctionName"] = "Другие работы";
            //    NewRow["DepartmentName"] = Security.CurrentUserName;
            //    NewRow["FunctionExecTypeID"] = 1;
            //    NewRow["ExecType"] = "Ежедневно";

            //    if (WorkDayDetailsDT.Select("FunctionID = 0").Count() == 0)
            //    {
            //        NewRow["Hours"] = 0;
            //        NewRow["Minutes"] = 0;
            //        NewRow["IsComplete"] = false;
            //    }
            //    else
            //    {
            //        NewRow["Hours"] = Convert.ToInt32(Convert.ToInt32(WorkDayDetailsDT.Select("FunctionID = 0")[0]["Minutes"]) / 60);
            //        NewRow["Minutes"] = Convert.ToInt32(WorkDayDetailsDT.Select("FunctionID = 0")[0]["Minutes"])
            //                            - Convert.ToInt32(NewRow["Hours"]) * 60;
            //        NewRow["IsComplete"] = WorkDayDetailsDT.Select("FunctionID = 0")[0]["IsComplete"];
            //    }

            //    MyFunctionDataTable.Rows.Add(NewRow);
            //}
        }

        public DataTable MyFunctionDataTableDT => MyFunctionDataTable.Copy();

        public bool IsTimesheetHoursSaved(int UserID)
        {
            var b = false;
            var SelectCommand = @"SELECT TOP 1 * FROM WorkDays WHERE UserID=" + UserID + " ORDER BY DayStartDateTime DESC";
            using (var DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    if (DA.Fill(DT) > 0)
                    {
                        if (Convert.ToDecimal(DT.Rows[0]["TimesheetHours"]) > 0)
                            b = true;
                    }
                    else
                        b = true;
                }
            }
            return b;
        }

        public static bool StartWorkDay(int UserID)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM WorkDays WHERE CAST(DayStartDateTime AS DATE) = CAST(GetDate() AS DATE) AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
            {
                using (new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        if (DA.Fill(DT) > 0)
                            return false;

                        var NewRow = DT.NewRow();
                        NewRow["UserID"] = UserID;
                        NewRow["DayStartDateTime"] = Security.GetCurrentDate();
                        NewRow["DayStartFactDateTime"] = NewRow["DayStartDateTime"];
                        DT.Rows.Add(NewRow);


                        DA.Update(DT);

                        return true;
                    }
                }
            }
        }

        public static bool StartWorkDay(int UserID, DateTime DateTime, string Notes)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM WorkDays WHERE CAST(DayStartDateTime AS DATE) = CAST(GetDate() AS DATE) AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
            {
                using (new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        if (DA.Fill(DT) > 0)
                            return false;

                        var NewRow = DT.NewRow();
                        NewRow["UserID"] = UserID;
                        NewRow["DayStartDateTime"] = DateTime;
                        NewRow["DayStartFactDateTime"] = Security.GetCurrentDate();
                        NewRow["DayStartNotes"] = Notes;
                        DT.Rows.Add(NewRow);

                        DA.Update(DT);

                        return true;
                    }
                }
            }
        }


        public static bool BreakStartWorkDay(int UserID)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM WorkDays WHERE CAST(DayStartDateTime AS DATE) = CAST(GetDate() AS DATE) AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
            {
                using (new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        if (DA.Fill(DT) < 1)
                            return false;

                        DT.Rows[0]["DayBreakStartDateTime"] = Security.GetCurrentDate();
                        DT.Rows[0]["DayBreakStartFactDateTime"] = DT.Rows[0]["DayBreakStartDateTime"];

                        DA.Update(DT);

                        return true;
                    }
                }
            }
        }

        public static bool BreakStartWorkDay(int UserID, DateTime DateTime, string Notes)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM WorkDays WHERE CAST(DayStartDateTime AS DATE) = CAST(GetDate() AS DATE) AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
            {
                using (new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        if (DA.Fill(DT) < 1)
                            return false;

                        DT.Rows[0]["DayBreakStartDateTime"] = DateTime;
                        DT.Rows[0]["DayBreakStartFactDateTime"] = Security.GetCurrentDate();
                        DT.Rows[0]["DayBreakStartNotes"] = Notes;

                        DA.Update(DT);

                        return true;
                    }
                }
            }
        }

        public static bool BreakStartWorkDay(int UserID, DateTime DateTime, string Notes, bool bOverdued)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM WorkDays WHERE CAST(DayStartDateTime AS DATE) = '" + DateTime.ToString("yyyy-MM-dd") + "' AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
            {
                using (new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        if (DA.Fill(DT) < 1)
                            return false;

                        DT.Rows[0]["DayBreakStartDateTime"] = DateTime;
                        DT.Rows[0]["DayBreakStartFactDateTime"] = Security.GetCurrentDate();
                        DT.Rows[0]["DayBreakStartNotes"] = Notes;

                        DA.Update(DT);

                        return true;
                    }
                }
            }
        }


        public static bool IsDayOverdued(int UserID, ref DateTime ActDateTime)
        {
            using (var DA = new SqlDataAdapter("SELECT TOP 1 * FROM WorkDays WHERE CAST(DayStartDateTime AS DATE) < CAST(GetDATE() AS DATE) AND DayEndDateTime IS NULL AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
            {
                using (new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        if (DT.Rows.Count > 0)
                        {
                            DT.Rows[0]["DayEndDateTime"] = Security.GetCurrentDate();
                            DT.Rows[0]["DayEndFactDateTime"] = DT.Rows[0]["DayEndDateTime"];

                            ActDateTime = Convert.ToDateTime(DT.Rows[0]["DayStartDateTime"]);
                            return true;
                        }

                        return false;
                    }
                }
            }
        }


        public static void EndWorkDay(int UserID)
        {
            using (var DA = new SqlDataAdapter("SELECT TOP 1 * FROM WorkDays WHERE CAST(DayStartDateTime AS DATE) < CAST(GetDATE() AS DATE) AND DayEndDateTime IS NULL AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
            {
                using (new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        if (DT.Rows.Count > 0)
                        {
                            DT.Rows[0]["DayEndDateTime"] = Security.GetCurrentDate();
                            DT.Rows[0]["DayEndFactDateTime"] = DT.Rows[0]["DayEndDateTime"];

                            DA.Update(DT);

                            return;
                        }
                    }
                }
            }


            using (var DA = new SqlDataAdapter("SELECT * FROM WorkDays WHERE CAST(DayStartDateTime AS DATE) = CAST(GetDate() AS DATE) AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
            {
                using (new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        if (DA.Fill(DT) < 1)
                            return;

                        DT.Rows[0]["UserID"] = UserID;
                        DT.Rows[0]["DayEndDateTime"] = Security.GetCurrentDate();
                        DT.Rows[0]["DayEndFactDateTime"] = DT.Rows[0]["DayEndDateTime"];

                        DA.Update(DT);
                    }
                }
            }

        }

        public static void EndWorkDay(int UserID, DateTime DateTime, string Notes)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM WorkDays WHERE CAST(DayStartDateTime AS DATE) = CAST(GetDate() AS DATE) AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
            {
                using (new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        if (DA.Fill(DT) < 1)
                            return;

                        DT.Rows[0]["DayEndDateTime"] = DateTime;
                        DT.Rows[0]["DayEndFactDateTime"] = Security.GetCurrentDate();
                        DT.Rows[0]["DayEndNotes"] = Notes;
                        DA.Update(DT);
                    }
                }
            }

        }

        public static void EndWorkDay(int UserID, DateTime DateTime, string Notes, bool bOverdued)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM WorkDays WHERE CAST(DayStartDateTime AS DATE) = '" + DateTime.ToString("yyyy-MM-dd") + "'AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
            {
                using (new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        if (DA.Fill(DT) < 1)
                            return;

                        DT.Rows[0]["DayEndDateTime"] = DateTime;
                        DT.Rows[0]["DayEndFactDateTime"] = Security.GetCurrentDate();
                        DT.Rows[0]["DayEndNotes"] = Notes;
                        DA.Update(DT);
                    }
                }
            }

        }


        public static void CancelEndWorkDay(int UserID, DateTime DateTime)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM WorkDays WHERE CAST(DayStartDateTime AS DATE) = '" + DateTime.ToString("yyyy-MM-dd") + "' AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
            {
                using (new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        if (DA.Fill(DT) < 1)
                            return;

                        DT.Rows[0]["DayEndDateTime"] = DBNull.Value;
                        DT.Rows[0]["DayEndFactDateTime"] = DBNull.Value;
                        DT.Rows[0]["Saved"] = false;

                        DA.Update(DT);
                    }
                }
            }
        }

        public static void ContinueWorkDay(int UserID)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM WorkDays WHERE CAST(DayStartDateTime AS DATE) = CAST(GetDate() AS DATE) AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
            {
                using (new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        if (DA.Fill(DT) < 1)
                            return;

                        DT.Rows[0]["DayBreakEndDateTime"] = Security.GetCurrentDate();
                        DT.Rows[0]["DayBreakEndFactDateTime"] = DT.Rows[0]["DayBreakEndDateTime"];

                        DA.Update(DT);
                    }
                }
            }
        }

        public static void ContinueWorkDay(int UserID, DateTime DateTime, string Notes)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM WorkDays WHERE CAST(DayStartDateTime AS DATE) = CAST(GetDate() AS DATE) AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
            {
                using (new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        if (DA.Fill(DT) < 1)
                            return;

                        DT.Rows[0]["DayBreakEndDateTime"] = DateTime;
                        DT.Rows[0]["DayBreakEndFactDateTime"] = Security.GetCurrentDate();
                        DT.Rows[0]["DayContinueNotes"] = Notes;

                        DA.Update(DT);
                    }
                }
            }
        }

        public static void ContinueWorkDay(int UserID, DateTime DateTime, string Notes, bool bOverdued)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM WorkDays WHERE CAST(DayStartDateTime AS DATE) = '" + DateTime.ToString("yyyy-MM-dd") + "'AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
            {
                using (new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        if (DA.Fill(DT) < 1)
                            return;

                        DT.Rows[0]["DayBreakEndDateTime"] = DateTime;
                        DT.Rows[0]["DayBreakEndFactDateTime"] = Security.GetCurrentDate();
                        DT.Rows[0]["DayContinueNotes"] = Notes;

                        DA.Update(DT);
                    }
                }
            }
        }

        public void SaveCurrentWorkDay()
        {
            using (var DA = new SqlDataAdapter("SELECT TOP 1 * FROM WorkDays WHERE WorkDayID = " + CurrentWorkDayID, ConnectionStrings.LightConnectionString))
            {
                using (new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);
                        if (DT.Rows[0]["DayEndDateTime"] != DBNull.Value)
                        {
                            DT.Rows[0]["Saved"] = true;
                            DA.Update(DT);
                            CurrentWorkDayID = 0;
                        }
                    }
                }
            }
        }

        public string SaveMyFunction(bool Saved, InfiniumFunctionsContainer InfiniumFunctionsContainer)
        {
            foreach (DataRow Row in InfiniumFunctionsContainer.FunctionsDataTable.Rows)
            {
                if (Convert.ToBoolean(Row["IsComplete"]) && Convert.ToInt32(Row["Minutes"]) == 0 && Convert.ToInt32(Row["Hours"]) == 0)
                    return "Обязанность не может быть завершенной,\nесли на неё не затрачено время.";

                if (Saved && !Convert.ToBoolean(Row["IsComplete"]) && (Convert.ToInt32(Row["Minutes"]) != 0 || Convert.ToInt32(Row["Hours"]) != 0))
                    return "После завершения рабочего дня,\nвсе обязанности, на которые было затрачено\nвремя, должны быть завершены.";

                if (Saved && !Convert.ToBoolean(Row["IsComplete"]))
                    continue;

                var Rows = newWorkDayDetailsDT.Select("WorkDayID = " + CurrentWorkDayID + " AND FunctionID = " + Row["FunctionID"]);
                if (Rows.Count() > 0)
                    continue;
                var NewRow = newWorkDayDetailsDT.NewRow();
                NewRow["WorkDayID"] = CurrentWorkDayID;
                NewRow["UserID"] = Security.CurrentUserID;
                NewRow["FunctionID"] = Row["FunctionID"];
                NewRow["FactoryID"] = Row["FactoryID"];
                NewRow["Minutes"] = Convert.ToInt32(Row["Hours"]) * 60 + Convert.ToInt32(Row["Minutes"]);
                NewRow["IsComplete"] = Row["IsComplete"];

                if (Convert.ToInt32(Row["FunctionID1"]) == 0 && Convert.ToInt32(Row["FactoryID"]) == 1)
                    NewRow["Comments"] = ProfilComments;
                if (Convert.ToInt32(Row["FunctionID1"]) == 0 && Convert.ToInt32(Row["FactoryID"]) == 2)
                    NewRow["Comments"] = TPSComments;

                newWorkDayDetailsDT.Rows.Add(NewRow);
            }

            return "";
        }

        public void SaveWorkDayDetails()
        {
            using (var DA = new SqlDataAdapter("DELETE FROM WorkDayDetails WHERE UserID = " + Security.CurrentUserID + " AND WorkDayID = " + CurrentWorkDayID, ConnectionStrings.LightConnectionString))
            {
                DA.Fill(WorkDayDetailsDT);
            }

            using (var DA = new SqlDataAdapter("SELECT * FROM WorkDayDetails", ConnectionStrings.LightConnectionString))
            {
                using (new SqlCommandBuilder(DA))
                {
                    DA.Update(newWorkDayDetailsDT);
                }
            }
            newWorkDayDetailsDT.Clear();
        }


        public DayStatus GetDayStatus(int UserID)
        {
            var DS = new DayStatus();

            using (var DA = new SqlDataAdapter("SELECT TOP 1 * FROM WorkDays WHERE CAST(DayStartDateTime AS DATE) = CAST(GETDATE() AS DATE) AND DayEndDateTime IS NOT NULL AND  UserID = " + UserID + " ORDER BY WorkDayID DESC", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    if (DA.Fill(DT) > 0)
                    {
                        DS.DayStarted = Convert.ToDateTime(DT.Rows[0]["DayStartDateTime"]);

                        if (DT.Rows[0]["DayBreakStartDateTime"] != DBNull.Value)
                        {
                            DS.BreakStarted = Convert.ToDateTime(DT.Rows[0]["DayBreakStartDateTime"]);
                            DS.BreakEnded = Convert.ToDateTime(DT.Rows[0]["DayBreakEndDateTime"]);
                        }

                        DS.DayEnded = Convert.ToDateTime(DT.Rows[0]["DayEndDateTime"]);

                        DS.iDayStatus = sDayEnded;

                        if (DT.Rows[0]["DayBreakStartDateTime"] != DBNull.Value)
                            DS.bBreak = true;
                        else
                            DS.bBreak = false;

                        if (!Convert.ToBoolean(DT.Rows[0]["Saved"]))
                            DS.iDayStatus = sDayNotSaved;

                        if (Convert.ToBoolean(DT.Rows[0]["Saved"]))
                            DS.iDayStatus = sDaySaved;

                        CurrentWorkDayID = Convert.ToInt32(DT.Rows[0]["WorkDayID"]);

                        return DS;
                    }
                }
            }

            using (var DA = new SqlDataAdapter("SELECT TOP 1 * FROM WorkDays WHERE Saved = 0 AND UserID = " + UserID + " ORDER BY WorkDayID DESC", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    if (DT.Rows.Count != 0)
                    {
                        DS.iDayStatus = sDayStarted;

                        DS.DayStarted = Convert.ToDateTime(DT.Rows[0]["DayStartDateTime"]);

                        if (DT.Rows[0]["DayBreakStartDateTime"] != DBNull.Value)
                        {
                            DS.iDayStatus = sBreakStarted;
                            DS.BreakStarted = Convert.ToDateTime(DT.Rows[0]["DayBreakStartDateTime"]);
                            DS.bBreak = true;
                        }
                        else
                            DS.bBreak = false;

                        if (DT.Rows[0]["DayBreakEndDateTime"] != DBNull.Value)
                        {
                            DS.iDayStatus = sDayContinued;
                            DS.BreakEnded = Convert.ToDateTime(DT.Rows[0]["DayBreakEndDateTime"]);
                        }

                        if (DT.Rows[0]["DayEndDateTime"] != DBNull.Value)
                        {
                            DS.iDayStatus = sDayNotSaved;
                            DS.DayEnded = Convert.ToDateTime(DT.Rows[0]["DayEndDateTime"]);
                        }

                        if (Convert.ToBoolean(DT.Rows[0]["Saved"]))
                            DS.iDayStatus = sDaySaved;
                        CurrentWorkDayID = Convert.ToInt32(DT.Rows[0]["WorkDayID"]);
                    }
                    else
                        DS.iDayStatus = sDayNotStarted;
                }
            }

            return DS;
        }

        public DayStatus GetDayStatus(int UserID, DateTime date)
        {
            var DS = new DayStatus();

            using (var DA = new SqlDataAdapter("SELECT TOP 1 * FROM WorkDays WHERE UserID = " + UserID + " AND CAST(DayStartDateTime AS DATE) = '" + date.ToString("yyyy-MM-dd") + "'", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    if (DA.Fill(DT) > 0)
                    {
                        DS.bDayStarted = false;
                        DS.bBreakStarted = false;
                        DS.bBreakEnded = false;
                        DS.bDayEnded = false;

                        if (DT.Rows[0]["DayStartDateTime"] != DBNull.Value)
                        {
                            DS.iDayStatus = sDayStarted;
                            DS.bDayStarted = true;
                            DS.DayStarted = Convert.ToDateTime(DT.Rows[0]["DayStartDateTime"]);
                        }

                        if (DT.Rows[0]["DayBreakStartDateTime"] != DBNull.Value)
                        {
                            DS.iDayStatus = sBreakStarted;
                            DS.bBreakStarted = true;
                            DS.BreakStarted = Convert.ToDateTime(DT.Rows[0]["DayBreakStartDateTime"]);
                        }

                        if (DT.Rows[0]["DayBreakEndDateTime"] != DBNull.Value)
                        {
                            DS.iDayStatus = sDayContinued;
                            DS.bBreakEnded = true;
                            DS.BreakEnded = Convert.ToDateTime(DT.Rows[0]["DayBreakEndDateTime"]);
                        }

                        if (DT.Rows[0]["DayEndDateTime"] != DBNull.Value)
                        {
                            DS.iDayStatus = sDayEnded;
                            DS.bDayEnded = true;
                            DS.DayEnded = Convert.ToDateTime(DT.Rows[0]["DayEndDateTime"]);
                        }

                        DS.dTimesheetHours = Convert.ToDecimal(DT.Rows[0]["TimesheetHours"]);
                        DS.iWorkDayID = Convert.ToInt32(DT.Rows[0]["WorkDayID"]);
                    }
                    else
                    {
                        DS.iDayStatus = sDayNotStarted;
                        DS.iWorkDayID = -1;
                    }
                }
            }
            return DS;
        }

        public void SaveDay(int UserID, DayStatus dayStatus)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM WorkDays WHERE WorkDayID = " + dayStatus.iWorkDayID, ConnectionStrings.LightConnectionString))
            {
                using (new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        if (DA.Fill(DT) > 0)
                        {
                            DT.Rows[0]["DayStartDateTime"] = dayStatus.DayStarted;
                            if (DT.Rows[0]["DayStartFactDateTime"] == DBNull.Value)
                                DT.Rows[0]["DayStartFactDateTime"] = dayStatus.DayStarted;

                            DT.Rows[0]["DayBreakStartDateTime"] = dayStatus.BreakStarted;
                            if (DT.Rows[0]["DayBreakStartFactDateTime"] == DBNull.Value)
                                DT.Rows[0]["DayBreakStartFactDateTime"] = dayStatus.BreakStarted;

                            DT.Rows[0]["DayBreakEndDateTime"] = dayStatus.BreakEnded;
                            if (DT.Rows[0]["DayBreakEndFactDateTime"] == DBNull.Value)
                                DT.Rows[0]["DayBreakEndFactDateTime"] = dayStatus.BreakEnded;

                            DT.Rows[0]["DayEndDateTime"] = dayStatus.DayEnded;
                            if (DT.Rows[0]["DayEndFactDateTime"] == DBNull.Value)
                                DT.Rows[0]["DayEndFactDateTime"] = dayStatus.DayEnded;
                            DT.Rows[0]["TimesheetHours"] = dayStatus.dTimesheetHours;
                        }
                        else
                        {
                            var NewRow = DT.NewRow();
                            NewRow["Saved"] = true;
                            NewRow["UserID"] = UserID;
                            NewRow["TimesheetHours"] = dayStatus.dTimesheetHours;
                            NewRow["DayStartDateTime"] = dayStatus.DayStarted;
                            if (NewRow["DayStartFactDateTime"] == DBNull.Value)
                                NewRow["DayStartFactDateTime"] = dayStatus.DayStarted;
                            NewRow["DayBreakStartDateTime"] = dayStatus.BreakStarted;
                            if (NewRow["DayBreakStartFactDateTime"] == DBNull.Value)
                                NewRow["DayBreakStartFactDateTime"] = dayStatus.BreakStarted;
                            NewRow["DayBreakEndDateTime"] = dayStatus.BreakEnded;
                            if (NewRow["DayBreakEndFactDateTime"] == DBNull.Value)
                                NewRow["DayBreakEndFactDateTime"] = dayStatus.BreakEnded;
                            NewRow["DayEndDateTime"] = dayStatus.DayEnded;
                            if (NewRow["DayEndFactDateTime"] == DBNull.Value)
                                NewRow["DayEndFactDateTime"] = dayStatus.DayEnded;
                            DT.Rows.Add(NewRow);
                        }

                        DA.Update(DT);
                    }
                }
            }
        }

        public void SaveTimesheetHours(int WorkDayID, decimal TimesheetHours)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM WorkDays WHERE WorkDayID = " + WorkDayID, 
                       ConnectionStrings.LightConnectionString))
            {
                using (new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        if (DA.Fill(DT) > 0)
                        {
                            DT.Rows[0]["TimesheetHours"] = TimesheetHours;
                        }

                        DA.Update(DT);
                    }
                }
            }
        }

        public void SaveLastTimesheetHours(decimal TimesheetHours)
        {
            using (var DA = new SqlDataAdapter("SELECT TOP 1 * FROM WorkDays Order by WorkDayID desc", 
                       ConnectionStrings.LightConnectionString))
            {
                using (new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        if (DA.Fill(DT) > 0)
                        {
                            DT.Rows[0]["TimesheetHours"] = TimesheetHours;
                        }

                        DA.Update(DT);
                    }
                }
            }
        }

        public DayFactStatus GetStatusFactTime(ref InfiniumTimeLabel ChangeDayStartLabel, ref InfiniumTimeLabel ChangeBreakStartLabel, ref InfiniumTimeLabel ChangeBreakEndLabel, ref InfiniumTimeLabel ChangeDayEndLabel)
        {
            ChangeDayStartLabel.ForeColor = Color.Gray;
            ChangeDayStartLabel.Text = "без изменений";

            ChangeBreakStartLabel.ForeColor = Color.Gray;
            ChangeBreakStartLabel.Text = "без изменений";

            ChangeBreakEndLabel.ForeColor = Color.Gray;
            ChangeBreakEndLabel.Text = "без изменений";

            ChangeDayEndLabel.ForeColor = Color.Gray;
            ChangeDayEndLabel.Text = "без изменений";

            var DFS = new DayFactStatus();


            using (var DA = new SqlDataAdapter("SELECT TOP 1 * FROM WorkDays WHERE (CAST(DayStartDateTime AS DATE) = CAST(GetDate() AS DATE) OR Saved = 0) AND UserID = " + Security.CurrentUserID + " ORDER BY WorkDayID DESC", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    if (DA.Fill(DT) > 0)
                    {
                        DateTime Time1, Time2;

                        if (DT.Rows[0]["DayStartDateTime"] != DBNull.Value)
                        {
                            Time1 = Convert.ToDateTime(DT.Rows[0]["DayStartDateTime"]);
                            Time2 = Convert.ToDateTime(DT.Rows[0]["DayStartFactDateTime"]);
                            DFS.DayFactStarted = Time2;

                            if (Time1 != Time2)
                            {
                                ChangeDayStartLabel.ForeColor = Color.Red;
                                ChangeDayStartLabel.Text = "с " + Time2.ToString("HH:mm") + " на " + Time1.ToString("HH:mm");
                            }
                            DFS.DayStartNotes = DT.Rows[0]["DayStartNotes"].ToString();
                        }

                        if (DT.Rows[0]["DayBreakStartDateTime"] != DBNull.Value)
                        {
                            Time1 = Convert.ToDateTime(DT.Rows[0]["DayBreakStartDateTime"]);
                            Time2 = Convert.ToDateTime(DT.Rows[0]["DayBreakStartFactDateTime"]);
                            DFS.BreakFactStarted = Time2;

                            if (Time1 != Time2)
                            {
                                ChangeBreakStartLabel.ForeColor = Color.Red;
                                ChangeBreakStartLabel.Text = "с " + Time2.ToString("HH:mm") + " на " + Time1.ToString("HH:mm");
                            }
                            DFS.DayBreakStartNotes = DT.Rows[0]["DayBreakStartNotes"].ToString();
                        }

                        if (DT.Rows[0]["DayBreakEndDateTime"] != DBNull.Value)
                        {
                            Time1 = Convert.ToDateTime(DT.Rows[0]["DayBreakEndDateTime"]);
                            Time2 = Convert.ToDateTime(DT.Rows[0]["DayBreakEndFactDateTime"]);
                            DFS.BreakFactEnded = Time2;

                            if (Time1 != Time2)
                            {
                                ChangeBreakEndLabel.ForeColor = Color.Red;
                                ChangeBreakEndLabel.Text = "с " + Time2.ToString("HH:mm") + " на " + Time1.ToString("HH:mm");
                            }
                            DFS.DayContinueNotes = DT.Rows[0]["DayContinueNotes"].ToString();
                        }

                        if (DT.Rows[0]["DayEndDateTime"] != DBNull.Value)
                        {
                            Time1 = Convert.ToDateTime(DT.Rows[0]["DayEndDateTime"]);
                            Time2 = Convert.ToDateTime(DT.Rows[0]["DayEndFactDateTime"]);
                            DFS.DayFactEnded = Time2;

                            if (Time1 != Time2)
                            {
                                ChangeDayEndLabel.ForeColor = Color.Red;
                                ChangeDayEndLabel.Text = "с " + Time2.ToString("HH:mm") + " на " + Time1.ToString("HH:mm");
                            }
                            DFS.DayEndNotes = DT.Rows[0]["DayEndNotes"].ToString();
                        }
                    }
                }
            }
            return DFS;
        }

        public static void GetWeekTimeSheet(ref InfiniumTimeLabel MonTimeSheetlabel, ref InfiniumTimeLabel TueTimeSheetlabel, ref InfiniumTimeLabel WedTimeSheetlabel, ref InfiniumTimeLabel ThuTimeSheetlabel, ref InfiniumTimeLabel FriTimeSheetlabel, ref InfiniumTimeLabel SatTimeSheetlabel, ref InfiniumTimeLabel SunTimeSheetlabel)
        {
            MonTimeSheetlabel.ForeColor = Color.Gray;
            MonTimeSheetlabel.Text = "-- : --";
            TueTimeSheetlabel.ForeColor = Color.Gray;
            TueTimeSheetlabel.Text = "-- : --";
            WedTimeSheetlabel.ForeColor = Color.Gray;
            WedTimeSheetlabel.Text = "-- : --";
            ThuTimeSheetlabel.ForeColor = Color.Gray;
            ThuTimeSheetlabel.Text = "-- : --";
            FriTimeSheetlabel.ForeColor = Color.Gray;
            FriTimeSheetlabel.Text = "-- : --";
            SatTimeSheetlabel.ForeColor = Color.Gray;
            SatTimeSheetlabel.Text = "-- : --";
            SunTimeSheetlabel.ForeColor = Color.Gray;
            SunTimeSheetlabel.Text = "-- : --";

            DateTime DayStartDateTime;
            DateTime DayEndDateTime;
            DateTime DayBreakStartDateTime;
            DateTime DayBreakEndDateTime;
            TimeSpan TimeWork;

            DayStartDateTime = DateTime.Now;
            DayStartDateTime = DayStartDateTime.Subtract(new TimeSpan(Convert.ToInt32(DateTime.Now.DayOfWeek - 1), 0, 0, 0));

            using (var DA = new SqlDataAdapter("SELECT DayStartDateTime, DayEndDateTime, DayBreakStartDateTime, DayBreakEndDateTime FROM WorkDays where DayStartDateTime > '" + DayStartDateTime.ToString("MM-dd-yyyy") + "' and UserId = " + Security.CurrentUserID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                    for (var i = 0; i < DT.Rows.Count; i++)
                    {
                        if (DT.Rows[i]["DayStartDateTime"] != DBNull.Value && DT.Rows[i]["DayEndDateTime"] != DBNull.Value)
                        {
                            DayStartDateTime = (DateTime)DT.Rows[i]["DayStartDateTime"];
                            DayEndDateTime = (DateTime)DT.Rows[i]["DayEndDateTime"];
                            TimeWork = DayEndDateTime.TimeOfDay - DayStartDateTime.TimeOfDay;
                            if (DT.Rows[i]["DayBreakStartDateTime"].ToString() != "" && DT.Rows[i]["DayBreakEndDateTime"].ToString() != "")
                            {
                                DayBreakStartDateTime = (DateTime)DT.Rows[i]["DayBreakStartDateTime"];
                                DayBreakEndDateTime = (DateTime)DT.Rows[i]["DayBreakEndDateTime"];
                                TimeWork -= DayBreakEndDateTime.TimeOfDay - DayBreakStartDateTime.TimeOfDay;
                            }

                            if (TimeWork.Seconds > 29)
                            {
                                TimeWork.Add(new TimeSpan(0, 1, 0));
                            }

                            switch ((int)DayStartDateTime.DayOfWeek)
                            {
                                case 1:
                                    MonTimeSheetlabel.ForeColor = Color.ForestGreen;
                                    MonTimeSheetlabel.Text = TimeWork.Hours + " ч : " + TimeWork.Minutes.ToString("D2") + " м";
                                    break;
                                case 2:
                                    TueTimeSheetlabel.ForeColor = Color.ForestGreen;
                                    TueTimeSheetlabel.Text = TimeWork.Hours + " ч : " + TimeWork.Minutes.ToString("D2") + " м";
                                    break;
                                case 3:
                                    WedTimeSheetlabel.ForeColor = Color.ForestGreen;
                                    WedTimeSheetlabel.Text = TimeWork.Hours + " ч : " + TimeWork.Minutes.ToString("D2") + " м";
                                    break;
                                case 4:
                                    ThuTimeSheetlabel.ForeColor = Color.ForestGreen;
                                    ThuTimeSheetlabel.Text = TimeWork.Hours + " ч : " + TimeWork.Minutes.ToString("D2") + " м";
                                    break;
                                case 5:
                                    FriTimeSheetlabel.ForeColor = Color.ForestGreen;
                                    FriTimeSheetlabel.Text = TimeWork.Hours + " ч : " + TimeWork.Minutes.ToString("D2") + " м";
                                    break;
                                case 6:
                                    SatTimeSheetlabel.ForeColor = Color.ForestGreen;
                                    SatTimeSheetlabel.Text = TimeWork.Hours + " ч : " + TimeWork.Minutes.ToString("D2") + " м";
                                    break;
                                case 7:
                                    SunTimeSheetlabel.ForeColor = Color.ForestGreen;
                                    SunTimeSheetlabel.Text = TimeWork.Hours + " ч : " + TimeWork.Minutes.ToString("D2") + " м";
                                    break;
                            }
                        }
                    }
                }
            }
        }

        public static void ChangeStartTime(int UserID, DateTime NewDateTime, string Notes)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM WorkDays WHERE CAST(DayStartDateTime AS DATE) = '" + NewDateTime.ToString("yyyy-MM-dd") + "' AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
            {
                using (new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        if (DA.Fill(DT) > 0)
                        {
                            DT.Rows[0]["DayStartDateTime"] = NewDateTime;

                            if (DT.Rows[0]["DayStartDateTime"].ToString() == DT.Rows[0]["DayStartFactDateTime"].ToString())
                            {
                                DT.Rows[0]["DayStartNotes"] = "";
                            }
                            else
                                DT.Rows[0]["DayStartNotes"] = Notes;

                            DA.Update(DT);
                        }
                    }
                }
            }
        }

        public static void ChangeEndTime(int UserID, DateTime NewDateTime, string Notes)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM WorkDays WHERE CAST(DayStartDateTime AS DATE) = '" + NewDateTime.ToString("yyyy-MM-dd") + "' AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
            {
                using (new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        if (DA.Fill(DT) > 0)
                        {
                            DT.Rows[0]["DayEndDateTime"] = NewDateTime;

                            if (DT.Rows[0]["DayEndDateTime"].ToString() == DT.Rows[0]["DayEndFactDateTime"].ToString())
                            {
                                DT.Rows[0]["DayEndNotes"] = "";
                            }
                            else
                                DT.Rows[0]["DayEndNotes"] = Notes;

                            DA.Update(DT);
                        }
                    }
                }
            }
        }

        public static void ChangeBreakStartTime(int UserID, DateTime NewDateTime, string Notes)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM WorkDays WHERE CAST(DayStartDateTime AS DATE) = '" + NewDateTime.ToString("yyyy-MM-dd") + "' AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
            {
                using (new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        if (DA.Fill(DT) > 0)
                        {
                            DT.Rows[0]["DayBreakStartDateTime"] = NewDateTime;

                            if (DT.Rows[0]["DayBreakStartDateTime"].ToString() == DT.Rows[0]["DayBreakStartFactDateTime"].ToString())
                            {
                                DT.Rows[0]["DayBreakStartNotes"] = "";
                            }
                            else
                                DT.Rows[0]["DayBreakStartNotes"] = Notes;

                            DA.Update(DT);
                        }
                    }
                }
            }
        }

        public static void ChangeBreakEndTime(int UserID, DateTime NewDateTime, string Notes)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM WorkDays WHERE CAST(DayStartDateTime AS DATE) = '" + NewDateTime.ToString("yyyy-MM-dd") + "' AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
            {
                using (new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        if (DA.Fill(DT) > 0)
                        {
                            DT.Rows[0]["DayBreakEndDateTime"] = NewDateTime;

                            if (DT.Rows[0]["DayBreakEndDateTime"].ToString() == DT.Rows[0]["DayBreakEndFactDateTime"].ToString())
                            {
                                DT.Rows[0]["DayContinueNotes"] = "";
                            }
                            else
                                DT.Rows[0]["DayContinueNotes"] = Notes;

                            DA.Update(DT);
                        }
                    }
                }
            }
        }
    }





    public class LightFunctions
    {
        public DataTable DepartmentFunctionsDataTable;
        public BindingSource DepartmentFunctionsBindingSource;

        public DataTable UserFunctionsDataTable;
        public BindingSource UserFunctionsBindingSource;
        public BindingSource SelectedFunctionsBindingSource;

        public SqlDataAdapter SelectedFunctionsDA;
        public SqlCommandBuilder SelectedFunctionsCB;

        public int CurrentPercents;
        private DataTable SelectedFunctionsDataTable;
        private readonly int UserID;
        private int DepartmentID = -1;

        public LightFunctions(int iUserID)
        {
            UserID = iUserID;
            GetDepartmentID();
            Create();
            Fill();
            Binding();
        }

        private void GetDepartmentID()
        {
            using (var DA = new SqlDataAdapter("SELECT DepartmentID FROM Users WHERE UserID = " + UserID,
                ConnectionStrings.UsersConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    if (DT.Rows.Count < 1)
                        return;

                    DepartmentID = Convert.ToInt32(DT.Rows[0]["DepartmentID"]);
                }
            }
        }

        private void Create()
        {
            DepartmentFunctionsDataTable = new DataTable();
            UserFunctionsDataTable = new DataTable();
            SelectedFunctionsDataTable = new DataTable();

            DepartmentFunctionsBindingSource = new BindingSource();
        }

        private void Fill()
        {
            using (var DA = new SqlDataAdapter("SELECT UserFunctions.*, Functions.FunctionName FROM UserFunctions" +
                                               " INNER JOIN Functions ON UserFunctions.FunctionID = Functions.FunctionID WHERE UserFunctions.UserID = " + UserID +
                                               " ORDER BY FunctionName", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(UserFunctionsDataTable);
            }

            using (var DA = new SqlDataAdapter("SELECT * FROM Functions WHERE DepartmentID = " + DepartmentID +
                                               " ORDER BY FunctionName", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(DepartmentFunctionsDataTable);
            }

            FillCurrentFunctions();
        }

        private void Binding()
        {
            DepartmentFunctionsBindingSource.DataSource = DepartmentFunctionsDataTable;
            UserFunctionsBindingSource = new BindingSource
            {
                DataSource = UserFunctionsDataTable
            };
            SelectedFunctionsBindingSource = new BindingSource
            {
                DataSource = SelectedFunctionsDataTable
            };
        }

        public void GetCurrentPercents()
        {
            CurrentPercents = 0;

            foreach (DataRow Row in SelectedFunctionsDataTable.Rows)
            {
                if (Row.RowState == DataRowState.Deleted)
                    continue;

                CurrentPercents += Convert.ToInt32(Row["DayPercent"]);
            }
        }

        public void SetGrid(ref PercentageDataGrid Grid, bool Image)
        {
            var FunctionsColumn = new DataGridViewComboBoxColumn
            {
                Name = "FunctionsColumn",
                HeaderText = "Обязанность",
                DataPropertyName = "FunctionID",
                DataSource = new DataView(DepartmentFunctionsDataTable),
                ValueMember = "FunctionID",
                DisplayMember = "FunctionName",
                DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing,
                DisplayIndex = 0
            };
            Grid.Columns.Add(FunctionsColumn);

            Grid.Columns["WorkDayDetailID"].Visible = false;
            Grid.Columns["FunctionID"].Visible = false;
            Grid.Columns["WorkDayID"].Visible = false;
            Grid.Columns["UserID"].Visible = false;
            Grid.Columns["DayPercent"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            Grid.Columns["DayPercent"].Width = 60;
            Grid.Columns["DayPercent"].ReadOnly = true;

            if (Image)
            {
                var ImageColumn = new DataGridViewImageColumn
                {
                    AutoSizeMode = DataGridViewAutoSizeColumnMode.None,
                    Width = 40,
                    Name = "ImageColumn",
                    Image = Resources.CancelGrid
                };
                ImageColumn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Grid.Columns.Add(ImageColumn);
                Grid.Columns["ImageColumn"].DisplayIndex = 5;
            }

        }

        public void FillCurrentFunctions()
        {
            var WorkDayID = GetWorkDayID();

            SelectedFunctionsDataTable.Clear();

            if (WorkDayID == -1)
            {
                using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM WorkDayDetails WHERE UserID = " + UserID + " AND WorkDayDetails.WorkDayID = " + WorkDayID, ConnectionStrings.LightConnectionString))
                {
                    DA.Fill(SelectedFunctionsDataTable);
                }

                return;
            }

            if (SelectedFunctionsDataTable.Rows.Count > 0)
            {
                SelectedFunctionsDataTable.Clear();
                SelectedFunctionsDA.Fill(SelectedFunctionsDataTable);
            }
            else
            {
                SelectedFunctionsDA = new SqlDataAdapter("SELECT * FROM WorkDayDetails WHERE UserID = " + UserID + " AND WorkDayDetails.WorkDayID = " + WorkDayID, ConnectionStrings.LightConnectionString);
                SelectedFunctionsCB = new SqlCommandBuilder(SelectedFunctionsDA);
                SelectedFunctionsDA.Fill(SelectedFunctionsDataTable);
            }


            GetCurrentPercents();

        }

        public void ClearSelectedFunctions()
        {
            SelectedFunctionsDataTable.Clear();
        }

        public int GetWorkDayID()
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM WorkDays WHERE UserID = " + UserID + " ORDER BY WorkdayID DESC", ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        if (DA.Fill(DT) == 0)
                        {
                            return -1;
                        }

                        return Convert.ToInt32(DT.Rows[0]["WorkDayID"]);
                    }
                }
            }
        }

        public void SaveFunctions()
        {
            var WorkDayID = GetWorkDayID();

            SelectedFunctionsDA.Update(SelectedFunctionsDataTable);
            SelectedFunctionsDataTable.Clear();
            SelectedFunctionsDA.Fill(SelectedFunctionsDataTable);

        }

        public void AddFunction(int FunctionID, int Percent)
        {
            if (SelectedFunctionsDataTable.Select("FunctionID = " + FunctionID).Count() > 0)
                return;

            var WorkDayID = GetWorkDayID();

            var NewRow = SelectedFunctionsDataTable.NewRow();
            NewRow["FunctionID"] = FunctionID;
            NewRow["WorkDayID"] = WorkDayID;
            NewRow["UserID"] = UserID;
            NewRow["DayPercent"] = Percent;
            SelectedFunctionsDataTable.Rows.Add(NewRow);

            CurrentPercents = 0;

            GetCurrentPercents();
        }

    }





    public class LightAddFunctions
    {
        private DataTable FunctionsDataTable;
        public BindingSource FunctionsBindingSource;
        private readonly PercentageDataGrid FunctionsGrid;

        public LightAddFunctions(ref PercentageDataGrid tFunctionsGrid)
        {
            FunctionsGrid = tFunctionsGrid;
            Create();
            Fill();
            Binding();
            GridSettings();
        }

        private void Create()
        {
            FunctionsDataTable = new DataTable();

            FunctionsBindingSource = new BindingSource();
        }

        private void Fill()
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM Functions" +
                                               " WHERE Functions.UserID = " + Security.CurrentUserID, ConnectionStrings.LightConnectionString))
            {
                DA.Fill(FunctionsDataTable);
            }
        }

        private void Binding()
        {
            FunctionsBindingSource.DataSource = FunctionsDataTable;
            FunctionsGrid.DataSource = FunctionsBindingSource;
        }

        private void GridSettings()
        {
            if (FunctionsGrid != null)
            {
                FunctionsGrid.AutoGenerateColumns = false;

                var ImageColumn = new DataGridViewImageColumn
                {
                    AutoSizeMode = DataGridViewAutoSizeColumnMode.None,
                    Width = 40,
                    Name = "ImageColumn",
                    Image = Resources.CancelGrid
                };
                ImageColumn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                FunctionsGrid.Columns.Add(ImageColumn);

                FunctionsGrid.Columns["FunctionID"].Visible = false;
                FunctionsGrid.Columns["UserID"].Visible = false;
                FunctionsGrid.Columns["FunctionName"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                FunctionsGrid.Columns["FunctionName"].DisplayIndex = 0;
                FunctionsGrid.Columns["ImageColumn"].DisplayIndex = 1;
            }
        }

        public void SaveFunctions()
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM Functions" +
                                               " WHERE Functions.UserID = " + Security.CurrentUserID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    DA.Update(FunctionsDataTable);
                    FunctionsDataTable.Clear();
                    DA.Fill(FunctionsDataTable);
                }
            }
        }

        public void AddFunction(string Name)
        {
            var NewRow = FunctionsDataTable.NewRow();
            NewRow["UserID"] = Security.CurrentUserID;
            NewRow["FunctionName"] = Name;
            FunctionsDataTable.Rows.Add(NewRow);
        }
    }





    public class LightMessage
    {
        public DataTable UsersDataTable;
        public DataTable MessagesDataTable;
        public DataTable SelectedUsersDataTable;
        public BindingSource UsersBindingSource;
        public BindingSource SelectedUsersBindingSource;

        public int CurrentUserID = Security.CurrentUserID;
        public string CurrentUserName;
        private readonly UsersMessagesDataGrid SelectedUsersGrid;
        private readonly UsersDataGrid UsersListDataGrid;

        public LightMessage(ref UsersMessagesDataGrid tSelectedUsersGrid, ref UsersDataGrid tUsersListDataGrid)
        {
            UsersDataTable = new DataTable();

            using (var DA = new SqlDataAdapter("SELECT UserID, Name, Online FROM USERS WHERE Fired <> 1 ORDER BY Name ASC", ConnectionStrings.UsersConnectionString))
            {
                DA.Fill(UsersDataTable);
            }

            //UsersDataTable.Columns.Add(new DataColumn("OnlineStatus", Type.GetType("System.Boolean")));

            /////////
            //UsersDataTable.Rows[3]["OnlineStatus"] = true;

            CurrentUserName = Security.GetUserNameByID(CurrentUserID);

            MessagesDataTable = new DataTable();

            UsersBindingSource = new BindingSource
            {
                DataSource = UsersDataTable,
                Filter = "UserID <> " + CurrentUserID
            };
            SelectedUsersDataTable = new DataTable();
            SelectedUsersDataTable = UsersDataTable.Clone();
            SelectedUsersDataTable.Columns.Add(new DataColumn("UpdatesCount", Type.GetType("System.Int32")));

            SelectedUsersGrid = tSelectedUsersGrid;
            UsersListDataGrid = tUsersListDataGrid;

            SelectedUsersBindingSource = new BindingSource
            {
                DataSource = SelectedUsersDataTable
            };
            SelectedUsersGrid.DataSource = SelectedUsersBindingSource;

            UsersListDataGrid.DataSource = UsersBindingSource;

            SetSelectedUsersGrid();
            SetUsersGrid();
        }

        public void FillMessages(int SenderID)
        {
            MessagesDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT TOP 100 infiniu2_light.dbo.Messages.MessageID, infiniu2_light.dbo.Messages.Text, infiniu2_light.dbo.Messages.SenderUserID, infiniu2_light.dbo.Messages.RecipientUserID, infiniu2_light.dbo.Messages.SendDateTime, infiniu2_users.dbo.Users.Name AS SenderName FROM infiniu2_light.dbo.Messages INNER JOIN infiniu2_users.dbo.Users ON infiniu2_users.dbo.Users.UserID = infiniu2_light.dbo.Messages.SenderUserID WHERE (RecipientUserID = " + CurrentUserID + " AND SenderUserID = " + SenderID + ") OR (RecipientUserID = " + SenderID + " AND SenderUserID = " + CurrentUserID + ") ORDER BY SendDateTime DESC", ConnectionStrings.UsersConnectionString))
            {
                DA.Fill(MessagesDataTable);
            }

            foreach (DataRow Row in MessagesDataTable.Rows)
            {
                Row["Text"] = LightCrypto.Decrypt(Row["Text"].ToString(), true, GetUserNameByID(Convert.ToInt32(Row["SenderUserID"])));
            }
        }

        private string GetUserNameByID(int UserID)
        {
            return UsersDataTable.Select("UserID = " + UserID)[0]["Name"].ToString();
        }

        public void SetSelectedUsersGrid()
        {
            SelectedUsersGrid.AddColumns();

            SelectedUsersGrid.Columns["UserID"].Visible = false;
            SelectedUsersGrid.Columns["UpdatesCount"].Visible = false;
            SelectedUsersGrid.Columns["Online"].Visible = false;
            SelectedUsersGrid.Columns["CloseColumn"].DisplayIndex = 3;
            SelectedUsersGrid.Columns["OnlineColumn"].DisplayIndex = 0;
            SelectedUsersGrid.sNewMessagesColumnName = "UpdatesCount";
            SelectedUsersGrid.sOnlineStatusColumnName = "Online";
        }

        public void SetUsersGrid()
        {
            UsersListDataGrid.AddColumns();
            UsersListDataGrid.sOnlineStatusColumnName = "Online";
            UsersListDataGrid.Columns["OnlineColumn"].DisplayIndex = 0;
            UsersListDataGrid.Columns["UserID"].Visible = false;
            UsersListDataGrid.Columns["Online"].Visible = false;
        }

        public DataTable GetOnlineStatus()
        {
            using (var DA = new SqlDataAdapter("SELECT UserID, Online FROM Users WHERE Fired <> 1", ConnectionStrings.UsersConnectionString))
            {
                using (var DT = new DataTable())
                {
                    try
                    {
                        DA.Fill(DT);
                    }
                    catch
                    {

                    }

                    return DT;
                }
            }
        }

        public void AddUserToSelected(int UserID)
        {
            if (SelectedUsersDataTable.Select("UserID = " + UserID).Count() > 0)
            {
                SelectedUsersBindingSource.Position = SelectedUsersBindingSource.Find("UserID", UserID);

                return;
            }

            var Row = UsersDataTable.Select("UserID = " + UserID);

            SelectedUsersDataTable.ImportRow(Row[0]);

            SelectedUsersBindingSource.Position = SelectedUsersBindingSource.Find("UserID", UserID);
        }

        public void AddSenderToSelected(int UserID)
        {
            var sRow = SelectedUsersDataTable.Select("UserID = " + UserID);

            if (sRow.Count() > 0)
            {
                //SelectedUsersBindingSource.Position = SelectedUsersBindingSource.Find("UserID", UserID);
                sRow[0]["UpdatesCount"] = 1;
                return;
            }

            var Row = UsersDataTable.Select("UserID = " + UserID);

            var NewRow = SelectedUsersDataTable.NewRow();
            NewRow["UserID"] = Row[0]["UserID"];
            NewRow["Name"] = Row[0]["Name"];
            NewRow["UpdatesCount"] = 1;
            NewRow["Online"] = Row[0]["Online"];
            SelectedUsersDataTable.Rows.Add(NewRow);

            //SelectedUsersBindingSource.Position = SelectedUsersBindingSource.Find("UserID", UserID);
        }

        public void AddMessage(int RecipientUserID, string sText)
        {
            using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM Messages", ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        var NewRow = DT.NewRow();
                        NewRow["SenderUserID"] = CurrentUserID;
                        NewRow["RecipientUserID"] = RecipientUserID;
                        NewRow["Text"] = LightCrypto.Encrypt(sText, true, CurrentUserName);

                        var DateTime = Security.GetCurrentDate();

                        NewRow["SendDateTime"] = DateTime;
                        DT.Rows.Add(NewRow);

                        DA.Update(DT);

                        FillMessages(RecipientUserID);

                        var Row = MessagesDataTable.Select("SendDateTime = '" + DateTime.ToString("yyyy-MM-dd HH:mm:ss.fff") + "'");

                        using (var sDA = new SqlDataAdapter("SELECT TOP 0 * FROM SubscribesRecords", ConnectionStrings.LightConnectionString))
                        {
                            using (var sCB = new SqlCommandBuilder(sDA))
                            {
                                using (var sDT = new DataTable())
                                {
                                    sDA.Fill(sDT);

                                    var sNewRow = sDT.NewRow();
                                    sNewRow["SubscribesItemID"] = 5;
                                    sNewRow["TableItemID"] = Convert.ToInt32(Row[0][0]);
                                    sNewRow["UserID"] = RecipientUserID;
                                    sNewRow["UserTypeID"] = 0;
                                    sDT.Rows.Add(sNewRow);

                                    sDA.Update(sDT);
                                }
                            }
                        }
                    }
                }
            }
        }

        public static void SendMessage(int RecipientUserID, string sText)
        {
            using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM Messages", ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        var NewRow = DT.NewRow();
                        NewRow["SenderUserID"] = Security.CurrentUserID;
                        NewRow["RecipientUserID"] = RecipientUserID;
                        NewRow["Text"] = LightCrypto.Encrypt(sText, true, Security.CurrentUserName);

                        var DateTime = Security.GetCurrentDate();

                        NewRow["SendDateTime"] = DateTime;
                        DT.Rows.Add(NewRow);

                        DA.Update(DT);

                        using (var sDA = new SqlDataAdapter("SELECT MessageID FROM Messages WHERE SendDateTime = '" + DateTime.ToString("yyyy-MM-dd HH:mm:ss.fff") + "'", ConnectionStrings.LightConnectionString))
                        {
                            using (var sDT = new DataTable())
                            {
                                sDA.Fill(sDT);

                                ActiveNotifySystem.CreateSubscribeRecord(5, Convert.ToInt32(sDT.Rows[0][0]), RecipientUserID);
                            }
                        }
                    }
                }
            }
        }

        public int GetNewMessages()
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM Messages WHERE MessageID IN (SELECT TableItemID FROM SubscribesRecords WHERE SubscribesItemID = 5 AND UserTypeID = 0 AND UserID = " + CurrentUserID + ") ORDER BY SendDateTime DESC", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    if (DA.Fill(DT) == 0)
                        return 0;

                    foreach (DataRow Row in DT.Rows)
                    {
                        AddSenderToSelected(Convert.ToInt32(Row["SenderUserID"]));
                    }

                    return DT.Rows.Count;
                }
            }
        }

        public void ClearCurrentUpdates()
        {
            ((DataRowView)SelectedUsersBindingSource.Current)["UpdatesCount"] = 0;
        }

        public int GetCurrentUserID()
        {
            return Convert.ToInt32(((DataRowView)SelectedUsersBindingSource.Current).Row["UserID"]);
        }

        public void RemoveCurrent()
        {
            var Pos = SelectedUsersBindingSource.Position;
            SelectedUsersBindingSource.RemoveCurrent();

            //остается на позиции удаленного
            if (SelectedUsersBindingSource.Count > 0)
                if (Pos >= SelectedUsersBindingSource.Count)
                {
                    SelectedUsersBindingSource.MoveLast();
                    SelectedUsersGrid.Rows[SelectedUsersGrid.Rows.Count - 1].Selected = true;
                }
                else
                    SelectedUsersBindingSource.Position = Pos;

            ((DataTable)SelectedUsersBindingSource.DataSource).AcceptChanges();
        }


        public bool CheckOnline(DataTable DT)
        {
            var bNeedRefresh = false;

            if (DT == null)
                return false;

            if (DT.Rows.Count == 0)
                return false;

            foreach (DataRow Row in UsersDataTable.Rows)
            {
                if (Row["Online"] == DBNull.Value)
                {
                    Row["Online"] = DT.Select("UserID = " + Row["UserID"])[0]["Online"];
                    bNeedRefresh = true;
                    continue;
                }

                if (Convert.ToBoolean(DT.Select("UserID = " + Row["UserID"])[0]["Online"]) != Convert.ToBoolean(Row["Online"]))
                {
                    Row["Online"] = DT.Select("UserID = " + Row["UserID"])[0]["Online"];

                    bNeedRefresh = true;
                }
            }

            if (bNeedRefresh == false)
                return false;

            foreach (DataRow Row in SelectedUsersDataTable.Rows)
            {
                if (Row["Online"] == DBNull.Value)
                {
                    Row["Online"] = DT.Select("UserID = " + Row["UserID"])[0]["Online"];
                    continue;
                }

                if (Convert.ToBoolean(DT.Select("UserID = " + Row["UserID"])[0]["Online"]) != Convert.ToBoolean(Row["Online"]))
                {
                    Row["Online"] = DT.Select("UserID = " + Row["UserID"])[0]["Online"];
                }
            }

            return true;
        }


        public bool IsEmptyMessage(string sText)
        {
            if (sText.Length == 0)
                return true;

            var n = 0;

            foreach (var c in sText)
            {
                if (c == '\n' || c == '\r' || c == ' ')
                    n++;
            }

            if (n == sText.Length)
                return true;

            return false;
        }
    }




    public class InfiniumMessages
    {
        public DataTable UsersDataTable;
        public DataTable SelectedUsersDataTable;
        public DataTable MessagesDataTable;
        public DataTable FullUsersDataTable;

        public InfiniumMessages()
        {
            UsersDataTable = TablesManager.UsersDataTable.Copy();

            FullUsersDataTable = TablesManager.UsersDataTable.Copy();

            UsersDataTable.Rows.Remove(UsersDataTable.Select("UserID = " + Security.CurrentUserID)[0]);

            foreach (DataRow Row in UsersDataTable.Rows)
            {
                var Name = Row["Name"].ToString();
                var NewName = "";

                var g = 0;

                for (var i = 0; i < Name.Length; i++)
                {
                    if (Name[i] != ' ')
                        NewName += Name[i];
                    else
                    {
                        if (g > 0)
                            break;

                        NewName += Name[i];
                        g++;
                    }
                }

                Row["Name"] = NewName;
            }

            SelectedUsersDataTable = new DataTable();
            SelectedUsersDataTable.Columns.Add(new DataColumn("UserID", Type.GetType("System.Int32")));
            SelectedUsersDataTable.Columns.Add(new DataColumn("Name", Type.GetType("System.String")));
            SelectedUsersDataTable.Columns.Add(new DataColumn("Count", Type.GetType("System.Int32")));
            SelectedUsersDataTable.Columns.Add(new DataColumn("Online", Type.GetType("System.Boolean")));
            SelectedUsersDataTable.Columns.Add(new DataColumn("Photo", Type.GetType("System.Byte[]")));


            MessagesDataTable = new DataTable();
        }

        public bool IsNewMessages()
        {
            using (var DA = new SqlDataAdapter("SELECT TOP 1 SubscribesRecordID FROM SubscribesRecords WHERE SubscribesItemID = 5 AND UserID = " + Security.CurrentUserID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    return DA.Fill(DT) > 0;
                }
            }
        }

        public static void SendMessage(string Text, int UserID)
        {
            using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM Messages", ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        var NewRow = DT.NewRow();
                        NewRow["Text"] = Text;
                        NewRow["SenderUserID"] = Security.CurrentUserID;
                        NewRow["RecipientUserID"] = UserID;

                        var Date = Security.GetCurrentDate();

                        NewRow["SendDateTime"] = Date;
                        DT.Rows.Add(NewRow);

                        DA.Update(DT);

                        using (var mDA = new SqlDataAdapter("SELECT MessageID FROM Messages WHERE SendDateTime = '" + Date.ToString("yyyy-MM-dd HH:mm:ss.fff") + "'", ConnectionStrings.LightConnectionString))
                        {
                            using (var mDT = new DataTable())
                            {
                                mDA.Fill(mDT);

                                using (var uDA = new SqlDataAdapter("SELECT TOP 0 * FROM SubscribesRecords", ConnectionStrings.LightConnectionString))
                                {
                                    using (var uCB = new SqlCommandBuilder(uDA))
                                    {
                                        using (var uDT = new DataTable())
                                        {
                                            uDA.Fill(uDT);

                                            var cNewRow = uDT.NewRow();
                                            cNewRow["TableItemID"] = mDT.Rows[0]["MessageID"];
                                            cNewRow["UserID"] = UserID;
                                            cNewRow["SubscribesItemID"] = 5;
                                            uDT.Rows.Add(cNewRow);

                                            uDA.Update(uDT);
                                        }
                                    }
                                }
                            }
                        }

                    }
                }
            }
        }

        public int AddUserToSelected(int UserID)
        {
            var Row = SelectedUsersDataTable.Select("UserID = " + UserID);

            if (Row.Count() > 0)
                return SelectedUsersDataTable.Rows.IndexOf(Row[0]);

            var uRow = UsersDataTable.Select("UserID = " + UserID);

            var NewRow = SelectedUsersDataTable.NewRow();
            NewRow["UserID"] = uRow[0]["UserID"];
            NewRow["Name"] = uRow[0]["Name"];
            NewRow["Count"] = 0;
            NewRow["Online"] = uRow[0]["Online"];
            NewRow["Photo"] = uRow[0]["Photo"];
            SelectedUsersDataTable.Rows.Add(NewRow);

            return -1;
        }

        public void RemoveUserFromSelected(int UserID)
        {
            SelectedUsersDataTable.Rows.Remove(SelectedUsersDataTable.Select("UserID = " + UserID)[0]);
        }

        public void FillMessages(int UserID)
        {
            MessagesDataTable.Clear();

            using (var DA = new SqlDataAdapter(@"SELECT * FROM Messages WHERE MessageID IN (SELECT TOP 100 MessageID FROM Messages 
                WHERE (SenderUserID = " + UserID + "AND RecipientUserID = " + Security.CurrentUserID + ")" + " OR (SenderUserID = " + Security.CurrentUserID + " AND RecipientUserID = " + UserID + ") ORDER BY MessageID DESC) ORDER BY MessageID", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(MessagesDataTable);
            }
        }

        public bool FillSelectedUsers(int iUserID)
        {
            SelectedUsersDataTable.Clear();

            var bRes = false;

            using (var DA = new SqlDataAdapter("SELECT SenderUserID FROM Messages WHERE MessageID IN (SELECT TableItemID FROM SubscribesRecords WHERE UserTypeID = 0 AND SubscribesItemID = 5 AND UserID = " + Security.CurrentUserID + ")", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    foreach (DataRow Row in DT.Rows)
                    {
                        var sRow = SelectedUsersDataTable.Select("UserID = " + Row["SenderUserID"]);

                        if (Convert.ToInt32(Row["SenderUserID"]) == iUserID)
                            bRes = true;

                        if (sRow.Count() == 0)
                        {
                            var uRow = TablesManager.UsersDataTable.Select("UserID = " + Row["SenderUserID"])[0];

                            var NewRow = SelectedUsersDataTable.NewRow();
                            NewRow["UserID"] = uRow["UserID"];
                            NewRow["Name"] = uRow["Name"];
                            NewRow["Photo"] = uRow["Photo"];
                            NewRow["Online"] = uRow["Online"];
                            NewRow["Count"] = 1;
                            SelectedUsersDataTable.Rows.Add(NewRow);
                        }
                        else
                        {
                            sRow[0]["Count"] = Convert.ToInt32(sRow[0]["Count"]) + 1;
                        }
                    }
                }
            }

            return bRes;
        }

        public void ClearCount(int UserID)
        {
            SelectedUsersDataTable.Select("UserID = " + UserID)[0]["Count"] = 0;
        }

        public void ClearSubscribes(int UserID)
        {
            using (var DA = new SqlDataAdapter("DELETE FROM SubscribesRecords WHERE SubscribesItemID = 5 AND UserTypeID = 0 AND UserID = " + Security.CurrentUserID + " AND TableItemID IN (SELECT MessageID FROM Messages WHERE RecipientUserID = " + Security.CurrentUserID + " AND SenderUserID = " + UserID + "AND IsRead = 0)", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }


            using (var mDA = new SqlDataAdapter("SELECT * FROM Messages WHERE IsRead = 0 AND RecipientUserID = " + Security.CurrentUserID + " AND SenderUserID = " + UserID, ConnectionStrings.LightConnectionString))
            {
                using (var mCB = new SqlCommandBuilder(mDA))
                {
                    using (var mDT = new DataTable())
                    {
                        mDA.Fill(mDT);

                        foreach (DataRow Row in mDT.Rows)
                        {
                            Row["IsRead"] = true;
                        }

                        mDA.Update(mDT);
                    }
                }
            }


        }
    }




    public class LightNotes
    {
        public DataTable NotesDataTable;
        private readonly int CurrentUserID;

        public LightNotes(int UserID)
        {
            NotesDataTable = new DataTable();

            CurrentUserID = UserID;
            Fill();
        }

        public void Fill()
        {
            NotesDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT * FROM Notes WHERE UserID = " + CurrentUserID, ConnectionStrings.LightConnectionString))
            {
                DA.Fill(NotesDataTable);
            }
        }

        public string GetNotes(int NotesID)
        {
            return NotesDataTable.Select("NotesID = " + NotesID)[0]["NotesText"].ToString();
        }

        public void AddNotes(string Name, int Priority)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM Notes WHERE UserID = " + CurrentUserID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        var NewRow = DT.NewRow();
                        NewRow["UserID"] = CurrentUserID;
                        NewRow["CreationDateTime"] = Security.GetCurrentDate();
                        NewRow["NotesName"] = Name;
                        NewRow["Priority"] = Priority;

                        DT.Rows.Add(NewRow);

                        DA.Update(DT);
                    }
                }
            }

            Fill();
        }

        public void SaveNote(int NotesID, string Name, string NotesText)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM Notes WHERE NotesID = " + NotesID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        DT.Rows[0]["NotesName"] = Name;
                        DT.Rows[0]["NotesText"] = NotesText;

                        DA.Update(DT);
                    }
                }
            }

            Fill();
        }

        public int DeleteNotes(int NotesID)
        {
            var i = NotesDataTable.Rows.IndexOf(NotesDataTable.Select("NotesID = " + NotesID)[0]);

            using (var DA = new SqlDataAdapter("SELECT * FROM Notes WHERE NotesID =" + NotesID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);
                        DT.Rows[0].Delete();
                        DA.Update(DT);
                    }
                }
            }

            Fill();

            return i;
        }
    }





    public class ToolsSellersManager
    {
        private int CurrentToolsSellerID = -1;
        private int CurrentToolsSellerGroupID = -1;
        private int CurrentToolsSellerSubGroupID = -1;
        private DataTable ToolsSellersDataTable;
        private DataTable ToolsSellersGroupsDataTable;
        private DataTable ToolsSellersSubGroupsDataTable;
        private DataTable ToolsSellerInfoDataTable;
        private BindingSource ToolsSellersBindingSource;
        private BindingSource ToolsSellersGroupsBindingSource;
        private BindingSource ToolsSellersSubGroupsBindingSource;
        private readonly PercentageDataGrid ToolsSellersDataGrid;
        private readonly PercentageDataGrid ToolsSellerInfoDataGrid;
        private readonly PercentageDataGrid ToolsSellersGroupsDataGrid;
        private readonly PercentageDataGrid ToolsSellersSubGroupsDataGrid;
        private SqlDataAdapter ToolsSellersDA;
        private SqlDataAdapter ToolsSellersGroupsDA;
        private SqlDataAdapter ToolsSellersSubGroupsDA;
        private SqlCommandBuilder ToolsSellersCB;
        private SqlCommandBuilder ToolsSellersGroupsCB;
        private SqlCommandBuilder ToolsSellersSubGroupsCB;

        public ToolsSellersManager(
            ref PercentageDataGrid tToolsSellersDataGrid,
            ref PercentageDataGrid tToolsSellerInfoDataGrid,
            ref PercentageDataGrid tToolsSellersGroupsDataGrid,
            ref PercentageDataGrid tToolsSellersSubGroupsDataGrid)
        {
            ToolsSellersDataGrid = tToolsSellersDataGrid;
            ToolsSellerInfoDataGrid = tToolsSellerInfoDataGrid;
            ToolsSellersGroupsDataGrid = tToolsSellersGroupsDataGrid;
            ToolsSellersSubGroupsDataGrid = tToolsSellersSubGroupsDataGrid;

            Create();
            Fill();
            Binding();
            GridSettings();
        }

        private void Create()
        {
            ToolsSellersDataTable = new DataTable();
            ToolsSellersGroupsDataTable = new DataTable();
            ToolsSellersSubGroupsDataTable = new DataTable();
            ToolsSellerInfoDataTable = new DataTable("SellerInfoDataTable");

            ToolsSellerInfoDataTable.Columns.Add(new DataColumn("Name", Type.GetType("System.String")));
            ToolsSellerInfoDataTable.Columns.Add(new DataColumn("Position", Type.GetType("System.String")));
            ToolsSellerInfoDataTable.Columns.Add(new DataColumn("Phone", Type.GetType("System.String")));
            ToolsSellerInfoDataTable.Columns.Add(new DataColumn("Email", Type.GetType("System.String")));
            ToolsSellerInfoDataTable.Columns.Add(new DataColumn("ICQ", Type.GetType("System.String")));
            ToolsSellerInfoDataTable.Columns.Add(new DataColumn("Skype", Type.GetType("System.String")));

            ToolsSellersBindingSource = new BindingSource();
            ToolsSellersGroupsBindingSource = new BindingSource();
            ToolsSellersSubGroupsBindingSource = new BindingSource();
        }

        private void Fill()
        {
            ToolsSellersDA = new SqlDataAdapter("SELECT * FROM ToolsSellers ORDER BY ToolsSellerName",
                ConnectionStrings.LightConnectionString);
            ToolsSellersCB = new SqlCommandBuilder(ToolsSellersDA);
            ToolsSellersDA.Fill(ToolsSellersDataTable);

            ToolsSellersGroupsDA = new SqlDataAdapter("SELECT * FROM ToolsSellersGroups ORDER BY ToolsSellerGroupName",
                ConnectionStrings.LightConnectionString);
            ToolsSellersGroupsCB = new SqlCommandBuilder(ToolsSellersGroupsDA);
            ToolsSellersGroupsDA.Fill(ToolsSellersGroupsDataTable);

            ToolsSellersSubGroupsDA = new SqlDataAdapter("SELECT * FROM ToolsSellersSubGroups ORDER BY ToolsSellerSubGroupName",
                ConnectionStrings.LightConnectionString);
            ToolsSellersSubGroupsCB = new SqlCommandBuilder(ToolsSellersSubGroupsDA);
            ToolsSellersSubGroupsDA.Fill(ToolsSellersSubGroupsDataTable);
        }

        private void Binding()
        {
            ToolsSellersBindingSource.DataSource = ToolsSellersDataTable;
            ToolsSellersDataGrid.DataSource = ToolsSellersBindingSource;

            ToolsSellersGroupsBindingSource.DataSource = ToolsSellersGroupsDataTable;
            ToolsSellersGroupsDataGrid.DataSource = ToolsSellersGroupsBindingSource;

            ToolsSellersSubGroupsBindingSource.DataSource = ToolsSellersSubGroupsDataTable;
            ToolsSellersSubGroupsDataGrid.DataSource = ToolsSellersSubGroupsBindingSource;

            ToolsSellerInfoDataGrid.DataSource = new DataView(ToolsSellerInfoDataTable);
        }

        private void GridSettings()
        {
            ToolsSellersDataGrid.Columns["ToolsSellerID"].Visible = false;
            ToolsSellersDataGrid.Columns["ToolsSellerSubGroupID"].Visible = false;
            ToolsSellersDataGrid.Columns["Address"].Visible = false;
            ToolsSellersDataGrid.Columns["ContractDocNumber"].Visible = false;
            ToolsSellersDataGrid.Columns["Email"].Visible = false;
            ToolsSellersDataGrid.Columns["Site"].Visible = false;
            ToolsSellersDataGrid.Columns["ContactsXML"].Visible = false;
            ToolsSellersDataGrid.Columns["Notes"].Visible = false;
            ToolsSellersDataGrid.Columns["Country"].Visible = false;

            ToolsSellersDataGrid.ColumnHeadersVisible = false;

            ToolsSellerInfoDataGrid.Columns["Name"].MinimumWidth = 150;
            ToolsSellerInfoDataGrid.Columns["Name"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            ToolsSellerInfoDataGrid.Columns["Position"].MinimumWidth = 150;
            ToolsSellerInfoDataGrid.Columns["Position"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            ToolsSellerInfoDataGrid.Columns["Phone"].MinimumWidth = 150;
            ToolsSellerInfoDataGrid.Columns["Phone"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            ToolsSellerInfoDataGrid.Columns["Email"].MinimumWidth = 150;
            ToolsSellerInfoDataGrid.Columns["Email"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            ToolsSellerInfoDataGrid.Columns["ICQ"].MinimumWidth = 150;
            ToolsSellerInfoDataGrid.Columns["ICQ"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            ToolsSellerInfoDataGrid.Columns["Skype"].MinimumWidth = 150;
            ToolsSellerInfoDataGrid.Columns["Skype"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            ToolsSellerInfoDataGrid.Columns["Name"].HeaderText = "Имя";
            ToolsSellerInfoDataGrid.Columns["Position"].HeaderText = "Должность";
            ToolsSellerInfoDataGrid.Columns["Phone"].HeaderText = "Телефон";
            ToolsSellerInfoDataGrid.Columns["Email"].HeaderText = "E-mail";
            ToolsSellerInfoDataGrid.Columns["ICQ"].HeaderText = "ICQ";
            ToolsSellerInfoDataGrid.Columns["Skype"].HeaderText = "Skype";

            ToolsSellersDataGrid.Columns["ToolsSellerName"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            ToolsSellersDataGrid.Columns["ToolsSellerName"].MinimumWidth = 190;

            ToolsSellersGroupsDataGrid.Columns["ToolsSellerGroupID"].Visible = false;
            ToolsSellersGroupsDataGrid.Columns["ToolsSellerGroupName"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            ToolsSellersGroupsDataGrid.Columns["ToolsSellerGroupName"].MinimumWidth = 50;

            ToolsSellersSubGroupsDataGrid.Columns["ToolsSellerGroupID"].Visible = false;
            ToolsSellersSubGroupsDataGrid.Columns["ToolsSellerSubGroupID"].Visible = false;
            ToolsSellersSubGroupsDataGrid.Columns["ToolsSellerSubGroupName"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            ToolsSellersSubGroupsDataGrid.Columns["ToolsSellerSubGroupName"].MinimumWidth = 50;
        }

        #region Properties

        public DataTable InfoDataTableClone
        {
            get
            {
                var DT = new DataTable();
                if (ToolsSellerInfoDataTable != null)
                    DT = ToolsSellerInfoDataTable.Clone();
                return DT;
            }
        }

        public DataTable InfoDataTableCopy
        {
            get
            {
                var DT = new DataTable();
                if (ToolsSellerInfoDataTable != null)
                    DT = ToolsSellerInfoDataTable.Copy();
                return DT;
            }
        }

        public int ToolsSellersCount
        {
            get
            {
                var Count = 0;
                if (ToolsSellersBindingSource != null)
                {
                    Count = ToolsSellersBindingSource.Count;
                }
                return Count;
            }
        }

        public int ToolsSubGroupsCount
        {
            get
            {
                var Count = 0;
                if (ToolsSellersSubGroupsBindingSource != null)
                {
                    Count = ToolsSellersSubGroupsBindingSource.Count;
                }
                return Count;
            }
        }

        public string CurrentToolsSellerName
        {
            get
            {
                var ToolsSellerName = string.Empty;
                if (ToolsSellersBindingSource.Count > 0
                    && ToolsSellersBindingSource.Current != null
                    && ((DataRowView)ToolsSellersBindingSource.Current).Row["ToolsSellerName"] != DBNull.Value)
                {
                    ToolsSellerName = ((DataRowView)ToolsSellersBindingSource.Current).Row["ToolsSellerName"].ToString();
                }
                return ToolsSellerName;
            }
        }

        public string CurrentToolsSellerCountry
        {
            get
            {
                var Country = string.Empty;
                if (ToolsSellersBindingSource.Count > 0
                    && ToolsSellersBindingSource.Current != null
                    && ((DataRowView)ToolsSellersBindingSource.Current).Row["Country"] != DBNull.Value)
                {
                    Country = ((DataRowView)ToolsSellersBindingSource.Current).Row["Country"].ToString();
                }
                return Country;
            }
        }

        public string CurrentToolsSellerAddress
        {
            get
            {
                var Address = string.Empty;
                if (ToolsSellersBindingSource.Count > 0
                    && ToolsSellersBindingSource.Current != null
                    && ((DataRowView)ToolsSellersBindingSource.Current).Row["Address"] != DBNull.Value)
                {
                    Address = ((DataRowView)ToolsSellersBindingSource.Current).Row["Address"].ToString();
                }
                return Address;
            }
        }

        public string CurrentContractDocNumber
        {
            get
            {
                var ContractDocNumber = string.Empty;
                if (ToolsSellersBindingSource.Count > 0
                    && ToolsSellersBindingSource.Current != null
                    && ((DataRowView)ToolsSellersBindingSource.Current).Row["ContractDocNumber"] != DBNull.Value)
                {
                    ContractDocNumber = ((DataRowView)ToolsSellersBindingSource.Current).Row["ContractDocNumber"].ToString();
                }
                return ContractDocNumber;
            }
        }

        public string CurrentToolsSellerEmail
        {
            get
            {
                var Email = string.Empty;
                if (ToolsSellersBindingSource.Count > 0
                    && ToolsSellersBindingSource.Current != null
                    && ((DataRowView)ToolsSellersBindingSource.Current).Row["Email"] != DBNull.Value)
                {
                    Email = ((DataRowView)ToolsSellersBindingSource.Current).Row["Email"].ToString();
                }
                return Email;
            }
        }

        public string CurrentToolsSellerSite
        {
            get
            {
                var Site = string.Empty;
                if (ToolsSellersBindingSource.Count > 0
                    && ToolsSellersBindingSource.Current != null
                    && ((DataRowView)ToolsSellersBindingSource.Current).Row["Site"] != DBNull.Value)
                {
                    Site = ((DataRowView)ToolsSellersBindingSource.Current).Row["Site"].ToString();
                }
                return Site;
            }
        }

        public string CurrentToolsSellerNotes
        {
            get
            {
                var Notes = string.Empty;
                if (ToolsSellersBindingSource.Count > 0
                    && ToolsSellersBindingSource.Current != null
                    && ((DataRowView)ToolsSellersBindingSource.Current).Row["Notes"] != DBNull.Value)
                {
                    Notes = ((DataRowView)ToolsSellersBindingSource.Current).Row["Notes"].ToString();
                }
                return Notes;
            }
        }

        public string CurrentToolsSellerGroupName
        {
            get
            {
                var ToolsSellerGroupName = string.Empty;
                if (ToolsSellersGroupsBindingSource.Count > 0
                    && ToolsSellersGroupsBindingSource.Current != null
                    && ((DataRowView)ToolsSellersGroupsBindingSource.Current).Row["ToolsSellerGroupName"] != DBNull.Value)
                {
                    ToolsSellerGroupName = ((DataRowView)ToolsSellersGroupsBindingSource.Current).Row["ToolsSellerGroupName"].ToString();
                }
                return ToolsSellerGroupName;
            }
        }

        public string CurrentToolsSellerSubGroupName
        {
            get
            {
                var ToolsSellerSubGroupName = string.Empty;
                if (ToolsSellersSubGroupsBindingSource.Count > 0
                    && ToolsSellersSubGroupsBindingSource.Current != null
                    && ((DataRowView)ToolsSellersSubGroupsBindingSource.Current).Row["ToolsSellerSubGroupName"] != DBNull.Value)
                {
                    ToolsSellerSubGroupName = ((DataRowView)ToolsSellersSubGroupsBindingSource.Current).Row["ToolsSellerSubGroupName"].ToString();
                }
                return ToolsSellerSubGroupName;
            }
        }

        #endregion

        #region Get functions

        public void GetCurrentSeller()
        {
            if (ToolsSellersBindingSource.Count == 0)
            {
                CurrentToolsSellerID = -1;
                return;
            }
            if (((DataRowView)ToolsSellersBindingSource.Current).Row["ToolsSellerID"] == DBNull.Value)
                return;

            CurrentToolsSellerID = Convert.ToInt32(((DataRowView)ToolsSellersBindingSource.Current).Row["ToolsSellerID"]);
        }

        public void GetCurrentSellerGroup()
        {
            if (ToolsSellersGroupsBindingSource.Count == 0)
            {
                CurrentToolsSellerGroupID = -1;
                return;
            }

            if (((DataRowView)ToolsSellersGroupsBindingSource.Current).Row["ToolsSellerGroupID"] == DBNull.Value)
                return;

            CurrentToolsSellerGroupID = Convert.ToInt32(((DataRowView)ToolsSellersGroupsBindingSource.Current).Row["ToolsSellerGroupID"]);
        }

        public void GetCurrentSellerSubGroup()
        {
            if (ToolsSellersSubGroupsBindingSource.Count == 0)
            {
                CurrentToolsSellerSubGroupID = -1;
                return;
            }
            if (((DataRowView)ToolsSellersSubGroupsBindingSource.Current).Row["ToolsSellerSubGroupID"] == DBNull.Value)
                return;

            CurrentToolsSellerSubGroupID = Convert.ToInt32(((DataRowView)ToolsSellersSubGroupsBindingSource.Current).Row["ToolsSellerSubGroupID"]);
        }

        public void GetToolsSellerInfo()
        {
            ToolsSellerInfoDataTable.Clear();
            if (ToolsSellersBindingSource.Count > 0)
            {
                var ContactsXML = ((DataRowView)ToolsSellersBindingSource.Current).Row["ContactsXML"].ToString();
                if (ContactsXML.Length == 0)
                    return;

                using (var SR = new StringReader(ContactsXML))
                {
                    ToolsSellerInfoDataTable.ReadXml(SR);
                }
            }
        }

        #endregion

        #region Filter functions

        public void FilterSellerSubGroups()
        {
            GetCurrentSellerGroup();
            ToolsSellersSubGroupsBindingSource.Filter = "ToolsSellerGroupID = " + CurrentToolsSellerGroupID;
            ToolsSellersSubGroupsBindingSource.MoveFirst();
        }

        public void FilterSellers()
        {
            GetCurrentSellerSubGroup();
            ToolsSellersBindingSource.Filter = "ToolsSellerSubGroupID = " + CurrentToolsSellerSubGroupID;
            ToolsSellersBindingSource.MoveFirst();
        }

        #endregion

        #region Add functions

        public void AddSeller(string ToolsSellerName, string Country, string Address,
            string ContractDocNumber, string Email, string Site, string ContactsXML, string Notes)
        {
            var NewRow = ToolsSellersDataTable.NewRow();
            NewRow["ToolsSellerSubGroupID"] = CurrentToolsSellerSubGroupID;
            NewRow["ToolsSellerName"] = ToolsSellerName;
            NewRow["Country"] = Country;
            NewRow["Address"] = Address;
            NewRow["ContractDocNumber"] = ContractDocNumber;
            NewRow["Email"] = Email;
            NewRow["Site"] = Site;
            NewRow["ContactsXML"] = ContactsXML;
            NewRow["Notes"] = Notes;
            ToolsSellersDataTable.Rows.Add(NewRow);
            SaveSellers();
            ToolsSellersBindingSource.Position = ToolsSellersDataTable.Rows.Count;
        }

        public void AddSellerGroup(string ToolsSellerGroup)
        {
            var NewRow = ToolsSellersGroupsDataTable.NewRow();
            NewRow["ToolsSellerGroupName"] = ToolsSellerGroup;
            ToolsSellersGroupsDataTable.Rows.Add(NewRow);

            SaveSellerGroups();
            //ToolsSellersGroupsBindingSource.Position = ToolsSellersGroupsDataTable.Rows.Count;
        }

        public void AddSellerSubGroup(string ToolsSellerSubGroup)
        {
            var NewRow = ToolsSellersSubGroupsDataTable.NewRow();
            NewRow["ToolsSellerSubGroupName"] = ToolsSellerSubGroup;
            NewRow["ToolsSellerGroupID"] = CurrentToolsSellerGroupID;
            ToolsSellersSubGroupsDataTable.Rows.Add(NewRow);

            SaveSellerSubGroups();
        }

        #endregion

        #region Edit functions

        public void EditSeller(string ToolsSellerName, string Country, string Address,
            string ContractDocNumber, string Email, string Site, string ContactsXML, string Notes)
        {
            var row = ToolsSellersDataTable.Select("ToolsSellerID = " + CurrentToolsSellerID);
            if (row.Count() > 0)
            {
                row[0]["ToolsSellerName"] = ToolsSellerName;
                row[0]["Country"] = Country;
                row[0]["Address"] = Address;
                row[0]["ContractDocNumber"] = ContractDocNumber;
                row[0]["Email"] = Email;
                row[0]["Site"] = Site;
                row[0]["ContactsXML"] = ContactsXML;
                row[0]["Notes"] = Notes;
            }

            SaveSellers();
            GetToolsSellerInfo();
        }

        public void EditSellerGroup(string ToolsSellerGroup)
        {
            if (ToolsSellersGroupsBindingSource.Count == 0)
                return;

            ToolsSellersGroupsDataTable.Rows[ToolsSellersGroupsDataGrid.CurrentRow.Index]["ToolsSellerGroupName"] = ToolsSellerGroup;

            ToolsSellersGroupsDA.Update(ToolsSellersGroupsDataTable);
        }

        public void EditSellerSubGroup(string SellerSubGroup)
        {
            if (ToolsSellersSubGroupsBindingSource.Count == 0)
                return;

            var row = ToolsSellersSubGroupsDataTable.Select("ToolsSellerSubGroupID = " + CurrentToolsSellerSubGroupID);
            if (row.Count() > 0)
                row[0]["ToolsSellerSubGroupName"] = SellerSubGroup;

            ToolsSellersSubGroupsDA.Update(ToolsSellersSubGroupsDataTable);
        }

        #endregion

        #region Remove functions

        public void RemoveSeller()
        {
            ToolsSellersBindingSource.RemoveCurrent();
            SaveSellers();
        }

        public void RemoveSellerGroup()
        {
            //if (ToolsSellersGroupsBindingSource.Count > 0)
            //{
            //	ToolsSellersGroupsBindingSource.RemoveCurrent();
            //	SaveSellerGroups();
            //}


            foreach (var r1 in ToolsSellersGroupsDataTable.Select("ToolsSellerGroupID = " + CurrentToolsSellerGroupID))
            {
                foreach (var r2 in ToolsSellersSubGroupsDataTable.Select("ToolsSellerGroupID = " + CurrentToolsSellerGroupID))
                {
                    foreach (var r3 in ToolsSellersDataTable.Select("ToolsSellerSubGroupID = " + r2["ToolsSellerSubGroupID"]))
                    {
                        r3.Delete();
                    }
                    r2.Delete();
                }
                r1.Delete();
            }
            SaveSellers();
            SaveSellerSubGroups();
            SaveSellerGroups();
        }

        public void RemoveSellerSubGroup()
        {
            //if (ToolsSellersSubGroupsBindingSource.Count > 0)
            //{
            //	ToolsSellersSubGroupsBindingSource.RemoveCurrent();
            //	SaveSellerSubGroups();
            //}
            foreach (var r2 in ToolsSellersSubGroupsDataTable.Select("ToolsSellerSubGroupID = " + CurrentToolsSellerSubGroupID))
            {
                foreach (var r3 in ToolsSellersDataTable.Select("ToolsSellerSubGroupID = " + r2["ToolsSellerSubGroupID"]))
                {
                    r3.Delete();
                }
                r2.Delete();
            }
            SaveSellers();
            SaveSellerSubGroups();
        }

        #endregion

        #region Save functions

        public void SaveSellers()
        {
            ToolsSellersDA.Update(ToolsSellersDataTable);
            ToolsSellersDataTable.Clear();
            ToolsSellersDA.Fill(ToolsSellersDataTable);
        }

        public void SaveSellerGroups()
        {
            ToolsSellersGroupsDA.Update(ToolsSellersGroupsDataTable);
            ToolsSellersGroupsDataTable.Clear();
            ToolsSellersGroupsDA.Fill(ToolsSellersGroupsDataTable);
        }

        public void SaveSellerSubGroups()
        {
            ToolsSellersSubGroupsDA.Update(ToolsSellersSubGroupsDataTable);
            ToolsSellersSubGroupsDataTable.Clear();
            ToolsSellersSubGroupsDA.Fill(ToolsSellersSubGroupsDataTable);
        }

        #endregion

    }





    public class LightDevelopmentPlan
    {
        public DataTable DevelopmentDataTable, DepartmentTable;
        public BindingSource DevelopmentBindingSource, DepartmentBindingSource;

        //private SqlCommandBuilder CB;
        //private SqlDataAdapter DA;
        private readonly PercentageDataGrid DevelopmentPlanDataGrid;
        private readonly PercentageDataGrid DepartmentDataGrid;

        public LightDevelopmentPlan(ref PercentageDataGrid tDevelopmentPlanDataGrid, ref PercentageDataGrid tDepartmentDataGrid)
        {
            DevelopmentPlanDataGrid = tDevelopmentPlanDataGrid;
            DepartmentDataGrid = tDepartmentDataGrid;
            CreateAndFill();
            Binding();
            GridSettings();
        }

        public void CreateAndFill()
        {
            DevelopmentDataTable = new DataTable();
            using (var DA = new SqlDataAdapter("SELECT * FROM DevelopmentPlan", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(DevelopmentDataTable);
            }

            DepartmentTable = new DataTable();
            using (var DA = new SqlDataAdapter("SELECT DepartmentID, DepartmentName FROM Departments", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(DepartmentTable);
            }
        }

        public void Binding()
        {
            DevelopmentBindingSource = new BindingSource
            {
                DataSource = DevelopmentDataTable
            };
            DevelopmentPlanDataGrid.DataSource = DevelopmentBindingSource;

            DepartmentBindingSource = new BindingSource
            {
                DataSource = DepartmentTable
            };
            DepartmentDataGrid.DataSource = DepartmentBindingSource;
        }

        public void GridSettings()
        {
            DevelopmentPlanDataGrid.Columns["DevelopmentPlanID"].Visible = false;
            DevelopmentPlanDataGrid.Columns["ProjectName"].HeaderText = "Проект";
            DevelopmentPlanDataGrid.Columns["ApproximateDateStart"].Visible = false;
            DevelopmentPlanDataGrid.Columns["ApproximateDateEnd"].Visible = false;
            DevelopmentPlanDataGrid.Columns["AuthorID"].Visible = false;
            //DevelopmentPlanDataGrid.Columns["ApproximateDateStart"].HeaderText = "Примерная\r\nдата начала";
            //DevelopmentPlanDataGrid.Columns["ApproximateDateEnd"].HeaderText = "     Примерная\r\nдата завершения";
            DevelopmentPlanDataGrid.Columns["Coders"].Visible = false;
            DevelopmentPlanDataGrid.Columns["BetaTesters"].Visible = false;
            DevelopmentPlanDataGrid.Columns["Users"].Visible = false;
            DevelopmentPlanDataGrid.Columns["DepartmentID"].Visible = false;
            DevelopmentPlanDataGrid.Columns["ProjectDescription"].Visible = false;
            DevelopmentPlanDataGrid.Columns["FactDateStart"].HeaderText = "Дата начала";
            DevelopmentPlanDataGrid.Columns["FactDateEnd"].HeaderText = "Дата завершения";
            DevelopmentPlanDataGrid.Columns["PercentageComplete"].HeaderText = "Разработано, %";
            DevelopmentPlanDataGrid.Columns["PercentageImplementation"].HeaderText = "Внедрено, %";
            DevelopmentPlanDataGrid.Columns["PercentageVerify"].HeaderText = "Тест, %";
            DevelopmentPlanDataGrid.Columns["IsClosed"].HeaderText = "Завершено";

            foreach (DataGridViewColumn Column in DevelopmentPlanDataGrid.Columns)
            {
                Column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }

            //DevelopmentPlanDataGrid.Columns["ApproximateDateStart"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            //DevelopmentPlanDataGrid.Columns["ApproximateDateStart"].Width = 115;
            //DevelopmentPlanDataGrid.Columns["ApproximateDateEnd"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            //DevelopmentPlanDataGrid.Columns["ApproximateDateEnd"].Width = 145;
            DevelopmentPlanDataGrid.Columns["FactDateStart"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            DevelopmentPlanDataGrid.Columns["FactDateStart"].Width = 130;
            DevelopmentPlanDataGrid.Columns["FactDateEnd"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            DevelopmentPlanDataGrid.Columns["FactDateEnd"].Width = 150;
            DevelopmentPlanDataGrid.Columns["IsClosed"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            DevelopmentPlanDataGrid.Columns["IsClosed"].Width = 100;
            DevelopmentPlanDataGrid.Columns["PercentageComplete"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            DevelopmentPlanDataGrid.Columns["PercentageComplete"].Width = 150;
            DevelopmentPlanDataGrid.Columns["PercentageImplementation"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            DevelopmentPlanDataGrid.Columns["PercentageImplementation"].Width = 150;
            DevelopmentPlanDataGrid.Columns["PercentageVerify"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            DevelopmentPlanDataGrid.Columns["PercentageVerify"].Width = 150;

            DevelopmentPlanDataGrid.AddPercentageColumn("PercentageComplete");
            DevelopmentPlanDataGrid.AddPercentageColumn("PercentageImplementation");
            DevelopmentPlanDataGrid.AddPercentageColumn("PercentageVerify");

            DepartmentDataGrid.Columns["DepartmentID"].Visible = false;
            DepartmentDataGrid.Columns["DepartmentName"].HeaderText = "Отдел";
        }

        public void AddProject(string ProjectName, string Users, string ProjectDescription, string DepartmentID, DateTime ApproximateDateStart, DateTime ApproximateDateEnd)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM DevelopmentPlan", ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        var NewRow = DT.NewRow();
                        NewRow["ProjectName"] = ProjectName;
                        NewRow["FactDateStart"] = Security.GetCurrentDate();
                        NewRow["Users"] = Users;
                        NewRow["AuthorID"] = Security.CurrentUserID;
                        NewRow["PercentageComplete"] = 0;
                        NewRow["PercentageImplementation"] = 0;
                        NewRow["PercentageVerify"] = 0;
                        NewRow["IsClosed"] = false;
                        NewRow["ProjectDescription"] = ProjectDescription;
                        NewRow["DepartmentID"] = DepartmentID;
                        NewRow["ApproximateDateStart"] = ApproximateDateStart;
                        NewRow["ApproximateDateEnd"] = ApproximateDateEnd;
                        DT.Rows.Add(NewRow);
                        DA.Update(DT);
                    }
                }
            }
        }

        public void DeleteProject(string DevelopmentPlanID)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM DevelopmentPlan WHERE DevelopmentPlanID =" + DevelopmentPlanID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);
                        DT.Rows[0].Delete();
                        DA.Update(DT);
                    }
                }
            }
        }

        //public void AddProject(string ProjectName, DateTime ApproximateDateStart, DateTime ApproximateDateEnd, string Coders)
        //{
        //    DataRow NewRow = DevelopmentDataTable.NewRow();
        //    NewRow["ProjectName"] = ProjectName;
        //    NewRow["ApproximateDateStart"] = ApproximateDateStart;
        //    NewRow["ApproximateDateEnd"] = ApproximateDateEnd;
        //    NewRow["Coders"] = Coders;
        //    //NewRow["BetaTesters"] = BetaTesters;
        //    //NewRow["Users"] = Users;
        //    NewRow["PercentageComplete"] = 0;
        //    NewRow["PercentageImplementation"] = 0;
        //    NewRow["PercentageVerify"] = 0;
        //    NewRow["IsClosed"] = false;
        //    DevelopmentDataTable.Rows.Add(NewRow);
        //}

        public void UpdateDevelopmentPlanDataGrid(string DepartmentID)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM DevelopmentPlan WHERE DepartmentID =" + DepartmentID, ConnectionStrings.LightConnectionString))
            {
                DevelopmentDataTable.Clear();
                DA.Fill(DevelopmentDataTable);
            }
        }

        private void UpdateDevelopmentPlanDataGrid()
        {
            var FactDateEnd = default(DateTime);

            if (DevelopmentDataTable.Select("IsClosed = 1 and FactDateEnd is null").Count() > 0)
                FactDateEnd = Security.GetCurrentDate();

            foreach (var RowProject in DevelopmentDataTable.Select("IsClosed = 1 and FactDateEnd is null"))
                RowProject["FactDateEnd"] = FactDateEnd;

            foreach (var RowProject in DevelopmentDataTable.Select("IsClosed = 0 and FactDateEnd is not null"))
                RowProject["FactDateEnd"] = DBNull.Value;
        }

        //public void Save()
        //{
        //    DA.Update(DevelopmentDataTable);
        //    DevelopmentDataTable.Clear();
        //    DA.Fill(DevelopmentDataTable);
        //}

        public void Save()
        {
            UpdateDevelopmentPlanDataGrid();
            using (var DA = new SqlDataAdapter("SELECT * FROM DevelopmentPlan", ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    DA.Update(DevelopmentDataTable);
                }
            }
        }

        public void Delete()
        {
            DevelopmentBindingSource.RemoveCurrent();
        }
    }





    public class Tasks
    {
        public DataTable SelectedUsersTable;

        //public DataTable MyTasksDataTable;
        //private SqlCommandBuilder MyTasksCB;
        //private SqlDataAdapter MyTasksDA;

        public DataTable TasksDataTable;
        public BindingSource TasksBindingSource;
        private SqlCommandBuilder TasksCB;
        private SqlDataAdapter TasksDA;
        private readonly PercentageDataGrid TasksDataGrid;

        public Tasks(ref PercentageDataGrid tTasksDataGrid)
        {
            TasksDataGrid = tTasksDataGrid;

            CreateAndFill();
            Binding();
            GridSettings();

            SelectedUsersTable = new DataTable();
            SelectedUsersTable.Columns.Add("UserID", typeof(long));
        }

        public void CreateAndFill()
        {
            TasksDA = new SqlDataAdapter("SELECT * FROM Tasks", ConnectionStrings.LightConnectionString);
            TasksCB = new SqlCommandBuilder(TasksDA);
            TasksDataTable = new DataTable();
            TasksDA.Fill(TasksDataTable);

            //MyTasksDA = new SqlDataAdapter("SELECT TaskID FROM TaskMembers WHERE PerformerID = " + Security.CurrentUserID, ConnectionStrings.LightConnectionString);
            //MyTasksCB = new SqlCommandBuilder(MyTasksDA);
            //MyTasksDataTable = new DataTable();
            //MyTasksDA.Fill(MyTasksDataTable);
        }

        private void Binding()
        {
            TasksBindingSource = new BindingSource
            {
                DataSource = TasksDataTable
            };
            TasksDataGrid.DataSource = TasksBindingSource;
        }

        public void GridSettings()
        {
            TasksDataGrid.Columns["TaskID"].Visible = false;
            TasksDataGrid.Columns["DirectorID"].Visible = false;
            TasksDataGrid.Columns["TaskDescription"].Visible = false;
            TasksDataGrid.Columns["ReturnDescription"].Visible = false;
            TasksDataGrid.Columns["CreationDate"].Visible = false;
            TasksDataGrid.Columns["ExecutionDate"].Visible = false;
            TasksDataGrid.Columns["FactExecutionDate"].Visible = false;
            TasksDataGrid.Columns["ExecutionStatus"].Visible = false;

            foreach (DataGridViewColumn Column in TasksDataGrid.Columns)
            {
                Column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }

            TasksDataGrid.AddPercentageColumn("CompletPercentage");

            TasksDataGrid.Columns["TaskName"].DisplayIndex = 0;
            TasksDataGrid.Columns["CompletPercentage"].DisplayIndex = 1;

        }

        public DataTable TableUsers()
        {
            using (var DA = new SqlDataAdapter("SELECT UserID, Name FROM Users  WHERE Fired <> 1 ORDER BY Name", ConnectionStrings.UsersConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    return DT;
                }
            }
        }

        //Обновление таблицы TasksDataTable
        public void UpdateTasksGrid(bool Filter)
        {
            if (Filter)
            {
                using (var DA = new SqlDataAdapter("SELECT * FROM Tasks WHERE TaskID IN (SELECT TaskID FROM TaskMembers WHERE PerformerID = " + Security.CurrentUserID + ")", ConnectionStrings.LightConnectionString))
                {
                    TasksDataTable.Clear();
                    DA.Fill(TasksDataTable);
                }
            }
            else
            {
                using (var DA = new SqlDataAdapter("SELECT * FROM Tasks WHERE DirectorID = " + Security.CurrentUserID, ConnectionStrings.LightConnectionString))
                {
                    TasksDataTable.Clear();
                    DA.Fill(TasksDataTable);
                }
            }


        }

        ////Обновление таблицы MyTasksDataTable
        //public void UpdateMyTasksGrid()
        //{
        //    using (SqlDataAdapter DA = new SqlDataAdapter("SELECT TaskID FROM TaskMembers WHERE PerformerID = " + Security.CurrentUserID, ConnectionStrings.LightConnectionString))
        //    {
        //        MyTasksDataTable.Clear();
        //        DA.Fill(MyTasksDataTable);
        //    }
        //}

        //Обновление грида с исполнителями
        public void UpdateUsersGrid(string TaskID)
        {
            using (var DA = new SqlDataAdapter("SELECT PerformerID as UserID FROM TaskMembers WHERE TaskID = " + TaskID, ConnectionStrings.LightConnectionString))
            {
                SelectedUsersTable.Clear();
                DA.Fill(SelectedUsersTable);
            }
        }

        //Добавление задачи
        public void SaveTasks(string TasksName, string TasksDescription, DateTime dataTime)
        {
            using (var DA = new SqlDataAdapter("SELECT [TaskID], [DirectorID], [CreationDate], [ExecutionDate], [TaskName]"
                                               + ",[TaskDescription], [ExecutionStatus], [FactExecutionDate], CompletPercentage FROM Tasks", ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        var NewRow = DT.NewRow();
                        NewRow["DirectorID"] = Security.CurrentUserID;
                        NewRow["CreationDate"] = Security.GetCurrentDate();
                        NewRow["ExecutionDate"] = dataTime;
                        NewRow["TaskName"] = TasksName;
                        NewRow["TaskDescription"] = TasksDescription;
                        NewRow["ExecutionStatus"] = false;
                        NewRow["CompletPercentage"] = 0;
                        DT.Rows.Add(NewRow);
                        DA.Update(DT);

                        int TaskID;
                        using (var DA2 = new SqlDataAdapter("SELECT TaskID FROM Tasks where CreationDate = '" + ((DateTime)NewRow["CreationDate"]).ToString("yyyy-MM-dd HH:mm:ss.fff") + "'", ConnectionStrings.LightConnectionString))
                        {
                            using (var DT2 = new DataTable())
                            {
                                DA2.Fill(DT2);
                                TaskID = Convert.ToInt32(DT2.Rows[0][0]);
                            }
                        }

                        SaveTasksMembers(TaskID);
                    }
                }
            }
        }

        public void SaveTasksMembers(int TaskID)
        {
            using (var DA = new SqlDataAdapter("SELECT [TaskMembersID],[TaskID] ,[PerformerID] FROM TaskMembers", ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        for (var i = 0; i < SelectedUsersTable.Rows.Count; i++)
                        {
                            var NewRow = DT.NewRow();
                            NewRow["PerformerID"] = SelectedUsersTable.Rows[i][0];
                            NewRow["TaskID"] = TaskID;
                            DT.Rows.Add(NewRow);
                            if (Security.CurrentUserID != Convert.ToInt32(SelectedUsersTable.Rows[i][0]))
                                ActiveNotifySystem.CreateSubscribeRecordForOneUser("TasksButton", TaskID, Convert.ToInt32(SelectedUsersTable.Rows[i][0]));
                        }

                        DA.Update(DT);
                    }
                }
            }
        }

        //Удаление задачи
        public void DeleteTasks(string TaskID)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM Tasks WHERE TaskID =" + TaskID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);
                        DT.Rows[0].Delete();
                        DA.Update(DT);
                    }
                }
            }

            DeleteMembers(TaskID);

            SelectedUsersTable.Clear();
        }

        public void DeleteMembers(string TaskID)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM TaskMembers WHERE TaskID =" + TaskID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);
                        for (var i = 0; i < DT.Rows.Count; i++)
                            DT.Rows[i].Delete();
                        DA.Update(DT);
                    }
                }
            }
        }

        //Обновление задачи
        public void UpdateTasks(string TaskID, string TasksName, string TasksDescription, DateTime dataTime)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM Tasks WHERE TaskID =" + TaskID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        DT.Rows[0]["TaskName"] = TasksName;
                        DT.Rows[0]["TaskDescription"] = TasksDescription;
                        DT.Rows[0]["ExecutionDate"] = dataTime;

                        DA.Update(DT);
                    }

                    DeleteMembers(TaskID);
                    SaveTasksMembers(Convert.ToInt32(TaskID));
                }
            }
        }

        //Изменение процента готовности
        public void UpdateTasks(string TaskID, int CompletPercentage)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM Tasks WHERE TaskID =" + TaskID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        DT.Rows[0]["CompletPercentage"] = CompletPercentage;

                        DA.Update(DT);
                    }

                    DeleteMembers(TaskID);
                    SaveTasksMembers(Convert.ToInt32(TaskID));
                }
            }
        }

        //Завершение задачи
        public void CloseTasks(string CloseNotes, string TaskID)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM Tasks WHERE TaskID =" + TaskID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        DT.Rows[0]["FactExecutionDate"] = Security.GetCurrentDate().Date;
                        DT.Rows[0]["ExecutionStatus"] = true;
                        DT.Rows[0]["ReturnDescription"] = CloseNotes;
                        DT.Rows[0]["CompletPercentage"] = 100;

                        DA.Update(DT);
                    }
                }
            }
        }
    }




    public class DayPlannerTimesheet
    {
        private decimal Rate;
        private string StrRate = "0";

        public DataTable _timesheetDataTable;
        private DataTable _absJournalDataTable;
        private DataTable _absTypesTable;
        private DataTable _prodSheduleDataTable;
        private DataTable _absDayDataTable;
        private Excel Ex;

        public DayPlannerTimesheet()
        {
            Create();
            Fill();
        }

        private void Create()
        {
            _timesheetDataTable = new DataTable();
            _absJournalDataTable = new DataTable();
            _absTypesTable = new DataTable();
            _prodSheduleDataTable = new DataTable();
            _absDayDataTable = new DataTable();
            _absDayDataTable.Columns.Add(new DataColumn("AbsenceTypeID", Type.GetType("System.Int32")));
            _absDayDataTable.Columns.Add(new DataColumn("Hours", Type.GetType("System.Decimal")));
        }

        private void Fill()
        {
            using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM WorkDays ORDER BY DayStartDateTime",
                ConnectionStrings.LightConnectionString))
            {
                DA.Fill(_timesheetDataTable);
            }
            using (var da = new SqlDataAdapter(@"SELECT TOP 0 * FROM AbsencesJournal", ConnectionStrings.LightConnectionString))
            {
                da.Fill(_absJournalDataTable);
            }
            using (var da = new SqlDataAdapter(@"SELECT * FROM AbsenceTypes", ConnectionStrings.LightConnectionString))
            {
                da.Fill(_absTypesTable);
            }
            using (var da = new SqlDataAdapter(@"SELECT TOP 0 * FROM ProductionShedule", ConnectionStrings.LightConnectionString))
            {
                da.Fill(_prodSheduleDataTable);
            }
        }

        public DataTable DayStartDate()
        {
            using (var DA = new SqlDataAdapter("SELECT DayStartDateTime FROM WorkDays WHERE UserID = " + Security.CurrentUserID,
                ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                    return DT;
                }
            }
        }

        public void GetTimesheet(int userId, int Yearint, int Monthint)
        {
            //using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM WorkDays WHERE DATEPART(year,DayStartD/*ateTime)=" + Yearint + " and DATEPART(month,DayStartDateTime)=" + Monthint + " and UserID = " + userId + " ORDER BY DayStartDateTime",
            using (var DA = new SqlDataAdapter("SELECT * FROM WorkDays WHERE DATEPART(year,DayStartDateTime)=" + Yearint + " and UserID = " + userId + " ORDER BY DayStartDateTime",
                ConnectionStrings.LightConnectionString))
            {
                _timesheetDataTable.Clear();
                DA.Fill(_timesheetDataTable);
            }
        }

        public void ExportToExcel(DataGridView TimeSheetDataGrid)
        {
            ClearReport();
            Ex = new Excel();
            Ex.NewDocument(1);

            for (var i = 0; i < TimeSheetDataGrid.Columns.Count; i++)
            {
                Ex.WriteCell(1, TimeSheetDataGrid.Columns[i].HeaderText, 1, i + 1, 12, true);
                Ex.SetHorisontalAlignment(1, 1, i + 1, Excel.AlignHorizontal.xlCenter);
                Ex.SetBorderStyle(1, 1, i + 1, false, true, false, true, Excel.LineStyle.xlContinuous, Excel.BorderWeight.xlThin);
                if (TimeSheetDataGrid.Columns[i].DefaultCellStyle.BackColor == Color.Yellow)
                    for (var j = 0; j <= TimeSheetDataGrid.Rows.Count; j++)
                        Ex.SetColor(1, j + 1, i + 1, Excel.Color.Yellow);
            }

            for (var i = 1; i < TimeSheetDataGrid.Rows.Count + 1; i++)
            {
                Ex.WriteCell(1, TimeSheetDataGrid.Rows[i - 1].Cells[0].Value, i + 1, 1, 12, true);
                Ex.SetBorderStyle(1, i + 1, 1, false, true, false, true, Excel.LineStyle.xlContinuous, Excel.BorderWeight.xlThin);
            }

            for (var i = 1; i < TimeSheetDataGrid.Rows.Count + 1; i++)
            {
                for (var j = 1; j < TimeSheetDataGrid.Columns.Count - 1; j++)
                {
                    if (TimeSheetDataGrid.Rows[i - 1].Cells[j].Value.ToString() != "")
                    {
                        Ex.WriteCell(1, double.Parse(TimeSheetDataGrid.Rows[i - 1].Cells[j].Value.ToString()), i + 1, j + 1, 12, false);
                    }
                    Ex.SetHorisontalAlignment(1, i + 1, j + 1, Excel.AlignHorizontal.xlCenter);
                    Ex.SetBorderStyle(1, i + 1, j + 1, false, true, false, true, Excel.LineStyle.xlContinuous, Excel.BorderWeight.xlThin);
                }
            }

            for (var i = 1; i < TimeSheetDataGrid.Rows.Count + 1; i++)
            {
                if (TimeSheetDataGrid.Rows[i - 1].Cells[TimeSheetDataGrid.Columns.Count - 1].Value.ToString() != "")
                {
                    Ex.WriteCell(1, double.Parse(TimeSheetDataGrid.Rows[i - 1].Cells[TimeSheetDataGrid.Columns.Count - 1].Value.ToString()), i + 1, TimeSheetDataGrid.Columns.Count, 12, true);
                }
                Ex.SetHorisontalAlignment(1, i + 1, TimeSheetDataGrid.Columns.Count, Excel.AlignHorizontal.xlCenter);
                Ex.SetBorderStyle(1, i + 1, TimeSheetDataGrid.Columns.Count, false, true, false, true, Excel.LineStyle.xlContinuous, Excel.BorderWeight.xlThin);
            }
            Ex.AutoFit(1, 1, 1, TimeSheetDataGrid.Rows.Count + 1, 1);
            Ex.Visible = true;
        }

        public void GetAbsJournal(int userId, int Yearint, int Monthint)
        {
            //int Monthint = Convert.ToDateTime(month + " " + year).Month;
            //int Yearint = Convert.ToInt32(year);
            //using (SqlDataAdapter da = new SqlDataAdapter(@"SELECT * FROM AbsencesJournal WHERE" +
            //    " UserID=" + userId + " AND ((DATEPART(month, DateStart) = " + Monthint + " AND DATEPART(year, DateStart) = " + Yearint +
            //    ") OR (DATEPART(month, DateFinish) = " + Monthint + " AND DATEPART(year, DateFinish) = " + Yearint + "))",
            using (var da = new SqlDataAdapter(@"SELECT * FROM AbsencesJournal WHERE" +
                                               " UserID=" + userId + " AND ((DATEPART(year, DateStart) = " + Yearint +
                                               ") OR (DATEPART(year, DateFinish) = " + Yearint + "))", ConnectionStrings.LightConnectionString))
            {
                _absJournalDataTable.Clear();
                da.Fill(_absJournalDataTable);
            }
        }

        private bool AbsenceInThatDay(DateTime date)
        {
            var absenceTypeId = 0;
            decimal absenceHour = 0;
            var b = false;
            _absDayDataTable.Clear();

            for (var i = 0; i < _absJournalDataTable.Rows.Count; i++)
            {
                var dateStart = Convert.ToDateTime(_absJournalDataTable.Rows[i]["DateStart"]);
                var dateFinish = Convert.ToDateTime(_absJournalDataTable.Rows[i]["DateFinish"]);

                if (date.Date >= dateStart.Date && date.Date <= dateFinish.Date)
                {
                    absenceTypeId = Convert.ToInt32(_absJournalDataTable.Rows[i]["AbsenceTypeID"]);
                    absenceHour = Convert.ToDecimal(_absJournalDataTable.Rows[i]["Hours"]);
                    b = true;
                    var dataRow = _absDayDataTable.NewRow();
                    dataRow["AbsenceTypeID"] = absenceTypeId;
                    dataRow["Hours"] = absenceHour;
                    _absDayDataTable.Rows.Add(dataRow);
                }
            }

            //var tuple = new Tuple<bool, int, decimal>(b, absenceTypeId, absenceHour);
            return b;
        }
        
        public bool IsAbsenceVacation(DateTime date)
        {
            var absenceTypeId = 0;
            var b = false;

            if (_timesheetDataTable.Rows.Count == 0)
                return false;

            DateTime LastDateTime;
            LastDateTime = (DateTime)_timesheetDataTable.Rows[_timesheetDataTable.Rows.Count - 1]["DayStartDateTime"];

            for (var i = 0; i < _absJournalDataTable.Rows.Count; i++)
            {
                var dateStart = Convert.ToDateTime(_absJournalDataTable.Rows[i]["DateStart"]);
                var dateFinish = Convert.ToDateTime(_absJournalDataTable.Rows[i]["DateFinish"]);

                if (LastDateTime.Date >= dateStart.Date && LastDateTime.Date <= dateFinish.Date)
                {
                    absenceTypeId = Convert.ToInt32(_absJournalDataTable.Rows[i]["AbsenceTypeID"]);
                    if (absenceTypeId == 2)
                        b = true;
                }
            }
            return b;
        }

        public (bool, DateTime) GetLastDateTime()
        {
            bool b = false;
            DateTime LastDateTime = default;
            if (_timesheetDataTable.Rows.Count > 0)
            {
                b = true;
                LastDateTime = (DateTime)_timesheetDataTable.Rows[_timesheetDataTable.Rows.Count - 1]["DayStartDateTime"];
            }

            return (b, LastDateTime);
        }

        private Tuple<bool, decimal, decimal, decimal, bool> WorkdayInThatDay(DateTime date)
        {
            decimal TimeWorkHours = 0;
            decimal TimeBreakHours = 0;
            decimal TimesheetHours = 0;
            var saved = false;
            var b = false;

            for (var i = 0; i < _timesheetDataTable.Rows.Count; i++)
            {
                if (_timesheetDataTable.Rows[i]["DayEndDateTime"] == DBNull.Value) //если рабочий день не закончен
                    continue;

                DateTime DayStartDateTime;
                DayStartDateTime = (DateTime)_timesheetDataTable.Rows[i]["DayStartDateTime"];
                if (DayStartDateTime.Date != date.Date) //если не тот рабочий день
                    continue;

                DateTime DayEndDateTime;
                TimeSpan TimeWork;

                DayEndDateTime = (DateTime)_timesheetDataTable.Rows[i]["DayEndDateTime"];
                TimesheetHours = (decimal)_timesheetDataTable.Rows[i]["TimesheetHours"];
                saved = (bool)_timesheetDataTable.Rows[i]["Saved"];
                TimeWork = DayEndDateTime.TimeOfDay - DayStartDateTime.TimeOfDay;
                TimeWorkHours = (decimal)Math.Round(TimeWork.TotalHours, 1);

                if (_timesheetDataTable.Rows[i]["DayBreakStartDateTime"] != DBNull.Value && _timesheetDataTable.Rows[i]["DayBreakEndDateTime"] != DBNull.Value) //если  был обед
                {
                    DateTime DayBreakStartDateTime;
                    DateTime DayBreakEndDateTime;
                    TimeSpan TimeBreak;

                    DayBreakStartDateTime = (DateTime)_timesheetDataTable.Rows[i]["DayBreakStartDateTime"];
                    DayBreakEndDateTime = (DateTime)_timesheetDataTable.Rows[i]["DayBreakEndDateTime"];
                    TimeBreak = DayBreakEndDateTime.TimeOfDay - DayBreakStartDateTime.TimeOfDay;
                    TimeBreakHours = (decimal)Math.Round(TimeBreak.TotalHours, 1);
                }

                b = true;
                break;
            }

            var tuple = new Tuple<bool, decimal, decimal, decimal, bool>(b, TimeWorkHours, TimeBreakHours, TimesheetHours, saved);
            return tuple;
        }

        public void CalcOverwork(int Yearint, int Monthint, DateTime dateToday)
        {
            Labels = new List<TimesheetInfo>();

            decimal AllTimesheetHours = 0; // планово до сегодняшнего дня
            decimal AllPlanHours = 0; // планово до сегодняшнего дня
            decimal AllFactHours = 0; // рабочего времени до сегодняшнего дня
            decimal AllAbsenceHours = 0;
            decimal AbsenteeismHours = 0; // прогул и отгул
            decimal AllAbsenteeismHours = 0; // прогул и отгул
            decimal OvertimeHours = 0; // сверхурочные сегодня
            decimal OverworkHours = 0; // переработка
            decimal AllOvertimeHours = 0; // все сверурочные часы

            for (var i = 0; i < _prodSheduleDataTable.Rows.Count; i++)
            {
                if (Convert.ToInt32(_prodSheduleDataTable.Rows[i]["Day"]) > DateTime.DaysInMonth(Yearint, Convert.ToInt32(_prodSheduleDataTable.Rows[i]["Month"])))
                    continue;

                if (new DateTime(Yearint, Convert.ToInt32(_prodSheduleDataTable.Rows[i]["Month"]), Convert.ToInt32(_prodSheduleDataTable.Rows[i]["Day"])) < new DateTime(2021, 2, 14))
                    continue;

                AbsenteeismHours = 0;
                OvertimeHours = 0;
                var date = new DateTime(Yearint, Convert.ToInt32(_prodSheduleDataTable.Rows[i]["Month"]), Convert.ToInt32(_prodSheduleDataTable.Rows[i]["Day"]));

                var isAbsence = AbsenceInThatDay(date);
                decimal absenceHour = 0;
                var AbsenceShortName = string.Empty;
                var AbsenceFullName = string.Empty;
                decimal AbsenceHours = 0;

                var workdayTuple = WorkdayInThatDay(date);
                var isWorkday = workdayTuple.Item1;
                var workHours = workdayTuple.Item2;
                var breakHours = workdayTuple.Item3;
                var timesheetHours = workdayTuple.Item4;
                var saved = workdayTuple.Item5;
                var factHours = workHours - breakHours;

                var prodSheduleHours = GetHourInProdShedule(date) * Rate;

                for (var j = 0; j < _absDayDataTable.Rows.Count; j++)
                {
                    var absenceTypeId = Convert.ToInt32(_absDayDataTable.Rows[j]["AbsenceTypeID"]);
                    absenceHour = Convert.ToDecimal(_absDayDataTable.Rows[j]["Hours"]);
                    AbsenceShortName += GetAbsenceShortName(absenceTypeId) + " ";
                    AbsenceFullName += string.Format("{0}/({1})", GetAbsenceShortName(absenceTypeId), Convert.ToDecimal(absenceHour.ToString("0.####"))) + " ";

                    if (absenceTypeId == 14)
                    {
                        OvertimeHours += absenceHour;
                    }

                    if (absenceTypeId != 14 && absenceTypeId != 12)
                    {
                        AbsenceHours += absenceHour;
                        //AllAbsenceHours += absenceHour;
                    }

                    if (absenceTypeId == 12 || absenceTypeId == 13)
                    {
                        AbsenteeismHours += absenceHour;
                        //AllAbsenteeismHours += absenceHour;
                    }
                    if (absenceTypeId != 13 && absenceTypeId != 12)
                    {
                        if (prodSheduleHours != 0)
                            AllAbsenceHours += absenceHour;
                    }
                }

                AllOvertimeHours = OvertimeHours;
                if (date == dateToday.Date) //если это выбранный день
                {

                }
                else
                {
                    AllTimesheetHours += timesheetHours;
                    AllPlanHours += prodSheduleHours;
                    AllFactHours += factHours;
                    OverworkHours = AllTimesheetHours - AllPlanHours + AllAbsenceHours;
                }

                if (timesheetHours - factHours > 0)
                {
                    AllOvertimeHours += -(timesheetHours - factHours);
                    if (AllOvertimeHours >= timesheetHours - factHours)
                    {
                    }
                }
                if (Labels.Count == 0)
                {
                    OverworkHours = 0;
                }

                var dayInfo = new TimesheetInfo
                {
                    AllAbsenteeismHours = AllAbsenteeismHours,
                    AllAbsenceHours = AllAbsenceHours,
                    AAbsenceHours = AbsenceHours,
                    Date = date,
                    OvertimeHours = Convert.ToDecimal(OvertimeHours.ToString("0.####")),
                    AbsenteeismHours = Convert.ToDecimal(AbsenteeismHours.ToString("0.####")),
                    IsAbsence = isAbsence,
                    AbsenceFullName = AbsenceFullName,
                    AbsenceShortName = AbsenceShortName,
                    StrRate = StrRate,
                    AbsenceHours = Convert.ToDecimal(absenceHour.ToString("0.####")),
                    PlanHours = Convert.ToDecimal((prodSheduleHours).ToString("0.####")),
                    FactHours = factHours,
                    BreakHours = breakHours,
                    TimesheetHours = timesheetHours,
                    AllOvertimeHours = AllOvertimeHours,
                    AllFactHours = AllFactHours,
                    AllPlanHours = AllPlanHours,
                    Saved = saved,
                    OverworkHours = Convert.ToDecimal((OverworkHours * Rate).ToString("0.####"))
                };

                Labels.Add(dayInfo);
            }
        }

        public TimesheetInfo GetDayInfo(DateTime date)
        {
            var dayInfo = new TimesheetInfo();
            dayInfo = Labels.Find(item => item.Date == date);
            return dayInfo;
        }

        public void GetProdShedule(int Yearint, int Monthint)
        {
            //using (SqlDataAdapter da = new SqlDataAdapter(@"SELECT * FROM ProductionShedule WHERE" +
            //    " Year = " + Yearint + " and Month = " + Monthint, ConnectionStrings.LightConnectionString))
            using (var da = new SqlDataAdapter(@"SELECT * FROM ProductionShedule WHERE" +
                                               " Year = " + Yearint, ConnectionStrings.LightConnectionString))
            {
                _prodSheduleDataTable.Clear();
                da.Fill(_prodSheduleDataTable);
            }
        }

        private string GetAbsenceFullName(int id)
        {
            var name = string.Empty;
            var rows = _absTypesTable.Select("AbsenceTypeID = " + id);
            if (rows.Count() > 0)
                name = rows[0]["Description"].ToString();
            return name;
        }

        private string GetAbsenceShortName(int id)
        {
            var name = string.Empty;
            var rows = _absTypesTable.Select("AbsenceTypeID = " + id);
            if (rows.Count() > 0)
                name = rows[0]["ShortName"].ToString();
            return name;
        }

        private int GetHourInProdShedule(DateTime date)
        {
            var hour = -1;
            var rows = _prodSheduleDataTable.Select("Year = " + date.Year + " AND Month=" + date.Month + " AND Day=" + date.Day);
            if (rows.Count() > 0)
                hour = Convert.ToInt32(rows[0]["Hour"]);
            return hour;
        }

        public void GetRate(int userId)
        {
            var MyQueryText = @"SELECT StaffListID, FactoryID, PositionID, UserID, Rate FROM StaffList
                WHERE UserID=" + userId;

            using (var da = new SqlDataAdapter(MyQueryText, ConnectionStrings.LightConnectionString))
            {
                using (var dt = new DataTable())
                {
                    Rate = 0;
                    if (da.Fill(dt) == 1)
                    {
                        Rate = Convert.ToDecimal(dt.Rows[0]["Rate"]);
                        StrRate = Rate.ToString();

                    }
                    if (da.Fill(dt) == 2)
                    {
                        Rate = Convert.ToDecimal(dt.Rows[0]["Rate"]) + Convert.ToDecimal(dt.Rows[1]["Rate"]);
                        StrRate = dt.Rows[0]["Rate"] + "+" + dt.Rows[1]["Rate"];
                    }
                }
            }
        }

        public void ClearReport()
        {
            if (Ex != null)
            {
                Ex.Dispose();
                Ex = null;
            }
        }

        private List<TimesheetInfo> Labels;
    }


    public class ResultTimesheet
    {
        private string StrRate = "0";

        private DataTable _timesheetDataTable;
        private DataTable _absJournalDataTable;
        private DataTable _absTypesTable;
        private DataTable _prodSheduleDataTable;
        private DataTable _absDayDataTable;
        private DataTable _workDaysDataTable;
        private DataTable _usersDataTable;

        public ResultTimesheet()
        {
            Create();
            Fill();
        }

        private void Create()
        {
            _timesheetDataTable = new DataTable();
            _absJournalDataTable = new DataTable();
            _absTypesTable = new DataTable();
            _prodSheduleDataTable = new DataTable();
            _absDayDataTable = new DataTable();
            _absDayDataTable.Columns.Add(new DataColumn("AbsenceTypeID", Type.GetType("System.Int32")));
            _absDayDataTable.Columns.Add(new DataColumn("Hours", Type.GetType("System.Decimal")));

            _workDaysDataTable = new DataTable();
            _workDaysDataTable.Columns.Add(new DataColumn("Date", Type.GetType("System.DateTime")));
            _workDaysDataTable.Columns.Add(new DataColumn("ProdHours", Type.GetType("System.Int32")));
            _workDaysDataTable.Columns.Add(new DataColumn("ProdHoursString", Type.GetType("System.String")));
            _workDaysDataTable.Columns.Add(new DataColumn("TimesheetHoursExp", Type.GetType("System.String")));
            _workDaysDataTable.Columns.Add(new DataColumn("TimesheetHours", Type.GetType("System.Decimal")));
            _workDaysDataTable.Columns.Add(new DataColumn("FactHours", Type.GetType("System.Decimal")));
            _workDaysDataTable.Columns.Add(new DataColumn("OverworkHours", Type.GetType("System.Decimal")));
            _workDaysDataTable.Columns.Add(new DataColumn("FactOverworkHours", Type.GetType("System.Decimal")));
            var column1 = new DataColumn("IsWorkDay")
            {
                DataType = Type.GetType("System.Boolean"),
                DefaultValue = 0
            };
            var column2 = new DataColumn("IsAbsence")
            {
                DataType = Type.GetType("System.Boolean"),
                DefaultValue = 0
            };
            _workDaysDataTable.Columns.Add(column1);
            _workDaysDataTable.Columns.Add(column2);

            _usersDataTable = new DataTable();
            _usersDataTable.Columns.Add(new DataColumn("UserID", Type.GetType("System.Int32")));
            _usersDataTable.Columns.Add(new DataColumn("ShortName", Type.GetType("System.String")));
        }

        private void Fill()
        {
            using (var DA = new SqlDataAdapter("SELECT TOP 0 WorkDays.*, U.Name, U.ShortName FROM WorkDays INNER JOIN infiniu2_users.dbo.Users as U ON WorkDays.UserID=U.UserID",
                ConnectionStrings.LightConnectionString))
            {
                DA.Fill(_timesheetDataTable);
            }
            using (var da = new SqlDataAdapter(@"SELECT TOP 0 * FROM AbsencesJournal", ConnectionStrings.LightConnectionString))
            {
                da.Fill(_absJournalDataTable);
            }
            using (var da = new SqlDataAdapter(@"SELECT * FROM AbsenceTypes", ConnectionStrings.LightConnectionString))
            {
                da.Fill(_absTypesTable);
            }
            using (var da = new SqlDataAdapter(@"SELECT TOP 0 * FROM ProductionShedule", ConnectionStrings.LightConnectionString))
            {
                da.Fill(_prodSheduleDataTable);
            }
        }

        private void GetTimesheet(int Yearint, int Monthint)
        {
            using (var DA = new SqlDataAdapter("SELECT WorkDays.*, U.Name, U.ShortName FROM WorkDays INNER JOIN infiniu2_users.dbo.Users as U ON WorkDays.UserID=U.UserID WHERE DATEPART(year,DayStartDateTime)=" + Yearint + " and DATEPART(month,DayStartDateTime)=" + Monthint + " ORDER BY DayStartDateTime",
            //using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM WorkDays WHERE DATEPART(year,DayStartDateTime)=" + Yearint + " and UserID = " + userId + " ORDER BY DayStartDateTime",
                ConnectionStrings.LightConnectionString))
            {
                _timesheetDataTable.Clear();
                DA.Fill(_timesheetDataTable);
            }
        }

        private void GetUsers()
        {
            _usersDataTable.Clear();
            using (var DV = new DataView(_timesheetDataTable, string.Empty, "ShortName", DataViewRowState.CurrentRows))
            {
                _usersDataTable = DV.ToTable(true, "UserID", "ShortName");
            }
        }

        public void CreateUsersList(int Yearint, int Monthint, DateTime dateToday)
        {
            var currentMonth = 1;
            yearTimesheets = new List<OneMonthTimesheet>();
            while (currentMonth <= Monthint)
            {
                GetTimesheet(Yearint, currentMonth);
                GetUsers();
                GetProdShedule(Yearint, currentMonth);

                var AllProdHours = 0;
                monthTimesheet = new List<UserTimesheetInfo>();
                for (var i = 0; i < _usersDataTable.Rows.Count; i++)
                {
                    var UserID = Convert.ToInt32(_usersDataTable.Rows[i]["UserID"]);
                    var ShortName = _usersDataTable.Rows[i]["ShortName"].ToString();

                    var staffInfoTuple = StaffInfo(UserID);

                    var PositionName = staffInfoTuple.Item1;
                    var PositionID = staffInfoTuple.Item2;
                    var Rank = staffInfoTuple.Item3;
                    var Rate = staffInfoTuple.Item4;

                    GetAbsJournal(UserID, Yearint, currentMonth);
                    var table = GetWorkDaysTable(UserID, Rate, Yearint, currentMonth, dateToday);
                    var AllTimesheetDays = 0;
                    decimal AllTimesheetHours = 0;
                    decimal AllFactHours = 0;
                    AllProdHours = 0;
                    for (var x = 0; x < table.Rows.Count; x++)
                    {
                        if (Convert.ToBoolean(table.Rows[x]["IsWorkday"]))
                            AllTimesheetDays++;
                        AllProdHours += Convert.ToInt32(table.Rows[x]["ProdHours"]);
                        AllTimesheetHours += Convert.ToDecimal(table.Rows[x]["TimesheetHours"]);
                        AllFactHours += Convert.ToDecimal(table.Rows[x]["FactHours"]);
                    }
                    var AllOverworkHours = Convert.ToDecimal(table.Rows[table.Rows.Count - 1]["OverworkHours"]);
                    var AllFactOverworkHours = Convert.ToDecimal(table.Rows[table.Rows.Count - 1]["FactOverworkHours"]);
                    AllTimesheetHours = Convert.ToDecimal(AllTimesheetHours.ToString("0.####"));
                    AllFactHours = Convert.ToDecimal(AllFactHours.ToString("0.####"));
                    AllOverworkHours = Convert.ToDecimal(AllOverworkHours.ToString("0.####"));
                    AllFactOverworkHours = Convert.ToDecimal(AllFactOverworkHours.ToString("0.####"));

                    var timesheetInfo = new UserTimesheetInfo
                    {
                        UserID = UserID,
                        Name = ShortName,
                        PositionName = PositionName,
                        PositionID = PositionID,
                        Rank = Rank,
                        Rate = Rate,
                        WorkDaysDT = table,
                        AllTimesheetHours = AllTimesheetHours,
                        AllFactHours = AllFactHours,
                        AllOverworkHours = AllOverworkHours,
                        AllFactOverworkHours = AllFactOverworkHours,
                        AllTimesheetDays = AllTimesheetDays

                    };
                    monthTimesheet.Add(timesheetInfo);
                }

                var oneMonthTimesheet = new OneMonthTimesheet
                {
                    Month = currentMonth,
                    Year = Yearint,
                    AllProdHours = AllProdHours,
                    userTimesheets = monthTimesheet
                };

                yearTimesheets.Add(oneMonthTimesheet);

                currentMonth++;
            }
        }

        public void GetAbsJournal(int userId, int Yearint, int Monthint)
        {
            //int Monthint = Convert.ToDateTime(month + " " + year).Month;
            //int Yearint = Convert.ToInt32(year);
            using (var da = new SqlDataAdapter(@"SELECT * FROM AbsencesJournal WHERE" +
                                               " UserID=" + userId + " AND ((DATEPART(month, DateStart) = " + Monthint + " AND DATEPART(year, DateStart) = " + Yearint +
                                               ") OR (DATEPART(month, DateFinish) = " + Monthint + " AND DATEPART(year, DateFinish) = " + Yearint + "))", ConnectionStrings.LightConnectionString))
            {
                _absJournalDataTable.Clear();
                da.Fill(_absJournalDataTable);
            }
        }

        private bool AbsenceInThatDay(DateTime date)
        {
            var absenceTypeId = 0;
            decimal absenceHour = 0;
            var b = false;
            _absDayDataTable.Clear();

            for (var i = 0; i < _absJournalDataTable.Rows.Count; i++)
            {
                var dateStart = Convert.ToDateTime(_absJournalDataTable.Rows[i]["DateStart"]);
                var dateFinish = Convert.ToDateTime(_absJournalDataTable.Rows[i]["DateFinish"]);

                if (date.Date >= dateStart.Date && date.Date <= dateFinish.Date)
                {
                    absenceTypeId = Convert.ToInt32(_absJournalDataTable.Rows[i]["AbsenceTypeID"]);
                    absenceHour = Convert.ToDecimal(_absJournalDataTable.Rows[i]["Hours"]);
                    if (absenceTypeId != 14)
                        b = true;
                    var dataRow = _absDayDataTable.NewRow();
                    dataRow["AbsenceTypeID"] = absenceTypeId;
                    dataRow["Hours"] = absenceHour;
                    _absDayDataTable.Rows.Add(dataRow);
                }
            }

            //var tuple = new Tuple<bool, int, decimal>(b, absenceTypeId, absenceHour);
            return b;
        }

        private Tuple<bool, decimal, decimal> WorkdayInThatDay(int UserID, DateTime date)
        {
            decimal TimesheetHours = 0;
            decimal factHours = 0;
            var b = false;

            for (var i = 0; i < _timesheetDataTable.Rows.Count; i++)
            {
                if (Convert.ToInt32(_timesheetDataTable.Rows[i]["UserID"]) != UserID)
                    continue;

                if (_timesheetDataTable.Rows[i]["DayEndDateTime"] == DBNull.Value) //если рабочий день не закончен
                    continue;

                DateTime DayStartDateTime;
                DateTime DayEndDateTime;
                DayStartDateTime = (DateTime)_timesheetDataTable.Rows[i]["DayStartDateTime"];
                DayEndDateTime = (DateTime)_timesheetDataTable.Rows[i]["DayEndDateTime"];
                if (DayStartDateTime.Date != date.Date) //если не тот рабочий день
                    continue;

                var TimeWork = DayEndDateTime.TimeOfDay - DayStartDateTime.TimeOfDay;
                var TimeWorkHours = (decimal)Math.Round(TimeWork.TotalHours, 1);
                decimal TimeBreakHours = 0;
                if (_timesheetDataTable.Rows[i]["DayBreakStartDateTime"] != DBNull.Value && _timesheetDataTable.Rows[i]["DayBreakEndDateTime"] != DBNull.Value) //если  был обед
                {
                    DateTime DayBreakStartDateTime;
                    DateTime DayBreakEndDateTime;
                    TimeSpan TimeBreak;

                    DayBreakStartDateTime = (DateTime)_timesheetDataTable.Rows[i]["DayBreakStartDateTime"];
                    DayBreakEndDateTime = (DateTime)_timesheetDataTable.Rows[i]["DayBreakEndDateTime"];
                    TimeBreak = DayBreakEndDateTime.TimeOfDay - DayBreakStartDateTime.TimeOfDay;
                    TimeBreakHours = (decimal)Math.Round(TimeBreak.TotalHours, 1);
                }

                factHours = TimeWorkHours - TimeBreakHours;
                TimesheetHours = (decimal)_timesheetDataTable.Rows[i]["TimesheetHours"];

                b = true;
                break;
            }

            var tuple = new Tuple<bool, decimal, decimal>(b, factHours, TimesheetHours);
            return tuple;
        }

        public DataTable GetWorkDaysTable(int UserID, decimal Rate, int Yearint, int Monthint, DateTime dateToday)
        {
            var table = _workDaysDataTable.Clone();

            decimal AllTimesheetHours = 0; // по табелю до сегодняшнего дня
            decimal AllFactHours = 0; // отработано до сегодняшнего дня
            decimal AllPlanHours = 0; // планово до сегодняшнего дня
            decimal AllAbsenceHours = 0;
            decimal OverworkHours = 0; // переработка
            decimal FactOverworkHours = 0; // переработка факт

            for (var i = 0; i < _prodSheduleDataTable.Rows.Count; i++)
            {
                if (Convert.ToInt32(_prodSheduleDataTable.Rows[i]["Day"]) > DateTime.DaysInMonth(Yearint, Monthint))
                    continue;

                var date = new DateTime(Yearint, Monthint, Convert.ToInt32(_prodSheduleDataTable.Rows[i]["Day"]));

                if (date == dateToday.Date) //если это выбранный день
                    break;

                var isAbsence = AbsenceInThatDay(date);
                decimal absenceHour = 0;
                var AbsenceShortName = string.Empty;
                var AbsenceFullName = string.Empty;
                //decimal AbsenceHours = 0;

                for (var j = 0; j < _absDayDataTable.Rows.Count; j++)
                {
                    var absenceTypeId = Convert.ToInt32(_absDayDataTable.Rows[j]["AbsenceTypeID"]);
                    absenceHour = Convert.ToDecimal(_absDayDataTable.Rows[j]["Hours"]);
                    AbsenceShortName += GetAbsenceShortName(absenceTypeId) + " ";
                    AbsenceFullName += string.Format("{0}/({1})", GetAbsenceShortName(absenceTypeId), Convert.ToDecimal(absenceHour.ToString("0.####"))) + " ";

                    if (absenceTypeId != 13 && absenceTypeId != 12)
                    {
                        AllAbsenceHours += absenceHour;
                    }
                }

                var workdayTuple = WorkdayInThatDay(UserID, date);
                var isWorkday = workdayTuple.Item1;
                var factHours = workdayTuple.Item2;
                var timesheetHours = workdayTuple.Item3;

                var prodSheduleHours = Convert.ToInt32(_prodSheduleDataTable.Rows[i]["Hour"]) * Rate;

                if (date == dateToday.Date) //если это выбранный день
                {

                }
                else
                {
                    AllTimesheetHours += timesheetHours;
                    AllFactHours += factHours;
                    AllPlanHours += prodSheduleHours;
                    OverworkHours = AllTimesheetHours - AllPlanHours + AllAbsenceHours;
                    FactOverworkHours = AllFactHours - AllPlanHours + AllAbsenceHours;
                }

                var newRow = table.NewRow();
                newRow["Date"] = date.Date;
                if (Convert.ToInt32(_prodSheduleDataTable.Rows[i]["Hour"]) == 0) // если это выходной
                {
                    newRow["ProdHoursString"] = "Вых";
                    if (timesheetHours > 0)
                        newRow["ProdHoursString"] = "1";
                }
                else
                    newRow["ProdHoursString"] = "1";
                newRow["ProdHours"] = Convert.ToInt32(_prodSheduleDataTable.Rows[i]["Hour"]);

                if (timesheetHours > 0)
                    newRow["IsWorkday"] = true;
                newRow["IsAbsence"] = isAbsence;
                newRow["TimesheetHoursExp"] = AbsenceFullName + " " + timesheetHours;
                newRow["TimesheetHours"] = timesheetHours;
                newRow["FactHours"] = factHours;
                newRow["OverworkHours"] = OverworkHours;
                newRow["FactOverworkHours"] = FactOverworkHours;


                table.Rows.Add(newRow);
            }

            return table;
        }

        public UserTimesheetInfo GetDayInfo(int UserID)
        {
            var userInfo = new UserTimesheetInfo();
            userInfo = monthTimesheet.Find(item => item.UserID == UserID);
            return userInfo;
        }

        public void GetProdShedule(int Yearint, int Monthint)
        {
            using (var da = new SqlDataAdapter(@"SELECT * FROM ProductionShedule WHERE" +
                                               " Year = " + Yearint + " and Month = " + Monthint, ConnectionStrings.LightConnectionString))
            //using (SqlDataAdapter da = new SqlDataAdapter(@"SELECT * FROM ProductionShedule WHERE" +
            //    " Year = " + Yearint, ConnectionStrings.LightConnectionString))
            {
                _prodSheduleDataTable.Clear();
                da.Fill(_prodSheduleDataTable);
            }
        }

        private string GetAbsenceFullName(int id)
        {
            var name = string.Empty;
            var rows = _absTypesTable.Select("AbsenceTypeID = " + id);
            if (rows.Count() > 0)
                name = rows[0]["Description"].ToString();
            return name;
        }

        private string GetAbsenceShortName(int id)
        {
            var name = string.Empty;
            var rows = _absTypesTable.Select("AbsenceTypeID = " + id);
            if (rows.Count() > 0)
                name = rows[0]["ShortName"].ToString();
            return name;
        }

        private int GetHourInProdShedule(DateTime date)
        {
            var hour = -1;
            var rows = _prodSheduleDataTable.Select("Year = " + date.Year + " AND Month=" + date.Month + " AND Day=" + date.Day);
            if (rows.Count() > 0)
                hour = Convert.ToInt32(rows[0]["Hour"]);
            return hour;
        }

        public Tuple<string, int, int, decimal> StaffInfo(int userId)
        {
            var MyQueryText = @"SELECT StaffListID, FactoryID, StaffList.PositionID, Positions.Position, UserID, Rate, Rank FROM StaffList INNER JOIN Positions ON StaffList.PositionID=Positions.PositionID
                WHERE FactoryId=1 AND UserID=" + userId;

            var PositionName = string.Empty;
            var PositionID = 0;
            var Rank = 0;
            decimal Rate = 0;

            using (var da = new SqlDataAdapter(MyQueryText, ConnectionStrings.LightConnectionString))
            {
                using (var dt = new DataTable())
                {
                    if (da.Fill(dt) > 0)
                    {
                        PositionName = dt.Rows[0]["Position"].ToString();
                        PositionID = Convert.ToInt32(dt.Rows[0]["PositionID"]);
                        Rank = Convert.ToInt32(dt.Rows[0]["Rank"]);
                        Rate = Convert.ToDecimal(dt.Rows[0]["Rate"]);
                        StrRate = Rate.ToString();

                    }
                    //if (da.Fill(dt) == 2)
                    //{
                    //    Rate = Convert.ToDecimal(dt.Rows[0]["Rate"]) + Convert.ToDecimal(dt.Rows[1]["Rate"]);
                    //    StrRate = dt.Rows[0]["Rate"].ToString() + "+" + dt.Rows[1]["Rate"].ToString();
                    //}
                }
            }
            var tuple = new Tuple<string, int, int, decimal>(PositionName, PositionID, Rank, Rate);
            return tuple;
        }


        private List<UserTimesheetInfo> monthTimesheet;


        private List<OneMonthTimesheet> yearTimesheets;

        public List<OneMonthTimesheet> YearTimesheets => yearTimesheets;
    }


    public struct OneMonthTimesheet
    {
        public List<UserTimesheetInfo> userTimesheets;

        public int Year;

        public int Month;
        /// <summary>
        /// плановое время
        /// </summary>
        public int AllProdHours;
    }

    public class TimesheetReport
    {
        private string date = "";
        private string totalProdHours = "";
        private readonly string firm = "ООО \"ОМЦ-ПРОФИЛЬ\"";
        private readonly string bossPosition = "Директор ООО \"ОМЦ-ПРОФИЛЬ\"";
        private readonly string bossName = "Ф.А. Авдей";

        private readonly string specialistPosition = "Составил: спец-т по кадрам";
        private readonly string specialistName = "Т.И. Козловская";

        private readonly string assert = "УТВЕРЖДАЮ";
        private readonly string approve = "Согласовано:";

        private string monthHeader = "ГРАФИК/ТАБЕЛЬ РАБОТЫ ЗА ЯНВАРЬ 2021г.:";

        private readonly string user1 = "А.В. Литвиненко";
        private readonly string user2 = "Р.П. Егорченко";
        private readonly string user3 = "Д.М. Яблоков";
        private readonly string user4 = "Ю.К. Янкойть";
        private readonly string user5 = "А.А. Иоскевич";

        private OneMonthTimesheet monthTimesheet;

        public void CreateReport(int Yearint, int Monthint, ResultTimesheet resultTimesheet)
        {
            var hssfworkbook = new HSSFWorkbook();

            #region Create fonts and styles

            var HeaderF = hssfworkbook.CreateFont();
            HeaderF.FontHeightInPoints = 11;
            HeaderF.FontName = "Times New Roman";

            var SimpleF = hssfworkbook.CreateFont();
            SimpleF.FontHeightInPoints = 11;
            SimpleF.FontName = "Times New Roman";

            var SimpleBoldF = hssfworkbook.CreateFont();
            SimpleBoldF.FontHeightInPoints = 11;
            SimpleBoldF.Boldweight = 11 * 256;
            SimpleBoldF.FontName = "Times New Roman";

            var SimpleTopBorderCS = hssfworkbook.CreateCellStyle();
            SimpleTopBorderCS.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
            SimpleTopBorderCS.Alignment = HSSFCellStyle.ALIGN_CENTER;
            SimpleTopBorderCS.BorderBottom = HSSFCellStyle.BORDER_THIN;
            SimpleTopBorderCS.BottomBorderColor = HSSFColor.BLACK.index;
            SimpleTopBorderCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            SimpleTopBorderCS.LeftBorderColor = HSSFColor.BLACK.index;
            SimpleTopBorderCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            SimpleTopBorderCS.RightBorderColor = HSSFColor.BLACK.index;
            SimpleTopBorderCS.BorderTop = HSSFCellStyle.BORDER_MEDIUM;
            SimpleTopBorderCS.TopBorderColor = HSSFColor.BLACK.index;
            SimpleTopBorderCS.WrapText = true;
            SimpleTopBorderCS.SetFont(SimpleF);

            var SimpleBottomBorderCS = hssfworkbook.CreateCellStyle();
            SimpleBottomBorderCS.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.0");
            SimpleBottomBorderCS.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
            SimpleBottomBorderCS.Alignment = HSSFCellStyle.ALIGN_CENTER;
            SimpleBottomBorderCS.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            SimpleBottomBorderCS.BottomBorderColor = HSSFColor.BLACK.index;
            SimpleBottomBorderCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            SimpleBottomBorderCS.LeftBorderColor = HSSFColor.BLACK.index;
            SimpleBottomBorderCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            SimpleBottomBorderCS.RightBorderColor = HSSFColor.BLACK.index;
            SimpleBottomBorderCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            SimpleBottomBorderCS.TopBorderColor = HSSFColor.BLACK.index;
            SimpleBottomBorderCS.WrapText = true;
            SimpleBottomBorderCS.SetFont(SimpleBoldF);

            var MainWithBorderCS = hssfworkbook.CreateCellStyle();
            MainWithBorderCS.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
            MainWithBorderCS.Alignment = HSSFCellStyle.ALIGN_CENTER;
            MainWithBorderCS.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            MainWithBorderCS.BottomBorderColor = HSSFColor.BLACK.index;
            MainWithBorderCS.BorderLeft = HSSFCellStyle.BORDER_MEDIUM;
            MainWithBorderCS.LeftBorderColor = HSSFColor.BLACK.index;
            MainWithBorderCS.BorderRight = HSSFCellStyle.BORDER_MEDIUM;
            MainWithBorderCS.RightBorderColor = HSSFColor.BLACK.index;
            MainWithBorderCS.BorderTop = HSSFCellStyle.BORDER_MEDIUM;
            MainWithBorderCS.TopBorderColor = HSSFColor.BLACK.index;
            MainWithBorderCS.SetFont(SimpleF);

            var MainWithoutBorderCS = hssfworkbook.CreateCellStyle();
            MainWithoutBorderCS.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
            MainWithoutBorderCS.Alignment = HSSFCellStyle.ALIGN_CENTER;
            MainWithoutBorderCS.SetFont(SimpleF);

            var NameCS = hssfworkbook.CreateCellStyle();
            NameCS.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
            NameCS.Alignment = HSSFCellStyle.ALIGN_LEFT;
            NameCS.BorderBottom = HSSFCellStyle.BORDER_THIN;
            NameCS.BottomBorderColor = HSSFColor.BLACK.index;
            NameCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            NameCS.LeftBorderColor = HSSFColor.BLACK.index;
            NameCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            NameCS.RightBorderColor = HSSFColor.BLACK.index;
            NameCS.BorderTop = HSSFCellStyle.BORDER_MEDIUM;
            NameCS.TopBorderColor = HSSFColor.BLACK.index;
            NameCS.SetFont(SimpleF);

            var SerialNumberCS = hssfworkbook.CreateCellStyle();
            SerialNumberCS.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
            SerialNumberCS.Alignment = HSSFCellStyle.ALIGN_CENTER;
            SerialNumberCS.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            SerialNumberCS.BottomBorderColor = HSSFColor.BLACK.index;
            SerialNumberCS.BorderLeft = HSSFCellStyle.BORDER_MEDIUM;
            SerialNumberCS.LeftBorderColor = HSSFColor.BLACK.index;
            SerialNumberCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            SerialNumberCS.RightBorderColor = HSSFColor.BLACK.index;
            SerialNumberCS.BorderTop = HSSFCellStyle.BORDER_MEDIUM;
            SerialNumberCS.TopBorderColor = HSSFColor.BLACK.index;
            SerialNumberCS.SetFont(SimpleF);

            var PositionCS = hssfworkbook.CreateCellStyle();
            PositionCS.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
            PositionCS.Alignment = HSSFCellStyle.ALIGN_LEFT;
            PositionCS.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            PositionCS.BottomBorderColor = HSSFColor.BLACK.index;
            PositionCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            PositionCS.LeftBorderColor = HSSFColor.BLACK.index;
            PositionCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            PositionCS.RightBorderColor = HSSFColor.BLACK.index;
            PositionCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            PositionCS.TopBorderColor = HSSFColor.BLACK.index;
            PositionCS.SetFont(SimpleF);

            var SimpleCS = hssfworkbook.CreateCellStyle();
            SimpleCS.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
            SimpleCS.Alignment = HSSFCellStyle.ALIGN_CENTER;
            SimpleCS.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.0");
            SimpleCS.BorderBottom = HSSFCellStyle.BORDER_THIN;
            SimpleCS.BottomBorderColor = HSSFColor.BLACK.index;
            SimpleCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            SimpleCS.LeftBorderColor = HSSFColor.BLACK.index;
            SimpleCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            SimpleCS.RightBorderColor = HSSFColor.BLACK.index;
            SimpleCS.BorderTop = HSSFCellStyle.BORDER_MEDIUM;
            SimpleCS.TopBorderColor = HSSFColor.BLACK.index;
            SimpleCS.SetFont(SimpleF);

            var SimpleDecCS = hssfworkbook.CreateCellStyle();
            SimpleDecCS.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
            SimpleDecCS.Alignment = HSSFCellStyle.ALIGN_CENTER;
            SimpleDecCS.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.0");
            SimpleDecCS.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            SimpleDecCS.BottomBorderColor = HSSFColor.BLACK.index;
            SimpleDecCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            SimpleDecCS.LeftBorderColor = HSSFColor.BLACK.index;
            SimpleDecCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            SimpleDecCS.RightBorderColor = HSSFColor.BLACK.index;
            SimpleDecCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            SimpleDecCS.TopBorderColor = HSSFColor.BLACK.index;
            SimpleDecCS.SetFont(SimpleF);

            var RateCS = hssfworkbook.CreateCellStyle();
            RateCS.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
            RateCS.Alignment = HSSFCellStyle.ALIGN_CENTER;
            RateCS.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.000");
            RateCS.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            RateCS.BottomBorderColor = HSSFColor.BLACK.index;
            RateCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            RateCS.LeftBorderColor = HSSFColor.BLACK.index;
            RateCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            RateCS.RightBorderColor = HSSFColor.BLACK.index;
            RateCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            RateCS.TopBorderColor = HSSFColor.BLACK.index;
            RateCS.SetFont(SimpleF);

            var TimesheetHoursCS = hssfworkbook.CreateCellStyle();
            TimesheetHoursCS.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
            TimesheetHoursCS.Alignment = HSSFCellStyle.ALIGN_CENTER;
            TimesheetHoursCS.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.000");
            TimesheetHoursCS.BorderBottom = HSSFCellStyle.BORDER_THIN;
            TimesheetHoursCS.BottomBorderColor = HSSFColor.BLACK.index;
            TimesheetHoursCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            TimesheetHoursCS.LeftBorderColor = HSSFColor.BLACK.index;
            TimesheetHoursCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            TimesheetHoursCS.RightBorderColor = HSSFColor.BLACK.index;
            TimesheetHoursCS.BorderTop = HSSFCellStyle.BORDER_MEDIUM;
            TimesheetHoursCS.TopBorderColor = HSSFColor.BLACK.index;
            TimesheetHoursCS.SetFont(SimpleF);

            var OutputTopBorderCS = hssfworkbook.CreateCellStyle();
            OutputTopBorderCS.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
            OutputTopBorderCS.Alignment = HSSFCellStyle.ALIGN_CENTER;
            OutputTopBorderCS.BorderBottom = HSSFCellStyle.BORDER_THIN;
            OutputTopBorderCS.BottomBorderColor = HSSFColor.BLACK.index;
            OutputTopBorderCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            OutputTopBorderCS.LeftBorderColor = HSSFColor.BLACK.index;
            OutputTopBorderCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            OutputTopBorderCS.RightBorderColor = HSSFColor.BLACK.index;
            OutputTopBorderCS.BorderTop = HSSFCellStyle.BORDER_MEDIUM;
            OutputTopBorderCS.TopBorderColor = HSSFColor.BLACK.index;
            OutputTopBorderCS.FillForegroundColor = HSSFColor.YELLOW.index;
            OutputTopBorderCS.FillPattern = HSSFCellStyle.SOLID_FOREGROUND;
            OutputTopBorderCS.FillBackgroundColor = HSSFColor.YELLOW.index;
            OutputTopBorderCS.SetFont(SimpleF);

            var OutputBottomBorderCS = hssfworkbook.CreateCellStyle();
            OutputBottomBorderCS.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.0");
            OutputBottomBorderCS.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
            OutputBottomBorderCS.Alignment = HSSFCellStyle.ALIGN_CENTER;
            OutputBottomBorderCS.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            OutputBottomBorderCS.BottomBorderColor = HSSFColor.BLACK.index;
            OutputBottomBorderCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            OutputBottomBorderCS.LeftBorderColor = HSSFColor.BLACK.index;
            OutputBottomBorderCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            OutputBottomBorderCS.RightBorderColor = HSSFColor.BLACK.index;
            OutputBottomBorderCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            OutputBottomBorderCS.TopBorderColor = HSSFColor.BLACK.index;
            OutputBottomBorderCS.FillForegroundColor = HSSFColor.YELLOW.index;
            OutputBottomBorderCS.FillPattern = HSSFCellStyle.SOLID_FOREGROUND;
            OutputBottomBorderCS.FillBackgroundColor = HSSFColor.YELLOW.index;
            OutputBottomBorderCS.WrapText = true;
            OutputBottomBorderCS.SetFont(SimpleBoldF);

            var AbsenceTopBorderCS = hssfworkbook.CreateCellStyle();
            AbsenceTopBorderCS.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
            AbsenceTopBorderCS.Alignment = HSSFCellStyle.ALIGN_CENTER;
            AbsenceTopBorderCS.BorderBottom = HSSFCellStyle.BORDER_THIN;
            AbsenceTopBorderCS.BottomBorderColor = HSSFColor.BLACK.index;
            AbsenceTopBorderCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            AbsenceTopBorderCS.LeftBorderColor = HSSFColor.BLACK.index;
            AbsenceTopBorderCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            AbsenceTopBorderCS.RightBorderColor = HSSFColor.BLACK.index;
            AbsenceTopBorderCS.BorderTop = HSSFCellStyle.BORDER_MEDIUM;
            AbsenceTopBorderCS.TopBorderColor = HSSFColor.BLACK.index;
            AbsenceTopBorderCS.FillForegroundColor = HSSFColor.RED.index;
            AbsenceTopBorderCS.FillPattern = HSSFCellStyle.SOLID_FOREGROUND;
            AbsenceTopBorderCS.FillBackgroundColor = HSSFColor.YELLOW.index;
            AbsenceTopBorderCS.SetFont(SimpleF);

            var AbsenceBottomBorderCS = hssfworkbook.CreateCellStyle();
            AbsenceBottomBorderCS.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.0");
            AbsenceBottomBorderCS.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
            AbsenceBottomBorderCS.Alignment = HSSFCellStyle.ALIGN_CENTER;
            AbsenceBottomBorderCS.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            AbsenceBottomBorderCS.BottomBorderColor = HSSFColor.BLACK.index;
            AbsenceBottomBorderCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            AbsenceBottomBorderCS.LeftBorderColor = HSSFColor.BLACK.index;
            AbsenceBottomBorderCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            AbsenceBottomBorderCS.RightBorderColor = HSSFColor.BLACK.index;
            AbsenceBottomBorderCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            AbsenceBottomBorderCS.TopBorderColor = HSSFColor.BLACK.index;
            AbsenceBottomBorderCS.FillForegroundColor = HSSFColor.RED.index;
            AbsenceBottomBorderCS.FillPattern = HSSFCellStyle.SOLID_FOREGROUND;
            AbsenceBottomBorderCS.FillBackgroundColor = HSSFColor.YELLOW.index;
            AbsenceBottomBorderCS.WrapText = true;
            AbsenceBottomBorderCS.SetFont(SimpleBoldF);

            var OutputHeaderTopBorderCS = hssfworkbook.CreateCellStyle();
            OutputHeaderTopBorderCS.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
            OutputHeaderTopBorderCS.Alignment = HSSFCellStyle.ALIGN_CENTER;
            OutputHeaderTopBorderCS.BorderBottom = HSSFCellStyle.BORDER_THIN;
            OutputHeaderTopBorderCS.BottomBorderColor = HSSFColor.BLACK.index;
            OutputHeaderTopBorderCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            OutputHeaderTopBorderCS.LeftBorderColor = HSSFColor.BLACK.index;
            OutputHeaderTopBorderCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            OutputHeaderTopBorderCS.RightBorderColor = HSSFColor.BLACK.index;
            OutputHeaderTopBorderCS.BorderTop = HSSFCellStyle.BORDER_MEDIUM;
            OutputHeaderTopBorderCS.TopBorderColor = HSSFColor.BLACK.index;
            OutputHeaderTopBorderCS.FillForegroundColor = HSSFColor.YELLOW.index;
            OutputHeaderTopBorderCS.FillPattern = HSSFCellStyle.SOLID_FOREGROUND;
            OutputHeaderTopBorderCS.FillBackgroundColor = HSSFColor.YELLOW.index;
            OutputHeaderTopBorderCS.SetFont(SimpleF);

            var OutputHeaderBottomBorderCS = hssfworkbook.CreateCellStyle();
            OutputHeaderBottomBorderCS.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
            OutputHeaderBottomBorderCS.Alignment = HSSFCellStyle.ALIGN_CENTER;
            OutputHeaderBottomBorderCS.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            OutputHeaderBottomBorderCS.BottomBorderColor = HSSFColor.BLACK.index;
            OutputHeaderBottomBorderCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            OutputHeaderBottomBorderCS.LeftBorderColor = HSSFColor.BLACK.index;
            OutputHeaderBottomBorderCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            OutputHeaderBottomBorderCS.RightBorderColor = HSSFColor.BLACK.index;
            OutputHeaderBottomBorderCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            OutputHeaderBottomBorderCS.TopBorderColor = HSSFColor.BLACK.index;
            OutputHeaderBottomBorderCS.FillForegroundColor = HSSFColor.YELLOW.index;
            OutputHeaderBottomBorderCS.FillPattern = HSSFCellStyle.SOLID_FOREGROUND;
            OutputHeaderBottomBorderCS.FillBackgroundColor = HSSFColor.YELLOW.index;
            OutputHeaderBottomBorderCS.WrapText = true;
            OutputHeaderBottomBorderCS.SetFont(SimpleBoldF);

            var HeaderTopBorderCS = hssfworkbook.CreateCellStyle();
            HeaderTopBorderCS.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
            HeaderTopBorderCS.Alignment = HSSFCellStyle.ALIGN_CENTER;
            HeaderTopBorderCS.BorderBottom = HSSFCellStyle.BORDER_THIN;
            HeaderTopBorderCS.BottomBorderColor = HSSFColor.BLACK.index;
            HeaderTopBorderCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            HeaderTopBorderCS.LeftBorderColor = HSSFColor.BLACK.index;
            HeaderTopBorderCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            HeaderTopBorderCS.RightBorderColor = HSSFColor.BLACK.index;
            HeaderTopBorderCS.BorderTop = HSSFCellStyle.BORDER_MEDIUM;
            HeaderTopBorderCS.TopBorderColor = HSSFColor.BLACK.index;
            HeaderTopBorderCS.SetFont(SimpleF);

            var HeaderBottomBorderCS = hssfworkbook.CreateCellStyle();
            HeaderBottomBorderCS.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
            HeaderBottomBorderCS.Alignment = HSSFCellStyle.ALIGN_CENTER;
            HeaderBottomBorderCS.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            HeaderBottomBorderCS.BottomBorderColor = HSSFColor.BLACK.index;
            HeaderBottomBorderCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            HeaderBottomBorderCS.LeftBorderColor = HSSFColor.BLACK.index;
            HeaderBottomBorderCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            HeaderBottomBorderCS.RightBorderColor = HSSFColor.BLACK.index;
            HeaderBottomBorderCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            HeaderBottomBorderCS.TopBorderColor = HSSFColor.BLACK.index;
            HeaderBottomBorderCS.WrapText = true;
            HeaderBottomBorderCS.SetFont(SimpleBoldF);

            var ReportCS = hssfworkbook.CreateCellStyle();
            ReportCS.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            ReportCS.BottomBorderColor = HSSFColor.BLACK.index;
            ReportCS.SetFont(HeaderF);

            var HeaderWithoutBorderCS = hssfworkbook.CreateCellStyle();
            HeaderWithoutBorderCS.SetFont(HeaderF);

            var SimpleHeaderCS = hssfworkbook.CreateCellStyle();
            SimpleHeaderCS.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
            SimpleHeaderCS.Alignment = HSSFCellStyle.ALIGN_CENTER;
            SimpleHeaderCS.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            SimpleHeaderCS.BottomBorderColor = HSSFColor.BLACK.index;
            SimpleHeaderCS.BorderLeft = HSSFCellStyle.BORDER_MEDIUM;
            SimpleHeaderCS.LeftBorderColor = HSSFColor.BLACK.index;
            SimpleHeaderCS.BorderRight = HSSFCellStyle.BORDER_MEDIUM;
            SimpleHeaderCS.RightBorderColor = HSSFColor.BLACK.index;
            SimpleHeaderCS.BorderTop = HSSFCellStyle.BORDER_MEDIUM;
            SimpleHeaderCS.TopBorderColor = HSSFColor.BLACK.index;
            SimpleHeaderCS.WrapText = true;
            SimpleHeaderCS.SetFont(HeaderF);

            #endregion

            for (var j = 0; j < resultTimesheet.YearTimesheets.Count; j++)
            {

                monthTimesheet = resultTimesheet.YearTimesheets[j];
                if (monthTimesheet.userTimesheets.Count == 0)
                    continue;

                date = DateTime.Now.ToShortDateString();
                totalProdHours = monthTimesheet.AllProdHours.ToString();
                monthHeader = "ГРАФИК/ТАБЕЛЬ РАБОТЫ ЗА " + new DateTime(monthTimesheet.Year, monthTimesheet.Month, 1).ToString("MMMM") + " " + monthTimesheet.Year + "г:";

                var sYear = monthTimesheet.Year.ToString();

                var sMonth = monthTimesheet.Month.ToString().PadLeft(2, '0');


                var sheetName = sMonth + "-" + sYear;
                var sheet1 = hssfworkbook.CreateSheet(sheetName);
                sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;
                sheet1.SetZoom(85, 100);
                sheet1.SetMargin(HSSFSheet.LeftMargin, .12);
                sheet1.SetMargin(HSSFSheet.RightMargin, .07);
                sheet1.SetMargin(HSSFSheet.TopMargin, .20);
                sheet1.SetMargin(HSSFSheet.BottomMargin, .20);

                var DisplayIndex = 1;
                var RowIndex = 0;

                HSSFCell HeaderCell;
                HSSFCell Cell1;

                HeaderCell = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex);
                HeaderCell.SetCellValue(date);
                HeaderCell.CellStyle = MainWithBorderCS;

                HeaderCell = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex);
                HeaderCell.SetCellValue(totalProdHours);
                HeaderCell.CellStyle = MainWithBorderCS;

                HeaderCell = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex + 20);
                HeaderCell.SetCellValue("                   " + assert);
                HeaderCell.CellStyle = MainWithoutBorderCS;

                HeaderCell = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex + 20);
                HeaderCell.SetCellValue("_________________  " + bossName);
                HeaderCell.CellStyle = MainWithoutBorderCS;

                HeaderCell = sheet1.CreateRow(RowIndex + 2).CreateCell(DisplayIndex + 20);
                HeaderCell.SetCellValue(bossPosition);
                HeaderCell.CellStyle = MainWithoutBorderCS;

                HeaderCell = sheet1.CreateRow(RowIndex + 2).CreateCell(DisplayIndex + 2);
                HeaderCell.SetCellValue(firm);
                HeaderCell.CellStyle = MainWithoutBorderCS;

                HeaderCell = sheet1.CreateRow(RowIndex + 4).CreateCell(DisplayIndex + 20);
                HeaderCell.SetCellValue("_________________  " + specialistName);
                HeaderCell.CellStyle = MainWithoutBorderCS;

                HeaderCell = sheet1.CreateRow(RowIndex + 5).CreateCell(DisplayIndex + 20);
                HeaderCell.SetCellValue(specialistPosition);
                HeaderCell.CellStyle = MainWithoutBorderCS;

                HeaderCell = sheet1.CreateRow(RowIndex + 6).CreateCell(DisplayIndex + 10);
                HeaderCell.SetCellValue(monthHeader);
                HeaderCell.CellStyle = MainWithBorderCS;

                for (var i = DisplayIndex + 11; i <= DisplayIndex + 20; i++)
                {
                    HeaderCell = sheet1.CreateRow(RowIndex + 6).CreateCell(i);
                    HeaderCell.SetCellValue("");
                    HeaderCell.CellStyle = MainWithBorderCS;
                }
                sheet1.AddMergedRegion(new Region(RowIndex + 6, DisplayIndex + 10, RowIndex + 6, DisplayIndex + 20));

                DisplayIndex = 0;
                RowIndex = 7;

                HeaderCell = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex);
                HeaderCell.SetCellValue("");
                HeaderCell.CellStyle = SimpleHeaderCS;
                HeaderCell = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex);
                HeaderCell.SetCellValue("№ п/п");
                HeaderCell.CellStyle = SimpleHeaderCS;
                sheet1.AddMergedRegion(new Region(RowIndex, DisplayIndex, RowIndex + 1, DisplayIndex));
                DisplayIndex++;

                HeaderCell = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex);
                HeaderCell.SetCellValue("");
                HeaderCell.CellStyle = SimpleHeaderCS;
                HeaderCell = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex);
                HeaderCell.SetCellValue("Ф.И.О.");
                HeaderCell.CellStyle = SimpleHeaderCS;
                sheet1.AddMergedRegion(new Region(RowIndex, DisplayIndex, RowIndex + 1, DisplayIndex));
                DisplayIndex++;

                HeaderCell = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex);
                HeaderCell.SetCellValue("");
                HeaderCell.CellStyle = SimpleHeaderCS;
                HeaderCell = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex);
                HeaderCell.SetCellValue("Разряд Ставка");
                HeaderCell.CellStyle = SimpleHeaderCS;
                sheet1.AddMergedRegion(new Region(RowIndex, DisplayIndex, RowIndex + 1, DisplayIndex));
                DisplayIndex++;

                for (var i = 0; i < monthTimesheet.userTimesheets[0].WorkDaysDT.Rows.Count; i++)
                {
                    var prodHours = Convert.ToInt32(monthTimesheet.userTimesheets[0].WorkDaysDT.Rows[i]["ProdHours"]);
                    if (prodHours == 0)
                    {
                        HeaderCell = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex);
                        HeaderCell.SetCellValue((i + 1));
                        HeaderCell.CellStyle = OutputHeaderTopBorderCS;

                        HeaderCell = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex);
                        HeaderCell.SetCellValue(Convert.ToInt32(monthTimesheet.userTimesheets[0].WorkDaysDT.Rows[i]["ProdHours"]));
                        HeaderCell.CellStyle = OutputHeaderBottomBorderCS;
                    }
                    else
                    {
                        HeaderCell = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex);
                        HeaderCell.SetCellValue((i + 1));
                        HeaderCell.CellStyle = HeaderTopBorderCS;

                        HeaderCell = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex);
                        HeaderCell.SetCellValue(Convert.ToInt32(monthTimesheet.userTimesheets[0].WorkDaysDT.Rows[i]["ProdHours"]));
                        HeaderCell.CellStyle = HeaderBottomBorderCS;
                    }


                    DisplayIndex++;
                }

                HeaderCell = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex);
                HeaderCell.SetCellValue("");
                HeaderCell.CellStyle = SimpleHeaderCS;
                HeaderCell = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex);
                HeaderCell.SetCellValue("Отраб Дни");
                HeaderCell.CellStyle = SimpleHeaderCS;
                sheet1.AddMergedRegion(new Region(RowIndex, DisplayIndex, RowIndex + 1, DisplayIndex));
                DisplayIndex++;

                HeaderCell = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex);
                HeaderCell.SetCellValue("");
                HeaderCell.CellStyle = SimpleHeaderCS;
                HeaderCell = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex);
                HeaderCell.SetCellValue("Отраб Табель Часы");
                HeaderCell.CellStyle = SimpleHeaderCS;
                sheet1.AddMergedRegion(new Region(RowIndex, DisplayIndex, RowIndex + 1, DisplayIndex));
                DisplayIndex++;

                HeaderCell = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex);
                HeaderCell.SetCellValue("");
                HeaderCell.CellStyle = SimpleHeaderCS;
                HeaderCell = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex);
                HeaderCell.SetCellValue("Перераб Табель");
                HeaderCell.CellStyle = SimpleHeaderCS;
                sheet1.AddMergedRegion(new Region(RowIndex, DisplayIndex, RowIndex + 1, DisplayIndex));
                DisplayIndex++;

                HeaderCell = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex);
                HeaderCell.SetCellValue("");
                HeaderCell.CellStyle = SimpleHeaderCS;
                HeaderCell = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex);
                HeaderCell.SetCellValue("Отраб Факт Часы");
                HeaderCell.CellStyle = SimpleHeaderCS;
                sheet1.AddMergedRegion(new Region(RowIndex, DisplayIndex, RowIndex + 1, DisplayIndex));
                DisplayIndex++;

                HeaderCell = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex);
                HeaderCell.SetCellValue("");
                HeaderCell.CellStyle = SimpleHeaderCS;
                HeaderCell = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex);
                HeaderCell.SetCellValue("Перераб Факт");
                HeaderCell.CellStyle = SimpleHeaderCS;
                sheet1.AddMergedRegion(new Region(RowIndex, DisplayIndex, RowIndex + 1, DisplayIndex));
                DisplayIndex++;

                RowIndex++;
                RowIndex++;
                for (var i = 0; i < monthTimesheet.userTimesheets.Count; i++)
                {
                    DisplayIndex = 0;

                    Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex);
                    Cell1.SetCellValue((i + 1));
                    Cell1.CellStyle = SerialNumberCS;

                    Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex);
                    Cell1.SetCellValue("");
                    Cell1.CellStyle = SerialNumberCS;

                    sheet1.AddMergedRegion(new Region(RowIndex, DisplayIndex, RowIndex + 1, DisplayIndex));

                    DisplayIndex++;

                    Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex);
                    Cell1.SetCellValue(monthTimesheet.userTimesheets[i].Name);
                    Cell1.CellStyle = NameCS;
                    Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex);
                    Cell1.SetCellValue(monthTimesheet.userTimesheets[i].PositionName);
                    Cell1.CellStyle = PositionCS;

                    DisplayIndex++;

                    Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex);
                    Cell1.SetCellValue(monthTimesheet.userTimesheets[i].Rank);
                    Cell1.CellStyle = SimpleTopBorderCS;
                    Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex);
                    Cell1.SetCellValue(Convert.ToDouble(monthTimesheet.userTimesheets[i].Rate));
                    Cell1.CellStyle = RateCS;

                    DisplayIndex++;
                    var b = false;

                    for (var x = 0; x < monthTimesheet.userTimesheets[i].WorkDaysDT.Rows.Count; x++)
                    {
                        var isAbsence = Convert.ToBoolean(monthTimesheet.userTimesheets[i].WorkDaysDT.Rows[x]["IsAbsence"]);
                        var prodHours = Convert.ToInt32(monthTimesheet.userTimesheets[i].WorkDaysDT.Rows[x]["ProdHours"]);
                        if (prodHours == 0)
                        {
                            if (isAbsence)
                            {
                                Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex);
                                b = int.TryParse(monthTimesheet.userTimesheets[i].WorkDaysDT.Rows[x]["ProdHoursString"].ToString(), out var ProdHours);
                                if (b)
                                    Cell1.SetCellValue(Convert.ToInt32(ProdHours));
                                else
                                    Cell1.SetCellValue(monthTimesheet.userTimesheets[i].WorkDaysDT.Rows[x]["ProdHoursString"].ToString());
                                Cell1.CellStyle = AbsenceTopBorderCS;


                                Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex);
                                b = decimal.TryParse(monthTimesheet.userTimesheets[i].WorkDaysDT.Rows[x]["TimesheetHoursExp"].ToString(), out var TimesheetHoursExp);
                                if (b)
                                    Cell1.SetCellValue(Convert.ToDouble(TimesheetHoursExp));
                                else
                                    Cell1.SetCellValue(monthTimesheet.userTimesheets[i].WorkDaysDT.Rows[x]["TimesheetHoursExp"].ToString());

                                Cell1.CellStyle = AbsenceBottomBorderCS;
                            }
                            else
                            {
                                Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex);
                                b = int.TryParse(monthTimesheet.userTimesheets[i].WorkDaysDT.Rows[x]["ProdHoursString"].ToString(), out var ProdHours);
                                if (b)
                                    Cell1.SetCellValue(Convert.ToInt32(ProdHours));
                                else
                                    Cell1.SetCellValue(monthTimesheet.userTimesheets[i].WorkDaysDT.Rows[x]["ProdHoursString"].ToString());
                                Cell1.CellStyle = OutputTopBorderCS;


                                Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex);
                                b = decimal.TryParse(monthTimesheet.userTimesheets[i].WorkDaysDT.Rows[x]["TimesheetHoursExp"].ToString(), out var TimesheetHoursExp);
                                if (b)
                                    Cell1.SetCellValue(Convert.ToDouble(TimesheetHoursExp));
                                else
                                    Cell1.SetCellValue(monthTimesheet.userTimesheets[i].WorkDaysDT.Rows[x]["TimesheetHoursExp"].ToString());

                                Cell1.CellStyle = OutputBottomBorderCS;
                            }
                        }
                        else
                        {
                            if (isAbsence)
                            {
                                Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex);
                                b = int.TryParse(monthTimesheet.userTimesheets[i].WorkDaysDT.Rows[x]["ProdHoursString"].ToString(), out var ProdHours);
                                if (b)
                                    Cell1.SetCellValue(Convert.ToInt32(ProdHours));
                                else
                                    Cell1.SetCellValue(monthTimesheet.userTimesheets[i].WorkDaysDT.Rows[x]["ProdHoursString"].ToString());
                                Cell1.CellStyle = AbsenceTopBorderCS;


                                Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex);
                                b = decimal.TryParse(monthTimesheet.userTimesheets[i].WorkDaysDT.Rows[x]["TimesheetHoursExp"].ToString(), out var TimesheetHoursExp);
                                if (b)
                                    Cell1.SetCellValue(Convert.ToDouble(TimesheetHoursExp));
                                else
                                    Cell1.SetCellValue(monthTimesheet.userTimesheets[i].WorkDaysDT.Rows[x]["TimesheetHoursExp"].ToString());

                                Cell1.CellStyle = AbsenceBottomBorderCS;
                            }
                            else
                            {
                                Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex);
                                b = int.TryParse(monthTimesheet.userTimesheets[i].WorkDaysDT.Rows[x]["ProdHoursString"].ToString(), out var ProdHours);
                                if (b)
                                    Cell1.SetCellValue(Convert.ToInt32(ProdHours));
                                else
                                    Cell1.SetCellValue(monthTimesheet.userTimesheets[i].WorkDaysDT.Rows[x]["ProdHoursString"].ToString());
                                Cell1.CellStyle = SimpleTopBorderCS;


                                Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex);
                                b = decimal.TryParse(monthTimesheet.userTimesheets[i].WorkDaysDT.Rows[x]["TimesheetHoursExp"].ToString(), out var TimesheetHoursExp);
                                if (b)
                                    Cell1.SetCellValue(Convert.ToDouble(TimesheetHoursExp));
                                else
                                    Cell1.SetCellValue(monthTimesheet.userTimesheets[i].WorkDaysDT.Rows[x]["TimesheetHoursExp"].ToString());

                                Cell1.CellStyle = SimpleBottomBorderCS;
                            }
                        }

                        DisplayIndex++;
                    }

                    Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex);

                    b = int.TryParse(monthTimesheet.userTimesheets[i].AllTimesheetDays.ToString(), out var AllTimesheetDays);
                    if (b)
                        Cell1.SetCellValue(Convert.ToInt32(AllTimesheetDays));
                    else
                        Cell1.SetCellValue(monthTimesheet.userTimesheets[i].AllTimesheetDays.ToString());

                    Cell1.CellStyle = SimpleCS;
                    sheet1.AddMergedRegion(new Region(RowIndex, DisplayIndex, RowIndex + 1, DisplayIndex));

                    Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex);
                    Cell1.CellStyle = SimpleDecCS;

                    DisplayIndex++;

                    Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex);

                    b = decimal.TryParse(monthTimesheet.userTimesheets[i].AllTimesheetHours.ToString(), out var AllTimesheetHours);
                    if (b)
                        Cell1.SetCellValue(Convert.ToDouble(AllTimesheetHours));
                    else
                        Cell1.SetCellValue(monthTimesheet.userTimesheets[i].AllTimesheetHours.ToString());

                    Cell1.CellStyle = TimesheetHoursCS;

                    Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex);
                    Cell1.CellStyle = SimpleDecCS;
                    sheet1.AddMergedRegion(new Region(RowIndex, DisplayIndex, RowIndex + 1, DisplayIndex));

                    DisplayIndex++;

                    Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex);

                    b = decimal.TryParse(monthTimesheet.userTimesheets[i].AllOverworkHours.ToString(), out var AllOverworkHours);
                    if (b)
                        Cell1.SetCellValue(Convert.ToDouble(AllOverworkHours));
                    else
                        Cell1.SetCellValue(monthTimesheet.userTimesheets[i].AllOverworkHours.ToString());

                    Cell1.CellStyle = TimesheetHoursCS;

                    Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex);
                    Cell1.CellStyle = SimpleDecCS;
                    sheet1.AddMergedRegion(new Region(RowIndex, DisplayIndex, RowIndex + 1, DisplayIndex));

                    DisplayIndex++;

                    Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex);

                    b = decimal.TryParse(monthTimesheet.userTimesheets[i].AllFactHours.ToString(), out var AllFactHours);
                    if (b)
                        Cell1.SetCellValue(Convert.ToDouble(AllFactHours));
                    else
                        Cell1.SetCellValue(monthTimesheet.userTimesheets[i].AllFactHours.ToString());

                    Cell1.CellStyle = TimesheetHoursCS;

                    Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex);
                    Cell1.CellStyle = SimpleDecCS;
                    sheet1.AddMergedRegion(new Region(RowIndex, DisplayIndex, RowIndex + 1, DisplayIndex));

                    DisplayIndex++;

                    Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex);

                    b = decimal.TryParse(monthTimesheet.userTimesheets[i].AllFactOverworkHours.ToString(), out var AllFactOverworkHours);
                    if (b)
                        Cell1.SetCellValue(Convert.ToDouble(AllFactOverworkHours));
                    else
                        Cell1.SetCellValue(monthTimesheet.userTimesheets[i].AllFactOverworkHours.ToString());

                    Cell1.CellStyle = TimesheetHoursCS;

                    Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex);
                    Cell1.CellStyle = SimpleDecCS;
                    sheet1.AddMergedRegion(new Region(RowIndex, DisplayIndex, RowIndex + 1, DisplayIndex));

                    DisplayIndex++;

                    RowIndex++;
                    RowIndex++;


                }

                DisplayIndex = 0;

                HeaderCell = sheet1.CreateRow(RowIndex + 2).CreateCell(DisplayIndex + 1);
                HeaderCell.SetCellValue(approve);
                HeaderCell.CellStyle = MainWithoutBorderCS;

                HeaderCell = sheet1.CreateRow(RowIndex + 2).CreateCell(DisplayIndex + 3);
                HeaderCell.SetCellValue(user1);
                HeaderCell.CellStyle = MainWithoutBorderCS;

                HeaderCell = sheet1.CreateRow(RowIndex + 4).CreateCell(DisplayIndex + 3);
                HeaderCell.SetCellValue(user2);
                HeaderCell.CellStyle = MainWithoutBorderCS;

                HeaderCell = sheet1.CreateRow(RowIndex + 2).CreateCell(DisplayIndex + 7);
                HeaderCell.SetCellValue(user3);
                HeaderCell.CellStyle = MainWithoutBorderCS;

                HeaderCell = sheet1.CreateRow(RowIndex + 4).CreateCell(DisplayIndex + 7);
                HeaderCell.SetCellValue(user4);
                HeaderCell.CellStyle = MainWithoutBorderCS;

                HeaderCell = sheet1.CreateRow(RowIndex + 2).CreateCell(DisplayIndex + 11);
                HeaderCell.SetCellValue(user5);
                HeaderCell.CellStyle = MainWithoutBorderCS;

                sheet1.SetColumnWidth(DisplayIndex++, 7 * 256);
                sheet1.SetColumnWidth(DisplayIndex++, 22 * 256);
                sheet1.SetColumnWidth(DisplayIndex++, 10 * 256);
                for (var i = 0; i < monthTimesheet.userTimesheets[0].WorkDaysDT.Rows.Count; i++)
                {
                    sheet1.SetColumnWidth(DisplayIndex++, 8 * 256);

                }
                sheet1.SetColumnWidth(DisplayIndex++, 8 * 256);// Факт. отраб.
                sheet1.SetColumnWidth(DisplayIndex++, 8 * 256);// Часы
                sheet1.SetColumnWidth(DisplayIndex++, 9 * 256);


                sheet1.GetRow(8).Height = 3 * 256;
                sheet1.CreateFreezePane(3, 9, 3, 9);
            }

            var FileName = "Профиль-" + monthTimesheet.Year;
            var tempFolder = Environment.GetEnvironmentVariable("TEMP");
            var file = new FileInfo(tempFolder + @"\" + FileName + ".xls");

            var y = 1;
            while (file.Exists)
            {
                file = new FileInfo(tempFolder + @"\" + FileName + "(" + y++ + ").xls");
            }

            var NewFile = new FileStream(file.FullName, FileMode.Create);
            hssfworkbook.Write(NewFile);
            NewFile.Close();
            Process.Start(file.FullName);
        }
    }


    public struct UserTimesheetInfo
    {
        /// <summary>
        /// id работника
        /// </summary>
        public int UserID;

        /// <summary>
        /// ФИО
        /// </summary>
        public string Name;

        /// <summary>
        /// таблица рабочих дней
        /// </summary>
        public DataTable WorkDaysDT;

        /// <summary>
        /// Разряд
        /// </summary>
        public int Rank;

        /// <summary>
        /// Ставка
        /// </summary>
        public decimal Rate;

        /// <summary>
        /// id должности
        /// </summary>
        public int PositionID;

        /// <summary>
        /// должность
        /// </summary>
        public string PositionName;

        /// <summary>
        /// отработано дней по табелю
        /// </summary>
        public decimal AllTimesheetDays;

        /// <summary>
        /// отработано часов по табелю
        /// </summary>
        public decimal AllTimesheetHours;
        /// <summary>
        /// отработано часов фактически
        /// </summary>
        public decimal AllFactHours;
        /// <summary>
        /// переработка
        /// </summary>
        public decimal AllOverworkHours;
        /// <summary>
        /// переработка
        /// </summary>
        public decimal AllFactOverworkHours;
    }

    public struct TimesheetInfo
    {
        public DateTime Date;
        /// <summary>
        /// прогулы и отгулы
        /// </summary>
        public decimal AbsenteeismHours;
        /// <summary>
        /// все прогулы и отгулы
        /// </summary>
        public decimal AllAbsenteeismHours;
        /// <summary>
        /// сверхурочные
        /// </summary>
        public decimal OvertimeHours;
        /// <summary>
        /// все сверхурочные включая отчетную дату
        /// </summary>
        public decimal AllOvertimeHours;
        /// <summary>
        /// часы по неявке (без СУ и прогулов)
        /// </summary>
        public decimal AAbsenceHours;
        /// <summary>
        /// все часы по неявке (без СУ и прогулов) включая отчетную дату
        /// </summary>
        public decimal AllAbsenceHours;
        /// <summary>
        /// переработка до отчетной даты
        /// </summary>
        public decimal OverworkHours;
        /// <summary>
        /// планово в этот день
        /// </summary>
        public decimal PlanHours;
        /// <summary>
        /// планово до отчетной даты
        /// </summary>
        public decimal AllPlanHours;
        /// <summary>
        /// фактически в этот день
        /// </summary>
        public decimal FactHours;
        /// <summary>
        /// фактически в этот день
        /// </summary>
        public decimal AllFactHours;
        /// <summary>
        /// обед в этот день
        /// </summary>
        public decimal BreakHours;
        /// <summary>
        /// была неявка
        /// </summary>
        public bool IsAbsence;
        /// <summary>
        /// тип неявки
        /// </summary>
        public int AbsenceTypeID;
        /// <summary>
        /// 
        /// </summary>
        public string AbsenceFullName;
        /// <summary>
        /// 
        /// </summary>
        public string AbsenceShortName;
        /// <summary>
        /// ставка
        /// </summary>
        public string StrRate;
        /// <summary>
        /// часы по неявке
        /// </summary>
        public decimal AbsenceHours;
        /// <summary>
        /// в табель 
        /// </summary>
        public decimal TimesheetHours;
        /// <summary>
        /// день завершен
        /// </summary>
        public bool Saved;
    }


    public class InfiniumPictures
    {
        public DataTable AlbumsDataTable;
        public DataTable UsersDataTable;

        public DataTable AlbumsItemsDataTable;
        public DataTable PicturesItemsDataTable;
        private readonly FileManager FM;

        public InfiniumPictures()
        {
            AlbumsDataTable = new DataTable();
            UsersDataTable = new DataTable();

            AlbumsItemsDataTable = new DataTable();
            AlbumsItemsDataTable.Columns.Add(new DataColumn("AlbumID", Type.GetType("System.Int32")));
            AlbumsItemsDataTable.Columns.Add(new DataColumn("AlbumName", Type.GetType("System.String")));
            AlbumsItemsDataTable.Columns.Add(new DataColumn("PictureID", Type.GetType("System.Int32")));
            AlbumsItemsDataTable.Columns.Add(new DataColumn("AuthorID", Type.GetType("System.Int32")));
            AlbumsItemsDataTable.Columns.Add(new DataColumn("LikesCount", Type.GetType("System.Int32")));
            AlbumsItemsDataTable.Columns.Add(new DataColumn("CommentsCount", Type.GetType("System.Int32")));
            AlbumsItemsDataTable.Columns.Add(new DataColumn("AuthorName", Type.GetType("System.String")));
            AlbumsItemsDataTable.Columns.Add(new DataColumn("DateTime", Type.GetType("System.String")));
            AlbumsItemsDataTable.Columns.Add(new DataColumn("AuthorPhoto", Type.GetType("System.Byte[]")));
            AlbumsItemsDataTable.Columns.Add(new DataColumn("SmallSample1", Type.GetType("System.Byte[]")));
            AlbumsItemsDataTable.Columns.Add(new DataColumn("SmallSample2", Type.GetType("System.Byte[]")));
            AlbumsItemsDataTable.Columns.Add(new DataColumn("SmallSample3", Type.GetType("System.Byte[]")));
            AlbumsItemsDataTable.Columns.Add(new DataColumn("BigSample", Type.GetType("System.Byte[]")));
            AlbumsItemsDataTable.Columns.Add(new DataColumn("ActiveUser1", Type.GetType("System.Byte[]")));
            AlbumsItemsDataTable.Columns.Add(new DataColumn("ActiveUser2", Type.GetType("System.Byte[]")));
            AlbumsItemsDataTable.Columns.Add(new DataColumn("ActiveUser3", Type.GetType("System.Byte[]")));
            AlbumsItemsDataTable.Columns.Add(new DataColumn("ActiveUser4", Type.GetType("System.Byte[]")));

            PicturesItemsDataTable = new DataTable();
            PicturesItemsDataTable.Columns.Add(new DataColumn("AlbumID", Type.GetType("System.Int32")));
            PicturesItemsDataTable.Columns.Add(new DataColumn("PictureID", Type.GetType("System.Int32")));
            PicturesItemsDataTable.Columns.Add(new DataColumn("Image", Type.GetType("System.Byte[]")));
            PicturesItemsDataTable.Columns.Add(new DataColumn("LikesCount", Type.GetType("System.Int32")));
            PicturesItemsDataTable.Columns.Add(new DataColumn("FileSize", Type.GetType("System.Int64")));

            Fill();

            FM = new FileManager();
        }

        public void Fill()
        {
            UsersDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT UserID, Name, Photo FROM Users WHERE Fired <> 1 ", ConnectionStrings.UsersConnectionString))
            {
                DA.Fill(UsersDataTable);
            }
        }

        private Bitmap ResizeImage(Bitmap SourceImage, bool bBig)
        {
            //int iBigSampleWidth = 306;
            var iBigSampleHeight = 204;

            //int iSmallSampleWidth = 97;
            var iSmallSampleHeight = 64;

            var iSourceWidth = SourceImage.Width;
            var iSourceHeight = SourceImage.Height;

            var iNewHeight = 0;
            var iNewWidth = 0;

            if (bBig)
            {
                iNewHeight = iBigSampleHeight;
                iNewWidth = iSourceWidth * iNewHeight / iSourceHeight;
            }
            else
            {
                iNewHeight = iSmallSampleHeight;
                iNewWidth = iSourceWidth * iNewHeight / iSourceHeight;
            }

            var resImage = new Bitmap(iNewWidth, iNewHeight);

            //resize using proportions
            using (var gr = Graphics.FromImage(resImage))
            {
                gr.SmoothingMode = SmoothingMode.HighQuality;
                gr.InterpolationMode = InterpolationMode.HighQualityBicubic;
                gr.PixelOffsetMode = PixelOffsetMode.HighQuality;
                gr.DrawImage(SourceImage, new Rectangle(0, 0, iNewWidth, iNewHeight));
            }

            return resImage;
        }

        private Bitmap ResizePicSample(Bitmap SourceImage, int Width, int Height)
        {
            var iNewWidth = SourceImage.Width * Height / SourceImage.Height;

            var resImage = new Bitmap(iNewWidth, Height);

            //resize using proportions
            using (var gr = Graphics.FromImage(resImage))
            {
                gr.SmoothingMode = SmoothingMode.HighQuality;
                gr.InterpolationMode = InterpolationMode.HighQualityBicubic;
                gr.PixelOffsetMode = PixelOffsetMode.HighQuality;
                gr.DrawImage(SourceImage, new Rectangle(0, 0, iNewWidth, Height));
            }

            return resImage;
        }

        public void FillAlbums()
        {
            AlbumsDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT * FROM Albums", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(AlbumsDataTable);
            }

            foreach (DataRow Row in AlbumsDataTable.Rows)
            {
                var NewRow = AlbumsItemsDataTable.NewRow();

                NewRow["AlbumID"] = Row["AlbumID"];
                NewRow["AlbumName"] = Row["AlbumName"];
                NewRow["AuthorID"] = Row["Author"];
                NewRow["DateTime"] = Convert.ToDateTime(Row["CreationDateTime"]).ToString("dd MMMM yyyy HH:mm");
                NewRow["AuthorName"] = UsersDataTable.Select("UserID = " + Row["Author"])[0]["Name"];
                NewRow["AuthorPhoto"] = (byte[])UsersDataTable.Select("UserID = " + Row["Author"])[0]["Photo"];



                //pictures samples
                using (var DA = new SqlDataAdapter("SELECT TOP 4 * FROM Pictures WHERE AlbumID = " + Row["AlbumID"], ConnectionStrings.LightConnectionString))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        for (var i = 0; i < DT.Rows.Count; i++)
                        {
                            NewRow["PictureID"] = DT.Rows[i]["PictureID"];

                            if (i == 0)
                            {
                                try
                                {
                                    using (var ms = new MemoryStream(
                                            FM.ReadFile(Configs.DocumentsPath + "Pictures/" + DT.Rows[i]["PictureID"] + ".jpg",
                                                                          Convert.ToInt64(DT.Rows[i]["FileSize"]), Configs.FTPType)))
                                    {
                                        var bmp = (Bitmap)Image.FromStream(ms);

                                        bmp.RotateFlip(OrientationToFlipType(((int)ImageOrientation(bmp)).ToString()));

                                        var resultBMP = ResizeImage(bmp, true);

                                        var nms = new MemoryStream();

                                        var ImageCodecInfo = UserProfile.GetEncoderInfo("image/jpeg");
                                        var eEncoder1 = Encoder.Quality;

                                        var myEncoderParameter1 = new EncoderParameter(eEncoder1, 100L);
                                        var myEncoderParameters = new EncoderParameters(1);

                                        myEncoderParameters.Param[0] = myEncoderParameter1;

                                        resultBMP.Save(nms, ImageCodecInfo, myEncoderParameters);

                                        NewRow["BigSample"] = nms.ToArray();
                                        nms.Dispose();
                                        resultBMP.Dispose();
                                        bmp.Dispose();
                                    }
                                }
                                catch
                                {
                                    return;
                                }
                            }
                            else
                            {
                                try
                                {
                                    using (var ms = new MemoryStream(
                                            FM.ReadFile(Configs.DocumentsPath + "Pictures/" + DT.Rows[i]["PictureID"] + ".jpg",
                                                                          Convert.ToInt64(DT.Rows[i]["FileSize"]), Configs.FTPType)))
                                    {
                                        var bmp = (Bitmap)Image.FromStream(ms);
                                        bmp.RotateFlip(OrientationToFlipType(((int)ImageOrientation(bmp)).ToString()));
                                        var resultBMP = ResizeImage(bmp, false);

                                        var nms = new MemoryStream();

                                        var ImageCodecInfo = UserProfile.GetEncoderInfo("image/jpeg");
                                        var eEncoder1 = Encoder.Quality;

                                        var myEncoderParameter1 = new EncoderParameter(eEncoder1, 100L);
                                        var myEncoderParameters = new EncoderParameters(1);

                                        myEncoderParameters.Param[0] = myEncoderParameter1;

                                        resultBMP.Save(nms, ImageCodecInfo, myEncoderParameters);

                                        NewRow["SmallSample" + i] = nms.ToArray();
                                        nms.Dispose();
                                        resultBMP.Dispose();
                                        bmp.Dispose();
                                    }
                                }
                                catch
                                {
                                    return;
                                }
                            }
                        }

                    }
                }


                //likes
                {
                    using (var DA = new SqlDataAdapter("SELECT Count(PictureLikeID) FROM PicturesLikes " +
                                                       "WHERE PictureID IN (SELECT PictureID FROM Pictures WHERE AlbumID = " +
                                                       Row["AlbumID"] + ")", ConnectionStrings.LightConnectionString))
                    {
                        using (var DT = new DataTable())
                        {
                            DA.Fill(DT);

                            NewRow["LikesCount"] = DT.Rows[0][0];
                        }
                    }

                    AlbumsItemsDataTable.Rows.Add(NewRow);
                }


                var ActiveUsers = GetMostActiveUsers(Convert.ToInt32(Row["AlbumID"]));

                if (ActiveUsers.DefaultView.Count > 0)
                {
                    var c = 4;

                    if (ActiveUsers.DefaultView.Count < 4)
                        c = ActiveUsers.DefaultView.Count;

                    for (var i = 0; i < c; i++)
                    {
                        NewRow["ActiveUser" + (i + 1)] = (byte[])UsersDataTable.Select("UserID = " + ActiveUsers.DefaultView[i]["UserID"])[0]["Photo"];
                    }
                }
            }
        }

        public int RefreshLikes(int AlbumID)
        {
            using (var DA = new SqlDataAdapter("SELECT Count(PictureLikeID) FROM PicturesLikes " +
                                               "WHERE PictureID IN (SELECT PictureID FROM Pictures WHERE AlbumID = " +
                                               AlbumID + ")", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    AlbumsItemsDataTable.Select("AlbumID = " + AlbumID)[0]["LikesCount"] = DT.Rows[0][0];

                    return Convert.ToInt32(DT.Rows[0][0]);
                }
            }
        }

        public void FillPictures(int AlbumID)
        {
            PicturesItemsDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT * FROM Pictures WHERE AlbumID = " + AlbumID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    foreach (DataRow Row in DT.Rows)
                    {
                        var NewRow = PicturesItemsDataTable.NewRow();

                        NewRow["AlbumID"] = AlbumID;
                        NewRow["PictureID"] = Row["PictureID"];
                        NewRow["FileSize"] = Row["FileSize"];

                        using (var ms = new MemoryStream(
                                            FM.ReadFile(Configs.DocumentsPath + "Pictures/" + Row["PictureID"] + ".jpg",
                                                                          Convert.ToInt64(Row["FileSize"]), Configs.FTPType)))
                        {
                            var bmp = (Bitmap)Image.FromStream(ms);

                            bmp.RotateFlip(OrientationToFlipType(((int)ImageOrientation(bmp)).ToString()));

                            var resultBMP = ResizePicSample(bmp, 364, 242);

                            var nms = new MemoryStream();

                            var ImageCodecInfo = UserProfile.GetEncoderInfo("image/jpeg");
                            var eEncoder1 = Encoder.Quality;

                            var myEncoderParameter1 = new EncoderParameter(eEncoder1, 100L);
                            var myEncoderParameters = new EncoderParameters(1);

                            myEncoderParameters.Param[0] = myEncoderParameter1;

                            resultBMP.Save(nms, ImageCodecInfo, myEncoderParameters);



                            NewRow["Image"] = nms.ToArray();
                            nms.Dispose();
                            resultBMP.Dispose();
                            bmp.Dispose();
                        }


                        //likes
                        using (var lDA = new SqlDataAdapter("SELECT Count(PictureLikeID) FROM PicturesLikes " +
                                                            "WHERE PictureID = " + Row["PictureID"], ConnectionStrings.LightConnectionString))
                        {
                            using (var lDT = new DataTable())
                            {
                                lDA.Fill(lDT);

                                NewRow["LikesCount"] = lDT.Rows[0][0];
                            }
                        }

                        PicturesItemsDataTable.Rows.Add(NewRow);
                    }
                }



            }
        }

        private const int OrientationId = 0x0112;

        public enum ExifOrientations : byte
        {
            Unknown = 0,
            TopLeft = 1,
            TopRight = 2,
            BottomRight = 3,
            BottomLeft = 4,
            LeftTop = 5,
            RightTop = 6,
            RightBottom = 7,
            LeftBottom = 8,
        }

        public static ExifOrientations ImageOrientation(Image img)
        {
            // Get the index of the orientation property.
            var orientation_index = Array.IndexOf(img.PropertyIdList, OrientationId);

            // If there is no such property, return Unknown.
            if (orientation_index < 0) return ExifOrientations.Unknown;

            // Return the orientation value.
            return (ExifOrientations)img.GetPropertyItem(OrientationId).Value[0];
        }

        private static RotateFlipType OrientationToFlipType(string orientation)
        {
            switch (int.Parse(orientation))
            {
                case 1:
                    return RotateFlipType.RotateNoneFlipNone;
                case 2:
                    return RotateFlipType.RotateNoneFlipX;
                case 3:
                    return RotateFlipType.Rotate180FlipNone;
                case 4:
                    return RotateFlipType.Rotate180FlipX;
                case 5:
                    return RotateFlipType.Rotate90FlipX;
                case 6:
                    return RotateFlipType.Rotate90FlipNone;
                case 7:
                    return RotateFlipType.Rotate270FlipX;
                case 8:
                    return RotateFlipType.Rotate270FlipNone;
                default:
                    return RotateFlipType.RotateNoneFlipNone;
            }
        }

        public DataTable GetMostActiveUsers(int AlbumID)
        {
            var uDT = new DataTable();
            uDT.Columns.Add(new DataColumn("UserID", Type.GetType("System.Int32")));
            uDT.Columns.Add(new DataColumn("Count", Type.GetType("System.Int32")));


            using (var DA = new SqlDataAdapter("SELECT UserID FROM PicturesLikes " +
                                               "WHERE PictureID IN (SELECT PictureID FROM Pictures WHERE AlbumID = " +
                                               AlbumID + ")", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    foreach (DataRow Row in DT.Rows)
                    {
                        var fRow = uDT.Select("UserID = " + Row["UserID"]);

                        if (fRow.Count() > 0)
                            fRow[0]["Count"] = Convert.ToInt32(fRow[0]["Count"]) + 1;
                        else
                        {
                            var NewRow = uDT.NewRow();
                            NewRow["UserID"] = Row["UserID"];
                            NewRow["Count"] = 1;
                            uDT.Rows.Add(NewRow);
                        }
                    }
                }
            }


            uDT.DefaultView.Sort = "Count DESC";

            if (uDT.DefaultView.Count > 4)
            {
                for (var i = 4; i < uDT.DefaultView.Count; i++)
                {
                    uDT.DefaultView[i].Delete();
                }
            }

            uDT.AcceptChanges();

            return uDT;
        }

        public Image GetPicture(int PictureID)
        {
            using (var ms = new MemoryStream(
                   FM.ReadFile(Configs.DocumentsPath + "Pictures/" + PictureID + ".jpg",
                               Convert.ToInt64(PicturesItemsDataTable.Select("PictureID = " + PictureID)[0]["FileSize"]), Configs.FTPType)))
            {
                var I = Image.FromStream(ms);

                I.RotateFlip(OrientationToFlipType(((int)ImageOrientation(I)).ToString()));

                return I;
            }
        }

        public int GetNextPictureID(int PictureID)
        {
            var index = PicturesItemsDataTable.Rows.IndexOf(PicturesItemsDataTable.Select("PictureID = " + PictureID)[0]);

            if (index == PicturesItemsDataTable.Rows.Count - 1)
                return -1;

            return Convert.ToInt32(PicturesItemsDataTable.Rows[++index]["PictureID"]);
        }

        public int GetPreviousPictureID(int PictureID)
        {
            var index = PicturesItemsDataTable.Rows.IndexOf(PicturesItemsDataTable.Select("PictureID = " + PictureID)[0]);

            if (index == 0)
                return -1;

            return Convert.ToInt32(PicturesItemsDataTable.Rows[--index]["PictureID"]);
        }

        public int LikePicture(int PictureID)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM PicturesLikes WHERE PictureID = " + PictureID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        if (DT.Select("UserID = " + Security.CurrentUserID).Count() > 0)//i like
                        {
                            using (var dDA = new SqlDataAdapter("DELETE FROM PicturesLikes WHERE PictureID = " + PictureID +
                                                                " AND UserID = " + Security.CurrentUserID, ConnectionStrings.LightConnectionString))
                            {
                                using (var dDT = new DataTable())
                                {
                                    dDA.Fill(dDT);
                                }
                            }

                            return DT.Rows.Count - 1;
                        }

                        var NewRow = DT.NewRow();
                        NewRow["PictureID"] = PictureID;
                        NewRow["UserID"] = Security.CurrentUserID;
                        NewRow["DateTime"] = Security.GetCurrentDate();
                        DT.Rows.Add(NewRow);

                        DA.Update(DT);

                        return DT.Rows.Count;
                    }
                }
            }
        }
    }





    public class InfiniumFiles
    {
        public DataTable CurrentItemsDataTable;
        public DataTable DocumentAttributesDataTable;
        public DataTable CurrentSignsDataTable;
        public DataTable CurrentReadDataTable;
        public DataTable CurrentAttributesDataTable;
        public DataTable CurrentPermissionsUsersDataTable;
        public DataTable CurrentPermissionsDepsDataTable;
        public DataTable DepartmentsDataTable;
        public DataTable CategoriesDataTable;

        public FileManager FM;

        public DataTable UsersDataTable;

        public bool bFirstSign;

        public InfiniumFiles()
        {
            FM = new FileManager();

            CurrentItemsDataTable = new DataTable();
            CurrentItemsDataTable.Columns.Add(new DataColumn("ItemName", Type.GetType("System.String")));
            CurrentItemsDataTable.Columns.Add(new DataColumn("FileID", Type.GetType("System.Int32")));
            CurrentItemsDataTable.Columns.Add(new DataColumn("FolderID", Type.GetType("System.Int32")));
            CurrentItemsDataTable.Columns.Add(new DataColumn("Extension", Type.GetType("System.String")));
            CurrentItemsDataTable.Columns.Add(new DataColumn("FileSize", Type.GetType("System.Int32")));
            CurrentItemsDataTable.Columns.Add(new DataColumn("Checked", Type.GetType("System.Boolean")));
            CurrentItemsDataTable.Columns.Add(new DataColumn("Author", Type.GetType("System.String")));
            CurrentItemsDataTable.Columns.Add(new DataColumn("LastModified", Type.GetType("System.String")));

            UsersDataTable = new DataTable();




            DocumentAttributesDataTable = new DataTable();
            DocumentAttributesDataTable.Columns.Add(new DataColumn("AttributeName", Type.GetType("System.String")));
            DocumentAttributesDataTable.Columns.Add(new DataColumn("Value", Type.GetType("System.Boolean")));

            CurrentAttributesDataTable = new DataTable();
            CurrentAttributesDataTable = DocumentAttributesDataTable.Clone();

            CurrentSignsDataTable = new DataTable();
            CurrentSignsDataTable.Columns.Add(new DataColumn("UserID", Type.GetType("System.Int32")));
            CurrentSignsDataTable.Columns.Add(new DataColumn("SignDescription", Type.GetType("System.String")));
            CurrentSignsDataTable.Columns.Add(new DataColumn("FileID", Type.GetType("System.String")));
            CurrentSignsDataTable.Columns.Add(new DataColumn("IsSigned", Type.GetType("System.Boolean")));

            CurrentReadDataTable = new DataTable();
            CurrentReadDataTable.Columns.Add(new DataColumn("UserID", Type.GetType("System.Int32")));
            CurrentReadDataTable.Columns.Add(new DataColumn("SignDescription", Type.GetType("System.String")));
            CurrentReadDataTable.Columns.Add(new DataColumn("FileID", Type.GetType("System.String")));
            CurrentReadDataTable.Columns.Add(new DataColumn("IsSigned", Type.GetType("System.Boolean")));

            DepartmentsDataTable = new DataTable();

            CurrentPermissionsDepsDataTable = new DataTable();

            CurrentPermissionsDepsDataTable = new DataTable();
            CurrentPermissionsDepsDataTable.Columns.Add(new DataColumn("DepartmentID", Type.GetType("System.Int32")));
            CurrentPermissionsDepsDataTable.Columns.Add(new DataColumn("DepartmentName", Type.GetType("System.String")));


            CurrentPermissionsUsersDataTable = new DataTable();

            CurrentPermissionsUsersDataTable = new DataTable();
            CurrentPermissionsUsersDataTable.Columns.Add(new DataColumn("UserID", Type.GetType("System.Int32")));
            CurrentPermissionsUsersDataTable.Columns.Add(new DataColumn("UserName", Type.GetType("System.String")));

            CategoriesDataTable = new DataTable();
            CategoriesDataTable.Columns.Add(new DataColumn("Category", Type.GetType("System.String")));
            CategoriesDataTable.Columns.Add(new DataColumn("FolderID", Type.GetType("System.Int32")));

            Fill();
        }

        public void Fill()
        {
            {
                var NewRow = DocumentAttributesDataTable.NewRow();
                NewRow["AttributeName"] = "Подписи";
                NewRow["Value"] = false;
                DocumentAttributesDataTable.Rows.Add(NewRow);
            }

            {
                var NewRow = DocumentAttributesDataTable.NewRow();
                NewRow["AttributeName"] = "Ознакомлен";
                NewRow["Value"] = false;
                DocumentAttributesDataTable.Rows.Add(NewRow);
            }

            {
                var NewRow = DocumentAttributesDataTable.NewRow();
                NewRow["AttributeName"] = "Бумага";
                NewRow["Value"] = false;
                DocumentAttributesDataTable.Rows.Add(NewRow);
            }


            //categories
            {
                var NewRow = CategoriesDataTable.NewRow();
                NewRow["Category"] = "Общие файлы";
                NewRow["FolderID"] = 2;
                CategoriesDataTable.Rows.Add(NewRow);
            }

            {
                var NewRow = CategoriesDataTable.NewRow();
                NewRow["Category"] = "Отделы";
                NewRow["FolderID"] = 1;
                CategoriesDataTable.Rows.Add(NewRow);
            }

            {
                var NewRow = CategoriesDataTable.NewRow();
                NewRow["Category"] = "Личные файлы";
                NewRow["FolderID"] = -3;
                CategoriesDataTable.Rows.Add(NewRow);
            }

            {
                var NewRow = CategoriesDataTable.NewRow();
                NewRow["Category"] = "Клиенты";
                NewRow["FolderID"] = 4;
                CategoriesDataTable.Rows.Add(NewRow);
            }

            {
                var NewRow = CategoriesDataTable.NewRow();
                NewRow["Category"] = "ЗОВ";
                NewRow["FolderID"] = 5;
                CategoriesDataTable.Rows.Add(NewRow);
            }

            //{
            //    DataRow NewRow = CategoriesDataTable.NewRow();
            //    NewRow["Category"] = "Документооборот";
            //    NewRow["FolderID"] = 6;
            //    CategoriesDataTable.Rows.Add(NewRow);
            //}

            {
                var NewRow = CategoriesDataTable.NewRow();
                NewRow["Category"] = "На подпись";
                NewRow["FolderID"] = -1;
                CategoriesDataTable.Rows.Add(NewRow);
            }

            {
                var NewRow = CategoriesDataTable.NewRow();
                NewRow["Category"] = "На ознакомление";
                NewRow["FolderID"] = -2;
                CategoriesDataTable.Rows.Add(NewRow);
            }



            using (var DA = new SqlDataAdapter("SELECT UserID, DepartmentID, Name FROM Users ORDER BY Name ASC", ConnectionStrings.UsersConnectionString))
            {
                DA.Fill(UsersDataTable);
            }

            using (var DA = new SqlDataAdapter("SELECT DepartmentID, DepartmentName FROM Departments", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(DepartmentsDataTable);
            }
        }


        public void AddUploadPending(int FileID, int UserID)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM FilesUploadsPending WHERE FileID = " + FileID + " AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        if (DT.Rows.Count > 0)//last crashed uploads
                        {
                            foreach (DataRow Row in DT.Rows)
                            {
                                Row.Delete();
                            }

                            DA.Update(DT);

                            DT.Clear();
                            DT.AcceptChanges();

                            DA.Fill(DT);
                        }

                        var NewRow = DT.NewRow();
                        NewRow["UserID"] = UserID;
                        NewRow["FileID"] = FileID;
                        NewRow["DateTime"] = Security.GetCurrentDate();
                        DT.Rows.Add(NewRow);

                        DA.Update(DT);
                    }
                }
            }
        }

        public bool EnterFolder(int FolderID)//loads folders and files for selected folderid
        {
            CurrentItemsDataTable.Clear();

            var iRootFolderID = -1;

            using (var DA = new SqlDataAdapter("SELECT * FROM Folders WHERE FolderID = " + FolderID + " ORDER BY FolderName ASC", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    iRootFolderID = Convert.ToInt32(DT.Rows[0]["RootFolderID"]);
                }

            }

            if (iRootFolderID > 0)
            {
                var NewRow = CurrentItemsDataTable.NewRow();
                NewRow["ItemName"] = "[...]";
                NewRow["Extension"] = "root";
                NewRow["FolderID"] = iRootFolderID;
                CurrentItemsDataTable.Rows.Add(NewRow);
            }



            using (var DA = new SqlDataAdapter("SELECT * FROM Folders WHERE RootFolderID = " + FolderID + " ORDER BY FolderName ASC", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(CurrentItemsDataTable);
            }

            foreach (DataRow Row in CurrentItemsDataTable.Rows)
            {
                if (Row["ItemName"].ToString().Length == 0)
                    Row["ItemName"] = Row["FolderName"];

                Row["FileID"] = -1;
                Row["FileSize"] = -1;

                if (Row["Extension"].ToString() != "root")
                    Row["Extension"] = "folder";
                else
                    continue;

                var Author = Row["Author"].ToString();
                var name = string.Empty;
                var CreationDateTime = string.Empty;
                var rows = UsersDataTable.Select("UserID = " + Author);
                if (rows.Count() > 0)
                {
                    Row["Author"] = UsersDataTable.Select("UserID = " + Row["Author"])[0]["Name"] + "\n" +
                                                Convert.ToDateTime(Row["CreationDateTime"]).ToString("dd.MM.yyyy HH:mm:ss");

                    Row["LastModified"] = UsersDataTable.Select("UserID = " + Row["LastModifiedUserID"])[0]["Name"] + "\n" +
                                              Convert.ToDateTime(Row["LastModifiedDateTime"]).ToString("dd.MM.yyyy HH:mm:ss");
                }

            }

            using (var DA = new SqlDataAdapter("SELECT * FROM Files WHERE FolderID = " + FolderID + " ORDER BY FileName ASC", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    foreach (DataRow Row in DT.Rows)
                    {
                        var NewRow = CurrentItemsDataTable.NewRow();
                        NewRow["ItemName"] = Row["FileName"];
                        NewRow["FolderID"] = FolderID;
                        NewRow["Extension"] = Row["FileExtension"];
                        NewRow["FileID"] = Row["FileID"];
                        NewRow["FileSize"] = Row["FileSize"];

                        NewRow["Author"] = UsersDataTable.Select("UserID = " + Row["Author"])[0]["Name"] + "\n" +
                                           Convert.ToDateTime(Row["CreationDateTime"]).ToString("dd.MM.yyyy HH:mm:ss");

                        NewRow["LastModified"] = UsersDataTable.Select("UserID = " + Row["LastModifiedUserID"])[0]["Name"] + "\n" +
                                           Convert.ToDateTime(Row["LastModifiedDateTime"]).ToString("dd.MM.yyyy HH:mm:ss");

                        CurrentItemsDataTable.Rows.Add(NewRow);
                    }
                }
            }

            return true;
        }

        public int CreateFolder(int FolderID, string FolderName)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM Folders WHERE FolderID = " + FolderID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        var NewRow = DT.NewRow();
                        NewRow["RootFolderID"] = FolderID;
                        NewRow["FolderName"] = FolderName;
                        NewRow["FolderPath"] = DT.Rows[0]["FolderPath"] + "/" + FolderName;
                        NewRow["Author"] = Security.CurrentUserID;

                        var Date = Security.GetCurrentDate();

                        NewRow["CreationDateTime"] = Date;
                        NewRow["LastModifiedUserID"] = NewRow["Author"];
                        NewRow["LastModifiedDateTime"] = NewRow["CreationDateTime"];

                        SetLastModified(FolderID, Date, Convert.ToInt32(NewRow["Author"]));

                        NewRow["FTPHost"] = DT.Rows[0]["FTPHost"];
                        NewRow["ParentFolderID"] = DT.Rows[0]["ParentFolderID"];

                        var iPermission = GetInheritedPermission(FolderID);

                        NewRow["FolderPermission"] = iPermission;
                        NewRow["InheritedPermission"] = iPermission;

                        DT.Rows.Add(NewRow);

                        var sPath = "";
                        var FTPType = -1;

                        if (IsPathHost(FolderID))
                        {
                            sPath = Configs.DocumentsPathHost;
                            FTPType = 1;
                        }
                        else
                        {
                            sPath = Configs.DocumentsPath;
                            FTPType = Configs.FTPType;
                        }

                        var res = FM.CreateFolder(sPath + DT.Rows[0]["FolderPath"] + "/", FolderName, FTPType);

                        if (res == -1)
                            return -1;//directory exists

                        if (res == 0)
                            return 0;//error

                        DA.Update(DT);

                        SetPermissions(Convert.ToDateTime(NewRow["CreationDateTime"]));
                    }
                }
            }



            return 1;//ok
        }

        public int GetRootFolderID(int FolderID)
        {
            using (var DA = new SqlDataAdapter("SELECT RootFolderID FROM Folders WHERE FolderID = " + FolderID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    return Convert.ToInt32(DT.Rows[0]["RootFolderID"]);
                }
            }
        }

        public bool CheckFileExist(string[] FileNames, int FolderID)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM Files WHERE FolderID =" + FolderID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    foreach (var FileName in FileNames)
                    {
                        if (DT.Select("FileName = '" + Path.GetFileName(FileName) + "'").Count() > 0)
                        {
                            return true;
                        }
                    }
                }
            }

            return false;
        }

        public bool CanEditFile(int UserID, int FileID)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsAdmins WHERE UserID = " + Security.CurrentUserID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    if (DA.Fill(DT) > 0)
                        return true;
                }
            }

            using (var DA = new SqlDataAdapter("SELECT FileID, Author FROM Files WHERE FileID = " + FileID + " AND Author = " + UserID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    if (DA.Fill(DT) > 0)
                        return true;

                    return false;
                }
            }
        }

        public bool CheckFolderPermission(int UserID, int FolderID)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsAdmins WHERE UserID = " + Security.CurrentUserID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    if (DA.Fill(DT) > 0)
                        return true;
                }
            }

            using (var DA = new SqlDataAdapter("SELECT FolderID, FolderPermission FROM Folders WHERE FolderID = " + FolderID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    if (DT.Rows[0]["FolderPermission"].ToString() == "2") //all users
                        return true;

                    if (DT.Rows[0]["FolderPermission"].ToString() == "0") //admin only
                        return false;
                }
            }

            using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsPermissions WHERE FolderID = " + FolderID +
                                               " AND ((UserTypeID = 0 AND UserID = " + UserID + ") OR (UserTypeID = 1 AND UserID = " +
                                               UsersDataTable.Select("UserID = " + UserID)[0]["DepartmentID"] + "))", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    if (DA.Fill(DT) > 0)
                        return true;

                    return false;
                }
            }
        }//for rename and delete folder

        public bool CheckInheritedPermission(int UserID, int FolderID)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsAdmins WHERE UserID = " + Security.CurrentUserID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    if (DA.Fill(DT) > 0)
                        return true;
                }
            }

            using (var DA = new SqlDataAdapter("SELECT FolderID, InheritedPermission FROM Folders WHERE FolderID = " + FolderID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    if (DT.Rows[0]["InheritedPermission"].ToString() == "2") //all users
                        return true;

                    if (DT.Rows[0]["InheritedPermission"].ToString() == "0") //admin only
                        return false;
                }
            }

            using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsPermissions WHERE FolderID = " + FolderID +
                                               " AND ((UserTypeID = 0 AND UserID = " + UserID + ") OR (UserTypeID = 1 AND UserID = " +
                                               UsersDataTable.Select("UserID = " + UserID)[0]["DepartmentID"] + "))", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    if (DA.Fill(DT) > 0)
                        return true;

                    return false;
                }
            }
        }//for upload, rename, delete folder content

        public int CheckCheckedItems(DataTable ItemsDataTable)
        {
            if (ItemsDataTable.Select("Checked = 1 AND Extension = 'folder'").Count() > 0
               && ItemsDataTable.Select("Checked = 1 AND FileID > -1").Count() == 0)
                return -1;

            if (ItemsDataTable.Select("Checked = 1 AND Extension = 'folder'").Count() > 0
               && ItemsDataTable.Select("Checked = 1 AND FileID > -1").Count() > 0)
                return -2;

            if (ItemsDataTable.Select("Checked = 1 AND Extension = 'folder'").Count() == 0
               && ItemsDataTable.Select("Checked = 1 AND FileID > -1").Count() > 0)
                return 1;

            if (ItemsDataTable.Select("Checked = 1").Count() == 0)
                return 0;

            return -3;
        }

        public bool CheckFileExists(string Path)
        {
            var fi = new FileInfo(Path);

            return fi.Exists;
        }

        public bool CheckFilesExists(DataTable ItemsDataTable, string Path)
        {
            foreach (var Row in ItemsDataTable.Select("Checked = 1 AND FileID > -1"))
            {
                var fi = new FileInfo(Path + Row["ItemName"]);

                if (fi.Exists)
                    return true;
            }

            return false;
        }

        public static int CheckPersonalFolder(int UserID)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM Folders WHERE FolderName = '" + UserID + "'", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    if (DA.Fill(DT) == 0)
                        return -1;

                    return Convert.ToInt32(DT.Rows[0]["FolderID"]);
                }
            }
        }

        public void CreateClientFolders(string Name)
        {
            CreateFolder(4, Name);

            using (var DA = new SqlDataAdapter("SELECT FolderID FROM Folders WHERE FolderPath = '" + @"Клиенты/" + Name + "'", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    CreateFolder(Convert.ToInt32(DT.Rows[0]["FolderID"]), "Отгрузочные ведомости");
                    CreateFolder(Convert.ToInt32(DT.Rows[0]["FolderID"]), "Счет-фактуры");
                    CreateFolder(Convert.ToInt32(DT.Rows[0]["FolderID"]), "ТТН");
                    CreateFolder(Convert.ToInt32(DT.Rows[0]["FolderID"]), "Протоколы");
                    CreateFolder(Convert.ToInt32(DT.Rows[0]["FolderID"]), "Договоры");
                }
            }
        }

        public void CreateFolderClient(string Name, string FolderName)
        {
            using (var DA = new SqlDataAdapter("SELECT FolderID FROM Folders WHERE FolderPath = '" + @"Клиенты/" + Name + "'", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    CreateFolder(Convert.ToInt32(DT.Rows[0]["FolderID"]), FolderName);
                }
            }
        }

        public void RenameFolder(string FolderName)
        {

        }

        public void CreatePersonalFolder(int UserID)
        {
            using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM Folders", ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        var NewRow = DT.NewRow();
                        NewRow["FolderName"] = UserID.ToString();
                        NewRow["FolderPath"] = "Личные файлы/" + UserID;
                        NewRow["RootFolderID"] = 3;
                        NewRow["ParentFolderID"] = 3;
                        NewRow["Author"] = Security.CurrentUserID;

                        var Date = Security.GetCurrentDate();

                        NewRow["CreationDateTime"] = Date;
                        NewRow["LastModifiedDateTime"] = NewRow["CreationDateTime"];
                        NewRow["LastModifiedUserID"] = Security.CurrentUserID;
                        NewRow["FolderPermission"] = 1;
                        NewRow["InheritedPermission"] = 1;
                        DT.Rows.Add(NewRow);

                        DA.Update(DT);

                        SetLastModified(3, Date, Security.CurrentUserID);

                        SetPermissionsPersonalFolder(Convert.ToDateTime(NewRow["CreationDateTime"]), UserID);

                        FM.CreateFolder("Личные файлы/", UserID.ToString(), Configs.FTPType);
                    }
                }
            }
        }

        public bool CheckFileVersion(int FileID, int UserID)//true if ok to replace
        {
            using (var DA = new SqlDataAdapter("SELECT CreationDateTime, LastModifiedDateTime, LastModifiedUserID  FROM Files WHERE FileID = " + FileID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    if (Convert.ToDateTime(DT.Rows[0]["CreationDateTime"]) == Convert.ToDateTime(DT.Rows[0]["LastModifiedDateTime"]))
                        return true;

                    if (Convert.ToInt32(DT.Rows[0]["LastModifiedUserID"]) == UserID) //this user last modified
                        return true;

                    using (var fDA = new SqlDataAdapter("SELECT * FROM FilesDownloadsLog WHERE FileID = " + FileID + " AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
                    {
                        using (var fDT = new DataTable())
                        {
                            fDA.Fill(fDT);

                            if (fDT.Rows.Count == 0)//file never been downloded by this user (maybe he is author)
                            {
                                return false;
                            }

                            if (Convert.ToDateTime(fDT.Rows[fDT.Rows.Count - 1]["DateTime"]) >
                                Convert.ToDateTime(DT.Rows[0]["LastModifiedDateTime"]))//ok
                                return true;

                            return false;//someone replaced file after this user download it

                        }
                    }
                }
            }
        }

        public void ClearFileLog(int FileID)
        {
            using (var DA = new SqlDataAdapter("DELETE FROM FilesDownloadsLog WHERE FileID = " + FileID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }
        }

        public bool CheckUploadPending(int FileID)//false if no one uploads
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM FilesUploadsPending WHERE FileID = " + FileID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        if (DA.Fill(DT) > 0)//someone uploads
                        {
                            var DateTime = Convert.ToDateTime(DT.Rows[0]["DateTime"]);

                            if (DateTime.AddSeconds(20) < Security.GetCurrentDate())//someone's upload crashed
                            {
                                foreach (DataRow Row in DT.Rows)
                                {
                                    Row.Delete();
                                }

                                DA.Update(DT);

                                return false;
                            }

                            return true;
                        }

                        return false;
                    }
                }
            }
        }

        public void ClearUploadPending(int FileID, int UserID)
        {
            using (var DA = new SqlDataAdapter("DELETE FROM FilesUploadsPending WHERE FileID = " + FileID + " AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }
        }

        public void ClearAttributes()
        {
            foreach (DataRow Row in DocumentAttributesDataTable.Rows)
            {
                Row["Value"] = false;
            }
        }

        public void FillSignsFiles(int UserID)
        {
            CurrentItemsDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT * FROM Files WHERE FileID IN " +
                                               "(SELECT FileID FROM DocumentsSigns WHERE IsSigned = 0 AND SignType = 0 AND UserID = " + UserID + ")", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    foreach (DataRow Row in DT.Rows)
                    {
                        var NewRow = CurrentItemsDataTable.NewRow();
                        NewRow["ItemName"] = Row["FileName"];
                        NewRow["FolderID"] = -1;
                        NewRow["Extension"] = Row["FileExtension"];
                        NewRow["FileID"] = Row["FileID"];
                        NewRow["FileSize"] = Row["FileSize"];

                        NewRow["Author"] = UsersDataTable.Select("UserID = " + Row["Author"])[0]["Name"] + "\n" +
                                           Convert.ToDateTime(Row["CreationDateTime"]).ToString("dd.MM.yyyy HH:mm");

                        NewRow["LastModified"] = UsersDataTable.Select("UserID = " + Row["LastModifiedUserID"])[0]["Name"] + "\n" +
                                           Convert.ToDateTime(Row["LastModifiedDateTime"]).ToString("dd.MM.yyyy HH:mm");

                        CurrentItemsDataTable.Rows.Add(NewRow);
                    }
                }
            }
        }

        public void FillReadFiles(int UserID)
        {
            CurrentItemsDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT * FROM Files WHERE FileID IN " +
                                               "(SELECT FileID FROM DocumentsSigns WHERE IsSigned = 0 AND SignType = 1 AND UserID = " + UserID + ")", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    foreach (DataRow Row in DT.Rows)
                    {
                        var NewRow = CurrentItemsDataTable.NewRow();
                        NewRow["ItemName"] = Row["FileName"];
                        NewRow["FolderID"] = -1;
                        NewRow["Extension"] = Row["FileExtension"];
                        NewRow["FileID"] = Row["FileID"];
                        NewRow["FileSize"] = Row["FileSize"];

                        NewRow["Author"] = UsersDataTable.Select("UserID = " + Row["Author"])[0]["Name"] + "\n" +
                                           Convert.ToDateTime(Row["CreationDateTime"]).ToString("dd.MM.yyyy HH:mm");

                        NewRow["LastModified"] = UsersDataTable.Select("UserID = " + Row["LastModifiedUserID"])[0]["Name"] + "\n" +
                                           Convert.ToDateTime(Row["LastModifiedDateTime"]).ToString("dd.MM.yyyy HH:mm");

                        CurrentItemsDataTable.Rows.Add(NewRow);
                    }
                }
            }
        }

        public void FillAttributes(int FileID)
        {
            CurrentSignsDataTable.Clear();
            CurrentReadDataTable.Clear();
            CurrentAttributesDataTable.Clear();

            if (FileID == -1)
                return;

            using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsSigns WHERE SignType = 0 AND FileID = " + FileID, ConnectionStrings.LightConnectionString))
            {
                DA.Fill(CurrentSignsDataTable);
            }

            if (CurrentSignsDataTable.Rows.Count > 0)
            {
                var NewRow = CurrentAttributesDataTable.NewRow();
                NewRow["AttributeName"] = "Подписи";
                NewRow["Value"] = true;
                CurrentAttributesDataTable.Rows.Add(NewRow);
            }


            using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsSigns WHERE SignType = 1 AND FileID = " + FileID, ConnectionStrings.LightConnectionString))
            {
                DA.Fill(CurrentReadDataTable);
            }

            if (CurrentReadDataTable.Rows.Count > 0)
            {
                var NewRow = CurrentAttributesDataTable.NewRow();
                NewRow["AttributeName"] = "Ознакомлен";
                NewRow["Value"] = true;
                CurrentAttributesDataTable.Rows.Add(NewRow);
            }


            bFirstSign = false;

            using (var DA = new SqlDataAdapter("SELECT FileID, NeedPrint, FirstSign FROM Files WHERE FileID = " + FileID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    var NewRow = CurrentAttributesDataTable.NewRow();
                    NewRow["AttributeName"] = "Бумага";
                    NewRow["Value"] = DT.Rows[0]["NeedPrint"];
                    CurrentAttributesDataTable.Rows.Add(NewRow);

                    bFirstSign = Convert.ToBoolean(DT.Rows[0]["FirstSign"]);
                }
            }
        }

        public int FillPermissionsUsers(int FolderID)
        {
            CurrentPermissionsDepsDataTable.Clear();
            CurrentPermissionsUsersDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsPermissions WHERE FolderID = " + FolderID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    if (DA.Fill(DT) > 0)
                    {
                        var ARows = DT.Select("UserTypeID = -1");//admin

                        if (ARows.Count() > 0)
                            return -1;

                        var AlRows = DT.Select("UserTypeID = -2");//all

                        if (AlRows.Count() > 0)
                            return -2;


                        var DRows = DT.Select("UserTypeID = 1");

                        foreach (var DRow in DRows)
                        {
                            var NewRow = CurrentPermissionsDepsDataTable.NewRow();
                            NewRow["DepartmentID"] = DRow["UserID"];
                            NewRow["DepartmentName"] = DepartmentsDataTable.Select("DepartmentID = " + DRow["UserID"])[0]["DepartmentName"];
                            CurrentPermissionsDepsDataTable.Rows.Add(NewRow);
                        }


                        var URows = DT.Select("UserTypeID = 0");

                        foreach (var URow in URows)
                        {
                            var NewRow = CurrentPermissionsUsersDataTable.NewRow();
                            NewRow["UserID"] = URow["UserID"];
                            NewRow["UserName"] = UsersDataTable.Select("UserID = " + URow["UserID"])[0]["Name"];
                            CurrentPermissionsUsersDataTable.Rows.Add(NewRow);
                        }

                        return 1;
                    }
                }
            }

            return 0;
        }

        public int GetInheritedPermission(int RootFolderID)
        {
            using (var DA = new SqlDataAdapter("SELECT FolderID, InheritedPermission FROM Folders WHERE FolderID = " + RootFolderID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    return Convert.ToInt32(DT.Rows[0]["InheritedPermission"]);
                }
            }
        }

        public string GetFolderPath(int FolderID)
        {
            using (var DA = new SqlDataAdapter("SELECT FolderPath FROM Folders WHERE FolderID = " + FolderID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    return DT.Rows[0]["FolderPath"].ToString();
                }
            }
        }

        public int GetSignCount(int UserID)
        {
            using (var DA = new SqlDataAdapter("SELECT COUNT(DocumentSignID) FROM DocumentsSigns WHERE SignType = 0 AND IsSigned = 0 AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    return Convert.ToInt32(DT.Rows[0][0]);
                }
            }
        }

        public int GetReadCount(int UserID)
        {
            using (var DA = new SqlDataAdapter("SELECT COUNT(DocumentSignID) FROM DocumentsSigns WHERE SignType = 1 AND IsSigned = 0 AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    return Convert.ToInt32(DT.Rows[0][0]);
                }
            }
        }

        public bool IsAdmin(int UserID)
        {
            using (var DA = new SqlDataAdapter("SELECT UserID FROM DocumentsAdmins WHERE UserID = " + UserID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    return (DA.Fill(DT) > 0);
                }
            }
        }

        public bool IsPathHost(int FolderID)
        {
            using (var DA = new SqlDataAdapter("SELECT FTPHost FROM Folders WHERE FolderID = " + FolderID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    return Convert.ToBoolean(DT.Rows[0]["FTPHost"]);
                }
            }
        }


        public bool OpenFile(int FileID)
        {

            var bOK = false;

            using (var fDA = new SqlDataAdapter("SELECT FileSize, FileName, FolderID FROM Files WHERE FileID = " + FileID, ConnectionStrings.LightConnectionString))
            {
                using (var fDT = new DataTable())
                {
                    fDA.Fill(fDT);


                    using (var DA = new SqlDataAdapter("SELECT FolderPath FROM Folders WHERE FolderID = " + fDT.Rows[0]["FolderID"], ConnectionStrings.LightConnectionString))
                    {
                        using (var DT = new DataTable())
                        {
                            DA.Fill(DT);

                            var sPath = "";
                            var FTPType = -1;

                            if (IsPathHost(Convert.ToInt32(fDT.Rows[0]["FolderID"])))
                            {
                                sPath = Configs.DocumentsPathHost;
                                FTPType = 1;
                            }
                            else
                            {
                                sPath = Configs.DocumentsPath;
                                FTPType = Configs.FTPType;
                            }

                            bOK = FM.DownloadFile(sPath + DT.Rows[0]["FolderPath"] + "/" + fDT.Rows[0]["FileName"],
                                            Environment.GetEnvironmentVariable("TEMP") + "/" +
                                            fDT.Rows[0]["FileName"], Convert.ToInt64(fDT.Rows[0]["FileSize"]), FTPType);

                            if (bOK)
                            {
                                var myProcess = new Process();
                                myProcess.StartInfo.UseShellExecute = true;
                                myProcess.StartInfo.FileName = Environment.GetEnvironmentVariable("TEMP") + "/" +
                                        fDT.Rows[0]["FileName"];
                                myProcess.StartInfo.CreateNoWindow = true;
                                myProcess.Start();
                            }

                        }
                    }
                }
            }

            return bOK;
        }

        public void RemoveFile(int FolderID)
        {
            using (var DA = new SqlDataAdapter("SELECT FileID, FileName FROM Files WHERE FolderID = " + FolderID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    foreach (DataRow Row in DT.Rows)
                    {
                        RemoveFile(Convert.ToInt32(Row["FileID"]), FolderID, Row["FileName"].ToString());
                    }
                }
            }
        }

        public bool RemoveFolder(int FolderID)
        {
            var Path = "";

            using (var DA = new SqlDataAdapter("SELECT * FROM Folders WHERE ParentFolderID IN (SELECT ParentFolderID FROM Folders WHERE FolderID = " + FolderID + ")", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    var path = DT.Select("FolderID = " + FolderID)[0]["FolderPath"].ToString();

                    for (var i = DT.Rows.Count - 1; i >= 0; i--)
                    {
                        if (DT.Rows[i]["FolderPath"].ToString().Contains(path))
                        {
                            RemoveFile(Convert.ToInt32(DT.Rows[i]["FolderID"]));

                            Path = DT.Rows[i]["FolderPath"].ToString();

                            SetLastModified(Convert.ToInt32(DT.Rows[i]["RootFolderID"]), Security.GetCurrentDate(), Security.CurrentUserID);

                            var sPath = "";
                            var FTPType = -1;

                            if (IsPathHost(Convert.ToInt32(DT.Rows[i]["FolderID"])))
                            {
                                sPath = Configs.DocumentsPathHost;
                                FTPType = 1;
                            }
                            else
                            {
                                sPath = Configs.DocumentsPath;
                                FTPType = Configs.FTPType;
                            }

                            if (FM.DeleteFolder(sPath + Path, FTPType) == false)
                                return false;

                            using (var pDA = new SqlDataAdapter("DELETE FROM DocumentsPermissions WHERE FolderID = " + Convert.ToInt32(DT.Rows[i]["FolderID"]), ConnectionStrings.LightConnectionString))
                            {
                                using (var pDT = new DataTable())
                                {
                                    pDA.Fill(pDT);
                                }
                            }

                            using (var dDA = new SqlDataAdapter("DELETE FROM Folders WHERE FolderID = " + DT.Rows[i]["FolderID"], ConnectionStrings.LightConnectionString))
                            {
                                using (var dDT = new DataTable())
                                {
                                    dDA.Fill(dDT);
                                }
                            }
                        }
                    }
                }
            }

            return true;
        }

        public void RemoveFolder(DataTable CheckedDataTable)
        {
            foreach (var Row in CheckedDataTable.Select("Checked = 1 AND Extension = 'folder'"))
            {
                RemoveFolder(Convert.ToInt32(Row["FolderID"]));
            }
        }

        public void RemoveFile(DataTable CheckedDataTable)
        {
            foreach (var Row in CheckedDataTable.Select("Checked = 1 AND FileID > -1"))
            {
                RemoveFile(Convert.ToInt32(Row["FileID"]), Convert.ToInt32(Row["FolderID"]), Row["ItemName"].ToString());
            }
        }

        public void RemoveFile(int FileID, int FolderID, string FileName)
        {
            var Path = "";



            using (var DA = new SqlDataAdapter("DELETE FROM Files WHERE FileID = " + FileID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }

            using (var DA = new SqlDataAdapter("SELECT FolderID, FolderPath, LastModifiedDateTime, LastModifiedUserID FROM Folders WHERE FolderID = " + FolderID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        Path = DT.Rows[0]["FolderPath"].ToString();

                        SetLastModified(Convert.ToInt32(DT.Rows[0]["FolderID"]), Security.GetCurrentDate(), Security.CurrentUserID);

                        DA.Update(DT);
                    }
                }
            }

            ClearFileLog(FileID);


            using (var DA = new SqlDataAdapter("DELETE FROM DocumentsSigns WHERE FileID = " + FileID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }

            using (var DA = new SqlDataAdapter("DELETE FROM SubscribesRecords WHERE SubscribesItemID = 11 AND TableItemID = " + FileID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }

            var sPath = "";
            var FTPType = -1;

            if (IsPathHost(FolderID))
            {
                sPath = Configs.DocumentsPathHost;
                FTPType = 1;
            }
            else
            {
                sPath = Configs.DocumentsPath;
                FTPType = Configs.FTPType;
            }

            FM.DeleteFile(sPath + Path + "/" + FileName, FTPType);
        }

        public void RefreshUploadPending(int FileID, int UserID)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM FilesUploadsPending WHERE FileID = " + FileID + " AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        DT.Rows[0]["DateTime"] = Security.GetCurrentDate();

                        DA.Update(DT);
                    }
                }
            }
        }

        public bool ReplaceFile(int FileID, string FileName)
        {
            AddUploadPending(FileID, Security.CurrentUserID);

            FileInfo fi;

            //get file size
            try
            {
                fi = new FileInfo(FileName);
            }
            catch
            {
                return false;
            }

            var iFileSize = fi.Length;
            var FolderID = -1;

            using (var DA = new SqlDataAdapter("SELECT * FROM Files WHERE FileID = " + FileID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        FolderID = Convert.ToInt32(DT.Rows[0]["FolderID"]);

                        var sPath = "";
                        var FTPType = -1;

                        if (IsPathHost(Convert.ToInt32(DT.Rows[0]["FolderID"])))
                        {
                            sPath = Configs.DocumentsPathHost;
                            FTPType = 1;
                        }
                        else
                        {
                            sPath = Configs.DocumentsPath;
                            FTPType = Configs.FTPType;
                        }

                        //load file to ftp
                        if (FM.UploadFile(FileName, sPath +
                                GetFolderPath(Convert.ToInt32(DT.Rows[0]["FolderID"])) + "/" +
                                Path.GetFileNameWithoutExtension(FileName) + "_Ver" + Path.GetExtension(FileName), FTPType) == false)
                        {
                            FM.DeleteFile(sPath +
                                GetFolderPath(Convert.ToInt32(DT.Rows[0]["FolderID"])) + "/" +
                                Path.GetFileNameWithoutExtension(FileName) + "_Ver" + Path.GetExtension(FileName), FTPType);

                            return false;
                        }

                        //delete previous version
                        FM.DeleteFile(Configs.DocumentsPath +
                                GetFolderPath(Convert.ToInt32(DT.Rows[0]["FolderID"])) + "/" +
                                Path.GetFileName(FileName), Configs.FTPType);

                        //rename temp new version
                        FM.RenameFile(Configs.DocumentsPath +
                                GetFolderPath(Convert.ToInt32(DT.Rows[0]["FolderID"])) + "/" +
                                Path.GetFileNameWithoutExtension(FileName) + "_Ver" + Path.GetExtension(FileName),
                                Path.GetFileName(FileName), Configs.FTPType);


                        DT.Rows[0]["FileName"] = Path.GetFileName(FileName);
                        DT.Rows[0]["FileSize"] = iFileSize;

                        SetLastModified(FolderID, Security.GetCurrentDate(), Security.CurrentUserID);

                        DA.Update(DT);
                    }
                }
            }



            ClearUploadPending(FileID, Security.CurrentUserID);

            return true;
        }

        private string GetClientEmail(int ClientID)
        {
            using (var DA = new SqlDataAdapter("SELECT Email FROM Clients WHERE ClientID = " + ClientID, ConnectionStrings.MarketingReferenceConnectionString))
            {
                using (var DT = new DataTable())
                {
                    if (DA.Fill(DT) > 0)
                    {
                        return DT.Rows[0]["Email"].ToString();
                    }
                }
            }

            return null;
        }

        public int GetCurrentClientFolder(int FolderID)
        {
            using (var DA = new SqlDataAdapter("SELECT RootFolderID, FolderName FROM Folders WHERE FolderID = " + FolderID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    if (Convert.ToInt32(DT.Rows[0]["RootFolderID"]) != 4)
                    {
                        using (var dDA = new SqlDataAdapter("SELECT RootFolderID, FolderName FROM Folders WHERE FolderID = " + DT.Rows[0]["RootFolderID"], ConnectionStrings.LightConnectionString))
                        {
                            using (var dDT = new DataTable())
                            {
                                dDA.Fill(dDT);

                                using (var fDA = new SqlDataAdapter("SELECT ClientID, ClientName FROM Clients WHERE ClientFolderName = '" + dDT.Rows[0]["FolderName"] + "'", ConnectionStrings.MarketingReferenceConnectionString))
                                {
                                    using (var fDT = new DataTable())
                                    {
                                        fDA.Fill(fDT);

                                        return Convert.ToInt32(fDT.Rows[0]["ClientID"]);
                                    }
                                }
                            }
                        }
                    }

                    using (var fDA = new SqlDataAdapter("SELECT ClientID, ClientName FROM Clients WHERE ClientFolderName = '" + DT.Rows[0]["FolderName"] + "'", ConnectionStrings.MarketingReferenceConnectionString))
                    {
                        using (var fDT = new DataTable())
                        {
                            fDA.Fill(fDT);

                            return Convert.ToInt32(fDT.Rows[0]["ClientID"]);
                        }
                    }
                }
            }
        }

        public string SendEmailNotifyClient(int ClientID, string[] FileNames)
        {
            //string AccountPassword = "1290qpalzm";
            //string SenderEmail = "zovprofilreport@mail.ru";

            //string AccountPassword = "7026Gradus0462";
            //var AccountPassword = "foqwsulbjiuslnue";
            var AccountPassword = "lfbeecgxvmwvzlna";
            var SenderEmail = "infiniumdevelopers@gmail.com";

            var to = GetClientEmail(ClientID);

            if (to == null)
                return "-1";

            var from = SenderEmail;


            using (var message = new MailMessage(from, to + ", andrewromanchuk@mail.ru"))
            {
                message.Subject = "Infinium. Новые документы";
                message.Body = "Здравствуйте. Вам отправлены новые документы в программе Infinium. Agent.\r\n" +
                               "Если эта программа у Вас не установлена, обратитесь в отдел маркетинга для получения ссылки на скачивание программы или напишите нам на этот адрес\r\n" +
                               "Отправленные файлы также прикреплены к этому письму\r\n\nПисьмо сгенерировано автоматически, не надо отвечать на этот адрес. По всем вопросам обращайтесь на marketing.zovprofil@gmail.com";
                //SmtpClient client = new SmtpClient("smtp.mail.ru")
                var client = new SmtpClient("smtp.gmail.com", 587)
                {
                    UseDefaultCredentials = false,
                    Credentials = new NetworkCredential(SenderEmail, AccountPassword)
                };
                if (FileNames != null)
                    if (FileNames.Length > 0)
                        foreach (var name in FileNames)
                        {
                            var S = new UTF8Encoding().GetString(Encoding.Convert(Encoding.GetEncoding("UTF-16"), Encoding.UTF8, Encoding.GetEncoding("UTF-16").GetBytes(name)));

                            var attach = new Attachment(S,
                                                               MediaTypeNames.Application.Octet);
                            var disposition = attach.ContentDisposition;

                            message.Attachments.Add(attach);
                        }


                try
                {
                    client.Send(message);
                }

                catch (Exception e)
                {
                    return e.Message;
                }

                client.Dispose();
            }

            return "1";
        }



        public bool UploadFile(string[] FileNames, int FolderID, ref int CurrentUploadedFile)
        {
            foreach (var FileName in FileNames)
            {
                FileInfo fi;

                CurrentUploadedFile++;

                //get file size
                try
                {
                    fi = new FileInfo(FileName);
                }
                catch
                {
                    return false;
                }

                var iFileSize = fi.Length;

                var sPath = "";
                var FTPType = -1;

                if (IsPathHost(FolderID))
                {
                    sPath = Configs.DocumentsPathHost;
                    FTPType = 1;
                }
                else
                {
                    sPath = Configs.DocumentsPath;
                    FTPType = Configs.FTPType;
                }

                //load file to ftp
                if (FM.UploadFile(FileName, sPath +
                       GetFolderPath(FolderID) + "/" + Path.GetFileName(FileName), FTPType) == false)
                {
                    FM.DeleteFile(sPath +
                       GetFolderPath(FolderID) + "/" + Path.GetFileName(FileName), FTPType);

                    return false;
                }


                //add file to database
                using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM Files", ConnectionStrings.LightConnectionString))
                {
                    using (var CB = new SqlCommandBuilder(DA))
                    {
                        using (var DT = new DataTable())
                        {
                            DA.Fill(DT);

                            var NewRow = DT.NewRow();
                            NewRow["FileName"] = Path.GetFileName(FileName);
                            NewRow["FolderID"] = FolderID;
                            if (Path.GetExtension(FileName).Length > 0)
                                NewRow["FileExtension"] = Path.GetExtension(FileName).Substring(1, Path.GetExtension(FileName).Length - 1);
                            else
                                NewRow["FileExtension"] = "";
                            NewRow["FileSize"] = iFileSize;
                            NewRow["Author"] = Security.CurrentUserID;

                            var Date = Security.GetCurrentDate();

                            NewRow["CreationDateTime"] = Date;
                            NewRow["LastModifiedDateTime"] = Date;
                            NewRow["LastModifiedUserID"] = Security.CurrentUserID;
                            DT.Rows.Add(NewRow);

                            SetLastModified(FolderID, Date, Convert.ToInt32(NewRow["Author"]));

                            DA.Update(DT);
                        }
                    }
                }
            }

            return true;
        }

        public bool SaveFile(int FileID, string SaveToPath)
        {
            var bOk = false;

            using (var fDA = new SqlDataAdapter("SELECT FileSize, FileName, FolderID FROM Files WHERE FileID = " + FileID, ConnectionStrings.LightConnectionString))
            {
                using (var fDT = new DataTable())
                {
                    fDA.Fill(fDT);

                    using (var DA = new SqlDataAdapter("SELECT FolderPath FROM Folders WHERE FolderID = " + fDT.Rows[0]["FolderID"], ConnectionStrings.LightConnectionString))
                    {
                        using (var DT = new DataTable())
                        {
                            DA.Fill(DT);

                            var sPath = "";
                            var FTPType = -1;

                            if (IsPathHost(Convert.ToInt32(fDT.Rows[0]["FolderID"])))
                            {
                                sPath = Configs.DocumentsPathHost;
                                FTPType = 1;
                            }
                            else
                            {
                                sPath = Configs.DocumentsPath;
                                FTPType = Configs.FTPType;
                            }

                            bOk = FM.DownloadFile(sPath + DT.Rows[0]["FolderPath"] + "/" + fDT.Rows[0]["FileName"],
                                            SaveToPath, Convert.ToInt64(fDT.Rows[0]["FileSize"]), FTPType);

                            WriteDownloadLog(FileID, Security.CurrentUserID);
                        }
                    }
                }
            }


            return bOk;
        }

        public void SaveFiles(DataTable ItemsDataTable, string Path, ref int CurrentFile)
        {
            foreach (var Row in ItemsDataTable.Select("Checked = 1 AND FileID > -1"))
            {
                CurrentFile++;

                SaveFile(Convert.ToInt32(Row["FileID"]), Path + Row["ItemName"]);
            }
        }

        public void SetLastModified(int FolderID, DateTime Date, int UserID)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM Folders WHERE FolderID = " + FolderID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);


                        var iR = -1;
                        var iF = Convert.ToInt32(DT.Rows[0]["RootFolderID"]);

                        if (iF != 0)
                        {

                            for (var i = 0; i < 100; i++)
                            {
                                iR = SetRootFolderID(iF, Date, UserID);

                                if (iR != 0)
                                {
                                    iF = iR;
                                    continue;
                                }
                                //SetRootFolderID(iR, Date, UserID);

                                break;
                            }
                        }

                        DT.Rows[0]["LastModifiedDateTime"] = Date;
                        DT.Rows[0]["LastModifiedUserID"] = UserID;

                        DA.Update(DT);
                    }
                }
            }
        }

        public void SetPermissions(DateTime DateTime)
        {
            using (var DA = new SqlDataAdapter("SELECT FolderID, RootFolderID FROM Folders WHERE CreationDateTime = '" + DateTime.ToString("yyyy-MM-dd HH:mm:ss.fff") + "'", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    using (var pDA = new SqlDataAdapter("SELECT * FROM DocumentsPermissions WHERE FolderID = " + DT.Rows[0]["RootFolderID"], ConnectionStrings.LightConnectionString))
                    {
                        using (var pCB = new SqlCommandBuilder(pDA))
                        {
                            using (var pDT = new DataTable())
                            {
                                pDA.Fill(pDT);

                                var NewRow = pDT.NewRow();
                                NewRow["FolderID"] = DT.Rows[0]["FolderID"];
                                NewRow["UserID"] = pDT.Rows[0]["UserID"];
                                NewRow["UserTypeID"] = pDT.Rows[0]["UserTypeID"];
                                pDT.Rows.Add(NewRow);

                                pDA.Update(pDT);
                            }
                        }
                    }
                }
            }
        }

        public void SetPermissionsPersonalFolder(DateTime DateTime, int UserID)
        {
            using (var DA = new SqlDataAdapter("SELECT FolderID FROM Folders WHERE CreationDateTime = '" + DateTime.ToString("yyyy-MM-dd HH:mm:ss.fff") + "'", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    using (var pDA = new SqlDataAdapter("SELECT TOP 0 * FROM DocumentsPermissions", ConnectionStrings.LightConnectionString))
                    {
                        using (var pCB = new SqlCommandBuilder(pDA))
                        {
                            using (var pDT = new DataTable())
                            {
                                pDA.Fill(pDT);

                                var NewRow = pDT.NewRow();
                                NewRow["FolderID"] = DT.Rows[0]["FolderID"];
                                NewRow["UserID"] = UserID;
                                NewRow["UserTypeID"] = 0;
                                pDT.Rows.Add(NewRow);

                                pDA.Update(pDT);
                            }
                        }
                    }
                }
            }
        }

        public int SetRootFolderID(int FolderID, DateTime Date, int UserID)
        {
            using (var DA = new SqlDataAdapter("SELECT FolderID, RootFolderID, LastModifiedDateTime, LastModifiedUserID FROM Folders WHERE FolderID = " + FolderID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        DT.Rows[0]["LastModifiedDateTime"] = Date;
                        DT.Rows[0]["LastModifiedUserID"] = UserID;

                        DA.Update(DT);

                        return Convert.ToInt32(DT.Rows[0]["RootFolderID"]);
                    }
                }
            }
        }


        public void SetAttributes(string FileName, int FolderID, bool bFirstSign)
        {
            using (var DA = new SqlDataAdapter("SELECT FileID, NeedPrint, FirstSign FROM Files WHERE FileName = '" + FileName + "' AND FolderID = " + FolderID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        if (Convert.ToBoolean(DocumentAttributesDataTable.Select("AttributeName = 'Бумага'")[0]["Value"]))
                            DT.Rows[0]["NeedPrint"] = true;
                        else
                            DT.Rows[0]["NeedPrint"] = false;

                        DT.Rows[0]["FirstSign"] = bFirstSign;

                        //signs and readlist
                        if (CurrentSignsDataTable.Rows.Count > 0 || CurrentReadDataTable.Rows.Count > 0)
                        {
                            using (var sDA = new SqlDataAdapter("SELECT TOP 0 * FROM DocumentsSigns", ConnectionStrings.LightConnectionString))
                            {
                                using (var sCB = new SqlCommandBuilder(sDA))
                                {
                                    using (var sDT = new DataTable())
                                    {
                                        sDA.Fill(sDT);

                                        foreach (DataRow Row in CurrentSignsDataTable.Rows)
                                        {
                                            var NewRow = sDT.NewRow();
                                            NewRow["FileID"] = DT.Rows[0]["FileID"];
                                            NewRow["UserID"] = Row["UserID"];
                                            NewRow["SignDescription"] = "";
                                            NewRow["IsSigned"] = false;
                                            NewRow["SignType"] = 0;
                                            sDT.Rows.Add(NewRow);
                                        }

                                        foreach (DataRow Row in CurrentReadDataTable.Rows)
                                        {
                                            var NewRow = sDT.NewRow();
                                            NewRow["FileID"] = DT.Rows[0]["FileID"];
                                            NewRow["UserID"] = Row["UserID"];
                                            NewRow["SignDescription"] = "";
                                            NewRow["IsSigned"] = false;
                                            NewRow["SignType"] = 1;
                                            sDT.Rows.Add(NewRow);
                                        }

                                        SetSubscribesForSigns(sDT);

                                        sDA.Update(sDT);
                                    }
                                }
                            }
                        }

                        DA.Update(DT);
                    }
                }
            }

        }

        public void SignFile(int FileID, int UserID)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsSigns WHERE SignType = 0 AND FileID = " + FileID + " AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        DT.Rows[0]["IsSigned"] = true;
                        DT.Rows[0]["DateTime"] = Security.GetCurrentDate();

                        DA.Update(DT);

                        using (var fDA = new SqlDataAdapter("SELECT FileID, FirstSign FROM Files WHERE FileID = " + FileID, ConnectionStrings.LightConnectionString))
                        {
                            using (var fDT = new DataTable())
                            {
                                fDA.Fill(fDT);

                                if (Convert.ToBoolean(fDT.Rows[0]["FirstSign"]))
                                {
                                    using (var sDA = new SqlDataAdapter("DELETE FROM DocumentsSigns WHERE SignType = 0 AND FileID = " + FileID + " AND UserID <> " + UserID, ConnectionStrings.LightConnectionString))
                                    {
                                        using (var sDT = new DataTable())
                                        {
                                            sDA.Fill(sDT);
                                        }
                                    }


                                    using (var sDA = new SqlDataAdapter("DELETE FROM SubscribesRecords WHERE ModuleID = 147 AND TableItemID = " + FileID, ConnectionStrings.LightConnectionString))
                                    {
                                        using (var sDT = new DataTable())
                                        {
                                            sDA.Fill(sDT);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        public void SignReadFile(int FileID, int UserID)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsSigns WHERE SignType = 1 AND FileID = " + FileID + " AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        DT.Rows[0]["IsSigned"] = true;
                        DT.Rows[0]["DateTime"] = Security.GetCurrentDate();

                        DA.Update(DT);
                    }
                }
            }
        }

        public void SetSubscribesForSigns(DataTable CurrentDT)
        {
            using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM SubscribesRecords", ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        foreach (DataRow Row in CurrentDT.Rows)
                        {
                            var NewRow = DT.NewRow();
                            NewRow["UserID"] = Row["UserID"];
                            NewRow["TableItemID"] = Row["FileID"];
                            NewRow["SubscribesItemID"] = 11;
                            NewRow["UserTypeID"] = 0;
                            DT.Rows.Add(NewRow);
                        }

                        DA.Update(DT);
                    }
                }
            }
        }

        public void WriteDownloadLog(int FileID, int UserID)
        {
            using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM FilesDownloadsLog", ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        var NewRow = DT.NewRow();
                        NewRow["UserID"] = UserID;
                        NewRow["FileID"] = FileID;
                        NewRow["DateTime"] = Security.GetCurrentDate();
                        DT.Rows.Add(NewRow);

                        DA.Update(DT);
                    }
                }
            }
        }

    }





    public class ZOVNews
    {
        public DataTable UsersDataTable;
        public DataTable NewsDataTable;
        public DataTable NewsLikesDataTable;
        public DataTable AttachmentsDataTable;
        public DataTable CommentsDataTable;
        public DataTable CommentsLikesDataTable;
        public DataTable CommentsSubsRecordsDataTable;
        public DataTable NewsSubsRecordsDataTable;
        public DataTable CurrentCommentsDataTable;
        public DataTable ManagersDataTable;

        public FileManager FM = new FileManager();

        public ZOVNews()
        {
            NewsDataTable = new DataTable();
            NewsLikesDataTable = new DataTable();
            CommentsLikesDataTable = new DataTable();
            UsersDataTable = new DataTable();

            Fill();
            FillNews();
            FillLikes();
        }

        public void Fill()
        {
            UsersDataTable.Clear();

            UsersDataTable = new DataView(TablesManager.UsersDataTable).ToTable(true, "UserID", "Name", "Photo");

            if (UsersDataTable.Columns["SenderTypeID"] == null)
                UsersDataTable.Columns.Add(new DataColumn("SenderTypeID", Type.GetType("System.Int32")));

            foreach (DataRow Row in UsersDataTable.Rows)
                Row["SenderTypeID"] = 0;

            AttachmentsDataTable = new DataTable();

            using (var DA = new SqlDataAdapter("SELECT NewsAttachID, NewsID, FileName, FileSize FROM ZOVNewsAttachs", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(AttachmentsDataTable);
            }

            ManagersDataTable = new DataTable();

            using (var DA = new SqlDataAdapter("SELECT * FROM Managers", ConnectionStrings.ZOVReferenceConnectionString))
            {
                DA.Fill(ManagersDataTable);
            }

            CommentsDataTable = new DataTable();

            using (var DA = new SqlDataAdapter("SELECT * FROM ZOVNewsComments", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(CommentsDataTable);
            }

            CommentsSubsRecordsDataTable = new DataTable();

            using (var DA = new SqlDataAdapter("SELECT SubscribesRecords.*, ZOVNewsComments.NewsID FROM SubscribesRecords INNER JOIN ZOVNewsComments ON ZOVNewsComments.NewsCommentID = SubscribesRecords.TableItemID WHERE SubscribesItemID = 14 AND SubscribesRecords.UserTypeID = 0 AND SubscribesRecords.UserID = " + Security.CurrentUserID, ConnectionStrings.LightConnectionString))
            {
                DA.Fill(CommentsSubsRecordsDataTable);
            }

            NewsSubsRecordsDataTable = new DataTable();

            using (var sDA = new SqlDataAdapter("SELECT * FROM SubscribesRecords WHERE SubscribesItemID = 13 " +
                                                " AND UserTypeID = 0 AND UserID = " + Security.CurrentUserID, ConnectionStrings.LightConnectionString))
            {
                sDA.Fill(NewsSubsRecordsDataTable);
            }
        }

        public void ReloadNews(int Count)
        {
            NewsDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT TOP " + Count + " * FROM ZOVNews WHERE Pending <> 1 ORDER BY LastCommentDateTime DESC", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(NewsDataTable);
            }

            //subscribes///////////////////////
            if (NewsDataTable.Columns["New"] == null)
                NewsDataTable.Columns.Add(new DataColumn("New", Type.GetType("System.Int32")));

            if (NewsDataTable.Columns["NewComments"] == null)
                NewsDataTable.Columns.Add(new DataColumn("NewComments", Type.GetType("System.Int32")));

            foreach (DataRow Row in NewsSubsRecordsDataTable.Rows)
            {
                if (NewsDataTable.Select("NewsID = " + Row["TableItemID"]).Count() > 0)
                {
                    NewsDataTable.Select("NewsID = " + Row["TableItemID"])[0]["New"] = 1;
                }
            }

            foreach (DataRow Row in CommentsSubsRecordsDataTable.Rows)
            {
                if (NewsDataTable.Select("NewsID = " + Row["NewsID"]).Count() > 0)
                {
                    NewsDataTable.Select("NewsID = " + Row["NewsID"])[0]["NewComments"] =
                            CommentsSubsRecordsDataTable.Select("NewsID = " + Row["NewsID"]).Count();
                }
            }
            //////////////////////////////////
        }

        public void ReloadComments()
        {
            CommentsDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT * FROM ZOVNewsComments", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(CommentsDataTable);
            }
        }

        public void ReloadSubscribes()
        {
            CommentsSubsRecordsDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT SubscribesRecords.*, ZOVNewsComments.NewsID FROM SubscribesRecords INNER JOIN ZOVNewsComments ON ZOVNewsComments.NewsCommentID = SubscribesRecords.TableItemID WHERE SubscribesItemID = 14 AND SubscribesRecords.UserTypeID = 0 AND SubscribesRecords.UserID = " + Security.CurrentUserID, ConnectionStrings.LightConnectionString))
            {
                DA.Fill(CommentsSubsRecordsDataTable);
            }

            NewsSubsRecordsDataTable.Clear();

            using (var sDA = new SqlDataAdapter("SELECT * FROM SubscribesRecords WHERE SubscribesItemID = 13 " +
                                                " AND UserTypeID = 0 AND UserID = " + Security.CurrentUserID, ConnectionStrings.LightConnectionString))
            {
                sDA.Fill(NewsSubsRecordsDataTable);
            }
        }

        public void FillNews()
        {
            NewsDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT TOP 20 * FROM ZOVNews WHERE Pending <> 1 ORDER BY LastCommentDateTime DESC", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(NewsDataTable);
            }

            if (NewsDataTable.Columns["New"] == null)
                NewsDataTable.Columns.Add(new DataColumn("New", Type.GetType("System.Int32")));

            if (NewsDataTable.Columns["NewComments"] == null)
                NewsDataTable.Columns.Add(new DataColumn("NewComments", Type.GetType("System.Int32")));

            foreach (DataRow Row in NewsSubsRecordsDataTable.Rows)
            {
                if (NewsDataTable.Select("NewsID = " + Row["TableItemID"]).Count() > 0)
                {
                    NewsDataTable.Select("NewsID = " + Row["TableItemID"])[0]["New"] = 1;
                }
            }

            foreach (DataRow Row in CommentsSubsRecordsDataTable.Rows)
            {
                if (NewsDataTable.Select("NewsID = " + Row["NewsID"]).Count() > 0)
                {
                    NewsDataTable.Select("NewsID = " + Row["NewsID"])[0]["NewComments"] =
                            CommentsSubsRecordsDataTable.Select("NewsID = " + Row["NewsID"]).Count();
                }
            }
        }

        public void FillLikes()
        {
            NewsLikesDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT * FROM ZOVNewsLikes", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(NewsLikesDataTable);
            }


            CommentsLikesDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT * FROM ZOVNewsCommentsLikes", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(CommentsLikesDataTable);
            }
        }

        public void FillMoreNews(int Count)
        {
            NewsDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT TOP " + Count + " * FROM ZOVNews WHERE Pending <> 1 ORDER BY LastCommentDateTime DESC", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(NewsDataTable);
            }
        }

        public bool IsMoreNews(int Count)
        {
            using (var DA = new SqlDataAdapter("SELECT Count(NewsID) FROM ZOVNews", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    return Convert.ToInt32(DT.Rows[0][0]) > Count;
                }
            }
        }

        public DateTime AddNews(int SenderID, int SenderTypeID, string HeaderText, string BodyText, int NewsCategoryID)
        {
            DateTime DateTime;

            using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM ZOVNews", ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        var NewRow = DT.NewRow();
                        NewRow["SenderID"] = SenderID;
                        NewRow["SenderTypeID"] = SenderTypeID;
                        NewRow["HeaderText"] = HeaderText;
                        NewRow["BodyText"] = BodyText;
                        NewRow["Pending"] = true;

                        DateTime = Security.GetCurrentDate();

                        NewRow["DateTime"] = DateTime;
                        NewRow["LastCommentDateTime"] = DateTime;
                        DT.Rows.Add(NewRow);

                        DA.Update(DT);

                        FillNews();
                    }
                }
            }

            return DateTime;
        }

        public void AddSubscribeForNews(DateTime DateTime)
        {
            using (var sDA = new SqlDataAdapter("SELECT * FROM ZOVNews WHERE DateTime = '" + DateTime.ToString("yyyy-MM-dd HH:mm:ss.fff") + "'",
                                                                       ConnectionStrings.LightConnectionString))
            {
                using (var sDT = new DataTable())
                {
                    sDA.Fill(sDT);

                    using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM SubscribesRecords", ConnectionStrings.LightConnectionString))
                    {
                        using (var CB = new SqlCommandBuilder(DA))
                        {
                            using (var DT = new DataTable())
                            {
                                DA.Fill(DT);

                                using (var uDA = new SqlDataAdapter("SELECT * FROM ModulesAccess WHERE ModuleID = 156", ConnectionStrings.UsersConnectionString))
                                {
                                    using (var uDT = new DataTable())
                                    {
                                        uDA.Fill(uDT);

                                        foreach (DataRow Row in uDT.Rows)
                                        {
                                            if (Convert.ToInt32(Row["UserID"]) == Security.CurrentUserID)
                                                continue;

                                            var NewRow = DT.NewRow();
                                            NewRow["SubscribesItemID"] = 13;
                                            NewRow["TableItemID"] = sDT.Rows[0]["NewsID"];
                                            NewRow["UserID"] = Row["UserID"];
                                            NewRow["UserTypeID"] = 0;
                                            DT.Rows.Add(NewRow);
                                        }

                                        foreach (DataRow Row in ManagersDataTable.Rows)
                                        {
                                            var NewRow = DT.NewRow();
                                            NewRow["SubscribesItemID"] = 13;
                                            NewRow["TableItemID"] = sDT.Rows[0]["NewsID"];
                                            NewRow["UserID"] = Row["ManagerID"];
                                            NewRow["UserTypeID"] = 1;
                                            DT.Rows.Add(NewRow);
                                        }

                                        DA.Update(DT);
                                    }
                                }
                            }
                        }
                    }


                }
            }
        }

        public void ClearPending(DateTime DateTime)
        {
            using (var sDA = new SqlDataAdapter("SELECT TOP 1 * FROM ZOVNews WHERE DateTime = '" + DateTime.ToString("yyyy-MM-dd HH:mm:ss.fff") + "'",
                                                                      ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(sDA))
                {
                    using (var sDT = new DataTable())
                    {
                        sDA.Fill(sDT);

                        sDT.Rows[0]["Pending"] = false;

                        sDA.Update(sDT);
                    }

                }

            }
        }

        public void AddComments(int UserID, int UserTypeID, int NewsID, string Text)
        {
            DateTime Date;

            using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM ZOVNewsComments", ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        var NewRow = DT.NewRow();
                        NewRow["NewsID"] = NewsID;
                        NewRow["NewsComment"] = Text;
                        NewRow["UserID"] = UserID;
                        NewRow["UserTypeID"] = UserTypeID;

                        Date = Security.GetCurrentDate();

                        NewRow["DateTime"] = Date;
                        DT.Rows.Add(NewRow);

                        DA.Update(DT);
                    }
                }
            }

            using (var DA = new SqlDataAdapter("SELECT NewsID, LastCommentDateTime FROM ZOVNews WHERE NewsID =" + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        DT.Rows[0]["LastCommentDateTime"] = Date;

                        DA.Update(DT);
                    }
                }
            }

            AddNewsCommentsSubs(NewsID, UserID, Date);
        }


        public void LikeNews(int UserID, int NewsID)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM ZOVNewsLikes WHERE NewsID = " + NewsID + " AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        if (DT.Rows.Count > 0)//i like
                        {
                            using (var dDA = new SqlDataAdapter("DELETE FROM ZOVNewsLikes WHERE NewsID = " + NewsID + " AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
                            {
                                using (var dDT = new DataTable())
                                {
                                    dDA.Fill(dDT);
                                }
                            }

                            return;
                        }

                        var NewRow = DT.NewRow();
                        NewRow["NewsID"] = NewsID;
                        NewRow["UserID"] = UserID;
                        DT.Rows.Add(NewRow);

                        DA.Update(DT);
                    }
                }
            }
        }

        public void LikeComments(int UserID, int NewsID, int NewsCommentID)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM ZOVNewsCommentsLikes WHERE NewsCommentID = " + NewsCommentID + " AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        if (DT.Rows.Count > 0)
                        {
                            using (var dDA = new SqlDataAdapter("DELETE FROM ZOVNewsCommentsLikes WHERE NewsCommentID = " + NewsCommentID + " AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
                            {
                                using (var dDT = new DataTable())
                                {
                                    dDA.Fill(dDT);
                                }
                            }

                            return;
                        }

                        var NewRow = DT.NewRow();
                        NewRow["NewsID"] = NewsID;
                        NewRow["NewsCommentID"] = NewsCommentID;
                        NewRow["UserID"] = UserID;
                        DT.Rows.Add(NewRow);

                        DA.Update(DT);
                    }
                }
            }
        }


        public void EditComments(int UserID, int NewsCommentID, string Text)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM ZOVNewsComments WHERE NewsCommentID = " + NewsCommentID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        DT.Rows[0]["NewsComment"] = Text;

                        DA.Update(DT);
                    }
                }
            }
        }


        public bool Attach(DataTable AttachmentsDataTable, DateTime NewsDateTime, ref int CurrentUploadedFile)
        {
            if (AttachmentsDataTable.Rows.Count == 0)
                return true;

            var Ok = true;

            int NewsID;

            using (var DA = new SqlDataAdapter("SELECT NewsID FROM ZOVNews WHERE DateTime = '" + NewsDateTime.ToString("yyyy-MM-dd HH:mm:ss.fff") + "'", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    NewsID = Convert.ToInt32(DT.Rows[0]["NewsID"]);
                }
            }

            using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM ZOVNewsAttachs", ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        foreach (DataRow Row in AttachmentsDataTable.Rows)
                        {
                            FileInfo fi;

                            try
                            {
                                fi = new FileInfo(Row["Path"].ToString());
                                //WriteLog("UploadFile fileinfosize: " + Row["Path"].ToString(), Security.CurrentUserID, -1);
                            }
                            catch
                            {
                                Ok = false;
                                continue;
                            }

                            var NewRow = DT.NewRow();
                            NewRow["NewsID"] = NewsID;
                            NewRow["FileName"] = Row["FileName"];
                            NewRow["FileSize"] = fi.Length;
                            DT.Rows.Add(NewRow);
                        }

                        DA.Update(DT);
                    }
                }
            }

            //write to ftp
            using (var DA = new SqlDataAdapter("SELECT * FROM ZOVNewsAttachs WHERE NewsID = " + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    foreach (DataRow Row in AttachmentsDataTable.Rows)
                    {
                        CurrentUploadedFile++;

                        try
                        {
                            if (FM.UploadFile(Row["Path"].ToString(), Configs.DocumentsPath +
                                          FileManager.GetPath("ZOVNews") + "/" +
                                          DT.Select("FileName = '" + Row["FileName"] + "'")[0]["NewsAttachID"] + ".idf", Configs.FTPType) == false)
                            {
                                break;
                            }
                        }
                        catch
                        {
                            Ok = false;
                            break;
                        }
                    }

                    DA.Update(DT);
                }
            }

            ReloadAttachments();

            return Ok;
        }

        public bool EditAttachments(int NewsID, DataTable AttachmentsDataTable, ref int CurrentUploadedFile, ref int TotalFilesCount)
        {
            var Ok = false;

            using (var DA = new SqlDataAdapter("SELECT NewsAttachID, NewsID, FileName FROM ZOVNewsAttachs WHERE NewsID = " + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        foreach (DataRow Row in DT.Rows)
                        {
                            if (AttachmentsDataTable.Select("FileName = '" + Row["FileName"] + "'").Count() == 0)
                            {
                                FM.DeleteFile(Configs.DocumentsPath + FileManager.GetPath("ZOVNews") + "/" +
                                              DT.Select("FileName = '" + Row["FileName"] + "'")[0]["NewsAttachID"] + ".idf", Configs.FTPType);
                                Row.Delete();
                            }
                        }

                        DA.Update(DT);
                    }
                }
            }

            using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM ZOVNewsAttachs WHERE NewsID = " + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);
                        Ok = true;

                        //add new
                        foreach (DataRow Row in AttachmentsDataTable.Rows)
                        {
                            if (Row["Path"].ToString() == "server")
                                continue;

                            TotalFilesCount++;

                            FileInfo fi;

                            try
                            {
                                fi = new FileInfo(Row["Path"].ToString());

                            }
                            catch
                            {
                                Ok = false;
                                continue;
                            }

                            var NewRow = DT.NewRow();
                            NewRow["NewsID"] = NewsID;
                            NewRow["FileName"] = Row["FileName"];
                            NewRow["FileSize"] = fi.Length;
                            DT.Rows.Add(NewRow);
                        }

                        DA.Update(DT);
                    }
                }
            }


            //write to ftp
            using (var DA = new SqlDataAdapter("SELECT * FROM ZOVNewsAttachs WHERE NewsID = " + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    foreach (DataRow Row in AttachmentsDataTable.Rows)
                    {
                        if (Row["Path"].ToString() == "server")
                            continue;

                        CurrentUploadedFile++;

                        try
                        {
                            if (FM.UploadFile(Row["Path"].ToString(), Configs.DocumentsPath +
                                         FileManager.GetPath("ZOVNews") + "/" +
                                         DT.Select("FileName = '" + Row["FileName"] + "'")[0]["NewsAttachID"] + ".idf", Configs.FTPType) == false)
                                break;
                        }
                        catch
                        {
                            Ok = false;
                            break;
                        }
                    }

                    DA.Update(DT);
                }
            }


            ReloadAttachments();

            return Ok;
        }

        public void ReloadAttachments()
        {
            AttachmentsDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT NewsAttachID, NewsID, FileName, FileSize FROM ZOVNewsAttachs", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(AttachmentsDataTable);
            }
        }

        public void EditNews(int NewsID, int SenderID, int SenderTypeID, string HeaderText, string BodyText, int NewsCategoryID, DateTime DateTime)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM ZOVNews WHERE NewsID = " + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        if (DA.Fill(DT) == 0)
                            return;

                        DT.Rows[0]["HeaderText"] = HeaderText;
                        DT.Rows[0]["BodyText"] = BodyText;
                        DT.Rows[0]["DateTime"] = DateTime;

                        DA.Update(DT);

                        FillNews();
                    }
                }
            }
        }

        public static string TempPath()
        {
            return Environment.GetEnvironmentVariable("TEMP");
        }

        public int GetNewsUpdatesCount()
        {
            using (var DA = new SqlDataAdapter("SELECT Count (SubscribesRecordID) FROM SubscribesRecords WHERE SubscribesItemID = 13 OR SubscribesItemID = 14 AND UserTypeID = 0 AND UserID = " + Security.CurrentUserID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    if (DA.Fill(DT) > 0)
                        return Convert.ToInt32(DT.Rows[0][0]);
                }
            }

            return 0;
        }

        public int GetNewsIDByDateTime(DateTime DateTime)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM ZOVNews WHERE DateTime = '" + DateTime.ToString("yyyy-MM-dd HH:mm:ss.fff") + "'",
                                                                         ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    return Convert.ToInt32(DT.Rows[0]["NewsID"]);
                }
            }
        }

        public void RemoveNews(int NewsID)
        {
            //ActiveNotifySystem.DeleteSubscribesRecord(1);

            using (var DA = new SqlDataAdapter("DELETE FROM ZOVNewsLikes WHERE NewsID = " + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }

            using (var DA = new SqlDataAdapter("DELETE FROM ZOVNews WHERE NewsID = " + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);
                    }
                }
            }

            RemoveComments(NewsID);

            RemoveAttachments(NewsID);

            FillNews();

            ReloadAttachments();
        }

        public void RemoveComments(int NewsID)
        {
            DeleteCommentsSubs(NewsID);

            using (var DA = new SqlDataAdapter("DELETE FROM ZOVNewsCommentsLikes WHERE NewsID = " + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }

            using (var DA = new SqlDataAdapter("DELETE FROM ZOVNewsComments WHERE NewsID = " + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);
                    }
                }
            }


        }

        public void RemoveComment(int NewsCommentID)
        {
            using (var DA = new SqlDataAdapter("DELETE FROM ZOVNewsCommentsLikes WHERE NewsCommentID = " + NewsCommentID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }

            using (var DA = new SqlDataAdapter("DELETE FROM ZOVNewsComments WHERE NewsCommentID = " + NewsCommentID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }

            DeleteCommentsSub(NewsCommentID);
        }

        public void RemoveAttachments(int NewsID)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM ZOVNewsAttachs WHERE NewsID = " + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    try
                    {
                        foreach (DataRow Row in DT.Rows)
                        {
                            FM.DeleteFile(Configs.DocumentsPath + FileManager.GetPath("ZOVNews") + "/" +
                                          DT.Select("FileName = '" + Row["FileName"] + "'")[0]["NewsAttachID"] + ".idf", Configs.FTPType);
                        }
                    }
                    catch
                    {
                        using (var fDA = new SqlDataAdapter("DELETE FROM ZOVNewsAttachs WHERE NewsID = " + NewsID, ConnectionStrings.LightConnectionString))
                        {
                            using (var CB = new SqlCommandBuilder(fDA))
                            {
                                using (var fDT = new DataTable())
                                {
                                    fDA.Fill(fDT);
                                }
                            }
                        }

                        ReloadAttachments();

                        return;
                    }
                }
            }



            using (var DA = new SqlDataAdapter("DELETE FROM ZOVNewsAttachs WHERE NewsID = " + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);
                    }
                }
            }

            ReloadAttachments();
        }

        public void RemoveCurrentAttachments(int NewsID, DataTable AttachmentsDT)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM ZOVNewsAttachs WHERE NewsID = " + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        foreach (DataRow Row in AttachmentsDT.Rows)
                        {
                            if (Row["Path"].ToString() == "server")
                                continue;

                            FM.DeleteFile(Configs.DocumentsPath + FileManager.GetPath("ZOVNews") + "/" +
                                      DT.Select("FileName = '" + Row["FileName"] + "'")[0]["NewsAttachID"] + ".idf", Configs.FTPType);

                            DT.Select("FileName = '" + Row["FileName"] + "'")[0].Delete();
                        }

                        DA.Update(DT);
                    }
                }
            }

            ReloadAttachments();
        }

        public DataTable GetAttachments(int NewsID)
        {
            using (var DA = new SqlDataAdapter("SELECT NewsAttachID, FileName FROM ZOVNewsAttachs WHERE NewsID = " + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    return DT;
                }
            }
        }

        public int GetThisNewsSenderTypeID(int NewsID)
        {
            return Convert.ToInt32(NewsDataTable.Select("NewsID = " + NewsID)[0]["SenderTypeID"]);
        }

        public string GetThisNewsBodyText(int NewsID)
        {
            return NewsDataTable.Select("NewsID = " + NewsID)[0]["BodyText"].ToString();
        }

        public string GetThisNewsHeaderText(int NewsID)
        {
            return NewsDataTable.Select("NewsID = " + NewsID)[0]["HeaderText"].ToString();
        }

        public DateTime GetThisNewsDateTime(int NewsID)
        {
            return Convert.ToDateTime(NewsDataTable.Select("NewsID = " + NewsID)[0]["DateTime"]);
        }

        public int ShowAttachDownloadMenu(string FileName)
        {
            //0 cancel, 1 open, 2 save


            return AttachDownloadForm.Result;
        }

        public string GetAttachmentName(int NewsAttachID)
        {
            return AttachmentsDataTable.Select("NewsAttachID = " + NewsAttachID)[0]["FileName"].ToString();
        }

        public string SaveFile(int NewsAttachID)//temp folder
        {
            var tempFolder = Environment.GetEnvironmentVariable("TEMP");

            var FileName = "";

            using (var DA = new SqlDataAdapter("SELECT FileName, FileSize FROM ZOVNewsAttachs WHERE NewsAttachID = " + NewsAttachID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    FileName = DT.Rows[0]["FileName"].ToString();

                    try
                    {
                        FM.DownloadFile(Configs.DocumentsPath + FileManager.GetPath("ZOVNews") + "/" + NewsAttachID + ".idf",
                                            tempFolder + "\\" + FileName, Convert.ToInt64(DT.Rows[0]["FileSize"]), Configs.FTPType);
                    }
                    catch
                    {
                        return null;
                    }

                    //byte[] b = (byte[])DT.Rows[0]["FileBytes"];
                    //MemoryStream ms = new MemoryStream(b);

                    //FileName = DT.Rows[0]["FileName"].ToString();

                    //FileStream s2;

                    //try
                    //{
                    //    s2 = new FileStream(tempFolder + "\\" + FileName, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite);
                    //}
                    //catch 
                    //{
                    //    FileName = GetNewFileName(tempFolder, DT.Rows[0]["FileName"].ToString());
                    //    s2 = new FileStream(tempFolder + "\\" + FileName, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite);
                    //}

                    //s2.Write(ms.ToArray(), 0, ms.Capacity);

                    //s2.Close();
                }
            }

            return tempFolder + "\\" + FileName;
        }

        public void SaveFile(int NewsAttachID, string sDestFileName)
        {
            using (var DA = new SqlDataAdapter("SELECT FileName, FileSize FROM ZOVNewsAttachs WHERE NewsAttachID = " + NewsAttachID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    var FileName = DT.Rows[0]["FileName"].ToString();

                    try
                    {
                        FM.DownloadFile(Configs.DocumentsPath + FileManager.GetPath("ZOVNews") + "/" + NewsAttachID + ".idf",
                                           sDestFileName, Convert.ToInt64(DT.Rows[0]["FileSize"]), Configs.FTPType);
                    }
                    catch
                    {
                    }
                }
            }
        }

        private string SetNumber(string FileName, int Number)
        {
            var Ext = "";
            var Name = "";

            for (var i = FileName.Length - 1; i > 0; i--)
            {
                if (FileName[i] == '.')
                {
                    Ext = FileName.Substring(i + 1, FileName.Length - i - 1);
                    Name = FileName.Substring(0, i) + " (" + Number + ")." + Ext;
                    break;
                }
            }

            return Name;
        }

        private string GetNewFileName(string path, string FileName)
        {
            var fileInfo = new FileInfo(path + "\\" + FileName);

            if (!fileInfo.Exists)
                return FileName;

            var Ok = false;
            var n = 1;

            while (!Ok)
            {
                fileInfo = new FileInfo(path + "\\" + SetNumber(FileName, n));

                if (!fileInfo.Exists)
                    return SetNumber(FileName, n);

                n++;
            }

            return "";
        }

        public void DeleteCommentsSub(int NewsCommentID)
        {
            using (var DA = new SqlDataAdapter("DELETE FROM SubscribesRecords WHERE SubscribesItemID = 14 AND TableItemID = " + NewsCommentID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }
        }

        public void DeleteCommentsSubs(int NewsID)
        {
            using (var DA = new SqlDataAdapter("DELETE FROM SubscribesRecords WHERE SubscribesItemID = 14 AND TableItemID IN (SELECT NewsCommentID FROM ZOVNewsComments WHERE NewsID = " + NewsID + ")", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }
        }

        public void AddNewsCommentsSubs(int NewsID, int UserID, DateTime DateTime)
        {
            using (var sDA = new SqlDataAdapter("SELECT NewsCommentID FROM ZOVNewsComments WHERE DateTime = '" + DateTime.ToString("yyyy-MM-dd HH:mm:ss.fff") + "'", ConnectionStrings.LightConnectionString))
            {
                using (var sDT = new DataTable())
                {
                    sDA.Fill(sDT);

                    using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM SubscribesRecords", ConnectionStrings.LightConnectionString))
                    {
                        using (var CB = new SqlCommandBuilder(DA))
                        {
                            using (var DT = new DataTable())
                            {
                                DA.Fill(DT);

                                using (var uDA = new SqlDataAdapter("SELECT * FROM ModulesAccess WHERE ModuleID = 156", ConnectionStrings.UsersConnectionString))
                                {
                                    using (var uDT = new DataTable())
                                    {
                                        uDA.Fill(uDT);

                                        foreach (DataRow Row in uDT.Rows)
                                        {
                                            if (Convert.ToInt32(Row["UserID"]) == Security.CurrentUserID)
                                                continue;

                                            var NewRow = DT.NewRow();
                                            NewRow["SubscribesItemID"] = 14;
                                            NewRow["TableItemID"] = sDT.Rows[0]["NewsCommentID"];
                                            NewRow["UserID"] = Row["UserID"];
                                            NewRow["UserTypeID"] = 0;
                                            DT.Rows.Add(NewRow);
                                        }

                                        foreach (DataRow Row in ManagersDataTable.Rows)
                                        {
                                            var NewRow = DT.NewRow();
                                            NewRow["SubscribesItemID"] = 14;
                                            NewRow["TableItemID"] = sDT.Rows[0]["NewsCommentID"];
                                            NewRow["UserID"] = Row["ManagerID"];
                                            NewRow["UserTypeID"] = 1;
                                            DT.Rows.Add(NewRow);
                                        }

                                        DA.Update(DT);
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }





    public class ZOVMessages
    {
        public int CurrentUserID = Security.CurrentUserID;

        public DataTable ManagersDataTable;
        public BindingSource ManagersBindingSource;
        private SqlDataAdapter ManagersDA;
        private SqlCommandBuilder ManagersCB;

        public string CurrentUserName;

        public DataTable SelectedUsersDataTable;
        public BindingSource SelectedUsersBindingSource;
        public BindingSource UsersBindingSource;

        public DataTable MessagesDataTable, UsersDataTable;
        private readonly ClientsMessagesDataGrid SelectedUsersGrid;
        private readonly ClientsDataGrid UsersListDataGrid;

        public ZOVMessages(ref ClientsMessagesDataGrid tSelectedUsersGrid, ref ClientsDataGrid tUsersListDataGrid)
        {
            UsersDataTable = new DataTable();

            using (var DA = new SqlDataAdapter("SELECT UserID, Name FROM Users WHERE UserID IN (SELECT UserID FROM ModulesAccess WHERE ModuleID = 157) ORDER BY Name ASC", ConnectionStrings.UsersConnectionString))
            {
                DA.Fill(UsersDataTable);
            }

            CreateAndFill();
            Binding();

            ManagersDataTable.Columns.Add(new DataColumn("OnlineStatus", Type.GetType("System.Boolean")));

            CurrentUserName = Security.GetUserNameByID(CurrentUserID);

            MessagesDataTable = new DataTable();

            SelectedUsersDataTable = new DataTable();
            SelectedUsersDataTable = ManagersDataTable.Clone();
            SelectedUsersDataTable.Columns.Add(new DataColumn("UpdatesCount", Type.GetType("System.Int32")));

            SelectedUsersGrid = tSelectedUsersGrid;
            UsersListDataGrid = tUsersListDataGrid;

            SelectedUsersBindingSource = new BindingSource
            {
                DataSource = SelectedUsersDataTable
            };
            SelectedUsersGrid.DataSource = SelectedUsersBindingSource;

            UsersListDataGrid.DataSource = ManagersBindingSource;

            SetSelectedUsersGrid();
            SetUsersGrid();
        }

        private void CreateAndFill()
        {
            ManagersDataTable = new DataTable();
            ManagersDA = new SqlDataAdapter("SELECT ManagerID, Name FROM Managers", ConnectionStrings.ZOVReferenceConnectionString);
            ManagersCB = new SqlCommandBuilder(ManagersDA);
            ManagersDA.Fill(ManagersDataTable);
        }

        private void Binding()
        {
            ManagersBindingSource = new BindingSource { DataSource = ManagersDataTable };
        }

        public void SetSelectedUsersGrid()
        {
            SelectedUsersGrid.AddColumns();

            SelectedUsersGrid.Columns["ManagerID"].Visible = false;
            SelectedUsersGrid.Columns["UpdatesCount"].Visible = false;
            SelectedUsersGrid.Columns["OnlineStatus"].Visible = false;
            SelectedUsersGrid.Columns["CloseColumn"].DisplayIndex = 3;
            SelectedUsersGrid.Columns["OnlineColumn"].DisplayIndex = 0;
            SelectedUsersGrid.sNewMessagesColumnName = "UpdatesCount";
            SelectedUsersGrid.sOnlineStatusColumnName = "OnlineStatus";
        }

        public void SetUsersGrid()
        {
            UsersListDataGrid.AddColumns();
            UsersListDataGrid.sOnlineStatusColumnName = "OnlineStatus";
            UsersListDataGrid.Columns["OnlineColumn"].DisplayIndex = 0;
            UsersListDataGrid.Columns["ManagerID"].Visible = false;
            UsersListDataGrid.Columns["OnlineStatus"].Visible = false;
        }

        public void UpdateList()
        {
            ManagersDataTable.Clear();
            ManagersDA.Fill(ManagersDataTable);
        }

        public void AddUserToSelected(string ManagerID)
        {
            if (SelectedUsersDataTable.Select("ManagerID = " + ManagerID).Count() > 0)
            {
                SelectedUsersBindingSource.Position = SelectedUsersBindingSource.Find("ManagerID", ManagerID);

                return;
            }

            var Row = ManagersDataTable.Select("ManagerID = " + ManagerID);

            SelectedUsersDataTable.ImportRow(Row[0]);

            SelectedUsersBindingSource.Position = SelectedUsersBindingSource.Find("ManagerID", ManagerID);
        }

        public void RemoveCurrent()
        {
            var Pos = SelectedUsersBindingSource.Position;
            SelectedUsersBindingSource.RemoveCurrent();

            //остается на позиции удаленного
            if (SelectedUsersBindingSource.Count > 0)
                if (Pos >= SelectedUsersBindingSource.Count)
                {
                    SelectedUsersBindingSource.MoveLast();
                    SelectedUsersGrid.Rows[SelectedUsersGrid.Rows.Count - 1].Selected = true;
                }
                else
                    SelectedUsersBindingSource.Position = Pos;

            ((DataTable)SelectedUsersBindingSource.DataSource).AcceptChanges();
        }

        public void ClearCurrentUpdates()
        {
            ((DataRowView)SelectedUsersBindingSource.Current)["UpdatesCount"] = 0;
        }

        public void FillMessages(int SenderID)
        {
            MessagesDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT TOP 100 MessageID, SenderID, RecipientID,SenderTypeID, RecipientTypeID,"
                                               + " SendDateTime, MessageText, infiniu2_zovreference.dbo.Managers.Name AS SenderName, infiniu2_users.dbo.Users.Name AS SenderName2 "
                                               + " FROM infiniu2_zovreference.dbo.ZOVMessages"
                                               + " left JOIN infiniu2_zovreference.dbo.Managers ON infiniu2_zovreference.dbo.Managers.ManagerID = infiniu2_zovreference.dbo.ZOVMessages.SenderID"
                                               + " left JOIN infiniu2_users.dbo.Users ON infiniu2_users.dbo.Users.UserID = infiniu2_zovreference.dbo.ZOVMessages.SenderID and SenderTypeID = 0"
                                               + " WHERE (RecipientID = " + SenderID + " AND SenderID = " + CurrentUserID + ") OR (RecipientID = " + CurrentUserID + " AND SenderID = " + SenderID +
                                               ") ORDER BY SendDateTime DESC", ConnectionStrings.MarketingReferenceConnectionString))
            {
                DA.Fill(MessagesDataTable);
                for (var i = 0; i < MessagesDataTable.Rows.Count; i++)
                {
                    if (MessagesDataTable.Rows[i]["SenderName"] == DBNull.Value)
                        MessagesDataTable.Rows[i]["SenderName"] = MessagesDataTable.Rows[i]["SenderName2"];
                }
                MessagesDataTable.Columns.Remove("SenderName2");
            }
        }

        private string GetUserNameByID(int UserID)
        {
            return UsersDataTable.Select("UserID = " + UserID)[0]["Name"].ToString();
        }

        public bool IsEmptyMessage(string sText)
        {
            if (sText.Length == 0)
                return true;

            var n = 0;

            foreach (var c in sText)
            {
                if (c == '\n' || c == '\r' || c == ' ')
                    n++;
            }

            if (n == sText.Length)
                return true;

            return false;
        }

        public void AddMessage(int RecipientID, string sText)
        {
            using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM ZOVMessages", ConnectionStrings.ZOVReferenceConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        var NewRow = DT.NewRow();
                        NewRow["SenderID"] = CurrentUserID;
                        NewRow["RecipientID"] = RecipientID;
                        NewRow["MessageText"] = sText;
                        NewRow["SenderTypeID"] = false;
                        NewRow["RecipientTypeID"] = true;

                        var DateTime = Security.GetCurrentDate();

                        NewRow["SendDateTime"] = DateTime;
                        DT.Rows.Add(NewRow);

                        DA.Update(DT);

                        FillMessages(RecipientID);

                        var Row = MessagesDataTable.Select("SendDateTime = '" + DateTime.ToString("yyyy-MM-dd HH:mm:ss.fff") + "'");

                        using (var sDA = new SqlDataAdapter("SELECT TOP 0 * FROM SubscribesRecords", ConnectionStrings.LightConnectionString))
                        {
                            using (var sCB = new SqlCommandBuilder(sDA))
                            {
                                using (var sDT = new DataTable())
                                {
                                    sDA.Fill(sDT);

                                    var sNewRow = sDT.NewRow();
                                    sNewRow["SubscribesItemID"] = 15;
                                    sNewRow["TableItemID"] = Convert.ToInt32(Row[0][0]);
                                    sNewRow["UserID"] = RecipientID;
                                    sNewRow["UserTypeID"] = 1;
                                    sDT.Rows.Add(sNewRow);

                                    sDA.Update(sDT);
                                }
                            }
                        }
                    }
                }
            }
        }

        public void GetNewMessages()
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM infiniu2_zovreference.dbo.ZOVMessages WHERE infiniu2_zovreference.dbo.ZOVMessages.MessageID IN (SELECT TableItemID FROM SubscribesRecords WHERE SubscribesItemID = 15 AND UserTypeID = 0 AND UserID = " + CurrentUserID + ") ORDER BY SendDateTime DESC", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    if (DA.Fill(DT) == 0)
                        return;

                    foreach (DataRow Row in DT.Rows)
                    {
                        AddSenderToSelected(Convert.ToInt32(Row["SenderID"]));
                    }
                }
            }
        }

        public void AddSenderToSelected(int ManagerID)
        {
            var sRow = SelectedUsersDataTable.Select("ManagerID = " + ManagerID);

            if (sRow.Count() > 0)
            {
                sRow[0]["UpdatesCount"] = 1;
                return;
            }

            var Row = ManagersDataTable.Select("ManagerID = " + ManagerID);

            var NewRow = SelectedUsersDataTable.NewRow();
            NewRow["ManagerID"] = Row[0]["ManagerID"];
            NewRow["Name"] = Row[0]["Name"];
            NewRow["UpdatesCount"] = 1;
            SelectedUsersDataTable.Rows.Add(NewRow);
        }

        public void CheckOnline()
        {
            using (var DA = new SqlDataAdapter("SELECT ManagerID FROM Managers WHERE Online = 1", ConnectionStrings.ZOVReferenceConnectionString))
            {
                using (var DT = new DataTable())
                {
                    try
                    {
                        DA.Fill(DT);
                    }
                    catch
                    {
                        return;
                    }

                    foreach (DataRow Row in ManagersDataTable.Rows)
                    {
                        if (DT.Select("ManagerID = " + Row["ManagerID"]).Count() > 0)
                            Row["OnlineStatus"] = true;
                        else
                            Row["OnlineStatus"] = false;
                    }

                    foreach (DataRow Row in SelectedUsersDataTable.Rows)
                    {
                        if (DT.Select("ManagerID = " + Row["ManagerID"]).Count() > 0)
                            Row["OnlineStatus"] = true;
                        else
                            Row["OnlineStatus"] = false;
                    }
                }
            }
        }
    }





    public class MarketingNews
    {
        public DataTable UsersDataTable;
        public DataTable NewsDataTable;
        public DataTable NewsLikesDataTable;
        public DataTable AttachmentsDataTable;
        public DataTable CommentsDataTable;
        public DataTable CommentsLikesDataTable;
        public DataTable CommentsSubsRecordsDataTable;
        public DataTable NewsSubsRecordsDataTable;
        public DataTable CurrentCommentsDataTable;
        public DataTable ClientsDataTable;
        public DataTable ClientsManagersDataTable;
        public DataTable ClientsSelectDataTable;
        public DataTable ClientsNewsDataTable;
        public DataTable ClientsManagersNewsDataTable;

        public FileManager FM = new FileManager();

        public MarketingNews()
        {
            NewsDataTable = new DataTable();
            NewsLikesDataTable = new DataTable();
            CommentsLikesDataTable = new DataTable();
            UsersDataTable = new DataTable();
            ClientsNewsDataTable = new DataTable();
            ClientsManagersNewsDataTable = new DataTable();
            Fill();
        }

        public void Fill()
        {
            UsersDataTable.Clear();

            UsersDataTable = new DataView(TablesManager.UsersDataTable).ToTable(true, "UserID", "Name", "Photo");

            if (UsersDataTable.Columns["SenderTypeID"] == null)
                UsersDataTable.Columns.Add(new DataColumn("SenderTypeID", Type.GetType("System.Int32")));

            foreach (DataRow Row in UsersDataTable.Rows)
                Row["SenderTypeID"] = 0;

            AttachmentsDataTable = new DataTable();

            using (var DA = new SqlDataAdapter("SELECT NewsAttachID, NewsID, FileName, FileSize FROM NewsAttachs", ConnectionStrings.MarketingReferenceConnectionString))
            {
                DA.Fill(AttachmentsDataTable);
            }

            ClientsDataTable = new DataTable();
            using (var DA = new SqlDataAdapter("SELECT ClientID, ClientName FROM Clients ORDER BY ClientName ASC", ConnectionStrings.MarketingReferenceConnectionString))
            {
                DA.Fill(ClientsDataTable);
            }
            ClientsManagersDataTable = new DataTable();
            using (var DA = new SqlDataAdapter("SELECT * FROM ClientsManagers ORDER BY Name ASC", ConnectionStrings.MarketingReferenceConnectionString))
            {
                DA.Fill(ClientsManagersDataTable);
            }

            ClientsSelectDataTable = new DataTable();

            //ClientsSelectDataTable = ClientsDataTable.Copy();
            ClientsSelectDataTable.Columns.Add(new DataColumn("Check", Type.GetType("System.Boolean")));
            ClientsSelectDataTable.Columns.Add(new DataColumn("UserTypeID", Type.GetType("System.Int32")));
            ClientsSelectDataTable.Columns.Add(new DataColumn("ID", Type.GetType("System.Int32")));
            ClientsSelectDataTable.Columns.Add(new DataColumn("Name", Type.GetType("System.String")));
            for (var i = 0; i < ClientsDataTable.Rows.Count; i++)
            {
                var NewRow = ClientsSelectDataTable.NewRow();
                NewRow["Check"] = 0;
                NewRow["UserTypeID"] = 0;
                NewRow["ID"] = ClientsDataTable.Rows[i]["ClientID"];
                NewRow["Name"] = ClientsDataTable.Rows[i]["ClientName"];
                ClientsSelectDataTable.Rows.Add(NewRow);
            }

            for (var i = 0; i < ClientsManagersDataTable.Rows.Count; i++)
            {
                var NewRow = ClientsSelectDataTable.NewRow();
                NewRow["Check"] = 0;
                NewRow["UserTypeID"] = 1;
                NewRow["ID"] = ClientsManagersDataTable.Rows[i]["ManagerID"];
                NewRow["Name"] = ClientsManagersDataTable.Rows[i]["Name"];
                ClientsSelectDataTable.Rows.Add(NewRow);
            }

            CommentsDataTable = new DataTable();

            using (var DA = new SqlDataAdapter("SELECT * FROM NewsComments", ConnectionStrings.MarketingReferenceConnectionString))
            {
                DA.Fill(CommentsDataTable);
            }

            CommentsSubsRecordsDataTable = new DataTable();

            using (var DA = new SqlDataAdapter("SELECT SubscribesRecords.*, NewsComments.NewsID FROM SubscribesRecords INNER JOIN infiniu2_marketingreference.dbo.NewsComments ON infiniu2_marketingreference.dbo.NewsComments.NewsCommentID = SubscribesRecords.TableItemID WHERE SubscribesItemID = 17 AND SubscribesRecords.UserTypeID = 0 AND SubscribesRecords.UserID = " + Security.CurrentUserID, ConnectionStrings.LightConnectionString))
            {
                DA.Fill(CommentsSubsRecordsDataTable);
            }

            NewsSubsRecordsDataTable = new DataTable();

            using (var sDA = new SqlDataAdapter("SELECT * FROM SubscribesRecords WHERE SubscribesItemID = 16 " +
                                                " AND UserTypeID = 0 AND UserID = " + Security.CurrentUserID, ConnectionStrings.LightConnectionString))
            {
                sDA.Fill(NewsSubsRecordsDataTable);
            }
        }

        public void FillNewClientsNews()
        {
            ClientsNewsDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT ClientID, ClientName FROM Clients WHERE ClientID IN (SELECT SenderID FROM News WHERE NewsID IN (SELECT TableItemID FROM infiniu2_light.dbo.SubscribesRecords WHERE UserTypeID = 0 AND SubscribesItemID = 16 AND UserID = " + Security.CurrentUserID + ")) OR ClientID IN (SELECT SenderID FROM News WHERE NewsID IN (SELECT NewsID FROM NewsComments WHERE NewsCommentID IN (SELECT TableItemID FROM infiniu2_light.dbo.SubscribesRecords WHERE UserTypeID = 0 AND SubscribesItemID = 17 AND UserID = " + Security.CurrentUserID + ")))", ConnectionStrings.MarketingReferenceConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(ClientsNewsDataTable);
                }
            }
        }

        public void FillNewManagersNews()
        {
            ClientsManagersNewsDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT ManagerID, Name FROM ClientsManagers WHERE ManagerID IN (SELECT SenderID FROM News WHERE SenderTypeID=3 AND NewsID IN (SELECT TableItemID FROM infiniu2_light.dbo.SubscribesRecords WHERE UserTypeID = 0 AND SubscribesItemID = 16 AND UserID = " + Security.CurrentUserID + ")) OR ManagerID IN (SELECT SenderID FROM News WHERE SenderTypeID=3 AND NewsID IN (SELECT NewsID FROM NewsComments WHERE UserTypeID=3 AND NewsCommentID IN (SELECT TableItemID FROM infiniu2_light.dbo.SubscribesRecords WHERE UserTypeID = 0 AND SubscribesItemID = 17 AND UserID = " + Security.CurrentUserID + ")))", ConnectionStrings.MarketingReferenceConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(ClientsManagersNewsDataTable);
                }
            }
        }

        public void ReloadClientsNews(int Count, int ClientID)
        {
            NewsDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT TOP " + Count + " * FROM News WHERE ((RecipientTypeID=2 AND RecipientID=" + ClientID + ") OR (SenderTypeID=2 AND SenderID = " + ClientID + ")) AND Pending <> 1 ORDER BY LastCommentDateTime DESC", ConnectionStrings.MarketingReferenceConnectionString))
            {
                DA.Fill(NewsDataTable);
            }

            //subscribes///////////////////////
            if (NewsDataTable.Columns["New"] == null)
                NewsDataTable.Columns.Add(new DataColumn("New", Type.GetType("System.Int32")));

            if (NewsDataTable.Columns["NewComments"] == null)
                NewsDataTable.Columns.Add(new DataColumn("NewComments", Type.GetType("System.Int32")));

            foreach (DataRow Row in NewsSubsRecordsDataTable.Rows)
            {
                if (NewsDataTable.Select("NewsID = " + Row["TableItemID"]).Count() > 0)
                {
                    NewsDataTable.Select("NewsID = " + Row["TableItemID"])[0]["New"] = 1;
                }
            }

            foreach (DataRow Row in CommentsSubsRecordsDataTable.Rows)
            {
                if (NewsDataTable.Select("NewsID = " + Row["NewsID"]).Count() > 0)
                {
                    NewsDataTable.Select("NewsID = " + Row["NewsID"])[0]["NewComments"] =
                            CommentsSubsRecordsDataTable.Select("NewsID = " + Row["NewsID"]).Count();
                }
            }
            //////////////////////////////////
        }

        public void ReloadManagersNews(int Count, int ManagerID)
        {
            NewsDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT TOP " + Count + " * FROM News WHERE ((RecipientTypeID=3 AND RecipientID=" + ManagerID + ") OR (SenderTypeID=3 AND SenderID = " + ManagerID + ")) AND Pending <> 1 ORDER BY LastCommentDateTime DESC", ConnectionStrings.MarketingReferenceConnectionString))
            {
                DA.Fill(NewsDataTable);
            }

            //subscribes///////////////////////
            if (NewsDataTable.Columns["New"] == null)
                NewsDataTable.Columns.Add(new DataColumn("New", Type.GetType("System.Int32")));

            if (NewsDataTable.Columns["NewComments"] == null)
                NewsDataTable.Columns.Add(new DataColumn("NewComments", Type.GetType("System.Int32")));

            foreach (DataRow Row in NewsSubsRecordsDataTable.Rows)
            {
                if (NewsDataTable.Select("NewsID = " + Row["TableItemID"]).Count() > 0)
                {
                    NewsDataTable.Select("NewsID = " + Row["TableItemID"])[0]["New"] = 1;
                }
            }

            foreach (DataRow Row in CommentsSubsRecordsDataTable.Rows)
            {
                if (NewsDataTable.Select("NewsID = " + Row["NewsID"]).Count() > 0)
                {
                    NewsDataTable.Select("NewsID = " + Row["NewsID"])[0]["NewComments"] =
                            CommentsSubsRecordsDataTable.Select("NewsID = " + Row["NewsID"]).Count();
                }
            }
            //////////////////////////////////
        }

        public void ReloadComments()
        {
            CommentsDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT * FROM NewsComments", ConnectionStrings.MarketingReferenceConnectionString))
            {
                DA.Fill(CommentsDataTable);
            }
        }

        public void ReloadSubscribes()
        {
            CommentsSubsRecordsDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT SubscribesRecords.*, infiniu2_marketingreference.dbo.NewsComments.NewsID FROM SubscribesRecords INNER JOIN infiniu2_marketingreference.dbo.NewsComments ON infiniu2_marketingreference.dbo.NewsComments.NewsCommentID = SubscribesRecords.TableItemID WHERE SubscribesItemID = 17 AND SubscribesRecords.UserTypeID = 0 AND SubscribesRecords.UserID = " + Security.CurrentUserID, ConnectionStrings.LightConnectionString))
            {
                DA.Fill(CommentsSubsRecordsDataTable);
            }

            NewsSubsRecordsDataTable.Clear();

            using (var sDA = new SqlDataAdapter("SELECT * FROM SubscribesRecords WHERE SubscribesItemID = 16 " +
                                                " AND UserTypeID = 0 AND UserID = " + Security.CurrentUserID, ConnectionStrings.LightConnectionString))
            {
                sDA.Fill(NewsSubsRecordsDataTable);
            }
        }

        public void FillNews(int ClientID)
        {
            NewsDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT TOP 20 * FROM News WHERE ((RecipientTypeID=2 AND RecipientID=" + ClientID + ") OR (SenderTypeID=2 AND SenderID = " + ClientID + ")) AND Pending <> 1 ORDER BY LastCommentDateTime DESC", ConnectionStrings.MarketingReferenceConnectionString))
            {
                DA.Fill(NewsDataTable);
            }

            if (NewsDataTable.Columns["New"] == null)
                NewsDataTable.Columns.Add(new DataColumn("New", Type.GetType("System.Int32")));

            if (NewsDataTable.Columns["NewComments"] == null)
                NewsDataTable.Columns.Add(new DataColumn("NewComments", Type.GetType("System.Int32")));

            foreach (DataRow Row in NewsSubsRecordsDataTable.Rows)
            {
                if (NewsDataTable.Select("NewsID = " + Row["TableItemID"]).Count() > 0)
                {
                    NewsDataTable.Select("NewsID = " + Row["TableItemID"])[0]["New"] = 1;
                }
            }

            foreach (DataRow Row in CommentsSubsRecordsDataTable.Rows)
            {
                if (NewsDataTable.Select("NewsID = " + Row["NewsID"]).Count() > 0)
                {
                    NewsDataTable.Select("NewsID = " + Row["NewsID"])[0]["NewComments"] =
                            CommentsSubsRecordsDataTable.Select("NewsID = " + Row["NewsID"]).Count();
                }
            }
        }

        public void FillMNews(int ManagerID)
        {
            NewsDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT TOP 20 * FROM News WHERE ((RecipientTypeID=3 AND RecipientID=" + ManagerID + ") OR (SenderTypeID=3 AND SenderID = " + ManagerID + ")) AND Pending <> 1 ORDER BY LastCommentDateTime DESC", ConnectionStrings.MarketingReferenceConnectionString))
            {
                DA.Fill(NewsDataTable);
            }

            if (NewsDataTable.Columns["New"] == null)
                NewsDataTable.Columns.Add(new DataColumn("New", Type.GetType("System.Int32")));

            if (NewsDataTable.Columns["NewComments"] == null)
                NewsDataTable.Columns.Add(new DataColumn("NewComments", Type.GetType("System.Int32")));

            foreach (DataRow Row in NewsSubsRecordsDataTable.Rows)
            {
                if (NewsDataTable.Select("NewsID = " + Row["TableItemID"]).Count() > 0)
                {
                    NewsDataTable.Select("NewsID = " + Row["TableItemID"])[0]["New"] = 1;
                }
            }

            foreach (DataRow Row in CommentsSubsRecordsDataTable.Rows)
            {
                if (NewsDataTable.Select("NewsID = " + Row["NewsID"]).Count() > 0)
                {
                    NewsDataTable.Select("NewsID = " + Row["NewsID"])[0]["NewComments"] =
                            CommentsSubsRecordsDataTable.Select("NewsID = " + Row["NewsID"]).Count();
                }
            }
        }

        public void FillLikes()
        {
            NewsLikesDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT * FROM NewsLikes", ConnectionStrings.MarketingReferenceConnectionString))
            {
                DA.Fill(NewsLikesDataTable);
            }


            CommentsLikesDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT * FROM NewsCommentsLikes", ConnectionStrings.MarketingReferenceConnectionString))
            {
                DA.Fill(CommentsLikesDataTable);
            }
        }

        public void FillMoreNews(int Count, int ClientID)
        {
            NewsDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT TOP " + Count + " * FROM News WHERE ClientID = " + ClientID + " AND Pending <> 1 ORDER BY LastCommentDateTime DESC", ConnectionStrings.MarketingReferenceConnectionString))
            {
                DA.Fill(NewsDataTable);
            }
        }

        public bool IsMoreNews(int Count, int ClientID)
        {
            using (var DA = new SqlDataAdapter("SELECT Count(NewsID) FROM News WHERE RecipientID = " + ClientID,
                ConnectionStrings.MarketingReferenceConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    return Convert.ToInt32(DT.Rows[0][0]) > Count;
                }
            }
        }

        public DateTime AddNews(int SenderID, int SenderTypeID, string HeaderText, string BodyText, int RecipientID, int RecipientTypeID)
        {
            DateTime DateTime;

            using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM News", ConnectionStrings.MarketingReferenceConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        DateTime = Security.GetCurrentDate();

                        var NewRow = DT.NewRow();
                        NewRow["SenderID"] = SenderID;
                        NewRow["SenderTypeID"] = SenderTypeID;
                        NewRow["HeaderText"] = HeaderText;
                        NewRow["BodyText"] = BodyText;
                        NewRow["Pending"] = true;
                        NewRow["RecipientID"] = RecipientID;
                        NewRow["RecipientTypeID"] = RecipientTypeID;


                        NewRow["DateTime"] = DateTime;
                        NewRow["LastCommentDateTime"] = DateTime;
                        DT.Rows.Add(NewRow);


                        DA.Update(DT);
                    }
                }
            }

            return DateTime;
        }

        public void AddSubscribeForNews(DateTime DateTime, int UserTypeID, int ClientID)
        {
            using (var sDA = new SqlDataAdapter("SELECT * FROM News WHERE RecipientTypeID=" + UserTypeID + " AND RecipientID = " + ClientID + " AND DateTime = '" + DateTime.ToString("yyyy-MM-dd HH:mm:ss.fff") + "'",
                                                                       ConnectionStrings.MarketingReferenceConnectionString))
            {
                using (var sDT = new DataTable())
                {
                    sDA.Fill(sDT);

                    using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM SubscribesRecords", ConnectionStrings.LightConnectionString))
                    {
                        using (var CB = new SqlCommandBuilder(DA))
                        {
                            using (var DT = new DataTable())
                            {
                                DA.Fill(DT);

                                using (var uDA = new SqlDataAdapter("SELECT * FROM ModulesAccess WHERE ModuleID = 158", ConnectionStrings.UsersConnectionString))
                                {
                                    using (var uDT = new DataTable())
                                    {
                                        uDA.Fill(uDT);

                                        foreach (DataRow Row in uDT.Rows)
                                        {
                                            if (Convert.ToInt32(Row["UserID"]) == Security.CurrentUserID)
                                                continue;

                                            var NewRow = DT.NewRow();
                                            NewRow["SubscribesItemID"] = 16;
                                            NewRow["TableItemID"] = sDT.Rows[0]["NewsID"];
                                            NewRow["UserID"] = Row["UserID"];
                                            NewRow["UserTypeID"] = 0;
                                            DT.Rows.Add(NewRow);
                                        }

                                        DA.Update(DT);
                                    }
                                }


                                //subscribe for clientID or managers
                                {
                                    var NewRow = DT.NewRow();
                                    NewRow["SubscribesItemID"] = 16;
                                    NewRow["TableItemID"] = sDT.Rows[0]["NewsID"];
                                    NewRow["UserID"] = ClientID;
                                    NewRow["UserTypeID"] = UserTypeID;
                                    DT.Rows.Add(NewRow);
                                }

                                DA.Update(DT);


                            }
                        }
                    }


                }
            }
        }

        public void ClearPending(DateTime DateTime, int ClientID)
        {
            using (var sDA = new SqlDataAdapter("SELECT * FROM News WHERE DateTime = '" + DateTime.ToString("yyyy-MM-dd HH:mm:ss.fff") + "'",
                                                                      ConnectionStrings.MarketingReferenceConnectionString))
            {
                using (var CB = new SqlCommandBuilder(sDA))
                {
                    using (var sDT = new DataTable())
                    {
                        sDA.Fill(sDT);
                        for (var i = 0; i < sDT.Rows.Count; i++)
                        {
                            sDT.Rows[i]["Pending"] = false;
                        }
                        sDA.Update(sDT);
                    }

                }

            }
        }

        public void AddComments(int UserID, int UserTypeID, int NewsID, string Text)
        {
            DateTime Date;

            using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM NewsComments", ConnectionStrings.MarketingReferenceConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        var NewRow = DT.NewRow();
                        NewRow["NewsID"] = NewsID;
                        NewRow["NewsComment"] = Text;
                        NewRow["UserID"] = UserID;
                        NewRow["UserTypeID"] = UserTypeID;

                        Date = Security.GetCurrentDate();

                        NewRow["DateTime"] = Date;
                        DT.Rows.Add(NewRow);

                        DA.Update(DT);
                    }
                }
            }

            using (var DA = new SqlDataAdapter("SELECT NewsID, LastCommentDateTime FROM News WHERE NewsID =" + NewsID, ConnectionStrings.MarketingReferenceConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        DT.Rows[0]["LastCommentDateTime"] = Date;

                        DA.Update(DT);
                    }
                }
            }

            AddNewsCommentsSubs(NewsID, UserID, Date);
        }


        public void LikeNews(int UserID, int NewsID)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM NewsLikes WHERE NewsID = " + NewsID + " AND UserID = " + UserID, ConnectionStrings.MarketingReferenceConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        if (DT.Rows.Count > 0)//i like
                        {
                            using (var dDA = new SqlDataAdapter("DELETE FROM NewsLikes WHERE NewsID = " + NewsID + " AND UserID = " + UserID, ConnectionStrings.MarketingReferenceConnectionString))
                            {
                                using (var dDT = new DataTable())
                                {
                                    dDA.Fill(dDT);
                                }
                            }

                            return;
                        }

                        var NewRow = DT.NewRow();
                        NewRow["NewsID"] = NewsID;
                        NewRow["UserID"] = UserID;
                        DT.Rows.Add(NewRow);

                        DA.Update(DT);
                    }
                }
            }
        }

        public void LikeComments(int UserID, int NewsID, int NewsCommentID)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM NewsCommentsLikes WHERE NewsCommentID = " + NewsCommentID + " AND UserID = " + UserID, ConnectionStrings.MarketingReferenceConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        if (DT.Rows.Count > 0)
                        {
                            using (var dDA = new SqlDataAdapter("DELETE FROM NewsCommentsLikes WHERE NewsCommentID = " + NewsCommentID + " AND UserID = " + UserID, ConnectionStrings.MarketingReferenceConnectionString))
                            {
                                using (var dDT = new DataTable())
                                {
                                    dDA.Fill(dDT);
                                }
                            }

                            return;
                        }

                        var NewRow = DT.NewRow();
                        NewRow["NewsID"] = NewsID;
                        NewRow["NewsCommentID"] = NewsCommentID;
                        NewRow["UserID"] = UserID;
                        DT.Rows.Add(NewRow);

                        DA.Update(DT);
                    }
                }
            }
        }


        public void EditComments(int UserID, int NewsCommentID, string Text)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM NewsComments WHERE NewsCommentID = " + NewsCommentID, ConnectionStrings.MarketingReferenceConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        DT.Rows[0]["NewsComment"] = Text;

                        DA.Update(DT);
                    }
                }
            }
        }


        public bool Attach(DataTable AttachmentsDataTable, DateTime NewsDateTime, ref int CurrentUploadedFile, int ClientID)
        {
            if (AttachmentsDataTable.Rows.Count == 0)
                return true;

            var Ok = true;

            int NewsID;

            using (var DA = new SqlDataAdapter("SELECT NewsID FROM News WHERE ClientID = " + ClientID + " AND  DateTime = '" + NewsDateTime.ToString("yyyy-MM-dd HH:mm:ss.fff") + "'", ConnectionStrings.MarketingReferenceConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    NewsID = Convert.ToInt32(DT.Rows[0]["NewsID"]);
                }
            }

            using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM NewsAttachs", ConnectionStrings.MarketingReferenceConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        foreach (DataRow Row in AttachmentsDataTable.Rows)
                        {
                            FileInfo fi;

                            try
                            {
                                fi = new FileInfo(Row["Path"].ToString());
                                //WriteLog("UploadFile fileinfosize: " + Row["Path"].ToString(), Security.CurrentUserID, -1);
                            }
                            catch
                            {
                                Ok = false;
                                continue;
                            }

                            var NewRow = DT.NewRow();
                            NewRow["NewsID"] = NewsID;
                            NewRow["FileName"] = Row["FileName"];
                            NewRow["FileSize"] = fi.Length;
                            DT.Rows.Add(NewRow);
                        }

                        DA.Update(DT);
                    }
                }
            }

            //write to ftp
            using (var DA = new SqlDataAdapter("SELECT * FROM NewsAttachs WHERE NewsID = " + NewsID, ConnectionStrings.MarketingReferenceConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    foreach (DataRow Row in AttachmentsDataTable.Rows)
                    {
                        CurrentUploadedFile++;

                        try
                        {
                            if (FM.UploadFile(Row["Path"].ToString(), Configs.DocumentsPath +
                                          FileManager.GetPath("MarketingNews") + "/" +
                                          DT.Select("FileName = '" + Row["FileName"] + "'")[0]["NewsAttachID"] + ".idf", Configs.FTPType) == false)
                            {
                                break;
                            }
                        }
                        catch
                        {
                            Ok = false;
                            break;
                        }
                    }

                    DA.Update(DT);
                }
            }

            ReloadAttachments();

            return Ok;
        }

        public bool EditAttachments(int NewsID, DataTable AttachmentsDataTable, ref int CurrentUploadedFile, ref int TotalFilesCount)
        {
            var Ok = false;

            using (var DA = new SqlDataAdapter("SELECT NewsAttachID, NewsID, FileName FROM NewsAttachs WHERE NewsID = " + NewsID, ConnectionStrings.MarketingReferenceConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        foreach (DataRow Row in DT.Rows)
                        {
                            if (AttachmentsDataTable.Select("FileName = '" + Row["FileName"] + "'").Count() == 0)
                            {
                                FM.DeleteFile(Configs.DocumentsPath + FileManager.GetPath("MarketingNews") + "/" +
                                              DT.Select("FileName = '" + Row["FileName"] + "'")[0]["NewsAttachID"] + ".idf", Configs.FTPType);
                                Row.Delete();
                            }
                        }

                        DA.Update(DT);
                    }
                }
            }

            using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM NewsAttachs WHERE NewsID = " + NewsID, ConnectionStrings.MarketingReferenceConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);
                        Ok = true;

                        //add new
                        foreach (DataRow Row in AttachmentsDataTable.Rows)
                        {
                            if (Row["Path"].ToString() == "server")
                                continue;

                            TotalFilesCount++;

                            FileInfo fi;

                            try
                            {
                                fi = new FileInfo(Row["Path"].ToString());

                            }
                            catch
                            {
                                Ok = false;
                                continue;
                            }

                            var NewRow = DT.NewRow();
                            NewRow["NewsID"] = NewsID;
                            NewRow["FileName"] = Row["FileName"];
                            NewRow["FileSize"] = fi.Length;
                            DT.Rows.Add(NewRow);
                        }

                        DA.Update(DT);
                    }
                }
            }


            //write to ftp
            using (var DA = new SqlDataAdapter("SELECT * FROM NewsAttachs WHERE NewsID = " + NewsID, ConnectionStrings.MarketingReferenceConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    foreach (DataRow Row in AttachmentsDataTable.Rows)
                    {
                        if (Row["Path"].ToString() == "server")
                            continue;

                        CurrentUploadedFile++;

                        try
                        {
                            if (FM.UploadFile(Row["Path"].ToString(), Configs.DocumentsPath +
                                         FileManager.GetPath("MarketingNews") + "/" +
                                         DT.Select("FileName = '" + Row["FileName"] + "'")[0]["NewsAttachID"] + ".idf", Configs.FTPType) == false)
                                break;
                        }
                        catch
                        {
                            Ok = false;
                            break;
                        }
                    }

                    DA.Update(DT);
                }
            }


            ReloadAttachments();

            return Ok;
        }

        public void ReloadAttachments()
        {
            AttachmentsDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT NewsAttachID, NewsID, FileName, FileSize FROM NewsAttachs", ConnectionStrings.MarketingReferenceConnectionString))
            {
                DA.Fill(AttachmentsDataTable);
            }
        }

        public void EditNews(int NewsID, int SenderID, int SenderTypeID, string HeaderText, string BodyText, int NewsCategoryID, DateTime DateTime)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM News WHERE NewsID = " + NewsID, ConnectionStrings.MarketingReferenceConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        if (DA.Fill(DT) == 0)
                            return;

                        DT.Rows[0]["HeaderText"] = HeaderText;
                        DT.Rows[0]["BodyText"] = BodyText;
                        DT.Rows[0]["DateTime"] = DateTime;

                        DA.Update(DT);
                    }
                }
            }
        }

        public static string TempPath()
        {
            return Environment.GetEnvironmentVariable("TEMP");
        }

        public int GetNewsUpdatesCount()
        {
            using (var DA = new SqlDataAdapter("SELECT Count (SubscribesRecordID) FROM SubscribesRecords WHERE (SubscribesItemID = 16 OR SubscribesItemID = 17) AND UserTypeID = 0 AND UserID = " + Security.CurrentUserID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    if (DA.Fill(DT) > 0)
                        return Convert.ToInt32(DT.Rows[0][0]);
                }
            }

            return 0;
        }

        public int GetNewsIDByDateTime(DateTime DateTime)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM News WHERE DateTime = '" + DateTime.ToString("yyyy-MM-dd HH:mm:ss.fff") + "'",
                                                                         ConnectionStrings.MarketingReferenceConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    return Convert.ToInt32(DT.Rows[0]["NewsID"]);
                }
            }
        }

        public void RemoveNews(int NewsID)
        {
            //ActiveNotifySystem.DeleteSubscribesRecord(1);

            using (var DA = new SqlDataAdapter("DELETE FROM NewsLikes WHERE NewsID = " + NewsID, ConnectionStrings.MarketingReferenceConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }

            using (var DA = new SqlDataAdapter("DELETE FROM News WHERE NewsID = " + NewsID, ConnectionStrings.MarketingReferenceConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);
                    }
                }
            }

            RemoveComments(NewsID);

            RemoveAttachments(NewsID);

            ReloadAttachments();
        }

        public void RemoveComments(int NewsID)
        {
            DeleteNewsSub(NewsID);
            DeleteCommentsSubs(NewsID);


            using (var DA = new SqlDataAdapter("DELETE FROM NewsCommentsLikes WHERE NewsID = " + NewsID, ConnectionStrings.MarketingReferenceConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }

            using (var DA = new SqlDataAdapter("DELETE FROM NewsComments WHERE NewsID = " + NewsID, ConnectionStrings.MarketingReferenceConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);
                    }
                }
            }


        }

        public void RemoveComment(int NewsCommentID)
        {
            using (var DA = new SqlDataAdapter("DELETE FROM NewsCommentsLikes WHERE NewsCommentID = " + NewsCommentID, ConnectionStrings.MarketingReferenceConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }

            using (var DA = new SqlDataAdapter("DELETE FROM NewsComments WHERE NewsCommentID = " + NewsCommentID, ConnectionStrings.MarketingReferenceConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }

            DeleteCommentsSub(NewsCommentID);
        }

        public void RemoveAttachments(int NewsID)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM NewsAttachs WHERE NewsID = " + NewsID, ConnectionStrings.MarketingReferenceConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    try
                    {
                        foreach (DataRow Row in DT.Rows)
                        {
                            FM.DeleteFile(Configs.DocumentsPath + FileManager.GetPath("MarketingNews") + "/" +
                                          DT.Select("FileName = '" + Row["FileName"] + "'")[0]["NewsAttachID"] + ".idf", Configs.FTPType);
                        }
                    }
                    catch
                    {
                        using (var fDA = new SqlDataAdapter("DELETE FROM NewsAttachs WHERE NewsID = " + NewsID, ConnectionStrings.MarketingReferenceConnectionString))
                        {
                            using (var CB = new SqlCommandBuilder(fDA))
                            {
                                using (var fDT = new DataTable())
                                {
                                    fDA.Fill(fDT);
                                }
                            }
                        }

                        ReloadAttachments();

                        return;
                    }
                }
            }



            using (var DA = new SqlDataAdapter("DELETE FROM NewsAttachs WHERE NewsID = " + NewsID, ConnectionStrings.MarketingReferenceConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);
                    }
                }
            }

            ReloadAttachments();
        }

        public void RemoveCurrentAttachments(int NewsID, DataTable AttachmentsDT)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM NewsAttachs WHERE NewsID = " + NewsID, ConnectionStrings.MarketingReferenceConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        foreach (DataRow Row in AttachmentsDT.Rows)
                        {
                            if (Row["Path"].ToString() == "server")
                                continue;

                            FM.DeleteFile(Configs.DocumentsPath + FileManager.GetPath("MarketingNews") + "/" +
                                      DT.Select("FileName = '" + Row["FileName"] + "'")[0]["NewsAttachID"] + ".idf", Configs.FTPType);

                            DT.Select("FileName = '" + Row["FileName"] + "'")[0].Delete();
                        }

                        DA.Update(DT);
                    }
                }
            }

            ReloadAttachments();
        }

        public DataTable GetAttachments(int NewsID)
        {
            using (var DA = new SqlDataAdapter("SELECT NewsAttachID, FileName FROM NewsAttachs WHERE NewsID = " + NewsID, ConnectionStrings.MarketingReferenceConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    return DT;
                }
            }
        }

        public int GetThisNewsSenderTypeID(int NewsID)
        {
            return Convert.ToInt32(NewsDataTable.Select("NewsID = " + NewsID)[0]["SenderTypeID"]);
        }

        public string GetThisNewsBodyText(int NewsID)
        {
            return NewsDataTable.Select("NewsID = " + NewsID)[0]["BodyText"].ToString();
        }

        public string GetThisNewsHeaderText(int NewsID)
        {
            return NewsDataTable.Select("NewsID = " + NewsID)[0]["HeaderText"].ToString();
        }

        public DateTime GetThisNewsDateTime(int NewsID)
        {
            return Convert.ToDateTime(NewsDataTable.Select("NewsID = " + NewsID)[0]["DateTime"]);
        }

        public int ShowAttachDownloadMenu(string FileName)
        {
            //0 cancel, 1 open, 2 save


            return AttachDownloadForm.Result;
        }

        public string GetAttachmentName(int NewsAttachID)
        {
            return AttachmentsDataTable.Select("NewsAttachID = " + NewsAttachID)[0]["FileName"].ToString();
        }

        public string SaveFile(int NewsAttachID)//temp folder
        {
            var tempFolder = Environment.GetEnvironmentVariable("TEMP");

            var FileName = "";

            using (var DA = new SqlDataAdapter("SELECT FileName, FileSize FROM NewsAttachs WHERE NewsAttachID = " + NewsAttachID, ConnectionStrings.MarketingReferenceConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    FileName = DT.Rows[0]["FileName"].ToString();

                    try
                    {
                        FM.DownloadFile(Configs.DocumentsPath + FileManager.GetPath("MarketingNews") + "/" + NewsAttachID + ".idf",
                                            tempFolder + "\\" + FileName, Convert.ToInt64(DT.Rows[0]["FileSize"]), Configs.FTPType);
                    }
                    catch
                    {
                        return null;
                    }

                    //byte[] b = (byte[])DT.Rows[0]["FileBytes"];
                    //MemoryStream ms = new MemoryStream(b);

                    //FileName = DT.Rows[0]["FileName"].ToString();

                    //FileStream s2;

                    //try
                    //{
                    //    s2 = new FileStream(tempFolder + "\\" + FileName, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite);
                    //}
                    //catch 
                    //{
                    //    FileName = GetNewFileName(tempFolder, DT.Rows[0]["FileName"].ToString());
                    //    s2 = new FileStream(tempFolder + "\\" + FileName, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite);
                    //}

                    //s2.Write(ms.ToArray(), 0, ms.Capacity);

                    //s2.Close();
                }
            }

            return tempFolder + "\\" + FileName;
        }

        public void SaveFile(int NewsAttachID, string sDestFileName)
        {
            using (var DA = new SqlDataAdapter("SELECT FileName, FileSize FROM NewsAttachs WHERE NewsAttachID = " + NewsAttachID, ConnectionStrings.MarketingReferenceConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    var FileName = DT.Rows[0]["FileName"].ToString();

                    try
                    {
                        FM.DownloadFile(Configs.DocumentsPath + FileManager.GetPath("MarketingNews") + "/" + NewsAttachID + ".idf",
                                           sDestFileName, Convert.ToInt64(DT.Rows[0]["FileSize"]), Configs.FTPType);
                    }
                    catch
                    {
                    }
                }
            }
        }

        private string SetNumber(string FileName, int Number)
        {
            var Ext = "";
            var Name = "";

            for (var i = FileName.Length - 1; i > 0; i--)
            {
                if (FileName[i] == '.')
                {
                    Ext = FileName.Substring(i + 1, FileName.Length - i - 1);
                    Name = FileName.Substring(0, i) + " (" + Number + ")." + Ext;
                    break;
                }
            }

            return Name;
        }

        private string GetNewFileName(string path, string FileName)
        {
            var fileInfo = new FileInfo(path + "\\" + FileName);

            if (!fileInfo.Exists)
                return FileName;

            var Ok = false;
            var n = 1;

            while (!Ok)
            {
                fileInfo = new FileInfo(path + "\\" + SetNumber(FileName, n));

                if (!fileInfo.Exists)
                    return SetNumber(FileName, n);

                n++;
            }

            return "";
        }

        public void DeleteCommentsSub(int NewsCommentID)
        {
            using (var DA = new SqlDataAdapter("DELETE FROM SubscribesRecords WHERE SubscribesItemID = 17 AND TableItemID = " + NewsCommentID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }
        }

        public void DeleteCommentsSubs(int NewsID)
        {
            using (var DA = new SqlDataAdapter("DELETE FROM SubscribesRecords WHERE SubscribesItemID = 17 AND TableItemID IN (SELECT infiniu2_marketingreference.dbo.NewsComments.NewsCommentID FROM infiniu2_marketingreference.dbo.NewsComments WHERE NewsID = " + NewsID + ")", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }
        }

        public void DeleteNewsSub(int NewsID)
        {
            using (var DA = new SqlDataAdapter("DELETE FROM SubscribesRecords WHERE SubscribesItemID = 16 AND TableItemID = " + NewsID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }
        }

        public void AddNewsCommentsSubs(int NewsID, int UserID, DateTime DateTime)
        {
            using (var sDA = new SqlDataAdapter("SELECT NewsCommentID, NewsID FROM NewsComments WHERE DateTime = '" + DateTime.ToString("yyyy-MM-dd HH:mm:ss.fff") + "'", ConnectionStrings.MarketingReferenceConnectionString))
            {
                using (var sDT = new DataTable())
                {
                    sDA.Fill(sDT);

                    using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM SubscribesRecords", ConnectionStrings.LightConnectionString))
                    {
                        using (var CB = new SqlCommandBuilder(DA))
                        {
                            using (var DT = new DataTable())
                            {
                                DA.Fill(DT);

                                using (var uDA = new SqlDataAdapter("SELECT * FROM ModulesAccess WHERE ModuleID = 158", ConnectionStrings.UsersConnectionString))
                                {
                                    using (var uDT = new DataTable())
                                    {
                                        uDA.Fill(uDT);

                                        foreach (DataRow Row in uDT.Rows)
                                        {
                                            if (Convert.ToInt32(Row["UserID"]) == Security.CurrentUserID)
                                                continue;

                                            var NewRow = DT.NewRow();
                                            NewRow["SubscribesItemID"] = 17;
                                            NewRow["TableItemID"] = sDT.Rows[0]["NewsCommentID"];
                                            NewRow["UserID"] = Row["UserID"];
                                            NewRow["UserTypeID"] = 0;
                                            DT.Rows.Add(NewRow);
                                        }

                                        DA.Update(DT);
                                    }
                                }


                                using (var uDA = new SqlDataAdapter("SELECT ClientID FROM Clients WHERE ClientID IN (SELECT ClientID FROM News WHERE NewsID = " + sDT.Rows[0]["NewsID"] + ")", ConnectionStrings.MarketingReferenceConnectionString))
                                {
                                    using (var uDT = new DataTable())
                                    {
                                        uDA.Fill(uDT);

                                        foreach (DataRow Row in uDT.Rows)
                                        {
                                            var NewRow = DT.NewRow();
                                            NewRow["SubscribesItemID"] = 17;
                                            NewRow["TableItemID"] = sDT.Rows[0]["NewsCommentID"];
                                            NewRow["UserID"] = Row["ClientID"];
                                            NewRow["UserTypeID"] = 2;
                                            DT.Rows.Add(NewRow);
                                        }

                                        DA.Update(DT);
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }









    public class InfiniumDocuments
    {
        public DataTable DocumentTypesDataTable;
        public DataTable UsersDataTable;
        public DataTable CorrespondentsDataTable;
        public DataTable DocumentsStatesDataTable;
        public DataTable InnerDocumentsDataTable;
        public DataTable IncomeDocumentsDataTable;
        public DataTable OuterDocumentsDataTable;
        public DataTable FactoryTypesDataTable;
        public DataTable DocumentsMenuDataTable;
        public DataTable CurrentFilesDataTable;
        public DataTable CommentsDataTable;
        public DataTable DocumentsCategoriesDataTable;

        public DataTable UpdatesDocumentsDataTable;
        public DataTable UpdatesCommentsDataTable;
        public DataTable UpdatesFilesDataTable;
        public DataTable UpdatesRecipientsDataTable;
        public DataTable UpdatesCommentsFilesDataTable;
        public DataTable UpdatesConfirmsDataTable;
        public DataTable UpdatesConfirmsRecipientsDataTable;

        public DataTable CurrentUploadedFiles;

        public FileManager FM;

        public InfiniumDocuments()
        {
            FM = new FileManager();

            Fill();
        }

        public void Fill()
        {
            DocumentsMenuDataTable = new DataTable();
            DocumentsMenuDataTable.Columns.Add(new DataColumn("Name", Type.GetType("System.String")));
            DocumentsMenuDataTable.Columns.Add(new DataColumn("Count", Type.GetType("System.Int32")));
            DocumentsMenuDataTable.Columns.Add(new DataColumn("Image", Type.GetType("System.Byte[]")));

            {
                var NewRow = DocumentsMenuDataTable.NewRow();
                NewRow["Name"] = "Лента";
                NewRow["Count"] = 0;

                var ms = new MemoryStream();
                Resources.DocumentsMenuUpdates.Save(ms, ImageFormat.Png);

                NewRow["Image"] = ms.ToArray();
                DocumentsMenuDataTable.Rows.Add(NewRow);
                ms.Dispose();
            }

            {
                var NewRow = DocumentsMenuDataTable.NewRow();
                NewRow["Name"] = "Все документы";
                NewRow["Count"] = 0;
                var ms = new MemoryStream();
                Resources.DocumentsMenuAllDocs.Save(ms, ImageFormat.Png);

                NewRow["Image"] = ms.ToArray();
                DocumentsMenuDataTable.Rows.Add(NewRow);
                ms.Dispose();
            }

            {
                var NewRow = DocumentsMenuDataTable.NewRow();
                NewRow["Name"] = "Мои документы";
                NewRow["Count"] = 0;
                var ms = new MemoryStream();
                Resources.DocumentsMenuUser.Save(ms, ImageFormat.Png);

                NewRow["Image"] = ms.ToArray();
                DocumentsMenuDataTable.Rows.Add(NewRow);
                ms.Dispose();
            }

            {
                var NewRow = DocumentsMenuDataTable.NewRow();
                NewRow["Name"] = "Согласование";
                NewRow["Count"] = 0;
                var ms = new MemoryStream();
                Resources.DocumentsMenuConfirms.Save(ms, ImageFormat.Png);

                NewRow["Image"] = ms.ToArray();
                DocumentsMenuDataTable.Rows.Add(NewRow);
                ms.Dispose();
            }


            DocumentTypesDataTable = new DataTable();

            using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsTypes ORDER BY DocumentType", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DocumentTypesDataTable);
                }
            }

            DocumentsCategoriesDataTable = new DataTable();

            using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsCategories", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DocumentsCategoriesDataTable);
                }
            }

            DocumentsStatesDataTable = new DataTable();

            using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsStates", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DocumentsStatesDataTable);
                }
            }

            CorrespondentsDataTable = new DataTable();

            using (var DA = new SqlDataAdapter("SELECT * FROM Correspondents ORDER BY CorrespondentName", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(CorrespondentsDataTable);
                }
            }

            UsersDataTable = TablesManager.UsersDataTable;

            CurrentFilesDataTable = new DataTable();

            InnerDocumentsDataTable = new DataTable();
            InnerDocumentsDataTable.Columns.Add(new DataColumn("CommentsCount", Type.GetType("System.Int32")));
            InnerDocumentsDataTable.Columns.Add(new DataColumn("FilesCount", Type.GetType("System.Int32")));
            InnerDocumentsDataTable.Columns.Add(new DataColumn("RecipientsCount", Type.GetType("System.Int32")));

            IncomeDocumentsDataTable = new DataTable();
            IncomeDocumentsDataTable.Columns.Add(new DataColumn("CommentsCount", Type.GetType("System.Int32")));
            IncomeDocumentsDataTable.Columns.Add(new DataColumn("FilesCount", Type.GetType("System.Int32")));
            IncomeDocumentsDataTable.Columns.Add(new DataColumn("RecipientsCount", Type.GetType("System.Int32")));

            OuterDocumentsDataTable = new DataTable();
            OuterDocumentsDataTable.Columns.Add(new DataColumn("CommentsCount", Type.GetType("System.Int32")));
            OuterDocumentsDataTable.Columns.Add(new DataColumn("FilesCount", Type.GetType("System.Int32")));
            OuterDocumentsDataTable.Columns.Add(new DataColumn("RecipientsCount", Type.GetType("System.Int32")));

            CommentsDataTable = new DataTable();

            FactoryTypesDataTable = new DataTable();

            using (var DA = new SqlDataAdapter("SELECT FactoryID, Factory FROM Factory WHERE FactoryID <> 0", ConnectionStrings.CatalogConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(FactoryTypesDataTable);
                }
            }

            UpdatesDocumentsDataTable = new DataTable();
            UpdatesCommentsDataTable = new DataTable();
            UpdatesFilesDataTable = new DataTable();
            UpdatesRecipientsDataTable = new DataTable();
            UpdatesCommentsFilesDataTable = new DataTable();
            UpdatesConfirmsDataTable = new DataTable();
            UpdatesConfirmsRecipientsDataTable = new DataTable();

            UpdatesDocumentsDataTable.Columns.Add(new DataColumn("UpdateDateTime", Type.GetType("System.DateTime")));

            CurrentUploadedFiles = new DataTable();
            CurrentUploadedFiles.Columns.Add(new DataColumn("FileName", Type.GetType("System.String")));
        }


        public void FillInnerDocuments(string Filter)
        {
            InnerDocumentsDataTable.Clear();

            var sNewFilter = "";

            if (Filter.Length > 0)
                sNewFilter = " WHERE " + Filter + " ORDER BY InnerDocumentID DESC";
            else
                sNewFilter = " ORDER BY InnerDocumentID DESC";

            using (var DA = new SqlDataAdapter("SELECT InnerDocuments.InnerDocumentID AS DocumentID, InnerDocuments.* FROM InnerDocuments" + sNewFilter, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(InnerDocumentsDataTable);
                }
            }

            using (var cDA = new SqlDataAdapter("SELECT infiniu2_light.dbo.InnerDocuments.InnerDocumentID, " +
                                                " COUNT(infiniu2_light.dbo.DocumentsComments.DocumentCommentID) AS ComCount FROM infiniu2_light.dbo.DocumentsComments " +
                                                " LEFT OUTER JOIN infiniu2_light.dbo.InnerDocuments ON infiniu2_light.dbo.DocumentsComments.DocumentID = " +
                                                "infiniu2_light.dbo.InnerDocuments.InnerDocumentID WHERE infiniu2_light.dbo.DocumentsComments.DocumentCategoryID = 0 " +
                                                "AND infiniu2_light.dbo.DocumentsComments.DocumentID IN " +
                                                "(SELECT infiniu2_light.dbo.InnerDocuments.InnerDocumentID" + Filter + ") GROUP BY " +
                                                "infiniu2_light.dbo.InnerDocuments.InnerDocumentID", ConnectionStrings.LightConnectionString))
            {
                using (var cDT = new DataTable())
                {
                    cDA.Fill(cDT);

                    foreach (DataRow Row in InnerDocumentsDataTable.Rows)
                    {
                        var Rows = cDT.Select("InnerDocumentID = " + Row["InnerDocumentID"]);

                        if (Rows.Count() > 0)
                            Row["CommentsCount"] = Rows[0]["ComCount"];
                        else
                            Row["CommentsCount"] = 0;
                    }
                }
            }


            using (var cDA = new SqlDataAdapter("SELECT infiniu2_light.dbo.InnerDocuments.InnerDocumentID, " +
                                                " COUNT(infiniu2_light.dbo.DocumentsRecipients.DocumentRecipientID) AS RecCount FROM infiniu2_light.dbo.DocumentsRecipients " +
                                                " LEFT OUTER JOIN infiniu2_light.dbo.InnerDocuments ON infiniu2_light.dbo.DocumentsRecipients.DocumentID = " +
                                                "infiniu2_light.dbo.InnerDocuments.InnerDocumentID WHERE infiniu2_light.dbo.DocumentsRecipients.DocumentCategoryID = 0 " +
                                                "AND infiniu2_light.dbo.DocumentsRecipients.DocumentID IN " +
                                                "(SELECT infiniu2_light.dbo.InnerDocuments.InnerDocumentID" + Filter + ") GROUP BY " +
                                                "infiniu2_light.dbo.InnerDocuments.InnerDocumentID", ConnectionStrings.LightConnectionString))
            {
                using (var cDT = new DataTable())
                {
                    cDA.Fill(cDT);

                    foreach (DataRow Row in InnerDocumentsDataTable.Rows)
                    {
                        var Rows = cDT.Select("InnerDocumentID = " + Row["InnerDocumentID"]);

                        if (Rows.Count() > 0)
                            Row["RecipientsCount"] = Rows[0]["RecCount"];
                        else
                            Row["RecipientsCount"] = 0;
                    }
                }
            }


            using (var cDA = new SqlDataAdapter("SELECT infiniu2_light.dbo.InnerDocuments.InnerDocumentID, " +
                                                " COUNT(infiniu2_light.dbo.DocumentsFiles.DocumentFileID) AS FileCount FROM infiniu2_light.dbo.DocumentsFiles " +
                                                " LEFT OUTER JOIN infiniu2_light.dbo.InnerDocuments ON infiniu2_light.dbo.DocumentsFiles.DocumentID = " +
                                                "infiniu2_light.dbo.InnerDocuments.InnerDocumentID WHERE infiniu2_light.dbo.DocumentsFiles.DocumentCategoryID = 0 " +
                                                "AND infiniu2_light.dbo.DocumentsFiles.DocumentID IN " +
                                                "(SELECT infiniu2_light.dbo.InnerDocuments.InnerDocumentID" + Filter + ") GROUP BY " +
                                                "infiniu2_light.dbo.InnerDocuments.InnerDocumentID", ConnectionStrings.LightConnectionString))
            {
                using (var cDT = new DataTable())
                {
                    cDA.Fill(cDT);

                    foreach (DataRow Row in InnerDocumentsDataTable.Rows)
                    {
                        var Rows = cDT.Select("InnerDocumentID = " + Row["InnerDocumentID"]);

                        if (Rows.Count() > 0)
                            Row["FilesCount"] = Rows[0]["FileCount"];
                        else
                            Row["FilesCount"] = 0;
                    }
                }
            }
        }

        public void FillIncomeDocuments(string Filter)
        {
            IncomeDocumentsDataTable.Clear();

            var sNewFilter = "";

            if (Filter.Length > 0)
                sNewFilter = " WHERE " + Filter + " ORDER BY IncomeDocumentID DESC";
            else
                sNewFilter = " ORDER BY IncomeDocumentID DESC";

            using (var DA = new SqlDataAdapter("SELECT IncomeDocuments.IncomeDocumentID AS DocumentID, IncomeDocuments.* FROM IncomeDocuments" + sNewFilter, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(IncomeDocumentsDataTable);
                }
            }

            using (var cDA = new SqlDataAdapter("SELECT infiniu2_light.dbo.IncomeDocuments.IncomeDocumentID, " +
                                                " COUNT(infiniu2_light.dbo.DocumentsComments.DocumentCommentID) AS ComCount FROM infiniu2_light.dbo.DocumentsComments " +
                                                " LEFT OUTER JOIN infiniu2_light.dbo.IncomeDocuments ON infiniu2_light.dbo.DocumentsComments.DocumentID = " +
                                                "infiniu2_light.dbo.IncomeDocuments.IncomeDocumentID WHERE infiniu2_light.dbo.DocumentsComments.DocumentCategoryID = 1 " +
                                                "AND infiniu2_light.dbo.DocumentsComments.DocumentID IN " +
                                                "(SELECT infiniu2_light.dbo.IncomeDocuments.IncomeDocumentID" + Filter + ") GROUP BY " +
                                                "infiniu2_light.dbo.IncomeDocuments.IncomeDocumentID", ConnectionStrings.LightConnectionString))
            {
                using (var cDT = new DataTable())
                {
                    cDA.Fill(cDT);

                    foreach (DataRow Row in IncomeDocumentsDataTable.Rows)
                    {
                        var Rows = cDT.Select("IncomeDocumentID = " + Row["IncomeDocumentID"]);

                        if (Rows.Count() > 0)
                            Row["CommentsCount"] = Rows[0]["ComCount"];
                        else
                            Row["CommentsCount"] = 0;
                    }
                }
            }


            using (var cDA = new SqlDataAdapter("SELECT infiniu2_light.dbo.IncomeDocuments.IncomeDocumentID, " +
                                                " COUNT(infiniu2_light.dbo.DocumentsRecipients.DocumentRecipientID) AS RecCount FROM infiniu2_light.dbo.DocumentsRecipients " +
                                                " LEFT OUTER JOIN infiniu2_light.dbo.IncomeDocuments ON infiniu2_light.dbo.DocumentsRecipients.DocumentID = " +
                                                "infiniu2_light.dbo.IncomeDocuments.IncomeDocumentID WHERE infiniu2_light.dbo.DocumentsRecipients.DocumentCategoryID = 1 " +
                                                "AND infiniu2_light.dbo.DocumentsRecipients.DocumentID IN " +
                                                "(SELECT infiniu2_light.dbo.IncomeDocuments.IncomeDocumentID" + Filter + ") GROUP BY " +
                                                "infiniu2_light.dbo.IncomeDocuments.IncomeDocumentID", ConnectionStrings.LightConnectionString))
            {
                using (var cDT = new DataTable())
                {
                    cDA.Fill(cDT);

                    foreach (DataRow Row in IncomeDocumentsDataTable.Rows)
                    {
                        var Rows = cDT.Select("IncomeDocumentID = " + Row["IncomeDocumentID"]);

                        if (Rows.Count() > 0)
                            Row["RecipientsCount"] = Rows[0]["RecCount"];
                        else
                            Row["RecipientsCount"] = 0;
                    }
                }
            }


            using (var cDA = new SqlDataAdapter("SELECT infiniu2_light.dbo.IncomeDocuments.IncomeDocumentID, " +
                                                " COUNT(infiniu2_light.dbo.DocumentsFiles.DocumentFileID) AS FileCount FROM infiniu2_light.dbo.DocumentsFiles " +
                                                " LEFT OUTER JOIN infiniu2_light.dbo.IncomeDocuments ON infiniu2_light.dbo.DocumentsFiles.DocumentID = " +
                                                "infiniu2_light.dbo.IncomeDocuments.IncomeDocumentID WHERE infiniu2_light.dbo.DocumentsFiles.DocumentCategoryID = 1 " +
                                                "AND infiniu2_light.dbo.DocumentsFiles.DocumentID IN " +
                                                "(SELECT infiniu2_light.dbo.IncomeDocuments.IncomeDocumentID" + Filter + ") GROUP BY " +
                                                "infiniu2_light.dbo.IncomeDocuments.IncomeDocumentID", ConnectionStrings.LightConnectionString))
            {
                using (var cDT = new DataTable())
                {
                    cDA.Fill(cDT);

                    foreach (DataRow Row in IncomeDocumentsDataTable.Rows)
                    {
                        var Rows = cDT.Select("IncomeDocumentID = " + Row["IncomeDocumentID"]);

                        if (Rows.Count() > 0)
                            Row["FilesCount"] = Rows[0]["FileCount"];
                        else
                            Row["FilesCount"] = 0;
                    }
                }
            }
        }

        public void FillOuterDocuments(string Filter)
        {
            OuterDocumentsDataTable.Clear();

            var sNewFilter = "";

            if (Filter.Length > 0)
                sNewFilter = " WHERE " + Filter + " ORDER BY OuterDocumentID DESC";
            else
                sNewFilter = " ORDER BY OuterDocumentID DESC";

            using (var DA = new SqlDataAdapter("SELECT OuterDocuments.OuterDocumentID AS DocumentID, OuterDocuments.* FROM OuterDocuments" + sNewFilter, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(OuterDocumentsDataTable);
                }
            }

            using (var cDA = new SqlDataAdapter("SELECT infiniu2_light.dbo.OuterDocuments.OuterDocumentID, " +
                                                " COUNT(infiniu2_light.dbo.DocumentsComments.DocumentCommentID) AS ComCount FROM infiniu2_light.dbo.DocumentsComments " +
                                                " LEFT OUTER JOIN infiniu2_light.dbo.OuterDocuments ON infiniu2_light.dbo.DocumentsComments.DocumentID = " +
                                                "infiniu2_light.dbo.OuterDocuments.OuterDocumentID WHERE infiniu2_light.dbo.DocumentsComments.DocumentCategoryID = 2 " +
                                                "AND infiniu2_light.dbo.DocumentsComments.DocumentID IN " +
                                                "(SELECT infiniu2_light.dbo.OuterDocuments.OuterDocumentID" + Filter + ") GROUP BY " +
                                                "infiniu2_light.dbo.OuterDocuments.OuterDocumentID", ConnectionStrings.LightConnectionString))
            {
                using (var cDT = new DataTable())
                {
                    cDA.Fill(cDT);

                    foreach (DataRow Row in OuterDocumentsDataTable.Rows)
                    {
                        var Rows = cDT.Select("OuterDocumentID = " + Row["OuterDocumentID"]);

                        if (Rows.Count() > 0)
                            Row["CommentsCount"] = Rows[0]["ComCount"];
                        else
                            Row["CommentsCount"] = 0;
                    }
                }
            }


            using (var cDA = new SqlDataAdapter("SELECT infiniu2_light.dbo.OuterDocuments.OuterDocumentID, " +
                                                " COUNT(infiniu2_light.dbo.DocumentsRecipients.DocumentRecipientID) AS RecCount FROM infiniu2_light.dbo.DocumentsRecipients " +
                                                " LEFT OUTER JOIN infiniu2_light.dbo.OuterDocuments ON infiniu2_light.dbo.DocumentsRecipients.DocumentID = " +
                                                "infiniu2_light.dbo.OuterDocuments.OuterDocumentID WHERE infiniu2_light.dbo.DocumentsRecipients.DocumentCategoryID = 2 " +
                                                "AND infiniu2_light.dbo.DocumentsRecipients.DocumentID IN " +
                                                "(SELECT infiniu2_light.dbo.OuterDocuments.OuterDocumentID" + Filter + ") GROUP BY " +
                                                "infiniu2_light.dbo.OuterDocuments.OuterDocumentID", ConnectionStrings.LightConnectionString))
            {
                using (var cDT = new DataTable())
                {
                    cDA.Fill(cDT);

                    foreach (DataRow Row in OuterDocumentsDataTable.Rows)
                    {
                        var Rows = cDT.Select("OuterDocumentID = " + Row["OuterDocumentID"]);

                        if (Rows.Count() > 0)
                            Row["RecipientsCount"] = Rows[0]["RecCount"];
                        else
                            Row["RecipientsCount"] = 0;
                    }
                }
            }


            using (var cDA = new SqlDataAdapter("SELECT infiniu2_light.dbo.OuterDocuments.OuterDocumentID, " +
                                                " COUNT(infiniu2_light.dbo.DocumentsFiles.DocumentFileID) AS FileCount FROM infiniu2_light.dbo.DocumentsFiles " +
                                                " LEFT OUTER JOIN infiniu2_light.dbo.OuterDocuments ON infiniu2_light.dbo.DocumentsFiles.DocumentID = " +
                                                "infiniu2_light.dbo.OuterDocuments.OuterDocumentID WHERE infiniu2_light.dbo.DocumentsFiles.DocumentCategoryID = 2 " +
                                                "AND infiniu2_light.dbo.DocumentsFiles.DocumentID IN " +
                                                "(SELECT infiniu2_light.dbo.OuterDocuments.OuterDocumentID" + Filter + ") GROUP BY " +
                                                "infiniu2_light.dbo.OuterDocuments.OuterDocumentID", ConnectionStrings.LightConnectionString))
            {
                using (var cDT = new DataTable())
                {
                    cDA.Fill(cDT);

                    foreach (DataRow Row in OuterDocumentsDataTable.Rows)
                    {
                        var Rows = cDT.Select("OuterDocumentID = " + Row["OuterDocumentID"]);

                        if (Rows.Count() > 0)
                            Row["FilesCount"] = Rows[0]["FileCount"];
                        else
                            Row["FilesCount"] = 0;
                    }
                }
            }
        }

        public DataTable GetDocumentsRecipients(int DocumentCategoryID, int DocumentID)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsRecipients WHERE DocumentCategoryID = " + DocumentCategoryID + " AND DocumentID = " + DocumentID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    return DT;
                }
            }
        }

        public void TestMoveFiles()
        {
            using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM DocumentsFiles", ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);


                        using (var sDA = new SqlDataAdapter("SELECT * FROM IncomeDocumentsFiles", ConnectionStrings.LightConnectionString))
                        {
                            using (var sDT = new DataTable())
                            {
                                sDA.Fill(sDT);

                                foreach (DataRow Row in sDT.Rows)
                                {
                                    var NewRow = DT.NewRow();
                                    NewRow["FileName"] = Row["FileName"];
                                    NewRow["FileSize"] = Row["FileSize"];
                                    NewRow["DocumentCategoryID"] = 1;
                                    NewRow["DocumentID"] = Row["IncomeDocumentID"];
                                    DT.Rows.Add(NewRow);
                                }

                                DA.Update(DT);
                            }
                        }

                    }
                }
            }
        }

        public void TestMoveRecipients()
        {
            using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM DocumentsRecipients", ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);


                        using (var sDA = new SqlDataAdapter("SELECT * FROM OuterDocumentsRecipients", ConnectionStrings.LightConnectionString))
                        {
                            using (var sDT = new DataTable())
                            {
                                sDA.Fill(sDT);

                                foreach (DataRow Row in sDT.Rows)
                                {
                                    var NewRow = DT.NewRow();
                                    NewRow["DocumentID"] = Row["OuterDocumentID"];
                                    NewRow["DocumentCategoryID"] = 2;
                                    NewRow["UserID"] = Row["UserID"];
                                    DT.Rows.Add(NewRow);
                                }

                                DA.Update(DT);
                            }
                        }

                    }
                }
            }
        }

        public void FillUpdates(int DateType)//0 - new, 1 - today, 2 - week, 3 - month
        {
            UpdatesDocumentsDataTable.Clear();
            UpdatesCommentsDataTable.Clear();
            UpdatesFilesDataTable.Clear();
            UpdatesRecipientsDataTable.Clear();
            UpdatesCommentsFilesDataTable.Clear();
            UpdatesConfirmsDataTable.Clear();
            UpdatesConfirmsRecipientsDataTable.Clear();

            var DateTime = "";

            if (DateType == 0)
            {
                using (var DA = new SqlDataAdapter("SELECT InnerDocuments.*, InnerDocumentID AS DocumentID FROM InnerDocuments WHERE InnerDocumentID IN " +
                                                   "(SELECT TableItemID FROM SubscribesRecords WHERE SubscribesItemID = 18 AND UserID = " + Security.CurrentUserID + ")" +
                                                   "OR InnerDocumentID IN (SELECT DocumentID FROM infiniu2_light.dbo.DocumentsComments WHERE DocumentCategoryID = 0 AND DocumentCommentID IN " +
                                                   "(SELECT TableItemID FROM SubscribesRecords WHERE SubscribesItemID = 21 AND UserID = " + Security.CurrentUserID + "))", ConnectionStrings.LightConnectionString))
                {
                    DA.Fill(UpdatesDocumentsDataTable);
                }

                using (var DA = new SqlDataAdapter("SELECT IncomeDocuments.*, IncomeDocumentID AS DocumentID FROM IncomeDocuments WHERE IncomeDocumentID IN " +
                                                   "(SELECT TableItemID FROM SubscribesRecords WHERE SubscribesItemID = 19 AND UserID = " + Security.CurrentUserID + ")" +
                                                   "OR IncomeDocumentID IN (SELECT DocumentID FROM infiniu2_light.dbo.DocumentsComments WHERE DocumentCategoryID = 1 AND DocumentCommentID IN " +
                                                   "(SELECT TableItemID FROM SubscribesRecords WHERE SubscribesItemID = 21 AND UserID = " + Security.CurrentUserID + "))", ConnectionStrings.LightConnectionString))
                {
                    DA.Fill(UpdatesDocumentsDataTable);
                }

                using (var DA = new SqlDataAdapter("SELECT OuterDocuments.*, OuterDocumentID AS DocumentID FROM OuterDocuments WHERE OuterDocumentID IN " +
                                                   "(SELECT TableItemID FROM SubscribesRecords WHERE SubscribesItemID = 20 AND UserID = " + Security.CurrentUserID + ")" +
                                                   "OR OuterDocumentID IN (SELECT DocumentID FROM infiniu2_light.dbo.DocumentsComments WHERE DocumentCategoryID = 2 AND DocumentCommentID IN " +
                                                   "(SELECT TableItemID FROM SubscribesRecords WHERE SubscribesItemID = 21 AND UserID = " + Security.CurrentUserID + "))", ConnectionStrings.LightConnectionString))
                {
                    DA.Fill(UpdatesDocumentsDataTable);
                }

                using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsComments WHERE DocumentCommentID IN " +
                                                   "(SELECT TableItemID FROM SubscribesRecords WHERE SubscribesItemID = 21 AND UserID = " + Security.CurrentUserID + ")" +
                                                   "OR DocumentID IN (SELECT InnerDocumentID FROM InnerDocuments WHERE InnerDocumentID IN (SELECT TableItemID FROM SubscribesRecords WHERE UserID = " + Security.CurrentUserID + " AND SubscribesItemID = 18) OR InnerDocumentID IN (SELECT DocumentID FROM DocumentsComments WHERE DocumentCategoryID = 0 AND DocumentCommentID IN(SELECT TableItemID FROM SubscribesRecords WHERE SubscribesItemID = 21 AND UserID = " + Security.CurrentUserID + ")))" +
                                                   "OR DocumentID IN (SELECT IncomeDocumentID FROM IncomeDocuments WHERE IncomeDocumentID IN (SELECT TableItemID FROM SubscribesRecords WHERE UserID = " + Security.CurrentUserID + " AND SubscribesItemID = 19) OR IncomeDocumentID IN (SELECT DocumentID FROM DocumentsComments WHERE DocumentCategoryID = 1 AND DocumentCommentID IN(SELECT TableItemID FROM SubscribesRecords WHERE SubscribesItemID = 21 AND UserID = " + Security.CurrentUserID + ")))" +
                                                   "OR DocumentID IN (SELECT OuterDocumentID FROM OuterDocuments WHERE OuterDocumentID IN (SELECT TableItemID FROM SubscribesRecords WHERE UserID = " + Security.CurrentUserID + " AND SubscribesItemID = 20) OR OuterDocumentID IN (SELECT DocumentID FROM DocumentsComments WHERE DocumentCategoryID = 2 AND DocumentCommentID IN(SELECT TableItemID FROM SubscribesRecords WHERE SubscribesItemID = 21 AND UserID = " + Security.CurrentUserID + ")))" +
                                                   " ORDER BY DocumentCommentID DESC", ConnectionStrings.LightConnectionString))
                {
                    DA.Fill(UpdatesCommentsDataTable);

                    foreach (DataRow Row in UpdatesDocumentsDataTable.Rows)
                    {
                        var uRow = UpdatesCommentsDataTable.Select("DocumentID = " + Row["DocumentID"] + " AND DateTime > '" + Convert.ToDateTime(Row["DateTime"]) + "'");

                        if (uRow.Count() > 0)
                        {
                            var date = Convert.ToDateTime(Row["DateTime"]);

                            foreach (var uuRow in uRow)
                            {
                                if (date < Convert.ToDateTime(uuRow["DateTime"]))
                                    date = Convert.ToDateTime(uuRow["DateTime"]);
                            }

                            Row["UpdateDateTime"] = date;
                        }
                        else
                            Row["UpdateDateTime"] = Row["DateTime"];
                    }
                }

                UpdatesDocumentsDataTable.DefaultView.Sort = "UpdateDateTime DESC";

                using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsCommentsFiles WHERE DocumentCommentID IN (SELECT DocumentCommentID FROM DocumentsComments WHERE DocumentCommentID IN " +
                                                   "(SELECT TableItemID FROM SubscribesRecords WHERE SubscribesItemID = 21 AND UserID = " + Security.CurrentUserID + ")" +
                                                   "OR DocumentID IN (SELECT InnerDocumentID FROM InnerDocuments WHERE InnerDocumentID IN (SELECT TableItemID FROM SubscribesRecords WHERE UserID = " + Security.CurrentUserID + " AND SubscribesItemID = 18) OR InnerDocumentID IN (SELECT DocumentID FROM DocumentsComments WHERE DocumentCategoryID = 0 AND DocumentCommentID IN(SELECT TableItemID FROM SubscribesRecords WHERE SubscribesItemID = 21 AND UserID = " + Security.CurrentUserID + ")))" +
                                                   "OR DocumentID IN (SELECT IncomeDocumentID FROM IncomeDocuments WHERE IncomeDocumentID IN (SELECT TableItemID FROM SubscribesRecords WHERE UserID = " + Security.CurrentUserID + " AND SubscribesItemID = 19) OR IncomeDocumentID IN (SELECT DocumentID FROM DocumentsComments WHERE DocumentCategoryID = 1 AND DocumentCommentID IN(SELECT TableItemID FROM SubscribesRecords WHERE SubscribesItemID = 21 AND UserID = " + Security.CurrentUserID + ")))" +
                                                   "OR DocumentID IN (SELECT OuterDocumentID FROM OuterDocuments WHERE OuterDocumentID IN (SELECT TableItemID FROM SubscribesRecords WHERE UserID = " + Security.CurrentUserID + " AND SubscribesItemID = 20) OR OuterDocumentID IN (SELECT DocumentID FROM DocumentsComments WHERE DocumentCategoryID = 2 AND DocumentCommentID IN(SELECT TableItemID FROM SubscribesRecords WHERE SubscribesItemID = 21 AND UserID = " + Security.CurrentUserID + "))))", ConnectionStrings.LightConnectionString))
                {
                    DA.Fill(UpdatesCommentsFilesDataTable);
                }

                using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsFiles WHERE " +
                                                   "(DocumentCategoryID = 0 AND DocumentID IN (SELECT InnerDocumentID FROM InnerDocuments WHERE InnerDocumentID IN" +
                                                   "(SELECT TableItemID FROM SubscribesRecords WHERE SubscribesItemID = 18 AND UserID = " + Security.CurrentUserID + ")" +
                                                   "OR InnerDocumentID IN (SELECT DocumentID FROM infiniu2_light.dbo.DocumentsComments WHERE DocumentCategoryID = 0 AND DocumentCommentID IN " +
                                                   "(SELECT TableItemID FROM SubscribesRecords WHERE SubscribesItemID = 21 AND UserID = " + Security.CurrentUserID + "))))" +

                                                   "OR (DocumentCategoryID = 1 AND DocumentID IN (SELECT IncomeDocumentID FROM IncomeDocuments WHERE IncomeDocumentID IN" +
                                                   "(SELECT TableItemID FROM SubscribesRecords WHERE SubscribesItemID = 19 AND UserID = " + Security.CurrentUserID + ")" +
                                                   "OR IncomeDocumentID IN (SELECT DocumentID FROM infiniu2_light.dbo.DocumentsComments WHERE DocumentCategoryID = 1 AND DocumentCommentID IN " +
                                                   "(SELECT TableItemID FROM SubscribesRecords WHERE SubscribesItemID = 21 AND UserID = " + Security.CurrentUserID + "))))" +

                                                   "OR (DocumentCategoryID = 2 AND DocumentID IN (SELECT OuterDocumentID FROM OuterDocuments WHERE OuterDocumentID IN" +
                                                   "(SELECT TableItemID FROM SubscribesRecords WHERE SubscribesItemID = 20 AND UserID = " + Security.CurrentUserID + ")" +
                                                   "OR OuterDocumentID IN (SELECT DocumentID FROM infiniu2_light.dbo.DocumentsComments WHERE DocumentCategoryID = 2 AND DocumentCommentID IN " +
                                                   "(SELECT TableItemID FROM SubscribesRecords WHERE SubscribesItemID = 21 AND UserID = " + Security.CurrentUserID + "))))", ConnectionStrings.LightConnectionString))
                {
                    DA.Fill(UpdatesFilesDataTable);
                }


                using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsRecipients WHERE " +
                                                   "(DocumentCategoryID = 0 AND DocumentID IN " +
                                                   "(SELECT InnerDocumentID FROM InnerDocuments WHERE InnerDocumentID IN" +
                                                   "(SELECT TableItemID FROM SubscribesRecords WHERE SubscribesItemID = 18 AND UserID = " + Security.CurrentUserID +
                                                   ") OR InnerDocumentID IN (SELECT DocumentID FROM infiniu2_light.dbo.DocumentsComments WHERE DocumentCategoryID = 0 AND DocumentCommentID IN " +
                                                   "(SELECT TableItemID FROM SubscribesRecords WHERE SubscribesItemID = 21 AND UserID = " + Security.CurrentUserID + "))))" +
                                                   " OR " +

                                                   "(DocumentCategoryID = 1 AND DocumentID IN " +
                                                   "(SELECT IncomeDocumentID FROM IncomeDocuments WHERE IncomeDocumentID IN" +
                                                   "(SELECT TableItemID FROM SubscribesRecords WHERE SubscribesItemID = 19 AND UserID = " + Security.CurrentUserID +
                                                   ") OR IncomeDocumentID IN (SELECT DocumentID FROM infiniu2_light.dbo.DocumentsComments WHERE DocumentCategoryID = 1 AND DocumentCommentID IN " +
                                                   "(SELECT TableItemID FROM SubscribesRecords WHERE SubscribesItemID = 21 AND UserID = " + Security.CurrentUserID + "))))" +
                                                   " OR " +

                                                   "(DocumentCategoryID = 2 AND DocumentID IN " +
                                                   "(SELECT OuterDocumentID FROM OuterDocuments WHERE OuterDocumentID IN" +
                                                   "(SELECT TableItemID FROM SubscribesRecords WHERE SubscribesItemID = 20 AND UserID = " + Security.CurrentUserID +
                                                   ") OR OuterDocumentID IN (SELECT DocumentID FROM infiniu2_light.dbo.DocumentsComments WHERE DocumentCategoryID = 2 AND DocumentCommentID IN " +
                                                   "(SELECT TableItemID FROM SubscribesRecords WHERE SubscribesItemID = 21 AND UserID = " + Security.CurrentUserID + "))))"


                , ConnectionStrings.LightConnectionString))
                {
                    DA.Fill(UpdatesRecipientsDataTable);
                }


                using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsConfirmations WHERE " +
                                                   "(DocumentCategoryID = 0 AND DocumentID IN " +
                                                   "(SELECT InnerDocumentID FROM InnerDocuments WHERE InnerDocumentID IN" +
                                                   "(SELECT TableItemID FROM SubscribesRecords WHERE SubscribesItemID = 18 AND UserID = " + Security.CurrentUserID +
                                                   ") OR InnerDocumentID IN (SELECT DocumentID FROM infiniu2_light.dbo.DocumentsComments WHERE DocumentCategoryID = 0 AND DocumentCommentID IN " +
                                                   "(SELECT TableItemID FROM SubscribesRecords WHERE SubscribesItemID = 21 AND UserID = " + Security.CurrentUserID + "))))" +
                                                   " OR " +

                                                   "(DocumentCategoryID = 1 AND DocumentID IN " +
                                                   "(SELECT IncomeDocumentID FROM IncomeDocuments WHERE IncomeDocumentID IN" +
                                                   "(SELECT TableItemID FROM SubscribesRecords WHERE SubscribesItemID = 19 AND UserID = " + Security.CurrentUserID +
                                                   ") OR IncomeDocumentID IN (SELECT DocumentID FROM infiniu2_light.dbo.DocumentsComments WHERE DocumentCategoryID = 1 AND DocumentCommentID IN " +
                                                   "(SELECT TableItemID FROM SubscribesRecords WHERE SubscribesItemID = 21 AND UserID = " + Security.CurrentUserID + "))))" +
                                                   " OR " +

                                                   "(DocumentCategoryID = 2 AND DocumentID IN " +
                                                   "(SELECT OuterDocumentID FROM OuterDocuments WHERE OuterDocumentID IN" +
                                                   "(SELECT TableItemID FROM SubscribesRecords WHERE SubscribesItemID = 20 AND UserID = " + Security.CurrentUserID +
                                                   ") OR OuterDocumentID IN (SELECT DocumentID FROM infiniu2_light.dbo.DocumentsComments WHERE DocumentCategoryID = 2 AND DocumentCommentID IN " +
                                                   "(SELECT TableItemID FROM SubscribesRecords WHERE SubscribesItemID = 21 AND UserID = " + Security.CurrentUserID + "))))"


                , ConnectionStrings.LightConnectionString))
                {
                    DA.Fill(UpdatesConfirmsDataTable);
                }


                using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsConfirmationsRecipients WHERE DocumentConfirmationID IN (SELECT DocumentConfirmationID FROM DocumentsConfirmations WHERE " +
                                                   "(DocumentCategoryID = 0 AND DocumentID IN " +
                                                   "(SELECT InnerDocumentID FROM InnerDocuments WHERE InnerDocumentID IN" +
                                                   "(SELECT TableItemID FROM SubscribesRecords WHERE SubscribesItemID = 18 AND UserID = " + Security.CurrentUserID +
                                                   ") OR InnerDocumentID IN (SELECT DocumentID FROM infiniu2_light.dbo.DocumentsComments WHERE DocumentCategoryID = 0 AND DocumentCommentID IN " +
                                                   "(SELECT TableItemID FROM SubscribesRecords WHERE SubscribesItemID = 21 AND UserID = " + Security.CurrentUserID + "))))" +
                                                   " OR " +

                                                   "(DocumentCategoryID = 1 AND DocumentID IN " +
                                                   "(SELECT IncomeDocumentID FROM IncomeDocuments WHERE IncomeDocumentID IN" +
                                                   "(SELECT TableItemID FROM SubscribesRecords WHERE SubscribesItemID = 19 AND UserID = " + Security.CurrentUserID +
                                                   ") OR IncomeDocumentID IN (SELECT DocumentID FROM infiniu2_light.dbo.DocumentsComments WHERE DocumentCategoryID = 1 AND DocumentCommentID IN " +
                                                   "(SELECT TableItemID FROM SubscribesRecords WHERE SubscribesItemID = 21 AND UserID = " + Security.CurrentUserID + "))))" +
                                                   " OR " +

                                                   "(DocumentCategoryID = 2 AND DocumentID IN " +
                                                   "(SELECT OuterDocumentID FROM OuterDocuments WHERE OuterDocumentID IN" +
                                                   "(SELECT TableItemID FROM SubscribesRecords WHERE SubscribesItemID = 20 AND UserID = " + Security.CurrentUserID +
                                                   ") OR OuterDocumentID IN (SELECT DocumentID FROM infiniu2_light.dbo.DocumentsComments WHERE DocumentCategoryID = 2 AND DocumentCommentID IN " +
                                                   "(SELECT TableItemID FROM SubscribesRecords WHERE SubscribesItemID = 21 AND UserID = " + Security.CurrentUserID + ")))))"


                , ConnectionStrings.LightConnectionString))
                {
                    DA.Fill(UpdatesConfirmsRecipientsDataTable);
                }


                return;
            }






            if (DateType == 1)
                DateTime = "(CAST(DateTime AS DATE) = '" + Security.GetCurrentDate().ToString("yyyy-MM-dd") + "')";

            if (DateType == 2)
                DateTime = "(CAST(DateTime AS DATE) >= '" + Security.GetCurrentDate().AddDays(-6).ToString("yyyy-MM-dd") + "') AND (CAST(DateTime AS DATE) <= '" + Security.GetCurrentDate().ToString("yyyy-MM-dd") + "')";

            if (DateType == 3)
                DateTime = "(CAST(DateTime AS DATE) >= '" + Security.GetCurrentDate().AddDays(-30).ToString("yyyy-MM-dd") + "') AND (CAST(DateTime AS DATE) <= '" + Security.GetCurrentDate().ToString("yyyy-MM-dd") + "')";


            using (var DA = new SqlDataAdapter("SELECT InnerDocuments.*, InnerDocumentID AS DocumentID FROM InnerDocuments WHERE " + DateTime + "" +
                                               "OR InnerDocumentID IN (SELECT DocumentID FROM infiniu2_light.dbo.DocumentsComments WHERE " + DateTime + " AND DocumentCategoryID = 0)", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(UpdatesDocumentsDataTable);
            }

            using (var DA = new SqlDataAdapter("SELECT IncomeDocuments.*, IncomeDocumentID AS DocumentID FROM IncomeDocuments WHERE " + DateTime + "" +
                                               "OR IncomeDocumentID IN (SELECT DocumentID FROM infiniu2_light.dbo.DocumentsComments WHERE " + DateTime + " AND DocumentCategoryID = 1)", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(UpdatesDocumentsDataTable);
            }

            using (var DA = new SqlDataAdapter("SELECT OuterDocuments.*, OuterDocumentID AS DocumentID FROM OuterDocuments WHERE " + DateTime + "" +
                                               "OR OuterDocumentID IN (SELECT DocumentID FROM infiniu2_light.dbo.DocumentsComments WHERE " + DateTime + " AND DocumentCategoryID = 2)", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(UpdatesDocumentsDataTable);
            }

            using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsComments WHERE " + DateTime +
                                               " OR (DocumentCategoryID = 0 AND DocumentID IN (SELECT InnerDocumentID FROM InnerDocuments WHERE " + DateTime + " OR InnerDocumentID IN (SELECT DocumentID FROM infiniu2_light.dbo.DocumentsComments WHERE " + DateTime + " AND DocumentCategoryID = 0)))" +
                                               " OR (DocumentCategoryID = 1 AND DocumentID IN (SELECT IncomeDocumentID FROM IncomeDocuments WHERE " + DateTime + " OR IncomeDocumentID IN (SELECT DocumentID FROM infiniu2_light.dbo.DocumentsComments WHERE " + DateTime + " AND DocumentCategoryID = 1)))" +
                                               " OR (DocumentCategoryID = 2 AND DocumentID IN (SELECT OuterDocumentID FROM OuterDocuments WHERE " + DateTime + " OR OuterDocumentID IN (SELECT DocumentID FROM infiniu2_light.dbo.DocumentsComments WHERE " + DateTime + " AND DocumentCategoryID = 2))) ORDER BY DocumentCommentID DESC", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(UpdatesCommentsDataTable);

                foreach (DataRow Row in UpdatesDocumentsDataTable.Rows)
                {
                    var uRow = UpdatesCommentsDataTable.Select("DocumentID = " + Row["DocumentID"] + " AND DateTime > '" + Convert.ToDateTime(Row["DateTime"]) + "'");

                    if (uRow.Count() > 0)
                    {
                        var date = Convert.ToDateTime(Row["DateTime"]);

                        foreach (var uuRow in uRow)
                        {
                            if (date < Convert.ToDateTime(uuRow["DateTime"]))
                                date = Convert.ToDateTime(uuRow["DateTime"]);
                        }

                        Row["UpdateDateTime"] = date;
                    }
                    else
                        Row["UpdateDateTime"] = Row["DateTime"];
                }
            }

            UpdatesDocumentsDataTable.DefaultView.Sort = "UpdateDateTime DESC";

            using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsCommentsFiles WHERE DocumentCommentID IN (SELECT DocumentCommentID FROM DocumentsComments WHERE " + DateTime +
                                               " OR (DocumentCategoryID = 0 AND DocumentID IN (SELECT InnerDocumentID FROM InnerDocuments WHERE " + DateTime + " OR InnerDocumentID IN (SELECT DocumentID FROM infiniu2_light.dbo.DocumentsComments WHERE " + DateTime + " AND DocumentCategoryID = 0)))" +
                                               " OR (DocumentCategoryID = 1 AND DocumentID IN (SELECT IncomeDocumentID FROM IncomeDocuments WHERE " + DateTime + " OR IncomeDocumentID IN (SELECT DocumentID FROM infiniu2_light.dbo.DocumentsComments WHERE " + DateTime + " AND DocumentCategoryID = 1)))" +
                                               " OR (DocumentCategoryID = 2 AND DocumentID IN (SELECT OuterDocumentID FROM OuterDocuments WHERE " + DateTime + " OR OuterDocumentID IN (SELECT DocumentID FROM infiniu2_light.dbo.DocumentsComments WHERE " + DateTime + " AND DocumentCategoryID = 2)))) ORDER BY DocumentCommentID DESC", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(UpdatesCommentsFilesDataTable);
            }

            using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsFiles WHERE " +
                                               "(DocumentCategoryID = 0 AND DocumentID IN " +

                                               "(SELECT InnerDocumentID FROM InnerDocuments WHERE InnerDocumentID IN " +
                                               "(SELECT InnerDocumentID FROM infiniu2_light.dbo.InnerDocuments WHERE " + DateTime +
                                               "OR InnerDocumentID IN (SELECT DocumentID FROM infiniu2_light.dbo.DocumentsComments WHERE " + DateTime + ")))) OR " +

                                               "(DocumentCategoryID = 1 AND DocumentID IN " +

                                               "(SELECT IncomeDocumentID FROM IncomeDocuments WHERE IncomeDocumentID IN " +
                                               "(SELECT IncomeDocumentID FROM infiniu2_light.dbo.IncomeDocuments WHERE " + DateTime +
                                               "OR IncomeDocumentID IN (SELECT DocumentID FROM infiniu2_light.dbo.DocumentsComments WHERE " + DateTime + ")))) OR " +

                                               "(DocumentCategoryID = 2 AND DocumentID IN " +

                                               "(SELECT OuterDocumentID FROM OuterDocuments WHERE OuterDocumentID IN " +
                                               "(SELECT OuterDocumentID FROM infiniu2_light.dbo.OuterDocuments WHERE " + DateTime +
                                               "OR OuterDocumentID IN (SELECT DocumentID FROM infiniu2_light.dbo.DocumentsComments WHERE " + DateTime + "))))"
                                                                , ConnectionStrings.LightConnectionString))
            {
                DA.Fill(UpdatesFilesDataTable);
            }


            using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsRecipients WHERE " +
                                               "((DocumentCategoryID = 2 AND DocumentID IN (SELECT OuterDocumentID FROM infiniu2_light.dbo.OuterDocuments WHERE " + DateTime + "))" +
                                               " OR (DocumentCategoryID = 2 AND DocumentID IN (SELECT DocumentID FROM infiniu2_light.dbo.DocumentsComments WHERE DocumentCategoryID = 2 AND " + DateTime + ")))" +

                                               " OR ((DocumentCategoryID = 1 AND DocumentID IN (SELECT IncomeDocumentID FROM infiniu2_light.dbo.IncomeDocuments WHERE " + DateTime + "))" +
                                               " OR (DocumentCategoryID = 1 AND DocumentID IN (SELECT DocumentID FROM infiniu2_light.dbo.DocumentsComments WHERE DocumentCategoryID = 1 AND " + DateTime + ")))" +

                                               " OR ((DocumentCategoryID = 0 AND DocumentID IN (SELECT InnerDocumentID FROM infiniu2_light.dbo.InnerDocuments WHERE " + DateTime + "))" +
                                               " OR (DocumentCategoryID = 0 AND DocumentID IN (SELECT DocumentID FROM infiniu2_light.dbo.DocumentsComments WHERE DocumentCategoryID = 0 AND " + DateTime + ")))"
                                                                , ConnectionStrings.LightConnectionString))
            {
                DA.Fill(UpdatesRecipientsDataTable);
            }

            using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsConfirmations WHERE " +
                                               "((DocumentCategoryID = 2 AND DocumentID IN (SELECT OuterDocumentID FROM infiniu2_light.dbo.OuterDocuments WHERE " + DateTime + "))" +
                                               " OR (DocumentCategoryID = 2 AND DocumentID IN (SELECT DocumentID FROM infiniu2_light.dbo.DocumentsComments WHERE DocumentCategoryID = 2 AND " + DateTime + ")))" +

                                               " OR ((DocumentCategoryID = 1 AND DocumentID IN (SELECT IncomeDocumentID FROM infiniu2_light.dbo.IncomeDocuments WHERE " + DateTime + "))" +
                                               " OR (DocumentCategoryID = 1 AND DocumentID IN (SELECT DocumentID FROM infiniu2_light.dbo.DocumentsComments WHERE DocumentCategoryID = 1 AND " + DateTime + ")))" +

                                               " OR ((DocumentCategoryID = 0 AND DocumentID IN (SELECT InnerDocumentID FROM infiniu2_light.dbo.InnerDocuments WHERE " + DateTime + "))" +
                                               " OR (DocumentCategoryID = 0 AND DocumentID IN (SELECT DocumentID FROM infiniu2_light.dbo.DocumentsComments WHERE DocumentCategoryID = 0 AND " + DateTime + ")))"
                                                                , ConnectionStrings.LightConnectionString))
            {
                DA.Fill(UpdatesConfirmsDataTable);
            }

            using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsConfirmationsRecipients WHERE DocumentConfirmationID IN (SELECT DocumentConfirmationID FROM DocumentsConfirmations WHERE" +
                                               "((DocumentCategoryID = 2 AND DocumentID IN (SELECT OuterDocumentID FROM infiniu2_light.dbo.OuterDocuments WHERE " + DateTime + "))" +
                                               " OR (DocumentCategoryID = 2 AND DocumentID IN (SELECT DocumentID FROM infiniu2_light.dbo.DocumentsComments WHERE DocumentCategoryID = 2 AND " + DateTime + ")))" +

                                               " OR ((DocumentCategoryID = 1 AND DocumentID IN (SELECT IncomeDocumentID FROM infiniu2_light.dbo.IncomeDocuments WHERE " + DateTime + "))" +
                                               " OR (DocumentCategoryID = 1 AND DocumentID IN (SELECT DocumentID FROM infiniu2_light.dbo.DocumentsComments WHERE DocumentCategoryID = 1 AND " + DateTime + ")))" +

                                               " OR ((DocumentCategoryID = 0 AND DocumentID IN (SELECT InnerDocumentID FROM infiniu2_light.dbo.InnerDocuments WHERE " + DateTime + "))" +
                                               " OR (DocumentCategoryID = 0 AND DocumentID IN (SELECT DocumentID FROM infiniu2_light.dbo.DocumentsComments WHERE DocumentCategoryID = 0 AND " + DateTime + "))))"
                                                                , ConnectionStrings.LightConnectionString))
            {
                DA.Fill(UpdatesConfirmsRecipientsDataTable);
            }
        }

        public void FillConfirmUpdates(int StatusType)//0 - not confirmed, 1 - confirmed, 2 - canceled, 3 - your sign
        {
            UpdatesDocumentsDataTable.Clear();
            UpdatesCommentsDataTable.Clear();
            UpdatesFilesDataTable.Clear();
            UpdatesRecipientsDataTable.Clear();
            UpdatesCommentsFilesDataTable.Clear();
            UpdatesConfirmsDataTable.Clear();
            UpdatesConfirmsRecipientsDataTable.Clear();

            if (StatusType == 3)
            {
                using (var DA = new SqlDataAdapter("SELECT InnerDocuments.*, InnerDocuments.InnerDocumentID AS DocumentID FROM InnerDocuments WHERE InnerDocumentID IN (SELECT DocumentID FROM DocumentsConfirmations WHERE DocumentCategoryID = 0 AND DocumentConfirmationID IN (SELECT DocumentConfirmationID FROM DocumentsConfirmationsRecipients WHERE Status = 0 AND UserID = " + Security.CurrentUserID + "))", ConnectionStrings.LightConnectionString))
                {
                    DA.Fill(UpdatesDocumentsDataTable);
                }

                using (var DA = new SqlDataAdapter("SELECT IncomeDocuments.*, IncomeDocuments.IncomeDocumentID AS DocumentID FROM IncomeDocuments WHERE IncomeDocumentID IN (SELECT DocumentID FROM DocumentsConfirmations WHERE DocumentCategoryID = 1 AND DocumentConfirmationID IN (SELECT DocumentConfirmationID FROM DocumentsConfirmationsRecipients WHERE Status = 0 AND UserID = " + Security.CurrentUserID + "))", ConnectionStrings.LightConnectionString))
                {
                    DA.Fill(UpdatesDocumentsDataTable);
                }

                using (var DA = new SqlDataAdapter("SELECT OuterDocuments.*, OuterDocuments.OuterDocumentID AS DocumentID FROM OuterDocuments WHERE OuterDocumentID IN (SELECT DocumentID FROM DocumentsConfirmations WHERE DocumentCategoryID = 2 AND DocumentConfirmationID IN (SELECT DocumentConfirmationID FROM DocumentsConfirmationsRecipients WHERE Status = 0 AND UserID = " + Security.CurrentUserID + "))", ConnectionStrings.LightConnectionString))
                {
                    DA.Fill(UpdatesDocumentsDataTable);
                }

                using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsComments WHERE " +
                                                   "(DocumentCategoryID = 0 AND DocumentID IN (SELECT InnerDocumentID FROM InnerDocuments WHERE InnerDocumentID IN (SELECT DocumentID FROM DocumentsConfirmations WHERE DocumentCategoryID = 0 AND DocumentConfirmationID IN (SELECT DocumentConfirmationID FROM DocumentsConfirmationsRecipients WHERE Status = 0 AND UserID = " + Security.CurrentUserID + ")))) OR" +
                                                   "(DocumentCategoryID = 1 AND DocumentID IN (SELECT IncomeDocumentID FROM IncomeDocuments WHERE IncomeDocumentID IN (SELECT DocumentID FROM DocumentsConfirmations WHERE DocumentCategoryID = 1 AND DocumentConfirmationID IN (SELECT DocumentConfirmationID FROM DocumentsConfirmationsRecipients WHERE Status = 0 AND UserID = " + Security.CurrentUserID + ")))) OR" +
                                                   "(DocumentCategoryID = 2 AND DocumentID IN (SELECT OuterDocumentID FROM OuterDocuments WHERE OuterDocumentID IN (SELECT DocumentID FROM DocumentsConfirmations WHERE DocumentCategoryID = 2 AND DocumentConfirmationID IN (SELECT DocumentConfirmationID FROM DocumentsConfirmationsRecipients WHERE Status = 0 AND UserID = " + Security.CurrentUserID + "))))"
                        , ConnectionStrings.LightConnectionString))
                {
                    DA.Fill(UpdatesCommentsDataTable);
                }


                foreach (DataRow Row in UpdatesDocumentsDataTable.Rows)
                {
                    var uRow = UpdatesCommentsDataTable.Select("DocumentID = " + Row["DocumentID"] + " AND DateTime > '" + Convert.ToDateTime(Row["DateTime"]) + "'");

                    if (uRow.Count() > 0)
                    {
                        var date = Convert.ToDateTime(Row["DateTime"]);

                        foreach (var uuRow in uRow)
                        {
                            if (date < Convert.ToDateTime(uuRow["DateTime"]))
                                date = Convert.ToDateTime(uuRow["DateTime"]);
                        }

                        Row["UpdateDateTime"] = date;
                    }
                    else
                        Row["UpdateDateTime"] = Row["DateTime"];
                }

                UpdatesDocumentsDataTable.DefaultView.Sort = "UpdateDateTime DESC";

                using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsFiles WHERE " +
                                                   "(DocumentCategoryID = 0 AND DocumentID IN (SELECT InnerDocumentID FROM InnerDocuments WHERE InnerDocumentID IN (SELECT DocumentID FROM DocumentsConfirmations WHERE DocumentCategoryID = 0 AND DocumentConfirmationID IN (SELECT DocumentConfirmationID FROM DocumentsConfirmationsRecipients WHERE Status = 0 AND UserID = " + Security.CurrentUserID + ")))) OR" +
                                                   "(DocumentCategoryID = 1 AND DocumentID IN (SELECT IncomeDocumentID FROM IncomeDocuments WHERE IncomeDocumentID IN (SELECT DocumentID FROM DocumentsConfirmations WHERE DocumentCategoryID = 1 AND DocumentConfirmationID IN (SELECT DocumentConfirmationID FROM DocumentsConfirmationsRecipients WHERE Status = 0 AND UserID = " + Security.CurrentUserID + ")))) OR" +
                                                   "(DocumentCategoryID = 2 AND DocumentID IN (SELECT OuterDocumentID FROM OuterDocuments WHERE OuterDocumentID IN (SELECT DocumentID FROM DocumentsConfirmations WHERE DocumentCategoryID = 2 AND DocumentConfirmationID IN (SELECT DocumentConfirmationID FROM DocumentsConfirmationsRecipients WHERE Status = 0 AND UserID = " + Security.CurrentUserID + "))))"
                        , ConnectionStrings.LightConnectionString))
                {
                    DA.Fill(UpdatesFilesDataTable);
                }

                using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsRecipients WHERE " +
                                                   "(DocumentCategoryID = 0 AND DocumentID IN (SELECT InnerDocumentID FROM InnerDocuments WHERE InnerDocumentID IN (SELECT DocumentID FROM DocumentsConfirmations WHERE DocumentCategoryID = 0 AND DocumentConfirmationID IN (SELECT DocumentConfirmationID FROM DocumentsConfirmationsRecipients WHERE Status = 0 AND UserID = " + Security.CurrentUserID + ")))) OR" +
                                                   "(DocumentCategoryID = 1 AND DocumentID IN (SELECT IncomeDocumentID FROM IncomeDocuments WHERE IncomeDocumentID IN (SELECT DocumentID FROM DocumentsConfirmations WHERE DocumentCategoryID = 1 AND DocumentConfirmationID IN (SELECT DocumentConfirmationID FROM DocumentsConfirmationsRecipients WHERE Status = 0 AND UserID = " + Security.CurrentUserID + ")))) OR" +
                                                   "(DocumentCategoryID = 2 AND DocumentID IN (SELECT OuterDocumentID FROM OuterDocuments WHERE OuterDocumentID IN (SELECT DocumentID FROM DocumentsConfirmations WHERE DocumentCategoryID = 2 AND DocumentConfirmationID IN (SELECT DocumentConfirmationID FROM DocumentsConfirmationsRecipients WHERE Status = 0 AND UserID = " + Security.CurrentUserID + "))))"
                        , ConnectionStrings.LightConnectionString))
                {
                    DA.Fill(UpdatesRecipientsDataTable);
                }

                using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsCommentsFiles WHERE DocumentCommentID IN (SELECT DocumentCommentID FROM DocumentsComments WHERE " +
                                                   "(DocumentCategoryID = 0 AND DocumentID IN (SELECT InnerDocumentID FROM InnerDocuments WHERE InnerDocumentID IN (SELECT DocumentID FROM DocumentsConfirmations WHERE DocumentCategoryID = 0 AND DocumentConfirmationID IN (SELECT DocumentConfirmationID FROM DocumentsConfirmationsRecipients WHERE Status = 0 AND UserID = " + Security.CurrentUserID + ")))) OR" +
                                                   "(DocumentCategoryID = 1 AND DocumentID IN (SELECT IncomeDocumentID FROM IncomeDocuments WHERE IncomeDocumentID IN (SELECT DocumentID FROM DocumentsConfirmations WHERE DocumentCategoryID = 1 AND DocumentConfirmationID IN (SELECT DocumentConfirmationID FROM DocumentsConfirmationsRecipients WHERE Status = 0 AND UserID = " + Security.CurrentUserID + ")))) OR" +
                                                   "(DocumentCategoryID = 2 AND DocumentID IN (SELECT OuterDocumentID FROM OuterDocuments WHERE OuterDocumentID IN (SELECT DocumentID FROM DocumentsConfirmations WHERE DocumentCategoryID = 2 AND DocumentConfirmationID IN (SELECT DocumentConfirmationID FROM DocumentsConfirmationsRecipients WHERE Status = 0 AND UserID = " + Security.CurrentUserID + ")))))"
                        , ConnectionStrings.LightConnectionString))
                {
                    DA.Fill(UpdatesCommentsFilesDataTable);
                }


                using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsConfirmations WHERE DocumentConfirmationID IN (SELECT DocumentConfirmationID FROM DocumentsConfirmationsRecipients WHERE Status = 0 AND UserID = " + Security.CurrentUserID + ")", ConnectionStrings.LightConnectionString))
                {
                    DA.Fill(UpdatesConfirmsDataTable);
                }

                using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsConfirmationsRecipients WHERE DocumentConfirmationID IN (SELECT DocumentConfirmationID FROM DocumentsConfirmations WHERE DocumentConfirmationID IN (SELECT DocumentConfirmationID FROM DocumentsConfirmationsRecipients WHERE Status = 0 AND UserID = " + Security.CurrentUserID + "))", ConnectionStrings.LightConnectionString))
                {
                    DA.Fill(UpdatesConfirmsRecipientsDataTable);
                }
            }
            else
            {
                using (var DA = new SqlDataAdapter("SELECT InnerDocuments.*, InnerDocuments.InnerDocumentID AS DocumentID FROM InnerDocuments WHERE InnerDocumentID IN (SELECT DocumentID FROM DocumentsConfirmations WHERE DocumentCategoryID = 0 AND Status = " + StatusType + ")", ConnectionStrings.LightConnectionString))
                {
                    DA.Fill(UpdatesDocumentsDataTable);
                }

                using (var DA = new SqlDataAdapter("SELECT IncomeDocuments.*, IncomeDocuments.IncomeDocumentID AS DocumentID FROM IncomeDocuments WHERE IncomeDocumentID IN (SELECT DocumentID FROM DocumentsConfirmations WHERE DocumentCategoryID = 1 AND Status = " + StatusType + ")", ConnectionStrings.LightConnectionString))
                {
                    DA.Fill(UpdatesDocumentsDataTable);
                }

                using (var DA = new SqlDataAdapter("SELECT OuterDocuments.*, OuterDocuments.OuterDocumentID AS DocumentID FROM OuterDocuments WHERE OuterDocumentID IN (SELECT DocumentID FROM DocumentsConfirmations WHERE DocumentCategoryID = 2 AND Status = " + StatusType + ")", ConnectionStrings.LightConnectionString))
                {
                    DA.Fill(UpdatesDocumentsDataTable);
                }

                using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsComments WHERE " +
                                                   "(DocumentCategoryID = 0 AND DocumentID IN (SELECT InnerDocumentID FROM InnerDocuments WHERE InnerDocumentID IN (SELECT DocumentID FROM DocumentsConfirmations WHERE DocumentCategoryID = 0 AND Status = " + StatusType + "))) OR" +
                                                   "(DocumentCategoryID = 1 AND DocumentID IN (SELECT IncomeDocumentID FROM IncomeDocuments WHERE IncomeDocumentID IN (SELECT DocumentID FROM DocumentsConfirmations WHERE DocumentCategoryID = 1 AND Status = " + StatusType + "))) OR" +
                                                   "(DocumentCategoryID = 2 AND DocumentID IN (SELECT OuterDocumentID FROM OuterDocuments WHERE OuterDocumentID IN (SELECT DocumentID FROM DocumentsConfirmations WHERE DocumentCategoryID = 2 AND Status = " + StatusType + ")))"
                        , ConnectionStrings.LightConnectionString))
                {
                    DA.Fill(UpdatesCommentsDataTable);
                }


                foreach (DataRow Row in UpdatesDocumentsDataTable.Rows)
                {
                    var uRow = UpdatesCommentsDataTable.Select("DocumentID = " + Row["DocumentID"] + " AND DateTime > '" + Convert.ToDateTime(Row["DateTime"]) + "'");

                    if (uRow.Count() > 0)
                    {
                        var date = Convert.ToDateTime(Row["DateTime"]);

                        foreach (var uuRow in uRow)
                        {
                            if (date < Convert.ToDateTime(uuRow["DateTime"]))
                                date = Convert.ToDateTime(uuRow["DateTime"]);
                        }

                        Row["UpdateDateTime"] = date;
                    }
                    else
                        Row["UpdateDateTime"] = Row["DateTime"];
                }


                using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsFiles WHERE " +
                                                   "(DocumentCategoryID = 0 AND DocumentID IN (SELECT InnerDocumentID FROM InnerDocuments WHERE InnerDocumentID IN (SELECT DocumentID FROM DocumentsConfirmations WHERE DocumentCategoryID = 0 AND Status = " + StatusType + "))) OR" +
                                                   "(DocumentCategoryID = 1 AND DocumentID IN (SELECT IncomeDocumentID FROM IncomeDocuments WHERE IncomeDocumentID IN (SELECT DocumentID FROM DocumentsConfirmations WHERE DocumentCategoryID = 1 AND Status = " + StatusType + "))) OR" +
                                                   "(DocumentCategoryID = 2 AND DocumentID IN (SELECT OuterDocumentID FROM OuterDocuments WHERE OuterDocumentID IN (SELECT DocumentID FROM DocumentsConfirmations WHERE DocumentCategoryID = 2 AND Status = " + StatusType + ")))"
                        , ConnectionStrings.LightConnectionString))
                {
                    DA.Fill(UpdatesFilesDataTable);
                }

                using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsRecipients WHERE " +
                                                   "(DocumentCategoryID = 0 AND DocumentID IN (SELECT InnerDocumentID FROM InnerDocuments WHERE InnerDocumentID IN (SELECT DocumentID FROM DocumentsConfirmations WHERE DocumentCategoryID = 0 AND Status = " + StatusType + "))) OR" +
                                                   "(DocumentCategoryID = 1 AND DocumentID IN (SELECT IncomeDocumentID FROM IncomeDocuments WHERE IncomeDocumentID IN (SELECT DocumentID FROM DocumentsConfirmations WHERE DocumentCategoryID = 1 AND Status = " + StatusType + "))) OR" +
                                                   "(DocumentCategoryID = 2 AND DocumentID IN (SELECT OuterDocumentID FROM OuterDocuments WHERE OuterDocumentID IN (SELECT DocumentID FROM DocumentsConfirmations WHERE DocumentCategoryID = 2 AND Status = " + StatusType + ")))"
                        , ConnectionStrings.LightConnectionString))
                {
                    DA.Fill(UpdatesRecipientsDataTable);
                }

                using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsCommentsFiles WHERE DocumentCommentID IN (SELECT DocumentCommentID FROM DocumentsComments WHERE " +
                                                   "(DocumentCategoryID = 0 AND DocumentID IN (SELECT InnerDocumentID FROM InnerDocuments WHERE InnerDocumentID IN (SELECT DocumentID FROM DocumentsConfirmations WHERE DocumentCategoryID = 0 AND Status = " + StatusType + "))) OR" +
                                                   "(DocumentCategoryID = 1 AND DocumentID IN (SELECT IncomeDocumentID FROM IncomeDocuments WHERE IncomeDocumentID IN (SELECT DocumentID FROM DocumentsConfirmations WHERE DocumentCategoryID = 1 AND Status = " + StatusType + "))) OR" +
                                                   "(DocumentCategoryID = 2 AND DocumentID IN (SELECT OuterDocumentID FROM OuterDocuments WHERE OuterDocumentID IN (SELECT DocumentID FROM DocumentsConfirmations WHERE DocumentCategoryID = 2 AND Status = " + StatusType + "))))"
                        , ConnectionStrings.LightConnectionString))
                {
                    DA.Fill(UpdatesCommentsFilesDataTable);
                }


                using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsConfirmations WHERE Status = " + StatusType, ConnectionStrings.LightConnectionString))
                {
                    DA.Fill(UpdatesConfirmsDataTable);
                }

                using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsConfirmationsRecipients WHERE DocumentConfirmationID IN (SELECT DocumentConfirmationID FROM DocumentsConfirmations WHERE Status = " + StatusType + ")", ConnectionStrings.LightConnectionString))
                {
                    DA.Fill(UpdatesConfirmsRecipientsDataTable);
                }
            }


        }

        public void FillItemUpdate(int DocumentID, int DocumentCategoryID)
        {
            UpdatesDocumentsDataTable.Clear();
            UpdatesCommentsDataTable.Clear();
            UpdatesFilesDataTable.Clear();
            UpdatesRecipientsDataTable.Clear();
            UpdatesCommentsFilesDataTable.Clear();
            UpdatesConfirmsDataTable.Clear();
            UpdatesConfirmsRecipientsDataTable.Clear();

            if (DocumentCategoryID == 0)
                using (var DA = new SqlDataAdapter("SELECT InnerDocuments.*, InnerDocuments.InnerDocumentID AS DocumentID FROM InnerDocuments WHERE InnerDocumentID  = " + DocumentID, ConnectionStrings.LightConnectionString))
                {
                    DA.Fill(UpdatesDocumentsDataTable);
                }

            if (DocumentCategoryID == 1)
                using (var DA = new SqlDataAdapter("SELECT IncomeDocuments.*, IncomeDocuments.IncomeDocumentID AS DocumentID FROM IncomeDocuments WHERE IncomeDocumentID = " + DocumentID, ConnectionStrings.LightConnectionString))
                {
                    DA.Fill(UpdatesDocumentsDataTable);
                }

            if (DocumentCategoryID == 2)
                using (var DA = new SqlDataAdapter("SELECT OuterDocuments.*, OuterDocuments.OuterDocumentID AS DocumentID FROM OuterDocuments WHERE OuterDocumentID = " + DocumentID, ConnectionStrings.LightConnectionString))
                {
                    DA.Fill(UpdatesDocumentsDataTable);
                }

            using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsComments WHERE DocumentCategoryID = " + DocumentCategoryID + " AND DocumentID = " + DocumentID
                    , ConnectionStrings.LightConnectionString))
            {
                DA.Fill(UpdatesCommentsDataTable);
            }


            foreach (DataRow Row in UpdatesDocumentsDataTable.Rows)
            {
                var uRow = UpdatesCommentsDataTable.Select("DocumentID = " + Row["DocumentID"] + " AND DateTime > '" + Convert.ToDateTime(Row["DateTime"]) + "'");

                if (uRow.Count() > 0)
                {
                    var date = Convert.ToDateTime(Row["DateTime"]);

                    foreach (var uuRow in uRow)
                    {
                        if (date < Convert.ToDateTime(uuRow["DateTime"]))
                            date = Convert.ToDateTime(uuRow["DateTime"]);
                    }

                    Row["UpdateDateTime"] = date;
                }
                else
                    Row["UpdateDateTime"] = Row["DateTime"];
            }


            using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsFiles WHERE DocumentCategoryID = " + DocumentCategoryID + " AND DocumentID = " + DocumentID
                    , ConnectionStrings.LightConnectionString))
            {
                DA.Fill(UpdatesFilesDataTable);
            }

            using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsRecipients WHERE DocumentCategoryID = " + DocumentCategoryID + " AND DocumentID = " + DocumentID
                    , ConnectionStrings.LightConnectionString))
            {
                DA.Fill(UpdatesRecipientsDataTable);
            }

            using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsCommentsFiles WHERE DocumentCommentID IN (SELECT DocumentCommentID FROM DocumentsComments WHERE DocumentCategoryID = " + DocumentCategoryID + " AND DocumentID = " + DocumentID + ")"
                    , ConnectionStrings.LightConnectionString))
            {
                DA.Fill(UpdatesCommentsFilesDataTable);
            }


            using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsConfirmations WHERE DocumentCategoryID = " + DocumentCategoryID + " AND DocumentID = " + DocumentID, ConnectionStrings.LightConnectionString))
            {
                DA.Fill(UpdatesConfirmsDataTable);
            }

            using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsConfirmationsRecipients WHERE DocumentConfirmationID IN (SELECT DocumentConfirmationID FROM DocumentsConfirmations WHERE DocumentCategoryID = " + DocumentCategoryID + " AND DocumentID = " + DocumentID + ")", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(UpdatesConfirmsRecipientsDataTable);
            }
        }

        public void FillMyUpdates(int UserID, int Type)//0 - i'm a creator, 1 - i'm in recipient list
        {
            UpdatesDocumentsDataTable.Clear();
            UpdatesCommentsDataTable.Clear();
            UpdatesFilesDataTable.Clear();
            UpdatesRecipientsDataTable.Clear();
            UpdatesCommentsFilesDataTable.Clear();
            UpdatesConfirmsDataTable.Clear();
            UpdatesConfirmsRecipientsDataTable.Clear();

            if (Type == 0)
            {
                using (var DA = new SqlDataAdapter("SELECT InnerDocuments.*, InnerDocuments.InnerDocumentID AS DocumentID FROM InnerDocuments WHERE UserID = " + UserID, ConnectionStrings.LightConnectionString))
                {
                    DA.Fill(UpdatesDocumentsDataTable);
                }

                using (var DA = new SqlDataAdapter("SELECT IncomeDocuments.*, IncomeDocuments.IncomeDocumentID AS DocumentID FROM IncomeDocuments WHERE UserID = " + UserID, ConnectionStrings.LightConnectionString))
                {
                    DA.Fill(UpdatesDocumentsDataTable);
                }

                using (var DA = new SqlDataAdapter("SELECT OuterDocuments.*, OuterDocuments.OuterDocumentID AS DocumentID FROM OuterDocuments WHERE UserID = " + UserID, ConnectionStrings.LightConnectionString))
                {
                    DA.Fill(UpdatesDocumentsDataTable);
                }


                using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsComments WHERE " +
                                                   "(DocumentCategoryID = 0 AND DocumentID IN (SELECT InnerDocumentID FROM InnerDocuments WHERE UserID = " + UserID + ")) OR " +
                                                   "(DocumentCategoryID = 1 AND DocumentID IN (SELECT IncomeDocumentID FROM IncomeDocuments WHERE UserID = " + UserID + ")) OR " +
                                                   "(DocumentCategoryID = 2 AND DocumentID IN (SELECT OuterDocumentID FROM OuterDocuments WHERE UserID = " + UserID + "))"
                    , ConnectionStrings.LightConnectionString))
                {
                    DA.Fill(UpdatesCommentsDataTable);
                }

                foreach (DataRow Row in UpdatesDocumentsDataTable.Rows)
                {
                    var uRow = UpdatesCommentsDataTable.Select("DocumentID = " + Row["DocumentID"] + " AND DateTime > '" + Convert.ToDateTime(Row["DateTime"]) + "'");

                    if (uRow.Count() > 0)
                    {
                        var date = Convert.ToDateTime(Row["DateTime"]);

                        foreach (var uuRow in uRow)
                        {
                            if (date < Convert.ToDateTime(uuRow["DateTime"]))
                                date = Convert.ToDateTime(uuRow["DateTime"]);
                        }

                        Row["UpdateDateTime"] = date;
                    }
                    else
                        Row["UpdateDateTime"] = Row["DateTime"];
                }

                UpdatesDocumentsDataTable.DefaultView.Sort = "UpdateDateTime DESC";

                using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsFiles WHERE " +
                                                   "(DocumentCategoryID = 0 AND DocumentID IN (SELECT InnerDocumentID FROM InnerDocuments WHERE UserID = " + UserID + ")) OR" +
                                                   "(DocumentCategoryID = 1 AND DocumentID IN (SELECT IncomeDocumentID FROM IncomeDocuments WHERE UserID = " + UserID + ")) OR" +
                                                   "(DocumentCategoryID = 2 AND DocumentID IN (SELECT OuterDocumentID FROM OuterDocuments WHERE UserID = " + UserID + "))"
                        , ConnectionStrings.LightConnectionString))
                {
                    DA.Fill(UpdatesFilesDataTable);
                }


                using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsRecipients WHERE " +
                                                   "(DocumentCategoryID = 0 AND DocumentID IN (SELECT InnerDocumentID FROM InnerDocuments WHERE UserID = " + UserID + ")) OR" +
                                                   "(DocumentCategoryID = 1 AND DocumentID IN (SELECT IncomeDocumentID FROM IncomeDocuments WHERE UserID = " + UserID + ")) OR" +
                                                   "(DocumentCategoryID = 2 AND DocumentID IN (SELECT OuterDocumentID FROM OuterDocuments WHERE UserID = " + UserID + "))"
                        , ConnectionStrings.LightConnectionString))
                {
                    DA.Fill(UpdatesRecipientsDataTable);
                }

                using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsCommentsFiles WHERE DocumentCommentID IN (SELECT DocumentCommentID FROM DocumentsComments WHERE " +
                                                   "(DocumentCategoryID = 0 AND DocumentID IN (SELECT InnerDocumentID FROM InnerDocuments WHERE UserID = " + UserID + ")) OR " +
                                                   "(DocumentCategoryID = 1 AND DocumentID IN (SELECT IncomeDocumentID FROM IncomeDocuments WHERE UserID = " + UserID + ")) OR " +
                                                   "(DocumentCategoryID = 2 AND DocumentID IN (SELECT OuterDocumentID FROM OuterDocuments WHERE UserID = " + UserID + ")))"
                    , ConnectionStrings.LightConnectionString))
                {
                    DA.Fill(UpdatesCommentsFilesDataTable);
                }


                using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsConfirmations WHERE" +
                                                   "(DocumentCategoryID = 0 AND DocumentID IN (SELECT InnerDocumentID FROM InnerDocuments WHERE UserID = " + UserID + ")) OR " +
                                                   "(DocumentCategoryID = 1 AND DocumentID IN (SELECT IncomeDocumentID FROM IncomeDocuments WHERE UserID = " + UserID + ")) OR " +
                                                   "(DocumentCategoryID = 2 AND DocumentID IN (SELECT OuterDocumentID FROM OuterDocuments WHERE UserID = " + UserID + "))"
                    , ConnectionStrings.LightConnectionString))
                {
                    DA.Fill(UpdatesConfirmsDataTable);
                }

                using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsConfirmationsRecipients WHERE DocumentConfirmationID IN (SELECT DocumentConfirmationID FROM DocumentsConfirmations WHERE" +
                                                   "(DocumentCategoryID = 0 AND DocumentID IN (SELECT InnerDocumentID FROM InnerDocuments WHERE UserID = " + UserID + ")) OR " +
                                                   "(DocumentCategoryID = 1 AND DocumentID IN (SELECT IncomeDocumentID FROM IncomeDocuments WHERE UserID = " + UserID + ")) OR " +
                                                   "(DocumentCategoryID = 2 AND DocumentID IN (SELECT OuterDocumentID FROM OuterDocuments WHERE UserID = " + UserID + ")))"
                    , ConnectionStrings.LightConnectionString))
                {
                    DA.Fill(UpdatesConfirmsRecipientsDataTable);
                }
            }
            else
            {
                using (var DA = new SqlDataAdapter("SELECT InnerDocuments.*, InnerDocuments.InnerDocumentID AS DocumentID FROM InnerDocuments WHERE UserID != " + UserID + " AND InnerDocumentID IN (SELECT DocumentID FROM DocumentsRecipients WHERE DocumentCategoryID = 0 AND UserID = " + UserID + ")", ConnectionStrings.LightConnectionString))
                {
                    DA.Fill(UpdatesDocumentsDataTable);
                }

                using (var DA = new SqlDataAdapter("SELECT IncomeDocuments.*, IncomeDocuments.IncomeDocumentID AS DocumentID FROM IncomeDocuments WHERE UserID != " + UserID + " AND IncomeDocumentID IN (SELECT DocumentID FROM DocumentsRecipients WHERE DocumentCategoryID = 1 AND UserID = " + UserID + ")", ConnectionStrings.LightConnectionString))
                {
                    DA.Fill(UpdatesDocumentsDataTable);
                }

                using (var DA = new SqlDataAdapter("SELECT OuterDocuments.*, OuterDocuments.OuterDocumentID AS DocumentID FROM OuterDocuments WHERE UserID != " + UserID + " AND OuterDocumentID IN (SELECT DocumentID FROM DocumentsRecipients WHERE DocumentCategoryID = 2 AND UserID = " + UserID + ")", ConnectionStrings.LightConnectionString))
                {
                    DA.Fill(UpdatesDocumentsDataTable);
                }


                using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsComments WHERE " +
                                                   "(DocumentCategoryID = 0 AND DocumentID IN (SELECT InnerDocumentID FROM InnerDocuments WHERE UserID != " + UserID + " AND InnerDocumentID IN (SELECT DocumentID FROM DocumentsRecipients WHERE DocumentCategoryID = 0 AND UserID = " + UserID + "))) OR " +
                                                   "(DocumentCategoryID = 1 AND DocumentID IN (SELECT IncomeDocumentID FROM IncomeDocuments WHERE UserID != " + UserID + " AND IncomeDocumentID IN (SELECT DocumentID FROM DocumentsRecipients WHERE DocumentCategoryID = 1 AND UserID = " + UserID + "))) OR " +
                                                   "(DocumentCategoryID = 2 AND DocumentID IN (SELECT OuterDocumentID FROM OuterDocuments WHERE UserID != " + UserID + " AND OuterDocumentID IN (SELECT DocumentID FROM DocumentsRecipients WHERE DocumentCategoryID = 2 AND UserID = " + UserID + ")))"
                    , ConnectionStrings.LightConnectionString))
                {
                    DA.Fill(UpdatesCommentsDataTable);
                }

                foreach (DataRow Row in UpdatesDocumentsDataTable.Rows)
                {
                    var uRow = UpdatesCommentsDataTable.Select("DocumentID = " + Row["DocumentID"] + " AND DateTime > '" + Convert.ToDateTime(Row["DateTime"]) + "'");

                    if (uRow.Count() > 0)
                    {
                        var date = Convert.ToDateTime(Row["DateTime"]);

                        foreach (var uuRow in uRow)
                        {
                            if (date < Convert.ToDateTime(uuRow["DateTime"]))
                                date = Convert.ToDateTime(uuRow["DateTime"]);
                        }

                        Row["UpdateDateTime"] = date;
                    }
                    else
                        Row["UpdateDateTime"] = Row["DateTime"];
                }

                UpdatesDocumentsDataTable.DefaultView.Sort = "UpdateDateTime DESC";

                using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsFiles WHERE " +
                                                   "(DocumentCategoryID = 0 AND DocumentID IN (SELECT InnerDocumentID FROM InnerDocuments WHERE UserID != " + UserID + " AND InnerDocumentID IN (SELECT DocumentID FROM DocumentsRecipients WHERE DocumentCategoryID = 0 AND UserID = " + UserID + "))) OR " +
                                                   "(DocumentCategoryID = 1 AND DocumentID IN (SELECT IncomeDocumentID FROM IncomeDocuments WHERE UserID != " + UserID + " AND IncomeDocumentID IN (SELECT DocumentID FROM DocumentsRecipients WHERE DocumentCategoryID = 1 AND UserID = " + UserID + "))) OR " +
                                                   "(DocumentCategoryID = 2 AND DocumentID IN (SELECT OuterDocumentID FROM OuterDocuments WHERE UserID != " + UserID + " AND OuterDocumentID IN (SELECT DocumentID FROM DocumentsRecipients WHERE DocumentCategoryID = 2 AND UserID = " + UserID + ")))"
                    , ConnectionStrings.LightConnectionString))
                {
                    DA.Fill(UpdatesFilesDataTable);
                }


                using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsRecipients WHERE " +
                                                   "(DocumentCategoryID = 0 AND DocumentID IN (SELECT InnerDocumentID FROM InnerDocuments WHERE UserID != " + UserID + " AND InnerDocumentID IN (SELECT DocumentID FROM DocumentsRecipients WHERE DocumentCategoryID = 0 AND UserID = " + UserID + "))) OR " +
                                                   "(DocumentCategoryID = 1 AND DocumentID IN (SELECT IncomeDocumentID FROM IncomeDocuments WHERE UserID != " + UserID + " AND IncomeDocumentID IN (SELECT DocumentID FROM DocumentsRecipients WHERE DocumentCategoryID = 1 AND UserID = " + UserID + "))) OR " +
                                                   "(DocumentCategoryID = 2 AND DocumentID IN (SELECT OuterDocumentID FROM OuterDocuments WHERE UserID != " + UserID + " AND OuterDocumentID IN (SELECT DocumentID FROM DocumentsRecipients WHERE DocumentCategoryID = 2 AND UserID = " + UserID + ")))"
                    , ConnectionStrings.LightConnectionString))
                {
                    DA.Fill(UpdatesRecipientsDataTable);
                }

                using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsCommentsFiles WHERE DocumentCommentID IN (SELECT DocumentCommentID FROM DocumentsComments WHERE " +
                                                   "(DocumentCategoryID = 0 AND DocumentID IN (SELECT InnerDocumentID FROM InnerDocuments WHERE UserID != " + UserID + " AND InnerDocumentID IN (SELECT DocumentID FROM DocumentsRecipients WHERE DocumentCategoryID = 0 AND UserID = " + UserID + "))) OR " +
                                                   "(DocumentCategoryID = 1 AND DocumentID IN (SELECT IncomeDocumentID FROM IncomeDocuments WHERE UserID != " + UserID + " AND IncomeDocumentID IN (SELECT DocumentID FROM DocumentsRecipients WHERE DocumentCategoryID = 1 AND UserID = " + UserID + "))) OR " +
                                                   "(DocumentCategoryID = 2 AND DocumentID IN (SELECT OuterDocumentID FROM OuterDocuments WHERE UserID != " + UserID + " AND OuterDocumentID IN (SELECT DocumentID FROM DocumentsRecipients WHERE DocumentCategoryID = 2 AND UserID = " + UserID + "))))"
                    , ConnectionStrings.LightConnectionString))
                {
                    DA.Fill(UpdatesCommentsFilesDataTable);
                }


                using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsConfirmations WHERE " +
                                                   "(DocumentCategoryID = 0 AND DocumentID IN (SELECT InnerDocumentID FROM InnerDocuments WHERE UserID != " + UserID + " AND InnerDocumentID IN (SELECT DocumentID FROM DocumentsRecipients WHERE DocumentCategoryID = 0 AND UserID = " + UserID + "))) OR " +
                                                   "(DocumentCategoryID = 1 AND DocumentID IN (SELECT IncomeDocumentID FROM IncomeDocuments WHERE UserID != " + UserID + " AND IncomeDocumentID IN (SELECT DocumentID FROM DocumentsRecipients WHERE DocumentCategoryID = 1 AND UserID = " + UserID + "))) OR " +
                                                   "(DocumentCategoryID = 2 AND DocumentID IN (SELECT OuterDocumentID FROM OuterDocuments WHERE UserID != " + UserID + " AND OuterDocumentID IN (SELECT DocumentID FROM DocumentsRecipients WHERE DocumentCategoryID = 2 AND UserID = " + UserID + ")))"
                    , ConnectionStrings.LightConnectionString))
                {
                    DA.Fill(UpdatesConfirmsDataTable);
                }

                using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsConfirmationsRecipients WHERE DocumentConfirmationID IN (SELECT DocumentConfirmationID FROM DocumentsConfirmations WHERE " +
                                                   "(DocumentCategoryID = 0 AND DocumentID IN (SELECT InnerDocumentID FROM InnerDocuments WHERE UserID != " + UserID + " AND InnerDocumentID IN (SELECT DocumentID FROM DocumentsRecipients WHERE DocumentCategoryID = 0 AND UserID = " + UserID + "))) OR " +
                                                   "(DocumentCategoryID = 1 AND DocumentID IN (SELECT IncomeDocumentID FROM IncomeDocuments WHERE UserID != " + UserID + " AND IncomeDocumentID IN (SELECT DocumentID FROM DocumentsRecipients WHERE DocumentCategoryID = 1 AND UserID = " + UserID + "))) OR " +
                                                   "(DocumentCategoryID = 2 AND DocumentID IN (SELECT OuterDocumentID FROM OuterDocuments WHERE UserID != " + UserID + " AND OuterDocumentID IN (SELECT DocumentID FROM DocumentsRecipients WHERE DocumentCategoryID = 2 AND UserID = " + UserID + "))))"
                    , ConnectionStrings.LightConnectionString))
                {
                    DA.Fill(UpdatesConfirmsRecipientsDataTable);
                }
            }
        }

        public bool IsAccessGrantedInner(int UserID, int InnerDocumentID)
        {
            using (var DA = new SqlDataAdapter("SELECT UserID FROM InnerDocuments WHERE InnerDocumentID = " + InnerDocumentID + " AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    return Convert.ToBoolean(DA.Fill(DT) > 0);
                }
            }
        }

        public bool IsAccessGrantedIncome(int UserID, int IncomeDocumentID)
        {
            using (var DA = new SqlDataAdapter("SELECT UserID FROM IncomeDocuments WHERE IncomeDocumentID = " + IncomeDocumentID + " AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    return Convert.ToBoolean(DA.Fill(DT) > 0);
                }
            }
        }

        public bool IsAccessGrantedOuter(int UserID, int OuterDocumentID)
        {
            using (var DA = new SqlDataAdapter("SELECT UserID FROM OuterDocuments WHERE OuterDocumentID = " + OuterDocumentID + " AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    return Convert.ToBoolean(DA.Fill(DT) > 0);
                }
            }
        }


        public int GetNotSigned(int UserID)
        {
            using (var DA = new SqlDataAdapter("SELECT Count(DocumentConfirmationID) FROM DocumentsConfirmations WHERE DocumentConfirmationID IN (SELECT DocumentConfirmationID FROM DocumentsConfirmationsRecipients WHERE Status = 0 AND UserID = " + UserID + ")", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    if (DA.Fill(DT) == 0)
                        return 0;

                    return Convert.ToInt32(DT.Rows[0][0]);
                }
            }
        }

        //public void ClearInnerSubscribes()
        //{
        //    using (SqlDataAdapter DA = new SqlDataAdapter("DELETE FROM SubscribesRecords WHERE SubscribesItemID = 18 AND UserID = " + Security.CurrentUserID, ConnectionStrings.LightConnectionString))
        //    {
        //        using (DataTable DT = new DataTable())
        //        {
        //            DA.Fill(DT);
        //        }
        //    }
        //}

        //public void ClearOuterSubscribes()
        //{
        //    using (SqlDataAdapter DA = new SqlDataAdapter("DELETE FROM SubscribesRecords WHERE SubscribesItemID = 20 AND UserID = " + Security.CurrentUserID, ConnectionStrings.LightConnectionString))
        //    {
        //        using (DataTable DT = new DataTable())
        //        {
        //            DA.Fill(DT);
        //        }
        //    }
        //}

        //public void ClearIncomeSubscribes()
        //{
        //    using (SqlDataAdapter DA = new SqlDataAdapter("DELETE FROM SubscribesRecords WHERE SubscribesItemID = 19 AND UserID = " + Security.CurrentUserID, ConnectionStrings.LightConnectionString))
        //    {
        //        using (DataTable DT = new DataTable())
        //        {
        //            DA.Fill(DT);
        //        }
        //    }
        //}


        //public void FillCurrentInnerFiles(int InnerDocumentID)
        //{
        //    CurrentFilesDataTable.Clear();

        //    using(SqlDataAdapter DA = new SqlDataAdapter("SELECT FileName FROM InnerDocumentsFiles WHERE InnerDocumentID = " + InnerDocumentID, ConnectionStrings.LightConnectionString))
        //    {
        //        DA.Fill(CurrentFilesDataTable);
        //    }
        //}

        //public void FillCurrentIncomeFiles(int IncomeDocumentID)
        //{
        //    CurrentFilesDataTable.Clear();

        //    using (SqlDataAdapter DA = new SqlDataAdapter("SELECT FileName FROM IncomeDocumentsFiles WHERE IncomeDocumentID = " + IncomeDocumentID, ConnectionStrings.LightConnectionString))
        //    {
        //        DA.Fill(CurrentFilesDataTable);
        //    }
        //}

        //public void FillCurrentOuterFiles(int OuterDocumentID)
        //{
        //    CurrentFilesDataTable.Clear();

        //    using (SqlDataAdapter DA = new SqlDataAdapter("SELECT FileName FROM OuterDocumentsFiles WHERE OuterDocumentID = " + OuterDocumentID, ConnectionStrings.LightConnectionString))
        //    {
        //        DA.Fill(CurrentFilesDataTable);
        //    }
        //}


        public DataTable GetCommentFiles(int DocumentCommentID)
        {
            using (var DA = new SqlDataAdapter("SELECT FileSize, FileName FROM DocumentsCommentsFiles WHERE DocumentCommentID = " + DocumentCommentID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    using (var CurrentFilesDataTable = new DataTable())
                    {
                        CurrentFilesDataTable.Columns.Add(new DataColumn("FileName", Type.GetType("System.String")));
                        CurrentFilesDataTable.Columns.Add(new DataColumn("FilePath", Type.GetType("System.String")));
                        CurrentFilesDataTable.Columns.Add(new DataColumn("FileSizeText", Type.GetType("System.String")));
                        CurrentFilesDataTable.Columns.Add(new DataColumn("FileSize", Type.GetType("System.Int64")));
                        CurrentFilesDataTable.Columns.Add(new DataColumn("IsNew", Type.GetType("System.Boolean")));

                        DA.Fill(DT);

                        var Size = 0;

                        foreach (DataRow Row in DT.Rows)
                        {
                            var NewRow = CurrentFilesDataTable.NewRow();
                            NewRow["FileName"] = Row["FileName"];
                            NewRow["FilePath"] = "загружен";

                            Size = (int)(Convert.ToInt64(Row["FileSize"]) / 1024);
                            if (Size == 0)
                                Size = 1;

                            NewRow["FileSizeText"] = Size + " КБ";

                            NewRow["FileSize"] = Row["FileSize"];
                            NewRow["IsNew"] = false;


                            CurrentFilesDataTable.Rows.Add(NewRow);
                        }

                        return CurrentFilesDataTable;
                    }
                }
            }
        }

        //public string GetCommentText(int CommentID)
        //{
        //    using (SqlDataAdapter DA = new SqlDataAdapter("SELECT Text FROM DocumentsComments WHERE DocumentCommentID = " + CommentID, ConnectionStrings.LightConnectionString))
        //    {
        //        using (DataTable DT = new DataTable())
        //        {
        //            DA.Fill(DT);

        //            return DT.Rows[0]["Text"].ToString();
        //        }
        //    }
        //}

        public void RemoveComment(int DocumentCommentID)
        {
            foreach (DataRow Row in GetCommentFiles(DocumentCommentID).Rows)
            {
                FM.DeleteFile(Configs.DocumentsPath + "/Документооборот/Комментарии" +
                                    "/" + Row["FileName"], Configs.FTPType);
            }

            using (var DA = new SqlDataAdapter("DELETE FROM DocumentsCommentsFiles WHERE DocumentCommentID = " + DocumentCommentID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }

            using (var DA = new SqlDataAdapter("DELETE FROM SubscribesRecords WHERE SubscribesItemID = 21 AND TableItemID = " + DocumentCommentID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }

            using (var DA = new SqlDataAdapter("DELETE FROM DocumentsComments WHERE DocumentCommentID = " + DocumentCommentID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }
        }

        public void AddComment(int UserID, string Text, int DocumentID, int DocumentCategoryID, DataTable FilesDataTable, ref int iCurrentFile)
        {
            CurrentUploadedFiles.Clear();

            foreach (var Row in FilesDataTable.Select("IsNew = 1"))
            {
                var FileName = "";

                var i = 1;

                FileName = Row["FileName"].ToString();

                if (IsFileExist(FileName))
                {
                    while (IsFileExist(FileName))
                    {
                        FileName = Path.GetFileNameWithoutExtension(Row["FilePath"].ToString()) + "(" + i++ + ")" + Path.GetExtension(Row["FilePath"].ToString());
                    }
                }

                var NewRow = CurrentUploadedFiles.NewRow();
                NewRow["FileName"] = Configs.DocumentsPath + "/Документооборот/Комментарии" +
                                    "/" + FileName;
                CurrentUploadedFiles.Rows.Add(NewRow);

                //load file to ftp
                if (FM.UploadFile(Row["FilePath"].ToString(), Configs.DocumentsPath + "/Документооборот/Комментарии" +
                                    "/" + FileName, Configs.FTPType) == false)
                {
                    FM.DeleteFile(Configs.DocumentsPath + "/Документооборот/Комментарии" +
                                    "/" + FileName, Configs.FTPType);

                    return;
                }

                iCurrentFile++;
            }




            DateTime DateTime;

            using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM DocumentsComments", ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        var NewRow = DT.NewRow();
                        NewRow["UserID"] = UserID;
                        NewRow["Text"] = Text;

                        DateTime = Security.GetCurrentDate();

                        NewRow["DateTime"] = DateTime;
                        NewRow["DocumentID"] = DocumentID;
                        NewRow["DocumentCategoryID"] = DocumentCategoryID;
                        DT.Rows.Add(NewRow);

                        DA.Update(DT);


                    }
                }
            }

            var DocumentCommentID = -1;

            using (var DA = new SqlDataAdapter("SELECT DocumentCommentID FROM DocumentsComments WHERE DateTime = '" + DateTime.ToString("yyyy-MM-dd HH:mm:ss.fff") + "'", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    var s = DateTime.ToString("yyyy-MM-dd HH:mm:ss.fff");

                    DocumentCommentID = Convert.ToInt32(DT.Rows[0]["DocumentCommentID"]);
                }
            }

            AddCommentsSubscribes(DocumentCommentID, DocumentID, DocumentCategoryID);

            if (FilesDataTable.Select("IsNew = 1").Count() == 0)
                return;

            using (var sDA = new SqlDataAdapter("SELECT TOP 0 * FROM DocumentsCommentsFiles", ConnectionStrings.LightConnectionString))
            {
                using (var sCB = new SqlCommandBuilder(sDA))
                {
                    using (var sDT = new DataTable())
                    {
                        sDA.Fill(sDT);

                        foreach (var Row in FilesDataTable.Select("IsNew = 1"))
                        {
                            FileInfo fi;

                            var FileName = "";

                            var i = 1;

                            FileName = Row["FileName"].ToString();

                            if (IsFileExist(FileName))
                            {
                                while (IsFileExist(FileName))
                                {
                                    FileName = Path.GetFileNameWithoutExtension(Row["FilePath"].ToString()) + "(" + i++ + ")" + Path.GetExtension(Row["Path"].ToString());
                                }
                            }

                            //get file size
                            try
                            {
                                fi = new FileInfo(Row["FilePath"].ToString());
                            }
                            catch
                            {
                                return;
                            }

                            var iFileSize = fi.Length;

                            var NewRow = sDT.NewRow();
                            NewRow["FileName"] = FileName;
                            NewRow["DocumentCommentID"] = DocumentCommentID;
                            NewRow["FileSize"] = fi.Length;
                            sDT.Rows.Add(NewRow);
                        }

                        sDA.Update(sDT);
                    }
                }
            }

        }


        public void EditComment(int DocumentCommentID, string Text, DataTable FilesDataTable, ref int iCurrentFile)
        {
            //delete from ftp and database deleted files

            var ExistsFilesDT = GetCommentFiles(DocumentCommentID);

            foreach (DataRow Row in ExistsFilesDT.Rows)
            {
                if (FilesDataTable.Select("IsNew is null AND FileName = '" + Row["FileName"] + "'").Count() == 0)
                {
                    FM.DeleteFile(Configs.DocumentsPath + "/Документооборот/Комментарии" +
                                    "/" + Row["FileName"], Configs.FTPType);


                    using (var dDA = new SqlDataAdapter("DELETE FROM DocumentsCommentsFiles WHERE DocumentCommentID = " + DocumentCommentID + " AND FileName = '" + Row["FileName"] + "'", ConnectionStrings.LightConnectionString))
                    {
                        using (var dDT = new DataTable())
                        {
                            dDA.Fill(dDT);
                        }
                    }
                }
            }

            CurrentUploadedFiles.Clear();

            //add new files to ftp
            foreach (var Row in FilesDataTable.Select("IsNew = 1"))
            {
                var FileName = "";

                var i = 1;

                FileName = Row["FileName"].ToString();

                if (IsFileExist(FileName))
                {
                    while (IsFileExist(FileName))
                    {
                        FileName = Path.GetFileNameWithoutExtension(Row["FilePath"].ToString()) + "(" + i++ + ")" + Path.GetExtension(Row["FilePath"].ToString());
                    }
                }

                var NewRow = CurrentUploadedFiles.NewRow();
                NewRow["FileName"] = Configs.DocumentsPath + "/Документооборот/Комментарии" +
                                    "/" + FileName;
                CurrentUploadedFiles.Rows.Add(NewRow);

                //load file to ftp
                if (FM.UploadFile(Row["FilePath"].ToString(), Configs.DocumentsPath + "/Документооборот/Комментарии" +
                                    "/" + FileName, Configs.FTPType) == false)
                {
                    FM.DeleteFile(Configs.DocumentsPath + "/Документооборот/Комментарии" +
                                    "/" + FileName, Configs.FTPType);

                    return;
                }

                iCurrentFile++;
            }


            //add new files to database

            using (var sDA = new SqlDataAdapter("SELECT TOP 0 * FROM DocumentsCommentsFiles", ConnectionStrings.LightConnectionString))
            {
                using (var sCB = new SqlCommandBuilder(sDA))
                {
                    using (var sDT = new DataTable())
                    {
                        sDA.Fill(sDT);

                        foreach (var Row in FilesDataTable.Select("IsNew = 1"))
                        {
                            FileInfo fi;

                            var FileName = "";

                            var i = 1;

                            FileName = Row["FileName"].ToString();

                            if (IsFileExist(FileName))
                            {
                                while (IsFileExist(FileName))
                                {
                                    FileName = Path.GetFileNameWithoutExtension(Row["FilePath"].ToString()) + "(" + i++ + ")" + Path.GetExtension(Row["FilePath"].ToString());
                                }
                            }

                            //get file size
                            try
                            {
                                fi = new FileInfo(Row["FilePath"].ToString());
                            }
                            catch
                            {
                                return;
                            }

                            var iFileSize = fi.Length;

                            var NewRow = sDT.NewRow();
                            NewRow["FileName"] = FileName;
                            NewRow["DocumentCommentID"] = DocumentCommentID;
                            NewRow["FileSize"] = fi.Length;
                            sDT.Rows.Add(NewRow);
                        }

                        sDA.Update(sDT);
                    }
                }
            }


            //edit comments
            using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsComments WHERE DocumentCommentID = " + DocumentCommentID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        DT.Rows[0]["Text"] = Text;
                        //DT.Rows[0]["DateTime"] = Security.GetCurrentDate();

                        DA.Update(DT);
                    }
                }
            }
        }

        public void AddRecipients(int DocumentCategoryID, int DocumentID, DataTable RecipientsDT)
        {
            using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM DocumentsRecipients", ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        using (var sDA = new SqlDataAdapter("SELECT TOP 0 * FROM SubscribesRecords", ConnectionStrings.LightConnectionString))
                        {
                            using (var sCB = new SqlCommandBuilder(sDA))
                            {
                                using (var sDT = new DataTable())
                                {
                                    sDA.Fill(sDT);

                                    foreach (DataRow Row in RecipientsDT.Rows)
                                    {
                                        var NewRow = DT.NewRow();
                                        NewRow["UserID"] = Row["UserID"];
                                        NewRow["DocumentID"] = DocumentID;
                                        NewRow["DocumentCategoryID"] = DocumentCategoryID;
                                        DT.Rows.Add(NewRow);

                                        var sNewRow = sDT.NewRow();
                                        if (DocumentCategoryID == 0)
                                            sNewRow["SubscribesItemID"] = 18;
                                        if (DocumentCategoryID == 1)
                                            sNewRow["SubscribesItemID"] = 19;
                                        if (DocumentCategoryID == 2)
                                            sNewRow["SubscribesItemID"] = 20;
                                        sNewRow["TableItemID"] = DocumentID;
                                        sNewRow["UserID"] = Row["UserID"];
                                        sNewRow["UserTypeID"] = 0;
                                        sDT.Rows.Add(sNewRow);
                                    }

                                    sDA.Update(sDT);
                                }
                            }
                        }



                        DA.Update(DT);
                    }
                }
            }
        }

        public bool AddInnerDocument(int DocumentTypeID, DateTime DateTime, int UserID, DataTable RecipientsDT,
                                      string Description, string RegNumber, int DocumentStateID, DataTable FilesDT, ref int CurrentUploadedFile, int FactoryID)
        {
            using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM InnerDocuments", ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        var NewRow = DT.NewRow();
                        NewRow["DocumentTypeID"] = DocumentTypeID;
                        NewRow["DateTime"] = DateTime;
                        NewRow["UserID"] = UserID;
                        NewRow["Description"] = Description;
                        NewRow["RegNumber"] = RegNumber;
                        NewRow["DocumentsStateID"] = DocumentStateID;
                        NewRow["FactoryID"] = FactoryID;
                        DT.Rows.Add(NewRow);

                        DA.Update(DT);
                    }
                }
            }

            var InnerDocumentID = -1;

            using (var DA = new SqlDataAdapter("SELECT InnerDocumentID FROM InnerDocuments WHERE DateTime = '" + DateTime.ToString("yyyy-MM-dd HH:mm:ss.fff") + "'", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    InnerDocumentID = Convert.ToInt32(DT.Rows[0]["InnerDocumentID"]);

                    using (var sDA = new SqlDataAdapter("SELECT TOP 0 * FROM DocumentsFiles", ConnectionStrings.LightConnectionString))
                    {
                        using (var sCB = new SqlCommandBuilder(sDA))
                        {
                            using (var sDT = new DataTable())
                            {
                                sDA.Fill(sDT);

                                foreach (DataRow Row in FilesDT.Rows)
                                {
                                    FileInfo fi;

                                    CurrentUploadedFile++;

                                    var FileName = "";

                                    var i = 1;

                                    FileName = Row["FileName"].ToString();

                                    if (IsFileExist(FileName))
                                    {
                                        while (IsFileExist(FileName))
                                        {
                                            FileName = Path.GetFileNameWithoutExtension(Row["Path"].ToString()) + "(" + i++ + ")" + Path.GetExtension(Row["Path"].ToString());
                                        }
                                    }

                                    //get file size
                                    try
                                    {
                                        fi = new FileInfo(Row["Path"].ToString());
                                    }
                                    catch
                                    {
                                        return false;
                                    }

                                    var iFileSize = fi.Length;

                                    //load file to ftp
                                    if (FM.UploadFile(Row["Path"].ToString(), Configs.DocumentsPath + "/Документооборот/Документы" +
                                                      "/" + FileName, Configs.FTPType) == false)
                                    {
                                        FM.DeleteFile(Configs.DocumentsPath + "/Документооборот/Документы" +
                                                      "/" + FileName, Configs.FTPType);

                                        return false;
                                    }

                                    var NewRow = sDT.NewRow();
                                    NewRow["FileName"] = FileName;
                                    NewRow["DocumentID"] = DT.Rows[0]["InnerDocumentID"];
                                    NewRow["FileSize"] = fi.Length;
                                    NewRow["DocumentCategoryID"] = 0;
                                    sDT.Rows.Add(NewRow);
                                }

                                sDA.Update(sDT);
                            }
                        }
                    }
                }
            }

            using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM DocumentsRecipients", ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        using (var sDA = new SqlDataAdapter("SELECT TOP 0 * FROM SubscribesRecords", ConnectionStrings.LightConnectionString))
                        {
                            using (var sCB = new SqlCommandBuilder(sDA))
                            {
                                using (var sDT = new DataTable())
                                {
                                    sDA.Fill(sDT);

                                    foreach (DataRow Row in RecipientsDT.Rows)
                                    {
                                        var NewRow = DT.NewRow();
                                        NewRow["UserID"] = Row["UserID"];
                                        NewRow["DocumentID"] = InnerDocumentID;
                                        NewRow["DocumentCategoryID"] = 0;
                                        DT.Rows.Add(NewRow);

                                        var sNewRow = sDT.NewRow();
                                        sNewRow["SubscribesItemID"] = 18;
                                        sNewRow["TableItemID"] = InnerDocumentID;
                                        sNewRow["UserID"] = Row["UserID"];
                                        sNewRow["UserTypeID"] = 0;
                                        sDT.Rows.Add(sNewRow);
                                    }

                                    if (RecipientsDT.Select("UserID = " + Security.CurrentUserID).Count() == 0)
                                    {
                                        var NewRow = DT.NewRow();
                                        NewRow["UserID"] = Security.CurrentUserID;
                                        NewRow["DocumentID"] = InnerDocumentID;
                                        NewRow["DocumentCategoryID"] = 0;
                                        DT.Rows.Add(NewRow);

                                        var sNewRow = sDT.NewRow();
                                        sNewRow["SubscribesItemID"] = 18;
                                        sNewRow["TableItemID"] = InnerDocumentID;
                                        sNewRow["UserID"] = Security.CurrentUserID;
                                        sNewRow["UserTypeID"] = 0;
                                        sDT.Rows.Add(sNewRow);
                                    }

                                    sDA.Update(sDT);
                                }
                            }
                        }



                        DA.Update(DT);
                    }
                }
            }

            return true;
        }

        public bool EditInnerDocument(int InnerDocumentID, int DocumentTypeID, int UserID, DataTable RecipientsDT,
                                       string Description, string RegNumber, int DocumentStateID, DataTable FilesDT, ref int CurrentUploadedFile, int FactoryID)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM InnerDocuments WHERE InnerDocumentID = " + InnerDocumentID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        DT.Rows[0]["DocumentTypeID"] = DocumentTypeID;
                        DT.Rows[0]["UserID"] = UserID;
                        DT.Rows[0]["Description"] = Description;
                        DT.Rows[0]["RegNumber"] = RegNumber;
                        DT.Rows[0]["DocumentsStateID"] = DocumentStateID;
                        DT.Rows[0]["FactoryID"] = FactoryID;

                        DA.Update(DT);
                    }
                }
            }


            using (var sDA = new SqlDataAdapter("SELECT * FROM DocumentsFiles WHERE DocumentCategoryID = 0 AND DocumentID = " + InnerDocumentID, ConnectionStrings.LightConnectionString))
            {
                using (var sCB = new SqlCommandBuilder(sDA))
                {
                    using (var sDT = new DataTable())
                    {
                        sDA.Fill(sDT);

                        foreach (DataRow Row in sDT.Rows)
                        {

                            if (FilesDT.Select("IsNew is null" + " AND FileName = '" + Row["FileName"] + "'").Count() == 0)//deleted
                            {
                                FM.DeleteFile(Configs.DocumentsPath + "/Документооборот/Документы" +
                                            "/" + Row["FileName"], Configs.FTPType);

                                Row.Delete();
                            }
                        }

                        foreach (var Row in FilesDT.Select("IsNew = true"))
                        {
                            FileInfo fi;

                            CurrentUploadedFile++;

                            var FileName = "";

                            var i = 1;

                            FileName = Row["FileName"].ToString();

                            if (IsFileExist(FileName))
                            {
                                while (IsFileExist(FileName))
                                {
                                    FileName = Path.GetFileNameWithoutExtension(Row["Path"].ToString()) + "(" + i++ + ")" + Path.GetExtension(Row["Path"].ToString());
                                }
                            }

                            //get file size
                            try
                            {
                                fi = new FileInfo(Row["Path"].ToString());
                            }
                            catch
                            {
                                return false;
                            }

                            var iFileSize = fi.Length;

                            //load file to ftp
                            if (FM.UploadFile(Row["Path"].ToString(), Configs.DocumentsPath + "/Документооборот/Документы" +
                                                "/" + FileName, Configs.FTPType) == false)
                            {
                                FM.DeleteFile(Configs.DocumentsPath + "/Документооборот/Документы" +
                                                "/" + FileName, Configs.FTPType);

                                return false;
                            }

                            var NewRow = sDT.NewRow();
                            NewRow["FileName"] = FileName;
                            NewRow["DocumentID"] = InnerDocumentID;
                            NewRow["FileSize"] = fi.Length;
                            NewRow["DocumentCategoryID"] = 0;
                            sDT.Rows.Add(NewRow);
                        }

                        sDA.Update(sDT);
                    }
                }

                using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsRecipients WHERE DocumentCategoryID = 0 AND DocumentID = " + InnerDocumentID, ConnectionStrings.LightConnectionString))
                {
                    using (var CB = new SqlCommandBuilder(DA))
                    {
                        using (var DT = new DataTable())
                        {
                            DA.Fill(DT);

                            using (var rDA = new SqlDataAdapter("SELECT * FROM SubscribesRecords WHERE SubscribesItemID = 18 AND TableItemID = " + InnerDocumentID, ConnectionStrings.LightConnectionString))
                            {
                                using (var rCB = new SqlCommandBuilder(rDA))
                                {
                                    using (var rDT = new DataTable())
                                    {
                                        rDA.Fill(rDT);

                                        for (var i = 0; i < DT.Rows.Count; i++)
                                        {
                                            if (Convert.ToInt32(DT.Rows[i]["UserID"]) == Security.CurrentUserID)
                                                continue;

                                            if (RecipientsDT.Select("UserID = " + DT.Rows[i]["UserID"]).Count() == 0)
                                            {
                                                var rRow = rDT.Select("UserID = " + DT.Rows[i]["UserID"]);

                                                if (rRow.Count() > 0)
                                                    rRow[0].Delete();

                                                RemoveCommentsSubscribes(Convert.ToInt32(DT.Rows[i]["UserID"]), InnerDocumentID, 0);

                                                DT.Rows[i].Delete();
                                            }
                                        }

                                        DA.Update(DT);
                                        rDA.Update(rDT);

                                        foreach (DataRow Row in RecipientsDT.Rows)
                                        {
                                            if (DT.Select("UserID = " + Row["UserID"]).Count() == 0)
                                            {
                                                var NewRow = DT.NewRow();
                                                NewRow["UserID"] = Row["UserID"];
                                                NewRow["DocumentID"] = InnerDocumentID;
                                                NewRow["DocumentCategoryID"] = 0;
                                                DT.Rows.Add(NewRow);

                                                var rNewRow = rDT.NewRow();
                                                rNewRow["SubscribesItemID"] = 18;
                                                rNewRow["TableItemID"] = InnerDocumentID;
                                                rNewRow["UserID"] = Row["UserID"];
                                                rNewRow["UserTypeID"] = 0;
                                                rDT.Rows.Add(rNewRow);
                                            }
                                        }

                                        rDA.Update(rDT);

                                        DA.Update(DT);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return true;
        }

        public void GetEditInnerDocument(int InnerDocumentID, ref int DocumentTypeID, ref DateTime DateTime, ref int UserID,
                                          DataTable RecipientsDT, ref string Description, ref string RegNumber, ref int DocumentStateID, DataTable FilesDT, ref int FactoryID)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM InnerDocuments WHERE InnerDocumentID = " + InnerDocumentID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    DocumentTypeID = Convert.ToInt32(DT.Rows[0]["DocumentTypeID"]);
                    DateTime = Convert.ToDateTime(DT.Rows[0]["DateTime"]);
                    UserID = Convert.ToInt32(DT.Rows[0]["UserID"]);
                    Description = DT.Rows[0]["Description"].ToString();
                    RegNumber = DT.Rows[0]["RegNumber"].ToString();
                    DocumentStateID = Convert.ToInt32(DT.Rows[0]["DocumentsStateID"]);
                    FactoryID = Convert.ToInt32(DT.Rows[0]["FactoryID"]);
                }
            }

            using (var DA = new SqlDataAdapter("SELECT FileName FROM DocumentsFiles WHERE DocumentCategoryID = 0 AND DocumentID = " + InnerDocumentID, ConnectionStrings.LightConnectionString))
            {
                DA.Fill(FilesDT);
            }

            using (var DA = new SqlDataAdapter("SELECT UserID FROM DocumentsRecipients WHERE DocumentCategoryID = 0 AND DocumentID = " + InnerDocumentID, ConnectionStrings.LightConnectionString))
            {
                DA.Fill(RecipientsDT);
            }
        }



        public bool AddOuterDocument(int DocumentTypeID, DateTime DateTime, int UserID, int CorrespondentID, DataTable RecipientsDT,
                                     string Description, string RegNumber, int DocumentStateID, DataTable FilesDT, ref int CurrentUploadedFile, int FactoryID)
        {
            using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM OuterDocuments", ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        var NewRow = DT.NewRow();
                        NewRow["DocumentTypeID"] = DocumentTypeID;
                        NewRow["DateTime"] = DateTime;
                        NewRow["UserID"] = UserID;
                        NewRow["CorrespondentID"] = CorrespondentID;
                        NewRow["Description"] = Description;
                        NewRow["RegNumber"] = RegNumber;
                        NewRow["DocumentsStateID"] = DocumentStateID;
                        NewRow["FactoryID"] = FactoryID;
                        DT.Rows.Add(NewRow);

                        DA.Update(DT);
                    }
                }
            }

            var OuterDocumentID = -1;

            using (var DA = new SqlDataAdapter("SELECT OuterDocumentID FROM OuterDocuments WHERE DateTime = '" + DateTime.ToString("yyyy-MM-dd HH:mm:ss.fff") + "'", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    OuterDocumentID = Convert.ToInt32(DT.Rows[0]["OuterDocumentID"]);

                    using (var sDA = new SqlDataAdapter("SELECT TOP 0 * FROM DocumentsFiles", ConnectionStrings.LightConnectionString))
                    {
                        using (var sCB = new SqlCommandBuilder(sDA))
                        {
                            using (var sDT = new DataTable())
                            {
                                sDA.Fill(sDT);

                                foreach (DataRow Row in FilesDT.Rows)
                                {
                                    FileInfo fi;

                                    CurrentUploadedFile++;

                                    var FileName = "";

                                    var i = 1;

                                    FileName = Row["FileName"].ToString();

                                    if (IsFileExist(FileName))
                                    {
                                        while (IsFileExist(FileName))
                                        {
                                            FileName = Path.GetFileNameWithoutExtension(Row["Path"].ToString()) + "(" + i++ + ")" + Path.GetExtension(Row["Path"].ToString());
                                        }
                                    }

                                    //get file size
                                    try
                                    {
                                        fi = new FileInfo(Row["Path"].ToString());
                                    }
                                    catch
                                    {
                                        return false;
                                    }

                                    var iFileSize = fi.Length;

                                    //load file to ftp
                                    if (FM.UploadFile(Row["Path"].ToString(), Configs.DocumentsPath + "/Документооборот/Документы" +
                                                      "/" + FileName, Configs.FTPType) == false)
                                    {
                                        FM.DeleteFile(Configs.DocumentsPath + "/Документооборот/Документы" +
                                                      "/" + FileName, Configs.FTPType);

                                        return false;
                                    }

                                    var NewRow = sDT.NewRow();
                                    NewRow["FileName"] = FileName;
                                    NewRow["DocumentID"] = DT.Rows[0]["OuterDocumentID"];
                                    NewRow["FileSize"] = fi.Length;
                                    NewRow["DocumentCategoryID"] = 2;
                                    sDT.Rows.Add(NewRow);


                                }

                                sDA.Update(sDT);
                            }
                        }
                    }
                }
            }

            using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM DocumentsRecipients", ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        using (var sDA = new SqlDataAdapter("SELECT TOP 0 * FROM SubscribesRecords", ConnectionStrings.LightConnectionString))
                        {
                            using (var sCB = new SqlCommandBuilder(sDA))
                            {
                                using (var sDT = new DataTable())
                                {
                                    sDA.Fill(sDT);

                                    foreach (DataRow Row in RecipientsDT.Rows)
                                    {
                                        var NewRow = DT.NewRow();
                                        NewRow["UserID"] = Row["UserID"];
                                        NewRow["DocumentID"] = OuterDocumentID;
                                        NewRow["DocumentCategoryID"] = 2;
                                        DT.Rows.Add(NewRow);

                                        var sNewRow = sDT.NewRow();
                                        sNewRow["SubscribesItemID"] = 20;
                                        sNewRow["TableItemID"] = OuterDocumentID;
                                        sNewRow["UserID"] = Row["UserID"];
                                        sNewRow["UserTypeID"] = 0;
                                        sDT.Rows.Add(sNewRow);
                                    }

                                    if (RecipientsDT.Select("UserID = " + Security.CurrentUserID).Count() == 0)
                                    {
                                        var NewRow = DT.NewRow();
                                        NewRow["UserID"] = Security.CurrentUserID;
                                        NewRow["DocumentID"] = OuterDocumentID;
                                        NewRow["DocumentCategoryID"] = 2;
                                        DT.Rows.Add(NewRow);

                                        var sNewRow = sDT.NewRow();
                                        sNewRow["SubscribesItemID"] = 20;
                                        sNewRow["TableItemID"] = OuterDocumentID;
                                        sNewRow["UserID"] = Security.CurrentUserID;
                                        sNewRow["UserTypeID"] = 0;
                                        sDT.Rows.Add(sNewRow);
                                    }

                                    sDA.Update(sDT);
                                }
                            }
                        }



                        DA.Update(DT);
                    }
                }
            }

            return true;
        }

        public bool EditOuterDocument(int OuterDocumentID, int DocumentTypeID, int UserID, int CorrespondentID, DataTable RecipientsDT,
                                     string Description, string RegNumber, int DocumentStateID, DataTable FilesDT, ref int CurrentUploadedFile, int FactoryID)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM OuterDocuments WHERE OuterDocumentID = " + OuterDocumentID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        DT.Rows[0]["DocumentTypeID"] = DocumentTypeID;
                        DT.Rows[0]["UserID"] = UserID;
                        DT.Rows[0]["CorrespondentID"] = CorrespondentID;
                        DT.Rows[0]["Description"] = Description;
                        DT.Rows[0]["RegNumber"] = RegNumber;
                        DT.Rows[0]["DocumentsStateID"] = DocumentStateID;
                        DT.Rows[0]["FactoryID"] = FactoryID;

                        DA.Update(DT);
                    }
                }
            }

            using (var sDA = new SqlDataAdapter("SELECT * FROM DocumentsFiles WHERE DocumentCategoryID = 2 AND DocumentID = " + OuterDocumentID, ConnectionStrings.LightConnectionString))
            {
                using (var sCB = new SqlCommandBuilder(sDA))
                {
                    using (var sDT = new DataTable())
                    {
                        sDA.Fill(sDT);

                        foreach (DataRow Row in sDT.Rows)
                        {

                            if (FilesDT.Select("IsNew is null" + " AND FileName = '" + Row["FileName"] + "'").Count() == 0)//deleted
                            {
                                FM.DeleteFile(Configs.DocumentsPath + "/Документооборот/Документы" +
                                            "/" + Row["FileName"], Configs.FTPType);

                                Row.Delete();
                            }
                        }

                        foreach (var Row in FilesDT.Select("IsNew = true"))
                        {
                            FileInfo fi;

                            CurrentUploadedFile++;

                            var FileName = "";

                            var i = 1;

                            FileName = Row["FileName"].ToString();

                            if (IsFileExist(FileName))
                            {
                                while (IsFileExist(FileName))
                                {
                                    FileName = Path.GetFileNameWithoutExtension(Row["Path"].ToString()) + "(" + i++ + ")" + Path.GetExtension(Row["Path"].ToString());
                                }
                            }

                            //get file size
                            try
                            {
                                fi = new FileInfo(Row["Path"].ToString());
                            }
                            catch
                            {
                                return false;
                            }

                            var iFileSize = fi.Length;

                            //load file to ftp
                            if (FM.UploadFile(Row["Path"].ToString(), Configs.DocumentsPath + "/Документооборот/Документы" +
                                                "/" + FileName, Configs.FTPType) == false)
                            {
                                FM.DeleteFile(Configs.DocumentsPath + "/Документооборот/Документы" +
                                                "/" + FileName, Configs.FTPType);

                                return false;
                            }

                            var NewRow = sDT.NewRow();
                            NewRow["FileName"] = FileName;
                            NewRow["DocumentID"] = OuterDocumentID;
                            NewRow["FileSize"] = fi.Length;
                            NewRow["DocumentCategoryID"] = 2;
                            sDT.Rows.Add(NewRow);
                        }

                        sDA.Update(sDT);
                    }
                }

                using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsRecipients WHERE DocumentCategoryID = 2 AND DocumentID = " + OuterDocumentID, ConnectionStrings.LightConnectionString))
                {
                    using (var CB = new SqlCommandBuilder(DA))
                    {
                        using (var DT = new DataTable())
                        {
                            DA.Fill(DT);

                            using (var rDA = new SqlDataAdapter("SELECT * FROM SubscribesRecords WHERE SubscribesItemID = 20 AND TableItemID = " + OuterDocumentID, ConnectionStrings.LightConnectionString))
                            {
                                using (var rCB = new SqlCommandBuilder(rDA))
                                {
                                    using (var rDT = new DataTable())
                                    {
                                        rDA.Fill(rDT);

                                        for (var i = 0; i < DT.Rows.Count; i++)
                                        {
                                            if (Convert.ToInt32(DT.Rows[i]["UserID"]) == Security.CurrentUserID)
                                                continue;

                                            if (RecipientsDT.Select("UserID = " + DT.Rows[i]["UserID"]).Count() == 0)
                                            {
                                                var rRow = rDT.Select("UserID = " + DT.Rows[i]["UserID"]);

                                                if (rRow.Count() > 0)
                                                    rRow[0].Delete();

                                                RemoveCommentsSubscribes(Convert.ToInt32(DT.Rows[i]["UserID"]), OuterDocumentID, 2);

                                                DT.Rows[i].Delete();
                                            }
                                        }

                                        DA.Update(DT);
                                        rDA.Update(rDT);

                                        foreach (DataRow Row in RecipientsDT.Rows)
                                        {
                                            if (DT.Select("UserID = " + Row["UserID"]).Count() == 0)
                                            {
                                                var NewRow = DT.NewRow();
                                                NewRow["UserID"] = Row["UserID"];
                                                NewRow["DocumentID"] = OuterDocumentID;
                                                NewRow["DocumentCategoryID"] = 2;
                                                DT.Rows.Add(NewRow);

                                                var rNewRow = rDT.NewRow();
                                                rNewRow["SubscribesItemID"] = 20;
                                                rNewRow["TableItemID"] = OuterDocumentID;
                                                rNewRow["UserID"] = Row["UserID"];
                                                rNewRow["UserTypeID"] = 0;
                                                rDT.Rows.Add(rNewRow);
                                            }
                                        }

                                        rDA.Update(rDT);

                                        DA.Update(DT);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return true;
        }

        public void GetEditOuterDocument(int OuterDocumentID, ref int DocumentTypeID, ref DateTime DateTime, ref int UserID, ref int CorrespondentID,
                                         DataTable RecipientsDT, ref string Description, ref string RegNumber, ref int DocumentStateID, DataTable FilesDT, ref int FactoryID)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM OuterDocuments WHERE OuterDocumentID = " + OuterDocumentID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    DocumentTypeID = Convert.ToInt32(DT.Rows[0]["DocumentTypeID"]);
                    DateTime = Convert.ToDateTime(DT.Rows[0]["DateTime"]);
                    UserID = Convert.ToInt32(DT.Rows[0]["UserID"]);
                    CorrespondentID = Convert.ToInt32(DT.Rows[0]["CorrespondentID"]);
                    Description = DT.Rows[0]["Description"].ToString();
                    RegNumber = DT.Rows[0]["RegNumber"].ToString();
                    DocumentStateID = Convert.ToInt32(DT.Rows[0]["DocumentsStateID"]);
                    FactoryID = Convert.ToInt32(DT.Rows[0]["FactoryID"]);
                }
            }

            using (var DA = new SqlDataAdapter("SELECT FileName FROM DocumentsFiles WHERE DocumentCategoryID = 2 AND DocumentID = " + OuterDocumentID, ConnectionStrings.LightConnectionString))
            {
                DA.Fill(FilesDT);
            }

            using (var DA = new SqlDataAdapter("SELECT UserID FROM DocumentsRecipients WHERE DocumentCategoryID = 2 AND DocumentID = " + OuterDocumentID, ConnectionStrings.LightConnectionString))
            {
                DA.Fill(RecipientsDT);
            }
        }



        public bool AddIncomeDocument(int DocumentTypeID, DateTime DateTime, int UserID, int CorrespondentID, DataTable RecipientsDT,
                                       string Description, string RegNumber, string IncomeNumber, int DocumentStateID, DataTable FilesDT, ref int CurrentUploadedFile, int FactoryID)
        {
            using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM IncomeDocuments", ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        var NewRow = DT.NewRow();
                        NewRow["DocumentTypeID"] = DocumentTypeID;
                        NewRow["DateTime"] = DateTime;
                        NewRow["UserID"] = UserID;
                        NewRow["CorrespondentID"] = CorrespondentID;
                        NewRow["Description"] = Description;
                        NewRow["RegNumber"] = RegNumber;
                        NewRow["IncomeNumber"] = IncomeNumber;
                        NewRow["DocumentsStateID"] = DocumentStateID;
                        NewRow["FactoryID"] = FactoryID;
                        DT.Rows.Add(NewRow);

                        DA.Update(DT);
                    }
                }
            }

            var IncomeDocumentID = -1;

            using (var DA = new SqlDataAdapter("SELECT IncomeDocumentID FROM IncomeDocuments WHERE DateTime = '" + DateTime.ToString("yyyy-MM-dd HH:mm:ss.fff") + "'", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    IncomeDocumentID = Convert.ToInt32(DT.Rows[0]["IncomeDocumentID"]);

                    using (var sDA = new SqlDataAdapter("SELECT TOP 0 * FROM DocumentsFiles", ConnectionStrings.LightConnectionString))
                    {
                        using (var sCB = new SqlCommandBuilder(sDA))
                        {
                            using (var sDT = new DataTable())
                            {
                                sDA.Fill(sDT);

                                foreach (DataRow Row in FilesDT.Rows)
                                {
                                    FileInfo fi;

                                    CurrentUploadedFile++;

                                    var FileName = "";

                                    var i = 1;

                                    FileName = Row["FileName"].ToString();

                                    if (IsFileExist(FileName))
                                    {
                                        while (IsFileExist(FileName))
                                        {
                                            FileName = Path.GetFileNameWithoutExtension(Row["Path"].ToString()) + "(" + i++ + ")" + Path.GetExtension(Row["Path"].ToString());
                                        }
                                    }

                                    //get file size
                                    try
                                    {
                                        fi = new FileInfo(Row["Path"].ToString());
                                    }
                                    catch
                                    {
                                        return false;
                                    }

                                    var iFileSize = fi.Length;

                                    //load file to ftp
                                    if (FM.UploadFile(Row["Path"].ToString(), Configs.DocumentsPath + "/Документооборот/Документы" +
                                                      "/" + FileName, Configs.FTPType) == false)
                                    {
                                        FM.DeleteFile(Configs.DocumentsPath + "/Документооборот/Документы" +
                                                      "/" + FileName, Configs.FTPType);

                                        return false;
                                    }

                                    var NewRow = sDT.NewRow();
                                    NewRow["FileName"] = FileName;
                                    NewRow["DocumentID"] = DT.Rows[0]["IncomeDocumentID"];
                                    NewRow["FileSize"] = fi.Length;
                                    NewRow["DocumentCategoryID"] = 1;
                                    sDT.Rows.Add(NewRow);


                                }

                                sDA.Update(sDT);
                            }
                        }
                    }
                }
            }

            using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM DocumentsRecipients", ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        using (var sDA = new SqlDataAdapter("SELECT TOP 0 * FROM SubscribesRecords", ConnectionStrings.LightConnectionString))
                        {
                            using (var sCB = new SqlCommandBuilder(sDA))
                            {
                                using (var sDT = new DataTable())
                                {
                                    sDA.Fill(sDT);

                                    foreach (DataRow Row in RecipientsDT.Rows)
                                    {
                                        var NewRow = DT.NewRow();
                                        NewRow["UserID"] = Row["UserID"];
                                        NewRow["DocumentID"] = IncomeDocumentID;
                                        NewRow["DocumentCategoryID"] = 1;
                                        DT.Rows.Add(NewRow);

                                        var sNewRow = sDT.NewRow();
                                        sNewRow["SubscribesItemID"] = 19;
                                        sNewRow["TableItemID"] = IncomeDocumentID;
                                        sNewRow["UserID"] = Row["UserID"];
                                        sNewRow["UserTypeID"] = 0;
                                        sDT.Rows.Add(sNewRow);
                                    }

                                    if (RecipientsDT.Select("UserID = " + Security.CurrentUserID).Count() == 0)
                                    {
                                        var NewRow = DT.NewRow();
                                        NewRow["UserID"] = Security.CurrentUserID;
                                        NewRow["DocumentID"] = IncomeDocumentID;
                                        NewRow["DocumentCategoryID"] = 1;
                                        DT.Rows.Add(NewRow);

                                        var sNewRow = sDT.NewRow();
                                        sNewRow["SubscribesItemID"] = 19;
                                        sNewRow["TableItemID"] = IncomeDocumentID;
                                        sNewRow["UserID"] = Security.CurrentUserID;
                                        sNewRow["UserTypeID"] = 0;
                                        sDT.Rows.Add(sNewRow);
                                    }

                                    sDA.Update(sDT);
                                }
                            }
                        }



                        DA.Update(DT);
                    }
                }
            }

            return true;
        }

        public bool EditIncomeDocument(int IncomeDocumentID, int DocumentTypeID, int UserID, int CorrespondentID, DataTable RecipientsDT,
                                        string Description, string RegNumber, string IncomeNumber, int DocumentStateID, DataTable FilesDT, ref int CurrentUploadedFile, int FactoryID)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM IncomeDocuments WHERE IncomeDocumentID = " + IncomeDocumentID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        DT.Rows[0]["DocumentTypeID"] = DocumentTypeID;
                        DT.Rows[0]["UserID"] = UserID;
                        DT.Rows[0]["CorrespondentID"] = CorrespondentID;
                        DT.Rows[0]["Description"] = Description;
                        DT.Rows[0]["RegNumber"] = RegNumber;
                        DT.Rows[0]["IncomeNumber"] = IncomeNumber;
                        DT.Rows[0]["DocumentsStateID"] = DocumentStateID;
                        DT.Rows[0]["FactoryID"] = FactoryID;

                        DA.Update(DT);
                    }
                }
            }

            using (var sDA = new SqlDataAdapter("SELECT * FROM DocumentsFiles WHERE DocumentCategoryID = 1 AND DocumentID = " + IncomeDocumentID, ConnectionStrings.LightConnectionString))
            {
                using (var sCB = new SqlCommandBuilder(sDA))
                {
                    using (var sDT = new DataTable())
                    {
                        sDA.Fill(sDT);

                        foreach (DataRow Row in sDT.Rows)
                        {

                            if (FilesDT.Select("IsNew is null" + " AND FileName = '" + Row["FileName"] + "'").Count() == 0)//deleted
                            {
                                FM.DeleteFile(Configs.DocumentsPath + "/Документооборот/Документы" +
                                            "/" + Row["FileName"], Configs.FTPType);

                                Row.Delete();
                            }
                        }

                        foreach (var Row in FilesDT.Select("IsNew = true"))
                        {
                            FileInfo fi;

                            CurrentUploadedFile++;

                            var FileName = "";

                            var i = 1;

                            FileName = Row["FileName"].ToString();

                            if (IsFileExist(FileName))
                            {
                                while (IsFileExist(FileName))
                                {
                                    FileName = Path.GetFileNameWithoutExtension(Row["Path"].ToString()) + "(" + i++ + ")" + Path.GetExtension(Row["Path"].ToString());
                                }
                            }

                            //get file size
                            try
                            {
                                fi = new FileInfo(Row["Path"].ToString());
                            }
                            catch
                            {
                                return false;
                            }

                            var iFileSize = fi.Length;

                            //load file to ftp
                            if (FM.UploadFile(Row["Path"].ToString(), Configs.DocumentsPath + "/Документооборот/Документы" +
                                                "/" + FileName, Configs.FTPType) == false)
                            {
                                FM.DeleteFile(Configs.DocumentsPath + "/Документооборот/Документы" +
                                                "/" + FileName, Configs.FTPType);

                                return false;
                            }

                            var NewRow = sDT.NewRow();
                            NewRow["FileName"] = FileName;
                            NewRow["DocumentID"] = IncomeDocumentID;
                            NewRow["FileSize"] = fi.Length;
                            NewRow["DocumentCategoryID"] = 1;
                            sDT.Rows.Add(NewRow);
                        }

                        sDA.Update(sDT);
                    }
                }

                using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsRecipients WHERE DocumentCategoryID = 1 AND DocumentID = " + IncomeDocumentID, ConnectionStrings.LightConnectionString))
                {
                    using (var CB = new SqlCommandBuilder(DA))
                    {
                        using (var DT = new DataTable())
                        {
                            DA.Fill(DT);

                            using (var rDA = new SqlDataAdapter("SELECT * FROM SubscribesRecords WHERE SubscribesItemID = 19 AND TableItemID = " + IncomeDocumentID, ConnectionStrings.LightConnectionString))
                            {
                                using (var rCB = new SqlCommandBuilder(rDA))
                                {
                                    using (var rDT = new DataTable())
                                    {
                                        rDA.Fill(rDT);

                                        for (var i = 0; i < DT.Rows.Count; i++)
                                        {
                                            if (Convert.ToInt32(DT.Rows[i]["UserID"]) == Security.CurrentUserID)
                                                continue;

                                            if (RecipientsDT.Select("UserID = " + DT.Rows[i]["UserID"]).Count() == 0)
                                            {
                                                var rRow = rDT.Select("UserID = " + DT.Rows[i]["UserID"]);

                                                if (rRow.Count() > 0)
                                                    rRow[0].Delete();

                                                RemoveCommentsSubscribes(Convert.ToInt32(DT.Rows[i]["UserID"]), IncomeDocumentID, 1);

                                                DT.Rows[i].Delete();
                                            }
                                        }

                                        DA.Update(DT);
                                        rDA.Update(rDT);

                                        foreach (DataRow Row in RecipientsDT.Rows)
                                        {
                                            if (DT.Select("UserID = " + Row["UserID"]).Count() == 0)
                                            {
                                                var NewRow = DT.NewRow();
                                                NewRow["UserID"] = Row["UserID"];
                                                NewRow["DocumentID"] = IncomeDocumentID;
                                                NewRow["DocumentCategoryID"] = 1;
                                                DT.Rows.Add(NewRow);

                                                var rNewRow = rDT.NewRow();
                                                rNewRow["SubscribesItemID"] = 19;
                                                rNewRow["TableItemID"] = IncomeDocumentID;
                                                rNewRow["UserID"] = Row["UserID"];
                                                rNewRow["UserTypeID"] = 0;
                                                rDT.Rows.Add(rNewRow);
                                            }
                                        }

                                        rDA.Update(rDT);

                                        DA.Update(DT);
                                    }
                                }
                            }
                        }
                    }
                }
            }

            return true;
        }

        public void GetEditIncomeDocument(int IncomeDocumentID, ref int DocumentTypeID, ref DateTime DateTime, ref int UserID, ref int CorrespondentID,
                                           DataTable RecipientsDT, ref string Description, ref string RegNumber, ref string IncomeNumber, ref int DocumentStateID, DataTable FilesDT, ref int FactoryID)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM IncomeDocuments WHERE IncomeDocumentID = " + IncomeDocumentID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    DocumentTypeID = Convert.ToInt32(DT.Rows[0]["DocumentTypeID"]);
                    DateTime = Convert.ToDateTime(DT.Rows[0]["DateTime"]);
                    UserID = Convert.ToInt32(DT.Rows[0]["UserID"]);
                    CorrespondentID = Convert.ToInt32(DT.Rows[0]["CorrespondentID"]);
                    Description = DT.Rows[0]["Description"].ToString();
                    RegNumber = DT.Rows[0]["RegNumber"].ToString();
                    IncomeNumber = DT.Rows[0]["IncomeNumber"].ToString();
                    DocumentStateID = Convert.ToInt32(DT.Rows[0]["DocumentsStateID"]);
                    FactoryID = Convert.ToInt32(DT.Rows[0]["FactoryID"]);
                }
            }

            using (var DA = new SqlDataAdapter("SELECT FileName FROM DocumentsFiles WHERE DocumentCategoryID = 1 AND DocumentID = " + IncomeDocumentID, ConnectionStrings.LightConnectionString))
            {
                DA.Fill(FilesDT);
            }

            using (var DA = new SqlDataAdapter("SELECT UserID FROM DocumentsRecipients WHERE DocumentCategoryID = 1 AND DocumentID = " + IncomeDocumentID, ConnectionStrings.LightConnectionString))
            {
                DA.Fill(RecipientsDT);
            }
        }


        private void RemoveCommentsSubscribes(int UserID, int DocumentID, int DocumentCategoryID)
        {
            using (var DA = new SqlDataAdapter("DELETE FROM SubscribesRecords WHERE UserID = " + UserID + " AND SubscribesItemID = 21 AND TableItemID IN (SELECT DocumentCommentID FROM DocumentsComments WHERE DocumentCategoryID = " + DocumentCategoryID + " AND DocumentID = " + DocumentID + ")", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }
        }

        private void AddCommentsSubscribes(int DocumentCommentID, int DocumentID, int DocumentCategoryID)
        {
            using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM SubscribesRecords", ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        using (var rDA = new SqlDataAdapter("SELECT * FROM DocumentsRecipients WHERE DocumentCategoryID = " + DocumentCategoryID + " AND DocumentID = " + DocumentID, ConnectionStrings.LightConnectionString))
                        {
                            using (var rDT = new DataTable())
                            {
                                rDA.Fill(rDT);

                                foreach (DataRow Row in rDT.Rows)
                                {
                                    var NewRow = DT.NewRow();
                                    NewRow["SubscribesItemID"] = 21;
                                    NewRow["UserID"] = Row["UserID"];
                                    NewRow["TableItemID"] = DocumentCommentID;
                                    DT.Rows.Add(NewRow);
                                }

                                DA.Update(DT);
                            }
                        }
                    }
                }
            }
        }

        public bool IsFileExist(string Name)
        {
            using (var DA = new SqlDataAdapter("SELECT TOP 1 FileName FROM DocumentsFiles WHERE FileName = '" + Name + "'", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    return DA.Fill(DT) > 0;
                }
            }
        }

        public bool AddCorrespondent(string Name)
        {
            using (var DA = new SqlDataAdapter("SELECT TOP 1 * FROM Correspondents WHERE CorrespondentName = '" + Name + "'", ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        if (DT.Rows.Count > 0)
                            return false;

                        var NewRow = DT.NewRow();
                        NewRow["CorrespondentName"] = Name;
                        DT.Rows.Add(NewRow);

                        DA.Update(DT);
                    }
                }
            }

            return true;
        }


        public bool OpenFile(int DocumentFileID)
        {
            var tempFolder = Environment.GetEnvironmentVariable("TEMP");

            var bOK = false;

            using (var fDA = new SqlDataAdapter("SELECT * FROM DocumentsFiles WHERE DocumentFileID = " + DocumentFileID, ConnectionStrings.LightConnectionString))
            {
                using (var fDT = new DataTable())
                {
                    fDA.Fill(fDT);

                    bOK = FM.DownloadFile(Configs.DocumentsPath + "Документооборот/Документы/" + fDT.Rows[0]["FileName"],
                                    tempFolder + "//" +
                                    fDT.Rows[0]["FileName"], Convert.ToInt64(fDT.Rows[0]["FileSize"]), Configs.FTPType);

                    if (bOK)
                        try
                        {
                            var myProcess = new Process();
                            myProcess.StartInfo.UseShellExecute = true;
                            myProcess.StartInfo.FileName = tempFolder + "//" +
                                                                fDT.Rows[0]["FileName"];
                            myProcess.StartInfo.CreateNoWindow = true;
                            myProcess.Start();
                        }
                        catch (Exception ex)
                        {
                            InfiniumMessages.SendMessage("Ошибка скачивания файла документов UserID = " + Security.CurrentUserID + ", ID = " + DocumentFileID + ", Name = " + fDT.Rows[0]["FileName"] + " exception = " + ex.Message, 321);
                            return false;
                        }
                    else
                    {
                        InfiniumMessages.SendMessage("Ошибка скачивания файла документов UserID = " + Security.CurrentUserID + ", ID = " + DocumentFileID + ", Name = " + fDT.Rows[0]["FileName"], 321);
                        MessageBox.Show("Ошибка скачивания файла. Отчет отправлен разработчикам. Приносим извинения за неудобства.");
                    }
                }

            }

            return bOK;
        }

        public bool OpenCommentFile(int DocumentCommentFileID)
        {
            var tempFolder = Environment.GetEnvironmentVariable("TEMP");


            var bOK = false;

            using (var fDA = new SqlDataAdapter("SELECT * FROM DocumentsCommentsFiles WHERE DocumentCommentFileID = " + DocumentCommentFileID, ConnectionStrings.LightConnectionString))
            {
                using (var fDT = new DataTable())
                {
                    fDA.Fill(fDT);

                    bOK = FM.DownloadFile(Configs.DocumentsPath + "Документооборот/Комментарии/" + fDT.Rows[0]["FileName"],
                                    tempFolder + "//" +
                                    fDT.Rows[0]["FileName"], Convert.ToInt64(fDT.Rows[0]["FileSize"]), Configs.FTPType);

                    if (bOK)
                    {
                        Process.Start(tempFolder + "//" +
                                      fDT.Rows[0]["FileName"]);
                    }
                    else
                    {
                        InfiniumMessages.SendMessage("Ошибка скачивания файла комментария документов UserID = " + Security.CurrentUserID + ", ID = " + DocumentCommentFileID + ", Name = " + fDT.Rows[0]["FileName"], 321);
                        MessageBox.Show("Ошибка скачивания файла. Отчет отправлен разработчикам. Приносим извинения за неудобства.");
                    }
                }

            }

            return bOK;
        }


        public void RefillCorrespondents()
        {
            CorrespondentsDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT * FROM Correspondents ORDER BY CorrespondentName", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(CorrespondentsDataTable);
                }
            }
        }


        public void RemoveUploadedFiles()
        {
            foreach (DataRow Row in CurrentUploadedFiles.Rows)
            {
                FM.DeleteFile(Row["FileName"].ToString(), Configs.FTPType);
            }
        }


        public void RemoveInnerDocument(int InnerDocumentID)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsFiles WHERE DocumentCategoryID = 0 AND DocumentID = " + InnerDocumentID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        foreach (DataRow Row in DT.Rows)
                        {
                            FM.DeleteFile(Configs.DocumentsPath + "Документооборот/Документы/" + Row["FileName"], Configs.FTPType);
                            Row.Delete();
                        }

                        DA.Update(DT);
                    }
                }
            }

            using (var DA = new SqlDataAdapter("DELETE FROM DocumentsRecipients WHERE DocumentCategoryID = 0 AND DocumentID = " + InnerDocumentID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }

            using (var DA = new SqlDataAdapter("DELETE FROM SubscribesRecords WHERE SubscribesItemID = 18 AND TableItemID = " + InnerDocumentID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }

            using (var DA = new SqlDataAdapter("DELETE FROM SubscribesRecords WHERE SubscribesItemID = 21 AND TableItemID IN (SELECT DocumentCommentID FROM DocumentsComments WHERE DocumentCategoryID = 0 AND DocumentID = " + InnerDocumentID + ")", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }

            using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsCommentsFiles WHERE DocumentCommentID IN (SELECT DocumentCommentID FROM DocumentsComments WHERE DocumentCategoryID = 0 AND DocumentID = " + InnerDocumentID + ")", ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        foreach (DataRow Row in DT.Rows)
                        {
                            FM.DeleteFile(Configs.DocumentsPath + "Документооборот/Комментарии/" + Row["FileName"], Configs.FTPType);
                            Row.Delete();
                        }

                        DA.Update(DT);
                    }
                }
            }

            using (var DA = new SqlDataAdapter("DELETE FROM DocumentsComments WHERE DocumentCategoryID = 0 AND DocumentID = " + InnerDocumentID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }

            using (var DA = new SqlDataAdapter("DELETE FROM InnerDocuments WHERE InnerDocumentID = " + InnerDocumentID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }
        }

        public void RemoveIncomeDocument(int IncomeDocumentID)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsFiles WHERE DocumentCategoryID = 1 AND DocumentID = " + IncomeDocumentID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        foreach (DataRow Row in DT.Rows)
                        {
                            FM.DeleteFile(Configs.DocumentsPath + "Документооборот/Документы/" + Row["FileName"], Configs.FTPType);
                            Row.Delete();
                        }

                        DA.Update(DT);
                    }
                }
            }

            using (var DA = new SqlDataAdapter("DELETE FROM DocumentsRecipients WHERE DocumentCategoryID = 1 AND DocumentID = " + IncomeDocumentID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }



            using (var DA = new SqlDataAdapter("DELETE FROM SubscribesRecords WHERE SubscribesItemID = 19 AND TableItemID = " + IncomeDocumentID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }

            using (var DA = new SqlDataAdapter("DELETE FROM SubscribesRecords WHERE SubscribesItemID = 21 AND TableItemID IN (SELECT DocumentCommentID FROM DocumentsComments WHERE DocumentCategoryID = 1 AND DocumentID = " + IncomeDocumentID + ")", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }

            using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsCommentsFiles WHERE DocumentCommentID IN (SELECT DocumentCommentID FROM DocumentsComments WHERE DocumentCategoryID = 1 AND DocumentID = " + IncomeDocumentID + ")", ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        foreach (DataRow Row in DT.Rows)
                        {
                            FM.DeleteFile(Configs.DocumentsPath + "Документооборот/Комментарии/" + Row["FileName"], Configs.FTPType);
                            Row.Delete();
                        }

                        DA.Update(DT);
                    }
                }
            }

            using (var DA = new SqlDataAdapter("DELETE FROM DocumentsComments WHERE DocumentCategoryID = 1 AND DocumentID = " + IncomeDocumentID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }

            using (var DA = new SqlDataAdapter("DELETE FROM IncomeDocuments WHERE IncomeDocumentID = " + IncomeDocumentID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }
        }

        public void RemoveOuterDocument(int OuterDocumentID)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsFiles WHERE DocumentCategoryID = 2 AND DocumentID = " + OuterDocumentID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        foreach (DataRow Row in DT.Rows)
                        {
                            FM.DeleteFile(Configs.DocumentsPath + "Документооборот/Документы/" + Row["FileName"], Configs.FTPType);
                            Row.Delete();
                        }

                        DA.Update(DT);
                    }
                }
            }

            using (var DA = new SqlDataAdapter("DELETE FROM DocumentsRecipients WHERE DocumentCategoryID = 2 AND DocumentID = " + OuterDocumentID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }

            using (var DA = new SqlDataAdapter("DELETE FROM SubscribesRecords WHERE SubscribesItemID = 20 AND TableItemID = " + OuterDocumentID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }

            using (var DA = new SqlDataAdapter("DELETE FROM SubscribesRecords WHERE SubscribesItemID = 21 AND TableItemID IN (SELECT DocumentCommentID FROM DocumentsComments WHERE DocumentCategoryID = 2 AND DocumentID = " + OuterDocumentID + ")", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }

            using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsCommentsFiles WHERE DocumentCommentID IN (SELECT DocumentCommentID FROM DocumentsComments WHERE DocumentCategoryID = 2 AND DocumentID = " + OuterDocumentID + ")", ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        foreach (DataRow Row in DT.Rows)
                        {
                            FM.DeleteFile(Configs.DocumentsPath + "Документооборот/Комментарии/" + Row["FileName"], Configs.FTPType);
                            Row.Delete();
                        }

                        DA.Update(DT);
                    }
                }
            }

            using (var DA = new SqlDataAdapter("DELETE FROM DocumentsComments WHERE DocumentCategoryID = 2 AND DocumentID = " + OuterDocumentID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }

            using (var DA = new SqlDataAdapter("DELETE FROM OuterDocuments WHERE OuterDocumentID = " + OuterDocumentID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }
        }

        public int GetUpdatesCount(int UserID)
        {
            using (var DA = new SqlDataAdapter("SELECT Count(SubscribesRecordID) FROM SubscribesRecords WHERE (SubscribesItemID = 18 OR SubscribesItemID = 19 OR SubscribesItemID = 20 OR SubscribesItemID = 21)  AND UserID = " + UserID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    if (DA.Fill(DT) > 0)
                    {
                        return Convert.ToInt32(DT.Rows[0][0]);
                    }

                    return 0;
                }
            }
        }

        public void AddConfirm(DataTable RecipientsDT, int UserID, int DocumentID, int DocumentCategoryID)
        {
            var DateTime = Security.GetCurrentDate();

            using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM DocumentsConfirmations", ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        var NewRow = DT.NewRow();
                        NewRow["DocumentCategoryID"] = DocumentCategoryID;
                        NewRow["DocumentID"] = DocumentID;
                        NewRow["UserID"] = UserID;
                        NewRow["DateTime"] = DateTime;
                        DT.Rows.Add(NewRow);

                        DA.Update(DT);
                    }
                }
            }

            var DocumentConfirmationID = -1;

            using (var DA = new SqlDataAdapter("SELECT DocumentConfirmationID FROM DocumentsConfirmations WHERE DateTime = '" + DateTime.ToString("yyyy-MM-dd HH:mm:ss.fff") + "'", ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    DocumentConfirmationID = Convert.ToInt32(DT.Rows[0][0]);
                }
            }

            using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsRecipients WHERE DocumentID = " + DocumentID + " AND DocumentCategoryID = " + DocumentCategoryID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        using (var sDA = new SqlDataAdapter("SELECT TOP 0 * FROM SubscribesRecords", ConnectionStrings.LightConnectionString))
                        {
                            using (var sCB = new SqlCommandBuilder(sDA))
                            {
                                using (var sDT = new DataTable())
                                {
                                    sDA.Fill(sDT);

                                    foreach (DataRow Row in RecipientsDT.Rows)
                                    {
                                        if (DT.Select("UserID = " + Row["UserID"]).Count() == 0)
                                        {
                                            var NewRow = DT.NewRow();
                                            NewRow["DocumentID"] = DocumentID;
                                            NewRow["DocumentCategoryID"] = DocumentCategoryID;
                                            NewRow["UserID"] = Row["UserID"];
                                            DT.Rows.Add(NewRow);

                                            var sNewRow = sDT.NewRow();
                                            sNewRow["TableItemID"] = DocumentID;

                                            if (DocumentCategoryID == 0)
                                                sNewRow["SubscribesItemID"] = 18;
                                            if (DocumentCategoryID == 1)
                                                sNewRow["SubscribesItemID"] = 19;
                                            if (DocumentCategoryID == 2)
                                                sNewRow["SubscribesItemID"] = 20;

                                            sNewRow["UserID"] = Row["UserID"];
                                            sDT.Rows.Add(sNewRow);
                                        }
                                    }

                                    DA.Update(DT);
                                    sDA.Update(sDT);

                                }
                            }
                        }
                    }
                }
            }


            using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM DocumentsConfirmationsRecipients", ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        foreach (DataRow Row in RecipientsDT.Rows)
                        {
                            var NewRow = DT.NewRow();
                            NewRow["DocumentConfirmationID"] = DocumentConfirmationID;
                            NewRow["UserID"] = Row["UserID"];
                            DT.Rows.Add(NewRow);
                        }

                        DA.Update(DT);
                    }
                }
            }
        }

        public void SetConfirmStatus(int Status, int DocumentConfirmationRecipientID)
        {
            using (var DA = new SqlDataAdapter("SELECT DocumentConfirmationRecipientID, DateTime, Status FROM DocumentsConfirmationsRecipients WHERE DocumentConfirmationRecipientID = " + DocumentConfirmationRecipientID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        DT.Rows[0]["Status"] = Status;

                        if (Status == 0)
                            DT.Rows[0]["DateTime"] = DBNull.Value;
                        else
                            DT.Rows[0]["DateTime"] = Security.GetCurrentDate();

                        DA.Update(DT);
                    }
                }
            }

            using (var DA = new SqlDataAdapter("SELECT DocumentConfirmationID, Status FROM DocumentsConfirmations WHERE DocumentConfirmationID IN (SELECT TOP 1 DocumentConfirmationID FROM DocumentsConfirmationsRecipients WHERE DocumentConfirmationRecipientID = " + DocumentConfirmationRecipientID + ")", ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        using (var rDA = new SqlDataAdapter("SELECT DocumentConfirmationRecipientID, Status FROM DocumentsConfirmationsRecipients WHERE DocumentConfirmationID = " + DT.Rows[0]["DocumentConfirmationID"], ConnectionStrings.LightConnectionString))
                        {
                            using (var rDT = new DataTable())
                            {
                                rDA.Fill(rDT);

                                var rStatus = 0;

                                if (rDT.Select("Status = 1").Count() == rDT.Rows.Count)
                                    rStatus = 1;

                                if (rDT.Select("Status = 2").Count() > 0)
                                    rStatus = 2;

                                DT.Rows[0]["Status"] = rStatus;

                                DA.Update(DT);
                            }
                        }
                    }
                }
            }
        }

        public DataTable GetConfirmationRecipients(int DocumentConfirmationID)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsConfirmationsRecipients WHERE DocumentConfirmationID = " + DocumentConfirmationID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    return DT;
                }
            }
        }

        public void EditConfirm(int DocumentConfirmationID, DataTable RecipientsDT)
        {
            var DocumentID = -1;
            var DocumentCategoryID = -1;

            using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsConfirmations WHERE DocumentConfirmationID = " + DocumentConfirmationID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    DocumentID = Convert.ToInt32(DT.Rows[0]["DocumentID"]);
                    DocumentCategoryID = Convert.ToInt32(DT.Rows[0]["DocumentCategoryID"]);
                }
            }

            using (var DA = new SqlDataAdapter("SELECT * FROM DocumentsConfirmationsRecipients WHERE DocumentConfirmationID = " + DocumentConfirmationID, ConnectionStrings.LightConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        using (var rDA = new SqlDataAdapter("SELECT * FROM DocumentsRecipients WHERE DocumentID = " + DocumentID + " AND DocumentCategoryID = " + DocumentCategoryID, ConnectionStrings.LightConnectionString))
                        {
                            using (var rCB = new SqlCommandBuilder(rDA))
                            {
                                using (var rDT = new DataTable())
                                {
                                    rDA.Fill(rDT);

                                    using (var sDA = new SqlDataAdapter("SELECT TOP 0 * FROM SubscribesRecords", ConnectionStrings.LightConnectionString))
                                    {
                                        using (var sCB = new SqlCommandBuilder(sDA))
                                        {
                                            using (var sDT = new DataTable())
                                            {
                                                sDA.Fill(sDT);

                                                foreach (DataRow Row in RecipientsDT.Rows)
                                                {
                                                    if (DT.Select("UserID = " + Row["UserID"]).Count() == 0)
                                                    {
                                                        var NewRow = DT.NewRow();
                                                        NewRow["UserID"] = Row["UserID"];
                                                        NewRow["DocumentConfirmationID"] = DocumentConfirmationID;
                                                        DT.Rows.Add(NewRow);


                                                        if (rDT.Select("UserID = " + Row["UserID"]).Count() == 0)
                                                        {
                                                            var rNewRow = rDT.NewRow();
                                                            rNewRow["DocumentID"] = DocumentID;
                                                            rNewRow["DocumentCategoryID"] = DocumentCategoryID;
                                                            rNewRow["UserID"] = Row["UserID"];
                                                            rDT.Rows.Add(rNewRow);

                                                            var sNewRow = sDT.NewRow();
                                                            sNewRow["TableItemID"] = DocumentID;

                                                            if (DocumentCategoryID == 0)
                                                                sNewRow["SubscribesItemID"] = 18;
                                                            if (DocumentCategoryID == 1)
                                                                sNewRow["SubscribesItemID"] = 19;
                                                            if (DocumentCategoryID == 2)
                                                                sNewRow["SubscribesItemID"] = 20;

                                                            sNewRow["UserID"] = Row["UserID"];
                                                            sDT.Rows.Add(sNewRow);
                                                        }
                                                    }
                                                }

                                                rDA.Update(rDT);
                                                sDA.Update(sDT);

                                            }
                                        }
                                    }
                                }
                            }
                        }

                        foreach (DataRow Row in DT.Rows)
                        {
                            if (RecipientsDT.Select("UserID = " + Row["UserID"]).Count() == 0)
                            {
                                Row.Delete();
                            }
                        }

                        DA.Update(DT);
                    }
                }
            }
        }

        public void DeleteConfirm(int DocumentConfirmationID)
        {
            using (var DA = new SqlDataAdapter("DELETE FROM DocumentsConfirmationsRecipients WHERE DocumentConfirmationID = " + DocumentConfirmationID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }

            using (var DA = new SqlDataAdapter("DELETE FROM DocumentsConfirmations WHERE DocumentConfirmationID = " + DocumentConfirmationID, ConnectionStrings.LightConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                }
            }
        }
    }





    public class Contractors
    {
        public DataTable ContractorsDataTable;
        public DataTable ContactsDataTable;
        public DataTable CategoriesDataTable;
        public DataTable SubCategoriesDataTable;
        public DataTable CitiesDataTable;
        public DataTable CountriesDataTable;
        private SqlDataAdapter ContactsDA;
        private SqlCommandBuilder ContactsCB;
        public DataTable CurrentContactsDataTable;

        public Contractors()
        {
            CategoriesDataTable = new DataTable();
            SubCategoriesDataTable = new DataTable();
            ContractorsDataTable = new DataTable();
            CountriesDataTable = new DataTable();
            CitiesDataTable = new DataTable();
            ContactsDataTable = new DataTable();
            CurrentContactsDataTable = new DataTable();

            Fill();
        }

        public void Fill()
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM ContractorsCategories ORDER BY Name ASC", ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(CategoriesDataTable);
            }

            using (var DA = new SqlDataAdapter("SELECT * FROM ContractorsSubCategories ORDER BY Name ASC", ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(SubCategoriesDataTable);
            }

            using (var DA = new SqlDataAdapter("SELECT * FROM Countries ORDER BY Name ASC", ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(CountriesDataTable);
            }

            using (var DA = new SqlDataAdapter("SELECT * FROM Cities ORDER BY Name ASC", ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(CitiesDataTable);
            }

            using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM ContractorsContacts", ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(ContactsDataTable);
            }
        }

        public void RefillCategories()
        {
            CategoriesDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT * FROM ContractorsCategories ORDER BY Name ASC", ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(CategoriesDataTable);
            }
        }

        public void RefillSubCategories()
        {
            SubCategoriesDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT * FROM ContractorsSubCategories ORDER BY Name ASC", ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(SubCategoriesDataTable);
            }
        }

        public void RefillCountries()
        {
            CountriesDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT * FROM Countries ORDER BY Name ASC", ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(CountriesDataTable);
            }
        }

        public void RefillCities()
        {
            CitiesDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT * FROM Cities ORDER BY Name ASC", ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(CitiesDataTable);
            }
        }

        public void FillContractors(int ContractorSubCategoryID)
        {
            ContractorsDataTable.Clear();
            ContactsDataTable.Clear();

            using (var DA = new SqlDataAdapter("SELECT * FROM Contractors WHERE ContractorSubCategoryID = " + ContractorSubCategoryID + " ORDER BY Name ASC", ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(ContractorsDataTable);
            }

            using (var DA = new SqlDataAdapter("SELECT * FROM ContractorsContacts WHERE ContractorID IN (SELECT ContractorID FROM Contractors WHERE ContractorSubCategoryID = " + ContractorSubCategoryID + ")", ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(ContactsDataTable);
            }
        }

        public void AddCategory(string sName)
        {
            using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM ContractorsCategories", ConnectionStrings.CatalogConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        var NewRow = DT.NewRow();
                        NewRow["Name"] = sName;
                        DT.Rows.Add(NewRow);

                        DA.Update(DT);
                    }
                }
            }
        }

        public void AddSubCategory(string sName, int ContractorCategoryID)
        {
            using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM ContractorsSubCategories", ConnectionStrings.CatalogConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        var NewRow = DT.NewRow();
                        NewRow["Name"] = sName;
                        NewRow["ContractorCategoryID"] = ContractorCategoryID;
                        DT.Rows.Add(NewRow);

                        DA.Update(DT);
                    }
                }
            }
        }

        public void AddCountry(string sName)
        {
            using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM Countries", ConnectionStrings.CatalogConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        var NewRow = DT.NewRow();
                        NewRow["Name"] = sName;
                        DT.Rows.Add(NewRow);

                        DA.Update(DT);
                    }
                }
            }
        }

        public void AddCity(string sName, int iCountryID)
        {
            using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM Cities", ConnectionStrings.CatalogConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        var NewRow = DT.NewRow();
                        NewRow["Name"] = sName;
                        NewRow["CountryID"] = iCountryID;
                        DT.Rows.Add(NewRow);

                        DA.Update(DT);
                    }
                }
            }
        }

        public bool AddContractor(string sName, string Email, string Website, string Address, string Facebook, string Skype,
                                  int CountryID, int CityID, string UNN, string Description, int CategoryID, int SubCategoryID, DataTable ContactsDataTable)
        {
            using (var DA = new SqlDataAdapter("SELECT TOP 1 Name FROM Contractors WHERE Name = '" + sName + "'", ConnectionStrings.CatalogConnectionString))
            {
                using (var DT = new DataTable())
                {
                    if (DA.Fill(DT) > 0)
                        return false;
                }
            }

            using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM Contractors", ConnectionStrings.CatalogConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        var NewRow = DT.NewRow();
                        NewRow["Name"] = sName;
                        NewRow["Email"] = Email;
                        NewRow["Website"] = Website;
                        NewRow["Address"] = Address;
                        NewRow["Facebook"] = Facebook;
                        NewRow["Skype"] = Skype;
                        NewRow["CountryID"] = CountryID;
                        NewRow["CityID"] = CityID;
                        NewRow["UNN"] = UNN;
                        NewRow["Description"] = Description;
                        NewRow["ContractorCategoryID"] = CategoryID;
                        NewRow["ContractorSubCategoryID"] = SubCategoryID;
                        DT.Rows.Add(NewRow);

                        DA.Update(DT);
                    }
                }
            }

            if (ContactsDataTable.Rows.Count == 0)
                return true;

            var ContractorID = -1;

            using (var DA = new SqlDataAdapter("SELECT TOP 1 ContractorID FROM Contractors WHERE Name = '" + sName + "'", ConnectionStrings.CatalogConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    ContractorID = Convert.ToInt32(DT.Rows[0][0]);
                }
            }



            using (var DA = new SqlDataAdapter("SELECT TOP 0 * FROM ContractorsContacts", ConnectionStrings.CatalogConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        foreach (DataRow Row in ContactsDataTable.Rows)
                        {
                            var NewRow = DT.NewRow();
                            NewRow["ContractorID"] = ContractorID;
                            NewRow["Name"] = Row["Name"];
                            NewRow["Position"] = Row["Position"];
                            NewRow["Phone1"] = Row["Phone1"];
                            NewRow["Phone2"] = Row["Phone2"];
                            NewRow["Phone3"] = Row["Phone3"];
                            NewRow["Email"] = Row["Email"];
                            NewRow["Website"] = Row["Website"];
                            NewRow["Skype"] = Row["Website"];
                            NewRow["Description"] = Row["Name"];
                            DT.Rows.Add(NewRow);
                        }

                        DA.Update(DT);

                        return true;
                    }
                }
            }
        }

        public void GetEditContractor(int ContractorID, ref string sName, ref string Email, ref string Website, ref string Address, ref string Facebook, ref string Skype,
                                      ref int CountryID, ref int CityID, ref string UNN, ref string Description, ref int CategoryID, ref int SubCategoryID)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM Contractors WHERE ContractorID = " + ContractorID, ConnectionStrings.CatalogConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);

                    sName = DT.Rows[0]["Name"].ToString();
                    Email = DT.Rows[0]["Email"].ToString();
                    Website = DT.Rows[0]["Website"].ToString();
                    Address = DT.Rows[0]["Address"].ToString();
                    Facebook = DT.Rows[0]["Facebook"].ToString();
                    Skype = DT.Rows[0]["Skype"].ToString();
                    CountryID = Convert.ToInt32(DT.Rows[0]["CountryID"]);
                    CityID = Convert.ToInt32(DT.Rows[0]["CityID"]);
                    UNN = DT.Rows[0]["UNN"].ToString();
                    Description = DT.Rows[0]["Description"].ToString();
                    CategoryID = Convert.ToInt32(DT.Rows[0]["ContractorCategoryID"]);
                    SubCategoryID = Convert.ToInt32(DT.Rows[0]["ContractorSubCategoryID"]);
                }
            }

            if (ContactsDA != null)
            {
                ContactsDA.Dispose();
                ContactsCB.Dispose();
                CurrentContactsDataTable.Clear();
            }

            ContactsDA = new SqlDataAdapter("SELECT * FROM ContractorsContacts WHERE ContractorID = " + ContractorID, ConnectionStrings.CatalogConnectionString);
            ContactsCB = new SqlCommandBuilder(ContactsDA);
            ContactsDA.Fill(CurrentContactsDataTable);

        }

        public void EditContractor(int ContractorID, string sName, string Email, string Website, string Address, string Facebook, string Skype,
                                  int CountryID, int CityID, string UNN, string Description, int CategoryID, int SubCategoryID, DataTable ContactsDataTable)
        {
            using (var DA = new SqlDataAdapter("SELECT * FROM Contractors WHERE ContractorID = " + ContractorID, ConnectionStrings.CatalogConnectionString))
            {
                using (var CB = new SqlCommandBuilder(DA))
                {
                    using (var DT = new DataTable())
                    {
                        DA.Fill(DT);

                        DT.Rows[0]["Name"] = sName;
                        DT.Rows[0]["Email"] = Email;
                        DT.Rows[0]["Website"] = Website;
                        DT.Rows[0]["Address"] = Address;
                        DT.Rows[0]["Facebook"] = Facebook;
                        DT.Rows[0]["Skype"] = Skype;
                        DT.Rows[0]["CountryID"] = CountryID;
                        DT.Rows[0]["CityID"] = CityID;
                        DT.Rows[0]["UNN"] = UNN;
                        DT.Rows[0]["Description"] = Description;
                        DT.Rows[0]["ContractorCategoryID"] = CategoryID;
                        DT.Rows[0]["ContractorSubCategoryID"] = SubCategoryID;

                        DA.Update(DT);
                    }
                }
            }


            ContactsDA.Update(CurrentContactsDataTable);
        }

    }
}
