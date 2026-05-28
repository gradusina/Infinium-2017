using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Windows.Forms;

using Testing;

namespace Infinium
{
    public class InfiniumProjects
    {
        public DataTable ProjectsDataTable;
        public DataTable DepartmentsDataTable;
        public DataTable ProjectUsersMembersDataTable;
        public DataTable ProjectStatusesDataTable;
        public DataTable UsersDataTable;
        public DataTable UsersPhotoDataTable;
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
            UsersPhotoDataTable = new DataTable();
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
            UsersPhotoDataTable = TablesManager.UsersPhotoDataTable.Copy();
            using (var DA = new SqlDataAdapter("SELECT * FROM ProjectStatuses", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(ProjectStatusesDataTable);
            }

            UsersDataTable = new DataTable();
            using (var DA = new SqlDataAdapter("SELECT * FROM Users where fired=0 ORDER BY Name", ConnectionStrings.UsersConnectionString))
            {
                UsersDataTable.Clear();
                DA.Fill(UsersDataTable);
            }

            //UsersPhoto();
            using (var DA = new SqlDataAdapter("SELECT DepartmentID, DepartmentName FROM Departments", ConnectionStrings.LightConnectionString))
            {
                DA.Fill(DepartmentsDataTable);
            }
        }

        public void UsersPhoto()
        {
            var fm1 = new FileManager();
            var path = Configs.DocumentsPath + FileManager.GetPath("UsersPhoto");
            
            for (int i = 0; i < UsersDataTable.Rows.Count; i++)
            {
                var rows = UsersPhotoDataTable.Select("UserID = " + Security.CurrentUserID);
                if (!rows.Any())
                    continue;

                if (rows[0]["Photo"] == DBNull.Value)
                {
                    try
                    {
                        UsersDataTable.Rows[i]["Photo"] = fm1.ReadFile(path + "/" + "default.jpg",
                            Convert.ToInt64(rows[0]["FileSize"]), Configs.FTPType);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Ошибка здесь " + ex.Message);
                        return;
                    }
                }
                else
                {
                    try
                    {
                        UsersDataTable.Rows[i]["Photo"] = fm1.ReadFile(path + "/" + rows[0]["Photo"],
                            Convert.ToInt64(rows[0]["FileSize"]), Configs.FTPType);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Ошибка здесь " + ex.Message);
                        return;
                    }
                }
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

        //public Image GetUserPhoto(int UserID)
        //{
        //    var Rows = TablesManager.UsersPhotoDataTable.Select("UserID = " + UserID);
        //    if (!Rows.Any())
        //        return null;
        //    if (Rows[0]["Photo"] == DBNull.Value)
        //    {
        //        if (FM.FileExist(Configs.DocumentsPath + FileManager.GetPath("UsersPhoto") + "/" + "default.jpg",
        //                Configs.FTPType))
        //        {
        //            try
        //            {
        //                using (var ms = new MemoryStream(
        //                           FM.ReadFile(
        //                               Configs.DocumentsPath + FileManager.GetPath("UsersPhoto") + "/" + "default.jpg",
        //                               Convert.ToInt64(Rows[0]["FileSize"]), Configs.FTPType)))
        //                {
        //                    return Image.FromStream(ms);
        //                }
        //            }
        //            catch (Exception ex)
        //            {
        //                MessageBox.Show("Или здесь " + ex.Message);
        //                return null;
        //            }
        //        }
        //        else
        //        {
        //            return null;
        //        }
        //    }

        //    if (FM.FileExist(Configs.DocumentsPath + FileManager.GetPath("UsersPhoto") + "/" + Rows[0]["Photo"], Configs.FTPType))
        //    {
        //        try
        //        {
        //            using (var ms = new MemoryStream(
        //                FM.ReadFile(Configs.DocumentsPath + FileManager.GetPath("UsersPhoto") + "/" + Rows[0]["Photo"],
        //                Convert.ToInt64(Rows[0]["FileSize"]), Configs.FTPType)))
        //            {
        //                return Image.FromStream(ms);
        //            }
        //        }
        //        catch (Exception ex)
        //        {
        //            MessageBox.Show("Или здесь " + ex.Message);
        //            return null;
        //        }
        //    }
        //    else
        //    {
        //        return null;
        //    }
        //}

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
}
