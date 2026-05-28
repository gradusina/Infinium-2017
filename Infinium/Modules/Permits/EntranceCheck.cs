
using Infinium.Modules.Marketing.NewOrders;

using Microsoft.VisualBasic.ApplicationServices;

using System;
using System.Collections;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Printing;
using System.Windows.Forms;
using Infinium.Modules.Permits;

namespace Infinium.Modules.Entrance
{
    public class EntranceCheck
    {
        public int CurrentProductType;
        public int CurrentMainOrderId;
        public int CurrentMegaOrderId;

        private DataTable _entranceDt;

        public BindingSource EntranceBs;

        public struct EntranceInfo
        {
            public int UserId;
            public string UserName;
            public DateTime EntryDate;
            public DateTime QuitDate;
        }

        public EntranceInfo Entry { get; set; }

        public int UserId { get; set; }

        public EntranceCheck()
        {
            Initialize();
        }

        private void Initialize()
        {
            Create();
            Fill();
            Binding();
        }

        private void Create()
        {
            _entranceDt = new DataTable();
        }

        private void Fill()
        {
            var str = @$"SELECT TOP 0 UsersEntrance.*, users.ShortName from UsersEntrance
INNER JOIN infiniu2_users.dbo.Users as users ON UsersEntrance.userId = users.userId";

            using var da = new SqlDataAdapter(str, ConnectionStrings.LightConnectionString);
            da.Fill(_entranceDt);
        }

        private void Binding()
        {
            EntranceBs = new BindingSource
            {
                DataSource = _entranceDt
            };
        }

        public bool FilterEntrance(DateTime date)
        {
            _entranceDt.Clear();

            var str = @$"SELECT UsersEntrance.*, users.ShortName from UsersEntrance
INNER JOIN infiniu2_users.dbo.Users as users ON UsersEntrance.userId = users.userId
WHERE CAST(entryDate AS DATE) = '{date.ToString("yyyy-MM-dd")}'";

            using var da = new SqlDataAdapter(str, ConnectionStrings.LightConnectionString);

            da.Fill(_entranceDt);

            return _entranceDt.Rows.Count > 0;
        }

        public void UserEnter(int userId)
        {
            const string str = "SELECT TOP 0 * FROM UsersEntrance";

            using var da = new SqlDataAdapter(str, ConnectionStrings.LightConnectionString);
            using var cb = new SqlCommandBuilder(da);
            using var dt = new DataTable();
            da.Fill(dt);

            var newRow = dt.NewRow();
            newRow["userID"] = userId;
            newRow["entryDate"] = Security.GetCurrentDate();
            newRow["entry"] = 1;
            dt.Rows.Add(newRow);

            da.Update(dt);

            Entry = new EntranceInfo
            {
                UserId = userId,
                UserName = GetUserName(userId),
                EntryDate = Convert.ToDateTime(dt.Rows[0]["entryDate"])
            };
        }

        public void UserQuit(int userId)
        {
            var str = @$"SELECT UsersEntrance.* FROM UsersEntrance 
WHERE UsersEntrance.entry=1 and UsersEntrance.quit=0 and UsersEntrance.userId = {userId}";

            using var da = new SqlDataAdapter(str, ConnectionStrings.LightConnectionString);
            using var cb = new SqlCommandBuilder(da);
            using var dt = new DataTable();

            if (da.Fill(dt) == 0) return;

            dt.Rows[0]["quit"] = 1;
            dt.Rows[0]["quitDate"] = Security.GetCurrentDate();

            Entry = new EntranceInfo
            {
                UserId = userId,
                UserName = GetUserName(userId),
                EntryDate = Convert.ToDateTime(dt.Rows[0]["entryDate"]),
                QuitDate = Convert.ToDateTime(dt.Rows[0]["quitDate"])
            };

            da.Update(dt);
        }

        public bool IsUserEnter(int userId)
        {
            var str = @$"SELECT UsersEntrance.*, users.ShortName FROM UsersEntrance 
INNER JOIN infiniu2_users.dbo.Users as users ON UsersEntrance.userId = users.userId 
WHERE UsersEntrance.entry=1 and UsersEntrance.quit=0 and UsersEntrance.userId = {userId}";

            using var da = new SqlDataAdapter(str, ConnectionStrings.LightConnectionString);
            using var dt = new DataTable();

            if (da.Fill(dt) <= 0) return false;
            
            Entry = new EntranceInfo
            {
                UserId = userId,
                UserName = dt.Rows[0]["ShortName"].ToString(),
                EntryDate = Convert.ToDateTime(dt.Rows[0]["entryDate"])
            };

            return true;
        }

        public string GetUserName(int userId)
        {
            using var da = new SqlDataAdapter(@$"SELECT ShortName FROM Users WHERE userId = {userId}",
                ConnectionStrings.UsersConnectionString);
            using var dt = new DataTable();

            return da.Fill(dt) == 0 ? "" : dt.Rows[0]["ShortName"].ToString();
        }

        public void Clear()
        {
            _entranceDt.Clear();
        }
    }


    public struct UsersBarcodeInfo
    {
        public string BarcodeNumber;
    }


    public class PrintUsersBarcode
    {
        private Packages.Marketing.Barcode Barcode;
        public PrintDocument PD;

        public int PaperHeight = 488;
        public int PaperWidth = 794;

        public int CurrentLabelNumber = 0;

        public int PrintedCount = 0;

        public bool Printed = false;
        
        public ArrayList LabelInfo;

        public PrintUsersBarcode()
        {
            Barcode = new Modules.Packages.Marketing.Barcode();
            
            InitializePrinter();
            
            LabelInfo = new System.Collections.ArrayList();
        }

        private void InitializePrinter()
        {
            PD = new PrintDocument();
            PD.DefaultPageSettings.PaperSize = new PaperSize("Label", PaperWidth, PaperHeight);
            PD.DefaultPageSettings.Landscape = true;
            PD.DefaultPageSettings.Margins.Bottom = 0;
            PD.DefaultPageSettings.Margins.Top = 0;
            PD.DefaultPageSettings.Margins.Left = 0;
            PD.DefaultPageSettings.Margins.Right = 0;
        }
        
        public void ClearLabelInfo()
        {
            CurrentLabelNumber = 0;
            PrintedCount = 0;
            LabelInfo.Clear();
            GC.Collect();
        }

        public string GetBarcodeNumber(int barcodeType, int id)
        {
            var type = barcodeType.ToString().Length switch
            {
                1 => "00" + barcodeType,
                2 => "0" + barcodeType,
                3 => barcodeType.ToString(),
                _ => ""
            };

            var number = id.ToString().Length switch
            {
                1 => "00000000" + id,
                2 => "0000000" + id,
                3 => "000000" + id,
                4 => "00000" + id,
                5 => "0000" + id,
                6 => "000" + id,
                7 => "00" + id,
                8 => "0" + id,
                _ => ""
            };

            System.Text.StringBuilder barcodeNumber = new System.Text.StringBuilder(type);
            barcodeNumber.Append(number);

            return barcodeNumber.ToString();
        }

        public void AddLabelInfo(ref UsersBarcodeInfo labelInfoItem)
        {
            LabelInfo.Add(labelInfoItem);
        }

        public void PrintPage(object sender, PrintPageEventArgs ev)
        {
            if (PrintedCount >= LabelInfo.Count)
                return;
            else
                PrintedCount++;
            ev.Graphics.Clear(Color.White);
            ev.Graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;

            var Y = 0;
            
            ev.Graphics.DrawImage(Barcode.GetBarcode(Packages.Marketing.Barcode.BarcodeLength.Short, 35, ((UsersBarcodeInfo)LabelInfo[CurrentLabelNumber]).BarcodeNumber), 0, Y);
            Barcode.DrawBarcodeText(Packages.Marketing.Barcode.BarcodeLength.Short, ev.Graphics, ((UsersBarcodeInfo)LabelInfo[CurrentLabelNumber]).BarcodeNumber, 4, Y + 35);

            if (CurrentLabelNumber == LabelInfo.Count - 1 || PrintedCount >= LabelInfo.Count)
                ev.HasMorePages = false;

            if (CurrentLabelNumber < LabelInfo.Count - 1 && PrintedCount < LabelInfo.Count)
            {
                ev.HasMorePages = true;
                CurrentLabelNumber++;
            }
        }

        public void Print()
        {
            PD.DefaultPageSettings.PaperSize = new PaperSize("Label", PaperWidth + 40, PaperHeight);
            PD.DefaultPageSettings.Margins.Bottom = 0;
            PD.DefaultPageSettings.Margins.Top = 0;
            PD.DefaultPageSettings.Margins.Right = 0;
            PD.DefaultPageSettings.Margins.Left = 0;

            if (!Printed)
            {
                Printed = true;
                PD.PrintPage += new PrintPageEventHandler(PrintPage);
            }

            PD.Print();
        }
    }
}
