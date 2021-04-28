using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;

using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;

namespace Infinium.Modules.StatisticsMarketing.Reports
{
    public class WeightDispatchReport
    {
        private DataTable MarketingFrontsWeightDT = null;
        private DataTable MarketingDecorWeightDT = null;
        private DataTable ZOVFrontsWeightDT = null;
        private DataTable ZOVDecorWeightDT = null;

        private DataTable AllMarketingWeightDT = null;
        private DataTable AllZOVWeightDT = null;

        public WeightDispatchReport()
        {
            MarketingFrontsWeightDT = new DataTable();
            MarketingDecorWeightDT = new DataTable();
            ZOVFrontsWeightDT = new DataTable();
            ZOVDecorWeightDT = new DataTable();

            AllMarketingWeightDT = new DataTable();
            AllMarketingWeightDT.Columns.Add(new DataColumn("ClientName", Type.GetType("System.String")));
            AllMarketingWeightDT.Columns.Add(new DataColumn("FrontsWeight", Type.GetType("System.Decimal")));
            AllMarketingWeightDT.Columns.Add(new DataColumn("DecorWeight", Type.GetType("System.Decimal")));
            AllMarketingWeightDT.Columns.Add(new DataColumn("TotalWeight", Type.GetType("System.Decimal")));

            AllZOVWeightDT = new DataTable();
            AllZOVWeightDT.Columns.Add(new DataColumn("date", Type.GetType("System.String")));
            AllZOVWeightDT.Columns.Add(new DataColumn("ClientName", Type.GetType("System.String")));
            AllZOVWeightDT.Columns.Add(new DataColumn("FrontsWeight", Type.GetType("System.Decimal")));
            AllZOVWeightDT.Columns.Add(new DataColumn("DecorWeight", Type.GetType("System.Decimal")));
            AllZOVWeightDT.Columns.Add(new DataColumn("TotalWeight", Type.GetType("System.Decimal")));
        }

        public void MarketingWeightDispatch(DateTime date1, DateTime date2)
        {
            string Filter = " AND CAST(Packages.DispatchDateTime AS date) >= '" + date1.ToString("yyyy-MM-dd") +
                " 00:00' AND CAST(Packages.DispatchDateTime AS date) <= '" + date2.ToString("yyyy-MM-dd") + " 23:59'";

            string SelectCommand = @"SELECT infiniu2_marketingreference.dbo.Clients.ClientName, SUM(PackageDetails.Count * FrontsOrders.Weight / FrontsOrders.Count) AS Count
                FROM FrontsOrders INNER JOIN
                PackageDetails ON FrontsOrders.FrontsOrdersID = PackageDetails.OrderID INNER JOIN
                Packages ON PackageDetails.PackageID = Packages.PackageID AND Packages.ProductType = 0 " + Filter + @" INNER JOIN
                MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID INNER JOIN
                MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID INNER JOIN
                infiniu2_marketingreference.dbo.Clients ON MegaOrders.ClientID = infiniu2_marketingreference.dbo.Clients.ClientID
                GROUP BY infiniu2_marketingreference.dbo.Clients.ClientName
                ORDER BY infiniu2_marketingreference.dbo.Clients.ClientName";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(MarketingFrontsWeightDT);
            }
            SelectCommand = @"SELECT infiniu2_marketingreference.dbo.Clients.ClientName, SUM(PackageDetails.Count * DecorOrders.Weight / DecorOrders.Count) AS Count
                FROM DecorOrders INNER JOIN
                PackageDetails ON DecorOrders.DecorOrderID = PackageDetails.OrderID INNER JOIN
                Packages ON PackageDetails.PackageID = Packages.PackageID AND Packages.ProductType = 1 " + Filter + @" INNER JOIN
                MainOrders ON DecorOrders.MainOrderID = MainOrders.MainOrderID INNER JOIN
                MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID INNER JOIN
                infiniu2_marketingreference.dbo.Clients ON MegaOrders.ClientID = infiniu2_marketingreference.dbo.Clients.ClientID
                GROUP BY infiniu2_marketingreference.dbo.Clients.ClientName
                ORDER BY infiniu2_marketingreference.dbo.Clients.ClientName";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(MarketingDecorWeightDT);
            }
        }

        public void ZOVWeightDispatch(DateTime date1, DateTime date2)
        {
            string Filter = " AND CAST(Packages.DispatchDateTime AS date) >= '" + date1.ToString("yyyy-MM-dd") +
                " 00:00' AND CAST(Packages.DispatchDateTime AS date) <= '" + date2.ToString("yyyy-MM-dd") + " 23:59'";

            string SelectCommand = @"SELECT CONVERT(nvarchar, Packages.DispatchDateTime, 104) AS date,
                infiniu2_zovreference.dbo.Clients.ClientName, SUM(PackageDetails.Count * FrontsOrders.Weight / FrontsOrders.Count) AS Count
                FROM FrontsOrders INNER JOIN
                PackageDetails ON FrontsOrders.FrontsOrdersID = PackageDetails.OrderID INNER JOIN
                Packages ON PackageDetails.PackageID = Packages.PackageID AND Packages.ProductType = 0 " + Filter + @" INNER JOIN
                MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID INNER JOIN
                JoinMainOrders ON MainOrders.MainOrderID = JoinMainOrders.MarketMainOrderID INNER JOIN
                infiniu2_zovreference.dbo.Clients ON JoinMainOrders.ZOVClientID = infiniu2_zovreference.dbo.Clients.ClientID
                GROUP BY CONVERT(nvarchar, Packages.DispatchDateTime, 104), infiniu2_zovreference.dbo.Clients.ClientName";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(ZOVFrontsWeightDT);
            }
            SelectCommand = @"SELECT CONVERT(nvarchar, Packages.DispatchDateTime, 104) AS date,
                infiniu2_zovreference.dbo.Clients.ClientName, SUM(PackageDetails.Count * DecorOrders.Weight / DecorOrders.Count) AS Count
                FROM DecorOrders INNER JOIN
                PackageDetails ON DecorOrders.DecorOrderID = PackageDetails.OrderID INNER JOIN
                Packages ON PackageDetails.PackageID = Packages.PackageID AND Packages.ProductType = 1 " + Filter + @" INNER JOIN
                MainOrders ON DecorOrders.MainOrderID = MainOrders.MainOrderID INNER JOIN
                JoinMainOrders ON MainOrders.MainOrderID = JoinMainOrders.MarketMainOrderID INNER JOIN
                infiniu2_zovreference.dbo.Clients ON JoinMainOrders.ZOVClientID = infiniu2_zovreference.dbo.Clients.ClientID
                GROUP BY CONVERT(nvarchar, Packages.DispatchDateTime, 104), infiniu2_zovreference.dbo.Clients.ClientName";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(ZOVDecorWeightDT);
            }
        }

        private void CollectMarketingWeight()
        {
            for (int i = 0; i < MarketingFrontsWeightDT.Rows.Count; i++)
            {
                string rr = MarketingFrontsWeightDT.Rows[i]["ClientName"].ToString();
                DataRow[] ItemsRows = AllMarketingWeightDT.Select("ClientName = '" + rr + "'");
                if (ItemsRows.Count() > 0)
                {
                    foreach (DataRow item in ItemsRows)
                    {
                        item["FrontsWeight"] = MarketingFrontsWeightDT.Rows[i]["Count"];
                    }
                }
                else
                {
                    DataRow NewRow = AllMarketingWeightDT.NewRow();
                    NewRow["ClientName"] = MarketingFrontsWeightDT.Rows[i]["ClientName"];
                    NewRow["FrontsWeight"] = MarketingFrontsWeightDT.Rows[i]["Count"];
                    AllMarketingWeightDT.Rows.Add(NewRow);
                }
            }
            for (int i = 0; i < MarketingDecorWeightDT.Rows.Count; i++)
            {
                string rr = MarketingDecorWeightDT.Rows[i]["ClientName"].ToString();
                DataRow[] ItemsRows = AllMarketingWeightDT.Select("ClientName = '" + rr + "'");
                if (ItemsRows.Count() > 0)
                {
                    foreach (DataRow item in ItemsRows)
                    {
                        item["DecorWeight"] = MarketingDecorWeightDT.Rows[i]["Count"];
                    }
                }
                else
                {
                    DataRow NewRow = AllMarketingWeightDT.NewRow();
                    NewRow["ClientName"] = MarketingDecorWeightDT.Rows[i]["ClientName"];
                    NewRow["DecorWeight"] = MarketingDecorWeightDT.Rows[i]["Count"];
                    AllMarketingWeightDT.Rows.Add(NewRow);
                }
            }

            using (DataView DV = new DataView(AllMarketingWeightDT.Copy()))
            {
                DV.Sort = "ClientName";
                AllMarketingWeightDT.Clear();
                AllMarketingWeightDT = DV.ToTable();
            }
            foreach (DataRow item in AllMarketingWeightDT.Rows)
            {
                if (item["FrontsWeight"] == DBNull.Value)
                    item["FrontsWeight"] = 0;
                if (item["DecorWeight"] == DBNull.Value)
                    item["DecorWeight"] = 0;
                decimal TotalWeight = Decimal.Round((Convert.ToDecimal(item["FrontsWeight"]) + Convert.ToDecimal(item["DecorWeight"])) / 1000, 2, MidpointRounding.AwayFromZero);
                item["TotalWeight"] = TotalWeight;
            }
        }

        private void CollectZOVWeight()
        {
            for (int i = 0; i < ZOVFrontsWeightDT.Rows.Count; i++)
            {
                string date = ZOVFrontsWeightDT.Rows[i]["date"].ToString();
                string ClientName = ZOVFrontsWeightDT.Rows[i]["ClientName"].ToString();
                DataRow[] ItemsRows = AllZOVWeightDT.Select("date = '" + date + "'" + " AND ClientName = '" + ClientName + "'");
                if (ItemsRows.Count() > 0)
                {
                    foreach (DataRow item in ItemsRows)
                    {
                        item["FrontsWeight"] = ZOVFrontsWeightDT.Rows[i]["Count"];
                    }
                }
                else
                {
                    DataRow NewRow = AllZOVWeightDT.NewRow();
                    NewRow["date"] = ZOVFrontsWeightDT.Rows[i]["date"];
                    NewRow["ClientName"] = ZOVFrontsWeightDT.Rows[i]["ClientName"];
                    NewRow["FrontsWeight"] = ZOVFrontsWeightDT.Rows[i]["Count"];
                    AllZOVWeightDT.Rows.Add(NewRow);
                }
            }
            for (int i = 0; i < ZOVDecorWeightDT.Rows.Count; i++)
            {
                string date = ZOVDecorWeightDT.Rows[i]["date"].ToString();
                string ClientName = ZOVDecorWeightDT.Rows[i]["ClientName"].ToString();
                DataRow[] ItemsRows = AllZOVWeightDT.Select("date = '" + date + "'" + " AND ClientName = '" + ClientName + "'");
                if (ItemsRows.Count() > 0)
                {
                    foreach (DataRow item in ItemsRows)
                    {
                        item["DecorWeight"] = ZOVDecorWeightDT.Rows[i]["Count"];
                    }
                }
                else
                {
                    DataRow NewRow = AllZOVWeightDT.NewRow();
                    NewRow["date"] = ZOVDecorWeightDT.Rows[i]["date"];
                    NewRow["ClientName"] = ZOVDecorWeightDT.Rows[i]["ClientName"];
                    NewRow["DecorWeight"] = ZOVDecorWeightDT.Rows[i]["Count"];
                    AllZOVWeightDT.Rows.Add(NewRow);
                }
            }

            using (DataView DV = new DataView(AllZOVWeightDT.Copy()))
            {
                DV.Sort = "date, ClientName";
                AllZOVWeightDT.Clear();
                AllZOVWeightDT = DV.ToTable();
            }
            foreach (DataRow item in AllZOVWeightDT.Rows)
            {
                if (item["FrontsWeight"] == DBNull.Value)
                    item["FrontsWeight"] = 0;
                if (item["DecorWeight"] == DBNull.Value)
                    item["DecorWeight"] = 0;

                decimal TotalWeight = Decimal.Round((Convert.ToDecimal(item["FrontsWeight"]) + Convert.ToDecimal(item["DecorWeight"])) / 1000, 2, MidpointRounding.AwayFromZero);
                item["TotalWeight"] = TotalWeight;
            }
        }

        public void ClearReport()
        {
            MarketingFrontsWeightDT.Clear();
            MarketingDecorWeightDT.Clear();
            ZOVFrontsWeightDT.Clear();
            ZOVDecorWeightDT.Clear();
            AllMarketingWeightDT.Clear();
            AllZOVWeightDT.Clear();
        }

        public void WeightReport(DateTime date1, DateTime date2, string FileName)
        {
            ClearReport();

            MarketingWeightDispatch(date1, date2);
            ZOVWeightDispatch(date1, date2);
            CollectMarketingWeight();
            CollectZOVWeight();

            string MainOrdersList = string.Empty;

            int pos = 0;

            //Export to excel
            HSSFWorkbook hssfworkbook = new HSSFWorkbook();

            ////create a entry of DocumentSummaryInformation
            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = "NPOI Team";
            hssfworkbook.DocumentSummaryInformation = dsi;

            ////create a entry of SummaryInformation
            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
            si.Subject = "NPOI SDK Example";
            hssfworkbook.SummaryInformation = si;

            #region Create fonts and styles

            HSSFFont HeaderF1 = hssfworkbook.CreateFont();
            HeaderF1.FontHeightInPoints = 11;
            HeaderF1.Boldweight = 11 * 256;
            HeaderF1.FontName = "Calibri";

            HSSFFont HeaderF2 = hssfworkbook.CreateFont();
            HeaderF2.FontHeightInPoints = 10;
            HeaderF2.Boldweight = 10 * 256;
            HeaderF2.FontName = "Calibri";

            HSSFFont HeaderF3 = hssfworkbook.CreateFont();
            HeaderF3.FontHeightInPoints = 9;
            HeaderF3.Boldweight = 9 * 256;
            HeaderF3.FontName = "Calibri";

            HSSFFont SimpleF = hssfworkbook.CreateFont();
            SimpleF.FontHeightInPoints = 10;
            SimpleF.FontName = "Calibri";

            HSSFCellStyle SimpleCS = hssfworkbook.CreateCellStyle();
            SimpleCS.BorderBottom = HSSFCellStyle.BORDER_THIN;
            SimpleCS.BottomBorderColor = HSSFColor.BLACK.index;
            SimpleCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            SimpleCS.LeftBorderColor = HSSFColor.BLACK.index;
            SimpleCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            SimpleCS.RightBorderColor = HSSFColor.BLACK.index;
            SimpleCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            SimpleCS.TopBorderColor = HSSFColor.BLACK.index;
            SimpleCS.SetFont(SimpleF);

            HSSFCellStyle CountCS = hssfworkbook.CreateCellStyle();
            CountCS.DataFormat = hssfworkbook.CreateDataFormat().GetFormat("### ### ##0.00");
            CountCS.BorderBottom = HSSFCellStyle.BORDER_THIN;
            CountCS.BottomBorderColor = HSSFColor.BLACK.index;
            CountCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            CountCS.LeftBorderColor = HSSFColor.BLACK.index;
            CountCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            CountCS.RightBorderColor = HSSFColor.BLACK.index;
            CountCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            CountCS.TopBorderColor = HSSFColor.BLACK.index;
            CountCS.SetFont(SimpleF);

            HSSFCellStyle WeightCS = hssfworkbook.CreateCellStyle();
            WeightCS.DataFormat = hssfworkbook.CreateDataFormat().GetFormat("### ### ##0.00");
            WeightCS.BorderBottom = HSSFCellStyle.BORDER_THIN;
            WeightCS.BottomBorderColor = HSSFColor.BLACK.index;
            WeightCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            WeightCS.LeftBorderColor = HSSFColor.BLACK.index;
            WeightCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            WeightCS.RightBorderColor = HSSFColor.BLACK.index;
            WeightCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            WeightCS.TopBorderColor = HSSFColor.BLACK.index;
            WeightCS.SetFont(SimpleF);

            HSSFCellStyle PriceBelCS = hssfworkbook.CreateCellStyle();
            PriceBelCS.DataFormat = hssfworkbook.CreateDataFormat().GetFormat("### ### ##0");
            PriceBelCS.BorderBottom = HSSFCellStyle.BORDER_THIN;
            PriceBelCS.BottomBorderColor = HSSFColor.BLACK.index;
            PriceBelCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            PriceBelCS.LeftBorderColor = HSSFColor.BLACK.index;
            PriceBelCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            PriceBelCS.RightBorderColor = HSSFColor.BLACK.index;
            PriceBelCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            PriceBelCS.TopBorderColor = HSSFColor.BLACK.index;
            PriceBelCS.SetFont(SimpleF);

            HSSFCellStyle PriceForeignCS = hssfworkbook.CreateCellStyle();
            PriceForeignCS.DataFormat = hssfworkbook.CreateDataFormat().GetFormat("### ### ##0.00");
            PriceForeignCS.BorderBottom = HSSFCellStyle.BORDER_THIN;
            PriceForeignCS.BottomBorderColor = HSSFColor.BLACK.index;
            PriceForeignCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            PriceForeignCS.LeftBorderColor = HSSFColor.BLACK.index;
            PriceForeignCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            PriceForeignCS.RightBorderColor = HSSFColor.BLACK.index;
            PriceForeignCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            PriceForeignCS.TopBorderColor = HSSFColor.BLACK.index;
            PriceForeignCS.SetFont(SimpleF);

            HSSFCellStyle ReportCS1 = hssfworkbook.CreateCellStyle();
            ReportCS1.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            ReportCS1.BottomBorderColor = HSSFColor.BLACK.index;
            ReportCS1.SetFont(HeaderF1);

            HSSFCellStyle ReportCS2 = hssfworkbook.CreateCellStyle();
            ReportCS2.SetFont(HeaderF1);

            HSSFCellStyle SummaryWithoutBorderBelCS = hssfworkbook.CreateCellStyle();
            SummaryWithoutBorderBelCS.DataFormat = hssfworkbook.CreateDataFormat().GetFormat("### ### ##0");
            SummaryWithoutBorderBelCS.SetFont(HeaderF2);

            HSSFCellStyle SummaryWithoutBorderForeignCS = hssfworkbook.CreateCellStyle();
            SummaryWithoutBorderForeignCS.DataFormat = hssfworkbook.CreateDataFormat().GetFormat("### ### ##0.00");
            SummaryWithoutBorderForeignCS.SetFont(HeaderF2);

            HSSFCellStyle SummaryWeightCS = hssfworkbook.CreateCellStyle();
            SummaryWeightCS.DataFormat = hssfworkbook.CreateDataFormat().GetFormat("### ### ##0.00");
            SummaryWeightCS.SetFont(HeaderF2);

            HSSFCellStyle SummaryWithBorderBelCS = hssfworkbook.CreateCellStyle();
            SummaryWithBorderBelCS.DataFormat = hssfworkbook.CreateDataFormat().GetFormat("### ### ##0");
            SummaryWithBorderBelCS.BorderBottom = HSSFCellStyle.BORDER_THIN;
            SummaryWithBorderBelCS.BottomBorderColor = HSSFColor.BLACK.index;
            SummaryWithBorderBelCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            SummaryWithBorderBelCS.LeftBorderColor = HSSFColor.BLACK.index;
            SummaryWithBorderBelCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            SummaryWithBorderBelCS.RightBorderColor = HSSFColor.BLACK.index;
            SummaryWithBorderBelCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            SummaryWithBorderBelCS.TopBorderColor = HSSFColor.BLACK.index;
            SummaryWithBorderBelCS.WrapText = true;
            SummaryWithBorderBelCS.SetFont(HeaderF2);

            HSSFCellStyle SummaryWithBorderForeignCS = hssfworkbook.CreateCellStyle();
            SummaryWithBorderForeignCS.DataFormat = hssfworkbook.CreateDataFormat().GetFormat("### ### ##0.00");
            SummaryWithBorderForeignCS.BorderBottom = HSSFCellStyle.BORDER_THIN;
            SummaryWithBorderForeignCS.BottomBorderColor = HSSFColor.BLACK.index;
            SummaryWithBorderForeignCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            SummaryWithBorderForeignCS.LeftBorderColor = HSSFColor.BLACK.index;
            SummaryWithBorderForeignCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            SummaryWithBorderForeignCS.RightBorderColor = HSSFColor.BLACK.index;
            SummaryWithBorderForeignCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            SummaryWithBorderForeignCS.TopBorderColor = HSSFColor.BLACK.index;
            SummaryWithBorderForeignCS.WrapText = true;
            SummaryWithBorderForeignCS.SetFont(HeaderF2);

            HSSFCellStyle SimpleHeaderCS = hssfworkbook.CreateCellStyle();
            SimpleHeaderCS.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            SimpleHeaderCS.BottomBorderColor = HSSFColor.BLACK.index;
            SimpleHeaderCS.BorderLeft = HSSFCellStyle.BORDER_MEDIUM;
            SimpleHeaderCS.LeftBorderColor = HSSFColor.BLACK.index;
            SimpleHeaderCS.BorderRight = HSSFCellStyle.BORDER_MEDIUM;
            SimpleHeaderCS.RightBorderColor = HSSFColor.BLACK.index;
            SimpleHeaderCS.BorderTop = HSSFCellStyle.BORDER_MEDIUM;
            SimpleHeaderCS.TopBorderColor = HSSFColor.BLACK.index;
            //SimpleHeaderCS.WrapText = true;
            SimpleHeaderCS.SetFont(HeaderF3);

            #endregion Create fonts and styles

            HSSFSheet sheet1 = hssfworkbook.CreateSheet("Маркетинг");
            sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

            sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
            sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
            sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
            sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

            sheet1.SetColumnWidth(0, 71 * 256);
            sheet1.SetColumnWidth(1, 13 * 256);
            sheet1.SetColumnWidth(2, 13 * 256);
            sheet1.SetColumnWidth(3, 14 * 256);

            HSSFSheet sheet2 = hssfworkbook.CreateSheet("ЗОВ");
            sheet2.PrintSetup.PaperSize = (short)PaperSizeType.A4;

            sheet2.SetMargin(HSSFSheet.LeftMargin, (double).12);
            sheet2.SetMargin(HSSFSheet.RightMargin, (double).07);
            sheet2.SetMargin(HSSFSheet.TopMargin, (double).20);
            sheet2.SetMargin(HSSFSheet.BottomMargin, (double).20);

            sheet2.SetColumnWidth(0, 13 * 256);
            sheet2.SetColumnWidth(1, 50 * 256);
            sheet2.SetColumnWidth(2, 13 * 256);
            sheet2.SetColumnWidth(3, 14 * 256);
            sheet2.SetColumnWidth(4, 14 * 256);

            HSSFCell Cell1;

            if (AllMarketingWeightDT.Rows.Count > 0)
            {
                //Профиль
                pos += 2;

                Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                Cell1.SetCellValue("Клиент");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(1);
                Cell1.SetCellValue("Фасады, кг");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(2);
                Cell1.SetCellValue("Декор, кг");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(3);
                Cell1.SetCellValue("Вес общий, т");
                Cell1.CellStyle = SimpleHeaderCS;

                pos++;

                for (int i = 0; i < AllMarketingWeightDT.Rows.Count; i++)
                {
                    Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                    Cell1.SetCellValue(AllMarketingWeightDT.Rows[i]["ClientName"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(1);
                    Cell1.SetCellValue(Convert.ToDouble(AllMarketingWeightDT.Rows[i]["FrontsWeight"]));
                    Cell1.CellStyle = CountCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(2);
                    Cell1.SetCellValue(Convert.ToDouble(AllMarketingWeightDT.Rows[i]["DecorWeight"]));
                    Cell1.CellStyle = CountCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(3);
                    Cell1.SetCellValue(Convert.ToDouble(AllMarketingWeightDT.Rows[i]["TotalWeight"]));
                    Cell1.CellStyle = CountCS;
                    pos++;
                }
            }
            pos = 0;

            if (AllZOVWeightDT.Rows.Count > 0)
            {
                //Профиль
                pos += 2;

                Cell1 = sheet2.CreateRow(pos).CreateCell(0);
                Cell1.SetCellValue("Дата");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet2.CreateRow(pos).CreateCell(1);
                Cell1.SetCellValue("Клиент");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet2.CreateRow(pos).CreateCell(2);
                Cell1.SetCellValue("Фасады, кг");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet2.CreateRow(pos).CreateCell(3);
                Cell1.SetCellValue("Декор, кг");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet2.CreateRow(pos).CreateCell(4);
                Cell1.SetCellValue("Вес общий, т");
                Cell1.CellStyle = SimpleHeaderCS;

                pos++;

                for (int i = 0; i < AllZOVWeightDT.Rows.Count; i++)
                {
                    Cell1 = sheet2.CreateRow(pos).CreateCell(0);
                    Cell1.SetCellValue(AllZOVWeightDT.Rows[i]["date"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet2.CreateRow(pos).CreateCell(1);
                    Cell1.SetCellValue(AllZOVWeightDT.Rows[i]["ClientName"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet2.CreateRow(pos).CreateCell(2);
                    Cell1.SetCellValue(Convert.ToDouble(AllZOVWeightDT.Rows[i]["FrontsWeight"]));
                    Cell1.CellStyle = CountCS;
                    Cell1 = sheet2.CreateRow(pos).CreateCell(3);
                    Cell1.SetCellValue(Convert.ToDouble(AllZOVWeightDT.Rows[i]["DecorWeight"]));
                    Cell1.CellStyle = CountCS;
                    Cell1 = sheet2.CreateRow(pos).CreateCell(4);
                    Cell1.SetCellValue(Convert.ToDouble(AllZOVWeightDT.Rows[i]["TotalWeight"]));
                    Cell1.CellStyle = CountCS;
                    pos++;
                }
            }

            pos++;

            string tempFolder = System.Environment.GetEnvironmentVariable("TEMP");
            FileInfo file = new FileInfo(tempFolder + @"\" + FileName + ".xls");
            int j = 1;
            while (file.Exists == true)
            {
                file = new FileInfo(tempFolder + @"\" + FileName + "(" + j++ + ").xls");
            }

            FileStream NewFile = new FileStream(file.FullName, FileMode.Create);
            hssfworkbook.Write(NewFile);
            NewFile.Close();

            System.Diagnostics.Process.Start(file.FullName);
            ClearReport();
        }
    }
}