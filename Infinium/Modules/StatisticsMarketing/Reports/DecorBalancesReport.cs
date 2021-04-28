using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;

using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;

namespace Infinium.Modules.StatisticsMarketing.Reports
{
    public class DecorBalancesReport
    {
        private DataTable BalancesDT = null;
        private HSSFWorkbook hssfworkbook = new HSSFWorkbook();

        public DecorBalancesReport(string FileName)
        {
            BalancesDT = new DataTable();
            BalancesMarketingReport(false);
            BalancesZOVReport(true);
            ////create a entry of DocumentSummaryInformation
            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = "NPOI Team";
            hssfworkbook.DocumentSummaryInformation = dsi;

            ////create a entry of SummaryInformation
            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
            si.Subject = "NPOI SDK Example";
            hssfworkbook.SummaryInformation = si;

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
        }

        private void BalancesMarketingReport(bool bZOV)
        {
            int pos = 0;

            string SheetName = "Маркетинг";
            string ConnectionString = ConnectionStrings.MarketingOrdersConnectionString;

            #region Create fonts and styles

            HSSFFont HeaderF = hssfworkbook.CreateFont();
            HeaderF.FontHeightInPoints = 9;
            HeaderF.Boldweight = 9 * 256;
            HeaderF.FontName = "Calibri";

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

            HSSFCellStyle CountCS1 = hssfworkbook.CreateCellStyle();
            CountCS1.DataFormat = hssfworkbook.CreateDataFormat().GetFormat("### ### ##0.00");
            CountCS1.BorderBottom = HSSFCellStyle.BORDER_THIN;
            CountCS1.BottomBorderColor = HSSFColor.BLACK.index;
            CountCS1.BorderLeft = HSSFCellStyle.BORDER_THIN;
            CountCS1.LeftBorderColor = HSSFColor.BLACK.index;
            CountCS1.BorderRight = HSSFCellStyle.BORDER_THIN;
            CountCS1.RightBorderColor = HSSFColor.BLACK.index;
            CountCS1.BorderTop = HSSFCellStyle.BORDER_THIN;
            CountCS1.TopBorderColor = HSSFColor.BLACK.index;
            CountCS1.SetFont(SimpleF);

            HSSFCellStyle CountCS2 = hssfworkbook.CreateCellStyle();
            CountCS2.DataFormat = hssfworkbook.CreateDataFormat().GetFormat("### ### ##0.0000");
            CountCS2.BorderBottom = HSSFCellStyle.BORDER_THIN;
            CountCS2.BottomBorderColor = HSSFColor.BLACK.index;
            CountCS2.BorderLeft = HSSFCellStyle.BORDER_THIN;
            CountCS2.LeftBorderColor = HSSFColor.BLACK.index;
            CountCS2.BorderRight = HSSFCellStyle.BORDER_THIN;
            CountCS2.RightBorderColor = HSSFColor.BLACK.index;
            CountCS2.BorderTop = HSSFCellStyle.BORDER_THIN;
            CountCS2.TopBorderColor = HSSFColor.BLACK.index;
            CountCS2.SetFont(SimpleF);

            HSSFCellStyle CountCS3 = hssfworkbook.CreateCellStyle();
            CountCS3.DataFormat = hssfworkbook.CreateDataFormat().GetFormat("### ### ##0");
            CountCS3.BorderBottom = HSSFCellStyle.BORDER_THIN;
            CountCS3.BottomBorderColor = HSSFColor.BLACK.index;
            CountCS3.BorderLeft = HSSFCellStyle.BORDER_THIN;
            CountCS3.LeftBorderColor = HSSFColor.BLACK.index;
            CountCS3.BorderRight = HSSFCellStyle.BORDER_THIN;
            CountCS3.RightBorderColor = HSSFColor.BLACK.index;
            CountCS3.BorderTop = HSSFCellStyle.BORDER_THIN;
            CountCS3.TopBorderColor = HSSFColor.BLACK.index;
            CountCS3.SetFont(SimpleF);

            HSSFCellStyle SimpleHeaderCS = hssfworkbook.CreateCellStyle();
            SimpleHeaderCS.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            SimpleHeaderCS.BottomBorderColor = HSSFColor.BLACK.index;
            SimpleHeaderCS.BorderLeft = HSSFCellStyle.BORDER_MEDIUM;
            SimpleHeaderCS.LeftBorderColor = HSSFColor.BLACK.index;
            SimpleHeaderCS.BorderRight = HSSFCellStyle.BORDER_MEDIUM;
            SimpleHeaderCS.RightBorderColor = HSSFColor.BLACK.index;
            SimpleHeaderCS.BorderTop = HSSFCellStyle.BORDER_MEDIUM;
            SimpleHeaderCS.TopBorderColor = HSSFColor.BLACK.index;
            SimpleHeaderCS.SetFont(HeaderF);

            #endregion Create fonts and styles

            HSSFSheet sheet1 = hssfworkbook.CreateSheet(SheetName);
            sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

            sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
            sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
            sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
            sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

            sheet1.SetColumnWidth(0, 20 * 256);
            sheet1.SetColumnWidth(1, 30 * 256);
            sheet1.SetColumnWidth(2, 20 * 256);
            sheet1.SetColumnWidth(3, 20 * 256);

            HSSFCell Cell1;

            string SelectCommand = @"SELECT infiniu2_catalog.dbo.DecorProducts.ProductName AS ProductName, Decor.TechStoreName AS Decor, Colors.TechStoreName AS Color, Patina.PatinaName AS Patina,
                SUM(dbo.PackageDetails.Count) AS Expr2
                FROM dbo.PackageDetails INNER JOIN
                dbo.DecorOrders ON dbo.PackageDetails.OrderID = dbo.DecorOrders.DecorOrderID AND dbo.DecorOrders.FactoryID = 2 AND dbo.DecorOrders.Length = - 1 AND dbo.DecorOrders.Height = - 1 AND
                dbo.DecorOrders.Width = - 1 INNER JOIN
                dbo.Packages ON dbo.PackageDetails.PackageID = dbo.Packages.PackageID AND dbo.Packages.ProductType = 1 AND dbo.Packages.PackageStatusID IN (1, 2, 4) INNER JOIN
                infiniu2_catalog.dbo.DecorProducts ON dbo.DecorOrders.ProductID = infiniu2_catalog.dbo.DecorProducts.ProductID INNER JOIN
                infiniu2_catalog.dbo.TechStore AS Decor ON dbo.DecorOrders.DecorID = Decor.TechStoreID LEFT OUTER JOIN
                infiniu2_catalog.dbo.TechStore AS Colors ON dbo.DecorOrders.ColorID = Colors.TechStoreID LEFT OUTER JOIN
                infiniu2_catalog.dbo.Patina AS Patina ON dbo.DecorOrders.PatinaID = Patina.PatinaID
                GROUP BY infiniu2_catalog.dbo.DecorProducts.ProductName, Decor.TechStoreName, Colors.TechStoreName, Patina.PatinaName
                ORDER BY ProductName, Decor, Color, Patina";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionString))
            {
                BalancesDT.Dispose();
                BalancesDT = new DataTable();
                DA.Fill(BalancesDT);
            }
            if (BalancesDT.Rows.Count > 0)
            {
                Cell1 = sheet1.CreateRow(pos).CreateCell(4);
                Cell1.SetCellValue("Кол-во");
                Cell1.CellStyle = SimpleHeaderCS;
                pos++;
                for (int i = 0; i < BalancesDT.Rows.Count; i++)
                {
                    Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                    Cell1.SetCellValue(BalancesDT.Rows[i]["ProductName"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(1);
                    Cell1.SetCellValue(BalancesDT.Rows[i]["Decor"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(2);
                    Cell1.SetCellValue(BalancesDT.Rows[i]["Color"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(3);
                    Cell1.SetCellValue(BalancesDT.Rows[i]["Patina"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(4);
                    Cell1.SetCellValue(Convert.ToInt32(BalancesDT.Rows[i]["Expr2"]));
                    Cell1.CellStyle = CountCS3;
                    pos++;
                }
                pos += 2;
            }

            SelectCommand = @"SELECT infiniu2_catalog.dbo.DecorProducts.ProductName AS ProductName, Decor.TechStoreName AS Decor, Colors.TechStoreName AS Color, Patina.PatinaName AS Patina, dbo.DecorOrders.Length,
                SUM(dbo.PackageDetails.Count) AS Expr2 FROM dbo.PackageDetails INNER JOIN
                dbo.DecorOrders ON dbo.PackageDetails.OrderID = dbo.DecorOrders.DecorOrderID AND dbo.DecorOrders.FactoryID = 2 AND
                dbo.DecorOrders.Length <> - 1 AND dbo.DecorOrders.Height = - 1 AND dbo.DecorOrders.Width = - 1 INNER JOIN
                dbo.Packages ON dbo.PackageDetails.PackageID = dbo.Packages.PackageID AND dbo.Packages.ProductType = 1 AND dbo.Packages.PackageStatusID IN (1, 2, 4) INNER JOIN

                infiniu2_catalog.dbo.DecorProducts ON dbo.DecorOrders.ProductID = infiniu2_catalog.dbo.DecorProducts.ProductID INNER JOIN
                infiniu2_catalog.dbo.TechStore AS Decor ON dbo.DecorOrders.DecorID = Decor.TechStoreID LEFT OUTER JOIN
                infiniu2_catalog.dbo.TechStore AS Colors ON dbo.DecorOrders.ColorID = Colors.TechStoreID LEFT OUTER JOIN
                infiniu2_catalog.dbo.Patina AS Patina ON dbo.DecorOrders.PatinaID = Patina.PatinaID
                GROUP BY infiniu2_catalog.dbo.DecorProducts.ProductName, Decor.TechStoreName, Colors.TechStoreName, Patina.PatinaName, dbo.DecorOrders.Length
                ORDER BY ProductName, Decor, Color, Patina, dbo.DecorOrders.Length";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionString))
            {
                BalancesDT.Dispose();
                BalancesDT = new DataTable();
                DA.Fill(BalancesDT);
            }
            if (BalancesDT.Rows.Count > 0)
            {
                Cell1 = sheet1.CreateRow(pos).CreateCell(4);
                Cell1.SetCellValue("Длина");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(5);
                Cell1.SetCellValue("Кол-во");
                Cell1.CellStyle = SimpleHeaderCS;
                pos++;
                for (int i = 0; i < BalancesDT.Rows.Count; i++)
                {
                    Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                    Cell1.SetCellValue(BalancesDT.Rows[i]["ProductName"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(1);
                    Cell1.SetCellValue(BalancesDT.Rows[i]["Decor"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(2);
                    Cell1.SetCellValue(BalancesDT.Rows[i]["Color"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(3);
                    Cell1.SetCellValue(BalancesDT.Rows[i]["Patina"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(4);
                    Cell1.SetCellValue(Convert.ToInt32(BalancesDT.Rows[i]["Length"]));
                    Cell1.CellStyle = CountCS3;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(5);
                    Cell1.SetCellValue(Convert.ToInt32(BalancesDT.Rows[i]["Expr2"]));
                    Cell1.CellStyle = CountCS3;
                    pos++;
                }
                pos += 2;
            }

            SelectCommand = @"SELECT infiniu2_catalog.dbo.DecorProducts.ProductName AS ProductName, Decor.TechStoreName AS Decor, Colors.TechStoreName AS Color, Patina.PatinaName AS Patina,
                SUM(CAST(dbo.DecorOrders.Length * dbo.DecorOrders.Width * dbo.PackageDetails.Count AS decimal(10,2))) AS Expr1, SUM(dbo.PackageDetails.Count) AS Expr2 FROM dbo.PackageDetails INNER JOIN
                dbo.DecorOrders ON dbo.PackageDetails.OrderID = dbo.DecorOrders.DecorOrderID AND dbo.DecorOrders.FactoryID = 2 AND dbo.DecorOrders.Length <> - 1 AND dbo.DecorOrders.Height = - 1 AND
                dbo.DecorOrders.Width <> - 1 INNER JOIN
                dbo.Packages ON dbo.PackageDetails.PackageID = dbo.Packages.PackageID AND dbo.Packages.ProductType = 1 AND dbo.Packages.PackageStatusID IN (1, 2, 4) INNER JOIN
                infiniu2_catalog.dbo.DecorConfig ON dbo.DecorOrders.DecorConfigID = infiniu2_catalog.dbo.DecorConfig.DecorConfigID AND infiniu2_catalog.dbo.DecorConfig.MeasureID = 2 INNER JOIN
                infiniu2_catalog.dbo.DecorProducts ON dbo.DecorOrders.ProductID = infiniu2_catalog.dbo.DecorProducts.ProductID INNER JOIN
                infiniu2_catalog.dbo.TechStore AS Decor ON dbo.DecorOrders.DecorID = Decor.TechStoreID LEFT OUTER JOIN
                infiniu2_catalog.dbo.TechStore AS Colors ON dbo.DecorOrders.ColorID = Colors.TechStoreID LEFT OUTER JOIN
                infiniu2_catalog.dbo.Patina AS Patina ON dbo.DecorOrders.PatinaID = Patina.PatinaID
                GROUP BY infiniu2_catalog.dbo.DecorProducts.ProductName, Decor.TechStoreName, Colors.TechStoreName, Patina.PatinaName
                ORDER BY ProductName, Decor, Color, Patina";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionString))
            {
                BalancesDT.Dispose();
                BalancesDT = new DataTable();
                DA.Fill(BalancesDT);
                foreach (DataRow item in BalancesDT.Rows)
                    item["Expr1"] = Convert.ToDecimal(item["Expr1"]) / 1000;
            }
            if (BalancesDT.Rows.Count > 0)
            {
                Cell1 = sheet1.CreateRow(pos).CreateCell(4);
                Cell1.SetCellValue("м.п.");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(5);
                Cell1.SetCellValue("Кол-во");
                Cell1.CellStyle = SimpleHeaderCS;
                pos++;
                for (int i = 0; i < BalancesDT.Rows.Count; i++)
                {
                    Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                    Cell1.SetCellValue(BalancesDT.Rows[i]["ProductName"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(1);
                    Cell1.SetCellValue(BalancesDT.Rows[i]["Decor"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(2);
                    Cell1.SetCellValue(BalancesDT.Rows[i]["Color"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(3);
                    Cell1.SetCellValue(BalancesDT.Rows[i]["Patina"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(4);
                    Cell1.SetCellValue(Convert.ToDouble(BalancesDT.Rows[i]["Expr1"]));
                    Cell1.CellStyle = CountCS1;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(5);
                    Cell1.SetCellValue(Convert.ToInt32(BalancesDT.Rows[i]["Expr2"]));
                    Cell1.CellStyle = CountCS3;
                    pos++;
                }
                pos += 2;
            }

            SelectCommand = @"SELECT infiniu2_catalog.dbo.DecorProducts.ProductName AS ProductName, Decor.TechStoreName AS Decor, Colors.TechStoreName AS Color, Patina.PatinaName AS Patina,
                SUM(CAST(dbo.DecorOrders.Length * dbo.DecorOrders.Width * dbo.PackageDetails.Count AS decimal(12,4))) AS Expr1, SUM(dbo.PackageDetails.Count) AS Expr2 FROM dbo.PackageDetails INNER JOIN
                dbo.DecorOrders ON dbo.PackageDetails.OrderID = dbo.DecorOrders.DecorOrderID AND dbo.DecorOrders.FactoryID = 2 AND dbo.DecorOrders.Length <> - 1 AND dbo.DecorOrders.Height = - 1 AND
                dbo.DecorOrders.Width <> - 1 INNER JOIN
                dbo.Packages ON dbo.PackageDetails.PackageID = dbo.Packages.PackageID AND dbo.Packages.ProductType = 1 AND dbo.Packages.PackageStatusID IN (1, 2, 4) INNER JOIN
                infiniu2_catalog.dbo.DecorConfig ON dbo.DecorOrders.DecorConfigID = infiniu2_catalog.dbo.DecorConfig.DecorConfigID AND infiniu2_catalog.dbo.DecorConfig.MeasureID = 1 INNER JOIN
                infiniu2_catalog.dbo.DecorProducts ON dbo.DecorOrders.ProductID = infiniu2_catalog.dbo.DecorProducts.ProductID INNER JOIN
                infiniu2_catalog.dbo.TechStore AS Decor ON dbo.DecorOrders.DecorID = Decor.TechStoreID LEFT OUTER JOIN
                infiniu2_catalog.dbo.TechStore AS Colors ON dbo.DecorOrders.ColorID = Colors.TechStoreID LEFT OUTER JOIN
                infiniu2_catalog.dbo.Patina AS Patina ON dbo.DecorOrders.PatinaID = Patina.PatinaID
                GROUP BY infiniu2_catalog.dbo.DecorProducts.ProductName, Decor.TechStoreName, Colors.TechStoreName, Patina.PatinaName
                ORDER BY ProductName, Decor, Color, Patina";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionString))
            {
                BalancesDT.Dispose();
                BalancesDT = new DataTable();
                DA.Fill(BalancesDT);
                foreach (DataRow item in BalancesDT.Rows)
                    item["Expr1"] = Convert.ToDecimal(item["Expr1"]) / 1000000;
            }
            if (BalancesDT.Rows.Count > 0)
            {
                Cell1 = sheet1.CreateRow(pos).CreateCell(4);
                Cell1.SetCellValue("м.кв.");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(5);
                Cell1.SetCellValue("Кол-во");
                Cell1.CellStyle = SimpleHeaderCS;
                pos++;
                for (int i = 0; i < BalancesDT.Rows.Count; i++)
                {
                    Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                    Cell1.SetCellValue(BalancesDT.Rows[i]["ProductName"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(1);
                    Cell1.SetCellValue(BalancesDT.Rows[i]["Decor"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(2);
                    Cell1.SetCellValue(BalancesDT.Rows[i]["Color"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(3);
                    Cell1.SetCellValue(BalancesDT.Rows[i]["Patina"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(4);
                    Cell1.SetCellValue(Convert.ToDouble(BalancesDT.Rows[i]["Expr1"]));
                    Cell1.CellStyle = CountCS2;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(5);
                    Cell1.SetCellValue(Convert.ToInt32(BalancesDT.Rows[i]["Expr2"]));
                    Cell1.CellStyle = CountCS3;
                    pos++;
                }
            }

            SelectCommand = @"SELECT infiniu2_catalog.dbo.DecorProducts.ProductName AS ProductName, Decor.TechStoreName AS Decor, Colors.TechStoreName AS Color, Patina.PatinaName AS Patina,
                SUM(CAST(dbo.DecorOrders.Height * dbo.DecorOrders.Width * dbo.PackageDetails.Count AS decimal(12,4))) AS Expr1, SUM(dbo.PackageDetails.Count) AS Expr2 FROM dbo.PackageDetails INNER JOIN
                dbo.DecorOrders ON dbo.PackageDetails.OrderID = dbo.DecorOrders.DecorOrderID AND dbo.DecorOrders.FactoryID = 2 AND
                dbo.DecorOrders.Length = - 1 AND dbo.DecorOrders.Height <> - 1 AND dbo.DecorOrders.Width <> - 1 INNER JOIN
                dbo.Packages ON dbo.PackageDetails.PackageID = dbo.Packages.PackageID AND dbo.Packages.ProductType = 1 AND dbo.Packages.PackageStatusID IN (1, 2, 4) INNER JOIN
                infiniu2_catalog.dbo.DecorProducts ON dbo.DecorOrders.ProductID = infiniu2_catalog.dbo.DecorProducts.ProductID INNER JOIN
                infiniu2_catalog.dbo.TechStore AS Decor ON dbo.DecorOrders.DecorID = Decor.TechStoreID LEFT OUTER JOIN
                infiniu2_catalog.dbo.TechStore AS Colors ON dbo.DecorOrders.ColorID = Colors.TechStoreID LEFT OUTER JOIN
                infiniu2_catalog.dbo.Patina AS Patina ON dbo.DecorOrders.PatinaID = Patina.PatinaID
                GROUP BY infiniu2_catalog.dbo.DecorProducts.ProductName, Decor.TechStoreName, Colors.TechStoreName, Patina.PatinaName
                ORDER BY ProductName, Decor, Color, Patina";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionString))
            {
                BalancesDT.Dispose();
                BalancesDT = new DataTable();
                DA.Fill(BalancesDT);
                foreach (DataRow item in BalancesDT.Rows)
                    item["Expr1"] = Convert.ToDecimal(item["Expr1"]) / 1000000;
            }
            if (BalancesDT.Rows.Count > 0)
            {
                for (int i = 0; i < BalancesDT.Rows.Count; i++)
                {
                    Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                    Cell1.SetCellValue(BalancesDT.Rows[i]["ProductName"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(1);
                    Cell1.SetCellValue(BalancesDT.Rows[i]["Decor"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(2);
                    Cell1.SetCellValue(BalancesDT.Rows[i]["Color"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(3);
                    Cell1.SetCellValue(BalancesDT.Rows[i]["Patina"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(4);
                    Cell1.SetCellValue(Convert.ToDouble(BalancesDT.Rows[i]["Expr1"]));
                    Cell1.CellStyle = CountCS2;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(5);
                    Cell1.SetCellValue(Convert.ToInt32(BalancesDT.Rows[i]["Expr2"]));
                    Cell1.CellStyle = CountCS3;
                    pos++;
                }
                pos += 1;
            }
        }

        private void BalancesZOVReport(bool bZOV)
        {
            int pos = 0;

            string SheetName = "ЗОВ";

            #region Create fonts and styles

            HSSFFont HeaderF = hssfworkbook.CreateFont();
            HeaderF.FontHeightInPoints = 9;
            HeaderF.Boldweight = 9 * 256;
            HeaderF.FontName = "Calibri";

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

            HSSFCellStyle CountCS1 = hssfworkbook.CreateCellStyle();
            CountCS1.DataFormat = hssfworkbook.CreateDataFormat().GetFormat("### ### ##0.00");
            CountCS1.BorderBottom = HSSFCellStyle.BORDER_THIN;
            CountCS1.BottomBorderColor = HSSFColor.BLACK.index;
            CountCS1.BorderLeft = HSSFCellStyle.BORDER_THIN;
            CountCS1.LeftBorderColor = HSSFColor.BLACK.index;
            CountCS1.BorderRight = HSSFCellStyle.BORDER_THIN;
            CountCS1.RightBorderColor = HSSFColor.BLACK.index;
            CountCS1.BorderTop = HSSFCellStyle.BORDER_THIN;
            CountCS1.TopBorderColor = HSSFColor.BLACK.index;
            CountCS1.SetFont(SimpleF);

            HSSFCellStyle CountCS2 = hssfworkbook.CreateCellStyle();
            CountCS2.DataFormat = hssfworkbook.CreateDataFormat().GetFormat("### ### ##0.0000");
            CountCS2.BorderBottom = HSSFCellStyle.BORDER_THIN;
            CountCS2.BottomBorderColor = HSSFColor.BLACK.index;
            CountCS2.BorderLeft = HSSFCellStyle.BORDER_THIN;
            CountCS2.LeftBorderColor = HSSFColor.BLACK.index;
            CountCS2.BorderRight = HSSFCellStyle.BORDER_THIN;
            CountCS2.RightBorderColor = HSSFColor.BLACK.index;
            CountCS2.BorderTop = HSSFCellStyle.BORDER_THIN;
            CountCS2.TopBorderColor = HSSFColor.BLACK.index;
            CountCS2.SetFont(SimpleF);

            HSSFCellStyle CountCS3 = hssfworkbook.CreateCellStyle();
            CountCS3.DataFormat = hssfworkbook.CreateDataFormat().GetFormat("### ### ##0");
            CountCS3.BorderBottom = HSSFCellStyle.BORDER_THIN;
            CountCS3.BottomBorderColor = HSSFColor.BLACK.index;
            CountCS3.BorderLeft = HSSFCellStyle.BORDER_THIN;
            CountCS3.LeftBorderColor = HSSFColor.BLACK.index;
            CountCS3.BorderRight = HSSFCellStyle.BORDER_THIN;
            CountCS3.RightBorderColor = HSSFColor.BLACK.index;
            CountCS3.BorderTop = HSSFCellStyle.BORDER_THIN;
            CountCS3.TopBorderColor = HSSFColor.BLACK.index;
            CountCS3.SetFont(SimpleF);

            HSSFCellStyle SimpleHeaderCS = hssfworkbook.CreateCellStyle();
            SimpleHeaderCS.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            SimpleHeaderCS.BottomBorderColor = HSSFColor.BLACK.index;
            SimpleHeaderCS.BorderLeft = HSSFCellStyle.BORDER_MEDIUM;
            SimpleHeaderCS.LeftBorderColor = HSSFColor.BLACK.index;
            SimpleHeaderCS.BorderRight = HSSFCellStyle.BORDER_MEDIUM;
            SimpleHeaderCS.RightBorderColor = HSSFColor.BLACK.index;
            SimpleHeaderCS.BorderTop = HSSFCellStyle.BORDER_MEDIUM;
            SimpleHeaderCS.TopBorderColor = HSSFColor.BLACK.index;
            SimpleHeaderCS.SetFont(HeaderF);

            #endregion Create fonts and styles

            HSSFSheet sheet1 = hssfworkbook.CreateSheet(SheetName);
            sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

            sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
            sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
            sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
            sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

            sheet1.SetColumnWidth(0, 20 * 256);
            sheet1.SetColumnWidth(1, 30 * 256);
            sheet1.SetColumnWidth(2, 20 * 256);
            sheet1.SetColumnWidth(3, 20 * 256);

            HSSFCell Cell1;

            string SelectCommand = @"SELECT infiniu2_catalog.dbo.DecorProducts.ProductName AS ProductName, Decor.TechStoreName AS Decor, Colors.TechStoreName AS Color, Patina.PatinaName AS Patina,
                SUM(dbo.PackageDetails.Count) AS Expr2
                FROM dbo.PackageDetails INNER JOIN
                dbo.DecorOrders ON dbo.PackageDetails.OrderID = dbo.DecorOrders.DecorOrderID AND dbo.DecorOrders.FactoryID = 2 AND dbo.DecorOrders.Length = - 1 AND dbo.DecorOrders.Height = - 1 AND
                dbo.DecorOrders.Width = - 1 INNER JOIN
                dbo.Packages ON dbo.PackageDetails.PackageID = dbo.Packages.PackageID AND dbo.Packages.ProductType = 1 AND dbo.Packages.PackageStatusID IN (1, 2, 4) INNER JOIN
                JoinMainOrders ON DecorOrders.MainOrderID = JoinMainOrders.MarketMainOrderID INNER JOIN
                infiniu2_catalog.dbo.DecorProducts ON dbo.DecorOrders.ProductID = infiniu2_catalog.dbo.DecorProducts.ProductID INNER JOIN
                infiniu2_catalog.dbo.TechStore AS Decor ON dbo.DecorOrders.DecorID = Decor.TechStoreID LEFT OUTER JOIN
                infiniu2_catalog.dbo.TechStore AS Colors ON dbo.DecorOrders.ColorID = Colors.TechStoreID LEFT OUTER JOIN
                infiniu2_catalog.dbo.Patina AS Patina ON dbo.DecorOrders.PatinaID = Patina.PatinaID
                GROUP BY infiniu2_catalog.dbo.DecorProducts.ProductName, Decor.TechStoreName, Colors.TechStoreName, Patina.PatinaName
                ORDER BY ProductName, Decor, Color, Patina";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                BalancesDT.Dispose();
                BalancesDT = new DataTable();
                DA.Fill(BalancesDT);
            }
            if (BalancesDT.Rows.Count > 0)
            {
                Cell1 = sheet1.CreateRow(pos).CreateCell(4);
                Cell1.SetCellValue("Кол-во");
                Cell1.CellStyle = SimpleHeaderCS;
                pos++;
                for (int i = 0; i < BalancesDT.Rows.Count; i++)
                {
                    Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                    Cell1.SetCellValue(BalancesDT.Rows[i]["ProductName"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(1);
                    Cell1.SetCellValue(BalancesDT.Rows[i]["Decor"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(2);
                    Cell1.SetCellValue(BalancesDT.Rows[i]["Color"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(3);
                    Cell1.SetCellValue(BalancesDT.Rows[i]["Patina"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(4);
                    Cell1.SetCellValue(Convert.ToInt32(BalancesDT.Rows[i]["Expr2"]));
                    Cell1.CellStyle = CountCS3;
                    pos++;
                }
                pos += 2;
            }

            SelectCommand = @"SELECT infiniu2_catalog.dbo.DecorProducts.ProductName AS ProductName, Decor.TechStoreName AS Decor, Colors.TechStoreName AS Color, Patina.PatinaName AS Patina, dbo.DecorOrders.Length,
                SUM(dbo.PackageDetails.Count) AS Expr2 FROM dbo.PackageDetails INNER JOIN
                dbo.DecorOrders ON dbo.PackageDetails.OrderID = dbo.DecorOrders.DecorOrderID AND dbo.DecorOrders.FactoryID = 2 AND
                dbo.DecorOrders.Length <> - 1 AND dbo.DecorOrders.Height = - 1 AND dbo.DecorOrders.Width = - 1 INNER JOIN
                dbo.Packages ON dbo.PackageDetails.PackageID = dbo.Packages.PackageID AND dbo.Packages.ProductType = 1 AND dbo.Packages.PackageStatusID IN (1, 2, 4) INNER JOIN
                JoinMainOrders ON DecorOrders.MainOrderID = JoinMainOrders.MarketMainOrderID INNER JOIN
                infiniu2_catalog.dbo.DecorProducts ON dbo.DecorOrders.ProductID = infiniu2_catalog.dbo.DecorProducts.ProductID INNER JOIN
                infiniu2_catalog.dbo.TechStore AS Decor ON dbo.DecorOrders.DecorID = Decor.TechStoreID LEFT OUTER JOIN
                infiniu2_catalog.dbo.TechStore AS Colors ON dbo.DecorOrders.ColorID = Colors.TechStoreID LEFT OUTER JOIN
                infiniu2_catalog.dbo.Patina AS Patina ON dbo.DecorOrders.PatinaID = Patina.PatinaID
                GROUP BY infiniu2_catalog.dbo.DecorProducts.ProductName, Decor.TechStoreName, Colors.TechStoreName, Patina.PatinaName, dbo.DecorOrders.Length
                ORDER BY ProductName, Decor, Color, Patina, dbo.DecorOrders.Length";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                BalancesDT.Dispose();
                BalancesDT = new DataTable();
                DA.Fill(BalancesDT);
            }
            if (BalancesDT.Rows.Count > 0)
            {
                Cell1 = sheet1.CreateRow(pos).CreateCell(4);
                Cell1.SetCellValue("Длина");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(5);
                Cell1.SetCellValue("Кол-во");
                Cell1.CellStyle = SimpleHeaderCS;
                pos++;
                for (int i = 0; i < BalancesDT.Rows.Count; i++)
                {
                    Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                    Cell1.SetCellValue(BalancesDT.Rows[i]["ProductName"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(1);
                    Cell1.SetCellValue(BalancesDT.Rows[i]["Decor"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(2);
                    Cell1.SetCellValue(BalancesDT.Rows[i]["Color"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(3);
                    Cell1.SetCellValue(BalancesDT.Rows[i]["Patina"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(4);
                    Cell1.SetCellValue(Convert.ToInt32(BalancesDT.Rows[i]["Length"]));
                    Cell1.CellStyle = CountCS3;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(5);
                    Cell1.SetCellValue(Convert.ToInt32(BalancesDT.Rows[i]["Expr2"]));
                    Cell1.CellStyle = CountCS3;
                    pos++;
                }
                pos += 2;
            }

            SelectCommand = @"SELECT infiniu2_catalog.dbo.DecorProducts.ProductName AS ProductName, Decor.TechStoreName AS Decor, Colors.TechStoreName AS Color, Patina.PatinaName AS Patina,
                SUM(CAST(dbo.DecorOrders.Length * dbo.DecorOrders.Width * dbo.PackageDetails.Count AS decimal(10,2))) AS Expr1, SUM(dbo.PackageDetails.Count) AS Expr2 FROM dbo.PackageDetails INNER JOIN
                dbo.DecorOrders ON dbo.PackageDetails.OrderID = dbo.DecorOrders.DecorOrderID AND dbo.DecorOrders.FactoryID = 2 AND dbo.DecorOrders.Length <> - 1 AND dbo.DecorOrders.Height = - 1 AND
                dbo.DecorOrders.Width <> - 1 INNER JOIN
                dbo.Packages ON dbo.PackageDetails.PackageID = dbo.Packages.PackageID AND dbo.Packages.ProductType = 1 AND dbo.Packages.PackageStatusID IN (1, 2, 4) INNER JOIN
                JoinMainOrders ON DecorOrders.MainOrderID = JoinMainOrders.MarketMainOrderID INNER JOIN
                infiniu2_catalog.dbo.DecorConfig ON dbo.DecorOrders.DecorConfigID = infiniu2_catalog.dbo.DecorConfig.DecorConfigID AND infiniu2_catalog.dbo.DecorConfig.MeasureID = 2 INNER JOIN
                infiniu2_catalog.dbo.DecorProducts ON dbo.DecorOrders.ProductID = infiniu2_catalog.dbo.DecorProducts.ProductID INNER JOIN
                infiniu2_catalog.dbo.TechStore AS Decor ON dbo.DecorOrders.DecorID = Decor.TechStoreID LEFT OUTER JOIN
                infiniu2_catalog.dbo.TechStore AS Colors ON dbo.DecorOrders.ColorID = Colors.TechStoreID LEFT OUTER JOIN
                infiniu2_catalog.dbo.Patina AS Patina ON dbo.DecorOrders.PatinaID = Patina.PatinaID
                GROUP BY infiniu2_catalog.dbo.DecorProducts.ProductName, Decor.TechStoreName, Colors.TechStoreName, Patina.PatinaName
                ORDER BY ProductName, Decor, Color, Patina";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                BalancesDT.Dispose();
                BalancesDT = new DataTable();
                DA.Fill(BalancesDT);
                foreach (DataRow item in BalancesDT.Rows)
                    item["Expr1"] = Convert.ToDecimal(item["Expr1"]) / 1000;
            }
            if (BalancesDT.Rows.Count > 0)
            {
                Cell1 = sheet1.CreateRow(pos).CreateCell(4);
                Cell1.SetCellValue("м.п.");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(5);
                Cell1.SetCellValue("Кол-во");
                Cell1.CellStyle = SimpleHeaderCS;
                pos++;
                for (int i = 0; i < BalancesDT.Rows.Count; i++)
                {
                    Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                    Cell1.SetCellValue(BalancesDT.Rows[i]["ProductName"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(1);
                    Cell1.SetCellValue(BalancesDT.Rows[i]["Decor"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(2);
                    Cell1.SetCellValue(BalancesDT.Rows[i]["Color"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(3);
                    Cell1.SetCellValue(BalancesDT.Rows[i]["Patina"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(4);
                    Cell1.SetCellValue(Convert.ToDouble(BalancesDT.Rows[i]["Expr1"]));
                    Cell1.CellStyle = CountCS1;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(5);
                    Cell1.SetCellValue(Convert.ToInt32(BalancesDT.Rows[i]["Expr2"]));
                    Cell1.CellStyle = CountCS3;
                    pos++;
                }
                pos += 2;
            }

            SelectCommand = @"SELECT infiniu2_catalog.dbo.DecorProducts.ProductName AS ProductName, Decor.TechStoreName AS Decor, Colors.TechStoreName AS Color, Patina.PatinaName AS Patina,
                SUM(CAST(dbo.DecorOrders.Length * dbo.DecorOrders.Width * dbo.PackageDetails.Count AS decimal(12,4))) AS Expr1, SUM(dbo.PackageDetails.Count) AS Expr2 FROM dbo.PackageDetails INNER JOIN
                dbo.DecorOrders ON dbo.PackageDetails.OrderID = dbo.DecorOrders.DecorOrderID AND dbo.DecorOrders.FactoryID = 2 AND dbo.DecorOrders.Length <> - 1 AND dbo.DecorOrders.Height = - 1 AND
                dbo.DecorOrders.Width <> - 1 INNER JOIN
                dbo.Packages ON dbo.PackageDetails.PackageID = dbo.Packages.PackageID AND dbo.Packages.ProductType = 1 AND dbo.Packages.PackageStatusID IN (1, 2, 4) INNER JOIN
                JoinMainOrders ON DecorOrders.MainOrderID = JoinMainOrders.MarketMainOrderID INNER JOIN
                infiniu2_catalog.dbo.DecorConfig ON dbo.DecorOrders.DecorConfigID = infiniu2_catalog.dbo.DecorConfig.DecorConfigID AND infiniu2_catalog.dbo.DecorConfig.MeasureID = 1 INNER JOIN
                infiniu2_catalog.dbo.DecorProducts ON dbo.DecorOrders.ProductID = infiniu2_catalog.dbo.DecorProducts.ProductID INNER JOIN
                infiniu2_catalog.dbo.TechStore AS Decor ON dbo.DecorOrders.DecorID = Decor.TechStoreID LEFT OUTER JOIN
                infiniu2_catalog.dbo.TechStore AS Colors ON dbo.DecorOrders.ColorID = Colors.TechStoreID LEFT OUTER JOIN
                infiniu2_catalog.dbo.Patina AS Patina ON dbo.DecorOrders.PatinaID = Patina.PatinaID
                GROUP BY infiniu2_catalog.dbo.DecorProducts.ProductName, Decor.TechStoreName, Colors.TechStoreName, Patina.PatinaName
                ORDER BY ProductName, Decor, Color, Patina";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                BalancesDT.Dispose();
                BalancesDT = new DataTable();
                DA.Fill(BalancesDT);
                foreach (DataRow item in BalancesDT.Rows)
                    item["Expr1"] = Convert.ToDecimal(item["Expr1"]) / 1000000;
            }
            if (BalancesDT.Rows.Count > 0)
            {
                Cell1 = sheet1.CreateRow(pos).CreateCell(4);
                Cell1.SetCellValue("м.кв.");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(5);
                Cell1.SetCellValue("Кол-во");
                Cell1.CellStyle = SimpleHeaderCS;
                pos++;
                for (int i = 0; i < BalancesDT.Rows.Count; i++)
                {
                    Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                    Cell1.SetCellValue(BalancesDT.Rows[i]["ProductName"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(1);
                    Cell1.SetCellValue(BalancesDT.Rows[i]["Decor"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(2);
                    Cell1.SetCellValue(BalancesDT.Rows[i]["Color"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(3);
                    Cell1.SetCellValue(BalancesDT.Rows[i]["Patina"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(4);
                    Cell1.SetCellValue(Convert.ToDouble(BalancesDT.Rows[i]["Expr1"]));
                    Cell1.CellStyle = CountCS2;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(5);
                    Cell1.SetCellValue(Convert.ToInt32(BalancesDT.Rows[i]["Expr2"]));
                    Cell1.CellStyle = CountCS3;
                    pos++;
                }
            }

            SelectCommand = @"SELECT infiniu2_catalog.dbo.DecorProducts.ProductName AS ProductName, Decor.TechStoreName AS Decor, Colors.TechStoreName AS Color, Patina.PatinaName AS Patina,
                SUM(CAST(dbo.DecorOrders.Height * dbo.DecorOrders.Width * dbo.PackageDetails.Count AS decimal(12,4))) AS Expr1, SUM(dbo.PackageDetails.Count) AS Expr2 FROM dbo.PackageDetails INNER JOIN
                dbo.DecorOrders ON dbo.PackageDetails.OrderID = dbo.DecorOrders.DecorOrderID AND dbo.DecorOrders.FactoryID = 2 AND
                dbo.DecorOrders.Length = - 1 AND dbo.DecorOrders.Height <> - 1 AND dbo.DecorOrders.Width <> - 1 INNER JOIN
                dbo.Packages ON dbo.PackageDetails.PackageID = dbo.Packages.PackageID AND dbo.Packages.ProductType = 1 AND dbo.Packages.PackageStatusID IN (1, 2, 4) INNER JOIN
                JoinMainOrders ON DecorOrders.MainOrderID = JoinMainOrders.MarketMainOrderID INNER JOIN
                infiniu2_catalog.dbo.DecorProducts ON dbo.DecorOrders.ProductID = infiniu2_catalog.dbo.DecorProducts.ProductID INNER JOIN
                infiniu2_catalog.dbo.TechStore AS Decor ON dbo.DecorOrders.DecorID = Decor.TechStoreID LEFT OUTER JOIN
                infiniu2_catalog.dbo.TechStore AS Colors ON dbo.DecorOrders.ColorID = Colors.TechStoreID LEFT OUTER JOIN
                infiniu2_catalog.dbo.Patina AS Patina ON dbo.DecorOrders.PatinaID = Patina.PatinaID
                GROUP BY infiniu2_catalog.dbo.DecorProducts.ProductName, Decor.TechStoreName, Colors.TechStoreName, Patina.PatinaName
                ORDER BY ProductName, Decor, Color, Patina";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                BalancesDT.Dispose();
                BalancesDT = new DataTable();
                DA.Fill(BalancesDT);
                foreach (DataRow item in BalancesDT.Rows)
                    item["Expr1"] = Convert.ToDecimal(item["Expr1"]) / 1000000;
            }
            if (BalancesDT.Rows.Count > 0)
            {
                for (int i = 0; i < BalancesDT.Rows.Count; i++)
                {
                    Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                    Cell1.SetCellValue(BalancesDT.Rows[i]["ProductName"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(1);
                    Cell1.SetCellValue(BalancesDT.Rows[i]["Decor"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(2);
                    Cell1.SetCellValue(BalancesDT.Rows[i]["Color"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(3);
                    Cell1.SetCellValue(BalancesDT.Rows[i]["Patina"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(4);
                    Cell1.SetCellValue(Convert.ToDouble(BalancesDT.Rows[i]["Expr1"]));
                    Cell1.CellStyle = CountCS2;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(5);
                    Cell1.SetCellValue(Convert.ToInt32(BalancesDT.Rows[i]["Expr2"]));
                    Cell1.CellStyle = CountCS3;
                    pos++;
                }
                pos += 1;
            }
        }
    }
}