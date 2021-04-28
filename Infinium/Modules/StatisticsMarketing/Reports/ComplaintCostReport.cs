using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;

using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;

namespace Infinium.Modules.StatisticsMarketing.Reports
{
    public class ComplaintCostReport
    {
        private DataTable ComplaintCostDT = null;

        public ComplaintCostReport()
        {
            ComplaintCostDT = new DataTable();
        }

        public void FilterByOrderDate(DateTime DateFrom, DateTime DateTo)
        {
            string Filter = " WHERE CAST(NewMegaOrders.OrderDate AS date) >= '" + DateFrom.ToString("yyyy-MM-dd") +
                " 00:00' AND CAST(NewMegaOrders.OrderDate AS date) <= '" + DateTo.ToString("yyyy-MM-dd") + " 23:59'";

            string SelectCommand = @"SELECT clients.ClientName, NewMegaOrders.OrderNumber, ClientsManagers.Name, NewMegaOrders.ConfirmDateTime, NewMegaOrders.ComplaintProfilCost + NewMegaOrders.ComplaintTPSCost, NewMegaOrders.TotalCost, CurrencyTypes.CurrencyType, NewMegaOrders.Rate, NewMegaOrders.PaymentRate, NewMegaOrders.CurrencyComplaintProfilCost + NewMegaOrders.CurrencyComplaintTPSCost, NewMegaOrders.CurrencyTotalCost
                                     infiniu2_marketingreference.dbo.Clients AS clients ON NewMegaOrders.ClientID = clients.ClientID INNER JOIN
                                     infiniu2_marketingreference.dbo.ClientsManagers AS ClientsManagers ON clients.ManagerID = ClientsManagers.ManagerID INNER JOIN
                                     infiniu2_catalog.dbo.CurrencyTypes AS CurrencyTypes ON NewMegaOrders.CurrencyTypeID = CurrencyTypes.CurrencyTypeID
            " + Filter + @" AND (NewMegaOrders.ComplaintProfilCost > 0 OR NewMegaOrders.ComplaintTPSCost > 0 OR NewMegaOrders.IsComplaint = 1)
            ORDER BY clients.ClientName, NewMegaOrders.OrderNumber";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                ComplaintCostDT.Clear();
                DA.Fill(ComplaintCostDT);
            }
        }

        public void FilterByConfirmDate(DateTime DateFrom, DateTime DateTo)
        {
            string Filter = " WHERE CAST(MegaOrders.ConfirmDateTime AS date) >= '" + DateFrom.ToString("yyyy-MM-dd") +
                " 00:00' AND CAST(MegaOrders.ConfirmDateTime AS date) <= '" + DateTo.ToString("yyyy-MM-dd") + " 23:59'";

            string SelectCommand = @"SELECT clients.ClientName, MegaOrders.OrderNumber, ClientsManagers.Name, MegaOrders.ConfirmDateTime, MegaOrders.ComplaintProfilCost + MegaOrders.ComplaintTPSCost, MegaOrders.TotalCost, CurrencyTypes.CurrencyType, MegaOrders.Rate, MegaOrders.PaymentRate, MegaOrders.CurrencyComplaintProfilCost + MegaOrders.CurrencyComplaintTPSCost, MegaOrders.CurrencyTotalCost
                                     infiniu2_marketingreference.dbo.Clients AS clients ON MegaOrders.ClientID = clients.ClientID INNER JOIN
                                     infiniu2_marketingreference.dbo.ClientsManagers AS ClientsManagers ON clients.ManagerID = ClientsManagers.ManagerID INNER JOIN
                                     infiniu2_catalog.dbo.CurrencyTypes AS CurrencyTypes ON MegaOrders.CurrencyTypeID = CurrencyTypes.CurrencyTypeID
            " + Filter + @" AND (MegaOrders.ComplaintProfilCost > 0 OR MegaOrders.ComplaintTPSCost > 0 OR MegaOrders.IsComplaint = 1)
            ORDER BY clients.ClientName, MegaOrders.OrderNumber";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                ComplaintCostDT.Clear();
                DA.Fill(ComplaintCostDT);
            }
        }

        public void FilterByPlanDispatch(DateTime DateFrom, DateTime DateTo)
        {
            string Filter = " WHERE ((CAST(ProfilDispatchDate AS DATE) >= '" + DateFrom.ToString("yyyy-MM-dd") +
                    "' AND CAST(ProfilDispatchDate AS DATE) <= '" + DateTo.ToString("yyyy-MM-dd") + "')" +
                    " OR (CAST(TPSDispatchDate AS DATE) >= '" + DateFrom.ToString("yyyy-MM-dd") +
                    "' AND CAST(TPSDispatchDate AS DATE) <= '" + DateTo.ToString("yyyy-MM-dd") + "')) ";

            string SelectCommand = @"SELECT clients.ClientName, MegaOrders.OrderNumber, ClientsManagers.Name, MegaOrders.ConfirmDateTime, MegaOrders.ComplaintProfilCost + MegaOrders.ComplaintTPSCost, MegaOrders.TotalCost, CurrencyTypes.CurrencyType, MegaOrders.Rate, MegaOrders.PaymentRate, MegaOrders.CurrencyComplaintProfilCost + MegaOrders.CurrencyComplaintTPSCost, MegaOrders.CurrencyTotalCost
                                     infiniu2_marketingreference.dbo.Clients AS clients ON MegaOrders.ClientID = clients.ClientID INNER JOIN
                                     infiniu2_marketingreference.dbo.ClientsManagers AS ClientsManagers ON clients.ManagerID = ClientsManagers.ManagerID INNER JOIN
                                     infiniu2_catalog.dbo.CurrencyTypes AS CurrencyTypes ON MegaOrders.CurrencyTypeID = CurrencyTypes.CurrencyTypeID
            " + Filter + @" AND (MegaOrders.ComplaintProfilCost > 0 OR MegaOrders.ComplaintTPSCost > 0 OR MegaOrders.IsComplaint = 1)
            ORDER BY clients.ClientName, MegaOrders.OrderNumber";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                ComplaintCostDT.Clear();
                DA.Fill(ComplaintCostDT);
            }
        }

        public void FilterByOnProduction(DateTime DateFrom, DateTime DateTo)
        {
            string Filter = " WHERE ((CAST(ProfilOnProductionDate AS Date) >= '" + DateFrom.ToString("yyyy-MM-dd") + "' AND CAST(ProfilOnProductionDate AS Date) <= '" + DateTo.ToString("yyyy-MM-dd") +
                    "') OR (CAST(TPSOnProductionDate AS Date) >= '" + DateFrom.ToString("yyyy-MM-dd") + "' AND CAST(TPSOnProductionDate AS Date) <= '" + DateTo.ToString("yyyy-MM-dd") + "'))";
            string SelectCommand = @"SELECT clients.ClientName, MegaOrders.OrderNumber, ClientsManagers.Name, MegaOrders.ConfirmDateTime, MegaOrders.ComplaintProfilCost + MegaOrders.ComplaintTPSCost, MegaOrders.TotalCost, CurrencyTypes.CurrencyType, MegaOrders.Rate, MegaOrders.PaymentRate, MegaOrders.CurrencyComplaintProfilCost + MegaOrders.CurrencyComplaintTPSCost, MegaOrders.CurrencyTotalCost
                                     infiniu2_marketingreference.dbo.Clients AS clients ON MegaOrders.ClientID = clients.ClientID INNER JOIN
                                     infiniu2_marketingreference.dbo.ClientsManagers AS ClientsManagers ON clients.ManagerID = ClientsManagers.ManagerID INNER JOIN
                                     infiniu2_catalog.dbo.CurrencyTypes AS CurrencyTypes ON MegaOrders.CurrencyTypeID = CurrencyTypes.CurrencyTypeID
            " + Filter + @" AND (MegaOrders.ComplaintProfilCost > 0 OR MegaOrders.ComplaintTPSCost > 0 OR MegaOrders.IsComplaint = 1)
            ORDER BY clients.ClientName, MegaOrders.OrderNumber";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                ComplaintCostDT.Clear();
                DA.Fill(ComplaintCostDT);
            }
        }

        public void FilterByOnAgreement(DateTime DateFrom, DateTime DateTo)
        {
            string Filter = " WHERE CAST(MegaOrders.OnAgreementDateTime AS date) >= '" + DateFrom.ToString("yyyy-MM-dd") +
                " 00:00' AND CAST(MegaOrders.OnAgreementDateTime AS date) <= '" + DateTo.ToString("yyyy-MM-dd") + " 23:59'";

            string SelectCommand = @"SELECT clients.ClientName, MegaOrders.OrderNumber, ClientsManagers.Name, MegaOrders.ConfirmDateTime, MegaOrders.ComplaintProfilCost + MegaOrders.ComplaintTPSCost, MegaOrders.TotalCost, CurrencyTypes.CurrencyType, MegaOrders.Rate, MegaOrders.PaymentRate, MegaOrders.CurrencyComplaintProfilCost + MegaOrders.CurrencyComplaintTPSCost, MegaOrders.CurrencyTotalCost
            FROM            MegaOrders INNER JOIN
                                     infiniu2_marketingreference.dbo.Clients AS clients ON MegaOrders.ClientID = clients.ClientID INNER JOIN
                                     infiniu2_marketingreference.dbo.ClientsManagers AS ClientsManagers ON clients.ManagerID = ClientsManagers.ManagerID INNER JOIN
                                     infiniu2_catalog.dbo.CurrencyTypes AS CurrencyTypes ON MegaOrders.CurrencyTypeID = CurrencyTypes.CurrencyTypeID
            " + Filter + @" AND (MegaOrders.ComplaintProfilCost > 0 OR MegaOrders.ComplaintTPSCost > 0 OR MegaOrders.IsComplaint = 1)
            ORDER BY clients.ClientName, MegaOrders.OrderNumber";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                ComplaintCostDT.Clear();
                DA.Fill(ComplaintCostDT);
            }
        }

        public void FilterByPackages(DateTime DateFrom, DateTime DateTo, int PackageStatusID)
        {
            string Date = "PackingDateTime";

            if (PackageStatusID == 2)
                Date = "StorageDateTime";
            if (PackageStatusID == 3)
                Date = "DispatchDateTime";
            if (PackageStatusID == 4)
                Date = "ExpeditionDateTime";
            string Filter = " WHERE MegaOrderID IN (SELECT MegaOrderID FROM MainOrders WHERE MainOrderID IN (SELECT MainOrderID FROM Packages WHERE CAST(" + Date + " AS DATE) >= '" + DateFrom.ToString("yyyy-MM-dd") +
                "' AND CAST(" + Date + " AS DATE) <= '" + DateTo.ToString("yyyy-MM-dd") + "'))"; ;

            string SelectCommand = @"SELECT clients.ClientName, MegaOrders.OrderNumber, ClientsManagers.Name, MegaOrders.ConfirmDateTime, MegaOrders.ComplaintProfilCost + MegaOrders.ComplaintTPSCost, MegaOrders.TotalCost, CurrencyTypes.CurrencyType, MegaOrders.Rate, MegaOrders.PaymentRate, MegaOrders.CurrencyComplaintProfilCost + MegaOrders.CurrencyComplaintTPSCost, MegaOrders.CurrencyTotalCost
            FROM            MegaOrders INNER JOIN
                                     infiniu2_marketingreference.dbo.Clients AS clients ON MegaOrders.ClientID = clients.ClientID INNER JOIN
                                     infiniu2_marketingreference.dbo.ClientsManagers AS ClientsManagers ON clients.ManagerID = ClientsManagers.ManagerID INNER JOIN
                                     infiniu2_catalog.dbo.CurrencyTypes AS CurrencyTypes ON MegaOrders.CurrencyTypeID = CurrencyTypes.CurrencyTypeID
            " + Filter + @" AND (MegaOrders.ComplaintProfilCost > 0 OR MegaOrders.ComplaintTPSCost > 0 OR MegaOrders.IsComplaint = 1)
            ORDER BY clients.ClientName, MegaOrders.OrderNumber";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                ComplaintCostDT.Clear();
                DA.Fill(ComplaintCostDT);
            }
        }

        public bool HasData
        {
            get
            {
                return ComplaintCostDT.Rows.Count > 0;
            }
        }

        public void Report(string FileName)
        {
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

            int DisplayIndex = 0;
            sheet1.SetColumnWidth(DisplayIndex++, 40 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 9 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 15 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 14 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 14 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 14 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 9 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 9 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 9 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 13 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 13 * 256);

            HSSFCell Cell1;

            DisplayIndex = 0;
            if (ComplaintCostDT.Rows.Count > 0)
            {
                pos += 2;

                Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                Cell1.SetCellValue("Клиент");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                Cell1.SetCellValue("№ заказа");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                Cell1.SetCellValue("Менеджер");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                Cell1.SetCellValue("Дата согласования");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                Cell1.SetCellValue("Сумма рекламации");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                Cell1.SetCellValue("Сумма заказа");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                Cell1.SetCellValue("Валюта");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                Cell1.SetCellValue("Курс");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                Cell1.SetCellValue("Расчетный курс");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                Cell1.SetCellValue("Итого сумма рекламации");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                Cell1.SetCellValue("Итого сумма заказа");
                Cell1.CellStyle = SimpleHeaderCS;

                pos++;

                int ColumnCount = ComplaintCostDT.Columns.Count;
                for (int x = 0; x < ComplaintCostDT.Rows.Count; x++)
                {
                    for (int y = 0; y < ColumnCount; y++)
                    {
                        Type t = ComplaintCostDT.Rows[x][y].GetType();

                        if (t.Name == "Decimal")
                        {
                            HSSFCell cell = sheet1.CreateRow(pos).CreateCell(y);
                            cell.SetCellValue(Convert.ToDouble(ComplaintCostDT.Rows[x][y]));

                            cell.CellStyle = CountCS;
                            continue;
                        }
                        if (t.Name == "Int32")
                        {
                            HSSFCell cell = sheet1.CreateRow(pos).CreateCell(y);
                            cell.SetCellValue(Convert.ToInt32(ComplaintCostDT.Rows[x][y]));
                            cell.CellStyle = SimpleCS;
                            continue;
                        }
                        if (t.Name == "Boolean")
                        {
                            bool b = Convert.ToBoolean(ComplaintCostDT.Rows[x][y]);
                            string str = "Да";
                            if (!b)
                                str = "Нет";
                            HSSFCell cell = sheet1.CreateRow(pos).CreateCell(y);
                            cell.SetCellValue(str);
                            cell.CellStyle = SimpleCS;
                            continue;
                        }
                        if (t.Name == "DateTime")
                        {
                            string dateTime = Convert.ToDateTime(ComplaintCostDT.Rows[x][y]).ToShortDateString();
                            HSSFCell cell = sheet1.CreateRow(pos).CreateCell(y);
                            cell.SetCellValue(dateTime);
                            cell.CellStyle = SimpleCS;
                            continue;
                        }

                        if (t.Name == "String" || t.Name == "DBNull")
                        {
                            HSSFCell cell = sheet1.CreateRow(pos).CreateCell(y);
                            cell.SetCellValue(ComplaintCostDT.Rows[x][y].ToString());
                            cell.CellStyle = SimpleCS;
                            continue;
                        }
                    }
                    pos++;
                }

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
        }
    }
}