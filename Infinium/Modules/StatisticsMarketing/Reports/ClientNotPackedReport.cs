using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;

using System;
using System.Collections;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;

namespace Infinium.Modules.StatisticsMarketing.Reports
{
    public class ClientNotPackedReport
    {
        private DataTable ZOVClientsDataTable = null;
        private DataTable FrontsDT = null;
        private DataTable FrontsProfilDT = null;
        private DataTable FrontsTPSDT = null;
        private DataTable DecorDT = null;
        private DataTable DecorProfilDT = null;
        private DataTable DecorTPSDT = null;

        public ClientNotPackedReport()
        {
            FrontsDT = new DataTable();
            DecorDT = new DataTable();
        }

        public void FillTables(DateTime date1, DateTime date2, ArrayList MClients, ArrayList MClientGroups)
        {
            ZOVClientsDataTable = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT Clients.*, ClientsGroups.ClientGroupName, Managers.Name FROM Clients INNER JOIN
                         dbo.ClientsGroups ON dbo.Clients.ClientGroupID = dbo.ClientsGroups.ClientGroupID INNER JOIN
                         dbo.Managers ON dbo.Clients.ManagerID = dbo.Managers.ManagerID",
                ConnectionStrings.ZOVReferenceConnectionString))
            {
                DA.Fill(ZOVClientsDataTable);
            }

            string MClientFilter = string.Empty;
            if (MClients.Count > 0)
            {
                MClientFilter = " AND MegaOrders.ClientID IN (" + string.Join(",", MClients.OfType<Int32>().ToArray()) + ")";
            }
            if (MClientGroups.Count > 0)
            {
                MClientFilter = " AND MegaOrders.ClientID IN" +
                    " (SELECT ClientID FROM infiniu2_marketingreference.dbo.Clients" +
                    " WHERE ClientGroupID IN (" + string.Join(",", MClientGroups.OfType<Int32>().ToArray()) + "))";
            }
            if (MClients.Count < 1 && MClientGroups.Count < 1)
                MClientFilter = " AND MegaOrders.ClientID = -1";

            string Filter = " AND (MainOrders.DocDateTime >= '" + date1.ToString("yyyy-MM-dd") +
                " 23:59:59' AND MainOrders.DocDateTime <= '" + date2.ToString("yyyy-MM-dd") + " 23:59:59')";

            string SelectCommand = @"SELECT infiniu2_marketingreference.dbo.Clients.ClientName, dbo.MegaOrders.OrderNumber, MainOrders.DocDateTime, dbo.PackageDetails.PackageID, dbo.PackageDetails.PackNumber, dbo.FrontsOrders.FrontsOrdersID,
                        dbo.PackageDetails.Count, FrontsOrders.Square*dbo.PackageDetails.Count/FrontsOrders.Count as Square, FrontsOrders.FactoryID,
                         infiniu2_catalog.dbo.FrontsConfig.AccountingName, infiniu2_catalog.dbo.FrontsConfig.InvNumber, infiniu2_catalog.dbo.Measures.Measure, Fronts.TechStoreName, Colors.TechStoreName AS Expr35,
                         InsetTypes.TechStoreName AS Expr36, InsetColors.TechStoreName AS Expr37, TechnoInsetTypes.TechStoreName AS Expr38, TechnoInsetColors.TechStoreName AS Expr1, ZOVClientID
FROM dbo.PackageDetails INNER JOIN
                         dbo.FrontsOrders ON dbo.PackageDetails.OrderID = dbo.FrontsOrders.FrontsOrdersID INNER JOIN
                         dbo.Packages ON dbo.PackageDetails.PackageID = dbo.Packages.PackageID AND Packages.ProductType = 0 AND Packages.PackageStatusID=0 LEFT OUTER JOIN
                         JoinMainOrders ON dbo.FrontsOrders.MainOrderID = JoinMainOrders.MarketMainOrderID LEFT OUTER JOIN
                         infiniu2_catalog.dbo.TechStore AS Fronts ON dbo.FrontsOrders.FrontID = Fronts.TechStoreID LEFT OUTER JOIN
                         infiniu2_catalog.dbo.TechStore AS Colors ON dbo.FrontsOrders.ColorID = Colors.TechStoreID LEFT OUTER JOIN
                         infiniu2_catalog.dbo.TechStore AS InsetTypes ON dbo.FrontsOrders.InsetTypeID = InsetTypes.TechStoreID LEFT OUTER JOIN
                         infiniu2_catalog.dbo.TechStore AS InsetColors ON dbo.FrontsOrders.InsetColorID = InsetColors.TechStoreID LEFT OUTER JOIN
                         infiniu2_catalog.dbo.TechStore AS TechnoInsetTypes ON dbo.FrontsOrders.TechnoInsetTypeID = TechnoInsetTypes.TechStoreID LEFT OUTER JOIN
                         infiniu2_catalog.dbo.TechStore AS TechnoInsetColors ON dbo.FrontsOrders.TechnoInsetColorID = TechnoInsetColors.TechStoreID INNER JOIN
                         infiniu2_catalog.dbo.FrontsConfig ON dbo.FrontsOrders.FrontConfigID = infiniu2_catalog.dbo.FrontsConfig.FrontConfigID INNER JOIN
                         infiniu2_catalog.dbo.Measures ON infiniu2_catalog.dbo.FrontsConfig.MeasureID = infiniu2_catalog.dbo.Measures.MeasureID INNER JOIN
                         dbo.MainOrders ON dbo.FrontsOrders.MainOrderID = dbo.MainOrders.MainOrderID " + Filter + @" INNER JOIN
                         dbo.MegaOrders ON dbo.MainOrders.MegaOrderID = dbo.MegaOrders.MegaOrderID " + MClientFilter + @" INNER JOIN
                         infiniu2_marketingreference.dbo.Clients ON dbo.MegaOrders.ClientID = infiniu2_marketingreference.dbo.Clients.ClientID
ORDER BY infiniu2_marketingreference.dbo.Clients.ClientName, dbo.MegaOrders.OrderNumber, dbo.PackageDetails.PackageID";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(FrontsDT);
            }
            FrontsDT.Columns.Add(new DataColumn(("ZOVClientName"), System.Type.GetType("System.String")));
            for (int i = 0; i < FrontsDT.Rows.Count; i++)
            {
                if (FrontsDT.Rows[i]["ZOVClientID"] == DBNull.Value)
                    FrontsDT.Rows[i]["ZOVClientID"] = -1;
                else
                    FrontsDT.Rows[i]["ZOVClientName"] = GetZOVClientName(Convert.ToInt32(FrontsDT.Rows[i]["ZOVClientID"]));
            }

            SelectCommand = @"SELECT infiniu2_marketingreference.dbo.Clients.ClientName, dbo.MegaOrders.OrderNumber, MainOrders.DocDateTime, dbo.PackageDetails.PackageID, dbo.PackageDetails.PackNumber, dbo.DecorOrders.DecorOrderID,
                         DecorOrders.FactoryID, infiniu2_catalog.dbo.DecorConfig.AccountingName, infiniu2_catalog.dbo.DecorConfig.InvNumber, infiniu2_catalog.dbo.Measures.Measure, Fronts.TechStoreName, Colors.TechStoreName AS Expr35, dbo.DecorOrders.Length, dbo.DecorOrders.Height, dbo.DecorOrders.Width, dbo.PackageDetails.Count, ZOVClientID
FROM dbo.PackageDetails INNER JOIN
                         dbo.DecorOrders ON dbo.PackageDetails.OrderID = dbo.DecorOrders.DecorOrderID INNER JOIN
                         dbo.Packages ON dbo.PackageDetails.PackageID = dbo.Packages.PackageID AND Packages.ProductType = 1 AND Packages.PackageStatusID=0 LEFT OUTER JOIN
                         JoinMainOrders ON dbo.DecorOrders.MainOrderID = JoinMainOrders.MarketMainOrderID LEFT OUTER JOIN
                         infiniu2_catalog.dbo.TechStore AS Fronts ON dbo.DecorOrders.DecorID = Fronts.TechStoreID LEFT OUTER JOIN
                         infiniu2_catalog.dbo.TechStore AS Colors ON dbo.DecorOrders.ColorID = Colors.TechStoreID INNER JOIN
                         infiniu2_catalog.dbo.DecorConfig ON dbo.DecorOrders.DecorConfigID = infiniu2_catalog.dbo.DecorConfig.DecorConfigID INNER JOIN
                        infiniu2_catalog.dbo.Measures ON infiniu2_catalog.dbo.DecorConfig.MeasureID = infiniu2_catalog.dbo.Measures.MeasureID INNER JOIN
                         dbo.MainOrders ON dbo.DecorOrders.MainOrderID = dbo.MainOrders.MainOrderID " + Filter + @" INNER JOIN
                         dbo.MegaOrders ON dbo.MainOrders.MegaOrderID = dbo.MegaOrders.MegaOrderID " + MClientFilter + @" INNER JOIN
                         infiniu2_marketingreference.dbo.Clients ON dbo.MegaOrders.ClientID = infiniu2_marketingreference.dbo.Clients.ClientID
ORDER BY infiniu2_marketingreference.dbo.Clients.ClientName, dbo.MegaOrders.OrderNumber, dbo.PackageDetails.PackageID";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(DecorDT);
            }
            DecorDT.Columns.Add(new DataColumn(("ZOVClientName"), System.Type.GetType("System.String")));
            for (int i = 0; i < DecorDT.Rows.Count; i++)
            {
                if (DecorDT.Rows[i]["ZOVClientID"] == DBNull.Value)
                    DecorDT.Rows[i]["ZOVClientID"] = -1;
                else
                    DecorDT.Rows[i]["ZOVClientName"] = GetZOVClientName(Convert.ToInt32(DecorDT.Rows[i]["ZOVClientID"]));
            }
            if (FrontsProfilDT == null)
                FrontsProfilDT = FrontsDT.Clone();
            if (FrontsTPSDT == null)
                FrontsTPSDT = FrontsDT.Clone();
            if (DecorProfilDT == null)
                DecorProfilDT = DecorDT.Clone();
            if (DecorTPSDT == null)
                DecorTPSDT = DecorDT.Clone();
        }

        private string GetZOVClientName(int ClientID)
        {
            DataRow[] Rows = ZOVClientsDataTable.Select("ClientID = " + ClientID);
            if (Rows.Count() > 0)
                return Rows[0]["ClientName"].ToString();
            else
                return string.Empty;
        }

        private void DivideByFactory()
        {
            DataRow[] ItemsRows = FrontsDT.Select("FactoryID = 1");
            foreach (DataRow item in ItemsRows)
            {
                DataRow NewRow = FrontsProfilDT.NewRow();
                NewRow.ItemArray = item.ItemArray;
                FrontsProfilDT.Rows.Add(NewRow);
            }
            ItemsRows = FrontsDT.Select("FactoryID = 2");
            foreach (DataRow item in ItemsRows)
            {
                DataRow NewRow = FrontsTPSDT.NewRow();
                NewRow.ItemArray = item.ItemArray;
                FrontsTPSDT.Rows.Add(NewRow);
            }
            ItemsRows = DecorDT.Select("FactoryID = 1");
            foreach (DataRow item in ItemsRows)
            {
                DataRow NewRow = DecorProfilDT.NewRow();
                NewRow.ItemArray = item.ItemArray;
                DecorProfilDT.Rows.Add(NewRow);
            }
            ItemsRows = DecorDT.Select("FactoryID = 2");
            foreach (DataRow item in ItemsRows)
            {
                DataRow NewRow = DecorTPSDT.NewRow();
                NewRow.ItemArray = item.ItemArray;
                DecorTPSDT.Rows.Add(NewRow);
            }
        }

        public bool HasData
        {
            get
            {
                return FrontsDT.Rows.Count > 0 || DecorDT.Rows.Count > 0;
            }
        }

        private void ClearReport()
        {
            FrontsDT.Clear();
            FrontsProfilDT.Clear();
            FrontsTPSDT.Clear();
            DecorDT.Clear();
            DecorProfilDT.Clear();
            DecorTPSDT.Clear();
        }

        public void Report(string FileName)
        {
            DivideByFactory();
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

            HSSFCellStyle CountDecCS = hssfworkbook.CreateCellStyle();
            CountDecCS.DataFormat = hssfworkbook.CreateDataFormat().GetFormat("### ### ##0.000");
            CountDecCS.BorderBottom = HSSFCellStyle.BORDER_THIN;
            CountDecCS.BottomBorderColor = HSSFColor.BLACK.index;
            CountDecCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            CountDecCS.LeftBorderColor = HSSFColor.BLACK.index;
            CountDecCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            CountDecCS.RightBorderColor = HSSFColor.BLACK.index;
            CountDecCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            CountDecCS.TopBorderColor = HSSFColor.BLACK.index;
            CountDecCS.SetFont(SimpleF);

            HSSFCellStyle CountCS = hssfworkbook.CreateCellStyle();
            CountCS.DataFormat = hssfworkbook.CreateDataFormat().GetFormat("### ### ##0");
            CountCS.BorderBottom = HSSFCellStyle.BORDER_THIN;
            CountCS.BottomBorderColor = HSSFColor.BLACK.index;
            CountCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            CountCS.LeftBorderColor = HSSFColor.BLACK.index;
            CountCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            CountCS.RightBorderColor = HSSFColor.BLACK.index;
            CountCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            CountCS.TopBorderColor = HSSFColor.BLACK.index;
            CountCS.SetFont(SimpleF);

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

            HSSFCell Cell1;

            if (FrontsProfilDT.Rows.Count > 0)
            {
                HSSFSheet sheet1 = hssfworkbook.CreateSheet("Фасады, Профиль");
                sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

                sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
                sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
                sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
                sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

                pos += 2;
                int ColIndex = 0;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("Клиент");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("Клиент ЗОВ");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("№ заказа");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("Дата создания");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("Бухг.наим.");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("Инв.номер");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("Фасад");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("Цвет");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("Вставка");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("Цвет наполнителя");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("Вставка-2");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("Цвет наполнителя-2");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("Кол-во");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("Квадратура");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("Ед.изм.");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("ID упаковки");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("№ упаковки");
                Cell1.CellStyle = SimpleHeaderCS;

                pos++;

                for (int i = 0; i < FrontsProfilDT.Rows.Count; i++)
                {
                    ColIndex = 0;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(FrontsProfilDT.Rows[i]["ClientName"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(FrontsProfilDT.Rows[i]["ZOVClientName"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(Convert.ToInt32(FrontsProfilDT.Rows[i]["OrderNumber"]));
                    Cell1.CellStyle = CountCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(FrontsProfilDT.Rows[i]["DocDateTime"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(FrontsProfilDT.Rows[i]["AccountingName"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(FrontsProfilDT.Rows[i]["InvNumber"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(FrontsProfilDT.Rows[i]["TechStoreName"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(FrontsProfilDT.Rows[i]["Expr35"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(FrontsProfilDT.Rows[i]["Expr36"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(FrontsProfilDT.Rows[i]["Expr37"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(FrontsProfilDT.Rows[i]["Expr38"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(FrontsProfilDT.Rows[i]["Expr1"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(Convert.ToInt32(FrontsProfilDT.Rows[i]["Count"]));
                    Cell1.CellStyle = CountCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(Convert.ToDouble(FrontsProfilDT.Rows[i]["Square"]));
                    Cell1.CellStyle = CountDecCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(FrontsProfilDT.Rows[i]["Measure"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(Convert.ToInt32(FrontsProfilDT.Rows[i]["PackageID"]));
                    Cell1.CellStyle = CountCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(Convert.ToInt32(FrontsProfilDT.Rows[i]["PackNumber"]));
                    Cell1.CellStyle = CountCS;

                    pos++;
                }
            }
            pos = 0;
            if (FrontsTPSDT.Rows.Count > 0)
            {
                HSSFSheet sheet1 = hssfworkbook.CreateSheet("Фасады, ТПС");
                sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

                sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
                sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
                sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
                sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

                pos += 2;
                int ColIndex = 0;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("Клиент");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("Клиент ЗОВ");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("№ заказа");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("Дата создания");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("Бухг.наим.");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("Инв.номер");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("Фасад");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("Цвет");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("Вставка");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("Цвет наполнителя");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("Вставка-2");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("Цвет наполнителя-2");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("Кол-во");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("Квадратура");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("Ед.изм.");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("ID упаковки");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("№ упаковки");
                Cell1.CellStyle = SimpleHeaderCS;

                pos++;

                for (int i = 0; i < FrontsTPSDT.Rows.Count; i++)
                {
                    ColIndex = 0;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(FrontsTPSDT.Rows[i]["ClientName"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(FrontsTPSDT.Rows[i]["ZOVClientName"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(Convert.ToInt32(FrontsTPSDT.Rows[i]["OrderNumber"]));
                    Cell1.CellStyle = CountCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(FrontsTPSDT.Rows[i]["DocDateTime"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(FrontsTPSDT.Rows[i]["AccountingName"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(FrontsTPSDT.Rows[i]["InvNumber"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(FrontsTPSDT.Rows[i]["TechStoreName"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(FrontsTPSDT.Rows[i]["Expr35"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(FrontsTPSDT.Rows[i]["Expr36"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(FrontsTPSDT.Rows[i]["Expr37"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(FrontsTPSDT.Rows[i]["Expr38"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(FrontsTPSDT.Rows[i]["Expr1"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(Convert.ToInt32(FrontsTPSDT.Rows[i]["Count"]));
                    Cell1.CellStyle = CountCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(Convert.ToDouble(FrontsTPSDT.Rows[i]["Square"]));
                    Cell1.CellStyle = CountDecCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(FrontsTPSDT.Rows[i]["Measure"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(Convert.ToInt32(FrontsTPSDT.Rows[i]["PackageID"]));
                    Cell1.CellStyle = CountCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(Convert.ToInt32(FrontsTPSDT.Rows[i]["PackNumber"]));
                    Cell1.CellStyle = CountCS;

                    pos++;
                }
            }

            pos = 0;
            if (DecorProfilDT.Rows.Count > 0)
            {
                HSSFSheet sheet1 = hssfworkbook.CreateSheet("Декор, Профиль");
                sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

                sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
                sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
                sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
                sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

                pos += 2;
                int ColIndex = 0;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("Клиент");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("Клиент ЗОВ");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("№ заказа");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("Дата создания");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("Бухг.наим.");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("Инв.номер");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("Артикул");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("Цвет");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("Длина");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("Высота");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("Ширина");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("Кол-во");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("Ед.изм.");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("ID упаковки");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("№ упаковки");
                Cell1.CellStyle = SimpleHeaderCS;

                pos++;

                for (int i = 0; i < DecorProfilDT.Rows.Count; i++)
                {
                    ColIndex = 0;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(DecorProfilDT.Rows[i]["ClientName"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(DecorProfilDT.Rows[i]["ZOVClientName"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(Convert.ToInt32(DecorProfilDT.Rows[i]["OrderNumber"]));
                    Cell1.CellStyle = CountCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(DecorProfilDT.Rows[i]["DocDateTime"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(DecorProfilDT.Rows[i]["AccountingName"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(DecorProfilDT.Rows[i]["InvNumber"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(DecorProfilDT.Rows[i]["TechStoreName"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(DecorProfilDT.Rows[i]["Expr35"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(Convert.ToInt32(DecorProfilDT.Rows[i]["Length"]));
                    Cell1.CellStyle = CountCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(Convert.ToInt32(DecorProfilDT.Rows[i]["Height"]));
                    Cell1.CellStyle = CountCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(Convert.ToInt32(DecorProfilDT.Rows[i]["Width"]));
                    Cell1.CellStyle = CountCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(Convert.ToInt32(DecorProfilDT.Rows[i]["Count"]));
                    Cell1.CellStyle = CountCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(DecorProfilDT.Rows[i]["Measure"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(Convert.ToInt32(DecorProfilDT.Rows[i]["PackageID"]));
                    Cell1.CellStyle = CountCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(Convert.ToInt32(DecorProfilDT.Rows[i]["PackNumber"]));
                    Cell1.CellStyle = CountCS;

                    pos++;
                }
            }
            pos = 0;
            if (DecorTPSDT.Rows.Count > 0)
            {
                HSSFSheet sheet1 = hssfworkbook.CreateSheet("Декор, ТПС");
                sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

                sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
                sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
                sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
                sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

                pos += 2;
                int ColIndex = 0;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("Клиент");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("Клиент ЗОВ");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("№ заказа");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("Дата создания");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("Бухг.наим.");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("Инв.номер");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("Артикул");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("Цвет");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("Длина");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("Высота");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("Ширина");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("Кол-во");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("Ед.изм.");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("ID упаковки");
                Cell1.CellStyle = SimpleHeaderCS;
                Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                Cell1.SetCellValue("№ упаковки");
                Cell1.CellStyle = SimpleHeaderCS;

                pos++;

                for (int i = 0; i < DecorTPSDT.Rows.Count; i++)
                {
                    ColIndex = 0;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(DecorTPSDT.Rows[i]["ClientName"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(DecorTPSDT.Rows[i]["ZOVClientName"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(Convert.ToInt32(DecorTPSDT.Rows[i]["OrderNumber"]));
                    Cell1.CellStyle = CountCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(DecorTPSDT.Rows[i]["DocDateTime"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(DecorTPSDT.Rows[i]["AccountingName"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(DecorTPSDT.Rows[i]["InvNumber"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(DecorTPSDT.Rows[i]["TechStoreName"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(DecorTPSDT.Rows[i]["Expr35"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(Convert.ToInt32(DecorTPSDT.Rows[i]["Length"]));
                    Cell1.CellStyle = CountCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(Convert.ToInt32(DecorTPSDT.Rows[i]["Height"]));
                    Cell1.CellStyle = CountCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(Convert.ToInt32(DecorTPSDT.Rows[i]["Width"]));
                    Cell1.CellStyle = CountCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(Convert.ToInt32(DecorTPSDT.Rows[i]["Count"]));
                    Cell1.CellStyle = CountCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(DecorTPSDT.Rows[i]["Measure"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(Convert.ToInt32(DecorTPSDT.Rows[i]["PackageID"]));
                    Cell1.CellStyle = CountCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(ColIndex++);
                    Cell1.SetCellValue(Convert.ToInt32(DecorTPSDT.Rows[i]["PackNumber"]));
                    Cell1.CellStyle = CountCS;

                    pos++;
                }
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
            ClearReport();
        }
    }
}