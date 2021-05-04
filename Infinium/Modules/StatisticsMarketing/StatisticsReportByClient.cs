using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.UserModel.Contrib;
using NPOI.HSSF.Util;

using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;

namespace Infinium.Modules.StatisticsMarketing
{
    public class StatisticsReportByClient : IAllFrontParameterName
    {
        private DataTable ClientsDataTable = null;
        private ZOV.DecorCatalogOrder DecorCatalog = null;
        private DataTable[] DecorResultDataTable = null;
        private DataTable FrameColorsDataTable = null;
        private DataTable FrontsDataTable = null;
        private DataTable FrontsResultDataTable = null;
        private DataTable InsetColorsDataTable = null;
        private DataTable InsetTypesDataTable = null;
        private DataTable PatinaDataTable = null;
        private DataTable PatinaRALDataTable = null;
        private DataTable ZOVClientsDataTable = null;

        public StatisticsReportByClient(ref ZOV.DecorCatalogOrder tDecorCatalog)
        {
            DecorCatalog = tDecorCatalog;

            Create();

            ClientsDataTable = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT Clients.*, infiniu2_catalog.dbo.Countries.Name AS CountryName, ClientsManagers.ShortName FROM Clients INNER JOIN
                         infiniu2_catalog.dbo.Countries ON dbo.Clients.CountryID = infiniu2_catalog.dbo.Countries.CountryID INNER JOIN
                         ClientsManagers ON dbo.Clients.ManagerID = ClientsManagers.ManagerID",
                ConnectionStrings.MarketingReferenceConnectionString))
            {
                DA.Fill(ClientsDataTable);
            }

            ZOVClientsDataTable = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT Clients.*, ClientsGroups.ClientGroupName, Managers.Name FROM Clients INNER JOIN
                         dbo.ClientsGroups ON dbo.Clients.ClientGroupID = dbo.ClientsGroups.ClientGroupID INNER JOIN
                         dbo.Managers ON dbo.Clients.ManagerID = dbo.Managers.ManagerID",
                ConnectionStrings.ZOVReferenceConnectionString))
            {
                DA.Fill(ZOVClientsDataTable);
            }

            string SelectCommand = @"SELECT TechStoreID AS FrontID, TechStoreName AS FrontName FROM TechStore
                WHERE TechStoreID IN (SELECT FrontID FROM FrontsConfig WHERE Enabled = 1)
                ORDER BY TechStoreName";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(FrontsDataTable);
            }
            GetColorsDT();
            SelectCommand = @"SELECT * FROM Patina";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(PatinaDataTable);
            }
            PatinaRALDataTable = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM PatinaRAL WHERE Enabled=1",
                ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(PatinaRALDataTable);
            }
            foreach (DataRow item in PatinaRALDataTable.Rows)
            {
                DataRow NewRow = PatinaDataTable.NewRow();
                NewRow["PatinaID"] = item["PatinaRALID"];
                NewRow["PatinaName"] = item["PatinaRAL"];
                NewRow["DisplayName"] = item["DisplayName"];
                PatinaDataTable.Rows.Add(NewRow);
            }
            SelectCommand = @"SELECT * FROM InsetTypes";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(InsetTypesDataTable);
            }
            SelectCommand = @"SELECT * FROM InsetColors";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(InsetColorsDataTable);
            }
            CreateFrontsDataTable();
            CreateDecorDataTable();
        }

        public void CreateReportOrderNumber(DateTime DateFrom, DateTime DateTo, int FactoryID,
            DataTable FrontsOrdersDataTable, DataTable DecorOrdersDataTable, string FileName, bool ZOV)
        {
            if (FrontsOrdersDataTable.Rows.Count < 1 && DecorOrdersDataTable.Rows.Count < 1)
                return;

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

            HSSFFont ClientNameFont = hssfworkbook.CreateFont();
            ClientNameFont.FontHeightInPoints = 14;
            ClientNameFont.Boldweight = HSSFFont.BOLDWEIGHT_BOLD;
            ClientNameFont.FontName = "Calibri";

            HSSFFont MainFont = hssfworkbook.CreateFont();
            MainFont.FontHeightInPoints = 12;
            MainFont.Boldweight = HSSFFont.BOLDWEIGHT_BOLD;
            MainFont.FontName = "Calibri";

            HSSFCellStyle ClientNameStyle = hssfworkbook.CreateCellStyle();
            ClientNameStyle.SetFont(ClientNameFont);

            HSSFCellStyle MainStyle = hssfworkbook.CreateCellStyle();
            MainStyle.SetFont(MainFont);

            HSSFFont HeaderFont = hssfworkbook.CreateFont();
            HeaderFont.FontHeightInPoints = 12;
            HeaderFont.Boldweight = HSSFFont.BOLDWEIGHT_BOLD;
            HeaderFont.FontName = "Calibri";

            HSSFCellStyle HeaderStyle = hssfworkbook.CreateCellStyle();
            HeaderStyle.BorderBottom = HSSFCellStyle.BORDER_THIN;
            HeaderStyle.BottomBorderColor = HSSFColor.BLACK.index;
            HeaderStyle.BorderLeft = HSSFCellStyle.BORDER_THIN;
            HeaderStyle.LeftBorderColor = HSSFColor.BLACK.index;
            HeaderStyle.BorderRight = HSSFCellStyle.BORDER_THIN;
            HeaderStyle.RightBorderColor = HSSFColor.BLACK.index;
            HeaderStyle.BorderTop = HSSFCellStyle.BORDER_THIN;
            HeaderStyle.TopBorderColor = HSSFColor.BLACK.index;
            HeaderStyle.SetFont(HeaderFont);

            HSSFFont PackNumberFont = hssfworkbook.CreateFont();
            PackNumberFont.Boldweight = 11 * 256;
            PackNumberFont.FontName = "Calibri";

            HSSFCellStyle PackNumberStyle = hssfworkbook.CreateCellStyle();
            PackNumberStyle.BorderBottom = HSSFCellStyle.BORDER_THIN;
            PackNumberStyle.BottomBorderColor = HSSFColor.BLACK.index;
            PackNumberStyle.BorderLeft = HSSFCellStyle.BORDER_THIN;
            PackNumberStyle.LeftBorderColor = HSSFColor.BLACK.index;
            PackNumberStyle.BorderRight = HSSFCellStyle.BORDER_THIN;
            PackNumberStyle.RightBorderColor = HSSFColor.BLACK.index;
            PackNumberStyle.BorderTop = HSSFCellStyle.BORDER_THIN;
            PackNumberStyle.TopBorderColor = HSSFColor.BLACK.index;
            PackNumberStyle.Alignment = HSSFCellStyle.ALIGN_CENTER;
            PackNumberStyle.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
            PackNumberStyle.SetFont(PackNumberFont);

            HSSFFont SimpleFont = hssfworkbook.CreateFont();
            SimpleFont.FontHeightInPoints = 11;
            SimpleFont.FontName = "Calibri";

            HSSFCellStyle SimpleCellStyle = hssfworkbook.CreateCellStyle();
            SimpleCellStyle.BorderBottom = HSSFCellStyle.BORDER_THIN;
            SimpleCellStyle.BottomBorderColor = HSSFColor.BLACK.index;
            SimpleCellStyle.BorderLeft = HSSFCellStyle.BORDER_THIN;
            SimpleCellStyle.LeftBorderColor = HSSFColor.BLACK.index;
            SimpleCellStyle.BorderRight = HSSFCellStyle.BORDER_THIN;
            SimpleCellStyle.RightBorderColor = HSSFColor.BLACK.index;
            SimpleCellStyle.BorderTop = HSSFCellStyle.BORDER_THIN;
            SimpleCellStyle.TopBorderColor = HSSFColor.BLACK.index;
            SimpleCellStyle.SetFont(SimpleFont);

            HSSFCellStyle cellStyle = hssfworkbook.CreateCellStyle();
            cellStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.000");
            cellStyle.BorderBottom = HSSFCellStyle.BORDER_THIN;
            cellStyle.BottomBorderColor = HSSFColor.BLACK.index;
            cellStyle.BorderLeft = HSSFCellStyle.BORDER_THIN;
            cellStyle.LeftBorderColor = HSSFColor.BLACK.index;
            cellStyle.BorderRight = HSSFCellStyle.BORDER_THIN;
            cellStyle.RightBorderColor = HSSFColor.BLACK.index;
            cellStyle.BorderTop = HSSFCellStyle.BORDER_THIN;
            cellStyle.TopBorderColor = HSSFColor.BLACK.index;
            cellStyle.SetFont(SimpleFont);

            HSSFCellStyle GreyCellStyle = hssfworkbook.CreateCellStyle();
            GreyCellStyle.BorderBottom = HSSFCellStyle.BORDER_THIN;
            GreyCellStyle.BottomBorderColor = HSSFColor.BLACK.index;
            GreyCellStyle.BorderLeft = HSSFCellStyle.BORDER_THIN;
            GreyCellStyle.LeftBorderColor = HSSFColor.BLACK.index;
            GreyCellStyle.BorderRight = HSSFCellStyle.BORDER_THIN;
            GreyCellStyle.RightBorderColor = HSSFColor.BLACK.index;
            GreyCellStyle.BorderTop = HSSFCellStyle.BORDER_THIN;
            GreyCellStyle.TopBorderColor = HSSFColor.BLACK.index;
            GreyCellStyle.FillForegroundColor = HSSFColor.GREY_25_PERCENT.index;
            GreyCellStyle.FillPattern = HSSFCellStyle.SOLID_FOREGROUND;
            GreyCellStyle.FillBackgroundColor = HSSFColor.YELLOW.index;
            GreyCellStyle.SetFont(SimpleFont);

            HSSFFont TotalFont = hssfworkbook.CreateFont();
            TotalFont.FontHeightInPoints = 11;
            TotalFont.FontName = "Calibri";

            HSSFCellStyle TotalStyle = hssfworkbook.CreateCellStyle();
            TotalStyle.BorderTop = HSSFCellStyle.BORDER_MEDIUM;
            TotalStyle.TopBorderColor = HSSFColor.BLACK.index;
            TotalStyle.SetFont(TotalFont);

            HSSFCellStyle TempStyle = hssfworkbook.CreateCellStyle();
            TempStyle.SetFont(TotalFont);

            #endregion Create fonts and styles

            #region границы между упаковками

            HSSFCellStyle BottomMediumBorderCellStyle = hssfworkbook.CreateCellStyle();
            BottomMediumBorderCellStyle.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            BottomMediumBorderCellStyle.BottomBorderColor = HSSFColor.BLACK.index;
            BottomMediumBorderCellStyle.BorderLeft = HSSFCellStyle.BORDER_THIN;
            BottomMediumBorderCellStyle.LeftBorderColor = HSSFColor.BLACK.index;
            BottomMediumBorderCellStyle.BorderRight = HSSFCellStyle.BORDER_THIN;
            BottomMediumBorderCellStyle.RightBorderColor = HSSFColor.BLACK.index;
            BottomMediumBorderCellStyle.BorderTop = HSSFCellStyle.BORDER_THIN;
            BottomMediumBorderCellStyle.TopBorderColor = HSSFColor.BLACK.index;
            BottomMediumBorderCellStyle.SetFont(SimpleFont);

            HSSFCellStyle BottomMediumLeftBorderCellStyle = hssfworkbook.CreateCellStyle();
            BottomMediumLeftBorderCellStyle.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            BottomMediumLeftBorderCellStyle.BottomBorderColor = HSSFColor.BLACK.index;
            BottomMediumLeftBorderCellStyle.BorderLeft = HSSFCellStyle.BORDER_MEDIUM;
            BottomMediumLeftBorderCellStyle.LeftBorderColor = HSSFColor.BLACK.index;
            BottomMediumLeftBorderCellStyle.BorderRight = HSSFCellStyle.BORDER_THIN;
            BottomMediumLeftBorderCellStyle.RightBorderColor = HSSFColor.BLACK.index;
            BottomMediumLeftBorderCellStyle.BorderTop = HSSFCellStyle.BORDER_THIN;
            BottomMediumLeftBorderCellStyle.TopBorderColor = HSSFColor.BLACK.index;
            BottomMediumLeftBorderCellStyle.Alignment = HSSFCellStyle.ALIGN_CENTER;
            BottomMediumLeftBorderCellStyle.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
            BottomMediumLeftBorderCellStyle.SetFont(PackNumberFont);

            HSSFCellStyle BottomMediumRightBorderCellStyle = hssfworkbook.CreateCellStyle();
            BottomMediumRightBorderCellStyle.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            BottomMediumRightBorderCellStyle.BottomBorderColor = HSSFColor.BLACK.index;
            BottomMediumRightBorderCellStyle.BorderLeft = HSSFCellStyle.BORDER_THIN;
            BottomMediumRightBorderCellStyle.LeftBorderColor = HSSFColor.BLACK.index;
            BottomMediumRightBorderCellStyle.BorderRight = HSSFCellStyle.BORDER_MEDIUM;
            BottomMediumRightBorderCellStyle.RightBorderColor = HSSFColor.BLACK.index;
            BottomMediumRightBorderCellStyle.BorderTop = HSSFCellStyle.BORDER_THIN;
            BottomMediumRightBorderCellStyle.TopBorderColor = HSSFColor.BLACK.index;
            BottomMediumRightBorderCellStyle.SetFont(SimpleFont);

            HSSFCellStyle LeftMediumBorderCellStyle = hssfworkbook.CreateCellStyle();
            LeftMediumBorderCellStyle.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            LeftMediumBorderCellStyle.BottomBorderColor = HSSFColor.BLACK.index;
            LeftMediumBorderCellStyle.BorderLeft = HSSFCellStyle.BORDER_MEDIUM;
            LeftMediumBorderCellStyle.LeftBorderColor = HSSFColor.BLACK.index;
            LeftMediumBorderCellStyle.BorderRight = HSSFCellStyle.BORDER_THIN;
            LeftMediumBorderCellStyle.RightBorderColor = HSSFColor.BLACK.index;
            LeftMediumBorderCellStyle.BorderTop = HSSFCellStyle.BORDER_THIN;
            LeftMediumBorderCellStyle.TopBorderColor = HSSFColor.BLACK.index;
            LeftMediumBorderCellStyle.Alignment = HSSFCellStyle.ALIGN_CENTER;
            LeftMediumBorderCellStyle.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
            LeftMediumBorderCellStyle.SetFont(PackNumberFont);

            HSSFCellStyle RightMediumBorderCellStyle = hssfworkbook.CreateCellStyle();
            RightMediumBorderCellStyle.BorderBottom = HSSFCellStyle.BORDER_THIN;
            RightMediumBorderCellStyle.BottomBorderColor = HSSFColor.BLACK.index;
            RightMediumBorderCellStyle.BorderLeft = HSSFCellStyle.BORDER_THIN;
            RightMediumBorderCellStyle.LeftBorderColor = HSSFColor.BLACK.index;
            RightMediumBorderCellStyle.BorderRight = HSSFCellStyle.BORDER_MEDIUM;
            RightMediumBorderCellStyle.RightBorderColor = HSSFColor.BLACK.index;
            RightMediumBorderCellStyle.BorderTop = HSSFCellStyle.BORDER_THIN;
            RightMediumBorderCellStyle.TopBorderColor = HSSFColor.BLACK.index;
            RightMediumBorderCellStyle.SetFont(SimpleFont);

            HSSFCellStyle GreyBottomMediumBorderCellStyle = hssfworkbook.CreateCellStyle();
            GreyBottomMediumBorderCellStyle.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            GreyBottomMediumBorderCellStyle.BottomBorderColor = HSSFColor.BLACK.index;
            GreyBottomMediumBorderCellStyle.BorderLeft = HSSFCellStyle.BORDER_THIN;
            GreyBottomMediumBorderCellStyle.LeftBorderColor = HSSFColor.BLACK.index;
            GreyBottomMediumBorderCellStyle.BorderRight = HSSFCellStyle.BORDER_THIN;
            GreyBottomMediumBorderCellStyle.RightBorderColor = HSSFColor.BLACK.index;
            GreyBottomMediumBorderCellStyle.BorderTop = HSSFCellStyle.BORDER_THIN;
            GreyBottomMediumBorderCellStyle.TopBorderColor = HSSFColor.BLACK.index;
            GreyBottomMediumBorderCellStyle.FillForegroundColor = HSSFColor.GREY_25_PERCENT.index;
            GreyBottomMediumBorderCellStyle.FillPattern = HSSFCellStyle.SOLID_FOREGROUND;
            GreyBottomMediumBorderCellStyle.FillBackgroundColor = HSSFColor.YELLOW.index;
            GreyBottomMediumBorderCellStyle.SetFont(SimpleFont);

            HSSFCellStyle GreyBottomMediumLeftBorderCellStyle = hssfworkbook.CreateCellStyle();
            GreyBottomMediumLeftBorderCellStyle.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            GreyBottomMediumLeftBorderCellStyle.BottomBorderColor = HSSFColor.BLACK.index;
            GreyBottomMediumLeftBorderCellStyle.BorderLeft = HSSFCellStyle.BORDER_MEDIUM;
            GreyBottomMediumLeftBorderCellStyle.LeftBorderColor = HSSFColor.BLACK.index;
            GreyBottomMediumLeftBorderCellStyle.BorderRight = HSSFCellStyle.BORDER_THIN;
            GreyBottomMediumLeftBorderCellStyle.RightBorderColor = HSSFColor.BLACK.index;
            GreyBottomMediumLeftBorderCellStyle.BorderTop = HSSFCellStyle.BORDER_THIN;
            GreyBottomMediumLeftBorderCellStyle.TopBorderColor = HSSFColor.BLACK.index;
            GreyBottomMediumLeftBorderCellStyle.Alignment = HSSFCellStyle.ALIGN_CENTER;
            GreyBottomMediumLeftBorderCellStyle.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
            GreyBottomMediumLeftBorderCellStyle.SetFont(PackNumberFont);

            HSSFCellStyle GreyBottomMediumRightBorderCellStyle = hssfworkbook.CreateCellStyle();
            GreyBottomMediumRightBorderCellStyle.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            GreyBottomMediumRightBorderCellStyle.BottomBorderColor = HSSFColor.BLACK.index;
            GreyBottomMediumRightBorderCellStyle.BorderLeft = HSSFCellStyle.BORDER_THIN;
            GreyBottomMediumRightBorderCellStyle.LeftBorderColor = HSSFColor.BLACK.index;
            GreyBottomMediumRightBorderCellStyle.BorderRight = HSSFCellStyle.BORDER_MEDIUM;
            GreyBottomMediumRightBorderCellStyle.RightBorderColor = HSSFColor.BLACK.index;
            GreyBottomMediumRightBorderCellStyle.BorderTop = HSSFCellStyle.BORDER_THIN;
            GreyBottomMediumRightBorderCellStyle.TopBorderColor = HSSFColor.BLACK.index;
            GreyBottomMediumRightBorderCellStyle.FillForegroundColor = HSSFColor.GREY_25_PERCENT.index;
            GreyBottomMediumRightBorderCellStyle.FillPattern = HSSFCellStyle.SOLID_FOREGROUND;
            GreyBottomMediumRightBorderCellStyle.FillBackgroundColor = HSSFColor.YELLOW.index;
            GreyBottomMediumRightBorderCellStyle.SetFont(SimpleFont);

            HSSFCellStyle GreyLeftMediumBorderCellStyle = hssfworkbook.CreateCellStyle();
            GreyLeftMediumBorderCellStyle.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            GreyLeftMediumBorderCellStyle.BottomBorderColor = HSSFColor.BLACK.index;
            GreyLeftMediumBorderCellStyle.BorderLeft = HSSFCellStyle.BORDER_MEDIUM;
            GreyLeftMediumBorderCellStyle.LeftBorderColor = HSSFColor.BLACK.index;
            GreyLeftMediumBorderCellStyle.BorderRight = HSSFCellStyle.BORDER_THIN;
            GreyLeftMediumBorderCellStyle.RightBorderColor = HSSFColor.BLACK.index;
            GreyLeftMediumBorderCellStyle.BorderTop = HSSFCellStyle.BORDER_THIN;
            GreyLeftMediumBorderCellStyle.TopBorderColor = HSSFColor.BLACK.index;
            GreyLeftMediumBorderCellStyle.Alignment = HSSFCellStyle.ALIGN_CENTER;
            GreyLeftMediumBorderCellStyle.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
            GreyLeftMediumBorderCellStyle.SetFont(PackNumberFont);

            HSSFCellStyle GreyRightMediumBorderCellStyle = hssfworkbook.CreateCellStyle();
            GreyRightMediumBorderCellStyle.BorderBottom = HSSFCellStyle.BORDER_THIN;
            GreyRightMediumBorderCellStyle.BottomBorderColor = HSSFColor.BLACK.index;
            GreyRightMediumBorderCellStyle.BorderLeft = HSSFCellStyle.BORDER_THIN;
            GreyRightMediumBorderCellStyle.LeftBorderColor = HSSFColor.BLACK.index;
            GreyRightMediumBorderCellStyle.BorderRight = HSSFCellStyle.BORDER_MEDIUM;
            GreyRightMediumBorderCellStyle.RightBorderColor = HSSFColor.BLACK.index;
            GreyRightMediumBorderCellStyle.BorderTop = HSSFCellStyle.BORDER_THIN;
            GreyRightMediumBorderCellStyle.TopBorderColor = HSSFColor.BLACK.index;
            GreyRightMediumBorderCellStyle.FillForegroundColor = HSSFColor.GREY_25_PERCENT.index;
            GreyRightMediumBorderCellStyle.FillPattern = HSSFCellStyle.SOLID_FOREGROUND;
            GreyRightMediumBorderCellStyle.FillBackgroundColor = HSSFColor.YELLOW.index;
            GreyRightMediumBorderCellStyle.SetFont(SimpleFont);

            #endregion границы между упаковками

            //HSSFCell ConfirmCell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, s);
            //ConfirmCell.CellStyle = TempStyle;

            if (FrontsOrdersDataTable.Rows.Count > 0)
                FrontsReportOrderNumber(hssfworkbook, HeaderStyle, SimpleFont, SimpleCellStyle, cellStyle, FrontsOrdersDataTable, ZOV);

            if (DecorOrdersDataTable.Rows.Count > 0)
                DecorReportOrderNumber(hssfworkbook, HeaderStyle, SimpleFont, SimpleCellStyle, cellStyle, DecorOrdersDataTable, ZOV);

            string tempFolder = Environment.GetEnvironmentVariable("TEMP");
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
        
        public void CreateReport(DateTime DateFrom, DateTime DateTo, int FactoryID,
            DataTable FrontsOrdersDataTable, DataTable DecorOrdersDataTable, string FileName, bool ZOV)
        {
            if (FrontsOrdersDataTable.Rows.Count < 1 && DecorOrdersDataTable.Rows.Count < 1)
                return;

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

            HSSFFont ClientNameFont = hssfworkbook.CreateFont();
            ClientNameFont.FontHeightInPoints = 14;
            ClientNameFont.Boldweight = HSSFFont.BOLDWEIGHT_BOLD;
            ClientNameFont.FontName = "Calibri";

            HSSFFont MainFont = hssfworkbook.CreateFont();
            MainFont.FontHeightInPoints = 12;
            MainFont.Boldweight = HSSFFont.BOLDWEIGHT_BOLD;
            MainFont.FontName = "Calibri";

            HSSFCellStyle ClientNameStyle = hssfworkbook.CreateCellStyle();
            ClientNameStyle.SetFont(ClientNameFont);

            HSSFCellStyle MainStyle = hssfworkbook.CreateCellStyle();
            MainStyle.SetFont(MainFont);

            HSSFFont HeaderFont = hssfworkbook.CreateFont();
            HeaderFont.FontHeightInPoints = 12;
            HeaderFont.Boldweight = HSSFFont.BOLDWEIGHT_BOLD;
            HeaderFont.FontName = "Calibri";

            HSSFCellStyle HeaderStyle = hssfworkbook.CreateCellStyle();
            HeaderStyle.BorderBottom = HSSFCellStyle.BORDER_THIN;
            HeaderStyle.BottomBorderColor = HSSFColor.BLACK.index;
            HeaderStyle.BorderLeft = HSSFCellStyle.BORDER_THIN;
            HeaderStyle.LeftBorderColor = HSSFColor.BLACK.index;
            HeaderStyle.BorderRight = HSSFCellStyle.BORDER_THIN;
            HeaderStyle.RightBorderColor = HSSFColor.BLACK.index;
            HeaderStyle.BorderTop = HSSFCellStyle.BORDER_THIN;
            HeaderStyle.TopBorderColor = HSSFColor.BLACK.index;
            HeaderStyle.SetFont(HeaderFont);

            HSSFFont PackNumberFont = hssfworkbook.CreateFont();
            PackNumberFont.Boldweight = 11 * 256;
            PackNumberFont.FontName = "Calibri";

            HSSFCellStyle PackNumberStyle = hssfworkbook.CreateCellStyle();
            PackNumberStyle.BorderBottom = HSSFCellStyle.BORDER_THIN;
            PackNumberStyle.BottomBorderColor = HSSFColor.BLACK.index;
            PackNumberStyle.BorderLeft = HSSFCellStyle.BORDER_THIN;
            PackNumberStyle.LeftBorderColor = HSSFColor.BLACK.index;
            PackNumberStyle.BorderRight = HSSFCellStyle.BORDER_THIN;
            PackNumberStyle.RightBorderColor = HSSFColor.BLACK.index;
            PackNumberStyle.BorderTop = HSSFCellStyle.BORDER_THIN;
            PackNumberStyle.TopBorderColor = HSSFColor.BLACK.index;
            PackNumberStyle.Alignment = HSSFCellStyle.ALIGN_CENTER;
            PackNumberStyle.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
            PackNumberStyle.SetFont(PackNumberFont);

            HSSFFont SimpleFont = hssfworkbook.CreateFont();
            SimpleFont.FontHeightInPoints = 11;
            SimpleFont.FontName = "Calibri";

            HSSFCellStyle SimpleCellStyle = hssfworkbook.CreateCellStyle();
            SimpleCellStyle.BorderBottom = HSSFCellStyle.BORDER_THIN;
            SimpleCellStyle.BottomBorderColor = HSSFColor.BLACK.index;
            SimpleCellStyle.BorderLeft = HSSFCellStyle.BORDER_THIN;
            SimpleCellStyle.LeftBorderColor = HSSFColor.BLACK.index;
            SimpleCellStyle.BorderRight = HSSFCellStyle.BORDER_THIN;
            SimpleCellStyle.RightBorderColor = HSSFColor.BLACK.index;
            SimpleCellStyle.BorderTop = HSSFCellStyle.BORDER_THIN;
            SimpleCellStyle.TopBorderColor = HSSFColor.BLACK.index;
            SimpleCellStyle.SetFont(SimpleFont);

            HSSFCellStyle cellStyle = hssfworkbook.CreateCellStyle();
            cellStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.000");
            cellStyle.BorderBottom = HSSFCellStyle.BORDER_THIN;
            cellStyle.BottomBorderColor = HSSFColor.BLACK.index;
            cellStyle.BorderLeft = HSSFCellStyle.BORDER_THIN;
            cellStyle.LeftBorderColor = HSSFColor.BLACK.index;
            cellStyle.BorderRight = HSSFCellStyle.BORDER_THIN;
            cellStyle.RightBorderColor = HSSFColor.BLACK.index;
            cellStyle.BorderTop = HSSFCellStyle.BORDER_THIN;
            cellStyle.TopBorderColor = HSSFColor.BLACK.index;
            cellStyle.SetFont(SimpleFont);

            HSSFCellStyle GreyCellStyle = hssfworkbook.CreateCellStyle();
            GreyCellStyle.BorderBottom = HSSFCellStyle.BORDER_THIN;
            GreyCellStyle.BottomBorderColor = HSSFColor.BLACK.index;
            GreyCellStyle.BorderLeft = HSSFCellStyle.BORDER_THIN;
            GreyCellStyle.LeftBorderColor = HSSFColor.BLACK.index;
            GreyCellStyle.BorderRight = HSSFCellStyle.BORDER_THIN;
            GreyCellStyle.RightBorderColor = HSSFColor.BLACK.index;
            GreyCellStyle.BorderTop = HSSFCellStyle.BORDER_THIN;
            GreyCellStyle.TopBorderColor = HSSFColor.BLACK.index;
            GreyCellStyle.FillForegroundColor = HSSFColor.GREY_25_PERCENT.index;
            GreyCellStyle.FillPattern = HSSFCellStyle.SOLID_FOREGROUND;
            GreyCellStyle.FillBackgroundColor = HSSFColor.YELLOW.index;
            GreyCellStyle.SetFont(SimpleFont);

            HSSFFont TotalFont = hssfworkbook.CreateFont();
            TotalFont.FontHeightInPoints = 11;
            TotalFont.FontName = "Calibri";

            HSSFCellStyle TotalStyle = hssfworkbook.CreateCellStyle();
            TotalStyle.BorderTop = HSSFCellStyle.BORDER_MEDIUM;
            TotalStyle.TopBorderColor = HSSFColor.BLACK.index;
            TotalStyle.SetFont(TotalFont);

            HSSFCellStyle TempStyle = hssfworkbook.CreateCellStyle();
            TempStyle.SetFont(TotalFont);

            #endregion Create fonts and styles

            #region границы между упаковками

            HSSFCellStyle BottomMediumBorderCellStyle = hssfworkbook.CreateCellStyle();
            BottomMediumBorderCellStyle.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            BottomMediumBorderCellStyle.BottomBorderColor = HSSFColor.BLACK.index;
            BottomMediumBorderCellStyle.BorderLeft = HSSFCellStyle.BORDER_THIN;
            BottomMediumBorderCellStyle.LeftBorderColor = HSSFColor.BLACK.index;
            BottomMediumBorderCellStyle.BorderRight = HSSFCellStyle.BORDER_THIN;
            BottomMediumBorderCellStyle.RightBorderColor = HSSFColor.BLACK.index;
            BottomMediumBorderCellStyle.BorderTop = HSSFCellStyle.BORDER_THIN;
            BottomMediumBorderCellStyle.TopBorderColor = HSSFColor.BLACK.index;
            BottomMediumBorderCellStyle.SetFont(SimpleFont);

            HSSFCellStyle BottomMediumLeftBorderCellStyle = hssfworkbook.CreateCellStyle();
            BottomMediumLeftBorderCellStyle.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            BottomMediumLeftBorderCellStyle.BottomBorderColor = HSSFColor.BLACK.index;
            BottomMediumLeftBorderCellStyle.BorderLeft = HSSFCellStyle.BORDER_MEDIUM;
            BottomMediumLeftBorderCellStyle.LeftBorderColor = HSSFColor.BLACK.index;
            BottomMediumLeftBorderCellStyle.BorderRight = HSSFCellStyle.BORDER_THIN;
            BottomMediumLeftBorderCellStyle.RightBorderColor = HSSFColor.BLACK.index;
            BottomMediumLeftBorderCellStyle.BorderTop = HSSFCellStyle.BORDER_THIN;
            BottomMediumLeftBorderCellStyle.TopBorderColor = HSSFColor.BLACK.index;
            BottomMediumLeftBorderCellStyle.Alignment = HSSFCellStyle.ALIGN_CENTER;
            BottomMediumLeftBorderCellStyle.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
            BottomMediumLeftBorderCellStyle.SetFont(PackNumberFont);

            HSSFCellStyle BottomMediumRightBorderCellStyle = hssfworkbook.CreateCellStyle();
            BottomMediumRightBorderCellStyle.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            BottomMediumRightBorderCellStyle.BottomBorderColor = HSSFColor.BLACK.index;
            BottomMediumRightBorderCellStyle.BorderLeft = HSSFCellStyle.BORDER_THIN;
            BottomMediumRightBorderCellStyle.LeftBorderColor = HSSFColor.BLACK.index;
            BottomMediumRightBorderCellStyle.BorderRight = HSSFCellStyle.BORDER_MEDIUM;
            BottomMediumRightBorderCellStyle.RightBorderColor = HSSFColor.BLACK.index;
            BottomMediumRightBorderCellStyle.BorderTop = HSSFCellStyle.BORDER_THIN;
            BottomMediumRightBorderCellStyle.TopBorderColor = HSSFColor.BLACK.index;
            BottomMediumRightBorderCellStyle.SetFont(SimpleFont);

            HSSFCellStyle LeftMediumBorderCellStyle = hssfworkbook.CreateCellStyle();
            LeftMediumBorderCellStyle.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            LeftMediumBorderCellStyle.BottomBorderColor = HSSFColor.BLACK.index;
            LeftMediumBorderCellStyle.BorderLeft = HSSFCellStyle.BORDER_MEDIUM;
            LeftMediumBorderCellStyle.LeftBorderColor = HSSFColor.BLACK.index;
            LeftMediumBorderCellStyle.BorderRight = HSSFCellStyle.BORDER_THIN;
            LeftMediumBorderCellStyle.RightBorderColor = HSSFColor.BLACK.index;
            LeftMediumBorderCellStyle.BorderTop = HSSFCellStyle.BORDER_THIN;
            LeftMediumBorderCellStyle.TopBorderColor = HSSFColor.BLACK.index;
            LeftMediumBorderCellStyle.Alignment = HSSFCellStyle.ALIGN_CENTER;
            LeftMediumBorderCellStyle.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
            LeftMediumBorderCellStyle.SetFont(PackNumberFont);

            HSSFCellStyle RightMediumBorderCellStyle = hssfworkbook.CreateCellStyle();
            RightMediumBorderCellStyle.BorderBottom = HSSFCellStyle.BORDER_THIN;
            RightMediumBorderCellStyle.BottomBorderColor = HSSFColor.BLACK.index;
            RightMediumBorderCellStyle.BorderLeft = HSSFCellStyle.BORDER_THIN;
            RightMediumBorderCellStyle.LeftBorderColor = HSSFColor.BLACK.index;
            RightMediumBorderCellStyle.BorderRight = HSSFCellStyle.BORDER_MEDIUM;
            RightMediumBorderCellStyle.RightBorderColor = HSSFColor.BLACK.index;
            RightMediumBorderCellStyle.BorderTop = HSSFCellStyle.BORDER_THIN;
            RightMediumBorderCellStyle.TopBorderColor = HSSFColor.BLACK.index;
            RightMediumBorderCellStyle.SetFont(SimpleFont);

            HSSFCellStyle GreyBottomMediumBorderCellStyle = hssfworkbook.CreateCellStyle();
            GreyBottomMediumBorderCellStyle.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            GreyBottomMediumBorderCellStyle.BottomBorderColor = HSSFColor.BLACK.index;
            GreyBottomMediumBorderCellStyle.BorderLeft = HSSFCellStyle.BORDER_THIN;
            GreyBottomMediumBorderCellStyle.LeftBorderColor = HSSFColor.BLACK.index;
            GreyBottomMediumBorderCellStyle.BorderRight = HSSFCellStyle.BORDER_THIN;
            GreyBottomMediumBorderCellStyle.RightBorderColor = HSSFColor.BLACK.index;
            GreyBottomMediumBorderCellStyle.BorderTop = HSSFCellStyle.BORDER_THIN;
            GreyBottomMediumBorderCellStyle.TopBorderColor = HSSFColor.BLACK.index;
            GreyBottomMediumBorderCellStyle.FillForegroundColor = HSSFColor.GREY_25_PERCENT.index;
            GreyBottomMediumBorderCellStyle.FillPattern = HSSFCellStyle.SOLID_FOREGROUND;
            GreyBottomMediumBorderCellStyle.FillBackgroundColor = HSSFColor.YELLOW.index;
            GreyBottomMediumBorderCellStyle.SetFont(SimpleFont);

            HSSFCellStyle GreyBottomMediumLeftBorderCellStyle = hssfworkbook.CreateCellStyle();
            GreyBottomMediumLeftBorderCellStyle.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            GreyBottomMediumLeftBorderCellStyle.BottomBorderColor = HSSFColor.BLACK.index;
            GreyBottomMediumLeftBorderCellStyle.BorderLeft = HSSFCellStyle.BORDER_MEDIUM;
            GreyBottomMediumLeftBorderCellStyle.LeftBorderColor = HSSFColor.BLACK.index;
            GreyBottomMediumLeftBorderCellStyle.BorderRight = HSSFCellStyle.BORDER_THIN;
            GreyBottomMediumLeftBorderCellStyle.RightBorderColor = HSSFColor.BLACK.index;
            GreyBottomMediumLeftBorderCellStyle.BorderTop = HSSFCellStyle.BORDER_THIN;
            GreyBottomMediumLeftBorderCellStyle.TopBorderColor = HSSFColor.BLACK.index;
            GreyBottomMediumLeftBorderCellStyle.Alignment = HSSFCellStyle.ALIGN_CENTER;
            GreyBottomMediumLeftBorderCellStyle.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
            GreyBottomMediumLeftBorderCellStyle.SetFont(PackNumberFont);

            HSSFCellStyle GreyBottomMediumRightBorderCellStyle = hssfworkbook.CreateCellStyle();
            GreyBottomMediumRightBorderCellStyle.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            GreyBottomMediumRightBorderCellStyle.BottomBorderColor = HSSFColor.BLACK.index;
            GreyBottomMediumRightBorderCellStyle.BorderLeft = HSSFCellStyle.BORDER_THIN;
            GreyBottomMediumRightBorderCellStyle.LeftBorderColor = HSSFColor.BLACK.index;
            GreyBottomMediumRightBorderCellStyle.BorderRight = HSSFCellStyle.BORDER_MEDIUM;
            GreyBottomMediumRightBorderCellStyle.RightBorderColor = HSSFColor.BLACK.index;
            GreyBottomMediumRightBorderCellStyle.BorderTop = HSSFCellStyle.BORDER_THIN;
            GreyBottomMediumRightBorderCellStyle.TopBorderColor = HSSFColor.BLACK.index;
            GreyBottomMediumRightBorderCellStyle.FillForegroundColor = HSSFColor.GREY_25_PERCENT.index;
            GreyBottomMediumRightBorderCellStyle.FillPattern = HSSFCellStyle.SOLID_FOREGROUND;
            GreyBottomMediumRightBorderCellStyle.FillBackgroundColor = HSSFColor.YELLOW.index;
            GreyBottomMediumRightBorderCellStyle.SetFont(SimpleFont);

            HSSFCellStyle GreyLeftMediumBorderCellStyle = hssfworkbook.CreateCellStyle();
            GreyLeftMediumBorderCellStyle.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            GreyLeftMediumBorderCellStyle.BottomBorderColor = HSSFColor.BLACK.index;
            GreyLeftMediumBorderCellStyle.BorderLeft = HSSFCellStyle.BORDER_MEDIUM;
            GreyLeftMediumBorderCellStyle.LeftBorderColor = HSSFColor.BLACK.index;
            GreyLeftMediumBorderCellStyle.BorderRight = HSSFCellStyle.BORDER_THIN;
            GreyLeftMediumBorderCellStyle.RightBorderColor = HSSFColor.BLACK.index;
            GreyLeftMediumBorderCellStyle.BorderTop = HSSFCellStyle.BORDER_THIN;
            GreyLeftMediumBorderCellStyle.TopBorderColor = HSSFColor.BLACK.index;
            GreyLeftMediumBorderCellStyle.Alignment = HSSFCellStyle.ALIGN_CENTER;
            GreyLeftMediumBorderCellStyle.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
            GreyLeftMediumBorderCellStyle.SetFont(PackNumberFont);

            HSSFCellStyle GreyRightMediumBorderCellStyle = hssfworkbook.CreateCellStyle();
            GreyRightMediumBorderCellStyle.BorderBottom = HSSFCellStyle.BORDER_THIN;
            GreyRightMediumBorderCellStyle.BottomBorderColor = HSSFColor.BLACK.index;
            GreyRightMediumBorderCellStyle.BorderLeft = HSSFCellStyle.BORDER_THIN;
            GreyRightMediumBorderCellStyle.LeftBorderColor = HSSFColor.BLACK.index;
            GreyRightMediumBorderCellStyle.BorderRight = HSSFCellStyle.BORDER_MEDIUM;
            GreyRightMediumBorderCellStyle.RightBorderColor = HSSFColor.BLACK.index;
            GreyRightMediumBorderCellStyle.BorderTop = HSSFCellStyle.BORDER_THIN;
            GreyRightMediumBorderCellStyle.TopBorderColor = HSSFColor.BLACK.index;
            GreyRightMediumBorderCellStyle.FillForegroundColor = HSSFColor.GREY_25_PERCENT.index;
            GreyRightMediumBorderCellStyle.FillPattern = HSSFCellStyle.SOLID_FOREGROUND;
            GreyRightMediumBorderCellStyle.FillBackgroundColor = HSSFColor.YELLOW.index;
            GreyRightMediumBorderCellStyle.SetFont(SimpleFont);

            #endregion границы между упаковками

            //HSSFCell ConfirmCell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, s);
            //ConfirmCell.CellStyle = TempStyle;

            if (FrontsOrdersDataTable.Rows.Count > 0)
                FrontsReport(hssfworkbook, HeaderStyle, SimpleFont, SimpleCellStyle, cellStyle, FrontsOrdersDataTable, ZOV);

            if (DecorOrdersDataTable.Rows.Count > 0)
                DecorReport(hssfworkbook, HeaderStyle, SimpleFont, SimpleCellStyle, cellStyle, DecorOrdersDataTable, ZOV);

            string tempFolder = Environment.GetEnvironmentVariable("TEMP");
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

        public string GetColorName(int ColorID)
        {
            string ColorName = string.Empty;
            DataRow[] Rows = FrameColorsDataTable.Select("ColorID = " + ColorID);
            if (Rows.Count() > 0)
                ColorName = Rows[0]["ColorName"].ToString();
            return ColorName;
        }

        public string GetFront2Name(int TechnoProfileID)
        {
            string FrontName = "";
            try
            {
                DataRow[] Rows = FrontsDataTable.Select("FrontID = " + TechnoProfileID);
                FrontName = Rows[0]["FrontName"].ToString();
            }
            catch
            {
                return "";
            }
            return FrontName;
        }

        public string GetFrontName(int FrontID)
        {
            string FrontName = string.Empty;
            DataRow[] Rows = FrontsDataTable.Select("FrontID = " + FrontID);
            if (Rows.Count() > 0)
                FrontName = Rows[0]["FrontName"].ToString();
            else
                FrontName = string.Empty;
            return FrontName;
        }

        public string GetInsetColorName(int InsetColorID)
        {
            string ColorName = string.Empty;
            DataRow[] Rows = InsetColorsDataTable.Select("InsetColorID = " + InsetColorID);
            if (Rows.Count() > 0)
                ColorName = Rows[0]["InsetColorName"].ToString();
            return ColorName;
        }

        public string GetInsetTypeName(int InsetTypeID)
        {
            string InsetType = string.Empty;
            DataRow[] Rows = InsetTypesDataTable.Select("InsetTypeID = " + InsetTypeID);
            if (Rows.Count() > 0)
                InsetType = Rows[0]["InsetTypeName"].ToString();
            return InsetType;
        }

        public string GetPatinaName(int PatinaID)
        {
            string PatinaName = string.Empty;
            DataRow[] Rows = PatinaDataTable.Select("PatinaID = " + PatinaID);
            if (Rows.Count() > 0)
                PatinaName = Rows[0]["PatinaName"].ToString();
            return PatinaName;
        }

        private void Create()
        {
            FrontsDataTable = new DataTable();
            FrameColorsDataTable = new DataTable();
            PatinaDataTable = new DataTable();
            InsetTypesDataTable = new DataTable();
            InsetColorsDataTable = new DataTable();
            FrontsResultDataTable = new DataTable();
            DecorResultDataTable = new DataTable[DecorCatalog.DecorProductsCount];
        }

        private void CreateDecorDataTable()
        {
            for (int i = 0; i < DecorCatalog.DecorProductsCount; i++)
            {
                DecorResultDataTable[i] = new DataTable();

                DecorResultDataTable[i].Columns.Add(new DataColumn("ClientGroupName", Type.GetType("System.String")));
                DecorResultDataTable[i].Columns.Add(new DataColumn("Client", Type.GetType("System.String")));
                DecorResultDataTable[i].Columns.Add(new DataColumn("OrderNumber", Type.GetType("System.Int32")));
                DecorResultDataTable[i].Columns.Add(new DataColumn("DocNumber", Type.GetType("System.String")));
                DecorResultDataTable[i].Columns.Add(new DataColumn("DispatchDateTime", Type.GetType("System.String")));
                DecorResultDataTable[i].Columns.Add(new DataColumn("ZOVClientName", Type.GetType("System.String")));
                DecorResultDataTable[i].Columns.Add(new DataColumn("Product", Type.GetType("System.String")));
                DecorResultDataTable[i].Columns.Add(new DataColumn("Color", Type.GetType("System.String")));
                DecorResultDataTable[i].Columns.Add(new DataColumn("Patina", Type.GetType("System.String")));
                DecorResultDataTable[i].Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
                DecorResultDataTable[i].Columns.Add(new DataColumn("Width", Type.GetType("System.Int32")));
                DecorResultDataTable[i].Columns.Add(new DataColumn("Count", Type.GetType("System.Int32")));
                DecorResultDataTable[i].Columns.Add(new DataColumn("IsSample", Type.GetType("System.Int32")));
                DecorResultDataTable[i].Columns.Add(new DataColumn("AvgPrice", Type.GetType("System.Decimal")));
                DecorResultDataTable[i].Columns.Add(new DataColumn("SumCost", Type.GetType("System.Decimal")));
                DecorResultDataTable[i].Columns.Add(new DataColumn("ManagerName", Type.GetType("System.String")));
                DecorResultDataTable[i].Columns.Add(new DataColumn("CountryName", Type.GetType("System.String")));
            }
        }

        private void CreateFrontsDataTable()
        {
            FrontsResultDataTable = new DataTable();

            //FrontsResultDataTable.Columns.Add(new DataColumn(("FrontID"), System.Type.GetType("System.Int32")));
            //FrontsResultDataTable.Columns.Add(new DataColumn(("FrontTypeID"), System.Type.GetType("System.Int32")));
            //FrontsResultDataTable.Columns.Add(new DataColumn(("ColorID"), System.Type.GetType("System.Int32")));
            //FrontsResultDataTable.Columns.Add(new DataColumn(("InsetTypeID"), System.Type.GetType("System.Int32")));
            //FrontsResultDataTable.Columns.Add(new DataColumn(("InsetColorID"), System.Type.GetType("System.Int32")));
            //FrontsResultDataTable.Columns.Add(new DataColumn(("ClientID"), System.Type.GetType("System.Int32")));

            FrontsResultDataTable.Columns.Add(new DataColumn("ClientGroupName", Type.GetType("System.String")));
            FrontsResultDataTable.Columns.Add(new DataColumn("Client", Type.GetType("System.String")));
            FrontsResultDataTable.Columns.Add(new DataColumn("OrderNumber", Type.GetType("System.Int32")));
            FrontsResultDataTable.Columns.Add(new DataColumn("DocNumber", Type.GetType("System.String")));
            FrontsResultDataTable.Columns.Add(new DataColumn("DispatchDateTime", Type.GetType("System.String")));
            FrontsResultDataTable.Columns.Add(new DataColumn("ZOVClientName", Type.GetType("System.String")));
            FrontsResultDataTable.Columns.Add(new DataColumn("Front", Type.GetType("System.String")));
            FrontsResultDataTable.Columns.Add(new DataColumn("FrameColor", Type.GetType("System.String")));
            FrontsResultDataTable.Columns.Add(new DataColumn("TechnoColor", Type.GetType("System.String")));
            FrontsResultDataTable.Columns.Add(new DataColumn("Patina", Type.GetType("System.String")));
            FrontsResultDataTable.Columns.Add(new DataColumn("InsetType", Type.GetType("System.String")));
            FrontsResultDataTable.Columns.Add(new DataColumn("InsetColor", Type.GetType("System.String")));
            FrontsResultDataTable.Columns.Add(new DataColumn("TechnoInsetType", Type.GetType("System.String")));
            FrontsResultDataTable.Columns.Add(new DataColumn("TechnoInsetColor", Type.GetType("System.String")));
            FrontsResultDataTable.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            FrontsResultDataTable.Columns.Add(new DataColumn("Width", Type.GetType("System.Int32")));
            FrontsResultDataTable.Columns.Add(new DataColumn("Square", Type.GetType("System.Decimal")));
            FrontsResultDataTable.Columns.Add(new DataColumn("Count", Type.GetType("System.Int32")));
            FrontsResultDataTable.Columns.Add(new DataColumn("IsSample", Type.GetType("System.Int32")));
            FrontsResultDataTable.Columns.Add(new DataColumn("AvgPrice", Type.GetType("System.Decimal")));
            FrontsResultDataTable.Columns.Add(new DataColumn("SumCost", Type.GetType("System.Decimal")));
            FrontsResultDataTable.Columns.Add(new DataColumn("ManagerName", Type.GetType("System.String")));
            FrontsResultDataTable.Columns.Add(new DataColumn("CountryName", Type.GetType("System.String")));
        }

        private void DecorReportOrderNumber(HSSFWorkbook hssfworkbook, HSSFCellStyle HeaderStyle,
            HSSFFont SimpleFont, HSSFCellStyle SimpleCellStyle, HSSFCellStyle cellStyle,
            DataTable DecorOrdersDataTable, bool ZOV)
        {
            int RowIndex = 0;

            int DisplayIndex = 0;
            HSSFSheet sheet1 = hssfworkbook.CreateSheet("Декор");
            sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

            sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
            sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
            sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
            sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

            if (ZOV)
            {
                sheet1.SetColumnWidth(DisplayIndex++, 25 * 256);
                sheet1.SetColumnWidth(DisplayIndex++, 20 * 256);
                sheet1.SetColumnWidth(DisplayIndex++, 20 * 256);
                sheet1.SetColumnWidth(DisplayIndex++, 15 * 256);
            }
            else
            {
                sheet1.SetColumnWidth(DisplayIndex++, 35 * 256);
                sheet1.SetColumnWidth(DisplayIndex++, 20 * 256);
                sheet1.SetColumnWidth(DisplayIndex++, 20 * 256);
            }
            sheet1.SetColumnWidth(DisplayIndex++, 30 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 20 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 20 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 10 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 10 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 10 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 10 * 256);
            if (Security.PriceAccess)
            {
                sheet1.SetColumnWidth(DisplayIndex++, 17 * 256);
                sheet1.SetColumnWidth(DisplayIndex++, 15 * 256);
            }
            sheet1.SetColumnWidth(DisplayIndex++, 20 * 256);
            //декор

            FillDecorOrderNumber(DecorOrdersDataTable, ZOV);

            for (int c = 0; c < DecorCatalog.DecorProductsCount; c++)
            {
                if (DecorResultDataTable[c].Rows.Count == 0)
                    continue;
                DisplayIndex = 0;

                HSSFCell cell;
                if (ZOV)
                {
                    cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Группа");
                    cell.CellStyle = HeaderStyle;
                    cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Клиент");
                    cell.CellStyle = HeaderStyle;
                    cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "№ заказа");
                    cell.CellStyle = HeaderStyle;
                    cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Дата отгр");
                    cell.CellStyle = HeaderStyle;
                }
                else
                {
                    cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Клиент");
                    cell.CellStyle = HeaderStyle;
                    cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "№ заказа");
                    cell.CellStyle = HeaderStyle;
                    cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Клиент ЗОВ");
                    cell.CellStyle = HeaderStyle;
                }
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Название");
                cell.CellStyle = HeaderStyle;
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Цвет");
                cell.CellStyle = HeaderStyle;
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Патина");
                cell.CellStyle = HeaderStyle;
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Длин\\Выс.");
                cell.CellStyle = HeaderStyle;
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Шир.");
                cell.CellStyle = HeaderStyle;
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Кол.");
                cell.CellStyle = HeaderStyle;
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Образец");
                cell.CellStyle = HeaderStyle;
                if (Security.PriceAccess)
                {
                    HSSFCell cell22 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Ср.цена, евро");
                    cell22.CellStyle = HeaderStyle;
                    HSSFCell cell23 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Сумма, евро");
                    cell23.CellStyle = HeaderStyle;
                }
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Менеджер");
                cell.CellStyle = HeaderStyle;
                if (!ZOV)
                {
                    cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Страна");
                    cell.CellStyle = HeaderStyle;
                }

                RowIndex++;

                DataTable dt = DecorResultDataTable[c].Copy();
                if (ZOV)
                {
                    dt.Columns.Remove("CountryName");
                    dt.Columns.Remove("ZOVClientName");
                }
                else
                {
                    dt.Columns.Remove("ClientGroupName");
                    dt.Columns.Remove("DocNumber");
                    dt.Columns.Remove("DispatchDateTime");
                }

                int ColumnCount = dt.Columns.Count;

                if (!Security.PriceAccess)
                {
                    ColumnCount = dt.Columns.Count - 2;
                }

                //вывод заказов декора в excel
                for (int x = 0; x < dt.Rows.Count; x++)
                {
                    for (int y = 0; y < ColumnCount; y++)
                    {
                        //if (ZOV && DecorResultDataTable[c].Columns[y].ColumnName == "CountryName")
                        //    continue;

                        //if (!ZOV && DecorResultDataTable[c].Columns[y].ColumnName == "ClientGroupName")
                        //    continue;

                        //int ColumnIndex = y;

                        //if (y == 0)
                        //{
                        //    ColumnIndex = y;
                        //}
                        //else
                        //{
                        //    ColumnIndex = y + 1;
                        //}

                        Type t = dt.Rows[x][y].GetType();

                        //sheet1.CreateRow(RowIndex).CreateCell(1).CellStyle = SimpleCellStyle;

                        if (t.Name == "Decimal")
                        {
                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(Convert.ToDouble(dt.Rows[x][y]));

                            cell.CellStyle = cellStyle;
                            continue;
                        }
                        if (t.Name == "Int32")
                        {
                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(Convert.ToInt32(dt.Rows[x][y]));
                            cell.CellStyle = SimpleCellStyle;
                            continue;
                        }

                        if (t.Name == "String" || t.Name == "DBNull")
                        {
                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(dt.Rows[x][y].ToString());
                            cell.CellStyle = SimpleCellStyle;
                            continue;
                        }
                    }

                    RowIndex++;
                }
                RowIndex++;
            }
        }
        
        private void DecorReport(HSSFWorkbook hssfworkbook, HSSFCellStyle HeaderStyle,
            HSSFFont SimpleFont, HSSFCellStyle SimpleCellStyle, HSSFCellStyle cellStyle,
            DataTable DecorOrdersDataTable, bool ZOV)
        {
            int RowIndex = 0;

            int DisplayIndex = 0;
            HSSFSheet sheet1 = hssfworkbook.CreateSheet("Декор");
            sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

            sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
            sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
            sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
            sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

            if (ZOV)
            {
                sheet1.SetColumnWidth(DisplayIndex++, 25 * 256);
                sheet1.SetColumnWidth(DisplayIndex++, 20 * 256);
                sheet1.SetColumnWidth(DisplayIndex++, 20 * 256);
                sheet1.SetColumnWidth(DisplayIndex++, 15 * 256);
            }
            else
            {
                sheet1.SetColumnWidth(DisplayIndex++, 35 * 256);
                sheet1.SetColumnWidth(DisplayIndex++, 20 * 256);
            }
            sheet1.SetColumnWidth(DisplayIndex++, 30 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 20 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 20 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 10 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 10 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 10 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 10 * 256);
            if (Security.PriceAccess)
            {
                sheet1.SetColumnWidth(DisplayIndex++, 17 * 256);
                sheet1.SetColumnWidth(DisplayIndex++, 15 * 256);
            }
            sheet1.SetColumnWidth(DisplayIndex++, 20 * 256);
            //декор

            FillDecor(DecorOrdersDataTable, ZOV);

            for (int c = 0; c < DecorCatalog.DecorProductsCount; c++)
            {
                if (DecorResultDataTable[c].Rows.Count == 0)
                    continue;
                DisplayIndex = 0;

                HSSFCell cell;
                if (ZOV)
                {
                    cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Группа");
                    cell.CellStyle = HeaderStyle;
                    cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Клиент");
                    cell.CellStyle = HeaderStyle;
                    cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "№ заказа");
                    cell.CellStyle = HeaderStyle;
                    cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Дата отгр");
                    cell.CellStyle = HeaderStyle;
                }
                else
                {
                    cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Клиент");
                    cell.CellStyle = HeaderStyle;
                    cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Клиент ЗОВ");
                    cell.CellStyle = HeaderStyle;
                }
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Название");
                cell.CellStyle = HeaderStyle;
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Цвет");
                cell.CellStyle = HeaderStyle;
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Патина");
                cell.CellStyle = HeaderStyle;
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Длин\\Выс.");
                cell.CellStyle = HeaderStyle;
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Шир.");
                cell.CellStyle = HeaderStyle;
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Кол.");
                cell.CellStyle = HeaderStyle;
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Образец");
                cell.CellStyle = HeaderStyle;
                if (Security.PriceAccess)
                {
                    HSSFCell cell22 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Ср.цена, евро");
                    cell22.CellStyle = HeaderStyle;
                    HSSFCell cell23 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Сумма, евро");
                    cell23.CellStyle = HeaderStyle;
                }
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Менеджер");
                cell.CellStyle = HeaderStyle;
                if (!ZOV)
                {
                    cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Страна");
                    cell.CellStyle = HeaderStyle;
                }

                RowIndex++;

                DataTable dt = DecorResultDataTable[c].Copy();
                if (ZOV)
                {
                    dt.Columns.Remove("CountryName");
                    dt.Columns.Remove("ZOVClientName");
                }
                else
                {
                    dt.Columns.Remove("ClientGroupName");
                    dt.Columns.Remove("DocNumber");
                    dt.Columns.Remove("DispatchDateTime");
                }

                int ColumnCount = dt.Columns.Count;

                if (!Security.PriceAccess)
                {
                    ColumnCount = dt.Columns.Count - 2;
                }

                //вывод заказов декора в excel
                for (int x = 0; x < dt.Rows.Count; x++)
                {
                    for (int y = 0; y < ColumnCount; y++)
                    {
                        //if (ZOV && DecorResultDataTable[c].Columns[y].ColumnName == "CountryName")
                        //    continue;

                        //if (!ZOV && DecorResultDataTable[c].Columns[y].ColumnName == "ClientGroupName")
                        //    continue;

                        //int ColumnIndex = y;

                        //if (y == 0)
                        //{
                        //    ColumnIndex = y;
                        //}
                        //else
                        //{
                        //    ColumnIndex = y + 1;
                        //}

                        Type t = dt.Rows[x][y].GetType();

                        //sheet1.CreateRow(RowIndex).CreateCell(1).CellStyle = SimpleCellStyle;

                        if (t.Name == "Decimal")
                        {
                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(Convert.ToDouble(dt.Rows[x][y]));

                            cell.CellStyle = cellStyle;
                            continue;
                        }
                        if (t.Name == "Int32")
                        {
                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(Convert.ToInt32(dt.Rows[x][y]));
                            cell.CellStyle = SimpleCellStyle;
                            continue;
                        }

                        if (t.Name == "String" || t.Name == "DBNull")
                        {
                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(dt.Rows[x][y].ToString());
                            cell.CellStyle = SimpleCellStyle;
                            continue;
                        }
                    }

                    RowIndex++;
                }
                RowIndex++;
            }
        }

        private void FillDecorOrderNumber(DataTable DecorOrdersDataTable, bool ZOV)
        {
            for (int i = 0; i < DecorCatalog.DecorProductsCount; i++)
            {
                DecorResultDataTable[i].Clear();
                DecorResultDataTable[i].AcceptChanges();
            }

            if (DecorOrdersDataTable.Rows.Count < 1)
                return;

            if (!ZOV)
                FillMDecorOrderNumber(DecorOrdersDataTable);
            if (ZOV)
                FillZDecor(DecorOrdersDataTable);

            for (int i = 0; i < DecorCatalog.DecorProductsCount; i++)
            {
                DataView DV1 = new DataView(DecorResultDataTable[i])
                {
                    Sort = "Product, Color, Height, Width, Client, Count"
                };
                DecorResultDataTable[i] = DV1.ToTable();
                DV1.Dispose();
            }
        }
        
        private void FillDecor(DataTable DecorOrdersDataTable, bool ZOV)
        {
            for (int i = 0; i < DecorCatalog.DecorProductsCount; i++)
            {
                DecorResultDataTable[i].Clear();
                DecorResultDataTable[i].AcceptChanges();
            }

            if (DecorOrdersDataTable.Rows.Count < 1)
                return;

            if (!ZOV)
                FillMDecor(DecorOrdersDataTable);
            if (ZOV)
                FillZDecor(DecorOrdersDataTable);

            for (int i = 0; i < DecorCatalog.DecorProductsCount; i++)
            {
                DataView DV1 = new DataView(DecorResultDataTable[i])
                {
                    Sort = "Product, Color, Height, Width, Client, Count"
                };
                DecorResultDataTable[i] = DV1.ToTable();
                DV1.Dispose();
            }
        }

        private void FillFrontsOrderNumber(DataTable FrontsOrdersDataTable, bool ZOV)
        {
            FrontsResultDataTable.Clear();

            if (FrontsOrdersDataTable.Rows.Count < 1)
                return;

            if (!ZOV)
                FillMFrontsOrderNumber(FrontsOrdersDataTable);
            if (ZOV)
                FillZFronts(FrontsOrdersDataTable);

            DataView DV1 = new DataView(FrontsResultDataTable)
            {
                Sort = "Front, FrameColor, TechnoColor, Patina, InsetType, InsetColor, TechnoInsetType, TechnoInsetColor, Height, Width, Count"
            };
            FrontsResultDataTable = DV1.ToTable();
            DV1.Dispose();
        }
        
        private void FillFronts(DataTable FrontsOrdersDataTable, bool ZOV)
        {
            FrontsResultDataTable.Clear();

            if (FrontsOrdersDataTable.Rows.Count < 1)
                return;

            if (!ZOV)
                FillMFronts(FrontsOrdersDataTable);
            if (ZOV)
                FillZFronts(FrontsOrdersDataTable);

            DataView DV1 = new DataView(FrontsResultDataTable)
            {
                Sort = "Front, FrameColor, TechnoColor, Patina, InsetType, InsetColor, TechnoInsetType, TechnoInsetColor, Height, Width, Count"
            };
            FrontsResultDataTable = DV1.ToTable();
            DV1.Dispose();
        }

        private void FillMDecorOrderNumber(DataTable DecorOrdersDataTable)
        {
            if (DecorOrdersDataTable.Rows.Count < 1)
                return;

            string QueryString = string.Empty;
            string Product = string.Empty;
            string Decor = string.Empty;
            string Color = string.Empty;
            string Patina = string.Empty;

            int Height = 0;
            int Width = 0;
            int Count = 0;

            string ManagerName = "Менеджер";
            string ClientName = "Клиент";
            string CountryName = "Страна";
            string ZOVClientName = string.Empty;

            decimal DecorCost = 0;
            decimal SumCost = 0;
            decimal SumCount = 0;
            decimal AvgPrice = 0;

            DataTable TempDecorProductsDT = new DataTable();
            DataTable Table = new DataTable();

            DataView DV3 = new DataView(DecorCatalog.DecorProductsDataTable)
            {
                Sort = "ProductName"
            };
            TempDecorProductsDT = DV3.ToTable();
            DV3.Dispose();

            for (int i = 0; i < DecorOrdersDataTable.Rows.Count; i++)
            {
                if (DecorOrdersDataTable.Rows[i]["ZOVClientID"] == DBNull.Value)
                    DecorOrdersDataTable.Rows[i]["ZOVClientID"] = -1;
            }

            for (int i = 0; i < DecorCatalog.DecorProductsCount; i++)
            {
                using (DataView DV = new DataView(DecorOrdersDataTable,
                    "ProductID = " + Convert.ToInt32(TempDecorProductsDT.Rows[i]["ProductID"]),
                    string.Empty, DataViewRowState.CurrentRows))
                {
                    Table = DV.ToTable(true, new string[] { "DecorID", "ColorID", "PatinaID", "MeasureID", "Length", "Height", "Width", "IsSample", "ClientID", "OrderNumber", "ZOVClientID" });
                }

                for (int j = 0; j < Table.Rows.Count; j++)
                {
                    QueryString = "ProductID=" + Convert.ToInt32(TempDecorProductsDT.Rows[i]["ProductID"]) +
                        " AND DecorID=" + Convert.ToInt32(Table.Rows[j]["DecorID"]) +
                        " AND ColorID=" + Convert.ToInt32(Table.Rows[j]["ColorID"]) +
                        " AND PatinaID=" + Convert.ToInt32(Table.Rows[j]["PatinaID"]) +
                        " AND MeasureID=" + Convert.ToInt32(Table.Rows[j]["MeasureID"]) +
                        " AND Length=" + Convert.ToInt32(Table.Rows[j]["Length"]) +
                        " AND Height=" + Convert.ToInt32(Table.Rows[j]["Height"]) +
                        " AND Width=" + Convert.ToInt32(Table.Rows[j]["Width"]) +
                        " AND IsSample=" + Convert.ToInt32(Table.Rows[j]["IsSample"]) +
                        " AND OrderNumber=" + Convert.ToInt32(Table.Rows[j]["OrderNumber"]) +
                        " AND ClientID=" + Convert.ToInt32(Table.Rows[j]["ClientID"]) +
                        " AND ZOVClientID = " + Convert.ToInt32(Table.Rows[j]["ZOVClientID"]);

                    DataRow[] Rows = DecorOrdersDataTable.Select(QueryString);

                    if (Rows.Count() == 0)
                        continue;

                    foreach (DataRow Row in Rows)
                    {
                        SumCount = 0;
                        SumCost = 0;

                        foreach (DataRow row in Rows)
                        {
                            if (Convert.ToInt32(row["MeasureID"]) == 1)
                            {
                                SumCount += Convert.ToDecimal(row["Height"]) * Convert.ToDecimal(row["Width"]) * Convert.ToInt32(row["Count"]) / 1000000;
                            }

                            if (Convert.ToInt32(row["MeasureID"]) == 3)
                            {
                                SumCount += Convert.ToInt32(row["Count"]);
                            }

                            if (Convert.ToInt32(row["MeasureID"]) == 2)
                            {
                                //нет параметра "высота"
                                if (row["Height"].ToString() == "-1")
                                    SumCount += Convert.ToDecimal(row["Length"]) * Convert.ToDecimal(row["Count"]) / 1000;
                                else
                                    SumCount += Convert.ToDecimal(row["Height"]) * Convert.ToDecimal(row["Count"]) / 1000;
                            }
                            SumCost += Convert.ToDecimal(row["Cost"]);
                        }

                        Product = TempDecorProductsDT.Rows[i]["ProductName"].ToString() + " " +
                            DecorCatalog.GetItemName(Convert.ToInt32(Row["DecorID"]));
                        Count = Convert.ToInt32(Row["Count"]);
                        DecorCost = Convert.ToDecimal(Row["Cost"]);

                        if (SumCount != 0)
                            AvgPrice = SumCost / SumCount;

                        QueryString = "Product = '" + Product + "'";

                        if (DecorCatalog.HasParameter(Convert.ToInt32(Row["ProductID"]), "ColorID"))
                        {
                            Color = GetColorName(Convert.ToInt32(Row["ColorID"]));
                            QueryString += " AND Color = '" + Color.ToString() + "'";
                            Patina = "-";
                            if (Convert.ToInt32(Row["PatinaID"]) != -1)
                            {
                                Patina = GetPatinaName(Convert.ToInt32(Row["PatinaID"]));
                                QueryString += $" AND Patina = '{ Patina.ToString() }'";
                            }
                        }

                        if (DecorCatalog.HasParameter(Convert.ToInt32(Row["ProductID"]), "Height"))
                        {
                            Height = Convert.ToInt32(Row["Height"]);
                            QueryString += " AND Height = '" + Height.ToString() + "'";
                        }

                        if (DecorCatalog.HasParameter(Convert.ToInt32(Row["ProductID"]), "Length"))
                        {
                            Height = Convert.ToInt32(Row["Length"]);
                            QueryString += " AND Height = '" + Height.ToString() + "'";
                        }

                        if (DecorCatalog.HasParameter(Convert.ToInt32(Row["ProductID"]), "Width"))
                        {
                            Width = Convert.ToInt32(Row["Width"]);
                            QueryString += " AND Width = '" + Width.ToString() + "'";
                        }

                        ManagerName = GetManagerName(Convert.ToInt32(Row["ClientID"]));
                        CountryName = GetCountryName(Convert.ToInt32(Row["ClientID"]));
                        ClientName = GetClientName(Convert.ToInt32(Row["ClientID"]));

                        ZOVClientName = string.Empty;
                        if (Convert.ToInt32(Row["ClientID"]) == 145)
                        {
                            if (Convert.ToInt32(Row["ZOVClientID"]) == -1)
                                ZOVClientName = "Маркетинг.Заказы";
                            else
                                ZOVClientName = GetZOVClientName(Convert.ToInt32(Row["ZOVClientID"]));
                        }

                        QueryString += " AND Client = '" + ClientName + "'";
                        DataRow[] dRow = DecorResultDataTable[i].Select(QueryString);

                        if (dRow.Count() == 0)
                        {
                            DataRow NewRow = DecorResultDataTable[i].NewRow();

                            NewRow["Product"] = Product;
                            NewRow["ManagerName"] = ManagerName;
                            NewRow["CountryName"] = CountryName;
                            NewRow["Client"] = ClientName;
                            NewRow["ZOVClientName"] = ZOVClientName;

                            if (DecorCatalog.HasParameter(Convert.ToInt32(Row["ProductID"]), "Height"))
                            {
                                if (Convert.ToInt32(Row["Height"]) != -1)
                                    NewRow["Height"] = Row["Height"];
                                if (Convert.ToInt32(Row["Length"]) != -1)
                                    NewRow["Height"] = Row["Length"];
                            }
                            if (DecorCatalog.HasParameter(Convert.ToInt32(Row["ProductID"]), "Length"))
                            {
                                if (Convert.ToInt32(Row["Length"]) != -1)
                                    NewRow["Height"] = Row["Length"];
                                if (Convert.ToInt32(Row["Height"]) != -1)
                                    NewRow["Height"] = Row["Height"];
                            }
                            if (DecorCatalog.HasParameter(Convert.ToInt32(Row["ProductID"]), "Width"))
                                NewRow["Width"] = Row["Width"];

                            if (DecorCatalog.HasParameter(Convert.ToInt32(Row["ProductID"]), "ColorID"))
                                NewRow["Color"] = Color;
                            NewRow["Patina"] = Patina;
                            NewRow["Count"] = Row["Count"];
                            NewRow["OrderNumber"] = Row["OrderNumber"];
                            NewRow["IsSample"] = Row["IsSample"];
                            NewRow["AvgPrice"] = decimal.Round(AvgPrice, 2, MidpointRounding.AwayFromZero);
                            NewRow["SumCost"] = Row["Cost"];

                            DecorResultDataTable[i].Rows.Add(NewRow);
                        }
                        else
                        {
                            dRow[0]["Count"] = Convert.ToDecimal(dRow[0]["Count"]) + Count;
                            dRow[0]["SumCost"] = decimal.Round(Convert.ToDecimal(dRow[0]["SumCost"]) + DecorCost, 3, MidpointRounding.AwayFromZero);
                        }
                    }
                }
            }

            TempDecorProductsDT.Dispose();
        }
        
        private void FillMDecor(DataTable DecorOrdersDataTable)
        {
            if (DecorOrdersDataTable.Rows.Count < 1)
                return;

            string QueryString = string.Empty;
            string Product = string.Empty;
            string Decor = string.Empty;
            string Color = string.Empty;
            string Patina = string.Empty;

            int Height = 0;
            int Width = 0;
            int Count = 0;

            string ManagerName = "Менеджер";
            string ClientName = "Клиент";
            string CountryName = "Страна";
            string ZOVClientName = string.Empty;

            decimal DecorCost = 0;
            decimal SumCost = 0;
            decimal SumCount = 0;
            decimal AvgPrice = 0;

            DataTable TempDecorProductsDT = new DataTable();
            DataTable Table = new DataTable();

            DataView DV3 = new DataView(DecorCatalog.DecorProductsDataTable)
            {
                Sort = "ProductName"
            };
            TempDecorProductsDT = DV3.ToTable();
            DV3.Dispose();

            for (int i = 0; i < DecorOrdersDataTable.Rows.Count; i++)
            {
                if (DecorOrdersDataTable.Rows[i]["ZOVClientID"] == DBNull.Value)
                    DecorOrdersDataTable.Rows[i]["ZOVClientID"] = -1;
            }

            for (int i = 0; i < DecorCatalog.DecorProductsCount; i++)
            {
                using (DataView DV = new DataView(DecorOrdersDataTable,
                    "ProductID = " + Convert.ToInt32(TempDecorProductsDT.Rows[i]["ProductID"]),
                    string.Empty, DataViewRowState.CurrentRows))
                {
                    Table = DV.ToTable(true, 
                        new string[] { "DecorID", "ColorID", "PatinaID", "MeasureID", "Length", "Height", "Width", "IsSample", "ClientID", "ZOVClientID" });
                }

                for (int j = 0; j < Table.Rows.Count; j++)
                {
                    QueryString = "ProductID=" + Convert.ToInt32(TempDecorProductsDT.Rows[i]["ProductID"]) +
                        " AND DecorID=" + Convert.ToInt32(Table.Rows[j]["DecorID"]) +
                        " AND ColorID=" + Convert.ToInt32(Table.Rows[j]["ColorID"]) +
                        " AND PatinaID=" + Convert.ToInt32(Table.Rows[j]["PatinaID"]) +
                        " AND MeasureID=" + Convert.ToInt32(Table.Rows[j]["MeasureID"]) +
                        " AND Length=" + Convert.ToInt32(Table.Rows[j]["Length"]) +
                        " AND Height=" + Convert.ToInt32(Table.Rows[j]["Height"]) +
                        " AND Width=" + Convert.ToInt32(Table.Rows[j]["Width"]) +
                        " AND IsSample=" + Convert.ToInt32(Table.Rows[j]["IsSample"]) +
                        " AND ClientID=" + Convert.ToInt32(Table.Rows[j]["ClientID"]) +
                        " AND ZOVClientID = " + Convert.ToInt32(Table.Rows[j]["ZOVClientID"]);

                    DataRow[] Rows = DecorOrdersDataTable.Select(QueryString);

                    if (Rows.Count() == 0)
                        continue;

                    foreach (DataRow Row in Rows)
                    {
                        SumCount = 0;
                        SumCost = 0;

                        foreach (DataRow row in Rows)
                        {
                            if (Convert.ToInt32(row["MeasureID"]) == 1)
                            {
                                SumCount += Convert.ToDecimal(row["Height"]) * Convert.ToDecimal(row["Width"]) * Convert.ToInt32(row["Count"]) / 1000000;
                            }

                            if (Convert.ToInt32(row["MeasureID"]) == 3)
                            {
                                SumCount += Convert.ToInt32(row["Count"]);
                            }

                            if (Convert.ToInt32(row["MeasureID"]) == 2)
                            {
                                //нет параметра "высота"
                                if (row["Height"].ToString() == "-1")
                                    SumCount += Convert.ToDecimal(row["Length"]) * Convert.ToDecimal(row["Count"]) / 1000;
                                else
                                    SumCount += Convert.ToDecimal(row["Height"]) * Convert.ToDecimal(row["Count"]) / 1000;
                            }
                            SumCost += Convert.ToDecimal(row["Cost"]);
                        }

                        Product = TempDecorProductsDT.Rows[i]["ProductName"].ToString() + " " +
                            DecorCatalog.GetItemName(Convert.ToInt32(Row["DecorID"]));
                        Count = Convert.ToInt32(Row["Count"]);
                        DecorCost = Convert.ToDecimal(Row["Cost"]);

                        if (SumCount != 0)
                            AvgPrice = SumCost / SumCount;

                        QueryString = "Product = '" + Product + "'";

                        if (DecorCatalog.HasParameter(Convert.ToInt32(Row["ProductID"]), "ColorID"))
                        {
                            Color = GetColorName(Convert.ToInt32(Row["ColorID"]));
                            QueryString += " AND Color = '" + Color.ToString() + "'";
                            Patina = "-";
                            if (Convert.ToInt32(Row["PatinaID"]) != -1)
                            {
                                Patina = GetPatinaName(Convert.ToInt32(Row["PatinaID"]));
                                QueryString += $" AND Patina = '{ Patina.ToString() }'";
                            }
                        }

                        if (DecorCatalog.HasParameter(Convert.ToInt32(Row["ProductID"]), "Height"))
                        {
                            Height = Convert.ToInt32(Row["Height"]);
                            QueryString += " AND Height = '" + Height.ToString() + "'";
                        }

                        if (DecorCatalog.HasParameter(Convert.ToInt32(Row["ProductID"]), "Length"))
                        {
                            Height = Convert.ToInt32(Row["Length"]);
                            QueryString += " AND Height = '" + Height.ToString() + "'";
                        }

                        if (DecorCatalog.HasParameter(Convert.ToInt32(Row["ProductID"]), "Width"))
                        {
                            Width = Convert.ToInt32(Row["Width"]);
                            QueryString += " AND Width = '" + Width.ToString() + "'";
                        }

                        ManagerName = GetManagerName(Convert.ToInt32(Row["ClientID"]));
                        CountryName = GetCountryName(Convert.ToInt32(Row["ClientID"]));
                        ClientName = GetClientName(Convert.ToInt32(Row["ClientID"]));

                        ZOVClientName = string.Empty;
                        if (Convert.ToInt32(Row["ClientID"]) == 145)
                        {
                            if (Convert.ToInt32(Row["ZOVClientID"]) == -1)
                                ZOVClientName = "Маркетинг.Заказы";
                            else
                                ZOVClientName = GetZOVClientName(Convert.ToInt32(Row["ZOVClientID"]));
                        }

                        QueryString += " AND Client = '" + ClientName + "'";
                        DataRow[] dRow = DecorResultDataTable[i].Select(QueryString);

                        if (dRow.Count() == 0)
                        {
                            DataRow NewRow = DecorResultDataTable[i].NewRow();

                            NewRow["Product"] = Product;
                            NewRow["ManagerName"] = ManagerName;
                            NewRow["CountryName"] = CountryName;
                            NewRow["Client"] = ClientName;
                            NewRow["ZOVClientName"] = ZOVClientName;

                            if (DecorCatalog.HasParameter(Convert.ToInt32(Row["ProductID"]), "Height"))
                            {
                                if (Convert.ToInt32(Row["Height"]) != -1)
                                    NewRow["Height"] = Row["Height"];
                                if (Convert.ToInt32(Row["Length"]) != -1)
                                    NewRow["Height"] = Row["Length"];
                            }
                            if (DecorCatalog.HasParameter(Convert.ToInt32(Row["ProductID"]), "Length"))
                            {
                                if (Convert.ToInt32(Row["Length"]) != -1)
                                    NewRow["Height"] = Row["Length"];
                                if (Convert.ToInt32(Row["Height"]) != -1)
                                    NewRow["Height"] = Row["Height"];
                            }
                            if (DecorCatalog.HasParameter(Convert.ToInt32(Row["ProductID"]), "Width"))
                                NewRow["Width"] = Row["Width"];

                            if (DecorCatalog.HasParameter(Convert.ToInt32(Row["ProductID"]), "ColorID"))
                                NewRow["Color"] = Color;
                            NewRow["Patina"] = Patina;
                            NewRow["Count"] = Row["Count"];
                            NewRow["IsSample"] = Row["IsSample"];
                            NewRow["AvgPrice"] = decimal.Round(AvgPrice, 2, MidpointRounding.AwayFromZero);
                            NewRow["SumCost"] = Row["Cost"];

                            DecorResultDataTable[i].Rows.Add(NewRow);
                        }
                        else
                        {
                            dRow[0]["Count"] = Convert.ToDecimal(dRow[0]["Count"]) + Count;
                            dRow[0]["SumCost"] = decimal.Round(Convert.ToDecimal(dRow[0]["SumCost"]) + DecorCost, 3, MidpointRounding.AwayFromZero);
                        }
                    }
                }
            }

            TempDecorProductsDT.Dispose();
        }
        
        private void FillMFrontsOrderNumber(DataTable FrontsOrdersDataTable)
        {
            if (FrontsOrdersDataTable.Rows.Count < 1)
                return;

            string Front = string.Empty;
            string FrameColor = string.Empty;
            string TechnoColor = string.Empty;
            string Patina = string.Empty;
            string InsetType = string.Empty;
            string InsetColor = string.Empty;
            string TechnoInsetType = string.Empty;
            string TechnoInsetColor = string.Empty;

            string ManagerName = "Менеджер";
            string ClientName = "Клиент";
            string ZOVClientName = string.Empty;
            string CountryName = "Страна";

            string QueryString = string.Empty;

            decimal FrontCost = 0;
            decimal AvgPrice = 0;
            decimal FrontSquare = 0;
            int FrontCount = 0;

            DataTable Table = new DataTable();
            for (int i = 0; i < FrontsOrdersDataTable.Rows.Count; i++)
            {
                if (FrontsOrdersDataTable.Rows[i]["ZOVClientID"] == DBNull.Value)
                    FrontsOrdersDataTable.Rows[i]["ZOVClientID"] = -1;
            }
            //using (DataView DV = new DataView(FrontsOrdersDataTable, string.Empty, string.Empty, DataViewRowState.CurrentRows))
            //{
            //    Table = DV.ToTable(true, new string[] { "FrontID",
            //        "ColorID", "TechnoColorID", "PatinaID","InsetTypeID", "InsetColorID", "TechnoInsetTypeID", "TechnoInsetColorID", "Height", "Width", "ClientID"});
            //}
            //Table.Clear();
            using (DataView DV = new DataView(FrontsOrdersDataTable, string.Empty, string.Empty, DataViewRowState.CurrentRows))
            {
                Table = DV.ToTable(true, new string[] { "FrontID",
                    "ColorID", "TechnoColorID", "PatinaID","InsetTypeID", "InsetColorID", "TechnoInsetTypeID", "TechnoInsetColorID", "Height", "Width", "IsSample", "ClientID", "OrderNumber", "ZOVClientID"});
            }

            for (int i = 0; i < Table.Rows.Count; i++)
            {
                QueryString = "FrontID=" + Convert.ToInt32(Table.Rows[i]["FrontID"]) +
                    " AND ColorID=" + Convert.ToInt32(Table.Rows[i]["ColorID"]) +
                    " AND TechnoColorID=" + Convert.ToInt32(Table.Rows[i]["TechnoColorID"]) +
                    " AND PatinaID=" + Convert.ToInt32(Table.Rows[i]["PatinaID"]) +
                    " AND InsetTypeID=" + Convert.ToInt32(Table.Rows[i]["InsetTypeID"]) +
                    " AND InsetColorID=" + Convert.ToInt32(Table.Rows[i]["InsetColorID"]) +
                    " AND TechnoInsetTypeID=" + Convert.ToInt32(Table.Rows[i]["TechnoInsetTypeID"]) +
                    " AND TechnoInsetColorID=" + Convert.ToInt32(Table.Rows[i]["TechnoInsetColorID"]) +
                    " AND Height = '" + Table.Rows[i]["Height"].ToString() + "'" +
                    " AND Width = '" + Table.Rows[i]["Width"].ToString() + "'" +
                    " AND IsSample=" + Convert.ToInt32(Table.Rows[i]["IsSample"]) +
                    " AND ClientID=" + Convert.ToInt32(Table.Rows[i]["ClientID"]) +
                    " AND OrderNumber=" + Convert.ToInt32(Table.Rows[i]["OrderNumber"]) +
                    " AND ZOVClientID=" + Convert.ToInt32(Table.Rows[i]["ZOVClientID"]);

                DataRow[] Rows = FrontsOrdersDataTable.Select(QueryString);
                if (Rows.Count() != 0)
                {
                    foreach (DataRow row in Rows)
                    {
                        FrontCost += Convert.ToDecimal(row["Cost"]);
                        FrontSquare += Convert.ToDecimal(row["Square"]);
                        FrontCount += Convert.ToInt32(row["Count"]);
                    }
                    if (Convert.ToInt32(Rows[0]["MeasureID"]) == 1)
                    {
                        if (FrontSquare != 0)
                            AvgPrice = FrontCost / FrontSquare;
                    }
                    if (Convert.ToInt32(Rows[0]["MeasureID"]) == 3)
                    {
                        if (FrontCount != 0)
                            AvgPrice = FrontCost / FrontCount;
                    }
                    Front = GetFrontName(Convert.ToInt32(Table.Rows[i]["FrontID"]));
                    FrameColor = GetColorName(Convert.ToInt32(Table.Rows[i]["ColorID"]));
                    TechnoColor = GetColorName(Convert.ToInt32(Table.Rows[i]["TechnoColorID"]));
                    Patina = GetPatinaName(Convert.ToInt32(Table.Rows[i]["PatinaID"]));
                    InsetType = GetInsetTypeName(Convert.ToInt32(Table.Rows[i]["InsetTypeID"]));
                    InsetColor = GetInsetColorName(Convert.ToInt32(Table.Rows[i]["InsetColorID"]));
                    TechnoInsetType = GetInsetTypeName(Convert.ToInt32(Table.Rows[i]["TechnoInsetTypeID"]));
                    TechnoInsetColor = GetInsetColorName(Convert.ToInt32(Table.Rows[i]["TechnoInsetColorID"]));

                    ManagerName = GetManagerName(Convert.ToInt32(Table.Rows[i]["ClientID"]));
                    CountryName = GetCountryName(Convert.ToInt32(Table.Rows[i]["ClientID"]));
                    ClientName = GetClientName(Convert.ToInt32(Table.Rows[i]["ClientID"]));

                    ZOVClientName = string.Empty;
                    if (Convert.ToInt32(Table.Rows[i]["ClientID"]) == 145)
                    {
                        if (Convert.ToInt32(Table.Rows[i]["ZOVClientID"]) == -1)
                            ZOVClientName = "Маркетинг.Заказы";
                        else
                            ZOVClientName = GetZOVClientName(Convert.ToInt32(Table.Rows[i]["ZOVClientID"]));
                    }
                    //if (FrontType == "Прямой")
                    //    FrontType = string.Empty;

                    DataRow NewRow = FrontsResultDataTable.NewRow();
                    NewRow["Front"] = Front;
                    NewRow["Patina"] = Patina;
                    NewRow["FrameColor"] = FrameColor;
                    NewRow["TechnoColor"] = TechnoColor;
                    NewRow["InsetType"] = InsetType;
                    NewRow["InsetColor"] = InsetColor;
                    NewRow["TechnoInsetType"] = TechnoInsetType;
                    NewRow["TechnoInsetColor"] = TechnoInsetColor;
                    NewRow["IsSample"] = Convert.ToInt32(Table.Rows[i]["IsSample"]);
                    NewRow["Height"] = Convert.ToInt32(Table.Rows[i]["Height"]);
                    NewRow["Width"] = Convert.ToInt32(Table.Rows[i]["Width"]);
                    NewRow["OrderNumber"] = Convert.ToInt32(Table.Rows[i]["OrderNumber"]);
                    NewRow["Square"] = decimal.Round(FrontSquare, 3, MidpointRounding.AwayFromZero);
                    NewRow["Count"] = FrontCount;
                    NewRow["ManagerName"] = ManagerName;
                    NewRow["CountryName"] = CountryName;
                    NewRow["Client"] = ClientName;
                    NewRow["ZOVClientName"] = ZOVClientName;
                    NewRow["AvgPrice"] = decimal.Round(AvgPrice, 2, MidpointRounding.AwayFromZero);
                    NewRow["SumCost"] = decimal.Round(FrontCost, 3, MidpointRounding.AwayFromZero);
                    FrontsResultDataTable.Rows.Add(NewRow);

                    ClientName = string.Empty;
                    AvgPrice = 0;
                    FrontCost = 0;
                    FrontSquare = 0;
                    FrontCount = 0;
                }
            }
        }
                
        private void FillMFronts(DataTable FrontsOrdersDataTable)
        {
            if (FrontsOrdersDataTable.Rows.Count < 1)
                return;

            string Front = string.Empty;
            string FrameColor = string.Empty;
            string TechnoColor = string.Empty;
            string Patina = string.Empty;
            string InsetType = string.Empty;
            string InsetColor = string.Empty;
            string TechnoInsetType = string.Empty;
            string TechnoInsetColor = string.Empty;

            string ManagerName = "Менеджер";
            string ClientName = "Клиент";
            string ZOVClientName = string.Empty;
            string CountryName = "Страна";

            string QueryString = string.Empty;

            decimal FrontCost = 0;
            decimal AvgPrice = 0;
            decimal FrontSquare = 0;
            int FrontCount = 0;

            DataTable Table = new DataTable();
            for (int i = 0; i < FrontsOrdersDataTable.Rows.Count; i++)
            {
                if (FrontsOrdersDataTable.Rows[i]["ZOVClientID"] == DBNull.Value)
                    FrontsOrdersDataTable.Rows[i]["ZOVClientID"] = -1;
            }
            //using (DataView DV = new DataView(FrontsOrdersDataTable, string.Empty, string.Empty, DataViewRowState.CurrentRows))
            //{
            //    Table = DV.ToTable(true, new string[] { "FrontID",
            //        "ColorID", "TechnoColorID", "PatinaID","InsetTypeID", "InsetColorID", "TechnoInsetTypeID", "TechnoInsetColorID", "Height", "Width", "ClientID"});
            //}
            //Table.Clear();
            using (DataView DV = new DataView(FrontsOrdersDataTable, string.Empty, string.Empty, DataViewRowState.CurrentRows))
            {
                Table = DV.ToTable(true, 
                    new string[] { "FrontID",
                    "ColorID", "TechnoColorID", "PatinaID","InsetTypeID", "InsetColorID", "TechnoInsetTypeID", "TechnoInsetColorID", "Height", "Width", "IsSample", "ClientID", "ZOVClientID"});
            }

            for (int i = 0; i < Table.Rows.Count; i++)
            {
                QueryString = "FrontID=" + Convert.ToInt32(Table.Rows[i]["FrontID"]) +
                    " AND ColorID=" + Convert.ToInt32(Table.Rows[i]["ColorID"]) +
                    " AND TechnoColorID=" + Convert.ToInt32(Table.Rows[i]["TechnoColorID"]) +
                    " AND PatinaID=" + Convert.ToInt32(Table.Rows[i]["PatinaID"]) +
                    " AND InsetTypeID=" + Convert.ToInt32(Table.Rows[i]["InsetTypeID"]) +
                    " AND InsetColorID=" + Convert.ToInt32(Table.Rows[i]["InsetColorID"]) +
                    " AND TechnoInsetTypeID=" + Convert.ToInt32(Table.Rows[i]["TechnoInsetTypeID"]) +
                    " AND TechnoInsetColorID=" + Convert.ToInt32(Table.Rows[i]["TechnoInsetColorID"]) +
                    " AND Height = '" + Table.Rows[i]["Height"].ToString() + "'" +
                    " AND Width = '" + Table.Rows[i]["Width"].ToString() + "'" +
                    " AND IsSample=" + Convert.ToInt32(Table.Rows[i]["IsSample"]) +
                    " AND ClientID=" + Convert.ToInt32(Table.Rows[i]["ClientID"]) +
                    " AND ZOVClientID=" + Convert.ToInt32(Table.Rows[i]["ZOVClientID"]);

                DataRow[] Rows = FrontsOrdersDataTable.Select(QueryString);
                if (Rows.Count() != 0)
                {
                    foreach (DataRow row in Rows)
                    {
                        FrontCost += Convert.ToDecimal(row["Cost"]);
                        FrontSquare += Convert.ToDecimal(row["Square"]);
                        FrontCount += Convert.ToInt32(row["Count"]);
                    }
                    if (Convert.ToInt32(Rows[0]["MeasureID"]) == 1)
                    {
                        if (FrontSquare != 0)
                            AvgPrice = FrontCost / FrontSquare;
                    }
                    if (Convert.ToInt32(Rows[0]["MeasureID"]) == 3)
                    {
                        if (FrontCount != 0)
                            AvgPrice = FrontCost / FrontCount;
                    }
                    Front = GetFrontName(Convert.ToInt32(Table.Rows[i]["FrontID"]));
                    FrameColor = GetColorName(Convert.ToInt32(Table.Rows[i]["ColorID"]));
                    TechnoColor = GetColorName(Convert.ToInt32(Table.Rows[i]["TechnoColorID"]));
                    Patina = GetPatinaName(Convert.ToInt32(Table.Rows[i]["PatinaID"]));
                    InsetType = GetInsetTypeName(Convert.ToInt32(Table.Rows[i]["InsetTypeID"]));
                    InsetColor = GetInsetColorName(Convert.ToInt32(Table.Rows[i]["InsetColorID"]));
                    TechnoInsetType = GetInsetTypeName(Convert.ToInt32(Table.Rows[i]["TechnoInsetTypeID"]));
                    TechnoInsetColor = GetInsetColorName(Convert.ToInt32(Table.Rows[i]["TechnoInsetColorID"]));

                    ManagerName = GetManagerName(Convert.ToInt32(Table.Rows[i]["ClientID"]));
                    CountryName = GetCountryName(Convert.ToInt32(Table.Rows[i]["ClientID"]));
                    ClientName = GetClientName(Convert.ToInt32(Table.Rows[i]["ClientID"]));

                    ZOVClientName = string.Empty;
                    if (Convert.ToInt32(Table.Rows[i]["ClientID"]) == 145)
                    {
                        if (Convert.ToInt32(Table.Rows[i]["ZOVClientID"]) == -1)
                            ZOVClientName = "Маркетинг.Заказы";
                        else
                            ZOVClientName = GetZOVClientName(Convert.ToInt32(Table.Rows[i]["ZOVClientID"]));
                    }
                    //if (FrontType == "Прямой")
                    //    FrontType = string.Empty;

                    DataRow NewRow = FrontsResultDataTable.NewRow();
                    NewRow["Front"] = Front;
                    NewRow["Patina"] = Patina;
                    NewRow["FrameColor"] = FrameColor;
                    NewRow["TechnoColor"] = TechnoColor;
                    NewRow["InsetType"] = InsetType;
                    NewRow["InsetColor"] = InsetColor;
                    NewRow["TechnoInsetType"] = TechnoInsetType;
                    NewRow["TechnoInsetColor"] = TechnoInsetColor;
                    NewRow["IsSample"] = Convert.ToInt32(Table.Rows[i]["IsSample"]);
                    NewRow["Height"] = Convert.ToInt32(Table.Rows[i]["Height"]);
                    NewRow["Width"] = Convert.ToInt32(Table.Rows[i]["Width"]);
                    NewRow["Square"] = decimal.Round(FrontSquare, 3, MidpointRounding.AwayFromZero);
                    NewRow["Count"] = FrontCount;
                    NewRow["ManagerName"] = ManagerName;
                    NewRow["CountryName"] = CountryName;
                    NewRow["Client"] = ClientName;
                    NewRow["ZOVClientName"] = ZOVClientName;
                    NewRow["AvgPrice"] = decimal.Round(AvgPrice, 2, MidpointRounding.AwayFromZero);
                    NewRow["SumCost"] = decimal.Round(FrontCost, 3, MidpointRounding.AwayFromZero);
                    FrontsResultDataTable.Rows.Add(NewRow);

                    ClientName = string.Empty;
                    AvgPrice = 0;
                    FrontCost = 0;
                    FrontSquare = 0;
                    FrontCount = 0;
                }
            }
        }

        private void FillZDecor(DataTable DecorOrdersDataTable)
        {
            if (DecorOrdersDataTable.Rows.Count < 1)
                return;

            string QueryString = string.Empty;
            string Product = string.Empty;
            string Decor = string.Empty;
            string Color = string.Empty;
            string Patina = string.Empty;

            int Height = 0;
            int Width = 0;
            int Count = 0;

            string ManagerName = "Менеджер ЗОВ";
            string ClientName = "Клиент ЗОВ";
            string ClientGroupName = "Группа ЗОВ";

            decimal DecorCost = 0;
            decimal SumCost = 0;
            decimal SumCount = 0;
            decimal AvgPrice = 0;

            DataTable TempDecorProductsDT = new DataTable();
            DataTable Table = new DataTable();

            DataView DV3 = new DataView(DecorCatalog.DecorProductsDataTable)
            {
                Sort = "ProductName"
            };
            TempDecorProductsDT = DV3.ToTable();
            DV3.Dispose();

            DataColumnCollection columns = DecorOrdersDataTable.Columns;
            if (columns.Contains("DispatchDateTime"))
            {
                for (int i = 0; i < DecorCatalog.DecorProductsCount; i++)
                {
                    using (DataView DV = new DataView(DecorOrdersDataTable,
                        "ProductID = " + Convert.ToInt32(TempDecorProductsDT.Rows[i]["ProductID"]),
                        string.Empty, DataViewRowState.CurrentRows))
                    {
                        Table = DV.ToTable(true, new string[] { "DecorID", "ColorID", "PatinaID", "MeasureID", "Length", "Height", "Width", "IsSample", "ZOVClientID", "DocNumber", "DispatchDateTime" });
                    }

                    for (int j = 0; j < Table.Rows.Count; j++)
                    {
                        QueryString = "ProductID=" + Convert.ToInt32(TempDecorProductsDT.Rows[i]["ProductID"]) +
                            " AND DecorID=" + Convert.ToInt32(Table.Rows[j]["DecorID"]) +
                            " AND ColorID=" + Convert.ToInt32(Table.Rows[j]["ColorID"]) +
                            " AND PatinaID=" + Convert.ToInt32(Table.Rows[j]["PatinaID"]) +
                            " AND MeasureID=" + Convert.ToInt32(Table.Rows[j]["MeasureID"]) +
                            " AND Length=" + Convert.ToInt32(Table.Rows[j]["Length"]) +
                            " AND Height=" + Convert.ToInt32(Table.Rows[j]["Height"]) +
                            " AND Width=" + Convert.ToInt32(Table.Rows[j]["Width"]) +
                            " AND IsSample=" + Convert.ToInt32(Table.Rows[j]["IsSample"]) +
                            " AND DocNumber = '" + Table.Rows[j]["DocNumber"].ToString() + "'" +
                            " AND DispatchDateTime = '" + Table.Rows[j]["DispatchDateTime"].ToString() + "'" +
                            " AND ZOVClientID=" + Convert.ToInt32(Table.Rows[j]["ZOVClientID"]);

                        DataRow[] Rows = DecorOrdersDataTable.Select(QueryString);

                        if (Rows.Count() == 0)
                            continue;

                        foreach (DataRow Row in Rows)
                        {
                            SumCount = 0;
                            SumCost = 0;

                            foreach (DataRow row in Rows)
                            {
                                if (Convert.ToInt32(row["MeasureID"]) == 1)
                                {
                                    SumCount += Convert.ToDecimal(row["Height"]) * Convert.ToDecimal(row["Width"]) * Convert.ToInt32(row["Count"]) / 1000000;
                                }

                                if (Convert.ToInt32(row["MeasureID"]) == 3)
                                {
                                    SumCount += Convert.ToInt32(row["Count"]);
                                }

                                if (Convert.ToInt32(row["MeasureID"]) == 2)
                                {
                                    //нет параметра "высота"
                                    if (row["Height"].ToString() == "-1")
                                        SumCount += Convert.ToDecimal(row["Length"]) * Convert.ToDecimal(row["Count"]) / 1000;
                                    else
                                        SumCount += Convert.ToDecimal(row["Height"]) * Convert.ToDecimal(row["Count"]) / 1000;
                                }
                                SumCost += Convert.ToDecimal(row["Cost"]);
                            }

                            Product = TempDecorProductsDT.Rows[i]["ProductName"].ToString() + " " +
                                DecorCatalog.GetItemName(Convert.ToInt32(Row["DecorID"]));
                            Count = Convert.ToInt32(Row["Count"]);
                            DecorCost = Convert.ToDecimal(Row["Cost"]);

                            DecorCost = DecorCost * 100 / 120;

                            if (SumCount != 0)
                                AvgPrice = SumCost / SumCount;

                            ManagerName = GetZOVManagerName(Convert.ToInt32(Row["ZOVClientID"]));
                            ClientName = GetZOVClientName(Convert.ToInt32(Row["ZOVClientID"]));
                            ClientGroupName = GetZOVClientGroupName(Convert.ToInt32(Row["ZOVClientID"]));

                            QueryString = "Product = '" + Product + "'";

                            if (DecorCatalog.HasParameter(Convert.ToInt32(Row["ProductID"]), "ColorID"))
                            {
                                Color = GetColorName(Convert.ToInt32(Row["ColorID"]));
                                QueryString += " AND Color = '" + Color.ToString() + "'";
                                Patina = "-";
                                if (Convert.ToInt32(Row["PatinaID"]) != -1)
                                {
                                    Patina = GetPatinaName(Convert.ToInt32(Row["PatinaID"]));
                                    QueryString += $" AND Patina = '{ Patina.ToString() }'";
                                }
                            }

                            if (DecorCatalog.HasParameter(Convert.ToInt32(Row["ProductID"]), "Height"))
                            {
                                Height = Convert.ToInt32(Row["Height"]);
                                QueryString += " AND Height = '" + Height.ToString() + "'";
                            }

                            if (DecorCatalog.HasParameter(Convert.ToInt32(Row["ProductID"]), "Length"))
                            {
                                Height = Convert.ToInt32(Row["Length"]);
                                QueryString += " AND Height = '" + Height.ToString() + "'";
                            }

                            if (DecorCatalog.HasParameter(Convert.ToInt32(Row["ProductID"]), "Width"))
                            {
                                Width = Convert.ToInt32(Row["Width"]);
                                QueryString += " AND Width = '" + Width.ToString() + "'";
                            }

                            DataRow[] dRow = DecorResultDataTable[i].Select(QueryString);

                            if (dRow.Count() == 0)
                            {
                                DataRow NewRow = DecorResultDataTable[i].NewRow();

                                NewRow["Product"] = Product;
                                NewRow["ManagerName"] = ManagerName;
                                NewRow["Client"] = ClientName;
                                NewRow["ClientGroupName"] = ClientGroupName;
                                NewRow["DocNumber"] = Row["DocNumber"].ToString();
                                NewRow["DispatchDateTime"] = Row["DispatchDateTime"].ToString();

                                if (DecorCatalog.HasParameter(Convert.ToInt32(Row["ProductID"]), "Height"))
                                {
                                    if (Convert.ToInt32(Row["Height"]) != -1)
                                        NewRow["Height"] = Row["Height"];
                                    if (Convert.ToInt32(Row["Length"]) != -1)
                                        NewRow["Height"] = Row["Length"];
                                }
                                if (DecorCatalog.HasParameter(Convert.ToInt32(Row["ProductID"]), "Length"))
                                {
                                    if (Convert.ToInt32(Row["Length"]) != -1)
                                        NewRow["Height"] = Row["Length"];
                                    if (Convert.ToInt32(Row["Height"]) != -1)
                                        NewRow["Height"] = Row["Height"];
                                }
                                if (DecorCatalog.HasParameter(Convert.ToInt32(Row["ProductID"]), "Width"))
                                    NewRow["Width"] = Row["Width"];

                                if (DecorCatalog.HasParameter(Convert.ToInt32(Row["ProductID"]), "ColorID"))
                                    NewRow["Color"] = Color;

                                NewRow["Patina"] = Patina;
                                NewRow["Count"] = Row["Count"];
                                NewRow["IsSample"] = Row["IsSample"];
                                NewRow["AvgPrice"] = decimal.Round(AvgPrice, 2, MidpointRounding.AwayFromZero);
                                NewRow["SumCost"] = decimal.Round(Convert.ToDecimal(Row["Cost"]), 3, MidpointRounding.AwayFromZero);

                                DecorResultDataTable[i].Rows.Add(NewRow);
                            }
                            else
                            {
                                dRow[0]["Count"] = Convert.ToDecimal(dRow[0]["Count"]) + Count;
                                dRow[0]["SumCost"] = decimal.Round(Convert.ToDecimal(dRow[0]["SumCost"]) + DecorCost, 3, MidpointRounding.AwayFromZero);
                            }
                        }
                    }
                }
            }
            else
            {
                for (int i = 0; i < DecorCatalog.DecorProductsCount; i++)
                {
                    using (DataView DV = new DataView(DecorOrdersDataTable,
                        "ProductID = " + Convert.ToInt32(TempDecorProductsDT.Rows[i]["ProductID"]),
                        string.Empty, DataViewRowState.CurrentRows))
                    {
                        Table = DV.ToTable(true, new string[] { "DecorID", "ColorID", "PatinaID", "MeasureID", "Length", "Height", "Width", "ZOVClientID", "DocNumber" });
                    }

                    for (int j = 0; j < Table.Rows.Count; j++)
                    {
                        QueryString = "ProductID=" + Convert.ToInt32(TempDecorProductsDT.Rows[i]["ProductID"]) +
                            " AND DecorID=" + Convert.ToInt32(Table.Rows[j]["DecorID"]) +
                            " AND ColorID=" + Convert.ToInt32(Table.Rows[j]["ColorID"]) +
                            " AND PatinaID=" + Convert.ToInt32(Table.Rows[j]["PatinaID"]) +
                            " AND MeasureID=" + Convert.ToInt32(Table.Rows[j]["MeasureID"]) +
                            " AND Length=" + Convert.ToInt32(Table.Rows[j]["Length"]) +
                            " AND Height=" + Convert.ToInt32(Table.Rows[j]["Height"]) +
                            " AND Width=" + Convert.ToInt32(Table.Rows[j]["Width"]) +
                            " AND DocNumber = '" + Table.Rows[j]["DocNumber"].ToString() + "'" +
                            " AND ZOVClientID=" + Convert.ToInt32(Table.Rows[j]["ZOVClientID"]);

                        DataRow[] Rows = DecorOrdersDataTable.Select(QueryString);

                        if (Rows.Count() == 0)
                            continue;

                        foreach (DataRow Row in Rows)
                        {
                            SumCount = 0;
                            SumCost = 0;

                            foreach (DataRow row in Rows)
                            {
                                if (Convert.ToInt32(row["MeasureID"]) == 1)
                                {
                                    SumCount += Convert.ToDecimal(row["Height"]) * Convert.ToDecimal(row["Width"]) * Convert.ToInt32(row["Count"]) / 1000000;
                                }

                                if (Convert.ToInt32(row["MeasureID"]) == 3)
                                {
                                    SumCount += Convert.ToInt32(row["Count"]);
                                }

                                if (Convert.ToInt32(row["MeasureID"]) == 2)
                                {
                                    //нет параметра "высота"
                                    if (row["Height"].ToString() == "-1")
                                        SumCount += Convert.ToDecimal(row["Length"]) * Convert.ToDecimal(row["Count"]) / 1000;
                                    else
                                        SumCount += Convert.ToDecimal(row["Height"]) * Convert.ToDecimal(row["Count"]) / 1000;
                                }
                                SumCost += Convert.ToDecimal(row["Cost"]);
                            }

                            Product = TempDecorProductsDT.Rows[i]["ProductName"].ToString() + " " +
                                DecorCatalog.GetItemName(Convert.ToInt32(Row["DecorID"]));
                            Count = Convert.ToInt32(Row["Count"]);
                            DecorCost = Convert.ToDecimal(Row["Cost"]);

                            DecorCost = DecorCost * 100 / 120;

                            if (SumCount != 0)
                                AvgPrice = SumCost / SumCount;

                            ManagerName = GetZOVManagerName(Convert.ToInt32(Row["ZOVClientID"]));
                            ClientName = GetZOVClientName(Convert.ToInt32(Row["ZOVClientID"]));
                            ClientGroupName = GetZOVClientGroupName(Convert.ToInt32(Row["ZOVClientID"]));

                            QueryString = "Product = '" + Product + "'";

                            if (DecorCatalog.HasParameter(Convert.ToInt32(Row["ProductID"]), "ColorID"))
                            {
                                Color = GetColorName(Convert.ToInt32(Row["ColorID"]));
                                QueryString += " AND Color = '" + Color.ToString() + "'";
                                Patina = "-";
                                if (Convert.ToInt32(Row["PatinaID"]) != -1)
                                {
                                    Patina = GetPatinaName(Convert.ToInt32(Row["PatinaID"]));
                                    QueryString += $" AND Patina = '{ Patina.ToString() }'";
                                }
                            }

                            if (DecorCatalog.HasParameter(Convert.ToInt32(Row["ProductID"]), "Height"))
                            {
                                Height = Convert.ToInt32(Row["Height"]);
                                QueryString += " AND Height = '" + Height.ToString() + "'";
                            }

                            if (DecorCatalog.HasParameter(Convert.ToInt32(Row["ProductID"]), "Length"))
                            {
                                Height = Convert.ToInt32(Row["Length"]);
                                QueryString += " AND Height = '" + Height.ToString() + "'";
                            }

                            if (DecorCatalog.HasParameter(Convert.ToInt32(Row["ProductID"]), "Width"))
                            {
                                Width = Convert.ToInt32(Row["Width"]);
                                QueryString += " AND Width = '" + Width.ToString() + "'";
                            }

                            DataRow[] dRow = DecorResultDataTable[i].Select(QueryString);

                            if (dRow.Count() == 0)
                            {
                                DataRow NewRow = DecorResultDataTable[i].NewRow();

                                NewRow["Product"] = Product;
                                NewRow["ManagerName"] = ManagerName;
                                NewRow["Client"] = ClientName;
                                NewRow["ClientGroupName"] = ClientGroupName;
                                NewRow["DocNumber"] = Row["DocNumber"].ToString();

                                if (DecorCatalog.HasParameter(Convert.ToInt32(Row["ProductID"]), "Height"))
                                {
                                    if (Convert.ToInt32(Row["Height"]) != -1)
                                        NewRow["Height"] = Row["Height"];
                                    if (Convert.ToInt32(Row["Length"]) != -1)
                                        NewRow["Height"] = Row["Length"];
                                }
                                if (DecorCatalog.HasParameter(Convert.ToInt32(Row["ProductID"]), "Length"))
                                {
                                    if (Convert.ToInt32(Row["Length"]) != -1)
                                        NewRow["Height"] = Row["Length"];
                                    if (Convert.ToInt32(Row["Height"]) != -1)
                                        NewRow["Height"] = Row["Height"];
                                }
                                if (DecorCatalog.HasParameter(Convert.ToInt32(Row["ProductID"]), "Width"))
                                    NewRow["Width"] = Row["Width"];

                                if (DecorCatalog.HasParameter(Convert.ToInt32(Row["ProductID"]), "ColorID"))
                                    NewRow["Color"] = Color;

                                NewRow["Patina"] = Patina;
                                NewRow["Count"] = Row["Count"];
                                NewRow["AvgPrice"] = decimal.Round(AvgPrice, 2, MidpointRounding.AwayFromZero);
                                NewRow["SumCost"] = decimal.Round(Convert.ToDecimal(Row["Cost"]), 3, MidpointRounding.AwayFromZero); ;

                                DecorResultDataTable[i].Rows.Add(NewRow);
                            }
                            else
                            {
                                dRow[0]["Count"] = Convert.ToDecimal(dRow[0]["Count"]) + Count;
                                dRow[0]["SumCost"] = decimal.Round(Convert.ToDecimal(dRow[0]["SumCost"]) + DecorCost, 3, MidpointRounding.AwayFromZero);
                            }
                        }
                    }
                }
            }

            TempDecorProductsDT.Dispose();
        }

        private void FillZFronts(DataTable FrontsOrdersDataTable)
        {
            if (FrontsOrdersDataTable.Rows.Count < 1)
                return;

            string Front = string.Empty;
            string FrameColor = string.Empty;
            string TechnoColor = string.Empty;
            string Patina = string.Empty;
            string InsetType = string.Empty;
            string InsetColor = string.Empty;
            string TechnoInsetType = string.Empty;
            string TechnoInsetColor = string.Empty;

            string ManagerName = "Менеджер ЗОВ";
            string ClientName = "Клиент ЗОВ";
            string ClientGroupName = "Группа ЗОВ";

            string QueryString = string.Empty;

            decimal FrontCost = 0;
            decimal AvgPrice = 0;
            decimal FrontSquare = 0;
            int FrontCount = 0;

            DataTable Table = new DataTable();

            DataColumnCollection columns = FrontsOrdersDataTable.Columns;
            if (columns.Contains("DispatchDateTime"))
            {
                using (DataView DV = new DataView(FrontsOrdersDataTable, string.Empty, string.Empty, DataViewRowState.CurrentRows))
                {
                    Table = DV.ToTable(true, new string[] { "FrontID",
                    "ColorID", "TechnoColorID", "PatinaID","InsetTypeID", "InsetColorID", "TechnoInsetTypeID",
                        "TechnoInsetColorID", "Height", "Width", "IsSample", "ZOVClientID", "DocNumber", "DispatchDateTime"});
                }

                for (int i = 0; i < Table.Rows.Count; i++)
                {
                    QueryString = "FrontID=" + Convert.ToInt32(Table.Rows[i]["FrontID"]) +
                        " AND ColorID=" + Convert.ToInt32(Table.Rows[i]["ColorID"]) +
                        " AND TechnoColorID=" + Convert.ToInt32(Table.Rows[i]["TechnoColorID"]) +
                        " AND PatinaID=" + Convert.ToInt32(Table.Rows[i]["PatinaID"]) +
                        " AND InsetTypeID=" + Convert.ToInt32(Table.Rows[i]["InsetTypeID"]) +
                        " AND InsetColorID=" + Convert.ToInt32(Table.Rows[i]["InsetColorID"]) +
                        " AND TechnoInsetTypeID=" + Convert.ToInt32(Table.Rows[i]["TechnoInsetTypeID"]) +
                        " AND TechnoInsetColorID=" + Convert.ToInt32(Table.Rows[i]["TechnoInsetColorID"]) +
                        " AND DocNumber = '" + Table.Rows[i]["DocNumber"].ToString() + "'" +
                        " AND DispatchDateTime = '" + Table.Rows[i]["DispatchDateTime"].ToString() + "'" +
                        " AND Height = '" + Table.Rows[i]["Height"].ToString() + "'" +
                        " AND Width = '" + Table.Rows[i]["Width"].ToString() + "' AND ZOVClientID=" + Convert.ToInt32(Table.Rows[i]["ZOVClientID"]) +
                        " AND IsSample=" + Convert.ToInt32(Table.Rows[i]["IsSample"]);

                    DataRow[] Rows = FrontsOrdersDataTable.Select(QueryString);
                    if (Rows.Count() != 0)
                    {
                        foreach (DataRow row in Rows)
                        {
                            FrontCost += Convert.ToDecimal(row["Cost"]);
                            FrontSquare += Convert.ToDecimal(row["Square"]);
                            FrontCount += Convert.ToInt32(row["Count"]);
                        }

                        FrontCost = FrontCost * 100 / 120;
                        if (Convert.ToInt32(Rows[0]["MeasureID"]) == 1)
                        {
                            if (FrontSquare != 0)
                                AvgPrice = FrontCost / FrontSquare;
                        }
                        if (Convert.ToInt32(Rows[0]["MeasureID"]) == 3)
                        {
                            if (FrontCount != 0)
                                AvgPrice = FrontCost / FrontCount;
                        }
                        Front = GetFrontName(Convert.ToInt32(Table.Rows[i]["FrontID"]));
                        FrameColor = GetColorName(Convert.ToInt32(Table.Rows[i]["ColorID"]));
                        TechnoColor = GetColorName(Convert.ToInt32(Table.Rows[i]["TechnoColorID"]));
                        Patina = GetPatinaName(Convert.ToInt32(Table.Rows[i]["PatinaID"]));
                        InsetType = GetInsetTypeName(Convert.ToInt32(Table.Rows[i]["InsetTypeID"]));
                        InsetColor = GetInsetColorName(Convert.ToInt32(Table.Rows[i]["InsetColorID"]));
                        TechnoInsetType = GetInsetTypeName(Convert.ToInt32(Table.Rows[i]["TechnoInsetTypeID"]));
                        TechnoInsetColor = GetInsetColorName(Convert.ToInt32(Table.Rows[i]["TechnoInsetColorID"]));

                        ManagerName = GetZOVManagerName(Convert.ToInt32(Table.Rows[i]["ZOVClientID"]));
                        ClientName = GetZOVClientName(Convert.ToInt32(Table.Rows[i]["ZOVClientID"]));
                        ClientGroupName = GetZOVClientGroupName(Convert.ToInt32(Table.Rows[i]["ZOVClientID"]));
                        //if (FrontType == "Прямой")
                        //    FrontType = string.Empty;

                        DataRow NewRow = FrontsResultDataTable.NewRow();
                        NewRow["Front"] = Front;
                        NewRow["Patina"] = Patina;
                        NewRow["FrameColor"] = FrameColor;
                        NewRow["TechnoColor"] = TechnoColor;
                        NewRow["InsetType"] = InsetType;
                        NewRow["InsetColor"] = InsetColor;
                        NewRow["TechnoInsetType"] = TechnoInsetType;
                        NewRow["TechnoInsetColor"] = TechnoInsetColor;
                        NewRow["IsSample"] = Convert.ToInt32(Table.Rows[i]["IsSample"]);
                        NewRow["Height"] = Convert.ToInt32(Table.Rows[i]["Height"]);
                        NewRow["Width"] = Convert.ToInt32(Table.Rows[i]["Width"]);
                        NewRow["Square"] = decimal.Round(FrontSquare, 3, MidpointRounding.AwayFromZero);
                        NewRow["Count"] = FrontCount;
                        NewRow["ManagerName"] = ManagerName;
                        NewRow["Client"] = ClientName;
                        NewRow["ClientGroupName"] = ClientGroupName;
                        NewRow["DocNumber"] = Table.Rows[i]["DocNumber"].ToString();
                        NewRow["DispatchDateTime"] = Table.Rows[i]["DispatchDateTime"].ToString();
                        NewRow["AvgPrice"] = decimal.Round(AvgPrice, 2, MidpointRounding.AwayFromZero);
                        NewRow["SumCost"] = decimal.Round(FrontCost, 3, MidpointRounding.AwayFromZero);
                        FrontsResultDataTable.Rows.Add(NewRow);

                        ClientName = string.Empty;
                        AvgPrice = 0;
                        FrontCost = 0;
                        FrontSquare = 0;
                        FrontCount = 0;
                    }
                }
            }
            else
            {
                using (DataView DV = new DataView(FrontsOrdersDataTable, string.Empty, string.Empty, DataViewRowState.CurrentRows))
                {
                    Table = DV.ToTable(true, new string[] { "FrontID",
                    "ColorID", "TechnoColorID", "PatinaID","InsetTypeID", "InsetColorID", "TechnoInsetTypeID", "TechnoInsetColorID", "Height", "Width", "ZOVClientID", "DocNumber"});
                }

                for (int i = 0; i < Table.Rows.Count; i++)
                {
                    QueryString = "FrontID=" + Convert.ToInt32(Table.Rows[i]["FrontID"]) +
                        " AND ColorID=" + Convert.ToInt32(Table.Rows[i]["ColorID"]) +
                        " AND TechnoColorID=" + Convert.ToInt32(Table.Rows[i]["TechnoColorID"]) +
                        " AND PatinaID=" + Convert.ToInt32(Table.Rows[i]["PatinaID"]) +
                        " AND InsetTypeID=" + Convert.ToInt32(Table.Rows[i]["InsetTypeID"]) +
                        " AND InsetColorID=" + Convert.ToInt32(Table.Rows[i]["InsetColorID"]) +
                        " AND TechnoInsetTypeID=" + Convert.ToInt32(Table.Rows[i]["TechnoInsetTypeID"]) +
                        " AND TechnoInsetColorID=" + Convert.ToInt32(Table.Rows[i]["TechnoInsetColorID"]) +
                        " AND DocNumber = '" + Table.Rows[i]["DocNumber"].ToString() + "'" +
                        " AND Height = '" + Table.Rows[i]["Height"].ToString() + "'" +
                        " AND Width = '" + Table.Rows[i]["Width"].ToString() + "' AND ZOVClientID=" + Convert.ToInt32(Table.Rows[i]["ZOVClientID"]);

                    DataRow[] Rows = FrontsOrdersDataTable.Select(QueryString);
                    if (Rows.Count() != 0)
                    {
                        foreach (DataRow row in Rows)
                        {
                            FrontCost += Convert.ToDecimal(row["Cost"]);
                            FrontSquare += Convert.ToDecimal(row["Square"]);
                            FrontCount += Convert.ToInt32(row["Count"]);
                        }

                        FrontCost = FrontCost * 100 / 120;
                        if (Convert.ToInt32(Rows[0]["MeasureID"]) == 1)
                        {
                            if (FrontSquare != 0)
                                AvgPrice = FrontCost / FrontSquare;
                        }
                        if (Convert.ToInt32(Rows[0]["MeasureID"]) == 3)
                        {
                            if (FrontCount != 0)
                                AvgPrice = FrontCost / FrontCount;
                        }
                        Front = GetFrontName(Convert.ToInt32(Table.Rows[i]["FrontID"]));
                        FrameColor = GetColorName(Convert.ToInt32(Table.Rows[i]["ColorID"]));
                        TechnoColor = GetColorName(Convert.ToInt32(Table.Rows[i]["TechnoColorID"]));
                        Patina = GetPatinaName(Convert.ToInt32(Table.Rows[i]["PatinaID"]));
                        InsetType = GetInsetTypeName(Convert.ToInt32(Table.Rows[i]["InsetTypeID"]));
                        InsetColor = GetInsetColorName(Convert.ToInt32(Table.Rows[i]["InsetColorID"]));
                        TechnoInsetType = GetInsetTypeName(Convert.ToInt32(Table.Rows[i]["TechnoInsetTypeID"]));
                        TechnoInsetColor = GetInsetColorName(Convert.ToInt32(Table.Rows[i]["TechnoInsetColorID"]));

                        ManagerName = GetZOVManagerName(Convert.ToInt32(Table.Rows[i]["ZOVClientID"]));
                        ClientName = GetZOVClientName(Convert.ToInt32(Table.Rows[i]["ZOVClientID"]));
                        ClientGroupName = GetZOVClientGroupName(Convert.ToInt32(Table.Rows[i]["ZOVClientID"]));
                        //if (FrontType == "Прямой")
                        //    FrontType = string.Empty;

                        DataRow NewRow = FrontsResultDataTable.NewRow();
                        NewRow["Front"] = Front;
                        NewRow["Patina"] = Patina;
                        NewRow["FrameColor"] = FrameColor;
                        NewRow["TechnoColor"] = TechnoColor;
                        NewRow["InsetType"] = InsetType;
                        NewRow["InsetColor"] = InsetColor;
                        NewRow["TechnoInsetType"] = TechnoInsetType;
                        NewRow["TechnoInsetColor"] = TechnoInsetColor;
                        NewRow["Height"] = Convert.ToInt32(Table.Rows[i]["Height"]);
                        NewRow["Width"] = Convert.ToInt32(Table.Rows[i]["Width"]);
                        NewRow["Square"] = decimal.Round(FrontSquare, 3, MidpointRounding.AwayFromZero);
                        NewRow["Count"] = FrontCount;
                        NewRow["ManagerName"] = ManagerName;
                        NewRow["Client"] = ClientName;
                        NewRow["ClientGroupName"] = ClientGroupName;
                        NewRow["DocNumber"] = Table.Rows[i]["DocNumber"].ToString();
                        NewRow["AvgPrice"] = decimal.Round(AvgPrice, 2, MidpointRounding.AwayFromZero);
                        NewRow["SumCost"] = decimal.Round(FrontCost, 3, MidpointRounding.AwayFromZero);
                        FrontsResultDataTable.Rows.Add(NewRow);

                        ClientName = string.Empty;
                        AvgPrice = 0;
                        FrontCost = 0;
                        FrontSquare = 0;
                        FrontCount = 0;
                    }
                }
            }
        }

        private void FrontsReportOrderNumber(HSSFWorkbook hssfworkbook, HSSFCellStyle HeaderStyle, 
            HSSFFont SimpleFont, HSSFCellStyle SimpleCellStyle, HSSFCellStyle cellStyle,
            DataTable FrontsOrdersDataTable, bool ZOV)
        {
            int RowIndex = 0;

            HSSFSheet sheet1 = hssfworkbook.CreateSheet("Фасады");
            sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

            sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
            sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
            sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
            sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

            int DisplayIndex = 0;

            if (ZOV)
            {
                sheet1.SetColumnWidth(DisplayIndex++, 25 * 256);
                sheet1.SetColumnWidth(DisplayIndex++, 20 * 256);
                sheet1.SetColumnWidth(DisplayIndex++, 20 * 256);
                sheet1.SetColumnWidth(DisplayIndex++, 15 * 256);
            }
            else
            {
                sheet1.SetColumnWidth(DisplayIndex++, 35 * 256);
                sheet1.SetColumnWidth(DisplayIndex++, 20 * 256);
                sheet1.SetColumnWidth(DisplayIndex++, 20 * 256);
            }

            sheet1.SetColumnWidth(DisplayIndex++, 32 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 28 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 18 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 14 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 17 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 18 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 15 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 16 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 6 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 6 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 8 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 8 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 10 * 256);

            sheet1.SetColumnWidth(DisplayIndex++, 16 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 14 * 256);

            decimal Square = 0;
            int FrontsCount = 0;
            int CurvedCount = 0;

            FillFrontsOrderNumber(FrontsOrdersDataTable, ZOV);

            if (FrontsResultDataTable.Rows.Count != 0)
            {
                DisplayIndex = 0;

                HSSFCell cell4;
                if (ZOV)
                {
                    cell4 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Группа");
                    cell4.CellStyle = HeaderStyle;
                    cell4 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Клиент");
                    cell4.CellStyle = HeaderStyle;
                    cell4 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "№ заказа");
                    cell4.CellStyle = HeaderStyle;
                    cell4 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Дата отгр");
                    cell4.CellStyle = HeaderStyle;
                }
                else
                {
                    cell4 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Клиент");
                    cell4.CellStyle = HeaderStyle;
                    cell4 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "№ заказа");
                    cell4.CellStyle = HeaderStyle;
                    cell4 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Клиент ЗОВ");
                    cell4.CellStyle = HeaderStyle;
                }
                cell4 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Фасад");
                cell4.CellStyle = HeaderStyle;
                cell4 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Цвет профиля");
                cell4.CellStyle = HeaderStyle;
                cell4 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Цвет профиля-2");
                cell4.CellStyle = HeaderStyle;
                cell4 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Патина");
                cell4.CellStyle = HeaderStyle;
                cell4 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Тип наполнителя");
                cell4.CellStyle = HeaderStyle;
                cell4 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Цвет наполнителя");
                cell4.CellStyle = HeaderStyle;
                cell4 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Тип наполнителя-2");
                cell4.CellStyle = HeaderStyle;
                cell4 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Цвет наполнителя-2");
                cell4.CellStyle = HeaderStyle;
                cell4 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Выс.");
                cell4.CellStyle = HeaderStyle;
                cell4 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Шир.");
                cell4.CellStyle = HeaderStyle;
                cell4 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Площ.");
                cell4.CellStyle = HeaderStyle;
                cell4 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Кол.");
                cell4.CellStyle = HeaderStyle;
                cell4 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Образец");
                cell4.CellStyle = HeaderStyle;
                if (Security.PriceAccess)
                {
                    HSSFCell cell15 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Ср.цена, евро");
                    cell15.CellStyle = HeaderStyle;
                    HSSFCell cell16 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Сумма, евро");
                    cell16.CellStyle = HeaderStyle;
                }
                cell4 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Менеджер");
                cell4.CellStyle = HeaderStyle;
                if (!ZOV)
                {
                    cell4 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Страна");
                    cell4.CellStyle = HeaderStyle;
                }

                RowIndex++;

                Square = GetSquare(FrontsResultDataTable);
                FrontsCount = GetCount(FrontsResultDataTable, false);
                CurvedCount = GetCount(FrontsResultDataTable, true);

                DataTable dt = FrontsResultDataTable.Copy();
                if (ZOV)
                {
                    dt.Columns.Remove("CountryName");
                    dt.Columns.Remove("ZOVClientName");
                }
                else
                {
                    dt.Columns.Remove("ClientGroupName");
                    dt.Columns.Remove("DocNumber");
                    dt.Columns.Remove("DispatchDateTime");
                }

                int ColumnCount = dt.Columns.Count;

                if (!Security.PriceAccess)
                {
                    ColumnCount = dt.Columns.Count - 2;
                }

                //вывод заказов фасадов
                for (int x = 0; x < dt.Rows.Count; x++)
                {
                    if (dt.Rows.Count == 0)
                        break;

                    for (int y = 0; y < ColumnCount; y++)
                    {
                        //if (ZOV && FrontsResultDataTable.Columns[y].ColumnName == "CountryName")
                        //    continue;

                        //if (!ZOV && FrontsResultDataTable.Columns[y].ColumnName == "ClientGroupName")
                        //    continue;

                        Type t = dt.Rows[x][y].GetType();

                        if (t.Name == "Decimal")
                        {
                            HSSFCell cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(Convert.ToDouble(dt.Rows[x][y]));

                            cell.CellStyle = cellStyle;
                            continue;
                        }
                        if (t.Name == "Int32")
                        {
                            HSSFCell cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(Convert.ToInt32(dt.Rows[x][y]));
                            cell.CellStyle = SimpleCellStyle;
                            continue;
                        }

                        if (t.Name == "String" || t.Name == "DBNull")
                        {
                            HSSFCell cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(dt.Rows[x][y].ToString());
                            cell.CellStyle = SimpleCellStyle;
                            continue;
                        }
                    }
                    RowIndex++;
                }

                RowIndex++;

                HSSFCellStyle cellStyle1 = hssfworkbook.CreateCellStyle();
                cellStyle1.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.000");
                cellStyle1.SetFont(SimpleFont);

                HSSFCell cell17 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Итого: ");
                cell17.CellStyle = cellStyle1;
                HSSFCell cell18 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, "всего фасадов: " + FrontsCount + " шт.");
                cell18.CellStyle = cellStyle1;
                HSSFCell cell19 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, "в том числе гнутых: " + CurvedCount + " шт.");
                cell19.CellStyle = cellStyle1;

                if (Square > 0)
                {
                    HSSFCell cell20 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1,
                        "площадь: " + decimal.Round(Square, 3, MidpointRounding.AwayFromZero) + " м.кв.");

                    cell20.CellStyle = cellStyle1;
                }

                RowIndex++;
            }
        }

        private void FrontsReport(HSSFWorkbook hssfworkbook, HSSFCellStyle HeaderStyle, 
            HSSFFont SimpleFont, HSSFCellStyle SimpleCellStyle, HSSFCellStyle cellStyle,
            DataTable FrontsOrdersDataTable, bool ZOV)
        {
            int RowIndex = 0;

            HSSFSheet sheet1 = hssfworkbook.CreateSheet("Фасады");
            sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

            sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
            sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
            sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
            sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

            int DisplayIndex = 0;

            if (ZOV)
            {
                sheet1.SetColumnWidth(DisplayIndex++, 25 * 256);
                sheet1.SetColumnWidth(DisplayIndex++, 20 * 256);
                sheet1.SetColumnWidth(DisplayIndex++, 20 * 256);
                sheet1.SetColumnWidth(DisplayIndex++, 15 * 256);
            }
            else
            {
                sheet1.SetColumnWidth(DisplayIndex++, 35 * 256);
                sheet1.SetColumnWidth(DisplayIndex++, 20 * 256);
            }

            sheet1.SetColumnWidth(DisplayIndex++, 32 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 28 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 18 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 14 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 17 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 18 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 15 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 16 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 6 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 6 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 8 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 8 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 10 * 256);

            sheet1.SetColumnWidth(DisplayIndex++, 16 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 14 * 256);

            decimal Square = 0;
            int FrontsCount = 0;
            int CurvedCount = 0;

            FillFronts(FrontsOrdersDataTable, ZOV);

            if (FrontsResultDataTable.Rows.Count != 0)
            {
                DisplayIndex = 0;

                HSSFCell cell4;
                if (ZOV)
                {
                    cell4 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Группа");
                    cell4.CellStyle = HeaderStyle;
                    cell4 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Клиент");
                    cell4.CellStyle = HeaderStyle;
                    cell4 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "№ заказа");
                    cell4.CellStyle = HeaderStyle;
                    cell4 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Дата отгр");
                    cell4.CellStyle = HeaderStyle;
                }
                else
                {
                    cell4 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Клиент");
                    cell4.CellStyle = HeaderStyle;
                    cell4 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Клиент ЗОВ");
                    cell4.CellStyle = HeaderStyle;
                }
                cell4 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Фасад");
                cell4.CellStyle = HeaderStyle;
                cell4 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Цвет профиля");
                cell4.CellStyle = HeaderStyle;
                cell4 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Цвет профиля-2");
                cell4.CellStyle = HeaderStyle;
                cell4 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Патина");
                cell4.CellStyle = HeaderStyle;
                cell4 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Тип наполнителя");
                cell4.CellStyle = HeaderStyle;
                cell4 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Цвет наполнителя");
                cell4.CellStyle = HeaderStyle;
                cell4 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Тип наполнителя-2");
                cell4.CellStyle = HeaderStyle;
                cell4 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Цвет наполнителя-2");
                cell4.CellStyle = HeaderStyle;
                cell4 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Выс.");
                cell4.CellStyle = HeaderStyle;
                cell4 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Шир.");
                cell4.CellStyle = HeaderStyle;
                cell4 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Площ.");
                cell4.CellStyle = HeaderStyle;
                cell4 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Кол.");
                cell4.CellStyle = HeaderStyle;
                cell4 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Образец");
                cell4.CellStyle = HeaderStyle;
                if (Security.PriceAccess)
                {
                    HSSFCell cell15 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Ср.цена, евро");
                    cell15.CellStyle = HeaderStyle;
                    HSSFCell cell16 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Сумма, евро");
                    cell16.CellStyle = HeaderStyle;
                }
                cell4 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Менеджер");
                cell4.CellStyle = HeaderStyle;
                if (!ZOV)
                {
                    cell4 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Страна");
                    cell4.CellStyle = HeaderStyle;
                }

                RowIndex++;

                Square = GetSquare(FrontsResultDataTable);
                FrontsCount = GetCount(FrontsResultDataTable, false);
                CurvedCount = GetCount(FrontsResultDataTable, true);

                DataTable dt = FrontsResultDataTable.Copy();
                if (ZOV)
                {
                    dt.Columns.Remove("CountryName");
                    dt.Columns.Remove("ZOVClientName");
                }
                else
                {
                    dt.Columns.Remove("ClientGroupName");
                    dt.Columns.Remove("DocNumber");
                    dt.Columns.Remove("DispatchDateTime");
                }

                int ColumnCount = dt.Columns.Count;

                if (!Security.PriceAccess)
                {
                    ColumnCount = dt.Columns.Count - 2;
                }

                //вывод заказов фасадов
                for (int x = 0; x < dt.Rows.Count; x++)
                {
                    if (dt.Rows.Count == 0)
                        break;

                    for (int y = 0; y < ColumnCount; y++)
                    {
                        //if (ZOV && FrontsResultDataTable.Columns[y].ColumnName == "CountryName")
                        //    continue;

                        //if (!ZOV && FrontsResultDataTable.Columns[y].ColumnName == "ClientGroupName")
                        //    continue;

                        Type t = dt.Rows[x][y].GetType();

                        if (t.Name == "Decimal")
                        {
                            HSSFCell cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(Convert.ToDouble(dt.Rows[x][y]));

                            cell.CellStyle = cellStyle;
                            continue;
                        }
                        if (t.Name == "Int32")
                        {
                            HSSFCell cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(Convert.ToInt32(dt.Rows[x][y]));
                            cell.CellStyle = SimpleCellStyle;
                            continue;
                        }

                        if (t.Name == "String" || t.Name == "DBNull")
                        {
                            HSSFCell cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(dt.Rows[x][y].ToString());
                            cell.CellStyle = SimpleCellStyle;
                            continue;
                        }
                    }
                    RowIndex++;
                }

                RowIndex++;

                HSSFCellStyle cellStyle1 = hssfworkbook.CreateCellStyle();
                cellStyle1.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.000");
                cellStyle1.SetFont(SimpleFont);

                HSSFCell cell17 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Итого: ");
                cell17.CellStyle = cellStyle1;
                HSSFCell cell18 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, "всего фасадов: " + FrontsCount + " шт.");
                cell18.CellStyle = cellStyle1;
                HSSFCell cell19 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, "в том числе гнутых: " + CurvedCount + " шт.");
                cell19.CellStyle = cellStyle1;

                if (Square > 0)
                {
                    HSSFCell cell20 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1,
                        "площадь: " + decimal.Round(Square, 3, MidpointRounding.AwayFromZero) + " м.кв.");

                    cell20.CellStyle = cellStyle1;
                }

                RowIndex++;
            }
        }

        private string GetClientName(int ClientID)
        {
            DataRow[] Rows = ClientsDataTable.Select("ClientID = " + ClientID);
            if (Rows.Count() > 0)
                return Rows[0]["ClientName"].ToString();
            else
                return string.Empty;
        }

        private void GetColorsDT()
        {
            FrameColorsDataTable.Columns.Add(new DataColumn("ColorID", Type.GetType("System.Int64")));
            FrameColorsDataTable.Columns.Add(new DataColumn("ColorName", Type.GetType("System.String")));
            string SelectCommand = @"SELECT TechStoreID, TechStoreName FROM TechStore
                WHERE TechStoreSubGroupID IN (SELECT TechStoreSubGroupID FROM TechStoreSubGroups WHERE TechStoreGroupID = 11)
                ORDER BY TechStoreName";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.CatalogConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    DA.Fill(DT);
                    {
                        DataRow NewRow = FrameColorsDataTable.NewRow();
                        NewRow["ColorID"] = -1;
                        NewRow["ColorName"] = "-";
                        FrameColorsDataTable.Rows.Add(NewRow);
                    }
                    {
                        DataRow NewRow = FrameColorsDataTable.NewRow();
                        NewRow["ColorID"] = 0;
                        NewRow["ColorName"] = "на выбор";
                        FrameColorsDataTable.Rows.Add(NewRow);
                    }
                    for (int i = 0; i < DT.Rows.Count; i++)
                    {
                        DataRow NewRow = FrameColorsDataTable.NewRow();
                        NewRow["ColorID"] = Convert.ToInt64(DT.Rows[i]["TechStoreID"]);
                        NewRow["ColorName"] = DT.Rows[i]["TechStoreName"].ToString();
                        FrameColorsDataTable.Rows.Add(NewRow);
                    }
                }
            }
        }

        private int GetCount(DataTable DT, bool Curved)
        {
            int S = 0;

            foreach (DataRow Row in DT.Rows)
            {
                if (Curved)
                {
                    if (Convert.ToInt32(Row["Width"]) == -1)
                        S += Convert.ToInt32(Row["Count"]);
                }
                else
                    S += Convert.ToInt32(Row["Count"]);
            }

            return S;
        }

        private string GetCountryName(int ClientID)
        {
            DataRow[] Rows = ClientsDataTable.Select("ClientID = " + ClientID);
            if (Rows.Count() > 0)
                return Rows[0]["CountryName"].ToString();
            else
                return string.Empty;
        }

        private string GetManagerName(int ClientID)
        {
            DataRow[] Rows = ClientsDataTable.Select("ClientID = " + ClientID);
            if (Rows.Count() > 0)
                return Rows[0]["ShortName"].ToString();
            else
                return string.Empty;
        }

        private decimal GetSquare(DataTable DT)
        {
            decimal S = 0;

            foreach (DataRow Row in DT.Rows)
            {
                if (Row["Square"] != DBNull.Value)
                    S += Convert.ToDecimal(Row["Square"]);
            }

            return S;
        }

        private string GetZOVClientGroupName(int ClientID)
        {
            DataRow[] Rows = ZOVClientsDataTable.Select("ClientID = " + ClientID);
            if (Rows.Count() > 0)
                return Rows[0]["ClientGroupName"].ToString();
            else
                return string.Empty;
        }

        private string GetZOVClientName(int ClientID)
        {
            DataRow[] Rows = ZOVClientsDataTable.Select("ClientID = " + ClientID);
            if (Rows.Count() > 0)
                return Rows[0]["ClientName"].ToString();
            else
                return string.Empty;
        }

        private string GetZOVManagerName(int ClientID)
        {
            DataRow[] Rows = ZOVClientsDataTable.Select("ClientID = " + ClientID);
            if (Rows.Count() > 0)
                return Rows[0]["Name"].ToString();
            else
                return string.Empty;
        }
    }
}