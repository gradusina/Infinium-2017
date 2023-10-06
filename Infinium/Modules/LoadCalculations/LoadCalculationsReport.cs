using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;

using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Diagnostics;
using System.IO;
using ComponentFactory.Krypton.Toolkit;
using NPOI.HSSF.UserModel.Contrib;
using NPOI.HSSF.Model;

namespace Infinium.Modules.LoadCalculations
{
    public class LoadCalculationsModel
    {
        public int Index { get; set; }
        public int ClientId { get; set; }
        public string ClientName { get; set; }
        public string Manager { get; set; }
        public string Country { get; set; }
        public string City { get; set; }
        public string Email { get; set; }
        public int DelayOfPayment { get; set; }
        public string Unn { get; set; }
        public string ClientGroupName { get; set; }
        public int PriceGroup { get; set; }
        public bool Enabled { get; set; }
    }

    public class LoadCalculationsReport
    {
        private readonly HSSFCellStyle _simpleCs;
        private readonly HSSFCellStyle _simpleDecCs;
        private readonly HSSFCellStyle _simpleHeaderCs;
        private readonly HSSFCellStyle _simpleIntCs;
        private readonly HSSFWorkbook _workbook;

        public LoadCalculationsReport()
        {
            GetClients();
            PrepareColumns();

            _workbook = new HSSFWorkbook();

            var headerF = _workbook.CreateFont();
            headerF.FontHeightInPoints = 11;
            headerF.Boldweight = HSSFFont.BOLDWEIGHT_BOLD;
            headerF.FontName = "Times New Roman";

            var simpleF = _workbook.CreateFont();
            simpleF.FontHeightInPoints = 11;
            simpleF.FontName = "Times New Roman";

            var simpleBoldF = _workbook.CreateFont();
            simpleBoldF.FontHeightInPoints = 11;
            simpleBoldF.Boldweight = HSSFFont.BOLDWEIGHT_BOLD;
            simpleBoldF.FontName = "Times New Roman";

            _simpleCs = _workbook.CreateCellStyle();
            _simpleCs.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
            _simpleCs.Alignment = HSSFCellStyle.ALIGN_LEFT;
            _simpleCs.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.0");
            _simpleCs.BorderBottom = HSSFCellStyle.BORDER_THIN;
            _simpleCs.BottomBorderColor = HSSFColor.BLACK.index;
            _simpleCs.BorderLeft = HSSFCellStyle.BORDER_THIN;
            _simpleCs.LeftBorderColor = HSSFColor.BLACK.index;
            _simpleCs.BorderRight = HSSFCellStyle.BORDER_THIN;
            _simpleCs.RightBorderColor = HSSFColor.BLACK.index;
            _simpleCs.BorderTop = HSSFCellStyle.BORDER_THIN;
            _simpleCs.TopBorderColor = HSSFColor.BLACK.index;
            _simpleCs.WrapText = true;
            _simpleCs.SetFont(simpleF);

            _simpleIntCs = _workbook.CreateCellStyle();
            _simpleIntCs.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
            _simpleIntCs.Alignment = HSSFCellStyle.ALIGN_CENTER;
            _simpleIntCs.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.0");
            _simpleIntCs.BorderBottom = HSSFCellStyle.BORDER_THIN;
            _simpleIntCs.BottomBorderColor = HSSFColor.BLACK.index;
            _simpleIntCs.BorderLeft = HSSFCellStyle.BORDER_THIN;
            _simpleIntCs.LeftBorderColor = HSSFColor.BLACK.index;
            _simpleIntCs.BorderRight = HSSFCellStyle.BORDER_THIN;
            _simpleIntCs.RightBorderColor = HSSFColor.BLACK.index;
            _simpleIntCs.BorderTop = HSSFCellStyle.BORDER_THIN;
            _simpleIntCs.TopBorderColor = HSSFColor.BLACK.index;
            _simpleIntCs.WrapText = true;
            _simpleIntCs.SetFont(simpleF);

            _simpleDecCs = _workbook.CreateCellStyle();
            _simpleDecCs.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
            _simpleDecCs.Alignment = HSSFCellStyle.ALIGN_CENTER;
            _simpleDecCs.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.0");
            _simpleDecCs.BorderBottom = HSSFCellStyle.BORDER_THIN;
            _simpleDecCs.BottomBorderColor = HSSFColor.BLACK.index;
            _simpleDecCs.BorderLeft = HSSFCellStyle.BORDER_THIN;
            _simpleDecCs.LeftBorderColor = HSSFColor.BLACK.index;
            _simpleDecCs.BorderRight = HSSFCellStyle.BORDER_THIN;
            _simpleDecCs.RightBorderColor = HSSFColor.BLACK.index;
            _simpleDecCs.BorderTop = HSSFCellStyle.BORDER_THIN;
            _simpleDecCs.TopBorderColor = HSSFColor.BLACK.index;
            _simpleDecCs.SetFont(simpleF);

            _simpleHeaderCs = _workbook.CreateCellStyle();
            _simpleHeaderCs.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
            _simpleHeaderCs.Alignment = HSSFCellStyle.ALIGN_CENTER;
            _simpleHeaderCs.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            _simpleHeaderCs.BottomBorderColor = HSSFColor.BLACK.index;
            _simpleHeaderCs.BorderLeft = HSSFCellStyle.BORDER_MEDIUM;
            _simpleHeaderCs.LeftBorderColor = HSSFColor.BLACK.index;
            _simpleHeaderCs.BorderRight = HSSFCellStyle.BORDER_MEDIUM;
            _simpleHeaderCs.RightBorderColor = HSSFColor.BLACK.index;
            _simpleHeaderCs.BorderTop = HSSFCellStyle.BORDER_MEDIUM;
            _simpleHeaderCs.TopBorderColor = HSSFColor.BLACK.index;
            _simpleHeaderCs.SetFont(headerF);
        }

        public bool NeedStartFile { get; set; }
        private List<LoadCalculations.Sector> SectorsList;

        public List<LoadCalculationsModel> DataList { get; set; }

        private bool HasData => DataList.Count > 0;

        private DataTable _clientsDt;

        private readonly Dictionary<string, string> _engToRusDictionary = new();

        private void PrepareColumns()
        {
            _engToRusDictionary.Add("ClientID", "ID");
            _engToRusDictionary.Add("ClientName", "Клиент");
            _engToRusDictionary.Add("Manager", "Менеджер");
            _engToRusDictionary.Add("Country", "Страна");
            _engToRusDictionary.Add("City", "Город");
            _engToRusDictionary.Add("Email", "Email");
            _engToRusDictionary.Add("DelayOfPayment", "Отсрочка");
            _engToRusDictionary.Add("UNN", "UNN");
            _engToRusDictionary.Add("ClientGroupName", "Группа");
            _engToRusDictionary.Add("PriceGroup", "Ценовая группа");
            _engToRusDictionary.Add("Enabled", "Активен");
        }

        private void GetClients()
        {
            _clientsDt = new DataTable();

            const string selectCommand = @"SELECT Clients.ClientID, Clients.ClientName, ClientsManagers.Name as Manager, 
Countries.Name AS Country, Clients.City, Clients.Email, Clients.DelayOfPayment, Clients.UNN, 
ClientGroups.ClientGroupName, Clients.PriceGroup, Clients.Enabled FROM Clients AS Clients 
INNER JOIN ClientsManagers ON Clients.ManagerID = ClientsManagers.ManagerID 
INNER JOIN ClientGroups ON Clients.ClientGroupID = ClientGroups.ClientGroupID
INNER JOIN infiniu2_catalog.dbo.Countries AS Countries ON Clients.CountryID = Countries.CountryID order by Clients.ClientName";

            using var da = new SqlDataAdapter(selectCommand, ConnectionStrings.MarketingReferenceConnectionString);
            da.Fill(_clientsDt);
        }

        public void CreateReport()
        {
            if (!HasData) return;

            ClearWorkbook();

            const string sheetName = "Замечания";

            var sheet1 = _workbook.CreateSheet(sheetName);

            //var displayIndex = 0;
            //const int firstRowIndex = 2;
            var rowIndex = 2;

            //var row = sheet1.CreateRow(rowIndex++);

            //var headerCell = row.CreateCell(displayIndex++);
            //headerCell.SetCellValue("№ п/п");
            //headerCell.CellStyle = _simpleHeaderCs;

            //headerCell = row.CreateCell(displayIndex++);
            //headerCell.SetCellValue("Станок");
            //headerCell.CellStyle = _simpleHeaderCs;

            //headerCell = row.CreateCell(displayIndex++);
            //headerCell.SetCellValue("Дата");
            //headerCell.CellStyle = _simpleHeaderCs;

            //headerCell = row.CreateCell(displayIndex);
            //headerCell.SetCellValue("Замечание");
            //headerCell.CellStyle = _simpleHeaderCs;


            var i = 0;
            foreach (var column in _engToRusDictionary)
            {
                var cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(rowIndex), i, column.Value);
                cell.CellStyle = _simpleHeaderCs;
                i++;
            }

            foreach (var remark in DataList)
            {
                var row = sheet1.CreateRow(rowIndex);

                WriteRemark(row, remark);

                rowIndex++;
            }

            var lastRowIndex = rowIndex;

            //sheet1.SetAutoFilter(new CellRangeAddress(firstRowIndex, lastRowIndex, 1, 1));

            var displayIndex = 0;

            //if (_workbook.NumberOfSheets > 0)
            //{
            //    sheet1.AutoSizeColumn(displayIndex);
            //}

            sheet1.SetColumnWidth(3, 60 * 256);

            //sheet1.GetRow(2).Height = (short)3 * 256;
            sheet1.CreateFreezePane(0, 3);

            sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;
            sheet1.SetMargin(HSSFSheet.LeftMargin, .12);
            sheet1.SetMargin(HSSFSheet.RightMargin, .07);
            sheet1.SetMargin(HSSFSheet.TopMargin, .20);
            sheet1.SetMargin(HSSFSheet.BottomMargin, .20);

            var fileName = $"Клиенты {DateTime.Now.ToShortDateString()}";
            var tempFolder = Environment.GetEnvironmentVariable("TEMP");

            var fileInfo = GetFileInfo(tempFolder, fileName);

            SaveReport(fileInfo);

            if (NeedStartFile)
                StartFile(fileInfo);
        }

        private static FileInfo GetFileInfo(string tempFolder, string fileName)
        {
            FileInfo fileInfo = new($@"{tempFolder}\{fileName}.xls");

            var j = 1;
            while (fileInfo.Exists) fileInfo = new FileInfo($@"{tempFolder}\{fileName}({j++}).xls");

            return fileInfo;
        }

        private static void StartFile(FileSystemInfo fileInfo)
        {
            Process.Start(fileInfo.FullName);
        }

        private void ClearWorkbook()
        {
            if (_workbook.NumberOfSheets == 0)
                return;
            var sheet = _workbook.GetSheetAt(0);
            for (var i = 0; i < sheet.LastRowNum; i++)
            {
                var removingRow = sheet.GetRow(i);
                sheet.RemoveRow(removingRow);
            }
        }

        private void SaveReport(FileSystemInfo fileInfo)
        {
            FileStream newFile = new(fileInfo.FullName, FileMode.Create);
            _workbook.Write(newFile);
            newFile.Close();
        }

        private void WriteCell(HSSFRow row, int colIndex, int cellValue)
        {
            var cell = row.CreateCell(colIndex);
            cell.SetCellValue(cellValue);
            cell.CellStyle = _simpleIntCs;
        }

        private void WriteCell(HSSFRow row, int colIndex, string cellValue)
        {
            var cell = row.CreateCell(colIndex);
            cell.SetCellValue(cellValue);
            cell.CellStyle = _simpleCs;
        }

        private void WriteRemark(HSSFRow row, LoadCalculationsModel loadCalculations)
        {
            var colIndex = 0;
            
            WriteCell(row, colIndex++, loadCalculations.ClientId);
            WriteCell(row, colIndex++, loadCalculations.ClientName);
            WriteCell(row, colIndex++, loadCalculations.Manager);
            WriteCell(row, colIndex++, loadCalculations.Country);
            WriteCell(row, colIndex++, loadCalculations.City);
            WriteCell(row, colIndex++, loadCalculations.Email);
            WriteCell(row, colIndex++, loadCalculations.DelayOfPayment);
            WriteCell(row, colIndex++, loadCalculations.Unn);
            WriteCell(row, colIndex++, loadCalculations.ClientGroupName);
            WriteCell(row, colIndex++, loadCalculations.PriceGroup);
        }
    }
}