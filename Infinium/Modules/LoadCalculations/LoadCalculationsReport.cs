using DevExpress.XtraCharts.Native;
using DevExpress.XtraRichEdit.Model;

using Infinium.Modules.Marketing.Clients;

using NPOI.HSSF.Model;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.UserModel.Contrib;
using NPOI.HSSF.Util;

using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace Infinium.Modules.LoadCalculations
{

    public class LoadCalculationsReport
    {
        private readonly HSSFCellStyle _simpleCs;
        private readonly HSSFCellStyle _simpleDecCs;
        private readonly HSSFCellStyle _simpleHeader1Cs;
        private readonly HSSFCellStyle _simpleHeader2Cs;
        private readonly HSSFCellStyle _simpleHeader3Cs;
        private readonly HSSFCellStyle _simpleIntCs;
        private readonly HSSFWorkbook _workbook;

        public LoadCalculationsReport()
        {
            nfi = (NumberFormatInfo)CultureInfo.InvariantCulture.NumberFormat.Clone();
            nfi.NumberDecimalSeparator = ",";
            nfi.NumberGroupSeparator = " ";

            _workbook = new HSSFWorkbook();

            var header1F = _workbook.CreateFont();
            header1F.FontHeightInPoints = 11;
            header1F.Boldweight = HSSFFont.BOLDWEIGHT_BOLD;
            header1F.FontName = "Times New Roman";

            var header2F = _workbook.CreateFont();
            header2F.FontHeightInPoints = 12;
            header2F.FontName = "Times New Roman";

            var header3F = _workbook.CreateFont();
            header3F.FontHeightInPoints = 13;
            header3F.Boldweight = HSSFFont.BOLDWEIGHT_BOLD;
            header3F.FontName = "Times New Roman";

            var simpleF = _workbook.CreateFont();
            simpleF.FontHeightInPoints = 11;
            simpleF.FontName = "Times New Roman";

            var simpleBoldF = _workbook.CreateFont();
            simpleBoldF.FontHeightInPoints = 11;
            simpleBoldF.Boldweight = HSSFFont.BOLDWEIGHT_BOLD;
            simpleBoldF.FontName = "Times New Roman";

            _simpleCs = _workbook.CreateCellStyle();
            _simpleCs.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
            _simpleCs.Alignment = HSSFCellStyle.ALIGN_CENTER;
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

            _simpleHeader1Cs = _workbook.CreateCellStyle();
            _simpleHeader1Cs.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
            _simpleHeader1Cs.Alignment = HSSFCellStyle.ALIGN_CENTER;
            _simpleHeader1Cs.BorderBottom = HSSFCellStyle.BORDER_THIN;
            _simpleHeader1Cs.BottomBorderColor = HSSFColor.BLACK.index;
            _simpleHeader1Cs.BorderLeft = HSSFCellStyle.BORDER_THIN;
            _simpleHeader1Cs.LeftBorderColor = HSSFColor.BLACK.index;
            _simpleHeader1Cs.BorderRight = HSSFCellStyle.BORDER_THIN;
            _simpleHeader1Cs.RightBorderColor = HSSFColor.BLACK.index;
            _simpleHeader1Cs.BorderTop = HSSFCellStyle.BORDER_THIN;
            _simpleHeader1Cs.TopBorderColor = HSSFColor.BLACK.index;
            _simpleHeader1Cs.SetFont(header1F);

            _simpleHeader2Cs = _workbook.CreateCellStyle();
            _simpleHeader2Cs.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
            _simpleHeader2Cs.Alignment = HSSFCellStyle.ALIGN_CENTER;
            _simpleHeader2Cs.BorderBottom = HSSFCellStyle.BORDER_THIN;
            _simpleHeader2Cs.BottomBorderColor = HSSFColor.BLACK.index;
            _simpleHeader2Cs.BorderLeft = HSSFCellStyle.BORDER_THIN;
            _simpleHeader2Cs.LeftBorderColor = HSSFColor.BLACK.index;
            _simpleHeader2Cs.BorderRight = HSSFCellStyle.BORDER_THIN;
            _simpleHeader2Cs.RightBorderColor = HSSFColor.BLACK.index;
            _simpleHeader2Cs.BorderTop = HSSFCellStyle.BORDER_THIN;
            _simpleHeader2Cs.TopBorderColor = HSSFColor.BLACK.index;
            _simpleHeader2Cs.SetFont(header2F);

            _simpleHeader3Cs = _workbook.CreateCellStyle();
            _simpleHeader3Cs.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
            _simpleHeader3Cs.Alignment = HSSFCellStyle.ALIGN_CENTER;
            _simpleHeader3Cs.BorderBottom = HSSFCellStyle.BORDER_THIN;
            _simpleHeader3Cs.BottomBorderColor = HSSFColor.BLACK.index;
            _simpleHeader3Cs.BorderLeft = HSSFCellStyle.BORDER_THIN;
            _simpleHeader3Cs.LeftBorderColor = HSSFColor.BLACK.index;
            _simpleHeader3Cs.BorderRight = HSSFCellStyle.BORDER_THIN;
            _simpleHeader3Cs.RightBorderColor = HSSFColor.BLACK.index;
            _simpleHeader3Cs.BorderTop = HSSFCellStyle.BORDER_THIN;
            _simpleHeader3Cs.TopBorderColor = HSSFColor.BLACK.index;
            _simpleHeader3Cs.SetFont(header3F);
        }

        public bool NeedStartFile { get; set; }

        private NumberFormatInfo nfi;
        private const string DigitalFormat = "#0,0.000";

        public List<LoadCalculations.Sector> SectorsList { get; set; }

        public LoadCalculations _calculations { get; set; }

        private bool HasData => SectorsList.Count > 0;

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
            
            var sheetName = "Лист1";
            var sheet1 = _workbook.CreateSheet(sheetName);
            
            var colIndex = 4;
            var rowIndex = 2;
            var rIndex = 0;
            var row = sheet1.CreateRow(rowIndex);
            
            var notConfirmedRowIndex = 0;
            var forAgreedRowIndex = 0;
            var agreedRowIndex = 0;
            var onProductionRowIndex = 0;
            var inProductionRowIndex = 0;

            var rowIndex1 = rowIndex + 1;
            var rowIndex2 = rowIndex1 + 1;
            var rowIndex3 = rowIndex2 + 1;
            var rowIndex4 = rowIndex3 + 1;
            var rowIndex5 = rowIndex4 + 1;
            var rowIndex6 = rowIndex5 + 1;
            var rowIndex7 = rowIndex6 + 1;
            var rowIndex8 = rowIndex7 + 1;
            var row1 = sheet1.CreateRow(rowIndex1);
            var row2 = sheet1.CreateRow(rowIndex2);
            var row3 = sheet1.CreateRow(rowIndex3);
            var row4 = sheet1.CreateRow(rowIndex4);
            var row5 = sheet1.CreateRow(rowIndex5);
            var row6 = sheet1.CreateRow(rowIndex6);
            var row7 = sheet1.CreateRow(rowIndex7);
            var row8 = sheet1.CreateRow(rowIndex8);

            var list = new List<LoadCalculations.Sector>(SectorsList);
            list.Reverse();

            foreach (var sector in list)
            {
                var colFrom = colIndex;

                WriteCell(row, colIndex, sector.Name, _simpleHeader3Cs);
                WriteCell(row1, colIndex, sector.Percent + " %", _simpleHeader1Cs);
                WriteCell(row2, colIndex, "Σ " + sector.SumTotal, _simpleHeader1Cs);
                WriteCell(row4, colIndex, "Σ " + sector.SumRank, _simpleHeader1Cs);
                
                var machines = new List<LoadCalculations.Machine>(sector.Machines);
                machines.Reverse();

                foreach (var machine in machines)
                {
                    WriteCell(row6, colIndex, machine.Name, _simpleHeader2Cs);
                    WriteCell(row6, colIndex + 1, "");
                    WriteCell(row6, colIndex + 2, "");
                    WriteCell(row6, colIndex + 3, "");
                    sheet1.AddMergedRegion(new Region(rowIndex6, colIndex, rowIndex6, colIndex + 3));

                    WriteCell(row7, colIndex, "Σ " + machine.SumTotal, _simpleHeader3Cs);
                    WriteCell(row7, colIndex + 1, "");
                    WriteCell(row7, colIndex + 2, "");
                    WriteCell(row7, colIndex + 3, "");
                    sheet1.AddMergedRegion(new Region(rowIndex7, colIndex, rowIndex7, colIndex + 3));

                    foreach (DataRow r in machine.DataSource1.Rows)
                    {
                        WriteCell(row8, colIndex, (decimal)r["total1"]);
                        WriteCell(row8, colIndex + 1, (decimal)r["total2"]);
                        WriteCell(row8, colIndex + 2, (decimal)r["total3"]);
                        WriteCell(row8, colIndex + 3, (decimal)r["total4"]);
                    }

                    rIndex = rowIndex8 + 4;
                    var firstRowIndex = rIndex;
                    notConfirmedRowIndex = firstRowIndex - 2;
                    if (machine.DataSourceNotConfirmed != null)
                    {
                        foreach (DataRow r in ((DataTable)machine.DataSourceNotConfirmed.DataSource).Rows)
                        {
                            var rRow = sheet1.CreateRow(rIndex++);
                            WriteCell(rRow, colIndex, (decimal)r["total1"]);
                            WriteCell(rRow, colIndex + 1, (decimal)r["total2"]);
                            WriteCell(rRow, colIndex + 2, (decimal)r["total3"]);
                            WriteCell(rRow, colIndex + 3, (decimal)r["total4"]);
                        }
                    }
                    var lastRowIndex = rIndex - 1;
                    sheet1.GroupRow(firstRowIndex, lastRowIndex);

                    rIndex += 3;
                    firstRowIndex = rIndex;
                    forAgreedRowIndex = firstRowIndex - 2;
                    if (machine.DataSourceForAgreed != null)
                    {
                        foreach (DataRow r in ((DataTable)machine.DataSourceForAgreed.DataSource).Rows)
                        {
                            var rRow = sheet1.CreateRow(rIndex++);
                            WriteCell(rRow, colIndex, (decimal)r["total1"]);
                            WriteCell(rRow, colIndex + 1, (decimal)r["total2"]);
                            WriteCell(rRow, colIndex + 2, (decimal)r["total3"]);
                            WriteCell(rRow, colIndex + 3, (decimal)r["total4"]);
                        }
                    }
                    lastRowIndex = rIndex - 1;
                    sheet1.GroupRow(firstRowIndex, lastRowIndex);

                    rIndex += 3;
                    firstRowIndex = rIndex;
                    agreedRowIndex = firstRowIndex - 2;
                    if (machine.DataSourceAgreed != null)
                    {
                        foreach (DataRow r in ((DataTable)machine.DataSourceAgreed.DataSource).Rows)
                        {
                            var rRow = sheet1.CreateRow(rIndex++);
                            WriteCell(rRow, colIndex, (decimal)r["total1"]);
                            WriteCell(rRow, colIndex + 1, (decimal)r["total2"]);
                            WriteCell(rRow, colIndex + 2, (decimal)r["total3"]);
                            WriteCell(rRow, colIndex + 3, (decimal)r["total4"]);
                        }
                    }
                    lastRowIndex = rIndex - 1;
                    sheet1.GroupRow(firstRowIndex, lastRowIndex);

                    rIndex += 3;
                    firstRowIndex = rIndex;
                    onProductionRowIndex = firstRowIndex - 2;
                    if (machine.DataSourceOnProduction != null)
                    {
                        foreach (DataRow r in ((DataTable)machine.DataSourceOnProduction.DataSource).Rows)
                        {
                            var rRow = sheet1.CreateRow(rIndex++);
                            WriteCell(rRow, colIndex, (decimal)r["total1"]);
                            WriteCell(rRow, colIndex + 1, (decimal)r["total2"]);
                            WriteCell(rRow, colIndex + 2, (decimal)r["total3"]);
                            WriteCell(rRow, colIndex + 3, (decimal)r["total4"]);
                        }
                    }
                    lastRowIndex = rIndex - 1;
                    sheet1.GroupRow(firstRowIndex, lastRowIndex);

                    rIndex += 3;
                    firstRowIndex = rIndex;
                    inProductionRowIndex = firstRowIndex - 2;
                    if (machine.DataSourceInProduction != null)
                    {
                        foreach (DataRow r in ((DataTable)machine.DataSourceInProduction.DataSource).Rows)
                        {
                            var rRow = sheet1.CreateRow(rIndex++);
                            WriteCell(rRow, colIndex, (decimal)r["total1"]);
                            WriteCell(rRow, colIndex + 1, (decimal)r["total2"]);
                            WriteCell(rRow, colIndex + 2, (decimal)r["total3"]);
                            WriteCell(rRow, colIndex + 3, (decimal)r["total4"]);
                        }
                    }
                    lastRowIndex = rIndex - 1;
                    sheet1.GroupRow(firstRowIndex, lastRowIndex);

                    colIndex += 4;
                }

                var colTo = colIndex - 1;
                for (var i = colFrom + 1; i <= colTo; i++)
                    WriteCell(row, i, "");
                for (var i = colFrom + 1; i <= colTo; i++)
                    WriteCell(row1, i, "");
                for (var i = colFrom + 1; i <= colTo; i++)
                    WriteCell(row2, i, "");

                var ratio = (colTo + 1 - colFrom) / 4;

                WriteCell(row3, colFrom + 0 * ratio, "#0: " + sector.Total1);
                WriteCell(row3, colFrom + 1 * ratio, "#2: " + sector.Total2);
                WriteCell(row3, colFrom + 2 * ratio, "#3: " + sector.Total3);
                WriteCell(row3, colFrom + 3 * ratio, "#4: " + sector.Total4);
                for (var i = colFrom + 0 * ratio + 1; i < colFrom + 1 * ratio; i++)
                    WriteCell(row3, i, "");
                for (var i = colFrom + 1 * ratio + 1; i < colFrom + 2 * ratio; i++)
                    WriteCell(row3, i, "");
                for (var i = colFrom + 2 * ratio + 1; i < colFrom + 3 * ratio; i++)
                    WriteCell(row3, i, "");
                for (var i = colFrom + 3 * ratio + 1; i < colFrom + 4 * ratio; i++)
                    WriteCell(row3, i, "");
                sheet1.AddMergedRegion(new Region(rowIndex3, colFrom + 0 * ratio, rowIndex3, colFrom + 1 * ratio - 1));
                sheet1.AddMergedRegion(new Region(rowIndex3, colFrom + 1 * ratio, rowIndex3, colFrom + 2 * ratio - 1));
                sheet1.AddMergedRegion(new Region(rowIndex3, colFrom + 2 * ratio, rowIndex3, colFrom + 3 * ratio - 1));
                sheet1.AddMergedRegion(new Region(rowIndex3, colFrom + 3 * ratio, rowIndex3, colFrom + 4 * ratio - 1));

                for (var i = colFrom + 1; i <= colTo; i++)
                    WriteCell(row4, i, "");

                WriteCell(row5, colFrom + 0 * ratio, "#0: " + sector.Rank1);
                WriteCell(row5, colFrom + 1 * ratio, "#2: " + sector.Rank2);
                WriteCell(row5, colFrom + 2 * ratio, "#3: " + sector.Rank3);
                WriteCell(row5, colFrom + 3 * ratio, "#4: " + sector.Rank4);
                for (var i = colFrom + 0 * ratio + 1; i < colFrom + 1 * ratio; i++)
                    WriteCell(row5, i, "");
                for (var i = colFrom + 1 * ratio + 1; i < colFrom + 2 * ratio; i++)
                    WriteCell(row5, i, "");
                for (var i = colFrom + 2 * ratio + 1; i < colFrom + 3 * ratio; i++)
                    WriteCell(row5, i, "");
                for (var i = colFrom + 3 * ratio + 1; i < colFrom + 4 * ratio; i++)
                    WriteCell(row5, i, "");
                sheet1.AddMergedRegion(new Region(rowIndex5, colFrom + 0 * ratio, rowIndex5, colFrom + 1 * ratio - 1));
                sheet1.AddMergedRegion(new Region(rowIndex5, colFrom + 1 * ratio, rowIndex5, colFrom + 2 * ratio - 1));
                sheet1.AddMergedRegion(new Region(rowIndex5, colFrom + 2 * ratio, rowIndex5, colFrom + 3 * ratio - 1));
                sheet1.AddMergedRegion(new Region(rowIndex5, colFrom + 3 * ratio, rowIndex5, colFrom + 4 * ratio - 1));

                sheet1.AddMergedRegion(new Region(rowIndex, colFrom, rowIndex, colTo));
                sheet1.AddMergedRegion(new Region(rowIndex1, colFrom, rowIndex1, colTo));
                sheet1.AddMergedRegion(new Region(rowIndex2, colFrom, rowIndex2, colTo));
                sheet1.AddMergedRegion(new Region(rowIndex4, colFrom, rowIndex4, colTo));
            }

            PrintOrdersByStatus(sheet1, notConfirmedRowIndex, _calculations.ClientsNotConfirmedList, "Не подтверждено",
                 $@"{_calculations.NotConfirmedSumRank.ToString(DigitalFormat, nfi)}");
            PrintOrdersByStatus(sheet1, forAgreedRowIndex, _calculations.ClientsForAgreedList, "На согласовании", $@"{_calculations.ForAgreedSumRank.ToString(DigitalFormat, nfi)}");
            PrintOrdersByStatus(sheet1, agreedRowIndex, _calculations.ClientsAgreedList, "Согласовано", $@"{_calculations.AgreedSumRank.ToString(DigitalFormat, nfi)}");
            PrintOrdersByStatus(sheet1, onProductionRowIndex, _calculations.ClientsOnProductionList, "На производстве", $@"{_calculations.OnProductionSumRank.ToString(DigitalFormat, nfi)}");
            PrintOrdersByStatus(sheet1, inProductionRowIndex, _calculations.ClientsInProductionList, "В производстве", $@"{_calculations.InProductionSumRank.ToString(DigitalFormat, nfi)}");

            rIndex = notConfirmedRowIndex + 2;

            if (_calculations.ClientsNotConfirmedList.Count > 0)
                sheet1.SetRowGroupCollapsed(notConfirmedRowIndex + 3, true);
            if (_calculations.ClientsForAgreedList.Count > 0)
                sheet1.SetRowGroupCollapsed(forAgreedRowIndex + 3, true);
            if (_calculations.ClientsAgreedList.Count > 0)
                sheet1.SetRowGroupCollapsed(agreedRowIndex + 3, true);
            if (_calculations.ClientsOnProductionList.Count > 0)
                sheet1.SetRowGroupCollapsed(onProductionRowIndex + 3, true);
            if (_calculations.ClientsInProductionList.Count > 0)
                sheet1.SetRowGroupCollapsed(inProductionRowIndex + 3, true);

            var displayIndex = 0;
            sheet1.SetColumnWidth(displayIndex++, 35 * 256);
            sheet1.SetColumnWidth(displayIndex++, 8 * 256);
            sheet1.SetColumnWidth(displayIndex++, 13 * 256);
            sheet1.SetColumnWidth(displayIndex++, 13 * 256);

            sheet1.GetRow(rowIndex).Height = (short)4 * 256;
            sheet1.GetRow(rowIndex6).Height = (short)3 * 256;
            sheet1.GetRow(rowIndex7).Height = (short)3 * 256;
            
            sheet1.CreateFreezePane(0, rowIndex8);
            sheet1.CreateFreezePane(4, 0);

            sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;
            sheet1.SetMargin(HSSFSheet.LeftMargin, .12);
            sheet1.SetMargin(HSSFSheet.RightMargin, .07);
            sheet1.SetMargin(HSSFSheet.TopMargin, .20);
            sheet1.SetMargin(HSSFSheet.BottomMargin, .20);

            var fileName = $"Расчёты по нагрузке {DateTime.Now.ToShortDateString()}";
            var tempFolder = Environment.GetEnvironmentVariable("TEMP");

            var fileInfo = GetFileInfo(tempFolder, fileName);

            SaveReport(fileInfo);

            if (NeedStartFile)
                StartFile(fileInfo);
        }

        private void BatchSheet(LoadCalculations.Sector sector)
        {
            var sheetName = sector.ShortName;
            var sheet1 = _workbook.CreateSheet(sheetName);

            var colIndex = 4;
            var rowIndex = 2;
            var rIndex = 0;
            var row = sheet1.CreateRow(rowIndex);

            var notConfirmedRowIndex = 0;
            var forAgreedRowIndex = 0;
            var agreedRowIndex = 0;
            var onProductionRowIndex = 0;
            var inProductionRowIndex = 0;

            var rowIndex1 = rowIndex + 1;
            var rowIndex2 = rowIndex1 + 1;
            var rowIndex3 = rowIndex2 + 1;
            var rowIndex4 = rowIndex3 + 1;
            var rowIndex5 = rowIndex4 + 1;
            var rowIndex6 = rowIndex5 + 1;
            var rowIndex7 = rowIndex6 + 1;
            var rowIndex8 = rowIndex7 + 1;
            var row1 = sheet1.CreateRow(rowIndex1);
            var row2 = sheet1.CreateRow(rowIndex2);
            var row3 = sheet1.CreateRow(rowIndex3);
            var row4 = sheet1.CreateRow(rowIndex4);
            var row5 = sheet1.CreateRow(rowIndex5);
            var row6 = sheet1.CreateRow(rowIndex6);
            var row7 = sheet1.CreateRow(rowIndex7);
            var row8 = sheet1.CreateRow(rowIndex8);

            var colFrom = colIndex;

            WriteCell(row, colIndex, sector.Name, _simpleHeader3Cs);
            WriteCell(row1, colIndex, sector.Percent + " %", _simpleHeader1Cs);
            WriteCell(row2, colIndex, "Σ " + sector.SumTotal, _simpleHeader1Cs);
            WriteCell(row4, colIndex, "Σ " + sector.SumRank, _simpleHeader1Cs);
            var secId = sector.Id;
            if (sector.Machines == null)
                return;

            var machines = new List<LoadCalculations.Machine>(sector.Machines);
            machines.Reverse();

            foreach (var machine in machines)
            {
                WriteCell(row6, colIndex, machine.Name, _simpleHeader2Cs);
                WriteCell(row6, colIndex + 1, "");
                WriteCell(row6, colIndex + 2, "");
                WriteCell(row6, colIndex + 3, "");
                sheet1.AddMergedRegion(new Region(rowIndex6, colIndex, rowIndex6, colIndex + 3));

                WriteCell(row7, colIndex, "Σ " + machine.SumTotal, _simpleHeader3Cs);
                WriteCell(row7, colIndex + 1, "");
                WriteCell(row7, colIndex + 2, "");
                WriteCell(row7, colIndex + 3, "");
                sheet1.AddMergedRegion(new Region(rowIndex7, colIndex, rowIndex7, colIndex + 3));

                foreach (DataRow r in machine.DataSource1.Rows)
                {
                    WriteCell(row8, colIndex, (decimal)r["total1"]);
                    WriteCell(row8, colIndex + 1, (decimal)r["total2"]);
                    WriteCell(row8, colIndex + 2, (decimal)r["total3"]);
                    WriteCell(row8, colIndex + 3, (decimal)r["total4"]);
                }

                rIndex = rowIndex8 + 4;
                var firstRowIndex = rIndex;
                notConfirmedRowIndex = firstRowIndex - 2;
                if (machine.DataSourceNotConfirmed != null)
                {
                    foreach (DataRow r in ((DataTable)machine.DataSourceNotConfirmed.DataSource).Rows)
                    {
                        var rRow = sheet1.CreateRow(rIndex++);
                        WriteCell(rRow, colIndex, (decimal)r["total1"]);
                        WriteCell(rRow, colIndex + 1, (decimal)r["total2"]);
                        WriteCell(rRow, colIndex + 2, (decimal)r["total3"]);
                        WriteCell(rRow, colIndex + 3, (decimal)r["total4"]);
                    }
                }

                var lastRowIndex = rIndex - 1;
                sheet1.GroupRow(firstRowIndex, lastRowIndex);

                rIndex += 3;
                firstRowIndex = rIndex;
                forAgreedRowIndex = firstRowIndex - 2;
                if (machine.DataSourceForAgreed != null)
                {
                    foreach (DataRow r in ((DataTable)machine.DataSourceForAgreed.DataSource).Rows)
                    {
                        var rRow = sheet1.CreateRow(rIndex++);
                        WriteCell(rRow, colIndex, (decimal)r["total1"]);
                        WriteCell(rRow, colIndex + 1, (decimal)r["total2"]);
                        WriteCell(rRow, colIndex + 2, (decimal)r["total3"]);
                        WriteCell(rRow, colIndex + 3, (decimal)r["total4"]);
                    }
                }

                lastRowIndex = rIndex - 1;
                sheet1.GroupRow(firstRowIndex, lastRowIndex);

                rIndex += 3;
                firstRowIndex = rIndex;
                agreedRowIndex = firstRowIndex - 2;
                if (machine.DataSourceAgreed != null)
                {
                    foreach (DataRow r in ((DataTable)machine.DataSourceAgreed.DataSource).Rows)
                    {
                        var rRow = sheet1.CreateRow(rIndex++);
                        WriteCell(rRow, colIndex, (decimal)r["total1"]);
                        WriteCell(rRow, colIndex + 1, (decimal)r["total2"]);
                        WriteCell(rRow, colIndex + 2, (decimal)r["total3"]);
                        WriteCell(rRow, colIndex + 3, (decimal)r["total4"]);
                    }
                }

                lastRowIndex = rIndex - 1;
                sheet1.GroupRow(firstRowIndex, lastRowIndex);

                rIndex += 3;
                firstRowIndex = rIndex;
                onProductionRowIndex = firstRowIndex - 2;
                if (machine.DataSourceOnProduction != null)
                {
                    foreach (DataRow r in ((DataTable)machine.DataSourceOnProduction.DataSource).Rows)
                    {
                        var rRow = sheet1.CreateRow(rIndex++);
                        WriteCell(rRow, colIndex, (decimal)r["total1"]);
                        WriteCell(rRow, colIndex + 1, (decimal)r["total2"]);
                        WriteCell(rRow, colIndex + 2, (decimal)r["total3"]);
                        WriteCell(rRow, colIndex + 3, (decimal)r["total4"]);
                    }
                }

                lastRowIndex = rIndex - 1;
                sheet1.GroupRow(firstRowIndex, lastRowIndex);

                rIndex += 3;
                firstRowIndex = rIndex;
                inProductionRowIndex = firstRowIndex - 2;
                if (machine.DataSourceInProduction != null)
                {
                    foreach (DataRow r in ((DataTable)machine.DataSourceInProduction.DataSource).Rows)
                    {
                        var rRow = sheet1.CreateRow(rIndex++);
                        WriteCell(rRow, colIndex, (decimal)r["total1"]);
                        WriteCell(rRow, colIndex + 1, (decimal)r["total2"]);
                        WriteCell(rRow, colIndex + 2, (decimal)r["total3"]);
                        WriteCell(rRow, colIndex + 3, (decimal)r["total4"]);
                    }
                }

                lastRowIndex = rIndex - 1;
                sheet1.GroupRow(firstRowIndex, lastRowIndex);

                colIndex += 4;
            }

            var colTo = colIndex - 1;
            for (var i = colFrom + 1; i <= colTo; i++)
                WriteCell(row, i, "");
            for (var i = colFrom + 1; i <= colTo; i++)
                WriteCell(row1, i, "");
            for (var i = colFrom + 1; i <= colTo; i++)
                WriteCell(row2, i, "");

            var ratio = (colTo + 1 - colFrom) / 4;

            WriteCell(row3, colFrom + 0 * ratio, "#0: " + sector.Total1);
            WriteCell(row3, colFrom + 1 * ratio, "#2: " + sector.Total2);
            WriteCell(row3, colFrom + 2 * ratio, "#3: " + sector.Total3);
            WriteCell(row3, colFrom + 3 * ratio, "#4: " + sector.Total4);
            for (var i = colFrom + 0 * ratio + 1; i < colFrom + 1 * ratio; i++)
                WriteCell(row3, i, "");
            for (var i = colFrom + 1 * ratio + 1; i < colFrom + 2 * ratio; i++)
                WriteCell(row3, i, "");
            for (var i = colFrom + 2 * ratio + 1; i < colFrom + 3 * ratio; i++)
                WriteCell(row3, i, "");
            for (var i = colFrom + 3 * ratio + 1; i < colFrom + 4 * ratio; i++)
                WriteCell(row3, i, "");
            sheet1.AddMergedRegion(new Region(rowIndex3, colFrom + 0 * ratio, rowIndex3, colFrom + 1 * ratio - 1));
            sheet1.AddMergedRegion(new Region(rowIndex3, colFrom + 1 * ratio, rowIndex3, colFrom + 2 * ratio - 1));
            sheet1.AddMergedRegion(new Region(rowIndex3, colFrom + 2 * ratio, rowIndex3, colFrom + 3 * ratio - 1));
            sheet1.AddMergedRegion(new Region(rowIndex3, colFrom + 3 * ratio, rowIndex3, colFrom + 4 * ratio - 1));

            for (var i = colFrom + 1; i <= colTo; i++)
                WriteCell(row4, i, "");

            WriteCell(row5, colFrom + 0 * ratio, "#0: " + sector.Rank1);
            WriteCell(row5, colFrom + 1 * ratio, "#2: " + sector.Rank2);
            WriteCell(row5, colFrom + 2 * ratio, "#3: " + sector.Rank3);
            WriteCell(row5, colFrom + 3 * ratio, "#4: " + sector.Rank4);
            for (var i = colFrom + 0 * ratio + 1; i < colFrom + 1 * ratio; i++)
                WriteCell(row5, i, "");
            for (var i = colFrom + 1 * ratio + 1; i < colFrom + 2 * ratio; i++)
                WriteCell(row5, i, "");
            for (var i = colFrom + 2 * ratio + 1; i < colFrom + 3 * ratio; i++)
                WriteCell(row5, i, "");
            for (var i = colFrom + 3 * ratio + 1; i < colFrom + 4 * ratio; i++)
                WriteCell(row5, i, "");
            sheet1.AddMergedRegion(new Region(rowIndex5, colFrom + 0 * ratio, rowIndex5, colFrom + 1 * ratio - 1));
            sheet1.AddMergedRegion(new Region(rowIndex5, colFrom + 1 * ratio, rowIndex5, colFrom + 2 * ratio - 1));
            sheet1.AddMergedRegion(new Region(rowIndex5, colFrom + 2 * ratio, rowIndex5, colFrom + 3 * ratio - 1));
            sheet1.AddMergedRegion(new Region(rowIndex5, colFrom + 3 * ratio, rowIndex5, colFrom + 4 * ratio - 1));

            sheet1.AddMergedRegion(new Region(rowIndex, colFrom, rowIndex, colTo));
            sheet1.AddMergedRegion(new Region(rowIndex1, colFrom, rowIndex1, colTo));
            sheet1.AddMergedRegion(new Region(rowIndex2, colFrom, rowIndex2, colTo));
            sheet1.AddMergedRegion(new Region(rowIndex4, colFrom, rowIndex4, colTo));

            //PrintOrdersByStatus(sheet1, notConfirmedRowIndex, _calculations.ClientsNotConfirmedList, "Не подтверждено",
            //    $@"{_calculations.NotConfirmedSumRank.ToString(DigitalFormat, nfi)}");
            //PrintOrdersByStatus(sheet1, forAgreedRowIndex, _calculations.ClientsForAgreedList, "На согласовании",
            //    $@"{_calculations.ForAgreedSumRank.ToString(DigitalFormat, nfi)}");
            //PrintOrdersByStatus(sheet1, agreedRowIndex, _calculations.ClientsAgreedList, "Согласовано",
            //    $@"{_calculations.AgreedSumRank.ToString(DigitalFormat, nfi)}");
            //PrintOrdersByStatus(sheet1, onProductionRowIndex, _calculations.ClientsOnProductionList, "На производстве",
            //    $@"{_calculations.OnProductionSumRank.ToString(DigitalFormat, nfi)}");
            //PrintOrdersByStatus(sheet1, inProductionRowIndex, _calculations.ClientsInProductionList, "В производстве",
            //    $@"{_calculations.InProductionSumRank.ToString(DigitalFormat, nfi)}");

            //rIndex = notConfirmedRowIndex + 2;

            if (_calculations.ClientsNotConfirmedList.Count > 1)
                sheet1.SetRowGroupCollapsed(notConfirmedRowIndex + 3, true);
            if (_calculations.ClientsForAgreedList.Count > 1)
                sheet1.SetRowGroupCollapsed(forAgreedRowIndex + 3, true);
            if (_calculations.ClientsAgreedList.Count > 1)
                sheet1.SetRowGroupCollapsed(agreedRowIndex + 3, true);
            if (_calculations.ClientsOnProductionList.Count > 1)
                sheet1.SetRowGroupCollapsed(onProductionRowIndex + 3, true);
            if (_calculations.ClientsInProductionList.Count > 1)
                sheet1.SetRowGroupCollapsed(inProductionRowIndex + 3, true);

            var displayIndex = 0;
            sheet1.SetColumnWidth(displayIndex++, 35 * 256);
            sheet1.SetColumnWidth(displayIndex++, 8 * 256);
            sheet1.SetColumnWidth(displayIndex++, 13 * 256);
            sheet1.SetColumnWidth(displayIndex++, 13 * 256);

            sheet1.GetRow(rowIndex).Height = (short)4 * 256;
            sheet1.GetRow(rowIndex6).Height = (short)3 * 256;
            sheet1.GetRow(rowIndex7).Height = (short)3 * 256;

            sheet1.CreateFreezePane(0, rowIndex8);
            sheet1.CreateFreezePane(4, 0);

            sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;
            sheet1.SetMargin(HSSFSheet.LeftMargin, .12);
            sheet1.SetMargin(HSSFSheet.RightMargin, .07);
            sheet1.SetMargin(HSSFSheet.TopMargin, .20);
            sheet1.SetMargin(HSSFSheet.BottomMargin, .20);
        }

        public void BatchReport(int megaBatchId)
        {
            var list = new List<LoadCalculations.Sector>(SectorsList);
            list.Reverse();

            foreach (var sector in list)
            {
                BatchSheet(sector);
            }

            var fileName = $"Группа партий №{megaBatchId}. Расчёты по нагрузке";
            var tempFolder = Environment.GetEnvironmentVariable("TEMP");

            var fileInfo = GetFileInfo(tempFolder, fileName);

            SaveReport(fileInfo);

            if (NeedStartFile)
                StartFile(fileInfo);
        }
        
        private void PrintOrdersByStatus(HSSFSheet sheet1, int rowIndex, BindingSource bs, string status, string count)
        {
            var row = sheet1.CreateRow(rowIndex);
            var row1 = sheet1.CreateRow(rowIndex + 1);
            WriteCell(row, 0, status, _simpleHeader2Cs);
            WriteCell(row, 1, "");
            WriteCell(row, 2, "");
            WriteCell(row, 3, "");
            WriteCell(row1, 0, count, _simpleHeader3Cs);
            WriteCell(row1, 1, "");
            WriteCell(row1, 2, "");
            WriteCell(row1, 3, "");

            var rIndex = rowIndex + 2;
            foreach (DataRow r in ((DataTable)bs.DataSource).Rows)
            {
                var rRow = sheet1.GetRow(rIndex++);
                WriteCell(rRow, 0, r["clientName"].ToString());
                WriteCell(rRow, 1, (int)r["orderNumber"]);
                WriteCell(rRow, 2, r["orderDate"].ToString());
                WriteCell(rRow, 3, (decimal)r["sumTotal"]);
            }

            sheet1.AddMergedRegion(new Region(rowIndex, 0, rowIndex, 3));
            sheet1.AddMergedRegion(new Region(rowIndex + 1, 0, rowIndex + 1, 3));
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
        
        private void SaveReport(FileSystemInfo fileInfo)
        {
            FileStream newFile = new(fileInfo.FullName, FileMode.Create);
            _workbook.Write(newFile);
            newFile.Close();
        }

        private void WriteCell(HSSFRow row, int colIndex, decimal cellValue)
        {
            var cell = row.CreateCell(colIndex);
            cell.SetCellValue((double)cellValue);
            cell.CellStyle = _simpleDecCs;
        }

        private void WriteCell(HSSFRow row, int colIndex, decimal cellValue, HSSFCellStyle cs)
        {
            var cell = row.CreateCell(colIndex);
            cell.SetCellValue((double)cellValue);
            cell.CellStyle = cs;
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

        private void WriteCell(HSSFRow row, int colIndex, string cellValue, HSSFCellStyle cs)
        {
            var cell = row.CreateCell(colIndex);
            cell.SetCellValue(cellValue);
            cell.CellStyle = cs;
        }
    }
}