using NPOI.HSSF.Model;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;

using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Diagnostics;
using System.Drawing.Printing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraExport;

namespace Infinium.Modules.LoadCalculations
{
    public class ExcelLoadCalculationsReport
    {
        private HSSFCellStyle _simpleCs;
        private HSSFCellStyle _simpleDecCs;
        private HSSFCellStyle _simpleHeader1Cs;
        private HSSFCellStyle _simpleHeader2Cs;
        private HSSFCellStyle _simpleHeader3Cs;
        private HSSFCellStyle _simpleIntCs;
        private HSSFWorkbook _workbook;

        public ExcelLoadCalculationsReport()
        {
        }

        private void CreateStyles()
        {
            nfi = (NumberFormatInfo)CultureInfo.InvariantCulture.NumberFormat.Clone();
            nfi.NumberDecimalSeparator = ",";
            nfi.NumberGroupSeparator = " ";
            
            var header1F = _workbook.CreateFont();
            header1F.FontHeightInPoints = 9;
            header1F.Boldweight = HSSFFont.BOLDWEIGHT_BOLD;
            header1F.FontName = "Arial";

            var header2F = _workbook.CreateFont();
            header2F.FontHeightInPoints = 12;
            header2F.FontName = "Arial";

            var header3F = _workbook.CreateFont();
            header3F.FontHeightInPoints = 13;
            header3F.Boldweight = HSSFFont.BOLDWEIGHT_BOLD;
            header3F.FontName = "Arial";

            var simpleF = _workbook.CreateFont();
            simpleF.FontHeightInPoints = 8;
            simpleF.FontName = "Arial";

            var simpleBoldF = _workbook.CreateFont();
            simpleBoldF.FontHeightInPoints = 11;
            simpleBoldF.Boldweight = HSSFFont.BOLDWEIGHT_BOLD;
            simpleBoldF.FontName = "Times New Roman";

            _simpleCs = _workbook.CreateCellStyle();
            _simpleCs.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
            _simpleCs.Alignment = HSSFCellStyle.ALIGN_RIGHT;
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
            _simpleIntCs.Alignment = HSSFCellStyle.ALIGN_RIGHT;
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
            _simpleDecCs.Alignment = HSSFCellStyle.ALIGN_RIGHT;
            _simpleDecCs.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.0000000");
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

        public ExcelLoadCalculations _calculations { get; set; }

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
        
        public void CreateReport(DataTable ReportDt, FileInfo fileInfo)
        {
            var sheet1 = _workbook.CloneSheet(0);
            var sheetNameExist = false;
            var sheetName = "РЕЗУЛЬТАТ";
            
            for (var i = 0; i < _workbook.NumberOfSheets; i++)
            {
                if (sheetName == _workbook.GetSheetName(i))
                    sheetNameExist = true;
            }

            var j = 1;
            while (sheetNameExist)
            {
                sheetName = $"РЕЗУЛЬТАТ({j++})";
                sheetNameExist = false;
                for (var i = 0; i < _workbook.NumberOfSheets; i++)
                {
                    if (sheetName == _workbook.GetSheetName(i))
                        sheetNameExist = true;
                }
            }

            _workbook.SetSheetName(_workbook.GetSheetIndex(sheet1), sheetName);
            
            var colIndex0 = 0;
            var colIndex4 = 4;
            var colIndex5 = 5;
            var colIndex6 = 6;
            var colIndex7 = 7;
            var colIndex8 = 8;
            var colIndex9 = 9;
            var colIndex10 = 10;
            var colIndex11 = 11;
            var colIndex12 = 12;
            
            var row = sheet1.GetRow(0);
            WriteCell(row, colIndex4, ReportDt.Columns[5].ColumnName, _simpleHeader1Cs);
            WriteCell(row, colIndex5, ReportDt.Columns[6].ColumnName, _simpleHeader1Cs);
            WriteCell(row, colIndex6, ReportDt.Columns[7].ColumnName, _simpleHeader1Cs);
            WriteCell(row, colIndex7, ReportDt.Columns[8].ColumnName, _simpleHeader1Cs);
            WriteCell(row, colIndex8, ReportDt.Columns[9].ColumnName, _simpleHeader1Cs);
            WriteCell(row, colIndex9, ReportDt.Columns[10].ColumnName, _simpleHeader1Cs);
            WriteCell(row, colIndex10, ReportDt.Columns[11].ColumnName, _simpleHeader1Cs);
            WriteCell(row, colIndex11, ReportDt.Columns[12].ColumnName, _simpleHeader1Cs);
            WriteCell(row, colIndex12, "ИТОГО", _simpleHeader1Cs);

            decimal totalSector1 = 0;
            decimal totalSector2 = 0;
            decimal totalSector3 = 0;
            decimal totalSector4 = 0;
            decimal totalSector5 = 0;
            decimal totalSector6 = 0;
            decimal totalSector7 = 0;
            decimal totalSector8 = 0;
            decimal totalSector9 = 0;
            decimal total1 = 0;
            decimal total2 = 0;
            decimal total3 = 0;
            decimal total4 = 0;
            for (var i = 0; i < ReportDt.Rows.Count; i++)
            {
                row = sheet1.GetRow(i + 1);
                WriteCell(row, colIndex4, Convert.ToDecimal(ReportDt.Rows[i][5]));
                WriteCell(row, colIndex5, Convert.ToDecimal(ReportDt.Rows[i][6]));
                WriteCell(row, colIndex6, Convert.ToDecimal(ReportDt.Rows[i][7]));
                WriteCell(row, colIndex7, Convert.ToDecimal(ReportDt.Rows[i][8]));
                WriteCell(row, colIndex8, Convert.ToDecimal(ReportDt.Rows[i][9]));
                WriteCell(row, colIndex9, Convert.ToDecimal(ReportDt.Rows[i][10]));
                WriteCell(row, colIndex10, Convert.ToDecimal(ReportDt.Rows[i][11]));
                WriteCell(row, colIndex11, Convert.ToDecimal(ReportDt.Rows[i][12]));
                WriteCell(row, colIndex12, Convert.ToDecimal(ReportDt.Rows[i][13]));

                total1 += Convert.ToDecimal(ReportDt.Rows[i][14]);
                total2 += Convert.ToDecimal(ReportDt.Rows[i][15]);
                total3 += Convert.ToDecimal(ReportDt.Rows[i][16]);
                total4 += Convert.ToDecimal(ReportDt.Rows[i][17]);

                if (ReportDt.Rows[i][2].ToString().Length == 0)
                {
                    totalSector1 += Convert.ToDecimal(ReportDt.Rows[i][5]);
                    totalSector2 += Convert.ToDecimal(ReportDt.Rows[i][6]);
                    totalSector3 += Convert.ToDecimal(ReportDt.Rows[i][7]);
                    totalSector4 += Convert.ToDecimal(ReportDt.Rows[i][8]);
                    totalSector5 += Convert.ToDecimal(ReportDt.Rows[i][9]);
                    totalSector6 += Convert.ToDecimal(ReportDt.Rows[i][10]);
                    totalSector7 += Convert.ToDecimal(ReportDt.Rows[i][11]);
                    totalSector8 += Convert.ToDecimal(ReportDt.Rows[i][12]);
                    totalSector9 += Convert.ToDecimal(ReportDt.Rows[i][13]);
                }
            }

            var rowIndex = ReportDt.Rows.Count + 1;
            row = sheet1.CreateRow(rowIndex);
            WriteCell(row, colIndex4, totalSector1);
            WriteCell(row, colIndex5, totalSector2);
            WriteCell(row, colIndex6, totalSector3);
            WriteCell(row, colIndex7, totalSector4);
            WriteCell(row, colIndex8, totalSector5);
            WriteCell(row, colIndex9, totalSector6);
            WriteCell(row, colIndex10, totalSector7);
            WriteCell(row, colIndex11, totalSector8);
            WriteCell(row, colIndex12, totalSector9);

            ++rowIndex;
            row = sheet1.CreateRow(++rowIndex);
            WriteCell(row, colIndex0, "0");
            WriteCell(row, colIndex12, total1);
            row = sheet1.CreateRow(++rowIndex);
            WriteCell(row, colIndex0, "2");
            WriteCell(row, colIndex12, total2);
            row = sheet1.CreateRow(++rowIndex);
            WriteCell(row, colIndex0, "3");
            WriteCell(row, colIndex12, total3);
            row = sheet1.CreateRow(++rowIndex);
            WriteCell(row, colIndex0, "4");
            WriteCell(row, colIndex12, total4);

            sheet1.CreateFreezePane(0, 1);

            sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;
            sheet1.SetMargin(HSSFSheet.LeftMargin, .12);
            sheet1.SetMargin(HSSFSheet.RightMargin, .07);
            sheet1.SetMargin(HSSFSheet.TopMargin, .20);
            sheet1.SetMargin(HSSFSheet.BottomMargin, .20);

            var fileName = $"Расчёт времени {DateTime.Now.ToShortDateString()}";
            var tempFolder = Environment.GetEnvironmentVariable("TEMP");

            var fi = GetFileInfo(fileInfo.DirectoryName, fileName);

            SaveReport(fi);

            if (NeedStartFile)
                StartFile(fi);
        }

        public HSSFSheet ExportFile(FileInfo fileInfo)
        {
            using (var fileStream = new FileStream(fileInfo.FullName, FileMode.Open, FileAccess.Read))
            {
                _workbook = new HSSFWorkbook(fileStream);
            }
            CreateStyles();

            var sheet = _workbook.GetSheetAt(0);

            return sheet;
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
