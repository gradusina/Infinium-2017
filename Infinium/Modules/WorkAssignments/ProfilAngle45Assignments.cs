using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.UserModel.Contrib;
using NPOI.HSSF.Util;

using System;
using System.Collections;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;

namespace Infinium.Modules.WorkAssignments
{
    public class ProfilAngle45Assignments : IAllFrontParameterName
    {
        #region Public Constructors

        public ProfilAngle45Assignments()
        {
        }

        #endregion Public Constructors

        #region Public Fields

        public DataTable FrameColorsDataTable = null;
        public DataTable FrontsDataTable = null;
        public DataTable InsetColorsDataTable = null;
        public DataTable InsetTypesDataTable = null;
        public DataTable PatinaDataTable = null;
        public DataTable StandardImpostDataTable = null;
        public DataTable TechnoInsetColorsDataTable = null;
        public DataTable TechnoInsetTypesDataTable = null;

        #endregion Public Fields

        #region Private Fields

        private DataTable Additional1DT;
        private DataTable Additional2DT;
        private DataTable Additional3DT;
        private DataTable Additional4DT;
        private DataTable Additional5DT;
        private DataTable AlbyGridsDT;
        private DataTable AlbyOrdersDT;
        private DataTable AlbySimpleDT;
        private DataTable AlbyVitrinaDT;
        private DataTable AntaliaGridsDT;
        private DataTable AntaliaOrdersDT;
        private DataTable AntaliaSimpleDT;
        private DataTable AntaliaVitrinaDT;
        private DataTable ArchDecorOrdersDT;
        private DataTable AssemblyDT;
        private DataTable BagetWithAngelOrdersDT;
        private DataTable BagetWithAngleAssemblyDT;
        private DataTable BergamoGlassDT;
        private DataTable BergamoGridsDT;
        private DataTable BergamoOrdersDT;
        private DataTable BergamoSimpleDT;
        private DataTable BergamoVitrinaDT;
        private DataTable BostonGlassDT;
        private DataTable BostonGridsDT;
        private DataTable BostonOrdersDT;
        private DataTable BostonSimpleDT;
        private DataTable BostonVitrinaDT;
        private DataTable BrunoGridsDT;
        private DataTable BrunoOrdersDT;
        private DataTable BrunoSimpleDT;
        private DataTable BrunoVitrinaDT;
        private DataTable ComecDT;
        private DateTime CurrentDate;
        private DataTable DecorAssemblyDT;
        private DataTable DecorDT;
        private DataTable DecorParametersDT;
        private DataTable DeyingDT;
        private DataTable DistMainOrdersDT;
        private DataTable ep018Marsel1GridsDT;
        private DataTable ep018Marsel1OrdersDT;
        private DataTable ep018Marsel1SimpleDT;
        private DataTable ep018Marsel1VitrinaDT;
        private DataTable ep041GridsDT;
        private DataTable ep041OrdersDT;
        private DataTable ep041SimpleDT;
        private DataTable ep041VitrinaDT;
        private DataTable ep043ShervudGridsDT;
        private DataTable ep043ShervudOrdersDT;
        private DataTable ep043ShervudSimpleDT;
        private DataTable ep043ShervudVitrinaDT;
        private DataTable ep066Marsel4GridsDT;
        private DataTable ep066Marsel4OrdersDT;
        private DataTable ep066Marsel4SimpleDT;
        private DataTable ep066Marsel4VitrinaDT;
        private DataTable ep071GridsDT;
        private DataTable ep071OrdersDT;
        private DataTable ep071SimpleDT;
        private DataTable ep071VitrinaDT;
        private DataTable ep110JersyGridsDT;
        private DataTable ep110JersyOrdersDT;
        private DataTable ep110JersySimpleDT;
        private DataTable ep110JersyVitrinaDT;
        private DataTable ep111GridsDT;
        private DataTable ep111OrdersDT;
        private DataTable ep111SimpleDT;
        private DataTable ep111VitrinaDT;
        private DataTable ep112GridsDT;
        private DataTable ep112OrdersDT;
        private DataTable ep112SimpleDT;
        private DataTable ep112VitrinaDT;
        private DataTable ep206GridsDT;
        private DataTable ep206OrdersDT;
        private DataTable ep206SimpleDT;
        private DataTable ep206VitrinaDT;
        private DataTable ep216GridsDT;
        private DataTable ep216OrdersDT;
        private DataTable ep216SimpleDT;
        private DataTable ep216VitrinaDT;
        private DataTable epFoxGridsDT;
        private DataTable epFoxOrdersDT;
        private DataTable epFoxSimpleDT;
        private DataTable epFoxVitrinaDT;
        private DataTable epsh406Techno4GridsDT;
        private DataTable epsh406Techno4OrdersDT;
        private DataTable epsh406Techno4SimpleDT;
        private DataTable epsh406Techno4VitrinaDT;
        private DataTable FatGridsDT;
        private DataTable FatOrdersDT;
        private DataTable FatSimpleDT;
        private DataTable FatVitrinaDT;
        private FileManager FM = new FileManager();
        private ArrayList FrontsID;
        private DataTable FrontsOrdersDT;
        private DataTable GridsDecorOrdersDT;
        private DataTable InsetDT;
        private DataTable LeonGridsDT;
        private DataTable LeonOrdersDT;
        private DataTable LeonSimpleDT;
        private DataTable LeonVitrinaDT;
        private DataTable LimogGridsDT;
        private DataTable LimogOrdersDT;
        private DataTable LimogSimpleDT;
        private DataTable LimogVitrinaDT;
        private DataTable LukGridsDT;
        private DataTable LukOrdersDT;
        private DataTable LukPVHGridsDT;
        private DataTable LukPVHOrdersDT;
        private DataTable LukPVHSimpleDT;
        private DataTable LukPVHVitrinaDT;
        private DataTable LukSimpleDT;
        private DataTable LukVitrinaDT;
        private DataTable MilanoGridsDT;
        private DataTable MilanoOrdersDT;
        private DataTable MilanoSimpleDT;
        private DataTable MilanoVitrinaDT;
        private DataTable Nord95GridsDT;
        private DataTable Nord95OrdersDT;
        private DataTable Nord95SimpleDT;
        private DataTable Nord95VitrinaDT;
        private DataTable NotArchDecorOrdersDT;
        private DataTable PatinaRALDataTable = null;
        //DataTable ColorsDT;
        //DataTable FrontsDT;

        private DataTable PragaGridsDT;
        private DataTable PragaOrdersDT;
        private DataTable PragaSimpleDT;
        private DataTable PragaVitrinaDT;
        private DataTable ProfileNamesDT;
        private DataTable RapidDT;
        private DataTable SigmaGlassDT;
        private DataTable SigmaGridsDT;
        private DataTable SigmaOrdersDT;
        private DataTable SigmaSimpleDT;
        private DataTable SigmaVitrinaDT;
        private DataTable SummOrdersDT;
        private DataTable TechnoNGridsDT;
        private DataTable TechnoNOrdersDT;
        private DataTable TechnoNSimpleDT;
        private DataTable TechnoNVitrinaDT;
        private DataTable UrbanGridsDT;
        private DataTable UrbanOrdersDT;
        private DataTable UrbanSimpleDT;
        private DataTable UrbanVitrinaDT;
        private DataTable VeneciaGridsDT;
        private DataTable VeneciaOrdersDT;
        private DataTable VeneciaSimpleDT;
        private DataTable VeneciaVitrinaDT;

        #endregion Private Fields

        #region Public Properties

        public ArrayList GetFrontsID
        {
            set => FrontsID = value;
        }

        #endregion Public Properties

        #region Public Methods

        public static int LoadImage(string path, HSSFWorkbook wb)
        {
            FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read);
            byte[] buffer = new byte[file.Length];
            file.Read(buffer, 0, (int)file.Length);
            return wb.AddPicture(buffer, HSSFWorkbook.PICTURE_TYPE_JPEG);
        }

        public void AdditionsToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            int WorkAssignmentID, string BatchName, string ClientName)
        {
            Additional1DT.Clear();
            CollectAdditions0(ref Additional1DT, 28);
            CollectAdditions1(ref Additional1DT, 138);

            Additional2DT.Clear();
            CollectAdditions2(ref Additional2DT, 138);

            Additional3DT.Clear();
            CollectAdditions3(ref Additional3DT, 138);

            Additional4DT.Clear();
            CollectAdditions4(ref Additional4DT, 138);

            Additional5DT.Clear();
            CollectAdditions5(ref Additional5DT, 138);

            if (Additional1DT.Rows.Count == 0 && Additional2DT.Rows.Count == 0 && Additional3DT.Rows.Count == 0
                && Additional4DT.Rows.Count == 0 && Additional5DT.Rows.Count == 0)
                return;

            int RowIndex = 0;

            HSSFSheet sheet1 = hssfworkbook.CreateSheet("Доп. задания");
            sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

            sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
            sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
            sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
            sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

            sheet1.SetColumnWidth(0, 25 * 256);
            sheet1.SetColumnWidth(1, 10 * 256);
            sheet1.SetColumnWidth(2, 11 * 256);
            sheet1.SetColumnWidth(3, 11 * 256);
            sheet1.SetColumnWidth(4, 11 * 256);
            sheet1.SetColumnWidth(5, 7 * 256);
            sheet1.SetColumnWidth(6, 11 * 256);

            {
                HSSFPatriarch patriarch = sheet1.CreateDrawingPatriarch();
                int ind = 0;

                CreateBarcode(WorkAssignmentID);
                HSSFClientAnchor anchor = new HSSFClientAnchor(0, 1, 1000, 8, 5, ind, 6, ind + 3)
                {
                    AnchorType = 2
                };
                string BarcodeNumber = GetBarcodeNumber(24, WorkAssignmentID);
                string barcodeName = System.Environment.GetEnvironmentVariable("TEMP") + @"/" + BarcodeNumber + ".jpg";
                HSSFPicture picture = patriarch.CreatePicture(anchor, LoadImage(barcodeName, hssfworkbook));

                HSSFFont BarcodeFont = hssfworkbook.CreateFont();
                BarcodeFont.FontHeightInPoints = 12;
                BarcodeFont.FontName = "Calibri";

                HSSFCellStyle BarcodeStyle = hssfworkbook.CreateCellStyle();
                BarcodeStyle.Alignment = HSSFCellStyle.ALIGN_CENTER;
                BarcodeStyle.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
                BarcodeStyle.SetFont(BarcodeFont);

                HSSFCell barcodeCell = sheet1.CreateRow(ind + 3).CreateCell(5);
                barcodeCell.SetCellValue(BarcodeNumber.ToString());
                barcodeCell.CellStyle = BarcodeStyle;
                sheet1.AddMergedRegion(new NPOI.HSSF.Util.Region(ind + 3, 5, ind + 3, 6));
            }

            if (Additional1DT.Rows.Count > 0)
            {
                AdditionsToExcelSingly(ref hssfworkbook,
                    Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, Additional1DT, WorkAssignmentID, BatchName, ClientName, "Резка стекла", "Стекло", ref RowIndex);
                RowIndex++;
            }

            if (Additional2DT.Rows.Count > 0)
            {
                AdditionsToExcelSingly(ref hssfworkbook,
                    Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, Additional2DT, WorkAssignmentID, BatchName, ClientName, "Пила", "Пила", ref RowIndex);
                RowIndex++;
            }

            if (Additional2DT.Rows.Count > 0)
            {
                AdditionsToExcelSingly(ref hssfworkbook,
                    Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, Additional2DT, WorkAssignmentID, BatchName, ClientName, "Клей", "Клей", ref RowIndex);
                RowIndex++;
            }

            if (Additional3DT.Rows.Count > 0)
            {
                AdditionsToExcelSingly(ref hssfworkbook,
                    Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, Additional3DT, WorkAssignmentID, BatchName, ClientName, "Пресс", "Пресс", ref RowIndex);
                RowIndex++;
            }

            if (Additional4DT.Rows.Count > 0)
            {
                AdditionsToExcelSingly(ref hssfworkbook,
                    Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, Additional4DT, WorkAssignmentID, BatchName, ClientName, "Фрезерование ПП03-404", "Пресс", ref RowIndex);
                RowIndex++;
            }

            if (Additional5DT.Rows.Count > 0)
            {
                AdditionsToExcelSingly(ref hssfworkbook,
                    Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, Additional5DT, WorkAssignmentID, BatchName, ClientName, "Покраска", "Пресс", ref RowIndex);
            }
        }

        public void AdditionsToExcelSingly(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            ref HSSFSheet sheet1, DataTable DT, int WorkAssignmentID, string BatchName, string ClientName, string PageName, string colName, ref int RowIndex)
        {
            HSSFCell cell = null;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Плановое время выполнения:");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, PageName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, ClientName);
            cell.CellStyle = CalibriBold11CS;

            int displayIndex = 0;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), displayIndex++, "Профиль");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), displayIndex++, "Цвет профиля");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), displayIndex++, "Цвет наполнителя");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), displayIndex++, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), displayIndex++, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), displayIndex++, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), displayIndex++, colName);
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), displayIndex++, "Работник");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            int CType = 0;
            int Height = 0;
            int Width = 0;
            int Count = 0;
            int AllTotalAmount = 0;
            decimal AllTotalSquare = 0;

            if (DT.Rows.Count > 0)
            {
                CType = Convert.ToInt32(DT.Rows[0]["TechnoInsetColorID"]);
            }

            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value && DT.Rows[x]["Height"] != DBNull.Value && DT.Rows[x]["Width"] != DBNull.Value)
                {
                    Count = Convert.ToInt32(DT.Rows[x]["Count"]);
                    Height = Convert.ToInt32(DT.Rows[x]["Height"]);
                    Width = Convert.ToInt32(DT.Rows[x]["Width"]);
                    AllTotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    AllTotalSquare += Convert.ToDecimal(DT.Rows[x]["Square"]);
                }

                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "ColorID" || DT.Columns[y].ColumnName == "TechnoInsetColorID")
                        continue;
                    Type t = DT.Rows[x][y].GetType();

                    if (t.Name == "Decimal")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToDouble(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderDecCS;
                        continue;
                    }
                    if (t.Name == "Int32")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToInt32(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (t.Name == "String" || t.Name == "DBNull")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(DT.Rows[x][y].ToString());
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }
                }

                if (x + 1 <= DT.Rows.Count - 1 && CType != Convert.ToInt32(DT.Rows[x + 1]["TechnoInsetColorID"]))
                {
                    RowIndex++;
                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "ColorID" || DT.Columns[y].ColumnName == "TechnoInsetColorID")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }
                }

                if (x == DT.Rows.Count - 1)
                {
                    RowIndex++;
                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "ColorID" || DT.Columns[y].ColumnName == "TechnoInsetColorID")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    RowIndex++;
                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "ColorID" || DT.Columns[y].ColumnName == "TechnoInsetColorID")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;

                    AllTotalSquare = decimal.Round(AllTotalSquare, 3, MidpointRounding.AwayFromZero);

                    cell = sheet1.CreateRow(RowIndex).CreateCell(6);
                    cell.SetCellValue(Convert.ToDouble(AllTotalSquare));
                    cell.CellStyle = TableHeaderDecCS;
                }
                RowIndex++;
            }

            CalibriBold11CS = hssfworkbook.CreateCellStyle();
            CalibriBold11CS.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.BottomBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.BorderLeft = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.LeftBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.BorderRight = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.RightBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.BorderTop = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.TopBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.SetFont(CalibriBold11F);
        }

        public void AllInsetsToExcelSingly(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            ref HSSFSheet sheet1, DataTable DT, int WorkAssignmentID, string DispatchDate, string BatchName, string ClientName, string TableName, ref int RowIndex)
        {
            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Плановое время выполнения:");
            cell.CellStyle = Calibri11CS;
            RowIndex++;

            if (DispatchDate.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Отгрузка: " + DispatchDate);
                cell.CellStyle = CalibriBold11CS;
            }

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, TableName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, ClientName);
            cell.CellStyle = CalibriBold11CS;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Цвет наполнителя");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Работник");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            int TotalAmount = 0;
            int AllTotalAmount = 0;
            string str = string.Empty;

            int CType = -1;
            int DifferentDecorCount = 0;

            using (DataView DV = new DataView(DT))
            {
                DifferentDecorCount = DV.ToTable(true, new string[] { "InsetColorID" }).Rows.Count;
            }

            if (DT.Rows.Count > 0)
            {
                CType = Convert.ToInt32(DT.Rows[0]["InsetColorID"]);
            }
            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value)
                {
                    TotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    AllTotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                }

                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "InsetColorID" || DT.Columns[y].ColumnName == "TechnoInsetColorID" ||
                        DT.Columns[y].ColumnName == "GlassCount" || DT.Columns[y].ColumnName == "MegaCount")
                        continue;
                    Type t = DT.Rows[x][y].GetType();

                    if (DT.Columns[y].ColumnName == "Name")
                        str = DT.Rows[x][y].ToString();

                    if (t.Name == "Decimal")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToDouble(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderDecCS;
                        continue;
                    }
                    if (t.Name == "Int32")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToInt32(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (t.Name == "String" || t.Name == "DBNull")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(DT.Rows[x][y].ToString());
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }
                }

                if (x + 1 <= DT.Rows.Count - 1)
                {
                    if (CType != Convert.ToInt32(DT.Rows[x + 1]["InsetColorID"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "InsetColorID" || DT.Columns[y].ColumnName == "TechnoInsetColorID" ||
                                DT.Columns[y].ColumnName == "GlassCount" || DT.Columns[y].ColumnName == "MegaCount")
                                continue;
                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        CType = Convert.ToInt32(DT.Rows[x + 1]["InsetColorID"]);
                        TotalAmount = 0;

                        RowIndex++;
                    }
                }

                if (x == DT.Rows.Count - 1)
                {
                    if (DifferentDecorCount > 1)
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "InsetColorID" || DT.Columns[y].ColumnName == "TechnoInsetColorID" ||
                                DT.Columns[y].ColumnName == "GlassCount" || DT.Columns[y].ColumnName == "MegaCount")
                                continue;
                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;
                    }
                    RowIndex++;

                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "InsetColorID" || DT.Columns[y].ColumnName == "TechnoInsetColorID" ||
                            DT.Columns[y].ColumnName == "GlassCount" || DT.Columns[y].ColumnName == "MegaCount")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;
                }
                RowIndex++;
            }
        }

        public void ArchDecorAssemblyToExcelSingly(ref HSSFWorkbook hssfworkbook, ref HSSFSheet sheet1,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            DataTable DT, int WorkAssignmentID, string BatchName, string ClientName, string OperationName, string OrderName, string Notes, ref int RowIndex)
        {
            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
            cell.CellStyle = Calibri11CS;

            RowIndex++;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, ClientName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Заказ:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, OrderName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 4, OperationName);
            cell.CellStyle = CalibriBold11CS;
            if (Notes.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Примечание: " + Notes);
                cell.CellStyle = CalibriBold11CS;
            }

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Наименование");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Цвет");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Длин./Выс.");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 5, "Прим.");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            int DecorID = -1;
            int AllTotalAmount = 0;
            int TotalAmount = 0;
            int DifferentDecorCount = 0;

            using (DataView DV = new DataView(DT))
            {
                DifferentDecorCount = DV.ToTable(true, new string[] { "DecorID" }).Rows.Count;
            }

            if (DT.Rows.Count > 0)
            {
                DecorID = Convert.ToInt32(DT.Rows[0]["DecorID"]);
            }

            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value)
                {
                    AllTotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    TotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                }
                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "DecorID")
                        continue;

                    Type t = DT.Rows[x][y].GetType();

                    if (t.Name == "Decimal")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToDouble(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderDecCS;
                        continue;
                    }
                    if (t.Name == "Int32")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToInt32(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (t.Name == "String" || t.Name == "DBNull")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(DT.Rows[x][y].ToString());
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }
                }

                if (x + 1 <= DT.Rows.Count - 1)
                {
                    if (DecorID != Convert.ToInt32(DT.Rows[x + 1]["DecorID"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "DecorID")
                                continue;

                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        DecorID = Convert.ToInt32(DT.Rows[x + 1]["DecorID"]);
                        TotalAmount = 0;
                        RowIndex++;
                    }
                }

                if (x == DT.Rows.Count - 1)
                {
                    if (DifferentDecorCount > 1)
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "DecorID")
                                continue;

                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;
                    }
                    RowIndex++;

                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "DecorID")
                            continue;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;
                }
                RowIndex++;
            }
        }

        public void Assembly1ToExcelSingly(ref HSSFWorkbook hssfworkbook, ref HSSFSheet sheet1,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            DataTable DT, int WorkAssignmentID, string DispatchDate, string BatchName, string ClientName, string PageName, ref int RowIndex)
        {
            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Плановое время выполнения:");
            cell.CellStyle = Calibri11CS;
            RowIndex++;
            if (DispatchDate.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Отгрузка: " + DispatchDate);
                cell.CellStyle = CalibriBold11CS;
            }
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, PageName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, ClientName);
            cell.CellStyle = CalibriBold11CS;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Профиль");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Цвет профиля");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Цвет наполнителя");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 5, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 6, "Зачистка");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 7, "Обклад витрин");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            int CType = -1;
            int AllTotalAmount = 0;
            int TotalAmount = 0;
            decimal AllTotalSquare = 0;
            decimal TotalSquare = 0;

            if (DT.Rows.Count > 0)
            {
                CType = Convert.ToInt32(DT.Rows[0]["ColorType"]);
            }

            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value)
                {
                    AllTotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    TotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                }
                if (DT.Rows[x]["Square"] != DBNull.Value)
                {
                    AllTotalSquare += Convert.ToDecimal(DT.Rows[x]["Square"]);
                    TotalSquare += Convert.ToDecimal(DT.Rows[x]["Square"]);
                }

                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "ColorType" || DT.Columns[y].ColumnName == "Square")
                        continue;

                    Type t = DT.Rows[x][y].GetType();

                    if (t.Name == "Decimal")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToDouble(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderDecCS;
                        continue;
                    }
                    if (t.Name == "Int32")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToInt32(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (t.Name == "String" || t.Name == "DBNull")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(DT.Rows[x][y].ToString());
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }
                }
                if (x + 1 <= DT.Rows.Count - 1)
                {
                    if (CType != Convert.ToInt32(DT.Rows[x + 1]["ColorType"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "ColorType" || DT.Columns[y].ColumnName == "Square")
                                continue;
                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        CType = Convert.ToInt32(DT.Rows[x + 1]["ColorType"]);
                        TotalAmount = 0;
                        TotalSquare = 0;

                        RowIndex++;
                    }
                }

                if (x == DT.Rows.Count - 1)
                {
                    RowIndex++;
                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "ColorType" || DT.Columns[y].ColumnName == "Square")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Итого:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                    cell.SetCellValue(TotalAmount);
                    cell.CellStyle = TableHeaderCS;

                    RowIndex++;

                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "ColorType" || DT.Columns[y].ColumnName == "Square")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;
                }
                RowIndex++;
            }
        }

        public void Assembly2ToExcelSingly(ref HSSFWorkbook hssfworkbook, ref HSSFSheet sheet1,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            DataTable DT, int WorkAssignmentID, string DispatchDate, string BatchName, string ClientName, string PageName, ref int RowIndex)
        {
            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Плановое время выполнения:");
            cell.CellStyle = Calibri11CS;
            RowIndex++;
            if (DispatchDate.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Отгрузка: " + DispatchDate);
                cell.CellStyle = CalibriBold11CS;
            }
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, PageName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, ClientName);
            cell.CellStyle = CalibriBold11CS;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Профиль");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Цвет профиля");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Цвет наполнителя");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 5, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 6, "Сверление");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 7, "Работник");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            int CType = -1;
            int AllTotalAmount = 0;
            int TotalAmount = 0;
            decimal AllTotalSquare = 0;
            decimal TotalSquare = 0;

            if (DT.Rows.Count > 0)
            {
                CType = Convert.ToInt32(DT.Rows[0]["ColorType"]);
            }

            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value)
                {
                    AllTotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    TotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                }
                if (DT.Rows[x]["Square"] != DBNull.Value)
                {
                    AllTotalSquare += Convert.ToDecimal(DT.Rows[x]["Square"]);
                    TotalSquare += Convert.ToDecimal(DT.Rows[x]["Square"]);
                }

                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "FrontType" || DT.Columns[y].ColumnName == "ColorType" || DT.Columns[y].ColumnName == "Square")
                        continue;

                    Type t = DT.Rows[x][y].GetType();

                    if (t.Name == "Decimal")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToDouble(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderDecCS;
                        continue;
                    }
                    if (t.Name == "Int32")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToInt32(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (t.Name == "String" || t.Name == "DBNull")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(DT.Rows[x][y].ToString());
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }
                }
                if (x + 1 <= DT.Rows.Count - 1)
                {
                    if (CType != Convert.ToInt32(DT.Rows[x + 1]["ColorType"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "ColorType" || DT.Columns[y].ColumnName == "FrontType" || DT.Columns[y].ColumnName == "Square")
                                continue;
                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        CType = Convert.ToInt32(DT.Rows[x + 1]["ColorType"]);
                        TotalAmount = 0;
                        TotalSquare = 0;

                        RowIndex++;
                    }
                }

                if (x == DT.Rows.Count - 1)
                {
                    RowIndex++;
                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "ColorType" || DT.Columns[y].ColumnName == "Square")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Итого:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                    cell.SetCellValue(TotalAmount);
                    cell.CellStyle = TableHeaderCS;

                    RowIndex++;

                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "ColorType" || DT.Columns[y].ColumnName == "Square")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;
                }
                RowIndex++;
            }

            CalibriBold11CS = hssfworkbook.CreateCellStyle();
            CalibriBold11CS.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.BottomBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.BorderLeft = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.LeftBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.BorderRight = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.RightBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.BorderTop = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.TopBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.SetFont(CalibriBold11F);
        }

        public void AssemblyToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            DataTable DT1, DataTable DT2, int WorkAssignmentID, string DispatchDate, string BatchName, string ClientName)
        {
            int RowIndex = 0;
            HSSFSheet sheet1 = hssfworkbook.CreateSheet("Зачистка");
            sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

            sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
            sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
            sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
            sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

            sheet1.SetColumnWidth(0, 20 * 256);
            sheet1.SetColumnWidth(1, 17 * 256);
            sheet1.SetColumnWidth(2, 17 * 256);
            sheet1.SetColumnWidth(3, 6 * 256);
            sheet1.SetColumnWidth(4, 6 * 256);
            sheet1.SetColumnWidth(5, 6 * 256);
            sheet1.SetColumnWidth(6, 13 * 256);
            sheet1.SetColumnWidth(7, 13 * 256);

            {
                HSSFPatriarch patriarch = sheet1.CreateDrawingPatriarch();
                int ind = 0;

                CreateBarcode(WorkAssignmentID);
                HSSFClientAnchor anchor = new HSSFClientAnchor(0, 1, 1000, 8, 5, ind, 6, ind + 3)
                {
                    AnchorType = 2
                };
                string BarcodeNumber = GetBarcodeNumber(24, WorkAssignmentID);
                string barcodeName = System.Environment.GetEnvironmentVariable("TEMP") + @"/" + BarcodeNumber + ".jpg";
                HSSFPicture picture = patriarch.CreatePicture(anchor, LoadImage(barcodeName, hssfworkbook));

                HSSFFont BarcodeFont = hssfworkbook.CreateFont();
                BarcodeFont.FontHeightInPoints = 12;
                BarcodeFont.FontName = "Calibri";

                HSSFCellStyle BarcodeStyle = hssfworkbook.CreateCellStyle();
                BarcodeStyle.Alignment = HSSFCellStyle.ALIGN_CENTER;
                BarcodeStyle.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
                BarcodeStyle.SetFont(BarcodeFont);

                HSSFCell barcodeCell = sheet1.CreateRow(ind + 3).CreateCell(5);
                barcodeCell.SetCellValue(BarcodeNumber.ToString());
                barcodeCell.CellStyle = BarcodeStyle;
                sheet1.AddMergedRegion(new NPOI.HSSF.Util.Region(ind + 3, 5, ind + 3, 6));
            }

            DataTable DT = DT1.Copy();
            DataColumn Col1 = new DataColumn();
            DataColumn Col2 = new DataColumn();

            Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
            Col2 = DT.Columns.Add("Col2", System.Type.GetType("System.String"));
            Col1.SetOrdinal(6);
            Col2.SetOrdinal(7);

            if (DT.Rows.Count > 0)
            {
                Assembly1ToExcelSingly(ref hssfworkbook, ref sheet1,
                    Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DT, WorkAssignmentID, DispatchDate, BatchName, ClientName, "Зачистка", ref RowIndex);
                RowIndex++;
                RowIndex++;
            }

            DT.Dispose();
            Col1.Dispose();
            Col2.Dispose();
            DT = DT2.Copy();
            Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
            Col2 = DT.Columns.Add("Col2", System.Type.GetType("System.String"));
            Col1.SetOrdinal(6);
            Col2.SetOrdinal(7);

            if (DT.Rows.Count > 0)
            {
                Assembly1ToExcelSingly(ref hssfworkbook, ref sheet1,
                    Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DT, WorkAssignmentID, DispatchDate, BatchName, ClientName, "Зачистка", ref RowIndex);
            }

            RowIndex = 0;
            HSSFSheet sheet2 = hssfworkbook.CreateSheet("Сборка");
            sheet2.PrintSetup.PaperSize = (short)PaperSizeType.A4;

            sheet2.SetMargin(HSSFSheet.LeftMargin, (double).12);
            sheet2.SetMargin(HSSFSheet.RightMargin, (double).07);
            sheet2.SetMargin(HSSFSheet.TopMargin, (double).20);
            sheet2.SetMargin(HSSFSheet.BottomMargin, (double).20);

            sheet2.SetColumnWidth(0, 20 * 256);
            sheet2.SetColumnWidth(1, 17 * 256);
            sheet2.SetColumnWidth(2, 17 * 256);
            sheet2.SetColumnWidth(3, 6 * 256);
            sheet2.SetColumnWidth(4, 6 * 256);
            sheet2.SetColumnWidth(5, 6 * 256);
            sheet2.SetColumnWidth(6, 13 * 256);
            sheet2.SetColumnWidth(7, 13 * 256);

            {
                HSSFPatriarch patriarch = sheet2.CreateDrawingPatriarch();
                int ind = 0;

                CreateBarcode(WorkAssignmentID);
                HSSFClientAnchor anchor = new HSSFClientAnchor(0, 1, 1000, 8, 5, ind, 6, ind + 3)
                {
                    AnchorType = 2
                };
                string BarcodeNumber = GetBarcodeNumber(24, WorkAssignmentID);
                string barcodeName = System.Environment.GetEnvironmentVariable("TEMP") + @"/" + BarcodeNumber + ".jpg";
                HSSFPicture picture = patriarch.CreatePicture(anchor, LoadImage(barcodeName, hssfworkbook));

                HSSFFont BarcodeFont = hssfworkbook.CreateFont();
                BarcodeFont.FontHeightInPoints = 12;
                BarcodeFont.FontName = "Calibri";

                HSSFCellStyle BarcodeStyle = hssfworkbook.CreateCellStyle();
                BarcodeStyle.Alignment = HSSFCellStyle.ALIGN_CENTER;
                BarcodeStyle.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
                BarcodeStyle.SetFont(BarcodeFont);

                HSSFCell barcodeCell = sheet2.CreateRow(ind + 3).CreateCell(5);
                barcodeCell.SetCellValue(BarcodeNumber.ToString());
                barcodeCell.CellStyle = BarcodeStyle;
                sheet2.AddMergedRegion(new NPOI.HSSF.Util.Region(ind + 3, 5, ind + 3, 6));
            }

            DT.Dispose();
            Col1.Dispose();
            Col2.Dispose();
            DT = DT1.Copy();
            Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
            Col2 = DT.Columns.Add("Col2", System.Type.GetType("System.String"));
            Col1.SetOrdinal(6);
            Col2.SetOrdinal(7);

            if (DT.Rows.Count > 0)
            {
                Assembly2ToExcelSingly(ref hssfworkbook, ref sheet2,
                    Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DT, WorkAssignmentID, DispatchDate, BatchName, ClientName, "Сборка", ref RowIndex);
                RowIndex++;
                RowIndex++;
            }

            DT.Dispose();
            Col1.Dispose();
            Col2.Dispose();
            DT = DT2.Copy();
            Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
            Col2 = DT.Columns.Add("Col2", System.Type.GetType("System.String"));
            Col1.SetOrdinal(6);
            Col2.SetOrdinal(7);

            if (DT.Rows.Count > 0)
            {
                Assembly2ToExcelSingly(ref hssfworkbook, ref sheet2,
                    Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DT, WorkAssignmentID, DispatchDate, BatchName, ClientName, "Сборка", ref RowIndex);
            }
        }

        public void BagetWithAngleAssemblyToExcelSingly(ref HSSFWorkbook hssfworkbook, ref HSSFSheet sheet1,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            DataTable DT, int WorkAssignmentID, string BatchName, string ClientName, string OperationName, string OrderName, string Notes, ref int RowIndex)
        {
            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
            cell.CellStyle = Calibri11CS;

            RowIndex++;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, ClientName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Заказ:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, OrderName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 4, OperationName);
            cell.CellStyle = CalibriBold11CS;
            if (Notes.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Примечание: " + Notes);
                cell.CellStyle = CalibriBold11CS;
            }

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Наименование");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Цвет");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Длин./Выс.");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Л. угол");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 5, "П. угол");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 6, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 7, "Прим.");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            int DecorID = -1;
            int AllTotalAmount = 0;
            int TotalAmount = 0;
            int DifferentDecorCount = 0;

            using (DataView DV = new DataView(DT))
            {
                DifferentDecorCount = DV.ToTable(true, new string[] { "DecorID" }).Rows.Count;
            }
            if (DT.Rows.Count > 0)
            {
                DecorID = Convert.ToInt32(DT.Rows[0]["DecorID"]);
            }

            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value)
                {
                    AllTotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    TotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                }
                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "DecorID")
                        continue;

                    Type t = DT.Rows[x][y].GetType();

                    if (t.Name == "Decimal")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToDouble(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderDecCS;
                        continue;
                    }
                    if (t.Name == "Int32")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToInt32(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (t.Name == "String" || t.Name == "DBNull")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(DT.Rows[x][y].ToString());
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }
                }

                if (x + 1 <= DT.Rows.Count - 1)
                {
                    if (DecorID != Convert.ToInt32(DT.Rows[x + 1]["DecorID"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "DecorID")
                                continue;

                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(6);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        DecorID = Convert.ToInt32(DT.Rows[x + 1]["DecorID"]);
                        TotalAmount = 0;
                        RowIndex++;
                    }
                }

                if (x == DT.Rows.Count - 1)
                {
                    if (DifferentDecorCount > 1)
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "DecorID")
                                continue;

                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(6);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;
                    }
                    RowIndex++;

                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "DecorID")
                            continue;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(6);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;
                }
                RowIndex++;
            }
        }

        public void ClearOrders()
        {
            BagetWithAngelOrdersDT.Clear();
            NotArchDecorOrdersDT.Clear();
            ArchDecorOrdersDT.Clear();
            GridsDecorOrdersDT.Clear();

            FrontsID.Clear();
            TechnoNOrdersDT.Clear();
            AntaliaOrdersDT.Clear();
            Nord95OrdersDT.Clear();
            epFoxOrdersDT.Clear();
            VeneciaOrdersDT.Clear();
            BergamoOrdersDT.Clear();
            ep041OrdersDT.Clear();
            ep071OrdersDT.Clear();
            ep206OrdersDT.Clear();
            ep216OrdersDT.Clear();
            ep111OrdersDT.Clear();
            BostonOrdersDT.Clear();
            LeonOrdersDT.Clear();
            LimogOrdersDT.Clear();
            ep066Marsel4OrdersDT.Clear();
            ep110JersyOrdersDT.Clear();
            ep018Marsel1OrdersDT.Clear();
            ep043ShervudOrdersDT.Clear();
            ep112OrdersDT.Clear();
            UrbanOrdersDT.Clear();
            AlbyOrdersDT.Clear();
            BrunoOrdersDT.Clear();
            epsh406Techno4OrdersDT.Clear();
            LukOrdersDT.Clear();
            LukPVHOrdersDT.Clear();
            MilanoOrdersDT.Clear();
            PragaOrdersDT.Clear();
            SigmaOrdersDT.Clear();
            FatOrdersDT.Clear();
        }

        public void ComecToExcelSingly(ref HSSFWorkbook hssfworkbook, ref HSSFSheet sheet1,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            DataTable DT, int WorkAssignmentID, string SheetName, string DispatchDate, string BatchName, string ClientName, ref int RowIndex)
        {
            HSSFCell cell = null;
            int DisplayIndex = 0;

            if (DispatchDate.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Отгрузка: " + DispatchDate);
                cell.CellStyle = CalibriBold11CS;
            }
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, SheetName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, ClientName);
            cell.CellStyle = CalibriBold11CS;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Профиль");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Цвет профиля");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Работник");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            decimal SticksCount = 0;
            int CType = 0;
            int FType = 0;
            int PType = 0;
            int TotalAmount = 0;
            int AllTotalAmount = 0;
            int Count = 0;
            int Height = 0;

            if (DT.Rows.Count > 0)
            {
                CType = Convert.ToInt32(DT.Rows[0]["ColorType"]);
                FType = Convert.ToInt32(DT.Rows[0]["FrontType"]);
                PType = Convert.ToInt32(DT.Rows[0]["ProfileType"]);
            }

            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value && DT.Rows[x]["Height"] != DBNull.Value)
                {
                    Count = Convert.ToInt32(DT.Rows[x]["Count"]);
                    Height = Convert.ToInt32(DT.Rows[x]["Height"]);
                    SticksCount += (Height + 4) * Count;
                    TotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    AllTotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    Count = Convert.ToInt32(DT.Rows[x]["Count"]);
                }

                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "ProfileType" || DT.Columns[y].ColumnName == "ColorType"
                        || DT.Columns[y].ColumnName == "FrontType" || DT.Columns[y].ColumnName == "VitrinaCount"
                        || DT.Columns[y].ColumnName == "BoxCount" || DT.Columns[y].ColumnName == "ImpostCount")
                        continue;
                    Type t = DT.Rows[x][y].GetType();

                    if (t.Name == "Decimal")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToDouble(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderDecCS;
                        continue;
                    }
                    if (t.Name == "Int32")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToInt32(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (t.Name == "String" || t.Name == "DBNull")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(DT.Rows[x][y].ToString());
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }
                }

                if (DT.Rows[x]["ColorType"] != DBNull.Value)
                    CType = Convert.ToInt32(DT.Rows[x]["ColorType"]);
                if (DT.Rows[x]["FrontType"] != DBNull.Value)
                    FType = Convert.ToInt32(DT.Rows[x]["FrontType"]);
                if (DT.Rows[x]["ProfileType"] != DBNull.Value)
                    PType = Convert.ToInt32(DT.Rows[x]["ProfileType"]);
                if (x + 1 <= DT.Rows.Count - 1 &&
                    (FType != Convert.ToInt32(DT.Rows[x + 1]["FrontType"]) || PType != Convert.ToInt32(DT.Rows[x + 1]["ProfileType"]) || CType != Convert.ToInt32(DT.Rows[x + 1]["ColorType"])))
                {
                    RowIndex++;
                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "ProfileType" || DT.Columns[y].ColumnName == "ColorType"
                        || DT.Columns[y].ColumnName == "FrontType" || DT.Columns[y].ColumnName == "VitrinaCount"
                        || DT.Columns[y].ColumnName == "BoxCount" || DT.Columns[y].ColumnName == "ImpostCount")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    SticksCount = SticksCount * 1.15m / 2620;
                    SticksCount = decimal.Round(SticksCount, 1, MidpointRounding.AwayFromZero);
                    cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                    cell.SetCellValue(SticksCount + " палок");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Итого:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                    cell.SetCellValue(TotalAmount);
                    cell.CellStyle = TableHeaderCS;
                    RowIndex++;

                    CType = Convert.ToInt32(DT.Rows[x + 1]["ColorType"]);
                    PType = Convert.ToInt32(DT.Rows[x + 1]["ProfileType"]);
                    Count = 0;
                    Height = 0;
                    SticksCount = 0;
                    TotalAmount = 0;
                }

                if (x == DT.Rows.Count - 1)
                {
                    RowIndex++;
                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "ProfileType" || DT.Columns[y].ColumnName == "ColorType"
                        || DT.Columns[y].ColumnName == "FrontType" || DT.Columns[y].ColumnName == "VitrinaCount"
                        || DT.Columns[y].ColumnName == "BoxCount" || DT.Columns[y].ColumnName == "ImpostCount")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    SticksCount = SticksCount * 1.15m / 2620;
                    SticksCount = decimal.Round(SticksCount, 1, MidpointRounding.AwayFromZero);
                    cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                    cell.SetCellValue(SticksCount + " палок");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Итого:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                    cell.SetCellValue(TotalAmount);
                    cell.CellStyle = TableHeaderCS;

                    RowIndex++;
                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "ProfileType" || DT.Columns[y].ColumnName == "ColorType"
                        || DT.Columns[y].ColumnName == "FrontType" || DT.Columns[y].ColumnName == "VitrinaCount"
                        || DT.Columns[y].ColumnName == "BoxCount" || DT.Columns[y].ColumnName == "ImpostCount")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;
                }
                RowIndex++;
            }
        }

        public void CreateBarcode(int id)
        {
            Barcode Barcode = new Barcode();
            string BarcodeNumber = GetBarcodeNumber(24, id);
            string FileName = System.Environment.GetEnvironmentVariable("TEMP") + @"/" + BarcodeNumber + ".jpg";
            Image img = Barcode.GetBarcode(Barcode.BarcodeLength.Short, 45, BarcodeNumber);
            if (!File.Exists(FileName))
                img.Save(FileName, System.Drawing.Imaging.ImageFormat.Jpeg);
        }

        public void CreateExcel(int WorkAssignmentID, string ClientName, string BatchName, ref string sSourceFileName)
        {
            GetCurrentDate();

            System.Diagnostics.Stopwatch sw = new System.Diagnostics.Stopwatch();
            sw.Start();
            HSSFWorkbook hssfworkbook = new HSSFWorkbook();

            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = "NPOI Team";
            hssfworkbook.DocumentSummaryInformation = dsi;

            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
            si.Subject = "NPOI SDK Example";
            hssfworkbook.SummaryInformation = si;

            #region Create fonts and styles

            HSSFFont Calibri11F = hssfworkbook.CreateFont();
            Calibri11F.FontHeightInPoints = 11;
            Calibri11F.FontName = "Calibri";

            HSSFFont CalibriBold11F = hssfworkbook.CreateFont();
            CalibriBold11F.FontHeightInPoints = 11;
            CalibriBold11F.Boldweight = HSSFFont.BOLDWEIGHT_BOLD;
            CalibriBold11F.FontName = "Calibri";

            HSSFFont Serif8F = hssfworkbook.CreateFont();
            Serif8F.FontHeightInPoints = 8;
            Serif8F.FontName = "MS Sans Serif";

            HSSFFont Serif10F = hssfworkbook.CreateFont();
            Serif10F.FontHeightInPoints = 10;
            Serif10F.FontName = "MS Sans Serif";

            HSSFFont SerifBold10F = hssfworkbook.CreateFont();
            SerifBold10F.FontHeightInPoints = 10;
            SerifBold10F.FontName = "MS Sans Serif";
            SerifBold10F.Boldweight = HSSFFont.BOLDWEIGHT_BOLD;

            HSSFCellStyle Calibri11CS = hssfworkbook.CreateCellStyle();
            Calibri11CS.SetFont(Calibri11F);

            HSSFCellStyle CalibriBold11CS = hssfworkbook.CreateCellStyle();
            CalibriBold11CS.SetFont(CalibriBold11F);

            HSSFCellStyle TableHeaderCS = hssfworkbook.CreateCellStyle();
            TableHeaderCS.BorderBottom = HSSFCellStyle.BORDER_THIN;
            TableHeaderCS.BottomBorderColor = HSSFColor.BLACK.index;
            TableHeaderCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            TableHeaderCS.LeftBorderColor = HSSFColor.BLACK.index;
            TableHeaderCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            TableHeaderCS.RightBorderColor = HSSFColor.BLACK.index;
            TableHeaderCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            TableHeaderCS.TopBorderColor = HSSFColor.BLACK.index;
            TableHeaderCS.SetFont(Serif8F);

            HSSFCellStyle TableHeaderDecCS = hssfworkbook.CreateCellStyle();
            TableHeaderDecCS.BorderBottom = HSSFCellStyle.BORDER_THIN;
            TableHeaderDecCS.BottomBorderColor = HSSFColor.BLACK.index;
            TableHeaderDecCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            TableHeaderDecCS.LeftBorderColor = HSSFColor.BLACK.index;
            TableHeaderDecCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            TableHeaderDecCS.RightBorderColor = HSSFColor.BLACK.index;
            TableHeaderDecCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            TableHeaderDecCS.TopBorderColor = HSSFColor.BLACK.index;
            TableHeaderDecCS.SetFont(Serif8F);

            HSSFCellStyle TableContentCS = hssfworkbook.CreateCellStyle();
            TableContentCS.BorderBottom = HSSFCellStyle.BORDER_THIN;
            TableContentCS.BottomBorderColor = HSSFColor.BLACK.index;
            TableContentCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            TableContentCS.LeftBorderColor = HSSFColor.BLACK.index;
            TableContentCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            TableContentCS.RightBorderColor = HSSFColor.BLACK.index;
            TableContentCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            TableContentCS.TopBorderColor = HSSFColor.BLACK.index;
            TableContentCS.SetFont(SerifBold10F);

            HSSFCellStyle TotamAmountCS = hssfworkbook.CreateCellStyle();
            TotamAmountCS.Alignment = HSSFCellStyle.ALIGN_RIGHT;
            TotamAmountCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            TotamAmountCS.LeftBorderColor = HSSFColor.BLACK.index;
            TotamAmountCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            TotamAmountCS.RightBorderColor = HSSFColor.BLACK.index;
            TotamAmountCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            TotamAmountCS.TopBorderColor = HSSFColor.BLACK.index;
            TotamAmountCS.SetFont(SerifBold10F);

            HSSFCellStyle TotamAmountDecCS = hssfworkbook.CreateCellStyle();
            TotamAmountDecCS.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.000");
            TotamAmountCS.Alignment = HSSFCellStyle.ALIGN_RIGHT;
            TotamAmountCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            TotamAmountCS.LeftBorderColor = HSSFColor.BLACK.index;
            TotamAmountCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            TotamAmountCS.RightBorderColor = HSSFColor.BLACK.index;
            TotamAmountCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            TotamAmountCS.TopBorderColor = HSSFColor.BLACK.index;
            TotamAmountCS.SetFont(SerifBold10F);

            #endregion Create fonts and styles

            TechnoNSimpleDT.Clear();
            GetSimpleFronts(TechnoNOrdersDT, ref TechnoNSimpleDT);
            TechnoNVitrinaDT.Clear();
            GetVitrinaFronts(TechnoNOrdersDT, ref TechnoNVitrinaDT);
            TechnoNGridsDT.Clear();

            AntaliaSimpleDT.Clear();
            GetSimpleFronts(AntaliaOrdersDT, ref AntaliaSimpleDT);
            AntaliaVitrinaDT.Clear();
            GetVitrinaFronts(AntaliaOrdersDT, ref AntaliaVitrinaDT);
            AntaliaGridsDT.Clear();
            GetGridFronts(AntaliaOrdersDT, ref AntaliaGridsDT);

            Nord95SimpleDT.Clear();
            GetSimpleFronts(Nord95OrdersDT, ref Nord95SimpleDT);
            Nord95VitrinaDT.Clear();
            GetVitrinaFronts(Nord95OrdersDT, ref Nord95VitrinaDT);
            Nord95GridsDT.Clear();
            GetGridFronts(Nord95OrdersDT, ref Nord95GridsDT);

            epFoxSimpleDT.Clear();
            GetSimpleFronts(epFoxOrdersDT, ref epFoxSimpleDT);
            epFoxVitrinaDT.Clear();
            GetVitrinaFronts(epFoxOrdersDT, ref epFoxVitrinaDT);
            epFoxGridsDT.Clear();
            GetGridFronts(epFoxOrdersDT, ref epFoxGridsDT);

            VeneciaSimpleDT.Clear();
            GetSimpleFronts(VeneciaOrdersDT, ref VeneciaSimpleDT);
            VeneciaVitrinaDT.Clear();
            GetVitrinaFronts(VeneciaOrdersDT, ref VeneciaVitrinaDT);
            VeneciaGridsDT.Clear();
            GetGridFronts(VeneciaOrdersDT, ref VeneciaGridsDT);

            BergamoSimpleDT.Clear();
            GetSimpleFronts(BergamoOrdersDT, ref BergamoSimpleDT);
            BergamoVitrinaDT.Clear();
            GetVitrinaFronts(BergamoOrdersDT, ref BergamoVitrinaDT);
            BergamoGridsDT.Clear();
            GetGridFronts(BergamoOrdersDT, ref BergamoGridsDT);
            BergamoGlassDT.Clear();
            GetGlassFronts(BergamoOrdersDT, ref BergamoGlassDT);

            ep041SimpleDT.Clear();
            GetSimpleFronts(ep041OrdersDT, ref ep041SimpleDT);
            ep041VitrinaDT.Clear();
            GetVitrinaFronts(ep041OrdersDT, ref ep041VitrinaDT);
            ep041GridsDT.Clear();
            GetGridFronts(ep041OrdersDT, ref ep041GridsDT);

            ep071SimpleDT.Clear();
            GetSimpleFronts(ep071OrdersDT, ref ep071SimpleDT);
            ep071VitrinaDT.Clear();
            GetVitrinaFronts(ep071OrdersDT, ref ep071VitrinaDT);
            ep071GridsDT.Clear();
            GetGridFronts(ep071OrdersDT, ref ep071GridsDT);

            ep206SimpleDT.Clear();
            GetSimpleFronts(ep206OrdersDT, ref ep206SimpleDT);
            ep206VitrinaDT.Clear();
            GetVitrinaFronts(ep206OrdersDT, ref ep206VitrinaDT);
            ep206GridsDT.Clear();
            GetGridFronts(ep206OrdersDT, ref ep206GridsDT);

            ep216SimpleDT.Clear();
            GetSimpleFronts(ep216OrdersDT, ref ep216SimpleDT);
            ep216VitrinaDT.Clear();
            GetVitrinaFronts(ep216OrdersDT, ref ep216VitrinaDT);
            ep216GridsDT.Clear();
            GetGridFronts(ep216OrdersDT, ref ep216GridsDT);

            ep111SimpleDT.Clear();
            GetSimpleFronts(ep111OrdersDT, ref ep111SimpleDT);
            ep111VitrinaDT.Clear();
            GetVitrinaFronts(ep111OrdersDT, ref ep111VitrinaDT);
            ep111GridsDT.Clear();
            GetGridFronts(ep111OrdersDT, ref ep111GridsDT);

            BostonSimpleDT.Clear();
            GetSimpleFronts(BostonOrdersDT, ref BostonSimpleDT);
            BostonVitrinaDT.Clear();
            GetVitrinaFronts(BostonOrdersDT, ref BostonVitrinaDT);
            BostonGridsDT.Clear();
            GetGridFronts(BostonOrdersDT, ref BostonGridsDT);
            BostonGlassDT.Clear();
            GetGlassFronts(BostonOrdersDT, ref BostonGlassDT);

            LeonSimpleDT.Clear();
            GetSimpleFronts(LeonOrdersDT, ref LeonSimpleDT);
            LeonVitrinaDT.Clear();
            GetVitrinaFronts(LeonOrdersDT, ref LeonVitrinaDT);
            LeonGridsDT.Clear();
            GetGridFronts(LeonOrdersDT, ref LeonGridsDT);

            LimogSimpleDT.Clear();
            GetSimpleFronts(LimogOrdersDT, ref LimogSimpleDT);
            LimogVitrinaDT.Clear();
            GetVitrinaFronts(LimogOrdersDT, ref LimogVitrinaDT);
            LimogGridsDT.Clear();
            GetGridFronts(LimogOrdersDT, ref LimogGridsDT);

            ep066Marsel4SimpleDT.Clear();
            GetSimpleFronts(ep066Marsel4OrdersDT, ref ep066Marsel4SimpleDT);
            ep066Marsel4VitrinaDT.Clear();
            GetVitrinaFronts(ep066Marsel4OrdersDT, ref ep066Marsel4VitrinaDT);
            ep066Marsel4GridsDT.Clear();
            GetGridFronts(ep066Marsel4OrdersDT, ref ep066Marsel4GridsDT);

            ep110JersySimpleDT.Clear();
            GetSimpleFronts(ep110JersyOrdersDT, ref ep110JersySimpleDT);
            ep110JersyVitrinaDT.Clear();
            GetVitrinaFronts(ep110JersyOrdersDT, ref ep110JersyVitrinaDT);
            ep110JersyGridsDT.Clear();
            GetGridFronts(ep110JersyOrdersDT, ref ep110JersyGridsDT);

            ep018Marsel1SimpleDT.Clear();
            GetSimpleFronts(ep018Marsel1OrdersDT, ref ep018Marsel1SimpleDT);
            ep018Marsel1VitrinaDT.Clear();
            GetVitrinaFronts(ep018Marsel1OrdersDT, ref ep018Marsel1VitrinaDT);
            ep018Marsel1GridsDT.Clear();
            GetGridFronts(ep018Marsel1OrdersDT, ref ep018Marsel1GridsDT);

            ep043ShervudSimpleDT.Clear();
            GetSimpleFronts(ep043ShervudOrdersDT, ref ep043ShervudSimpleDT);
            ep043ShervudVitrinaDT.Clear();
            GetVitrinaFronts(ep043ShervudOrdersDT, ref ep043ShervudVitrinaDT);
            ep043ShervudGridsDT.Clear();
            GetGridFronts(ep043ShervudOrdersDT, ref ep043ShervudGridsDT);

            ep112SimpleDT.Clear();
            GetSimpleFronts(ep112OrdersDT, ref ep112SimpleDT);
            ep112VitrinaDT.Clear();
            GetVitrinaFronts(ep112OrdersDT, ref ep112VitrinaDT);
            ep112GridsDT.Clear();
            GetGridFronts(ep112OrdersDT, ref ep112GridsDT);

            UrbanSimpleDT.Clear();
            GetSimpleFronts(UrbanOrdersDT, ref UrbanSimpleDT);
            UrbanVitrinaDT.Clear();
            GetVitrinaFronts(UrbanOrdersDT, ref UrbanVitrinaDT);
            UrbanGridsDT.Clear();
            GetGridFronts(UrbanOrdersDT, ref UrbanGridsDT);

            AlbySimpleDT.Clear();
            GetSimpleFronts(AlbyOrdersDT, ref AlbySimpleDT);
            AlbyVitrinaDT.Clear();
            GetVitrinaFronts(AlbyOrdersDT, ref AlbyVitrinaDT);
            AlbyGridsDT.Clear();
            GetGridFronts(AlbyOrdersDT, ref AlbyGridsDT);

            BrunoSimpleDT.Clear();
            GetSimpleFronts(BrunoOrdersDT, ref BrunoSimpleDT);
            BrunoVitrinaDT.Clear();
            GetVitrinaFronts(BrunoOrdersDT, ref BrunoVitrinaDT);
            BrunoGridsDT.Clear();
            GetGridFronts(BrunoOrdersDT, ref BrunoGridsDT);

            epsh406Techno4SimpleDT.Clear();
            GetSimpleFronts(epsh406Techno4OrdersDT, ref epsh406Techno4SimpleDT);
            epsh406Techno4VitrinaDT.Clear();
            GetVitrinaFronts(epsh406Techno4OrdersDT, ref epsh406Techno4VitrinaDT);
            epsh406Techno4GridsDT.Clear();
            GetGridFronts(epsh406Techno4OrdersDT, ref epsh406Techno4GridsDT);

            LukSimpleDT.Clear();
            GetSimpleFronts(LukOrdersDT, ref LukSimpleDT);
            LukVitrinaDT.Clear();
            GetVitrinaFronts(LukOrdersDT, ref LukVitrinaDT);
            LukGridsDT.Clear();
            GetGridFronts(LukOrdersDT, ref LukGridsDT);

            LukPVHSimpleDT.Clear();
            GetSimpleFronts(LukPVHOrdersDT, ref LukPVHSimpleDT);
            LukPVHVitrinaDT.Clear();
            GetVitrinaFronts(LukPVHOrdersDT, ref LukPVHVitrinaDT);
            LukPVHGridsDT.Clear();
            GetGridFronts(LukPVHOrdersDT, ref LukPVHGridsDT);

            MilanoSimpleDT.Clear();
            GetSimpleFronts(MilanoOrdersDT, ref MilanoSimpleDT);
            MilanoVitrinaDT.Clear();
            GetVitrinaFronts(MilanoOrdersDT, ref MilanoVitrinaDT);
            MilanoGridsDT.Clear();
            GetGridFronts(MilanoOrdersDT, ref MilanoGridsDT);

            PragaSimpleDT.Clear();
            GetSimpleFronts(PragaOrdersDT, ref PragaSimpleDT);
            PragaVitrinaDT.Clear();
            GetVitrinaFronts(PragaOrdersDT, ref PragaVitrinaDT);
            PragaGridsDT.Clear();
            GetGridFronts(PragaOrdersDT, ref PragaGridsDT);

            SigmaSimpleDT.Clear();
            GetSimpleFronts(SigmaOrdersDT, ref SigmaSimpleDT);
            SigmaVitrinaDT.Clear();
            GetVitrinaFronts(SigmaOrdersDT, ref SigmaVitrinaDT);
            SigmaGridsDT.Clear();
            GetGridFronts(SigmaOrdersDT, ref SigmaGridsDT);
            SigmaGlassDT.Clear();
            GetGlassFronts(SigmaOrdersDT, ref SigmaGlassDT);

            FatSimpleDT.Clear();
            GetSimpleFronts(FatOrdersDT, ref FatSimpleDT);
            FatVitrinaDT.Clear();
            GetVitrinaFronts(FatOrdersDT, ref FatVitrinaDT);
            FatGridsDT.Clear();
            GetGridFronts(FatOrdersDT, ref FatGridsDT);

            if (TechnoNOrdersDT.Rows.Count == 0 && AntaliaOrdersDT.Rows.Count == 0 &&
                Nord95OrdersDT.Rows.Count == 0 && epFoxOrdersDT.Rows.Count == 0 && VeneciaOrdersDT.Rows.Count == 0 && BergamoOrdersDT.Rows.Count == 0 &&
                ep041OrdersDT.Rows.Count == 0 && ep071OrdersDT.Rows.Count == 0 && ep206OrdersDT.Rows.Count == 0 && ep216OrdersDT.Rows.Count == 0 &&
                 ep111OrdersDT.Rows.Count == 0 &&
                BostonOrdersDT.Rows.Count == 0 && LeonOrdersDT.Rows.Count == 0
                && LimogOrdersDT.Rows.Count == 0 && LukOrdersDT.Rows.Count == 0 && LukPVHOrdersDT.Rows.Count == 0 && MilanoOrdersDT.Rows.Count == 0
                && ep066Marsel4OrdersDT.Rows.Count == 0 && ep110JersyOrdersDT.Rows.Count == 0
                && ep018Marsel1OrdersDT.Rows.Count == 0 && ep043ShervudOrdersDT.Rows.Count == 0 && ep112OrdersDT.Rows.Count == 0 &&
                UrbanOrdersDT.Rows.Count == 0 && AlbyOrdersDT.Rows.Count == 0 && BrunoOrdersDT.Rows.Count == 0 && epsh406Techno4OrdersDT.Rows.Count == 0
                && PragaOrdersDT.Rows.Count == 0 && SigmaOrdersDT.Rows.Count == 0 && FatOrdersDT.Rows.Count == 0 &&
                BagetWithAngelOrdersDT.Rows.Count == 0 && NotArchDecorOrdersDT.Rows.Count == 0 && ArchDecorOrdersDT.Rows.Count == 0 && GridsDecorOrdersDT.Rows.Count == 0)
                return;

            string DispatchDate = string.Empty;
            if (ClientName == "ЗОВ" || ClientName == "Маркетинг + ЗОВ")
            {
                string FrontsFilterString = "(SELECT MainOrderID FROM FrontsOrders WHERE FactoryID=1 AND NOT (FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Marsel1) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Marsel3) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Marsel4) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Marsel5) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Porto) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Monte) +
                       " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Techno1) +
                       " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Shervud) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Techno2) +
                       " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Techno4) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.pFox) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Techno5) +
                       " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.PR1) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.PR2) +
                       " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.PR3) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.PRU8) + "))";
                string SelectCommand = @"SELECT DispatchDate, MegaOrderID FROM MegaOrders
                    WHERE MegaOrderID IN (SELECT MegaOrderID FROM MainOrders
                    WHERE MainOrderID IN" + FrontsFilterString + " AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE ProfilWorkAssignmentID=" + WorkAssignmentID + ")))";

                using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                    ConnectionStrings.ZOVOrdersConnectionString))
                {
                    using (DataTable DT = new DataTable())
                    {
                        if (DA.Fill(DT) > 0 && DT.Rows[0]["DispatchDate"] != DBNull.Value)
                            DispatchDate = Convert.ToDateTime(DT.Rows[0]["DispatchDate"]).ToString("dd.MM.yyyy");
                    }
                }
            }

            DataTable DistFrameColorsDT = DistFrameColorsTable(TechnoNOrdersDT, true);
            DataTable DT1 = AssemblyDT.Clone();
            DataTable DT2 = AssemblyDT.Clone();

            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), TechnoNSimpleDT, ref DT1);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), TechnoNVitrinaDT, ref DT2);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), TechnoNGridsDT, ref DT1);
            }
            DistFrameColorsDT.Clear();
            DistFrameColorsDT = DistFrameColorsTable(AntaliaOrdersDT, true);
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), AntaliaSimpleDT, ref DT1);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), AntaliaVitrinaDT, ref DT2);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), AntaliaGridsDT, ref DT1);
            }
            DistFrameColorsDT.Clear();
            DistFrameColorsDT = DistFrameColorsTable(Nord95OrdersDT, true);
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), Nord95SimpleDT, ref DT1);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), Nord95VitrinaDT, ref DT2);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), Nord95GridsDT, ref DT1);
            }
            DistFrameColorsDT.Clear();
            DistFrameColorsDT = DistFrameColorsTable(epFoxOrdersDT, true);
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), epFoxSimpleDT, ref DT1);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), epFoxVitrinaDT, ref DT2);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), epFoxGridsDT, ref DT1);
            }
            DistFrameColorsDT.Clear();
            DistFrameColorsDT = DistFrameColorsTable(VeneciaOrdersDT, true);
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), VeneciaSimpleDT, ref DT1);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), VeneciaVitrinaDT, ref DT2);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), VeneciaGridsDT, ref DT1);
            }
            DistFrameColorsDT.Clear();
            DistFrameColorsDT = DistFrameColorsTable(BergamoOrdersDT, true);
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), BergamoSimpleDT, ref DT1);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), BergamoVitrinaDT, ref DT2);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), BergamoGridsDT, ref DT1);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), BergamoGlassDT, ref DT1);
            }
            DistFrameColorsDT.Clear();
            DistFrameColorsDT = DistFrameColorsTable(ep041OrdersDT, true);
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), ep041SimpleDT, ref DT1);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), ep041VitrinaDT, ref DT2);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), ep041GridsDT, ref DT1);
            }
            DistFrameColorsDT.Clear();
            DistFrameColorsDT = DistFrameColorsTable(ep071OrdersDT, true);
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), ep071SimpleDT, ref DT1);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), ep071VitrinaDT, ref DT2);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), ep071GridsDT, ref DT1);
            }
            DistFrameColorsDT.Clear();
            DistFrameColorsDT = DistFrameColorsTable(ep206OrdersDT, true);
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), ep206SimpleDT, ref DT1);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), ep206VitrinaDT, ref DT2);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), ep206GridsDT, ref DT1);
            }
            DistFrameColorsDT.Clear();
            DistFrameColorsDT = DistFrameColorsTable(ep216OrdersDT, true);
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), ep216SimpleDT, ref DT1);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), ep216VitrinaDT, ref DT2);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), ep216GridsDT, ref DT1);
            }
            DistFrameColorsDT.Clear();
            DistFrameColorsDT = DistFrameColorsTable(ep111OrdersDT, true);
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), ep111SimpleDT, ref DT1);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), ep111VitrinaDT, ref DT2);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), ep111GridsDT, ref DT1);
            }
            DistFrameColorsDT.Clear();
            DistFrameColorsDT = DistFrameColorsTable(BostonOrdersDT, true);
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), BostonSimpleDT, ref DT1);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), BostonVitrinaDT, ref DT2);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), BostonGridsDT, ref DT1);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), BostonGlassDT, ref DT1);
            }
            DistFrameColorsDT.Clear();
            DistFrameColorsDT = DistFrameColorsTable(LeonOrdersDT, true);
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), LeonSimpleDT, ref DT1);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), LeonVitrinaDT, ref DT2);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), LeonGridsDT, ref DT1);
            }
            DistFrameColorsDT.Clear();
            DistFrameColorsDT = DistFrameColorsTable(LimogOrdersDT, true);
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), LimogSimpleDT, ref DT1);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), LimogVitrinaDT, ref DT2);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), LimogGridsDT, ref DT1);
            }
            DistFrameColorsDT.Clear();
            DistFrameColorsDT = DistFrameColorsTable(ep066Marsel4OrdersDT, true);
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), ep066Marsel4SimpleDT, ref DT1);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), ep066Marsel4VitrinaDT, ref DT2);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), ep066Marsel4GridsDT, ref DT1);
            }
            DistFrameColorsDT.Clear();
            DistFrameColorsDT = DistFrameColorsTable(ep110JersyOrdersDT, true);
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), ep110JersySimpleDT, ref DT1);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), ep110JersyVitrinaDT, ref DT2);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), ep110JersyGridsDT, ref DT1);
            }
            DistFrameColorsDT.Clear();
            DistFrameColorsDT = DistFrameColorsTable(ep018Marsel1OrdersDT, true);
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), ep018Marsel1SimpleDT, ref DT1);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), ep018Marsel1VitrinaDT, ref DT2);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), ep018Marsel1GridsDT, ref DT1);
            }
            DistFrameColorsDT.Clear();
            DistFrameColorsDT = DistFrameColorsTable(ep043ShervudOrdersDT, true);
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), ep043ShervudSimpleDT, ref DT1);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), ep043ShervudVitrinaDT, ref DT2);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), ep043ShervudGridsDT, ref DT1);
            }
            DistFrameColorsDT.Clear();
            DistFrameColorsDT = DistFrameColorsTable(ep112OrdersDT, true);
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), ep112SimpleDT, ref DT1);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), ep112VitrinaDT, ref DT2);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), ep112GridsDT, ref DT1);
            }
            DistFrameColorsDT.Clear();
            DistFrameColorsDT = DistFrameColorsTable(UrbanOrdersDT, true);
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), UrbanSimpleDT, ref DT1);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), UrbanVitrinaDT, ref DT2);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), UrbanGridsDT, ref DT1);
            }
            DistFrameColorsDT.Clear();
            DistFrameColorsDT = DistFrameColorsTable(AlbyOrdersDT, true);
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), AlbySimpleDT, ref DT1);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), AlbyVitrinaDT, ref DT2);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), AlbyGridsDT, ref DT1);
            }
            DistFrameColorsDT.Clear();
            DistFrameColorsDT = DistFrameColorsTable(BrunoOrdersDT, true);
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), BrunoSimpleDT, ref DT1);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), BrunoVitrinaDT, ref DT2);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), BrunoGridsDT, ref DT1);
            }
            DistFrameColorsDT.Clear();
            DistFrameColorsDT = DistFrameColorsTable(epsh406Techno4OrdersDT, true);
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), epsh406Techno4SimpleDT, ref DT1);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), epsh406Techno4VitrinaDT, ref DT2);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), epsh406Techno4GridsDT, ref DT1);
            }
            DistFrameColorsDT.Clear();
            DistFrameColorsDT = DistFrameColorsTable(LukOrdersDT, true);
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), LukSimpleDT, ref DT1);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), LukVitrinaDT, ref DT2);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), LukGridsDT, ref DT1);
            }
            DistFrameColorsDT.Clear();
            DistFrameColorsDT = DistFrameColorsTable(LukPVHOrdersDT, true);
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), LukPVHSimpleDT, ref DT1);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), LukPVHVitrinaDT, ref DT2);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), LukPVHGridsDT, ref DT1);
            }
            DistFrameColorsDT.Clear();
            DistFrameColorsDT = DistFrameColorsTable(MilanoOrdersDT, true);
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), MilanoSimpleDT, ref DT1);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), MilanoVitrinaDT, ref DT2);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), MilanoGridsDT, ref DT1);
            }
            DistFrameColorsDT.Clear();
            DistFrameColorsDT = DistFrameColorsTable(PragaOrdersDT, true);
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), PragaSimpleDT, ref DT1);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), PragaVitrinaDT, ref DT2);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), PragaGridsDT, ref DT1);
            }
            DistFrameColorsDT.Clear();
            DistFrameColorsDT = DistFrameColorsTable(SigmaOrdersDT, true);
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), SigmaSimpleDT, ref DT1);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), SigmaVitrinaDT, ref DT2);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), SigmaGridsDT, ref DT1);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), SigmaGlassDT, ref DT1);
            }
            DistFrameColorsDT.Clear();
            DistFrameColorsDT = DistFrameColorsTable(FatOrdersDT, true);
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), FatSimpleDT, ref DT1);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), FatVitrinaDT, ref DT2);
                AssemblySingly(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), FatGridsDT, ref DT1);
            }

            BagetWithAngleAssemblyByMainOrderToExcel(ref hssfworkbook, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, BatchName);

            NotArchDecorAssemblyByMainOrderToExcel(ref hssfworkbook, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, BatchName);

            ArchDecorAssemblyByMainOrderToExcel(ref hssfworkbook, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, BatchName);

            GridsDecorAssemblyByMainOrderToExcel(ref hssfworkbook, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, BatchName);

            MartinToExcel(ref hssfworkbook,
                Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS,
                WorkAssignmentID, DispatchDate, BatchName, ClientName);

            RapidToExcel(ref hssfworkbook,
                Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS,
                WorkAssignmentID, DispatchDate, BatchName, ClientName);

            InsetToExcel(ref hssfworkbook,
                Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS,
                WorkAssignmentID, DispatchDate, BatchName, ClientName);

            AdditionsToExcel(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, BatchName, ClientName);

            AssemblyToExcel(ref hssfworkbook,
                Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS,
                DT1, DT2, WorkAssignmentID, DispatchDate, BatchName, ClientName);

            OrdersSummaryInfoToExcel(ref hssfworkbook,
                Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS,
                WorkAssignmentID, DispatchDate, BatchName, ClientName);

            DeyingByMainOrderToExcel(ref hssfworkbook,
                Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, BatchName);

            GetMainOrdersSummary(ref hssfworkbook,
                Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS,
                WorkAssignmentID, BatchName);

            string FileName = WorkAssignmentID + " " + BatchName + "  Угол 45";
            string tempFolder = @"\\192.168.1.6\Public\USERS_2016\_ДЕЙСТВУЮЩИЕ\ПРОИЗВОДСТВО\ПРОФИЛЬ\инфиниум\";
            //string tempFolder = System.Environment.GetEnvironmentVariable("TEMP");

            string CurrentMonthName = DateTime.Now.ToString("MMMM");
            tempFolder = Path.Combine(tempFolder, CurrentMonthName);
            if (!(Directory.Exists(tempFolder)))
            {
                Directory.CreateDirectory(tempFolder);
            }
            FileInfo file = new FileInfo(tempFolder + @"\" + FileName + ".xls");
            int j = 1;
            while (file.Exists == true)
            {
                file = new FileInfo(tempFolder + @"\" + FileName + "(" + j++ + ").xls");
            }

            FileStream NewFile = new FileStream(file.FullName, FileMode.Create);
            hssfworkbook.Write(NewFile);
            NewFile.Close();

            sw.Stop();
            System.Diagnostics.Process.Start(file.FullName);
        }

        public void DeyingByMainOrderToExcelSingly(ref HSSFWorkbook hssfworkbook, ref HSSFSheet sheet1,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            DataTable DT, int WorkAssignmentID, string BatchName, string ClientName, string OrderName, string Notes, ref int RowIndex)
        {
            DataTable TempDT = new DataTable();
            DataColumn Col1 = new DataColumn("Col1", System.Type.GetType("System.String"));
            DataColumn Col2 = new DataColumn("Col2", System.Type.GetType("System.String"));
            DataColumn Col3 = new DataColumn("Col3", System.Type.GetType("System.String"));

            if (DT.Rows.Count > 0)
            {
                TempDT.Dispose();
                Col1.Dispose();
                Col2.Dispose();
                Col3.Dispose();
                TempDT = DT.Copy();
                Col1 = TempDT.Columns.Add("Col1", System.Type.GetType("System.String"));
                Col2 = TempDT.Columns.Add("Col2", System.Type.GetType("System.String"));
                Col1.SetOrdinal(6);
                Col2.SetOrdinal(7);
                TempDT.Columns["Square"].SetOrdinal(8);

                DyeingPackingToExcel(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, TempDT, WorkAssignmentID, BatchName, ClientName, OrderName,
                    "Упаковка. (" + Security.CurrentUserShortName + " от " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + ")", Notes, ref RowIndex);
                RowIndex++;

                TempDT.Dispose();
                Col1.Dispose();
                Col2.Dispose();
                Col3.Dispose();
                TempDT = DT.Copy();
                Col1 = TempDT.Columns.Add("Col1", System.Type.GetType("System.String"));
                Col1.SetOrdinal(6);
                TempDT.Columns["Square"].SetOrdinal(7);
                TempDT.Columns["Notes"].SetOrdinal(8);

                DyeingBoringToExcel(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, TempDT, WorkAssignmentID, BatchName, ClientName, OrderName,
                    "Сверление. (" + Security.CurrentUserShortName + " от " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + ")", Notes, ref RowIndex);
            }

            RowIndex++;
        }

        public void DyeingBoringToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            ref HSSFSheet sheet1, DataTable DT, int WorkAssignmentID, string BatchName, string ClientName, string OrderName, string PageName, string Notes,
            ref int RowIndex)
        {
            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
            cell.CellStyle = Calibri11CS;

            RowIndex++;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, PageName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, ClientName);
            cell.CellStyle = CalibriBold11CS;
            if (OrderName.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Заказ:");
                cell.CellStyle = CalibriBold11CS;
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, OrderName);
                cell.CellStyle = CalibriBold11CS;
            }
            if (Notes.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Примечание: " + Notes);
                cell.CellStyle = CalibriBold11CS;
            }

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Профиль");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Цвет профиля");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Цвет наполнителя");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 5, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 6, "Сверление");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 7, "м.кв.");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 8, "Прим.");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            int CType = -1;
            int AllTotalAmount = 0;
            int TotalAmount = 0;
            decimal AllTotalSquare = 0;
            decimal TotalSquare = 0;
            int DifferentDecorCount = 0;

            using (DataView DV = new DataView(DT))
            {
                DifferentDecorCount = DV.ToTable(true, new string[] { "ColorType" }).Rows.Count;
            }

            if (DT.Rows.Count > 0)
                CType = Convert.ToInt32(DT.Rows[0]["ColorType"]);

            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value)
                {
                    AllTotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    TotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                }
                if (DT.Rows[x]["Square"] != DBNull.Value)
                {
                    AllTotalSquare += Convert.ToDecimal(DT.Rows[x]["Square"]);
                    TotalSquare += Convert.ToDecimal(DT.Rows[x]["Square"]);
                }

                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "ColorType")
                        continue;

                    Type t = DT.Rows[x][y].GetType();

                    if (t.Name == "Decimal")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToDouble(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderDecCS;
                        continue;
                    }
                    if (t.Name == "Int32")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToInt32(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (t.Name == "String" || t.Name == "DBNull")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(DT.Rows[x][y].ToString());
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }
                }

                if (x + 1 <= DT.Rows.Count - 1)
                {
                    if (CType != Convert.ToInt32(DT.Rows[x + 1]["ColorType"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "ColorType")
                                continue;
                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(7);
                        cell.SetCellValue(Convert.ToDouble(TotalSquare));
                        cell.CellStyle = TableHeaderCS;

                        CType = Convert.ToInt32(DT.Rows[x + 1]["ColorType"]);
                        TotalAmount = 0;
                        TotalSquare = 0;
                        RowIndex++;
                    }
                }

                if (x == DT.Rows.Count - 1)
                {
                    if (DifferentDecorCount > 1)
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "ColorType")
                                continue;
                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(7);
                        cell.SetCellValue(Convert.ToDouble(TotalSquare));
                        cell.CellStyle = TableHeaderDecCS;
                    }
                    RowIndex++;

                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "ColorType")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(7);
                    cell.SetCellValue(Convert.ToDouble(AllTotalSquare));
                    cell.CellStyle = TableHeaderDecCS;
                }
                RowIndex++;
            }
        }

        public void DyeingPackingToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            ref HSSFSheet sheet1, DataTable DT, int WorkAssignmentID, string BatchName, string ClientName, string OrderName, string PageName, string Notes,
            ref int RowIndex)
        {
            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
            cell.CellStyle = Calibri11CS;

            RowIndex++;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, PageName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, ClientName);
            cell.CellStyle = CalibriBold11CS;
            if (OrderName.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Заказ:");
                cell.CellStyle = CalibriBold11CS;
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, OrderName);
                cell.CellStyle = CalibriBold11CS;
            }
            if (Notes.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Примечание: " + Notes);
                cell.CellStyle = CalibriBold11CS;
            }

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Профиль");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Цвет профиля");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Цвет наполнителя");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 5, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 6, "Пленка");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 7, "Упаковка");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 8, "м.кв.");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            int CType = -1;
            int AllTotalAmount = 0;
            int TotalAmount = 0;
            decimal AllTotalSquare = 0;
            decimal TotalSquare = 0;
            int DifferentDecorCount = 0;

            using (DataView DV = new DataView(DT))
            {
                DifferentDecorCount = DV.ToTable(true, new string[] { "ColorType" }).Rows.Count;
            }

            if (DT.Rows.Count > 0)
                CType = Convert.ToInt32(DT.Rows[0]["ColorType"]);

            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value)
                {
                    AllTotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    TotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                }
                if (DT.Rows[x]["Square"] != DBNull.Value)
                {
                    AllTotalSquare += Convert.ToDecimal(DT.Rows[x]["Square"]);
                    TotalSquare += Convert.ToDecimal(DT.Rows[x]["Square"]);
                }

                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "ColorType" || DT.Columns[y].ColumnName == "Notes")
                        continue;

                    Type t = DT.Rows[x][y].GetType();

                    if (t.Name == "Decimal")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToDouble(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderDecCS;
                        continue;
                    }
                    if (t.Name == "Int32")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToInt32(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (t.Name == "String" || t.Name == "DBNull")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(DT.Rows[x][y].ToString());
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }
                }

                if (x + 1 <= DT.Rows.Count - 1)
                {
                    if (CType != Convert.ToInt32(DT.Rows[x + 1]["ColorType"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "ColorType" || DT.Columns[y].ColumnName == "Notes")
                                continue;

                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(8);
                        cell.SetCellValue(Convert.ToDouble(TotalSquare));
                        cell.CellStyle = TableHeaderCS;

                        CType = Convert.ToInt32(DT.Rows[x + 1]["ColorType"]);
                        TotalAmount = 0;
                        TotalSquare = 0;
                        RowIndex++;
                    }
                }

                if (x == DT.Rows.Count - 1)
                {
                    if (DifferentDecorCount > 1)
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "ColorType" || DT.Columns[y].ColumnName == "Notes")
                                continue;

                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(8);
                        cell.SetCellValue(Convert.ToDouble(TotalSquare));
                        cell.CellStyle = TableHeaderDecCS;
                    }
                    RowIndex++;

                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "ColorType" || DT.Columns[y].ColumnName == "Notes")
                            continue;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(8);
                    cell.SetCellValue(Convert.ToDouble(AllTotalSquare));
                    cell.CellStyle = TableHeaderDecCS;
                }
                RowIndex++;
            }
        }

        public void FilenkaToExcelSingly(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            ref HSSFSheet sheet1, DataTable DT, int WorkAssignmentID, string DispatchDate, string BatchName, string ClientName, string TableName, ref int RowIndex)
        {
            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Плановое время выполнения:");
            cell.CellStyle = Calibri11CS;

            if (DispatchDate.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Отгрузка: " + DispatchDate);
                cell.CellStyle = CalibriBold11CS;
            }

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, TableName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, ClientName);
            cell.CellStyle = CalibriBold11CS;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Цвет наполнителя");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Работник");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            decimal TotalSquare = 0;
            decimal AllTotalSquare = 0;
            string str = string.Empty;

            int CType = -1;
            int DifferentDecorCount = 0;

            using (DataView DV = new DataView(DT))
            {
                DifferentDecorCount = DV.ToTable(true, new string[] { "InsetColorID" }).Rows.Count;
            }

            if (DT.Rows.Count > 0)
            {
                CType = Convert.ToInt32(DT.Rows[0]["InsetColorID"]);
            }
            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value)
                {
                    TotalSquare += Convert.ToDecimal(DT.Rows[x]["Height"]) * Convert.ToDecimal(DT.Rows[x]["Width"]) * Convert.ToDecimal(DT.Rows[x]["Count"]) / 1000000;
                    AllTotalSquare += Convert.ToDecimal(DT.Rows[x]["Height"]) * Convert.ToDecimal(DT.Rows[x]["Width"]) * Convert.ToDecimal(DT.Rows[x]["Count"]) / 1000000;
                }

                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "InsetColorID" || DT.Columns[y].ColumnName == "TechnoInsetColorID" ||
                        DT.Columns[y].ColumnName == "GlassCount" || DT.Columns[y].ColumnName == "MegaCount")
                        continue;
                    Type t = DT.Rows[x][y].GetType();

                    if (DT.Columns[y].ColumnName == "Name")
                        str = DT.Rows[x][y].ToString();

                    if (t.Name == "Decimal")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToDouble(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderDecCS;
                        continue;
                    }
                    if (t.Name == "Int32")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToInt32(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (t.Name == "String" || t.Name == "DBNull")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(DT.Rows[x][y].ToString());
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }
                }

                if (x + 1 <= DT.Rows.Count - 1)
                {
                    if (CType != Convert.ToInt32(DT.Rows[x + 1]["InsetColorID"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "InsetColorID" || DT.Columns[y].ColumnName == "TechnoInsetColorID" ||
                                DT.Columns[y].ColumnName == "GlassCount" || DT.Columns[y].ColumnName == "MegaCount")
                                continue;
                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        TotalSquare = decimal.Round(TotalSquare, 2, MidpointRounding.AwayFromZero);

                        cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                        cell.SetCellValue(Convert.ToDouble(TotalSquare));
                        cell.CellStyle = TableHeaderDecCS;

                        CType = Convert.ToInt32(DT.Rows[x + 1]["InsetColorID"]);
                        TotalSquare = 0;

                        RowIndex++;
                    }
                }

                if (x == DT.Rows.Count - 1)
                {
                    if (DifferentDecorCount > 1)
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "InsetColorID" || DT.Columns[y].ColumnName == "TechnoInsetColorID" ||
                                DT.Columns[y].ColumnName == "GlassCount" || DT.Columns[y].ColumnName == "MegaCount")
                                continue;
                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        TotalSquare = decimal.Round(TotalSquare, 2, MidpointRounding.AwayFromZero);

                        cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                        cell.SetCellValue(Convert.ToDouble(TotalSquare));
                        cell.CellStyle = TableHeaderDecCS;
                    }
                    RowIndex++;

                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "InsetColorID" || DT.Columns[y].ColumnName == "TechnoInsetColorID" ||
                            DT.Columns[y].ColumnName == "GlassCount" || DT.Columns[y].ColumnName == "MegaCount")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    AllTotalSquare = decimal.Round(AllTotalSquare, 2, MidpointRounding.AwayFromZero);

                    cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                    cell.SetCellValue(Convert.ToDouble(AllTotalSquare));
                    cell.CellStyle = TableHeaderDecCS;
                }
                RowIndex++;
            }
        }

        public string GetBarcodeNumber(int BarcodeType, int iNumber)
        {
            string Type = "";
            if (BarcodeType.ToString().Length == 1)
                Type = "00" + BarcodeType.ToString();
            if (BarcodeType.ToString().Length == 2)
                Type = "0" + BarcodeType.ToString();
            if (BarcodeType.ToString().Length == 3)
                Type = BarcodeType.ToString();

            string sNumber = "";
            if (iNumber.ToString().Length == 1)
                sNumber = "00000000" + iNumber.ToString();
            if (iNumber.ToString().Length == 2)
                sNumber = "0000000" + iNumber.ToString();
            if (iNumber.ToString().Length == 3)
                sNumber = "000000" + iNumber.ToString();
            if (iNumber.ToString().Length == 4)
                sNumber = "00000" + iNumber.ToString();
            if (iNumber.ToString().Length == 5)
                sNumber = "0000" + iNumber.ToString();
            if (iNumber.ToString().Length == 6)
                sNumber = "000" + iNumber.ToString();
            if (iNumber.ToString().Length == 7)
                sNumber = "00" + iNumber.ToString();
            if (iNumber.ToString().Length == 8)
                sNumber = "0" + iNumber.ToString();

            StringBuilder BarcodeNumber = new StringBuilder(Type);
            BarcodeNumber.Append(sNumber);

            return BarcodeNumber.ToString();
        }

        public string GetColorName(int ColorID)
        {
            string ColorName = string.Empty;
            try
            {
                DataRow[] Rows = FrameColorsDataTable.Select("ColorID = " + ColorID);
                ColorName = Rows[0]["ColorName"].ToString();
            }
            catch
            {
                return string.Empty;
            }
            return ColorName;
        }

        public void GetCurrentDate()
        {
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT GETDATE()", ConnectionStrings.UsersConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    DA.Fill(DT);

                    CurrentDate = Convert.ToDateTime(DT.Rows[0][0]);
                }
            }
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
            string FrontName = "";
            try
            {
                DataRow[] Rows = FrontsDataTable.Select("FrontID = " + FrontID);
                FrontName = Rows[0]["FrontName"].ToString();
            }
            catch
            {
                return "";
            }
            return FrontName;
        }

        public string GetInsetColorName(int ColorID)
        {
            string ColorName = string.Empty;
            try
            {
                DataRow[] Rows = InsetColorsDataTable.Select("InsetColorID = " + ColorID);
                ColorName = Rows[0]["InsetColorName"].ToString();
            }
            catch
            {
                return string.Empty;
            }
            return ColorName;
        }

        public string GetInsetTypeName(int InsetTypeID)
        {
            string InsetType = string.Empty;
            try
            {
                DataRow[] Rows = InsetTypesDataTable.Select("InsetTypeID = " + InsetTypeID);
                InsetType = Rows[0]["InsetTypeName"].ToString();
            }
            catch
            {
                return string.Empty;
            }
            return InsetType;
        }

        public void GetMainOrdersSummary(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            int WorkAssignmentID, string BatchName)
        {
            int MainOrderID = 0;
            int OrderNumber = 0;
            string ClientName = string.Empty;
            string DispatchDate = string.Empty;
            string OrderName = string.Empty;
            string Notes = string.Empty;
            string SelectCommand = string.Empty;
            DataTable DistClientNamesDT = new DataTable();
            DataTable DistMainOrdersDT = new DataTable();
            DataTable DT = new DataTable();

            //            SelectCommand = @"SELECT infiniu2_marketingreference.dbo.Clients.ClientName, MegaOrders.ClientID, MegaOrders.OrderNumber, MainOrders.MainOrderID, MainOrders.Notes AS MNotes,
            //                FrontsOrdersID, FrontsOrders.MainOrderID, FrontID, FrontTypeID, InsetTypeID,
            //                ColorID, InsetColorID, AdditionalColorID, Height, Width, Count, FrontConfigID, FrontsOrders.Notes FROM FrontsOrders
            //                INNER JOIN MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID
            //                INNER JOIN MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID
            //                INNER JOIN infiniu2_marketingreference.dbo.Clients ON MegaOrders.ClientID=infiniu2_marketingreference.dbo.Clients.ClientID
            //                INNER JOIN BatchDetails ON FrontsOrders.MainOrderID = BatchDetails.MainOrderID AND BatchDetails.FactoryID = 1
            //                INNER JOIN Batch ON BatchDetails.BatchID = Batch.BatchID AND Batch.ProfilWorkAssignmentID = " + WorkAssignmentID +
            //                @" WHERE FrontsOrders.FactoryID=1 AND FrontID IN (" +
            //                Convert.ToInt32(Fronts.Antalia) + "," + Convert.ToInt32(Fronts.Fat) + "," + Convert.ToInt32(Fronts.Leon) + "," + Convert.ToInt32(Fronts.LeonF) + "," +
            //                Convert.ToInt32(Fronts.Limog) + "," + Convert.ToInt32(Fronts.Luk) + "," + Convert.ToInt32(Fronts.LukPVH) + "," +
            //                Convert.ToInt32(Fronts.Milano) + "," + Convert.ToInt32(Fronts.MilanoK) + "," + Convert.ToInt32(Fronts.MilanoKF) + "," +
            //                Convert.ToInt32(Fronts.Praga) + "," + Convert.ToInt32(Fronts.Sigma) + "," + Convert.ToInt32(Fronts.Venecia) + "," + Convert.ToInt32(Fronts.VeneciaF) + ")";

            SelectCommand = @"SELECT infiniu2_marketingreference.dbo.Clients.ClientName, MegaOrders.ClientID, MegaOrders.OrderNumber, MainOrders.MainOrderID, MainOrders.Notes AS MNotes,
                            FrontsOrdersID, FrontsOrders.MainOrderID, FrontID, TechnoProfileID, PatinaID, InsetTypeID,
                            ColorID, TechnoColorID, InsetColorID, TechnoInsetTypeID, TechnoInsetColorID, Height, Width, Count, FrontConfigID, FrontsOrders.Notes FROM FrontsOrders
                            INNER JOIN MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID
                            INNER JOIN MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID
                            INNER JOIN infiniu2_marketingreference.dbo.Clients ON MegaOrders.ClientID=infiniu2_marketingreference.dbo.Clients.ClientID
                            INNER JOIN BatchDetails ON FrontsOrders.MainOrderID = BatchDetails.MainOrderID AND BatchDetails.FactoryID = 1
                            INNER JOIN Batch ON BatchDetails.BatchID = Batch.BatchID AND Batch.ProfilWorkAssignmentID = " + WorkAssignmentID +
                @" WHERE FrontsOrders.FactoryID=1 AND FrontID IN (" + string.Join(",", FrontsID.OfType<int>().ToArray()) + ")";

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(DT);
            }
            if (DT.Rows.Count > 0)
            {
                DataTable TempFrontsOrdersDT = DT.Clone();
                using (DataView DV = new DataView(DT))
                {
                    DV.Sort = "ClientName";
                    DistClientNamesDT = DV.ToTable(true, new string[] { "ClientName", "ClientID" });
                }

                for (int i = 0; i < DistClientNamesDT.Rows.Count; i++)
                {
                    ClientName = DistClientNamesDT.Rows[i]["ClientName"].ToString();

                    int RowIndex = 0;
                    HSSFSheet sheet1 = hssfworkbook.CreateSheet(ClientName.Replace("/", "-"));
                    sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

                    sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
                    sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
                    sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
                    sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

                    sheet1.SetColumnWidth(0, 25 * 256);
                    sheet1.SetColumnWidth(1, 11 * 256);
                    sheet1.SetColumnWidth(2, 25 * 256);
                    sheet1.SetColumnWidth(3, 15 * 256);
                    sheet1.SetColumnWidth(4, 6 * 256);
                    sheet1.SetColumnWidth(5, 6 * 256);
                    sheet1.SetColumnWidth(6, 6 * 256);

                    using (DataView DV = new DataView(DT, "ClientID=" + DistClientNamesDT.Rows[i]["ClientID"], "MainOrderID", DataViewRowState.CurrentRows))
                    {
                        DistMainOrdersDT = DV.ToTable(true, new string[] { "MainOrderID" });
                    }

                    for (int j = 0; j < DistMainOrdersDT.Rows.Count; j++)
                    {
                        MainOrderID = Convert.ToInt32(DistMainOrdersDT.Rows[j]["MainOrderID"]);
                        DataRow[] Frows = DT.Select("MainOrderID=" + MainOrderID);
                        if (Frows.Count() == 0)
                            continue;
                        OrderNumber = Convert.ToInt32(Frows[0]["OrderNumber"]);
                        Notes = Frows[0]["MNotes"].ToString();
                        OrderName = "№" + OrderNumber.ToString() + "-" + MainOrderID;

                        TempFrontsOrdersDT.Clear();
                        FrontsOrdersDT.Clear();
                        foreach (DataRow row in Frows)
                            TempFrontsOrdersDT.Rows.Add(row.ItemArray);
                        CollectMainOrders(TempFrontsOrdersDT, ref FrontsOrdersDT);

                        MainOrdersSummaryInfoToExcel(ref hssfworkbook, ref sheet1,
                            Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, FrontsOrdersDT,
                            WorkAssignmentID, DispatchDate, BatchName, ClientName, OrderName, Notes, ref RowIndex);
                        RowIndex++;
                    }
                }
            }

            DistMainOrdersDT.Clear();
            DistClientNamesDT.Clear();
            DT.Clear();

            //            SelectCommand = @"SELECT infiniu2_zovreference.dbo.Clients.ClientName, MainOrders.ClientID, MainOrders.DocNumber, MegaOrders.DispatchDate, MainOrders.Notes AS MNotes,
            //                FrontsOrdersID, FrontsOrders.MainOrderID, FrontID, FrontTypeID, InsetTypeID,
            //                ColorID, InsetColorID, AdditionalColorID, Height, Width, Count, FrontConfigID, FrontsOrders.Notes FROM FrontsOrders
            //                INNER JOIN MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID
            //                INNER JOIN MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID
            //                INNER JOIN infiniu2_zovreference.dbo.Clients ON MainOrders.ClientID=infiniu2_zovreference.dbo.Clients.ClientID
            //                INNER JOIN BatchDetails ON FrontsOrders.MainOrderID = BatchDetails.MainOrderID AND BatchDetails.FactoryID = 1
            //                INNER JOIN Batch ON BatchDetails.BatchID = Batch.BatchID AND Batch.ProfilWorkAssignmentID = " + WorkAssignmentID +
            //                @" WHERE FrontsOrders.FactoryID=1 AND FrontID IN (" +
            //                Convert.ToInt32(Fronts.Antalia) + "," + Convert.ToInt32(Fronts.Fat) + "," + Convert.ToInt32(Fronts.Leon) + "," + Convert.ToInt32(Fronts.LeonF) + "," +
            //                Convert.ToInt32(Fronts.Limog) + "," + Convert.ToInt32(Fronts.Luk) + "," + Convert.ToInt32(Fronts.LukPVH) + "," +
            //                Convert.ToInt32(Fronts.Milano) + "," + Convert.ToInt32(Fronts.MilanoK) + "," + Convert.ToInt32(Fronts.MilanoKF) + "," +
            //                Convert.ToInt32(Fronts.Praga) + "," + Convert.ToInt32(Fronts.Sigma) + "," + Convert.ToInt32(Fronts.Venecia) + "," + Convert.ToInt32(Fronts.VeneciaF) + ")";

            SelectCommand = @"SELECT infiniu2_zovreference.dbo.Clients.ClientName, MainOrders.ClientID, MainOrders.DocNumber, MegaOrders.DispatchDate, MainOrders.Notes AS MNotes,
                FrontsOrdersID, FrontsOrders.MainOrderID, FrontID, TechnoProfileID, PatinaID, InsetTypeID,
                ColorID, TechnoColorID, InsetColorID, TechnoInsetTypeID, TechnoInsetColorID, Height, Width, Count, FrontConfigID, FrontsOrders.Notes FROM FrontsOrders
                INNER JOIN MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID
                INNER JOIN MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID
                INNER JOIN infiniu2_zovreference.dbo.Clients ON MainOrders.ClientID=infiniu2_zovreference.dbo.Clients.ClientID
                INNER JOIN BatchDetails ON FrontsOrders.MainOrderID = BatchDetails.MainOrderID AND BatchDetails.FactoryID = 1
                INNER JOIN Batch ON BatchDetails.BatchID = Batch.BatchID AND Batch.ProfilWorkAssignmentID = " + WorkAssignmentID +
                @" WHERE FrontsOrders.FactoryID=1 AND FrontID IN (" + string.Join(",", FrontsID.OfType<int>().ToArray()) + ")";

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.ZOVOrdersConnectionString))
            {
                DA.Fill(DT);
            }
            if (DT.Rows.Count > 0)
            {
                DataTable TempFrontsOrdersDT = DT.Clone();
                using (DataView DV = new DataView(DT))
                {
                    DV.Sort = "ClientName";
                    DistClientNamesDT = DV.ToTable(true, new string[] { "ClientName", "ClientID" });
                }

                for (int i = 0; i < DistClientNamesDT.Rows.Count; i++)
                {
                    ClientName = DistClientNamesDT.Rows[i]["ClientName"].ToString();

                    int RowIndex = 0;
                    HSSFSheet sheet1 = hssfworkbook.CreateSheet(ClientName.Replace("/", "-"));
                    sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;
                    sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
                    sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
                    sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
                    sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

                    sheet1.SetColumnWidth(0, 25 * 256);
                    sheet1.SetColumnWidth(1, 11 * 256);
                    sheet1.SetColumnWidth(2, 25 * 256);
                    sheet1.SetColumnWidth(3, 15 * 256);
                    sheet1.SetColumnWidth(4, 6 * 256);
                    sheet1.SetColumnWidth(5, 6 * 256);
                    sheet1.SetColumnWidth(6, 6 * 256);

                    using (DataView DV = new DataView(DT, "ClientID=" + DistClientNamesDT.Rows[i]["ClientID"], "MainOrderID", DataViewRowState.CurrentRows))
                    {
                        DistMainOrdersDT = DV.ToTable(true, new string[] { "MainOrderID" });
                    }

                    for (int j = 0; j < DistMainOrdersDT.Rows.Count; j++)
                    {
                        MainOrderID = Convert.ToInt32(DistMainOrdersDT.Rows[j]["MainOrderID"]);
                        DataRow[] Frows = DT.Select("MainOrderID=" + MainOrderID);
                        if (Frows.Count() == 0)
                            continue;
                        if (Frows[0]["DispatchDate"] != DBNull.Value)
                            DispatchDate = Convert.ToDateTime(Frows[0]["DispatchDate"]).ToString("dd.MM.yyyy");
                        Notes = Frows[0]["MNotes"].ToString();
                        OrderName = Frows[0]["DocNumber"].ToString();

                        TempFrontsOrdersDT.Clear();
                        FrontsOrdersDT.Clear();
                        foreach (DataRow row in Frows)
                            TempFrontsOrdersDT.Rows.Add(row.ItemArray);
                        CollectMainOrders(TempFrontsOrdersDT, ref FrontsOrdersDT);

                        MainOrdersSummaryInfoToExcel(ref hssfworkbook, ref sheet1,
                            Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, FrontsOrdersDT,
                            WorkAssignmentID, DispatchDate, BatchName, ClientName, OrderName, Notes, ref RowIndex);
                        RowIndex++;
                    }
                }
            }
        }

        public string GetMarketClientName(int MainOrderID)
        {
            string ClientName = string.Empty;

            using (DataTable DT = new DataTable())
            {
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT ClientName FROM Clients WHERE ClientID = " +
                    " (SELECT ClientID FROM infiniu2_marketingorders.dbo.MegaOrders" +
                    " WHERE MegaOrderID=(SELECT TOP 1 MegaOrderID FROM infiniu2_marketingorders.dbo.MainOrders WHERE MainOrderID = " + MainOrderID + "))",
                    ConnectionStrings.MarketingReferenceConnectionString))
                {
                    DA.Fill(DT);
                    if (DT.Rows.Count > 0)
                        ClientName = DT.Rows[0]["ClientName"].ToString();
                }
            }

            return ClientName;
        }

        public bool GetOrders(int WorkAssignmentID)
        {
            for (int i = 0; i < FrontsID.Count; i++)
            {
                if (Convert.ToInt32(Convert.ToInt32(FrontsID[i])) == Convert.ToInt32(Fronts.TechnoN))
                    GetFrontsOrders(ref TechnoNOrdersDT, WorkAssignmentID, Fronts.TechnoN);
                if (Convert.ToInt32(Convert.ToInt32(FrontsID[i])) == Convert.ToInt32(Fronts.Antalia))
                    GetFrontsOrders(ref AntaliaOrdersDT, WorkAssignmentID, Fronts.Antalia);
                if (Convert.ToInt32(Convert.ToInt32(FrontsID[i])) == Convert.ToInt32(Fronts.Nord95))
                    GetFrontsOrders(ref Nord95OrdersDT, WorkAssignmentID, Fronts.Nord95);
                if (Convert.ToInt32(Convert.ToInt32(FrontsID[i])) == Convert.ToInt32(Fronts.epFox))
                    GetFrontsOrders(ref epFoxOrdersDT, WorkAssignmentID, Fronts.epFox);
                if (Convert.ToInt32(Convert.ToInt32(FrontsID[i])) == Convert.ToInt32(Fronts.Venecia))
                    GetFrontsOrders(ref VeneciaOrdersDT, WorkAssignmentID, Fronts.Venecia);
                if (Convert.ToInt32(Convert.ToInt32(FrontsID[i])) == Convert.ToInt32(Fronts.Bergamo))
                    GetFrontsOrders(ref BergamoOrdersDT, WorkAssignmentID, Fronts.Bergamo);
                if (Convert.ToInt32(Convert.ToInt32(FrontsID[i])) == Convert.ToInt32(Fronts.ep041))
                    GetFrontsOrders(ref ep041OrdersDT, WorkAssignmentID, Fronts.ep041);
                if (Convert.ToInt32(Convert.ToInt32(FrontsID[i])) == Convert.ToInt32(Fronts.ep071))
                    GetFrontsOrders(ref ep071OrdersDT, WorkAssignmentID, Fronts.ep071);
                if (Convert.ToInt32(Convert.ToInt32(FrontsID[i])) == Convert.ToInt32(Fronts.ep206))
                    GetFrontsOrders(ref ep206OrdersDT, WorkAssignmentID, Fronts.ep206);
                if (Convert.ToInt32(Convert.ToInt32(FrontsID[i])) == Convert.ToInt32(Fronts.ep216))
                    GetFrontsOrders(ref ep216OrdersDT, WorkAssignmentID, Fronts.ep216);
                if (Convert.ToInt32(Convert.ToInt32(FrontsID[i])) == Convert.ToInt32(Fronts.ep111))
                    GetFrontsOrders(ref ep111OrdersDT, WorkAssignmentID, Fronts.ep111);
                if (Convert.ToInt32(Convert.ToInt32(FrontsID[i])) == Convert.ToInt32(Fronts.Boston))
                    GetFrontsOrders(ref BostonOrdersDT, WorkAssignmentID, Fronts.Boston);
                if (Convert.ToInt32(FrontsID[i]) == Convert.ToInt32(Fronts.Leon))
                    GetFrontsOrders(ref LeonOrdersDT, WorkAssignmentID, Fronts.Leon);
                if (Convert.ToInt32(FrontsID[i]) == Convert.ToInt32(Fronts.Limog))
                    GetFrontsOrders(ref LimogOrdersDT, WorkAssignmentID, Fronts.Limog);
                if (Convert.ToInt32(FrontsID[i]) == Convert.ToInt32(Fronts.ep066Marsel4))
                    GetFrontsOrders(ref ep066Marsel4OrdersDT, WorkAssignmentID, Fronts.ep066Marsel4);
                if (Convert.ToInt32(FrontsID[i]) == Convert.ToInt32(Fronts.ep110Jersy))
                    GetFrontsOrders(ref ep110JersyOrdersDT, WorkAssignmentID, Fronts.ep110Jersy);
                if (Convert.ToInt32(FrontsID[i]) == Convert.ToInt32(Fronts.ep018Marsel1))
                    GetFrontsOrders(ref ep018Marsel1OrdersDT, WorkAssignmentID, Fronts.ep018Marsel1);
                if (Convert.ToInt32(FrontsID[i]) == Convert.ToInt32(Fronts.ep043Shervud))
                    GetFrontsOrders(ref ep043ShervudOrdersDT, WorkAssignmentID, Fronts.ep043Shervud);
                if (Convert.ToInt32(FrontsID[i]) == Convert.ToInt32(Fronts.ep112))
                    GetFrontsOrders(ref ep112OrdersDT, WorkAssignmentID, Fronts.ep112);
                if (Convert.ToInt32(FrontsID[i]) == Convert.ToInt32(Fronts.Urban))
                    GetFrontsOrders(ref UrbanOrdersDT, WorkAssignmentID, Fronts.Urban);
                if (Convert.ToInt32(FrontsID[i]) == Convert.ToInt32(Fronts.Alby))
                    GetFrontsOrders(ref AlbyOrdersDT, WorkAssignmentID, Fronts.Alby);
                if (Convert.ToInt32(FrontsID[i]) == Convert.ToInt32(Fronts.Bruno))
                    GetFrontsOrders(ref BrunoOrdersDT, WorkAssignmentID, Fronts.Bruno);
                if (Convert.ToInt32(FrontsID[i]) == Convert.ToInt32(Fronts.epsh406Techno4))
                    GetFrontsOrders(ref epsh406Techno4OrdersDT, WorkAssignmentID, Fronts.epsh406Techno4);
                if (Convert.ToInt32(FrontsID[i]) == Convert.ToInt32(Fronts.Luk))
                    GetFrontsOrders(ref LukOrdersDT, WorkAssignmentID, Fronts.Luk);
                if (Convert.ToInt32(FrontsID[i]) == Convert.ToInt32(Fronts.LukPVH))
                    GetFrontsOrders(ref LukPVHOrdersDT, WorkAssignmentID, Fronts.LukPVH);
                if (Convert.ToInt32(FrontsID[i]) == Convert.ToInt32(Fronts.Milano))
                    GetFrontsOrders(ref MilanoOrdersDT, WorkAssignmentID, Fronts.Milano);
                if (Convert.ToInt32(FrontsID[i]) == Convert.ToInt32(Fronts.Praga))
                    GetFrontsOrders(ref PragaOrdersDT, WorkAssignmentID, Fronts.Praga);
                if (Convert.ToInt32(FrontsID[i]) == Convert.ToInt32(Fronts.Sigma))
                    GetFrontsOrders(ref SigmaOrdersDT, WorkAssignmentID, Fronts.Sigma);
                if (Convert.ToInt32(FrontsID[i]) == Convert.ToInt32(Fronts.Fat))
                    GetFrontsOrders(ref FatOrdersDT, WorkAssignmentID, Fronts.Fat);
            }

            ProfileNamesDT.Clear();
            GetProfileNames(ref ProfileNamesDT, WorkAssignmentID);

            GetNotArchDecorOrders(ref NotArchDecorOrdersDT, WorkAssignmentID, 1);
            GetBagetWithAngleOrders(ref BagetWithAngelOrdersDT, WorkAssignmentID, 1);
            GetArchDecorOrders(ref ArchDecorOrdersDT, WorkAssignmentID, 1);
            GetGridsDecorOrders(ref GridsDecorOrdersDT, WorkAssignmentID, 1);

            if (TechnoNOrdersDT.Rows.Count == 0 && AntaliaOrdersDT.Rows.Count == 0 &&
                Nord95OrdersDT.Rows.Count == 0 && epFoxOrdersDT.Rows.Count == 0 && VeneciaOrdersDT.Rows.Count == 0 && BergamoOrdersDT.Rows.Count == 0 &&
                ep041OrdersDT.Rows.Count == 0 && ep071OrdersDT.Rows.Count == 0 && ep206OrdersDT.Rows.Count == 0 && ep216OrdersDT.Rows.Count == 0 && ep111OrdersDT.Rows.Count == 0 &&
                BostonOrdersDT.Rows.Count == 0 && LeonOrdersDT.Rows.Count == 0 &&
                LimogOrdersDT.Rows.Count == 0 && LukOrdersDT.Rows.Count == 0 && LukPVHOrdersDT.Rows.Count == 0 &&
                ep066Marsel4OrdersDT.Rows.Count == 0 && ep110JersyOrdersDT.Rows.Count == 0 &&
                ep018Marsel1OrdersDT.Rows.Count == 0 && ep043ShervudOrdersDT.Rows.Count == 0 && ep112OrdersDT.Rows.Count == 0 &&
                UrbanOrdersDT.Rows.Count == 0 && AlbyOrdersDT.Rows.Count == 0 && BrunoOrdersDT.Rows.Count == 0 && epsh406Techno4OrdersDT.Rows.Count == 0 &&
                MilanoOrdersDT.Rows.Count == 0 && PragaOrdersDT.Rows.Count == 0 && SigmaOrdersDT.Rows.Count == 0 &&
                FatOrdersDT.Rows.Count == 0 &&
                BagetWithAngelOrdersDT.Rows.Count == 0 && NotArchDecorOrdersDT.Rows.Count == 0 && ArchDecorOrdersDT.Rows.Count == 0 && GridsDecorOrdersDT.Rows.Count == 0)
                return false;
            else
                return true;
        }

        public string GetPatinaName(int PatinaID)
        {
            string FrontType = string.Empty;
            try
            {
                DataRow[] Rows = PatinaDataTable.Select("PatinaID = " + PatinaID);
                FrontType = Rows[0]["PatinaName"].ToString();
            }
            catch
            {
                return string.Empty;
            }
            return FrontType;
        }

        public string GetZOVClientName(int MainOrderID)
        {
            string ClientName = string.Empty;

            using (DataTable DT = new DataTable())
            {
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT ClientName FROM Clients WHERE ClientID = " +
                    " (SELECT ClientID FROM infiniu2_zovorders.dbo.MainOrders WHERE MainOrderID = " + MainOrderID + ")",
                    ConnectionStrings.ZOVReferenceConnectionString))
                {
                    DA.Fill(DT);
                    if (DT.Rows.Count > 0)
                        ClientName = DT.Rows[0]["ClientName"].ToString();
                }
            }

            return ClientName;
        }

        public void GridsDecorAssemblyToExcelSingly(ref HSSFWorkbook hssfworkbook, ref HSSFSheet sheet1,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            DataTable DT, int WorkAssignmentID, string BatchName, string ClientName, string OperationName, string OrderName, string Notes, ref int RowIndex)
        {
            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
            cell.CellStyle = Calibri11CS;

            RowIndex++;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, ClientName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Заказ:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, OrderName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 4, OperationName);
            cell.CellStyle = CalibriBold11CS;
            if (Notes.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Примечание: " + Notes);
                cell.CellStyle = CalibriBold11CS;
            }

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Наименование");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Цвет");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Длин./Выс.");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 5, "Прим.");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            int DecorID = -1;
            int AllTotalAmount = 0;
            int TotalAmount = 0;
            int DifferentDecorCount = 0;

            using (DataView DV = new DataView(DT))
            {
                DifferentDecorCount = DV.ToTable(true, new string[] { "DecorID" }).Rows.Count;
            }

            if (DT.Rows.Count > 0)
            {
                DecorID = Convert.ToInt32(DT.Rows[0]["DecorID"]);
            }

            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value)
                {
                    AllTotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    TotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                }
                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "DecorID")
                        continue;

                    Type t = DT.Rows[x][y].GetType();

                    if (t.Name == "Decimal")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToDouble(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderDecCS;
                        continue;
                    }
                    if (t.Name == "Int32")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToInt32(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (t.Name == "String" || t.Name == "DBNull")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(DT.Rows[x][y].ToString());
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }
                }

                if (x + 1 <= DT.Rows.Count - 1)
                {
                    if (DecorID != Convert.ToInt32(DT.Rows[x + 1]["DecorID"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "DecorID")
                                continue;

                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        DecorID = Convert.ToInt32(DT.Rows[x + 1]["DecorID"]);
                        TotalAmount = 0;
                        RowIndex++;
                    }
                }

                if (x == DT.Rows.Count - 1)
                {
                    if (DifferentDecorCount > 1)
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "DecorID")
                                continue;

                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;
                    }
                    RowIndex++;

                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "DecorID")
                            continue;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;
                }
                RowIndex++;
            }
        }

        public void GridsToExcelSingly(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            ref HSSFSheet sheet1, DataTable DT, int WorkAssignmentID, string DispatchDate, string BatchName, string ClientName, string TableName, ref int RowIndex)
        {
            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Плановое время выполнения:");
            cell.CellStyle = Calibri11CS;

            if (DispatchDate.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Отгрузка: " + DispatchDate);
                cell.CellStyle = CalibriBold11CS;
            }

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, TableName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, ClientName);
            cell.CellStyle = CalibriBold11CS;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Цвет наполнителя");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Работник");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            decimal TotalSquare = 0;
            decimal AllTotalSquare = 0;
            string str = string.Empty;

            int CType = -1;
            int DifferentDecorCount = 0;

            using (DataView DV = new DataView(DT))
            {
                DifferentDecorCount = DV.ToTable(true, new string[] { "InsetColorID" }).Rows.Count;
            }

            if (DT.Rows.Count > 0)
            {
                CType = Convert.ToInt32(DT.Rows[0]["InsetColorID"]);
            }
            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value)
                {
                    TotalSquare += Convert.ToDecimal(DT.Rows[x]["Height"]) * Convert.ToDecimal(DT.Rows[x]["Width"]) * Convert.ToDecimal(DT.Rows[x]["Count"]) / 1000000;
                    AllTotalSquare += Convert.ToDecimal(DT.Rows[x]["Height"]) * Convert.ToDecimal(DT.Rows[x]["Width"]) * Convert.ToDecimal(DT.Rows[x]["Count"]) / 1000000;
                }

                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "InsetColorID" || DT.Columns[y].ColumnName == "TechnoInsetColorID" ||
                        DT.Columns[y].ColumnName == "GlassCount" || DT.Columns[y].ColumnName == "MegaCount")
                        continue;
                    Type t = DT.Rows[x][y].GetType();

                    if (DT.Columns[y].ColumnName == "Name")
                        str = DT.Rows[x][y].ToString();

                    if (t.Name == "Decimal")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToDouble(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderDecCS;
                        continue;
                    }
                    if (t.Name == "Int32")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToInt32(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (t.Name == "String" || t.Name == "DBNull")
                    {
                        if (DT.Rows[x][y].ToString().IndexOf("3х4") != -1)
                        {
                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(DT.Rows[x][y].ToString());
                            cell.CellStyle = CalibriBold11CS;
                        }
                        else
                        {
                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(DT.Rows[x][y].ToString());
                            cell.CellStyle = TableHeaderCS;
                        }
                        continue;
                    }
                }

                if (x + 1 <= DT.Rows.Count - 1)
                {
                    if (CType != Convert.ToInt32(DT.Rows[x + 1]["InsetColorID"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "InsetColorID" || DT.Columns[y].ColumnName == "TechnoInsetColorID" ||
                                DT.Columns[y].ColumnName == "GlassCount" || DT.Columns[y].ColumnName == "MegaCount")
                                continue;
                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        TotalSquare = decimal.Round(TotalSquare, 2, MidpointRounding.AwayFromZero);

                        cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                        cell.SetCellValue(Convert.ToDouble(TotalSquare));
                        cell.CellStyle = TableHeaderDecCS;

                        CType = Convert.ToInt32(DT.Rows[x + 1]["InsetColorID"]);
                        TotalSquare = 0;

                        RowIndex++;
                    }
                }

                if (x == DT.Rows.Count - 1)
                {
                    if (DifferentDecorCount > 1)
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "InsetColorID" || DT.Columns[y].ColumnName == "TechnoInsetColorID" ||
                                DT.Columns[y].ColumnName == "GlassCount" || DT.Columns[y].ColumnName == "MegaCount")
                                continue;
                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        TotalSquare = decimal.Round(TotalSquare, 2, MidpointRounding.AwayFromZero);

                        cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                        cell.SetCellValue(Convert.ToDouble(TotalSquare));
                        cell.CellStyle = TableHeaderDecCS;
                    }
                    RowIndex++;

                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "InsetColorID" || DT.Columns[y].ColumnName == "TechnoInsetColorID" ||
                            DT.Columns[y].ColumnName == "GlassCount" || DT.Columns[y].ColumnName == "MegaCount")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    AllTotalSquare = decimal.Round(AllTotalSquare, 2, MidpointRounding.AwayFromZero);

                    cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                    cell.SetCellValue(Convert.ToDouble(AllTotalSquare));
                    cell.CellStyle = TableHeaderDecCS;
                }
                RowIndex++;
            }
        }

        public bool HasParameter(int ProductID, string Parameter)
        {
            DataRow[] Rows = DecorParametersDT.Select("ProductID = " + ProductID);

            return Convert.ToBoolean(Rows[0][Parameter]);
        }

        public void Initialize()
        {
            Create();
            Fill();
        }

        public void InsetToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            int WorkAssignmentID, string DispatchDate, string BatchName, string ClientName)
        {
            int RowIndex = 0;

            HSSFSheet sheet1 = hssfworkbook.CreateSheet("Вставка");
            sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

            sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
            sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
            sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
            sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

            sheet1.SetColumnWidth(0, 25 * 256);
            sheet1.SetColumnWidth(1, 11 * 256);
            sheet1.SetColumnWidth(2, 11 * 256);
            sheet1.SetColumnWidth(3, 7 * 256);
            sheet1.SetColumnWidth(4, 12 * 256);

            {
                HSSFPatriarch patriarch = sheet1.CreateDrawingPatriarch();
                int ind = 0;

                CreateBarcode(WorkAssignmentID);
                HSSFClientAnchor anchor = new HSSFClientAnchor(0, 1, 1000, 8, 5, ind, 6, ind + 3)
                {
                    AnchorType = 2
                };
                string BarcodeNumber = GetBarcodeNumber(24, WorkAssignmentID);
                string barcodeName = System.Environment.GetEnvironmentVariable("TEMP") + @"/" + BarcodeNumber + ".jpg";
                HSSFPicture picture = patriarch.CreatePicture(anchor, LoadImage(barcodeName, hssfworkbook));

                HSSFFont BarcodeFont = hssfworkbook.CreateFont();
                BarcodeFont.FontHeightInPoints = 12;
                BarcodeFont.FontName = "Calibri";

                HSSFCellStyle BarcodeStyle = hssfworkbook.CreateCellStyle();
                BarcodeStyle.Alignment = HSSFCellStyle.ALIGN_CENTER;
                BarcodeStyle.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
                BarcodeStyle.SetFont(BarcodeFont);

                HSSFCell barcodeCell = sheet1.CreateRow(ind + 3).CreateCell(5);
                barcodeCell.SetCellValue(BarcodeNumber.ToString());
                barcodeCell.CellStyle = BarcodeStyle;
                sheet1.AddMergedRegion(new NPOI.HSSF.Util.Region(ind + 3, 5, ind + 3, 6));
            }

            InsetDT.Clear();
            CollectAllInsets(ref InsetDT);

            DataTable DT = InsetDT.Copy();
            DataColumn Col1 = new DataColumn();
            DataColumn Col2 = new DataColumn();

            Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
            Col1.SetOrdinal(4);

            if (DT.Rows.Count > 0)
            {
                AllInsetsToExcelSingly(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT, WorkAssignmentID, DispatchDate, BatchName, ClientName, "Вставка", ref RowIndex);
                RowIndex++;
                RowIndex++;
            }
            InsetDT.Clear();
            CollectInsetsGridsOnly(ref InsetDT);

            DT.Dispose();
            Col1.Dispose();
            DT = InsetDT.Copy();
            Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
            Col1.SetOrdinal(4);

            if (DT.Rows.Count > 0)
            {
                GridsToExcelSingly(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT, WorkAssignmentID, DispatchDate, BatchName, ClientName, "Вставка", ref RowIndex);
                RowIndex++;
                RowIndex++;
            }

            InsetDT.Clear();
            CollectFilenkaOnly(ref InsetDT);

            DT.Dispose();
            Col1.Dispose();
            DT = InsetDT.Copy();
            Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
            Col1.SetOrdinal(4);

            if (DT.Rows.Count > 0)
            {
                FilenkaToExcelSingly(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT, WorkAssignmentID, DispatchDate, BatchName, ClientName, "Вставка", ref RowIndex);
                RowIndex++;
                RowIndex++;
            }

            InsetDT.Clear();
            CollectPressOnly(ref InsetDT);

            DT.Dispose();
            Col1.Dispose();
            DT = InsetDT.Copy();
            Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
            Col1.SetOrdinal(4);

            if (DT.Rows.Count > 0)
            {
                PressToExcelSingly(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT, WorkAssignmentID, DispatchDate, BatchName, ClientName, "Вставка", ref RowIndex);
                RowIndex++;
                RowIndex++;
            }

            RowIndex++;
        }

        public void MainOrdersSummaryInfoToExcel(ref HSSFWorkbook hssfworkbook, ref HSSFSheet sheet1,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            DataTable DT, int WorkAssignmentID, string DispatchDate, string BatchName, string ClientName, string OrderName, string Notes, ref int RowIndex)
        {
            HSSFCell cell = null;

            if (DispatchDate.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Отгрузка: " + DispatchDate);
                cell.CellStyle = CalibriBold11CS;
            }

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, "Заказы");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, ClientName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Заказ:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, OrderName);
            cell.CellStyle = CalibriBold11CS;
            if (Notes.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Примечание: " + Notes);
                cell.CellStyle = CalibriBold11CS;
            }

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Профиль");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Цвет профиля");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Цвет наполнителя");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Техно цвет");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 5, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 6, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 7, "Квадратура");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            decimal TotalSquare = 0;
            int TotalAmount = 0;

            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value)
                {
                    TotalSquare += Convert.ToDecimal(DT.Rows[x]["Square"]);
                    TotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                }

                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    Type t = DT.Rows[x][y].GetType();

                    if (t.Name == "Decimal")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToDouble(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderDecCS;
                        continue;
                    }
                    if (t.Name == "Int32")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToInt32(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (t.Name == "String" || t.Name == "DBNull")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(DT.Rows[x][y].ToString());
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }
                }

                if (x == DT.Rows.Count - 1)
                {
                    RowIndex++;
                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Итого:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(6);
                    cell.SetCellValue(TotalAmount);
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(7);
                    cell.SetCellValue(Convert.ToDouble(TotalSquare));
                    cell.CellStyle = TableHeaderCS;
                }
                RowIndex++;
            }
        }

        public void MartinToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            int WorkAssignmentID, string DispatchDate, string BatchName, string ClientName)
        {
            int RowIndex = 0;

            HSSFSheet sheet1 = hssfworkbook.CreateSheet("MARTIN");
            sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

            sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
            sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
            sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
            sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

            sheet1.SetColumnWidth(0, 25 * 256);
            sheet1.SetColumnWidth(1, 11 * 256);
            sheet1.SetColumnWidth(2, 6 * 256);
            sheet1.SetColumnWidth(3, 6 * 256);
            sheet1.SetColumnWidth(4, 16 * 256);

            {
                HSSFPatriarch patriarch = sheet1.CreateDrawingPatriarch();
                int ind = 0;

                CreateBarcode(WorkAssignmentID);
                HSSFClientAnchor anchor = new HSSFClientAnchor(0, 1, 1000, 8, 5, ind, 6, ind + 3)
                {
                    AnchorType = 2
                };
                string BarcodeNumber = GetBarcodeNumber(24, WorkAssignmentID);
                string barcodeName = System.Environment.GetEnvironmentVariable("TEMP") + @"/" + BarcodeNumber + ".jpg";
                HSSFPicture picture = patriarch.CreatePicture(anchor, LoadImage(barcodeName, hssfworkbook));

                HSSFFont BarcodeFont = hssfworkbook.CreateFont();
                BarcodeFont.FontHeightInPoints = 12;
                BarcodeFont.FontName = "Calibri";

                HSSFCellStyle BarcodeStyle = hssfworkbook.CreateCellStyle();
                BarcodeStyle.Alignment = HSSFCellStyle.ALIGN_CENTER;
                BarcodeStyle.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
                BarcodeStyle.SetFont(BarcodeFont);

                HSSFCell barcodeCell = sheet1.CreateRow(ind + 3).CreateCell(5);
                barcodeCell.SetCellValue(BarcodeNumber.ToString());
                barcodeCell.CellStyle = BarcodeStyle;
                sheet1.AddMergedRegion(new NPOI.HSSF.Util.Region(ind + 3, 5, ind + 3, 6));
            }

            RapidDT.Clear();
            CombineRapidSimple(ref RapidDT, false);
            CombineRapidFilenka(ref RapidDT);
            //CombineRapidBoxes(ref RapidDT);
            CombineRapidFilenkaBoxes(ref RapidDT);

            DataTable DT = RapidDT.Copy();

            if (DT.Rows.Count > 0)
            {
                RapidToExcelSingly(ref hssfworkbook, ref sheet1,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DT, WorkAssignmentID, "MARTIN", DispatchDate, BatchName, ClientName, ref RowIndex);
                RowIndex++;
                RowIndex++;
            }

            RapidDT.Clear();
            CombineRapidImpostSimple(ref RapidDT, false);

            DT = RapidDT.Copy();

            if (DT.Rows.Count > 0)
            {
                RapidToExcelSingly(ref hssfworkbook, ref sheet1,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DT, WorkAssignmentID, "MARTIN", DispatchDate, BatchName, ClientName, ref RowIndex);
                RowIndex++;
                RowIndex++;
            }
        }

        public void MartinToExcelSingly(ref HSSFWorkbook hssfworkbook, ref HSSFSheet sheet1,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            DataTable DT, int WorkAssignmentID, string DispatchDate, string BatchName, string ClientName, ref int RowIndex)
        {
            HSSFCell cell = null;

            if (DispatchDate.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Отгрузка: " + DispatchDate);
                cell.CellStyle = CalibriBold11CS;
            }
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, "РапидТ");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, ClientName);
            cell.CellStyle = CalibriBold11CS;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Профиль");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Цвет профиля");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Работник");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            decimal SticksCount = 0;
            int CType = 0;
            int FType = 0;
            int PType = 0;
            int TotalAmount = 0;
            int AllTotalAmount = 0;
            int Count = 0;
            int Height = 0;

            if (DT.Rows.Count > 0)
            {
                CType = Convert.ToInt32(DT.Rows[0]["ColorType"]);
                FType = Convert.ToInt32(DT.Rows[0]["FrontType"]);
                PType = Convert.ToInt32(DT.Rows[0]["ProfileType"]);
            }

            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value && DT.Rows[x]["Height"] != DBNull.Value)
                {
                    Count = Convert.ToInt32(DT.Rows[x]["Count"]);
                    Height = Convert.ToInt32(DT.Rows[x]["Height"]);
                    SticksCount += (Height + 4) * Count;
                    TotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    AllTotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    Count = Convert.ToInt32(DT.Rows[x]["Count"]);
                }

                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "ProfileType" || DT.Columns[y].ColumnName == "ColorType"
                        || DT.Columns[y].ColumnName == "FrontType" || DT.Columns[y].ColumnName == "VitrinaCount" || DT.Columns[y].ColumnName == "BoxCount")
                        continue;
                    Type t = DT.Rows[x][y].GetType();

                    if (t.Name == "Decimal")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToDouble(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderDecCS;
                        continue;
                    }
                    if (t.Name == "Int32")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToInt32(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (t.Name == "String" || t.Name == "DBNull")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(DT.Rows[x][y].ToString());
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }
                }

                if (x + 1 <= DT.Rows.Count - 1 &&
                    (FType != Convert.ToInt32(DT.Rows[x + 1]["FrontType"]) || PType != Convert.ToInt32(DT.Rows[x + 1]["ProfileType"]) || CType != Convert.ToInt32(DT.Rows[x + 1]["ColorType"])))
                {
                    RowIndex++;
                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "ProfileType" || DT.Columns[y].ColumnName == "ColorType"
                            || DT.Columns[y].ColumnName == "FrontType" || DT.Columns[y].ColumnName == "VitrinaCount" || DT.Columns[y].ColumnName == "BoxCount")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    SticksCount = SticksCount * 1.15m / 2620;
                    SticksCount = decimal.Round(SticksCount, 1, MidpointRounding.AwayFromZero);
                    cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                    cell.SetCellValue(SticksCount + " палок");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Итого:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                    cell.SetCellValue(TotalAmount);
                    cell.CellStyle = TableHeaderCS;
                    RowIndex++;

                    CType = Convert.ToInt32(DT.Rows[x + 1]["ColorType"]);
                    PType = Convert.ToInt32(DT.Rows[x + 1]["ProfileType"]);
                    Count = 0;
                    Height = 0;
                    SticksCount = 0;
                    TotalAmount = 0;
                }

                if (x == DT.Rows.Count - 1)
                {
                    RowIndex++;
                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "ProfileType" || DT.Columns[y].ColumnName == "ColorType"
                            || DT.Columns[y].ColumnName == "FrontType" || DT.Columns[y].ColumnName == "VitrinaCount" || DT.Columns[y].ColumnName == "BoxCount")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    SticksCount = SticksCount * 1.15m / 2620;
                    SticksCount = decimal.Round(SticksCount, 1, MidpointRounding.AwayFromZero);
                    cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                    cell.SetCellValue(SticksCount + " палок");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Итого:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                    cell.SetCellValue(TotalAmount);
                    cell.CellStyle = TableHeaderCS;

                    RowIndex++;
                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "ProfileType" || DT.Columns[y].ColumnName == "ColorType"
                            || DT.Columns[y].ColumnName == "FrontType" || DT.Columns[y].ColumnName == "VitrinaCount" || DT.Columns[y].ColumnName == "BoxCount")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;
                }
                RowIndex++;
            }
        }

        public void NotArchDecorAssemblyToExcelSingly(ref HSSFWorkbook hssfworkbook, ref HSSFSheet sheet1,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            DataTable DT, int WorkAssignmentID, string BatchName, string ClientName, string OperationName, string OrderName, string Notes, ref int RowIndex)
        {
            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
            cell.CellStyle = Calibri11CS;

            RowIndex++;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, ClientName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Заказ:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, OrderName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 4, OperationName);
            cell.CellStyle = CalibriBold11CS;
            if (Notes.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Примечание: " + Notes);
                cell.CellStyle = CalibriBold11CS;
            }

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Наименование");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Цвет");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Длин./Выс.");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 5, "Прим.");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            int DecorID = -1;
            int AllTotalAmount = 0;
            int TotalAmount = 0;
            int DifferentDecorCount = 0;

            using (DataView DV = new DataView(DT))
            {
                DifferentDecorCount = DV.ToTable(true, new string[] { "DecorID" }).Rows.Count;
            }
            if (DT.Rows.Count > 0)
            {
                DecorID = Convert.ToInt32(DT.Rows[0]["DecorID"]);
            }

            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value)
                {
                    AllTotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    TotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                }
                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "DecorID")
                        continue;

                    Type t = DT.Rows[x][y].GetType();

                    if (t.Name == "Decimal")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToDouble(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderDecCS;
                        continue;
                    }
                    if (t.Name == "Int32")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToInt32(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (t.Name == "String" || t.Name == "DBNull")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(DT.Rows[x][y].ToString());
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }
                }

                if (x + 1 <= DT.Rows.Count - 1)
                {
                    if (DecorID != Convert.ToInt32(DT.Rows[x + 1]["DecorID"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "DecorID")
                                continue;

                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        DecorID = Convert.ToInt32(DT.Rows[x + 1]["DecorID"]);
                        TotalAmount = 0;
                        RowIndex++;
                    }
                }

                if (x == DT.Rows.Count - 1)
                {
                    if (DifferentDecorCount > 1)
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "DecorID")
                                continue;

                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;
                    }
                    RowIndex++;

                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "DecorID")
                            continue;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;
                }
                RowIndex++;
            }
        }

        public DataTable OrderedFrameColors(DataTable SourceDT)
        {
            DataTable OrderedDT = SourceDT.Copy();
            OrderedDT.Columns.Add(new DataColumn("ColorName", Type.GetType("System.String")));

            for (int i = 0; i < OrderedDT.Rows.Count; i++)
                OrderedDT.Rows[i]["ColorName"] = GetColorName(Convert.ToInt32(OrderedDT.Rows[i]["ColorID"]));

            using (DataView DV = new DataView(OrderedDT.Copy()))
            {
                DV.Sort = "ColorName";
                OrderedDT.Clear();
                OrderedDT = DV.ToTable();
            }

            return OrderedDT;
        }

        public DataTable OrderedTechnoFrameColors(DataTable SourceDT)
        {
            DataTable OrderedDT = SourceDT.Copy();
            OrderedDT.Columns.Add(new DataColumn("ColorName", Type.GetType("System.String")));

            for (int i = 0; i < OrderedDT.Rows.Count; i++)
                OrderedDT.Rows[i]["ColorName"] = GetColorName(Convert.ToInt32(OrderedDT.Rows[i]["TechnoColorID"]));

            using (DataView DV = new DataView(OrderedDT.Copy()))
            {
                DV.Sort = "ColorName";
                OrderedDT.Clear();
                OrderedDT = DV.ToTable();
            }

            return OrderedDT;
        }

        public void OrdersSummaryInfoToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            int WorkAssignmentID, string DispatchDate, string BatchName, string ClientName)
        {
            int RowIndex = 0;
            HSSFSheet sheet1 = hssfworkbook.CreateSheet("Заказы");
            sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

            sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
            sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
            sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
            sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

            sheet1.SetColumnWidth(0, 20 * 256);

            {
                HSSFPatriarch patriarch = sheet1.CreateDrawingPatriarch();
                int ind = 0;

                CreateBarcode(WorkAssignmentID);
                HSSFClientAnchor anchor = new HSSFClientAnchor(0, 1, 1000, 8, 5, ind, 6, ind + 3)
                {
                    AnchorType = 2
                };
                string BarcodeNumber = GetBarcodeNumber(24, WorkAssignmentID);
                string barcodeName = System.Environment.GetEnvironmentVariable("TEMP") + @"/" + BarcodeNumber + ".jpg";
                HSSFPicture picture = patriarch.CreatePicture(anchor, LoadImage(barcodeName, hssfworkbook));

                HSSFFont BarcodeFont = hssfworkbook.CreateFont();
                BarcodeFont.FontHeightInPoints = 12;
                BarcodeFont.FontName = "Calibri";

                HSSFCellStyle BarcodeStyle = hssfworkbook.CreateCellStyle();
                BarcodeStyle.Alignment = HSSFCellStyle.ALIGN_CENTER;
                BarcodeStyle.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
                BarcodeStyle.SetFont(BarcodeFont);

                HSSFCell barcodeCell = sheet1.CreateRow(ind + 3).CreateCell(5);
                barcodeCell.SetCellValue(BarcodeNumber.ToString());
                barcodeCell.CellStyle = BarcodeStyle;
                sheet1.AddMergedRegion(new NPOI.HSSF.Util.Region(ind + 3, 5, ind + 3, 6));
            }

            decimal AllSquare = 0;
            string FrontName = string.Empty;

            if (TechnoNOrdersDT.Rows.Count > 0)
            {
                FrontName = GetFrontName(Convert.ToInt32(TechnoNOrdersDT.Rows[0]["FrontID"]));
                SummaryOrders(TechnoNOrdersDT, TechnoNSimpleDT, TechnoNVitrinaDT, TechnoNGridsDT, FrontName, ref AllSquare);
                OrdersToExcelSingly(ref hssfworkbook, ref sheet1,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, DispatchDate, BatchName, ClientName, ref RowIndex);
                RowIndex++;
                RowIndex++;
            }
            if (AntaliaOrdersDT.Rows.Count > 0)
            {
                FrontName = GetFrontName(Convert.ToInt32(AntaliaOrdersDT.Rows[0]["FrontID"]));
                SummaryOrders(AntaliaOrdersDT, AntaliaSimpleDT, AntaliaVitrinaDT, AntaliaGridsDT, FrontName, ref AllSquare);
                OrdersToExcelSingly(ref hssfworkbook, ref sheet1,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, DispatchDate, BatchName, ClientName, ref RowIndex);
                RowIndex++;
                RowIndex++;
            }
            if (Nord95OrdersDT.Rows.Count > 0)
            {
                FrontName = GetFrontName(Convert.ToInt32(Nord95OrdersDT.Rows[0]["FrontID"]));
                SummaryOrders(Nord95OrdersDT, Nord95SimpleDT, Nord95VitrinaDT, Nord95GridsDT, FrontName, ref AllSquare);
                OrdersToExcelSingly(ref hssfworkbook, ref sheet1,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, DispatchDate, BatchName, ClientName, ref RowIndex);
                RowIndex++;
                RowIndex++;
            }
            if (epFoxOrdersDT.Rows.Count > 0)
            {
                FrontName = GetFrontName(Convert.ToInt32(epFoxOrdersDT.Rows[0]["FrontID"]));
                SummaryOrders(epFoxOrdersDT, epFoxSimpleDT, epFoxVitrinaDT, epFoxGridsDT, FrontName, ref AllSquare);
                OrdersToExcelSingly(ref hssfworkbook, ref sheet1,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, DispatchDate, BatchName, ClientName, ref RowIndex);
                RowIndex++;
                RowIndex++;
            }

            if (BostonOrdersDT.Rows.Count > 0)
            {
                FrontName = GetFrontName(Convert.ToInt32(BostonOrdersDT.Rows[0]["FrontID"]));
                SummaryOrders(BostonOrdersDT, BostonSimpleDT, BostonVitrinaDT, BostonGridsDT, BostonGlassDT, FrontName, ref AllSquare);
                OrdersToExcelSingly(ref hssfworkbook, ref sheet1,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, DispatchDate, BatchName, ClientName, ref RowIndex);
                RowIndex++;
                RowIndex++;
            }

            if (BergamoOrdersDT.Rows.Count > 0)
            {
                FrontName = GetFrontName(Convert.ToInt32(BergamoOrdersDT.Rows[0]["FrontID"]));
                SummaryOrders(BergamoOrdersDT, BergamoSimpleDT, BergamoVitrinaDT, BergamoGridsDT, BergamoGlassDT, FrontName, ref AllSquare);
                OrdersToExcelSingly(ref hssfworkbook, ref sheet1,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, DispatchDate, BatchName, ClientName, ref RowIndex);
                RowIndex++;
                RowIndex++;
            }

            if (ep041OrdersDT.Rows.Count > 0)
            {
                FrontName = GetFrontName(Convert.ToInt32(ep041OrdersDT.Rows[0]["FrontID"]));
                SummaryOrders(ep041OrdersDT, ep041SimpleDT, ep041VitrinaDT, ep041GridsDT, FrontName, ref AllSquare);
                OrdersToExcelSingly(ref hssfworkbook, ref sheet1,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, DispatchDate, BatchName, ClientName, ref RowIndex);
                RowIndex++;
                RowIndex++;
            }

            if (ep071OrdersDT.Rows.Count > 0)
            {
                FrontName = GetFrontName(Convert.ToInt32(ep071OrdersDT.Rows[0]["FrontID"]));
                SummaryOrders(ep071OrdersDT, ep071SimpleDT, ep071VitrinaDT, ep071GridsDT, FrontName, ref AllSquare);
                OrdersToExcelSingly(ref hssfworkbook, ref sheet1,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, DispatchDate, BatchName, ClientName, ref RowIndex);
                RowIndex++;
                RowIndex++;
            }

            if (ep206OrdersDT.Rows.Count > 0)
            {
                FrontName = GetFrontName(Convert.ToInt32(ep206OrdersDT.Rows[0]["FrontID"]));
                SummaryOrders(ep206OrdersDT, ep206SimpleDT, ep206VitrinaDT, ep206GridsDT, FrontName, ref AllSquare);
                OrdersToExcelSingly(ref hssfworkbook, ref sheet1,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, DispatchDate, BatchName, ClientName, ref RowIndex);
                RowIndex++;
                RowIndex++;
            }

            if (ep216OrdersDT.Rows.Count > 0)
            {
                FrontName = GetFrontName(Convert.ToInt32(ep216OrdersDT.Rows[0]["FrontID"]));
                SummaryOrders(ep216OrdersDT, ep216SimpleDT, ep216VitrinaDT, ep216GridsDT, FrontName, ref AllSquare);
                OrdersToExcelSingly(ref hssfworkbook, ref sheet1,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, DispatchDate, BatchName, ClientName, ref RowIndex);
                RowIndex++;
                RowIndex++;
            }

            if (ep111OrdersDT.Rows.Count > 0)
            {
                FrontName = GetFrontName(Convert.ToInt32(ep111OrdersDT.Rows[0]["FrontID"]));
                SummaryOrders(ep111OrdersDT, ep111SimpleDT, ep111VitrinaDT, ep111GridsDT, FrontName, ref AllSquare);
                OrdersToExcelSingly(ref hssfworkbook, ref sheet1,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, DispatchDate, BatchName, ClientName, ref RowIndex);
                RowIndex++;
                RowIndex++;
            }
            if (VeneciaOrdersDT.Rows.Count > 0)
            {
                DataTable TempDT = new DataTable();
                using (DataView DV = new DataView(VeneciaOrdersDT))
                {
                    TempDT = DV.ToTable(true, new string[] { "ProfileID" });
                }
                for (int i = 0; i < TempDT.Rows.Count; i++)
                {
                    int ProfileID = Convert.ToInt32(TempDT.Rows[i]["ProfileID"]);
                    DataTable TempOrdersDT = MilanoOrdersDT.Clone();
                    DataTable TempSimpleDT = MilanoOrdersDT.Clone();
                    DataTable TempVitrinaDT = MilanoOrdersDT.Clone();
                    DataTable TempGridsDT = MilanoOrdersDT.Clone();
                    DataRow[] rows = VeneciaOrdersDT.Select("ProfileID=" + ProfileID);
                    if (rows.Count() > 0)
                    {
                        foreach (DataRow item in rows)
                            TempOrdersDT.Rows.Add(item.ItemArray);

                        rows = VeneciaSimpleDT.Select("ProfileID=" + ProfileID);
                        foreach (DataRow item in rows)
                            TempSimpleDT.Rows.Add(item.ItemArray);

                        rows = VeneciaVitrinaDT.Select("ProfileID=" + ProfileID);
                        foreach (DataRow item in rows)
                            TempVitrinaDT.Rows.Add(item.ItemArray);

                        rows = VeneciaGridsDT.Select("ProfileID=" + ProfileID);
                        foreach (DataRow item in rows)
                            TempGridsDT.Rows.Add(item.ItemArray);

                        FrontName = GetFrontName(Convert.ToInt32(TempOrdersDT.Rows[0]["FrontID"]));
                        SummaryOrders(TempOrdersDT, TempSimpleDT, TempVitrinaDT, TempGridsDT, FrontName, ref AllSquare);
                        OrdersToExcelSingly(ref hssfworkbook, ref sheet1,
                                Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, DispatchDate, BatchName, ClientName, ref RowIndex);
                        RowIndex++;
                        RowIndex++;
                    }
                    TempOrdersDT.Dispose();
                    TempSimpleDT.Dispose();
                    TempVitrinaDT.Dispose();
                    TempGridsDT.Dispose();
                }
                TempDT.Dispose();
            }
            if (LeonOrdersDT.Rows.Count > 0)
            {
                DataTable TempDT = new DataTable();
                using (DataView DV = new DataView(LeonOrdersDT))
                {
                    TempDT = DV.ToTable(true, new string[] { "ProfileID" });
                }
                for (int i = 0; i < TempDT.Rows.Count; i++)
                {
                    int ProfileID = Convert.ToInt32(TempDT.Rows[i]["ProfileID"]);
                    DataTable TempOrdersDT = MilanoOrdersDT.Clone();
                    DataTable TempSimpleDT = MilanoOrdersDT.Clone();
                    DataTable TempVitrinaDT = MilanoOrdersDT.Clone();
                    DataTable TempGridsDT = MilanoOrdersDT.Clone();
                    DataRow[] rows = LeonOrdersDT.Select("ProfileID=" + ProfileID);
                    if (rows.Count() > 0)
                    {
                        foreach (DataRow item in rows)
                            TempOrdersDT.Rows.Add(item.ItemArray);

                        rows = LeonSimpleDT.Select("ProfileID=" + ProfileID);
                        foreach (DataRow item in rows)
                            TempSimpleDT.Rows.Add(item.ItemArray);

                        rows = LeonVitrinaDT.Select("ProfileID=" + ProfileID);
                        foreach (DataRow item in rows)
                            TempVitrinaDT.Rows.Add(item.ItemArray);

                        rows = LeonGridsDT.Select("ProfileID=" + ProfileID);
                        foreach (DataRow item in rows)
                            TempGridsDT.Rows.Add(item.ItemArray);

                        FrontName = GetFrontName(Convert.ToInt32(TempOrdersDT.Rows[0]["FrontID"]));
                        SummaryOrders(TempOrdersDT, TempSimpleDT, TempVitrinaDT, TempGridsDT, FrontName, ref AllSquare);
                        OrdersToExcelSingly(ref hssfworkbook, ref sheet1,
                                Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, DispatchDate, BatchName, ClientName, ref RowIndex);
                        RowIndex++;
                        RowIndex++;
                    }
                    TempOrdersDT.Dispose();
                    TempSimpleDT.Dispose();
                    TempVitrinaDT.Dispose();
                    TempGridsDT.Dispose();
                }
                TempDT.Dispose();
            }
            if (LimogOrdersDT.Rows.Count > 0)
            {
                FrontName = GetFrontName(Convert.ToInt32(LimogOrdersDT.Rows[0]["FrontID"]));
                SummaryOrders(LimogOrdersDT, LimogSimpleDT, LimogVitrinaDT, LimogGridsDT, FrontName, ref AllSquare);
                OrdersToExcelSingly(ref hssfworkbook, ref sheet1,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, DispatchDate, BatchName, ClientName, ref RowIndex);
                RowIndex++;
                RowIndex++;
            }
            if (ep066Marsel4OrdersDT.Rows.Count > 0)
            {
                FrontName = GetFrontName(Convert.ToInt32(ep066Marsel4OrdersDT.Rows[0]["FrontID"]));
                SummaryOrders(ep066Marsel4OrdersDT, ep066Marsel4SimpleDT, ep066Marsel4VitrinaDT, ep066Marsel4GridsDT, FrontName, ref AllSquare);
                OrdersToExcelSingly(ref hssfworkbook, ref sheet1,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, DispatchDate, BatchName, ClientName, ref RowIndex);
                RowIndex++;
                RowIndex++;
            }
            if (ep110JersyOrdersDT.Rows.Count > 0)
            {
                FrontName = GetFrontName(Convert.ToInt32(ep110JersyOrdersDT.Rows[0]["FrontID"]));
                SummaryOrders(ep110JersyOrdersDT, ep110JersySimpleDT, ep110JersyVitrinaDT, ep110JersyGridsDT, FrontName, ref AllSquare);
                OrdersToExcelSingly(ref hssfworkbook, ref sheet1,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, DispatchDate, BatchName, ClientName, ref RowIndex);
                RowIndex++;
                RowIndex++;
            }
            if (ep018Marsel1OrdersDT.Rows.Count > 0)
            {
                FrontName = GetFrontName(Convert.ToInt32(ep018Marsel1OrdersDT.Rows[0]["FrontID"]));
                SummaryOrders(ep018Marsel1OrdersDT, ep018Marsel1SimpleDT, ep018Marsel1VitrinaDT, ep018Marsel1GridsDT, FrontName, ref AllSquare);
                OrdersToExcelSingly(ref hssfworkbook, ref sheet1,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, DispatchDate, BatchName, ClientName, ref RowIndex);
                RowIndex++;
                RowIndex++;
            }
            if (ep043ShervudOrdersDT.Rows.Count > 0)
            {
                FrontName = GetFrontName(Convert.ToInt32(ep043ShervudOrdersDT.Rows[0]["FrontID"]));
                SummaryOrders(ep043ShervudOrdersDT, ep043ShervudSimpleDT, ep043ShervudVitrinaDT, ep043ShervudGridsDT, FrontName, ref AllSquare);
                OrdersToExcelSingly(ref hssfworkbook, ref sheet1,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, DispatchDate, BatchName, ClientName, ref RowIndex);
                RowIndex++;
                RowIndex++;
            }
            if (ep112OrdersDT.Rows.Count > 0)
            {
                FrontName = GetFrontName(Convert.ToInt32(ep112OrdersDT.Rows[0]["FrontID"]));
                SummaryOrders(ep112OrdersDT, ep112SimpleDT, ep112VitrinaDT, ep112GridsDT, FrontName, ref AllSquare);
                OrdersToExcelSingly(ref hssfworkbook, ref sheet1,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, DispatchDate, BatchName, ClientName, ref RowIndex);
                RowIndex++;
                RowIndex++;
            }
            if (UrbanOrdersDT.Rows.Count > 0)
            {
                FrontName = GetFrontName(Convert.ToInt32(UrbanOrdersDT.Rows[0]["FrontID"]));
                SummaryOrders(UrbanOrdersDT, UrbanSimpleDT, UrbanVitrinaDT, UrbanGridsDT, FrontName, ref AllSquare);
                OrdersToExcelSingly(ref hssfworkbook, ref sheet1,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, DispatchDate, BatchName, ClientName, ref RowIndex);
                RowIndex++;
                RowIndex++;
            }
            if (AlbyOrdersDT.Rows.Count > 0)
            {
                FrontName = GetFrontName(Convert.ToInt32(AlbyOrdersDT.Rows[0]["FrontID"]));
                SummaryOrders(AlbyOrdersDT, AlbySimpleDT, AlbyVitrinaDT, AlbyGridsDT, FrontName, ref AllSquare);
                OrdersToExcelSingly(ref hssfworkbook, ref sheet1,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, DispatchDate, BatchName, ClientName, ref RowIndex);
                RowIndex++;
                RowIndex++;
            }
            if (BrunoOrdersDT.Rows.Count > 0)
            {
                FrontName = GetFrontName(Convert.ToInt32(BrunoOrdersDT.Rows[0]["FrontID"]));
                SummaryOrders(BrunoOrdersDT, BrunoSimpleDT, BrunoVitrinaDT, BrunoGridsDT, FrontName, ref AllSquare);
                OrdersToExcelSingly(ref hssfworkbook, ref sheet1,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, DispatchDate, BatchName, ClientName, ref RowIndex);
                RowIndex++;
                RowIndex++;
            }
            if (epsh406Techno4OrdersDT.Rows.Count > 0)
            {
                FrontName = GetFrontName(Convert.ToInt32(epsh406Techno4OrdersDT.Rows[0]["FrontID"]));
                SummaryOrders(epsh406Techno4OrdersDT, epsh406Techno4SimpleDT, epsh406Techno4VitrinaDT, epsh406Techno4GridsDT, FrontName, ref AllSquare);
                OrdersToExcelSingly(ref hssfworkbook, ref sheet1,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, DispatchDate, BatchName, ClientName, ref RowIndex);
                RowIndex++;
                RowIndex++;
            }
            if (LukOrdersDT.Rows.Count > 0)
            {
                FrontName = GetFrontName(Convert.ToInt32(LukOrdersDT.Rows[0]["FrontID"]));
                SummaryOrders(LukOrdersDT, LukSimpleDT, LukVitrinaDT, LukGridsDT, FrontName, ref AllSquare);
                OrdersToExcelSingly(ref hssfworkbook, ref sheet1,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, DispatchDate, BatchName, ClientName, ref RowIndex);
                RowIndex++;
                RowIndex++;
            }
            if (LukPVHOrdersDT.Rows.Count > 0)
            {
                FrontName = GetFrontName(Convert.ToInt32(LukPVHOrdersDT.Rows[0]["FrontID"]));
                SummaryOrders(LukPVHOrdersDT, LukPVHSimpleDT, LukPVHVitrinaDT, LukPVHGridsDT, FrontName, ref AllSquare);
                OrdersToExcelSingly(ref hssfworkbook, ref sheet1,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, DispatchDate, BatchName, ClientName, ref RowIndex);
                RowIndex++;
                RowIndex++;
            }
            if (MilanoOrdersDT.Rows.Count > 0)
            {
                DataTable TempDT = new DataTable();
                using (DataView DV = new DataView(MilanoOrdersDT))
                {
                    TempDT = DV.ToTable(true, new string[] { "ProfileID" });
                }
                for (int i = 0; i < TempDT.Rows.Count; i++)
                {
                    int ProfileID = Convert.ToInt32(TempDT.Rows[i]["ProfileID"]);
                    DataTable TempOrdersDT = MilanoOrdersDT.Clone();
                    DataTable TempSimpleDT = MilanoOrdersDT.Clone();
                    DataTable TempVitrinaDT = MilanoOrdersDT.Clone();
                    DataTable TempGridsDT = MilanoOrdersDT.Clone();
                    DataRow[] rows = MilanoOrdersDT.Select("ProfileID=" + ProfileID);
                    if (rows.Count() > 0)
                    {
                        foreach (DataRow item in rows)
                            TempOrdersDT.Rows.Add(item.ItemArray);

                        rows = MilanoSimpleDT.Select("ProfileID=" + ProfileID);
                        foreach (DataRow item in rows)
                            TempSimpleDT.Rows.Add(item.ItemArray);

                        rows = MilanoVitrinaDT.Select("ProfileID=" + ProfileID);
                        foreach (DataRow item in rows)
                            TempVitrinaDT.Rows.Add(item.ItemArray);

                        rows = MilanoGridsDT.Select("ProfileID=" + ProfileID);
                        foreach (DataRow item in rows)
                            TempGridsDT.Rows.Add(item.ItemArray);

                        FrontName = GetFrontName(Convert.ToInt32(TempOrdersDT.Rows[0]["FrontID"]));
                        SummaryOrders(TempOrdersDT, TempSimpleDT, TempVitrinaDT, TempGridsDT, FrontName, ref AllSquare);
                        OrdersToExcelSingly(ref hssfworkbook, ref sheet1,
                                Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, DispatchDate, BatchName, ClientName, ref RowIndex);
                        RowIndex++;
                        RowIndex++;
                    }
                    TempOrdersDT.Dispose();
                    TempSimpleDT.Dispose();
                    TempVitrinaDT.Dispose();
                    TempGridsDT.Dispose();
                }
                TempDT.Dispose();
            }
            if (PragaOrdersDT.Rows.Count > 0)
            {
                FrontName = GetFrontName(Convert.ToInt32(PragaOrdersDT.Rows[0]["FrontID"]));
                SummaryOrders(PragaOrdersDT, PragaSimpleDT, PragaVitrinaDT, PragaGridsDT, FrontName, ref AllSquare);
                OrdersToExcelSingly(ref hssfworkbook, ref sheet1,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, DispatchDate, BatchName, ClientName, ref RowIndex);
                RowIndex++;
                RowIndex++;
            }
            if (SigmaOrdersDT.Rows.Count > 0)
            {
                FrontName = GetFrontName(Convert.ToInt32(SigmaOrdersDT.Rows[0]["FrontID"]));
                SummaryOrders(SigmaOrdersDT, SigmaSimpleDT, SigmaVitrinaDT, SigmaGlassDT, FrontName, ref AllSquare);
                OrdersToExcelSingly(ref hssfworkbook, ref sheet1,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, DispatchDate, BatchName, ClientName, ref RowIndex);
                RowIndex++;
                RowIndex++;
            }
            if (FatOrdersDT.Rows.Count > 0)
            {
                FrontName = GetFrontName(Convert.ToInt32(FatOrdersDT.Rows[0]["FrontID"]));
                SummaryOrders(FatOrdersDT, FatSimpleDT, FatVitrinaDT, FatGridsDT, FrontName, ref AllSquare);
                OrdersToExcelSingly(ref hssfworkbook, ref sheet1,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, DispatchDate, BatchName, ClientName, ref RowIndex);
                RowIndex++;
                RowIndex++;
            }
            AllSquare = decimal.Round(AllSquare, 3, MidpointRounding.AwayFromZero);
            OrdersToExcelSingly(ref hssfworkbook, CalibriBold11CS, CalibriBold11CS, ref sheet1, AllSquare, ref RowIndex);
        }

        public void OrdersToExcelSingly(ref HSSFWorkbook hssfworkbook, ref HSSFSheet sheet1,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            int WorkAssignmentID, string DispatchDate, string BatchName, string ClientName, ref int RowIndex)
        {
            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Плановое время выполнения:");
            cell.CellStyle = Calibri11CS;
            if (DispatchDate.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Отгрузка: " + DispatchDate);
                cell.CellStyle = CalibriBold11CS;
            }
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, "Заказы");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, ClientName);
            cell.CellStyle = CalibriBold11CS;

            int ColumnIndex = -1;
            string ColumnName = string.Empty;

            for (int x = 0; x < SummOrdersDT.Columns.Count; x++)
            {
                if (SummOrdersDT.Columns[x].ColumnName == "Height" || SummOrdersDT.Columns[x].ColumnName == "Width")
                    continue;
                ColumnIndex++;
                ColumnName = SummOrdersDT.Columns[x].ColumnName;
                if (Contains(ColumnName, "_", StringComparison.OrdinalIgnoreCase))
                {
                    ColumnName = ColumnName.Substring(0, ColumnName.Length - 2);
                }
                if (ColumnName == "Sizes")
                {
                    ColumnName = "Размер";
                    cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), ColumnIndex, ColumnName);
                    cell.CellStyle = TableHeaderCS;
                    continue;
                }
                if (ColumnName == "TotalAmount")
                {
                    ColumnName = "Итого";
                    cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), ColumnIndex, ColumnName);
                    cell.CellStyle = TableHeaderCS;
                    continue;
                }
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), ColumnIndex, ColumnName);
                cell.CellStyle = TableHeaderCS;
                sheet1.SetColumnWidth(ColumnIndex, 19 * 256);
            }
            RowIndex++;
            TableHeaderCS.Alignment = HSSFCellStyle.ALIGN_LEFT;
            TableHeaderDecCS.Alignment = HSSFCellStyle.ALIGN_LEFT;

            HSSFFont FirstColF = hssfworkbook.CreateFont();
            FirstColF.FontHeightInPoints = 12;
            FirstColF.FontName = "MS Sans Serif";

            HSSFCellStyle FirstColCS = hssfworkbook.CreateCellStyle();
            FirstColCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            FirstColCS.LeftBorderColor = HSSFColor.BLACK.index;
            FirstColCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            FirstColCS.RightBorderColor = HSSFColor.BLACK.index;
            FirstColCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            FirstColCS.TopBorderColor = HSSFColor.BLACK.index;
            FirstColCS.BorderBottom = HSSFCellStyle.BORDER_THIN;
            FirstColCS.BottomBorderColor = HSSFColor.BLACK.index;
            FirstColCS.SetFont(FirstColF);

            for (int x = 0; x < SummOrdersDT.Rows.Count; x++)
            {
                ColumnIndex = -1;
                for (int y = 0; y < SummOrdersDT.Columns.Count; y++)
                {
                    if (SummOrdersDT.Columns[y].ColumnName == "Height" || SummOrdersDT.Columns[y].ColumnName == "Width")
                        continue;
                    Type t = SummOrdersDT.Rows[x][y].GetType();

                    ColumnIndex++;

                    if (x == SummOrdersDT.Rows.Count - 1 && int.TryParse(SummOrdersDT.Rows[x][y].ToString(), out int IntValue))
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(ColumnIndex);
                        cell.SetCellValue(IntValue);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (x == SummOrdersDT.Rows.Count - 2 && double.TryParse(SummOrdersDT.Rows[x][y].ToString(), out double DecValue))
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(ColumnIndex);
                        cell.SetCellValue(DecValue);
                        cell.CellStyle = TableHeaderDecCS;
                        continue;
                    }

                    if (int.TryParse(SummOrdersDT.Rows[x][y].ToString(), out IntValue))
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(ColumnIndex);
                        cell.SetCellValue(IntValue);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (t.Name == "String" || t.Name == "DBNull")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(ColumnIndex);
                        cell.SetCellValue(SummOrdersDT.Rows[x][y].ToString());

                        if (ColumnIndex == 0)
                            cell.CellStyle = FirstColCS;
                        else
                            cell.CellStyle = TableHeaderCS;
                        continue;
                    }
                }
                RowIndex++;
            }
        }

        public void OrdersToExcelSingly(ref HSSFWorkbook hssfworkbook, HSSFCellStyle CalibriBold11CS, HSSFCellStyle TableHeaderDecCS,
            ref HSSFSheet sheet1, decimal AllSquare, ref int RowIndex)
        {
            CalibriBold11CS.Alignment = HSSFCellStyle.ALIGN_LEFT;
            TableHeaderDecCS.Alignment = HSSFCellStyle.ALIGN_LEFT;

            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "ИТОГО:");
            cell.CellStyle = CalibriBold11CS;

            cell = sheet1.CreateRow(RowIndex).CreateCell(1);
            cell.SetCellValue(Convert.ToDouble(AllSquare));
            cell.CellStyle = TableHeaderDecCS;
        }

        public void PressToExcelSingly(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            ref HSSFSheet sheet1, DataTable DT, int WorkAssignmentID, string DispatchDate, string BatchName, string ClientName, string TableName, ref int RowIndex)
        {
            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Плановое время выполнения:");
            cell.CellStyle = Calibri11CS;

            if (DispatchDate.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Отгрузка: " + DispatchDate);
                cell.CellStyle = CalibriBold11CS;
            }

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, TableName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, ClientName);
            cell.CellStyle = CalibriBold11CS;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Цвет наполнителя");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Работник");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            decimal TotalSquare = 0;
            decimal AllTotalSquare = 0;
            string str = string.Empty;

            int CType = -1;
            int DifferentDecorCount = 0;

            using (DataView DV = new DataView(DT))
            {
                DifferentDecorCount = DV.ToTable(true, new string[] { "InsetColorID" }).Rows.Count;
            }

            if (DT.Rows.Count > 0)
            {
                CType = Convert.ToInt32(DT.Rows[0]["InsetColorID"]);
            }
            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value)
                {
                    TotalSquare += Convert.ToDecimal(DT.Rows[x]["Height"]) * Convert.ToDecimal(DT.Rows[x]["Width"]) * Convert.ToDecimal(DT.Rows[x]["Count"]) / 1000000;
                    AllTotalSquare += Convert.ToDecimal(DT.Rows[x]["Height"]) * Convert.ToDecimal(DT.Rows[x]["Width"]) * Convert.ToDecimal(DT.Rows[x]["Count"]) / 1000000;
                }

                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "InsetColorID" || DT.Columns[y].ColumnName == "TechnoInsetColorID" ||
                        DT.Columns[y].ColumnName == "GlassCount" || DT.Columns[y].ColumnName == "MegaCount")
                        continue;
                    Type t = DT.Rows[x][y].GetType();

                    if (DT.Columns[y].ColumnName == "Name")
                        str = DT.Rows[x][y].ToString();

                    if (t.Name == "Decimal")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToDouble(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderDecCS;
                        continue;
                    }
                    if (t.Name == "Int32")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToInt32(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (t.Name == "String" || t.Name == "DBNull")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(DT.Rows[x][y].ToString());
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }
                }

                if (x + 1 <= DT.Rows.Count - 1)
                {
                    if (CType != Convert.ToInt32(DT.Rows[x + 1]["InsetColorID"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "InsetColorID" || DT.Columns[y].ColumnName == "TechnoInsetColorID" ||
                                DT.Columns[y].ColumnName == "GlassCount" || DT.Columns[y].ColumnName == "MegaCount")
                                continue;
                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        TotalSquare = decimal.Round(TotalSquare, 2, MidpointRounding.AwayFromZero);

                        cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                        cell.SetCellValue(Convert.ToDouble(TotalSquare));
                        cell.CellStyle = TableHeaderDecCS;

                        CType = Convert.ToInt32(DT.Rows[x + 1]["InsetColorID"]);
                        TotalSquare = 0;

                        RowIndex++;
                    }
                }

                if (x == DT.Rows.Count - 1)
                {
                    if (DifferentDecorCount > 1)
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "InsetColorID" || DT.Columns[y].ColumnName == "TechnoInsetColorID" ||
                                DT.Columns[y].ColumnName == "GlassCount" || DT.Columns[y].ColumnName == "MegaCount")
                                continue;
                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        TotalSquare = decimal.Round(TotalSquare, 2, MidpointRounding.AwayFromZero);

                        cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                        cell.SetCellValue(Convert.ToDouble(TotalSquare));
                        cell.CellStyle = TableHeaderDecCS;
                    }
                    RowIndex++;

                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "InsetColorID" || DT.Columns[y].ColumnName == "TechnoInsetColorID" ||
                            DT.Columns[y].ColumnName == "GlassCount" || DT.Columns[y].ColumnName == "MegaCount")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    AllTotalSquare = decimal.Round(AllTotalSquare, 2, MidpointRounding.AwayFromZero);

                    cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                    cell.SetCellValue(Convert.ToDouble(AllTotalSquare));
                    cell.CellStyle = TableHeaderDecCS;
                }
                RowIndex++;
            }
        }

        public void RapidToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            int WorkAssignmentID, string DispatchDate, string BatchName, string ClientName)
        {
            int RowIndex = 0;

            HSSFSheet sheet1 = hssfworkbook.CreateSheet("Рапид");
            sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

            sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
            sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
            sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
            sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

            sheet1.SetColumnWidth(0, 25 * 256);
            sheet1.SetColumnWidth(1, 11 * 256);
            sheet1.SetColumnWidth(2, 6 * 256);
            sheet1.SetColumnWidth(3, 6 * 256);
            sheet1.SetColumnWidth(4, 16 * 256);

            {
                HSSFPatriarch patriarch = sheet1.CreateDrawingPatriarch();
                int ind = 0;

                CreateBarcode(WorkAssignmentID);
                HSSFClientAnchor anchor = new HSSFClientAnchor(0, 1, 1000, 8, 5, ind, 6, ind + 3)
                {
                    AnchorType = 2
                };
                string BarcodeNumber = GetBarcodeNumber(24, WorkAssignmentID);
                string barcodeName = System.Environment.GetEnvironmentVariable("TEMP") + @"/" + BarcodeNumber + ".jpg";
                HSSFPicture picture = patriarch.CreatePicture(anchor, LoadImage(barcodeName, hssfworkbook));

                HSSFFont BarcodeFont = hssfworkbook.CreateFont();
                BarcodeFont.FontHeightInPoints = 12;
                BarcodeFont.FontName = "Calibri";

                HSSFCellStyle BarcodeStyle = hssfworkbook.CreateCellStyle();
                BarcodeStyle.Alignment = HSSFCellStyle.ALIGN_CENTER;
                BarcodeStyle.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
                BarcodeStyle.SetFont(BarcodeFont);

                HSSFCell barcodeCell = sheet1.CreateRow(ind + 3).CreateCell(5);
                barcodeCell.SetCellValue(BarcodeNumber.ToString());
                barcodeCell.CellStyle = BarcodeStyle;
                sheet1.AddMergedRegion(new NPOI.HSSF.Util.Region(ind + 3, 5, ind + 3, 6));
            }

            RapidDT.Clear();
            CombineRapidSimple(ref RapidDT, true);
            CombineRapidFilenka(ref RapidDT);
            CombineRapidBoxes(ref RapidDT);
            CombineRapidFilenkaBoxes(ref RapidDT);

            DataTable DT = RapidDT.Copy();

            if (DT.Rows.Count > 0)
            {
                RapidToExcelSingly(ref hssfworkbook, ref sheet1,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DT, WorkAssignmentID, "РапидТ", DispatchDate, BatchName, ClientName, ref RowIndex);
                RowIndex++;
                RowIndex++;
            }

            ComecDT.Clear();
            CombineComecSimple(ref ComecDT, true);
            CombineComecFilenka(ref ComecDT);

            DT = new DataTable();
            DT = ComecDT.Copy();

            if (DT.Rows.Count > 0)
            {
                ComecToExcelSingly(ref hssfworkbook, ref sheet1,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DT, WorkAssignmentID, "Comec", DispatchDate, BatchName, ClientName, ref RowIndex);
                RowIndex++;
                RowIndex++;
            }
        }

        public void RapidToExcelSingly(ref HSSFWorkbook hssfworkbook, ref HSSFSheet sheet1,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            DataTable DT, int WorkAssignmentID, string SheetName, string DispatchDate, string BatchName, string ClientName, ref int RowIndex)
        {
            HSSFCell cell = null;

            if (DispatchDate.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Отгрузка: " + DispatchDate);
                cell.CellStyle = CalibriBold11CS;
            }
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, SheetName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 2, ClientName);
            cell.CellStyle = CalibriBold11CS;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Профиль");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Цвет профиля");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Работник");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            decimal SticksCount = 0;
            int CType = 0;
            int FType = 0;
            int PType = 0;
            int TotalAmount = 0;
            int AllTotalAmount = 0;
            int Count = 0;
            int Height = 0;

            if (DT.Rows.Count > 0)
            {
                CType = Convert.ToInt32(DT.Rows[0]["ColorType"]);
                FType = Convert.ToInt32(DT.Rows[0]["FrontType"]);
                PType = Convert.ToInt32(DT.Rows[0]["ProfileType"]);
            }

            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value && DT.Rows[x]["Height"] != DBNull.Value)
                {
                    Count = Convert.ToInt32(DT.Rows[x]["Count"]);
                    Height = Convert.ToInt32(DT.Rows[x]["Height"]);
                    SticksCount += (Height + 4) * Count;
                    TotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    AllTotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    Count = Convert.ToInt32(DT.Rows[x]["Count"]);
                }

                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "ProfileType" || DT.Columns[y].ColumnName == "ColorType"
                        || DT.Columns[y].ColumnName == "FrontType" || DT.Columns[y].ColumnName == "VitrinaCount"
                        || DT.Columns[y].ColumnName == "BoxCount" || DT.Columns[y].ColumnName == "ImpostCount")
                        continue;
                    Type t = DT.Rows[x][y].GetType();

                    if (t.Name == "Decimal")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToDouble(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderDecCS;
                        continue;
                    }
                    if (t.Name == "Int32")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToInt32(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (t.Name == "String" || t.Name == "DBNull")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(DT.Rows[x][y].ToString());
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }
                }

                if (DT.Rows[x]["ColorType"] != DBNull.Value)
                    CType = Convert.ToInt32(DT.Rows[x]["ColorType"]);
                if (DT.Rows[x]["FrontType"] != DBNull.Value)
                    FType = Convert.ToInt32(DT.Rows[x]["FrontType"]);
                if (DT.Rows[x]["ProfileType"] != DBNull.Value)
                    PType = Convert.ToInt32(DT.Rows[x]["ProfileType"]);
                if (x + 1 <= DT.Rows.Count - 1 &&
                    (FType != Convert.ToInt32(DT.Rows[x + 1]["FrontType"]) || PType != Convert.ToInt32(DT.Rows[x + 1]["ProfileType"]) || CType != Convert.ToInt32(DT.Rows[x + 1]["ColorType"])))
                {
                    RowIndex++;
                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "ProfileType" || DT.Columns[y].ColumnName == "ColorType"
                            || DT.Columns[y].ColumnName == "FrontType" || DT.Columns[y].ColumnName == "VitrinaCount"
                            || DT.Columns[y].ColumnName == "BoxCount" || DT.Columns[y].ColumnName == "ImpostCount")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    SticksCount = SticksCount * 1.15m / 2620;
                    SticksCount = decimal.Round(SticksCount, 1, MidpointRounding.AwayFromZero);
                    cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                    cell.SetCellValue(SticksCount + " палок");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Итого:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                    cell.SetCellValue(TotalAmount);
                    cell.CellStyle = TableHeaderCS;
                    RowIndex++;

                    CType = Convert.ToInt32(DT.Rows[x + 1]["ColorType"]);
                    PType = Convert.ToInt32(DT.Rows[x + 1]["ProfileType"]);
                    Count = 0;
                    Height = 0;
                    SticksCount = 0;
                    TotalAmount = 0;
                }

                if (x == DT.Rows.Count - 1)
                {
                    RowIndex++;
                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "ProfileType" || DT.Columns[y].ColumnName == "ColorType"
                            || DT.Columns[y].ColumnName == "FrontType" || DT.Columns[y].ColumnName == "VitrinaCount"
                            || DT.Columns[y].ColumnName == "BoxCount" || DT.Columns[y].ColumnName == "ImpostCount")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    SticksCount = SticksCount * 1.15m / 2620;
                    SticksCount = decimal.Round(SticksCount, 1, MidpointRounding.AwayFromZero);
                    cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                    cell.SetCellValue(SticksCount + " палок");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Итого:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                    cell.SetCellValue(TotalAmount);
                    cell.CellStyle = TableHeaderCS;

                    RowIndex++;
                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "ProfileType" || DT.Columns[y].ColumnName == "ColorType"
                            || DT.Columns[y].ColumnName == "FrontType" || DT.Columns[y].ColumnName == "VitrinaCount"
                            || DT.Columns[y].ColumnName == "BoxCount" || DT.Columns[y].ColumnName == "ImpostCount")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;
                }
                RowIndex++;
            }
        }

        #endregion Public Methods

        #region Private Methods

        private void AdditionsSyngly(DataTable SourceDT, ref DataTable DestinationDT, string Profile, string InsetColor, int Margin)
        {
            DataTable DT1 = new DataTable();

            using (DataView DV = new DataView(SourceDT))
            {
                DT1 = DV.ToTable(true, new string[] { "ColorID", "TechnoInsetColorID", "Height", "Width" });
            }

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                DataRow[] Srows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                    " AND TechnoInsetColorID=" + Convert.ToInt32(DT1.Rows[i]["TechnoInsetColorID"]) + " AND Width=" + Convert.ToInt32(DT1.Rows[i]["Width"]) +
                    " AND Height=" + Convert.ToInt32(DT1.Rows[i]["Height"]));
                if (Srows.Count() > 0)
                {
                    int Count = 0;

                    foreach (DataRow item in Srows)
                        Count += Convert.ToInt32(item["Count"]);

                    int Height = Convert.ToInt32(DT1.Rows[i]["Height"]) - Margin;
                    int Width = Convert.ToInt32(DT1.Rows[i]["Width"]) - Margin;

                    decimal Square = Convert.ToDecimal(Height) * Convert.ToDecimal(Width) * Convert.ToDecimal(Count) / 1000000;
                    Square = decimal.Round(Square, 3, MidpointRounding.AwayFromZero);

                    DataRow[] rows = DestinationDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                    " AND TechnoInsetColorID=" + Convert.ToInt32(DT1.Rows[i]["TechnoInsetColorID"]) + " AND Width=" + Width + " AND Height=" + Height);
                    if (rows.Count() == 0)
                    {
                        DataRow NewRow = DestinationDT.NewRow();
                        NewRow["Profile"] = Profile;
                        NewRow["Color"] = "-";
                        NewRow["InsetColor"] = InsetColor;
                        NewRow["Height"] = Height;
                        NewRow["Width"] = Width;
                        NewRow["Count"] = Count;
                        NewRow["Square"] = Square;
                        NewRow["ColorID"] = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                        NewRow["TechnoInsetColorID"] = Convert.ToInt32(DT1.Rows[i]["TechnoInsetColorID"]);
                        DestinationDT.Rows.Add(NewRow);
                    }
                    else
                    {
                        rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count;
                        rows[0]["Square"] = Convert.ToDecimal(rows[0]["Square"]) + Square;
                    }
                }
            }
            DT1.Dispose();
        }

        private void AllInsets(DataTable SourceDT, ref DataTable DestinationDT, int Margin, bool OrderASC)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            DataTable DT3 = new DataTable();
            DataTable DT4 = new DataTable();

            using (DataView DV = new DataView(SourceDT))
            {
                DT1 = DV.ToTable(true, new string[] { "InsetTypeID" });
            }

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                string InsetTypeName = "";
                InsetTypeName = GetInsetTypeName(Convert.ToInt32(DT1.Rows[i]["InsetTypeID"])) + " ";
                using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]), "InsetColorID", DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "InsetColorID" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                        " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]), "TechnoInsetColorID", DataViewRowState.CurrentRows))
                    {
                        DT3 = DV.ToTable(true, new string[] { "TechnoInsetColorID" });
                    }
                    for (int x = 0; x < DT3.Rows.Count; x++)
                    {
                        using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                            " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) +
                            " AND TechnoInsetColorID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]), "Height, Width", DataViewRowState.CurrentRows))
                        {
                            DT4 = DV.ToTable(true, new string[] { "Height", "Width" });
                        }
                        for (int y = 0; y < DT4.Rows.Count; y++)
                        {
                            DataRow[] Srows = SourceDT.Select("InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                                " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND TechnoInsetColorID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]) +
                                " AND Height=" + Convert.ToInt32(DT4.Rows[y]["Height"]) + " AND Width=" + Convert.ToInt32(DT4.Rows[y]["Width"]));
                            if (Srows.Count() == 0)
                                continue;

                            int Count = 0;
                            int Height = Convert.ToInt32(DT4.Rows[y]["Height"]) - Margin;
                            int Width = Convert.ToInt32(DT4.Rows[y]["Width"]) - Margin;
                            int TechnoInsetColorID = Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]);

                            if (Height < 10 || Width < 10)
                                continue;

                            if (Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) == 2069 || Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) == 2070 || Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) == 2071
                                 || Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) == 2073 || Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) == 2075 || Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) == 2077
                                 || Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) == 2233 || Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) == 3644)
                            {
                                if (Height >= 100 && Width >= 100)
                                    continue;
                                if (Width > 900)
                                    continue;
                            }
                            string InsetColor = GetInsetColorName(Convert.ToInt32(DT2.Rows[j]["InsetColorID"]));

                            foreach (DataRow item in Srows)
                                Count += Convert.ToInt32(item["Count"]);

                            if (TechnoInsetColorID == 128)
                            {
                                InsetColor = "мега " + InsetColor + " вит";
                            }

                            DataRow[] rows = DestinationDT.Select("InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) +
                                " AND TechnoInsetColorID=" + TechnoInsetColorID +
                                " AND Height=" + Height + " AND Width=" + Width);
                            if (rows.Count() == 0)
                            {
                                DataRow NewRow = DestinationDT.NewRow();
                                NewRow["Name"] = InsetTypeName + InsetColor;
                                NewRow["Height"] = Height;
                                NewRow["Width"] = Width;
                                NewRow["Count"] = Count;
                                NewRow["FrontID"] = Convert.ToInt32(Srows[0]["FrontID"]);
                                NewRow["InsetColorID"] = Convert.ToInt32(DT2.Rows[j]["InsetColorID"]);
                                NewRow["TechnoInsetColorID"] = TechnoInsetColorID;
                                DestinationDT.Rows.Add(NewRow);
                            }
                            else
                            {
                                rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count;
                            }
                        }
                    }
                }
            }
        }

        private void ArchDecorAssemblyByMainOrderToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            int WorkAssignmentID, string BatchName)
        {
            if (ArchDecorOrdersDT.Rows.Count == 0)
                return;

            DataTable DistMainOrdersDT = DistMainOrdersTable(ArchDecorOrdersDT, true);
            DataTable DT = ArchDecorOrdersDT.Clone();
            DataTable ZOVOrdersNames = new DataTable();
            DataTable MarketOrdersNames = new DataTable();
            string MainOrdersID = string.Empty;
            string ClientName = string.Empty;
            string OrderName = string.Empty;
            string Notes = string.Empty;

            if (DistMainOrdersDT.Select("GroupType=0").Count() > 0)
            {
                MainOrdersID = string.Empty;
                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=0"))
                    MainOrdersID += Convert.ToInt32(item["MainOrderID"]) + ",";
                MainOrdersID = MainOrdersID.Substring(0, MainOrdersID.Length - 1);
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT ClientName, DocNumber, MainOrderID FROM MainOrders" +
                    " INNER JOIN infiniu2_zovreference.dbo.Clients ON MainOrders.ClientID=infiniu2_zovreference.dbo.Clients.ClientID" +
                    " WHERE MainOrderID IN (" + MainOrdersID + ")", ConnectionStrings.ZOVOrdersConnectionString))
                {
                    DA.Fill(ZOVOrdersNames);
                }

                int RowIndex = 0;
                HSSFSheet sheet1 = hssfworkbook.CreateSheet("Арки ЗОВ");
                sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

                sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
                sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
                sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
                sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

                sheet1.SetColumnWidth(0, 30 * 256);
                sheet1.SetColumnWidth(1, 20 * 256);
                sheet1.SetColumnWidth(2, 9 * 256);
                sheet1.SetColumnWidth(3, 6 * 256);
                sheet1.SetColumnWidth(4, 6 * 256);

                {
                    HSSFPatriarch patriarch = sheet1.CreateDrawingPatriarch();
                    int ind = 0;

                    CreateBarcode(WorkAssignmentID);
                    HSSFClientAnchor anchor = new HSSFClientAnchor(0, 1, 1000, 8, 5, ind, 6, ind + 3)
                    {
                        AnchorType = 2
                    };
                    string BarcodeNumber = GetBarcodeNumber(24, WorkAssignmentID);
                    string barcodeName = System.Environment.GetEnvironmentVariable("TEMP") + @"/" + BarcodeNumber + ".jpg";
                    HSSFPicture picture = patriarch.CreatePicture(anchor, LoadImage(barcodeName, hssfworkbook));

                    HSSFFont BarcodeFont = hssfworkbook.CreateFont();
                    BarcodeFont.FontHeightInPoints = 12;
                    BarcodeFont.FontName = "Calibri";

                    HSSFCellStyle BarcodeStyle = hssfworkbook.CreateCellStyle();
                    BarcodeStyle.Alignment = HSSFCellStyle.ALIGN_CENTER;
                    BarcodeStyle.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
                    BarcodeStyle.SetFont(BarcodeFont);

                    HSSFCell barcodeCell = sheet1.CreateRow(ind + 3).CreateCell(5);
                    barcodeCell.SetCellValue(BarcodeNumber.ToString());
                    barcodeCell.CellStyle = BarcodeStyle;
                    sheet1.AddMergedRegion(new NPOI.HSSF.Util.Region(ind + 3, 5, ind + 3, 6));
                }

                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=0"))
                {
                    DecorAssemblyDT.Clear();
                    DT.Clear();
                    int MainOrderID = Convert.ToInt32(item["MainOrderID"]);
                    DataRow[] rows = ArchDecorOrdersDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item1 in rows)
                        DT.Rows.Add(item1.ItemArray);
                    AssemblyDecorCollect(DT, ref DecorAssemblyDT);

                    DataRow[] CRows = ZOVOrdersNames.Select("MainOrderID=" + Convert.ToInt32(item["MainOrderID"]));
                    if (CRows.Count() > 0)
                    {
                        ClientName = CRows[0]["ClientName"].ToString();
                        OrderName = CRows[0]["DocNumber"].ToString();
                    }
                    if (DecorAssemblyDT.Rows.Count > 0)
                    {
                        ArchDecorAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                             Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DecorAssemblyDT,
                             WorkAssignmentID, BatchName, ClientName, string.Empty, OrderName, Notes, ref RowIndex);
                        RowIndex++;
                        ArchDecorAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                            Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DecorAssemblyDT,
                            WorkAssignmentID, BatchName, ClientName, "ДУБЛЬ", OrderName, Notes, ref RowIndex);
                    }
                    RowIndex++;
                }
            }
            if (DistMainOrdersDT.Select("GroupType=1").Count() > 0)
            {
                MainOrdersID = string.Empty;
                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=1"))
                    MainOrdersID += Convert.ToInt32(item["MainOrderID"]) + ",";
                MainOrdersID = MainOrdersID.Substring(0, MainOrdersID.Length - 1);
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT ClientName, OrderNumber, MainOrderID, Notes FROM MainOrders" +
                    " INNER JOIN MegaOrders ON MainOrders.MegaOrderID=MegaOrders.MegaOrderID" +
                    " INNER JOIN infiniu2_marketingreference.dbo.Clients ON MegaOrders.ClientID=infiniu2_marketingreference.dbo.Clients.ClientID" +
                    " WHERE MainOrderID IN (" + MainOrdersID + ")", ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DA.Fill(MarketOrdersNames);
                }

                int RowIndex = 0;
                HSSFSheet sheet1 = hssfworkbook.CreateSheet("Арки Маркетинг");
                sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

                sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
                sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
                sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
                sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

                sheet1.SetColumnWidth(0, 30 * 256);
                sheet1.SetColumnWidth(1, 20 * 256);
                sheet1.SetColumnWidth(2, 9 * 256);
                sheet1.SetColumnWidth(3, 6 * 256);
                sheet1.SetColumnWidth(4, 6 * 256);

                {
                    HSSFPatriarch patriarch = sheet1.CreateDrawingPatriarch();
                    int ind = 0;

                    CreateBarcode(WorkAssignmentID);
                    HSSFClientAnchor anchor = new HSSFClientAnchor(0, 1, 1000, 8, 5, ind, 6, ind + 3)
                    {
                        AnchorType = 2
                    };
                    string BarcodeNumber = GetBarcodeNumber(24, WorkAssignmentID);
                    string barcodeName = System.Environment.GetEnvironmentVariable("TEMP") + @"/" + BarcodeNumber + ".jpg";
                    HSSFPicture picture = patriarch.CreatePicture(anchor, LoadImage(barcodeName, hssfworkbook));

                    HSSFFont BarcodeFont = hssfworkbook.CreateFont();
                    BarcodeFont.FontHeightInPoints = 12;
                    BarcodeFont.FontName = "Calibri";

                    HSSFCellStyle BarcodeStyle = hssfworkbook.CreateCellStyle();
                    BarcodeStyle.Alignment = HSSFCellStyle.ALIGN_CENTER;
                    BarcodeStyle.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
                    BarcodeStyle.SetFont(BarcodeFont);

                    HSSFCell barcodeCell = sheet1.CreateRow(ind + 3).CreateCell(5);
                    barcodeCell.SetCellValue(BarcodeNumber.ToString());
                    barcodeCell.CellStyle = BarcodeStyle;
                    sheet1.AddMergedRegion(new NPOI.HSSF.Util.Region(ind + 3, 5, ind + 3, 6));
                }

                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=1"))
                {
                    DecorAssemblyDT.Clear();
                    DT.Clear();
                    int MainOrderID = Convert.ToInt32(item["MainOrderID"]);
                    DataRow[] rows = ArchDecorOrdersDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item1 in rows)
                        DT.Rows.Add(item1.ItemArray);
                    AssemblyDecorCollect(DT, ref DecorAssemblyDT);

                    DataRow[] CRows = MarketOrdersNames.Select("MainOrderID=" + MainOrderID);
                    if (CRows.Count() > 0)
                    {
                        ClientName = CRows[0]["ClientName"].ToString();
                        Notes = CRows[0]["Notes"].ToString();
                        OrderName = "№" + CRows[0]["OrderNumber"].ToString() + "-" + MainOrderID;
                    }
                    if (DecorAssemblyDT.Rows.Count > 0)
                    {
                        ArchDecorAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                             Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DecorAssemblyDT,
                             WorkAssignmentID, BatchName, ClientName, string.Empty, OrderName, Notes, ref RowIndex);
                        RowIndex++;
                        ArchDecorAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                            Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DecorAssemblyDT,
                            WorkAssignmentID, BatchName, ClientName, "ДУБЛЬ", OrderName, Notes, ref RowIndex);
                    }
                    RowIndex++;
                }
            }
        }

        private void AssemblyBagetWithAngleCollect(DataTable SourceDT, ref DataTable DestinationDT)
        {
            DataTable DT1 = new DataTable();

            using (DataView DV = new DataView(SourceDT))
            {
                DT1 = DV.ToTable(true, new string[] { "DecorID", "ColorID", "PatinaID", "Length", "Height", "Width", "LeftAngle", "RightAngle", "Notes" });
            }
            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                int Count = 0;
                string filter = "DecorID=" + Convert.ToInt32(DT1.Rows[i]["DecorID"]) + " AND ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                    " AND Length=" + Convert.ToInt32(DT1.Rows[i]["Length"]) + " AND Height=" + Convert.ToInt32(DT1.Rows[i]["Height"]) +
                    " AND Width=" + Convert.ToInt32(DT1.Rows[i]["Width"]) +
                    " AND LeftAngle=" + Convert.ToInt32(DT1.Rows[i]["LeftAngle"]) +
                    " AND RightAngle=" + Convert.ToInt32(DT1.Rows[i]["RightAngle"]) +
                    " AND (Notes='' OR Notes IS NULL)";
                if (DT1.Rows[i]["Notes"] != DBNull.Value && DT1.Rows[i]["Notes"].ToString().Length > 0)
                {
                    filter = "DecorID=" + Convert.ToInt32(DT1.Rows[i]["DecorID"]) + " AND ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                      " AND Length=" + Convert.ToInt32(DT1.Rows[i]["Length"]) + " AND Height=" + Convert.ToInt32(DT1.Rows[i]["Height"]) +
                      " AND Width=" + Convert.ToInt32(DT1.Rows[i]["Width"]) +
                    " AND LeftAngle=" + Convert.ToInt32(DT1.Rows[i]["LeftAngle"]) +
                    " AND RightAngle=" + Convert.ToInt32(DT1.Rows[i]["RightAngle"]) +
                    " AND Notes='" + DT1.Rows[i]["Notes"] + "'";
                }
                DataRow[] rows = SourceDT.Select(filter);
                if (rows.Count() == 0)
                    continue;

                foreach (DataRow item in rows)
                    Count += Convert.ToInt32(item["Count"]);

                string Color = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                if (Convert.ToInt32(DT1.Rows[i]["PatinaID"]) != -1)
                    Color += " " + GetPatinaName(Convert.ToInt32(DT1.Rows[i]["PatinaID"]));

                DataRow NewRow = DestinationDT.NewRow();
                NewRow["DecorID"] = Convert.ToInt32(DT1.Rows[i]["DecorID"]);
                NewRow["Name"] = GetDecorName(Convert.ToInt32(DT1.Rows[i]["DecorID"]));
                if (HasParameter(Convert.ToInt32(rows[0]["ProductID"]), "ColorID"))
                    NewRow["Color"] = Color;
                //if (HasParameter(Convert.ToInt32(rows[0]["ProductID"]), "Height"))
                //    NewRow["Height"] = DT1.Rows[i]["Height"];
                //if (HasParameter(Convert.ToInt32(rows[0]["ProductID"]), "Length"))
                //    NewRow["Height"] = DT1.Rows[i]["Length"];
                //if (HasParameter(Convert.ToInt32(rows[0]["ProductID"]), "Width"))
                //    NewRow["Width"] = DT1.Rows[i]["Width"];

                if (HasParameter(Convert.ToInt32(rows[0]["ProductID"]), "Height"))
                {
                    if (Convert.ToInt32(DT1.Rows[i]["Height"]) != -1)
                        NewRow["Height"] = DT1.Rows[i]["Height"];
                    if (Convert.ToInt32(DT1.Rows[i]["Length"]) != -1)
                        NewRow["Height"] = DT1.Rows[i]["Length"];
                }
                if (HasParameter(Convert.ToInt32(rows[0]["ProductID"]), "Length"))
                {
                    if (Convert.ToInt32(DT1.Rows[i]["Length"]) != -1)
                        NewRow["Height"] = DT1.Rows[i]["Length"];
                    if (Convert.ToInt32(DT1.Rows[i]["Height"]) != -1)
                        NewRow["Height"] = DT1.Rows[i]["Height"];
                }
                if (HasParameter(Convert.ToInt32(rows[0]["ProductID"]), "Width") && Convert.ToInt32(DT1.Rows[i]["Width"]) != -1)
                    NewRow["Width"] = DT1.Rows[i]["Width"];

                NewRow["Count"] = Count;
                NewRow["LeftAngle"] = DT1.Rows[i]["LeftAngle"];
                NewRow["RightAngle"] = DT1.Rows[i]["RightAngle"];
                NewRow["Notes"] = DT1.Rows[i]["Notes"];
                DestinationDT.Rows.Add(NewRow);
            }
            using (DataView DV = new DataView(DestinationDT.Copy()))
            {
                DV.Sort = "Name, Color, Height, Width, LeftAngle, RightAngle";
                DestinationDT.Clear();
                DestinationDT = DV.ToTable();
            }
        }

        private void AssemblyDecorCollect(DataTable SourceDT, ref DataTable DestinationDT)
        {
            DataTable DT1 = new DataTable();

            using (DataView DV = new DataView(SourceDT))
            {
                DT1 = DV.ToTable(true, new string[] { "DecorID", "ColorID", "PatinaID", "Length", "Height", "Width", "Notes" });
            }
            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                int Count = 0;
                string filter = "DecorID=" + Convert.ToInt32(DT1.Rows[i]["DecorID"]) + " AND ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                    " AND Length=" + Convert.ToInt32(DT1.Rows[i]["Length"]) + " AND Height=" + Convert.ToInt32(DT1.Rows[i]["Height"]) +
                    " AND Width=" + Convert.ToInt32(DT1.Rows[i]["Width"]) + " AND (Notes='' OR Notes IS NULL)";
                if (DT1.Rows[i]["Notes"] != DBNull.Value && DT1.Rows[i]["Notes"].ToString().Length > 0)
                {
                    filter = "DecorID=" + Convert.ToInt32(DT1.Rows[i]["DecorID"]) + " AND ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                      " AND Length=" + Convert.ToInt32(DT1.Rows[i]["Length"]) + " AND Height=" + Convert.ToInt32(DT1.Rows[i]["Height"]) +
                      " AND Width=" + Convert.ToInt32(DT1.Rows[i]["Width"]) + " AND Notes='" + DT1.Rows[i]["Notes"] + "'";
                }
                DataRow[] rows = SourceDT.Select(filter);
                if (rows.Count() == 0)
                    continue;

                foreach (DataRow item in rows)
                    Count += Convert.ToInt32(item["Count"]);

                string Color = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                if (Convert.ToInt32(DT1.Rows[i]["PatinaID"]) != -1)
                    Color += " " + GetPatinaName(Convert.ToInt32(DT1.Rows[i]["PatinaID"]));

                DataRow NewRow = DestinationDT.NewRow();
                NewRow["DecorID"] = Convert.ToInt32(DT1.Rows[i]["DecorID"]);
                NewRow["Name"] = GetDecorName(Convert.ToInt32(DT1.Rows[i]["DecorID"]));
                if (HasParameter(Convert.ToInt32(rows[0]["ProductID"]), "ColorID"))
                    NewRow["Color"] = Color;
                //if (HasParameter(Convert.ToInt32(rows[0]["ProductID"]), "Height"))
                //    NewRow["Height"] = DT1.Rows[i]["Height"];
                //if (HasParameter(Convert.ToInt32(rows[0]["ProductID"]), "Length"))
                //    NewRow["Height"] = DT1.Rows[i]["Length"];
                //if (HasParameter(Convert.ToInt32(rows[0]["ProductID"]), "Width"))
                //    NewRow["Width"] = DT1.Rows[i]["Width"];

                if (HasParameter(Convert.ToInt32(rows[0]["ProductID"]), "Height"))
                {
                    if (Convert.ToInt32(DT1.Rows[i]["Height"]) != -1)
                        NewRow["Height"] = DT1.Rows[i]["Height"];
                    if (Convert.ToInt32(DT1.Rows[i]["Length"]) != -1)
                        NewRow["Height"] = DT1.Rows[i]["Length"];
                }
                if (HasParameter(Convert.ToInt32(rows[0]["ProductID"]), "Length"))
                {
                    if (Convert.ToInt32(DT1.Rows[i]["Length"]) != -1)
                        NewRow["Height"] = DT1.Rows[i]["Length"];
                    if (Convert.ToInt32(DT1.Rows[i]["Height"]) != -1)
                        NewRow["Height"] = DT1.Rows[i]["Height"];
                }
                if (HasParameter(Convert.ToInt32(rows[0]["ProductID"]), "Width") && Convert.ToInt32(DT1.Rows[i]["Width"]) != -1)
                    NewRow["Width"] = DT1.Rows[i]["Width"];

                NewRow["Count"] = Count;
                NewRow["Notes"] = DT1.Rows[i]["Notes"];
                DestinationDT.Rows.Add(NewRow);
            }
            using (DataView DV = new DataView(DestinationDT.Copy()))
            {
                DV.Sort = "Name, Color, Height, Width";
                DestinationDT.Clear();
                DestinationDT = DV.ToTable();
            }
        }

        private void AssemblySingly(int ColorID, DataTable SourceDT, ref DataTable DestinationDT)
        {
            DataTable DT2 = new DataTable();

            using (DataView DV = new DataView(SourceDT, "ColorID=" + ColorID, "ColorID, TechnoColorID, InsetTypeID, InsetColorID, TechnoInsetTypeID, TechnoInsetColorID, Height, Width", DataViewRowState.CurrentRows))
            {
                DT2 = DV.ToTable(true, new string[] { "ColorID", "TechnoColorID", "InsetTypeID", "InsetColorID", "TechnoInsetTypeID", "TechnoInsetColorID", "Height", "Width" });
            }
            for (int j = 0; j < DT2.Rows.Count; j++)
            {
                decimal Square = 0;
                int Count = 0;
                string Color = string.Empty;
                string InsetColor = string.Empty;
                //витрины
                DataRow[] rows = SourceDT.Select("ColorID=" + ColorID + " AND TechnoColorID=" + Convert.ToInt32(DT2.Rows[j]["TechnoColorID"]) + " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) +
                    " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND TechnoInsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["TechnoInsetTypeID"]) +
                    " AND TechnoInsetColorID=" + Convert.ToInt32(DT2.Rows[j]["TechnoInsetColorID"]) + " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                    " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]));
                foreach (DataRow item in rows)
                {
                    Count += Convert.ToInt32(item["Count"]);
                    Square += Convert.ToDecimal(item["Height"]) * Convert.ToDecimal(item["Width"]) * Convert.ToDecimal(item["Count"]) / 1000000;
                }
                Square = decimal.Round(Square, 3, MidpointRounding.AwayFromZero);
                Color = GetInsetColorName(Convert.ToInt32(rows[0]["InsetColorID"]));
                InsetColor = GetInsetTypeName(Convert.ToInt32(rows[0]["InsetTypeID"])) + " " + Color;

                if (Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) == 2 && Convert.ToInt32(DT2.Rows[j]["TechnoInsetTypeID"]) != -1)
                    InsetColor = GetInsetTypeName(Convert.ToInt32(DT2.Rows[j]["TechnoInsetTypeID"])) + " " + GetInsetColorName(Convert.ToInt32(DT2.Rows[j]["InsetColorID"]));
                //int GroupID = Convert.ToInt32(ItemInsetTypesDataTable.Select("InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]))[0]["GroupID"]);
                //if (GroupID == 7 || GroupID == 8)
                //    InsetColor = "фил " + GetInsetTypeName(Convert.ToInt32(rows[0]["InsetTypeID"])) + " " + Color;
                //if (GroupID == 16 || GroupID == 17)
                //    InsetColor = GetInsetTypeName(Convert.ToInt32(rows[0]["InsetTypeID"])) + " 45 " + Color;
                //if (GroupID == 18 || GroupID == 19)
                //    InsetColor = GetInsetTypeName(Convert.ToInt32(rows[0]["InsetTypeID"])) + " 90 " + Color;
                //if (GroupID == 12)
                //    InsetColor = GetInsetTypeName(Convert.ToInt32(rows[0]["InsetTypeID"])) + " " + Color;
                if (Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) == 1)
                    InsetColor = "Витрина";

                DataRow NewRow = DestinationDT.NewRow();
                NewRow["ColorType"] = ColorID;
                string FrontName = GetFrontName(Convert.ToInt32(rows[0]["FrontID"]));
                if (Convert.ToInt32(rows[0]["TechnoColorID"]) != -1)
                    FrontName += " ИМП.";
                if (Convert.ToInt32(rows[0]["TechnoColorID"]) == -1)
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"]));
                if (Convert.ToInt32(rows[0]["TechnoColorID"]) != -1)
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"])) + "/" + GetColorName(Convert.ToInt32(rows[0]["TechnoColorID"]));
                //if (Convert.ToInt32(rows[0]["TechnoColorID"]) != -1 && Convert.ToInt32(rows[0]["TechnoColorID"]) != Convert.ToInt32(rows[0]["ColorID"]))
                //    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"])) + "/" + GetColorName(Convert.ToInt32(rows[0]["TechnoColorID"]));
                NewRow["Name"] = FrontName;
                NewRow["InsetColor"] = InsetColor;
                NewRow["Height"] = rows[0]["Height"];
                NewRow["Width"] = rows[0]["Width"];
                NewRow["Count"] = Count;
                NewRow["Square"] = Square;
                DestinationDT.Rows.Add(NewRow);
            }
        }

        private void BagetWithAngleAssemblyByMainOrderToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            int WorkAssignmentID, string BatchName)
        {
            if (BagetWithAngelOrdersDT.Rows.Count == 0)
                return;

            DataTable DistMainOrdersDT = DistMainOrdersTable(BagetWithAngelOrdersDT, true);
            DataTable DT = BagetWithAngelOrdersDT.Clone();
            DataTable ZOVOrdersNames = new DataTable();
            DataTable MarketOrdersNames = new DataTable();
            string MainOrdersID = string.Empty;
            string ClientName = string.Empty;
            string OrderName = string.Empty;
            string Notes = string.Empty;

            if (DistMainOrdersDT.Select("GroupType=0").Count() > 0)
            {
                MainOrdersID = string.Empty;
                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=0"))
                    MainOrdersID += Convert.ToInt32(item["MainOrderID"]) + ",";
                MainOrdersID = MainOrdersID.Substring(0, MainOrdersID.Length - 1);
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT ClientName, DocNumber, MainOrderID FROM MainOrders" +
                    " INNER JOIN infiniu2_zovreference.dbo.Clients ON MainOrders.ClientID=infiniu2_zovreference.dbo.Clients.ClientID" +
                    " WHERE MainOrderID IN (" + MainOrdersID + ")", ConnectionStrings.ZOVOrdersConnectionString))
                {
                    DA.Fill(ZOVOrdersNames);
                }

                int RowIndex = 0;
                HSSFSheet sheet1 = hssfworkbook.CreateSheet("Не арки ЗОВ");
                sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

                sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
                sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
                sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
                sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

                sheet1.SetColumnWidth(0, 30 * 256);
                sheet1.SetColumnWidth(1, 20 * 256);
                sheet1.SetColumnWidth(2, 9 * 256);
                sheet1.SetColumnWidth(3, 6 * 256);
                sheet1.SetColumnWidth(4, 6 * 256);
                sheet1.SetColumnWidth(5, 6 * 256);
                sheet1.SetColumnWidth(6, 6 * 256);

                {
                    HSSFPatriarch patriarch = sheet1.CreateDrawingPatriarch();
                    int ind = 0;

                    CreateBarcode(WorkAssignmentID);
                    HSSFClientAnchor anchor = new HSSFClientAnchor(0, 1, 1000, 8, 5, ind, 6, ind + 3)
                    {
                        AnchorType = 2
                    };
                    string BarcodeNumber = GetBarcodeNumber(24, WorkAssignmentID);
                    string barcodeName = System.Environment.GetEnvironmentVariable("TEMP") + @"/" + BarcodeNumber + ".jpg";
                    HSSFPicture picture = patriarch.CreatePicture(anchor, LoadImage(barcodeName, hssfworkbook));

                    HSSFFont BarcodeFont = hssfworkbook.CreateFont();
                    BarcodeFont.FontHeightInPoints = 12;
                    BarcodeFont.FontName = "Calibri";

                    HSSFCellStyle BarcodeStyle = hssfworkbook.CreateCellStyle();
                    BarcodeStyle.Alignment = HSSFCellStyle.ALIGN_CENTER;
                    BarcodeStyle.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
                    BarcodeStyle.SetFont(BarcodeFont);

                    HSSFCell barcodeCell = sheet1.CreateRow(ind + 3).CreateCell(5);
                    barcodeCell.SetCellValue(BarcodeNumber.ToString());
                    barcodeCell.CellStyle = BarcodeStyle;
                    sheet1.AddMergedRegion(new NPOI.HSSF.Util.Region(ind + 3, 5, ind + 3, 6));
                }

                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=0"))
                {
                    BagetWithAngleAssemblyDT.Clear();
                    DT.Clear();
                    int MainOrderID = Convert.ToInt32(item["MainOrderID"]);
                    DataRow[] rows = BagetWithAngelOrdersDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item1 in rows)
                        DT.Rows.Add(item1.ItemArray);
                    AssemblyBagetWithAngleCollect(DT, ref BagetWithAngleAssemblyDT);

                    DataRow[] CRows = ZOVOrdersNames.Select("MainOrderID=" + Convert.ToInt32(item["MainOrderID"]));
                    if (CRows.Count() > 0)
                    {
                        ClientName = CRows[0]["ClientName"].ToString();
                        OrderName = CRows[0]["DocNumber"].ToString();
                    }
                    if (BagetWithAngleAssemblyDT.Rows.Count > 0)
                    {
                        BagetWithAngleAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                             Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, BagetWithAngleAssemblyDT,
                             WorkAssignmentID, BatchName, ClientName, string.Empty, OrderName, Notes, ref RowIndex);
                        RowIndex++;
                        BagetWithAngleAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                            Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, BagetWithAngleAssemblyDT,
                            WorkAssignmentID, BatchName, ClientName, "ДУБЛЬ", OrderName, Notes, ref RowIndex);
                    }
                    RowIndex++;
                }
            }
            if (DistMainOrdersDT.Select("GroupType=1").Count() > 0)
            {
                MainOrdersID = string.Empty;
                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=1"))
                    MainOrdersID += Convert.ToInt32(item["MainOrderID"]) + ",";
                MainOrdersID = MainOrdersID.Substring(0, MainOrdersID.Length - 1);
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT ClientName, OrderNumber, MainOrderID, Notes FROM MainOrders" +
                    " INNER JOIN MegaOrders ON MainOrders.MegaOrderID=MegaOrders.MegaOrderID" +
                    " INNER JOIN infiniu2_marketingreference.dbo.Clients ON MegaOrders.ClientID=infiniu2_marketingreference.dbo.Clients.ClientID" +
                    " WHERE MainOrderID IN (" + MainOrdersID + ")", ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DA.Fill(MarketOrdersNames);
                }

                int RowIndex = 0;
                HSSFSheet sheet1 = hssfworkbook.CreateSheet("Багет с запилом Маркетинг");
                sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

                sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
                sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
                sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
                sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

                sheet1.SetColumnWidth(0, 30 * 256);
                sheet1.SetColumnWidth(1, 20 * 256);
                sheet1.SetColumnWidth(2, 9 * 256);
                sheet1.SetColumnWidth(3, 6 * 256);
                sheet1.SetColumnWidth(4, 6 * 256);
                sheet1.SetColumnWidth(5, 6 * 256);
                sheet1.SetColumnWidth(6, 6 * 256);

                {
                    HSSFPatriarch patriarch = sheet1.CreateDrawingPatriarch();
                    int ind = 0;

                    CreateBarcode(WorkAssignmentID);
                    HSSFClientAnchor anchor = new HSSFClientAnchor(0, 1, 1000, 8, 5, ind, 6, ind + 3)
                    {
                        AnchorType = 2
                    };
                    string BarcodeNumber = GetBarcodeNumber(24, WorkAssignmentID);
                    string barcodeName = System.Environment.GetEnvironmentVariable("TEMP") + @"/" + BarcodeNumber + ".jpg";
                    HSSFPicture picture = patriarch.CreatePicture(anchor, LoadImage(barcodeName, hssfworkbook));

                    HSSFFont BarcodeFont = hssfworkbook.CreateFont();
                    BarcodeFont.FontHeightInPoints = 12;
                    BarcodeFont.FontName = "Calibri";

                    HSSFCellStyle BarcodeStyle = hssfworkbook.CreateCellStyle();
                    BarcodeStyle.Alignment = HSSFCellStyle.ALIGN_CENTER;
                    BarcodeStyle.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
                    BarcodeStyle.SetFont(BarcodeFont);

                    HSSFCell barcodeCell = sheet1.CreateRow(ind + 3).CreateCell(5);
                    barcodeCell.SetCellValue(BarcodeNumber.ToString());
                    barcodeCell.CellStyle = BarcodeStyle;
                    sheet1.AddMergedRegion(new NPOI.HSSF.Util.Region(ind + 3, 5, ind + 3, 6));
                }

                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=1"))
                {
                    BagetWithAngleAssemblyDT.Clear();
                    DT.Clear();
                    int MainOrderID = Convert.ToInt32(item["MainOrderID"]);
                    DataRow[] rows = BagetWithAngelOrdersDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item1 in rows)
                        DT.Rows.Add(item1.ItemArray);
                    AssemblyBagetWithAngleCollect(DT, ref BagetWithAngleAssemblyDT);

                    DataRow[] CRows = MarketOrdersNames.Select("MainOrderID=" + MainOrderID);
                    if (CRows.Count() > 0)
                    {
                        ClientName = CRows[0]["ClientName"].ToString();
                        Notes = CRows[0]["Notes"].ToString();
                        OrderName = "№" + CRows[0]["OrderNumber"].ToString() + "-" + MainOrderID;
                    }
                    if (BagetWithAngleAssemblyDT.Rows.Count > 0)
                    {
                        BagetWithAngleAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                             Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, BagetWithAngleAssemblyDT,
                             WorkAssignmentID, BatchName, ClientName, string.Empty, OrderName, Notes, ref RowIndex);
                        RowIndex++;
                        BagetWithAngleAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                            Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, BagetWithAngleAssemblyDT,
                            WorkAssignmentID, BatchName, ClientName, "ДУБЛЬ", OrderName, Notes, ref RowIndex);
                    }
                    RowIndex++;
                }
            }
        }

        private void CollectAdditions0(ref DataTable DestinationDT, int Margin)
        {
            if (BostonGlassDT.Rows.Count > 0)
                AdditionsSyngly(BostonGlassDT, ref DestinationDT, "Стекло 4 мм", "-", Margin);

            for (int i = 1; i < DestinationDT.Rows.Count; i++)
            {
                if (DestinationDT.Rows[i]["Profile"].ToString() == DestinationDT.Rows[i - 1]["Profile"].ToString() &&
                    DestinationDT.Rows[i]["InsetColor"].ToString() == DestinationDT.Rows[i - 1]["InsetColor"].ToString())
                {
                    DestinationDT.Rows[i]["Profile"] = string.Empty;
                    DestinationDT.Rows[i]["Color"] = string.Empty;
                    DestinationDT.Rows[i]["InsetColor"] = string.Empty;
                }
            }
        }

        private void CollectAdditions1(ref DataTable DestinationDT, int Margin)
        {
            if (BergamoGlassDT.Rows.Count > 0)
                AdditionsSyngly(BergamoGlassDT, ref DestinationDT, "Стекло 4 мм", "-", Margin);

            for (int i = 1; i < DestinationDT.Rows.Count; i++)
            {
                if (DestinationDT.Rows[i]["Profile"].ToString() == DestinationDT.Rows[i - 1]["Profile"].ToString() &&
                    DestinationDT.Rows[i]["InsetColor"].ToString() == DestinationDT.Rows[i - 1]["InsetColor"].ToString())
                {
                    DestinationDT.Rows[i]["Profile"] = string.Empty;
                    DestinationDT.Rows[i]["Color"] = string.Empty;
                    DestinationDT.Rows[i]["InsetColor"] = string.Empty;
                }
            }
        }

        private void CollectAdditions2(ref DataTable DestinationDT, int Margin)
        {
            if (BergamoGlassDT.Rows.Count > 0)
                AdditionsSyngly(BergamoGlassDT, ref DestinationDT, "МДФ 3 мм", "-", Margin);

            for (int i = 1; i < DestinationDT.Rows.Count; i++)
            {
                if (DestinationDT.Rows[i]["Profile"].ToString() == DestinationDT.Rows[i - 1]["Profile"].ToString() &&
                    DestinationDT.Rows[i]["InsetColor"].ToString() == DestinationDT.Rows[i - 1]["InsetColor"].ToString())
                {
                    DestinationDT.Rows[i]["Profile"] = string.Empty;
                    DestinationDT.Rows[i]["Color"] = string.Empty;
                    DestinationDT.Rows[i]["InsetColor"] = string.Empty;
                }
            }
        }

        private void CollectAdditions3(ref DataTable DestinationDT, int Margin)
        {
            if (BergamoGlassDT.Rows.Count > 0)
                AdditionsSyngly(BergamoGlassDT, ref DestinationDT, "МДФ 3 мм", "ПП Дубок", Margin);

            for (int i = 1; i < DestinationDT.Rows.Count; i++)
            {
                if (DestinationDT.Rows[i]["Profile"].ToString() == DestinationDT.Rows[i - 1]["Profile"].ToString() &&
                    DestinationDT.Rows[i]["InsetColor"].ToString() == DestinationDT.Rows[i - 1]["InsetColor"].ToString())
                {
                    DestinationDT.Rows[i]["Profile"] = string.Empty;
                    DestinationDT.Rows[i]["Color"] = string.Empty;
                    DestinationDT.Rows[i]["InsetColor"] = string.Empty;
                }
            }
        }

        private void CollectAdditions4(ref DataTable DestinationDT, int Margin)
        {
            if (BergamoGlassDT.Rows.Count > 0)
                AdditionsSyngly(BergamoGlassDT, ref DestinationDT, "ПП03-404", "ПП Дубок", Margin);

            for (int i = 1; i < DestinationDT.Rows.Count; i++)
            {
                if (DestinationDT.Rows[i]["Profile"].ToString() == DestinationDT.Rows[i - 1]["Profile"].ToString() &&
                    DestinationDT.Rows[i]["InsetColor"].ToString() == DestinationDT.Rows[i - 1]["InsetColor"].ToString())
                {
                    DestinationDT.Rows[i]["Profile"] = string.Empty;
                    DestinationDT.Rows[i]["Color"] = string.Empty;
                    DestinationDT.Rows[i]["InsetColor"] = string.Empty;
                }
            }
        }

        private void CollectAdditions5(ref DataTable DestinationDT, int Margin)
        {
            if (BergamoGlassDT.Rows.Count > 0)
                AdditionsSyngly(BergamoGlassDT, ref DestinationDT, "ПП03-404", "Сахара", Margin);

            for (int i = 1; i < DestinationDT.Rows.Count; i++)
            {
                if (DestinationDT.Rows[i]["Profile"].ToString() == DestinationDT.Rows[i - 1]["Profile"].ToString() &&
                    DestinationDT.Rows[i]["InsetColor"].ToString() == DestinationDT.Rows[i - 1]["InsetColor"].ToString())
                {
                    DestinationDT.Rows[i]["Profile"] = string.Empty;
                    DestinationDT.Rows[i]["Color"] = string.Empty;
                    DestinationDT.Rows[i]["InsetColor"] = string.Empty;
                }
            }
        }

        private void CollectAllInsets(ref DataTable DestinationDT)
        {
            if (TechnoNSimpleDT.Rows.Count > 0)
                AllInsets(TechnoNSimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.TechnoNMargin), true);
            if (AntaliaSimpleDT.Rows.Count > 0)
                AllInsets(AntaliaSimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.AntaliaMargin), true);
            if (Nord95SimpleDT.Rows.Count > 0)
                AllInsets(Nord95SimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.Nord95Margin), true);
            if (epFoxSimpleDT.Rows.Count > 0)
                AllInsets(epFoxSimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.epFoxMargin), true);
            //if (BergamoSimpleDT.Rows.Count > 0)
            //    AllInsets(BergamoSimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.BergamoMargin), true);
            if (BostonSimpleDT.Rows.Count > 0)
                AllInsets(BostonSimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.BostonMargin), true);

            DataTable DT = BergamoSimpleDT.Clone();

            DataRow[] rows = BergamoSimpleDT.Select("(Height>150 OR Width>150)");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            AllInsets(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.BergamoMargin), true);

            DT.Clear();
            rows = BergamoSimpleDT.Select("(Height<=150 OR Width<=150)");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            AllInsets(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.BergamoNarrowMargin), true);

            if (ep041SimpleDT.Rows.Count > 0)
                AllInsets(ep041SimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep041Margin), true);

            if (ep071SimpleDT.Rows.Count > 0)
                AllInsets(ep071SimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep071Margin), true);

            if (ep206SimpleDT.Rows.Count > 0)
                AllInsets(ep206SimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep206Margin), true);

            if (ep216SimpleDT.Rows.Count > 0)
                AllInsets(ep216SimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep216Margin), true);

            if (ep111SimpleDT.Rows.Count > 0)
                AllInsets(ep111SimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep111Margin), true);

            if (VeneciaSimpleDT.Rows.Count > 0)
            {
                DataTable TempDT = new DataTable();
                TempDT.Clear();
                using (DataView DV = new DataView(VeneciaOrdersDT))
                {
                    TempDT = DV.ToTable(true, new string[] { "ProfileID" });
                }
                for (int i = 0; i < TempDT.Rows.Count; i++)
                {
                    int ProfileID = Convert.ToInt32(TempDT.Rows[i]["ProfileID"]);
                    DataTable TempSimpleDT = VeneciaSimpleDT.Clone();
                    rows = VeneciaSimpleDT.Select("ProfileID=" + ProfileID);
                    if (rows.Count() > 0)
                    {
                        foreach (DataRow item in rows)
                            TempSimpleDT.Rows.Add(item.ItemArray);
                        AllInsets(TempSimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.VeneciaMargin), true);
                    }
                    TempSimpleDT.Dispose();
                }
            }
            if (LeonSimpleDT.Rows.Count > 0)
                AllInsets(LeonSimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.LeonMargin), true);
            if (LimogSimpleDT.Rows.Count > 0)
                AllInsets(LimogSimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.LimogMargin), true);
            if (ep066Marsel4SimpleDT.Rows.Count > 0)
                AllInsets(ep066Marsel4SimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep066Marsel4Margin), true);
            if (ep110JersySimpleDT.Rows.Count > 0)
                AllInsets(ep110JersySimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep110JersyMargin), true);
            if (ep018Marsel1SimpleDT.Rows.Count > 0)
                AllInsets(ep018Marsel1SimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep018Marsel1Margin), true);
            if (ep043ShervudSimpleDT.Rows.Count > 0)
                AllInsets(ep043ShervudSimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep043ShervudMargin), true);
            if (ep112SimpleDT.Rows.Count > 0)
                AllInsets(ep112SimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep112Margin), true);
            if (UrbanSimpleDT.Rows.Count > 0)
                AllInsets(UrbanSimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.UrbanMargin), true);
            if (AlbySimpleDT.Rows.Count > 0)
                AllInsets(AlbySimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.AlbyMargin), true);
            if (BrunoSimpleDT.Rows.Count > 0)
                AllInsets(BrunoSimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.BrunoMargin), true);
            if (epsh406Techno4SimpleDT.Rows.Count > 0)
                AllInsets(epsh406Techno4SimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.epsh406Techno4Margin), true);
            if (LukSimpleDT.Rows.Count > 0)
                AllInsets(LukSimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.LukMargin), true);
            if (LukPVHSimpleDT.Rows.Count > 0)
                AllInsets(LukPVHSimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.LukMargin), true);
            if (MilanoSimpleDT.Rows.Count > 0)
                AllInsets(MilanoSimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.MilanoMargin), true);
            if (PragaSimpleDT.Rows.Count > 0)
                AllInsets(PragaSimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.PragaMargin), true);
            if (SigmaSimpleDT.Rows.Count > 0)
                AllInsets(SigmaSimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.SigmaMargin), true);
            if (FatSimpleDT.Rows.Count > 0)
                AllInsets(FatSimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.FatMargin), true);

            using (DataView DV = new DataView(DestinationDT.Copy()))
            {
                DV.Sort = "Name, Height, Width";
                DestinationDT.Clear();
                DestinationDT = DV.ToTable();
            }

            if (DestinationDT.Rows.Count == 0)
                return;

            for (int i = 1; i < DestinationDT.Rows.Count; i++)
            {
                if (Convert.ToInt32(DestinationDT.Rows[i]["InsetColorID"]) == Convert.ToInt32(DestinationDT.Rows[i - 1]["InsetColorID"]))
                    DestinationDT.Rows[i]["Name"] = string.Empty;
            }
        }

        private void CollectDeying(int ColorID, DataTable SourceDT, ref DataTable DestinationDT, string AdditionalName)
        {
            DataTable DT2 = new DataTable();
            //Витрины сначала
            using (DataView DV = new DataView(SourceDT, "InsetTypeID=1 AND ColorID=" + ColorID,
                "InsetTypeID, InsetColorID, Height, Width", DataViewRowState.CurrentRows))
            {
                DT2 = DV.ToTable(true, new string[] { "PatinaID", "InsetTypeID", "InsetColorID", "Height", "Width", "Notes" });
            }
            for (int j = 0; j < DT2.Rows.Count; j++)
            {
                decimal Square = 0;
                int Count = 0;
                string InsetColor = string.Empty;
                string filter1 = "ColorID=" + ColorID + " AND PatinaID=" + Convert.ToInt32(DT2.Rows[j]["PatinaID"]) + " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) +
                            " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                            " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) + " AND (Notes='' OR Notes IS NULL)";

                if (DT2.Rows[j]["Notes"] != DBNull.Value && DT2.Rows[j]["Notes"].ToString().Length > 0)
                    filter1 = "ColorID=" + ColorID + " AND PatinaID=" + Convert.ToInt32(DT2.Rows[j]["PatinaID"]) + " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) +
                    " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                    " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) + " AND Notes='" + DT2.Rows[j]["Notes"] + "'";

                DataRow[] rows = SourceDT.Select(filter1);
                foreach (DataRow item in rows)
                {
                    Count += Convert.ToInt32(item["Count"]);
                    Square += Convert.ToDecimal(item["Height"]) * Convert.ToDecimal(item["Width"]) * Convert.ToDecimal(item["Count"]) / 1000000;
                }
                Square = decimal.Round(Square, 3, MidpointRounding.AwayFromZero);

                DataRow NewRow = DestinationDT.NewRow();
                NewRow["ColorType"] = ColorID;
                NewRow["Name"] = GetFrontName(Convert.ToInt32(rows[0]["FrontID"])) + AdditionalName;
                if (Convert.ToInt32(rows[0]["PatinaID"]) == -1)
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"]));
                else
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"])) + "+" + GetPatinaName(Convert.ToInt32(rows[0]["PatinaID"]));
                if (Convert.ToInt32(rows[0]["InsetTypeID"]) == 1)
                    NewRow["InsetColor"] = "Витрина";
                else
                    NewRow["InsetColor"] = GetInsetTypeName(Convert.ToInt32(rows[0]["InsetTypeID"])) + " " + GetInsetColorName(Convert.ToInt32(rows[0]["InsetColorID"]));
                NewRow["Height"] = rows[0]["Height"];
                NewRow["Width"] = rows[0]["Width"];
                NewRow["Count"] = Count;
                NewRow["Square"] = Square;
                NewRow["Notes"] = rows[0]["Notes"];
                DestinationDT.Rows.Add(NewRow);
            }

            using (DataView DV = new DataView(SourceDT, "InsetTypeID<>1 AND ColorID=" + ColorID,
                "InsetTypeID, InsetColorID, Height, Width", DataViewRowState.CurrentRows))
            {
                DT2 = DV.ToTable(true, new string[] { "PatinaID", "InsetTypeID", "InsetColorID", "Height", "Width", "Notes" });
            }
            for (int j = 0; j < DT2.Rows.Count; j++)
            {
                decimal Square = 0;
                int Count = 0;
                string InsetColor = string.Empty;
                string filter1 = "ColorID=" + ColorID + " AND PatinaID=" + Convert.ToInt32(DT2.Rows[j]["PatinaID"]) + " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) +
                            " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                            " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) + " AND (Notes='' OR Notes IS NULL)";

                if (DT2.Rows[j]["Notes"] != DBNull.Value && DT2.Rows[j]["Notes"].ToString().Length > 0)
                    filter1 = "ColorID=" + ColorID + " AND PatinaID=" + Convert.ToInt32(DT2.Rows[j]["PatinaID"]) + " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) +
                    " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                    " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) + " AND Notes='" + DT2.Rows[j]["Notes"] + "'";

                DataRow[] rows = SourceDT.Select(filter1);
                foreach (DataRow item in rows)
                {
                    Count += Convert.ToInt32(item["Count"]);
                    Square += Convert.ToDecimal(item["Height"]) * Convert.ToDecimal(item["Width"]) * Convert.ToDecimal(item["Count"]) / 1000000;
                }
                Square = decimal.Round(Square, 3, MidpointRounding.AwayFromZero);

                DataRow NewRow = DestinationDT.NewRow();
                NewRow["ColorType"] = ColorID;
                NewRow["Name"] = GetFrontName(Convert.ToInt32(rows[0]["FrontID"])) + AdditionalName;
                if (Convert.ToInt32(rows[0]["PatinaID"]) == -1)
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"]));
                else
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"])) + "+" + GetPatinaName(Convert.ToInt32(rows[0]["PatinaID"]));
                if (Convert.ToInt32(rows[0]["InsetTypeID"]) == 1)
                    NewRow["InsetColor"] = "Витрина";
                else
                    NewRow["InsetColor"] = GetInsetTypeName(Convert.ToInt32(rows[0]["InsetTypeID"])) + " " + GetInsetColorName(Convert.ToInt32(rows[0]["InsetColorID"]));
                NewRow["Height"] = rows[0]["Height"];
                NewRow["Width"] = rows[0]["Width"];
                NewRow["Count"] = Count;
                NewRow["Square"] = Square;
                NewRow["Notes"] = rows[0]["Notes"];
                DestinationDT.Rows.Add(NewRow);
            }
        }

        private void CollectFilenkaOnly(ref DataTable DestinationDT)
        {
            if (AntaliaSimpleDT.Rows.Count > 0)
                InsetsFilenkaOnly(AntaliaSimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.AntaliaMargin), true, string.Empty);
            if (Nord95SimpleDT.Rows.Count > 0)
                InsetsFilenkaOnly(Nord95SimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.Nord95Margin), true, string.Empty);
            if (epFoxSimpleDT.Rows.Count > 0)
                InsetsFilenkaOnly(epFoxSimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.epFoxMargin), true, string.Empty);
            //if (BergamoSimpleDT.Rows.Count > 0)
            //    InsetsFilenkaOnly(BergamoSimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.BergamoMargin), true, string.Empty);

            DataTable DT = BergamoSimpleDT.Clone();

            DataRow[] rows = BergamoSimpleDT.Select("(Height>150 OR Width>150)");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            InsetsFilenkaOnly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.BergamoMargin), true, string.Empty);

            DT.Clear();
            rows = BergamoSimpleDT.Select("(Height<=150 OR Width<=150)");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            InsetsFilenkaOnly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.BergamoNarrowMargin), true, string.Empty);

            if (ep041SimpleDT.Rows.Count > 0)
                InsetsFilenkaOnly(ep041SimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep041Margin), true, string.Empty);

            if (ep071SimpleDT.Rows.Count > 0)
                InsetsFilenkaOnly(ep071SimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep071Margin), true, string.Empty);

            if (ep206SimpleDT.Rows.Count > 0)
                InsetsFilenkaOnly(ep206SimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep206Margin), true, string.Empty);

            if (ep216SimpleDT.Rows.Count > 0)
                InsetsFilenkaOnly(ep216SimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep216Margin), true, string.Empty);

            if (ep111SimpleDT.Rows.Count > 0)
                InsetsFilenkaOnly(ep111SimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep111Margin), true, string.Empty);

            if (BostonSimpleDT.Rows.Count > 0)
                InsetsFilenkaOnly(BostonSimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.BostonMargin), true, string.Empty);
            if (VeneciaSimpleDT.Rows.Count > 0)
                InsetsFilenkaOnly(VeneciaSimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.VeneciaMargin), true, " БРВ 10");
            if (LeonSimpleDT.Rows.Count > 0)
                InsetsFilenkaOnly(LeonSimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.LeonMargin), true, " Леон/де Гроб 8");
            if (LimogSimpleDT.Rows.Count > 0)
                InsetsFilenkaOnly(LimogSimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.LimogMargin), true, string.Empty);
            if (ep066Marsel4SimpleDT.Rows.Count > 0)
                InsetsFilenkaOnly(ep066Marsel4SimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep066Marsel4Margin), true, string.Empty);
            if (ep110JersySimpleDT.Rows.Count > 0)
                InsetsFilenkaOnly(ep110JersySimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep110JersyMargin), true, string.Empty);
            if (ep018Marsel1SimpleDT.Rows.Count > 0)
                InsetsFilenkaOnly(ep018Marsel1SimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep018Marsel1Margin), true, string.Empty);
            if (ep043ShervudSimpleDT.Rows.Count > 0)
                InsetsFilenkaOnly(ep043ShervudSimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep043ShervudMargin), true, string.Empty);
            if (ep112SimpleDT.Rows.Count > 0)
                InsetsFilenkaOnly(ep112SimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep112Margin), true, string.Empty);
            if (UrbanSimpleDT.Rows.Count > 0)
                InsetsFilenkaOnly(UrbanSimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.UrbanMargin), true, string.Empty);
            if (AlbySimpleDT.Rows.Count > 0)
                InsetsFilenkaOnly(AlbySimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.AlbyMargin), true, string.Empty);
            if (BrunoSimpleDT.Rows.Count > 0)
                InsetsFilenkaOnly(BrunoSimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.BrunoMargin), true, string.Empty);
            if (epsh406Techno4SimpleDT.Rows.Count > 0)
                InsetsFilenkaOnly(epsh406Techno4SimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.epsh406Techno4Margin), true, string.Empty);
            if (LukSimpleDT.Rows.Count > 0)
                InsetsFilenkaOnly(LukSimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.LukMargin), true, string.Empty);
            if (LukPVHSimpleDT.Rows.Count > 0)
                InsetsFilenkaOnly(LukPVHSimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.LukMargin), true, string.Empty);
            if (MilanoSimpleDT.Rows.Count > 0)
                InsetsFilenkaOnly(MilanoSimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.MilanoMargin), true, " Милано");
            if (PragaSimpleDT.Rows.Count > 0)
                InsetsFilenkaOnly(PragaSimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.PragaMargin), true, string.Empty);
            if (SigmaSimpleDT.Rows.Count > 0)
                InsetsFilenkaOnly(SigmaSimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.SigmaMargin), true, string.Empty);
            if (FatSimpleDT.Rows.Count > 0)
                InsetsFilenkaOnly(FatSimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.FatMargin), true, string.Empty);

            if (DestinationDT.Rows.Count == 0)
                return;

            for (int i = 1; i < DestinationDT.Rows.Count; i++)
            {
                if (Convert.ToInt32(DestinationDT.Rows[i]["FrontID"]) == Convert.ToInt32(DestinationDT.Rows[i - 1]["FrontID"]) &&
                    Convert.ToInt32(DestinationDT.Rows[i]["InsetColorID"]) == Convert.ToInt32(DestinationDT.Rows[i - 1]["InsetColorID"]) &&
                    Convert.ToInt32(DestinationDT.Rows[i]["TechnoInsetColorID"]) == Convert.ToInt32(DestinationDT.Rows[i - 1]["TechnoInsetColorID"]))
                    DestinationDT.Rows[i]["Name"] = string.Empty;
            }
        }

        private void CollectInsetsGridsOnly(ref DataTable DestinationDT)
        {
            if (AntaliaGridsDT.Rows.Count > 0)
                InsetsGridsOnly(AntaliaGridsDT, ref DestinationDT, Convert.ToInt32(FrontMargins.AntaliaMargin), true);
            if (Nord95GridsDT.Rows.Count > 0)
                InsetsGridsOnly(Nord95GridsDT, ref DestinationDT, Convert.ToInt32(FrontMargins.Nord95Margin), true);
            if (epFoxGridsDT.Rows.Count > 0)
                InsetsGridsOnly(epFoxGridsDT, ref DestinationDT, Convert.ToInt32(FrontMargins.epFoxMargin), true);
            if (VeneciaGridsDT.Rows.Count > 0)
                InsetsGridsOnly(VeneciaGridsDT, ref DestinationDT, Convert.ToInt32(FrontMargins.VeneciaMargin), true);
            //if (BergamoGridsDT.Rows.Count > 0)
            //    InsetsGridsOnly(BergamoGridsDT, ref DestinationDT, Convert.ToInt32(FrontMargins.BergamoMargin), true);
            DataTable DT = BergamoGridsDT.Clone();

            DataRow[] rows = BergamoGridsDT.Select("(Height>150 OR Width>150)");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            InsetsGridsOnly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.BergamoMargin), true);

            DT.Clear();
            rows = BergamoGridsDT.Select("(Height<=150 OR Width<=150)");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            InsetsGridsOnly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.BergamoNarrowMargin), true);

            if (ep041GridsDT.Rows.Count > 0)
                InsetsGridsOnly(ep041GridsDT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep041Margin), true);
            if (ep071GridsDT.Rows.Count > 0)
                InsetsGridsOnly(ep071GridsDT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep071Margin), true);
            if (ep206GridsDT.Rows.Count > 0)
                InsetsGridsOnly(ep206GridsDT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep206Margin), true);
            if (ep216GridsDT.Rows.Count > 0)
                InsetsGridsOnly(ep216GridsDT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep216Margin), true);
            if (ep111GridsDT.Rows.Count > 0)
                InsetsGridsOnly(ep111GridsDT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep111Margin), true);

            if (BostonGridsDT.Rows.Count > 0)
                InsetsGridsOnly(BostonGridsDT, ref DestinationDT, Convert.ToInt32(FrontMargins.BostonMargin), true);
            if (LeonGridsDT.Rows.Count > 0)
                InsetsGridsOnly(LeonGridsDT, ref DestinationDT, Convert.ToInt32(FrontMargins.LeonMargin), true);
            if (LimogGridsDT.Rows.Count > 0)
                InsetsGridsOnly(LimogGridsDT, ref DestinationDT, Convert.ToInt32(FrontMargins.LimogMargin), true);
            if (ep066Marsel4GridsDT.Rows.Count > 0)
                InsetsGridsOnly(ep066Marsel4GridsDT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep066Marsel4Margin), true);
            if (ep110JersyGridsDT.Rows.Count > 0)
                InsetsGridsOnly(ep110JersyGridsDT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep110JersyMargin), true);
            if (ep018Marsel1GridsDT.Rows.Count > 0)
                InsetsGridsOnly(ep018Marsel1GridsDT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep018Marsel1Margin), true);
            if (ep043ShervudGridsDT.Rows.Count > 0)
                InsetsGridsOnly(ep043ShervudGridsDT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep043ShervudMargin), true);
            if (ep112GridsDT.Rows.Count > 0)
                InsetsGridsOnly(ep112GridsDT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep112Margin), true);
            if (UrbanGridsDT.Rows.Count > 0)
                InsetsGridsOnly(UrbanGridsDT, ref DestinationDT, Convert.ToInt32(FrontMargins.UrbanMargin), true);
            if (AlbyGridsDT.Rows.Count > 0)
                InsetsGridsOnly(AlbyGridsDT, ref DestinationDT, Convert.ToInt32(FrontMargins.AlbyMargin), true);
            if (BrunoGridsDT.Rows.Count > 0)
                InsetsGridsOnly(BrunoGridsDT, ref DestinationDT, Convert.ToInt32(FrontMargins.BrunoMargin), true);
            if (epsh406Techno4GridsDT.Rows.Count > 0)
                InsetsGridsOnly(epsh406Techno4GridsDT, ref DestinationDT, Convert.ToInt32(FrontMargins.epsh406Techno4Margin), true);
            if (LukGridsDT.Rows.Count > 0)
                InsetsGridsOnly(LukGridsDT, ref DestinationDT, Convert.ToInt32(FrontMargins.LukMargin), true);
            if (LukPVHGridsDT.Rows.Count > 0)
                InsetsGridsOnly(LukPVHGridsDT, ref DestinationDT, Convert.ToInt32(FrontMargins.LukMargin), true);
            if (MilanoGridsDT.Rows.Count > 0)
                InsetsGridsOnly(MilanoGridsDT, ref DestinationDT, Convert.ToInt32(FrontMargins.MilanoMargin), true);
            if (PragaGridsDT.Rows.Count > 0)
                InsetsGridsOnly(PragaGridsDT, ref DestinationDT, Convert.ToInt32(FrontMargins.PragaMargin), true);
            if (SigmaGridsDT.Rows.Count > 0)
                InsetsGridsOnly(SigmaGridsDT, ref DestinationDT, Convert.ToInt32(FrontMargins.SigmaMargin), true);
            if (FatGridsDT.Rows.Count > 0)
                InsetsGridsOnly(FatGridsDT, ref DestinationDT, Convert.ToInt32(FrontMargins.FatMargin), true);

            if (DestinationDT.Rows.Count == 0)
                return;

            for (int i = 1; i < DestinationDT.Rows.Count; i++)
            {
                if (Convert.ToInt32(DestinationDT.Rows[i]["FrontID"]) == Convert.ToInt32(DestinationDT.Rows[i - 1]["FrontID"]) &&
                    Convert.ToInt32(DestinationDT.Rows[i]["InsetColorID"]) == Convert.ToInt32(DestinationDT.Rows[i - 1]["InsetColorID"]) &&
                    Convert.ToInt32(DestinationDT.Rows[i]["TechnoInsetColorID"]) == Convert.ToInt32(DestinationDT.Rows[i - 1]["TechnoInsetColorID"]))
                    DestinationDT.Rows[i]["Name"] = string.Empty;
            }
        }

        private void CollectMainOrders(DataTable SourceDT, ref DataTable DestinationDT)
        {
            DataTable DT = new DataTable();

            using (DataView DV = new DataView(SourceDT, string.Empty, "FrontID, ColorID, TechnoColorID, InsetTypeID, InsetColorID, TechnoInsetTypeID, TechnoInsetColorID, Height, Width", DataViewRowState.CurrentRows))
            {
                DT = DV.ToTable(true, new string[] { "FrontID", "ColorID", "TechnoColorID", "InsetTypeID", "InsetColorID", "TechnoInsetTypeID", "TechnoInsetColorID", "Height", "Width" });
            }
            for (int i = 0; i < DT.Rows.Count; i++)
            {
                decimal Square = 0;
                int Count = 0;

                DataRow[] rows = SourceDT.Select("FrontID=" + Convert.ToInt32(DT.Rows[i]["FrontID"]) + " AND ColorID=" + Convert.ToInt32(DT.Rows[i]["ColorID"]) + " AND TechnoColorID=" + Convert.ToInt32(DT.Rows[i]["TechnoColorID"]) +
                    " AND ColorID=" + Convert.ToInt32(DT.Rows[i]["ColorID"]) + " AND InsetTypeID=" + Convert.ToInt32(DT.Rows[i]["InsetTypeID"]) +
                    " AND InsetColorID=" + Convert.ToInt32(DT.Rows[i]["InsetColorID"]) + " AND TechnoInsetTypeID=" + Convert.ToInt32(DT.Rows[i]["TechnoInsetTypeID"]) + " AND TechnoInsetColorID=" + Convert.ToInt32(DT.Rows[i]["TechnoInsetColorID"]) +
                    " AND Height=" + Convert.ToInt32(DT.Rows[i]["Height"]) +
                    " AND Width=" + Convert.ToInt32(DT.Rows[i]["Width"]));
                foreach (DataRow item in rows)
                {
                    Count += Convert.ToInt32(item["Count"]);
                    Square += Convert.ToDecimal(item["Height"]) * Convert.ToDecimal(item["Width"]) * Convert.ToDecimal(item["Count"]) / 1000000;
                }
                Square = decimal.Round(Square, 3, MidpointRounding.AwayFromZero);

                DataRow NewRow = DestinationDT.NewRow();
                NewRow["Name"] = GetFrontName(Convert.ToInt32(rows[0]["FrontID"]));
                if (Convert.ToInt32(rows[0]["TechnoColorID"]) == -1)
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"]));
                if (Convert.ToInt32(rows[0]["TechnoColorID"]) != -1 && Convert.ToInt32(rows[0]["TechnoColorID"]) != Convert.ToInt32(rows[0]["ColorID"]))
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"])) + "/" + GetColorName(Convert.ToInt32(rows[0]["TechnoColorID"]));
                NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"]));
                NewRow["InsetColor"] = GetInsetTypeName(Convert.ToInt32(rows[0]["InsetTypeID"])) + " " + GetInsetColorName(Convert.ToInt32(rows[0]["InsetColorID"]));
                if (Convert.ToInt32(rows[0]["InsetTypeID"]) == 1)
                    NewRow["InsetColor"] = "Витрина";
                //if (Convert.ToInt32(rows[0]["InsetTypeID"]) == 2)
                //    NewRow["InsetColor"] = "Стекло";
                NewRow["TechnoColor"] = GetInsetTypeName(Convert.ToInt32(rows[0]["TechnoInsetTypeID"])) + " " + GetInsetColorName(Convert.ToInt32(rows[0]["TechnoInsetColorID"]));
                NewRow["Height"] = rows[0]["Height"];
                NewRow["Width"] = rows[0]["Width"];
                NewRow["Count"] = Count;
                NewRow["Square"] = Square;
                DestinationDT.Rows.Add(NewRow);
            }
            DestinationDT.DefaultView.Sort = "Name, FrameColor, InsetColor, TechnoColor, Height, Width";
        }

        private void CollectOrders(DataTable DistinctSizesDT, DataTable SourceDT, ref DataTable DestinationDT, int FrontType, string FrontName)
        {
            int InsetTypeID = 0;
            string ColName = string.Empty;
            string FrameColor = string.Empty;
            string InsetColor = string.Empty;

            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            DataTable DT3 = new DataTable();

            for (int y = 0; y < DistinctSizesDT.Rows.Count; y++)
            {
                using (DataView DV = new DataView(SourceDT))
                {
                    DT1 = DV.ToTable(true, new string[] { "ColorID", "TechnoColorID" });
                }
                for (int i = 0; i < DT1.Rows.Count; i++)
                {
                    using (DataView DV = new DataView(SourceDT, "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) + " AND TechnoColorID=" + Convert.ToInt32(DT1.Rows[i]["TechnoColorID"]), string.Empty, DataViewRowState.CurrentRows))
                    {
                        DT2 = DV.ToTable(true, new string[] { "InsetTypeID" });
                    }
                    for (int j = 0; j < DT2.Rows.Count; j++)
                    {
                        using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]), string.Empty, DataViewRowState.CurrentRows))
                        {
                            DT3 = DV.ToTable(true, new string[] { "InsetColorID", "TechnoInsetTypeID", "TechnoInsetColorID" });
                            for (int x = 0; x < DT3.Rows.Count; x++)
                            {
                                DataRow[] rows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) + " AND TechnoColorID=" + Convert.ToInt32(DT1.Rows[i]["TechnoColorID"]) + " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) + " AND InsetColorID=" + Convert.ToInt32(DT3.Rows[x]["InsetColorID"]) + " AND TechnoInsetTypeID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetTypeID"]) + " AND TechnoInsetColorID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]));

                                if (rows.Count() > 0)
                                {
                                    InsetTypeID = Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]);
                                    int TechnoInsetTypeID = Convert.ToInt32(DT3.Rows[x]["TechnoInsetTypeID"]);
                                    FrameColor = GetColorName(Convert.ToInt32(rows[0]["ColorID"]));

                                    if (Convert.ToInt32(rows[0]["TechnoColorID"]) != -1)
                                        FrameColor = "ИМП. " + GetColorName(Convert.ToInt32(rows[0]["ColorID"])) + "/" + GetColorName(Convert.ToInt32(rows[0]["TechnoColorID"]));

                                    int GroupID = Convert.ToInt32(InsetTypesDataTable.Select("InsetTypeID=" + InsetTypeID)[0]["GroupID"]);
                                    switch (GroupID)
                                    {
                                        case -1:
                                            InsetColor = "Витрина";
                                            break;

                                        case 3:
                                            InsetColor = GetInsetColorName(Convert.ToInt32(rows[0]["InsetColorID"]));
                                            break;

                                        case 4:
                                            InsetColor = GetInsetColorName(Convert.ToInt32(rows[0]["InsetColorID"]));
                                            break;

                                        case 7:
                                            InsetColor = "фил " + GetInsetColorName(Convert.ToInt32(rows[0]["InsetColorID"]));
                                            break;

                                        case 8:
                                            InsetColor = "фил " + GetInsetColorName(Convert.ToInt32(rows[0]["InsetColorID"]));
                                            break;

                                        case 16:
                                            InsetColor = "реш " + GetInsetColorName(Convert.ToInt32(rows[0]["InsetColorID"]));
                                            break;

                                        case 17:
                                            InsetColor = "реш " + GetInsetColorName(Convert.ToInt32(rows[0]["InsetColorID"]));
                                            break;

                                        case 18:
                                            InsetColor = "реш " + GetInsetColorName(Convert.ToInt32(rows[0]["InsetColorID"]));
                                            break;

                                        case 19:
                                            InsetColor = "реш " + GetInsetColorName(Convert.ToInt32(rows[0]["InsetColorID"]));
                                            break;

                                        case 12:
                                            InsetColor = "люкс " + GetInsetColorName(Convert.ToInt32(rows[0]["InsetColorID"]));
                                            break;

                                        case 13:
                                            InsetColor = "мега " + GetInsetColorName(Convert.ToInt32(rows[0]["InsetColorID"]));
                                            break;

                                        default:
                                            InsetColor = GetInsetColorName(Convert.ToInt32(rows[0]["InsetColorID"]));
                                            break;
                                    }

                                    if (InsetTypeID == 2 && TechnoInsetTypeID != -1)
                                        InsetColor = GetInsetTypeName(TechnoInsetTypeID) + " " + GetInsetColorName(Convert.ToInt32(DT3.Rows[x]["InsetColorID"]));

                                    ColName = FrameColor + "(" + InsetColor + ")_" + FrontType;
                                    if (!DestinationDT.Columns.Contains(ColName))
                                        DestinationDT.Columns.Add(new DataColumn(ColName, Type.GetType("System.String")));

                                    DestinationDT.Rows[0][ColName] = FrontName;

                                    DataRow[] Srows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) + " AND TechnoColorID=" + Convert.ToInt32(DT1.Rows[i]["TechnoColorID"]) + " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) + " AND InsetColorID=" + Convert.ToInt32(DT3.Rows[x]["InsetColorID"]) +
                                        " AND TechnoInsetTypeID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetTypeID"]) + " AND TechnoInsetColorID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]) +
                                        " AND Height=" + Convert.ToInt32(DistinctSizesDT.Rows[y]["Height"]) + " AND Width=" + Convert.ToInt32(DistinctSizesDT.Rows[y]["Width"]));
                                    if (Srows.Count() > 0)
                                    {
                                        int Count = 0;
                                        foreach (DataRow item in Srows)
                                            Count += Convert.ToInt32(item["Count"]);

                                        DataRow[] Drows = DestinationDT.Select("Sizes='" + DistinctSizesDT.Rows[y]["Height"].ToString() + " X " + DistinctSizesDT.Rows[y]["Width"].ToString() + "'");
                                        if (Drows.Count() == 0)
                                        {
                                            DataRow NewRow = DestinationDT.NewRow();
                                            NewRow["Sizes"] = DistinctSizesDT.Rows[y]["Height"].ToString() + " X " + DistinctSizesDT.Rows[y]["Width"].ToString();
                                            NewRow["Height"] = DistinctSizesDT.Rows[y]["Height"];
                                            NewRow["Width"] = DistinctSizesDT.Rows[y]["Width"];
                                            NewRow[ColName] = Count;
                                            DestinationDT.Rows.Add(NewRow);
                                        }
                                        else
                                        {
                                            Drows[0][ColName] = Count;
                                        }
                                    }
                                }
                                else
                                    continue;
                            }
                        }
                    }
                }
            }
        }

        private void CollectPressOnly(ref DataTable DestinationDT)
        {
            if (TechnoNSimpleDT.Rows.Count > 0)
                InsetsPressOnly(TechnoNSimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.TechnoNMargin), true);
            if (AntaliaSimpleDT.Rows.Count > 0)
                InsetsPressOnly(AntaliaSimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.AntaliaMargin), true);
            if (Nord95SimpleDT.Rows.Count > 0)
                InsetsPressOnly(Nord95SimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.Nord95Margin), true);
            if (epFoxSimpleDT.Rows.Count > 0)
                InsetsPressOnly(epFoxSimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.epFoxMargin), true);
            if (VeneciaSimpleDT.Rows.Count > 0)
                InsetsPressOnly(VeneciaSimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.VeneciaMargin), true);
            //if (BergamoSimpleDT.Rows.Count > 0)
            //    InsetsPressOnly(BergamoSimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.BergamoMargin), true);
            DataTable DT = BergamoSimpleDT.Clone();

            DataRow[] rows = BergamoSimpleDT.Select("(Height>150 OR Width>150)");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            AllInsets(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.BergamoMargin), true);

            DT.Clear();
            rows = BergamoSimpleDT.Select("(Height<=150 OR Width<=150)");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            AllInsets(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.BergamoNarrowMargin), true);

            if (ep041SimpleDT.Rows.Count > 0)
                InsetsPressOnly(ep041SimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep041Margin), true);
            if (ep071SimpleDT.Rows.Count > 0)
                InsetsPressOnly(ep071SimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep071Margin), true);
            if (ep206SimpleDT.Rows.Count > 0)
                InsetsPressOnly(ep206SimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep206Margin), true);
            if (ep216SimpleDT.Rows.Count > 0)
                InsetsPressOnly(ep216SimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep216Margin), true);
            if (ep111SimpleDT.Rows.Count > 0)
                InsetsPressOnly(ep111SimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep111Margin), true);

            if (BostonSimpleDT.Rows.Count > 0)
                InsetsPressOnly(BostonSimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.BostonMargin), true);
            if (LeonSimpleDT.Rows.Count > 0)
                InsetsPressOnly(LeonSimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.LeonMargin), true);
            if (LimogSimpleDT.Rows.Count > 0)
                InsetsPressOnly(LimogSimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.LimogMargin), true);
            if (ep066Marsel4SimpleDT.Rows.Count > 0)
                InsetsPressOnly(ep066Marsel4SimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep066Marsel4Margin), true);
            if (ep110JersySimpleDT.Rows.Count > 0)
                InsetsPressOnly(ep110JersySimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep110JersyMargin), true);
            if (ep018Marsel1SimpleDT.Rows.Count > 0)
                InsetsPressOnly(ep018Marsel1SimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep018Marsel1Margin), true);
            if (ep043ShervudSimpleDT.Rows.Count > 0)
                InsetsPressOnly(ep043ShervudSimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep043ShervudMargin), true);
            if (ep112SimpleDT.Rows.Count > 0)
                InsetsPressOnly(ep112SimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep112Margin), true);
            if (UrbanSimpleDT.Rows.Count > 0)
                InsetsPressOnly(UrbanSimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.UrbanMargin), true);
            if (AlbySimpleDT.Rows.Count > 0)
                InsetsPressOnly(AlbySimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.AlbyMargin), true);
            if (BrunoSimpleDT.Rows.Count > 0)
                InsetsPressOnly(BrunoSimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.BrunoMargin), true);
            if (epsh406Techno4SimpleDT.Rows.Count > 0)
                InsetsPressOnly(epsh406Techno4SimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.epsh406Techno4Margin), true);
            if (LukSimpleDT.Rows.Count > 0)
                InsetsPressOnly(LukSimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.LukMargin), true);
            if (LukPVHSimpleDT.Rows.Count > 0)
                InsetsPressOnly(LukPVHSimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.LukMargin), true);
            if (MilanoSimpleDT.Rows.Count > 0)
                InsetsPressOnly(MilanoSimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.MilanoMargin), true);
            if (PragaSimpleDT.Rows.Count > 0)
                InsetsPressOnly(PragaSimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.PragaMargin), true);
            if (SigmaSimpleDT.Rows.Count > 0)
                InsetsPressOnly(SigmaSimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.SigmaMargin), true);
            if (FatSimpleDT.Rows.Count > 0)
                InsetsPressOnly(FatSimpleDT, ref DestinationDT, Convert.ToInt32(FrontMargins.FatMargin), true);

            if (DestinationDT.Rows.Count == 0)
                return;

            for (int i = 1; i < DestinationDT.Rows.Count; i++)
            {
                if (Convert.ToInt32(DestinationDT.Rows[i]["FrontID"]) == Convert.ToInt32(DestinationDT.Rows[i - 1]["FrontID"]) &&
                    Convert.ToInt32(DestinationDT.Rows[i]["InsetColorID"]) == Convert.ToInt32(DestinationDT.Rows[i - 1]["InsetColorID"]) &&
                    Convert.ToInt32(DestinationDT.Rows[i]["TechnoInsetColorID"]) == Convert.ToInt32(DestinationDT.Rows[i - 1]["TechnoInsetColorID"]))
                    DestinationDT.Rows[i]["Name"] = string.Empty;
            }
        }

        private void CombineAssembly(ref DataTable DestinationDT)
        {
        }

        private void CombineComecFilenka(ref DataTable DestinationDT)
        {
            string filter = @"TechnoProfileID<>-1 AND InsetTypeID IN (2069,2070,2071,2073,2075,2077,2233,3644,29043,29531)";

            DataTable DT = AntaliaOrdersDT.Clone();
            DataRow[] rows = AntaliaOrdersDT.Select(filter +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.AntaliaMargin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.AntaliaMargin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            string Front = GetFrontName(Convert.ToInt32(Fronts.Antalia));
            ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.AntaliaWidth), 1, Front, false);

            DT.Clear();
            rows = TechnoNOrdersDT.Select(filter +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.TechnoNMargin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.TechnoNMargin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.TechnoN));
            ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.TechnoNWidth), 1, Front, false);

            DT.Clear();
            rows = Nord95OrdersDT.Select(filter +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.Nord95Margin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.Nord95Margin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Nord95));
            ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.Nord95Width), 1, Front, false);

            DT.Clear();
            rows = epFoxOrdersDT.Select(filter +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.epFoxMargin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.epFoxMargin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.epFox));
            ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.epFoxWidth), 1, Front, false);

            DT.Clear();
            rows = BergamoOrdersDT.Select(filter +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.BergamoMargin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.BergamoMargin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Bergamo));
            ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.BergamoWidth), 1, Front, false);

            DT.Clear();
            rows = ep041OrdersDT.Select(filter +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.ep041Margin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.ep041Margin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep041));
            ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep041Width), 1, Front, false);

            DT.Clear();
            rows = ep071OrdersDT.Select(filter +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.ep071Margin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.ep071Margin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep071));
            ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep071Width), 1, Front, false);

            DT.Clear();
            rows = ep206OrdersDT.Select(filter +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.ep206Margin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.ep206Margin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep206));
            ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep206Width), 1, Front, false);

            DT.Clear();
            rows = ep216OrdersDT.Select(filter +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.ep216Margin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.ep216Margin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep216));
            ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep216Width), 1, Front, false);

            DT.Clear();
            rows = ep111OrdersDT.Select(filter +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.ep111Margin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.ep111Margin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep111));
            ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep111Width), 1, Front, false);

            DT.Clear();
            rows = BostonOrdersDT.Select(filter +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.BostonMargin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.BostonMargin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Boston));
            ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.BostonWidth), 1, Front, false);

            DataTable TempDT = new DataTable();
            TempDT.Clear();
            using (DataView DV = new DataView(VeneciaOrdersDT))
            {
                TempDT = DV.ToTable(true, new string[] { "ProfileID" });
            }
            for (int i = 0; i < TempDT.Rows.Count; i++)
            {
                int ProfileID = Convert.ToInt32(TempDT.Rows[i]["ProfileID"]);
                DT.Clear();
                rows = VeneciaOrdersDT.Select(filter + " AND ProfileID=" + ProfileID +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.VeneciaMargin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.VeneciaMargin) + 100) + ")");
                if (rows.Count() > 0)
                {
                    foreach (DataRow item in rows)
                        DT.Rows.Add(item.ItemArray);
                    Front = ProfileName(Convert.ToInt32(rows[0]["FrontConfigID"]));
                    ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.VeneciaWidth), 1, Front, false);
                }
            }

            TempDT.Clear();
            using (DataView DV = new DataView(LeonOrdersDT))
            {
                TempDT = DV.ToTable(true, new string[] { "ProfileID" });
            }
            for (int i = 0; i < TempDT.Rows.Count; i++)
            {
                int ProfileID = Convert.ToInt32(TempDT.Rows[i]["ProfileID"]);
                DT.Clear();
                rows = LeonOrdersDT.Select(filter + " AND ProfileID=" + ProfileID +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.LeonMargin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.LeonMargin) + 100) + ")");
                if (rows.Count() > 0)
                {
                    foreach (DataRow item in rows)
                        DT.Rows.Add(item.ItemArray);
                    Front = ProfileName(Convert.ToInt32(rows[0]["FrontConfigID"]));
                    ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.LeonWidth), 1, Front, false);
                }
            }

            DT.Clear();
            rows = LimogOrdersDT.Select(filter +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.LimogMargin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.LimogMargin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Limog));
            ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.LimogWidth), 1, Front, false);

            DT.Clear();
            rows = ep066Marsel4OrdersDT.Select(filter +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.ep066Marsel4Margin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.ep066Marsel4Margin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep066Marsel4));
            ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep066Marsel4Width), 1, Front, false);

            DT.Clear();
            rows = ep110JersyOrdersDT.Select(filter +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.ep110JersyMargin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.ep110JersyMargin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep110Jersy));
            ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep110JersyWidth), 1, Front, false);

            DT.Clear();
            rows = ep018Marsel1OrdersDT.Select(filter +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.ep018Marsel1Margin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.ep018Marsel1Margin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep018Marsel1));
            ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep018Marsel1Width), 1, Front, false);

            DT.Clear();
            rows = ep043ShervudOrdersDT.Select(filter +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.ep043ShervudMargin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.ep043ShervudMargin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep043Shervud));
            ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep043ShervudWidth), 1, Front, false);

            DT.Clear();
            rows = ep112OrdersDT.Select(filter +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.ep112Margin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.ep112Margin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep112));
            ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep112Width), 1, Front, false);

            DT.Clear();
            rows = UrbanOrdersDT.Select(filter +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.UrbanMargin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.UrbanMargin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Urban));
            ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.UrbanWidth), 1, Front, false);

            DT.Clear();
            rows = AlbyOrdersDT.Select(filter +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.AlbyMargin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.AlbyMargin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Alby));
            ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.AlbyWidth), 1, Front, false);

            DT.Clear();
            rows = BrunoOrdersDT.Select(filter +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.BrunoMargin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.BrunoMargin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Bruno));
            ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.BrunoWidth), 1, Front, false);

            DT.Clear();
            rows = epsh406Techno4OrdersDT.Select(filter +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.epsh406Techno4Margin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.epsh406Techno4Margin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.epsh406Techno4));
            ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.epsh406Techno4Width), 1, Front, false);

            DT.Clear();
            rows = LukOrdersDT.Select(filter +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.LukMargin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.LukMargin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Luk));
            ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.LukWidth), 1, Front, false);

            DT.Clear();
            rows = LukPVHOrdersDT.Select(filter +
                "  AND (Height>" + (Convert.ToInt32(FrontMargins.LukMargin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.LukMargin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.LukPVH));
            ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.LukWidth), 1, Front, false);

            TempDT.Clear();
            using (DataView DV = new DataView(MilanoOrdersDT))
            {
                TempDT = DV.ToTable(true, new string[] { "ProfileID" });
            }
            for (int i = 0; i < TempDT.Rows.Count; i++)
            {
                int ProfileID = Convert.ToInt32(TempDT.Rows[i]["ProfileID"]);
                DT.Clear();
                rows = MilanoOrdersDT.Select(filter + " AND ProfileID=" + ProfileID +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.MilanoMargin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.MilanoMargin) + 100) + ")");
                if (rows.Count() > 0)
                {
                    foreach (DataRow item in rows)
                        DT.Rows.Add(item.ItemArray);
                    Front = ProfileName(Convert.ToInt32(rows[0]["FrontConfigID"]));
                    ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.MilanoWidth), 1, Front, false);
                }
            }

            DT.Clear();
            rows = PragaOrdersDT.Select(filter +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.PragaMargin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.PragaMargin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Praga));
            ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.PragaWidth), 1, Front, false);

            DT.Clear();
            rows = SigmaOrdersDT.Select(filter +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.SigmaMargin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.SigmaMargin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Sigma));
            ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.SigmaWidth), 1, Front, false);

            DT.Clear();
            rows = FatOrdersDT.Select(filter +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.FatMargin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.FatMargin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Fat));
            ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.FatWidth), 1, Front, false);

            using (DataView DV = new DataView(DestinationDT.Copy()))
            {
                DV.Sort = "Front, Color, ProfileType, Height DESC";
                DestinationDT.Clear();
                DestinationDT = DV.ToTable();
            }

            for (int i = 0; i < DestinationDT.Rows.Count; i++)
            {
                if (Convert.ToInt32(DestinationDT.Rows[i]["VitrinaCount"]) > 0)
                    DestinationDT.Rows[i]["Notes"] = Convert.ToInt32(DestinationDT.Rows[i]["VitrinaCount"]) + " витр.";
                if (DestinationDT.Rows[i]["BoxCount"] != DBNull.Value && Convert.ToInt32(DestinationDT.Rows[i]["BoxCount"]) > 0)
                {
                    if (DestinationDT.Rows[i]["Notes"] != DBNull.Value)
                    {
                        if (DestinationDT.Rows[i]["Notes"].ToString().Length > 0)
                            DestinationDT.Rows[i]["Notes"] = DestinationDT.Rows[i]["Notes"].ToString() + ", " + Convert.ToInt32(DestinationDT.Rows[i]["BoxCount"]) + " шуф.";
                    }
                    else
                        DestinationDT.Rows[i]["Notes"] = Convert.ToInt32(DestinationDT.Rows[i]["BoxCount"]) + " шуф.";
                }
                if (DestinationDT.Rows[i]["ImpostCount"] != DBNull.Value && Convert.ToInt32(DestinationDT.Rows[i]["ImpostCount"]) > 0)
                {
                    if (DestinationDT.Rows[i]["Notes"] != DBNull.Value)
                    {
                        if (DestinationDT.Rows[i]["Notes"].ToString().Length > 0)
                            DestinationDT.Rows[i]["Notes"] = DestinationDT.Rows[i]["Notes"].ToString() + ", " + Convert.ToInt32(DestinationDT.Rows[i]["ImpostCount"]) + " имп.";
                    }
                    else
                        DestinationDT.Rows[i]["Notes"] = Convert.ToInt32(DestinationDT.Rows[i]["ImpostCount"]) + " имп.";
                }
                if (i == 0)
                    continue;
                if (Convert.ToInt32(DestinationDT.Rows[i]["ProfileType"]) == Convert.ToInt32(DestinationDT.Rows[i - 1]["ProfileType"]) &&
                    Convert.ToInt32(DestinationDT.Rows[i]["ColorType"]) == Convert.ToInt32(DestinationDT.Rows[i - 1]["ColorType"]) &&
                    Convert.ToInt32(DestinationDT.Rows[i]["FrontType"]) == Convert.ToInt32(DestinationDT.Rows[i - 1]["FrontType"]))
                {
                    DestinationDT.Rows[i]["Front"] = string.Empty;
                    DestinationDT.Rows[i]["Color"] = string.Empty;
                }
            }
        }

        /// <summary>
        /// вкладка Рапид, нижняя таблица с импостом
        /// </summary>
        /// <param name="DestinationDT"></param>
        /// <param name="IsBox"></param>
        private void CombineComecSimple(ref DataTable DestinationDT, bool IsBox)
        {
            string filter = @"TechnoProfileID<>-1 AND InsetTypeID NOT IN (2069,2070,2071,2073,2075,2077,2233,3644,29043,29531) AND InsetTypeID NOT IN (860) AND InsetTypeID NOT IN (685,686,687,688,29470,29471)";

            DataTable DT = AntaliaOrdersDT.Clone();
            DataRow[] rows = AntaliaOrdersDT.Select(filter);
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            string Front = GetFrontName(Convert.ToInt32(Fronts.Antalia));
            ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.AntaliaWidth), 1, Front, false);

            if (!IsBox)
            {
                DT.Clear();
                rows = TechnoNOrdersDT.Select(filter +
                    " AND Height>" + (Convert.ToInt32(FrontMargins.TechnoNWidth)) + " AND Width>" + (Convert.ToInt32(FrontMargins.TechnoNWidth)));
                foreach (DataRow item in rows)
                    DT.Rows.Add(item.ItemArray);
                Front = GetFrontName(Convert.ToInt32(Fronts.TechnoN));
                ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.TechnoNWidth), 1, Front, false);
            }

            DT.Clear();
            rows = Nord95OrdersDT.Select(filter);
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Nord95));
            ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.Nord95Width), 1, Front, false);

            DT.Clear();
            rows = epFoxOrdersDT.Select(filter);
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.epFox));
            ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.epFoxWidth), 1, Front, false);

            DT.Clear();
            rows = BergamoOrdersDT.Select(filter + " AND (Height>150 OR Width>150)");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Bergamo));
            ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.BergamoWidth), 1, Front, false);

            DT.Clear();
            rows = BergamoOrdersDT.Select(filter + " AND (Height<=150 OR Width<=150)");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = "ПШ-050 22х55 Бергамо";
            RapidSimpleBergamoSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.BergamoWidth), 1, 29548, Front, false);

            DT.Clear();
            rows = ep041OrdersDT.Select(filter);
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep041));
            ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep041Width), 1, Front, false);

            DT.Clear();
            rows = ep071OrdersDT.Select(filter);
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep071));
            ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep071Width), 1, Front, false);

            DT.Clear();
            rows = ep206OrdersDT.Select(filter);
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep206));
            ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep206Width), 1, Front, false);

            DT.Clear();
            rows = ep216OrdersDT.Select(filter);
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep216));
            ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep216Width), 1, Front, false);

            DT.Clear();
            rows = ep111OrdersDT.Select(filter);
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep111));
            ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep111Width), 1, Front, false);

            DT.Clear();
            rows = BostonOrdersDT.Select(filter);
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Boston));
            ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.BostonWidth), 1, Front, false);

            DataTable TempDT = new DataTable();
            TempDT.Clear();
            using (DataView DV = new DataView(VeneciaOrdersDT))
            {
                TempDT = DV.ToTable(true, new string[] { "ProfileID" });
            }
            for (int i = 0; i < TempDT.Rows.Count; i++)
            {
                int ProfileID = Convert.ToInt32(TempDT.Rows[i]["ProfileID"]);
                DT.Clear();
                rows = VeneciaOrdersDT.Select(filter + " AND ProfileID=" + ProfileID);
                if (rows.Count() > 0)
                {
                    foreach (DataRow item in rows)
                        DT.Rows.Add(item.ItemArray);
                    Front = ProfileName(Convert.ToInt32(rows[0]["FrontConfigID"]));
                    ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.VeneciaWidth), 1, Front, false);
                }
            }

            TempDT.Clear();
            using (DataView DV = new DataView(LeonOrdersDT))
            {
                TempDT = DV.ToTable(true, new string[] { "ProfileID" });
            }
            for (int i = 0; i < TempDT.Rows.Count; i++)
            {
                int ProfileID = Convert.ToInt32(TempDT.Rows[i]["ProfileID"]);
                DT.Clear();
                rows = LeonOrdersDT.Select(filter + " AND ProfileID=" + ProfileID);
                if (rows.Count() > 0)
                {
                    foreach (DataRow item in rows)
                        DT.Rows.Add(item.ItemArray);
                    Front = ProfileName(Convert.ToInt32(rows[0]["FrontConfigID"]));
                    ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.LeonWidth), 1, Front, false);
                }
            }

            DT.Clear();
            rows = LimogOrdersDT.Select(filter);
            //AND Height>" + (Convert.ToInt32(FrontMargins.LimogWidth) - 10) + " AND Width>" + (Convert.ToInt32(FrontMargins.LimogWidth) - 10));
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Limog));
            ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.LimogWidth), 1, Front, false);

            DT.Clear();
            rows = ep066Marsel4OrdersDT.Select(filter);
            //AND Height>" + (Convert.ToInt32(FrontMargins.ep066Marsel4Width) - 10) + " AND Width>" + (Convert.ToInt32(FrontMargins.ep066Marsel4Width) - 10));
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep066Marsel4));
            ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep066Marsel4Width), 1, Front, false);

            DT.Clear();
            rows = ep110JersyOrdersDT.Select(filter);
            //AND Height>" + (Convert.ToInt32(FrontMargins.ep110JersyWidth) - 10) + " AND Width>" + (Convert.ToInt32(FrontMargins.ep110JersyWidth) - 10));
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep110Jersy));
            ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep110JersyWidth), 1, Front, false);

            DT.Clear();
            rows = ep018Marsel1OrdersDT.Select(filter);
            //AND Height>" + (Convert.ToInt32(FrontMargins.ep018Marsel1Width) - 10) + " AND Width>" + (Convert.ToInt32(FrontMargins.ep018Marsel1Width) - 10));
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep018Marsel1));
            ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep018Marsel1Width), 1, Front, false);

            DT.Clear();
            rows = ep043ShervudOrdersDT.Select(filter);
            //AND Height>" + (Convert.ToInt32(FrontMargins.ep043ShervudWidth) - 10) + " AND Width>" + (Convert.ToInt32(FrontMargins.ep043ShervudWidth) - 10));
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep043Shervud));
            ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep043ShervudWidth), 1, Front, false);

            DT.Clear();
            rows = ep112OrdersDT.Select(filter);
            //AND Height>" + (Convert.ToInt32(FrontMargins.ep112Width) - 10) + " AND Width>" + (Convert.ToInt32(FrontMargins.ep112Width) - 10));
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep112));
            ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep112Width), 1, Front, false);

            DT.Clear();
            rows = UrbanOrdersDT.Select(filter);
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Urban));
            ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.UrbanWidth), 1, Front, false);

            DT.Clear();
            rows = AlbyOrdersDT.Select(filter);
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Alby));
            ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.AlbyWidth), 1, Front, false);

            DT.Clear();
            rows = BrunoOrdersDT.Select(filter);
            //AND Height>" + (Convert.ToInt32(FrontMargins.BrunoWidth) - 10) + " AND Width>" + (Convert.ToInt32(FrontMargins.BrunoWidth) - 10));
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Bruno));
            ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.BrunoWidth), 1, Front, false);

            DT.Clear();
            rows = epsh406Techno4OrdersDT.Select(filter);
            //AND Height>" + (Convert.ToInt32(FrontMargins.epsh406Techno4Width) - 10) + " AND Width>" + (Convert.ToInt32(FrontMargins.epsh406Techno4Width) - 10));
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.epsh406Techno4));
            ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.epsh406Techno4Width), 1, Front, false);

            DT.Clear();
            rows = LukOrdersDT.Select(filter);
            //AND Height>" + (Convert.ToInt32(FrontMargins.LukWidth) - 10) + " AND Width>" + (Convert.ToInt32(FrontMargins.LukWidth) - 10));
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Luk));
            ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.LukWidth), 1, Front, false);

            DT.Clear();
            rows = LukPVHOrdersDT.Select(filter);
            //AND Height>" + (Convert.ToInt32(FrontMargins.LukWidth) - 10) + " AND Width>" + (Convert.ToInt32(FrontMargins.LukWidth) - 10));
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.LukPVH));
            ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.LukWidth), 1, Front, false);

            TempDT.Clear();
            using (DataView DV = new DataView(MilanoOrdersDT))
            {
                TempDT = DV.ToTable(true, new string[] { "ProfileID" });
            }
            for (int i = 0; i < TempDT.Rows.Count; i++)
            {
                int ProfileID = Convert.ToInt32(TempDT.Rows[i]["ProfileID"]);
                DT.Clear();
                rows = MilanoOrdersDT.Select(filter + " AND ProfileID=" + ProfileID);
                if (rows.Count() > 0)
                {
                    foreach (DataRow item in rows)
                        DT.Rows.Add(item.ItemArray);
                    Front = ProfileName(Convert.ToInt32(rows[0]["FrontConfigID"]));
                    ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.MilanoWidth), 1, Front, false);
                }
            }

            DT.Clear();
            rows = PragaOrdersDT.Select(filter);
            //AND Height>" + (Convert.ToInt32(FrontMargins.PragaWidth) - 10) + " AND Width>" + (Convert.ToInt32(FrontMargins.PragaWidth) - 10));
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Praga));
            ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.PragaWidth), 1, Front, false);

            DT.Clear();
            rows = SigmaOrdersDT.Select(filter);
            //AND Height>" + (Convert.ToInt32(FrontMargins.SigmaWidth) - 10) + " AND Width>" + (Convert.ToInt32(FrontMargins.SigmaWidth) - 10));
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Sigma));
            ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.SigmaWidth), 1, Front, false);

            DT.Clear();
            rows = FatOrdersDT.Select(filter);
            //AND Height>" + (Convert.ToInt32(FrontMargins.FatWidth) - 10) + " AND Width>" + (Convert.ToInt32(FrontMargins.FatWidth) - 10));
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Fat));
            ComecSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.FatWidth), 1, Front, false);

            //CombineRapidGrids(ref RapidDT);
            //CombineRapidVitrina(ref RapidDT);
        }

        private void CombineRapidBoxes(ref DataTable DestinationDT)
        {
            string filter = @"InsetTypeID NOT IN (2069,2070,2071,2073,2075,2077,2233,3644,29043,29531) AND InsetTypeID NOT IN (860) AND InsetTypeID NOT IN (685,686,687,688,29470,29471)";
            DataTable DT = AntaliaOrdersDT.Clone();
            DataRow[] rows = TechnoNOrdersDT.Select(filter +
                " AND Height<=" + (Convert.ToInt32(FrontMargins.TechnoNWidth)) + " OR Width<=" + (Convert.ToInt32(FrontMargins.TechnoNWidth)));
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            string Front = GetFrontName(Convert.ToInt32(Fronts.TechnoN));
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.TechnoNWidth), 2, Front, false);
        }

        private void CombineRapidFilenka(ref DataTable DestinationDT)
        {
            string filter = @"InsetTypeID IN (2069,2070,2071,2073,2075,2077,2233,3644,29043,29531)";

            DataTable DT = AntaliaOrdersDT.Clone();
            DataRow[] rows = AntaliaOrdersDT.Select(filter +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.AntaliaMargin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.AntaliaMargin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            string Front = GetFrontName(Convert.ToInt32(Fronts.Antalia));
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.AntaliaWidth), 1, Front, false);

            DT.Clear();
            rows = TechnoNOrdersDT.Select(filter +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.TechnoNMargin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.TechnoNMargin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.TechnoN));
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.TechnoNWidth), 1, Front, false);

            DT.Clear();
            rows = Nord95OrdersDT.Select(filter +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.Nord95Margin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.Nord95Margin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Nord95));
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.Nord95Width), 1, Front, false);

            DT.Clear();
            rows = epFoxOrdersDT.Select(filter +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.epFoxMargin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.epFoxMargin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.epFox));
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.epFoxWidth), 1, Front, false);

            DT.Clear();
            rows = BergamoOrdersDT.Select(filter +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.BergamoMargin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.BergamoMargin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Bergamo));
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.BergamoWidth), 1, Front, false);

            DT.Clear();
            rows = ep041OrdersDT.Select(filter +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.ep041Margin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.ep041Margin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep041));
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep041Width), 1, Front, false);

            DT.Clear();
            rows = ep071OrdersDT.Select(filter +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.ep071Margin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.ep071Margin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep071));
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep071Width), 1, Front, false);

            DT.Clear();
            rows = ep206OrdersDT.Select(filter +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.ep206Margin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.ep206Margin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep206));
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep206Width), 1, Front, false);

            DT.Clear();
            rows = ep216OrdersDT.Select(filter +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.ep216Margin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.ep216Margin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep216));
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep216Width), 1, Front, false);

            DT.Clear();
            rows = ep111OrdersDT.Select(filter +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.ep111Margin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.ep111Margin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep111));
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep111Width), 1, Front, false);

            DT.Clear();
            rows = BostonOrdersDT.Select(filter +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.BostonMargin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.BostonMargin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Boston));
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.BostonWidth), 1, Front, false);

            DataTable TempDT = new DataTable();
            TempDT.Clear();
            using (DataView DV = new DataView(VeneciaOrdersDT))
            {
                TempDT = DV.ToTable(true, new string[] { "ProfileID" });
            }
            for (int i = 0; i < TempDT.Rows.Count; i++)
            {
                int ProfileID = Convert.ToInt32(TempDT.Rows[i]["ProfileID"]);
                DT.Clear();
                rows = VeneciaOrdersDT.Select(filter + " AND ProfileID=" + ProfileID +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.VeneciaMargin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.VeneciaMargin) + 100) + ")");
                if (rows.Count() > 0)
                {
                    foreach (DataRow item in rows)
                        DT.Rows.Add(item.ItemArray);
                    Front = ProfileName(Convert.ToInt32(rows[0]["FrontConfigID"]));
                    RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.VeneciaWidth), 1, Front, false);
                }
            }

            TempDT.Clear();
            using (DataView DV = new DataView(LeonOrdersDT))
            {
                TempDT = DV.ToTable(true, new string[] { "ProfileID" });
            }
            for (int i = 0; i < TempDT.Rows.Count; i++)
            {
                int ProfileID = Convert.ToInt32(TempDT.Rows[i]["ProfileID"]);
                DT.Clear();
                rows = LeonOrdersDT.Select(filter + " AND ProfileID=" + ProfileID +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.LeonMargin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.LeonMargin) + 100) + ")");
                if (rows.Count() > 0)
                {
                    foreach (DataRow item in rows)
                        DT.Rows.Add(item.ItemArray);
                    Front = ProfileName(Convert.ToInt32(rows[0]["FrontConfigID"]));
                    RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.LeonWidth), 1, Front, false);
                }
            }

            DT.Clear();
            rows = LimogOrdersDT.Select(filter +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.LimogMargin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.LimogMargin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Limog));
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.LimogWidth), 1, Front, false);

            DT.Clear();
            rows = ep066Marsel4OrdersDT.Select(filter +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.ep066Marsel4Margin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.ep066Marsel4Margin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep066Marsel4));
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep066Marsel4Width), 1, Front, false);

            DT.Clear();
            rows = ep110JersyOrdersDT.Select(filter +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.ep110JersyMargin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.ep110JersyMargin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep110Jersy));
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep110JersyWidth), 1, Front, false);

            DT.Clear();
            rows = ep018Marsel1OrdersDT.Select(filter +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.ep018Marsel1Margin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.ep018Marsel1Margin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep018Marsel1));
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep018Marsel1Width), 1, Front, false);

            DT.Clear();
            rows = ep043ShervudOrdersDT.Select(filter +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.ep043ShervudMargin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.ep043ShervudMargin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep043Shervud));
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep043ShervudWidth), 1, Front, false);

            DT.Clear();
            rows = ep112OrdersDT.Select(filter +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.ep112Margin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.ep112Margin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep112));
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep112Width), 1, Front, false);

            DT.Clear();
            rows = UrbanOrdersDT.Select(filter +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.UrbanMargin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.UrbanMargin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Urban));
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.UrbanWidth), 1, Front, false);

            DT.Clear();
            rows = AlbyOrdersDT.Select(filter +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.AlbyMargin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.AlbyMargin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Alby));
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.AlbyWidth), 1, Front, false);

            DT.Clear();
            rows = BrunoOrdersDT.Select(filter +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.BrunoMargin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.BrunoMargin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Bruno));
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.BrunoWidth), 1, Front, false);

            DT.Clear();
            rows = epsh406Techno4OrdersDT.Select(filter +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.epsh406Techno4Margin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.epsh406Techno4Margin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.epsh406Techno4));
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.epsh406Techno4Width), 1, Front, false);

            DT.Clear();
            rows = LukOrdersDT.Select(filter +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.LukMargin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.LukMargin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Luk));
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.LukWidth), 1, Front, false);

            DT.Clear();
            rows = LukPVHOrdersDT.Select(filter +
                "  AND (Height>" + (Convert.ToInt32(FrontMargins.LukMargin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.LukMargin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.LukPVH));
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.LukWidth), 1, Front, false);

            TempDT.Clear();
            using (DataView DV = new DataView(MilanoOrdersDT))
            {
                TempDT = DV.ToTable(true, new string[] { "ProfileID" });
            }
            for (int i = 0; i < TempDT.Rows.Count; i++)
            {
                int ProfileID = Convert.ToInt32(TempDT.Rows[i]["ProfileID"]);
                DT.Clear();
                rows = MilanoOrdersDT.Select(filter + " AND ProfileID=" + ProfileID +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.MilanoMargin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.MilanoMargin) + 100) + ")");
                if (rows.Count() > 0)
                {
                    foreach (DataRow item in rows)
                        DT.Rows.Add(item.ItemArray);
                    Front = ProfileName(Convert.ToInt32(rows[0]["FrontConfigID"]));
                    RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.MilanoWidth), 1, Front, false);
                }
            }

            DT.Clear();
            rows = PragaOrdersDT.Select(filter +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.PragaMargin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.PragaMargin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Praga));
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.PragaWidth), 1, Front, false);

            DT.Clear();
            rows = SigmaOrdersDT.Select(filter +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.SigmaMargin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.SigmaMargin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Sigma));
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.SigmaWidth), 1, Front, false);

            DT.Clear();
            rows = FatOrdersDT.Select(filter +
                " AND (Height>" + (Convert.ToInt32(FrontMargins.FatMargin) + 100) + " AND Width>" + (Convert.ToInt32(FrontMargins.FatMargin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Fat));
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.FatWidth), 1, Front, false);
        }

        private void CombineRapidFilenkaBoxes(ref DataTable DestinationDT)
        {
            string filter = @"InsetTypeID IN (2069,2070,2071,2073,2075,2077,2233,3644,29043,29531)";
            DataTable DT = AntaliaOrdersDT.Clone();
            DataRow[] rows = AntaliaOrdersDT.Select(filter +
                " AND (Height<=" + (Convert.ToInt32(FrontMargins.AntaliaMargin) + 100) + " OR Width<=" + (Convert.ToInt32(FrontMargins.AntaliaMargin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            string Front = GetFrontName(Convert.ToInt32(Fronts.Antalia)) + " ШУФ";
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.AntaliaWidth), 2, Front, false);

            DT.Clear();
            rows = TechnoNOrdersDT.Select(filter +
                " AND (Height<=" + (Convert.ToInt32(FrontMargins.TechnoNMargin) + 100) + " OR Width<=" + (Convert.ToInt32(FrontMargins.TechnoNMargin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.TechnoN)) + " ШУФ";
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.TechnoNWidth), 2, Front, false);

            DT.Clear();
            rows = Nord95OrdersDT.Select(filter +
                " AND (Height<=" + (Convert.ToInt32(FrontMargins.Nord95Margin) + 100) + " OR Width<=" + (Convert.ToInt32(FrontMargins.Nord95Margin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Nord95)) + " ШУФ";
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.Nord95Width), 2, Front, false);

            DT.Clear();
            rows = epFoxOrdersDT.Select(filter +
                " AND (Height<=" + (Convert.ToInt32(FrontMargins.epFoxMargin) + 100) + " OR Width<=" + (Convert.ToInt32(FrontMargins.epFoxMargin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.epFox)) + " ШУФ";
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.epFoxWidth), 2, Front, false);

            DT.Clear();
            rows = BergamoOrdersDT.Select(filter +
                " AND (Height<=" + (Convert.ToInt32(FrontMargins.BergamoMargin) + 100) + " OR Width<=" + (Convert.ToInt32(FrontMargins.BergamoMargin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Bergamo)) + " ШУФ";
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.BergamoWidth), 2, Front, false);

            DT.Clear();
            rows = ep041OrdersDT.Select(filter +
                " AND (Height<=" + (Convert.ToInt32(FrontMargins.ep041Margin) + 100) + " OR Width<=" + (Convert.ToInt32(FrontMargins.ep041Margin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep041)) + " ШУФ";
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep041Width), 2, Front, false);

            DT.Clear();
            rows = ep071OrdersDT.Select(filter +
                " AND (Height<=" + (Convert.ToInt32(FrontMargins.ep071Margin) + 100) + " OR Width<=" + (Convert.ToInt32(FrontMargins.ep071Margin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep071)) + " ШУФ";
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep071Width), 2, Front, false);

            DT.Clear();
            rows = ep206OrdersDT.Select(filter +
                " AND (Height<=" + (Convert.ToInt32(FrontMargins.ep206Margin) + 100) + " OR Width<=" + (Convert.ToInt32(FrontMargins.ep206Margin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep206)) + " ШУФ";
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep206Width), 2, Front, false);

            DT.Clear();
            rows = ep216OrdersDT.Select(filter +
                " AND (Height<=" + (Convert.ToInt32(FrontMargins.ep216Margin) + 100) + " OR Width<=" + (Convert.ToInt32(FrontMargins.ep216Margin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep216)) + " ШУФ";
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep216Width), 2, Front, false);

            DT.Clear();
            rows = ep111OrdersDT.Select(filter +
                " AND (Height<=" + (Convert.ToInt32(FrontMargins.ep111Margin) + 100) + " OR Width<=" + (Convert.ToInt32(FrontMargins.ep111Margin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep111)) + " ШУФ";
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep111Width), 2, Front, false);

            DT.Clear();
            rows = BostonOrdersDT.Select(filter +
                " AND (Height<=" + (Convert.ToInt32(FrontMargins.BostonMargin) + 100) + " OR Width<=" + (Convert.ToInt32(FrontMargins.BostonMargin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Boston)) + " ШУФ";
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.BostonWidth), 2, Front, false);

            DataTable TempDT = new DataTable();
            TempDT.Clear();
            using (DataView DV = new DataView(LeonOrdersDT))
            {
                TempDT = DV.ToTable(true, new string[] { "ProfileID" });
            }
            for (int i = 0; i < TempDT.Rows.Count; i++)
            {
                int ProfileID = Convert.ToInt32(TempDT.Rows[i]["ProfileID"]);
                DT.Clear();
                rows = LeonOrdersDT.Select(filter + " AND ProfileID=" + ProfileID +
                " AND (Height<=" + (Convert.ToInt32(FrontMargins.LeonMargin) + 100) + " OR Width<=" + (Convert.ToInt32(FrontMargins.LeonMargin) + 100) + ")");
                if (rows.Count() > 0)
                {
                    foreach (DataRow item in rows)
                        DT.Rows.Add(item.ItemArray);
                    Front = ProfileName(Convert.ToInt32(rows[0]["FrontConfigID"])) + " ШУФ";
                    RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.LeonWidth), 2, Front, false);
                }
            }

            DT.Clear();
            rows = LimogOrdersDT.Select(filter +
                " AND (Height<=" + (Convert.ToInt32(FrontMargins.LimogMargin) + 100) + " OR Width<=" + (Convert.ToInt32(FrontMargins.LimogMargin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Limog)) + " ШУФ";
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.LimogWidth), 2, Front, false);

            DT.Clear();
            rows = ep066Marsel4OrdersDT.Select(filter +
                " AND (Height<=" + (Convert.ToInt32(FrontMargins.ep066Marsel4Margin) + 100) + " OR Width<=" + (Convert.ToInt32(FrontMargins.ep066Marsel4Margin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep066Marsel4)) + " ШУФ";
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep066Marsel4Width), 2, Front, false);

            DT.Clear();
            rows = ep110JersyOrdersDT.Select(filter +
                " AND (Height<=" + (Convert.ToInt32(FrontMargins.ep110JersyMargin) + 100) + " OR Width<=" + (Convert.ToInt32(FrontMargins.ep110JersyMargin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep110Jersy)) + " ШУФ";
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep110JersyWidth), 2, Front, false);

            DT.Clear();
            rows = ep018Marsel1OrdersDT.Select(filter +
                " AND (Height<=" + (Convert.ToInt32(FrontMargins.ep018Marsel1Margin) + 100) + " OR Width<=" + (Convert.ToInt32(FrontMargins.ep018Marsel1Margin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep018Marsel1)) + " ШУФ";
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep018Marsel1Width), 2, Front, false);

            DT.Clear();
            rows = ep043ShervudOrdersDT.Select(filter +
                " AND (Height<=" + (Convert.ToInt32(FrontMargins.ep043ShervudMargin) + 100) + " OR Width<=" + (Convert.ToInt32(FrontMargins.ep043ShervudMargin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep043Shervud)) + " ШУФ";
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep043ShervudWidth), 2, Front, false);

            DT.Clear();
            rows = ep112OrdersDT.Select(filter +
                " AND (Height<=" + (Convert.ToInt32(FrontMargins.ep112Margin) + 100) + " OR Width<=" + (Convert.ToInt32(FrontMargins.ep112Margin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep112)) + " ШУФ";
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep112Width), 2, Front, false);

            DT.Clear();
            rows = UrbanOrdersDT.Select(filter +
                " AND (Height<=" + (Convert.ToInt32(FrontMargins.UrbanMargin) + 100) + " OR Width<=" + (Convert.ToInt32(FrontMargins.UrbanMargin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Urban)) + " ШУФ";
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.UrbanWidth), 2, Front, false);

            DT.Clear();
            rows = AlbyOrdersDT.Select(filter +
                " AND (Height<=" + (Convert.ToInt32(FrontMargins.AlbyMargin) + 100) + " OR Width<=" + (Convert.ToInt32(FrontMargins.AlbyMargin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Alby)) + " ШУФ";
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.AlbyWidth), 2, Front, false);

            DT.Clear();
            rows = BrunoOrdersDT.Select(filter +
                " AND (Height<=" + (Convert.ToInt32(FrontMargins.BrunoMargin) + 100) + " OR Width<=" + (Convert.ToInt32(FrontMargins.BrunoMargin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Bruno)) + " ШУФ";
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.BrunoWidth), 2, Front, false);

            DT.Clear();
            rows = epsh406Techno4OrdersDT.Select(filter +
                " AND (Height<=" + (Convert.ToInt32(FrontMargins.epsh406Techno4Margin) + 100) + " OR Width<=" + (Convert.ToInt32(FrontMargins.epsh406Techno4Margin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.epsh406Techno4)) + " ШУФ";
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.epsh406Techno4Width), 2, Front, false);

            DT.Clear();
            rows = LukOrdersDT.Select(filter +
                " AND (Height<=" + (Convert.ToInt32(FrontMargins.LukMargin) + 100) + " OR Width<=" + (Convert.ToInt32(FrontMargins.LukMargin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Luk)) + " ШУФ";
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.LukWidth), 2, Front, false);

            DT.Clear();
            rows = LukPVHOrdersDT.Select(filter +
                " AND (Height<=" + (Convert.ToInt32(FrontMargins.LukMargin) + 100) + " OR Width<=" + (Convert.ToInt32(FrontMargins.LukMargin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.LukPVH)) + " ШУФ";
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.LukWidth), 2, Front, false);

            TempDT.Clear();
            using (DataView DV = new DataView(MilanoOrdersDT))
            {
                TempDT = DV.ToTable(true, new string[] { "ProfileID" });
            }
            for (int i = 0; i < TempDT.Rows.Count; i++)
            {
                int ProfileID = Convert.ToInt32(TempDT.Rows[i]["ProfileID"]);
                DT.Clear();
                rows = MilanoOrdersDT.Select(filter + " AND ProfileID=" + ProfileID +
                " AND (Height<=" + (Convert.ToInt32(FrontMargins.MilanoMargin) + 100) + " OR Width<=" + (Convert.ToInt32(FrontMargins.MilanoMargin) + 100) + ")");
                if (rows.Count() > 0)
                {
                    foreach (DataRow item in rows)
                        DT.Rows.Add(item.ItemArray);
                    Front = ProfileName(Convert.ToInt32(rows[0]["FrontConfigID"])) + " ШУФ";
                    RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.MilanoWidth), 2, Front, false);
                }
            }

            //DT.Clear();
            //rows = MilanoOrdersDT.Select(filter +
            //    " AND (Height<=" + (Convert.ToInt32(FrontMargins.MilanoMargin) + 100) + " OR Width<=" + (Convert.ToInt32(FrontMargins.MilanoMargin) + 100) + ")");
            //foreach (DataRow item in rows)
            //    DT.Rows.Add(item.ItemArray);
            //Front = GetFrontName(Convert.ToInt32(Fronts.Milano)) + " ШУФ";
            //RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.MilanoWidth), 2, Front, false);

            DT.Clear();
            rows = PragaOrdersDT.Select(filter +
                " AND (Height<=" + (Convert.ToInt32(FrontMargins.PragaMargin) + 100) + " OR Width<=" + (Convert.ToInt32(FrontMargins.PragaMargin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Praga)) + " ШУФ";
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.PragaWidth), 2, Front, false);

            DT.Clear();
            rows = SigmaOrdersDT.Select(filter +
                " AND (Height<=" + (Convert.ToInt32(FrontMargins.SigmaMargin) + 100) + " OR Width<=" + (Convert.ToInt32(FrontMargins.SigmaMargin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Sigma)) + " ШУФ";
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.SigmaWidth), 2, Front, false);

            DT.Clear();
            rows = FatOrdersDT.Select(filter +
                " AND (Height<=" + (Convert.ToInt32(FrontMargins.FatMargin) + 100) + " OR Width<=" + (Convert.ToInt32(FrontMargins.FatMargin) + 100) + ")");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Fat)) + " ШУФ";
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.FatWidth), 2, Front, false);

            using (DataView DV = new DataView(DestinationDT.Copy()))
            {
                DV.Sort = "Front, Color, ProfileType, Height DESC";
                DestinationDT.Clear();
                DestinationDT = DV.ToTable();
            }

            for (int i = 0; i < DestinationDT.Rows.Count; i++)
            {
                if (Convert.ToInt32(DestinationDT.Rows[i]["VitrinaCount"]) > 0)
                    DestinationDT.Rows[i]["Notes"] = Convert.ToInt32(DestinationDT.Rows[i]["VitrinaCount"]) + " витр.";
                if (DestinationDT.Rows[i]["BoxCount"] != DBNull.Value && Convert.ToInt32(DestinationDT.Rows[i]["BoxCount"]) > 0)
                {
                    if (DestinationDT.Rows[i]["Notes"] != DBNull.Value)
                    {
                        if (DestinationDT.Rows[i]["Notes"].ToString().Length > 0)
                            DestinationDT.Rows[i]["Notes"] = DestinationDT.Rows[i]["Notes"].ToString() + ", " + Convert.ToInt32(DestinationDT.Rows[i]["BoxCount"]) + " шуф.";
                    }
                    else
                        DestinationDT.Rows[i]["Notes"] = Convert.ToInt32(DestinationDT.Rows[i]["BoxCount"]) + " шуф.";
                }
                if (DestinationDT.Rows[i]["ImpostCount"] != DBNull.Value && Convert.ToInt32(DestinationDT.Rows[i]["ImpostCount"]) > 0)
                {
                    if (DestinationDT.Rows[i]["Notes"] != DBNull.Value)
                    {
                        if (DestinationDT.Rows[i]["Notes"].ToString().Length > 0)
                            DestinationDT.Rows[i]["Notes"] = DestinationDT.Rows[i]["Notes"].ToString() + ", " + Convert.ToInt32(DestinationDT.Rows[i]["ImpostCount"]) + " имп.";
                    }
                    else
                        DestinationDT.Rows[i]["Notes"] = Convert.ToInt32(DestinationDT.Rows[i]["ImpostCount"]) + " имп.";
                }
                if (i == 0)
                    continue;
                if (Convert.ToInt32(DestinationDT.Rows[i]["ProfileType"]) == Convert.ToInt32(DestinationDT.Rows[i - 1]["ProfileType"]) &&
                    Convert.ToInt32(DestinationDT.Rows[i]["ColorType"]) == Convert.ToInt32(DestinationDT.Rows[i - 1]["ColorType"]) &&
                    Convert.ToInt32(DestinationDT.Rows[i]["FrontType"]) == Convert.ToInt32(DestinationDT.Rows[i - 1]["FrontType"]))
                {
                    DestinationDT.Rows[i]["Front"] = string.Empty;
                    DestinationDT.Rows[i]["Color"] = string.Empty;
                }
            }
        }

        private void CombineRapidGrids(ref DataTable DestinationDT)
        {
            if (AntaliaGridsDT.Rows.Count > 0)
            {
                string Front = GetFrontName(Convert.ToInt32(Fronts.Antalia)) + " РЕШ";
                RapidGridsSingly(AntaliaGridsDT, ref DestinationDT, Convert.ToInt32(FrontMargins.AntaliaWidth), 3, Front, false);
            }

            if (Nord95GridsDT.Rows.Count > 0)
            {
                string Front = GetFrontName(Convert.ToInt32(Fronts.Nord95)) + " РЕШ";
                RapidGridsSingly(Nord95GridsDT, ref DestinationDT, Convert.ToInt32(FrontMargins.Nord95Width), 3, Front, false);
            }
            if (epFoxGridsDT.Rows.Count > 0)
            {
                string Front = GetFrontName(Convert.ToInt32(Fronts.epFox)) + " РЕШ";
                RapidGridsSingly(epFoxGridsDT, ref DestinationDT, Convert.ToInt32(FrontMargins.epFoxWidth), 3, Front, false);
            }

            if (BergamoGridsDT.Rows.Count > 0)
            {
                string Front = GetFrontName(Convert.ToInt32(Fronts.Bergamo)) + " РЕШ";
                RapidGridsSingly(BergamoGridsDT, ref DestinationDT, Convert.ToInt32(FrontMargins.BergamoWidth), 3, Front, false);
            }

            if (ep041GridsDT.Rows.Count > 0)
            {
                string Front = GetFrontName(Convert.ToInt32(Fronts.ep041)) + " РЕШ";
                RapidGridsSingly(ep041GridsDT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep041Width), 3, Front, false);
            }

            if (ep071GridsDT.Rows.Count > 0)
            {
                string Front = GetFrontName(Convert.ToInt32(Fronts.ep071)) + " РЕШ";
                RapidGridsSingly(ep071GridsDT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep071Width), 3, Front, false);
            }

            if (ep206GridsDT.Rows.Count > 0)
            {
                string Front = GetFrontName(Convert.ToInt32(Fronts.ep206)) + " РЕШ";
                RapidGridsSingly(ep206GridsDT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep206Width), 3, Front, false);
            }

            if (ep216GridsDT.Rows.Count > 0)
            {
                string Front = GetFrontName(Convert.ToInt32(Fronts.ep216)) + " РЕШ";
                RapidGridsSingly(ep216GridsDT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep216Width), 3, Front, false);
            }

            if (ep111GridsDT.Rows.Count > 0)
            {
                string Front = GetFrontName(Convert.ToInt32(Fronts.ep111)) + " РЕШ";
                RapidGridsSingly(ep111GridsDT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep111Width), 3, Front, false);
            }

            if (BostonGridsDT.Rows.Count > 0)
            {
                string Front = GetFrontName(Convert.ToInt32(Fronts.Boston)) + " РЕШ";
                RapidGridsSingly(BostonGridsDT, ref DestinationDT, Convert.ToInt32(FrontMargins.BostonWidth), 3, Front, false);
            }

            DataTable DT = ep216OrdersDT.Clone();
            DataTable TempDT = new DataTable();
            TempDT.Clear();
            using (DataView DV = new DataView(VeneciaGridsDT))
            {
                TempDT = DV.ToTable(true, new string[] { "ProfileID" });
            }
            for (int i = 0; i < TempDT.Rows.Count; i++)
            {
                int ProfileID = Convert.ToInt32(TempDT.Rows[i]["ProfileID"]);
                DT.Clear();
                DataRow[] rows = VeneciaGridsDT.Select("ProfileID=" + ProfileID);
                if (rows.Count() > 0)
                {
                    foreach (DataRow item in rows)
                        DT.Rows.Add(item.ItemArray);
                    string Front = ProfileName(Convert.ToInt32(rows[0]["FrontConfigID"])) + " РЕШ";
                    RapidGridsSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.VeneciaWidth), 3, Front, false);
                }
            }

            TempDT.Clear();
            using (DataView DV = new DataView(LeonGridsDT))
            {
                TempDT = DV.ToTable(true, new string[] { "ProfileID" });
            }
            for (int i = 0; i < TempDT.Rows.Count; i++)
            {
                int ProfileID = Convert.ToInt32(TempDT.Rows[i]["ProfileID"]);
                DT.Clear();
                DataRow[] rows = LeonGridsDT.Select("ProfileID=" + ProfileID);
                if (rows.Count() > 0)
                {
                    foreach (DataRow item in rows)
                        DT.Rows.Add(item.ItemArray);
                    string Front = ProfileName(Convert.ToInt32(rows[0]["FrontConfigID"])) + " РЕШ";
                    RapidGridsSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.LeonWidth), 3, Front, false);
                }
            }

            if (LimogGridsDT.Rows.Count > 0)
            {
                string Front = GetFrontName(Convert.ToInt32(Fronts.Limog)) + " РЕШ";
                RapidGridsSingly(LimogGridsDT, ref DestinationDT, Convert.ToInt32(FrontMargins.LimogWidth), 3, Front, false);
            }
            if (ep066Marsel4GridsDT.Rows.Count > 0)
            {
                string Front = GetFrontName(Convert.ToInt32(Fronts.ep066Marsel4)) + " РЕШ";
                RapidGridsSingly(ep066Marsel4GridsDT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep066Marsel4Width), 3, Front, false);
            }
            if (ep110JersyGridsDT.Rows.Count > 0)
            {
                string Front = GetFrontName(Convert.ToInt32(Fronts.ep110Jersy)) + " РЕШ";
                RapidGridsSingly(ep110JersyGridsDT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep110JersyWidth), 3, Front, false);
            }
            if (ep018Marsel1GridsDT.Rows.Count > 0)
            {
                string Front = GetFrontName(Convert.ToInt32(Fronts.ep018Marsel1)) + " РЕШ";
                RapidGridsSingly(ep018Marsel1GridsDT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep018Marsel1Width), 3, Front, false);
            }
            if (ep043ShervudGridsDT.Rows.Count > 0)
            {
                string Front = GetFrontName(Convert.ToInt32(Fronts.ep043Shervud)) + " РЕШ";
                RapidGridsSingly(ep043ShervudGridsDT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep043ShervudWidth), 3, Front, false);
            }
            if (ep112GridsDT.Rows.Count > 0)
            {
                string Front = GetFrontName(Convert.ToInt32(Fronts.ep112)) + " РЕШ";
                RapidGridsSingly(ep112GridsDT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep112Width), 3, Front, false);
            }
            if (UrbanGridsDT.Rows.Count > 0)
            {
                string Front = GetFrontName(Convert.ToInt32(Fronts.Urban)) + " РЕШ";
                RapidGridsSingly(UrbanGridsDT, ref DestinationDT, Convert.ToInt32(FrontMargins.UrbanWidth), 3, Front, false);
            }
            if (AlbyGridsDT.Rows.Count > 0)
            {
                string Front = GetFrontName(Convert.ToInt32(Fronts.Alby)) + " РЕШ";
                RapidGridsSingly(AlbyGridsDT, ref DestinationDT, Convert.ToInt32(FrontMargins.AlbyWidth), 3, Front, false);
            }
            if (BrunoGridsDT.Rows.Count > 0)
            {
                string Front = GetFrontName(Convert.ToInt32(Fronts.Bruno)) + " РЕШ";
                RapidGridsSingly(BrunoGridsDT, ref DestinationDT, Convert.ToInt32(FrontMargins.BrunoWidth), 3, Front, false);
            }
            if (epsh406Techno4GridsDT.Rows.Count > 0)
            {
                string Front = GetFrontName(Convert.ToInt32(Fronts.epsh406Techno4)) + " РЕШ";
                RapidGridsSingly(epsh406Techno4GridsDT, ref DestinationDT, Convert.ToInt32(FrontMargins.epsh406Techno4Width), 3, Front, false);
            }
            if (LukGridsDT.Rows.Count > 0)
            {
                string Front = GetFrontName(Convert.ToInt32(Fronts.Luk)) + " РЕШ";
                RapidGridsSingly(LukGridsDT, ref DestinationDT, Convert.ToInt32(FrontMargins.LukWidth), 3, Front, false);
            }
            if (LukPVHGridsDT.Rows.Count > 0)
            {
                string Front = GetFrontName(Convert.ToInt32(Fronts.LukPVH)) + " РЕШ";
                RapidGridsSingly(LukPVHGridsDT, ref DestinationDT, Convert.ToInt32(FrontMargins.LukWidth), 3, Front, false);
            }
            TempDT.Clear();
            using (DataView DV = new DataView(MilanoGridsDT))
            {
                TempDT = DV.ToTable(true, new string[] { "ProfileID" });
            }
            for (int i = 0; i < TempDT.Rows.Count; i++)
            {
                int ProfileID = Convert.ToInt32(TempDT.Rows[i]["ProfileID"]);
                DT.Clear();
                DataRow[] rows = MilanoGridsDT.Select("ProfileID=" + ProfileID);
                if (rows.Count() > 0)
                {
                    foreach (DataRow item in rows)
                        DT.Rows.Add(item.ItemArray);
                    string Front = ProfileName(Convert.ToInt32(rows[0]["FrontConfigID"])) + " РЕШ";
                    RapidGridsSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.MilanoWidth), 3, Front, false);
                }
            }

            if (PragaGridsDT.Rows.Count > 0)
            {
                string Front = GetFrontName(Convert.ToInt32(Fronts.Praga)) + " РЕШ";
                RapidGridsSingly(PragaGridsDT, ref DestinationDT, Convert.ToInt32(FrontMargins.PragaWidth), 3, Front, false);
            }
            if (SigmaGridsDT.Rows.Count > 0)
            {
                string Front = GetFrontName(Convert.ToInt32(Fronts.Sigma)) + " РЕШ";
                RapidGridsSingly(SigmaGridsDT, ref DestinationDT, Convert.ToInt32(FrontMargins.SigmaWidth), 3, Front, false);
            }
            if (FatGridsDT.Rows.Count > 0)
            {
                string Front = GetFrontName(Convert.ToInt32(Fronts.Fat)) + " РЕШ";
                RapidGridsSingly(FatGridsDT, ref DestinationDT, Convert.ToInt32(FrontMargins.FatWidth), 3, Front, false);
            }
        }

        /// <summary>
        /// вкладка Мартин, нижняя таблица с импостом
        /// </summary>
        /// <param name="DestinationDT"></param>
        /// <param name="IsBox"></param>
        private void CombineRapidImpostSimple(ref DataTable DestinationDT, bool IsBox)
        {
            string filter = @"TechnoProfileID<>-1 AND InsetTypeID NOT IN (2069,2070,2071,2073,2075,2077,2233,3644,29043,29531) AND InsetTypeID NOT IN (860) AND InsetTypeID NOT IN (685,686,687,688,29470,29471)";

            DataTable DT = AntaliaOrdersDT.Clone();
            DataRow[] rows = AntaliaOrdersDT.Select(filter);
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            string Front = GetFrontName(Convert.ToInt32(Fronts.Antalia));
            RapidSimpleImpostSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.AntaliaWidth), 1, Front, false);

            if (!IsBox)
            {
                DT.Clear();
                rows = TechnoNOrdersDT.Select(filter +
                    " AND Height>" + (Convert.ToInt32(FrontMargins.TechnoNWidth)) + " AND Width>" + (Convert.ToInt32(FrontMargins.TechnoNWidth)));
                foreach (DataRow item in rows)
                    DT.Rows.Add(item.ItemArray);
                Front = GetFrontName(Convert.ToInt32(Fronts.TechnoN));
                RapidSimpleImpostSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.TechnoNWidth), 1, Front, false);
            }

            DT.Clear();
            rows = Nord95OrdersDT.Select(filter);
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Nord95));
            RapidSimpleImpostSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.Nord95Width), 1, Front, false);

            DT.Clear();
            rows = epFoxOrdersDT.Select(filter);
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.epFox));
            RapidSimpleImpostSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.epFoxWidth), 1, Front, false);

            DT.Clear();
            rows = BergamoOrdersDT.Select(filter + " AND (Height>150 OR Width>150)");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Bergamo));
            RapidSimpleImpostSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.BergamoWidth), 1, Front, false);

            DT.Clear();
            rows = BergamoOrdersDT.Select(filter + " AND (Height<=150 OR Width<=150)");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = "ПШ-050 22х55 Бергамо";
            RapidSimpleBergamoSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.BergamoWidth), 1, 29548, Front, false);

            DT.Clear();
            rows = ep041OrdersDT.Select(filter);
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep041));
            RapidSimpleImpostSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep041Width), 1, Front, false);

            DT.Clear();
            rows = ep071OrdersDT.Select(filter);
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep071));
            RapidSimpleImpostSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep071Width), 1, Front, false);

            DT.Clear();
            rows = ep206OrdersDT.Select(filter);
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep206));
            RapidSimpleImpostSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep206Width), 1, Front, false);

            DT.Clear();
            rows = ep216OrdersDT.Select(filter);
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep216));
            RapidSimpleImpostSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep216Width), 1, Front, false);

            DT.Clear();
            rows = ep111OrdersDT.Select(filter);
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep111));
            RapidSimpleImpostSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep111Width), 1, Front, false);

            DT.Clear();
            rows = BostonOrdersDT.Select(filter);
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Boston));
            RapidSimpleImpostSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.BostonWidth), 1, Front, false);

            DataTable TempDT = new DataTable();
            TempDT.Clear();
            using (DataView DV = new DataView(VeneciaOrdersDT))
            {
                TempDT = DV.ToTable(true, new string[] { "ProfileID" });
            }
            for (int i = 0; i < TempDT.Rows.Count; i++)
            {
                int ProfileID = Convert.ToInt32(TempDT.Rows[i]["ProfileID"]);
                DT.Clear();
                rows = VeneciaOrdersDT.Select(filter + " AND ProfileID=" + ProfileID);
                if (rows.Count() > 0)
                {
                    foreach (DataRow item in rows)
                        DT.Rows.Add(item.ItemArray);
                    Front = ProfileName(Convert.ToInt32(rows[0]["FrontConfigID"]));
                    RapidSimpleImpostSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.VeneciaWidth), 1, Front, false);
                }
            }

            TempDT.Clear();
            using (DataView DV = new DataView(LeonOrdersDT))
            {
                TempDT = DV.ToTable(true, new string[] { "ProfileID" });
            }
            for (int i = 0; i < TempDT.Rows.Count; i++)
            {
                int ProfileID = Convert.ToInt32(TempDT.Rows[i]["ProfileID"]);
                DT.Clear();
                rows = LeonOrdersDT.Select(filter + " AND ProfileID=" + ProfileID);
                if (rows.Count() > 0)
                {
                    foreach (DataRow item in rows)
                        DT.Rows.Add(item.ItemArray);
                    Front = ProfileName(Convert.ToInt32(rows[0]["FrontConfigID"]));
                    RapidSimpleImpostSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.LeonWidth), 1, Front, false);
                }
            }

            DT.Clear();
            rows = LimogOrdersDT.Select(filter);
            //AND Height>" + (Convert.ToInt32(FrontMargins.LimogWidth) - 10) + " AND Width>" + (Convert.ToInt32(FrontMargins.LimogWidth) - 10));
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Limog));
            RapidSimpleImpostSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.LimogWidth), 1, Front, false);

            DT.Clear();
            rows = ep066Marsel4OrdersDT.Select(filter);
            //AND Height>" + (Convert.ToInt32(FrontMargins.ep066Marsel4Width) - 10) + " AND Width>" + (Convert.ToInt32(FrontMargins.ep066Marsel4Width) - 10));
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep066Marsel4));
            RapidSimpleImpostSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep066Marsel4Width), 1, Front, false);

            DT.Clear();
            rows = ep110JersyOrdersDT.Select(filter);
            //AND Height>" + (Convert.ToInt32(FrontMargins.ep110JersyWidth) - 10) + " AND Width>" + (Convert.ToInt32(FrontMargins.ep110JersyWidth) - 10));
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep110Jersy));
            RapidSimpleImpostSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep110JersyWidth), 1, Front, false);

            DT.Clear();
            rows = ep018Marsel1OrdersDT.Select(filter);
            //AND Height>" + (Convert.ToInt32(FrontMargins.ep018Marsel1Width) - 10) + " AND Width>" + (Convert.ToInt32(FrontMargins.ep018Marsel1Width) - 10));
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep018Marsel1));
            RapidSimpleImpostSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep018Marsel1Width), 1, Front, false);

            DT.Clear();
            rows = ep043ShervudOrdersDT.Select(filter);
            //AND Height>" + (Convert.ToInt32(FrontMargins.ep043ShervudWidth) - 10) + " AND Width>" + (Convert.ToInt32(FrontMargins.ep043ShervudWidth) - 10));
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep043Shervud));
            RapidSimpleImpostSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep043ShervudWidth), 1, Front, false);

            DT.Clear();
            rows = ep112OrdersDT.Select(filter);
            //AND Height>" + (Convert.ToInt32(FrontMargins.ep112Width) - 10) + " AND Width>" + (Convert.ToInt32(FrontMargins.ep112Width) - 10));
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep112));
            RapidSimpleImpostSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep112Width), 1, Front, false);

            DT.Clear();
            rows = UrbanOrdersDT.Select(filter);
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Urban));
            RapidSimpleImpostSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.UrbanWidth), 1, Front, false);

            DT.Clear();
            rows = AlbyOrdersDT.Select(filter);
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Alby));
            RapidSimpleImpostSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.AlbyWidth), 1, Front, false);

            DT.Clear();
            rows = BrunoOrdersDT.Select(filter);
            //AND Height>" + (Convert.ToInt32(FrontMargins.BrunoWidth) - 10) + " AND Width>" + (Convert.ToInt32(FrontMargins.BrunoWidth) - 10));
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Bruno));
            RapidSimpleImpostSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.BrunoWidth), 1, Front, false);

            DT.Clear();
            rows = epsh406Techno4OrdersDT.Select(filter);
            //AND Height>" + (Convert.ToInt32(FrontMargins.epsh406Techno4Width) - 10) + " AND Width>" + (Convert.ToInt32(FrontMargins.epsh406Techno4Width) - 10));
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.epsh406Techno4));
            RapidSimpleImpostSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.epsh406Techno4Width), 1, Front, false);

            DT.Clear();
            rows = LukOrdersDT.Select(filter);
            //AND Height>" + (Convert.ToInt32(FrontMargins.LukWidth) - 10) + " AND Width>" + (Convert.ToInt32(FrontMargins.LukWidth) - 10));
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Luk));
            RapidSimpleImpostSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.LukWidth), 1, Front, false);

            DT.Clear();
            rows = LukPVHOrdersDT.Select(filter);
            //AND Height>" + (Convert.ToInt32(FrontMargins.LukWidth) - 10) + " AND Width>" + (Convert.ToInt32(FrontMargins.LukWidth) - 10));
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.LukPVH));
            RapidSimpleImpostSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.LukWidth), 1, Front, false);

            TempDT.Clear();
            using (DataView DV = new DataView(MilanoOrdersDT))
            {
                TempDT = DV.ToTable(true, new string[] { "ProfileID" });
            }
            for (int i = 0; i < TempDT.Rows.Count; i++)
            {
                int ProfileID = Convert.ToInt32(TempDT.Rows[i]["ProfileID"]);
                DT.Clear();
                rows = MilanoOrdersDT.Select(filter + " AND ProfileID=" + ProfileID);
                if (rows.Count() > 0)
                {
                    foreach (DataRow item in rows)
                        DT.Rows.Add(item.ItemArray);
                    Front = ProfileName(Convert.ToInt32(rows[0]["FrontConfigID"]));
                    RapidSimpleImpostSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.MilanoWidth), 1, Front, false);
                }
            }

            DT.Clear();
            rows = PragaOrdersDT.Select(filter);
            //AND Height>" + (Convert.ToInt32(FrontMargins.PragaWidth) - 10) + " AND Width>" + (Convert.ToInt32(FrontMargins.PragaWidth) - 10));
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Praga));
            RapidSimpleImpostSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.PragaWidth), 1, Front, false);

            DT.Clear();
            rows = SigmaOrdersDT.Select(filter);
            //AND Height>" + (Convert.ToInt32(FrontMargins.SigmaWidth) - 10) + " AND Width>" + (Convert.ToInt32(FrontMargins.SigmaWidth) - 10));
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Sigma));
            RapidSimpleImpostSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.SigmaWidth), 1, Front, false);

            DT.Clear();
            rows = FatOrdersDT.Select(filter);
            //AND Height>" + (Convert.ToInt32(FrontMargins.FatWidth) - 10) + " AND Width>" + (Convert.ToInt32(FrontMargins.FatWidth) - 10));
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Fat));
            RapidSimpleImpostSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.FatWidth), 1, Front, false);

            for (int i = 0; i < DestinationDT.Rows.Count; i++)
            {
                if (DestinationDT.Rows[i]["VitrinaCount"] != DBNull.Value && Convert.ToInt32(DestinationDT.Rows[i]["VitrinaCount"]) > 0)
                    DestinationDT.Rows[i]["Notes"] = Convert.ToInt32(DestinationDT.Rows[i]["VitrinaCount"]) + " витр.";
                if (DestinationDT.Rows[i]["BoxCount"] != DBNull.Value && Convert.ToInt32(DestinationDT.Rows[i]["BoxCount"]) > 0)
                {
                    if (DestinationDT.Rows[i]["Notes"] != DBNull.Value)
                    {
                        if (DestinationDT.Rows[i]["Notes"].ToString().Length > 0)
                            DestinationDT.Rows[i]["Notes"] = DestinationDT.Rows[i]["Notes"].ToString() + ", " + Convert.ToInt32(DestinationDT.Rows[i]["BoxCount"]) + " шуф.";
                    }
                    else
                        DestinationDT.Rows[i]["Notes"] = Convert.ToInt32(DestinationDT.Rows[i]["BoxCount"]) + " шуф.";
                }
                if (DestinationDT.Rows[i]["ImpostCount"] != DBNull.Value && Convert.ToInt32(DestinationDT.Rows[i]["ImpostCount"]) > 0)
                {
                    if (DestinationDT.Rows[i]["Notes"] != DBNull.Value)
                    {
                        if (DestinationDT.Rows[i]["Notes"].ToString().Length > 0)
                            DestinationDT.Rows[i]["Notes"] = DestinationDT.Rows[i]["Notes"].ToString() + ", " + Convert.ToInt32(DestinationDT.Rows[i]["ImpostCount"]) + " имп.";
                    }
                    else
                        DestinationDT.Rows[i]["Notes"] = Convert.ToInt32(DestinationDT.Rows[i]["ImpostCount"]) + " имп.";
                }
                if (i == 0)
                    continue;
                if (Convert.ToInt32(DestinationDT.Rows[i]["ProfileType"]) == Convert.ToInt32(DestinationDT.Rows[i - 1]["ProfileType"]) &&
                    Convert.ToInt32(DestinationDT.Rows[i]["ColorType"]) == Convert.ToInt32(DestinationDT.Rows[i - 1]["ColorType"]) &&
                    Convert.ToInt32(DestinationDT.Rows[i]["FrontType"]) == Convert.ToInt32(DestinationDT.Rows[i - 1]["FrontType"]))
                {
                    DestinationDT.Rows[i]["Front"] = string.Empty;
                    DestinationDT.Rows[i]["Color"] = string.Empty;
                }
            }
        }

        private void CombineRapidSimple(ref DataTable DestinationDT, bool IsBox)
        {
            string filter = @"InsetTypeID NOT IN (2069,2070,2071,2073,2075,2077,2233,3644,29043,29531) AND InsetTypeID NOT IN (860) AND InsetTypeID NOT IN (685,686,687,688,29470,29471)";

            DataTable DT = AntaliaOrdersDT.Clone();
            DataRow[] rows = AntaliaOrdersDT.Select(filter);
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            string Front = GetFrontName(Convert.ToInt32(Fronts.Antalia));
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.AntaliaWidth), 1, Front, false);

            if (!IsBox)
            {
                DT.Clear();
                rows = TechnoNOrdersDT.Select(filter +
                    " AND Height>" + (Convert.ToInt32(FrontMargins.TechnoNWidth)) + " AND Width>" + (Convert.ToInt32(FrontMargins.TechnoNWidth)));
                foreach (DataRow item in rows)
                    DT.Rows.Add(item.ItemArray);
                Front = GetFrontName(Convert.ToInt32(Fronts.TechnoN));
                RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.TechnoNWidth), 1, Front, false);
            }

            DT.Clear();
            rows = Nord95OrdersDT.Select(filter);
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Nord95));
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.Nord95Width), 1, Front, false);

            DT.Clear();
            rows = epFoxOrdersDT.Select(filter);
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.epFox));
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.epFoxWidth), 1, Front, false);

            DT.Clear();
            rows = BergamoOrdersDT.Select(filter + " AND (Height>150 OR Width>150)");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Bergamo));
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.BergamoWidth), 1, Front, false);

            DT.Clear();
            rows = BergamoOrdersDT.Select(filter + " AND (Height<=150 OR Width<=150)");
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = "ПШ-050 22х55 Бергамо";
            RapidSimpleBergamoSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.BergamoWidth), 1, 29548, Front, false);

            DT.Clear();
            rows = ep041OrdersDT.Select(filter);
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep041));
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep041Width), 1, Front, false);

            DT.Clear();
            rows = ep071OrdersDT.Select(filter);
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep071));
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep071Width), 1, Front, false);

            DT.Clear();
            rows = ep206OrdersDT.Select(filter);
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep206));
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep206Width), 1, Front, false);

            DT.Clear();
            rows = ep216OrdersDT.Select(filter);
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep216));
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep216Width), 1, Front, false);

            DT.Clear();
            rows = ep111OrdersDT.Select(filter);
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep111));
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep111Width), 1, Front, false);

            DT.Clear();
            rows = BostonOrdersDT.Select(filter);
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Boston));
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.BostonWidth), 1, Front, false);

            DataTable TempDT = new DataTable();
            TempDT.Clear();
            using (DataView DV = new DataView(VeneciaOrdersDT))
            {
                TempDT = DV.ToTable(true, new string[] { "ProfileID" });
            }
            for (int i = 0; i < TempDT.Rows.Count; i++)
            {
                int ProfileID = Convert.ToInt32(TempDT.Rows[i]["ProfileID"]);
                DT.Clear();
                rows = VeneciaOrdersDT.Select(filter + " AND ProfileID=" + ProfileID);
                if (rows.Count() > 0)
                {
                    foreach (DataRow item in rows)
                        DT.Rows.Add(item.ItemArray);
                    Front = ProfileName(Convert.ToInt32(rows[0]["FrontConfigID"]));
                    RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.VeneciaWidth), 1, Front, false);
                }
            }

            TempDT.Clear();
            using (DataView DV = new DataView(LeonOrdersDT))
            {
                TempDT = DV.ToTable(true, new string[] { "ProfileID" });
            }
            for (int i = 0; i < TempDT.Rows.Count; i++)
            {
                int ProfileID = Convert.ToInt32(TempDT.Rows[i]["ProfileID"]);
                DT.Clear();
                rows = LeonOrdersDT.Select(filter + " AND ProfileID=" + ProfileID);
                if (rows.Count() > 0)
                {
                    foreach (DataRow item in rows)
                        DT.Rows.Add(item.ItemArray);
                    Front = ProfileName(Convert.ToInt32(rows[0]["FrontConfigID"]));
                    RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.LeonWidth), 1, Front, false);
                }
            }

            DT.Clear();
            rows = LimogOrdersDT.Select(filter);
            //AND Height>" + (Convert.ToInt32(FrontMargins.LimogWidth) - 10) + " AND Width>" + (Convert.ToInt32(FrontMargins.LimogWidth) - 10));
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Limog));
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.LimogWidth), 1, Front, false);

            DT.Clear();
            rows = ep066Marsel4OrdersDT.Select(filter);
            //AND Height>" + (Convert.ToInt32(FrontMargins.ep066Marsel4Width) - 10) + " AND Width>" + (Convert.ToInt32(FrontMargins.ep066Marsel4Width) - 10));
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep066Marsel4));
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep066Marsel4Width), 1, Front, false);

            DT.Clear();
            rows = ep110JersyOrdersDT.Select(filter);
            //AND Height>" + (Convert.ToInt32(FrontMargins.ep110JersyWidth) - 10) + " AND Width>" + (Convert.ToInt32(FrontMargins.ep110JersyWidth) - 10));
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep110Jersy));
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep110JersyWidth), 1, Front, false);

            DT.Clear();
            rows = ep018Marsel1OrdersDT.Select(filter);
            //AND Height>" + (Convert.ToInt32(FrontMargins.ep018Marsel1Width) - 10) + " AND Width>" + (Convert.ToInt32(FrontMargins.ep018Marsel1Width) - 10));
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep018Marsel1));
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep018Marsel1Width), 1, Front, false);

            DT.Clear();
            rows = ep043ShervudOrdersDT.Select(filter);
            //AND Height>" + (Convert.ToInt32(FrontMargins.ep043ShervudWidth) - 10) + " AND Width>" + (Convert.ToInt32(FrontMargins.ep043ShervudWidth) - 10));
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep043Shervud));
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep043ShervudWidth), 1, Front, false);

            DT.Clear();
            rows = ep112OrdersDT.Select(filter);
            //AND Height>" + (Convert.ToInt32(FrontMargins.ep112Width) - 10) + " AND Width>" + (Convert.ToInt32(FrontMargins.ep112Width) - 10));
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.ep112));
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.ep112Width), 1, Front, false);

            DT.Clear();
            rows = UrbanOrdersDT.Select(filter);
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Urban));
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.UrbanWidth), 1, Front, false);

            DT.Clear();
            rows = AlbyOrdersDT.Select(filter);
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Alby));
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.AlbyWidth), 1, Front, false);

            DT.Clear();
            rows = BrunoOrdersDT.Select(filter);
            //AND Height>" + (Convert.ToInt32(FrontMargins.BrunoWidth) - 10) + " AND Width>" + (Convert.ToInt32(FrontMargins.BrunoWidth) - 10));
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Bruno));
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.BrunoWidth), 1, Front, false);

            DT.Clear();
            rows = epsh406Techno4OrdersDT.Select(filter);
            //AND Height>" + (Convert.ToInt32(FrontMargins.epsh406Techno4Width) - 10) + " AND Width>" + (Convert.ToInt32(FrontMargins.epsh406Techno4Width) - 10));
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.epsh406Techno4));
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.epsh406Techno4Width), 1, Front, false);

            DT.Clear();
            rows = LukOrdersDT.Select(filter);
            //AND Height>" + (Convert.ToInt32(FrontMargins.LukWidth) - 10) + " AND Width>" + (Convert.ToInt32(FrontMargins.LukWidth) - 10));
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Luk));
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.LukWidth), 1, Front, false);

            DT.Clear();
            rows = LukPVHOrdersDT.Select(filter);
            //AND Height>" + (Convert.ToInt32(FrontMargins.LukWidth) - 10) + " AND Width>" + (Convert.ToInt32(FrontMargins.LukWidth) - 10));
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.LukPVH));
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.LukWidth), 1, Front, false);

            TempDT.Clear();
            using (DataView DV = new DataView(MilanoOrdersDT))
            {
                TempDT = DV.ToTable(true, new string[] { "ProfileID" });
            }
            for (int i = 0; i < TempDT.Rows.Count; i++)
            {
                int ProfileID = Convert.ToInt32(TempDT.Rows[i]["ProfileID"]);
                DT.Clear();
                rows = MilanoOrdersDT.Select(filter + " AND ProfileID=" + ProfileID);
                if (rows.Count() > 0)
                {
                    foreach (DataRow item in rows)
                        DT.Rows.Add(item.ItemArray);
                    Front = ProfileName(Convert.ToInt32(rows[0]["FrontConfigID"]));
                    RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.MilanoWidth), 1, Front, false);
                }
            }

            DT.Clear();
            rows = PragaOrdersDT.Select(filter);
            //AND Height>" + (Convert.ToInt32(FrontMargins.PragaWidth) - 10) + " AND Width>" + (Convert.ToInt32(FrontMargins.PragaWidth) - 10));
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Praga));
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.PragaWidth), 1, Front, false);

            DT.Clear();
            rows = SigmaOrdersDT.Select(filter);
            //AND Height>" + (Convert.ToInt32(FrontMargins.SigmaWidth) - 10) + " AND Width>" + (Convert.ToInt32(FrontMargins.SigmaWidth) - 10));
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Sigma));
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.SigmaWidth), 1, Front, false);

            DT.Clear();
            rows = FatOrdersDT.Select(filter);
            //AND Height>" + (Convert.ToInt32(FrontMargins.FatWidth) - 10) + " AND Width>" + (Convert.ToInt32(FrontMargins.FatWidth) - 10));
            foreach (DataRow item in rows)
                DT.Rows.Add(item.ItemArray);
            Front = GetFrontName(Convert.ToInt32(Fronts.Fat));
            RapidSimpleSingly(DT, ref DestinationDT, Convert.ToInt32(FrontMargins.FatWidth), 1, Front, false);

            CombineRapidGrids(ref RapidDT);
            //CombineRapidVitrina(ref RapidDT);
        }

        private void ComecSimpleSingly(DataTable SourceDT, ref DataTable DestinationDT, int WidthMargin, int FrontType, string Front, bool OrderASC)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = DistSizesTable(SourceDT, OrderASC);

            using (DataView DV = new DataView(SourceDT))
            {
                DV.RowFilter = "TechnoColorID <> -1";
                DT1 = DV.ToTable(true, new string[] { "TechnoColorID" });
            }
            DT1 = OrderedTechnoFrameColors(DT1);

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    DataRow[] Srows = SourceDT.Select("TechnoColorID=" + Convert.ToInt32(DT1.Rows[i]["TechnoColorID"]) +
                        " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) + " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]));
                    if (Srows.Count() > 0)
                    {
                        int Count = 0;
                        int BoxCount = 0;
                        int VitrinaCount = 0;
                        int Height = Convert.ToInt32(DT2.Rows[j]["Height"]);
                        int Width = Convert.ToInt32(DT2.Rows[j]["Width"]);
                        int count = StandardImpostCount(Convert.ToInt32(Srows[0]["FrontID"]), Height, Width);

                        //Front = ProfileName(Convert.ToInt32(SourceDT.Rows[0]["FrontConfigID"]), Convert.ToInt32(SourceDT.Rows[0]["TechnoProfileID"]));
                        string Color = GetColorName(Convert.ToInt32(DT1.Rows[i]["TechnoColorID"]));
                        string Notes = string.Empty;
                        foreach (DataRow item in Srows)
                        {
                            if (Convert.ToInt32(item["InsetTypeID"]) != 1 && (Convert.ToInt32(item["Height"]) - 10 < WidthMargin || Convert.ToInt32(item["Width"]) - 10 < WidthMargin))
                                BoxCount += Convert.ToInt32(item["Count"]);
                            if (Convert.ToInt32(item["InsetTypeID"]) == 1)
                                VitrinaCount += Convert.ToInt32(item["Count"]);

                            Count += Convert.ToInt32(item["Count"]);
                        }

                        //if (Height <= WidthMargin)
                        //    Height = WidthMargin;

                        DataRow[] rows = DestinationDT.Select("Front='" + Front + "' AND Color='" + Color + "' AND Height=" + Height + " AND Width=" + Width);
                        if (rows.Count() == 0)
                        {
                            DataRow NewRow = DestinationDT.NewRow();
                            NewRow["Front"] = Front;
                            NewRow["Color"] = Color;
                            NewRow["Height"] = Height;
                            NewRow["Width"] = Width;
                            NewRow["Count"] = Count;
                            NewRow["BoxCount"] = BoxCount * 2;
                            NewRow["VitrinaCount"] = VitrinaCount * 2;
                            NewRow["ImpostCount"] = count;
                            NewRow["ProfileType"] = Convert.ToInt32(SourceDT.Rows[0]["FrontID"]);
                            NewRow["ColorType"] = Convert.ToInt32(DT1.Rows[i]["TechnoColorID"]);
                            NewRow["FrontType"] = FrontType;
                            DestinationDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count;
                            rows[0]["BoxCount"] = Convert.ToInt32(rows[0]["BoxCount"]) + BoxCount * 2;
                            rows[0]["VitrinaCount"] = Convert.ToInt32(rows[0]["VitrinaCount"]) + VitrinaCount * 2;
                            rows[0]["ImpostCount"] = count;
                        }
                    }
                }
            }
        }

        private bool Contains(string source, string toCheck, StringComparison comp)
        {
            return source.IndexOf(toCheck, comp) >= 0;
        }

        private void Create()
        {
            DistMainOrdersDT = new DataTable();
            DistMainOrdersDT.Columns.Add(new DataColumn("MainOrderID", Type.GetType("System.Int32")));
            DistMainOrdersDT.Columns.Add(new DataColumn("GroupType", Type.GetType("System.Int32")));

            ProfileNamesDT = new DataTable();

            TechnoNVitrinaDT = new DataTable();
            TechnoNGridsDT = new DataTable();
            TechnoNSimpleDT = new DataTable();
            TechnoNOrdersDT = new DataTable();

            AntaliaVitrinaDT = new DataTable();
            AntaliaGridsDT = new DataTable();
            AntaliaSimpleDT = new DataTable();
            AntaliaOrdersDT = new DataTable();

            Nord95VitrinaDT = new DataTable();
            Nord95GridsDT = new DataTable();
            Nord95SimpleDT = new DataTable();
            Nord95OrdersDT = new DataTable();

            epFoxVitrinaDT = new DataTable();
            epFoxGridsDT = new DataTable();
            epFoxSimpleDT = new DataTable();
            epFoxOrdersDT = new DataTable();

            VeneciaVitrinaDT = new DataTable();
            VeneciaGridsDT = new DataTable();
            VeneciaSimpleDT = new DataTable();
            VeneciaOrdersDT = new DataTable();

            BergamoVitrinaDT = new DataTable();
            BergamoGridsDT = new DataTable();
            BergamoSimpleDT = new DataTable();
            BergamoOrdersDT = new DataTable();
            BergamoGlassDT = new DataTable();

            ep071VitrinaDT = new DataTable();
            ep071GridsDT = new DataTable();
            ep071SimpleDT = new DataTable();
            ep071OrdersDT = new DataTable();

            ep206VitrinaDT = new DataTable();
            ep206GridsDT = new DataTable();
            ep206SimpleDT = new DataTable();
            ep206OrdersDT = new DataTable();

            ep041VitrinaDT = new DataTable();
            ep041GridsDT = new DataTable();
            ep041SimpleDT = new DataTable();
            ep041OrdersDT = new DataTable();

            ep216VitrinaDT = new DataTable();
            ep216GridsDT = new DataTable();
            ep216SimpleDT = new DataTable();
            ep216OrdersDT = new DataTable();

            ep111VitrinaDT = new DataTable();
            ep111GridsDT = new DataTable();
            ep111SimpleDT = new DataTable();
            ep111OrdersDT = new DataTable();

            BostonVitrinaDT = new DataTable();
            BostonGridsDT = new DataTable();
            BostonSimpleDT = new DataTable();
            BostonOrdersDT = new DataTable();

            LeonVitrinaDT = new DataTable();
            LeonGridsDT = new DataTable();
            LeonSimpleDT = new DataTable();
            LeonOrdersDT = new DataTable();

            LimogVitrinaDT = new DataTable();
            LimogGridsDT = new DataTable();
            LimogSimpleDT = new DataTable();
            LimogOrdersDT = new DataTable();

            ep066Marsel4VitrinaDT = new DataTable();
            ep066Marsel4GridsDT = new DataTable();
            ep066Marsel4SimpleDT = new DataTable();
            ep066Marsel4OrdersDT = new DataTable();

            ep110JersyVitrinaDT = new DataTable();
            ep110JersyGridsDT = new DataTable();
            ep110JersySimpleDT = new DataTable();
            ep110JersyOrdersDT = new DataTable();

            ep018Marsel1VitrinaDT = new DataTable();
            ep018Marsel1GridsDT = new DataTable();
            ep018Marsel1SimpleDT = new DataTable();
            ep018Marsel1OrdersDT = new DataTable();

            ep043ShervudVitrinaDT = new DataTable();
            ep043ShervudGridsDT = new DataTable();
            ep043ShervudSimpleDT = new DataTable();
            ep043ShervudOrdersDT = new DataTable();

            ep112VitrinaDT = new DataTable();
            ep112GridsDT = new DataTable();
            ep112SimpleDT = new DataTable();
            ep112OrdersDT = new DataTable();

            UrbanVitrinaDT = new DataTable();
            UrbanGridsDT = new DataTable();
            UrbanSimpleDT = new DataTable();
            UrbanOrdersDT = new DataTable();

            AlbyVitrinaDT = new DataTable();
            AlbyGridsDT = new DataTable();
            AlbySimpleDT = new DataTable();
            AlbyOrdersDT = new DataTable();

            BrunoVitrinaDT = new DataTable();
            BrunoGridsDT = new DataTable();
            BrunoSimpleDT = new DataTable();
            BrunoOrdersDT = new DataTable();

            epsh406Techno4VitrinaDT = new DataTable();
            epsh406Techno4GridsDT = new DataTable();
            epsh406Techno4SimpleDT = new DataTable();
            epsh406Techno4OrdersDT = new DataTable();

            LukVitrinaDT = new DataTable();
            LukGridsDT = new DataTable();
            LukSimpleDT = new DataTable();
            LukOrdersDT = new DataTable();

            LukPVHVitrinaDT = new DataTable();
            LukPVHGridsDT = new DataTable();
            LukPVHSimpleDT = new DataTable();
            LukPVHOrdersDT = new DataTable();

            MilanoVitrinaDT = new DataTable();
            MilanoGridsDT = new DataTable();
            MilanoSimpleDT = new DataTable();
            MilanoOrdersDT = new DataTable();

            PragaVitrinaDT = new DataTable();
            PragaGridsDT = new DataTable();
            PragaSimpleDT = new DataTable();
            PragaOrdersDT = new DataTable();

            SigmaVitrinaDT = new DataTable();
            SigmaGlassDT = new DataTable();
            BostonGlassDT = new DataTable();
            SigmaGridsDT = new DataTable();
            SigmaSimpleDT = new DataTable();
            SigmaOrdersDT = new DataTable();

            FatVitrinaDT = new DataTable();
            FatGridsDT = new DataTable();
            FatSimpleDT = new DataTable();
            FatOrdersDT = new DataTable();

            BagetWithAngelOrdersDT = new DataTable();
            NotArchDecorOrdersDT = new DataTable();
            ArchDecorOrdersDT = new DataTable();
            GridsDecorOrdersDT = new DataTable();

            DeyingDT = new DataTable();
            DeyingDT.Columns.Add(new DataColumn("Name", Type.GetType("System.String")));
            DeyingDT.Columns.Add(new DataColumn("FrameColor", Type.GetType("System.String")));
            DeyingDT.Columns.Add(new DataColumn("InsetColor", Type.GetType("System.String")));
            DeyingDT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            DeyingDT.Columns.Add(new DataColumn("Width", Type.GetType("System.Int32")));
            DeyingDT.Columns.Add(new DataColumn("Count", Type.GetType("System.Int32")));
            DeyingDT.Columns.Add(new DataColumn("Square", Type.GetType("System.Decimal")));
            DeyingDT.Columns.Add(new DataColumn("Notes", Type.GetType("System.String")));
            DeyingDT.Columns.Add(new DataColumn("ColorType", Type.GetType("System.Int32")));

            DecorAssemblyDT = new DataTable();
            DecorAssemblyDT.Columns.Add(new DataColumn("Name", Type.GetType("System.String")));
            DecorAssemblyDT.Columns.Add(new DataColumn("Color", Type.GetType("System.String")));
            DecorAssemblyDT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            DecorAssemblyDT.Columns.Add(new DataColumn("Width", Type.GetType("System.Int32")));
            DecorAssemblyDT.Columns.Add(new DataColumn("Count", Type.GetType("System.Int32")));
            DecorAssemblyDT.Columns.Add(new DataColumn("Notes", Type.GetType("System.String")));
            DecorAssemblyDT.Columns.Add(new DataColumn("DecorID", Type.GetType("System.Int32")));

            BagetWithAngleAssemblyDT = new DataTable();
            BagetWithAngleAssemblyDT.Columns.Add(new DataColumn("Name", Type.GetType("System.String")));
            BagetWithAngleAssemblyDT.Columns.Add(new DataColumn("Color", Type.GetType("System.String")));
            BagetWithAngleAssemblyDT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            BagetWithAngleAssemblyDT.Columns.Add(new DataColumn("Width", Type.GetType("System.Int32")));
            BagetWithAngleAssemblyDT.Columns.Add(new DataColumn("LeftAngle", Type.GetType("System.Decimal")));
            BagetWithAngleAssemblyDT.Columns.Add(new DataColumn("RightAngle", Type.GetType("System.Decimal")));
            BagetWithAngleAssemblyDT.Columns.Add(new DataColumn("Count", Type.GetType("System.Int32")));
            BagetWithAngleAssemblyDT.Columns.Add(new DataColumn("Notes", Type.GetType("System.String")));
            BagetWithAngleAssemblyDT.Columns.Add(new DataColumn("DecorID", Type.GetType("System.Int32")));

            Additional1DT = new DataTable();
            Additional1DT.Columns.Add(new DataColumn("Profile", Type.GetType("System.String")));
            Additional1DT.Columns.Add(new DataColumn("Color", Type.GetType("System.String")));
            Additional1DT.Columns.Add(new DataColumn("InsetColor", Type.GetType("System.String")));
            Additional1DT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            Additional1DT.Columns.Add(new DataColumn("Width", Type.GetType("System.Int32")));
            Additional1DT.Columns.Add(new DataColumn("Count", Type.GetType("System.Int32")));
            Additional1DT.Columns.Add(new DataColumn("Square", Type.GetType("System.Decimal")));
            Additional1DT.Columns.Add(new DataColumn("Worker", Type.GetType("System.String")));
            Additional1DT.Columns.Add(new DataColumn("ColorID", Type.GetType("System.Int32")));
            Additional1DT.Columns.Add(new DataColumn("TechnoInsetColorID", Type.GetType("System.Int32")));
            Additional2DT = Additional1DT.Clone();
            Additional3DT = Additional1DT.Clone();
            Additional4DT = Additional1DT.Clone();
            Additional5DT = Additional1DT.Clone();

            RapidDT = new DataTable();
            RapidDT.Columns.Add(new DataColumn("Front", Type.GetType("System.String")));
            RapidDT.Columns.Add(new DataColumn("Color", Type.GetType("System.String")));
            RapidDT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            RapidDT.Columns.Add(new DataColumn("Count", Type.GetType("System.Int32")));
            RapidDT.Columns.Add(new DataColumn("Notes", Type.GetType("System.String")));
            RapidDT.Columns.Add(new DataColumn("BoxCount", Type.GetType("System.Int32")));
            RapidDT.Columns.Add(new DataColumn("VitrinaCount", Type.GetType("System.Int32")));
            RapidDT.Columns.Add(new DataColumn("ImpostCount", Type.GetType("System.Int32")));
            RapidDT.Columns.Add(new DataColumn("ProfileType", Type.GetType("System.Int32")));
            RapidDT.Columns.Add(new DataColumn("ColorType", Type.GetType("System.Int32")));
            RapidDT.Columns.Add(new DataColumn("FrontType", Type.GetType("System.Int32")));

            ComecDT = new DataTable();
            ComecDT.Columns.Add(new DataColumn("Front", Type.GetType("System.String")));
            ComecDT.Columns.Add(new DataColumn("Color", Type.GetType("System.String")));
            ComecDT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            ComecDT.Columns.Add(new DataColumn("Width", Type.GetType("System.Int32")));
            ComecDT.Columns.Add(new DataColumn("Count", Type.GetType("System.Int32")));
            ComecDT.Columns.Add(new DataColumn("Notes", Type.GetType("System.String")));
            ComecDT.Columns.Add(new DataColumn("BoxCount", Type.GetType("System.Int32")));
            ComecDT.Columns.Add(new DataColumn("VitrinaCount", Type.GetType("System.Int32")));
            ComecDT.Columns.Add(new DataColumn("ImpostCount", Type.GetType("System.Int32")));
            ComecDT.Columns.Add(new DataColumn("ProfileType", Type.GetType("System.Int32")));
            ComecDT.Columns.Add(new DataColumn("ColorType", Type.GetType("System.Int32")));
            ComecDT.Columns.Add(new DataColumn("FrontType", Type.GetType("System.Int32")));

            InsetDT = new DataTable();
            InsetDT.Columns.Add(new DataColumn("Name", Type.GetType("System.String")));
            InsetDT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            InsetDT.Columns.Add(new DataColumn("Width", Type.GetType("System.Int32")));
            InsetDT.Columns.Add(new DataColumn("Count", Type.GetType("System.Int32")));
            InsetDT.Columns.Add(new DataColumn("GlassCount", Type.GetType("System.Int32")));
            InsetDT.Columns.Add(new DataColumn("MegaCount", Type.GetType("System.Int32")));
            InsetDT.Columns.Add(new DataColumn("FrontID", Type.GetType("System.Int32")));
            InsetDT.Columns.Add(new DataColumn("InsetColorID", Type.GetType("System.Int32")));
            InsetDT.Columns.Add(new DataColumn("TechnoInsetColorID", Type.GetType("System.Int32")));

            AssemblyDT = new DataTable();
            AssemblyDT.Columns.Add(new DataColumn("Name", Type.GetType("System.String")));
            AssemblyDT.Columns.Add(new DataColumn("FrameColor", Type.GetType("System.String")));
            AssemblyDT.Columns.Add(new DataColumn("InsetColor", Type.GetType("System.String")));
            AssemblyDT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            AssemblyDT.Columns.Add(new DataColumn("Width", Type.GetType("System.Int32")));
            AssemblyDT.Columns.Add(new DataColumn("Count", Type.GetType("System.Int32")));
            AssemblyDT.Columns.Add(new DataColumn("Square", Type.GetType("System.Decimal")));
            AssemblyDT.Columns.Add(new DataColumn("ColorType", Type.GetType("System.Int32")));

            FrontsOrdersDT = new DataTable();
            FrontsOrdersDT.Columns.Add(new DataColumn("Name", Type.GetType("System.String")));
            FrontsOrdersDT.Columns.Add(new DataColumn("FrameColor", Type.GetType("System.String")));
            FrontsOrdersDT.Columns.Add(new DataColumn("InsetColor", Type.GetType("System.String")));
            FrontsOrdersDT.Columns.Add(new DataColumn("TechnoColor", Type.GetType("System.String")));
            FrontsOrdersDT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            FrontsOrdersDT.Columns.Add(new DataColumn("Width", Type.GetType("System.Int32")));
            FrontsOrdersDT.Columns.Add(new DataColumn("Count", Type.GetType("System.Int32")));
            FrontsOrdersDT.Columns.Add(new DataColumn("Square", Type.GetType("System.Decimal")));

            SummOrdersDT = new DataTable();

            FrontsID = new ArrayList();
        }

        private void DeyingByMainOrderToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            int WorkAssignmentID, string BatchName)
        {
            DistMainOrdersDT.Clear();

            DistMainOrdersTable(TechnoNOrdersDT);
            DistMainOrdersTable(AntaliaOrdersDT);
            DistMainOrdersTable(Nord95OrdersDT);
            DistMainOrdersTable(epFoxOrdersDT);
            DistMainOrdersTable(VeneciaOrdersDT);
            DistMainOrdersTable(BergamoOrdersDT);
            DistMainOrdersTable(ep041OrdersDT);
            DistMainOrdersTable(ep071OrdersDT);
            DistMainOrdersTable(ep206OrdersDT);
            DistMainOrdersTable(ep216OrdersDT);
            DistMainOrdersTable(ep111OrdersDT);
            DistMainOrdersTable(BostonOrdersDT);
            DistMainOrdersTable(LeonOrdersDT);
            DistMainOrdersTable(LimogOrdersDT);
            DistMainOrdersTable(LukOrdersDT);
            DistMainOrdersTable(LukPVHOrdersDT);
            DistMainOrdersTable(MilanoOrdersDT);
            DistMainOrdersTable(ep066Marsel4OrdersDT);
            DistMainOrdersTable(ep110JersyOrdersDT);
            DistMainOrdersTable(ep018Marsel1OrdersDT);
            DistMainOrdersTable(ep043ShervudOrdersDT);
            DistMainOrdersTable(ep112OrdersDT);
            DistMainOrdersTable(UrbanOrdersDT);
            DistMainOrdersTable(AlbyOrdersDT);
            DistMainOrdersTable(BrunoOrdersDT);
            DistMainOrdersTable(epsh406Techno4OrdersDT);
            DistMainOrdersTable(PragaOrdersDT);
            DistMainOrdersTable(SigmaOrdersDT);
            DistMainOrdersTable(FatOrdersDT);

            using (DataView DV = new DataView(DistMainOrdersDT.Copy()))
            {
                DistMainOrdersDT.Clear();
                DV.Sort = "MainOrderID ASC";
                DistMainOrdersDT = DV.ToTable(true, new string[] { "MainOrderID", "GroupType" });
            }

            DataTable DT = TechnoNOrdersDT.Clone();
            DataTable DT1 = new DataTable();

            DataTable MarketOrdersNames = new DataTable();
            string MainOrdersID = string.Empty;
            string ClientName = string.Empty;
            string OrderName = string.Empty;

            if (DistMainOrdersDT.Select("GroupType=1").Count() > 0)
            {
                MainOrdersID = string.Empty;
                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=1"))
                    MainOrdersID += Convert.ToInt32(item["MainOrderID"]) + ",";
                MainOrdersID = MainOrdersID.Substring(0, MainOrdersID.Length - 1);
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT MegaOrders.ClientID, ClientName, MegaOrders.OrderNumber, MainOrders.MainOrderID, MainOrders.Notes, Batch.MegaBatchID, Batch.BatchID FROM MainOrders" +
                    " INNER JOIN BatchDetails ON MainOrders.MainOrderID=BatchDetails.MainOrderID" +
                    " INNER JOIN Batch ON BatchDetails.BatchID=Batch.BatchID" +
                    " INNER JOIN MegaOrders ON MainOrders.MegaOrderID=MegaOrders.MegaOrderID" +
                    " INNER JOIN infiniu2_marketingreference.dbo.Clients ON MegaOrders.ClientID=infiniu2_marketingreference.dbo.Clients.ClientID" +
                    " WHERE MainOrders.MainOrderID IN (" + MainOrdersID + ")", ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DA.Fill(MarketOrdersNames);
                }
            }

            if (DistMainOrdersDT.Select("GroupType=1").Count() > 0)
            {
                int RowIndex = 0;

                HSSFSheet sheet1 = hssfworkbook.CreateSheet("Маркет");
                sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

                sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
                sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
                sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
                sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

                sheet1.SetColumnWidth(0, 25 * 256);
                sheet1.SetColumnWidth(1, 20 * 256);
                sheet1.SetColumnWidth(2, 20 * 256);
                sheet1.SetColumnWidth(3, 6 * 256);
                sheet1.SetColumnWidth(4, 6 * 256);
                sheet1.SetColumnWidth(5, 6 * 256);

                {
                    HSSFPatriarch patriarch = sheet1.CreateDrawingPatriarch();
                    int ind = 0;

                    CreateBarcode(WorkAssignmentID);
                    HSSFClientAnchor anchor = new HSSFClientAnchor(0, 1, 1000, 8, 5, ind, 6, ind + 3)
                    {
                        AnchorType = 2
                    };
                    string BarcodeNumber = GetBarcodeNumber(24, WorkAssignmentID);
                    string barcodeName = System.Environment.GetEnvironmentVariable("TEMP") + @"/" + BarcodeNumber + ".jpg";
                    HSSFPicture picture = patriarch.CreatePicture(anchor, LoadImage(barcodeName, hssfworkbook));

                    HSSFFont BarcodeFont = hssfworkbook.CreateFont();
                    BarcodeFont.FontHeightInPoints = 12;
                    BarcodeFont.FontName = "Calibri";

                    HSSFCellStyle BarcodeStyle = hssfworkbook.CreateCellStyle();
                    BarcodeStyle.Alignment = HSSFCellStyle.ALIGN_CENTER;
                    BarcodeStyle.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
                    BarcodeStyle.SetFont(BarcodeFont);

                    HSSFCell barcodeCell = sheet1.CreateRow(ind + 3).CreateCell(5);
                    barcodeCell.SetCellValue(BarcodeNumber.ToString());
                    barcodeCell.CellStyle = BarcodeStyle;
                    sheet1.AddMergedRegion(new NPOI.HSSF.Util.Region(ind + 3, 5, ind + 3, 6));
                }

                for (int i = 0; i < DistMainOrdersDT.Rows.Count; i++)
                {
                    if (Convert.ToInt32(DistMainOrdersDT.Rows[i]["GroupType"]) == 0)
                        continue;
                    int MainOrderID = Convert.ToInt32(DistMainOrdersDT.Rows[i]["MainOrderID"]);
                    int MegaBatchID = 0;
                    int BatchID = 0;
                    string Notes = string.Empty;

                    DeyingDT.Clear();
                    using (DataView DV = new DataView(TechnoNOrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = TechnoNSimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = TechnoNGridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                    }

                    using (DataView DV = new DataView(AntaliaOrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = AntaliaSimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = AntaliaGridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                    }

                    using (DataView DV = new DataView(Nord95OrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = Nord95SimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = Nord95GridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                    }

                    using (DataView DV = new DataView(epFoxOrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = epFoxSimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = epFoxGridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                    }

                    using (DataView DV = new DataView(VeneciaOrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = VeneciaSimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = VeneciaGridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                    }

                    using (DataView DV = new DataView(BergamoOrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = BergamoSimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = BergamoGridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");

                        DT.Clear();
                        rows = BergamoGlassDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);
                    }

                    using (DataView DV = new DataView(ep041OrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = ep041SimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = ep041GridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                    }

                    using (DataView DV = new DataView(ep071OrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = ep071SimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = ep071GridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                    }

                    using (DataView DV = new DataView(ep206OrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = ep206SimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = ep206GridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                    }

                    using (DataView DV = new DataView(ep216OrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = ep216SimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = ep216GridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                    }

                    using (DataView DV = new DataView(ep111OrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = ep111SimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = ep111GridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                    }

                    using (DataView DV = new DataView(BostonOrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = BostonSimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = BostonGridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");

                        DT.Clear();
                        rows = BostonGlassDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);
                    }

                    using (DataView DV = new DataView(LeonOrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = LeonSimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = LeonGridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                    }

                    using (DataView DV = new DataView(LimogOrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = LimogSimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = LimogGridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                    }

                    using (DataView DV = new DataView(LukOrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = LukSimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = LukGridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                    }

                    using (DataView DV = new DataView(LukPVHOrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = LukPVHSimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = LukPVHGridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                    }

                    using (DataView DV = new DataView(MilanoOrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = MilanoSimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = MilanoGridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                    }

                    using (DataView DV = new DataView(ep066Marsel4OrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = ep066Marsel4SimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = ep066Marsel4GridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                    }

                    using (DataView DV = new DataView(ep110JersyOrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = ep110JersySimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = ep110JersyGridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                    }

                    using (DataView DV = new DataView(ep018Marsel1OrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = ep018Marsel1SimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = ep018Marsel1GridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                    }

                    using (DataView DV = new DataView(ep043ShervudOrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = ep043ShervudSimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = ep043ShervudGridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                    }

                    using (DataView DV = new DataView(ep112OrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = ep112SimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = ep112GridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                    }

                    using (DataView DV = new DataView(UrbanOrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = UrbanSimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = UrbanGridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                    }

                    using (DataView DV = new DataView(BrunoOrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = BrunoSimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = BrunoGridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                    }

                    using (DataView DV = new DataView(epsh406Techno4OrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = epsh406Techno4SimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = epsh406Techno4GridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                    }

                    using (DataView DV = new DataView(PragaOrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = PragaSimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = PragaGridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                    }

                    using (DataView DV = new DataView(SigmaOrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = SigmaSimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = SigmaGridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");

                        DT.Clear();
                        rows = SigmaGlassDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);
                    }

                    using (DataView DV = new DataView(FatOrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = FatSimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = FatGridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                    }

                    string C = "Маркетинг ";
                    DataRow[] CRows = MarketOrdersNames.Select("MainOrderID=" + Convert.ToInt32(DistMainOrdersDT.Rows[i]["MainOrderID"]));
                    if (CRows.Count() > 0)
                    {
                        MegaBatchID = Convert.ToInt32(CRows[0]["MegaBatchID"]);
                        BatchID = Convert.ToInt32(CRows[0]["BatchID"]);
                        ClientName = CRows[0]["ClientName"].ToString();
                        Notes = CRows[0]["Notes"].ToString();
                        OrderName = "№" + CRows[0]["OrderNumber"].ToString() + "-" + MainOrderID;
                        if (Convert.ToInt32(CRows[0]["ClientID"]) == 101)
                            C = "Москва-1 ";
                    }
                    if (DeyingDT.Rows.Count > 0)
                        DeyingByMainOrderToExcelSingly(ref hssfworkbook, ref sheet1,
                            Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DeyingDT, WorkAssignmentID,
                            C + MegaBatchID + ", " + BatchID + ", " + MainOrderID, ClientName, OrderName, Notes, ref RowIndex);
                }
            }
        }

        private DataTable DistFrameColorsTable(DataTable SourceDT, bool OrderASC)
        {
            int ColorID = 0;
            DataTable DT = new DataTable();
            DT.Columns.Add(new DataColumn("ColorID", Type.GetType("System.Int32")));
            foreach (DataRow Row in SourceDT.Rows)
            {
                if (int.TryParse(Row["ColorID"].ToString(), out ColorID))
                {
                    DataRow NewRow = DT.NewRow();
                    NewRow["ColorID"] = ColorID;
                    DT.Rows.Add(NewRow);
                }
            }
            using (DataView DV = new DataView(DT.Copy()))
            {
                DT.Clear();
                if (OrderASC)
                    DV.Sort = "ColorID ASC";
                else
                    DV.Sort = "ColorID DESC";
                DT = DV.ToTable(true, new string[] { "ColorID" });
            }
            return DT;
        }

        private DataTable DistHeightAndWidth(DataTable SourceDT, bool OrderASC)
        {
            int Height = 0;
            DataTable DT = new DataTable();
            DT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            foreach (DataRow Row in SourceDT.Rows)
            {
                if (int.TryParse(Row["Height"].ToString(), out Height))
                {
                    DataRow NewRow = DT.NewRow();
                    NewRow["Height"] = Height;
                    DT.Rows.Add(NewRow);
                }
                if (int.TryParse(Row["Width"].ToString(), out Height))
                {
                    DataRow NewRow = DT.NewRow();
                    NewRow["Height"] = Height;
                    DT.Rows.Add(NewRow);
                }
            }
            using (DataView DV = new DataView(DT.Copy()))
            {
                DT.Clear();
                if (OrderASC)
                    DV.Sort = "Height ASC";
                else
                    DV.Sort = "Height DESC";
                DT = DV.ToTable(true, new string[] { "Height" });
            }
            return DT;
        }

        private DataTable DistInsetColorsTable(DataTable SourceDT, bool OrderASC)
        {
            int InsetColorID = 0;
            DataTable DT = new DataTable();
            DT.Columns.Add(new DataColumn("InsetColorID", Type.GetType("System.Int32")));
            foreach (DataRow Row in SourceDT.Rows)
            {
                //if (Convert.ToInt32(Row["InsetTypeID"]) != 2 && Convert.ToInt32(Row["InsetTypeID"]) != 5 && Convert.ToInt32(Row["InsetTypeID"]) != 6
                //    && Convert.ToInt32(Row["InsetTypeID"]) != 9 && Convert.ToInt32(Row["InsetTypeID"]) != 10 && Convert.ToInt32(Row["InsetTypeID"]) != 11)
                //    continue;

                if (int.TryParse(Row["InsetColorID"].ToString(), out InsetColorID))
                {
                    DataRow NewRow = DT.NewRow();
                    NewRow["InsetColorID"] = InsetColorID;
                    DT.Rows.Add(NewRow);
                }
            }
            using (DataView DV = new DataView(DT.Copy()))
            {
                DT.Clear();
                if (OrderASC)
                    DV.Sort = "InsetColorID ASC";
                else
                    DV.Sort = "InsetColorID DESC";
                DT = DV.ToTable(true, new string[] { "InsetColorID" });
            }
            return DT;
        }

        private DataTable DistMainOrdersTable(DataTable SourceDT, bool OrderASC)
        {
            int MainOrderID = 0;
            DataTable DT = new DataTable();
            DT.Columns.Add(new DataColumn("MainOrderID", Type.GetType("System.Int32")));
            DT.Columns.Add(new DataColumn("GroupType", Type.GetType("System.Int32")));
            foreach (DataRow Row in SourceDT.Rows)
            {
                if (int.TryParse(Row["MainOrderID"].ToString(), out MainOrderID))
                {
                    DataRow NewRow = DT.NewRow();
                    NewRow["MainOrderID"] = MainOrderID;
                    NewRow["GroupType"] = Convert.ToInt32(Row["GroupType"]);
                    DT.Rows.Add(NewRow);
                }
            }
            using (DataView DV = new DataView(DT.Copy()))
            {
                DT.Clear();
                if (OrderASC)
                    DV.Sort = "MainOrderID ASC";
                else
                    DV.Sort = "MainOrderID DESC";
                DT = DV.ToTable(true, new string[] { "MainOrderID", "GroupType" });
            }
            return DT;
        }

        private void DistMainOrdersTable(DataTable SourceDT1)
        {
            int MainOrderID = 0;
            foreach (DataRow Row in SourceDT1.Rows)
            {
                if (int.TryParse(Row["MainOrderID"].ToString(), out MainOrderID))
                {
                    DataRow NewRow = DistMainOrdersDT.NewRow();
                    NewRow["MainOrderID"] = MainOrderID;
                    NewRow["GroupType"] = Convert.ToInt32(Row["GroupType"]);
                    DistMainOrdersDT.Rows.Add(NewRow);
                }
            }
        }

        private DataTable DistSizesTable(DataTable SourceDT, bool OrderASC)
        {
            DataTable DT = new DataTable();
            DT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            DT.Columns.Add(new DataColumn("Width", Type.GetType("System.Int32")));
            foreach (DataRow Row in SourceDT.Rows)
            {
                DataRow NewRow = DT.NewRow();
                NewRow["Height"] = Row["Height"];
                NewRow["Width"] = Row["Width"];
                DT.Rows.Add(NewRow);
            }
            using (DataView DV = new DataView(DT.Copy()))
            {
                DT.Clear();
                if (OrderASC)
                    DV.Sort = "Height ASC, Width ASC";
                else
                    DV.Sort = "Height DESC, Width DESC";
                DT = DV.ToTable(true, new string[] { "Height", "Width" });
            }
            return DT;
        }

        private DataTable DistWidth(DataTable SourceDT, bool OrderASC)
        {
            int Height = 0;
            DataTable DT = new DataTable();
            DT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            foreach (DataRow Row in SourceDT.Rows)
            {
                if (int.TryParse(Row["Height"].ToString(), out Height))
                {
                    DataRow NewRow = DT.NewRow();
                    NewRow["Height"] = Height;
                    DT.Rows.Add(NewRow);
                }
                if (int.TryParse(Row["Width"].ToString(), out Height))
                {
                    DataRow NewRow = DT.NewRow();
                    NewRow["Height"] = Height;
                    DT.Rows.Add(NewRow);
                }
            }
            using (DataView DV = new DataView(DT.Copy()))
            {
                DT.Clear();
                if (OrderASC)
                    DV.Sort = "Height ASC";
                else
                    DV.Sort = "Height DESC";
                DT = DV.ToTable(true, new string[] { "Height" });
            }
            return DT;
        }

        private void Fill()
        {
            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT TOP 0 infiniu2_catalog.dbo.TechStore.TechStoreName, infiniu2_catalog.dbo.FrontsConfig.FrontConfigID, infiniu2_catalog.dbo.FrontsConfig.ProfileID FROM infiniu2_catalog.dbo.TechStore
                    INNER JOIN infiniu2_catalog.dbo.FrontsConfig ON infiniu2_catalog.dbo.TechStore.TechStoreID = infiniu2_catalog.dbo.FrontsConfig.ProfileID",
                ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(ProfileNamesDT);
            }

            DecorDT = new DataTable();
            string SelectCommand = @"SELECT DISTINCT TechStore.TechStoreID AS DecorID, TechStore.TechStoreName AS Name, DecorConfig.ProductID FROM TechStore
                INNER JOIN DecorConfig ON TechStore.TechStoreID = DecorConfig.DecorID AND Enabled = 1 ORDER BY TechStoreName";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(DecorDT);
            }
            DecorParametersDT = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT * FROM DecorParameters",
                ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(DecorParametersDT);
            }

            SelectCommand = @"SELECT DISTINCT TechStoreID AS FrontID, TechStoreName AS FrontName FROM TechStore
                WHERE TechStoreID IN (SELECT FrontID FROM FrontsConfig WHERE Enabled = 1) ORDER BY TechStoreName";
            FrontsDataTable = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(FrontsDataTable);
            }
            PatinaDataTable = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM Patina",
                ConnectionStrings.CatalogConnectionString))
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
            GetColorsDT();
            GetInsetColorsDT();
            InsetTypesDataTable = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM InsetTypes",
                ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(InsetTypesDataTable);
            }
            TechnoInsetTypesDataTable = InsetTypesDataTable.Copy();
            TechnoInsetColorsDataTable = InsetColorsDataTable.Copy();

            StandardImpostDataTable = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM StandardImpost45",
                ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(StandardImpostDataTable);
            }

            SelectCommand = @"SELECT TOP 0 FrontsOrders.FrontsOrdersID, FrontsOrders.MainOrderID, FrontsOrders.FrontID, FrontsOrders.TechnoProfileID, FrontsOrders.PatinaID, FrontsOrders.InsetTypeID,
                FrontsOrders.ColorID, FrontsOrders.TechnoColorID, FrontsOrders.InsetColorID, FrontsOrders.TechnoInsetTypeID, FrontsOrders.TechnoInsetColorID, FrontsOrders.Height, FrontsOrders.Width, FrontsOrders.Count, FrontsOrders.FrontConfigID, infiniu2_catalog.dbo.FrontsConfig.ProfileID, FrontsOrders.Notes FROM FrontsOrders
                INNER JOIN infiniu2_catalog.dbo.FrontsConfig ON FrontsOrders.FrontConfigID=infiniu2_catalog.dbo.FrontsConfig.FrontConfigID";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.ZOVOrdersConnectionString))
            {
                DA.Fill(LeonOrdersDT);
                LeonOrdersDT.Columns.Add(new DataColumn("GroupType", Type.GetType("System.Int32")));

                TechnoNVitrinaDT = LeonOrdersDT.Clone();
                TechnoNGridsDT = LeonOrdersDT.Clone();
                TechnoNSimpleDT = LeonOrdersDT.Clone();
                TechnoNOrdersDT = LeonOrdersDT.Clone();

                AntaliaVitrinaDT = LeonOrdersDT.Clone();
                AntaliaGridsDT = LeonOrdersDT.Clone();
                AntaliaSimpleDT = LeonOrdersDT.Clone();
                AntaliaOrdersDT = LeonOrdersDT.Clone();

                Nord95VitrinaDT = LeonOrdersDT.Clone();
                Nord95GridsDT = LeonOrdersDT.Clone();
                Nord95SimpleDT = LeonOrdersDT.Clone();
                Nord95OrdersDT = LeonOrdersDT.Clone();

                epFoxVitrinaDT = LeonOrdersDT.Clone();
                epFoxGridsDT = LeonOrdersDT.Clone();
                epFoxSimpleDT = LeonOrdersDT.Clone();
                epFoxOrdersDT = LeonOrdersDT.Clone();

                VeneciaVitrinaDT = LeonOrdersDT.Clone();
                VeneciaGridsDT = LeonOrdersDT.Clone();
                VeneciaSimpleDT = LeonOrdersDT.Clone();
                VeneciaOrdersDT = LeonOrdersDT.Clone();

                BergamoVitrinaDT = LeonOrdersDT.Clone();
                BergamoGridsDT = LeonOrdersDT.Clone();
                BergamoSimpleDT = LeonOrdersDT.Clone();
                BergamoOrdersDT = LeonOrdersDT.Clone();
                BergamoGlassDT = LeonOrdersDT.Clone();

                ep071VitrinaDT = LeonOrdersDT.Clone();
                ep071GridsDT = LeonOrdersDT.Clone();
                ep071SimpleDT = LeonOrdersDT.Clone();
                ep071OrdersDT = LeonOrdersDT.Clone();

                ep206VitrinaDT = LeonOrdersDT.Clone();
                ep206GridsDT = LeonOrdersDT.Clone();
                ep206SimpleDT = LeonOrdersDT.Clone();
                ep206OrdersDT = LeonOrdersDT.Clone();

                ep216VitrinaDT = LeonOrdersDT.Clone();
                ep216GridsDT = LeonOrdersDT.Clone();
                ep216SimpleDT = LeonOrdersDT.Clone();
                ep216OrdersDT = LeonOrdersDT.Clone();

                ep111VitrinaDT = LeonOrdersDT.Clone();
                ep111GridsDT = LeonOrdersDT.Clone();
                ep111SimpleDT = LeonOrdersDT.Clone();
                ep111OrdersDT = LeonOrdersDT.Clone();

                ep041VitrinaDT = LeonOrdersDT.Clone();
                ep041GridsDT = LeonOrdersDT.Clone();
                ep041SimpleDT = LeonOrdersDT.Clone();
                ep041OrdersDT = LeonOrdersDT.Clone();

                BostonVitrinaDT = LeonOrdersDT.Clone();
                BostonGridsDT = LeonOrdersDT.Clone();
                BostonSimpleDT = LeonOrdersDT.Clone();
                BostonOrdersDT = LeonOrdersDT.Clone();

                LeonVitrinaDT = LeonOrdersDT.Clone();
                LeonGridsDT = LeonOrdersDT.Clone();
                LeonSimpleDT = LeonOrdersDT.Clone();

                LimogVitrinaDT = LeonOrdersDT.Clone();
                LimogGridsDT = LeonOrdersDT.Clone();
                LimogSimpleDT = LeonOrdersDT.Clone();
                LimogOrdersDT = LeonOrdersDT.Clone();

                ep066Marsel4VitrinaDT = LeonOrdersDT.Clone();
                ep066Marsel4GridsDT = LeonOrdersDT.Clone();
                ep066Marsel4SimpleDT = LeonOrdersDT.Clone();
                ep066Marsel4OrdersDT = LeonOrdersDT.Clone();

                ep110JersyVitrinaDT = LeonOrdersDT.Clone();
                ep110JersyGridsDT = LeonOrdersDT.Clone();
                ep110JersySimpleDT = LeonOrdersDT.Clone();
                ep110JersyOrdersDT = LeonOrdersDT.Clone();

                ep018Marsel1VitrinaDT = LeonOrdersDT.Clone();
                ep018Marsel1GridsDT = LeonOrdersDT.Clone();
                ep018Marsel1SimpleDT = LeonOrdersDT.Clone();
                ep018Marsel1OrdersDT = LeonOrdersDT.Clone();

                ep043ShervudVitrinaDT = LeonOrdersDT.Clone();
                ep043ShervudGridsDT = LeonOrdersDT.Clone();
                ep043ShervudSimpleDT = LeonOrdersDT.Clone();
                ep043ShervudOrdersDT = LeonOrdersDT.Clone();

                ep112VitrinaDT = LeonOrdersDT.Clone();
                ep112GridsDT = LeonOrdersDT.Clone();
                ep112SimpleDT = LeonOrdersDT.Clone();
                ep112OrdersDT = LeonOrdersDT.Clone();

                UrbanVitrinaDT = LeonOrdersDT.Clone();
                UrbanGridsDT = LeonOrdersDT.Clone();
                UrbanSimpleDT = LeonOrdersDT.Clone();
                UrbanOrdersDT = LeonOrdersDT.Clone();

                AlbyVitrinaDT = LeonOrdersDT.Clone();
                AlbyGridsDT = LeonOrdersDT.Clone();
                AlbySimpleDT = LeonOrdersDT.Clone();
                AlbyOrdersDT = LeonOrdersDT.Clone();

                BrunoVitrinaDT = LeonOrdersDT.Clone();
                BrunoGridsDT = LeonOrdersDT.Clone();
                BrunoSimpleDT = LeonOrdersDT.Clone();
                BrunoOrdersDT = LeonOrdersDT.Clone();

                epsh406Techno4VitrinaDT = LeonOrdersDT.Clone();
                epsh406Techno4GridsDT = LeonOrdersDT.Clone();
                epsh406Techno4SimpleDT = LeonOrdersDT.Clone();
                epsh406Techno4OrdersDT = LeonOrdersDT.Clone();

                LukVitrinaDT = LeonOrdersDT.Clone();
                LukGridsDT = LeonOrdersDT.Clone();
                LukSimpleDT = LeonOrdersDT.Clone();
                LukOrdersDT = LeonOrdersDT.Clone();

                LukPVHVitrinaDT = LeonOrdersDT.Clone();
                LukPVHGridsDT = LeonOrdersDT.Clone();
                LukPVHSimpleDT = LeonOrdersDT.Clone();
                LukPVHOrdersDT = LeonOrdersDT.Clone();

                MilanoVitrinaDT = LeonOrdersDT.Clone();
                MilanoGridsDT = LeonOrdersDT.Clone();
                MilanoSimpleDT = LeonOrdersDT.Clone();
                MilanoOrdersDT = LeonOrdersDT.Clone();

                PragaVitrinaDT = LeonOrdersDT.Clone();
                PragaGridsDT = LeonOrdersDT.Clone();
                PragaSimpleDT = LeonOrdersDT.Clone();
                PragaOrdersDT = LeonOrdersDT.Clone();

                SigmaVitrinaDT = LeonOrdersDT.Clone();
                SigmaGlassDT = LeonOrdersDT.Clone();
                BostonGlassDT = LeonOrdersDT.Clone();
                SigmaGridsDT = LeonOrdersDT.Clone();
                SigmaSimpleDT = LeonOrdersDT.Clone();
                SigmaOrdersDT = LeonOrdersDT.Clone();

                FatVitrinaDT = LeonOrdersDT.Clone();
                FatGridsDT = LeonOrdersDT.Clone();
                FatSimpleDT = LeonOrdersDT.Clone();
                FatOrdersDT = LeonOrdersDT.Clone();
            }

            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT TOP 0 DecorOrderID, MainOrderID, ProductID, DecorID, DecorConfigID, ColorID, PatinaID,
                Height, Length, Width, LeftAngle, RightAngle, Count, Notes FROM DecorOrders",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(BagetWithAngelOrdersDT);
                BagetWithAngelOrdersDT.Columns.Add(new DataColumn("GroupType", Type.GetType("System.Int32")));
            }
            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT TOP 0 DecorOrderID, MainOrderID, ProductID, DecorID, DecorConfigID, ColorID, PatinaID,
                Height, Length, Width, Count, Notes FROM DecorOrders",
                ConnectionStrings.ZOVOrdersConnectionString))
            {
                DA.Fill(NotArchDecorOrdersDT);
                NotArchDecorOrdersDT.Columns.Add(new DataColumn("GroupType", Type.GetType("System.Int32")));
                ArchDecorOrdersDT = NotArchDecorOrdersDT.Clone();
                GridsDecorOrdersDT = NotArchDecorOrdersDT.Clone();
            }
        }

        private void GetArchDecorOrders(ref DataTable DestinationDT, int WorkAssignmentID, int FactoryID)
        {
            string SelectCommand = string.Empty;
            DataTable DT = new DataTable();

            SelectCommand = @"SELECT DecorOrderID, MainOrderID, ProductID, DecorID, DecorConfigID, ColorID, PatinaID,
                Height, Length, Width, Count, Notes FROM DecorOrders
                WHERE FactoryID=" + FactoryID + " AND ProductID IN (31, 4, 18, 32) AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE ProfilWorkAssignmentID=" + WorkAssignmentID + "))";
            if (FactoryID == 2)
                SelectCommand = @"SELECT DecorOrderID, MainOrderID, ProductID, DecorID, DecorConfigID, ColorID, PatinaID,
                    Height, Length, Width, Count, Notes FROM DecorOrders
                    WHERE FactoryID=" + FactoryID + " AND ProductID IN (31, 4, 18, 32) AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE TPSWorkAssignmentID=" + WorkAssignmentID + "))";

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.ZOVOrdersConnectionString))
            {
                DA.Fill(DT);
            }
            foreach (DataRow item in DT.Rows)
            {
                DataRow NewRow = DestinationDT.NewRow();
                NewRow.ItemArray = item.ItemArray;
                NewRow["GroupType"] = 0;
                DestinationDT.Rows.Add(NewRow);
            }

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                DT.Clear();
                DA.Fill(DT);
            }
            foreach (DataRow item in DT.Rows)
            {
                DataRow NewRow = DestinationDT.NewRow();
                NewRow.ItemArray = item.ItemArray;
                NewRow["GroupType"] = 1;
                DestinationDT.Rows.Add(NewRow);
            }
        }

        private void GetBagetWithAngleOrders(ref DataTable DestinationDT, int WorkAssignmentID, int FactoryID)
        {
            string SelectCommand = string.Empty;
            DataTable DT = new DataTable();

            SelectCommand = @"SELECT DecorOrderID, MainOrderID, ProductID, DecorID, DecorConfigID, ColorID, PatinaID,
                Height, Length, Width, LeftAngle, RightAngle, Count, Notes FROM DecorOrders
                WHERE FactoryID=" + FactoryID + " AND ProductID IN (1) AND (LeftAngle<>0 OR RightAngle<>0) AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE ProfilWorkAssignmentID=" + WorkAssignmentID + "))";
            if (FactoryID == 2)
                SelectCommand = @"SELECT DecorOrderID, MainOrderID, ProductID, DecorID, DecorConfigID, ColorID, PatinaID,
                    Height, Length, Width, LeftAngle, RightAngle, Count, Notes FROM DecorOrders
                    WHERE FactoryID=" + FactoryID + " AND ProductID IN (1) AND (LeftAngle<>0 OR RightAngle<>0) AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE TPSWorkAssignmentID=" + WorkAssignmentID + "))";

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.ZOVOrdersConnectionString))
            {
                DA.Fill(DT);
            }
            foreach (DataRow item in DT.Rows)
            {
                DataRow NewRow = DestinationDT.NewRow();
                NewRow.ItemArray = item.ItemArray;
                NewRow["GroupType"] = 0;
                DestinationDT.Rows.Add(NewRow);
            }

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                DT.Clear();
                DA.Fill(DT);
            }
            foreach (DataRow item in DT.Rows)
            {
                DataRow NewRow = DestinationDT.NewRow();
                NewRow.ItemArray = item.ItemArray;
                NewRow["GroupType"] = 1;
                DestinationDT.Rows.Add(NewRow);
            }
        }

        private void GetColorsDT()
        {
            FrameColorsDataTable = new DataTable();
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
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM InsetColors WHERE GroupID IN (2,3,4,5,6,21,22)", ConnectionStrings.CatalogConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    DA.Fill(DT);

                    for (int i = 0; i < DT.Rows.Count; i++)
                    {
                        DataRow[] rows = FrameColorsDataTable.Select("ColorID=" + Convert.ToInt32(DT.Rows[i]["InsetColorID"]));
                        if (rows.Count() == 0)
                        {
                            DataRow NewRow = FrameColorsDataTable.NewRow();
                            NewRow["ColorID"] = Convert.ToInt64(DT.Rows[i]["InsetColorID"]);
                            NewRow["ColorName"] = DT.Rows[i]["InsetColorName"].ToString();
                            FrameColorsDataTable.Rows.Add(NewRow);
                        }
                    }
                }
            }
        }

        private string GetDecorName(int ID)
        {
            DataRow[] rows = DecorDT.Select("DecorID=" + ID);
            if (rows.Count() > 0)
                return rows[0]["Name"].ToString();
            else
                return string.Empty;
        }

        private string GetFileName(string sDestFolder, string ExcelName)
        {
            string sExtension = ".xls";
            string sFileName = ExcelName;

            int j = 1;
            while (FM.FileExist(sDestFolder + "/" + sFileName + sExtension, Configs.FTPType))
            {
                sFileName = ExcelName + "(" + j++ + ")";
            }
            sFileName = sFileName + sExtension;
            return sFileName;
        }

        private void GetFrontsOrders(ref DataTable DestinationDT, int WorkAssignmentID, Fronts Front)
        {
            string SelectCommand = string.Empty;
            DataTable DT = new DataTable();

            SelectCommand = @"SELECT FrontsOrders.FrontsOrdersID, FrontsOrders.MainOrderID, FrontsOrders.FrontID, FrontsOrders.TechnoProfileID, FrontsOrders.PatinaID, FrontsOrders.InsetTypeID,
                FrontsOrders.ColorID, FrontsOrders.TechnoColorID, FrontsOrders.InsetColorID, FrontsOrders.TechnoInsetTypeID, FrontsOrders.TechnoInsetColorID, FrontsOrders.Height, FrontsOrders.Width, FrontsOrders.Count, FrontsOrders.FrontConfigID, infiniu2_catalog.dbo.FrontsConfig.ProfileID, FrontsOrders.Notes FROM FrontsOrders
                INNER JOIN infiniu2_catalog.dbo.FrontsConfig ON FrontsOrders.FrontConfigID=infiniu2_catalog.dbo.FrontsConfig.FrontConfigID
                WHERE FrontsOrders.FrontID=" + Convert.ToInt32(Front) +
                " AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE ProfilWorkAssignmentID=" + WorkAssignmentID + "))";

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.ZOVOrdersConnectionString))
            {
                DA.Fill(DT);
            }
            foreach (DataRow item in DT.Rows)
            {
                if (item["Notes"] != DBNull.Value && item["Notes"].ToString().Length == 0)
                    item["Notes"] = DBNull.Value;
                DataRow NewRow = DestinationDT.NewRow();
                NewRow.ItemArray = item.ItemArray;
                NewRow["GroupType"] = 0;
                DestinationDT.Rows.Add(NewRow);
            }

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                DT.Clear();
                DA.Fill(DT);
            }
            foreach (DataRow item in DT.Rows)
            {
                if (item["Notes"] != DBNull.Value && item["Notes"].ToString().Length == 0)
                    item["Notes"] = DBNull.Value;
                DataRow NewRow = DestinationDT.NewRow();
                NewRow.ItemArray = item.ItemArray;
                NewRow["GroupType"] = 1;
                DestinationDT.Rows.Add(NewRow);
            }
        }

        private void GetGlassFronts(DataTable SourceDT, ref DataTable DestinationDT)
        {
            DataRow[] rows = SourceDT.Select("InsetTypeID IN (2)");
            foreach (DataRow dr in rows)
                DestinationDT.Rows.Add(dr.ItemArray);
        }

        private void GetGridFronts(DataTable SourceDT, ref DataTable DestinationDT)
        {
            DataRow[] rows = SourceDT.Select("InsetTypeID IN (685,686,687,688,29470,29471)");
            foreach (DataRow dr in rows)
                DestinationDT.Rows.Add(dr.ItemArray);
        }

        private void GetGridsDecorOrders(ref DataTable DestinationDT, int WorkAssignmentID, int FactoryID)
        {
            string SelectCommand = string.Empty;
            DataTable DT = new DataTable();

            SelectCommand = @"SELECT DecorOrderID, MainOrderID, ProductID, DecorID, DecorConfigID, ColorID, PatinaID,
                Height, Length, Width, Count, Notes FROM DecorOrders
                WHERE FactoryID=" + FactoryID + " AND ProductID IN (10, 11, 12) AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE ProfilWorkAssignmentID=" + WorkAssignmentID + "))";
            if (FactoryID == 2)
                SelectCommand = @"SELECT DecorOrderID, MainOrderID, ProductID, DecorID, DecorConfigID, ColorID, PatinaID,
                    Height, Length, Width, Count, Notes FROM DecorOrders
                    WHERE FactoryID=" + FactoryID + " AND ProductID IN (10, 11, 12) AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE TPSWorkAssignmentID=" + WorkAssignmentID + "))";

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.ZOVOrdersConnectionString))
            {
                DA.Fill(DT);
            }
            foreach (DataRow item in DT.Rows)
            {
                DataRow NewRow = DestinationDT.NewRow();
                NewRow.ItemArray = item.ItemArray;
                NewRow["GroupType"] = 0;
                DestinationDT.Rows.Add(NewRow);
            }

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                DT.Clear();
                DA.Fill(DT);
            }
            foreach (DataRow item in DT.Rows)
            {
                DataRow NewRow = DestinationDT.NewRow();
                NewRow.ItemArray = item.ItemArray;
                NewRow["GroupType"] = 1;
                DestinationDT.Rows.Add(NewRow);
            }
        }

        private void GetInsetColorsDT()
        {
            InsetColorsDataTable = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT InsetColors.InsetColorID, InsetColors.GroupID, infiniu2_catalog.dbo.TechStore.TechStoreName AS InsetColorName FROM InsetColors" +
                " INNER JOIN infiniu2_catalog.dbo.TechStore ON InsetColors.InsetColorID = infiniu2_catalog.dbo.TechStore.TechStoreID ORDER BY TechStoreName", ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(InsetColorsDataTable);
                {
                    DataRow NewRow = InsetColorsDataTable.NewRow();
                    NewRow["InsetColorID"] = -1;
                    NewRow["GroupID"] = -1;
                    NewRow["InsetColorName"] = "-";
                    InsetColorsDataTable.Rows.Add(NewRow);
                }
                {
                    DataRow NewRow = InsetColorsDataTable.NewRow();
                    NewRow["InsetColorID"] = 0;
                    NewRow["GroupID"] = -1;
                    NewRow["InsetColorName"] = "на выбор";
                    InsetColorsDataTable.Rows.Add(NewRow);
                }
            }
        }

        private void GetLuxMegaPlankaFronts(DataTable SourceDT, ref DataTable DestinationDT)
        {
            DataRow[] rows = SourceDT.Select("InsetTypeID IN (860,862,4310)");
            foreach (DataRow dr in rows)
                DestinationDT.Rows.Add(dr.ItemArray);
        }

        private void GetNotArchDecorOrders(ref DataTable DestinationDT, int WorkAssignmentID, int FactoryID)
        {
            string SelectCommand = string.Empty;
            DataTable DT = new DataTable();

            SelectCommand = @"SELECT DecorOrderID, MainOrderID, ProductID, DecorID, DecorConfigID, ColorID, PatinaID,
                Height, Length, Width, Count, Notes FROM DecorOrders
                WHERE FactoryID=" + FactoryID + " AND NOT (ProductID IN (1) AND (LeftAngle<>0 OR RightAngle<>0)) AND ProductID NOT IN (31, 4, 18, 32, 10, 11, 12) AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE ProfilWorkAssignmentID=" + WorkAssignmentID + "))";
            if (FactoryID == 2)
                SelectCommand = @"SELECT DecorOrderID, MainOrderID, ProductID, DecorID, DecorConfigID, ColorID, PatinaID,
                    Height, Length, Width, Count, Notes FROM DecorOrders
                    WHERE FactoryID=" + FactoryID + " AND NOT (ProductID IN (1) AND (LeftAngle<>0 OR RightAngle<>0)) AND ProductID NOT IN (31, 4, 18, 32, 10, 11, 12) AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE TPSWorkAssignmentID=" + WorkAssignmentID + "))";

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.ZOVOrdersConnectionString))
            {
                DA.Fill(DT);
            }
            foreach (DataRow item in DT.Rows)
            {
                DataRow NewRow = DestinationDT.NewRow();
                NewRow.ItemArray = item.ItemArray;
                NewRow["GroupType"] = 0;
                DestinationDT.Rows.Add(NewRow);
            }

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                DT.Clear();
                DA.Fill(DT);
            }
            foreach (DataRow item in DT.Rows)
            {
                DataRow NewRow = DestinationDT.NewRow();
                NewRow.ItemArray = item.ItemArray;
                NewRow["GroupType"] = 1;
                DestinationDT.Rows.Add(NewRow);
            }
        }

        private string GetOrderName(int MainOrderID, int GroupType)
        {
            string name = string.Empty;
            string ConnectionString = ConnectionStrings.ZOVOrdersConnectionString;
            string SelectCommand = string.Empty;
            DataTable DT = new DataTable();

            if (GroupType == 1)
                ConnectionString = ConnectionStrings.MarketingOrdersConnectionString;
            SelectCommand = @"SELECT MegaBatchID, BatchID FROM Batch WHERE BatchID IN (SELECT BatchID FROM BatchDetails WHERE MainOrderID = " + MainOrderID + ")";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionString))
            {
                if (DA.Fill(DT) > 0 && DT.Rows[0]["MegaBatchID"] != DBNull.Value && DT.Rows[0]["BatchID"] != DBNull.Value)
                    name = DT.Rows[0]["MegaBatchID"].ToString() + ", " + DT.Rows[0]["BatchID"] + ", " + MainOrderID;
            }
            return name;
        }

        private void GetProfileNames(ref DataTable DestinationDT, int WorkAssignmentID)
        {
            string SelectCommand = string.Empty;
            DataTable DT = new DataTable();

            SelectCommand = @"SELECT infiniu2_catalog.dbo.TechStore.TechStoreName, infiniu2_catalog.dbo.FrontsConfig.FrontConfigID, infiniu2_catalog.dbo.FrontsConfig.ProfileID FROM infiniu2_catalog.dbo.TechStore
                INNER JOIN infiniu2_catalog.dbo.FrontsConfig ON infiniu2_catalog.dbo.TechStore.TechStoreID = infiniu2_catalog.dbo.FrontsConfig.ProfileID AND infiniu2_catalog.dbo.FrontsConfig.FrontConfigID IN (SELECT FrontConfigID FROM FrontsOrders
                WHERE FrontsOrders.FactoryID=1 AND FrontID IN (" +
                Convert.ToInt32(Fronts.TechnoN) + "," +
                Convert.ToInt32(Fronts.Antalia) + "," + Convert.ToInt32(Fronts.epFox) + "," + Convert.ToInt32(Fronts.Nord95) + "," + Convert.ToInt32(Fronts.Fat) + "," + Convert.ToInt32(Fronts.Leon) + "," +
                Convert.ToInt32(Fronts.Limog) + "," +
                Convert.ToInt32(Fronts.ep066Marsel4) + "," + Convert.ToInt32(Fronts.ep110Jersy) + "," +
                Convert.ToInt32(Fronts.ep018Marsel1) + "," + Convert.ToInt32(Fronts.ep043Shervud) + "," +
                Convert.ToInt32(Fronts.ep018Marsel1) + "," + Convert.ToInt32(Fronts.ep112) + "," +
                Convert.ToInt32(Fronts.Urban) + "," + Convert.ToInt32(Fronts.Alby) + "," + Convert.ToInt32(Fronts.Bruno) + "," +
                Convert.ToInt32(Fronts.epsh406Techno4) + "," +
                Convert.ToInt32(Fronts.Luk) + "," + Convert.ToInt32(Fronts.LukPVH) + "," +
                Convert.ToInt32(Fronts.Milano) + "," +
                Convert.ToInt32(Fronts.Praga) + "," + Convert.ToInt32(Fronts.Sigma) + "," + Convert.ToInt32(Fronts.Venecia) + "," + Convert.ToInt32(Fronts.Bergamo) + "," +
                Convert.ToInt32(Fronts.ep041) + "," + Convert.ToInt32(Fronts.ep071) + "," + Convert.ToInt32(Fronts.ep206) + "," + Convert.ToInt32(Fronts.ep216) + "," + Convert.ToInt32(Fronts.ep111) + "," + Convert.ToInt32(Fronts.Boston) + "))";

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.ZOVOrdersConnectionString))
            {
                DA.Fill(DT);
            }
            foreach (DataRow item in DT.Rows)
            {
                DataRow NewRow = DestinationDT.NewRow();
                NewRow.ItemArray = item.ItemArray;
                DestinationDT.Rows.Add(NewRow);
            }

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                DT.Clear();
                DA.Fill(DT);
            }
            foreach (DataRow item in DT.Rows)
            {
                DataRow NewRow = DestinationDT.NewRow();
                NewRow.ItemArray = item.ItemArray;
                DestinationDT.Rows.Add(NewRow);
            }

            SelectCommand = @"SELECT infiniu2_catalog.dbo.TechStore.TechStoreName, infiniu2_catalog.dbo.FrontsConfig.FrontConfigID, infiniu2_catalog.dbo.FrontsConfig.TechnoProfileID AS ProfileID FROM infiniu2_catalog.dbo.TechStore
                INNER JOIN infiniu2_catalog.dbo.FrontsConfig ON infiniu2_catalog.dbo.TechStore.TechStoreID = infiniu2_catalog.dbo.FrontsConfig.TechnoProfileID AND infiniu2_catalog.dbo.FrontsConfig.FrontConfigID IN (SELECT FrontConfigID FROM FrontsOrders
                WHERE FrontsOrders.FactoryID=1 AND FrontID IN (" +
                Convert.ToInt32(Fronts.TechnoN) + "," +
                Convert.ToInt32(Fronts.Antalia) + "," + Convert.ToInt32(Fronts.epFox) + "," + Convert.ToInt32(Fronts.Nord95) + "," + Convert.ToInt32(Fronts.Fat) + "," + Convert.ToInt32(Fronts.Leon) + "," +
                Convert.ToInt32(Fronts.Limog) + "," +
                Convert.ToInt32(Fronts.ep066Marsel4) + "," + Convert.ToInt32(Fronts.ep110Jersy) + "," +
                Convert.ToInt32(Fronts.ep018Marsel1) + "," + Convert.ToInt32(Fronts.ep043Shervud) + "," +
                Convert.ToInt32(Fronts.ep018Marsel1) + "," + Convert.ToInt32(Fronts.ep112) + "," +
                Convert.ToInt32(Fronts.Urban) + "," + Convert.ToInt32(Fronts.Alby) + "," + Convert.ToInt32(Fronts.Bruno) + "," +
                Convert.ToInt32(Fronts.epsh406Techno4) + "," +
                Convert.ToInt32(Fronts.Luk) + "," + Convert.ToInt32(Fronts.LukPVH) + "," +
                Convert.ToInt32(Fronts.Milano) + "," +
                Convert.ToInt32(Fronts.Praga) + "," + Convert.ToInt32(Fronts.Sigma) + "," + Convert.ToInt32(Fronts.Venecia) + "," + Convert.ToInt32(Fronts.Bergamo) + "," +
                Convert.ToInt32(Fronts.ep041) + "," + Convert.ToInt32(Fronts.ep071) + "," + Convert.ToInt32(Fronts.ep206) + "," + Convert.ToInt32(Fronts.ep216) + "," + Convert.ToInt32(Fronts.ep111) + "," + Convert.ToInt32(Fronts.Boston) + "))";

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                DT.Clear();
                DA.Fill(DT);
            }
            foreach (DataRow item in DT.Rows)
            {
                DataRow NewRow = DestinationDT.NewRow();
                NewRow.ItemArray = item.ItemArray;
                DestinationDT.Rows.Add(NewRow);
            }
        }

        private void GetSimpleFronts(DataTable SourceDT, ref DataTable DestinationDT)
        {
            DataRow[] rows = SourceDT.Select("InsetTypeID NOT IN (1,2,685,686,687,688,29470,29471)");
            foreach (DataRow dr in rows)
                DestinationDT.Rows.Add(dr.ItemArray);
        }

        private void GetVitrinaFronts(DataTable SourceDT, ref DataTable DestinationDT)
        {
            DataRow[] rows = SourceDT.Select("InsetTypeID=1");
            foreach (DataRow dr in rows)
                DestinationDT.Rows.Add(dr.ItemArray);
        }

        private void GridsDecorAssemblyByMainOrderToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            int WorkAssignmentID, string BatchName)
        {
            if (GridsDecorOrdersDT.Rows.Count == 0)
                return;

            DataTable DistMainOrdersDT = DistMainOrdersTable(GridsDecorOrdersDT, true);
            DataTable DT = GridsDecorOrdersDT.Clone();
            DataTable ZOVOrdersNames = new DataTable();
            DataTable MarketOrdersNames = new DataTable();
            string MainOrdersID = string.Empty;
            string ClientName = string.Empty;
            string OrderName = string.Empty;
            string Notes = string.Empty;

            if (DistMainOrdersDT.Select("GroupType=0").Count() > 0)
            {
                MainOrdersID = string.Empty;
                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=0"))
                    MainOrdersID += Convert.ToInt32(item["MainOrderID"]) + ",";
                MainOrdersID = MainOrdersID.Substring(0, MainOrdersID.Length - 1);
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT ClientName, DocNumber, MainOrderID FROM MainOrders" +
                    " INNER JOIN infiniu2_zovreference.dbo.Clients ON MainOrders.ClientID=infiniu2_zovreference.dbo.Clients.ClientID" +
                    " WHERE MainOrderID IN (" + MainOrdersID + ")", ConnectionStrings.ZOVOrdersConnectionString))
                {
                    DA.Fill(ZOVOrdersNames);
                }

                int RowIndex = 0;
                HSSFSheet sheet1 = hssfworkbook.CreateSheet("Решетки1 ЗОВ");
                sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

                sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
                sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
                sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
                sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

                sheet1.SetColumnWidth(0, 30 * 256);
                sheet1.SetColumnWidth(1, 20 * 256);
                sheet1.SetColumnWidth(2, 9 * 256);
                sheet1.SetColumnWidth(3, 6 * 256);
                sheet1.SetColumnWidth(4, 6 * 256);

                {
                    HSSFPatriarch patriarch = sheet1.CreateDrawingPatriarch();
                    int ind = 0;

                    CreateBarcode(WorkAssignmentID);
                    HSSFClientAnchor anchor = new HSSFClientAnchor(0, 1, 1000, 8, 5, ind, 6, ind + 3)
                    {
                        AnchorType = 2
                    };
                    string BarcodeNumber = GetBarcodeNumber(24, WorkAssignmentID);
                    string barcodeName = System.Environment.GetEnvironmentVariable("TEMP") + @"/" + BarcodeNumber + ".jpg";
                    HSSFPicture picture = patriarch.CreatePicture(anchor, LoadImage(barcodeName, hssfworkbook));

                    HSSFFont BarcodeFont = hssfworkbook.CreateFont();
                    BarcodeFont.FontHeightInPoints = 12;
                    BarcodeFont.FontName = "Calibri";

                    HSSFCellStyle BarcodeStyle = hssfworkbook.CreateCellStyle();
                    BarcodeStyle.Alignment = HSSFCellStyle.ALIGN_CENTER;
                    BarcodeStyle.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
                    BarcodeStyle.SetFont(BarcodeFont);

                    HSSFCell barcodeCell = sheet1.CreateRow(ind + 3).CreateCell(5);
                    barcodeCell.SetCellValue(BarcodeNumber.ToString());
                    barcodeCell.CellStyle = BarcodeStyle;
                    sheet1.AddMergedRegion(new NPOI.HSSF.Util.Region(ind + 3, 5, ind + 3, 6));
                }

                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=0"))
                {
                    DecorAssemblyDT.Clear();
                    DT.Clear();
                    int MainOrderID = Convert.ToInt32(item["MainOrderID"]);
                    DataRow[] rows = GridsDecorOrdersDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item1 in rows)
                        DT.Rows.Add(item1.ItemArray);
                    AssemblyDecorCollect(DT, ref DecorAssemblyDT);

                    DataRow[] CRows = ZOVOrdersNames.Select("MainOrderID=" + Convert.ToInt32(item["MainOrderID"]));
                    if (CRows.Count() > 0)
                    {
                        ClientName = CRows[0]["ClientName"].ToString();
                        OrderName = CRows[0]["DocNumber"].ToString();
                    }
                    if (DecorAssemblyDT.Rows.Count > 0)
                    {
                        GridsDecorAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                             Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DecorAssemblyDT,
                             WorkAssignmentID, BatchName, ClientName, string.Empty, OrderName, Notes, ref RowIndex);
                        RowIndex++;
                        GridsDecorAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                            Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DecorAssemblyDT,
                            WorkAssignmentID, BatchName, ClientName, "ДУБЛЬ", OrderName, Notes, ref RowIndex);
                    }
                    RowIndex++;
                }
            }
            if (DistMainOrdersDT.Select("GroupType=1").Count() > 0)
            {
                MainOrdersID = string.Empty;
                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=1"))
                    MainOrdersID += Convert.ToInt32(item["MainOrderID"]) + ",";
                MainOrdersID = MainOrdersID.Substring(0, MainOrdersID.Length - 1);
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT ClientName, OrderNumber, MainOrderID, Notes FROM MainOrders" +
                    " INNER JOIN MegaOrders ON MainOrders.MegaOrderID=MegaOrders.MegaOrderID" +
                    " INNER JOIN infiniu2_marketingreference.dbo.Clients ON MegaOrders.ClientID=infiniu2_marketingreference.dbo.Clients.ClientID" +
                    " WHERE MainOrderID IN (" + MainOrdersID + ")", ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DA.Fill(MarketOrdersNames);
                }

                int RowIndex = 0;
                HSSFSheet sheet1 = hssfworkbook.CreateSheet("Решетки1 Маркетинг");
                sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

                sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
                sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
                sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
                sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

                sheet1.SetColumnWidth(0, 30 * 256);
                sheet1.SetColumnWidth(1, 20 * 256);
                sheet1.SetColumnWidth(2, 9 * 256);
                sheet1.SetColumnWidth(3, 6 * 256);
                sheet1.SetColumnWidth(4, 6 * 256);

                {
                    HSSFPatriarch patriarch = sheet1.CreateDrawingPatriarch();
                    int ind = 0;

                    CreateBarcode(WorkAssignmentID);
                    HSSFClientAnchor anchor = new HSSFClientAnchor(0, 1, 1000, 8, 5, ind, 6, ind + 3)
                    {
                        AnchorType = 2
                    };
                    string BarcodeNumber = GetBarcodeNumber(24, WorkAssignmentID);
                    string barcodeName = System.Environment.GetEnvironmentVariable("TEMP") + @"/" + BarcodeNumber + ".jpg";
                    HSSFPicture picture = patriarch.CreatePicture(anchor, LoadImage(barcodeName, hssfworkbook));

                    HSSFFont BarcodeFont = hssfworkbook.CreateFont();
                    BarcodeFont.FontHeightInPoints = 12;
                    BarcodeFont.FontName = "Calibri";

                    HSSFCellStyle BarcodeStyle = hssfworkbook.CreateCellStyle();
                    BarcodeStyle.Alignment = HSSFCellStyle.ALIGN_CENTER;
                    BarcodeStyle.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
                    BarcodeStyle.SetFont(BarcodeFont);

                    HSSFCell barcodeCell = sheet1.CreateRow(ind + 3).CreateCell(5);
                    barcodeCell.SetCellValue(BarcodeNumber.ToString());
                    barcodeCell.CellStyle = BarcodeStyle;
                    sheet1.AddMergedRegion(new NPOI.HSSF.Util.Region(ind + 3, 5, ind + 3, 6));
                }

                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=1"))
                {
                    DecorAssemblyDT.Clear();
                    DT.Clear();
                    int MainOrderID = Convert.ToInt32(item["MainOrderID"]);
                    DataRow[] rows = GridsDecorOrdersDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item1 in rows)
                        DT.Rows.Add(item1.ItemArray);
                    AssemblyDecorCollect(DT, ref DecorAssemblyDT);

                    DataRow[] CRows = MarketOrdersNames.Select("MainOrderID=" + MainOrderID);
                    if (CRows.Count() > 0)
                    {
                        ClientName = CRows[0]["ClientName"].ToString();
                        Notes = CRows[0]["Notes"].ToString();
                        OrderName = "№" + CRows[0]["OrderNumber"].ToString() + "-" + MainOrderID;
                    }
                    if (DecorAssemblyDT.Rows.Count > 0)
                    {
                        GridsDecorAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                             Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DecorAssemblyDT,
                             WorkAssignmentID, BatchName, ClientName, string.Empty, OrderName, Notes, ref RowIndex);
                        RowIndex++;
                        GridsDecorAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                            Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DecorAssemblyDT,
                            WorkAssignmentID, BatchName, ClientName, "ДУБЛЬ", OrderName, Notes, ref RowIndex);
                    }
                    RowIndex++;
                }
            }
        }

        private void InsetsFilenkaOnly(DataTable SourceDT, ref DataTable DestinationDT,
            int WidthMargin, bool OrderASC, string AdditionalName)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            DataTable DT3 = new DataTable();
            DataTable DT4 = new DataTable();

            using (DataView DV = new DataView(SourceDT, "InsetTypeID IN (2069,2070,2071,2073,2075,2077,2233,3644,29043,29531)", "InsetTypeID", DataViewRowState.CurrentRows))
            {
                DT1 = DV.ToTable(true, new string[] { "InsetTypeID" });
            }

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                string InsetTypeName = "";
                InsetTypeName = GetInsetTypeName(Convert.ToInt32(DT1.Rows[i]["InsetTypeID"])) + " ";
                using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]), "InsetColorID", DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "InsetColorID" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                        " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]), "TechnoInsetColorID", DataViewRowState.CurrentRows))
                    {
                        DT3 = DV.ToTable(true, new string[] { "TechnoInsetColorID" });
                    }
                    for (int x = 0; x < DT3.Rows.Count; x++)
                    {
                        using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                            " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) +
                            " AND TechnoInsetColorID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]), "Height, Width", DataViewRowState.CurrentRows))
                        {
                            DT4 = DV.ToTable(true, new string[] { "Height", "Width" });
                        }
                        for (int y = 0; y < DT4.Rows.Count; y++)
                        {
                            DataRow[] Srows = SourceDT.Select("InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                                " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND TechnoInsetColorID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]) +
                                " AND Height=" + Convert.ToInt32(DT4.Rows[y]["Height"]) + " AND Width=" + Convert.ToInt32(DT4.Rows[y]["Width"]));
                            if (Srows.Count() == 0)
                                continue;

                            int Count = 0;
                            int Height = Convert.ToInt32(DT4.Rows[y]["Height"]) - WidthMargin;
                            int Width = Convert.ToInt32(DT4.Rows[y]["Width"]) - WidthMargin;
                            int TechnoInsetColorID = Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]);
                            string InsetColor = "фил " + GetInsetColorName(Convert.ToInt32(DT2.Rows[j]["InsetColorID"])) + AdditionalName;

                            if (Height < 114 || Width < 114)
                                continue;

                            foreach (DataRow item in Srows)
                                Count += Convert.ToInt32(item["Count"]);

                            DataRow[] rows = DestinationDT.Select("InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) +
                                " AND TechnoInsetColorID=" + TechnoInsetColorID +
                                " AND Height=" + Height + " AND Width=" + Width);
                            if (rows.Count() == 0)
                            {
                                DataRow NewRow = DestinationDT.NewRow();
                                NewRow["Name"] = InsetTypeName + InsetColor;
                                NewRow["Height"] = Height;
                                NewRow["Width"] = Width;
                                NewRow["Count"] = Count;
                                NewRow["FrontID"] = Convert.ToInt32(Srows[0]["FrontID"]);
                                NewRow["InsetColorID"] = Convert.ToInt32(DT2.Rows[j]["InsetColorID"]);
                                NewRow["TechnoInsetColorID"] = TechnoInsetColorID;
                                DestinationDT.Rows.Add(NewRow);
                            }
                            else
                            {
                                rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count;
                            }
                        }
                    }
                }
            }
        }

        private void InsetsGridsOnly(DataTable SourceDT, ref DataTable DestinationDT, int Margin, bool OrderASC)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            DataTable DT3 = new DataTable();
            DataTable DT4 = new DataTable();

            using (DataView DV = new DataView(SourceDT))
            {
                DT1 = DV.ToTable(true, new string[] { "InsetTypeID" });
            }

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                string InsetTypeName = "";
                InsetTypeName = GetInsetTypeName(Convert.ToInt32(DT1.Rows[i]["InsetTypeID"])) + " ";
                using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]), "InsetColorID", DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "InsetColorID" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                        " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]), "TechnoInsetColorID", DataViewRowState.CurrentRows))
                    {
                        DT3 = DV.ToTable(true, new string[] { "TechnoInsetColorID" });
                    }
                    for (int x = 0; x < DT3.Rows.Count; x++)
                    {
                        using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                            " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) +
                            " AND TechnoInsetColorID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]), "Height, Width", DataViewRowState.CurrentRows))
                        {
                            DT4 = DV.ToTable(true, new string[] { "Height", "Width" });
                        }
                        for (int y = 0; y < DT4.Rows.Count; y++)
                        {
                            DataRow[] Srows = SourceDT.Select("InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                                " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND TechnoInsetColorID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]) +
                                " AND Height=" + Convert.ToInt32(DT4.Rows[y]["Height"]) + " AND Width=" + Convert.ToInt32(DT4.Rows[y]["Width"]));
                            if (Srows.Count() == 0)
                                continue;

                            int Count = 0;
                            int Height = Convert.ToInt32(DT4.Rows[y]["Height"]) - Margin;
                            int Width = Convert.ToInt32(DT4.Rows[y]["Width"]) - Margin;
                            int TechnoInsetColorID = Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]);

                            if (Height < 10 || Width < 10)
                                continue;

                            string InsetColor = GetInsetColorName(Convert.ToInt32(DT2.Rows[j]["InsetColorID"]));

                            if (Height < 10 || Width < 10)
                                continue;

                            foreach (DataRow item in Srows)
                                Count += Convert.ToInt32(item["Count"]);

                            string Name = string.Empty;
                            if (Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) == 685 || Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) == 688 || Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) == 29470)
                                Name = GetInsetTypeName(Convert.ToInt32(Srows[0]["InsetTypeID"])) + " 45 " + GetInsetColorName(Convert.ToInt32(DT2.Rows[j]["InsetColorID"]));
                            if (Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) == 686 || Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) == 687 || Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) == 29471)
                                Name = GetInsetTypeName(Convert.ToInt32(Srows[0]["InsetTypeID"])) + " 90 " + GetInsetColorName(Convert.ToInt32(DT2.Rows[j]["InsetColorID"]));
                            if (Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) == 860)
                                Name = GetInsetTypeName(Convert.ToInt32(Srows[0]["InsetTypeID"])) + " люкс " + GetInsetColorName(Convert.ToInt32(DT2.Rows[j]["InsetColorID"]));
                            DataRow[] rows = DestinationDT.Select("InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) +
                                " AND Height=" + Height + " AND Width=" + Width);
                            if (rows.Count() == 0)
                            {
                                DataRow NewRow = DestinationDT.NewRow();
                                NewRow["Name"] = InsetTypeName + InsetColor;
                                NewRow["Height"] = Height;
                                NewRow["Width"] = Width;
                                NewRow["Count"] = Count;
                                NewRow["FrontID"] = Convert.ToInt32(Srows[0]["FrontID"]);
                                NewRow["InsetColorID"] = Convert.ToInt32(DT2.Rows[j]["InsetColorID"]);
                                NewRow["TechnoInsetColorID"] = TechnoInsetColorID;
                                DestinationDT.Rows.Add(NewRow);
                            }
                            else
                            {
                                rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count;
                            }
                        }
                    }
                }
            }
        }

        private void InsetsPressOnly(DataTable SourceDT, ref DataTable DestinationDT,
            int WidthMargin, bool OrderASC)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            DataTable DT3 = new DataTable();
            DataTable DT4 = new DataTable();

            using (DataView DV = new DataView(SourceDT))
            {
                DT1 = DV.ToTable(true, new string[] { "InsetTypeID" });
            }

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                string InsetTypeName = "";
                InsetTypeName = GetInsetTypeName(Convert.ToInt32(DT1.Rows[i]["InsetTypeID"])) + " ";
                using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]), "InsetColorID", DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "InsetColorID" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    int InsetColorID = Convert.ToInt32(DT2.Rows[j]["InsetColorID"]);
                    if (InsetColorID != 44 || InsetColorID != 54 || InsetColorID != 51 || InsetColorID != 257 || InsetColorID != 57 ||
                        InsetColorID != 43 || InsetColorID != 45 || InsetColorID != 238 || InsetColorID != 52 || InsetColorID != 219 ||
                        InsetColorID != 218 || InsetColorID != 340 || InsetColorID != 220)
                        continue;

                    using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                        " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]), "TechnoInsetColorID", DataViewRowState.CurrentRows))
                    {
                        DT3 = DV.ToTable(true, new string[] { "TechnoInsetColorID" });
                    }
                    for (int x = 0; x < DT3.Rows.Count; x++)
                    {
                        using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                            " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) +
                            " AND TechnoInsetColorID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]), "Height, Width", DataViewRowState.CurrentRows))
                        {
                            DT4 = DV.ToTable(true, new string[] { "Height", "Width" });
                        }
                        for (int y = 0; y < DT4.Rows.Count; y++)
                        {
                            DataRow[] Srows = SourceDT.Select("InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                                " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND TechnoInsetColorID=" + Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]) +
                                " AND Height=" + Convert.ToInt32(DT4.Rows[y]["Height"]) + " AND Width=" + Convert.ToInt32(DT4.Rows[y]["Width"]));
                            if (Srows.Count() == 0)
                                continue;

                            int Count = 0;
                            int Height = Convert.ToInt32(DT4.Rows[y]["Height"]) - WidthMargin;
                            int Width = Convert.ToInt32(DT4.Rows[y]["Width"]) - WidthMargin;
                            int TechnoInsetColorID = Convert.ToInt32(DT3.Rows[x]["TechnoInsetColorID"]);
                            string InsetColor = GetInsetColorName(Convert.ToInt32(DT2.Rows[j]["InsetColorID"]));

                            if (Height < 10 || Width < 10)
                                continue;

                            if (Width <= 900)
                                continue;

                            foreach (DataRow item in Srows)
                                Count += Convert.ToInt32(item["Count"]);

                            DataRow[] rows = DestinationDT.Select("InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) +
                                " AND TechnoInsetColorID=" + TechnoInsetColorID +
                                " AND Height=" + Height + " AND Width=" + Width);
                            if (rows.Count() == 0)
                            {
                                DataRow NewRow = DestinationDT.NewRow();
                                NewRow["Name"] = InsetTypeName + InsetColor;
                                NewRow["Height"] = Height;
                                NewRow["Width"] = Width;
                                NewRow["Count"] = Count;
                                NewRow["FrontID"] = Convert.ToInt32(Srows[0]["FrontID"]);
                                NewRow["InsetColorID"] = Convert.ToInt32(DT2.Rows[j]["InsetColorID"]);
                                NewRow["TechnoInsetColorID"] = TechnoInsetColorID;
                                DestinationDT.Rows.Add(NewRow);
                            }
                            else
                            {
                                rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count;
                            }
                        }
                    }
                }
            }
        }

        private void NotArchDecorAssemblyByMainOrderToExcel(ref HSSFWorkbook hssfworkbook,
                                                                            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            int WorkAssignmentID, string BatchName)
        {
            if (NotArchDecorOrdersDT.Rows.Count == 0)
                return;

            DataTable DistMainOrdersDT = DistMainOrdersTable(NotArchDecorOrdersDT, true);
            DataTable DT = NotArchDecorOrdersDT.Clone();
            DataTable ZOVOrdersNames = new DataTable();
            DataTable MarketOrdersNames = new DataTable();
            string MainOrdersID = string.Empty;
            string ClientName = string.Empty;
            string OrderName = string.Empty;
            string Notes = string.Empty;

            if (DistMainOrdersDT.Select("GroupType=0").Count() > 0)
            {
                MainOrdersID = string.Empty;
                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=0"))
                    MainOrdersID += Convert.ToInt32(item["MainOrderID"]) + ",";
                MainOrdersID = MainOrdersID.Substring(0, MainOrdersID.Length - 1);
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT ClientName, DocNumber, MainOrderID FROM MainOrders" +
                    " INNER JOIN infiniu2_zovreference.dbo.Clients ON MainOrders.ClientID=infiniu2_zovreference.dbo.Clients.ClientID" +
                    " WHERE MainOrderID IN (" + MainOrdersID + ")", ConnectionStrings.ZOVOrdersConnectionString))
                {
                    DA.Fill(ZOVOrdersNames);
                }

                int RowIndex = 0;
                HSSFSheet sheet1 = hssfworkbook.CreateSheet("Не арки ЗОВ");
                sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

                sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
                sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
                sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
                sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

                sheet1.SetColumnWidth(0, 30 * 256);
                sheet1.SetColumnWidth(1, 20 * 256);
                sheet1.SetColumnWidth(2, 9 * 256);
                sheet1.SetColumnWidth(3, 6 * 256);
                sheet1.SetColumnWidth(4, 6 * 256);

                {
                    HSSFPatriarch patriarch = sheet1.CreateDrawingPatriarch();
                    int ind = 0;

                    CreateBarcode(WorkAssignmentID);
                    HSSFClientAnchor anchor = new HSSFClientAnchor(0, 1, 1000, 8, 5, ind, 6, ind + 3)
                    {
                        AnchorType = 2
                    };
                    string BarcodeNumber = GetBarcodeNumber(24, WorkAssignmentID);
                    string barcodeName = System.Environment.GetEnvironmentVariable("TEMP") + @"/" + BarcodeNumber + ".jpg";
                    HSSFPicture picture = patriarch.CreatePicture(anchor, LoadImage(barcodeName, hssfworkbook));

                    HSSFFont BarcodeFont = hssfworkbook.CreateFont();
                    BarcodeFont.FontHeightInPoints = 12;
                    BarcodeFont.FontName = "Calibri";

                    HSSFCellStyle BarcodeStyle = hssfworkbook.CreateCellStyle();
                    BarcodeStyle.Alignment = HSSFCellStyle.ALIGN_CENTER;
                    BarcodeStyle.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
                    BarcodeStyle.SetFont(BarcodeFont);

                    HSSFCell barcodeCell = sheet1.CreateRow(ind + 3).CreateCell(5);
                    barcodeCell.SetCellValue(BarcodeNumber.ToString());
                    barcodeCell.CellStyle = BarcodeStyle;
                    sheet1.AddMergedRegion(new NPOI.HSSF.Util.Region(ind + 3, 5, ind + 3, 6));
                }

                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=0"))
                {
                    DecorAssemblyDT.Clear();
                    DT.Clear();
                    int MainOrderID = Convert.ToInt32(item["MainOrderID"]);
                    DataRow[] rows = NotArchDecorOrdersDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item1 in rows)
                        DT.Rows.Add(item1.ItemArray);
                    AssemblyDecorCollect(DT, ref DecorAssemblyDT);

                    DataRow[] CRows = ZOVOrdersNames.Select("MainOrderID=" + Convert.ToInt32(item["MainOrderID"]));
                    if (CRows.Count() > 0)
                    {
                        ClientName = CRows[0]["ClientName"].ToString();
                        OrderName = CRows[0]["DocNumber"].ToString();
                    }
                    if (DecorAssemblyDT.Rows.Count > 0)
                    {
                        NotArchDecorAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                             Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DecorAssemblyDT,
                             WorkAssignmentID, BatchName, ClientName, string.Empty, OrderName, Notes, ref RowIndex);
                        RowIndex++;
                        NotArchDecorAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                            Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DecorAssemblyDT,
                            WorkAssignmentID, BatchName, ClientName, "ДУБЛЬ", OrderName, Notes, ref RowIndex);
                    }
                    RowIndex++;
                }
            }
            if (DistMainOrdersDT.Select("GroupType=1").Count() > 0)
            {
                MainOrdersID = string.Empty;
                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=1"))
                    MainOrdersID += Convert.ToInt32(item["MainOrderID"]) + ",";
                MainOrdersID = MainOrdersID.Substring(0, MainOrdersID.Length - 1);
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT ClientName, OrderNumber, MainOrderID, Notes FROM MainOrders" +
                    " INNER JOIN MegaOrders ON MainOrders.MegaOrderID=MegaOrders.MegaOrderID" +
                    " INNER JOIN infiniu2_marketingreference.dbo.Clients ON MegaOrders.ClientID=infiniu2_marketingreference.dbo.Clients.ClientID" +
                    " WHERE MainOrderID IN (" + MainOrdersID + ")", ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DA.Fill(MarketOrdersNames);
                }

                int RowIndex = 0;
                HSSFSheet sheet1 = hssfworkbook.CreateSheet("Не арки Маркетинг");
                sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

                sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
                sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
                sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
                sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

                sheet1.SetColumnWidth(0, 30 * 256);
                sheet1.SetColumnWidth(1, 20 * 256);
                sheet1.SetColumnWidth(2, 9 * 256);
                sheet1.SetColumnWidth(3, 6 * 256);
                sheet1.SetColumnWidth(4, 6 * 256);

                {
                    HSSFPatriarch patriarch = sheet1.CreateDrawingPatriarch();
                    int ind = 0;

                    CreateBarcode(WorkAssignmentID);
                    HSSFClientAnchor anchor = new HSSFClientAnchor(0, 1, 1000, 8, 5, ind, 6, ind + 3)
                    {
                        AnchorType = 2
                    };
                    string BarcodeNumber = GetBarcodeNumber(24, WorkAssignmentID);
                    string barcodeName = System.Environment.GetEnvironmentVariable("TEMP") + @"/" + BarcodeNumber + ".jpg";
                    HSSFPicture picture = patriarch.CreatePicture(anchor, LoadImage(barcodeName, hssfworkbook));

                    HSSFFont BarcodeFont = hssfworkbook.CreateFont();
                    BarcodeFont.FontHeightInPoints = 12;
                    BarcodeFont.FontName = "Calibri";

                    HSSFCellStyle BarcodeStyle = hssfworkbook.CreateCellStyle();
                    BarcodeStyle.Alignment = HSSFCellStyle.ALIGN_CENTER;
                    BarcodeStyle.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
                    BarcodeStyle.SetFont(BarcodeFont);

                    HSSFCell barcodeCell = sheet1.CreateRow(ind + 3).CreateCell(5);
                    barcodeCell.SetCellValue(BarcodeNumber.ToString());
                    barcodeCell.CellStyle = BarcodeStyle;
                    sheet1.AddMergedRegion(new NPOI.HSSF.Util.Region(ind + 3, 5, ind + 3, 6));
                }

                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=1"))
                {
                    DecorAssemblyDT.Clear();
                    DT.Clear();
                    int MainOrderID = Convert.ToInt32(item["MainOrderID"]);
                    DataRow[] rows = NotArchDecorOrdersDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item1 in rows)
                        DT.Rows.Add(item1.ItemArray);
                    AssemblyDecorCollect(DT, ref DecorAssemblyDT);

                    DataRow[] CRows = MarketOrdersNames.Select("MainOrderID=" + MainOrderID);
                    if (CRows.Count() > 0)
                    {
                        ClientName = CRows[0]["ClientName"].ToString();
                        Notes = CRows[0]["Notes"].ToString();
                        OrderName = "№" + CRows[0]["OrderNumber"].ToString() + "-" + MainOrderID;
                    }
                    if (DecorAssemblyDT.Rows.Count > 0)
                    {
                        NotArchDecorAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                             Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DecorAssemblyDT,
                             WorkAssignmentID, BatchName, ClientName, string.Empty, OrderName, Notes, ref RowIndex);
                        RowIndex++;
                        NotArchDecorAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                            Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DecorAssemblyDT,
                            WorkAssignmentID, BatchName, ClientName, "ДУБЛЬ", OrderName, Notes, ref RowIndex);
                    }
                    RowIndex++;
                }
            }
        }

        private string ProfileName(int ID)
        {
            string name = string.Empty;
            DataRow[] rows = ProfileNamesDT.Select("FrontConfigID=" + ID);
            if (rows.Count() > 0)
                name = rows[0]["TechStoreName"].ToString();
            return name;
        }

        private string ProfileName(int ID, int ProfileID)
        {
            string name = string.Empty;
            DataRow[] rows = ProfileNamesDT.Select("FrontConfigID=" + ID + " AND ProfileID=" + ProfileID);
            if (rows.Count() > 0)
                name = rows[0]["TechStoreName"].ToString();
            return name;
        }

        private void RapidGridsSingly(DataTable SourceDT, ref DataTable DestinationDT, int WidthMargin, int FrontType, string Front, bool OrderASC)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = DistHeightAndWidth(SourceDT, OrderASC);

            using (DataView DV = new DataView(SourceDT))
            {
                DT1 = DV.ToTable(true, new string[] { "ColorID" });
            }
            DT1 = OrderedFrameColors(DT1);

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    DataRow[] Srows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]));
                    if (Srows.Count() > 0)
                    {
                        int Count = 0;
                        int Height = Convert.ToInt32(DT2.Rows[j]["Height"]);
                        //string Front = ProfileName(Convert.ToInt32(SourceDT.Rows[0]["FrontID"]));
                        string Color = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                        string Notes = string.Empty;
                        foreach (DataRow item in Srows)
                        {
                            Count += Convert.ToInt32(item["Count"]);
                        }

                        //if (Height <= WidthMargin)
                        //    Height = WidthMargin;

                        DataRow[] rows = DestinationDT.Select("Front='" + Front + "' AND Color='" + Color + "' AND Height=" + Height);
                        if (rows.Count() == 0)
                        {
                            DataRow NewRow = DestinationDT.NewRow();
                            NewRow["Front"] = Front;
                            NewRow["Color"] = Color;
                            NewRow["Height"] = Height;
                            NewRow["Count"] = Count * 2;
                            NewRow["VitrinaCount"] = 0;
                            NewRow["ProfileType"] = Convert.ToInt32(SourceDT.Rows[0]["FrontID"]);
                            NewRow["ColorType"] = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                            NewRow["FrontType"] = FrontType;
                            DestinationDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count * 2;
                            rows[0]["VitrinaCount"] = 0;
                        }
                    }
                    Srows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Height"]));
                    if (Srows.Count() > 0)
                    {
                        int Count = 0;
                        int Height = Convert.ToInt32(DT2.Rows[j]["Height"]);
                        //string Front = ProfileName(Convert.ToInt32(SourceDT.Rows[0]["FrontID"]));
                        string Color = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                        string Notes = string.Empty;
                        foreach (DataRow item in Srows)
                        {
                            Count += Convert.ToInt32(item["Count"]);
                        }

                        //if (Height <= WidthMargin)
                        //    Height = WidthMargin;

                        DataRow[] rows = DestinationDT.Select("Front='" + Front + "' AND Color='" + Color + "' AND Height=" + Height);
                        if (rows.Count() == 0)
                        {
                            DataRow NewRow = DestinationDT.NewRow();
                            NewRow["Front"] = Front;
                            NewRow["Color"] = Color;
                            NewRow["Height"] = Height;
                            NewRow["Count"] = Count * 2;
                            NewRow["VitrinaCount"] = 0;
                            NewRow["ProfileType"] = Convert.ToInt32(SourceDT.Rows[0]["FrontID"]);
                            NewRow["ColorType"] = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                            NewRow["FrontType"] = FrontType;
                            DestinationDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count * 2;
                            rows[0]["VitrinaCount"] = 0;
                        }
                    }
                }
            }
        }

        private void RapidSimpleBergamoSingly(DataTable SourceDT, ref DataTable DestinationDT, int WidthMargin, int FrontType, int ProfileType, string Front, bool OrderASC)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = DistHeightAndWidth(SourceDT, OrderASC);

            using (DataView DV = new DataView(SourceDT))
            {
                DT1 = DV.ToTable(true, new string[] { "ColorID" });
            }
            DT1 = OrderedFrameColors(DT1);

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    DataRow[] Srows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]));
                    if (Srows.Count() > 0)
                    {
                        int Count = 0;
                        int BoxCount = 0;
                        int VitrinaCount = 0;
                        int Height = Convert.ToInt32(DT2.Rows[j]["Height"]);
                        //string Front = ProfileName(Convert.ToInt32(SourceDT.Rows[0]["FrontID"]));
                        string Color = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                        string Notes = string.Empty;
                        foreach (DataRow item in Srows)
                        {
                            if (Convert.ToInt32(item["InsetTypeID"]) != 1 && (Convert.ToInt32(item["Height"]) - 10 < WidthMargin || Convert.ToInt32(item["Width"]) - 10 < WidthMargin))
                                BoxCount += Convert.ToInt32(item["Count"]);
                            if (Convert.ToInt32(item["InsetTypeID"]) == 1)
                                VitrinaCount += Convert.ToInt32(item["Count"]);
                            Count += Convert.ToInt32(item["Count"]);
                        }

                        //if (Height <= WidthMargin)
                        //    Height = WidthMargin;

                        DataRow[] rows = DestinationDT.Select("Front='" + Front + "' AND Color='" + Color + "' AND Height=" + Height);
                        if (rows.Count() == 0)
                        {
                            DataRow NewRow = DestinationDT.NewRow();
                            NewRow["Front"] = Front;
                            NewRow["Color"] = Color;
                            NewRow["Height"] = Height;
                            NewRow["Count"] = Count * 2;
                            NewRow["BoxCount"] = BoxCount * 2;
                            NewRow["VitrinaCount"] = VitrinaCount * 2;
                            NewRow["ProfileType"] = ProfileType;
                            NewRow["ColorType"] = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                            NewRow["FrontType"] = FrontType;
                            DestinationDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count * 2;
                            rows[0]["BoxCount"] = Convert.ToInt32(rows[0]["BoxCount"]) + BoxCount * 2;
                            rows[0]["VitrinaCount"] = Convert.ToInt32(rows[0]["VitrinaCount"]) + VitrinaCount * 2;
                        }
                    }
                    Srows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Height"]));
                    if (Srows.Count() > 0)
                    {
                        int Count = 0;
                        int BoxCount = 0;
                        int VitrinaCount = 0;
                        int Height = Convert.ToInt32(DT2.Rows[j]["Height"]);
                        //string Front = ProfileName(Convert.ToInt32(SourceDT.Rows[0]["FrontID"]));
                        string Color = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                        string Notes = string.Empty;
                        foreach (DataRow item in Srows)
                        {
                            if (Convert.ToInt32(item["InsetTypeID"]) != 1 && (Convert.ToInt32(item["Height"]) - 10 < WidthMargin || Convert.ToInt32(item["Width"]) - 10 < WidthMargin))
                                BoxCount += Convert.ToInt32(item["Count"]);
                            if (Convert.ToInt32(item["InsetTypeID"]) == 1)
                                VitrinaCount += Convert.ToInt32(item["Count"]);
                            Count += Convert.ToInt32(item["Count"]);
                        }

                        //if (Height <= WidthMargin)
                        //    Height = WidthMargin;

                        DataRow[] rows = DestinationDT.Select("Front='" + Front + "' AND Color='" + Color + "' AND Height=" + Height);
                        if (rows.Count() == 0)
                        {
                            DataRow NewRow = DestinationDT.NewRow();
                            NewRow["Front"] = Front;
                            NewRow["Color"] = Color;
                            NewRow["Height"] = Height;
                            NewRow["Count"] = Count * 2;
                            NewRow["BoxCount"] = BoxCount * 2;
                            NewRow["VitrinaCount"] = VitrinaCount * 2;
                            NewRow["ProfileType"] = ProfileType;
                            NewRow["ColorType"] = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                            NewRow["FrontType"] = FrontType;
                            DestinationDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count * 2;
                            rows[0]["BoxCount"] = Convert.ToInt32(rows[0]["BoxCount"]) + BoxCount * 2;
                            rows[0]["VitrinaCount"] = Convert.ToInt32(rows[0]["VitrinaCount"]) + VitrinaCount * 2;
                        }
                    }
                }
            }
        }

        private void RapidSimpleImpostSingly(DataTable SourceDT, ref DataTable DestinationDT, int WidthMargin, int FrontType, string Front, bool OrderASC)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = DistSizesTable(SourceDT, OrderASC);

            using (DataView DV = new DataView(SourceDT))
            {
                DV.RowFilter = "TechnoColorID <> -1";
                DT1 = DV.ToTable(true, new string[] { "TechnoColorID" });
            }
            DT1 = OrderedTechnoFrameColors(DT1);

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    DataRow[] Srows = SourceDT.Select("TechnoColorID=" + Convert.ToInt32(DT1.Rows[i]["TechnoColorID"]) +
                        " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) + " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]));
                    if (Srows.Count() > 0)
                    {
                        int Count = 0;
                        int BoxCount = 0;
                        int VitrinaCount = 0;
                        int Height = Convert.ToInt32(DT2.Rows[j]["Height"]);
                        int Width = Convert.ToInt32(DT2.Rows[j]["Width"]);
                        int impostCount = StandardImpostCount(Convert.ToInt32(Srows[0]["FrontID"]), Height, Width);

                        Front = ProfileName(Convert.ToInt32(SourceDT.Rows[0]["FrontConfigID"]), Convert.ToInt32(SourceDT.Rows[0]["TechnoProfileID"]));
                        string Color = GetColorName(Convert.ToInt32(DT1.Rows[i]["TechnoColorID"]));
                        string Notes = string.Empty;
                        foreach (DataRow item in Srows)
                        {
                            if (Convert.ToInt32(item["InsetTypeID"]) != 1 && (Convert.ToInt32(item["Height"]) - 10 < WidthMargin || Convert.ToInt32(item["Width"]) - 10 < WidthMargin))
                                BoxCount += Convert.ToInt32(item["Count"]);
                            if (Convert.ToInt32(item["InsetTypeID"]) == 1)
                                VitrinaCount += Convert.ToInt32(item["Count"]);

                            Count += Convert.ToInt32(item["Count"]);
                        }

                        //if (Height <= WidthMargin)
                        //    Height = WidthMargin;

                        int impostLength = Width - 100;

                        if (Height < Width)
                        {
                            impostLength = Height - 100;
                        }
                        string impostFilter = "Front='" + Front + "' AND Color='" + Color + "' AND Height=" + impostLength;

                        DataRow[] rows = DestinationDT.Select(impostFilter);
                        if (rows.Count() == 0)
                        {
                            DataRow NewRow = DestinationDT.NewRow();
                            NewRow["Front"] = Front;
                            NewRow["Color"] = Color;
                            NewRow["Height"] = impostLength;
                            NewRow["Count"] = Count * impostCount;
                            NewRow["BoxCount"] = BoxCount * 2;
                            NewRow["VitrinaCount"] = VitrinaCount * 2;
                            NewRow["ProfileType"] = Convert.ToInt32(SourceDT.Rows[0]["FrontID"]);
                            NewRow["ColorType"] = Convert.ToInt32(DT1.Rows[i]["TechnoColorID"]);
                            NewRow["FrontType"] = FrontType;
                            DestinationDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count * impostCount;
                            rows[0]["BoxCount"] = Convert.ToInt32(rows[0]["BoxCount"]) + BoxCount * 2;
                            rows[0]["VitrinaCount"] = Convert.ToInt32(rows[0]["VitrinaCount"]) + VitrinaCount * 2;
                        }
                    }
                }
            }
        }

        private void RapidSimpleSingly(DataTable SourceDT, ref DataTable DestinationDT, int WidthMargin, int FrontType, string Front, bool OrderASC)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = DistHeightAndWidth(SourceDT, OrderASC);

            using (DataView DV = new DataView(SourceDT))
            {
                //DV.RowFilter = "TechnoColorID = -1";
                DT1 = DV.ToTable(true, new string[] { "ColorID" });
            }
            DT1 = OrderedFrameColors(DT1);

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    DataRow[] Srows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]));
                    if (Srows.Count() > 0)
                    {
                        int Count = 0;
                        int BoxCount = 0;
                        int VitrinaCount = 0;
                        int ImpostCount = 0;
                        int Height = Convert.ToInt32(DT2.Rows[j]["Height"]);

                        //string Front = ProfileName(Convert.ToInt32(SourceDT.Rows[0]["FrontID"]));
                        string Color = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                        string Notes = string.Empty;
                        foreach (DataRow item in Srows)
                        {
                            if (Convert.ToInt32(item["InsetTypeID"]) != 1 && (Convert.ToInt32(item["Height"]) - 10 < WidthMargin || Convert.ToInt32(item["Width"]) - 10 < WidthMargin))
                                BoxCount += Convert.ToInt32(item["Count"]);
                            if (Convert.ToInt32(item["InsetTypeID"]) == 1)
                                VitrinaCount += Convert.ToInt32(item["Count"]);

                            if (Convert.ToInt32(item["TechnoColorID"]) != -1)
                            {
                                ImpostCount += Convert.ToInt32(item["Count"]) * StandardImpostCount(Convert.ToInt32(item["FrontID"]),
                                    Convert.ToInt32(item["Height"]),
                                    Convert.ToInt32(item["Width"]));
                            }
                            Count += Convert.ToInt32(item["Count"]);
                        }

                        //if (Height <= WidthMargin)
                        //    Height = WidthMargin;

                        DataRow[] rows = DestinationDT.Select("Front='" + Front + "' AND Color='" + Color + "' AND Height=" + Height);
                        if (rows.Count() == 0)
                        {
                            DataRow NewRow = DestinationDT.NewRow();
                            NewRow["Front"] = Front;
                            NewRow["Color"] = Color;
                            NewRow["Height"] = Height;
                            NewRow["Count"] = Count * 2;
                            NewRow["BoxCount"] = BoxCount * 2;
                            NewRow["VitrinaCount"] = VitrinaCount * 2;
                            NewRow["ImpostCount"] = ImpostCount * 2;
                            NewRow["ProfileType"] = Convert.ToInt32(SourceDT.Rows[0]["FrontID"]);
                            NewRow["ColorType"] = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                            NewRow["FrontType"] = FrontType;
                            DestinationDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count * 2;
                            rows[0]["BoxCount"] = Convert.ToInt32(rows[0]["BoxCount"]) + BoxCount * 2;
                            rows[0]["VitrinaCount"] = Convert.ToInt32(rows[0]["VitrinaCount"]) + VitrinaCount * 2;
                            rows[0]["ImpostCount"] = Convert.ToInt32(rows[0]["ImpostCount"]) + ImpostCount * 2;
                        }
                    }
                    Srows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Height"]));
                    if (Srows.Count() > 0)
                    {
                        int Count = 0;
                        int BoxCount = 0;
                        int VitrinaCount = 0;
                        int ImpostCount = 0;
                        int Height = Convert.ToInt32(DT2.Rows[j]["Height"]);
                        //string Front = ProfileName(Convert.ToInt32(SourceDT.Rows[0]["FrontID"]));
                        string Color = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                        string Notes = string.Empty;
                        foreach (DataRow item in Srows)
                        {
                            if (Convert.ToInt32(item["InsetTypeID"]) != 1 && (Convert.ToInt32(item["Height"]) - 10 < WidthMargin || Convert.ToInt32(item["Width"]) - 10 < WidthMargin))
                                BoxCount += Convert.ToInt32(item["Count"]);
                            if (Convert.ToInt32(item["InsetTypeID"]) == 1)
                                VitrinaCount += Convert.ToInt32(item["Count"]);

                            if (Convert.ToInt32(item["TechnoColorID"]) != -1)
                            {
                                ImpostCount += Convert.ToInt32(item["Count"]) * StandardImpostCount(Convert.ToInt32(item["FrontID"]),
                                    Convert.ToInt32(item["Height"]),
                                    Convert.ToInt32(item["Width"]));
                            }
                            Count += Convert.ToInt32(item["Count"]);
                        }

                        //if (Height <= WidthMargin)
                        //    Height = WidthMargin;

                        DataRow[] rows = DestinationDT.Select("Front='" + Front + "' AND Color='" + Color + "' AND Height=" + Height);
                        if (rows.Count() == 0)
                        {
                            DataRow NewRow = DestinationDT.NewRow();
                            NewRow["Front"] = Front;
                            NewRow["Color"] = Color;
                            NewRow["Height"] = Height;
                            NewRow["Count"] = Count * 2;
                            NewRow["BoxCount"] = BoxCount * 2;
                            NewRow["VitrinaCount"] = VitrinaCount * 2;
                            NewRow["ImpostCount"] = ImpostCount * 2;
                            NewRow["ProfileType"] = Convert.ToInt32(SourceDT.Rows[0]["FrontID"]);
                            NewRow["ColorType"] = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                            NewRow["FrontType"] = FrontType;
                            DestinationDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count * 2;
                            rows[0]["BoxCount"] = Convert.ToInt32(rows[0]["BoxCount"]) + BoxCount * 2;
                            rows[0]["VitrinaCount"] = Convert.ToInt32(rows[0]["VitrinaCount"]) + VitrinaCount * 2;
                            rows[0]["ImpostCount"] = Convert.ToInt32(rows[0]["ImpostCount"]) + ImpostCount * 2;
                        }
                    }
                }
            }
        }

        private void RapidVitrinaSingly(DataTable SourceDT, ref DataTable DestinationDT, int WidthMargin, int FrontType, string Front, bool OrderASC)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = DistHeightAndWidth(SourceDT, OrderASC);

            using (DataView DV = new DataView(SourceDT))
            {
                DT1 = DV.ToTable(true, new string[] { "ColorID" });
            }
            DT1 = OrderedFrameColors(DT1);

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    DataRow[] Srows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]));
                    if (Srows.Count() > 0)
                    {
                        int Count = 0;
                        int VitrinaCount = 0;
                        int Height = Convert.ToInt32(DT2.Rows[j]["Height"]);
                        //string Front = ProfileName(Convert.ToInt32(SourceDT.Rows[0]["FrontID"]));
                        string Color = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                        string Notes = string.Empty;
                        foreach (DataRow item in Srows)
                        {
                            if (Convert.ToInt32(item["InsetTypeID"]) == 1)
                                VitrinaCount += Convert.ToInt32(item["Count"]);
                            Count += Convert.ToInt32(item["Count"]);
                        }

                        //if (Height <= WidthMargin)
                        //    Height = WidthMargin;

                        DataRow[] rows = DestinationDT.Select("Front='" + Front + "' AND Color='" + Color + "' AND Height=" + Height);
                        if (rows.Count() == 0)
                        {
                            DataRow NewRow = DestinationDT.NewRow();
                            NewRow["Front"] = Front;
                            NewRow["Color"] = Color;
                            NewRow["Height"] = Height;
                            NewRow["Count"] = Count * 2;
                            NewRow["VitrinaCount"] = VitrinaCount * 2;
                            NewRow["ProfileType"] = Convert.ToInt32(SourceDT.Rows[0]["FrontID"]);
                            NewRow["ColorType"] = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                            NewRow["FrontType"] = FrontType;
                            DestinationDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count * 2;
                            rows[0]["VitrinaCount"] = Convert.ToInt32(rows[0]["VitrinaCount"]) + VitrinaCount * 2;
                        }
                    }
                    Srows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Height"]));
                    if (Srows.Count() > 0)
                    {
                        int Count = 0;
                        int VitrinaCount = 0;
                        int Height = Convert.ToInt32(DT2.Rows[j]["Height"]);
                        //string Front = ProfileName(Convert.ToInt32(SourceDT.Rows[0]["FrontID"]));
                        string Color = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                        string Notes = string.Empty;
                        foreach (DataRow item in Srows)
                        {
                            if (Convert.ToInt32(item["InsetTypeID"]) == 1)
                                VitrinaCount += Convert.ToInt32(item["Count"]);
                            Count += Convert.ToInt32(item["Count"]);
                        }

                        //if (Height <= WidthMargin)
                        //    Height = WidthMargin;

                        DataRow[] rows = DestinationDT.Select("Front='" + Front + "' AND Color='" + Color + "' AND Height=" + Height);
                        if (rows.Count() == 0)
                        {
                            DataRow NewRow = DestinationDT.NewRow();
                            NewRow["Front"] = Front;
                            NewRow["Color"] = Color;
                            NewRow["Height"] = Height;
                            NewRow["Count"] = Count * 2;
                            NewRow["VitrinaCount"] = VitrinaCount * 2;
                            NewRow["ProfileType"] = Convert.ToInt32(SourceDT.Rows[0]["FrontID"]);
                            NewRow["ColorType"] = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                            NewRow["FrontType"] = FrontType;
                            DestinationDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count * 2;
                            rows[0]["VitrinaCount"] = Convert.ToInt32(rows[0]["VitrinaCount"]) + VitrinaCount * 2;
                        }
                    }
                }
            }
        }

        private int StandardImpostCount(int ID, int Height, int Width)
        {
            int count = -1;
            DataRow[] rows = StandardImpostDataTable.Select("FrontID=" + ID + " AND Height=" + Height + " AND Width=" + Width);
            if (rows.Count() > 0)
                count = Convert.ToInt32(rows[0]["Count"]);
            return count;
        }

        private void SummaryOrders(DataTable SourceDT, DataTable SimpleDT, DataTable VitrinaDT, DataTable GridsDT, string FrontName, ref decimal AllSquare)
        {
            SummOrdersDT.Dispose();
            SummOrdersDT = new DataTable();
            SummOrdersDT.Columns.Add(new DataColumn("Sizes", Type.GetType("System.String")));
            SummOrdersDT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            SummOrdersDT.Columns.Add(new DataColumn("Width", Type.GetType("System.Int32")));

            DataRow NewRow1 = SummOrdersDT.NewRow();
            NewRow1["Sizes"] = "Профиль";
            SummOrdersDT.Rows.Add(NewRow1);

            DataTable DistinctSizesDT = DistSizesTable(SourceDT, true);

            CollectOrders(DistinctSizesDT, SimpleDT, ref SummOrdersDT, 2, FrontName);
            CollectOrders(DistinctSizesDT, VitrinaDT, ref SummOrdersDT, 1, FrontName);
            CollectOrders(DistinctSizesDT, GridsDT, ref SummOrdersDT, 3, FrontName);
            SummOrdersDT.Columns.Add(new DataColumn("TotalAmount", Type.GetType("System.String")));

            using (DataView DV = new DataView(SummOrdersDT.Copy()))
            {
                SummOrdersDT.Clear();
                DV.Sort = "Height, Width";
                SummOrdersDT = DV.ToTable();
            }

            DataRow NewRow2 = SummOrdersDT.NewRow();
            NewRow2["Sizes"] = "Квадратура";
            SummOrdersDT.Rows.Add(NewRow2);

            DataRow NewRow3 = SummOrdersDT.NewRow();
            NewRow3["Sizes"] = "Кол-во";
            SummOrdersDT.Rows.Add(NewRow3);

            decimal TotalSquare = 0;
            int TotalCount = 0;

            for (int y = 0; y < SummOrdersDT.Columns.Count; y++)
            {
                if (SummOrdersDT.Columns[y].ColumnName == "Sizes" || SummOrdersDT.Columns[y].ColumnName == "TotalAmount"
                    || SummOrdersDT.Columns[y].ColumnName == "Height" || SummOrdersDT.Columns[y].ColumnName == "Width")
                    continue;
                decimal Square = 0;
                int Count = 0;

                for (int x = 0; x < SummOrdersDT.Rows.Count; x++)
                {
                    if (SummOrdersDT.Rows[x]["Height"] != DBNull.Value && SummOrdersDT.Rows[x]["Width"] != DBNull.Value && SummOrdersDT.Rows[x][y] != DBNull.Value)
                    {
                        Square += Convert.ToDecimal(SummOrdersDT.Rows[x]["Height"]) * Convert.ToDecimal(SummOrdersDT.Rows[x]["Width"]) * Convert.ToDecimal(SummOrdersDT.Rows[x][y]) / 1000000;
                        Count += Convert.ToInt32(SummOrdersDT.Rows[x][y]);
                        TotalSquare += Convert.ToDecimal(SummOrdersDT.Rows[x]["Height"]) * Convert.ToDecimal(SummOrdersDT.Rows[x]["Width"]) * Convert.ToDecimal(SummOrdersDT.Rows[x][y]) / 1000000;
                        TotalCount += Convert.ToInt32(SummOrdersDT.Rows[x][y]);
                    }
                }
                Square = decimal.Round(Square, 3, MidpointRounding.AwayFromZero);
                SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 1][y] = Count;
                SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 2][y] = Square;
            }
            AllSquare += TotalSquare;
            TotalSquare = decimal.Round(TotalSquare, 3, MidpointRounding.AwayFromZero);
            SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 1]["TotalAmount"] = TotalCount;
            SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 2]["TotalAmount"] = TotalSquare;
        }

        private void SummaryOrders(DataTable SourceDT, DataTable SimpleDT, DataTable VitrinaDT, DataTable GridsDT, DataTable GlassDT, string FrontName, ref decimal AllSquare)
        {
            SummOrdersDT.Dispose();
            SummOrdersDT = new DataTable();
            SummOrdersDT.Columns.Add(new DataColumn("Sizes", Type.GetType("System.String")));
            SummOrdersDT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            SummOrdersDT.Columns.Add(new DataColumn("Width", Type.GetType("System.Int32")));

            DataRow NewRow1 = SummOrdersDT.NewRow();
            NewRow1["Sizes"] = "Профиль";
            SummOrdersDT.Rows.Add(NewRow1);

            DataTable DistinctSizesDT = DistSizesTable(SourceDT, true);

            CollectOrders(DistinctSizesDT, SimpleDT, ref SummOrdersDT, 2, FrontName);
            CollectOrders(DistinctSizesDT, VitrinaDT, ref SummOrdersDT, 1, FrontName);
            CollectOrders(DistinctSizesDT, GridsDT, ref SummOrdersDT, 3, FrontName);
            CollectOrders(DistinctSizesDT, GlassDT, ref SummOrdersDT, 4, FrontName);
            SummOrdersDT.Columns.Add(new DataColumn("TotalAmount", Type.GetType("System.String")));

            using (DataView DV = new DataView(SummOrdersDT.Copy()))
            {
                SummOrdersDT.Clear();
                DV.Sort = "Height, Width";
                SummOrdersDT = DV.ToTable();
            }

            DataRow NewRow2 = SummOrdersDT.NewRow();
            NewRow2["Sizes"] = "Квадратура";
            SummOrdersDT.Rows.Add(NewRow2);

            DataRow NewRow3 = SummOrdersDT.NewRow();
            NewRow3["Sizes"] = "Кол-во";
            SummOrdersDT.Rows.Add(NewRow3);

            decimal TotalSquare = 0;
            int TotalCount = 0;

            for (int y = 0; y < SummOrdersDT.Columns.Count; y++)
            {
                if (SummOrdersDT.Columns[y].ColumnName == "Sizes" || SummOrdersDT.Columns[y].ColumnName == "TotalAmount"
                    || SummOrdersDT.Columns[y].ColumnName == "Height" || SummOrdersDT.Columns[y].ColumnName == "Width")
                    continue;
                decimal Square = 0;
                int Count = 0;

                for (int x = 0; x < SummOrdersDT.Rows.Count; x++)
                {
                    if (SummOrdersDT.Rows[x]["Height"] != DBNull.Value && SummOrdersDT.Rows[x]["Width"] != DBNull.Value && SummOrdersDT.Rows[x][y] != DBNull.Value)
                    {
                        Square += Convert.ToDecimal(SummOrdersDT.Rows[x]["Height"]) * Convert.ToDecimal(SummOrdersDT.Rows[x]["Width"]) * Convert.ToDecimal(SummOrdersDT.Rows[x][y]) / 1000000;
                        Count += Convert.ToInt32(SummOrdersDT.Rows[x][y]);
                        TotalSquare += Convert.ToDecimal(SummOrdersDT.Rows[x]["Height"]) * Convert.ToDecimal(SummOrdersDT.Rows[x]["Width"]) * Convert.ToDecimal(SummOrdersDT.Rows[x][y]) / 1000000;
                        TotalCount += Convert.ToInt32(SummOrdersDT.Rows[x][y]);
                    }
                }
                Square = decimal.Round(Square, 3, MidpointRounding.AwayFromZero);
                SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 1][y] = Count;
                SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 2][y] = Square;
            }
            AllSquare += TotalSquare;
            TotalSquare = decimal.Round(TotalSquare, 3, MidpointRounding.AwayFromZero);
            SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 1]["TotalAmount"] = TotalCount;
            SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 2]["TotalAmount"] = TotalSquare;
        }

        private void TotalSum(DataTable SourceDT, ref DataTable DestinationDT, string ProfileName1, string ProfileName2,
                                                                                            int WidthMargin, int WidthMin, int HeightMargin)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            string SizesASC = string.Empty;

            using (DataView DV = new DataView(SourceDT))
            {
                DT1 = DV.ToTable(true, new string[] { "ColorID" });
            }

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                SizesASC = "Height ASC";

                using (DataView DV = new DataView(SourceDT, "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]), SizesASC, DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "Height" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    DataRow[] Srows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]));
                    if (Srows.Count() == 0)
                        continue;

                    decimal SticksCount = 0;
                    int Count = 0;
                    int Height = Convert.ToInt32(DT2.Rows[j]["Height"]) - 1;
                    string FrameColor = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                    foreach (DataRow item in Srows)
                        Count += Convert.ToInt32(item["Count"]);

                    if (Height <= HeightMargin)
                        Height = HeightMargin;

                    SticksCount = (Height + 4) * Count * 2;
                    SticksCount = SticksCount * 1.15m / 2620;
                    SticksCount = decimal.Round(SticksCount, 1, MidpointRounding.AwayFromZero);

                    DataRow[] rows = DestinationDT.Select("Front='" + ProfileName2 + "' AND Color='" + FrameColor + "'");
                    if (rows.Count() == 0)
                    {
                        DataRow NewRow = DestinationDT.NewRow();
                        NewRow["Front"] = ProfileName2;
                        NewRow["Color"] = FrameColor;
                        NewRow["SticksCount"] = SticksCount;
                        DestinationDT.Rows.Add(NewRow);
                    }
                    else
                        rows[0]["SticksCount"] = Convert.ToDecimal(rows[0]["SticksCount"]) + SticksCount;
                }

                SizesASC = "Width ASC";

                using (DataView DV = new DataView(SourceDT, "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]), SizesASC, DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "Width" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    DataRow[] Srows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]));
                    if (Srows.Count() == 0)
                        continue;

                    decimal SticksCount = 0;
                    int Count = 0;
                    int Height = Convert.ToInt32(DT2.Rows[j]["Width"]) - WidthMargin;
                    string FrameColor = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                    foreach (DataRow item in Srows)
                        Count += Convert.ToInt32(item["Count"]);

                    if (Height <= WidthMin)
                        Height = WidthMin;

                    SticksCount = (Height + 4) * Count * 2;
                    SticksCount = SticksCount * 1.15m / 2620;
                    SticksCount = decimal.Round(SticksCount, 1, MidpointRounding.AwayFromZero);

                    DataRow[] rows = DestinationDT.Select("Front='" + ProfileName1 + "' AND Color='" + FrameColor + "'");
                    if (rows.Count() == 0)
                    {
                        DataRow NewRow = DestinationDT.NewRow();
                        NewRow["Front"] = ProfileName1;
                        NewRow["Color"] = FrameColor;
                        NewRow["SticksCount"] = SticksCount;
                        DestinationDT.Rows.Add(NewRow);
                    }
                    else
                        rows[0]["SticksCount"] = Convert.ToDecimal(rows[0]["SticksCount"]) + SticksCount;
                }
            }
        }

        private void TotalSumTechno4(DataTable SourceDT, ref DataTable DestinationDT, string ProfileName1, string ProfileName2,
            int WidthMargin, int WidthMin, int HeightMargin, int HeightNarrowMargin)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            string SizesASC = string.Empty;

            using (DataView DV = new DataView(SourceDT))
            {
                DT1 = DV.ToTable(true, new string[] { "ColorID" });
            }

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                SizesASC = "Height ASC";

                using (DataView DV = new DataView(SourceDT, "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]), SizesASC, DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "Height" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    DataRow[] Srows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]));
                    if (Srows.Count() == 0)
                        continue;

                    decimal SticksCount = 0;
                    int Count = 0;
                    int Height = Convert.ToInt32(DT2.Rows[j]["Height"]) - 1;
                    string FrameColor = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                    foreach (DataRow item in Srows)
                        Count += Convert.ToInt32(item["Count"]);

                    SticksCount = (Height + 4) * Count * 2;
                    SticksCount = SticksCount * 1.15m / 2620;
                    SticksCount = decimal.Round(SticksCount, 1, MidpointRounding.AwayFromZero);

                    DataRow[] rows = DestinationDT.Select("Front='" + ProfileName2 + "' AND Color='" + FrameColor + "'");
                    if (rows.Count() == 0)
                    {
                        DataRow NewRow = DestinationDT.NewRow();
                        NewRow["Front"] = ProfileName2;
                        NewRow["Color"] = FrameColor;
                        NewRow["SticksCount"] = SticksCount;
                        DestinationDT.Rows.Add(NewRow);
                    }
                    else
                        rows[0]["SticksCount"] = Convert.ToDecimal(rows[0]["SticksCount"]) + SticksCount;
                }

                SizesASC = "Width ASC";

                using (DataView DV = new DataView(SourceDT, "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]), SizesASC, DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "Width" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    DataRow[] Srows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]));
                    if (Srows.Count() == 0)
                        continue;

                    decimal SticksCount = 0;
                    int Count = 0;
                    int Height = Convert.ToInt32(DT2.Rows[j]["Width"]);
                    string FrameColor = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                    foreach (DataRow item in Srows)
                        Count += Convert.ToInt32(item["Count"]);

                    if (Height < HeightMargin)
                        Height = Height - HeightNarrowMargin;
                    else
                        Height = Height - WidthMargin;
                    if (Height <= WidthMin)
                        Height = WidthMin;

                    SticksCount = (Height + 4) * Count * 2;
                    SticksCount = SticksCount * 1.15m / 2620;
                    SticksCount = decimal.Round(SticksCount, 1, MidpointRounding.AwayFromZero);

                    DataRow[] rows = DestinationDT.Select("Front='" + ProfileName1 + "' AND Color='" + FrameColor + "'");
                    if (rows.Count() == 0)
                    {
                        DataRow NewRow = DestinationDT.NewRow();
                        NewRow["Front"] = ProfileName1;
                        NewRow["Color"] = FrameColor;
                        NewRow["SticksCount"] = SticksCount;
                        DestinationDT.Rows.Add(NewRow);
                    }
                    else
                        rows[0]["SticksCount"] = Convert.ToDecimal(rows[0]["SticksCount"]) + SticksCount;
                }
            }
        }

        #endregion Private Methods
    }
}