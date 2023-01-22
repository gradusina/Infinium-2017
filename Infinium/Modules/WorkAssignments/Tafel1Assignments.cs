using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.UserModel.Contrib;
using NPOI.HSSF.Util;

using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;

namespace Infinium.Modules.WorkAssignments
{
    public class Tafel1Assignments : IAllFrontParameterName
    {
        public DataTable FrameColorsDataTable = null;
        public DataTable FrontsDataTable = null;
        public DataTable InsetColorsDataTable = null;
        public DataTable InsetTypesDataTable = null;
        public DataTable PatinaDataTable = null;
        public DataTable TechStoreDataTable = null;
        private readonly string _appovingUser = "Егорченко Р.П.";
        private DataTable ArchDecorOrdersDT;
        private DateTime CurrentDate;
        private DataTable DecorAssemblyDT;
        private DataTable DecorDT;
        private DataTable DecorParametersDT;
        private DataTable DeyingDT;
        private readonly FileManager FM = new FileManager();
        private DataTable GridsDecorOrdersDT;
        private bool HeightLess180 = false;
        private DataTable InsetDT;
        private DataTable NotArchDecorOrdersDT;
        private DataTable NotCurvedAssemblyDT;
        private DataTable NotCurvedOrdersDT;
        private DataTable PatinaRALDataTable = null;
        private DataTable SimpleDT;
        public int WorkAssignmentID { get; set; } = 0;

        public Tafel1Assignments()
        {
        }

        public void ArchDecorAssemblyToExcelSingly(ref HSSFWorkbook hssfworkbook, ref HSSFSheet sheet1,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            DataTable DT, string BatchName, string ClientName, string OperationName, string OrderName, string Notes, ref int RowIndex)
        {
            HSSFCell cell = null;

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

        private void AssemblyOrders(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            ref HSSFSheet sheet1, DataTable DT, string BatchName, string ClientName, ref int RowIndex)
        {
            HSSFCell cell = null;
            
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, $"Распечатал: Дата/время { CurrentDate.ToString("dd.MM.yyyy HH:mm") }" +
                $" \r\n ФИО: { Security.CurrentUserShortName}");
            cell.CellStyle = Calibri11CS;

            RowIndex++;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, _appovingUser);
            cell.CellStyle = Calibri11CS;
            RowIndex++;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Заказ");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, $"План. отгрузка: {DateTime.Now.AddDays(14).ToShortDateString()}");
            cell.CellStyle = Calibri11CS;
            RowIndex++;

            int displayIndex = 0;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), displayIndex, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), displayIndex + 1, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), displayIndex, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), displayIndex + 1, ClientName);
            cell.CellStyle = CalibriBold11CS;

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
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), displayIndex++, "м.кв.");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), displayIndex++, "Примечание");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            int CType = -1;
            int AllTotalAmount = 0;
            int TotalAmount = 0;
            decimal AllTotalSquare = 0;
            decimal TotalSquare = 0;

            if (NotCurvedAssemblyDT.Rows.Count > 0)
                CType = Convert.ToInt32(NotCurvedAssemblyDT.Rows[0]["ColorType"]);

            for (int x = 0; x < NotCurvedAssemblyDT.Rows.Count; x++)
            {
                if (NotCurvedAssemblyDT.Rows[x]["Count"] != DBNull.Value)
                {
                    AllTotalAmount += Convert.ToInt32(NotCurvedAssemblyDT.Rows[x]["Count"]);
                    TotalAmount += Convert.ToInt32(NotCurvedAssemblyDT.Rows[x]["Count"]);
                }
                if (NotCurvedAssemblyDT.Rows[x]["Square"] != DBNull.Value)
                {
                    AllTotalSquare += Convert.ToDecimal(NotCurvedAssemblyDT.Rows[x]["Square"]);
                    TotalSquare += Convert.ToDecimal(NotCurvedAssemblyDT.Rows[x]["Square"]);
                }

                for (int y = 0; y < NotCurvedAssemblyDT.Columns.Count; y++)
                {
                    if (NotCurvedAssemblyDT.Columns[y].ColumnName == "ColorType")
                        continue;

                    Type t = NotCurvedAssemblyDT.Rows[x][y].GetType();

                    if (t.Name == "Decimal")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToDouble(NotCurvedAssemblyDT.Rows[x][y]));
                        cell.CellStyle = TableHeaderDecCS;
                        continue;
                    }
                    if (t.Name == "Int32")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToInt32(NotCurvedAssemblyDT.Rows[x][y]));
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (t.Name == "String" || t.Name == "DBNull")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(NotCurvedAssemblyDT.Rows[x][y].ToString());
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }
                }

                if (x + 1 <= NotCurvedAssemblyDT.Rows.Count - 1)
                {
                    if (CType != Convert.ToInt32(NotCurvedAssemblyDT.Rows[x + 1]["ColorType"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < NotCurvedAssemblyDT.Columns.Count; y++)
                        {
                            if (NotCurvedAssemblyDT.Columns[y].ColumnName == "ColorType")
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

                        cell = sheet1.CreateRow(RowIndex).CreateCell(6);
                        cell.SetCellValue(Convert.ToDouble(TotalSquare));
                        cell.CellStyle = TableHeaderCS;

                        CType = Convert.ToInt32(NotCurvedAssemblyDT.Rows[x + 1]["ColorType"]);
                        TotalAmount = 0;
                        TotalSquare = 0;
                        RowIndex++;
                    }
                }

                if (x == NotCurvedAssemblyDT.Rows.Count - 1)
                {
                    RowIndex++;
                    for (int y = 0; y < NotCurvedAssemblyDT.Columns.Count; y++)
                    {
                        if (NotCurvedAssemblyDT.Columns[y].ColumnName == "ColorType")
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

                    cell = sheet1.CreateRow(RowIndex).CreateCell(6);
                    cell.SetCellValue(Convert.ToDouble(TotalSquare));
                    cell.CellStyle = TableHeaderDecCS;

                    RowIndex++;

                    for (int y = 0; y < NotCurvedAssemblyDT.Columns.Count; y++)
                    {
                        if (NotCurvedAssemblyDT.Columns[y].ColumnName == "ColorType")
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

        private void AssemblyRover(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            ref HSSFSheet sheet1, DataTable DT, string BatchName, string ClientName, PlanningTimeFund timeFund, ref int RowIndex)
        {
            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, $"Распечатал: Дата/время { CurrentDate.ToString("dd.MM.yyyy HH:mm") }" +
                $" \r\n ФИО: { Security.CurrentUserShortName}");
            cell.CellStyle = Calibri11CS;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Плановое время выполнения:");
            cell.CellStyle = Calibri11CS;

            cell = sheet1.CreateRow(RowIndex).CreateCell(3);
            cell.SetCellValue(Convert.ToDouble(timeFund.time));
            cell.CellStyle = Calibri11CS;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "ч/ч");
            cell.CellStyle = Calibri11CS;
            RowIndex++;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Планово-премиальный фонд:");
            cell.CellStyle = Calibri11CS;

            cell = sheet1.CreateRow(RowIndex).CreateCell(3);
            cell.SetCellValue(Convert.ToDouble(timeFund.fund));
            cell.CellStyle = Calibri11CS;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "руб.");
            cell.CellStyle = Calibri11CS;

            RowIndex++;
            RowIndex++;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, _appovingUser);
            cell.CellStyle = Calibri11CS;
            RowIndex++;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Ровер");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, $"План. отгрузка: {DateTime.Now.AddDays(14).ToShortDateString()}");
            cell.CellStyle = Calibri11CS;
            RowIndex++;

            int displayIndex = 0;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), displayIndex, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), displayIndex + 1, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), displayIndex, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), displayIndex + 1, ClientName);
            cell.CellStyle = CalibriBold11CS;

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
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), displayIndex++, "м.кв.");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), displayIndex++, "Примечание");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            int CType = -1;
            int AllTotalAmount = 0;
            int TotalAmount = 0;
            decimal AllTotalSquare = 0;
            decimal TotalSquare = 0;

            if (NotCurvedAssemblyDT.Rows.Count > 0)
                CType = Convert.ToInt32(NotCurvedAssemblyDT.Rows[0]["ColorType"]);

            for (int x = 0; x < NotCurvedAssemblyDT.Rows.Count; x++)
            {
                if (NotCurvedAssemblyDT.Rows[x]["Count"] != DBNull.Value)
                {
                    AllTotalAmount += Convert.ToInt32(NotCurvedAssemblyDT.Rows[x]["Count"]);
                    TotalAmount += Convert.ToInt32(NotCurvedAssemblyDT.Rows[x]["Count"]);
                }
                if (NotCurvedAssemblyDT.Rows[x]["Square"] != DBNull.Value)
                {
                    AllTotalSquare += Convert.ToDecimal(NotCurvedAssemblyDT.Rows[x]["Square"]);
                    TotalSquare += Convert.ToDecimal(NotCurvedAssemblyDT.Rows[x]["Square"]);
                }

                for (int y = 0; y < NotCurvedAssemblyDT.Columns.Count; y++)
                {
                    if (NotCurvedAssemblyDT.Columns[y].ColumnName == "ColorType")
                        continue;

                    Type t = NotCurvedAssemblyDT.Rows[x][y].GetType();

                    if (t.Name == "Decimal")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToDouble(NotCurvedAssemblyDT.Rows[x][y]));
                        cell.CellStyle = TableHeaderDecCS;
                        continue;
                    }
                    if (t.Name == "Int32")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToInt32(NotCurvedAssemblyDT.Rows[x][y]));
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (t.Name == "String" || t.Name == "DBNull")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(NotCurvedAssemblyDT.Rows[x][y].ToString());
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }
                }

                if (x + 1 <= NotCurvedAssemblyDT.Rows.Count - 1)
                {
                    if (CType != Convert.ToInt32(NotCurvedAssemblyDT.Rows[x + 1]["ColorType"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < NotCurvedAssemblyDT.Columns.Count; y++)
                        {
                            if (NotCurvedAssemblyDT.Columns[y].ColumnName == "ColorType")
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

                        cell = sheet1.CreateRow(RowIndex).CreateCell(6);
                        cell.SetCellValue(Convert.ToDouble(TotalSquare));
                        cell.CellStyle = TableHeaderCS;

                        CType = Convert.ToInt32(NotCurvedAssemblyDT.Rows[x + 1]["ColorType"]);
                        TotalAmount = 0;
                        TotalSquare = 0;
                        RowIndex++;
                    }
                }

                if (x == NotCurvedAssemblyDT.Rows.Count - 1)
                {
                    RowIndex++;
                    for (int y = 0; y < NotCurvedAssemblyDT.Columns.Count; y++)
                    {
                        if (NotCurvedAssemblyDT.Columns[y].ColumnName == "ColorType")
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

                    cell = sheet1.CreateRow(RowIndex).CreateCell(6);
                    cell.SetCellValue(Convert.ToDouble(TotalSquare));
                    cell.CellStyle = TableHeaderDecCS;

                    RowIndex++;

                    for (int y = 0; y < NotCurvedAssemblyDT.Columns.Count; y++)
                    {
                        if (NotCurvedAssemblyDT.Columns[y].ColumnName == "ColorType")
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

        private void AssemblySaw(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            ref HSSFSheet sheet1, DataTable DT, string BatchName, string ClientName, PlanningTimeFund timeFund, ref int RowIndex)
        {
            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, $"Распечатал: Дата/время { CurrentDate.ToString("dd.MM.yyyy HH:mm") }" +
                $" \r\n ФИО: { Security.CurrentUserShortName}");
            cell.CellStyle = Calibri11CS;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Плановое время выполнения:");
            cell.CellStyle = Calibri11CS;

            cell = sheet1.CreateRow(RowIndex).CreateCell(3);
            cell.SetCellValue(Convert.ToDouble(timeFund.time));
            cell.CellStyle = Calibri11CS;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "ч/ч");
            cell.CellStyle = Calibri11CS;
            RowIndex++;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Планово-премиальный фонд:");
            cell.CellStyle = Calibri11CS;

            cell = sheet1.CreateRow(RowIndex).CreateCell(3);
            cell.SetCellValue(Convert.ToDouble(timeFund.fund));
            cell.CellStyle = Calibri11CS;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "руб.");
            cell.CellStyle = Calibri11CS;

            RowIndex++;
            RowIndex++;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, _appovingUser);
            cell.CellStyle = Calibri11CS;
            RowIndex++;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Пила");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, $"План. отгрузка: {DateTime.Now.AddDays(14).ToShortDateString()}");
            cell.CellStyle = Calibri11CS;
            RowIndex++;

            int displayIndex = 0;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), displayIndex, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), displayIndex + 1, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), displayIndex, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), displayIndex + 1, ClientName);
            cell.CellStyle = CalibriBold11CS;

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
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), displayIndex++, "м.кв.");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), displayIndex++, "Примечание");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            int CType = -1;
            int AllTotalAmount = 0;
            int TotalAmount = 0;
            decimal AllTotalSquare = 0;
            decimal TotalSquare = 0;

            if (NotCurvedAssemblyDT.Rows.Count > 0)
                CType = Convert.ToInt32(NotCurvedAssemblyDT.Rows[0]["ColorType"]);

            for (int x = 0; x < NotCurvedAssemblyDT.Rows.Count; x++)
            {
                if (NotCurvedAssemblyDT.Rows[x]["Count"] != DBNull.Value)
                {
                    AllTotalAmount += Convert.ToInt32(NotCurvedAssemblyDT.Rows[x]["Count"]);
                    TotalAmount += Convert.ToInt32(NotCurvedAssemblyDT.Rows[x]["Count"]);
                }
                if (NotCurvedAssemblyDT.Rows[x]["Square"] != DBNull.Value)
                {
                    AllTotalSquare += Convert.ToDecimal(NotCurvedAssemblyDT.Rows[x]["Square"]);
                    TotalSquare += Convert.ToDecimal(NotCurvedAssemblyDT.Rows[x]["Square"]);
                }

                for (int y = 0; y < NotCurvedAssemblyDT.Columns.Count; y++)
                {
                    if (NotCurvedAssemblyDT.Columns[y].ColumnName == "ColorType")
                        continue;

                    Type t = NotCurvedAssemblyDT.Rows[x][y].GetType();

                    if (t.Name == "Decimal")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToDouble(NotCurvedAssemblyDT.Rows[x][y]));
                        cell.CellStyle = TableHeaderDecCS;
                        continue;
                    }
                    if (t.Name == "Int32")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToInt32(NotCurvedAssemblyDT.Rows[x][y]));
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (t.Name == "String" || t.Name == "DBNull")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(NotCurvedAssemblyDT.Rows[x][y].ToString());
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }
                }

                if (x + 1 <= NotCurvedAssemblyDT.Rows.Count - 1)
                {
                    if (CType != Convert.ToInt32(NotCurvedAssemblyDT.Rows[x + 1]["ColorType"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < NotCurvedAssemblyDT.Columns.Count; y++)
                        {
                            if (NotCurvedAssemblyDT.Columns[y].ColumnName == "ColorType")
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

                        cell = sheet1.CreateRow(RowIndex).CreateCell(6);
                        cell.SetCellValue(Convert.ToDouble(TotalSquare));
                        cell.CellStyle = TableHeaderCS;

                        CType = Convert.ToInt32(NotCurvedAssemblyDT.Rows[x + 1]["ColorType"]);
                        TotalAmount = 0;
                        TotalSquare = 0;
                        RowIndex++;
                    }
                }

                if (x == NotCurvedAssemblyDT.Rows.Count - 1)
                {
                    RowIndex++;
                    for (int y = 0; y < NotCurvedAssemblyDT.Columns.Count; y++)
                    {
                        if (NotCurvedAssemblyDT.Columns[y].ColumnName == "ColorType")
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

                    cell = sheet1.CreateRow(RowIndex).CreateCell(6);
                    cell.SetCellValue(Convert.ToDouble(TotalSquare));
                    cell.CellStyle = TableHeaderDecCS;

                    RowIndex++;

                    for (int y = 0; y < NotCurvedAssemblyDT.Columns.Count; y++)
                    {
                        if (NotCurvedAssemblyDT.Columns[y].ColumnName == "ColorType")
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

        public void ClearOrders()
        {
            NotCurvedOrdersDT.Clear();
            //CurvedOrdersDT.Clear();
            NotArchDecorOrdersDT.Clear();
            ArchDecorOrdersDT.Clear();
            GridsDecorOrdersDT.Clear();
        }

        public void CreateExcel(string ClientName, string BatchName, ref string sSourceFileName)
        {
            GetCurrentDate();
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

            HSSFCellStyle WorkerColumnCS = hssfworkbook.CreateCellStyle();
            WorkerColumnCS.Alignment = HSSFCellStyle.ALIGN_RIGHT;
            WorkerColumnCS.BorderBottom = HSSFCellStyle.BORDER_THIN;
            WorkerColumnCS.BottomBorderColor = HSSFColor.BLACK.index;
            WorkerColumnCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            WorkerColumnCS.LeftBorderColor = HSSFColor.BLACK.index;
            WorkerColumnCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            WorkerColumnCS.RightBorderColor = HSSFColor.BLACK.index;
            WorkerColumnCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            WorkerColumnCS.TopBorderColor = HSSFColor.BLACK.index;
            WorkerColumnCS.SetFont(Serif10F);

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

            NotArchDecorAssemblyByMainOrderToExcel(ref hssfworkbook, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, BatchName);

            ArchDecorAssemblyByMainOrderToExcel(ref hssfworkbook, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, BatchName);

            //CurvedAssemblyToExcel(ref hssfworkbook,
            //    Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, BatchName);

            GridsDecorAssemblyByMainOrderToExcel(ref hssfworkbook, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, BatchName);

            SimpleDT.Clear();
            GetSimpleFronts(NotCurvedOrdersDT, ref SimpleDT);

            //GridsDT.Clear();
            //GetGridFronts(NotCurvedOrdersDT, ref GridsDT);

            //AppliqueDT.Clear();
            //GetAppliqueFronts(NotCurvedOrdersDT, ref AppliqueDT);

            NotCurvedAssemblyToExcel(ref hssfworkbook,
                Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, BatchName, ClientName);

            //InsetDT.Clear();
            //InsetCollectGridFronts(ref InsetDT, false);
            //if (InsetDT.Rows.Count > 0)
            //    InsetToExcel(ref hssfworkbook,
            //            Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, BatchName, ClientName);

            DeyingToExcel(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, BatchName, ClientName, "Покраска");

            DeyingByMainOrderToExcel(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, BatchName);

            string FileName = WorkAssignmentID + " " + BatchName + "  Угол 45";
            string tempFolder = @"\\192.168.1.6\Public\_ДЕЙСТВУЮЩИЕ\ПРОИЗВОДСТВО\ТПС\инфиниум\";
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
                file = new FileInfo(tempFolder + @"\" + FileName + "(" + j++ + ").xls");

            FileStream NewFile = new FileStream(file.FullName, FileMode.Create);
            hssfworkbook.Write(NewFile);
            NewFile.Close();

            System.Diagnostics.Process.Start(file.FullName);

            //string sSourceFolder = System.Environment.GetEnvironmentVariable("TEMP");
            //string sFolderPath = "Общие файлы/Производство/Задания в работу";
            //string sDestFolder = Configs.DocumentsPath + sFolderPath;
            //sSourceFileName = GetFileName(sDestFolder, BatchName);

            //FileInfo file = new FileInfo(sSourceFolder + @"\" + sSourceFileName);
            //FileStream NewFile = new FileStream(file.FullName, FileMode.Create);
            //hssfworkbook.Write(NewFile);
            //NewFile.Close();
        }

        private void CurvedAssembly1ToExcelSingly(ref HSSFWorkbook hssfworkbook, ref HSSFSheet sheet1,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            DataTable DT, string BatchName, string ClientName, string OperationName, string PageName, ref int RowIndex)
        {
            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, PageName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, ClientName);
            cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Партия:");
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, BatchName);
            //cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 4, OperationName);
            cell.CellStyle = CalibriBold11CS;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Фасад");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Цвет профиля");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Тип наполнителя");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Цвет наполнителя");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 5, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            int ColorID = -1;
            int AllTotalAmount = 0;
            int TotalAmount = 0;
            int DifferentDecorCount = 0;

            using (DataView DV = new DataView(DT))
            {
                DifferentDecorCount = DV.ToTable(true, new string[] { "ColorID" }).Rows.Count;
            }

            if (DT.Rows.Count > 0)
            {
                ColorID = Convert.ToInt32(DT.Rows[0]["ColorID"]);
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
                    if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "ColorID" ||
                        DT.Columns[y].ColumnName == "InsetTypeID" || DT.Columns[y].ColumnName == "InsetColorID")
                        continue;

                    Type t = DT.Rows[x][y].GetType();

                    if (t.Name == "Decimal")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToDouble(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderDecCS;
                        continue;
                    }
                    if (t.Name == "Int32" || t.Name == "Int64")
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
                    if (ColorID != Convert.ToInt32(DT.Rows[x + 1]["ColorID"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "ColorID" ||
                                DT.Columns[y].ColumnName == "InsetTypeID" || DT.Columns[y].ColumnName == "InsetColorID")
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

                        ColorID = Convert.ToInt32(DT.Rows[x + 1]["ColorID"]);
                        TotalAmount = 0;
                    }
                }

                if (x == DT.Rows.Count - 1)
                {
                    if (DifferentDecorCount > 1)
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "ColorID" ||
                                DT.Columns[y].ColumnName == "InsetTypeID" || DT.Columns[y].ColumnName == "InsetColorID")
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
                    }
                    RowIndex++;

                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "ColorID" ||
                            DT.Columns[y].ColumnName == "InsetTypeID" || DT.Columns[y].ColumnName == "InsetColorID")
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

        private void CurvedAssembly2ToExcelSingly(ref HSSFWorkbook hssfworkbook, ref HSSFSheet sheet1,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            DataTable DT, string BatchName, string ClientName, string OrderName, string PageName, string Notes, ref int RowIndex)
        {
            HSSFCell cell = null;

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
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Заказ:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, OrderName);
            cell.CellStyle = CalibriBold11CS;
            if (Notes.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Примечание: " + Notes);
                cell.CellStyle = CalibriBold11CS;
            }

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Фасад");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Цвет профиля");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Тип наполнителя");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Цвет наполнителя");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 5, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            int ColorID = -1;
            int AllTotalAmount = 0;
            int TotalAmount = 0;
            int DifferentDecorCount = 0;

            using (DataView DV = new DataView(DT))
            {
                DifferentDecorCount = DV.ToTable(true, new string[] { "ColorID" }).Rows.Count;
            }

            if (DT.Rows.Count > 0)
            {
                ColorID = Convert.ToInt32(DT.Rows[0]["ColorID"]);
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
                    if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "ColorID" ||
                        DT.Columns[y].ColumnName == "InsetTypeID" || DT.Columns[y].ColumnName == "InsetColorID")
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
                    if (ColorID != Convert.ToInt32(DT.Rows[x + 1]["ColorID"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "ColorID" ||
                                DT.Columns[y].ColumnName == "InsetTypeID" || DT.Columns[y].ColumnName == "InsetColorID")
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

                        ColorID = Convert.ToInt32(DT.Rows[x + 1]["ColorID"]);
                        TotalAmount = 0;
                    }
                }

                if (x == DT.Rows.Count - 1)
                {
                    if (DifferentDecorCount > 1)
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "ColorID" ||
                                DT.Columns[y].ColumnName == "InsetTypeID" || DT.Columns[y].ColumnName == "InsetColorID")
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
                    }
                    RowIndex++;

                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "ColorID" ||
                            DT.Columns[y].ColumnName == "InsetTypeID" || DT.Columns[y].ColumnName == "InsetColorID")
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

        private void DeyingByMainOrderToExcelSingly(ref HSSFWorkbook hssfworkbook, ref HSSFSheet sheet1,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            DataTable DT, string BatchName, string ClientName, string OrderName, string Notes, ref int RowIndex)
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

                decimal div1 = 2.5m;
                decimal div2 = 2.14m;

                PlanningTimeFund timeFund = CalcTimeFund(TempDT, div1, div2);

                DyeingPackingToExcel(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, TempDT, BatchName, ClientName, OrderName,
                    "Упаковка", Notes, timeFund, ref RowIndex);
                RowIndex++;
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

                div1 = 4.5m;
                div2 = 2.14m;

                timeFund = CalcTimeFund(TempDT, div1, div2);

                DyeingBoringToExcel(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, TempDT, BatchName, ClientName, OrderName,
                    "Сверление", Notes, timeFund, ref RowIndex);
                RowIndex++;
                RowIndex++;
            }

            RowIndex++;
        }

        private void DeyingToExcelSingly(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            string BatchName, string ClientName, string PageName)
        {
            int RowIndex = 0;

            HSSFSheet sheet1 = hssfworkbook.CreateSheet(PageName);
            sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

            sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
            sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
            sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
            sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

            sheet1.SetColumnWidth(0, 20 * 256);
            sheet1.SetColumnWidth(1, 20 * 256);
            sheet1.SetColumnWidth(2, 20 * 256);
            sheet1.SetColumnWidth(3, 6 * 256);
            sheet1.SetColumnWidth(4, 6 * 256);
            sheet1.SetColumnWidth(5, 6 * 256);

            DataTable DT = new DataTable();
            DataColumn Col1 = new DataColumn("Col1", System.Type.GetType("System.String"));
            DataColumn Col2 = new DataColumn("Col2", System.Type.GetType("System.String"));
            DataColumn Col3 = new DataColumn("Col3", System.Type.GetType("System.String"));

            if (DeyingDT.Rows.Count > 0)
            {
                DT.Dispose();
                Col1.Dispose();
                Col2.Dispose();
                Col3.Dispose();
                DT = DeyingDT.Copy();
                Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
                Col1.SetOrdinal(6);
                DT.Columns["Square"].SetOrdinal(7);
                DyeingWomen1ToExcel(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT, BatchName, ClientName, string.Empty,
                    "Шлифование торцев. и пласти х3", string.Empty, ref RowIndex);
                RowIndex++;
                RowIndex++;

                DT.Dispose();
                Col1.Dispose();
                Col2.Dispose();
                Col3.Dispose();
                DT = DeyingDT.Copy();
                Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
                Col1.SetOrdinal(6);
                DT.Columns["Square"].SetOrdinal(7);
                DyeingWomen2ToExcel(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT, BatchName, ClientName, string.Empty,
                    "Нанесение грунта на торцы и пласти х2", string.Empty, ref RowIndex);
                RowIndex++;
                RowIndex++;

                DT.Dispose();
                Col1.Dispose();
                Col2.Dispose();
                Col3.Dispose();
                DT = DeyingDT.Copy();
                Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
                Col2 = DT.Columns.Add("Col2", System.Type.GetType("System.String"));
                Col3 = DT.Columns.Add("Col3", System.Type.GetType("System.String"));
                Col1.SetOrdinal(6);
                Col2.SetOrdinal(7);
                Col3.SetOrdinal(8);
                DT.Columns["Square"].SetOrdinal(9);
                DyeingMen1ToExcel(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT, BatchName, ClientName, string.Empty,
                    "нанесение эмали", string.Empty, ref RowIndex);
                RowIndex++;
                RowIndex++;

                DT.Dispose();
                Col1.Dispose();
                Col2.Dispose();
                Col3.Dispose();
                DT = DeyingDT.Copy();
                Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
                Col1.SetOrdinal(6);
                DT.Columns["Square"].SetOrdinal(7);
                DyeingWomen3ToExcel(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT, BatchName, ClientName, string.Empty,
                    "Нанесение глянцевого лака", string.Empty, ref RowIndex);
                RowIndex++;
                RowIndex++;

                DT.Dispose();
                Col1.Dispose();
                Col2.Dispose();
                Col3.Dispose();
                DT = DeyingDT.Copy();
                Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
                Col1.SetOrdinal(6);
                DT.Columns["Square"].SetOrdinal(7);
                DyeingMen2ToExcel(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT, BatchName, ClientName, string.Empty,
                    "Шлифовка и полировка", string.Empty, ref RowIndex);
                RowIndex++;
                RowIndex++;

                DT.Dispose();
                Col1.Dispose();
                Col2.Dispose();
                Col3.Dispose();
                DT = DeyingDT.Copy();
                Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
                Col2 = DT.Columns.Add("Col2", System.Type.GetType("System.String"));
                Col1.SetOrdinal(6);
                Col2.SetOrdinal(7);
                DT.Columns["Square"].SetOrdinal(8);
                DyeingPackingToExcel(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT, BatchName, ClientName, string.Empty,
                    "Упаковка", string.Empty, ref RowIndex);
                RowIndex++;
                RowIndex++;

                //DT.Dispose();
                //Col1.Dispose();
                //Col2.Dispose();
                //Col3.Dispose();
                //DT = DeyingDT.Copy();
                //Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
                //Col1.SetOrdinal(6);
                //DT.Columns["Square"].SetOrdinal(7);
                //DT.Columns["Notes"].SetOrdinal(8);
                //DyeingBoringToExcel(ref hssfworkbook,
                //        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT, ClientName, BatchName,
                //    "Сверление. (" + Security.CurrentUserShortName + " от " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + ")", ref RowIndex);
            }

            RowIndex++;
        }

        private void DyeingBoringToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            ref HSSFSheet sheet1, DataTable DT, string BatchName, string ClientName, string OrderName, string PageName, string Notes, PlanningTimeFund timeFund,
            ref int RowIndex)
        {
            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, $"Распечатал: Дата/время { CurrentDate.ToString("dd.MM.yyyy HH:mm") }" +
                $" \r\n ФИО: { Security.CurrentUserShortName}");
            cell.CellStyle = Calibri11CS;

            if (OrderName.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Плановое время выполнения:");
                cell.CellStyle = Calibri11CS;

                cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                cell.SetCellValue(Convert.ToDouble(timeFund.time));
                cell.CellStyle = Calibri11CS;

                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "ч/ч");
                cell.CellStyle = Calibri11CS;
                RowIndex++;

                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Планово-премиальный фонд:");
                cell.CellStyle = Calibri11CS;

                cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                cell.SetCellValue(Convert.ToDouble(timeFund.fund));
                cell.CellStyle = Calibri11CS;

                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "руб.");
                cell.CellStyle = Calibri11CS;

                RowIndex++;
            }
            RowIndex++;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, _appovingUser);
            cell.CellStyle = Calibri11CS;
            RowIndex++;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, PageName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, $"План. отгрузка: {DateTime.Now.AddDays(14).ToShortDateString()}");
            cell.CellStyle = Calibri11CS;
            RowIndex++;

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

        private void DyeingMen1ToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            ref HSSFSheet sheet1, DataTable DT, string BatchName, string ClientName, string OrderName, string PageName, string Notes,
            ref int RowIndex)
        {
            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, $"Распечатал: Дата/время { CurrentDate.ToString("dd.MM.yyyy HH:mm") }" +
                $" \r\n ФИО: { Security.CurrentUserShortName}");
            cell.CellStyle = Calibri11CS;

            RowIndex++;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, _appovingUser);
            cell.CellStyle = Calibri11CS;
            RowIndex++;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, PageName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, $"План. отгрузка: {DateTime.Now.AddDays(14).ToShortDateString()}");
            cell.CellStyle = Calibri11CS;
            RowIndex++;

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
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 6, "Гр.в.");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 7, "Гр.н.");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 8, "Патина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 9, "м.кв.");
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

                        cell = sheet1.CreateRow(RowIndex).CreateCell(9);
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

                        cell = sheet1.CreateRow(RowIndex).CreateCell(7);
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

                    cell = sheet1.CreateRow(RowIndex).CreateCell(9);
                    cell.SetCellValue(Convert.ToDouble(AllTotalSquare));
                    cell.CellStyle = TableHeaderDecCS;
                }
                RowIndex++;
            }
        }

        private void DyeingMen2ToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            ref HSSFSheet sheet1, DataTable DT, string BatchName, string ClientName, string OrderName, string PageName, string Notes,
            ref int RowIndex)
        {
            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, $"Распечатал: Дата/время { CurrentDate.ToString("dd.MM.yyyy HH:mm") }" +
                $" \r\n ФИО: { Security.CurrentUserShortName}");
            cell.CellStyle = Calibri11CS;

            RowIndex++;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, _appovingUser);
            cell.CellStyle = Calibri11CS;
            RowIndex++;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, PageName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, $"План. отгрузка: {DateTime.Now.AddDays(14).ToShortDateString()}");
            cell.CellStyle = Calibri11CS;
            RowIndex++;

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
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 6, "Лак");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 7, "м.кв.");
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

                        cell = sheet1.CreateRow(RowIndex).CreateCell(7);
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

                    cell = sheet1.CreateRow(RowIndex).CreateCell(7);
                    cell.SetCellValue(Convert.ToDouble(AllTotalSquare));
                    cell.CellStyle = TableHeaderDecCS;
                }
                RowIndex++;
            }
        }

        private void DyeingPackingToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            ref HSSFSheet sheet1, DataTable DT, string BatchName, string ClientName, string OrderName, string PageName, string Notes, 
            ref int RowIndex)
        {
            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, $"Распечатал: Дата/время { CurrentDate.ToString("dd.MM.yyyy HH:mm") }" +
                $" \r\n ФИО: { Security.CurrentUserShortName}");
            cell.CellStyle = Calibri11CS;
            
            RowIndex++;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, _appovingUser);
            cell.CellStyle = Calibri11CS;
            RowIndex++;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, PageName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, $"План. отгрузка: {DateTime.Now.AddDays(14).ToShortDateString()}");
            cell.CellStyle = Calibri11CS;
            RowIndex++;

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
        
        private void DyeingPackingToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            ref HSSFSheet sheet1, DataTable DT, string BatchName, string ClientName, string OrderName, string PageName, string Notes, PlanningTimeFund timeFund,
            ref int RowIndex)
        {
            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, $"Распечатал: Дата/время { CurrentDate.ToString("dd.MM.yyyy HH:mm") }" +
                $" \r\n ФИО: { Security.CurrentUserShortName}");
            cell.CellStyle = Calibri11CS;

            if (OrderName.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Плановое время выполнения:");
                cell.CellStyle = Calibri11CS;

                cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                cell.SetCellValue(Convert.ToDouble(timeFund.time));
                cell.CellStyle = Calibri11CS;

                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "ч/ч");
                cell.CellStyle = Calibri11CS;
                RowIndex++;

                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Планово-премиальный фонд:");
                cell.CellStyle = Calibri11CS;

                cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                cell.SetCellValue(Convert.ToDouble(timeFund.fund));
                cell.CellStyle = Calibri11CS;

                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "руб.");
                cell.CellStyle = Calibri11CS;

                RowIndex++;
            }

            RowIndex++;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, _appovingUser);
            cell.CellStyle = Calibri11CS;
            RowIndex++;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, PageName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, $"План. отгрузка: {DateTime.Now.AddDays(14).ToShortDateString()}");
            cell.CellStyle = Calibri11CS;
            RowIndex++;

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

        private void DyeingWomen1ToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            ref HSSFSheet sheet1, DataTable DT, string BatchName, string ClientName, string OrderName, string PageName, string Notes,
            ref int RowIndex)
        {
            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, $"Распечатал: Дата/время { CurrentDate.ToString("dd.MM.yyyy HH:mm") }" +
                $" \r\n ФИО: { Security.CurrentUserShortName}");
            cell.CellStyle = Calibri11CS;

            RowIndex++;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, _appovingUser);
            cell.CellStyle = Calibri11CS;
            RowIndex++;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, PageName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, $"План. отгрузка: {DateTime.Now.AddDays(14).ToShortDateString()}");
            cell.CellStyle = Calibri11CS;
            RowIndex++;

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
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 6, "Зачистка");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 7, "м.кв.");
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

                        cell = sheet1.CreateRow(RowIndex).CreateCell(7);
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

                    cell = sheet1.CreateRow(RowIndex).CreateCell(7);
                    cell.SetCellValue(Convert.ToDouble(AllTotalSquare));
                    cell.CellStyle = TableHeaderDecCS;
                }
                RowIndex++;
            }
        }

        private void DyeingWomen2ToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            ref HSSFSheet sheet1, DataTable DT, string BatchName, string ClientName, string OrderName, string PageName, string Notes,
            ref int RowIndex)
        {
            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, $"Распечатал: Дата/время { CurrentDate.ToString("dd.MM.yyyy HH:mm") }" +
                $" \r\n ФИО: { Security.CurrentUserShortName}");
            cell.CellStyle = Calibri11CS;

            RowIndex++;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, _appovingUser);
            cell.CellStyle = Calibri11CS;
            RowIndex++;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, PageName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, $"План. отгрузка: {DateTime.Now.AddDays(14).ToShortDateString()}");
            cell.CellStyle = Calibri11CS;
            RowIndex++;

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
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 6, "Обезжиривание");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 7, "м.кв.");
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

                        cell = sheet1.CreateRow(RowIndex).CreateCell(7);
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

                    cell = sheet1.CreateRow(RowIndex).CreateCell(7);
                    cell.SetCellValue(Convert.ToDouble(AllTotalSquare));
                    cell.CellStyle = TableHeaderDecCS;
                }
                RowIndex++;
            }
        }

        private void DyeingWomen3ToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            ref HSSFSheet sheet1, DataTable DT, string BatchName, string ClientName, string OrderName, string PageName, string Notes,
            ref int RowIndex)
        {
            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, $"Распечатал: Дата/время { CurrentDate.ToString("dd.MM.yyyy HH:mm") }" +
                $" \r\n ФИО: { Security.CurrentUserShortName}");
            cell.CellStyle = Calibri11CS;

            RowIndex++;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, _appovingUser);
            cell.CellStyle = Calibri11CS;
            RowIndex++;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, PageName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, $"План. отгрузка: {DateTime.Now.AddDays(14).ToShortDateString()}");
            cell.CellStyle = Calibri11CS;
            RowIndex++;

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
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 6, "Протирка патины");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 7, "м.кв.");
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

                        cell = sheet1.CreateRow(RowIndex).CreateCell(7);
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

                    cell = sheet1.CreateRow(RowIndex).CreateCell(7);
                    cell.SetCellValue(Convert.ToDouble(AllTotalSquare));
                    cell.CellStyle = TableHeaderDecCS;
                }
                RowIndex++;
            }
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

        public bool GetOrders(int FactoryID)
        {
            GetNotCurvedFrontsOrders(ref NotCurvedOrdersDT, FactoryID);
            //GetCurvedFrontsOrders(ref CurvedOrdersDT, FactoryID);
            GetNotArchDecorOrders(ref NotArchDecorOrdersDT, FactoryID);
            GetArchDecorOrders(ref ArchDecorOrdersDT, FactoryID);
            GetGridsDecorOrders(ref GridsDecorOrdersDT, FactoryID);

            //ProfileNamesDT.Clear();
            //GetProfileNames(ref ProfileNamesDT, FactoryID, Fronts.Geneva);

            //InsetTypeNamesDT.Clear();
            //GetInsetTypeNames(ref InsetTypeNamesDT, FactoryID, Fronts.Geneva);

            if (NotCurvedOrdersDT.Rows.Count == 0 && NotArchDecorOrdersDT.Rows.Count == 0 && ArchDecorOrdersDT.Rows.Count == 0 && GridsDecorOrdersDT.Rows.Count == 0)
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
            DataTable DT, string BatchName, string ClientName, string OperationName, string OrderName, string Notes, ref int RowIndex)
        {
            HSSFCell cell = null;

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
            string BatchName, string ClientName)
        {
            int RowIndex = 0;

            HSSFSheet sheet1 = hssfworkbook.CreateSheet("Решетки2");
            sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

            sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
            sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
            sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
            sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

            sheet1.SetColumnWidth(0, 35 * 256);
            sheet1.SetColumnWidth(1, 11 * 256);
            sheet1.SetColumnWidth(2, 11 * 256);
            sheet1.SetColumnWidth(3, 7 * 256);
            sheet1.SetColumnWidth(4, 23 * 256);

            InsetToExcelSingly(ref hssfworkbook, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS,
                    ref sheet1, InsetDT, BatchName, ClientName, "Сборка решеток Женева", ref RowIndex);

            RowIndex++;
            RowIndex++;

            InsetToExcelSingly(ref hssfworkbook, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS,
                    ref sheet1, InsetDT, BatchName, ClientName, "Пила DFTP-400", ref RowIndex);
        }

        public void InsetToExcelSingly(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            ref HSSFSheet sheet1, DataTable DT,
            string BatchName, string ClientName, string PageName, ref int RowIndex)
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

            int CType = -1;
            int AllTotalAmount = 0;
            int TotalAmount = 0;

            if (DT.Rows.Count > 0)
                CType = Convert.ToInt32(DT.Rows[0]["ColorType"]);

            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value)
                {
                    AllTotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    TotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
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

                        cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        CType = Convert.ToInt32(DT.Rows[x + 1]["ColorType"]);
                        TotalAmount = 0;
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

                    cell = sheet1.CreateRow(RowIndex).CreateCell(3);
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

                    cell = sheet1.CreateRow(RowIndex).CreateCell(3);
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

            //RowIndex++;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "задание начали:");
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //sheet1.AddMergedRegion(new NPOI.HSSF.Util.Region(RowIndex, 0, RowIndex, 1));
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "задание закончили:");
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //sheet1.AddMergedRegion(new NPOI.HSSF.Util.Region(RowIndex, 2, RowIndex, 4));
            //RowIndex++;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "работало человек:");
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //sheet1.AddMergedRegion(new NPOI.HSSF.Util.Region(RowIndex, 0, RowIndex, 1));
            //RowIndex++;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "№ станка:");
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //sheet1.AddMergedRegion(new NPOI.HSSF.Util.Region(RowIndex, 0, RowIndex, 1));
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "№ операции:");
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //sheet1.AddMergedRegion(new NPOI.HSSF.Util.Region(RowIndex, 2, RowIndex, 4));
        }

        public void NotArchDecorAssemblyToExcelSingly(ref HSSFWorkbook hssfworkbook, ref HSSFSheet sheet1,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            DataTable DT, string BatchName, string ClientName, string OperationName, string OrderName, string Notes, ref int RowIndex)
        {
            HSSFCell cell = null;

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

        public void NotCurvedAssemblyToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            string BatchName, string ClientName)
        {
            if (NotCurvedOrdersDT.Rows.Count == 0)
                return;

            int RowIndex = 0;
            HSSFSheet sheet1 = hssfworkbook.CreateSheet("Сборка");
            sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

            sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
            sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
            sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
            sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

            sheet1.SetColumnWidth(0, 20 * 256);
            sheet1.SetColumnWidth(1, 20 * 256);
            sheet1.SetColumnWidth(2, 20 * 256);
            sheet1.SetColumnWidth(3, 6 * 256);
            sheet1.SetColumnWidth(4, 6 * 256);
            sheet1.SetColumnWidth(5, 6 * 256);
            sheet1.SetColumnWidth(6, 6 * 256);
            sheet1.SetColumnWidth(7, 18 * 256);

            DataTable DistFrameColorsDT = DistFrameColorsTable(NotCurvedOrdersDT, true);
            NotCurvedAssemblyDT.Clear();
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                HeightLess180 = false;
                AssemblySimpleCollect(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), SimpleDT, ref NotCurvedAssemblyDT, 4);
                //AssemblyGridsCollect(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), GridsDT, ref NotCurvedAssemblyDT, 4);
                //AssemblyAppliqueCollect(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), AppliqueDT, ref NotCurvedAssemblyDT, 4);
            }

            decimal div1 = 8;
            decimal div2 = 2.14m;

            PlanningTimeFund timeFund = CalcTimeFund(NotCurvedAssemblyDT, div1, div2);

            if (NotCurvedAssemblyDT.Rows.Count > 0)
            {
                RowIndex++;
                AssemblySaw(ref hssfworkbook, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS,
                    ref sheet1, NotCurvedAssemblyDT, BatchName, ClientName, timeFund, ref RowIndex);
                RowIndex++;
            }

            NotCurvedAssemblyDT.Clear();
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                AssemblySimpleCollect(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), SimpleDT, ref NotCurvedAssemblyDT, 0);
                //AssemblyGridsCollect(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), GridsDT, ref NotCurvedAssemblyDT, 0);
                //AssemblyAppliqueCollect(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), AppliqueDT, ref NotCurvedAssemblyDT, 0);
            }
            if (NotCurvedAssemblyDT.Rows.Count > 0)
            {
                div1 = 2.2m;
                div2 = 2.14m;

                timeFund = CalcTimeFund(NotCurvedAssemblyDT, div1, div2);

                RowIndex++;
                AssemblyRover(ref hssfworkbook, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS,
                    ref sheet1, NotCurvedAssemblyDT, BatchName, ClientName, timeFund, ref RowIndex);
                RowIndex++;
                RowIndex++;
                
                AssemblyOrders(ref hssfworkbook, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS,
                     ref sheet1, NotCurvedAssemblyDT, BatchName, ClientName, ref RowIndex);
            }
        }

        private void ArchDecorAssemblyByMainOrderToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            string BatchName)
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
                             BatchName, ClientName, string.Empty, OrderName, Notes, ref RowIndex);
                        RowIndex++;
                        ArchDecorAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                            Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DecorAssemblyDT,
                            BatchName, ClientName, "ДУБЛЬ", OrderName, Notes, ref RowIndex);
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
                             BatchName, ClientName, string.Empty, OrderName, Notes, ref RowIndex);
                        RowIndex++;
                        ArchDecorAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                            Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DecorAssemblyDT,
                            BatchName, ClientName, "ДУБЛЬ", OrderName, Notes, ref RowIndex);
                    }
                    RowIndex++;
                }
            }
        }

        private void AssemblyAppliqueCollect(int ColorID, DataTable SourceDT, ref DataTable DestinationDT, int Addmission)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();

            using (DataView DV = new DataView(SourceDT))
            {
                DT1 = DV.ToTable(true, new string[] { "FrontID" });
            }
            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                int FrontID = Convert.ToInt32(DT1.Rows[i]["FrontID"]);
                string Name = GetFrontName(Convert.ToInt32(DT1.Rows[i]["FrontID"]));

                using (DataView DV = new DataView(SourceDT, "FrontID=" + Convert.ToInt32(DT1.Rows[i]["FrontID"]) + " AND ColorID=" + ColorID, "ColorID, InsetTypeID DESC, InsetColorID, Height, Width", DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "ColorID", "PatinaID", "InsetTypeID", "InsetColorID", "Height", "Width" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    decimal Square = 0;
                    int Count = 0;
                    DataRow[] rows = SourceDT.Select("FrontID=" + Convert.ToInt32(DT1.Rows[i]["FrontID"]) +
                        " AND ColorID=" + Convert.ToInt32(DT2.Rows[j]["ColorID"]) + " AND PatinaID=" + Convert.ToInt32(DT2.Rows[j]["PatinaID"]) + " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) +
                        " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                        " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]));
                    if (rows.Count() == 0)
                        continue;

                    foreach (DataRow item in rows)
                    {
                        Count += Convert.ToInt32(item["Count"]);
                        Square += (Addmission + Convert.ToDecimal(item["Height"])) * (Addmission + Convert.ToDecimal(item["Width"])) * Convert.ToDecimal(item["Count"]) / 1000000;
                    }
                    Square = decimal.Round(Square, 3, MidpointRounding.AwayFromZero);

                    DataRow NewRow = DestinationDT.NewRow();
                    NewRow["ColorType"] = ColorID;
                    NewRow["Name"] = Name;
                    if (Convert.ToInt32(rows[0]["PatinaID"]) == -1)
                        NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"]));
                    else
                        NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"])) + "+" + GetPatinaName(Convert.ToInt32(rows[0]["PatinaID"]));
                    NewRow["InsetColor"] = "-";
                    NewRow["Height"] = Convert.ToInt32(rows[0]["Height"]) + Addmission;
                    NewRow["Width"] = Convert.ToInt32(rows[0]["Width"]) + Addmission;
                    NewRow["Count"] = Count;
                    NewRow["Notes"] = rows[0]["Notes"].ToString();
                    NewRow["Square"] = Square;
                    DestinationDT.Rows.Add(NewRow);
                }
            }
        }

        private void AssemblyCurvedCollect(DataTable SourceDT, ref DataTable DestinationDT)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            DataTable DT3 = new DataTable();
            DataTable DT4 = new DataTable();

            using (DataView DV = new DataView(SourceDT, string.Empty, "ColorID, PatinaID", DataViewRowState.CurrentRows))
            {
                DT1 = DV.ToTable(true, new string[] { "ColorID", "PatinaID" });
            }
            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                using (DataView DV = new DataView(SourceDT, "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) + " AND PatinaID=" + Convert.ToInt32(DT1.Rows[i]["PatinaID"]), "InsetTypeID", DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "InsetTypeID" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    using (DataView DV = new DataView(SourceDT, "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) + " AND PatinaID=" + Convert.ToInt32(DT1.Rows[i]["PatinaID"]) +
                        " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]), "InsetColorID", DataViewRowState.CurrentRows))
                    {
                        DT3 = DV.ToTable(true, new string[] { "InsetColorID" });
                    }
                    for (int x = 0; x < DT3.Rows.Count; x++)
                    {
                        using (DataView DV = new DataView(SourceDT, "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) + " AND PatinaID=" + Convert.ToInt32(DT1.Rows[i]["PatinaID"]) +
                            " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) +
                            " AND InsetColorID=" + Convert.ToInt32(DT3.Rows[x]["InsetColorID"]), "Height", DataViewRowState.CurrentRows))
                        {
                            DT4 = DV.ToTable(true, new string[] { "Height", "Width" });
                        }
                        for (int y = 0; y < DT4.Rows.Count; y++)
                        {
                            DataRow[] Srows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) + " AND PatinaID=" + Convert.ToInt32(DT1.Rows[i]["PatinaID"]) +
                                " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) +
                                " AND InsetColorID=" + Convert.ToInt32(DT3.Rows[x]["InsetColorID"]) +
                                " AND Height=" + Convert.ToInt32(DT4.Rows[y]["Height"]));
                            if (Srows.Count() == 0)
                                continue;

                            int Count = 0;
                            foreach (DataRow item in Srows)
                                Count += Convert.ToInt32(item["Count"]);

                            DataRow[] rows = DestinationDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) + " AND PatinaID=" + Convert.ToInt32(DT1.Rows[i]["PatinaID"]) +
                                " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) +
                                " AND InsetColorID=" + Convert.ToInt32(DT3.Rows[x]["InsetColorID"]) +
                                " AND Height=" + Convert.ToInt32(DT4.Rows[y]["Height"]));
                            if (rows.Count() == 0)
                            {
                                DataRow NewRow = DestinationDT.NewRow();
                                NewRow["Front"] = GetFrontName(Convert.ToInt32(Srows[0]["FrontID"]));
                                if (Convert.ToInt32(DT1.Rows[i]["PatinaID"]) == -1)
                                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                                else
                                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"])) + "+" + GetPatinaName(Convert.ToInt32(DT1.Rows[i]["PatinaID"]));
                                NewRow["InsetType"] = GetInsetTypeName(Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]));
                                NewRow["InsetColor"] = GetColorName(Convert.ToInt32(DT3.Rows[x]["InsetColorID"]));
                                NewRow["Height"] = Convert.ToInt32(DT4.Rows[y]["Height"]);
                                NewRow["Count"] = Count;
                                NewRow["FrontID"] = Convert.ToInt32(Srows[0]["FrontID"]);
                                NewRow["ColorID"] = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                                NewRow["InsetTypeID"] = Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]);
                                NewRow["InsetColorID"] = Convert.ToInt32(DT3.Rows[x]["InsetColorID"]);
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
            DT1.Dispose();
            DT2.Dispose();
            DT3.Dispose();
            DT4.Dispose();
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

        private void AssemblyGridsCollect(int ColorID, DataTable SourceDT, ref DataTable DestinationDT, int Addmission)
        {
            DataTable DT2 = new DataTable();

            using (DataView DV = new DataView(SourceDT, "FrontID NOT IN (3729) AND ColorID=" + ColorID, "ColorID, InsetTypeID, InsetColorID, Height, Width", DataViewRowState.CurrentRows))
            {
                DT2 = DV.ToTable(true, new string[] { "ColorID", "PatinaID", "InsetTypeID", "InsetColorID", "Height", "Width" });
            }
            for (int j = 0; j < DT2.Rows.Count; j++)
            {
                decimal Square = 0;
                int Count = 0;
                //решетка пп 45
                DataRow[] rows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT2.Rows[j]["ColorID"]) + " AND PatinaID=" + Convert.ToInt32(DT2.Rows[j]["PatinaID"]) + " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) +
                    " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                    " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]));
                foreach (DataRow item in rows)
                {
                    Count += Convert.ToInt32(item["Count"]);
                    Square += (Addmission + Convert.ToDecimal(item["Height"])) * (Addmission + Convert.ToDecimal(item["Width"])) * Convert.ToDecimal(item["Count"]) / 1000000;
                }
                Square = decimal.Round(Square, 3, MidpointRounding.AwayFromZero);

                DataRow NewRow = DestinationDT.NewRow();
                NewRow["ColorType"] = ColorID;
                NewRow["Name"] = GetFrontName(Convert.ToInt32(rows[0]["FrontID"])) + " РЕШ";
                if (Convert.ToInt32(rows[0]["PatinaID"]) == -1)
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"]));
                else
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"])) + "+" + GetPatinaName(Convert.ToInt32(rows[0]["PatinaID"]));
                NewRow["InsetColor"] = GetInsetTypeName(Convert.ToInt32(rows[0]["InsetTypeID"])) + " " + GetInsetColorName(Convert.ToInt32(rows[0]["InsetColorID"]));
                NewRow["Height"] = Convert.ToInt32(rows[0]["Height"]) + Addmission;
                NewRow["Width"] = Convert.ToInt32(rows[0]["Width"]) + Addmission;
                NewRow["Count"] = Count;
                if (rows[0]["Notes"].ToString().Length == 0)
                    NewRow["Notes"] = "Силик уплотнитель";
                else
                    NewRow["Notes"] = rows[0]["Notes"].ToString() + ", Силик уплотнитель";
                NewRow["Square"] = Square;
                DestinationDT.Rows.Add(NewRow);
            }
            using (DataView DV = new DataView(SourceDT, "FrontID IN (3729) AND ColorID=" + ColorID, "ColorID, InsetTypeID, InsetColorID, Height, Width", DataViewRowState.CurrentRows))
            {
                DT2 = DV.ToTable(true, new string[] { "ColorID", "PatinaID", "InsetTypeID", "InsetColorID", "Height", "Width" });
            }
            for (int j = 0; j < DT2.Rows.Count; j++)
            {
                decimal Square = 0;
                int Count = 0;
                //решетка овал
                DataRow[] rows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT2.Rows[j]["ColorID"]) + " AND PatinaID=" + Convert.ToInt32(DT2.Rows[j]["PatinaID"]) + " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) +
                    " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                    " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]));
                foreach (DataRow item in rows)
                {
                    Count += Convert.ToInt32(item["Count"]);
                    Square += (Addmission + Convert.ToDecimal(item["Height"])) * (Addmission + Convert.ToDecimal(item["Width"])) * Convert.ToDecimal(item["Count"]) / 1000000;
                }
                Square = decimal.Round(Square, 3, MidpointRounding.AwayFromZero);

                DataRow NewRow = DestinationDT.NewRow();
                NewRow["ColorType"] = ColorID;
                NewRow["Name"] = GetFrontName(Convert.ToInt32(rows[0]["FrontID"]));
                if (Convert.ToInt32(rows[0]["PatinaID"]) == -1)
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"]));
                else
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"])) + "+" + GetPatinaName(Convert.ToInt32(rows[0]["PatinaID"]));
                NewRow["InsetColor"] = GetInsetTypeName(Convert.ToInt32(rows[0]["InsetTypeID"])) + " " + GetInsetColorName(Convert.ToInt32(rows[0]["InsetColorID"])) + " ОВАЛ";
                NewRow["Height"] = Convert.ToInt32(rows[0]["Height"]) + Addmission;
                NewRow["Width"] = Convert.ToInt32(rows[0]["Width"]) + Addmission;
                NewRow["Count"] = Count;
                if (rows[0]["Notes"].ToString().Length == 0)
                    NewRow["Notes"] = "Силик уплотнитель";
                else
                    NewRow["Notes"] = rows[0]["Notes"].ToString() + ", Силик уплотнитель";
                NewRow["Square"] = Square;
                DestinationDT.Rows.Add(NewRow);
            }
        }

        private void AssemblySimpleCollect(int ColorID, DataTable SourceDT, ref DataTable DestinationDT, int Addmission)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();

            using (DataView DV = new DataView(SourceDT))
            {
                DT1 = DV.ToTable(true, new string[] { "FrontID" });
            }
            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                string Name = GetFrontName(Convert.ToInt32(DT1.Rows[i]["FrontID"]));
                using (DataView DV = new DataView(SourceDT, "FrontID=" + Convert.ToInt32(DT1.Rows[i]["FrontID"]) + " AND InsetTypeID=1 AND ColorID=" + ColorID, "ColorID, InsetColorID, Height, Width", DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "ColorID", "PatinaID", "InsetTypeID", "InsetColorID", "Height", "Width", "Notes" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    {
                        decimal Square = 0;
                        int Count = 0;
                        //витрины не дуэт и не трио
                        string filter = "FrontID=" + Convert.ToInt32(DT1.Rows[i]["FrontID"]) +
                            " AND ColorID=" + Convert.ToInt32(DT2.Rows[j]["ColorID"]) + " AND PatinaID=" + Convert.ToInt32(DT2.Rows[j]["PatinaID"]) + " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) +
                            " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                            " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) + " AND (Notes='' OR Notes IS NULL)";
                        if (DT2.Rows[j]["Notes"] != DBNull.Value && DT2.Rows[j]["Notes"].ToString().Length > 0)
                            filter = "FrontID=" + Convert.ToInt32(DT1.Rows[i]["FrontID"]) +
                            " AND ColorID=" + Convert.ToInt32(DT2.Rows[j]["ColorID"]) + " AND PatinaID=" + Convert.ToInt32(DT2.Rows[j]["PatinaID"]) + " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) +
                            " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                            " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) + " AND Notes='" + DT2.Rows[j]["Notes"] + "'";

                        DataRow[] rows = SourceDT.Select(filter);
                        if (rows.Count() == 0)
                            continue;

                        foreach (DataRow item in rows)
                        {
                            Count += Convert.ToInt32(item["Count"]);
                            Square += (Addmission + Convert.ToDecimal(item["Height"])) * (Addmission + Convert.ToDecimal(item["Width"])) * Convert.ToDecimal(item["Count"]) / 1000000;
                        }
                        Square = decimal.Round(Square, 3, MidpointRounding.AwayFromZero);
                        DataRow NewRow = DestinationDT.NewRow();
                        NewRow["ColorType"] = ColorID;
                        NewRow["Name"] = Name;
                        if (Convert.ToInt32(rows[0]["PatinaID"]) == -1)
                            NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"]));
                        else
                            NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"])) + "+" + GetPatinaName(Convert.ToInt32(rows[0]["PatinaID"]));
                        NewRow["InsetColor"] = "Витрина";
                        NewRow["Height"] = Convert.ToInt32(rows[0]["Height"]) + Addmission;
                        NewRow["Width"] = Convert.ToInt32(rows[0]["Width"]) + Addmission;
                        NewRow["Count"] = Count;
                        NewRow["Notes"] = rows[0]["Notes"].ToString();
                        if (rows[0]["Notes"].ToString().Length == 0)
                            NewRow["Notes"] = "текстура не важна";
                        else
                            NewRow["Notes"] = rows[0]["Notes"].ToString() + ", текстура не важна";
                        NewRow["Square"] = Square;
                        DestinationDT.Rows.Add(NewRow);
                    }
                }

                using (DataView DV = new DataView(SourceDT, "FrontID=" + Convert.ToInt32(DT1.Rows[i]["FrontID"]) + " AND InsetTypeID=-1 AND ColorID=" + ColorID, "ColorID, InsetColorID, Height, Width", DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "ColorID", "PatinaID", "InsetTypeID", "InsetColorID", "Height", "Width", "Notes" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    {
                        decimal Square = 0;
                        int Count = 0;
                        //без наполнителя
                        string filter = "FrontID=" + Convert.ToInt32(DT1.Rows[i]["FrontID"]) +
                            " AND ColorID=" + Convert.ToInt32(DT2.Rows[j]["ColorID"]) + " AND PatinaID=" + Convert.ToInt32(DT2.Rows[j]["PatinaID"]) + " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) +
                            " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                            " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) + " AND (Notes='' OR Notes IS NULL)";
                        if (DT2.Rows[j]["Notes"] != DBNull.Value && DT2.Rows[j]["Notes"].ToString().Length > 0)
                            filter = "FrontID=" + Convert.ToInt32(DT1.Rows[i]["FrontID"]) +
                            " AND ColorID=" + Convert.ToInt32(DT2.Rows[j]["ColorID"]) + " AND PatinaID=" + Convert.ToInt32(DT2.Rows[j]["PatinaID"]) + " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) +
                            " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                            " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) + " AND Notes='" + DT2.Rows[j]["Notes"] + "'";

                        DataRow[] rows = SourceDT.Select(filter);
                        if (rows.Count() == 0)
                            continue;

                        foreach (DataRow item in rows)
                        {
                            Count += Convert.ToInt32(item["Count"]);
                            Square += (Addmission + Convert.ToDecimal(item["Height"])) * (Addmission + Convert.ToDecimal(item["Width"])) * Convert.ToDecimal(item["Count"]) / 1000000;
                        }
                        Square = decimal.Round(Square, 3, MidpointRounding.AwayFromZero);
                        DataRow NewRow = DestinationDT.NewRow();
                        NewRow["ColorType"] = ColorID;
                        NewRow["Name"] = Name;
                        if (Convert.ToInt32(rows[0]["PatinaID"]) == -1)
                            NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"]));
                        else
                            NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"])) + "+" + GetPatinaName(Convert.ToInt32(rows[0]["PatinaID"]));
                        NewRow["InsetColor"] = "-";

                        bool bHeightLess180 = false;
                        bool bWidthLess180 = false;

                        if (Convert.ToInt32(rows[0]["Height"]) <= 176)
                            bHeightLess180 = true;
                        if (Convert.ToInt32(rows[0]["Width"]) <= 176)
                            bWidthLess180 = true;

                        NewRow["Notes"] = rows[0]["Notes"].ToString();
                        NewRow["Height"] = Convert.ToInt32(rows[0]["Height"]) + Addmission;
                        NewRow["Width"] = Convert.ToInt32(rows[0]["Width"]) + Addmission;
                        NewRow["Count"] = Count;
                        if (bHeightLess180)
                        {
                            HeightLess180 = true;
                            if (rows[0]["Notes"].ToString().Length == 0)
                                NewRow["Notes"] = "текстура не важна";
                            else
                                NewRow["Notes"] = rows[0]["Notes"].ToString() + ", текстура не важна";
                        }
                        if (bWidthLess180)
                        {
                            if (!HeightLess180)
                            {
                                if (rows[0]["Notes"].ToString().Length == 0)
                                    NewRow["Notes"] = "текстура не важна";
                                else
                                    NewRow["Notes"] = rows[0]["Notes"].ToString() + ", текстура не важна";
                            }
                        }
                        if (!bHeightLess180 && !bWidthLess180)
                            NewRow["Notes"] = rows[0]["Notes"].ToString();

                        NewRow["Square"] = Square;
                        DestinationDT.Rows.Add(NewRow);
                    }
                }
            }
        }

        private PlanningTimeFund CalcTimeFund(DataTable dt, decimal div1, decimal div2)
        {
            decimal count = 0;
            PlanningTimeFund timeFund = new PlanningTimeFund();
            for (int x = 0; x < dt.Rows.Count; x++)
            {
                if (dt.Rows[x]["Square"] != DBNull.Value)
                    count += Convert.ToDecimal(dt.Rows[x]["Square"]);
            }
            timeFund.time = 0;
            if (div1 != 0)
                timeFund.time = decimal.Round(count / div1, 3, MidpointRounding.AwayFromZero);
            timeFund.fund = decimal.Round(timeFund.time * div2, 3, MidpointRounding.AwayFromZero);
            return timeFund;
        }
                
        private void CollectDeying(DataTable SourceDT, ref DataTable DestinationDT)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            DataTable DT3 = new DataTable();

            using (DataView DV = new DataView(SourceDT, string.Empty, "ColorID, PatinaID", DataViewRowState.CurrentRows))
            {
                DT1 = DV.ToTable(true, new string[] { "ColorID", "PatinaID" });
            }
            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                using (DataView DV = new DataView(SourceDT, "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) + " AND PatinaID=" + Convert.ToInt32(DT1.Rows[i]["PatinaID"]), "FrontID", DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "FrontID" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    string Name = GetFrontName(Convert.ToInt32(DT2.Rows[j]["FrontID"]));
                    //сначала витрины
                    int FrontID = Convert.ToInt32(DT2.Rows[j]["FrontID"]);
                    using (DataView DV = new DataView(SourceDT, "InsetTypeID=1 AND FrontID=" + Convert.ToInt32(DT2.Rows[j]["FrontID"]) +
                        " AND ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) + " AND PatinaID=" + Convert.ToInt32(DT1.Rows[i]["PatinaID"]),
                        "InsetTypeID, InsetColorID, Height, Width", DataViewRowState.CurrentRows))
                    {
                        DT3 = DV.ToTable(true, new string[] { "InsetTypeID", "InsetColorID", "Height", "Width", "Notes" });
                    }

                    for (int c = 0; c < DT3.Rows.Count; c++)
                    {
                        decimal Square = 0;
                        int Count = 0;
                        string filter = "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                            " AND PatinaID=" + Convert.ToInt32(DT1.Rows[i]["PatinaID"]) +
                            " AND FrontID=" + Convert.ToInt32(DT2.Rows[j]["FrontID"]) +
                            " AND InsetTypeID=" + Convert.ToInt32(DT3.Rows[c]["InsetTypeID"]) +
                            " AND InsetColorID=" + Convert.ToInt32(DT3.Rows[c]["InsetColorID"]) + " AND Height=" + Convert.ToInt32(DT3.Rows[c]["Height"]) +
                            " AND Width=" + Convert.ToInt32(DT3.Rows[c]["Width"]) + " AND (Notes='' OR Notes IS NULL)";
                        if (DT3.Rows[c]["Notes"] != DBNull.Value && DT3.Rows[c]["Notes"].ToString().Length > 0)
                            filter = "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                            " AND PatinaID=" + Convert.ToInt32(DT1.Rows[i]["PatinaID"]) +
                            " AND FrontID=" + Convert.ToInt32(DT2.Rows[j]["FrontID"]) +
                            " AND InsetTypeID=" + Convert.ToInt32(DT3.Rows[c]["InsetTypeID"]) +
                            " AND InsetColorID=" + Convert.ToInt32(DT3.Rows[c]["InsetColorID"]) + " AND Height=" + Convert.ToInt32(DT3.Rows[c]["Height"]) +
                            " AND Width=" + Convert.ToInt32(DT3.Rows[c]["Width"]) + " AND Notes='" + DT3.Rows[c]["Notes"] + "'";

                        DataRow[] rows = SourceDT.Select(filter);
                        if (rows.Count() == 0)
                            continue;
                        foreach (DataRow item in rows)
                        {
                            Count += Convert.ToInt32(item["Count"]);
                            Square += Convert.ToDecimal(item["Height"]) * Convert.ToDecimal(item["Width"]) * Convert.ToDecimal(item["Count"]) / 1000000;
                        }
                        Square = decimal.Round(Square, 3, MidpointRounding.AwayFromZero);

                        DataRow NewRow = DestinationDT.NewRow();
                        NewRow["ColorType"] = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                        if (Convert.ToInt32(DT1.Rows[i]["PatinaID"]) == -1)
                            NewRow["FrameColor"] = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                        else
                            NewRow["FrameColor"] = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"])) + "+" + GetPatinaName(Convert.ToInt32(DT1.Rows[i]["PatinaID"]));

                        NewRow["Name"] = Name;
                        NewRow["InsetColor"] = "Витрина";
                        NewRow["Height"] = rows[0]["Height"];
                        NewRow["Width"] = rows[0]["Width"];
                        NewRow["Count"] = Count;
                        NewRow["Square"] = Square;
                        NewRow["Notes"] = rows[0]["Notes"];
                        DestinationDT.Rows.Add(NewRow);
                    }

                    using (DataView DV = new DataView(SourceDT, "InsetTypeID<>1 AND FrontID=" + Convert.ToInt32(DT2.Rows[j]["FrontID"]) +
                        " AND ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) + " AND PatinaID=" + Convert.ToInt32(DT1.Rows[i]["PatinaID"]),
                        "InsetTypeID, InsetColorID, Height, Width", DataViewRowState.CurrentRows))
                    {
                        DT3 = DV.ToTable(true, new string[] { "InsetTypeID", "InsetColorID", "Height", "Width", "Notes" });
                    }
                    for (int c = 0; c < DT3.Rows.Count; c++)
                    {
                        decimal Square = 0;
                        int Count = 0;
                        string InsetType = string.Empty;
                        string InsetColor = string.Empty;

                        string filter = "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                            " AND PatinaID=" + Convert.ToInt32(DT1.Rows[i]["PatinaID"]) +
                            " AND FrontID=" + Convert.ToInt32(DT2.Rows[j]["FrontID"]) +
                            " AND InsetTypeID=" + Convert.ToInt32(DT3.Rows[c]["InsetTypeID"]) +
                            " AND InsetColorID=" + Convert.ToInt32(DT3.Rows[c]["InsetColorID"]) + " AND Height=" + Convert.ToInt32(DT3.Rows[c]["Height"]) +
                            " AND Width=" + Convert.ToInt32(DT3.Rows[c]["Width"]) + " AND (Notes='' OR Notes IS NULL)";
                        if (DT3.Rows[c]["Notes"] != DBNull.Value && DT3.Rows[c]["Notes"].ToString().Length > 0)
                            filter = "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                            " AND PatinaID=" + Convert.ToInt32(DT1.Rows[i]["PatinaID"]) +
                            " AND FrontID=" + Convert.ToInt32(DT2.Rows[j]["FrontID"]) +
                            " AND InsetTypeID=" + Convert.ToInt32(DT3.Rows[c]["InsetTypeID"]) +
                            " AND InsetColorID=" + Convert.ToInt32(DT3.Rows[c]["InsetColorID"]) + " AND Height=" + Convert.ToInt32(DT3.Rows[c]["Height"]) +
                            " AND Width=" + Convert.ToInt32(DT3.Rows[c]["Width"]) + " AND Notes='" + DT3.Rows[c]["Notes"] + "'";

                        DataRow[] rows = SourceDT.Select(filter);
                        //DataRow[] rows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        //    " AND PatinaID=" + Convert.ToInt32(DT1.Rows[i]["PatinaID"]) +
                        //    " AND FrontID=" + Convert.ToInt32(DT2.Rows[j]["FrontID"]) +
                        //    " AND InsetTypeID=" + Convert.ToInt32(DT3.Rows[c]["InsetTypeID"]) +
                        //    " AND InsetColorID=" + Convert.ToInt32(DT3.Rows[c]["InsetColorID"]) + " AND Height=" + Convert.ToInt32(DT3.Rows[c]["Height"]) +
                        //    " AND Width=" + Convert.ToInt32(DT3.Rows[c]["Width"]));
                        if (rows.Count() == 0)
                            continue;
                        foreach (DataRow item in rows)
                        {
                            Count += Convert.ToInt32(item["Count"]);
                            Square += Convert.ToDecimal(item["Height"]) * Convert.ToDecimal(item["Width"]) * Convert.ToDecimal(item["Count"]) / 1000000;
                        }
                        Square = decimal.Round(Square, 3, MidpointRounding.AwayFromZero);

                        DataRow NewRow = DestinationDT.NewRow();
                        NewRow["ColorType"] = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                        if (Convert.ToInt32(DT1.Rows[i]["PatinaID"]) == -1)
                            NewRow["FrameColor"] = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                        else
                            NewRow["FrameColor"] = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"])) + "+" + GetPatinaName(Convert.ToInt32(DT1.Rows[i]["PatinaID"]));
                        InsetType = GetInsetTypeName(Convert.ToInt32(rows[0]["InsetTypeID"]));
                        if (Convert.ToInt32(rows[0]["InsetTypeID"]) != -1)
                        {
                            if (Convert.ToInt32(rows[0]["InsetColorID"]) != -1)
                                InsetColor = " " + GetInsetColorName(Convert.ToInt32(rows[0]["InsetColorID"]));
                        }
                        else
                        {
                            InsetType = string.Empty;
                        }
                        if (Convert.ToInt32(DT3.Rows[c]["InsetTypeID"]) == 685 || Convert.ToInt32(DT3.Rows[c]["InsetTypeID"]) == 686 || Convert.ToInt32(DT3.Rows[c]["InsetTypeID"]) == 687 || Convert.ToInt32(DT3.Rows[c]["InsetTypeID"]) == 688 || Convert.ToInt32(DT3.Rows[c]["InsetTypeID"]) == 29470 || Convert.ToInt32(DT3.Rows[c]["InsetTypeID"]) == 29471)
                            Name = GetFrontName(Convert.ToInt32(rows[0]["FrontID"])) + " РЕШ";
                        NewRow["Name"] = Name;
                        NewRow["InsetColor"] = InsetType + InsetColor;
                        NewRow["Height"] = rows[0]["Height"];
                        NewRow["Width"] = rows[0]["Width"];
                        NewRow["Count"] = Count;
                        NewRow["Square"] = Square;
                        NewRow["Notes"] = rows[0]["Notes"];
                        DestinationDT.Rows.Add(NewRow);
                    }
                }
            }
        }

        private void Create()
        {
            FrontsDataTable = new DataTable();
            FrameColorsDataTable = new DataTable();
            PatinaDataTable = new DataTable();
            InsetTypesDataTable = new DataTable();
            InsetColorsDataTable = new DataTable();
            DecorDT = new DataTable();
            DecorParametersDT = new DataTable();            
            SimpleDT = new DataTable();
            NotCurvedOrdersDT = new DataTable();
            NotArchDecorOrdersDT = new DataTable();
            ArchDecorOrdersDT = new DataTable();
            GridsDecorOrdersDT = new DataTable();

            InsetDT = new DataTable();
            InsetDT.Columns.Add(new DataColumn("Name", Type.GetType("System.String")));
            InsetDT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            InsetDT.Columns.Add(new DataColumn("Width", Type.GetType("System.Int32")));
            InsetDT.Columns.Add(new DataColumn("Count", Type.GetType("System.Int32")));
            InsetDT.Columns.Add(new DataColumn("Worker", Type.GetType("System.String")));
            InsetDT.Columns.Add(new DataColumn("ColorType", Type.GetType("System.Int32")));

            NotCurvedAssemblyDT = new DataTable();
            NotCurvedAssemblyDT.Columns.Add(new DataColumn("Name", Type.GetType("System.String")));
            NotCurvedAssemblyDT.Columns.Add(new DataColumn("FrameColor", Type.GetType("System.String")));
            NotCurvedAssemblyDT.Columns.Add(new DataColumn("InsetColor", Type.GetType("System.String")));
            NotCurvedAssemblyDT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            NotCurvedAssemblyDT.Columns.Add(new DataColumn("Width", Type.GetType("System.Int32")));
            NotCurvedAssemblyDT.Columns.Add(new DataColumn("Count", Type.GetType("System.Int32")));
            NotCurvedAssemblyDT.Columns.Add(new DataColumn("Square", Type.GetType("System.Decimal")));
            NotCurvedAssemblyDT.Columns.Add(new DataColumn("Notes", Type.GetType("System.String")));
            //NotCurvedAssemblyDT.Columns.Add(new DataColumn("Worker", Type.GetType("System.String")));
            NotCurvedAssemblyDT.Columns.Add(new DataColumn("ColorType", Type.GetType("System.Int32")));
            
            DecorAssemblyDT = new DataTable();
            DecorAssemblyDT.Columns.Add(new DataColumn("Name", Type.GetType("System.String")));
            DecorAssemblyDT.Columns.Add(new DataColumn("Color", Type.GetType("System.String")));
            DecorAssemblyDT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            DecorAssemblyDT.Columns.Add(new DataColumn("Width", Type.GetType("System.Int32")));
            DecorAssemblyDT.Columns.Add(new DataColumn("Count", Type.GetType("System.Int32")));
            DecorAssemblyDT.Columns.Add(new DataColumn("Notes", Type.GetType("System.String")));
            DecorAssemblyDT.Columns.Add(new DataColumn("DecorID", Type.GetType("System.Int32")));

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
        }

        private void DeyingByMainOrderToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            string BatchName)
        {
            DataTable DistMainOrdersDT = DistMainOrdersTable(NotCurvedOrdersDT, true);
            DataTable DT = NotCurvedOrdersDT.Clone();
            DataTable ZOVOrdersNames = new DataTable();
            DataTable MarketOrdersNames = new DataTable();
            string MainOrdersID = string.Empty;
            string ClientName = string.Empty;
            string OrderName = string.Empty;

            if (DistMainOrdersDT.Select("GroupType=0").Count() > 0)
            {
                MainOrdersID = string.Empty;
                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=0"))
                    MainOrdersID += Convert.ToInt32(item["MainOrderID"]) + ",";
                MainOrdersID = MainOrdersID.Substring(0, MainOrdersID.Length - 1);
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT ClientName, MainOrders.DocNumber, MainOrders.MainOrderID, Batch.MegaBatchID, Batch.BatchID FROM MainOrders" +
                    " INNER JOIN BatchDetails ON MainOrders.MainOrderID=BatchDetails.MainOrderID" +
                    " INNER JOIN Batch ON BatchDetails.BatchID=Batch.BatchID" +
                    " INNER JOIN infiniu2_zovreference.dbo.Clients ON MainOrders.CLientID=infiniu2_zovreference.dbo.Clients.ClientID" +
                    " WHERE MainOrders.MainOrderID IN (" + MainOrdersID + ")", ConnectionStrings.ZOVOrdersConnectionString))
                {
                    DA.Fill(ZOVOrdersNames);
                }

                int RowIndex = 0;

                HSSFSheet sheet1 = hssfworkbook.CreateSheet("ЗОВ");
                sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

                sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
                sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
                sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
                sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

                sheet1.SetColumnWidth(0, 20 * 256);
                sheet1.SetColumnWidth(1, 20 * 256);
                sheet1.SetColumnWidth(2, 20 * 256);
                sheet1.SetColumnWidth(3, 6 * 256);
                sheet1.SetColumnWidth(4, 6 * 256);
                sheet1.SetColumnWidth(5, 6 * 256);

                HSSFCell cell = null;
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
                cell.CellStyle = Calibri11CS;
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Плановое время выполнения:");
                cell.CellStyle = Calibri11CS;
                RowIndex++;
                RowIndex++;
                for (int i = 0; i < DistMainOrdersDT.Rows.Count; i++)
                {
                    if (Convert.ToInt32(DistMainOrdersDT.Rows[i]["GroupType"]) == 1)
                        continue;
                    int MainOrderID = Convert.ToInt32(DistMainOrdersDT.Rows[i]["MainOrderID"]);
                    int MegaBatchID = 0;
                    int BatchID = 0;
                    DeyingDT.Clear();
                    DT.Clear();
                    DataRow[] rows = NotCurvedOrdersDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item in rows)
                        DT.Rows.Add(item.ItemArray);
                    CollectDeying(DT, ref DeyingDT);

                    DataRow[] CRows = ZOVOrdersNames.Select("MainOrderID=" + MainOrderID);
                    if (CRows.Count() > 0)
                    {
                        MegaBatchID = Convert.ToInt32(CRows[0]["MegaBatchID"]);
                        BatchID = Convert.ToInt32(CRows[0]["BatchID"]);
                        ClientName = CRows[0]["ClientName"].ToString();
                        OrderName = CRows[0]["DocNumber"].ToString();
                    }
                    if (DeyingDT.Rows.Count > 0)
                        DeyingByMainOrderToExcelSingly(ref hssfworkbook, ref sheet1,
                            Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DeyingDT, 
                            "ЗОВ " + MegaBatchID + ", " + BatchID + ", " + MainOrderID, ClientName, OrderName, string.Empty, ref RowIndex);
                }
            }
            if (DistMainOrdersDT.Select("GroupType=1").Count() > 0)
            {
                MainOrdersID = string.Empty;
                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=1"))
                    MainOrdersID += Convert.ToInt32(item["MainOrderID"]) + ",";
                MainOrdersID = MainOrdersID.Substring(0, MainOrdersID.Length - 1);
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT ClientName, MegaOrders.OrderNumber, MainOrders.MainOrderID, MainOrders.Notes, Batch.MegaBatchID, Batch.BatchID FROM MainOrders" +
                    " INNER JOIN BatchDetails ON MainOrders.MainOrderID=BatchDetails.MainOrderID" +
                    " INNER JOIN Batch ON BatchDetails.BatchID=Batch.BatchID" +
                    " INNER JOIN MegaOrders ON MainOrders.MegaOrderID=MegaOrders.MegaOrderID" +
                    " INNER JOIN infiniu2_marketingreference.dbo.Clients ON MegaOrders.ClientID=infiniu2_marketingreference.dbo.Clients.ClientID" +
                    " WHERE MainOrders.MainOrderID IN (" + MainOrdersID + ")", ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DA.Fill(MarketOrdersNames);
                }

                int RowIndex = 0;

                HSSFSheet sheet1 = hssfworkbook.CreateSheet("Маркет");
                sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

                sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
                sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
                sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
                sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

                sheet1.SetColumnWidth(0, 20 * 256);
                sheet1.SetColumnWidth(1, 20 * 256);
                sheet1.SetColumnWidth(2, 20 * 256);
                sheet1.SetColumnWidth(3, 6 * 256);
                sheet1.SetColumnWidth(4, 6 * 256);
                sheet1.SetColumnWidth(5, 6 * 256);
                
                for (int i = 0; i < DistMainOrdersDT.Rows.Count; i++)
                {
                    if (Convert.ToInt32(DistMainOrdersDT.Rows[i]["GroupType"]) == 0)
                        continue;
                    int MainOrderID = Convert.ToInt32(DistMainOrdersDT.Rows[i]["MainOrderID"]);
                    int MegaBatchID = 0;
                    int BatchID = 0;
                    DeyingDT.Clear();
                    DT.Clear();
                    string Notes = string.Empty;
                    DataRow[] rows = NotCurvedOrdersDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item in rows)
                        DT.Rows.Add(item.ItemArray);
                    CollectDeying(DT, ref DeyingDT);

                    DataRow[] CRows = MarketOrdersNames.Select("MainOrderID=" + MainOrderID);
                    if (CRows.Count() > 0)
                    {
                        MegaBatchID = Convert.ToInt32(CRows[0]["MegaBatchID"]);
                        BatchID = Convert.ToInt32(CRows[0]["BatchID"]);
                        ClientName = CRows[0]["ClientName"].ToString();
                        Notes = CRows[0]["Notes"].ToString();
                        OrderName = "№" + CRows[0]["OrderNumber"].ToString() + "-" + MainOrderID;
                    }
                    if (DeyingDT.Rows.Count > 0)
                        DeyingByMainOrderToExcelSingly(ref hssfworkbook, ref sheet1,
                            Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DeyingDT, 
                            "ЗОВ " + MegaBatchID + ", " + BatchID + ", " + MainOrderID, ClientName, OrderName, Notes, ref RowIndex);
                }
            }
        }

        private void DeyingToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            string BatchName, string ClientName, string PageName)
        {
            DeyingDT.Clear();
            CollectDeying(NotCurvedOrdersDT, ref DeyingDT);
            if (DeyingDT.Rows.Count > 0)
                DeyingToExcelSingly(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, BatchName, ClientName, "Покраска");
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

        private DataTable DistInsetColorsTable(DataTable SourceDT, bool OrderASC)
        {
            int InsetColorID = 0;
            DataTable DT = new DataTable();
            DT.Columns.Add(new DataColumn("InsetColorID", Type.GetType("System.Int32")));
            foreach (DataRow Row in SourceDT.Rows)
            {
                //if (Convert.ToInt32(Row["InsetTypeID"]) != 2 && Convert.ToInt32(Row["InsetTypeID"]) != 5 && Convert.ToInt32(Row["InsetTypeID"]) != 6 && Convert.ToInt32(Row["InsetTypeID"]) != 8)
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

        private void Fill()
        {
            string SelectCommand = @"SELECT TechStoreID AS FrontID, TechStoreName AS FrontName FROM TechStore
                WHERE TechStoreID IN (SELECT FrontID FROM FrontsConfig WHERE Enabled = 1)
                ORDER BY TechStoreName";
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
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT PatinaRAL.*, Patina.Patina FROM PatinaRAL INNER JOIN Patina ON Patina.PatinaID=PatinaRAL.PatinaID WHERE PatinaRAL.Enabled=1",
                ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(PatinaRALDataTable);
            }
            foreach (DataRow item in PatinaRALDataTable.Rows)
            {
                DataRow NewRow = PatinaDataTable.NewRow();
                NewRow["PatinaID"] = item["PatinaRALID"];
                NewRow["PatinaName"] = item["PatinaRAL"]; NewRow["Patina"] = item["Patina"];
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
            TechStoreDataTable = new DataTable();
            //using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM TechStore",
            //    ConnectionStrings.CatalogConnectionString))
            //{
            //    DA.Fill(TechStoreDataTable);
            //}
            TechStoreDataTable = TablesManager.TechStoreDataTable;

            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT * FROM DecorParameters",
                ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(DecorParametersDT);
            }
            DecorDT = new DataTable();
            SelectCommand = @"SELECT DISTINCT TechStore.TechStoreID AS DecorID, TechStore.TechStoreName AS Name, DecorConfig.ProductID FROM TechStore
                INNER JOIN DecorConfig ON TechStore.TechStoreID = DecorConfig.DecorID AND Enabled = 1 ORDER BY TechStoreName";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(DecorDT);
            }

            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT TOP 0 FrontsOrdersID, MainOrderID, FrontID, PatinaID, InsetTypeID,
                ColorID, InsetColorID, Height, Width, Count, FrontConfigID, Notes FROM FrontsOrders",
                ConnectionStrings.ZOVOrdersConnectionString))
            {
                DA.Fill(NotCurvedOrdersDT);
                NotCurvedOrdersDT.Columns.Add(new DataColumn("GroupType", Type.GetType("System.Int32")));

                //CurvedOrdersDT = NotCurvedOrdersDT.Clone();
                //AppliqueDT = NotCurvedOrdersDT.Clone();
                //GridsDT = NotCurvedOrdersDT.Clone();
                SimpleDT = NotCurvedOrdersDT.Clone();
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

        private void GetArchDecorOrders(ref DataTable DestinationDT, int FactoryID)
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

        private void GetGridMargins(int FrontID, ref int MarginHeight, ref int MarginWidth)
        {
            DataRow[] Rows = TechStoreDataTable.Select("TechStoreID = " + FrontID);
            if (Rows.Count() > 0)
            {
                if (Rows[0]["InsetHeightAdmission"] != DBNull.Value)
                    MarginHeight = Convert.ToInt32(Rows[0]["InsetHeightAdmission"]);
                if (Rows[0]["InsetWidthAdmission"] != DBNull.Value)
                    MarginWidth = Convert.ToInt32(Rows[0]["InsetWidthAdmission"]);
            }
        }

        private void GetGridsDecorOrders(ref DataTable DestinationDT, int FactoryID)
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

        private void GetNotArchDecorOrders(ref DataTable DestinationDT, int FactoryID)
        {
            string SelectCommand = string.Empty;
            DataTable DT = new DataTable();

            SelectCommand = @"SELECT DecorOrderID, MainOrderID, ProductID, DecorID, DecorConfigID, ColorID, PatinaID,
                Height, Length, Width, Count, Notes FROM DecorOrders
                WHERE FactoryID=" + FactoryID + " AND ProductID NOT IN (31, 4, 18, 32, 10, 11, 12) AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE ProfilWorkAssignmentID=" + WorkAssignmentID + "))";
            if (FactoryID == 2)
                SelectCommand = @"SELECT DecorOrderID, MainOrderID, ProductID, DecorID, DecorConfigID, ColorID, PatinaID,
                    Height, Length, Width, Count, Notes FROM DecorOrders
                    WHERE FactoryID=" + FactoryID + " AND ProductID NOT IN (31, 4, 18, 32, 10, 11, 12) AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE TPSWorkAssignmentID=" + WorkAssignmentID + "))";

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
        
        private void GetNotCurvedFrontsOrders(ref DataTable DestinationDT, int FactoryID)
        {
            string SelectCommand = string.Empty;
            DataTable DT = new DataTable();

            SelectCommand = $"SELECT FrontsOrdersID, MainOrderID, FrontID, PatinaID, InsetTypeID, " +
                $"ColorID, InsetColorID, Height, Width, Count, FrontConfigID, Notes FROM FrontsOrders " +
                $" WHERE Width<>-1 AND FrontID IN " + 
                $" ({Convert.ToInt32(Fronts.Tafel1_19)}, {Convert.ToInt32(Fronts.Tafel1Gl_19)}, {Convert.ToInt32(Fronts.Tafel1R1)}, " +
                $" {Convert.ToInt32(Fronts.Tafel1R1Gl)}, {Convert.ToInt32(Fronts.Tafel1R2)}, {Convert.ToInt32(Fronts.Tafel1R2Gl)}, " +
                $" {Convert.ToInt32(Fronts.Tafel1_16)}, {Convert.ToInt32(Fronts.Tafel1Gl_16)}) " +
                $" AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE ProfilWorkAssignmentID= { WorkAssignmentID }))";
            if (FactoryID == 2)
                SelectCommand = $"SELECT FrontsOrdersID, MainOrderID, FrontID, PatinaID, InsetTypeID, " +
               $"ColorID, InsetColorID, Height, Width, Count, FrontConfigID, Notes FROM FrontsOrders " +
               $" WHERE Width<>-1 AND FrontID IN " +
               $" ({Convert.ToInt32(Fronts.Tafel1_19)}, {Convert.ToInt32(Fronts.Tafel1Gl_19)}, {Convert.ToInt32(Fronts.Tafel1R1)}, " +
               $" {Convert.ToInt32(Fronts.Tafel1R1Gl)}, {Convert.ToInt32(Fronts.Tafel1R2)}, {Convert.ToInt32(Fronts.Tafel1R2Gl)}, " +
               $" {Convert.ToInt32(Fronts.Tafel1_16)}, {Convert.ToInt32(Fronts.Tafel1Gl_16)}) " +
               $" AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE TPSWorkAssignmentID= { WorkAssignmentID }))";
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

        private void GetSimpleFronts(DataTable SourceDT, ref DataTable DestinationDT)
        {
            DataRow[] rows = SourceDT.Select();
            foreach (DataRow dr in rows)
                DestinationDT.Rows.Add(dr.ItemArray);
        }

        private void GridsDecorAssemblyByMainOrderToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            string BatchName)
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
                             BatchName, ClientName, string.Empty, OrderName, Notes, ref RowIndex);
                        RowIndex++;
                        GridsDecorAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                            Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DecorAssemblyDT,
                            BatchName, ClientName, "ДУБЛЬ", OrderName, Notes, ref RowIndex);
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
                             BatchName, ClientName, string.Empty, OrderName, Notes, ref RowIndex);
                        RowIndex++;
                        GridsDecorAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                            Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DecorAssemblyDT,
                            BatchName, ClientName, "ДУБЛЬ", OrderName, Notes, ref RowIndex);
                    }
                    RowIndex++;
                }
            }
        }

        private void NotArchDecorAssemblyByMainOrderToExcel(ref HSSFWorkbook hssfworkbook,
                    HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            string BatchName)
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
                             BatchName, ClientName, string.Empty, OrderName, Notes, ref RowIndex);
                        RowIndex++;
                        NotArchDecorAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                            Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DecorAssemblyDT,
                            BatchName, ClientName, "ДУБЛЬ", OrderName, Notes, ref RowIndex);
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
                             BatchName, ClientName, string.Empty, OrderName, Notes, ref RowIndex);
                        RowIndex++;
                        NotArchDecorAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                            Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DecorAssemblyDT,
                            BatchName, ClientName, "ДУБЛЬ", OrderName, Notes, ref RowIndex);
                    }
                    RowIndex++;
                }
            }
        }

        private struct PlanningTimeFund
        {
            public decimal fund;
            public decimal time;
        }
    }

    public class TafelAssignments : IFirstProfilName, IColorName, IPatinaName
    {
        public DataTable FrameColorsDataTable = null;
        public DataTable FrontsDataTable = null;
        public DataTable PatinaDataTable = null;
        private DataTable ArchDecorOrdersDT;
        private DataTable AssemblyDT;
        private DateTime CurrentDate;
        private readonly string _appovingUser = "Егорченко Р.П.";
        private DataTable DecorAssemblyDT;
        private DataTable DecorDT;
        private DataTable DecorParametersDT;
        private readonly FileManager FM = new FileManager();
        private DataTable GridsDecorOrdersDT;
        private DataTable NotArchDecorOrdersDT;
        private DataTable NotCurvedOrdersDT;
        private DataTable PatinaRALDataTable = null;
        private DataTable RoughCutDT;

        private int _workAssignmentID = 0;
        public int WorkAssignmentID { get => _workAssignmentID; set => _workAssignmentID = value; }

        public TafelAssignments()
        {
        }

        public void ArchDecorAssemblyToExcelSingly(ref HSSFWorkbook hssfworkbook, ref HSSFSheet sheet1,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            DataTable DT, string BatchName, string ClientName, string OperationName, string OrderName, string Notes, ref int RowIndex)
        {
            HSSFCell cell = null;

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

        public void AssemblyToExcel(ref HSSFWorkbook hssfworkbook, ref HSSFSheet sheet1,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            DataTable DT, string BatchName, string ClientName, string OrderName, string PageName, string Notes, ref int RowIndex)
        {
            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, $"Распечатал: Дата/время { CurrentDate.ToString("dd.MM.yyyy HH:mm") }" +
                $" \r\n ФИО: { Security.CurrentUserShortName}");
            cell.CellStyle = Calibri11CS;

            RowIndex++;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, _appovingUser);
            cell.CellStyle = Calibri11CS;
            RowIndex++;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, PageName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, $"План. отгрузка: {DateTime.Now.AddDays(14).ToShortDateString()}");
            cell.CellStyle = Calibri11CS;
            RowIndex++;


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

            if (Notes.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(++RowIndex), 0, "Примечание: " + Notes);
                cell.CellStyle = CalibriBold11CS;
            }

            RowIndex++;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Профиль");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Цвет профиля");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 5, "м.кв.");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 6, "Прим.");
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
                DifferentDecorCount = DV.ToTable(true, new string[] { "ColorID" }).Rows.Count;
            }

            if (DT.Rows.Count > 0)
                CType = Convert.ToInt32(DT.Rows[0]["ColorID"]);

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
                    if (DT.Columns[y].ColumnName == "ColorID")
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
                    if (CType != Convert.ToInt32(DT.Rows[x + 1]["ColorID"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "ColorID")
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

                        cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                        cell.SetCellValue(Convert.ToDouble(TotalSquare));
                        cell.CellStyle = TableHeaderCS;

                        CType = Convert.ToInt32(DT.Rows[x + 1]["ColorID"]);
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
                            if (DT.Columns[y].ColumnName == "ColorID")
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

                        cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                        cell.SetCellValue(Convert.ToDouble(TotalSquare));
                        cell.CellStyle = TableHeaderDecCS;
                    }
                    RowIndex++;

                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "ColorID")
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

                    cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                    cell.SetCellValue(Convert.ToDouble(AllTotalSquare));
                    cell.CellStyle = TableHeaderDecCS;
                }
                RowIndex++;
            }
        }

        public void ClearOrders()
        {
            NotCurvedOrdersDT.Clear();
            NotArchDecorOrdersDT.Clear();
            ArchDecorOrdersDT.Clear();
            GridsDecorOrdersDT.Clear();
        }

        public void CreateExcel(string ClientName, string BatchName, ref string sSourceFileName)
        {
            GetCurrentDate();
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

            HSSFCellStyle WorkerColumnCS = hssfworkbook.CreateCellStyle();
            WorkerColumnCS.Alignment = HSSFCellStyle.ALIGN_RIGHT;
            WorkerColumnCS.BorderBottom = HSSFCellStyle.BORDER_THIN;
            WorkerColumnCS.BottomBorderColor = HSSFColor.BLACK.index;
            WorkerColumnCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            WorkerColumnCS.LeftBorderColor = HSSFColor.BLACK.index;
            WorkerColumnCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            WorkerColumnCS.RightBorderColor = HSSFColor.BLACK.index;
            WorkerColumnCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            WorkerColumnCS.TopBorderColor = HSSFColor.BLACK.index;
            WorkerColumnCS.SetFont(Serif10F);

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

            NotArchDecorAssemblyByMainOrderToExcel(ref hssfworkbook, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, BatchName);

            ArchDecorAssemblyByMainOrderToExcel(ref hssfworkbook, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, BatchName);

            GridsDecorAssemblyByMainOrderToExcel(ref hssfworkbook, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, BatchName);

            AssemblyByMainOrderToExcel(ref hssfworkbook,
                Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, BatchName);

            string FileName = WorkAssignmentID + " " + BatchName;
            string tempFolder = @"\\192.168.1.6\Public\_ДЕЙСТВУЮЩИЕ\ПРОИЗВОДСТВО\ТПС\инфиниум\";
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
                file = new FileInfo(tempFolder + @"\" + FileName + "(" + j++ + ").xls");

            FileStream NewFile = new FileStream(file.FullName, FileMode.Create);
            hssfworkbook.Write(NewFile);
            NewFile.Close();

            System.Diagnostics.Process.Start(file.FullName);

            //string sSourceFolder = System.Environment.GetEnvironmentVariable("TEMP");
            //string sFolderPath = "Общие файлы/Производство/Задания в работу";
            //string sDestFolder = Configs.DocumentsPath + sFolderPath;
            //sSourceFileName = GetFileName(sDestFolder, BatchName);

            //FileInfo file = new FileInfo(sSourceFolder + @"\" + sSourceFileName);
            //FileStream NewFile = new FileStream(file.FullName, FileMode.Create);
            //hssfworkbook.Write(NewFile);
            //NewFile.Close();
        }

        public string GetColorName(int ColorID)
        {
            string ColorName = string.Empty;
            DataRow[] Rows = FrameColorsDataTable.Select("ColorID = " + ColorID);
            if (Rows.Count() > 0)
                ColorName = Rows[0]["ColorName"].ToString();
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

        public string GetFrontName(int FrontID)
        {
            string FrontName = string.Empty;
            DataRow[] Rows = FrontsDataTable.Select("FrontID = " + FrontID);
            if (Rows.Count() > 0)
                FrontName = Rows[0]["FrontName"].ToString();
            return FrontName;
        }

        public bool GetOrders(DataTable EditFrontOrdersDT, int FactoryID)
        {
            NotCurvedOrdersDT = EditFrontOrdersDT.Copy();
            //GetNotCurvedFrontsOrders(ref NotCurvedOrdersDT, FactoryID, Fronts.Tafel3);
            //GetNotCurvedFrontsOrders(ref NotCurvedOrdersDT, FactoryID, Fronts.Tafel2);
            //GetNotCurvedFrontsOrders(ref NotCurvedOrdersDT, FactoryID, Fronts.Tafel3Gl);
            GetNotArchDecorOrders(ref NotArchDecorOrdersDT, FactoryID);
            GetArchDecorOrders(ref ArchDecorOrdersDT, FactoryID);
            GetGridsDecorOrders(ref GridsDecorOrdersDT, FactoryID);

            if (NotCurvedOrdersDT.Rows.Count == 0 && NotArchDecorOrdersDT.Rows.Count == 0 && ArchDecorOrdersDT.Rows.Count == 0 && GridsDecorOrdersDT.Rows.Count == 0)
                return false;
            else
                return true;
        }

        public string GetPatinaName(int PatinaID)
        {
            string PatinaName = string.Empty;
            DataRow[] Rows = PatinaDataTable.Select("PatinaID = " + PatinaID);
            if (Rows.Count() > 0)
                PatinaName = Rows[0]["PatinaName"].ToString();
            return PatinaName;
        }

        public void GridsDecorAssemblyToExcelSingly(ref HSSFWorkbook hssfworkbook, ref HSSFSheet sheet1,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            DataTable DT, string BatchName, string ClientName, string OperationName, string OrderName, string Notes, ref int RowIndex)
        {
            HSSFCell cell = null;

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

        public void NotArchDecorAssemblyToExcelSingly(ref HSSFWorkbook hssfworkbook, ref HSSFSheet sheet1,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            DataTable DT, string BatchName, string ClientName, string OperationName, string OrderName, string Notes, ref int RowIndex)
        {
            HSSFCell cell = null;

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

        public void RoughCutToExcel(ref HSSFWorkbook hssfworkbook, ref HSSFSheet sheet1,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            DataTable DT, string BatchName, string ClientName, string OrderName, string PageName, string Notes, ref int RowIndex)
        {
            HSSFCell cell = null;

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
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Заказ:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, OrderName);
            cell.CellStyle = CalibriBold11CS;

            if (Notes.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(++RowIndex), 0, "Примечание: " + Notes);
                cell.CellStyle = CalibriBold11CS;
            }

            HSSFCellStyle CalibriBold11CS1 = hssfworkbook.CreateCellStyle();
            CalibriBold11CS1.Alignment = HSSFCellStyle.ALIGN_CENTER;
            CalibriBold11CS1.BorderBottom = HSSFCellStyle.BORDER_THIN;
            CalibriBold11CS1.BottomBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS1.BorderLeft = HSSFCellStyle.BORDER_THIN;
            CalibriBold11CS1.LeftBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS1.BorderRight = HSSFCellStyle.BORDER_THIN;
            CalibriBold11CS1.RightBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS1.BorderTop = HSSFCellStyle.BORDER_THIN;
            CalibriBold11CS1.TopBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS1.SetFont(CalibriBold11F);

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "черновой распил");
            cell.CellStyle = CalibriBold11CS1;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, string.Empty);
            cell.CellStyle = CalibriBold11CS1;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, string.Empty);
            cell.CellStyle = CalibriBold11CS1;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 5, string.Empty);
            cell.CellStyle = CalibriBold11CS1;
            sheet1.AddMergedRegion(new NPOI.HSSF.Util.Region(RowIndex, 2, RowIndex, 5));

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 6, "чистовой размер");
            cell.CellStyle = CalibriBold11CS1;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 7, string.Empty);
            cell.CellStyle = CalibriBold11CS1;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 8, string.Empty);
            cell.CellStyle = CalibriBold11CS1;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 9, string.Empty);
            cell.CellStyle = CalibriBold11CS1;
            sheet1.AddMergedRegion(new NPOI.HSSF.Util.Region(RowIndex, 6, RowIndex, 9));

            RowIndex++;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Профиль");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Цвет профиля");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 5, "м.кв.");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 6, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 7, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 8, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 9, "м.кв.");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 10, "Прим.");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            int CType = -1;
            int AllTotalAmount1 = 0;
            int TotalAmount1 = 0;
            decimal AllTotalSquare1 = 0;
            decimal TotalSquare1 = 0;

            int AllTotalAmount2 = 0;
            int TotalAmount2 = 0;
            decimal AllTotalSquare2 = 0;
            decimal TotalSquare2 = 0;

            int DifferentDecorCount = 0;

            using (DataView DV = new DataView(DT))
            {
                DifferentDecorCount = DV.ToTable(true, new string[] { "ColorID" }).Rows.Count;
            }

            if (DT.Rows.Count > 0)
                CType = Convert.ToInt32(DT.Rows[0]["ColorID"]);

            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count1"] != DBNull.Value)
                {
                    AllTotalAmount1 += Convert.ToInt32(DT.Rows[x]["Count1"]);
                    TotalAmount1 += Convert.ToInt32(DT.Rows[x]["Count1"]);
                }
                if (DT.Rows[x]["Square1"] != DBNull.Value)
                {
                    AllTotalSquare1 += Convert.ToDecimal(DT.Rows[x]["Square1"]);
                    TotalSquare1 += Convert.ToDecimal(DT.Rows[x]["Square1"]);
                }
                if (DT.Rows[x]["Count2"] != DBNull.Value)
                {
                    AllTotalAmount2 += Convert.ToInt32(DT.Rows[x]["Count2"]);
                    TotalAmount2 += Convert.ToInt32(DT.Rows[x]["Count2"]);
                }
                if (DT.Rows[x]["Square2"] != DBNull.Value)
                {
                    AllTotalSquare2 += Convert.ToDecimal(DT.Rows[x]["Square2"]);
                    TotalSquare2 += Convert.ToDecimal(DT.Rows[x]["Square2"]);
                }

                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "ColorID")
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
                    if (CType != Convert.ToInt32(DT.Rows[x + 1]["ColorID"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "ColorID")
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
                        cell.SetCellValue(TotalAmount1);
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                        cell.SetCellValue(Convert.ToDouble(TotalSquare1));
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(8);
                        cell.SetCellValue(TotalAmount2);
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(9);
                        cell.SetCellValue(Convert.ToDouble(TotalSquare2));
                        cell.CellStyle = TableHeaderCS;

                        CType = Convert.ToInt32(DT.Rows[x + 1]["ColorID"]);
                        TotalAmount1 = 0;
                        TotalSquare1 = 0;
                        TotalAmount2 = 0;
                        TotalSquare2 = 0;
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
                            if (DT.Columns[y].ColumnName == "ColorID")
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
                        cell.SetCellValue(TotalAmount1);
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                        cell.SetCellValue(Convert.ToDouble(TotalSquare1));
                        cell.CellStyle = TableHeaderDecCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(8);
                        cell.SetCellValue(TotalAmount2);
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(9);
                        cell.SetCellValue(Convert.ToDouble(TotalSquare2));
                        cell.CellStyle = TableHeaderDecCS;
                    }
                    RowIndex++;

                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "ColorID")
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
                    cell.SetCellValue(AllTotalAmount1);
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                    cell.SetCellValue(Convert.ToDouble(AllTotalSquare1));
                    cell.CellStyle = TableHeaderDecCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(8);
                    cell.SetCellValue(AllTotalAmount2);
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(9);
                    cell.SetCellValue(Convert.ToDouble(AllTotalSquare2));
                    cell.CellStyle = TableHeaderDecCS;
                }
                RowIndex++;
            }
        }

        private void ArchDecorAssemblyByMainOrderToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            string BatchName)
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

                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=0"))
                {
                    DecorAssemblyDT.Clear();
                    DT.Clear();
                    int MainOrderID = Convert.ToInt32(item["MainOrderID"]);
                    DataRow[] rows = ArchDecorOrdersDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item1 in rows)
                        DT.Rows.Add(item1.ItemArray);
                    GroupAndSortDecor(DT, ref DecorAssemblyDT);

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
                             BatchName, ClientName, string.Empty, OrderName, Notes, ref RowIndex);
                        RowIndex++;
                        ArchDecorAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                            Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DecorAssemblyDT,
                            BatchName, ClientName, "ДУБЛЬ", OrderName, Notes, ref RowIndex);
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

                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=1"))
                {
                    DecorAssemblyDT.Clear();
                    DT.Clear();
                    int MainOrderID = Convert.ToInt32(item["MainOrderID"]);
                    DataRow[] rows = ArchDecorOrdersDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item1 in rows)
                        DT.Rows.Add(item1.ItemArray);
                    GroupAndSortDecor(DT, ref DecorAssemblyDT);

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
                             BatchName, ClientName, string.Empty, OrderName, Notes, ref RowIndex);
                        RowIndex++;
                        ArchDecorAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                            Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DecorAssemblyDT,
                            BatchName, ClientName, "ДУБЛЬ", OrderName, Notes, ref RowIndex);
                    }
                    RowIndex++;
                }
            }
        }

        private void AssemblyByMainOrderToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            string BatchName)
        {
            DataTable DistMainOrdersDT = DistMainOrdersTable(NotCurvedOrdersDT, true);
            DataTable DT = NotCurvedOrdersDT.Clone();
            DataTable DT1 = NotCurvedOrdersDT.Clone();
            DataTable ZOVOrdersNames = new DataTable();
            DataTable MarketOrdersNames = new DataTable();
            string MainOrdersID = string.Empty;
            string ClientName = string.Empty;
            string OrderName = string.Empty;

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
            }

            if (DistMainOrdersDT.Select("GroupType=0").Count() > 0)
            {
                for (int i = 0; i < DistMainOrdersDT.Rows.Count; i++)
                {
                    if (Convert.ToInt32(DistMainOrdersDT.Rows[i]["GroupType"]) == 1)
                        continue;
                    int MainOrderID = Convert.ToInt32(DistMainOrdersDT.Rows[i]["MainOrderID"]);

                    DataRow[] CRows = ZOVOrdersNames.Select("MainOrderID=" + MainOrderID);
                    if (CRows.Count() > 0)
                    {
                        ClientName = CRows[0]["ClientName"].ToString();
                        OrderName = CRows[0]["DocNumber"].ToString();
                    }

                    string Notes = string.Empty;

                    RoughCutDT.Clear();
                    DT.Clear();
                    DataTable DistGroupNumbersDT = DistGroupNumbersTable(NotCurvedOrdersDT.Select("MainOrderID=" + MainOrderID), true);
                    for (int j = 0; j < DistGroupNumbersDT.Rows.Count; j++)
                    {
                        bool CollectByHeight = false;
                        int GroupNumber = Convert.ToInt32(DistGroupNumbersDT.Rows[j]["GroupNumber"]);
                        DataRow[] rows = NotCurvedOrdersDT.Select("MainOrderID=" + MainOrderID + " AND GroupNumber=" + GroupNumber);
                        if (rows.Count() == 0)
                            continue;
                        int Height = Convert.ToInt32(rows[0]["Height"]);
                        int Width = Convert.ToInt32(rows[0]["Width"]);
                        DT1.Clear();
                        foreach (DataRow item in rows)
                        {
                            if (Height == Convert.ToInt32(item["Height"]))
                                CollectByHeight = true;
                            else
                                CollectByHeight = false;
                            if (Height == Convert.ToInt32(item["Height"]) && Width == Convert.ToInt32(item["Width"]))
                            {
                                if (Height > Width)
                                    CollectByHeight = true;
                                else
                                    CollectByHeight = false;
                            }
                            DT1.Rows.Add(item.ItemArray);
                            DT.Rows.Add(item.ItemArray);
                        }
                        if (rows.Count() == 1)
                        {
                            RoughCutCoolect(DT1, ref RoughCutDT);
                        }
                        else
                        {
                            if (CollectByHeight)
                                RoughCutCoolectByHeight(DT1, ref RoughCutDT);
                            else
                                RoughCutCoolectByWidth(DT1, ref RoughCutDT);
                        }
                    }
                    if (DT.Rows.Count == 0)
                        continue;

                    int RowIndex = 0;

                    HSSFSheet sheet1 = hssfworkbook.CreateSheet(OrderName.Replace("/", "-"));
                    sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

                    sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
                    sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
                    sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
                    sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

                    sheet1.SetColumnWidth(0, 16 * 256);
                    sheet1.SetColumnWidth(1, 18 * 256);
                    sheet1.SetColumnWidth(2, 6 * 256);
                    sheet1.SetColumnWidth(3, 6 * 256);
                    sheet1.SetColumnWidth(4, 6 * 256);
                    sheet1.SetColumnWidth(5, 7 * 256);
                    sheet1.SetColumnWidth(6, 6 * 256);
                    sheet1.SetColumnWidth(7, 6 * 256);
                    sheet1.SetColumnWidth(8, 6 * 256);
                    sheet1.SetColumnWidth(9, 7 * 256);

                    HSSFCell cell = null;
                    cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
                    cell.CellStyle = Calibri11CS;
                    cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Плановое время выполнения:");
                    cell.CellStyle = Calibri11CS;
                    RowIndex++;
                    RowIndex++;

                    if (RoughCutDT.Rows.Count > 0)
                    {
                        RoughCutToExcel(ref hssfworkbook, ref sheet1, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS,
                            RoughCutDT, BatchName, ClientName, OrderName, "ДСТП-18", string.Empty, ref RowIndex);
                        RowIndex++;
                    }

                    AssemblyDT.Clear();
                    SetOfCovers(RoughCutDT, ref AssemblyDT);

                    if (AssemblyDT.Rows.Count > 0)
                    {
                        AssemblyToExcel(ref hssfworkbook, ref sheet1, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS,
                            AssemblyDT, BatchName, ClientName, OrderName, "Рубашки", string.Empty, ref RowIndex);
                        RowIndex++;
                    }

                    AssemblyDT.Clear();
                    CalibrationPlate(RoughCutDT, ref AssemblyDT);

                    if (AssemblyDT.Rows.Count > 0)
                    {
                        AssemblyToExcel(ref hssfworkbook, ref sheet1, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS,
                            AssemblyDT, BatchName, ClientName, OrderName, "Калибровка", string.Empty, ref RowIndex);
                        RowIndex++;
                    }

                    AssemblyDT.Clear();
                    GroupAndSortFronts(DT, ref AssemblyDT);

                    if (AssemblyDT.Rows.Count > 0)
                    {
                        AssemblyToExcel(ref hssfworkbook, ref sheet1, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS,
                            AssemblyDT, BatchName, ClientName, OrderName, "Оклейка", string.Empty, ref RowIndex);
                        RowIndex++;
                    }

                    AssemblyDT.Clear();
                    GroupAndSortGlossFronts(DT, ref AssemblyDT);

                    if (AssemblyDT.Rows.Count > 0)
                    {
                        AssemblyToExcel(ref hssfworkbook, ref sheet1, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS,
                            AssemblyDT, BatchName, ClientName, OrderName, "Отделка кромок", string.Empty, ref RowIndex);
                        RowIndex++;
                    }

                    AssemblyDT.Clear();
                    GroupAndSortNotGlossFronts(DT, ref AssemblyDT);

                    if (AssemblyDT.Rows.Count > 0)
                    {
                        AssemblyToExcel(ref hssfworkbook, ref sheet1, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS,
                            AssemblyDT, BatchName, ClientName, OrderName, "Отделка кромок", string.Empty, ref RowIndex);
                        RowIndex++;

                        AssemblyToExcel(ref hssfworkbook, ref sheet1, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS,
                            AssemblyDT, BatchName, ClientName, OrderName, "Лак глянец, шлифовка и полировка глянца", string.Empty, ref RowIndex);
                        RowIndex++;
                    }

                    AssemblyDT.Clear();
                    GroupAndSortColoredFronts(DT, ref AssemblyDT);

                    if (AssemblyDT.Rows.Count > 0)
                    {
                        AssemblyToExcel(ref hssfworkbook, ref sheet1, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS,
                            AssemblyDT, BatchName, ClientName, OrderName, "Бейц, Лак", string.Empty, ref RowIndex);
                        RowIndex++;
                    }

                    AssemblyDT.Clear();
                    GroupAndSortFronts(DT, ref AssemblyDT);

                    if (AssemblyDT.Rows.Count > 0)
                    {
                        AssemblyToExcel(ref hssfworkbook, ref sheet1, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS,
                            AssemblyDT, BatchName, ClientName, OrderName, "Сверление и упаковка", string.Empty, ref RowIndex);
                        RowIndex++;
                    }

                    AssemblyDT.Clear();
                    GroupAndSortFrontsWithHands(DT, ref AssemblyDT);

                    if (AssemblyDT.Rows.Count > 0)
                    {
                        AssemblyToExcel(ref hssfworkbook, ref sheet1, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS,
                            AssemblyDT, BatchName, ClientName, OrderName, "Тафель с ручкой УКВ-7", string.Empty, ref RowIndex);
                        RowIndex++;
                    }
                }
            }

            if (DistMainOrdersDT.Select("GroupType=1").Count() > 0)
            {
                for (int i = 0; i < DistMainOrdersDT.Rows.Count; i++)
                {
                    if (Convert.ToInt32(DistMainOrdersDT.Rows[i]["GroupType"]) == 0)
                        continue;
                    int MainOrderID = Convert.ToInt32(DistMainOrdersDT.Rows[i]["MainOrderID"]);
                    string Notes = string.Empty;

                    DataRow[] CRows = MarketOrdersNames.Select("MainOrderID=" + MainOrderID);
                    if (CRows.Count() > 0)
                    {
                        ClientName = CRows[0]["ClientName"].ToString();
                        Notes = CRows[0]["Notes"].ToString();
                        OrderName = "№" + CRows[0]["OrderNumber"].ToString() + "-" + MainOrderID;
                    }

                    RoughCutDT.Clear();
                    DT.Clear();
                    DataTable DistGroupNumbersDT = DistGroupNumbersTable(NotCurvedOrdersDT.Select("MainOrderID=" + MainOrderID), true);
                    for (int j = 0; j < DistGroupNumbersDT.Rows.Count; j++)
                    {
                        bool CollectByHeight = false;
                        int GroupNumber = Convert.ToInt32(DistGroupNumbersDT.Rows[j]["GroupNumber"]);
                        DataRow[] rows = NotCurvedOrdersDT.Select("MainOrderID=" + MainOrderID + " AND GroupNumber=" + GroupNumber);
                        if (rows.Count() == 0)
                            continue;
                        int Height = Convert.ToInt32(rows[0]["Height"]);
                        int Width = Convert.ToInt32(rows[0]["Width"]);
                        DT1.Clear();
                        foreach (DataRow item in rows)
                        {
                            if (Height == Convert.ToInt32(item["Height"]))
                                CollectByHeight = true;
                            else
                                CollectByHeight = false;
                            if (Height == Convert.ToInt32(item["Height"]) && Width == Convert.ToInt32(item["Width"]))
                            {
                                if (Height > Width)
                                    CollectByHeight = true;
                                else
                                    CollectByHeight = false;
                            }
                            DT1.Rows.Add(item.ItemArray);
                            DT.Rows.Add(item.ItemArray);
                        }
                        if (rows.Count() == 1)
                        {
                            RoughCutCoolect(DT1, ref RoughCutDT);
                        }
                        else
                        {
                            if (CollectByHeight)
                                RoughCutCoolectByHeight(DT1, ref RoughCutDT);
                            else
                                RoughCutCoolectByWidth(DT1, ref RoughCutDT);
                        }
                    }

                    if (DT.Rows.Count == 0)
                        continue;

                    int RowIndex = 0;

                    HSSFSheet sheet1 = hssfworkbook.CreateSheet(OrderName.Replace("/", "-"));
                    sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

                    sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
                    sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
                    sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
                    sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

                    sheet1.SetColumnWidth(0, 16 * 256);
                    sheet1.SetColumnWidth(1, 18 * 256);
                    sheet1.SetColumnWidth(2, 6 * 256);
                    sheet1.SetColumnWidth(3, 6 * 256);
                    sheet1.SetColumnWidth(4, 6 * 256);
                    sheet1.SetColumnWidth(5, 7 * 256);
                    sheet1.SetColumnWidth(6, 6 * 256);
                    sheet1.SetColumnWidth(7, 6 * 256);
                    sheet1.SetColumnWidth(8, 6 * 256);
                    sheet1.SetColumnWidth(9, 7 * 256);

                    HSSFCell cell = null;
                    cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
                    cell.CellStyle = Calibri11CS;
                    cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Плановое время выполнения:");
                    cell.CellStyle = Calibri11CS;
                    RowIndex++;
                    RowIndex++;

                    if (RoughCutDT.Rows.Count > 0)
                    {
                        RoughCutToExcel(ref hssfworkbook, ref sheet1, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS,
                            RoughCutDT, BatchName, ClientName, OrderName, "ДСТП-18", string.Empty, ref RowIndex);
                        RowIndex++;
                    }

                    AssemblyDT.Clear();
                    SetOfCovers(RoughCutDT, ref AssemblyDT);

                    if (AssemblyDT.Rows.Count > 0)
                    {
                        AssemblyToExcel(ref hssfworkbook, ref sheet1, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS,
                            AssemblyDT, BatchName, ClientName, OrderName, "Рубашки", string.Empty, ref RowIndex);
                        RowIndex++;
                    }

                    AssemblyDT.Clear();
                    CalibrationPlate(RoughCutDT, ref AssemblyDT);

                    if (AssemblyDT.Rows.Count > 0)
                    {
                        AssemblyToExcel(ref hssfworkbook, ref sheet1, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS,
                            AssemblyDT, BatchName, ClientName, OrderName, "Калибровка", string.Empty, ref RowIndex);
                        RowIndex++;
                    }

                    AssemblyDT.Clear();
                    GroupAndSortFronts(DT, ref AssemblyDT);

                    if (AssemblyDT.Rows.Count > 0)
                    {
                        AssemblyToExcel(ref hssfworkbook, ref sheet1, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS,
                            AssemblyDT, BatchName, ClientName, OrderName, "Оклейка", string.Empty, ref RowIndex);
                        RowIndex++;
                    }

                    AssemblyDT.Clear();
                    GroupAndSortGlossFronts(DT, ref AssemblyDT);

                    if (AssemblyDT.Rows.Count > 0)
                    {
                        AssemblyToExcel(ref hssfworkbook, ref sheet1, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS,
                            AssemblyDT, BatchName, ClientName, OrderName, "Отделка кромок", string.Empty, ref RowIndex);
                        RowIndex++;
                    }

                    AssemblyDT.Clear();
                    GroupAndSortNotGlossFronts(DT, ref AssemblyDT);

                    if (AssemblyDT.Rows.Count > 0)
                    {
                        AssemblyToExcel(ref hssfworkbook, ref sheet1, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS,
                            AssemblyDT, BatchName, ClientName, OrderName, "Отделка кромок", string.Empty, ref RowIndex);
                        RowIndex++;

                        AssemblyToExcel(ref hssfworkbook, ref sheet1, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS,
                            AssemblyDT, BatchName, ClientName, OrderName, "Лак глянец, шлифовка и полировка глянца", string.Empty, ref RowIndex);
                        RowIndex++;
                    }

                    AssemblyDT.Clear();
                    GroupAndSortColoredFronts(DT, ref AssemblyDT);

                    if (AssemblyDT.Rows.Count > 0)
                    {
                        AssemblyToExcel(ref hssfworkbook, ref sheet1, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS,
                            AssemblyDT, BatchName, ClientName, OrderName, "Бейц, Лак", string.Empty, ref RowIndex);
                        RowIndex++;
                    }

                    AssemblyDT.Clear();
                    GroupAndSortFronts(DT, ref AssemblyDT);

                    if (AssemblyDT.Rows.Count > 0)
                    {
                        AssemblyToExcel(ref hssfworkbook, ref sheet1, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS,
                            AssemblyDT, BatchName, ClientName, OrderName, "Сверление и упаковка", string.Empty, ref RowIndex);
                        RowIndex++;
                    }

                    AssemblyDT.Clear();
                    GroupAndSortFrontsWithHands(DT, ref AssemblyDT);

                    if (AssemblyDT.Rows.Count > 0)
                    {
                        AssemblyToExcel(ref hssfworkbook, ref sheet1, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS,
                            AssemblyDT, BatchName, ClientName, OrderName, "Тафель с ручкой УКВ-7", string.Empty, ref RowIndex);
                        RowIndex++;
                    }
                }
            }
        }

        private void CalibrationPlate(DataTable SourceDT, ref DataTable DestinationDT)
        {
            for (int i = 0; i < SourceDT.Rows.Count; i++)
            {
                if (SourceDT.Rows[i]["Height1"] == DBNull.Value)
                    continue;
                DataRow NewRow = DestinationDT.NewRow();
                NewRow["Name"] = SourceDT.Rows[i]["Name"];
                NewRow["FrameColor"] = SourceDT.Rows[i]["FrameColor"];
                NewRow["Height"] = SourceDT.Rows[i]["Height1"];
                NewRow["Width"] = SourceDT.Rows[i]["Width1"];
                NewRow["Count"] = SourceDT.Rows[i]["Count1"];
                NewRow["Square"] = SourceDT.Rows[i]["Square1"];
                NewRow["Notes"] = SourceDT.Rows[i]["Notes"];
                NewRow["ColorID"] = SourceDT.Rows[i]["ColorID"];
                DestinationDT.Rows.Add(NewRow);
            }
        }

        private void Create()
        {
            DecorDT = new DataTable();
            DecorParametersDT = new DataTable();

            NotCurvedOrdersDT = new DataTable();
            NotArchDecorOrdersDT = new DataTable();
            ArchDecorOrdersDT = new DataTable();
            GridsDecorOrdersDT = new DataTable();

            DecorAssemblyDT = new DataTable();
            DecorAssemblyDT.Columns.Add(new DataColumn("Name", Type.GetType("System.String")));
            DecorAssemblyDT.Columns.Add(new DataColumn("Color", Type.GetType("System.String")));
            DecorAssemblyDT.Columns.Add(new DataColumn("Patina", Type.GetType("System.String")));
            DecorAssemblyDT.Columns.Add(new DataColumn("Length", Type.GetType("System.Int32")));
            DecorAssemblyDT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            DecorAssemblyDT.Columns.Add(new DataColumn("Width", Type.GetType("System.Int32")));
            DecorAssemblyDT.Columns.Add(new DataColumn("Count", Type.GetType("System.Int32")));
            DecorAssemblyDT.Columns.Add(new DataColumn("Notes", Type.GetType("System.String")));
            DecorAssemblyDT.Columns.Add(new DataColumn("DecorID", Type.GetType("System.Int32")));

            AssemblyDT = new DataTable();
            AssemblyDT.Columns.Add(new DataColumn("Name", Type.GetType("System.String")));
            AssemblyDT.Columns.Add(new DataColumn("FrameColor", Type.GetType("System.String")));
            AssemblyDT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            AssemblyDT.Columns.Add(new DataColumn("Width", Type.GetType("System.Int32")));
            AssemblyDT.Columns.Add(new DataColumn("Count", Type.GetType("System.Int32")));
            AssemblyDT.Columns.Add(new DataColumn("Square", Type.GetType("System.Decimal")));
            AssemblyDT.Columns.Add(new DataColumn("Notes", Type.GetType("System.String")));
            AssemblyDT.Columns.Add(new DataColumn("ColorID", Type.GetType("System.Int32")));

            RoughCutDT = new DataTable();
            RoughCutDT.Columns.Add(new DataColumn("Name", Type.GetType("System.String")));
            RoughCutDT.Columns.Add(new DataColumn("FrameColor", Type.GetType("System.String")));
            RoughCutDT.Columns.Add(new DataColumn("Height1", Type.GetType("System.Int32")));
            RoughCutDT.Columns.Add(new DataColumn("Width1", Type.GetType("System.Int32")));
            RoughCutDT.Columns.Add(new DataColumn("Count1", Type.GetType("System.Int32")));
            RoughCutDT.Columns.Add(new DataColumn("Square1", Type.GetType("System.Decimal")));
            RoughCutDT.Columns.Add(new DataColumn("Height2", Type.GetType("System.Int32")));
            RoughCutDT.Columns.Add(new DataColumn("Width2", Type.GetType("System.Int32")));
            RoughCutDT.Columns.Add(new DataColumn("Count2", Type.GetType("System.Int32")));
            RoughCutDT.Columns.Add(new DataColumn("Square2", Type.GetType("System.Decimal")));
            RoughCutDT.Columns.Add(new DataColumn("Notes", Type.GetType("System.String")));
            RoughCutDT.Columns.Add(new DataColumn("ColorID", Type.GetType("System.Int32")));
        }

        private DataTable DistGroupNumbersTable(DataRow[] Rows, bool OrderASC)
        {
            int GroupNumber = 0;
            DataTable DT = new DataTable();
            DT.Columns.Add(new DataColumn("GroupNumber", Type.GetType("System.Int32")));
            foreach (DataRow Row in Rows)
            {
                if (int.TryParse(Row["GroupNumber"].ToString(), out GroupNumber))
                {
                    DataRow NewRow = DT.NewRow();
                    NewRow["GroupNumber"] = GroupNumber;
                    DT.Rows.Add(NewRow);
                }
            }
            using (DataView DV = new DataView(DT.Copy()))
            {
                DT.Clear();
                if (OrderASC)
                    DV.Sort = "GroupNumber ASC";
                else
                    DV.Sort = "GroupNumber DESC";
                DT = DV.ToTable(true, new string[] { "GroupNumber" });
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

        private void Fill()
        {
            FrontsDataTable = new DataTable();
            PatinaDataTable = new DataTable();
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
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT PatinaRAL.*, Patina.Patina FROM PatinaRAL INNER JOIN Patina ON Patina.PatinaID=PatinaRAL.PatinaID WHERE PatinaRAL.Enabled=1",
                ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(PatinaRALDataTable);
            }
            foreach (DataRow item in PatinaRALDataTable.Rows)
            {
                DataRow NewRow = PatinaDataTable.NewRow();
                NewRow["PatinaID"] = item["PatinaRALID"];
                NewRow["PatinaName"] = item["PatinaRAL"]; NewRow["Patina"] = item["Patina"];
                NewRow["DisplayName"] = item["DisplayName"];
                PatinaDataTable.Rows.Add(NewRow);
            }
            DecorDT = new DataTable();
            SelectCommand = @"SELECT DISTINCT TechStore.TechStoreID AS DecorID, TechStore.TechStoreName AS Name, DecorConfig.ProductID FROM TechStore
                INNER JOIN DecorConfig ON TechStore.TechStoreID = DecorConfig.DecorID ORDER BY TechStoreName";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(DecorDT);
            }
            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT * FROM DecorParameters",
                ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(DecorParametersDT);
            }
            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT TOP 0 FrontsOrdersID, MainOrderID,
                FrontID, ColorID, PatinaID, InsetTypeID, InsetColorID, TechnoColorID, TechnoInsetTypeID, TechnoInsetColorID,
                Height, Width, Count, FrontConfigID, Square, FactoryID, IsNonStandard, Notes FROM FrontsOrders",
                ConnectionStrings.ZOVOrdersConnectionString))
            {
                DA.Fill(NotCurvedOrdersDT);
                NotCurvedOrdersDT.Columns.Add(new DataColumn("GroupType", Type.GetType("System.Int32")));
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

        private void GetArchDecorOrders(ref DataTable DestinationDT, int FactoryID)
        {
            string SelectCommand = string.Empty;
            DataTable DT = new DataTable();

            SelectCommand = @"SELECT DecorOrderID, MainOrderID, ProductID, DecorID, DecorConfigID, ColorID, PatinaID,
                Height, Length, Width, Count, Notes FROM DecorOrders
                WHERE FactoryID=" + FactoryID + " AND ProductID IN (31, 4, 18, 32) AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE ProfilWorkAssignmentID=" + WorkAssignmentID + "))";
            if (FactoryID == 2)
                SelectCommand = @"SELECT DecorOrderID, MainOrderID, ProductID, DecorID, DecorConfigID, ColorID,  PatinaID,
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

        private void GetGridsDecorOrders(ref DataTable DestinationDT, int FactoryID)
        {
            string SelectCommand = string.Empty;
            DataTable DT = new DataTable();

            SelectCommand = @"SELECT DecorOrderID, MainOrderID, ProductID, DecorID, DecorConfigID, ColorID,  PatinaID,
                Height, Length, Width, Count, Notes FROM DecorOrders
                WHERE FactoryID=" + FactoryID + " AND ProductID IN (10, 11, 12) AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE ProfilWorkAssignmentID=" + WorkAssignmentID + "))";
            if (FactoryID == 2)
                SelectCommand = @"SELECT DecorOrderID, MainOrderID, ProductID, DecorID, DecorConfigID, ColorID,  PatinaID,
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

        private void GetNotArchDecorOrders(ref DataTable DestinationDT, int FactoryID)
        {
            string SelectCommand = string.Empty;
            DataTable DT = new DataTable();

            SelectCommand = @"SELECT DecorOrderID, MainOrderID, ProductID, DecorID, DecorConfigID, ColorID,  PatinaID,
                Height, Length, Width, Count, Notes FROM DecorOrders
                WHERE FactoryID=" + FactoryID + " AND ProductID NOT IN (31, 4, 18, 32, 10, 11, 12) AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE ProfilWorkAssignmentID=" + WorkAssignmentID + "))";
            if (FactoryID == 2)
                SelectCommand = @"SELECT DecorOrderID, MainOrderID, ProductID, DecorID, DecorConfigID, ColorID,  PatinaID,
                    Height, Length, Width, Count, Notes FROM DecorOrders
                    WHERE FactoryID=" + FactoryID + " AND ProductID NOT IN (31, 4, 18, 32, 10, 11, 12) AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE TPSWorkAssignmentID=" + WorkAssignmentID + "))";

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

        private void GetNotCurvedFrontsOrders(ref DataTable DestinationDT, int FactoryID, Fronts Front)
        {
            string SelectCommand = string.Empty;
            DataTable DT = new DataTable();

            SelectCommand = @"SELECT FrontsOrdersID, MainOrderID,
                FrontID, ColorID, PatinaID, InsetTypeID, InsetColorID, TechnoColorID, TechnoInsetTypeID, TechnoInsetColorID,
                Height, Width, Count, FrontConfigID, Square, FactoryID, IsNonStandard, Notes FROM FrontsOrders
                WHERE Width<>-1 AND FrontID=" + Convert.ToInt32(Front) +
                " AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE ProfilWorkAssignmentID=" + WorkAssignmentID + "))";
            if (FactoryID == 2)
                SelectCommand = @"SELECT FrontsOrdersID, MainOrderID,
                FrontID, ColorID, PatinaID, InsetTypeID, InsetColorID, TechnoColorID, TechnoInsetTypeID, TechnoInsetColorID,
                Height, Width, Count, FrontConfigID, Square, FactoryID, IsNonStandard, Notes FROM FrontsOrders
                    WHERE Width<>-1 AND FrontID=" + Convert.ToInt32(Front) +
                    " AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE TPSWorkAssignmentID=" + WorkAssignmentID + "))";
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

        private void GridsDecorAssemblyByMainOrderToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            string BatchName)
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

                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=0"))
                {
                    DecorAssemblyDT.Clear();
                    DT.Clear();
                    int MainOrderID = Convert.ToInt32(item["MainOrderID"]);
                    DataRow[] rows = GridsDecorOrdersDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item1 in rows)
                        DT.Rows.Add(item1.ItemArray);
                    GroupAndSortDecor(DT, ref DecorAssemblyDT);

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
                             BatchName, ClientName, string.Empty, OrderName, Notes, ref RowIndex);
                        RowIndex++;
                        GridsDecorAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                            Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DecorAssemblyDT,
                            BatchName, ClientName, "ДУБЛЬ", OrderName, Notes, ref RowIndex);
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

                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=1"))
                {
                    DecorAssemblyDT.Clear();
                    DT.Clear();
                    int MainOrderID = Convert.ToInt32(item["MainOrderID"]);
                    DataRow[] rows = GridsDecorOrdersDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item1 in rows)
                        DT.Rows.Add(item1.ItemArray);
                    GroupAndSortDecor(DT, ref DecorAssemblyDT);

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
                             BatchName, ClientName, string.Empty, OrderName, Notes, ref RowIndex);
                        RowIndex++;
                        GridsDecorAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                            Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DecorAssemblyDT,
                            BatchName, ClientName, "ДУБЛЬ", OrderName, Notes, ref RowIndex);
                    }
                    RowIndex++;
                }
            }
        }

        private void GroupAndSortColoredFronts(DataTable SourceDT, ref DataTable DestinationDT)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            int H = 0;
            int W = 0;
            string N = string.Empty;
            //ШП КР Дуб Венге
            using (DataView DV = new DataView(SourceDT, @"(ColorID = 1890 AND PatinaID = 6) OR (ColorID = 1890 AND PatinaID = 7)", string.Empty, DataViewRowState.CurrentRows))
            {
                DT1 = DV.ToTable(true, new string[] { "FrontID", "ColorID", "PatinaID" });
            }
            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                H = 0;
                W = 0;
                N = string.Empty;
                int FrontID = Convert.ToInt32(DT1.Rows[i]["FrontID"]);
                int ColorID = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                int PatinaID = Convert.ToInt32(DT1.Rows[i]["PatinaID"]);
                using (DataView DV = new DataView(SourceDT, "ColorID=" + ColorID + " AND FrontID=" + FrontID + " AND PatinaID=" + PatinaID,
                    "InsetTypeID, InsetColorID, Height, Width", DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "Height", "Width", "Notes" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    if (Convert.ToInt32(DT2.Rows[j]["Height"]) == H && Convert.ToInt32(DT2.Rows[j]["Width"]) == W &&
                        DT2.Rows[j]["Notes"].ToString() == N)
                        continue;

                    H = Convert.ToInt32(DT2.Rows[j]["Height"]);
                    W = Convert.ToInt32(DT2.Rows[j]["Width"]);
                    N = DT2.Rows[j]["Notes"].ToString();

                    decimal Square = 0;
                    int Count = 0;

                    string filter = "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) + " AND FrontID=" + Convert.ToInt32(DT1.Rows[i]["FrontID"]) + " AND PatinaID=" + Convert.ToInt32(DT1.Rows[i]["PatinaID"]) +
                        " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                        " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) + " AND (Notes='' OR Notes IS NULL)";
                    if (DT2.Rows[j]["Notes"] != DBNull.Value && DT2.Rows[j]["Notes"].ToString().Length > 0)
                        filter = "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) + " AND FrontID=" + Convert.ToInt32(DT1.Rows[i]["FrontID"]) + " AND PatinaID=" + Convert.ToInt32(DT1.Rows[i]["PatinaID"]) +
                        " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                        " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) +
                        " AND Notes='" + DT2.Rows[j]["Notes"].ToString() + "'";
                    DataRow[] rows = SourceDT.Select(filter);
                    foreach (DataRow item in rows)
                    {
                        Count += Convert.ToInt32(item["Count"]);
                        Square += Convert.ToDecimal(item["Height"]) * Convert.ToDecimal(item["Width"]) * Convert.ToDecimal(item["Count"]) / 1000000;
                    }
                    Square = decimal.Round(Square, 3, MidpointRounding.AwayFromZero);

                    DataRow NewRow = DestinationDT.NewRow();
                    NewRow["Name"] = GetFrontName(Convert.ToInt32(DT1.Rows[i]["FrontID"]));
                    if (Convert.ToInt32(DT1.Rows[i]["PatinaID"]) != -1)
                        NewRow["FrameColor"] = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"])) + " " + GetPatinaName(Convert.ToInt32(DT1.Rows[i]["PatinaID"]));
                    else
                        NewRow["FrameColor"] = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                    NewRow["Height"] = Convert.ToInt32(DT2.Rows[j]["Height"]);
                    NewRow["Width"] = Convert.ToInt32(DT2.Rows[j]["Width"]);
                    NewRow["Count"] = Count;
                    NewRow["Square"] = Square;
                    NewRow["Notes"] = DT2.Rows[j]["Notes"];
                    NewRow["ColorID"] = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                    DestinationDT.Rows.Add(NewRow);
                }
            }
            using (DataView DV = new DataView(DestinationDT.Copy()))
            {
                DV.Sort = "Name, FrameColor, Height, Width";
                DestinationDT.Clear();
                DestinationDT = DV.ToTable();
            }
        }

        private void GroupAndSortDecor(DataTable SourceDT, ref DataTable DestinationDT)
        {
            DataTable DT1 = new DataTable();

            using (DataView DV = new DataView(SourceDT))
            {
                DT1 = DV.ToTable(true, new string[] { "DecorID", "ColorID", "PatinaID", "Length", "Height", "Width", "Notes" });
            }
            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                int Count = 0;
                string filter = "DecorID=" + Convert.ToInt32(DT1.Rows[i]["DecorID"]) + " AND ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) + " AND PatinaID=" + Convert.ToInt32(DT1.Rows[i]["PatinaID"]) +
                    " AND Length=" + Convert.ToInt32(DT1.Rows[i]["Length"]) + " AND Height=" + Convert.ToInt32(DT1.Rows[i]["Height"]) +
                    " AND Width=" + Convert.ToInt32(DT1.Rows[i]["Width"]) + " AND (Notes='' OR Notes IS NULL)";
                if (DT1.Rows[i]["Notes"] != DBNull.Value && DT1.Rows[i]["Notes"].ToString().Length > 0)
                {
                    filter = "DecorID=" + Convert.ToInt32(DT1.Rows[i]["DecorID"]) + " AND ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) + " AND PatinaID=" + Convert.ToInt32(DT1.Rows[i]["PatinaID"]) +
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
                //if (HasParameter(Convert.ToInt32(rows[0]["ProductID"]), "ColorID"))
                NewRow["Color"] = Color;
                //if (HasParameter(Convert.ToInt32(rows[0]["ProductID"]), "Height"))
                NewRow["Height"] = DT1.Rows[i]["Height"];
                //if (HasParameter(Convert.ToInt32(rows[0]["ProductID"]), "Length"))
                NewRow["Length"] = DT1.Rows[i]["Length"];
                //if (HasParameter(Convert.ToInt32(rows[0]["ProductID"]), "Width"))
                NewRow["Width"] = DT1.Rows[i]["Width"];
                NewRow["Count"] = Count;
                NewRow["Notes"] = DT1.Rows[i]["Notes"];
                DestinationDT.Rows.Add(NewRow);
            }
            using (DataView DV = new DataView(DestinationDT.Copy()))
            {
                DV.Sort = "Name, Color, Patina, Length, Height, Width";
                DestinationDT.Clear();
                DestinationDT = DV.ToTable();
            }
        }

        private void GroupAndSortFronts(DataTable SourceDT, ref DataTable DestinationDT)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            int H = 0;
            int W = 0;
            string N = string.Empty;

            using (DataView DV = new DataView(SourceDT, string.Empty, "FrontID, ColorID, PatinaID", DataViewRowState.CurrentRows))
            {
                DT1 = DV.ToTable(true, new string[] { "FrontID", "ColorID", "PatinaID" });
            }
            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                H = 0;
                W = 0;
                N = string.Empty;
                int FrontID = Convert.ToInt32(DT1.Rows[i]["FrontID"]);
                int ColorID = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                int PatinaID = Convert.ToInt32(DT1.Rows[i]["PatinaID"]);
                using (DataView DV = new DataView(SourceDT, "ColorID=" + ColorID + " AND FrontID=" + FrontID + " AND PatinaID=" + PatinaID,
                    "Height, Width, Notes", DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "Height", "Width", "Notes" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    if (Convert.ToInt32(DT2.Rows[j]["Height"]) == H && Convert.ToInt32(DT2.Rows[j]["Width"]) == W &&
                        DT2.Rows[j]["Notes"].ToString() == N)
                        continue;

                    H = Convert.ToInt32(DT2.Rows[j]["Height"]);
                    W = Convert.ToInt32(DT2.Rows[j]["Width"]);
                    N = DT2.Rows[j]["Notes"].ToString();

                    decimal Square = 0;
                    int Count = 0;
                    string filter = "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) + " AND FrontID=" + Convert.ToInt32(DT1.Rows[i]["FrontID"]) + " AND PatinaID=" + Convert.ToInt32(DT1.Rows[i]["PatinaID"]) +
                        " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                        " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) + " AND (Notes='' OR Notes IS NULL)";
                    if (DT2.Rows[j]["Notes"] != DBNull.Value && DT2.Rows[j]["Notes"].ToString().Length > 0)
                        filter = "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) + " AND FrontID=" + Convert.ToInt32(DT1.Rows[i]["FrontID"]) + " AND PatinaID=" + Convert.ToInt32(DT1.Rows[i]["PatinaID"]) +
                        " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                        " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) +
                        " AND Notes='" + DT2.Rows[j]["Notes"].ToString() + "'";

                    DataRow[] rows = SourceDT.Select(filter);
                    foreach (DataRow item in rows)
                    {
                        Count += Convert.ToInt32(item["Count"]);
                        Square += Convert.ToDecimal(item["Height"]) * Convert.ToDecimal(item["Width"]) * Convert.ToDecimal(item["Count"]) / 1000000;
                    }
                    Square = decimal.Round(Square, 3, MidpointRounding.AwayFromZero);

                    DataRow NewRow = DestinationDT.NewRow();
                    NewRow["Name"] = GetFrontName(Convert.ToInt32(DT1.Rows[i]["FrontID"]));
                    if (Convert.ToInt32(DT1.Rows[i]["PatinaID"]) != -1)
                        NewRow["FrameColor"] = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"])) + " " + GetPatinaName(Convert.ToInt32(DT1.Rows[i]["PatinaID"]));
                    else
                        NewRow["FrameColor"] = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                    NewRow["Height"] = Convert.ToInt32(DT2.Rows[j]["Height"]);
                    NewRow["Width"] = Convert.ToInt32(DT2.Rows[j]["Width"]);
                    NewRow["Count"] = Count;
                    NewRow["Square"] = Square;
                    NewRow["Notes"] = DT2.Rows[j]["Notes"];
                    NewRow["ColorID"] = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                    DestinationDT.Rows.Add(NewRow);
                }
            }
            using (DataView DV = new DataView(DestinationDT.Copy()))
            {
                DV.Sort = "Name, FrameColor, Height, Width";
                DestinationDT.Clear();
                DestinationDT = DV.ToTable();
            }
        }

        private void GroupAndSortFrontsWithHands(DataTable SourceDT, ref DataTable DestinationDT)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();

            using (DataView DV = new DataView(SourceDT, "Notes='укв-7'", "ColorID", DataViewRowState.CurrentRows))
            {
                DT1 = DV.ToTable(true, new string[] { "FrontID", "ColorID", "PatinaID" });
            }
            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                int FrontID = Convert.ToInt32(DT1.Rows[i]["FrontID"]);
                int ColorID = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                int PatinaID = Convert.ToInt32(DT1.Rows[i]["PatinaID"]);
                using (DataView DV = new DataView(SourceDT, "ColorID=" + ColorID + " AND FrontID=" + FrontID + " AND PatinaID=" + PatinaID,
                    "InsetTypeID, InsetColorID, Height, Width", DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "InsetTypeID", "InsetColorID", "Height", "Width" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    decimal Square = 0;
                    int Count = 0;

                    DataRow[] rows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) + " AND FrontID=" + Convert.ToInt32(DT1.Rows[i]["FrontID"]) + " AND PatinaID=" + Convert.ToInt32(DT1.Rows[i]["PatinaID"]) +
                        " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) +
                        " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) +
                        " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                        " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) + " AND Notes='укв-7'");
                    if (rows.Count() == 0)
                        continue;
                    foreach (DataRow item in rows)
                    {
                        Count += Convert.ToInt32(item["Count"]);
                        Square += Convert.ToDecimal(item["Height"]) * Convert.ToDecimal(item["Width"]) * Convert.ToDecimal(item["Count"]) / 1000000;
                    }
                    Square = decimal.Round(Square, 3, MidpointRounding.AwayFromZero);

                    DataRow NewRow = DestinationDT.NewRow();
                    NewRow["Name"] = GetFrontName(Convert.ToInt32(DT1.Rows[i]["FrontID"]));
                    if (Convert.ToInt32(DT1.Rows[i]["PatinaID"]) != -1)
                        NewRow["FrameColor"] = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"])) + " " + GetPatinaName(Convert.ToInt32(DT1.Rows[i]["PatinaID"]));
                    else
                        NewRow["FrameColor"] = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                    NewRow["Height"] = Convert.ToInt32(DT2.Rows[j]["Height"]);
                    NewRow["Width"] = Convert.ToInt32(DT2.Rows[j]["Width"]);
                    NewRow["Count"] = Count;
                    NewRow["Square"] = Square;
                    NewRow["Notes"] = rows[0]["Notes"];
                    NewRow["ColorID"] = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                    DestinationDT.Rows.Add(NewRow);
                }
            }
            using (DataView DV = new DataView(DestinationDT.Copy()))
            {
                DV.Sort = "Name, FrameColor, Height, Width";
                DestinationDT.Clear();
                DestinationDT = DV.ToTable();
            }
        }

        private void GroupAndSortGlossFronts(DataTable SourceDT, ref DataTable DestinationDT)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            int H = 0;
            int W = 0;
            string N = string.Empty;
            //ШП Американо и т.д.
            using (DataView DV = new DataView(SourceDT, @"FrontID=3663 AND ColorID IN (3694,1881,3695,3696,3697,3698,1893,1890,1884,3699,1882,1885,3700,3701,1891,1886,3702,1889,3703,1894,1892,3704,3705)", string.Empty, DataViewRowState.CurrentRows))
            {
                DT1 = DV.ToTable(true, new string[] { "ColorID" });
            }
            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                H = 0;
                W = 0;
                N = string.Empty;
                using (DataView DV = new DataView(SourceDT, "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]),
                    "Height, Width, Notes", DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "Height", "Width", "Notes" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    if (Convert.ToInt32(DT2.Rows[j]["Height"]) == H && Convert.ToInt32(DT2.Rows[j]["Width"]) == W &&
                        DT2.Rows[j]["Notes"].ToString() == N)
                        continue;

                    H = Convert.ToInt32(DT2.Rows[j]["Height"]);
                    W = Convert.ToInt32(DT2.Rows[j]["Width"]);
                    N = DT2.Rows[j]["Notes"].ToString();

                    decimal Square = 0;
                    int Count = 0;

                    string filter = "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                        " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) + " AND (Notes='' OR Notes IS NULL)";
                    if (DT2.Rows[j]["Notes"] != DBNull.Value && DT2.Rows[j]["Notes"].ToString().Length > 0)
                        filter = "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                        " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) +
                        " AND Notes='" + DT2.Rows[j]["Notes"].ToString() + "'";
                    DataRow[] rows = SourceDT.Select(filter);
                    foreach (DataRow item in rows)
                    {
                        Count += Convert.ToInt32(item["Count"]);
                        Square += Convert.ToDecimal(item["Height"]) * Convert.ToDecimal(item["Width"]) * Convert.ToDecimal(item["Count"]) / 1000000;
                    }
                    Square = decimal.Round(Square, 3, MidpointRounding.AwayFromZero);

                    DataRow NewRow = DestinationDT.NewRow();
                    NewRow["Name"] = GetFrontName(3663);
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                    NewRow["Height"] = Convert.ToInt32(DT2.Rows[j]["Height"]);
                    NewRow["Width"] = Convert.ToInt32(DT2.Rows[j]["Width"]);
                    NewRow["Count"] = Count;
                    NewRow["Square"] = Square;
                    NewRow["Notes"] = DT2.Rows[j]["Notes"];
                    NewRow["ColorID"] = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                    DestinationDT.Rows.Add(NewRow);
                }
            }
            using (DataView DV = new DataView(DestinationDT.Copy()))
            {
                DV.Sort = "Name, FrameColor, Height, Width";
                DestinationDT.Clear();
                DestinationDT = DV.ToTable();
            }
        }

        private void GroupAndSortNotGlossFronts(DataTable SourceDT, ref DataTable DestinationDT)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            int H = 0;
            int W = 0;
            string N = string.Empty;
            //ГЛЯНЕЦ
            using (DataView DV = new DataView(SourceDT, @"FrontID=3664 AND ColorID IN (1881,1893,1883,1884,1882,1885,1886,1889,1894)", string.Empty, DataViewRowState.CurrentRows))
            {
                DT1 = DV.ToTable(true, new string[] { "ColorID" });
            }
            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                H = 0;
                W = 0;
                N = string.Empty;
                using (DataView DV = new DataView(SourceDT, "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]),
                    "Height, Width, Notes", DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "Height", "Width", "Notes" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    if (Convert.ToInt32(DT2.Rows[j]["Height"]) == H && Convert.ToInt32(DT2.Rows[j]["Width"]) == W &&
                        DT2.Rows[j]["Notes"].ToString() == N)
                        continue;

                    H = Convert.ToInt32(DT2.Rows[j]["Height"]);
                    W = Convert.ToInt32(DT2.Rows[j]["Width"]);
                    N = DT2.Rows[j]["Notes"].ToString();

                    decimal Square = 0;
                    int Count = 0;

                    string filter = "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                        " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) + " AND (Notes='' OR Notes IS NULL)";
                    if (DT2.Rows[j]["Notes"] != DBNull.Value && DT2.Rows[j]["Notes"].ToString().Length > 0)
                        filter = "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                        " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) +
                        " AND Notes='" + DT2.Rows[j]["Notes"].ToString() + "'";
                    DataRow[] rows = SourceDT.Select(filter);
                    foreach (DataRow item in rows)
                    {
                        Count += Convert.ToInt32(item["Count"]);
                        Square += Convert.ToDecimal(item["Height"]) * Convert.ToDecimal(item["Width"]) * Convert.ToDecimal(item["Count"]) / 1000000;
                    }
                    Square = decimal.Round(Square, 3, MidpointRounding.AwayFromZero);

                    DataRow NewRow = DestinationDT.NewRow();
                    NewRow["Name"] = GetFrontName(3664);
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                    NewRow["Height"] = Convert.ToInt32(DT2.Rows[j]["Height"]);
                    NewRow["Width"] = Convert.ToInt32(DT2.Rows[j]["Width"]);
                    NewRow["Count"] = Count;
                    NewRow["Square"] = Square;
                    NewRow["Notes"] = DT2.Rows[j]["Notes"];
                    NewRow["ColorID"] = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                    DestinationDT.Rows.Add(NewRow);
                }
            }
            using (DataView DV = new DataView(DestinationDT.Copy()))
            {
                DV.Sort = "Name, FrameColor, Height, Width";
                DestinationDT.Clear();
                DestinationDT = DV.ToTable();
            }
        }

        private void NotArchDecorAssemblyByMainOrderToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            string BatchName)
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

                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=0"))
                {
                    DecorAssemblyDT.Clear();
                    DT.Clear();
                    int MainOrderID = Convert.ToInt32(item["MainOrderID"]);
                    DataRow[] rows = NotArchDecorOrdersDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item1 in rows)
                        DT.Rows.Add(item1.ItemArray);
                    GroupAndSortDecor(DT, ref DecorAssemblyDT);

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
                             BatchName, ClientName, string.Empty, OrderName, Notes, ref RowIndex);
                        RowIndex++;
                        NotArchDecorAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                            Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DecorAssemblyDT,
                            BatchName, ClientName, "ДУБЛЬ", OrderName, Notes, ref RowIndex);
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

                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=1"))
                {
                    DecorAssemblyDT.Clear();
                    DT.Clear();
                    int MainOrderID = Convert.ToInt32(item["MainOrderID"]);
                    DataRow[] rows = NotArchDecorOrdersDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item1 in rows)
                        DT.Rows.Add(item1.ItemArray);
                    GroupAndSortDecor(DT, ref DecorAssemblyDT);

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
                             BatchName, ClientName, string.Empty, OrderName, Notes, ref RowIndex);
                        RowIndex++;
                        NotArchDecorAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                            Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DecorAssemblyDT,
                            BatchName, ClientName, "ДУБЛЬ", OrderName, Notes, ref RowIndex);
                    }
                    RowIndex++;
                }
            }
        }

        private void RoughCutCoolect(DataTable SourceDT, ref DataTable DestinationDT)
        {
            int Admission = 5;
            int HandsAdmission = 0;
            int Deduction = 0;
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            int H = 0;
            int W = 0;
            int TotalHeight = 0;
            int TotalCount = 1;
            int TotalWidth = 0;
            decimal Square1 = 0;
            string N = string.Empty;

            using (DataView DV = new DataView(SourceDT, string.Empty, "FrontID, ColorID, PatinaID", DataViewRowState.CurrentRows))
            {
                DT1 = DV.ToTable(true, new string[] { "FrontID", "ColorID", "PatinaID" });
            }
            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                H = 0;
                W = 0;
                N = string.Empty;
                int FrontID = Convert.ToInt32(DT1.Rows[i]["FrontID"]);
                int ColorID = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                int PatinaID = Convert.ToInt32(DT1.Rows[i]["PatinaID"]);

                using (DataView DV = new DataView(SourceDT, "ColorID=" + ColorID + " AND FrontID=" + FrontID + " AND PatinaID=" + PatinaID,
                    "Height, Width, Notes", DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "Height", "Width", "Notes" });
                }
                bool FirstRow = true;
                Square1 = 0;
                Admission = (DT2.Rows.Count - 1) * 5;
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    int Count = 0;
                    HandsAdmission = 0;

                    string filter = "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) + " AND FrontID=" + Convert.ToInt32(DT1.Rows[i]["FrontID"]) + " AND PatinaID=" + Convert.ToInt32(DT1.Rows[i]["PatinaID"]) +
                           " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                           " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) + " AND (Notes='' OR Notes IS NULL OR Notes<>'УКВ-7')";
                    if (DT2.Rows[j]["Notes"] != DBNull.Value && DT2.Rows[j]["Notes"].ToString() == "УКВ-7")
                    {
                        filter = "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) + " AND FrontID=" + Convert.ToInt32(DT1.Rows[i]["FrontID"]) + " AND PatinaID=" + Convert.ToInt32(DT1.Rows[i]["PatinaID"]) +
                            " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                            " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) +
                            " AND Notes='УКВ-7'";
                        Admission += 20;
                        HandsAdmission = -36;
                    }
                    DataRow[] rows = SourceDT.Select(filter);
                    foreach (DataRow item in rows)
                    {
                        Count += Convert.ToInt32(item["Count"]);
                    }
                    Admission += 5 * (Count - 1);

                    if (Convert.ToInt32(DT2.Rows[j]["Width"]) * Count + Admission + 10 > 735)
                    {
                        if (Convert.ToInt32(DT2.Rows[j]["Height"]) * Count + Admission + 10 > 735)
                        {
                            TotalCount = Count;
                            TotalHeight = Convert.ToInt32(DT2.Rows[j]["Height"]) + HandsAdmission;
                            TotalWidth = Convert.ToInt32(DT2.Rows[j]["Width"]);
                        }
                        else
                        {
                            TotalWidth = Convert.ToInt32(DT2.Rows[j]["Width"]) + HandsAdmission;
                            TotalHeight += Convert.ToInt32(DT2.Rows[j]["Height"]) * Count;
                        }
                    }
                    else
                    {
                        TotalHeight = Convert.ToInt32(DT2.Rows[j]["Height"]) + HandsAdmission;
                        TotalWidth += Convert.ToInt32(DT2.Rows[j]["Width"]) * Count;
                    }
                }
                TotalWidth = TotalWidth + Admission;
                Square1 += Convert.ToDecimal(TotalHeight) * Convert.ToDecimal(TotalWidth) * TotalCount / 1000000;
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    if (Convert.ToInt32(DT2.Rows[j]["Height"]) == H && Convert.ToInt32(DT2.Rows[j]["Width"]) == W &&
                        DT2.Rows[j]["Notes"].ToString() == N)
                        continue;

                    H = Convert.ToInt32(DT2.Rows[j]["Height"]);
                    W = Convert.ToInt32(DT2.Rows[j]["Width"]);
                    N = DT2.Rows[j]["Notes"].ToString();

                    decimal Square2 = 0;
                    int Count = 0;

                    if (Convert.ToInt32(DT2.Rows[j]["Height"]) >= 40 && Convert.ToInt32(DT2.Rows[j]["Width"]) >= 40)
                        Deduction = 0;
                    else
                    {
                        if (ColorID == 1881 || ColorID == 1893 || ColorID == 1883 || ColorID == 1884 || ColorID == 1882 || ColorID == 1885 ||
                            ColorID == 1886 || ColorID == 1889 || ColorID == 1894 || (ColorID == 1890 && PatinaID == 6) || (ColorID == 1890 && PatinaID == 7))
                            Deduction = 2;

                        if (ColorID == 3694 || ColorID == 3695 || ColorID == 3696 || ColorID == 3697 || ColorID == 3698 || ColorID == 1890 ||
                            ColorID == 3699 || ColorID == 3700 || ColorID == 3701 || ColorID == 1891 || ColorID == 3702 || ColorID == 3703 ||
                            ColorID == 1892 || ColorID == 3704 || ColorID == 3705)
                            Deduction = 1;
                    }

                    string filter = "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) + " AND FrontID=" + Convert.ToInt32(DT1.Rows[i]["FrontID"]) + " AND PatinaID=" + Convert.ToInt32(DT1.Rows[i]["PatinaID"]) +
                        " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                        " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) + " AND (Notes='' OR Notes IS NULL OR Notes<>'УКВ-7')";
                    if (DT2.Rows[j]["Notes"] != DBNull.Value && DT2.Rows[j]["Notes"].ToString() == "УКВ-7")
                    {
                        filter = "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) + " AND FrontID=" + Convert.ToInt32(DT1.Rows[i]["FrontID"]) + " AND PatinaID=" + Convert.ToInt32(DT1.Rows[i]["PatinaID"]) +
                            " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                            " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) +
                            " AND Notes='УКВ-7'";
                    }
                    DataRow[] rows = SourceDT.Select(filter);
                    foreach (DataRow item in rows)
                    {
                        Count += Convert.ToInt32(item["Count"]);
                        Square2 += Convert.ToDecimal(item["Height"]) * (Convert.ToDecimal(item["Width"]) - Deduction) * Convert.ToDecimal(item["Count"]) / 1000000;
                    }
                    Square2 = decimal.Round(Square2, 3, MidpointRounding.AwayFromZero);

                    DataRow NewRow = DestinationDT.NewRow();
                    NewRow["Name"] = GetFrontName(Convert.ToInt32(DT1.Rows[i]["FrontID"]));
                    if (Convert.ToInt32(DT1.Rows[i]["PatinaID"]) != -1)
                        NewRow["FrameColor"] = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"])) + " " + GetPatinaName(Convert.ToInt32(DT1.Rows[i]["PatinaID"]));
                    else
                        NewRow["FrameColor"] = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                    if (FirstRow)
                    {
                        Square1 = decimal.Round(Square1, 3, MidpointRounding.AwayFromZero);
                        NewRow["Height1"] = TotalWidth + 10;
                        NewRow["Width1"] = TotalHeight + 10;
                        NewRow["Count1"] = TotalCount;
                        NewRow["Square1"] = Square1;
                    }
                    NewRow["Height2"] = Convert.ToInt32(DT2.Rows[j]["Width"]) - Deduction;
                    NewRow["Width2"] = Convert.ToInt32(DT2.Rows[j]["Height"]);
                    NewRow["Count2"] = Count;
                    NewRow["Square2"] = Square2;
                    NewRow["Notes"] = DT2.Rows[j]["Notes"];
                    NewRow["ColorID"] = ColorID;
                    DestinationDT.Rows.Add(NewRow);
                    FirstRow = false;
                }
            }
        }

        private void RoughCutCoolectByHeight(DataTable SourceDT, ref DataTable DestinationDT)
        {
            int Admission = 5;
            int HandsAdmission = 0;
            int Deduction = 0;
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            int H = 0;
            int W = 0;
            int TotalHeight = 0;
            int TotalCount2 = 0;
            int TotalCount1 = 1;
            int TotalWidth = 0;
            decimal Square1 = 0;
            string N = string.Empty;

            using (DataView DV = new DataView(SourceDT, string.Empty, "FrontID, ColorID, PatinaID", DataViewRowState.CurrentRows))
            {
                DT1 = DV.ToTable(true, new string[] { "FrontID", "ColorID", "PatinaID" });
            }
            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                H = 0;
                W = 0;
                N = string.Empty;
                int FrontID = Convert.ToInt32(DT1.Rows[i]["FrontID"]);
                int ColorID = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                int PatinaID = Convert.ToInt32(DT1.Rows[i]["PatinaID"]);

                using (DataView DV = new DataView(SourceDT, "ColorID=" + ColorID + " AND FrontID=" + FrontID + " AND PatinaID=" + PatinaID,
                    "Height, Width, Notes", DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "Height", "Width", "Notes" });
                }
                bool FirstRow = true;
                Square1 = 0;
                Admission = 5;
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    int Count = 0;
                    HandsAdmission = 0;

                    string filter = "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) + " AND FrontID=" + Convert.ToInt32(DT1.Rows[i]["FrontID"]) + " AND PatinaID=" + Convert.ToInt32(DT1.Rows[i]["PatinaID"]) +
                           " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                           " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) + " AND (Notes='' OR Notes IS NULL OR Notes<>'УКВ-7')";
                    if (DT2.Rows[j]["Notes"] != DBNull.Value && DT2.Rows[j]["Notes"].ToString() == "УКВ-7")
                    {
                        filter = "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) + " AND FrontID=" + Convert.ToInt32(DT1.Rows[i]["FrontID"]) + " AND PatinaID=" + Convert.ToInt32(DT1.Rows[i]["PatinaID"]) +
                            " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                            " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) +
                            " AND Notes='УКВ-7'";
                        Admission += 20;
                        HandsAdmission = -36;
                    }
                    DataRow[] rows = SourceDT.Select(filter);
                    foreach (DataRow item in rows)
                    {
                        Count += Convert.ToInt32(item["Count"]);
                    }

                    TotalWidth += Convert.ToInt32(DT2.Rows[j]["Width"]) * Count;
                    TotalHeight = Convert.ToInt32(DT2.Rows[j]["Height"]) + HandsAdmission;
                    TotalCount2 += Count;
                }
                if (TotalWidth + (TotalCount2 - 1) * Admission + 10 > 735)
                    TotalCount1 = TotalCount2;
                TotalWidth = TotalWidth + (TotalCount2 - 1) * Admission;
                Square1 += Convert.ToDecimal(TotalHeight) * Convert.ToDecimal(TotalWidth) * TotalCount1 / 1000000;
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    if (Convert.ToInt32(DT2.Rows[j]["Height"]) == H && Convert.ToInt32(DT2.Rows[j]["Width"]) == W &&
                        DT2.Rows[j]["Notes"].ToString() == N)
                        continue;

                    H = Convert.ToInt32(DT2.Rows[j]["Height"]);
                    W = Convert.ToInt32(DT2.Rows[j]["Width"]);
                    N = DT2.Rows[j]["Notes"].ToString();

                    decimal Square2 = 0;
                    int Count = 0;

                    if (Convert.ToInt32(DT2.Rows[j]["Height"]) >= 40 && Convert.ToInt32(DT2.Rows[j]["Width"]) >= 40)
                        Deduction = 0;
                    else
                    {
                        if (ColorID == 1881 || ColorID == 1893 || ColorID == 1883 || ColorID == 1884 || ColorID == 1882 || ColorID == 1885 ||
                            ColorID == 1886 || ColorID == 1889 || ColorID == 1894 || (ColorID == 1890 && PatinaID == 6) || (ColorID == 1890 && PatinaID == 7))
                            Deduction = 2;

                        if (ColorID == 3694 || ColorID == 3695 || ColorID == 3696 || ColorID == 3697 || ColorID == 3698 || ColorID == 1890 ||
                            ColorID == 3699 || ColorID == 3700 || ColorID == 3701 || ColorID == 1891 || ColorID == 3702 || ColorID == 3703 ||
                            ColorID == 1892 || ColorID == 3704 || ColorID == 3705)
                            Deduction = 1;
                    }

                    string filter = "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) + " AND FrontID=" + Convert.ToInt32(DT1.Rows[i]["FrontID"]) + " AND PatinaID=" + Convert.ToInt32(DT1.Rows[i]["PatinaID"]) +
                        " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                        " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) + " AND (Notes='' OR Notes IS NULL OR Notes<>'УКВ-7')";
                    if (DT2.Rows[j]["Notes"] != DBNull.Value && DT2.Rows[j]["Notes"].ToString() == "УКВ-7")
                    {
                        filter = "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) + " AND FrontID=" + Convert.ToInt32(DT1.Rows[i]["FrontID"]) + " AND PatinaID=" + Convert.ToInt32(DT1.Rows[i]["PatinaID"]) +
                            " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                            " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) +
                            " AND Notes='УКВ-7'";
                    }
                    DataRow[] rows = SourceDT.Select(filter);
                    foreach (DataRow item in rows)
                    {
                        Count += Convert.ToInt32(item["Count"]);
                        Square2 += Convert.ToDecimal(item["Height"]) * (Convert.ToDecimal(item["Width"]) - Deduction) * Convert.ToDecimal(item["Count"]) / 1000000;
                    }
                    Square2 = decimal.Round(Square2, 3, MidpointRounding.AwayFromZero);

                    DataRow NewRow = DestinationDT.NewRow();
                    NewRow["Name"] = GetFrontName(Convert.ToInt32(DT1.Rows[i]["FrontID"]));
                    if (Convert.ToInt32(DT1.Rows[i]["PatinaID"]) != -1)
                        NewRow["FrameColor"] = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"])) + " " + GetPatinaName(Convert.ToInt32(DT1.Rows[i]["PatinaID"]));
                    else
                        NewRow["FrameColor"] = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                    if (FirstRow)
                    {
                        Square1 = decimal.Round(Square1, 3, MidpointRounding.AwayFromZero);
                        NewRow["Height1"] = TotalWidth + 10;
                        NewRow["Width1"] = TotalHeight + 10;
                        NewRow["Count1"] = TotalCount1;
                        NewRow["Square1"] = Square1;
                    }
                    NewRow["Height2"] = Convert.ToInt32(DT2.Rows[j]["Width"]) - Deduction;
                    NewRow["Width2"] = Convert.ToInt32(DT2.Rows[j]["Height"]);
                    NewRow["Count2"] = Count;
                    NewRow["Square2"] = Square2;
                    NewRow["Notes"] = DT2.Rows[j]["Notes"];
                    NewRow["ColorID"] = ColorID;
                    DestinationDT.Rows.Add(NewRow);
                    FirstRow = false;
                }
            }
        }

        private void RoughCutCoolectByWidth(DataTable SourceDT, ref DataTable DestinationDT)
        {
            int Admission = 5;
            int HandsAdmission = 0;
            int Deduction = 0;
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            int H = 0;
            int W = 0;
            int TotalHeight = 0;
            int TotalCount2 = 0;
            int TotalCount1 = 1;
            int TotalWidth = 0;
            decimal Square1 = 0;
            string N = string.Empty;

            using (DataView DV = new DataView(SourceDT, string.Empty, "FrontID, ColorID, PatinaID", DataViewRowState.CurrentRows))
            {
                DT1 = DV.ToTable(true, new string[] { "FrontID", "ColorID", "PatinaID" });
            }
            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                H = 0;
                W = 0;
                N = string.Empty;
                int FrontID = Convert.ToInt32(DT1.Rows[i]["FrontID"]);
                int ColorID = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                int PatinaID = Convert.ToInt32(DT1.Rows[i]["PatinaID"]);

                using (DataView DV = new DataView(SourceDT, "ColorID=" + ColorID + " AND FrontID=" + FrontID + " AND PatinaID=" + PatinaID,
                    "Height, Width, Notes", DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "Height", "Width", "Notes" });
                }
                bool FirstRow = true;
                Square1 = 0;
                Admission = 5;
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    int Count = 0;
                    HandsAdmission = 0;

                    string filter = "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) + " AND FrontID=" + Convert.ToInt32(DT1.Rows[i]["FrontID"]) + " AND PatinaID=" + Convert.ToInt32(DT1.Rows[i]["PatinaID"]) +
                           " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                           " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) + " AND (Notes='' OR Notes IS NULL OR Notes<>'УКВ-7')";
                    if (DT2.Rows[j]["Notes"] != DBNull.Value && DT2.Rows[j]["Notes"].ToString() == "УКВ-7")
                    {
                        filter = "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) + " AND FrontID=" + Convert.ToInt32(DT1.Rows[i]["FrontID"]) + " AND PatinaID=" + Convert.ToInt32(DT1.Rows[i]["PatinaID"]) +
                            " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                            " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) +
                            " AND Notes='УКВ-7'";
                        Admission += 20;
                        HandsAdmission = -36;
                    }
                    DataRow[] rows = SourceDT.Select(filter);
                    foreach (DataRow item in rows)
                    {
                        Count += Convert.ToInt32(item["Count"]);
                    }

                    TotalHeight += Convert.ToInt32(DT2.Rows[j]["Height"]) * Count;
                    TotalWidth = Convert.ToInt32(DT2.Rows[j]["Width"]) + HandsAdmission;
                    TotalCount2 += Count;
                }
                if (TotalHeight + (TotalCount2 - 1) * Admission + 10 > 735)
                    TotalCount1 = TotalCount2;
                TotalHeight = TotalHeight + (TotalCount2 - 1) * Admission;
                Square1 += Convert.ToDecimal(TotalHeight) * Convert.ToDecimal(TotalWidth) * TotalCount1 / 1000000;
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    if (Convert.ToInt32(DT2.Rows[j]["Height"]) == H && Convert.ToInt32(DT2.Rows[j]["Width"]) == W &&
                        DT2.Rows[j]["Notes"].ToString() == N)
                        continue;

                    H = Convert.ToInt32(DT2.Rows[j]["Height"]);
                    W = Convert.ToInt32(DT2.Rows[j]["Width"]);
                    N = DT2.Rows[j]["Notes"].ToString();

                    decimal Square2 = 0;
                    int Count = 0;

                    if (Convert.ToInt32(DT2.Rows[j]["Height"]) >= 40 && Convert.ToInt32(DT2.Rows[j]["Width"]) >= 40)
                        Deduction = 0;
                    else
                    {
                        if (ColorID == 1881 || ColorID == 1893 || ColorID == 1883 || ColorID == 1884 || ColorID == 1882 || ColorID == 1885 ||
                            ColorID == 1886 || ColorID == 1889 || ColorID == 1894 || (ColorID == 1890 && PatinaID == 6) || (ColorID == 1890 && PatinaID == 7))
                            Deduction = 2;

                        if (ColorID == 3694 || ColorID == 3695 || ColorID == 3696 || ColorID == 3697 || ColorID == 3698 || ColorID == 1890 ||
                            ColorID == 3699 || ColorID == 3700 || ColorID == 3701 || ColorID == 1891 || ColorID == 3702 || ColorID == 3703 ||
                            ColorID == 1892 || ColorID == 3704 || ColorID == 3705)
                            Deduction = 1;
                    }

                    string filter = "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) + " AND FrontID=" + Convert.ToInt32(DT1.Rows[i]["FrontID"]) + " AND PatinaID=" + Convert.ToInt32(DT1.Rows[i]["PatinaID"]) +
                        " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                        " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) + " AND (Notes='' OR Notes IS NULL OR Notes<>'УКВ-7')";
                    if (DT2.Rows[j]["Notes"] != DBNull.Value && DT2.Rows[j]["Notes"].ToString() == "УКВ-7")
                    {
                        filter = "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) + " AND FrontID=" + Convert.ToInt32(DT1.Rows[i]["FrontID"]) + " AND PatinaID=" + Convert.ToInt32(DT1.Rows[i]["PatinaID"]) +
                            " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                            " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) +
                            " AND Notes='УКВ-7'";
                    }
                    DataRow[] rows = SourceDT.Select(filter);
                    foreach (DataRow item in rows)
                    {
                        Count += Convert.ToInt32(item["Count"]);
                        Square2 += Convert.ToDecimal(item["Height"]) * (Convert.ToDecimal(item["Width"]) - Deduction) * Convert.ToDecimal(item["Count"]) / 1000000;
                    }
                    Square2 = decimal.Round(Square2, 3, MidpointRounding.AwayFromZero);

                    DataRow NewRow = DestinationDT.NewRow();
                    NewRow["Name"] = GetFrontName(Convert.ToInt32(DT1.Rows[i]["FrontID"]));
                    if (Convert.ToInt32(DT1.Rows[i]["PatinaID"]) != -1)
                        NewRow["FrameColor"] = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"])) + " " + GetPatinaName(Convert.ToInt32(DT1.Rows[i]["PatinaID"]));
                    else
                        NewRow["FrameColor"] = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                    if (FirstRow)
                    {
                        Square1 = decimal.Round(Square1, 3, MidpointRounding.AwayFromZero);
                        NewRow["Height1"] = TotalWidth + 10;
                        NewRow["Width1"] = TotalHeight + 10;
                        NewRow["Count1"] = TotalCount1;
                        NewRow["Square1"] = Square1;
                    }
                    NewRow["Height2"] = Convert.ToInt32(DT2.Rows[j]["Width"]) - Deduction;
                    NewRow["Width2"] = Convert.ToInt32(DT2.Rows[j]["Height"]);
                    NewRow["Count2"] = Count;
                    NewRow["Square2"] = Square2;
                    NewRow["Notes"] = DT2.Rows[j]["Notes"];
                    NewRow["ColorID"] = ColorID;
                    DestinationDT.Rows.Add(NewRow);
                    FirstRow = false;
                }
            }
        }

        private void SetOfCovers(DataTable SourceDT, ref DataTable DestinationDT)
        {
            for (int i = 0; i < SourceDT.Rows.Count; i++)
            {
                if (SourceDT.Rows[i]["Height1"] == DBNull.Value)
                    continue;
                DataRow NewRow = DestinationDT.NewRow();
                NewRow["Name"] = SourceDT.Rows[i]["Name"];
                NewRow["FrameColor"] = SourceDT.Rows[i]["FrameColor"];
                NewRow["Height"] = Convert.ToInt32(SourceDT.Rows[i]["Height1"]) + 10;
                NewRow["Width"] = Convert.ToInt32(SourceDT.Rows[i]["Width1"]) + 10;
                NewRow["Count"] = Convert.ToInt32(SourceDT.Rows[i]["Count1"]) * 2;
                NewRow["Square"] = decimal.Round(Convert.ToDecimal(SourceDT.Rows[i]["Square1"]) * 2, 3, MidpointRounding.AwayFromZero);
                NewRow["Notes"] = SourceDT.Rows[i]["Notes"];
                NewRow["ColorID"] = SourceDT.Rows[i]["ColorID"];
                DestinationDT.Rows.Add(NewRow);
            }
        }
    }
}