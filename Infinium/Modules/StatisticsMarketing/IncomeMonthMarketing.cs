using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;

namespace Infinium.Modules.StatisticsMarketing
{
    public class IncomeMonthMarketing
    {
        private int FactoryID = 0;

        public DataTable IncomeTotalDataTable = null;
        public DataTable IncomeDataTable = null;

        public DataTable FrontsOrdersDataTable = null;
        public DataTable DecorOrdersDataTable = null;

        public DataTable DecorConfigDataTable = null;

        public BindingSource IncomeBindingSource = null;

        private PercentageDataGrid IncomeDataGrid;

        public IncomeMonthMarketing(ref PercentageDataGrid tIncomeDataGrid)
        {
            IncomeDataGrid = tIncomeDataGrid;

            Initialize();
        }

        public int Factory
        {
            get { return FactoryID; }
            set { FactoryID = value; }
        }

        private void Create()
        {
            IncomeTotalDataTable = new DataTable();
            IncomeTotalDataTable.Columns.Add(new DataColumn("Date", Type.GetType("System.String")));
            IncomeTotalDataTable.Columns.Add(new DataColumn("Cost", Type.GetType("System.Decimal")));
            IncomeTotalDataTable.Columns.Add(new DataColumn("Date1", Type.GetType("System.DateTime")));

            IncomeDataTable = new DataTable();
            IncomeDataTable.Columns.Add(new DataColumn("DateTime", Type.GetType("System.DateTime")));
            IncomeDataTable.Columns.Add(new DataColumn("Date", Type.GetType("System.String")));
            IncomeDataTable.Columns.Add(new DataColumn("Square", Type.GetType("System.Decimal")));
            IncomeDataTable.Columns.Add(new DataColumn("DecorLinearCount", Type.GetType("System.Decimal")));
            IncomeDataTable.Columns.Add(new DataColumn("DecorItemsCount", Type.GetType("System.Decimal")));
            IncomeDataTable.Columns.Add(new DataColumn("FrontsCost", Type.GetType("System.Decimal")));
            IncomeDataTable.Columns.Add(new DataColumn("DecorCost", Type.GetType("System.Decimal")));
            IncomeDataTable.Columns.Add(new DataColumn("TotalCost", Type.GetType("System.Decimal")));

            DecorConfigDataTable = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT DecorConfigID, MeasureID FROM DecorConfig",
                ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(DecorConfigDataTable);
            }

            FrontsOrdersDataTable = new DataTable();
            DecorOrdersDataTable = new DataTable();
        }

        public void Fill()
        {
            FrontsOrdersDataTable.Clear();
            DecorOrdersDataTable.Clear();
            IncomeTotalDataTable.Clear();
            IncomeDataTable.Clear();

            string FactoryFilter = string.Empty;

            if (FactoryID != 0)
                FactoryFilter = " WHERE NewFrontsOrders.FactoryID = " + FactoryID;

            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT CAST(NewMegaOrders.OrderDate AS DATE) AS OrderDate, NewFrontsOrders.Cost," +
                " NewFrontsOrders.Square, NewFrontsOrders.FactoryID FROM NewFrontsOrders" +
                " INNER JOIN NewMainOrders ON NewFrontsOrders.MainOrderID = NewMainOrders.MainOrderID" +
                " INNER JOIN NewMegaOrders ON NewMainOrders.MegaOrderID = NewMegaOrders.MegaOrderID" + FactoryFilter,

                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(FrontsOrdersDataTable);
            }

            if (FactoryID != 0)
                FactoryFilter = " WHERE NewDecorOrders.FactoryID = " + FactoryID;
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT CAST(NewMegaOrders.OrderDate AS DATE) AS OrderDate, NewDecorOrders.Cost," +
                " NewDecorOrders.Count, NewDecorOrders.Length, NewDecorOrders.FactoryID, MeasureID FROM NewDecorOrders" +
                " INNER JOIN NewMainOrders ON NewDecorOrders.MainOrderID = NewMainOrders.MainOrderID" +
                " INNER JOIN NewMegaOrders ON NewMainOrders.MegaOrderID = NewMegaOrders.MegaOrderID" +
                " INNER JOIN infiniu2_catalog.dbo.DecorConfig ON NewDecorOrders.DecorConfigID = infiniu2_catalog.dbo.DecorConfig.DecorConfigID" + FactoryFilter,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(DecorOrdersDataTable);
            }

            foreach (DataRow Row in FrontsOrdersDataTable.Rows)
            {
                string Date = Convert.ToDateTime(Row["OrderDate"]).ToString("yyyy. MMMM");

                DataRow[] iRows = IncomeTotalDataTable.Select("Date = '" + Date + "'");

                if (iRows.Count() == 0)
                {
                    DataRow NewRow = IncomeTotalDataTable.NewRow();
                    NewRow["Date"] = Date;
                    NewRow["Cost"] = Convert.ToDecimal(Row["Cost"]);
                    IncomeTotalDataTable.Rows.Add(NewRow);
                }
                else
                {
                    iRows[0]["Cost"] = Convert.ToDecimal(iRows[0]["Cost"]) + Convert.ToDecimal(Row["Cost"]);
                }
            }

            foreach (DataRow Row in DecorOrdersDataTable.Rows)
            {
                string Date = Convert.ToDateTime(Row["OrderDate"]).ToString("yyyy. MMMM");

                DataRow[] iRows = IncomeTotalDataTable.Select("Date = '" + Date + "'");

                if (iRows.Count() == 0)
                {
                    DataRow NewRow = IncomeTotalDataTable.NewRow();
                    NewRow["Date"] = Date;
                    NewRow["Cost"] = Convert.ToDecimal(Row["Cost"]);
                    IncomeTotalDataTable.Rows.Add(NewRow);
                }
                else
                {
                    iRows[0]["Cost"] = Convert.ToDecimal(iRows[0]["Cost"]) + Convert.ToDecimal(Row["Cost"]);
                }
            }

            FactoryFilter = string.Empty;

            //using (DataView DV = new DataView(IncomeTotalDataTable.Copy()))
            //{
            //    DV.Sort = "Date1";
            //    IncomeTotalDataTable.Clear();
            //    IncomeTotalDataTable = DV.ToTable();
            //}

            foreach (DataRow Row in IncomeTotalDataTable.Rows)
            {
                string D = Convert.ToDateTime(Row["Date"]).ToString("yyyy-MM") + "-%";

                if (FactoryID != 0)
                    FactoryFilter = " AND NewFrontsOrders.FactoryID = " + FactoryID;
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT CAST(NewMegaOrders.OrderDate AS DATE) AS OrderDate, NewFrontsOrders.Cost," +
                " NewFrontsOrders.Square, NewFrontsOrders.FactoryID FROM NewFrontsOrders" +
                " INNER JOIN NewMainOrders ON NewFrontsOrders.MainOrderID = NewMainOrders.MainOrderID" +
                " INNER JOIN NewMegaOrders ON NewMainOrders.MegaOrderID = NewMegaOrders.MegaOrderID" +
                " WHERE CAST(NewMegaOrders.OrderDate AS DATE) LIKE '" + D + "'" + FactoryFilter,
                ConnectionStrings.MarketingOrdersConnectionString))
                {
                    FrontsOrdersDataTable.Clear();
                    DA.Fill(FrontsOrdersDataTable);
                }

                if (FactoryID != 0)
                    FactoryFilter = " AND NewDecorOrders.FactoryID = " + FactoryID;
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT CAST(NewMegaOrders.OrderDate AS DATE) AS OrderDate, NewDecorOrders.Cost," +
                    " NewDecorOrders.Count, NewDecorOrders.Length, NewDecorOrders.FactoryID, MeasureID FROM NewDecorOrders" +
                    " INNER JOIN NewMainOrders ON NewDecorOrders.MainOrderID = NewMainOrders.MainOrderID" +
                    " INNER JOIN NewMegaOrders ON NewMainOrders.MegaOrderID = NewMegaOrders.MegaOrderID" +
                    " INNER JOIN infiniu2_catalog.dbo.DecorConfig ON NewDecorOrders.DecorConfigID = infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                    " WHERE CAST(NewMegaOrders.OrderDate AS DATE) LIKE '" + D + "'" + FactoryFilter,
                    ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DecorOrdersDataTable.Clear();
                    DA.Fill(DecorOrdersDataTable);
                }

                decimal Square = 0;
                decimal DecorLinearCount = 0;
                decimal DecorItemsCount = 0;
                decimal FrontsCost = 0;
                decimal DecorCost = 0;

                foreach (DataRow mRow in FrontsOrdersDataTable.Rows)
                {
                    Square += Convert.ToDecimal(mRow["Square"]);
                    FrontsCost += Convert.ToDecimal(mRow["Cost"]);
                }

                foreach (DataRow dRow in DecorOrdersDataTable.Rows)
                {
                    DecorCost += Convert.ToDecimal(dRow["Cost"]);
                    if (dRow["MeasureID"].ToString() == "2")
                        DecorLinearCount += Convert.ToDecimal(dRow["Count"]) * Convert.ToDecimal(dRow["Length"]) / 1000;
                    else
                        DecorItemsCount += Convert.ToDecimal(dRow["Count"]);
                }

                DataRow NewRow = IncomeDataTable.NewRow();
                NewRow["DateTime"] = Row["Date"];
                NewRow["Date"] = Convert.ToDateTime(Row["Date"]).ToString("yyyy. MMMM");
                NewRow["Square"] = Decimal.Round(Square, 2, MidpointRounding.AwayFromZero);
                NewRow["FrontsCost"] = Decimal.Round(FrontsCost, 2, MidpointRounding.AwayFromZero);
                NewRow["DecorCost"] = Decimal.Round(DecorCost, 2, MidpointRounding.AwayFromZero);
                NewRow["TotalCost"] = Decimal.Round(FrontsCost + DecorCost, 2, MidpointRounding.AwayFromZero);
                NewRow["DecorLinearCount"] = Decimal.Round(DecorLinearCount, 2, MidpointRounding.AwayFromZero);
                NewRow["DecorItemsCount"] = Decimal.Round(DecorItemsCount, 2, MidpointRounding.AwayFromZero);
                IncomeDataTable.Rows.Add(NewRow);
            }
        }

        private void Binding()
        {
            IncomeBindingSource = new BindingSource()
            {
                DataSource = IncomeDataTable
            };
            IncomeDataGrid.DataSource = IncomeBindingSource;
        }

        private void SetGrids()
        {
            foreach (DataGridViewColumn Column in IncomeDataGrid.Columns)
            {
                Column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }

            IncomeDataGrid.Columns["DateTime"].Visible = false;

            IncomeDataGrid.Columns["Date"].HeaderText = "Месяц";
            IncomeDataGrid.Columns["Square"].HeaderText = "Квадратура";
            IncomeDataGrid.Columns["FrontsCost"].HeaderText = "Фасады, €";
            IncomeDataGrid.Columns["DecorCost"].HeaderText = "Декор, €";
            IncomeDataGrid.Columns["TotalCost"].HeaderText = "Итого, €";
            IncomeDataGrid.Columns["DecorLinearCount"].HeaderText = "Декор, м.п.";
            IncomeDataGrid.Columns["DecorItemsCount"].HeaderText = "Декор, шт.";

            //ResultDataGrid.Columns["Client"].HeaderText = "Клиент";

            NumberFormatInfo nfi1 = new NumberFormatInfo()
            {
                CurrencyGroupSeparator = " ",
                CurrencySymbol = string.Empty,
                CurrencyDecimalDigits = 1
            };
            NumberFormatInfo nfi2 = new NumberFormatInfo()
            {
                CurrencyGroupSeparator = " ",
                CurrencySymbol = string.Empty,
                CurrencyDecimalDigits = 1
            };
            IncomeDataGrid.Columns["Square"].DefaultCellStyle.Format = "C";
            IncomeDataGrid.Columns["Square"].DefaultCellStyle.FormatProvider = nfi1;

            IncomeDataGrid.Columns["FrontsCost"].DefaultCellStyle.Format = "C";
            IncomeDataGrid.Columns["FrontsCost"].DefaultCellStyle.FormatProvider = nfi2;

            IncomeDataGrid.Columns["DecorCost"].DefaultCellStyle.Format = "C";
            IncomeDataGrid.Columns["DecorCost"].DefaultCellStyle.FormatProvider = nfi2;

            IncomeDataGrid.Columns["TotalCost"].DefaultCellStyle.Format = "C";
            IncomeDataGrid.Columns["TotalCost"].DefaultCellStyle.FormatProvider = nfi2;

            IncomeDataGrid.Columns["DecorLinearCount"].DefaultCellStyle.Format = "C";
            IncomeDataGrid.Columns["DecorLinearCount"].DefaultCellStyle.FormatProvider = nfi1;

            IncomeDataGrid.Columns["DecorItemsCount"].DefaultCellStyle.Format = "C";
            IncomeDataGrid.Columns["DecorItemsCount"].DefaultCellStyle.FormatProvider = nfi2;

            //IncomeDataGrid.Columns["Square"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //IncomeDataGrid.Columns["FrontsCost"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //IncomeDataGrid.Columns["DecorCost"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //IncomeDataGrid.Columns["TotalCost"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //IncomeDataGrid.Columns["DecorLinearCount"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //IncomeDataGrid.Columns["DecorItemsCount"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //IncomeDataGrid.Columns["Square"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        }

        public void Initialize()
        {
            Create();
            Fill();
            Binding();
            SetGrids();
        }
    }
}