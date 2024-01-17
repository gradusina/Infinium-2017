using Infinium.Modules.Marketing.Orders;

using System;
using System.Data;
using System.Data.SqlClient;
using System.Linq;

namespace Infinium.Modules.ExpeditionMarketing.StandardReport
{
    public class DecorReport
    {
        private string ProfilCurrencyCode = "0";
        private string TPSCurrencyCode = "0";
        private decimal PaymentRate = 0;

        private string UNN = string.Empty;
        //int ClientID = 0;

        public DataTable CurrencyTypesDT = null;

        private DecorCatalogOrder DecorCatalogOrder = null;

        public DataTable ProfilReportDataTable = null;
        public DataTable TPSReportDataTable = null;
        private DataTable DecorProductsDataTable = null;
        private DataTable DecorConfigDataTable = null;
        private DataTable MeasuresDataTable = null;
        private DataTable DecorOrdersDataTable = null;
        private DataTable CoverTypesDataTable = null;
        private DataTable DecorDataTable = null;

        public DecorReport(ref DecorCatalogOrder tDecorCatalogOrder)
        {
            DecorCatalogOrder = tDecorCatalogOrder;

            Create();
            CreateReportDataTable();

            CurrencyTypesDT = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM CurrencyTypes",
                ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(CurrencyTypesDT);
            }
        }

        private void Create()
        {
            DecorOrdersDataTable = new DataTable();

            string SelectCommand = @"SELECT ProductID, ProductName, MeasureID, ReportParam FROM DecorProducts" +
                " WHERE (ProductID IN (SELECT ProductID FROM DecorConfig)) ORDER BY ProductName ASC";
            DecorProductsDataTable = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(DecorProductsDataTable);
            }
            DecorDataTable = new DataTable();
            SelectCommand = @"SELECT DISTINCT TechStore.TechStoreID AS DecorID, TechStore.TechStoreName AS Name, DecorConfig.ProductID FROM TechStore 
                INNER JOIN DecorConfig ON TechStore.TechStoreID = DecorConfig.DecorID ORDER BY TechStoreName";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(DecorDataTable);
            }
            MeasuresDataTable = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM Measures",
                ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(MeasuresDataTable);
            }

            DecorConfigDataTable = new DataTable();
            //using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM DecorConfig",
            //    ConnectionStrings.CatalogConnectionString))
            //{
            //    DA.Fill(DecorConfigDataTable);
            //}
            DecorConfigDataTable = TablesManager.DecorConfigDataTableAll;

            CoverTypesDataTable = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM CoverTypes",
                ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(CoverTypesDataTable);
            }
        }

        private void CreateReportDataTable()
        {
            ProfilReportDataTable = new DataTable();
            ProfilReportDataTable.Columns.Add(new DataColumn("UNN", Type.GetType("System.String")));
            ProfilReportDataTable.Columns.Add(new DataColumn("CurrencyCode", Type.GetType("System.String")));
            ProfilReportDataTable.Columns.Add(new DataColumn("TPSCurCode", Type.GetType("System.String")));
            ProfilReportDataTable.Columns.Add(new DataColumn("InvNumber", Type.GetType("System.String")));
            ProfilReportDataTable.Columns.Add(new DataColumn("AccountingName", Type.GetType("System.String")));
            ProfilReportDataTable.Columns.Add(new DataColumn("Name", Type.GetType("System.String")));
            ProfilReportDataTable.Columns.Add(new DataColumn("Count", Type.GetType("System.Decimal")));
            ProfilReportDataTable.Columns.Add(new DataColumn("Measure", Type.GetType("System.String")));
            ProfilReportDataTable.Columns.Add(new DataColumn("OriginalPrice", Type.GetType("System.String")));
            ProfilReportDataTable.Columns.Add(new DataColumn("PriceWithTransport", Type.GetType("System.String")));
            ProfilReportDataTable.Columns.Add(new DataColumn("CostWithTransport", Type.GetType("System.String")));
            ProfilReportDataTable.Columns.Add(new DataColumn("DiscountVolume", Type.GetType("System.String")));
            ProfilReportDataTable.Columns.Add(new DataColumn("TotalDiscount", Type.GetType("System.String")));
            ProfilReportDataTable.Columns.Add(new DataColumn("Price", Type.GetType("System.String")));
            ProfilReportDataTable.Columns.Add(new DataColumn("Cost", Type.GetType("System.String")));
            ProfilReportDataTable.Columns.Add(new DataColumn("Weight", Type.GetType("System.String")));
            ProfilReportDataTable.Columns.Add(new DataColumn("PaymentRate", Type.GetType("System.Decimal")));
            TPSReportDataTable = ProfilReportDataTable.Clone();
        }

        public void ClearReport()
        {
            DecorOrdersDataTable.Clear();

            ProfilReportDataTable.Clear();
            TPSReportDataTable.Clear();

            ProfilReportDataTable.AcceptChanges();
            TPSReportDataTable.AcceptChanges();
        }

        private bool IsProfil(int DecorConfigID)
        {
            DataRow[] Rows = DecorConfigDataTable.Select("DecorConfigID = " + DecorConfigID.ToString());

            if (Rows[0]["FactoryID"].ToString() == "1")
                return true;
            return false;
        }

        private int GetReportMeasureTypeID(int DecorConfigID)
        {
            DataRow[] Row = DecorConfigDataTable.Select("DecorConfigID = " + DecorConfigID);
            if (Row.Length > 0)
                return Convert.ToInt32(Row[0]["ReportMeasureID"]);//1 м.кв.  2 м.п. 3 шт.
            else
            {
                System.Windows.Forms.MessageBox.Show(@$"Ошибка конфигурации: не найдена DecorConfigID={DecorConfigID}. 
                        DecorConfigDataTable.Count={DecorConfigDataTable.Rows.Count}");
                return -1;
            }
        }

        private int GetMeasureTypeID(int DecorConfigID)
        {
            DataRow[] Row = DecorConfigDataTable.Select("DecorConfigID = " + DecorConfigID);

            return Convert.ToInt32(Row[0]["MeasureID"]);//1 м.кв.  2 м.п. 3 шт.
        }

        private string GetItemName(int DecorID)
        {
            DataRow[] Row = DecorDataTable.Select("DecorID = " + DecorID);

            return Row[0]["Name"].ToString();
        }

        private string GetProductName(int ProductID)
        {
            DataRow[] Row = DecorProductsDataTable.Select("ProductID = " + ProductID);

            return Row[0]["ProductName"].ToString();
        }

        public decimal GetDecorWeight(DataRow DecorOrderRow)
        {
            if (DecorOrderRow["Weight"] == DBNull.Value)
            {
                System.Windows.Forms.MessageBox.Show("Ошибка конфигурации: нет веса в DecorConfigID=" + DecorOrderRow["DecorConfigID"].ToString() +
                    ". Вес будет выставлен в 0.");
                return 0;
            }
            decimal Weight = 0;

            DataRow[] Row = DecorConfigDataTable.Select("DecorConfigID = " + DecorOrderRow["DecorConfigID"].ToString());

            if (Row[0]["Weight"] == DBNull.Value)
            {
                System.Windows.Forms.MessageBox.Show("Ошибка конфигурации: нет веса в DecorConfigID=" + DecorOrderRow["DecorConfigID"].ToString() +
                    ". Вес будет выставлен в 0.");
                return 0;
            }
            if (Row[0]["WeightMeasureID"].ToString() == "1")
            {
                if (Convert.ToDecimal(DecorOrderRow["Height"]) != -1)
                    Weight = Convert.ToDecimal(DecorOrderRow["Height"]) * Convert.ToDecimal(DecorOrderRow["Width"]) / 1000000
                         * Convert.ToDecimal(Row[0]["Weight"]) * Convert.ToDecimal(DecorOrderRow["Count"]);
                if (Convert.ToDecimal(DecorOrderRow["Length"]) != -1)
                    Weight = Convert.ToDecimal(DecorOrderRow["Length"]) * Convert.ToDecimal(DecorOrderRow["Width"]) / 1000000
                         * Convert.ToDecimal(Row[0]["Weight"]) * Convert.ToDecimal(DecorOrderRow["Count"]);
            }
            decimal L = 0;

            if (Row[0]["WeightMeasureID"].ToString() == "2")
            {

                L = 0;

                L = Convert.ToDecimal(DecorOrderRow["Length"]);

                if (L == -1)
                    L = Convert.ToDecimal(DecorOrderRow["Height"]);

                Weight = Convert.ToDecimal(L) / 1000 * Convert.ToDecimal(Row[0]["Weight"]) *
                         Convert.ToDecimal(DecorOrderRow["Count"]);

            }
            if (Row[0]["WeightMeasureID"].ToString() == "3")
                Weight = Convert.ToDecimal(Row[0]["Weight"]) * Convert.ToDecimal(DecorOrderRow["Count"]);

            return Weight;
        }

        public void CreateParamsTable(string Params, DataTable DT)
        {
            string Param = null;

            for (int i = 0; i < Params.Length; i++)
            {
                if (Params[i] != ';')
                    Param += Params[i];

                if (Params[i] == ';' || i == Params.Length - 1)
                {
                    if (Param.Length > 0)
                    {
                        DT.Columns.Add(new DataColumn(Param, Type.GetType("System.Int32")));
                        Param = "";
                    }
                }
            }

            DT.Columns.Add(new DataColumn("UNN", Type.GetType("System.String")));
            DT.Columns.Add(new DataColumn("CurrencyCode", Type.GetType("System.String")));
            DT.Columns.Add(new DataColumn("TPSCurCode", Type.GetType("System.String")));
            DT.Columns.Add(new DataColumn("InvNumber", Type.GetType("System.String")));
            DT.Columns.Add(new DataColumn("AccountingName", Type.GetType("System.String")));
            DT.Columns.Add(new DataColumn("OriginalPrice", Type.GetType("System.String")));
            DT.Columns.Add(new DataColumn("PriceWithTransport", Type.GetType("System.String")));
            DT.Columns.Add(new DataColumn("CostWithTransport", Type.GetType("System.String")));
            DT.Columns.Add(new DataColumn("DiscountVolume", Type.GetType("System.String")));
            DT.Columns.Add(new DataColumn("TotalDiscount", Type.GetType("System.String")));
            DT.Columns.Add(new DataColumn("ProductID", Type.GetType("System.Int32")));
            DT.Columns.Add(new DataColumn("DecorID", Type.GetType("System.Int32")));
            DT.Columns.Add(new DataColumn("TotalCount", Type.GetType("System.Int32")));
            DT.Columns.Add(new DataColumn("Count", Type.GetType("System.Decimal")));
            DT.Columns.Add(new DataColumn("Cost", Type.GetType("System.Decimal")));
            DT.Columns.Add(new DataColumn("Price", Type.GetType("System.Decimal")));
            DT.Columns.Add(new DataColumn("Weight", Type.GetType("System.Decimal")));
            DT.Columns.Add(new DataColumn("PaymentRate", Type.GetType("System.Decimal")));
        }

        private void GroupCoverTypes(DataRow[] Rows, int MeasureTypeID, int DecorID)
        {
            DataTable PDT = new DataTable();
            DataTable TDT = new DataTable();

            PDT = DecorOrdersDataTable.Clone();
            TDT = DecorOrdersDataTable.Clone();

            PDT.Columns.Remove("Count");
            PDT.Columns.Add(new DataColumn("TotalCount", Type.GetType("System.Int32")));
            PDT.Columns.Add(new DataColumn("Count", Type.GetType("System.Decimal")));
            PDT.Columns.Add(new DataColumn("UNN", Type.GetType("System.String")));
            PDT.Columns.Add(new DataColumn("CurrencyCode", Type.GetType("System.String")));
            PDT.Columns.Add(new DataColumn("TPSCurCode", Type.GetType("System.String")));

            TDT.Columns.Remove("Count");
            TDT.Columns.Add(new DataColumn("TotalCount", Type.GetType("System.Int32")));
            TDT.Columns.Add(new DataColumn("Count", Type.GetType("System.Decimal")));
            TDT.Columns.Add(new DataColumn("UNN", Type.GetType("System.String")));
            TDT.Columns.Add(new DataColumn("CurrencyCode", Type.GetType("System.String")));
            TDT.Columns.Add(new DataColumn("TPSCurCode", Type.GetType("System.String")));


            for (int r = 0; r < Rows.Count(); r++)
            {
                string InvNumber = Rows[r]["InvNumber"].ToString();
                //м.п.
                if (MeasureTypeID == 2)
                {
                    decimal L = 0;

                    L = Convert.ToDecimal(Rows[r]["Length"]);

                    if (L == -1)
                        L = Convert.ToDecimal(Rows[r]["Height"]);

                    if (IsProfil(Convert.ToInt32(Rows[r]["DecorConfigID"])))
                    {
                        if (PDT.Rows.Count == 0)
                        {
                            DataRow NewRow = PDT.NewRow();
                            NewRow["UNN"] = UNN;
                            NewRow["PaymentRate"] = PaymentRate;
                            NewRow["AccountingName"] = Rows[r]["AccountingName"].ToString();
                            NewRow["InvNumber"] = Rows[r]["InvNumber"].ToString();
                            NewRow["CurrencyCode"] = ProfilCurrencyCode;
                            NewRow["ProductID"] = Rows[r]["ProductID"];
                            NewRow["DecorID"] = Rows[r]["DecorID"];
                            NewRow["Count"] = Convert.ToDecimal(Rows[r]["Count"]) * L;
                            NewRow["OriginalPrice"] = Convert.ToDecimal(Rows[r]["OriginalPrice"]);
                            NewRow["DiscountVolume"] = Convert.ToDecimal(Rows[r]["DiscountVolume"]);
                            NewRow["TotalCount"] = Convert.ToDecimal(Rows[r]["Count"]);
                            NewRow["TotalDiscount"] = Convert.ToDecimal(Rows[r]["TotalDiscount"]);
                            NewRow["Price"] = Convert.ToDecimal(Rows[r]["Price"]);
                            NewRow["Cost"] = Convert.ToDecimal(Rows[r]["Cost"]);
                            NewRow["PriceWithTransport"] = Convert.ToDecimal(Rows[r]["PriceWithTransport"]);
                            NewRow["CostWithTransport"] = Convert.ToDecimal(Rows[r]["CostWithTransport"]);
                            NewRow["Weight"] = GetDecorWeight(Rows[r]);
                            PDT.Rows.Add(NewRow);
                            continue;
                        }
                        else
                        {
                            PDT.Rows[0]["OriginalPrice"] = (Convert.ToDecimal(PDT.Rows[0]["OriginalPrice"]) * Convert.ToDecimal(PDT.Rows[0]["TotalCount"]) + Convert.ToDecimal(Rows[r]["Count"]) * Convert.ToDecimal(Rows[r]["OriginalPrice"])) /
                                (Convert.ToDecimal(PDT.Rows[0]["TotalCount"]) + Convert.ToDecimal(Rows[r]["Count"]));
                            PDT.Rows[0]["Price"] = (Convert.ToDecimal(PDT.Rows[0]["Price"]) * Convert.ToDecimal(PDT.Rows[0]["TotalCount"]) + Convert.ToDecimal(Rows[r]["Count"]) * Convert.ToDecimal(Rows[r]["Price"])) /
                                (Convert.ToDecimal(PDT.Rows[0]["TotalCount"]) + Convert.ToDecimal(Rows[r]["Count"]));
                            PDT.Rows[0]["TotalCount"] = Convert.ToDecimal(PDT.Rows[0]["TotalCount"]) +
                                Convert.ToDecimal(Rows[r]["Count"]);
                            PDT.Rows[0]["Count"] = Convert.ToDecimal(PDT.Rows[0]["Count"]) +
                                                                      Convert.ToDecimal(Rows[r]["Count"]) * L;

                            PDT.Rows[0]["Cost"] = Convert.ToDecimal(PDT.Rows[0]["Cost"]) + Convert.ToDecimal(Rows[r]["Cost"]);
                            PDT.Rows[0]["CostWithTransport"] = Convert.ToDecimal(PDT.Rows[0]["CostWithTransport"]) + Convert.ToDecimal(Rows[r]["CostWithTransport"]);
                            PDT.Rows[0]["Weight"] = Convert.ToDecimal(PDT.Rows[0]["Weight"]) +
                                                         GetDecorWeight(Rows[r]);

                            continue;
                        }

                    }
                    else
                    {
                        if (TDT.Rows.Count == 0)
                        {
                            DataRow NewRow = TDT.NewRow();
                            NewRow["UNN"] = UNN;
                            NewRow["PaymentRate"] = PaymentRate;
                            NewRow["AccountingName"] = Rows[r]["AccountingName"].ToString();
                            NewRow["InvNumber"] = Rows[r]["InvNumber"].ToString();
                            NewRow["CurrencyCode"] = ProfilCurrencyCode;
                            NewRow["TPSCurCode"] = TPSCurrencyCode;
                            NewRow["ProductID"] = Rows[r]["ProductID"];
                            NewRow["DecorID"] = Rows[r]["DecorID"];
                            NewRow["Count"] = Convert.ToDecimal(Rows[r]["Count"]) * L;
                            NewRow["OriginalPrice"] = Convert.ToDecimal(Rows[r]["OriginalPrice"]);
                            NewRow["PriceWithTransport"] = Convert.ToDecimal(Rows[r]["PriceWithTransport"]);
                            NewRow["CostWithTransport"] = Convert.ToDecimal(Rows[r]["CostWithTransport"]);
                            NewRow["DiscountVolume"] = Convert.ToDecimal(Rows[r]["DiscountVolume"]);
                            NewRow["TotalCount"] = Convert.ToDecimal(Rows[r]["Count"]);
                            NewRow["TotalDiscount"] = Convert.ToDecimal(Rows[r]["TotalDiscount"]);
                            NewRow["Price"] = Convert.ToDecimal(Rows[r]["Price"]);
                            NewRow["Cost"] = Convert.ToDecimal(Rows[r]["Cost"]);
                            NewRow["Weight"] = GetDecorWeight(Rows[r]);
                            TDT.Rows.Add(NewRow);
                            continue;
                        }
                        else
                        {
                            TDT.Rows[0]["OriginalPrice"] = (Convert.ToDecimal(TDT.Rows[0]["OriginalPrice"]) * Convert.ToDecimal(TDT.Rows[0]["TotalCount"]) + Convert.ToDecimal(Rows[r]["Count"]) * Convert.ToDecimal(Rows[r]["OriginalPrice"])) /
                                (Convert.ToDecimal(TDT.Rows[0]["TotalCount"]) + Convert.ToDecimal(Rows[r]["Count"]));
                            TDT.Rows[0]["Price"] = (Convert.ToDecimal(TDT.Rows[0]["Price"]) * Convert.ToDecimal(TDT.Rows[0]["TotalCount"]) + Convert.ToDecimal(Rows[r]["Count"]) * Convert.ToDecimal(Rows[r]["Price"])) /
                                (Convert.ToDecimal(TDT.Rows[0]["TotalCount"]) + Convert.ToDecimal(Rows[r]["Count"]));
                            TDT.Rows[0]["TotalCount"] = Convert.ToDecimal(TDT.Rows[0]["TotalCount"]) +
                                Convert.ToDecimal(Rows[r]["Count"]);
                            TDT.Rows[0]["Count"] = Convert.ToDecimal(TDT.Rows[0]["Count"]) +
                                                                      Convert.ToDecimal(Rows[r]["Count"]) * L;

                            TDT.Rows[0]["Cost"] = Convert.ToDecimal(TDT.Rows[0]["Cost"]) + Convert.ToDecimal(Rows[r]["Cost"]);
                            TDT.Rows[0]["CostWithTransport"] = Convert.ToDecimal(TDT.Rows[0]["CostWithTransport"]) + Convert.ToDecimal(Rows[r]["CostWithTransport"]);
                            TDT.Rows[0]["Weight"] = Convert.ToDecimal(TDT.Rows[0]["Weight"]) +
                                                         GetDecorWeight(Rows[r]);

                            continue;
                        }
                    }
                }

                //шт.
                if (MeasureTypeID == 3)
                {
                    //get_parametrized_data function only
                }

                //м.кв.
                if (MeasureTypeID == 1)
                {
                    if (IsProfil(Convert.ToInt32(Rows[r]["DecorConfigID"])))
                    {
                        if (PDT.Rows.Count == 0)
                        {
                            DataRow NewRow = PDT.NewRow();
                            NewRow["UNN"] = UNN;
                            NewRow["PaymentRate"] = PaymentRate;
                            NewRow["AccountingName"] = Rows[r]["AccountingName"].ToString();
                            NewRow["InvNumber"] = Rows[r]["InvNumber"].ToString();
                            NewRow["CurrencyCode"] = ProfilCurrencyCode;
                            if (GetMeasureTypeID(Convert.ToInt32(Rows[0]["DecorConfigID"])) == 1)
                            {
                                if (Convert.ToDecimal(Rows[r]["Height"]) != -1)
                                {
                                    decimal d = Convert.ToDecimal(Rows[r]["Height"]) * Convert.ToDecimal(Rows[r]["Width"]) / 1000000;
                                    d = Decimal.Round(d, 3, MidpointRounding.AwayFromZero);
                                    decimal Square = d * Convert.ToDecimal(Rows[r]["Count"]);
                                    NewRow["Count"] = Square;
                                }
                                if (Convert.ToDecimal(Rows[r]["Length"]) != -1)
                                {
                                    decimal d = Convert.ToDecimal(Rows[r]["Length"]) * Convert.ToDecimal(Rows[r]["Width"]) / 1000000;
                                    d = Decimal.Round(d, 3, MidpointRounding.AwayFromZero);
                                    decimal Square = d * Convert.ToDecimal(Rows[r]["Count"]);
                                    NewRow["Count"] = Square;
                                }
                            }

                            if (GetMeasureTypeID(Convert.ToInt32(Rows[0]["DecorConfigID"])) == 2)
                            {
                                decimal L = 0;

                                L = Convert.ToDecimal(Rows[r]["Length"]);

                                if (L == -1)
                                    L = Convert.ToDecimal(Rows[r]["Height"]);
                                decimal d = L * Convert.ToDecimal(Rows[r]["Width"]) / 1000000;
                                d = Decimal.Round(d, 3, MidpointRounding.AwayFromZero);
                                decimal Square = d * Convert.ToDecimal(Rows[r]["Count"]);
                                NewRow["Count"] = Square;
                            }

                            if (GetMeasureTypeID(Convert.ToInt32(Rows[0]["DecorConfigID"])) == 3)
                            {

                                decimal H = 0;

                                if (Convert.ToDecimal(Rows[r]["Height"]) != -1)
                                    H = Convert.ToDecimal(Rows[r]["Height"]);

                                if (Convert.ToDecimal(Rows[r]["Length"]) != -1)
                                    H = Convert.ToDecimal(Rows[r]["Length"]);
                                decimal d = H * Convert.ToDecimal(Rows[r]["Width"]) / 1000000;
                                d = Decimal.Round(d, 3, MidpointRounding.AwayFromZero);
                                decimal Square = d * Convert.ToDecimal(Rows[r]["Count"]);
                                NewRow["Count"] = Square;
                                NewRow["CostWithTransport"] = Convert.ToDecimal(Rows[r]["PriceWithTransport"]);
                            }

                            NewRow["ProductID"] = Rows[r]["ProductID"];
                            NewRow["DecorID"] = Rows[r]["DecorID"];
                            NewRow["OriginalPrice"] = Convert.ToDecimal(Rows[r]["OriginalPrice"]);
                            NewRow["PriceWithTransport"] = Convert.ToDecimal(Rows[r]["PriceWithTransport"]);
                            NewRow["CostWithTransport"] = Convert.ToDecimal(Rows[r]["CostWithTransport"]);
                            NewRow["DiscountVolume"] = Convert.ToDecimal(Rows[r]["DiscountVolume"]);
                            NewRow["TotalCount"] = Convert.ToDecimal(Rows[r]["Count"]);
                            NewRow["TotalDiscount"] = Convert.ToDecimal(Rows[r]["TotalDiscount"]);
                            NewRow["Price"] = Convert.ToDecimal(Rows[r]["Price"]);
                            NewRow["Cost"] = Convert.ToDecimal(Rows[r]["Cost"]);
                            NewRow["Weight"] = GetDecorWeight(Rows[r]);
                            PDT.Rows.Add(NewRow);
                            continue;
                        }
                        else     //if no color parameter (hands e.g.)
                        {
                            if (GetMeasureTypeID(Convert.ToInt32(Rows[0]["DecorConfigID"])) == 1)
                            {
                                if (Convert.ToDecimal(Rows[r]["Height"]) != -1)
                                {
                                    decimal d = Convert.ToDecimal(Rows[r]["Height"]) * Convert.ToDecimal(Rows[r]["Width"]) / 1000000;
                                    d = Decimal.Round(d, 3, MidpointRounding.AwayFromZero);
                                    decimal Square = d * Convert.ToDecimal(Rows[r]["Count"]);
                                    PDT.Rows[0]["Count"] = Convert.ToDecimal(PDT.Rows[0]["Count"]) + Square;
                                }
                                if (Convert.ToDecimal(Rows[r]["Length"]) != -1)
                                {
                                    decimal d = Convert.ToDecimal(Rows[r]["Length"]) * Convert.ToDecimal(Rows[r]["Width"]) / 1000000;
                                    d = Decimal.Round(d, 3, MidpointRounding.AwayFromZero);
                                    decimal Square = d * Convert.ToDecimal(Rows[r]["Count"]);
                                    PDT.Rows[0]["Count"] = Convert.ToDecimal(PDT.Rows[0]["Count"]) + Square;
                                }
                            }

                            if (GetMeasureTypeID(Convert.ToInt32(Rows[0]["DecorConfigID"])) == 2)
                            {
                                decimal L = 0;

                                L = Convert.ToDecimal(Rows[r]["Length"]);

                                if (L == -1)
                                    L = Convert.ToDecimal(Rows[r]["Height"]);

                                decimal d = L * Convert.ToDecimal(Rows[r]["Width"]) / 1000000;
                                d = Decimal.Round(d, 3, MidpointRounding.AwayFromZero);
                                decimal Square = d * Convert.ToDecimal(Rows[r]["Count"]);
                                PDT.Rows[0]["Count"] = Convert.ToDecimal(PDT.Rows[0]["Count"]) + Square;
                            }

                            if (GetMeasureTypeID(Convert.ToInt32(Rows[0]["DecorConfigID"])) == 3)
                            {

                                decimal H = 0;

                                if (Convert.ToDecimal(Rows[r]["Height"]) != -1)
                                    H = Convert.ToDecimal(Rows[r]["Height"]);

                                if (Convert.ToDecimal(Rows[r]["Length"]) != -1)
                                    H = Convert.ToDecimal(Rows[r]["Length"]);
                                decimal d = H * Convert.ToDecimal(Rows[r]["Width"]) / 1000000;
                                d = Decimal.Round(d, 3, MidpointRounding.AwayFromZero);
                                decimal Square = d * Convert.ToDecimal(Rows[r]["Count"]);
                                PDT.Rows[0]["Count"] = Convert.ToDecimal(PDT.Rows[0]["Count"]) + Square;
                            }

                            PDT.Rows[0]["OriginalPrice"] = (Convert.ToDecimal(PDT.Rows[0]["OriginalPrice"]) * Convert.ToDecimal(PDT.Rows[0]["TotalCount"]) + Convert.ToDecimal(Rows[r]["Count"]) * Convert.ToDecimal(Rows[r]["OriginalPrice"])) /
                                (Convert.ToDecimal(PDT.Rows[0]["TotalCount"]) + Convert.ToDecimal(Rows[r]["Count"]));
                            PDT.Rows[0]["Price"] = (Convert.ToDecimal(PDT.Rows[0]["Price"]) * Convert.ToDecimal(PDT.Rows[0]["TotalCount"]) + Convert.ToDecimal(Rows[r]["Count"]) * Convert.ToDecimal(Rows[r]["Price"])) /
                                (Convert.ToDecimal(PDT.Rows[0]["TotalCount"]) + Convert.ToDecimal(Rows[r]["Count"]));
                            PDT.Rows[0]["TotalCount"] = Convert.ToDecimal(PDT.Rows[0]["TotalCount"]) +
                                Convert.ToDecimal(Rows[r]["Count"]);

                            PDT.Rows[0]["Cost"] = Convert.ToDecimal(PDT.Rows[0]["Cost"]) + Convert.ToDecimal(Rows[r]["Cost"]);
                            PDT.Rows[0]["CostWithTransport"] = Convert.ToDecimal(PDT.Rows[0]["CostWithTransport"]) + Convert.ToDecimal(Rows[r]["CostWithTransport"]);
                            PDT.Rows[0]["Weight"] = Convert.ToDecimal(PDT.Rows[0]["Weight"]) +
                                                         GetDecorWeight(Rows[r]);

                            continue;
                        }
                    }
                    else
                    {
                        if (TDT.Rows.Count == 0)
                        {
                            DataRow NewRow = TDT.NewRow();
                            NewRow["UNN"] = UNN;
                            NewRow["PaymentRate"] = PaymentRate;
                            NewRow["AccountingName"] = Rows[r]["AccountingName"].ToString();
                            NewRow["InvNumber"] = Rows[r]["InvNumber"].ToString();
                            NewRow["CurrencyCode"] = ProfilCurrencyCode;
                            NewRow["TPSCurCode"] = TPSCurrencyCode;
                            if (GetMeasureTypeID(Convert.ToInt32(Rows[0]["DecorConfigID"])) == 1)
                            {
                                if (Convert.ToDecimal(Rows[r]["Height"]) != -1)
                                {
                                    decimal d = Convert.ToDecimal(Rows[r]["Height"]) * Convert.ToDecimal(Rows[r]["Width"]) / 1000000;
                                    d = Decimal.Round(d, 3, MidpointRounding.AwayFromZero);
                                    decimal Square = d * Convert.ToDecimal(Rows[r]["Count"]);
                                    NewRow["Count"] = Square;
                                }
                                if (Convert.ToDecimal(Rows[r]["Length"]) != -1)
                                {
                                    decimal d = Convert.ToDecimal(Rows[r]["Length"]) * Convert.ToDecimal(Rows[r]["Width"]) / 1000000;
                                    d = Decimal.Round(d, 3, MidpointRounding.AwayFromZero);
                                    decimal Square = d * Convert.ToDecimal(Rows[r]["Count"]);
                                    NewRow["Count"] = Square;
                                }
                            }

                            if (GetMeasureTypeID(Convert.ToInt32(Rows[0]["DecorConfigID"])) == 2)
                            {
                                decimal L = 0;

                                L = Convert.ToDecimal(Rows[r]["Length"]);

                                if (L == -1)
                                    L = Convert.ToDecimal(Rows[r]["Height"]);
                                decimal d = L * Convert.ToDecimal(Rows[r]["Width"]) / 1000000;
                                d = Decimal.Round(d, 3, MidpointRounding.AwayFromZero);
                                decimal Square = d * Convert.ToDecimal(Rows[r]["Count"]);
                                NewRow["Count"] = Square;
                            }

                            if (GetMeasureTypeID(Convert.ToInt32(Rows[0]["DecorConfigID"])) == 3)
                            {

                                decimal H = 0;

                                if (Convert.ToDecimal(Rows[r]["Height"]) != -1)
                                    H = Convert.ToDecimal(Rows[r]["Height"]);

                                if (Convert.ToDecimal(Rows[r]["Length"]) != -1)
                                    H = Convert.ToDecimal(Rows[r]["Length"]);
                                decimal d = H * Convert.ToDecimal(Rows[r]["Width"]) / 1000000;
                                d = Decimal.Round(d, 3, MidpointRounding.AwayFromZero);
                                decimal Square = d * Convert.ToDecimal(Rows[r]["Count"]);
                                NewRow["Count"] = Square;

                            }

                            NewRow["ProductID"] = Rows[r]["ProductID"];
                            NewRow["DecorID"] = Rows[r]["DecorID"];
                            NewRow["OriginalPrice"] = Convert.ToDecimal(Rows[r]["OriginalPrice"]);
                            NewRow["PriceWithTransport"] = Convert.ToDecimal(Rows[r]["PriceWithTransport"]);
                            NewRow["CostWithTransport"] = Convert.ToDecimal(Rows[r]["CostWithTransport"]);
                            NewRow["DiscountVolume"] = Convert.ToDecimal(Rows[r]["DiscountVolume"]);
                            NewRow["TotalCount"] = Convert.ToDecimal(Rows[r]["Count"]);
                            NewRow["TotalDiscount"] = Convert.ToDecimal(Rows[r]["TotalDiscount"]);
                            NewRow["Price"] = Convert.ToDecimal(Rows[r]["Price"]);
                            NewRow["Cost"] = Convert.ToDecimal(Rows[r]["Cost"]);
                            NewRow["Weight"] = GetDecorWeight(Rows[r]);
                            TDT.Rows.Add(NewRow);
                            continue;
                        }
                        else //if no color parameter (hands e.g.)
                        {
                            if (GetMeasureTypeID(Convert.ToInt32(Rows[0]["DecorConfigID"])) == 1)
                            {
                                if (Convert.ToDecimal(Rows[r]["Height"]) != -1)
                                {
                                    decimal d = Convert.ToDecimal(Rows[r]["Height"]) * Convert.ToDecimal(Rows[r]["Width"]) / 1000000;
                                    d = Decimal.Round(d, 3, MidpointRounding.AwayFromZero);
                                    decimal Square = d * Convert.ToDecimal(Rows[r]["Count"]);
                                    TDT.Rows[0]["Count"] = Convert.ToDecimal(TDT.Rows[0]["Count"]) + Square;
                                }
                                if (Convert.ToDecimal(Rows[r]["Length"]) != -1)
                                {
                                    decimal d = Convert.ToDecimal(Rows[r]["Length"]) * Convert.ToDecimal(Rows[r]["Width"]) / 1000000;
                                    d = Decimal.Round(d, 3, MidpointRounding.AwayFromZero);
                                    decimal Square = d * Convert.ToDecimal(Rows[r]["Count"]);
                                    TDT.Rows[0]["Count"] = Convert.ToDecimal(TDT.Rows[0]["Count"]) + Square;
                                }
                            }

                            if (GetMeasureTypeID(Convert.ToInt32(Rows[0]["DecorConfigID"])) == 2)
                            {
                                decimal L = 0;

                                L = Convert.ToDecimal(Rows[r]["Length"]);

                                if (L == -1)
                                    L = Convert.ToDecimal(Rows[r]["Height"]);
                                decimal d = L * Convert.ToDecimal(Rows[r]["Width"]) / 1000000;
                                d = Decimal.Round(d, 3, MidpointRounding.AwayFromZero);
                                decimal Square = d * Convert.ToDecimal(Rows[r]["Count"]);
                                TDT.Rows[0]["Count"] = Convert.ToDecimal(TDT.Rows[0]["Count"]) + Square;
                            }

                            if (GetMeasureTypeID(Convert.ToInt32(Rows[0]["DecorConfigID"])) == 3)
                            {

                                decimal H = 0;

                                if (Convert.ToDecimal(Rows[r]["Height"]) != -1)
                                    H = Convert.ToDecimal(Rows[r]["Height"]);

                                if (Convert.ToDecimal(Rows[r]["Length"]) != -1)
                                    H = Convert.ToDecimal(Rows[r]["Length"]);

                                decimal d = H * Convert.ToDecimal(Rows[r]["Width"]) / 1000000;
                                d = Decimal.Round(d, 3, MidpointRounding.AwayFromZero);
                                decimal Square = d * Convert.ToDecimal(Rows[r]["Count"]);
                                TDT.Rows[0]["Count"] = Convert.ToDecimal(TDT.Rows[0]["Count"]) + Square;
                            }
                            TDT.Rows[0]["OriginalPrice"] = (Convert.ToDecimal(TDT.Rows[0]["OriginalPrice"]) * Convert.ToDecimal(TDT.Rows[0]["TotalCount"]) + Convert.ToDecimal(Rows[r]["Count"]) * Convert.ToDecimal(Rows[r]["OriginalPrice"])) /
                                (Convert.ToDecimal(TDT.Rows[0]["TotalCount"]) + Convert.ToDecimal(Rows[r]["Count"]));
                            TDT.Rows[0]["Price"] = (Convert.ToDecimal(TDT.Rows[0]["Price"]) * Convert.ToDecimal(TDT.Rows[0]["TotalCount"]) + Convert.ToDecimal(Rows[r]["Count"]) * Convert.ToDecimal(Rows[r]["Price"])) /
                                (Convert.ToDecimal(TDT.Rows[0]["TotalCount"]) + Convert.ToDecimal(Rows[r]["Count"]));
                            TDT.Rows[0]["TotalCount"] = Convert.ToDecimal(TDT.Rows[0]["TotalCount"]) +
                                Convert.ToDecimal(Rows[r]["Count"]);

                            TDT.Rows[0]["Cost"] = Convert.ToDecimal(TDT.Rows[0]["Cost"]) + Convert.ToDecimal(Rows[r]["Cost"]);
                            TDT.Rows[0]["CostWithTransport"] = Convert.ToDecimal(TDT.Rows[0]["CostWithTransport"]) + Convert.ToDecimal(Rows[r]["CostWithTransport"]);
                            TDT.Rows[0]["Weight"] = Convert.ToDecimal(TDT.Rows[0]["Weight"]) +
                                                         GetDecorWeight(Rows[r]);

                            continue;
                        }
                    }
                }
            }




            //REPORT TABLE
            //м.п.
            if (MeasureTypeID == 2)
            {
                if (PDT.Rows.Count > 0)
                {
                    for (int i = 0; i < PDT.Rows.Count; i++)
                    {
                        DataRow NewRow = ProfilReportDataTable.NewRow();

                        NewRow["UNN"] = UNN;
                        NewRow["PaymentRate"] = PaymentRate;
                        NewRow["AccountingName"] = PDT.Rows[i]["AccountingName"].ToString();
                        NewRow["InvNumber"] = PDT.Rows[i]["InvNumber"].ToString();
                        NewRow["CurrencyCode"] = ProfilCurrencyCode;
                        NewRow["Name"] = GetProductName(Convert.ToInt32(PDT.Rows[i]["ProductID"])) + " " +
                            GetItemName(DecorID);

                        NewRow["Count"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[i]["Count"]) / 1000, 3, MidpointRounding.AwayFromZero);
                        NewRow["Measure"] = "м.п.";
                        NewRow["OriginalPrice"] = Convert.ToDecimal(PDT.Rows[i]["OriginalPrice"]);
                        NewRow["DiscountVolume"] = Convert.ToDecimal(PDT.Rows[i]["DiscountVolume"]);
                        NewRow["TotalDiscount"] = Convert.ToDecimal(PDT.Rows[i]["TotalDiscount"]);
                        NewRow["Price"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[i]["Cost"]) / (Convert.ToDecimal(PDT.Rows[i]["Count"]) / 1000), 2, MidpointRounding.AwayFromZero);
                        NewRow["Cost"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[i]["Cost"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["PriceWithTransport"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[i]["CostWithTransport"]) / (Convert.ToDecimal(PDT.Rows[i]["Count"]) / 1000), 2, MidpointRounding.AwayFromZero);
                        NewRow["CostWithTransport"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[i]["CostWithTransport"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["Weight"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[i]["Weight"]), 3, MidpointRounding.AwayFromZero);

                        ProfilReportDataTable.Rows.Add(NewRow);
                    }
                }

                if (TDT.Rows.Count > 0)
                {
                    for (int i = 0; i < TDT.Rows.Count; i++)
                    {
                        DataRow NewRow = TPSReportDataTable.NewRow();
                        NewRow["UNN"] = UNN;
                        NewRow["PaymentRate"] = PaymentRate;
                        NewRow["AccountingName"] = TDT.Rows[i]["AccountingName"].ToString();
                        NewRow["InvNumber"] = TDT.Rows[i]["InvNumber"].ToString();
                        NewRow["CurrencyCode"] = ProfilCurrencyCode;
                        NewRow["TPSCurCode"] = TPSCurrencyCode;
                        NewRow["Name"] = GetProductName(Convert.ToInt32(TDT.Rows[i]["ProductID"])) + " " +
                            GetItemName(DecorID);

                        NewRow["Count"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[i]["Count"]) / 1000, 3, MidpointRounding.AwayFromZero);
                        NewRow["Measure"] = "м.п.";
                        NewRow["OriginalPrice"] = Convert.ToDecimal(TDT.Rows[i]["OriginalPrice"]);
                        NewRow["PriceWithTransport"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[i]["CostWithTransport"]) / (Convert.ToDecimal(TDT.Rows[i]["Count"]) / 1000), 2, MidpointRounding.AwayFromZero);
                        NewRow["CostWithTransport"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[i]["CostWithTransport"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["DiscountVolume"] = Convert.ToDecimal(TDT.Rows[i]["DiscountVolume"]);
                        NewRow["TotalDiscount"] = Convert.ToDecimal(TDT.Rows[i]["TotalDiscount"]);
                        NewRow["Price"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[i]["Cost"]) / (Convert.ToDecimal(TDT.Rows[i]["Count"]) / 1000), 2, MidpointRounding.AwayFromZero);
                        NewRow["Cost"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[i]["Cost"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["Weight"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[i]["Weight"]), 3, MidpointRounding.AwayFromZero);

                        TPSReportDataTable.Rows.Add(NewRow);
                    }
                }
            }

            //шт.
            if (MeasureTypeID == 3)
            {
                //get_parametrized_data function only
            }

            //м.кв.
            if (MeasureTypeID == 1)
            {
                if (PDT.Rows.Count > 0)
                {
                    for (int i = 0; i < PDT.Rows.Count; i++)
                    {
                        DataRow NewRow = ProfilReportDataTable.NewRow();

                        NewRow["UNN"] = UNN;
                        NewRow["PaymentRate"] = PaymentRate;
                        NewRow["AccountingName"] = PDT.Rows[i]["AccountingName"].ToString();
                        NewRow["InvNumber"] = PDT.Rows[i]["InvNumber"].ToString();
                        NewRow["CurrencyCode"] = ProfilCurrencyCode;
                        if (PDT.Rows[i]["ProductID"].ToString() == "10" || PDT.Rows[i]["ProductID"].ToString() == "11" ||
                            PDT.Rows[i]["ProductID"].ToString() == "12")
                            NewRow["Name"] = GetProductName(Convert.ToInt32(PDT.Rows[i]["ProductID"]));
                        else
                            NewRow["Name"] = GetProductName(Convert.ToInt32(PDT.Rows[i]["ProductID"])) + " " +
                                             GetItemName(DecorID);

                        NewRow["Count"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[i]["Count"]), 3, MidpointRounding.AwayFromZero);
                        NewRow["Measure"] = "м.кв.";
                        NewRow["OriginalPrice"] = Convert.ToDecimal(PDT.Rows[i]["OriginalPrice"]);
                        NewRow["PriceWithTransport"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[i]["CostWithTransport"]) / Convert.ToDecimal(PDT.Rows[i]["Count"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["CostWithTransport"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[i]["CostWithTransport"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["DiscountVolume"] = Convert.ToDecimal(PDT.Rows[i]["DiscountVolume"]);
                        NewRow["TotalDiscount"] = Convert.ToDecimal(PDT.Rows[i]["TotalDiscount"]);
                        NewRow["Price"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[i]["Cost"]) / Convert.ToDecimal(PDT.Rows[i]["Count"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["Cost"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[i]["Cost"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["Weight"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[i]["Weight"]), 3, MidpointRounding.AwayFromZero);

                        ProfilReportDataTable.Rows.Add(NewRow);
                    }
                }

                if (TDT.Rows.Count > 0)
                {
                    for (int i = 0; i < TDT.Rows.Count; i++)
                    {
                        DataRow NewRow = TPSReportDataTable.NewRow();

                        NewRow["UNN"] = UNN;
                        NewRow["PaymentRate"] = PaymentRate;
                        NewRow["AccountingName"] = TDT.Rows[i]["AccountingName"].ToString();
                        NewRow["InvNumber"] = TDT.Rows[i]["InvNumber"].ToString();
                        NewRow["CurrencyCode"] = ProfilCurrencyCode;
                        NewRow["TPSCurCode"] = TPSCurrencyCode;
                        if (TDT.Rows[i]["ProductID"].ToString() == "10" || TDT.Rows[i]["ProductID"].ToString() == "11" ||
                            TDT.Rows[i]["ProductID"].ToString() == "12")
                            NewRow["Name"] = GetProductName(Convert.ToInt32(TDT.Rows[i]["ProductID"]));
                        else
                            NewRow["Name"] = GetProductName(Convert.ToInt32(TDT.Rows[i]["ProductID"])) + " " +
                                             GetItemName(DecorID);

                        NewRow["Count"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[i]["Count"]), 3, MidpointRounding.AwayFromZero);
                        NewRow["Measure"] = "м.кв.";
                        NewRow["OriginalPrice"] = Convert.ToDecimal(TDT.Rows[i]["OriginalPrice"]);
                        NewRow["PriceWithTransport"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[i]["CostWithTransport"]) / Convert.ToDecimal(TDT.Rows[i]["Count"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["CostWithTransport"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[i]["CostWithTransport"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["DiscountVolume"] = Convert.ToDecimal(TDT.Rows[i]["DiscountVolume"]);
                        NewRow["TotalDiscount"] = Convert.ToDecimal(TDT.Rows[i]["TotalDiscount"]);
                        NewRow["Price"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[i]["Cost"]) / Convert.ToDecimal(TDT.Rows[i]["Count"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["Cost"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[i]["Cost"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["Weight"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[i]["Weight"]), 3, MidpointRounding.AwayFromZero);

                        TPSReportDataTable.Rows.Add(NewRow);
                    }
                }
            }

            PDT.Dispose();
            TDT.Dispose();
        }

        private void GetParametrizedData(DataRow[] Rows, DataTable PDT, DataTable TDT, int DecorID)
        {
            string p1 = "";
            string p2 = "";
            string p3 = "";

            if (PDT.Columns["Height"] != null)
                p1 = "Height";

            if (PDT.Columns["Length"] != null)
                p1 = "Length";

            if (PDT.Columns["Width"] != null)
                p2 = "Width";

            if (p1.Length > 0 && p2.Length == 0)
                p3 = p1;

            if (p1.Length == 0 && p2.Length > 0)
                p3 = p2;



            if (p1.Length == 0 && p2.Length == 0 && p3.Length == 0)
            {
                for (int r = 0; r < Rows.Count(); r++)
                {
                    if (IsProfil(Convert.ToInt32(Rows[r]["DecorConfigID"])))
                    {
                        if (PDT.Rows.Count == 0)
                        {
                            DataRow NewRow = PDT.NewRow();

                            NewRow["UNN"] = UNN;
                            NewRow["PaymentRate"] = PaymentRate;
                            NewRow["AccountingName"] = Rows[r]["AccountingName"].ToString();
                            NewRow["InvNumber"] = Rows[r]["InvNumber"].ToString();
                            NewRow["CurrencyCode"] = ProfilCurrencyCode;
                            NewRow["ProductID"] = Rows[r]["ProductID"];
                            NewRow["DecorID"] = Rows[r]["DecorID"];
                            NewRow["Count"] = Convert.ToDecimal(Rows[r]["Count"]);
                            NewRow["Price"] = Convert.ToDecimal(Rows[r]["Price"]);
                            NewRow["Cost"] = Convert.ToDecimal(Rows[r]["Cost"]);
                            NewRow["OriginalPrice"] = Convert.ToDecimal(Rows[r]["OriginalPrice"]);
                            NewRow["PriceWithTransport"] = Convert.ToDecimal(Rows[r]["PriceWithTransport"]);
                            NewRow["CostWithTransport"] = Convert.ToDecimal(Rows[r]["CostWithTransport"]);
                            NewRow["DiscountVolume"] = Convert.ToDecimal(Rows[r]["DiscountVolume"]);
                            NewRow["TotalCount"] = Convert.ToDecimal(Rows[r]["Count"]);
                            NewRow["TotalDiscount"] = Convert.ToDecimal(Rows[r]["TotalDiscount"]);
                            NewRow["Weight"] = GetDecorWeight(Rows[r]);
                            PDT.Rows.Add(NewRow);
                            continue;
                        }
                        else
                        {
                            PDT.Rows[0]["OriginalPrice"] = (Convert.ToDecimal(PDT.Rows[0]["OriginalPrice"]) * Convert.ToDecimal(PDT.Rows[0]["TotalCount"]) + Convert.ToDecimal(Rows[r]["Count"]) * Convert.ToDecimal(Rows[r]["OriginalPrice"])) /
                                (Convert.ToDecimal(PDT.Rows[0]["TotalCount"]) + Convert.ToDecimal(Rows[r]["Count"]));
                            PDT.Rows[0]["Price"] = (Convert.ToDecimal(PDT.Rows[0]["Price"]) * Convert.ToDecimal(PDT.Rows[0]["TotalCount"]) + Convert.ToDecimal(Rows[r]["Count"]) * Convert.ToDecimal(Rows[r]["Price"])) /
                                (Convert.ToDecimal(PDT.Rows[0]["TotalCount"]) + Convert.ToDecimal(Rows[r]["Count"]));
                            PDT.Rows[0]["TotalCount"] = Convert.ToDecimal(PDT.Rows[0]["TotalCount"]) +
                                Convert.ToDecimal(Rows[r]["Count"]);

                            PDT.Rows[0]["Count"] = Convert.ToDecimal(PDT.Rows[0]["Count"]) +
                                                            Convert.ToDecimal(Rows[r]["Count"]);
                            PDT.Rows[0]["Price"] = Convert.ToDecimal(Rows[r]["Price"]);
                            PDT.Rows[0]["Cost"] = Convert.ToDecimal(PDT.Rows[0]["Cost"]) +
                                                            Convert.ToDecimal(Rows[r]["Cost"]);
                            PDT.Rows[0]["CostWithTransport"] = Convert.ToDecimal(PDT.Rows[0]["CostWithTransport"]) + Convert.ToDecimal(Rows[r]["CostWithTransport"]);
                            PDT.Rows[0]["Weight"] = Convert.ToDecimal(PDT.Rows[0]["Weight"]) +
                                                            GetDecorWeight(Rows[r]);
                            continue;
                        }
                    }
                    else
                    {
                        if (TDT.Rows.Count == 0)
                        {
                            DataRow NewRow = TDT.NewRow();

                            NewRow["UNN"] = UNN;
                            NewRow["PaymentRate"] = PaymentRate;
                            NewRow["AccountingName"] = Rows[r]["AccountingName"].ToString();
                            NewRow["InvNumber"] = Rows[r]["InvNumber"].ToString();
                            NewRow["CurrencyCode"] = ProfilCurrencyCode;
                            NewRow["TPSCurCode"] = TPSCurrencyCode;
                            NewRow["ProductID"] = Rows[r]["ProductID"];
                            NewRow["DecorID"] = Rows[r]["DecorID"];
                            NewRow["Count"] = Convert.ToDecimal(Rows[r]["Count"]);
                            NewRow["Price"] = Convert.ToDecimal(Rows[r]["Price"]);
                            NewRow["Cost"] = Convert.ToDecimal(Rows[r]["Cost"]);
                            NewRow["OriginalPrice"] = Convert.ToDecimal(Rows[r]["OriginalPrice"]);
                            NewRow["PriceWithTransport"] = Convert.ToDecimal(Rows[r]["PriceWithTransport"]);
                            NewRow["CostWithTransport"] = Convert.ToDecimal(Rows[r]["CostWithTransport"]);
                            NewRow["DiscountVolume"] = Convert.ToDecimal(Rows[r]["DiscountVolume"]);
                            NewRow["TotalCount"] = Convert.ToDecimal(Rows[r]["Count"]);
                            NewRow["TotalDiscount"] = Convert.ToDecimal(Rows[r]["TotalDiscount"]);
                            NewRow["Weight"] = GetDecorWeight(Rows[r]);
                            TDT.Rows.Add(NewRow);
                            continue;
                        }
                        else
                        {
                            TDT.Rows[0]["OriginalPrice"] = (Convert.ToDecimal(TDT.Rows[0]["OriginalPrice"]) * Convert.ToDecimal(TDT.Rows[0]["TotalCount"]) + Convert.ToDecimal(Rows[r]["Count"]) * Convert.ToDecimal(Rows[r]["OriginalPrice"])) /
                                (Convert.ToDecimal(TDT.Rows[0]["TotalCount"]) + Convert.ToDecimal(Rows[r]["Count"]));
                            TDT.Rows[0]["Price"] = (Convert.ToDecimal(TDT.Rows[0]["Price"]) * Convert.ToDecimal(TDT.Rows[0]["TotalCount"]) + Convert.ToDecimal(Rows[r]["Count"]) * Convert.ToDecimal(Rows[r]["Price"])) /
                                (Convert.ToDecimal(TDT.Rows[0]["TotalCount"]) + Convert.ToDecimal(Rows[r]["Count"]));
                            TDT.Rows[0]["TotalCount"] = Convert.ToDecimal(TDT.Rows[0]["TotalCount"]) +
                                Convert.ToDecimal(Rows[r]["Count"]);
                            TDT.Rows[0]["Count"] = Convert.ToDecimal(TDT.Rows[0]["Count"]) +
                                                            Convert.ToDecimal(Rows[r]["Count"]);
                            TDT.Rows[0]["Cost"] = Convert.ToDecimal(TDT.Rows[0]["Cost"]) +
                                                            Convert.ToDecimal(Rows[r]["Cost"]);
                            TDT.Rows[0]["CostWithTransport"] = Convert.ToDecimal(TDT.Rows[0]["CostWithTransport"]) + Convert.ToDecimal(Rows[r]["CostWithTransport"]);
                            TDT.Rows[0]["Weight"] = Convert.ToDecimal(TDT.Rows[0]["Weight"]) +
                                                            GetDecorWeight(Rows[r]);
                            break;
                        }
                    }
                }

            }



            if (p1.Length > 0 && p2.Length > 0)
            {
                for (int r = 0; r < Rows.Count(); r++)
                {
                    if (IsProfil(Convert.ToInt32(Rows[r]["DecorConfigID"])))
                    {
                        if (PDT.Rows.Count == 0)
                        {
                            DataRow NewRow = PDT.NewRow();
                            NewRow[p1] = Convert.ToDecimal(Rows[r][p1]);
                            NewRow[p2] = Convert.ToDecimal(Rows[r][p2]);

                            NewRow["UNN"] = UNN;
                            NewRow["PaymentRate"] = PaymentRate;
                            NewRow["AccountingName"] = Rows[r]["AccountingName"].ToString();
                            NewRow["InvNumber"] = Rows[r]["InvNumber"].ToString();
                            NewRow["CurrencyCode"] = ProfilCurrencyCode;
                            NewRow["ProductID"] = Rows[r]["ProductID"];
                            NewRow["DecorID"] = Rows[r]["DecorID"];
                            NewRow["Count"] = Convert.ToDecimal(Rows[r]["Count"]);
                            NewRow["Price"] = Convert.ToDecimal(Rows[r]["Price"]);
                            NewRow["Cost"] = Convert.ToDecimal(Rows[r]["Cost"]);
                            NewRow["OriginalPrice"] = Convert.ToDecimal(Rows[r]["OriginalPrice"]);
                            NewRow["PriceWithTransport"] = Convert.ToDecimal(Rows[r]["PriceWithTransport"]);
                            NewRow["CostWithTransport"] = Convert.ToDecimal(Rows[r]["CostWithTransport"]);
                            NewRow["DiscountVolume"] = Convert.ToDecimal(Rows[r]["DiscountVolume"]);
                            NewRow["TotalCount"] = Convert.ToDecimal(Rows[r]["Count"]);
                            NewRow["TotalDiscount"] = Convert.ToDecimal(Rows[r]["TotalDiscount"]);
                            NewRow["Weight"] = GetDecorWeight(Rows[r]);
                            PDT.Rows.Add(NewRow);
                            continue;
                        }
                        else
                        {
                            if (Convert.ToInt32(PDT.Rows[0][p1]) == Convert.ToInt32(Rows[r][p1]) &&
                                        Convert.ToInt32(PDT.Rows[0][p2]) == Convert.ToInt32(Rows[r][p2]))
                            {
                                PDT.Rows[0]["OriginalPrice"] = (Convert.ToDecimal(PDT.Rows[0]["OriginalPrice"]) * Convert.ToDecimal(PDT.Rows[0]["TotalCount"]) + Convert.ToDecimal(Rows[r]["Count"]) * Convert.ToDecimal(Rows[r]["OriginalPrice"])) /
                                    (Convert.ToDecimal(PDT.Rows[0]["TotalCount"]) + Convert.ToDecimal(Rows[r]["Count"]));
                                PDT.Rows[0]["Price"] = (Convert.ToDecimal(PDT.Rows[0]["Price"]) * Convert.ToDecimal(PDT.Rows[0]["TotalCount"]) + Convert.ToDecimal(Rows[r]["Count"]) * Convert.ToDecimal(Rows[r]["Price"])) /
                                    (Convert.ToDecimal(PDT.Rows[0]["TotalCount"]) + Convert.ToDecimal(Rows[r]["Count"]));
                                PDT.Rows[0]["TotalCount"] = Convert.ToDecimal(PDT.Rows[0]["TotalCount"]) +
                                    Convert.ToDecimal(Rows[r]["Count"]);
                                PDT.Rows[0]["Count"] = Convert.ToDecimal(PDT.Rows[0]["Count"]) +
                                                              Convert.ToDecimal(Rows[r]["Count"]);
                                PDT.Rows[0]["Cost"] = Convert.ToDecimal(PDT.Rows[0]["Cost"]) +
                                                             Convert.ToDecimal(Rows[r]["Cost"]);
                                PDT.Rows[0]["CostWithTransport"] = Convert.ToDecimal(PDT.Rows[0]["CostWithTransport"]) + Convert.ToDecimal(Rows[r]["CostWithTransport"]);
                                PDT.Rows[0]["Weight"] = Convert.ToDecimal(PDT.Rows[0]["Weight"]) +
                                                             GetDecorWeight(Rows[r]);
                                continue;
                            }
                        }
                    }
                    else
                    {
                        if (TDT.Rows.Count == 0)
                        {
                            DataRow NewRow = TDT.NewRow();
                            NewRow[p1] = Convert.ToDecimal(Rows[r][p1]);
                            NewRow[p2] = Convert.ToDecimal(Rows[r][p2]);
                            NewRow["UNN"] = UNN;
                            NewRow["PaymentRate"] = PaymentRate;
                            NewRow["AccountingName"] = Rows[r]["AccountingName"].ToString();
                            NewRow["InvNumber"] = Rows[r]["InvNumber"].ToString();
                            NewRow["CurrencyCode"] = ProfilCurrencyCode;
                            NewRow["TPSCurCode"] = TPSCurrencyCode;
                            NewRow["ProductID"] = Rows[r]["ProductID"];
                            NewRow["DecorID"] = Rows[r]["DecorID"];
                            NewRow["Count"] = Convert.ToDecimal(Rows[r]["Count"]);
                            NewRow["Price"] = Convert.ToDecimal(Rows[r]["Price"]);
                            NewRow["Cost"] = Convert.ToDecimal(Rows[r]["Cost"]);
                            NewRow["OriginalPrice"] = Convert.ToDecimal(Rows[r]["OriginalPrice"]);
                            NewRow["PriceWithTransport"] = Convert.ToDecimal(Rows[r]["PriceWithTransport"]);
                            NewRow["CostWithTransport"] = Convert.ToDecimal(Rows[r]["CostWithTransport"]);
                            NewRow["DiscountVolume"] = Convert.ToDecimal(Rows[r]["DiscountVolume"]);
                            NewRow["TotalCount"] = Convert.ToDecimal(Rows[r]["Count"]);
                            NewRow["TotalDiscount"] = Convert.ToDecimal(Rows[r]["TotalDiscount"]);
                            NewRow["Weight"] = GetDecorWeight(Rows[r]);
                            TDT.Rows.Add(NewRow);
                            continue;
                        }
                        else
                        {
                            if (Convert.ToInt32(TDT.Rows[0][p1]) == Convert.ToInt32(Rows[r][p1]) &&
                                        Convert.ToInt32(TDT.Rows[0][p2]) == Convert.ToInt32(Rows[r][p2]))
                            {
                                TDT.Rows[0]["OriginalPrice"] = (Convert.ToDecimal(TDT.Rows[0]["OriginalPrice"]) * Convert.ToDecimal(TDT.Rows[0]["TotalCount"]) + Convert.ToDecimal(Rows[r]["Count"]) * Convert.ToDecimal(Rows[r]["OriginalPrice"])) /
                                    (Convert.ToDecimal(TDT.Rows[0]["TotalCount"]) + Convert.ToDecimal(Rows[r]["Count"]));
                                TDT.Rows[0]["Price"] = (Convert.ToDecimal(TDT.Rows[0]["Price"]) * Convert.ToDecimal(TDT.Rows[0]["TotalCount"]) + Convert.ToDecimal(Rows[r]["Count"]) * Convert.ToDecimal(Rows[r]["Price"])) /
                                    (Convert.ToDecimal(TDT.Rows[0]["TotalCount"]) + Convert.ToDecimal(Rows[r]["Count"]));
                                TDT.Rows[0]["TotalCount"] = Convert.ToDecimal(TDT.Rows[0]["TotalCount"]) +
                                    Convert.ToDecimal(Rows[r]["Count"]);
                                TDT.Rows[0]["Count"] = Convert.ToDecimal(TDT.Rows[0]["Count"]) +
                                                              Convert.ToDecimal(Rows[r]["Count"]);
                                TDT.Rows[0]["Cost"] = Convert.ToDecimal(TDT.Rows[0]["Cost"]) +
                                                             Convert.ToDecimal(Rows[r]["Cost"]);
                                TDT.Rows[0]["CostWithTransport"] = Convert.ToDecimal(TDT.Rows[0]["CostWithTransport"]) + Convert.ToDecimal(Rows[r]["CostWithTransport"]);
                                TDT.Rows[0]["Weight"] = Convert.ToDecimal(TDT.Rows[0]["Weight"]) +
                                                             GetDecorWeight(Rows[r]);
                                continue;
                            }
                        }
                    }
                }
            }

            if (p3.Length > 0)
            {
                for (int r = 0; r < Rows.Count(); r++)
                {
                    if (IsProfil(Convert.ToInt32(Rows[r]["DecorConfigID"])))
                    {
                        if (PDT.Rows.Count == 0)
                        {
                            DataRow NewRow = PDT.NewRow();
                            NewRow[p3] = Convert.ToDecimal(Rows[r][p3]);
                            NewRow["UNN"] = UNN;
                            NewRow["PaymentRate"] = PaymentRate;
                            NewRow["AccountingName"] = Rows[r]["AccountingName"].ToString();
                            NewRow["InvNumber"] = Rows[r]["InvNumber"].ToString();
                            NewRow["CurrencyCode"] = ProfilCurrencyCode;
                            NewRow["ProductID"] = Rows[r]["ProductID"];
                            NewRow["DecorID"] = Rows[r]["DecorID"];
                            NewRow["Count"] = Convert.ToDecimal(Rows[r]["Count"]);
                            NewRow["Price"] = Convert.ToDecimal(Rows[r]["Price"]);
                            NewRow["Cost"] = Convert.ToDecimal(Rows[r]["Cost"]);
                            NewRow["OriginalPrice"] = Convert.ToDecimal(Rows[r]["OriginalPrice"]);
                            NewRow["PriceWithTransport"] = Convert.ToDecimal(Rows[r]["PriceWithTransport"]);
                            NewRow["CostWithTransport"] = Convert.ToDecimal(Rows[r]["CostWithTransport"]);
                            NewRow["DiscountVolume"] = Convert.ToDecimal(Rows[r]["DiscountVolume"]);
                            NewRow["TotalCount"] = Convert.ToDecimal(Rows[r]["Count"]);
                            NewRow["TotalDiscount"] = Convert.ToDecimal(Rows[r]["TotalDiscount"]);
                            NewRow["Weight"] = GetDecorWeight(Rows[r]);
                            PDT.Rows.Add(NewRow);
                            continue;
                        }
                        else
                        {
                            if (Convert.ToInt32(PDT.Rows[0][p3]) == Convert.ToInt32(Rows[r][p3]))
                            {
                                PDT.Rows[0]["OriginalPrice"] = (Convert.ToDecimal(PDT.Rows[0]["OriginalPrice"]) * Convert.ToDecimal(PDT.Rows[0]["TotalCount"]) + Convert.ToDecimal(Rows[r]["Count"]) * Convert.ToDecimal(Rows[r]["OriginalPrice"])) /
                                    (Convert.ToDecimal(PDT.Rows[0]["TotalCount"]) + Convert.ToDecimal(Rows[r]["Count"]));
                                PDT.Rows[0]["Price"] = (Convert.ToDecimal(PDT.Rows[0]["Price"]) * Convert.ToDecimal(PDT.Rows[0]["TotalCount"]) + Convert.ToDecimal(Rows[r]["Count"]) * Convert.ToDecimal(Rows[r]["Price"])) /
                                    (Convert.ToDecimal(PDT.Rows[0]["TotalCount"]) + Convert.ToDecimal(Rows[r]["Count"]));
                                PDT.Rows[0]["TotalCount"] = Convert.ToDecimal(PDT.Rows[0]["TotalCount"]) +
                                    Convert.ToDecimal(Rows[r]["Count"]);
                                PDT.Rows[0]["Count"] = Convert.ToDecimal(PDT.Rows[0]["Count"]) +
                                                              Convert.ToDecimal(Rows[r]["Count"]);
                                PDT.Rows[0]["Cost"] = Convert.ToDecimal(PDT.Rows[0]["Cost"]) +
                                                             Convert.ToDecimal(Rows[r]["Cost"]);
                                PDT.Rows[0]["CostWithTransport"] = Convert.ToDecimal(PDT.Rows[0]["CostWithTransport"]) + Convert.ToDecimal(Rows[r]["CostWithTransport"]);
                                PDT.Rows[0]["Weight"] = Convert.ToDecimal(PDT.Rows[0]["Weight"]) +
                                                             GetDecorWeight(Rows[r]);
                                continue;
                            }
                        }
                    }
                    else
                    {
                        if (TDT.Rows.Count == 0)
                        {
                            DataRow NewRow = TDT.NewRow();
                            NewRow[p3] = Convert.ToDecimal(Rows[r][p3]);
                            NewRow["UNN"] = UNN;
                            NewRow["PaymentRate"] = PaymentRate;
                            NewRow["AccountingName"] = Rows[r]["AccountingName"].ToString();
                            NewRow["InvNumber"] = Rows[r]["InvNumber"].ToString();
                            NewRow["CurrencyCode"] = ProfilCurrencyCode;
                            NewRow["TPSCurCode"] = TPSCurrencyCode;
                            NewRow["ProductID"] = Rows[r]["ProductID"];
                            NewRow["DecorID"] = Rows[r]["DecorID"];
                            NewRow["Count"] = Convert.ToDecimal(Rows[r]["Count"]);
                            NewRow["Price"] = Convert.ToDecimal(Rows[r]["Price"]);
                            NewRow["Cost"] = Convert.ToDecimal(Rows[r]["Cost"]);
                            NewRow["OriginalPrice"] = Convert.ToDecimal(Rows[r]["OriginalPrice"]);
                            NewRow["PriceWithTransport"] = Convert.ToDecimal(Rows[r]["PriceWithTransport"]);
                            NewRow["CostWithTransport"] = Convert.ToDecimal(Rows[r]["CostWithTransport"]);
                            NewRow["DiscountVolume"] = Convert.ToDecimal(Rows[r]["DiscountVolume"]);
                            NewRow["TotalCount"] = Convert.ToDecimal(Rows[r]["Count"]);
                            NewRow["TotalDiscount"] = Convert.ToDecimal(Rows[r]["TotalDiscount"]);
                            NewRow["Weight"] = GetDecorWeight(Rows[r]);
                            TDT.Rows.Add(NewRow);
                            continue;
                        }
                        else
                        {
                            if (Convert.ToInt32(TDT.Rows[0][p3]) == Convert.ToInt32(Rows[r][p3]))
                            {
                                TDT.Rows[0]["OriginalPrice"] = (Convert.ToDecimal(TDT.Rows[0]["OriginalPrice"]) * Convert.ToDecimal(TDT.Rows[0]["TotalCount"]) + Convert.ToDecimal(Rows[r]["Count"]) * Convert.ToDecimal(Rows[r]["OriginalPrice"])) /
                                    (Convert.ToDecimal(TDT.Rows[0]["TotalCount"]) + Convert.ToDecimal(Rows[r]["Count"]));
                                TDT.Rows[0]["Price"] = (Convert.ToDecimal(TDT.Rows[0]["Price"]) * Convert.ToDecimal(TDT.Rows[0]["TotalCount"]) + Convert.ToDecimal(Rows[r]["Count"]) * Convert.ToDecimal(Rows[r]["Price"])) /
                                    (Convert.ToDecimal(TDT.Rows[0]["TotalCount"]) + Convert.ToDecimal(Rows[r]["Count"]));
                                TDT.Rows[0]["TotalCount"] = Convert.ToDecimal(TDT.Rows[0]["TotalCount"]) +
                                    Convert.ToDecimal(Rows[r]["Count"]);
                                TDT.Rows[0]["Count"] = Convert.ToDecimal(TDT.Rows[0]["Count"]) +
                                                              Convert.ToDecimal(Rows[r]["Count"]);
                                TDT.Rows[0]["Cost"] = Convert.ToDecimal(TDT.Rows[0]["Cost"]) +
                                                             Convert.ToDecimal(Rows[r]["Cost"]);
                                TDT.Rows[0]["CostWithTransport"] = Convert.ToDecimal(TDT.Rows[0]["CostWithTransport"]) + Convert.ToDecimal(Rows[r]["CostWithTransport"]);
                                TDT.Rows[0]["Weight"] = Convert.ToDecimal(TDT.Rows[0]["Weight"]) +
                                                             GetDecorWeight(Rows[r]);
                                continue;
                            }
                        }
                    }
                }
            }





            //REPORT TABLES
            if (PDT.Rows.Count > 0)
            {
                for (int g = 0; g < PDT.Rows.Count; g++)
                {
                    if (p1.Length > 0 && p2.Length > 0)
                    {
                        DataRow NewRow = ProfilReportDataTable.NewRow();

                        NewRow["UNN"] = UNN;
                        NewRow["PaymentRate"] = PaymentRate;
                        NewRow["AccountingName"] = PDT.Rows[g]["AccountingName"].ToString();
                        NewRow["InvNumber"] = PDT.Rows[g]["InvNumber"].ToString();
                        NewRow["CurrencyCode"] = ProfilCurrencyCode;
                        if (PDT.Rows[g]["ProductID"].ToString() == "10")
                            NewRow["Name"] = GetProductName(Convert.ToInt32(PDT.Rows[g]["ProductID"]));
                        else
                            NewRow["Name"] = GetProductName(Convert.ToInt32(PDT.Rows[g]["ProductID"])) + " " +
                                         GetItemName(Convert.ToInt32(DecorID)) + " " + PDT.Rows[g][p1] + "x" + PDT.Rows[g][p2];

                        NewRow["OriginalPrice"] = Convert.ToDecimal(PDT.Rows[g]["OriginalPrice"]);
                        NewRow["DiscountVolume"] = Convert.ToDecimal(PDT.Rows[g]["DiscountVolume"]);
                        NewRow["TotalDiscount"] = Convert.ToDecimal(PDT.Rows[g]["TotalDiscount"]);
                        NewRow["Count"] = PDT.Rows[g]["Count"];
                        NewRow["Measure"] = "шт.";
                        NewRow["Price"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[g]["Price"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["Cost"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[g]["Cost"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["PriceWithTransport"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[g]["CostWithTransport"]) / Convert.ToDecimal(PDT.Rows[g]["Count"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["CostWithTransport"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[g]["CostWithTransport"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["Weight"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[g]["Weight"]), 3, MidpointRounding.AwayFromZero);

                        ProfilReportDataTable.Rows.Add(NewRow);
                    }

                    if (p3.Length > 0)
                    {
                        DataRow NewRow = ProfilReportDataTable.NewRow();



                        NewRow["UNN"] = UNN;
                        NewRow["PaymentRate"] = PaymentRate;
                        NewRow["AccountingName"] = PDT.Rows[g]["AccountingName"].ToString();
                        NewRow["InvNumber"] = PDT.Rows[g]["InvNumber"].ToString();
                        NewRow["CurrencyCode"] = ProfilCurrencyCode;
                        if (PDT.Rows[g]["ProductID"].ToString() == "10")
                            NewRow["Name"] = GetProductName(Convert.ToInt32(PDT.Rows[g]["ProductID"]));
                        else
                            NewRow["Name"] = GetProductName(Convert.ToInt32(PDT.Rows[g]["ProductID"])) + " " +
                                         GetItemName(Convert.ToInt32(DecorID)) + " " + PDT.Rows[g][p3];

                        NewRow["OriginalPrice"] = Convert.ToDecimal(PDT.Rows[g]["OriginalPrice"]);
                        NewRow["DiscountVolume"] = Convert.ToDecimal(PDT.Rows[g]["DiscountVolume"]);
                        NewRow["TotalDiscount"] = Convert.ToDecimal(PDT.Rows[g]["TotalDiscount"]);
                        NewRow["Count"] = PDT.Rows[g]["Count"];
                        NewRow["Measure"] = "шт.";
                        NewRow["Price"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[g]["Price"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["Cost"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[g]["Cost"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["PriceWithTransport"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[g]["CostWithTransport"]) / Convert.ToDecimal(PDT.Rows[g]["Count"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["CostWithTransport"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[g]["CostWithTransport"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["Weight"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[g]["Weight"]), 3, MidpointRounding.AwayFromZero);

                        ProfilReportDataTable.Rows.Add(NewRow);
                    }

                    if (p1.Length == 0 && p2.Length == 0 && p3.Length == 0)
                    {
                        DataRow NewRow = ProfilReportDataTable.NewRow();

                        NewRow["UNN"] = UNN;
                        NewRow["PaymentRate"] = PaymentRate;
                        NewRow["AccountingName"] = PDT.Rows[g]["AccountingName"].ToString();
                        NewRow["InvNumber"] = PDT.Rows[g]["InvNumber"].ToString();
                        NewRow["CurrencyCode"] = ProfilCurrencyCode;
                        if (PDT.Rows[g]["ProductID"].ToString() == "10")
                            NewRow["Name"] = GetProductName(Convert.ToInt32(PDT.Rows[g]["ProductID"]));
                        else
                            NewRow["Name"] = GetProductName(Convert.ToInt32(PDT.Rows[g]["ProductID"])) + " " +
                                         GetItemName(Convert.ToInt32(DecorID));

                        NewRow["OriginalPrice"] = Convert.ToDecimal(PDT.Rows[g]["OriginalPrice"]);
                        NewRow["DiscountVolume"] = Convert.ToDecimal(PDT.Rows[g]["DiscountVolume"]);
                        NewRow["TotalDiscount"] = Convert.ToDecimal(PDT.Rows[g]["TotalDiscount"]);
                        NewRow["Count"] = PDT.Rows[g]["Count"];
                        NewRow["Measure"] = "шт.";
                        NewRow["Price"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[g]["Price"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["Cost"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[g]["Cost"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["PriceWithTransport"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[g]["CostWithTransport"]) / Convert.ToDecimal(PDT.Rows[g]["Count"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["CostWithTransport"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[g]["CostWithTransport"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["Weight"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[g]["Weight"]), 3, MidpointRounding.AwayFromZero);

                        ProfilReportDataTable.Rows.Add(NewRow);
                    }
                }
            }

            if (TDT.Rows.Count > 0)
            {
                for (int g = 0; g < TDT.Rows.Count; g++)
                {
                    if (p1.Length > 0 && p2.Length > 0)
                    {
                        DataRow NewRow = TPSReportDataTable.NewRow();

                        NewRow["UNN"] = UNN;
                        NewRow["PaymentRate"] = PaymentRate;
                        NewRow["AccountingName"] = TDT.Rows[g]["AccountingName"].ToString();
                        NewRow["InvNumber"] = TDT.Rows[g]["InvNumber"].ToString();
                        NewRow["CurrencyCode"] = ProfilCurrencyCode;
                        NewRow["TPSCurCode"] = TPSCurrencyCode;
                        if (TDT.Rows[g]["ProductID"].ToString() == "10")
                            NewRow["Name"] = GetProductName(Convert.ToInt32(TDT.Rows[g]["ProductID"]));
                        else
                            NewRow["Name"] = GetProductName(Convert.ToInt32(TDT.Rows[g]["ProductID"])) + " " +
                                         GetItemName(Convert.ToInt32(DecorID)) + " " + TDT.Rows[g][p1] + "x" + TDT.Rows[g][p2];

                        NewRow["OriginalPrice"] = Convert.ToDecimal(TDT.Rows[g]["OriginalPrice"]);
                        NewRow["DiscountVolume"] = Convert.ToDecimal(TDT.Rows[g]["DiscountVolume"]);
                        NewRow["TotalDiscount"] = Convert.ToDecimal(TDT.Rows[g]["TotalDiscount"]);
                        NewRow["Count"] = TDT.Rows[g]["Count"];
                        NewRow["Measure"] = "шт.";
                        NewRow["Price"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[g]["Price"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["Cost"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[g]["Cost"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["PriceWithTransport"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[g]["CostWithTransport"]) / Convert.ToDecimal(TDT.Rows[g]["Count"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["CostWithTransport"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[g]["CostWithTransport"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["Weight"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[g]["Weight"]), 3, MidpointRounding.AwayFromZero);

                        TPSReportDataTable.Rows.Add(NewRow);
                    }

                    if (p3.Length > 0)
                    {
                        DataRow NewRow = TPSReportDataTable.NewRow();

                        NewRow["UNN"] = UNN;
                        NewRow["PaymentRate"] = PaymentRate;
                        NewRow["AccountingName"] = TDT.Rows[g]["AccountingName"].ToString();
                        NewRow["InvNumber"] = TDT.Rows[g]["InvNumber"].ToString();
                        NewRow["CurrencyCode"] = ProfilCurrencyCode;
                        NewRow["TPSCurCode"] = TPSCurrencyCode;
                        if (TDT.Rows[g]["ProductID"].ToString() == "10")
                            NewRow["Name"] = GetProductName(Convert.ToInt32(TDT.Rows[g]["ProductID"]));
                        else
                            NewRow["Name"] = GetProductName(Convert.ToInt32(TDT.Rows[g]["ProductID"])) + " " +
                                         GetItemName(Convert.ToInt32(DecorID)) + " " + TDT.Rows[g][p3];

                        NewRow["OriginalPrice"] = Convert.ToDecimal(TDT.Rows[g]["OriginalPrice"]);
                        NewRow["DiscountVolume"] = Convert.ToDecimal(TDT.Rows[g]["DiscountVolume"]);
                        NewRow["TotalDiscount"] = Convert.ToDecimal(TDT.Rows[g]["TotalDiscount"]);
                        NewRow["Count"] = TDT.Rows[g]["Count"];
                        NewRow["Measure"] = "шт.";
                        NewRow["Price"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[g]["Price"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["Cost"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[g]["Cost"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["PriceWithTransport"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[g]["CostWithTransport"]) / Convert.ToDecimal(TDT.Rows[g]["Count"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["CostWithTransport"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[g]["CostWithTransport"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["Weight"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[g]["Weight"]), 3, MidpointRounding.AwayFromZero);

                        TPSReportDataTable.Rows.Add(NewRow);
                    }

                    if (p1.Length == 0 && p2.Length == 0 && p3.Length == 0)
                    {
                        DataRow NewRow = TPSReportDataTable.NewRow();

                        NewRow["UNN"] = UNN;
                        NewRow["PaymentRate"] = PaymentRate;
                        NewRow["AccountingName"] = TDT.Rows[g]["AccountingName"].ToString();
                        NewRow["InvNumber"] = TDT.Rows[g]["InvNumber"].ToString();
                        NewRow["CurrencyCode"] = ProfilCurrencyCode;
                        NewRow["TPSCurCode"] = TPSCurrencyCode;
                        if (TDT.Rows[g]["ProductID"].ToString() == "10")
                            NewRow["Name"] = GetProductName(Convert.ToInt32(TDT.Rows[g]["ProductID"]));
                        else
                            NewRow["Name"] = GetProductName(Convert.ToInt32(TDT.Rows[g]["ProductID"])) + " " +
                                         GetItemName(Convert.ToInt32(DecorID));

                        NewRow["OriginalPrice"] = Convert.ToDecimal(TDT.Rows[g]["OriginalPrice"]);
                        NewRow["DiscountVolume"] = Convert.ToDecimal(TDT.Rows[g]["DiscountVolume"]);
                        NewRow["TotalDiscount"] = Convert.ToDecimal(TDT.Rows[g]["TotalDiscount"]);
                        NewRow["Count"] = TDT.Rows[g]["Count"];
                        NewRow["Measure"] = "шт.";
                        NewRow["Price"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[g]["Price"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["Cost"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[g]["Cost"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["PriceWithTransport"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[g]["CostWithTransport"]) / Convert.ToDecimal(TDT.Rows[g]["Count"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["CostWithTransport"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[g]["CostWithTransport"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["Weight"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[g]["Weight"]), 3, MidpointRounding.AwayFromZero);

                        TPSReportDataTable.Rows.Add(NewRow);
                    }
                }
            }


        }

        private void Collect(bool ProfilVerify, bool TPSVerify, int ClientID, int DiscountPaymentConditionID)
        {
            DataTable DistRatesDT = new DataTable();
            DataTable Items = new DataTable();

            for (int i = 0; i < DecorOrdersDataTable.Rows.Count; i++)
            {
                if (DecorOrdersDataTable.Rows[i]["FactoryID"].ToString() == "1")//profil
                {
                    //if (ClientID != 145 && DiscountPaymentConditionID != 6 && !ProfilVerify)
                    //{
                    //    DecorOrdersDataTable.Rows[i]["PaymentRate"] = Convert.ToDecimal(DecorOrdersDataTable.Rows[i]["Rate"]) * 1.05m;
                    //}
                }
                if (DecorOrdersDataTable.Rows[i]["FactoryID"].ToString() == "2")//tps
                {
                    //if (ClientID != 145 && DiscountPaymentConditionID != 6 && !TPSVerify)
                    //{
                    //    DecorOrdersDataTable.Rows[i]["PaymentRate"] = Convert.ToDecimal(DecorOrdersDataTable.Rows[i]["Rate"]) * 1.05m;
                    //}
                }
            }
            using (DataView DV = new DataView(DecorOrdersDataTable))
            {
                Items = DV.ToTable(true, new string[] { "DecorID" });
            }

            //get count of different covertypes
            using (DataView DV = new DataView(DecorOrdersDataTable))
            {
                DistRatesDT = DV.ToTable(true, new string[] { "PaymentRate" });
            }

            for (int i = 0; i < Items.Rows.Count; i++)
            {
                int rr = Convert.ToInt32(Items.Rows[i]["DecorID"]);

                for (int j = 0; j < DistRatesDT.Rows.Count; j++)
                {
                    PaymentRate = Convert.ToDecimal(DistRatesDT.Rows[j]["PaymentRate"]);
                    DataRow[] ItemsRows = DecorOrdersDataTable.Select("PaymentRate='" + PaymentRate.ToString() + "' AND DecorID = " + Items.Rows[i]["DecorID"].ToString(),
                                                                          "Price ASC");
                    if (ItemsRows.Count() == 0)
                        continue;
                    int DecorConfigID = Convert.ToInt32(ItemsRows[0]["DecorConfigID"]);
                    //м.п.
                    if (GetReportMeasureTypeID(Convert.ToInt32(ItemsRows[0]["DecorConfigID"])) == 2)
                    {
                        GroupCoverTypes(ItemsRows, 2, Convert.ToInt32(Items.Rows[i]["DecorID"]));
                    }


                    //шт.
                    if (GetReportMeasureTypeID(Convert.ToInt32(ItemsRows[0]["DecorConfigID"])) == 3)
                    {
                        DataTable ParamTableProfil = new DataTable();
                        DataTable ParamTableTPS = new DataTable();

                        DataRow[] DCs = DecorConfigDataTable.Select("DecorConfigID = " +
                                                                                        ItemsRows[0]["DecorConfigID"].ToString());

                        CreateParamsTable(DCs[0]["ReportParam"].ToString(), ParamTableProfil);
                        CreateParamsTable(DCs[0]["ReportParam"].ToString(), ParamTableTPS);

                        GetParametrizedData(ItemsRows, ParamTableProfil, ParamTableTPS, Convert.ToInt32(Items.Rows[i]["DecorID"]));

                        ParamTableProfil.Dispose();
                        ParamTableTPS.Dispose();
                    }


                    //м.кв.
                    if (GetReportMeasureTypeID(Convert.ToInt32(ItemsRows[0]["DecorConfigID"])) == 1)
                    {
                        GroupCoverTypes(ItemsRows, 1, Convert.ToInt32(Items.Rows[i]["DecorID"]));
                    }
                }
            }

            Items.Dispose();
        }

        public void Report(int[] DispatchID, int CurrencyTypeID, bool ProfilVerify, bool TPSVerify, int ClientID, int DiscountPaymentConditionID)
        {
            DataRow[] rows = CurrencyTypesDT.Select("CurrencyTypeID = " + CurrencyTypeID);
            if (rows.Count() > 0)
            {
                ProfilCurrencyCode = rows[0]["CurrencyCode"].ToString();
                TPSCurrencyCode = rows[0]["TPSCurrencyCode"].ToString();
            }
            string SelectCommand = "SELECT UNN FROM Clients WHERE ClientID = " + ClientID;
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingReferenceConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    if (DA.Fill(DT) > 0)
                    {
                        if (DT.Rows[0]["UNN"] != DBNull.Value)
                            UNN = DT.Rows[0]["UNN"].ToString();
                    }
                }
            }

            SelectCommand = @"SELECT DISTINCT PackageDetailID, PackageDetails.Count AS Count, (DecorOrders.Cost * PackageDetails.Count / DecorOrders.Count) AS Cost, (DecorOrders.CostWithTransport * PackageDetails.Count / DecorOrders.Count) AS CostWithTransport, DecorOrders.*, infiniu2_catalog.dbo.DecorConfig.AccountingName, infiniu2_catalog.dbo.DecorConfig.InvNumber, MegaOrders.Rate, MegaOrders.PaymentRate FROM PackageDetails
                INNER JOIN DecorOrders ON PackageDetails.OrderID = DecorOrders.DecorOrderID
                INNER JOIN MainOrders ON DecorOrders.MainOrderID = MainOrders.MainOrderID
                INNER JOIN MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID
                INNER JOIN  Packages ON PackageDetails.PackageID = Packages.PackageID AND Packages.ProductType = 1
                AND Packages.DispatchID IN (" + string.Join(",", DispatchID) + @")
                INNER JOIN infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID = infiniu2_catalog.dbo.DecorConfig.DecorConfigID
                WHERE InvNumber IS NOT NULL ORDER BY InvNumber";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(DecorOrdersDataTable);
            }

            Collect(ProfilVerify, TPSVerify, ClientID, DiscountPaymentConditionID);

        }

        public void Report(int[] DispatchID, int CurrencyTypeID,
            bool ProfilVerify, bool TPSVerify, int ClientID, int DiscountPaymentConditionID, bool IsSample)
        {
            DecorOrdersDataTable.Clear();
            ProfilReportDataTable.Clear();
            TPSReportDataTable.Clear();
            DataRow[] rows = CurrencyTypesDT.Select("CurrencyTypeID = " + CurrencyTypeID);
            if (rows.Count() > 0)
            {
                ProfilCurrencyCode = rows[0]["CurrencyCode"].ToString();
                TPSCurrencyCode = rows[0]["TPSCurrencyCode"].ToString();
            }
            string SelectCommand = "SELECT UNN FROM Clients WHERE ClientID = " + ClientID;
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingReferenceConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    if (DA.Fill(DT) > 0)
                    {
                        if (DT.Rows[0]["UNN"] != DBNull.Value)
                            UNN = DT.Rows[0]["UNN"].ToString();
                    }
                }
            }

            SelectCommand = @"SELECT DISTINCT PackageDetailID, PackageDetails.Count AS Count, (DecorOrders.Cost * PackageDetails.Count / DecorOrders.Count) AS Cost, (DecorOrders.CostWithTransport * PackageDetails.Count / DecorOrders.Count) AS CostWithTransport, DecorOrders.*, infiniu2_catalog.dbo.DecorConfig.AccountingName, infiniu2_catalog.dbo.DecorConfig.InvNumber, MegaOrders.Rate, MegaOrders.PaymentRate FROM PackageDetails
                INNER JOIN DecorOrders ON PackageDetails.OrderID = DecorOrders.DecorOrderID
                INNER JOIN MainOrders ON DecorOrders.MainOrderID = MainOrders.MainOrderID
                INNER JOIN MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID
                INNER JOIN  Packages ON PackageDetails.PackageID = Packages.PackageID AND Packages.ProductType = 1
                AND Packages.DispatchID IN (" + string.Join(",", DispatchID) + @")
                INNER JOIN infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID = infiniu2_catalog.dbo.DecorConfig.DecorConfigID
                WHERE DecorOrders.IsSample=1 AND InvNumber IS NOT NULL ORDER BY InvNumber";
            if (!IsSample)
                SelectCommand = @"SELECT DISTINCT PackageDetailID, PackageDetails.Count AS Count, (DecorOrders.Cost * PackageDetails.Count / DecorOrders.Count) AS Cost, (DecorOrders.CostWithTransport * PackageDetails.Count / DecorOrders.Count) AS CostWithTransport, DecorOrders.*, infiniu2_catalog.dbo.DecorConfig.AccountingName, infiniu2_catalog.dbo.DecorConfig.InvNumber, MegaOrders.Rate, MegaOrders.PaymentRate FROM PackageDetails
                INNER JOIN DecorOrders ON PackageDetails.OrderID = DecorOrders.DecorOrderID
                INNER JOIN MainOrders ON DecorOrders.MainOrderID = MainOrders.MainOrderID
                INNER JOIN MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID
                INNER JOIN  Packages ON PackageDetails.PackageID = Packages.PackageID AND Packages.ProductType = 1
                AND Packages.DispatchID IN (" + string.Join(",", DispatchID) + @")
                INNER JOIN infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID = infiniu2_catalog.dbo.DecorConfig.DecorConfigID
                WHERE DecorOrders.IsSample=0 AND InvNumber IS NOT NULL ORDER BY InvNumber";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(DecorOrdersDataTable);
            }

            Collect(ProfilVerify, TPSVerify, ClientID, DiscountPaymentConditionID);

        }

    }
}
