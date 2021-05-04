using Infinium.Modules.Marketing.Orders;

using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;

using System;
using System.Data;
using System.Data.SqlClient;
using System.Text;

namespace Infinium.Modules.ExpeditionMarketing.StandardReport
{
    public class Report
    {
        //string ReportFilePath = string.Empty;
        //HSSFWorkbook hssfworkbook;
        decimal VAT = 1.0m;
        public FrontsReport FrontsReport = null;
        public DecorReport DecorReport = null;

        public DataTable CurrencyTypesDataTable = null;
        public DataTable ProfilReportTable = null;
        public DataTable TPSReportTable = null;

        public Report(ref DecorCatalogOrder DecorCatalogOrder, ref FrontsCalculate FC)
        {
            FrontsReport = new FrontsReport(ref FC);
            DecorReport = new DecorReport(ref DecorCatalogOrder);

            CreateProfilReportTable();

            CurrencyTypesDataTable = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM CurrencyTypes",
                ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(CurrencyTypesDataTable);
            }
        }

        private void CreateProfilReportTable()
        {
            ProfilReportTable = new DataTable();

            ProfilReportTable.Columns.Add(new DataColumn("UNN", Type.GetType("System.String")));
            ProfilReportTable.Columns.Add(new DataColumn("CurrencyCode", Type.GetType("System.String")));
            ProfilReportTable.Columns.Add(new DataColumn("TPSCurCode", Type.GetType("System.String")));
            ProfilReportTable.Columns.Add(new DataColumn("InvNumber", Type.GetType("System.String")));
            ProfilReportTable.Columns.Add(new DataColumn("AccountingName", Type.GetType("System.String")));
            ProfilReportTable.Columns.Add(new DataColumn("Name", Type.GetType("System.String")));
            ProfilReportTable.Columns.Add(new DataColumn("Count", Type.GetType("System.Decimal")));
            ProfilReportTable.Columns.Add(new DataColumn("Measure", Type.GetType("System.String")));
            ProfilReportTable.Columns.Add(new DataColumn("Price", Type.GetType("System.Decimal")));
            ProfilReportTable.Columns.Add(new DataColumn("Cost", Type.GetType("System.Decimal")));
            ProfilReportTable.Columns.Add(new DataColumn("Weight", Type.GetType("System.Decimal")));
            ProfilReportTable.Columns.Add(new DataColumn("OriginalPrice", Type.GetType("System.String")));
            ProfilReportTable.Columns.Add(new DataColumn("PriceWithTransport", Type.GetType("System.String")));
            ProfilReportTable.Columns.Add(new DataColumn("CostWithTransport", Type.GetType("System.String")));
            ProfilReportTable.Columns.Add(new DataColumn("DiscountVolume", Type.GetType("System.String")));
            ProfilReportTable.Columns.Add(new DataColumn("TotalDiscount", Type.GetType("System.String")));
            ProfilReportTable.Columns.Add(new DataColumn("PaymentRate", Type.GetType("System.Decimal")));
            ProfilReportTable.Columns.Add(new DataColumn("OriginalCount", Type.GetType("System.Decimal")));
            ProfilReportTable.Columns.Add(new DataColumn("OriginalCost", Type.GetType("System.Decimal")));
            ProfilReportTable.Columns.Add(new DataColumn("IsNonStandard", Type.GetType("System.Boolean")));
            ProfilReportTable.Columns.Add(new DataColumn("NonStandardMargin", Type.GetType("System.Decimal")));
            TPSReportTable = ProfilReportTable.Clone();
        }

        public void ConvertToCurrency(int CurrencyTypeID, bool IsNonStandard)
        {
            decimal Cost = 0;
            decimal Count = 0;
            decimal OriginalPrice = 0;
            decimal PriceWithTransport = 0;
            decimal CostWithTransport = 0;
            decimal Price = 0;
            decimal PaymentRate = 0;
            int DecCount = 2;
            for (int i = 0; i < ProfilReportTable.Rows.Count; i++)
            {
                if (IsNonStandard)
                {
                    if (ProfilReportTable.Rows[i]["IsNonStandard"] == DBNull.Value ||
                        (ProfilReportTable.Rows[i]["IsNonStandard"] != DBNull.Value && !Convert.ToBoolean(ProfilReportTable.Rows[i]["IsNonStandard"])))
                        continue;
                }
                if (!IsNonStandard)
                {
                    if (ProfilReportTable.Rows[i]["IsNonStandard"] != DBNull.Value && Convert.ToBoolean(ProfilReportTable.Rows[i]["IsNonStandard"]))
                        continue;
                }

                PaymentRate = Convert.ToDecimal(ProfilReportTable.Rows[i]["PaymentRate"]);
                Count = Convert.ToDecimal(ProfilReportTable.Rows[i]["Count"]);
                Cost = Convert.ToDecimal(ProfilReportTable.Rows[i]["Cost"]) * PaymentRate / VAT;
                if (ProfilReportTable.Rows[i]["OriginalPrice"] != DBNull.Value)
                    OriginalPrice = Convert.ToDecimal(ProfilReportTable.Rows[i]["OriginalPrice"]) * PaymentRate / VAT;
                if (ProfilReportTable.Rows[i]["PriceWithTransport"] != DBNull.Value)
                    PriceWithTransport = Convert.ToDecimal(ProfilReportTable.Rows[i]["PriceWithTransport"]) * PaymentRate / VAT;
                if (ProfilReportTable.Rows[i]["CostWithTransport"] != DBNull.Value)
                    CostWithTransport = Convert.ToDecimal(ProfilReportTable.Rows[i]["CostWithTransport"]) * PaymentRate / VAT;
                Cost = Math.Ceiling(Cost / 0.01m) * 0.01m;
                CostWithTransport = Math.Ceiling(CostWithTransport / 0.01m) * 0.01m;
                Price = Cost / Count;
                PriceWithTransport = CostWithTransport / Count;
                if (CurrencyTypeID == 0)
                {
                    DecCount = 0;
                    if (OriginalPrice != 0)
                        OriginalPrice = Math.Ceiling(OriginalPrice / 100.0m) * 100.0m;
                    if (CostWithTransport != 0)
                        CostWithTransport = Math.Ceiling(CostWithTransport / 100.0m) * 100.0m;
                    Price = Math.Ceiling(Price / 100.0m) * 100.0m;
                    Cost = Price * Count;
                    PriceWithTransport = Math.Ceiling(PriceWithTransport / 100.0m) * 100.0m;
                    CostWithTransport = PriceWithTransport * Count;
                }
                if (ProfilReportTable.Rows[i]["NonStandardMargin"] != DBNull.Value)
                {
                    decimal ExtraPrice = Convert.ToDecimal(ProfilReportTable.Rows[i]["NonStandardMargin"]);
                    Price = Price / (ExtraPrice / 100 + 1);
                    PriceWithTransport = PriceWithTransport / (ExtraPrice / 100 + 1);
                    ProfilReportTable.Rows[i]["NonStandardMargin"] = Convert.ToDecimal(ProfilReportTable.Rows[i]["NonStandardMargin"]);
                }
                ProfilReportTable.Rows[i]["OriginalPrice"] = Decimal.Round(OriginalPrice, DecCount, MidpointRounding.AwayFromZero);
                ProfilReportTable.Rows[i]["PriceWithTransport"] = Decimal.Round(PriceWithTransport, DecCount, MidpointRounding.AwayFromZero);
                ProfilReportTable.Rows[i]["CostWithTransport"] = Decimal.Round(CostWithTransport, DecCount, MidpointRounding.AwayFromZero);
                ProfilReportTable.Rows[i]["Price"] = Decimal.Round(Price, DecCount, MidpointRounding.AwayFromZero);
                ProfilReportTable.Rows[i]["Cost"] = Decimal.Round(Cost, DecCount, MidpointRounding.AwayFromZero);
            }

            for (int i = 0; i < TPSReportTable.Rows.Count; i++)
            {
                if (IsNonStandard)
                {
                    if (TPSReportTable.Rows[i]["IsNonStandard"] == DBNull.Value ||
                        (TPSReportTable.Rows[i]["IsNonStandard"] != DBNull.Value && !Convert.ToBoolean(TPSReportTable.Rows[i]["IsNonStandard"])))
                        continue;
                }
                if (!IsNonStandard)
                {
                    if (TPSReportTable.Rows[i]["IsNonStandard"] != DBNull.Value && Convert.ToBoolean(TPSReportTable.Rows[i]["IsNonStandard"]))
                        continue;
                }

                PaymentRate = Convert.ToDecimal(TPSReportTable.Rows[i]["PaymentRate"]);
                Count = Convert.ToDecimal(TPSReportTable.Rows[i]["Count"]);
                Cost = Convert.ToDecimal(TPSReportTable.Rows[i]["Cost"]) * PaymentRate / VAT;
                if (TPSReportTable.Rows[i]["OriginalPrice"] != DBNull.Value)
                    OriginalPrice = Convert.ToDecimal(TPSReportTable.Rows[i]["OriginalPrice"]) * PaymentRate / VAT;
                if (TPSReportTable.Rows[i]["PriceWithTransport"] != DBNull.Value)
                    PriceWithTransport = Convert.ToDecimal(TPSReportTable.Rows[i]["PriceWithTransport"]) * PaymentRate / VAT;
                if (TPSReportTable.Rows[i]["CostWithTransport"] != DBNull.Value)
                    CostWithTransport = Convert.ToDecimal(TPSReportTable.Rows[i]["CostWithTransport"]) * PaymentRate / VAT;
                Cost = Math.Ceiling(Cost / 0.01m) * 0.01m;
                CostWithTransport = Math.Ceiling(CostWithTransport / 0.01m) * 0.01m;
                Price = Cost / Count;
                PriceWithTransport = CostWithTransport / Count;
                if (CurrencyTypeID == 0)
                {
                    DecCount = 0;
                    if (OriginalPrice != 0)
                        OriginalPrice = Math.Ceiling(OriginalPrice / 100.0m) * 100.0m;
                    if (CostWithTransport != 0)
                        CostWithTransport = Math.Ceiling(CostWithTransport / 100.0m) * 100.0m;
                    Price = Math.Ceiling(Price / 100.0m) * 100.0m;
                    Cost = Price * Count;
                    PriceWithTransport = Math.Ceiling(PriceWithTransport / 100.0m) * 100.0m;
                    CostWithTransport = PriceWithTransport * Count;
                }
                if (TPSReportTable.Rows[i]["NonStandardMargin"] != DBNull.Value)
                {
                    decimal ExtraPrice = Convert.ToDecimal(TPSReportTable.Rows[i]["NonStandardMargin"]);
                    Price = Price / (ExtraPrice / 100 + 1);
                    PriceWithTransport = PriceWithTransport / (ExtraPrice / 100 + 1);
                    TPSReportTable.Rows[i]["NonStandardMargin"] = ExtraPrice = Convert.ToDecimal(TPSReportTable.Rows[i]["NonStandardMargin"]);
                }
                TPSReportTable.Rows[i]["OriginalPrice"] = Decimal.Round(OriginalPrice, DecCount, MidpointRounding.AwayFromZero);
                TPSReportTable.Rows[i]["PriceWithTransport"] = Decimal.Round(PriceWithTransport, DecCount, MidpointRounding.AwayFromZero);
                TPSReportTable.Rows[i]["CostWithTransport"] = Decimal.Round(CostWithTransport, DecCount, MidpointRounding.AwayFromZero);
                TPSReportTable.Rows[i]["Price"] = Decimal.Round(Price, DecCount, MidpointRounding.AwayFromZero);
                TPSReportTable.Rows[i]["Cost"] = Decimal.Round(Cost, DecCount, MidpointRounding.AwayFromZero);
            }
        }

        public void CollectGridsAndGlass()
        {
            decimal GridSquare = 0;
            decimal GridCost = 0;
            decimal GridCostWithTransport = 0;
            decimal GridWeight = 0;

            decimal LacomatSquare = 0;
            decimal LacomatCost = 0;
            decimal LacomatCostWithTransport = 0;
            decimal LacomatWeight = 0;
            decimal LacomatPrice = 0;
            decimal LacomatPriceWithTransport = 0;

            decimal KrizetSquare = 0;
            decimal KrizetCost = 0;
            decimal KrizetCostWithTransport = 0;
            decimal KrizetWeight = 0;
            decimal KrizetPrice = 0;
            decimal KrizetPriceWithTransport = 0;

            decimal FlutesSquare = 0;
            decimal FlutesCost = 0;
            decimal FlutesCostWithTransport = 0;
            decimal FlutesWeight = 0;
            decimal FlutesPrice = 0;
            decimal FlutesPriceWithTransport = 0;

            DataTable DistRatesDT = new DataTable();
            DataTable dt = ProfilReportTable.Clone();
            using (DataView DV = new DataView(ProfilReportTable))
            {
                DistRatesDT = DV.ToTable(true, new string[] { "PaymentRate" });
            }

            for (int j = 0; j < DistRatesDT.Rows.Count; j++)
            {
                decimal PaymentRate = Convert.ToDecimal(DistRatesDT.Rows[j]["PaymentRate"]);
                //collect

                GridSquare = 0;
                GridCost = 0;
                GridCostWithTransport = 0;
                GridWeight = 0;

                LacomatSquare = 0;
                LacomatCost = 0;
                LacomatCostWithTransport = 0;
                LacomatWeight = 0;
                LacomatPrice = 0;

                KrizetSquare = 0;
                KrizetCost = 0;
                KrizetCostWithTransport = 0;
                KrizetWeight = 0;
                KrizetPrice = 0;

                FlutesSquare = 0;
                FlutesCost = 0;
                FlutesCostWithTransport = 0;
                FlutesWeight = 0;
                FlutesPrice = 0;

                bool bb = false;
                for (int i = 0; i < ProfilReportTable.Rows.Count; i++)
                {
                    if (Convert.ToDecimal(ProfilReportTable.Rows[i]["PaymentRate"]) != PaymentRate)
                        continue;

                    bool b = false;
                    if (ProfilReportTable.Rows[i]["Name"].ToString().IndexOf("Решетка") > -1)
                    {
                        //GridCost += Convert.ToDecimal(ProfilReportTable.Rows[i]["Cost"]);
                        //GridCostWithTransport += Convert.ToDecimal(ProfilReportTable.Rows[i]["CostWithTransport"]);
                        GridSquare += Convert.ToDecimal(ProfilReportTable.Rows[i]["Count"]);
                        GridWeight += Convert.ToDecimal(ProfilReportTable.Rows[i]["Weight"]);

                        GridCost += Math.Ceiling(Convert.ToDecimal(ProfilReportTable.Rows[i]["OriginalPrice"]) * Convert.ToDecimal(ProfilReportTable.Rows[i]["Count"]) / 0.001m) * 0.001m;
                        GridCostWithTransport += Math.Ceiling(Convert.ToDecimal(ProfilReportTable.Rows[i]["PriceWithTransport"]) * Convert.ToDecimal(ProfilReportTable.Rows[i]["Count"]) / 0.001m) * 0.001m;

                        b = true;
                    }

                    if (ProfilReportTable.Rows[i]["Name"].ToString() == "Стекло Лакомат")
                    {
                        LacomatCost += Convert.ToDecimal(ProfilReportTable.Rows[i]["Cost"]);
                        LacomatCostWithTransport += Convert.ToDecimal(ProfilReportTable.Rows[i]["CostWithTransport"]);
                        LacomatSquare += Convert.ToDecimal(ProfilReportTable.Rows[i]["Count"]);
                        LacomatWeight += Convert.ToDecimal(ProfilReportTable.Rows[i]["Weight"]);
                        LacomatPrice = Convert.ToDecimal(ProfilReportTable.Rows[i]["Price"]);
                        LacomatPriceWithTransport = Convert.ToDecimal(ProfilReportTable.Rows[i]["PriceWithTransport"]);
                        b = true;
                    }

                    if (ProfilReportTable.Rows[i]["Name"].ToString() == "Стекло Кризет")
                    {
                        KrizetCost += Convert.ToDecimal(ProfilReportTable.Rows[i]["Cost"]);
                        KrizetCostWithTransport += Convert.ToDecimal(ProfilReportTable.Rows[i]["CostWithTransport"]);
                        KrizetSquare += Convert.ToDecimal(ProfilReportTable.Rows[i]["Count"]);
                        KrizetWeight += Convert.ToDecimal(ProfilReportTable.Rows[i]["Weight"]);
                        KrizetPrice = Convert.ToDecimal(ProfilReportTable.Rows[i]["Price"]);
                        KrizetPriceWithTransport = Convert.ToDecimal(ProfilReportTable.Rows[i]["PriceWithTransport"]);
                        b = true;
                    }

                    if (ProfilReportTable.Rows[i]["Name"].ToString() == "Стекло Флутес")
                    {
                        FlutesCost += Convert.ToDecimal(ProfilReportTable.Rows[i]["Cost"]);
                        FlutesCostWithTransport += Convert.ToDecimal(ProfilReportTable.Rows[i]["CostWithTransport"]);
                        FlutesSquare += Convert.ToDecimal(ProfilReportTable.Rows[i]["Count"]);
                        FlutesWeight += Convert.ToDecimal(ProfilReportTable.Rows[i]["Weight"]);
                        FlutesPrice = Convert.ToDecimal(ProfilReportTable.Rows[i]["Price"]);
                        FlutesPriceWithTransport = Convert.ToDecimal(ProfilReportTable.Rows[i]["PriceWithTransport"]);
                        b = true;
                    }

                    if (b)
                    {
                        ProfilReportTable.Rows[i].Delete();
                        ProfilReportTable.AcceptChanges();
                        i--;
                        bb = true;
                    }
                }

                if (bb)
                {
                    string AccountingName = string.Empty;
                    string InvNumber = string.Empty;
                    string UNN = string.Empty;
                    string ProfilCurrencyCode = string.Empty;
                    string TPSCurrencyCode = string.Empty;
                    //ADD to Table

                    if (ProfilReportTable.Rows.Count > 0)
                    {
                        AccountingName = ProfilReportTable.Rows[0]["AccountingName"].ToString();
                        InvNumber = ProfilReportTable.Rows[0]["InvNumber"].ToString();
                        UNN = ProfilReportTable.Rows[0]["UNN"].ToString();
                        ProfilCurrencyCode = ProfilReportTable.Rows[0]["CurrencyCode"].ToString();
                        TPSCurrencyCode = ProfilReportTable.Rows[0]["TPSCurCode"].ToString();
                    }
                    if (GridSquare > 0)
                    {
                        DataRow NewRow = dt.NewRow();
                        NewRow["UNN"] = UNN;
                        NewRow["PaymentRate"] = PaymentRate;
                        NewRow["AccountingName"] = AccountingName;
                        NewRow["InvNumber"] = InvNumber;
                        NewRow["CurrencyCode"] = ProfilCurrencyCode;
                        NewRow["TPSCurCode"] = TPSCurrencyCode;
                        NewRow["Name"] = "Решетка";
                        NewRow["Count"] = GridSquare;
                        NewRow["Measure"] = "м.кв.";
                        NewRow["Price"] = Decimal.Round(GridCost / GridSquare, 2, MidpointRounding.AwayFromZero);
                        NewRow["OriginalPrice"] = Decimal.Round(GridCost / GridSquare, 2, MidpointRounding.AwayFromZero);
                        NewRow["PriceWithTransport"] = Decimal.Round(GridCostWithTransport / GridSquare, 2, MidpointRounding.AwayFromZero);
                        NewRow["Cost"] = GridCost;
                        NewRow["CostWithTransport"] = GridCostWithTransport;
                        NewRow["Weight"] = Decimal.Round(GridWeight, 3, MidpointRounding.AwayFromZero);
                        dt.Rows.Add(NewRow);
                    }

                    GridSquare = 0;
                    GridCost = 0;
                    GridCost = 0;
                    GridWeight = 0;

                    if (LacomatSquare > 0)
                    {
                        DataRow NewRow = dt.NewRow();
                        NewRow["UNN"] = UNN;
                        NewRow["PaymentRate"] = PaymentRate;
                        NewRow["AccountingName"] = AccountingName;
                        NewRow["InvNumber"] = InvNumber;
                        NewRow["CurrencyCode"] = ProfilCurrencyCode;
                        NewRow["TPSCurCode"] = TPSCurrencyCode;
                        NewRow["Name"] = "Стекло Лакомат";
                        NewRow["Count"] = LacomatSquare;
                        NewRow["Measure"] = "м.кв.";
                        NewRow["Price"] = LacomatPrice;
                        NewRow["OriginalPrice"] = LacomatPrice;
                        NewRow["PriceWithTransport"] = LacomatPriceWithTransport;
                        NewRow["Cost"] = LacomatCost;
                        NewRow["CostWithTransport"] = LacomatCostWithTransport;
                        NewRow["Weight"] = Decimal.Round(LacomatWeight, 3, MidpointRounding.AwayFromZero);
                        dt.Rows.Add(NewRow);
                    }

                    LacomatSquare = 0;
                    LacomatCost = 0;
                    LacomatCost = 0;
                    LacomatWeight = 0;
                    LacomatPrice = 0;
                    LacomatPriceWithTransport = 0;

                    if (KrizetSquare > 0)
                    {
                        DataRow NewRow = dt.NewRow();
                        NewRow["UNN"] = UNN;
                        NewRow["PaymentRate"] = PaymentRate;
                        NewRow["AccountingName"] = AccountingName;
                        NewRow["InvNumber"] = InvNumber;
                        NewRow["CurrencyCode"] = ProfilCurrencyCode;
                        NewRow["TPSCurCode"] = TPSCurrencyCode;
                        NewRow["Name"] = "Стекло Кризет";
                        NewRow["Count"] = KrizetSquare;
                        NewRow["Measure"] = "м.кв.";
                        NewRow["Price"] = KrizetPrice;
                        NewRow["OriginalPrice"] = KrizetPrice;
                        NewRow["PriceWithTransport"] = KrizetPriceWithTransport;
                        NewRow["Cost"] = KrizetCost;
                        NewRow["CostWithTransport"] = KrizetCostWithTransport;
                        NewRow["Weight"] = Decimal.Round(KrizetWeight, 3, MidpointRounding.AwayFromZero);
                        dt.Rows.Add(NewRow);
                    }

                    KrizetSquare = 0;
                    KrizetCost = 0;
                    KrizetCostWithTransport = 0;
                    KrizetWeight = 0;
                    KrizetPrice = 0;
                    KrizetPriceWithTransport = 0;

                    if (FlutesSquare > 0)
                    {
                        DataRow NewRow = dt.NewRow();
                        NewRow["UNN"] = UNN;
                        NewRow["PaymentRate"] = PaymentRate;
                        NewRow["AccountingName"] = AccountingName;
                        NewRow["InvNumber"] = InvNumber;
                        NewRow["CurrencyCode"] = ProfilCurrencyCode;
                        NewRow["TPSCurCode"] = TPSCurrencyCode;
                        NewRow["Name"] = "Стекло Флутес";
                        NewRow["Count"] = FlutesSquare;
                        NewRow["Measure"] = "м.кв.";
                        NewRow["Price"] = FlutesPrice;
                        NewRow["OriginalPrice"] = FlutesPrice;
                        NewRow["PriceWithTransport"] = FlutesPriceWithTransport;
                        NewRow["Cost"] = FlutesCost;
                        NewRow["CostWithTransport"] = FlutesCostWithTransport;
                        NewRow["Weight"] = Decimal.Round(FlutesWeight, 3, MidpointRounding.AwayFromZero);
                        dt.Rows.Add(NewRow);
                    }

                    FlutesSquare = 0;
                    FlutesCost = 0;
                    FlutesCostWithTransport = 0;
                    FlutesWeight = 0;
                    FlutesPrice = 0;
                    FlutesPriceWithTransport = 0;
                }
            }
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                DataRow NewRow = ProfilReportTable.NewRow();
                NewRow["UNN"] = dt.Rows[i]["UNN"];
                NewRow["PaymentRate"] = dt.Rows[i]["PaymentRate"];
                NewRow["AccountingName"] = dt.Rows[i]["AccountingName"];
                NewRow["InvNumber"] = dt.Rows[i]["InvNumber"];
                NewRow["CurrencyCode"] = dt.Rows[i]["CurrencyCode"];
                NewRow["TPSCurCode"] = dt.Rows[i]["TPSCurCode"];
                NewRow["Name"] = dt.Rows[i]["Name"];
                NewRow["Count"] = dt.Rows[i]["Count"];
                NewRow["Measure"] = dt.Rows[i]["Measure"];
                NewRow["Price"] = dt.Rows[i]["Price"];
                NewRow["OriginalPrice"] = dt.Rows[i]["OriginalPrice"];
                NewRow["PriceWithTransport"] = dt.Rows[i]["PriceWithTransport"];
                NewRow["Cost"] = dt.Rows[i]["Cost"];
                NewRow["CostWithTransport"] = dt.Rows[i]["CostWithTransport"];
                NewRow["Weight"] = dt.Rows[i]["Weight"];
                ProfilReportTable.Rows.Add(NewRow);
            }

            DistRatesDT.Clear();
            dt.Clear();
            using (DataView DV = new DataView(TPSReportTable))
            {
                DistRatesDT = DV.ToTable(true, new string[] { "PaymentRate" });
            }

            for (int j = 0; j < DistRatesDT.Rows.Count; j++)
            {
                decimal PaymentRate = Convert.ToDecimal(DistRatesDT.Rows[j]["PaymentRate"]);
                //collect
                GridSquare = 0;
                GridCost = 0;
                GridCostWithTransport = 0;
                GridWeight = 0;

                LacomatSquare = 0;
                LacomatCost = 0;
                LacomatCostWithTransport = 0;
                LacomatWeight = 0;
                LacomatPrice = 0;
                LacomatPriceWithTransport = 0;

                KrizetSquare = 0;
                KrizetCost = 0;
                KrizetCostWithTransport = 0;
                KrizetWeight = 0;
                KrizetPrice = 0;
                KrizetPriceWithTransport = 0;

                FlutesSquare = 0;
                FlutesCost = 0;
                FlutesCostWithTransport = 0;
                FlutesWeight = 0;
                FlutesPrice = 0;
                FlutesPriceWithTransport = 0;

                bool bb = false;
                for (int i = 0; i < TPSReportTable.Rows.Count; i++)
                {
                    if (Convert.ToDecimal(TPSReportTable.Rows[i]["PaymentRate"]) != PaymentRate)
                        continue;

                    bool b = false;
                    if (TPSReportTable.Rows[i]["Name"].ToString().IndexOf("Решетка") > -1)
                    {
                        GridCost += Convert.ToDecimal(TPSReportTable.Rows[i]["Cost"]);
                        GridCostWithTransport += Convert.ToDecimal(TPSReportTable.Rows[i]["CostWithTransport"]);
                        GridSquare += Convert.ToDecimal(TPSReportTable.Rows[i]["Count"]);
                        GridWeight += Convert.ToDecimal(TPSReportTable.Rows[i]["Weight"]);
                        GridCost = 0;
                        GridCostWithTransport = 0;
                        GridCost += Math.Ceiling(Convert.ToDecimal(TPSReportTable.Rows[i]["OriginalPrice"]) * Convert.ToDecimal(TPSReportTable.Rows[i]["Count"]) / 0.001m) * 0.001m;
                        GridCostWithTransport += Math.Ceiling(Convert.ToDecimal(TPSReportTable.Rows[i]["PriceWithTransport"]) * Convert.ToDecimal(TPSReportTable.Rows[i]["Count"]) / 0.001m) * 0.001m;

                        b = true;
                    }

                    if (TPSReportTable.Rows[i]["Name"].ToString() == "Стекло Лакомат")
                    {
                        LacomatCost += Convert.ToDecimal(TPSReportTable.Rows[i]["Cost"]);
                        LacomatCostWithTransport += Convert.ToDecimal(TPSReportTable.Rows[i]["CostWithTransport"]);
                        LacomatSquare += Convert.ToDecimal(TPSReportTable.Rows[i]["Count"]);
                        LacomatWeight += Convert.ToDecimal(TPSReportTable.Rows[i]["Weight"]);
                        LacomatPrice = Convert.ToDecimal(TPSReportTable.Rows[i]["Price"]);
                        LacomatPriceWithTransport = Convert.ToDecimal(TPSReportTable.Rows[i]["PriceWithTransport"]);
                        b = true;
                    }

                    if (TPSReportTable.Rows[i]["Name"].ToString() == "Стекло Кризет")
                    {
                        KrizetCost += Convert.ToDecimal(TPSReportTable.Rows[i]["Cost"]);
                        KrizetCostWithTransport += Convert.ToDecimal(TPSReportTable.Rows[i]["CostWithTransport"]);
                        KrizetSquare += Convert.ToDecimal(TPSReportTable.Rows[i]["Count"]);
                        KrizetWeight += Convert.ToDecimal(TPSReportTable.Rows[i]["Weight"]);
                        KrizetPrice = Convert.ToDecimal(TPSReportTable.Rows[i]["Price"]);
                        KrizetPriceWithTransport = Convert.ToDecimal(TPSReportTable.Rows[i]["PriceWithTransport"]);
                        b = true;
                    }

                    if (TPSReportTable.Rows[i]["Name"].ToString() == "Стекло Флутес")
                    {
                        FlutesCost += Convert.ToDecimal(TPSReportTable.Rows[i]["Cost"]);
                        FlutesCostWithTransport += Convert.ToDecimal(TPSReportTable.Rows[i]["CostWithTransport"]);
                        FlutesSquare += Convert.ToDecimal(TPSReportTable.Rows[i]["Count"]);
                        FlutesWeight += Convert.ToDecimal(TPSReportTable.Rows[i]["Weight"]);
                        FlutesPrice = Convert.ToDecimal(TPSReportTable.Rows[i]["Price"]);
                        FlutesPriceWithTransport = Convert.ToDecimal(TPSReportTable.Rows[i]["PriceWithTransport"]);
                        b = true;
                    }

                    if (b)
                    {
                        TPSReportTable.Rows[i].Delete();
                        TPSReportTable.AcceptChanges();
                        i--;
                        bb = true;
                    }
                }

                if (bb)
                {
                    string AccountingName = string.Empty;
                    string InvNumber = string.Empty;
                    string UNN = string.Empty;
                    string ProfilCurrencyCode = string.Empty;
                    string TPSCurrencyCode = string.Empty;
                    //ADD to Table

                    if (TPSReportTable.Rows.Count > 0)
                    {
                        AccountingName = TPSReportTable.Rows[0]["AccountingName"].ToString();
                        InvNumber = TPSReportTable.Rows[0]["InvNumber"].ToString();
                        UNN = TPSReportTable.Rows[0]["UNN"].ToString();
                        ProfilCurrencyCode = TPSReportTable.Rows[0]["CurrencyCode"].ToString();
                        TPSCurrencyCode = TPSReportTable.Rows[0]["TPSCurCode"].ToString();
                    }

                    //ADD to Table
                    if (GridSquare > 0)
                    {
                        DataRow NewRow = dt.NewRow();
                        NewRow["UNN"] = UNN;
                        NewRow["PaymentRate"] = PaymentRate;
                        NewRow["AccountingName"] = AccountingName;
                        NewRow["InvNumber"] = InvNumber;
                        NewRow["CurrencyCode"] = ProfilCurrencyCode;
                        NewRow["TPSCurCode"] = TPSCurrencyCode;
                        NewRow["Name"] = "Решетка";
                        NewRow["Count"] = GridSquare;
                        NewRow["Measure"] = "м.кв.";
                        NewRow["OriginalPrice"] = Decimal.Round(GridCost / GridSquare, 2, MidpointRounding.AwayFromZero);
                        NewRow["Price"] = Decimal.Round(GridCost / GridSquare, 2, MidpointRounding.AwayFromZero);
                        NewRow["PriceWithTransport"] = Decimal.Round(GridCostWithTransport / GridSquare, 2, MidpointRounding.AwayFromZero);
                        NewRow["Cost"] = GridCost;
                        NewRow["CostWithTransport"] = GridCostWithTransport;
                        NewRow["Weight"] = Decimal.Round(GridWeight, 3, MidpointRounding.AwayFromZero);
                        dt.Rows.Add(NewRow);
                    }

                    if (LacomatSquare > 0)
                    {
                        DataRow NewRow = dt.NewRow();
                        NewRow["UNN"] = UNN;
                        NewRow["PaymentRate"] = PaymentRate;
                        NewRow["AccountingName"] = AccountingName;
                        NewRow["InvNumber"] = InvNumber;
                        NewRow["CurrencyCode"] = ProfilCurrencyCode;
                        NewRow["TPSCurCode"] = TPSCurrencyCode;
                        NewRow["Name"] = "Стекло Лакомат";
                        NewRow["Count"] = LacomatSquare;
                        NewRow["Measure"] = "м.кв.";
                        NewRow["Price"] = LacomatPrice;
                        NewRow["OriginalPrice"] = LacomatPrice;
                        NewRow["PriceWithTransport"] = LacomatPriceWithTransport;
                        NewRow["Cost"] = LacomatCost;
                        NewRow["CostWithTransport"] = LacomatCostWithTransport;
                        NewRow["Weight"] = Decimal.Round(LacomatWeight, 3, MidpointRounding.AwayFromZero);
                        dt.Rows.Add(NewRow);
                    }

                    if (KrizetSquare > 0)
                    {
                        DataRow NewRow = dt.NewRow();
                        NewRow["UNN"] = UNN;
                        NewRow["PaymentRate"] = PaymentRate;
                        NewRow["AccountingName"] = AccountingName;
                        NewRow["InvNumber"] = InvNumber;
                        NewRow["CurrencyCode"] = ProfilCurrencyCode;
                        NewRow["TPSCurCode"] = TPSCurrencyCode;
                        NewRow["Name"] = "Стекло Кризет";
                        NewRow["Count"] = KrizetSquare;
                        NewRow["Measure"] = "м.кв.";
                        NewRow["Price"] = KrizetPrice;
                        NewRow["OriginalPrice"] = KrizetPrice;
                        NewRow["PriceWithTransport"] = KrizetPriceWithTransport;
                        NewRow["Cost"] = KrizetCost;
                        NewRow["CostWithTransport"] = KrizetCostWithTransport;
                        NewRow["Weight"] = Decimal.Round(KrizetWeight, 3, MidpointRounding.AwayFromZero);
                        dt.Rows.Add(NewRow);
                    }

                    if (FlutesSquare > 0)
                    {
                        DataRow NewRow = dt.NewRow();
                        NewRow["UNN"] = UNN;
                        NewRow["PaymentRate"] = PaymentRate;
                        NewRow["AccountingName"] = AccountingName;
                        NewRow["InvNumber"] = InvNumber;
                        NewRow["CurrencyCode"] = ProfilCurrencyCode;
                        NewRow["TPSCurCode"] = TPSCurrencyCode;
                        NewRow["Name"] = "Стекло Флутес";
                        NewRow["Count"] = FlutesSquare;
                        NewRow["Measure"] = "м.кв.";
                        NewRow["Price"] = FlutesPrice;
                        NewRow["OriginalPrice"] = FlutesPrice;
                        NewRow["PriceWithTransport"] = FlutesPriceWithTransport;
                        NewRow["Cost"] = FlutesCost;
                        NewRow["CostWithTransport"] = FlutesCostWithTransport;
                        NewRow["Weight"] = Decimal.Round(FlutesWeight, 3, MidpointRounding.AwayFromZero);
                        dt.Rows.Add(NewRow);
                    }
                }
            }
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                DataRow NewRow = TPSReportTable.NewRow();
                NewRow["UNN"] = dt.Rows[i]["UNN"];
                NewRow["PaymentRate"] = dt.Rows[i]["PaymentRate"];
                NewRow["AccountingName"] = dt.Rows[i]["AccountingName"];
                NewRow["InvNumber"] = dt.Rows[i]["InvNumber"];
                NewRow["CurrencyCode"] = dt.Rows[i]["CurrencyCode"];
                NewRow["TPSCurCode"] = dt.Rows[i]["TPSCurCode"];
                NewRow["Name"] = dt.Rows[i]["Name"];
                NewRow["Count"] = dt.Rows[i]["Count"];
                NewRow["Measure"] = dt.Rows[i]["Measure"];
                NewRow["Price"] = dt.Rows[i]["Price"];
                NewRow["OriginalPrice"] = dt.Rows[i]["OriginalPrice"];
                NewRow["PriceWithTransport"] = dt.Rows[i]["PriceWithTransport"];
                NewRow["Cost"] = dt.Rows[i]["Cost"];
                NewRow["CostWithTransport"] = dt.Rows[i]["CostWithTransport"];
                NewRow["Weight"] = dt.Rows[i]["Weight"];
                TPSReportTable.Rows.Add(NewRow);
            }
        }

        public void AssignCost(decimal ComplaintProfilCost, decimal ComplaintTPSCost, decimal TransportCost, decimal AdditionalCost, decimal WeightProfil, decimal WeightTPS, decimal TotalWeight,
                               decimal TotalProfil, decimal TotalTPS, ref decimal TransportAndOtherProfil,
                               ref decimal TransportAndOtherTPS)
        {
            decimal Total = TransportCost + AdditionalCost;

            if (Total == 0 && ComplaintProfilCost == 0 && ComplaintTPSCost == 0)
                return;

            decimal pProfil = 0;
            decimal pTPS = 0;

            decimal cProfil = 0;
            decimal cTPS = 0;


            pProfil = WeightProfil / (TotalWeight / 100);
            pTPS = WeightTPS / (TotalWeight / 100);

            cProfil = Total / 100 * pProfil - ComplaintProfilCost;
            cTPS = Total / 100 * pTPS - ComplaintTPSCost;

            TransportAndOtherProfil = Decimal.Round(cProfil, 1, MidpointRounding.AwayFromZero);
            TransportAndOtherTPS = Decimal.Round(cTPS, 1, MidpointRounding.AwayFromZero);
        }

        private int GetMegaOrderID(int MainOrderID)
        {
            int MegaOrderID = 0;
            using (DataTable DT = new DataTable())
            {
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT MegaOrderID FROM MainOrders" +
                    " WHERE MainOrderID=" + MainOrderID, ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DA.Fill(DT);
                    if (DT.Rows.Count > 0)
                        MegaOrderID = Convert.ToInt32(DT.Rows[0]["MegaOrderID"]);
                }
            }
            return MegaOrderID;
        }

        private bool IsComplaint(int MegaOrderID)
        {
            bool IsComplaint = false;

            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT IsComplaint" +
                            " FROM MegaOrders WHERE MegaOrderID = " + MegaOrderID,
                            ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (SqlCommandBuilder CB = new SqlCommandBuilder(DA))
                {
                    using (DataTable DT = new DataTable())
                    {
                        if (DA.Fill(DT) == 0)
                            return false;

                        if (DBNull.Value != DT.Rows[0]["IsComplaint"] && Convert.ToInt32(DT.Rows[0]["IsComplaint"]) > 0)
                        {
                            IsComplaint = Convert.ToBoolean(DT.Rows[0]["IsComplaint"]);
                        }
                    }
                }
            }

            return IsComplaint;
        }

        private static string convertDefaultToDos(string src)
        {
            byte[] buffer;
            buffer = Encoding.Default.GetBytes(src);
            Encoding.Convert(Encoding.Default, Encoding.GetEncoding(866), buffer);
            return Encoding.Default.GetString(buffer);
        }

        private static string GetConnection(string path)
        {
            return "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + ";Extended Properties=dBASE 5.0;";
        }

        public void CreateReport(
            ref HSSFWorkbook hssfworkbook,
            ref HSSFSheet sheet1, int[] DispatchID, int[] OrderNumbers, int[] MainOrdersIDs, int ClientID, string ClientName,
            decimal ComplaintProfilCost, decimal ComplaintTPSCost, decimal TransportCost, decimal AdditionalCost,
            decimal TotalCost, int CurrencyTypeID, decimal TotalWeight, int pos, bool ProfilVerify, bool TPSVerify, int DiscountPaymentConditionID)
        {
            ClearReport();

            string MainOrdersList = string.Empty;

            string Currency = string.Empty;

            DataRow[] Row = CurrencyTypesDataTable.Select("CurrencyTypeID = " + CurrencyTypeID);

            Currency = Row[0]["CurrencyType"].ToString();
            //if (ClientID == 145 || ClientID == 258 || ClientID == 267)
            if (ClientID == 145)
            {
                VAT = 1.2m;
            }
            else
            {
                VAT = 1.0m;
            }
            TransportCost = TransportCost / VAT;
            AdditionalCost = AdditionalCost / VAT;
            FrontsReport.Report(DispatchID, CurrencyTypeID, ProfilVerify, TPSVerify, ClientID, DiscountPaymentConditionID);
            DecorReport.Report(DispatchID, CurrencyTypeID, ProfilVerify, TPSVerify, ClientID, DiscountPaymentConditionID);

            //PROFIL
            if (FrontsReport.ProfilReportDataTable.Rows.Count > 0)
            {
                for (int i = 0; i < FrontsReport.ProfilReportDataTable.Rows.Count; i++)
                {
                    ProfilReportTable.ImportRow(FrontsReport.ProfilReportDataTable.Rows[i]);
                }
            }

            if (DecorReport.ProfilReportDataTable.Rows.Count > 0)
            {
                for (int i = 0; i < DecorReport.ProfilReportDataTable.Rows.Count; i++)
                {
                    ProfilReportTable.ImportRow(DecorReport.ProfilReportDataTable.Rows[i]);
                }
            }


            //TPS
            if (FrontsReport.TPSReportDataTable.Rows.Count > 0)
            {
                for (int i = 0; i < FrontsReport.TPSReportDataTable.Rows.Count; i++)
                {
                    TPSReportTable.ImportRow(FrontsReport.TPSReportDataTable.Rows[i]);
                }
            }

            if (DecorReport.TPSReportDataTable.Rows.Count > 0)
            {
                for (int i = 0; i < DecorReport.TPSReportDataTable.Rows.Count; i++)
                {
                    TPSReportTable.ImportRow(DecorReport.TPSReportDataTable.Rows[i]);
                }
            }
            //F(@"D:\temp", "temp", TPSReportTable);
            CollectGridsAndGlass();
            ConvertToCurrency(CurrencyTypeID, false);
            ConvertToCurrency(CurrencyTypeID, true);

            decimal Total = TransportCost + AdditionalCost - ComplaintProfilCost - ComplaintTPSCost;
            decimal TotalProfil = 0;
            decimal TotalTPS = 0;
            decimal WeightProfil = 0;
            decimal WeightTPS = 0;

            decimal TransportAndOtherProfil = 0;
            decimal TransportAndOtherTPS = 0;

            for (int i = 0; i < ProfilReportTable.Rows.Count; i++)
            {
                TotalProfil += Convert.ToDecimal(ProfilReportTable.Rows[i]["CostWithTransport"]);
                WeightProfil += Convert.ToDecimal(ProfilReportTable.Rows[i]["Weight"]);
            }

            for (int i = 0; i < TPSReportTable.Rows.Count; i++)
            {
                TotalTPS += Convert.ToDecimal(TPSReportTable.Rows[i]["CostWithTransport"]);
                WeightTPS += Convert.ToDecimal(TPSReportTable.Rows[i]["Weight"]);
            }

            TotalCost = (TotalProfil + TotalTPS);

            //Assign COST
            AssignCost(ComplaintProfilCost, ComplaintTPSCost, TransportCost, AdditionalCost, WeightProfil, WeightTPS, TotalWeight, TotalProfil, TotalTPS, ref TransportAndOtherProfil,
                       ref TransportAndOtherTPS);

            decimal dd = Decimal.Round((WeightProfil + WeightTPS) / TotalWeight, 3, MidpointRounding.AwayFromZero);
            TotalProfil = Decimal.Round(TotalProfil, 3, MidpointRounding.AwayFromZero);
            TotalTPS = Decimal.Round(TotalTPS, 3, MidpointRounding.AwayFromZero);
            TransportCost = Decimal.Round(TransportCost * dd, 3, MidpointRounding.AwayFromZero);
            AdditionalCost = Decimal.Round(AdditionalCost * dd, 3, MidpointRounding.AwayFromZero);
            TotalCost = Decimal.Round(TotalCost, 3, MidpointRounding.AwayFromZero);

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
            SimpleF.FontHeightInPoints = 9;
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
            CountCS.DataFormat = hssfworkbook.CreateDataFormat().GetFormat("### ### ##0.000");
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

            HSSFCellStyle CurrencyCS = hssfworkbook.CreateCellStyle();
            CurrencyCS.DataFormat = hssfworkbook.CreateDataFormat().GetFormat("### ### ##0.0000");
            CurrencyCS.BorderBottom = HSSFCellStyle.BORDER_THIN;
            CurrencyCS.BottomBorderColor = HSSFColor.BLACK.index;
            CurrencyCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            CurrencyCS.LeftBorderColor = HSSFColor.BLACK.index;
            CurrencyCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            CurrencyCS.RightBorderColor = HSSFColor.BLACK.index;
            CurrencyCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            CurrencyCS.TopBorderColor = HSSFColor.BLACK.index;
            CurrencyCS.SetFont(SimpleF);

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

            #endregion

            HSSFCell Cell1 = sheet1.CreateRow(pos).CreateCell(0);
            Cell1.SetCellValue("Сводный отчет:");
            Cell1.CellStyle = ReportCS1;

            if (IsComplaint(GetMegaOrderID(MainOrdersIDs[0])))
            {
                Cell1 = sheet1.CreateRow(pos++).CreateCell(0);
                Cell1.SetCellValue("РЕКЛАМАЦИЯ");
                Cell1.CellStyle = ReportCS1;
            }

            int DisplayIndex = 0;
            if (ProfilReportTable.Rows.Count > 0)
            {
                //Профиль
                pos += 2;

                Cell1 = sheet1.CreateRow(pos++).CreateCell(0);
                Cell1.SetCellValue("ЗОВ-Профиль:");
                Cell1.CellStyle = SummaryWithoutBorderBelCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                Cell1.SetCellValue("Наименование");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                Cell1.SetCellValue("Кол-во");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                Cell1.SetCellValue("Ед. измер.");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                Cell1.SetCellValue("Цена нач, " + Currency);
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                Cell1.SetCellValue("Коэф.");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                Cell1.SetCellValue("Скидка, %");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                Cell1.SetCellValue("Цена с транс, " + Currency);
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                Cell1.SetCellValue("Нестандарт, %");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                Cell1.SetCellValue("Стоимость с транс, " + Currency);
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                Cell1.SetCellValue("Курс, " + Currency);
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos++).CreateCell(DisplayIndex++);
                Cell1.SetCellValue("Вес, кг.");
                Cell1.CellStyle = SimpleHeaderCS;

                for (int i = 0; i < ProfilReportTable.Rows.Count; i++)
                {
                    DisplayIndex = 0;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                    Cell1.SetCellValue(ProfilReportTable.Rows[i]["Name"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                    Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["Count"]));
                    Cell1.CellStyle = CountCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                    Cell1.SetCellValue(ProfilReportTable.Rows[i]["Measure"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    if (CurrencyTypeID == 0)
                    {
                        Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                        if (ProfilReportTable.Rows[i]["OriginalPrice"] != DBNull.Value)
                            Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["OriginalPrice"]));
                        Cell1.CellStyle = PriceBelCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                        if (ProfilReportTable.Rows[i]["DiscountVolume"] != DBNull.Value)
                            Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["DiscountVolume"]));
                        Cell1.CellStyle = WeightCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                        if (ProfilReportTable.Rows[i]["TotalDiscount"] != DBNull.Value)
                            Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["TotalDiscount"]));
                        Cell1.CellStyle = WeightCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["PriceWithTransport"]));
                        Cell1.CellStyle = PriceBelCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                        if (ProfilReportTable.Rows[i]["NonStandardMargin"] != DBNull.Value)
                        {
                            Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["NonStandardMargin"]));
                        }
                        Cell1.CellStyle = CountCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["CostWithTransport"]));
                        Cell1.CellStyle = PriceBelCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["PaymentRate"]));
                        Cell1.CellStyle = CurrencyCS;
                    }
                    else
                    {
                        Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                        if (ProfilReportTable.Rows[i]["OriginalPrice"] != DBNull.Value)
                            Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["OriginalPrice"]));
                        Cell1.CellStyle = PriceForeignCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                        if (ProfilReportTable.Rows[i]["DiscountVolume"] != DBNull.Value)
                            Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["DiscountVolume"]));
                        Cell1.CellStyle = PriceForeignCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                        if (ProfilReportTable.Rows[i]["TotalDiscount"] != DBNull.Value)
                            Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["TotalDiscount"]));
                        Cell1.CellStyle = PriceForeignCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["PriceWithTransport"]));
                        Cell1.CellStyle = PriceForeignCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                        if (ProfilReportTable.Rows[i]["NonStandardMargin"] != DBNull.Value)
                        {
                            Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["NonStandardMargin"]));
                        }
                        Cell1.CellStyle = CountCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["CostWithTransport"]));
                        Cell1.CellStyle = PriceForeignCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["PaymentRate"]));
                        Cell1.CellStyle = CurrencyCS;
                    }
                    Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                    Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["Weight"]));
                    Cell1.CellStyle = WeightCS;
                    pos++;
                }

                if (CurrencyTypeID == 0)
                {
                    Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                    Cell1.SetCellValue("Заказ, " + Currency + ":");
                    Cell1.CellStyle = SummaryWithoutBorderBelCS;
                    Cell1 = sheet1.CreateRow(pos++).CreateCell(8);
                    Cell1.SetCellValue(Convert.ToDouble(TotalProfil));
                    Cell1.CellStyle = SummaryWithoutBorderBelCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                    Cell1.SetCellValue("Транспорт, прочее, " + Currency + ":");
                    Cell1.CellStyle = SummaryWithoutBorderBelCS;
                    Cell1 = sheet1.CreateRow(pos++).CreateCell(8);
                    Cell1.SetCellValue(Convert.ToDouble(TransportAndOtherProfil));
                    Cell1.CellStyle = SummaryWithoutBorderBelCS;
                    Cell1 = sheet1.CreateRow(pos - 2).CreateCell(10);
                    Cell1.SetCellValue(Convert.ToDouble(WeightProfil));
                    Cell1.CellStyle = SummaryWeightCS;
                }
                else
                {
                    Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                    Cell1.SetCellValue("Заказ, " + Currency + ":");
                    Cell1.CellStyle = SummaryWithoutBorderForeignCS;
                    Cell1 = sheet1.CreateRow(pos++).CreateCell(8);
                    Cell1.SetCellValue(Convert.ToDouble(TotalProfil));
                    Cell1.CellStyle = SummaryWithoutBorderForeignCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                    Cell1.SetCellValue("Транспорт, прочее, " + Currency + ":");
                    Cell1.CellStyle = SummaryWithoutBorderForeignCS;
                    Cell1 = sheet1.CreateRow(pos++).CreateCell(8);
                    Cell1.SetCellValue(Convert.ToDouble(TransportAndOtherProfil));
                    Cell1.CellStyle = SummaryWithoutBorderForeignCS;
                    Cell1 = sheet1.CreateRow(pos - 2).CreateCell(10);
                    Cell1.SetCellValue(Convert.ToDouble(WeightProfil));
                    Cell1.CellStyle = SummaryWeightCS;
                }
            }

            DisplayIndex = 0;
            if (TPSReportTable.Rows.Count > 0)
            {
                //ТПС
                pos++;
                Cell1 = sheet1.CreateRow(pos++).CreateCell(0);
                Cell1.SetCellValue("ЗОВ-ТПС:");
                Cell1.CellStyle = SummaryWithoutBorderBelCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                Cell1.SetCellValue("Наименование");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                Cell1.SetCellValue("Кол-во");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                Cell1.SetCellValue("Ед. измер.");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                Cell1.SetCellValue("Цена нач, " + Currency);
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                Cell1.SetCellValue("Коэф.");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                Cell1.SetCellValue("Скидка, %");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                Cell1.SetCellValue("Цена с транс, " + Currency);
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                Cell1.SetCellValue("Нестандарт, %");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                Cell1.SetCellValue("Стоимость с транс, " + Currency);
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                Cell1.SetCellValue("Курс, " + Currency);
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos++).CreateCell(DisplayIndex++);
                Cell1.SetCellValue("Вес, кг.");
                Cell1.CellStyle = SimpleHeaderCS;

                for (int i = 0; i < TPSReportTable.Rows.Count; i++)
                {
                    DisplayIndex = 0;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                    Cell1.SetCellValue(TPSReportTable.Rows[i]["Name"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                    Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["Count"]));
                    Cell1.CellStyle = CountCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                    Cell1.SetCellValue(TPSReportTable.Rows[i]["Measure"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    if (CurrencyTypeID == 0)
                    {
                        Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                        if (TPSReportTable.Rows[i]["OriginalPrice"] != DBNull.Value)
                            Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["OriginalPrice"]));
                        Cell1.CellStyle = PriceBelCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                        if (TPSReportTable.Rows[i]["DiscountVolume"] != DBNull.Value)
                            Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["DiscountVolume"]));
                        Cell1.CellStyle = WeightCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                        if (TPSReportTable.Rows[i]["TotalDiscount"] != DBNull.Value)
                            Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["TotalDiscount"]));
                        Cell1.CellStyle = WeightCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["PriceWithTransport"]));
                        Cell1.CellStyle = PriceBelCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                        if (TPSReportTable.Rows[i]["NonStandardMargin"] != DBNull.Value)
                        {
                            Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["NonStandardMargin"]));
                        }
                        Cell1.CellStyle = CountCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["CostWithTransport"]));
                        Cell1.CellStyle = PriceBelCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["PaymentRate"]));
                        Cell1.CellStyle = CurrencyCS;
                    }
                    else
                    {
                        Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                        if (TPSReportTable.Rows[i]["OriginalPrice"] != DBNull.Value)
                            Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["OriginalPrice"]));
                        Cell1.CellStyle = PriceForeignCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                        if (TPSReportTable.Rows[i]["DiscountVolume"] != DBNull.Value)
                            Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["DiscountVolume"]));
                        Cell1.CellStyle = WeightCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                        if (TPSReportTable.Rows[i]["TotalDiscount"] != DBNull.Value)
                            Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["TotalDiscount"]));
                        Cell1.CellStyle = WeightCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["PriceWithTransport"]));
                        Cell1.CellStyle = PriceForeignCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                        if (TPSReportTable.Rows[i]["NonStandardMargin"] != DBNull.Value)
                        {
                            Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["NonStandardMargin"]));
                        }
                        Cell1.CellStyle = CountCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["CostWithTransport"]));
                        Cell1.CellStyle = PriceForeignCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["PaymentRate"]));
                        Cell1.CellStyle = CurrencyCS;
                    }
                    Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                    Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["Weight"]));
                    Cell1.CellStyle = WeightCS;
                    pos++;
                }

                if (CurrencyTypeID == 0)
                {
                    Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                    Cell1.SetCellValue("Заказ, " + Currency + ":");
                    Cell1.CellStyle = SummaryWithoutBorderBelCS;
                    Cell1 = sheet1.CreateRow(pos++).CreateCell(8);
                    Cell1.SetCellValue(Convert.ToDouble(TotalTPS));
                    Cell1.CellStyle = SummaryWithoutBorderBelCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                    Cell1.SetCellValue("Транспорт, прочее, " + Currency + ":");
                    Cell1.CellStyle = SummaryWithoutBorderBelCS;
                    Cell1 = sheet1.CreateRow(pos++).CreateCell(8);
                    Cell1.SetCellValue(Convert.ToDouble(TransportAndOtherTPS));
                    Cell1.CellStyle = SummaryWithoutBorderBelCS;
                    Cell1 = sheet1.CreateRow(pos - 2).CreateCell(10);
                    Cell1.SetCellValue(Convert.ToDouble(WeightTPS));
                    Cell1.CellStyle = SummaryWeightCS;
                }
                else
                {
                    Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                    Cell1.SetCellValue("Заказ, " + Currency + ":");
                    Cell1.CellStyle = SummaryWithoutBorderForeignCS;
                    Cell1 = sheet1.CreateRow(pos++).CreateCell(8);
                    Cell1.SetCellValue(Convert.ToDouble(TotalTPS));
                    Cell1.CellStyle = SummaryWithoutBorderForeignCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                    Cell1.SetCellValue("Транспорт, прочее, " + Currency + ":");
                    Cell1.CellStyle = SummaryWithoutBorderForeignCS;
                    Cell1 = sheet1.CreateRow(pos++).CreateCell(8);
                    Cell1.SetCellValue(Convert.ToDouble(TransportAndOtherTPS));
                    Cell1.CellStyle = SummaryWithoutBorderForeignCS;
                    Cell1 = sheet1.CreateRow(pos - 2).CreateCell(10);
                    Cell1.SetCellValue(Convert.ToDouble(WeightTPS));
                    Cell1.CellStyle = SummaryWeightCS;
                }
            }
            pos++;

            //if (Rate != 1)
            //{
            //    Cell1 = sheet1.CreateRow(pos++).CreateCell(0);
            //    Cell1.SetCellValue("Курс, 1 EUR = " + Rate + " " + Currency);
            //    Cell1.CellStyle = SummaryWithoutBorderBelCS;
            //    //Excel.WriteCell(1, "Курс, 1 EUR = " + Rate + " " + Currency, pos++, 1, 12, true);
            //}

            if (CurrencyTypeID == 0)
            {
                Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                Cell1.SetCellValue("Транспорт, " + Currency + ":");
                Cell1.CellStyle = SummaryWithBorderBelCS;

                Cell1 = sheet1.CreateRow(pos++).CreateCell(1);
                Cell1.SetCellValue(Convert.ToDouble(TransportCost));
                Cell1.CellStyle = SummaryWithBorderBelCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                Cell1.SetCellValue("Прочее, " + Currency + ":");
                Cell1.CellStyle = SummaryWithBorderBelCS;

                Cell1 = sheet1.CreateRow(pos++).CreateCell(1);
                Cell1.SetCellValue(Convert.ToDouble(AdditionalCost));
                Cell1.CellStyle = SummaryWithBorderBelCS;
            }
            else
            {
                Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                Cell1.SetCellValue("Транспорт, " + Currency + ":");
                Cell1.CellStyle = SummaryWithBorderForeignCS;

                Cell1 = sheet1.CreateRow(pos++).CreateCell(1);
                Cell1.SetCellValue(Convert.ToDouble(TransportCost));
                Cell1.CellStyle = SummaryWithBorderForeignCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                Cell1.SetCellValue("Прочее, " + Currency + ":");
                Cell1.CellStyle = SummaryWithBorderForeignCS;

                Cell1 = sheet1.CreateRow(pos++).CreateCell(1);
                Cell1.SetCellValue(Convert.ToDouble(AdditionalCost));
                Cell1.CellStyle = SummaryWithBorderForeignCS;
            }

            if (IsComplaint(GetMegaOrderID(MainOrdersIDs[0])))
            {
                TotalCost = 0;
            }

            if (CurrencyTypeID == 0)
            {
                Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                Cell1.SetCellValue("Итого, " + Currency + ":");
                Cell1.CellStyle = SummaryWithBorderBelCS;
                Cell1 = sheet1.CreateRow(pos++).CreateCell(1);
                Cell1.SetCellValue(Convert.ToDouble((TotalCost)));
                Cell1.CellStyle = SummaryWithBorderBelCS;
            }
            else
            {
                Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                Cell1.SetCellValue("Итого, " + Currency + ":");
                Cell1.CellStyle = SummaryWithBorderForeignCS;
                Cell1 = sheet1.CreateRow(pos++).CreateCell(1);
                Cell1.SetCellValue(Convert.ToDouble((TotalCost)));
                Cell1.CellStyle = SummaryWithBorderForeignCS;
            }

            ClearReport();
        }

        public void CreateReport(
            ref HSSFWorkbook hssfworkbook,
            ref HSSFSheet sheet1, int[] DispatchID, int[] OrderNumbers, int[] MainOrdersIDs, int ClientID, string ClientName,
            decimal ComplaintProfilCost, decimal ComplaintTPSCost, decimal TransportCost, decimal AdditionalCost,
            decimal TotalCost, int CurrencyTypeID, decimal TotalWeight, int pos, bool ProfilVerify, bool TPSVerify, int DiscountPaymentConditionID, bool IsSample)
        {
            ClearReport();

            string MainOrdersList = string.Empty;

            string Currency = string.Empty;

            DataRow[] Row = CurrencyTypesDataTable.Select("CurrencyTypeID = " + CurrencyTypeID);

            Currency = Row[0]["CurrencyType"].ToString();
            //if (ClientID == 145 || ClientID == 258 || ClientID == 267)
            if (ClientID == 145)
            {
                VAT = 1.2m;
            }
            else
            {
                VAT = 1.0m;
            }
            TransportCost = TransportCost / VAT;
            AdditionalCost = AdditionalCost / VAT;
            FrontsReport.Report(DispatchID, CurrencyTypeID, ProfilVerify, TPSVerify, ClientID, DiscountPaymentConditionID, IsSample);
            DecorReport.Report(DispatchID, CurrencyTypeID, ProfilVerify, TPSVerify, ClientID, DiscountPaymentConditionID, IsSample);

            //PROFIL
            if (FrontsReport.ProfilReportDataTable.Rows.Count > 0)
            {
                for (int i = 0; i < FrontsReport.ProfilReportDataTable.Rows.Count; i++)
                {
                    ProfilReportTable.ImportRow(FrontsReport.ProfilReportDataTable.Rows[i]);
                }
            }

            if (DecorReport.ProfilReportDataTable.Rows.Count > 0)
            {
                for (int i = 0; i < DecorReport.ProfilReportDataTable.Rows.Count; i++)
                {
                    ProfilReportTable.ImportRow(DecorReport.ProfilReportDataTable.Rows[i]);
                }
            }


            //TPS
            if (FrontsReport.TPSReportDataTable.Rows.Count > 0)
            {
                for (int i = 0; i < FrontsReport.TPSReportDataTable.Rows.Count; i++)
                {
                    TPSReportTable.ImportRow(FrontsReport.TPSReportDataTable.Rows[i]);
                }
            }

            if (DecorReport.TPSReportDataTable.Rows.Count > 0)
            {
                for (int i = 0; i < DecorReport.TPSReportDataTable.Rows.Count; i++)
                {
                    TPSReportTable.ImportRow(DecorReport.TPSReportDataTable.Rows[i]);
                }
            }
            //F(@"D:\temp", "temp", TPSReportTable);
            CollectGridsAndGlass();
            ConvertToCurrency(CurrencyTypeID, false);
            ConvertToCurrency(CurrencyTypeID, true);

            decimal Total = TransportCost + AdditionalCost - ComplaintProfilCost - ComplaintTPSCost;
            decimal TotalProfil = 0;
            decimal TotalTPS = 0;
            decimal WeightProfil = 0;
            decimal WeightTPS = 0;

            decimal TransportAndOtherProfil = 0;
            decimal TransportAndOtherTPS = 0;

            for (int i = 0; i < ProfilReportTable.Rows.Count; i++)
            {
                TotalProfil += Convert.ToDecimal(ProfilReportTable.Rows[i]["CostWithTransport"]);
                WeightProfil += Convert.ToDecimal(ProfilReportTable.Rows[i]["Weight"]);
            }

            for (int i = 0; i < TPSReportTable.Rows.Count; i++)
            {
                TotalTPS += Convert.ToDecimal(TPSReportTable.Rows[i]["CostWithTransport"]);
                WeightTPS += Convert.ToDecimal(TPSReportTable.Rows[i]["Weight"]);
            }

            TotalCost = (TotalProfil + TotalTPS);

            //Assign COST
            AssignCost(ComplaintProfilCost, ComplaintTPSCost, TransportCost, AdditionalCost, WeightProfil, WeightTPS, TotalWeight, TotalProfil, TotalTPS, ref TransportAndOtherProfil,
                       ref TransportAndOtherTPS);

            decimal dd = Decimal.Round((WeightProfil + WeightTPS) / TotalWeight, 3, MidpointRounding.AwayFromZero);
            TotalProfil = Decimal.Round(TotalProfil, 3, MidpointRounding.AwayFromZero);
            TotalTPS = Decimal.Round(TotalTPS, 3, MidpointRounding.AwayFromZero);
            TransportCost = Decimal.Round(TransportCost * dd, 3, MidpointRounding.AwayFromZero);
            AdditionalCost = Decimal.Round(AdditionalCost * dd, 3, MidpointRounding.AwayFromZero);
            TotalCost = Decimal.Round(TotalCost, 3, MidpointRounding.AwayFromZero);

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
            SimpleF.FontHeightInPoints = 9;
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
            CountCS.DataFormat = hssfworkbook.CreateDataFormat().GetFormat("### ### ##0.000");
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

            HSSFCellStyle CurrencyCS = hssfworkbook.CreateCellStyle();
            CurrencyCS.DataFormat = hssfworkbook.CreateDataFormat().GetFormat("### ### ##0.0000");
            CurrencyCS.BorderBottom = HSSFCellStyle.BORDER_THIN;
            CurrencyCS.BottomBorderColor = HSSFColor.BLACK.index;
            CurrencyCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            CurrencyCS.LeftBorderColor = HSSFColor.BLACK.index;
            CurrencyCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            CurrencyCS.RightBorderColor = HSSFColor.BLACK.index;
            CurrencyCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            CurrencyCS.TopBorderColor = HSSFColor.BLACK.index;
            CurrencyCS.SetFont(SimpleF);

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

            #endregion

            HSSFCell Cell1 = sheet1.CreateRow(pos).CreateCell(0);
            Cell1.SetCellValue("Сводный отчет:");
            Cell1.CellStyle = ReportCS1;

            if (IsComplaint(GetMegaOrderID(MainOrdersIDs[0])))
            {
                Cell1 = sheet1.CreateRow(pos++).CreateCell(0);
                Cell1.SetCellValue("РЕКЛАМАЦИЯ");
                Cell1.CellStyle = ReportCS1;
            }

            int DisplayIndex = 0;
            if (ProfilReportTable.Rows.Count > 0)
            {
                //Профиль
                pos += 2;

                Cell1 = sheet1.CreateRow(pos++).CreateCell(0);
                Cell1.SetCellValue("ЗОВ-Профиль:");
                Cell1.CellStyle = SummaryWithoutBorderBelCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                Cell1.SetCellValue("Наименование");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                Cell1.SetCellValue("Кол-во");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                Cell1.SetCellValue("Ед. измер.");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                Cell1.SetCellValue("Цена нач, " + Currency);
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                Cell1.SetCellValue("Коэф.");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                Cell1.SetCellValue("Скидка, %");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                Cell1.SetCellValue("Цена с транс, " + Currency);
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                Cell1.SetCellValue("Нестандарт, %");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                Cell1.SetCellValue("Стоимость с транс, " + Currency);
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                Cell1.SetCellValue("Курс, " + Currency);
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos++).CreateCell(DisplayIndex++);
                Cell1.SetCellValue("Вес, кг.");
                Cell1.CellStyle = SimpleHeaderCS;

                for (int i = 0; i < ProfilReportTable.Rows.Count; i++)
                {
                    DisplayIndex = 0;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                    Cell1.SetCellValue(ProfilReportTable.Rows[i]["Name"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                    Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["Count"]));
                    Cell1.CellStyle = CountCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                    Cell1.SetCellValue(ProfilReportTable.Rows[i]["Measure"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    if (CurrencyTypeID == 0)
                    {
                        Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                        if (ProfilReportTable.Rows[i]["OriginalPrice"] != DBNull.Value)
                            Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["OriginalPrice"]));
                        Cell1.CellStyle = PriceBelCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                        if (ProfilReportTable.Rows[i]["DiscountVolume"] != DBNull.Value)
                            Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["DiscountVolume"]));
                        Cell1.CellStyle = WeightCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                        if (ProfilReportTable.Rows[i]["TotalDiscount"] != DBNull.Value)
                            Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["TotalDiscount"]));
                        Cell1.CellStyle = WeightCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["PriceWithTransport"]));
                        Cell1.CellStyle = PriceBelCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                        if (ProfilReportTable.Rows[i]["NonStandardMargin"] != DBNull.Value)
                        {
                            Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["NonStandardMargin"]));
                        }
                        Cell1.CellStyle = CountCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["CostWithTransport"]));
                        Cell1.CellStyle = PriceBelCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["PaymentRate"]));
                        Cell1.CellStyle = CurrencyCS;
                    }
                    else
                    {
                        Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                        if (ProfilReportTable.Rows[i]["OriginalPrice"] != DBNull.Value)
                            Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["OriginalPrice"]));
                        Cell1.CellStyle = PriceForeignCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                        if (ProfilReportTable.Rows[i]["DiscountVolume"] != DBNull.Value)
                            Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["DiscountVolume"]));
                        Cell1.CellStyle = PriceForeignCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                        if (ProfilReportTable.Rows[i]["TotalDiscount"] != DBNull.Value)
                            Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["TotalDiscount"]));
                        Cell1.CellStyle = PriceForeignCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["PriceWithTransport"]));
                        Cell1.CellStyle = PriceForeignCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                        if (ProfilReportTable.Rows[i]["NonStandardMargin"] != DBNull.Value)
                        {
                            Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["NonStandardMargin"]));
                        }
                        Cell1.CellStyle = CountCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["CostWithTransport"]));
                        Cell1.CellStyle = PriceForeignCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["PaymentRate"]));
                        Cell1.CellStyle = CurrencyCS;
                    }
                    Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                    Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["Weight"]));
                    Cell1.CellStyle = WeightCS;
                    pos++;
                }

                if (CurrencyTypeID == 0)
                {
                    Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                    Cell1.SetCellValue("Заказ, " + Currency + ":");
                    Cell1.CellStyle = SummaryWithoutBorderBelCS;
                    Cell1 = sheet1.CreateRow(pos++).CreateCell(8);
                    Cell1.SetCellValue(Convert.ToDouble(TotalProfil));
                    Cell1.CellStyle = SummaryWithoutBorderBelCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                    Cell1.SetCellValue("Транспорт, прочее, " + Currency + ":");
                    Cell1.CellStyle = SummaryWithoutBorderBelCS;
                    Cell1 = sheet1.CreateRow(pos++).CreateCell(8);
                    Cell1.SetCellValue(Convert.ToDouble(TransportAndOtherProfil));
                    Cell1.CellStyle = SummaryWithoutBorderBelCS;
                    Cell1 = sheet1.CreateRow(pos - 2).CreateCell(10);
                    Cell1.SetCellValue(Convert.ToDouble(WeightProfil));
                    Cell1.CellStyle = SummaryWeightCS;
                }
                else
                {
                    Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                    Cell1.SetCellValue("Заказ, " + Currency + ":");
                    Cell1.CellStyle = SummaryWithoutBorderForeignCS;
                    Cell1 = sheet1.CreateRow(pos++).CreateCell(8);
                    Cell1.SetCellValue(Convert.ToDouble(TotalProfil));
                    Cell1.CellStyle = SummaryWithoutBorderForeignCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                    Cell1.SetCellValue("Транспорт, прочее, " + Currency + ":");
                    Cell1.CellStyle = SummaryWithoutBorderForeignCS;
                    Cell1 = sheet1.CreateRow(pos++).CreateCell(8);
                    Cell1.SetCellValue(Convert.ToDouble(TransportAndOtherProfil));
                    Cell1.CellStyle = SummaryWithoutBorderForeignCS;
                    Cell1 = sheet1.CreateRow(pos - 2).CreateCell(10);
                    Cell1.SetCellValue(Convert.ToDouble(WeightProfil));
                    Cell1.CellStyle = SummaryWeightCS;
                }
            }

            DisplayIndex = 0;
            if (TPSReportTable.Rows.Count > 0)
            {
                //ТПС
                pos++;
                Cell1 = sheet1.CreateRow(pos++).CreateCell(0);
                Cell1.SetCellValue("ЗОВ-ТПС:");
                Cell1.CellStyle = SummaryWithoutBorderBelCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                Cell1.SetCellValue("Наименование");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                Cell1.SetCellValue("Кол-во");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                Cell1.SetCellValue("Ед. измер.");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                Cell1.SetCellValue("Цена нач, " + Currency);
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                Cell1.SetCellValue("Коэф.");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                Cell1.SetCellValue("Скидка, %");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                Cell1.SetCellValue("Цена с транс, " + Currency);
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                Cell1.SetCellValue("Нестандарт, %");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                Cell1.SetCellValue("Стоимость с транс, " + Currency);
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                Cell1.SetCellValue("Курс, " + Currency);
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos++).CreateCell(DisplayIndex++);
                Cell1.SetCellValue("Вес, кг.");
                Cell1.CellStyle = SimpleHeaderCS;

                for (int i = 0; i < TPSReportTable.Rows.Count; i++)
                {
                    DisplayIndex = 0;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                    Cell1.SetCellValue(TPSReportTable.Rows[i]["Name"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                    Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["Count"]));
                    Cell1.CellStyle = CountCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                    Cell1.SetCellValue(TPSReportTable.Rows[i]["Measure"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    if (CurrencyTypeID == 0)
                    {
                        Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                        if (TPSReportTable.Rows[i]["OriginalPrice"] != DBNull.Value)
                            Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["OriginalPrice"]));
                        Cell1.CellStyle = PriceBelCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                        if (TPSReportTable.Rows[i]["DiscountVolume"] != DBNull.Value)
                            Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["DiscountVolume"]));
                        Cell1.CellStyle = WeightCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                        if (TPSReportTable.Rows[i]["TotalDiscount"] != DBNull.Value)
                            Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["TotalDiscount"]));
                        Cell1.CellStyle = WeightCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["PriceWithTransport"]));
                        Cell1.CellStyle = PriceBelCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                        if (TPSReportTable.Rows[i]["NonStandardMargin"] != DBNull.Value)
                        {
                            Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["NonStandardMargin"]));
                        }
                        Cell1.CellStyle = CountCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["CostWithTransport"]));
                        Cell1.CellStyle = PriceBelCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["PaymentRate"]));
                        Cell1.CellStyle = CurrencyCS;
                    }
                    else
                    {
                        Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                        if (TPSReportTable.Rows[i]["OriginalPrice"] != DBNull.Value)
                            Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["OriginalPrice"]));
                        Cell1.CellStyle = PriceForeignCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                        if (TPSReportTable.Rows[i]["DiscountVolume"] != DBNull.Value)
                            Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["DiscountVolume"]));
                        Cell1.CellStyle = WeightCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                        if (TPSReportTable.Rows[i]["TotalDiscount"] != DBNull.Value)
                            Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["TotalDiscount"]));
                        Cell1.CellStyle = WeightCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["PriceWithTransport"]));
                        Cell1.CellStyle = PriceForeignCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                        if (TPSReportTable.Rows[i]["NonStandardMargin"] != DBNull.Value)
                        {
                            Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["NonStandardMargin"]));
                        }
                        Cell1.CellStyle = CountCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["CostWithTransport"]));
                        Cell1.CellStyle = PriceForeignCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["PaymentRate"]));
                        Cell1.CellStyle = CurrencyCS;
                    }
                    Cell1 = sheet1.CreateRow(pos).CreateCell(DisplayIndex++);
                    Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["Weight"]));
                    Cell1.CellStyle = WeightCS;
                    pos++;
                }

                if (CurrencyTypeID == 0)
                {
                    Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                    Cell1.SetCellValue("Заказ, " + Currency + ":");
                    Cell1.CellStyle = SummaryWithoutBorderBelCS;
                    Cell1 = sheet1.CreateRow(pos++).CreateCell(8);
                    Cell1.SetCellValue(Convert.ToDouble(TotalTPS));
                    Cell1.CellStyle = SummaryWithoutBorderBelCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                    Cell1.SetCellValue("Транспорт, прочее, " + Currency + ":");
                    Cell1.CellStyle = SummaryWithoutBorderBelCS;
                    Cell1 = sheet1.CreateRow(pos++).CreateCell(8);
                    Cell1.SetCellValue(Convert.ToDouble(TransportAndOtherTPS));
                    Cell1.CellStyle = SummaryWithoutBorderBelCS;
                    Cell1 = sheet1.CreateRow(pos - 2).CreateCell(10);
                    Cell1.SetCellValue(Convert.ToDouble(WeightTPS));
                    Cell1.CellStyle = SummaryWeightCS;
                }
                else
                {
                    Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                    Cell1.SetCellValue("Заказ, " + Currency + ":");
                    Cell1.CellStyle = SummaryWithoutBorderForeignCS;
                    Cell1 = sheet1.CreateRow(pos++).CreateCell(8);
                    Cell1.SetCellValue(Convert.ToDouble(TotalTPS));
                    Cell1.CellStyle = SummaryWithoutBorderForeignCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                    Cell1.SetCellValue("Транспорт, прочее, " + Currency + ":");
                    Cell1.CellStyle = SummaryWithoutBorderForeignCS;
                    Cell1 = sheet1.CreateRow(pos++).CreateCell(8);
                    Cell1.SetCellValue(Convert.ToDouble(TransportAndOtherTPS));
                    Cell1.CellStyle = SummaryWithoutBorderForeignCS;
                    Cell1 = sheet1.CreateRow(pos - 2).CreateCell(10);
                    Cell1.SetCellValue(Convert.ToDouble(WeightTPS));
                    Cell1.CellStyle = SummaryWeightCS;
                }
            }
            pos++;

            //if (Rate != 1)
            //{
            //    Cell1 = sheet1.CreateRow(pos++).CreateCell(0);
            //    Cell1.SetCellValue("Курс, 1 EUR = " + Rate + " " + Currency);
            //    Cell1.CellStyle = SummaryWithoutBorderBelCS;
            //    //Excel.WriteCell(1, "Курс, 1 EUR = " + Rate + " " + Currency, pos++, 1, 12, true);
            //}

            if (CurrencyTypeID == 0)
            {
                Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                Cell1.SetCellValue("Транспорт, " + Currency + ":");
                Cell1.CellStyle = SummaryWithBorderBelCS;

                Cell1 = sheet1.CreateRow(pos++).CreateCell(1);
                Cell1.SetCellValue(Convert.ToDouble(TransportCost));
                Cell1.CellStyle = SummaryWithBorderBelCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                Cell1.SetCellValue("Прочее, " + Currency + ":");
                Cell1.CellStyle = SummaryWithBorderBelCS;

                Cell1 = sheet1.CreateRow(pos++).CreateCell(1);
                Cell1.SetCellValue(Convert.ToDouble(AdditionalCost));
                Cell1.CellStyle = SummaryWithBorderBelCS;
            }
            else
            {
                Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                Cell1.SetCellValue("Транспорт, " + Currency + ":");
                Cell1.CellStyle = SummaryWithBorderForeignCS;

                Cell1 = sheet1.CreateRow(pos++).CreateCell(1);
                Cell1.SetCellValue(Convert.ToDouble(TransportCost));
                Cell1.CellStyle = SummaryWithBorderForeignCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                Cell1.SetCellValue("Прочее, " + Currency + ":");
                Cell1.CellStyle = SummaryWithBorderForeignCS;

                Cell1 = sheet1.CreateRow(pos++).CreateCell(1);
                Cell1.SetCellValue(Convert.ToDouble(AdditionalCost));
                Cell1.CellStyle = SummaryWithBorderForeignCS;
            }

            if (IsComplaint(GetMegaOrderID(MainOrdersIDs[0])))
            {
                TotalCost = 0;
            }

            if (CurrencyTypeID == 0)
            {
                Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                Cell1.SetCellValue("Итого, " + Currency + ":");
                Cell1.CellStyle = SummaryWithBorderBelCS;
                Cell1 = sheet1.CreateRow(pos++).CreateCell(1);
                Cell1.SetCellValue(Convert.ToDouble((TotalCost)));
                Cell1.CellStyle = SummaryWithBorderBelCS;
            }
            else
            {
                Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                Cell1.SetCellValue("Итого, " + Currency + ":");
                Cell1.CellStyle = SummaryWithBorderForeignCS;
                Cell1 = sheet1.CreateRow(pos++).CreateCell(1);
                Cell1.SetCellValue(Convert.ToDouble((TotalCost)));
                Cell1.CellStyle = SummaryWithBorderForeignCS;
            }

            ClearReport();
        }

        public void ClearReport()
        {
            ProfilReportTable.Clear();
            TPSReportTable.Clear();

            FrontsReport.ClearReport();
            DecorReport.ClearReport();
        }
    }
}
