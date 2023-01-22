using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;

using System;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;

namespace Infinium.Modules.Marketing.NewOrders.InvoiceReportToDbf
{
    public class InvoiceReportToDbf
    {
        private decimal VAT = 1.0m;
        public InvoiceFrontsReportToDbf FrontsReport;
        public InvoiceDecorReportToDbf DecorReport = null;
        //HSSFWorkbook hssfworkbook;
        public DataTable CurrencyTypesDataTable = null;
        public DataTable ProfilReportTable = null;
        public DataTable TPSReportTable = null;

        public InvoiceReportToDbf(FrontsCatalogOrder FrontsCatalogOrder, DecorCatalogOrder DecorCatalogOrder)
        {
            FrontsReport = new InvoiceFrontsReportToDbf(FrontsCatalogOrder, DecorCatalogOrder);
            DecorReport = new InvoiceDecorReportToDbf(DecorCatalogOrder);

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
            ProfilReportTable.Columns.Add(new DataColumn("Count", Type.GetType("System.Decimal")));
            ProfilReportTable.Columns.Add(new DataColumn("Price", Type.GetType("System.Decimal")));
            ProfilReportTable.Columns.Add(new DataColumn("Cost", Type.GetType("System.Decimal")));
            ProfilReportTable.Columns.Add(new DataColumn("Weight", Type.GetType("System.Decimal")));
            ProfilReportTable.Columns.Add(new DataColumn("PackageCount", Type.GetType("System.Int32")));
            ProfilReportTable.Columns.Add(new DataColumn("OriginalPrice", Type.GetType("System.Decimal")));
            ProfilReportTable.Columns.Add(new DataColumn("PriceWithTransport", Type.GetType("System.Decimal")));
            ProfilReportTable.Columns.Add(new DataColumn("CostWithTransport", Type.GetType("System.Decimal")));
            ProfilReportTable.Columns.Add(new DataColumn("PaymentRate", Type.GetType("System.Decimal")));
            ProfilReportTable.Columns.Add(new DataColumn("IsNonStandard", Type.GetType("System.Boolean")));
            ProfilReportTable.Columns.Add(new DataColumn("NonStandardMargin", Type.GetType("System.Decimal")));
            TPSReportTable = ProfilReportTable.Clone();
        }

        public void Collect(DataTable table1, DataTable table2, bool IsNonStandard)
        {
            DataTable TempDT = ProfilReportTable.Clone();
            string IsNonStandardFilter = "(IsNonStandard=0 OR IsNonStandard IS NULL)";
            if (IsNonStandard)
                IsNonStandardFilter = "IsNonStandard=1";
            DataTable DistinctInvNumbersDT;
            using (DataView DV = new DataView(table1))
            {
                DV.RowFilter = IsNonStandardFilter;
                DistinctInvNumbersDT = DV.ToTable(true, new string[] { "AccountingName", "InvNumber", "PaymentRate" });
            }
            IsNonStandardFilter = " AND " + IsNonStandardFilter;
            for (int i = 0; i < DistinctInvNumbersDT.Rows.Count; i++)
            {
                decimal Cost = 0;
                decimal PriceWithTransport = 0;
                decimal CostWithTransport = 0;
                decimal Count = 0;
                decimal Price = 0;
                decimal Weight = 0;
                int DecCount = 2;
                int PackageCount = 0;
                //decimal OriginalPrice = Convert.ToDecimal(table1.Rows[0]["OriginalPrice"]);
                decimal NonStandardMargin = 0;
                string UNN = table1.Rows[0]["UNN"].ToString();
                string AccountingName = DistinctInvNumbersDT.Rows[i]["AccountingName"].ToString();
                string InvNumber = DistinctInvNumbersDT.Rows[i]["InvNumber"].ToString();
                string CurrencyCode = table1.Rows[0]["CurrencyCode"].ToString();
                decimal PaymentRate = Convert.ToDecimal(DistinctInvNumbersDT.Rows[i]["PaymentRate"]);
                if (table1.Rows[0]["NonStandardMargin"] != DBNull.Value)
                    NonStandardMargin = Convert.ToDecimal(table1.Rows[0]["NonStandardMargin"]);

                DataRow[] InvRows = table1.Select("InvNumber = '" + InvNumber + "' AND PaymentRate='" + PaymentRate + "' AND AccountingName='" + AccountingName + "'" + IsNonStandardFilter);
                if (InvRows.Count() > 0)
                {
                    for (int j = 0; j < InvRows.Count(); j++)
                    {
                        if (InvRows[j]["NonStandardMargin"] != DBNull.Value)
                            NonStandardMargin = Convert.ToDecimal(InvRows[j]["NonStandardMargin"]);
                        Weight += Convert.ToDecimal(InvRows[j]["Weight"]);
                        Count += Convert.ToDecimal(InvRows[j]["Count"]);
                        Cost += Convert.ToDecimal(InvRows[j]["Cost"]);
                        CostWithTransport += Convert.ToDecimal(InvRows[j]["CostWithTransport"]);
                    }
                    Cost = Cost * PaymentRate / VAT;
                    CostWithTransport = CostWithTransport * PaymentRate / VAT;
                    Cost = Math.Ceiling(Cost / 0.01m) * 0.01m;
                    CostWithTransport = Math.Ceiling(CostWithTransport / 0.01m) * 0.01m;
                    Price = Cost / Count;
                    PriceWithTransport = CostWithTransport / Count;
                    //if (Convert.ToInt32(CurrencyCode) == 974)
                    //{
                    //    DecCount = 0;
                    //    Price = Math.Ceiling(Price / 100.0m) * 100.0m;
                    //    Cost = Price * Count;
                    //    PriceWithTransport = Math.Ceiling(PriceWithTransport / 100.0m) * 100.0m;
                    //    CostWithTransport = PriceWithTransport * Count;
                    //}
                    Price = Decimal.Round(Price, DecCount, MidpointRounding.AwayFromZero);
                    Cost = Decimal.Round(Cost, DecCount, MidpointRounding.AwayFromZero);
                    PriceWithTransport = Decimal.Round(PriceWithTransport, DecCount, MidpointRounding.AwayFromZero);
                    CostWithTransport = Decimal.Round(CostWithTransport, DecCount, MidpointRounding.AwayFromZero);

                    DataRow NewRow = TempDT.NewRow();
                    NewRow["UNN"] = UNN;
                    NewRow["CurrencyCode"] = CurrencyCode;
                    NewRow["InvNumber"] = InvNumber;
                    NewRow["AccountingName"] = AccountingName;
                    NewRow["Count"] = Count;
                    NewRow["Price"] = Price;
                    NewRow["Cost"] = Cost;
                    NewRow["Weight"] = Weight;
                    NewRow["PackageCount"] = PackageCount;
                    //NewRow["OriginalPrice"] = OriginalPrice;
                    NewRow["PriceWithTransport"] = PriceWithTransport;
                    NewRow["CostWithTransport"] = CostWithTransport;
                    NewRow["PaymentRate"] = PaymentRate;
                    NewRow["IsNonStandard"] = IsNonStandard;
                    if (NonStandardMargin != 0)
                        NewRow["NonStandardMargin"] = NonStandardMargin;
                    TempDT.Rows.Add(NewRow);
                }
            }

            foreach (DataRow item in TempDT.Rows)
            {
                DataRow NewRow = ProfilReportTable.NewRow();
                NewRow.ItemArray = item.ItemArray;
                ProfilReportTable.Rows.Add(NewRow);
            }

            TempDT.Clear();
            IsNonStandardFilter = "(IsNonStandard=0 OR IsNonStandard IS NULL)";
            if (IsNonStandard)
                IsNonStandardFilter = "IsNonStandard=1";
            DistinctInvNumbersDT.Dispose();
            using (DataView DV = new DataView(table2))
            {
                DV.RowFilter = IsNonStandardFilter;
                DistinctInvNumbersDT = DV.ToTable(true, new string[] { "AccountingName", "InvNumber", "PaymentRate" });
            }
            IsNonStandardFilter = " AND " + IsNonStandardFilter;
            for (int i = 0; i < DistinctInvNumbersDT.Rows.Count; i++)
            {
                decimal Cost = 0;
                decimal PriceWithTransport = 0;
                decimal CostWithTransport = 0;
                decimal Count = 0;
                decimal Price = 0;
                decimal Weight = 0;
                int DecCount = 2;
                int PackageCount = 0;
                //decimal OriginalPrice = Convert.ToDecimal(table2.Rows[0]["OriginalPrice"]);
                decimal NonStandardMargin = 0;
                string UNN = table2.Rows[0]["UNN"].ToString();
                string AccountingName = DistinctInvNumbersDT.Rows[i]["AccountingName"].ToString();
                string InvNumber = DistinctInvNumbersDT.Rows[i]["InvNumber"].ToString();
                string CurrencyCode = table2.Rows[0]["CurrencyCode"].ToString();
                string TPSCurrencyCode = table2.Rows[0]["TPSCurCode"].ToString();
                decimal PaymentRate = Convert.ToDecimal(DistinctInvNumbersDT.Rows[i]["PaymentRate"]);
                DataRow[] InvRows = table2.Select("InvNumber = '" + InvNumber + "' AND PaymentRate='" + PaymentRate + "' AND AccountingName='" + AccountingName + "'" + IsNonStandardFilter);
                if (InvRows.Count() > 0)
                {
                    for (int j = 0; j < InvRows.Count(); j++)
                    {
                        if (InvRows[j]["NonStandardMargin"] != DBNull.Value)
                            NonStandardMargin = Convert.ToDecimal(InvRows[j]["NonStandardMargin"]);
                        Weight += Convert.ToDecimal(InvRows[j]["Weight"]);
                        Count += Convert.ToDecimal(InvRows[j]["Count"]);
                        Cost += Convert.ToDecimal(InvRows[j]["Cost"]) * PaymentRate / VAT;
                        CostWithTransport += Convert.ToDecimal(InvRows[j]["CostWithTransport"]) * PaymentRate / VAT;
                    }
                    Cost = Math.Ceiling(Cost / 0.01m) * 0.01m;
                    CostWithTransport = Math.Ceiling(CostWithTransport / 0.01m) * 0.01m;
                    Price = Cost / Count;
                    PriceWithTransport = CostWithTransport / Count;
                    //if (Convert.ToInt32(CurrencyCode) == 974)
                    //{
                    //    DecCount = 0;
                    //    Price = Math.Ceiling(Price / 100.0m) * 100.0m;
                    //    Cost = Price * Count;
                    //    PriceWithTransport = Math.Ceiling(PriceWithTransport / 100.0m) * 100.0m;
                    //    CostWithTransport = PriceWithTransport * Count;
                    //}
                    Price = Decimal.Round(Price, DecCount, MidpointRounding.AwayFromZero);
                    Cost = Decimal.Round(Cost, DecCount, MidpointRounding.AwayFromZero);
                    PriceWithTransport = Decimal.Round(PriceWithTransport, DecCount, MidpointRounding.AwayFromZero);
                    CostWithTransport = Decimal.Round(CostWithTransport, DecCount, MidpointRounding.AwayFromZero);

                    DataRow NewRow = TempDT.NewRow();
                    NewRow["UNN"] = UNN;
                    NewRow["CurrencyCode"] = CurrencyCode;
                    NewRow["TPSCurCode"] = TPSCurrencyCode;
                    NewRow["InvNumber"] = InvNumber;
                    NewRow["AccountingName"] = AccountingName;
                    NewRow["Count"] = Count;
                    NewRow["Price"] = Price;
                    NewRow["Cost"] = Cost;
                    NewRow["Weight"] = Weight;
                    NewRow["PackageCount"] = PackageCount;
                    //NewRow["OriginalPrice"] = OriginalPrice;
                    NewRow["PriceWithTransport"] = PriceWithTransport;
                    NewRow["CostWithTransport"] = CostWithTransport;
                    NewRow["PaymentRate"] = PaymentRate;
                    NewRow["IsNonStandard"] = IsNonStandard;
                    if (NonStandardMargin != 0)
                        NewRow["NonStandardMargin"] = NonStandardMargin;
                    TempDT.Rows.Add(NewRow);
                }
            }
            foreach (DataRow item in TempDT.Rows)
            {
                DataRow NewRow = TPSReportTable.NewRow();
                NewRow.ItemArray = item.ItemArray;
                TPSReportTable.Rows.Add(NewRow);
            }
        }

        public void AssignCost(decimal ComplaintProfilCost, decimal ComplaintTPSCost, decimal TransportCost, decimal AdditionalCost, decimal WeightProfil, decimal WeightTPS, decimal TotalWeight,
            ref decimal TransportAndOtherProfil, ref decimal TransportAndOtherTPS)
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
                            " FROM NewMegaOrders WHERE MegaOrderID = " + MegaOrderID,
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

        public DataTable GroupBy(string i_sGroupByColumn, string i_sAggregateColumn, DataTable i_dSourceTable)
        {
            DataView dv = new DataView(i_dSourceTable);

            //getting distinct values for group column
            DataTable dtGroup = dv.ToTable(true, new string[] { i_sGroupByColumn });

            //adding column for the row count
            dtGroup.Columns.Add("Count", typeof(int));

            //looping thru distinct values for the group, counting
            foreach (DataRow dr in dtGroup.Rows)
            {
                dr["Count"] = i_dSourceTable.Compute("Count(" + i_sAggregateColumn + ")", i_sGroupByColumn + " = '" + dr[i_sGroupByColumn] + "'");
            }
            dv.Dispose();
            dv = new DataView(dtGroup.Copy())
            {
                Sort = "Count"
            };
            dtGroup.Clear();
            dtGroup = dv.ToTable();

            //returning grouped/counted result
            return dtGroup;
        }

        public DataTable GetMegaOrdersTable(int[] MegaOrders)
        {
            DataTable ReturnDT = new DataTable();
            string SelectCommand = @"SELECT MegaOrders.MegaOrderID, MegaOrders.OrderNumber, MegaOrders.ClientID, 
                MegaOrders.ComplaintProfilCost, MegaOrders.ComplaintTPSCost, MegaOrders.TransportCost, MegaOrders.AdditionalCost, MegaOrders.CurrencyTypeID,
                MainOrders.MainOrderID, MainOrders.Weight
                FROM MainOrders
                INNER JOIN MegaOrders ON MainOrders.MegaOrderID=MegaOrders.MegaOrderID
                AND MegaOrders.MegaOrderID IN (" + string.Join(",", MegaOrders) + ")";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(ReturnDT);
                ReturnDT.Columns.Add(new DataColumn("ProfilFrontsWeight", Type.GetType("System.Decimal")));
                ReturnDT.Columns.Add(new DataColumn("TPSFrontsWeight", Type.GetType("System.Decimal")));
                ReturnDT.Columns.Add(new DataColumn("ProfilDecorWeight", Type.GetType("System.Decimal")));
                ReturnDT.Columns.Add(new DataColumn("TPSDecorWeight", Type.GetType("System.Decimal")));
                for (int i = 0; i < ReturnDT.Rows.Count; i++)
                {
                    decimal ProfilFrontsWeight = 0;
                    decimal TPSFrontsWeight = 0;
                    decimal ProfilDecorWeight = 0;
                    decimal TPSDecorWeight = 0;
                    int MainOrderID = Convert.ToInt32(ReturnDT.Rows[i]["MainOrderID"]);
                    SelectCommand = @"SELECT Square,Weight,FactoryID FROM FrontsOrders WHERE MainOrderID=" + MainOrderID;
                    using (SqlDataAdapter DA1 = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
                    {
                        using (DataTable DT = new DataTable())
                        {
                            DA1.Fill(DT);
                            for (int j = 0; j < DT.Rows.Count; j++)
                            {
                                if (Convert.ToInt32(DT.Rows[j]["FactoryID"]) == 1)
                                    ProfilFrontsWeight += Convert.ToDecimal(DT.Rows[j]["Square"]) * Convert.ToDecimal(0.7) + Convert.ToDecimal(DT.Rows[j]["Weight"]);
                                if (Convert.ToInt32(DT.Rows[j]["FactoryID"]) == 2)
                                    TPSFrontsWeight += Convert.ToDecimal(DT.Rows[j]["Square"]) * Convert.ToDecimal(0.7) + Convert.ToDecimal(DT.Rows[j]["Weight"]);
                            }
                        }
                    }
                    SelectCommand = @"SELECT Weight,FactoryID FROM DecorOrders WHERE MainOrderID=" + MainOrderID;
                    using (SqlDataAdapter DA1 = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
                    {
                        using (DataTable DT = new DataTable())
                        {
                            DA1.Fill(DT);
                            for (int j = 0; j < DT.Rows.Count; j++)
                            {
                                if (Convert.ToInt32(DT.Rows[j]["FactoryID"]) == 1)
                                    ProfilDecorWeight += Convert.ToDecimal(DT.Rows[j]["Weight"]);
                                if (Convert.ToInt32(DT.Rows[j]["FactoryID"]) == 2)
                                    TPSDecorWeight += Convert.ToDecimal(DT.Rows[j]["Weight"]);
                            }
                        }
                    }
                    ReturnDT.Rows[i]["ProfilFrontsWeight"] = ProfilFrontsWeight;
                    ReturnDT.Rows[i]["TPSFrontsWeight"] = TPSFrontsWeight;
                    ReturnDT.Rows[i]["ProfilDecorWeight"] = ProfilDecorWeight;
                    ReturnDT.Rows[i]["TPSDecorWeight"] = TPSDecorWeight;
                }
            }
            return ReturnDT;
        }

        public decimal CalcCurrencyCost(int MegaOrderID, int ClientID, decimal PaymentRate)
        {
            ClearReport();
            if (ClientID == 145)
            {
                VAT = 1.2m;
            }
            else
            {
                VAT = 1.0m;
            }

            FrontsReport.Report(MegaOrderID, PaymentRate);
            DecorReport.Report(MegaOrderID, PaymentRate);

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

            decimal TotalCost = 0;
            decimal TotalProfil = 0;
            decimal TotalTPS = 0;

            for (int i = 0; i < ProfilReportTable.Rows.Count; i++)
            {
                TotalProfil += Convert.ToDecimal(ProfilReportTable.Rows[i]["CostWithTransport"]);
            }

            for (int i = 0; i < TPSReportTable.Rows.Count; i++)
            {
                TotalTPS += Convert.ToDecimal(TPSReportTable.Rows[i]["CostWithTransport"]);
            }

            TotalCost = (TotalProfil + TotalTPS);
            TotalCost = Decimal.Round(TotalCost, 2, MidpointRounding.AwayFromZero);

            TotalCost = 0;
            TotalProfil = 0;
            TotalTPS = 0;

            DataTable table1 = ProfilReportTable.Copy();
            DataTable table2 = TPSReportTable.Copy();
            ProfilReportTable.Clear();
            TPSReportTable.Clear();
            Collect(table1, table2, false);
            Collect(table1, table2, true);

            using (DataView DV = new DataView(ProfilReportTable.Copy(), string.Empty, "AccountingName", DataViewRowState.CurrentRows))
            {
                ProfilReportTable.Clear();
                ProfilReportTable = DV.ToTable();
            }
            using (DataView DV = new DataView(TPSReportTable.Copy(), string.Empty, "AccountingName", DataViewRowState.CurrentRows))
            {
                TPSReportTable.Clear();
                TPSReportTable = DV.ToTable();
            }

            for (int i = 0; i < ProfilReportTable.Rows.Count; i++)
            {
                TotalProfil += Convert.ToDecimal(ProfilReportTable.Rows[i]["CostWithTransport"]);
            }

            for (int i = 0; i < TPSReportTable.Rows.Count; i++)
            {
                TotalTPS += Convert.ToDecimal(TPSReportTable.Rows[i]["CostWithTransport"]);
            }

            TotalCost = (TotalProfil + TotalTPS);
            TotalCost = Decimal.Round(TotalCost, 2, MidpointRounding.AwayFromZero);
            return TotalCost;
        }
        
        public decimal CalcCurrencyCost(int MegaOrderID, int ClientID, decimal PaymentRate, bool IsSample)
        {
            ClearReport();
            if (ClientID == 145)
            {
                VAT = 1.2m;
            }
            else
            {
                VAT = 1.0m;
            }

            FrontsReport.Report(MegaOrderID, PaymentRate, IsSample);
            DecorReport.Report(MegaOrderID, PaymentRate, IsSample);

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

            decimal TotalCost = 0;
            decimal TotalProfil = 0;
            decimal TotalTPS = 0;

            for (int i = 0; i < ProfilReportTable.Rows.Count; i++)
            {
                TotalProfil += Convert.ToDecimal(ProfilReportTable.Rows[i]["CostWithTransport"]);
            }

            for (int i = 0; i < TPSReportTable.Rows.Count; i++)
            {
                TotalTPS += Convert.ToDecimal(TPSReportTable.Rows[i]["CostWithTransport"]);
            }

            TotalCost = (TotalProfil + TotalTPS);
            TotalCost = Decimal.Round(TotalCost, 2, MidpointRounding.AwayFromZero);

            TotalCost = 0;
            TotalProfil = 0;
            TotalTPS = 0;

            DataTable table1 = ProfilReportTable.Copy();
            DataTable table2 = TPSReportTable.Copy();
            ProfilReportTable.Clear();
            TPSReportTable.Clear();
            Collect(table1, table2, false);
            Collect(table1, table2, true);

            using (DataView DV = new DataView(ProfilReportTable.Copy(), string.Empty, "AccountingName", DataViewRowState.CurrentRows))
            {
                ProfilReportTable.Clear();
                ProfilReportTable = DV.ToTable();
            }
            using (DataView DV = new DataView(TPSReportTable.Copy(), string.Empty, "AccountingName", DataViewRowState.CurrentRows))
            {
                TPSReportTable.Clear();
                TPSReportTable = DV.ToTable();
            }

            for (int i = 0; i < ProfilReportTable.Rows.Count; i++)
            {
                TotalProfil += Convert.ToDecimal(ProfilReportTable.Rows[i]["CostWithTransport"]);
            }

            for (int i = 0; i < TPSReportTable.Rows.Count; i++)
            {
                TotalTPS += Convert.ToDecimal(TPSReportTable.Rows[i]["CostWithTransport"]);
            }

            TotalCost = (TotalProfil + TotalTPS);
            TotalCost = Decimal.Round(TotalCost, 2, MidpointRounding.AwayFromZero);
            return TotalCost;
        }

        public void CreateReport(ref HSSFWorkbook hssfworkbook, int[] MegaOrders, int[] OrderNumbers, int[] MainOrdersIDs, int ClientID, string ClientName,
            decimal ComplaintProfilCost, decimal ComplaintTPSCost, decimal TransportCost, decimal AdditionalCost,
            decimal TotalCost, int CurrencyTypeID, decimal TotalWeight, bool IsSample, ref decimal TotalProfil1, ref decimal TotalTPS1)
        {
            ClearReport();
            DataTable ReturnDT = GetMegaOrdersTable(MegaOrders);
            string MainOrdersList = string.Empty;

            int pos = 0;

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
            FrontsReport.Report(MainOrdersIDs, IsSample);
            DecorReport.Report(MainOrdersIDs, IsSample);

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

            decimal Total = TransportCost + AdditionalCost - ComplaintProfilCost - ComplaintTPSCost;
            decimal TotalProfil = 0;
            decimal TotalTPS = 0;
            decimal WeightProfil = 0;
            decimal WeightTPS = 0;

            decimal TransportAndOtherProfil = 0;
            decimal TransportAndOtherTPS = 0;

            DataTable table1 = ProfilReportTable.Copy();
            DataTable table2 = TPSReportTable.Copy();
            ProfilReportTable.Clear();
            TPSReportTable.Clear();
            Collect(table1, table2, false);
            Collect(table1, table2, true);
            using (DataView DV = new DataView(ProfilReportTable.Copy(), string.Empty, "AccountingName", DataViewRowState.CurrentRows))
            {
                ProfilReportTable.Clear();
                ProfilReportTable = DV.ToTable();
            }
            using (DataView DV = new DataView(TPSReportTable.Copy(), string.Empty, "AccountingName", DataViewRowState.CurrentRows))
            {
                TPSReportTable.Clear();
                TPSReportTable = DV.ToTable();
            }

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
            AssignCost(ComplaintProfilCost, ComplaintTPSCost, TransportCost, AdditionalCost, WeightProfil, WeightTPS, TotalWeight, ref TransportAndOtherProfil,
                       ref TransportAndOtherTPS);

            TotalProfil1 = (TotalProfil);
            TotalTPS1 = (TotalTPS);
            //decimal dd = Decimal.Round((WeightProfil + WeightTPS) / TotalWeight, 3, MidpointRounding.AwayFromZero);
            decimal dd = 0;
            if (TotalWeight != 0)
                dd = Decimal.Round((WeightProfil + WeightTPS) / TotalWeight, 3, MidpointRounding.AwayFromZero);
            TotalProfil = Decimal.Round(TotalProfil, 3, MidpointRounding.AwayFromZero);
            TotalTPS = Decimal.Round(TotalTPS, 3, MidpointRounding.AwayFromZero);
            TransportCost = Decimal.Round(TransportCost * dd, 3, MidpointRounding.AwayFromZero);
            AdditionalCost = Decimal.Round(AdditionalCost * dd, 3, MidpointRounding.AwayFromZero);
            TotalCost = Decimal.Round(TotalCost, 3, MidpointRounding.AwayFromZero);

            for (int i = 0; i < ProfilReportTable.Rows.Count; i++)
                ProfilReportTable.Rows[i]["PackageCount"] = 0;
            for (int i = 0; i < TPSReportTable.Rows.Count; i++)
                TPSReportTable.Rows[i]["PackageCount"] = 0;

            DataTable dtGroup = new DataTable();
            DataTable dtMainOrders = new DataTable();
            DataTable DT = null;
            DataTable TempDT = new DataTable();
            DataTable dtDistInvNumbers = new DataTable();
            DataTable DistPackagesDT = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT PackageDetails.PackNumber, infiniu2_catalog.dbo.FrontsConfig.InvNumber, infiniu2_catalog.dbo.FrontsConfig.FactoryID, FrontsOrders.MainOrderID
                FROM PackageDetails INNER JOIN
                FrontsOrders ON PackageDetails.OrderID = FrontsOrders.FrontsOrdersID INNER JOIN
                infiniu2_catalog.dbo.FrontsConfig ON FrontsOrders.FrontConfigID = infiniu2_catalog.dbo.FrontsConfig.FrontConfigID
                WHERE PackageDetails.PackageID IN 
                (SELECT PackageID FROM Packages WHERE ProductType = 0 AND MainOrderID IN (" + string.Join(",", MainOrdersIDs) + "))",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(TempDT);
            }
            if (DT == null)
                DT = TempDT.Clone();
            foreach (DataRow item in TempDT.Rows)
                DT.Rows.Add(item.ItemArray);
            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT PackageDetails.PackNumber, infiniu2_catalog.dbo.DecorConfig.InvNumber, infiniu2_catalog.dbo.DecorConfig.FactoryID, DecorOrders.MainOrderID
                FROM PackageDetails INNER JOIN
                DecorOrders ON PackageDetails.OrderID = DecorOrders.DecorOrderID INNER JOIN
                infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID = infiniu2_catalog.dbo.DecorConfig.DecorConfigID
                WHERE PackageDetails.PackageID IN 
                (SELECT PackageID FROM Packages WHERE ProductType = 1 AND MainOrderID IN (" + string.Join(",", MainOrdersIDs) + "))",
                 ConnectionStrings.MarketingOrdersConnectionString))
            {
                TempDT.Clear();
                DA.Fill(TempDT);
            }
            if (DT == null)
                DT = TempDT.Clone();
            foreach (DataRow item in TempDT.Rows)
                DT.Rows.Add(item.ItemArray);

            for (int z = 0; z < MainOrdersIDs.Count(); z++)
            {
                using (DataView DV = new DataView(DT, "MainOrderID=" + MainOrdersIDs[z], string.Empty, DataViewRowState.CurrentRows))
                {
                    dtMainOrders.Clear();
                    dtMainOrders = DV.ToTable();
                }
                using (DataView DV = new DataView(dtMainOrders, "FactoryID=1", string.Empty, DataViewRowState.CurrentRows))
                {
                    DistPackagesDT.Clear();
                    DistPackagesDT = DV.ToTable(true, "PackNumber");
                }
                dtGroup.Clear();
                dtGroup = GroupBy("InvNumber", "PackNumber", dtMainOrders);
                for (int i = 0; i < DistPackagesDT.Rows.Count; i++)
                {
                    using (DataView DV = new DataView(dtMainOrders, "PackNumber=" + Convert.ToInt32(DistPackagesDT.Rows[i]["PackNumber"]), string.Empty, DataViewRowState.CurrentRows))
                    {
                        dtDistInvNumbers.Clear();
                        dtDistInvNumbers = DV.ToTable(true, "InvNumber");
                        string InvNumber = dtDistInvNumbers.Rows[0]["InvNumber"].ToString();
                        int PackageCount = 0;

                        DataRow[] rows0 = dtGroup.Select("InvNumber='" + InvNumber + "'");
                        if (rows0.Count() > 0)
                            PackageCount = Convert.ToInt32(rows0[0]["Count"]);

                        for (int x = 1; x < dtDistInvNumbers.Rows.Count; x++)
                        {
                            DataRow[] rows1 = dtGroup.Select("InvNumber='" + dtDistInvNumbers.Rows[x]["InvNumber"].ToString() + "'");
                            if (rows1.Count() > 0)
                            {
                                if (Convert.ToInt32(rows1[0]["Count"]) < PackageCount)
                                {
                                    PackageCount = Convert.ToInt32(rows1[0]["Count"]);
                                    InvNumber = dtDistInvNumbers.Rows[x]["InvNumber"].ToString();
                                }
                            }
                        }
                        PackageCount = 0;
                        for (int x = 0; x < dtDistInvNumbers.Rows.Count; x++)
                        {
                            DataRow[] rows1 = ProfilReportTable.Select("InvNumber='" + InvNumber + "'");
                            if (rows1.Count() > 0)
                            {
                                if (Convert.ToInt32(rows1[0]["PackageCount"]) < PackageCount)
                                {
                                    PackageCount = Convert.ToInt32(rows1[0]["PackageCount"]);
                                    InvNumber = dtDistInvNumbers.Rows[x]["InvNumber"].ToString();
                                }
                            }
                        }
                        DataRow[] rows = ProfilReportTable.Select("InvNumber='" + InvNumber + "'");
                        if (rows.Count() > 0)
                            rows[0]["PackageCount"] = Convert.ToInt32(rows[0]["PackageCount"]) + 1;
                    }
                }



                using (DataView DV = new DataView(dtMainOrders, "FactoryID=2", string.Empty, DataViewRowState.CurrentRows))
                {
                    DistPackagesDT.Clear();
                    DistPackagesDT = DV.ToTable(true, "PackNumber");
                }
                dtGroup.Clear();
                dtGroup = GroupBy("InvNumber", "PackNumber", dtMainOrders);
                for (int i = 0; i < DistPackagesDT.Rows.Count; i++)
                {
                    using (DataView DV = new DataView(dtMainOrders, "PackNumber=" + Convert.ToInt32(DistPackagesDT.Rows[i]["PackNumber"]), string.Empty, DataViewRowState.CurrentRows))
                    {
                        dtDistInvNumbers.Clear();
                        dtDistInvNumbers = DV.ToTable(true, "InvNumber");
                        string InvNumber = dtDistInvNumbers.Rows[0]["InvNumber"].ToString();
                        int PackageCount = 0;

                        DataRow[] rows0 = dtGroup.Select("InvNumber='" + InvNumber + "'");
                        if (rows0.Count() > 0)
                            PackageCount = Convert.ToInt32(rows0[0]["Count"]);

                        for (int x = 1; x < dtDistInvNumbers.Rows.Count; x++)
                        {
                            DataRow[] rows1 = dtGroup.Select("InvNumber='" + dtDistInvNumbers.Rows[x]["InvNumber"].ToString() + "'");
                            if (rows1.Count() > 0)
                            {
                                if (Convert.ToInt32(rows1[0]["Count"]) < PackageCount)
                                {
                                    PackageCount = Convert.ToInt32(rows1[0]["Count"]);
                                    InvNumber = dtDistInvNumbers.Rows[x]["InvNumber"].ToString();
                                }
                            }
                        }
                        PackageCount = 0;
                        for (int x = 0; x < dtDistInvNumbers.Rows.Count; x++)
                        {
                            DataRow[] rows1 = TPSReportTable.Select("InvNumber='" + InvNumber + "'");
                            if (rows1.Count() > 0)
                            {
                                if (Convert.ToInt32(rows1[0]["PackageCount"]) < PackageCount)
                                {
                                    PackageCount = Convert.ToInt32(rows1[0]["PackageCount"]);
                                    InvNumber = dtDistInvNumbers.Rows[x]["InvNumber"].ToString();
                                }
                            }
                        }
                        DataRow[] rows = TPSReportTable.Select("InvNumber='" + InvNumber + "'");
                        if (rows.Count() > 0)
                            rows[0]["PackageCount"] = Convert.ToInt32(rows[0]["PackageCount"]) + 1;
                    }
                }
            }

            #region 

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
            SimpleF.FontHeightInPoints = 8;
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

            HSSFSheet sheet1 = hssfworkbook.CreateSheet("Сводный отчет");
            sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

            sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
            sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
            sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
            sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

            sheet1.SetColumnWidth(0, 16 * 256);
            sheet1.SetColumnWidth(1, 12 * 256);
            sheet1.SetColumnWidth(2, 71 * 256);
            sheet1.SetColumnWidth(3, 8 * 256);
            sheet1.SetColumnWidth(4, 15 * 256);
            sheet1.SetColumnWidth(5, 12 * 256);
            sheet1.SetColumnWidth(6, 17 * 256);
            sheet1.SetColumnWidth(7, 15 * 256);
            sheet1.SetColumnWidth(8, 10 * 256);

            HSSFCell Cell1 = sheet1.CreateRow(pos).CreateCell(0);
            Cell1.SetCellValue("Клиент:");
            Cell1.CellStyle = ReportCS2;

            Cell1 = sheet1.CreateRow(pos++).CreateCell(1);
            Cell1.SetCellValue(ClientName);
            Cell1.CellStyle = ReportCS2;

            Cell1 = sheet1.CreateRow(pos).CreateCell(0);
            Cell1.SetCellValue("№ заказов:");
            Cell1.CellStyle = ReportCS1;

            Cell1 = sheet1.CreateRow(pos++).CreateCell(1);
            Cell1.SetCellValue(string.Join(", ", OrderNumbers));
            Cell1.CellStyle = ReportCS1;

            int RowCount = 1;


            Cell1 = sheet1.CreateRow(pos++).CreateCell(0);
            Cell1.SetCellValue("№ подзаказов: ");
            Cell1.CellStyle = SummaryWithoutBorderBelCS;

            for (int i = 0; i < MainOrdersIDs.Count(); i++)
            {
                MainOrdersList += MainOrdersIDs[i].ToString() + ", ";

                if (i > 0 && i % 9 == 0)
                {
                    Cell1 = sheet1.CreateRow(pos++).CreateCell(0);
                    Cell1.SetCellValue(MainOrdersList);
                    Cell1.CellStyle = SummaryWithoutBorderBelCS;

                    MainOrdersList = string.Empty;
                    RowCount++;
                }
                if (i == MainOrdersIDs.Count() - 1)
                {
                    if (MainOrdersList.Length > 0)
                        MainOrdersList = MainOrdersList.Remove(MainOrdersList.Length - 2, 2);
                    Cell1 = sheet1.CreateRow(pos++).CreateCell(0);
                    Cell1.SetCellValue(MainOrdersList);
                    Cell1.CellStyle = SummaryWithoutBorderBelCS;

                    MainOrdersList = string.Empty;
                    RowCount++;
                }
            }

            if (IsComplaint(GetMegaOrderID(MainOrdersIDs[0])))
            {
                Cell1 = sheet1.CreateRow(pos++).CreateCell(0);
                Cell1.SetCellValue("РЕКЛАМАЦИЯ");
                Cell1.CellStyle = ReportCS1;
            }

            int displayIndex = 0;
            if (ProfilReportTable.Rows.Count > 0)
            {
                //Профиль
                pos += 2;

                displayIndex = 0;

                Cell1 = sheet1.CreateRow(pos++).CreateCell(0);
                Cell1.SetCellValue("ОМЦ-ПРОФИЛЬ:");
                Cell1.CellStyle = SummaryWithoutBorderBelCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("УНН");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Инв.№");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Наименование");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Кол-во");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Цена с транс, " + Currency);
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Нестандарт, %");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Стоимость с транс, " + Currency);
                Cell1.CellStyle = SimpleHeaderCS;

                //Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                //Cell1.SetCellValue("Курс, " + Currency);
                //Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Вес, кг. ");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Кол-во уп. ");
                Cell1.CellStyle = SimpleHeaderCS;
                pos++;

                for (int i = 0; i < ProfilReportTable.Rows.Count; i++)
                {
                    displayIndex = 0;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(ProfilReportTable.Rows[i]["UNN"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(ProfilReportTable.Rows[i]["InvNumber"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(ProfilReportTable.Rows[i]["AccountingName"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["Count"]));
                    Cell1.CellStyle = CountCS;
                    if (CurrencyTypeID == 0)
                    {
                        Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                        Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["PriceWithTransport"]));
                        Cell1.CellStyle = PriceBelCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                        if (ProfilReportTable.Rows[i]["NonStandardMargin"] != DBNull.Value)
                        {
                            Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["NonStandardMargin"]));
                        }
                        Cell1.CellStyle = CountCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                        Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["CostWithTransport"]));
                        Cell1.CellStyle = PriceBelCS;
                        //Cell1 = sheet1.CreateRow(pos).CreateCell(7);
                        //Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["PaymentRate"]));
                        //Cell1.CellStyle = CurrencyCS;
                    }
                    else
                    {
                        Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                        Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["PriceWithTransport"]));
                        Cell1.CellStyle = PriceForeignCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                        if (ProfilReportTable.Rows[i]["NonStandardMargin"] != DBNull.Value)
                        {
                            Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["NonStandardMargin"]));
                        }
                        Cell1.CellStyle = CountCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                        Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["CostWithTransport"]));
                        Cell1.CellStyle = PriceForeignCS;
                        //Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                        //Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["PaymentRate"]));
                        //Cell1.CellStyle = CurrencyCS;
                    }
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["Weight"]));
                    Cell1.CellStyle = WeightCS;
                    if (ProfilReportTable.Rows[i]["PackageCount"] != DBNull.Value)
                    {
                        Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                        Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["PackageCount"]));
                        Cell1.CellStyle = SimpleCS;
                    }
                    pos++;
                }

                if (CurrencyTypeID == 0)
                {
                    displayIndex = 0;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                    Cell1.SetCellValue("Заказ, " + Currency + ":");
                    Cell1.CellStyle = SummaryWithoutBorderBelCS;
                    Cell1 = sheet1.CreateRow(pos++).CreateCell(6);
                    Cell1.SetCellValue(Convert.ToDouble(TotalProfil));
                    Cell1.CellStyle = SummaryWithoutBorderBelCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                    Cell1.SetCellValue("Транспорт, прочее, " + Currency + ":");
                    Cell1.CellStyle = SummaryWithoutBorderBelCS;
                    Cell1 = sheet1.CreateRow(pos++).CreateCell(6);
                    Cell1.SetCellValue(Convert.ToDouble(TransportAndOtherProfil));
                    Cell1.CellStyle = SummaryWithoutBorderBelCS;
                    Cell1 = sheet1.CreateRow(pos - 2).CreateCell(7);
                    Cell1.SetCellValue(Convert.ToDouble(WeightProfil));
                    Cell1.CellStyle = SummaryWeightCS;
                }
                else
                {
                    Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                    Cell1.SetCellValue("Заказ, " + Currency + ":");
                    Cell1.CellStyle = SummaryWithoutBorderForeignCS;
                    Cell1 = sheet1.CreateRow(pos++).CreateCell(6);
                    Cell1.SetCellValue(Convert.ToDouble(TotalProfil));
                    Cell1.CellStyle = SummaryWithoutBorderForeignCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                    Cell1.SetCellValue("Транспорт, прочее, " + Currency + ":");
                    Cell1.CellStyle = SummaryWithoutBorderForeignCS;
                    Cell1 = sheet1.CreateRow(pos++).CreateCell(6);
                    Cell1.SetCellValue(Convert.ToDouble(TransportAndOtherProfil));
                    Cell1.CellStyle = SummaryWithoutBorderForeignCS;
                    Cell1 = sheet1.CreateRow(pos - 2).CreateCell(7);
                    Cell1.SetCellValue(Convert.ToDouble(WeightProfil));
                    Cell1.CellStyle = SummaryWeightCS;
                }
            }

            if (TPSReportTable.Rows.Count > 0)
            {
                //ТПС
                pos++;

                Cell1 = sheet1.CreateRow(pos++).CreateCell(0);
                Cell1.SetCellValue("ЗОВ-ТПС:");
                Cell1.CellStyle = SummaryWithoutBorderBelCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                Cell1.SetCellValue("УНН");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(1);
                Cell1.SetCellValue("Инв.№");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(2);
                Cell1.SetCellValue("Наименование");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(3);
                Cell1.SetCellValue("Кол-во");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(4);
                Cell1.SetCellValue("Цена с транс, " + Currency);
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(5);
                Cell1.SetCellValue("Нестандарт, %");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(6);
                Cell1.SetCellValue("Стоимость с транс, " + Currency);
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(7);
                Cell1.SetCellValue("Курс, " + Currency);
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(8);
                Cell1.SetCellValue("Вес, кг. ");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(9);
                Cell1.SetCellValue("Кол-во уп. ");
                Cell1.CellStyle = SimpleHeaderCS;
                pos++;

                for (int i = 0; i < TPSReportTable.Rows.Count; i++)
                {
                    Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                    Cell1.SetCellValue(TPSReportTable.Rows[i]["UNN"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(1);
                    Cell1.SetCellValue(TPSReportTable.Rows[i]["InvNumber"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(2);
                    Cell1.SetCellValue(TPSReportTable.Rows[i]["AccountingName"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(3);
                    Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["Count"]));
                    Cell1.CellStyle = CountCS;
                    if (CurrencyTypeID == 0)
                    {
                        Cell1 = sheet1.CreateRow(pos).CreateCell(4);
                        Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["PriceWithTransport"]));
                        Cell1.CellStyle = PriceBelCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(5);
                        if (TPSReportTable.Rows[i]["NonStandardMargin"] != DBNull.Value)
                        {
                            Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["NonStandardMargin"]));
                        }
                        Cell1.CellStyle = CountCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(6);
                        Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["Cost"]));
                        Cell1.CellStyle = PriceBelCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(7);
                        Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["PaymentRate"]));
                        Cell1.CellStyle = CurrencyCS;
                    }
                    else
                    {
                        Cell1 = sheet1.CreateRow(pos).CreateCell(4);
                        Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["PriceWithTransport"]));
                        Cell1.CellStyle = PriceForeignCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(5);
                        if (TPSReportTable.Rows[i]["NonStandardMargin"] != DBNull.Value)
                        {
                            Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["NonStandardMargin"]));
                        }
                        Cell1.CellStyle = CountCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(6);
                        Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["CostWithTransport"]));
                        Cell1.CellStyle = PriceForeignCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(7);
                        Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["PaymentRate"]));
                        Cell1.CellStyle = CurrencyCS;
                    }
                    Cell1 = sheet1.CreateRow(pos).CreateCell(8);
                    Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["Weight"]));
                    Cell1.CellStyle = WeightCS;
                    if (TPSReportTable.Rows[i]["PackageCount"] != DBNull.Value)
                    {
                        Cell1 = sheet1.CreateRow(pos).CreateCell(9);
                        Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["PackageCount"]));
                        Cell1.CellStyle = SimpleCS;
                    }
                    pos++;
                }

                if (CurrencyTypeID == 0)
                {
                    Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                    Cell1.SetCellValue("Заказ, " + Currency + ":");
                    Cell1.CellStyle = SummaryWithoutBorderBelCS;
                    Cell1 = sheet1.CreateRow(pos++).CreateCell(6);
                    Cell1.SetCellValue(Convert.ToDouble(TotalTPS));
                    Cell1.CellStyle = SummaryWithoutBorderBelCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                    Cell1.SetCellValue("Транспорт, прочее, " + Currency + ":");
                    Cell1.CellStyle = SummaryWithoutBorderBelCS;
                    Cell1 = sheet1.CreateRow(pos++).CreateCell(6);
                    Cell1.SetCellValue(Convert.ToDouble(TransportAndOtherTPS));
                    Cell1.CellStyle = SummaryWithoutBorderBelCS;
                    Cell1 = sheet1.CreateRow(pos - 2).CreateCell(7);
                    Cell1.SetCellValue(Convert.ToDouble(WeightTPS));
                    Cell1.CellStyle = SummaryWeightCS;
                }
                else
                {
                    Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                    Cell1.SetCellValue("Заказ, " + Currency + ":");
                    Cell1.CellStyle = SummaryWithoutBorderForeignCS;
                    Cell1 = sheet1.CreateRow(pos++).CreateCell(6);
                    Cell1.SetCellValue(Convert.ToDouble(TotalTPS));
                    Cell1.CellStyle = SummaryWithoutBorderForeignCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                    Cell1.SetCellValue("Транспорт, прочее, " + Currency + ":");
                    Cell1.CellStyle = SummaryWithoutBorderForeignCS;
                    Cell1 = sheet1.CreateRow(pos++).CreateCell(6);
                    Cell1.SetCellValue(Convert.ToDouble(TransportAndOtherTPS));
                    Cell1.CellStyle = SummaryWithoutBorderForeignCS;
                    Cell1 = sheet1.CreateRow(pos - 2).CreateCell(7);
                    Cell1.SetCellValue(Convert.ToDouble(WeightTPS));
                    Cell1.CellStyle = SummaryWeightCS;
                }
            }
            pos++;

            //if (PaymentRate != 1)
            //{
            //    Cell1 = sheet1.CreateRow(pos++).CreateCell(0);
            //    Cell1.SetCellValue("Курс, 1 EUR = " + PaymentRate + " " + Currency);
            //    Cell1.CellStyle = SummaryWithoutBorderBelCS;
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
            #endregion
        }

        public void CreateReport(ref HSSFWorkbook hssfworkbook, int[] MegaOrders, int[] OrderNumbers, int[] MainOrdersIDs, int ClientID, string ClientName,
            decimal ComplaintProfilCost, decimal ComplaintTPSCost, decimal TransportCost, decimal AdditionalCost,
            decimal TotalCost, int CurrencyTypeID, decimal TotalWeight, ref decimal TotalProfil1, ref decimal TotalTPS1)
        {
            ClearReport();
            DataTable ReturnDT = GetMegaOrdersTable(MegaOrders);
            string MainOrdersList = string.Empty;

            int pos = 0;

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
            FrontsReport.Report(MainOrdersIDs);
            DecorReport.Report(MainOrdersIDs);

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

            decimal Total = TransportCost + AdditionalCost - ComplaintProfilCost - ComplaintTPSCost;
            decimal TotalProfil = 0;
            decimal TotalTPS = 0;
            decimal WeightProfil = 0;
            decimal WeightTPS = 0;

            decimal TransportAndOtherProfil = 0;
            decimal TransportAndOtherTPS = 0;

            DataTable table1 = ProfilReportTable.Copy();
            DataTable table2 = TPSReportTable.Copy();
            ProfilReportTable.Clear();
            TPSReportTable.Clear();
            Collect(table1, table2, false);
            Collect(table1, table2, true);
            using (DataView DV = new DataView(ProfilReportTable.Copy(), string.Empty, "AccountingName", DataViewRowState.CurrentRows))
            {
                ProfilReportTable.Clear();
                ProfilReportTable = DV.ToTable();
            }
            using (DataView DV = new DataView(TPSReportTable.Copy(), string.Empty, "AccountingName", DataViewRowState.CurrentRows))
            {
                TPSReportTable.Clear();
                TPSReportTable = DV.ToTable();
            }

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
            AssignCost(ComplaintProfilCost, ComplaintTPSCost, TransportCost, AdditionalCost, WeightProfil, WeightTPS, TotalWeight, ref TransportAndOtherProfil,
                       ref TransportAndOtherTPS);

            TotalProfil1 = (TotalProfil);
            TotalTPS1 = (TotalTPS);
            //decimal dd = Decimal.Round((WeightProfil + WeightTPS) / TotalWeight, 3, MidpointRounding.AwayFromZero);
            decimal dd = 0;
            if (TotalWeight != 0)
                dd = Decimal.Round((WeightProfil + WeightTPS) / TotalWeight, 3, MidpointRounding.AwayFromZero);
            TotalProfil = Decimal.Round(TotalProfil, 3, MidpointRounding.AwayFromZero);
            TotalTPS = Decimal.Round(TotalTPS, 3, MidpointRounding.AwayFromZero);
            TransportCost = Decimal.Round(TransportCost * dd, 3, MidpointRounding.AwayFromZero);
            AdditionalCost = Decimal.Round(AdditionalCost * dd, 3, MidpointRounding.AwayFromZero);
            TotalCost = Decimal.Round(TotalCost, 3, MidpointRounding.AwayFromZero);

            for (int i = 0; i < ProfilReportTable.Rows.Count; i++)
                ProfilReportTable.Rows[i]["PackageCount"] = 0;
            for (int i = 0; i < TPSReportTable.Rows.Count; i++)
                TPSReportTable.Rows[i]["PackageCount"] = 0;

            DataTable dtGroup = new DataTable();
            DataTable dtMainOrders = new DataTable();
            DataTable DT = null;
            DataTable TempDT = new DataTable();
            DataTable dtDistInvNumbers = new DataTable();
            DataTable DistPackagesDT = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT PackageDetails.PackNumber, infiniu2_catalog.dbo.FrontsConfig.InvNumber, infiniu2_catalog.dbo.FrontsConfig.FactoryID, FrontsOrders.MainOrderID
                FROM PackageDetails INNER JOIN
                FrontsOrders ON PackageDetails.OrderID = FrontsOrders.FrontsOrdersID INNER JOIN
                infiniu2_catalog.dbo.FrontsConfig ON FrontsOrders.FrontConfigID = infiniu2_catalog.dbo.FrontsConfig.FrontConfigID
                WHERE PackageDetails.PackageID IN 
                (SELECT PackageID FROM Packages WHERE ProductType = 0 AND MainOrderID IN (" + string.Join(",", MainOrdersIDs) + "))",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(TempDT);
            }
            if (DT == null)
                DT = TempDT.Clone();
            foreach (DataRow item in TempDT.Rows)
                DT.Rows.Add(item.ItemArray);
            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT PackageDetails.PackNumber, infiniu2_catalog.dbo.DecorConfig.InvNumber, infiniu2_catalog.dbo.DecorConfig.FactoryID, DecorOrders.MainOrderID
                FROM PackageDetails INNER JOIN
                DecorOrders ON PackageDetails.OrderID = DecorOrders.DecorOrderID INNER JOIN
                infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID = infiniu2_catalog.dbo.DecorConfig.DecorConfigID
                WHERE PackageDetails.PackageID IN 
                (SELECT PackageID FROM Packages WHERE ProductType = 1 AND MainOrderID IN (" + string.Join(",", MainOrdersIDs) + "))",
                 ConnectionStrings.MarketingOrdersConnectionString))
            {
                TempDT.Clear();
                DA.Fill(TempDT);
            }
            if (DT == null)
                DT = TempDT.Clone();
            foreach (DataRow item in TempDT.Rows)
                DT.Rows.Add(item.ItemArray);

            for (int z = 0; z < MainOrdersIDs.Count(); z++)
            {
                using (DataView DV = new DataView(DT, "MainOrderID=" + MainOrdersIDs[z], string.Empty, DataViewRowState.CurrentRows))
                {
                    dtMainOrders.Clear();
                    dtMainOrders = DV.ToTable();
                }
                using (DataView DV = new DataView(dtMainOrders, "FactoryID=1", string.Empty, DataViewRowState.CurrentRows))
                {
                    DistPackagesDT.Clear();
                    DistPackagesDT = DV.ToTable(true, "PackNumber");
                }
                dtGroup.Clear();
                dtGroup = GroupBy("InvNumber", "PackNumber", dtMainOrders);
                for (int i = 0; i < DistPackagesDT.Rows.Count; i++)
                {
                    using (DataView DV = new DataView(dtMainOrders, "PackNumber=" + Convert.ToInt32(DistPackagesDT.Rows[i]["PackNumber"]), string.Empty, DataViewRowState.CurrentRows))
                    {
                        dtDistInvNumbers.Clear();
                        dtDistInvNumbers = DV.ToTable(true, "InvNumber");
                        string InvNumber = dtDistInvNumbers.Rows[0]["InvNumber"].ToString();
                        int PackageCount = 0;

                        DataRow[] rows0 = dtGroup.Select("InvNumber='" + InvNumber + "'");
                        if (rows0.Count() > 0)
                            PackageCount = Convert.ToInt32(rows0[0]["Count"]);

                        for (int x = 1; x < dtDistInvNumbers.Rows.Count; x++)
                        {
                            DataRow[] rows1 = dtGroup.Select("InvNumber='" + dtDistInvNumbers.Rows[x]["InvNumber"].ToString() + "'");
                            if (rows1.Count() > 0)
                            {
                                if (Convert.ToInt32(rows1[0]["Count"]) < PackageCount)
                                {
                                    PackageCount = Convert.ToInt32(rows1[0]["Count"]);
                                    InvNumber = dtDistInvNumbers.Rows[x]["InvNumber"].ToString();
                                }
                            }
                        }
                        PackageCount = 0;
                        for (int x = 0; x < dtDistInvNumbers.Rows.Count; x++)
                        {
                            DataRow[] rows1 = ProfilReportTable.Select("InvNumber='" + InvNumber + "'");
                            if (rows1.Count() > 0)
                            {
                                if (Convert.ToInt32(rows1[0]["PackageCount"]) < PackageCount)
                                {
                                    PackageCount = Convert.ToInt32(rows1[0]["PackageCount"]);
                                    InvNumber = dtDistInvNumbers.Rows[x]["InvNumber"].ToString();
                                }
                            }
                        }
                        DataRow[] rows = ProfilReportTable.Select("InvNumber='" + InvNumber + "'");
                        if (rows.Count() > 0)
                            rows[0]["PackageCount"] = Convert.ToInt32(rows[0]["PackageCount"]) + 1;
                    }
                }



                using (DataView DV = new DataView(dtMainOrders, "FactoryID=2", string.Empty, DataViewRowState.CurrentRows))
                {
                    DistPackagesDT.Clear();
                    DistPackagesDT = DV.ToTable(true, "PackNumber");
                }
                dtGroup.Clear();
                dtGroup = GroupBy("InvNumber", "PackNumber", dtMainOrders);
                for (int i = 0; i < DistPackagesDT.Rows.Count; i++)
                {
                    using (DataView DV = new DataView(dtMainOrders, "PackNumber=" + Convert.ToInt32(DistPackagesDT.Rows[i]["PackNumber"]), string.Empty, DataViewRowState.CurrentRows))
                    {
                        dtDistInvNumbers.Clear();
                        dtDistInvNumbers = DV.ToTable(true, "InvNumber");
                        string InvNumber = dtDistInvNumbers.Rows[0]["InvNumber"].ToString();
                        int PackageCount = 0;

                        DataRow[] rows0 = dtGroup.Select("InvNumber='" + InvNumber + "'");
                        if (rows0.Count() > 0)
                            PackageCount = Convert.ToInt32(rows0[0]["Count"]);

                        for (int x = 1; x < dtDistInvNumbers.Rows.Count; x++)
                        {
                            DataRow[] rows1 = dtGroup.Select("InvNumber='" + dtDistInvNumbers.Rows[x]["InvNumber"].ToString() + "'");
                            if (rows1.Count() > 0)
                            {
                                if (Convert.ToInt32(rows1[0]["Count"]) < PackageCount)
                                {
                                    PackageCount = Convert.ToInt32(rows1[0]["Count"]);
                                    InvNumber = dtDistInvNumbers.Rows[x]["InvNumber"].ToString();
                                }
                            }
                        }
                        PackageCount = 0;
                        for (int x = 0; x < dtDistInvNumbers.Rows.Count; x++)
                        {
                            DataRow[] rows1 = TPSReportTable.Select("InvNumber='" + InvNumber + "'");
                            if (rows1.Count() > 0)
                            {
                                if (Convert.ToInt32(rows1[0]["PackageCount"]) < PackageCount)
                                {
                                    PackageCount = Convert.ToInt32(rows1[0]["PackageCount"]);
                                    InvNumber = dtDistInvNumbers.Rows[x]["InvNumber"].ToString();
                                }
                            }
                        }
                        DataRow[] rows = TPSReportTable.Select("InvNumber='" + InvNumber + "'");
                        if (rows.Count() > 0)
                            rows[0]["PackageCount"] = Convert.ToInt32(rows[0]["PackageCount"]) + 1;
                    }
                }
            }

            #region 

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
            SimpleF.FontHeightInPoints = 8;
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

            HSSFSheet sheet1 = hssfworkbook.CreateSheet("Сводный отчет");
            sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

            sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
            sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
            sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
            sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

            sheet1.SetColumnWidth(0, 16 * 256);
            sheet1.SetColumnWidth(1, 12 * 256);
            sheet1.SetColumnWidth(2, 71 * 256);
            sheet1.SetColumnWidth(3, 8 * 256);
            sheet1.SetColumnWidth(4, 15 * 256);
            sheet1.SetColumnWidth(5, 12 * 256);
            sheet1.SetColumnWidth(6, 17 * 256);
            sheet1.SetColumnWidth(7, 15 * 256);
            sheet1.SetColumnWidth(8, 10 * 256);

            HSSFCell Cell1 = sheet1.CreateRow(pos).CreateCell(0);
            Cell1.SetCellValue("Клиент:");
            Cell1.CellStyle = ReportCS2;

            Cell1 = sheet1.CreateRow(pos++).CreateCell(1);
            Cell1.SetCellValue(ClientName);
            Cell1.CellStyle = ReportCS2;

            Cell1 = sheet1.CreateRow(pos).CreateCell(0);
            Cell1.SetCellValue("№ заказов:");
            Cell1.CellStyle = ReportCS1;

            Cell1 = sheet1.CreateRow(pos++).CreateCell(1);
            Cell1.SetCellValue(string.Join(", ", OrderNumbers));
            Cell1.CellStyle = ReportCS1;

            int RowCount = 1;


            Cell1 = sheet1.CreateRow(pos++).CreateCell(0);
            Cell1.SetCellValue("№ подзаказов: ");
            Cell1.CellStyle = SummaryWithoutBorderBelCS;

            for (int i = 0; i < MainOrdersIDs.Count(); i++)
            {
                MainOrdersList += MainOrdersIDs[i].ToString() + ", ";

                if (i > 0 && i % 9 == 0)
                {
                    Cell1 = sheet1.CreateRow(pos++).CreateCell(0);
                    Cell1.SetCellValue(MainOrdersList);
                    Cell1.CellStyle = SummaryWithoutBorderBelCS;

                    MainOrdersList = string.Empty;
                    RowCount++;
                }
                if (i == MainOrdersIDs.Count() - 1)
                {
                    if (MainOrdersList.Length > 0)
                        MainOrdersList = MainOrdersList.Remove(MainOrdersList.Length - 2, 2);
                    Cell1 = sheet1.CreateRow(pos++).CreateCell(0);
                    Cell1.SetCellValue(MainOrdersList);
                    Cell1.CellStyle = SummaryWithoutBorderBelCS;

                    MainOrdersList = string.Empty;
                    RowCount++;
                }
            }

            if (IsComplaint(GetMegaOrderID(MainOrdersIDs[0])))
            {
                Cell1 = sheet1.CreateRow(pos++).CreateCell(0);
                Cell1.SetCellValue("РЕКЛАМАЦИЯ");
                Cell1.CellStyle = ReportCS1;
            }

            int displayIndex = 0;
            if (ProfilReportTable.Rows.Count > 0)
            {
                //Профиль
                pos += 2;
                displayIndex = 0;

                Cell1 = sheet1.CreateRow(pos++).CreateCell(0);
                Cell1.SetCellValue("ОМЦ-ПРОФИЛЬ:");
                Cell1.CellStyle = SummaryWithoutBorderBelCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("УНН");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Инв.№");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Наименование");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Кол-во");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Цена с транс, " + Currency);
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Нестандарт, %");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Стоимость с транс, " + Currency);
                Cell1.CellStyle = SimpleHeaderCS;

                //Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                //Cell1.SetCellValue("Курс, " + Currency);
                //Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Вес, кг. ");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Кол-во уп. ");
                Cell1.CellStyle = SimpleHeaderCS;
                pos++;

                for (int i = 0; i < ProfilReportTable.Rows.Count; i++)
                {
                    displayIndex = 0;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(ProfilReportTable.Rows[i]["UNN"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(ProfilReportTable.Rows[i]["InvNumber"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(ProfilReportTable.Rows[i]["AccountingName"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["Count"]));
                    Cell1.CellStyle = CountCS;
                    if (CurrencyTypeID == 0)
                    {
                        Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                        Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["PriceWithTransport"]));
                        Cell1.CellStyle = PriceBelCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                        if (ProfilReportTable.Rows[i]["NonStandardMargin"] != DBNull.Value)
                        {
                            Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["NonStandardMargin"]));
                        }
                        Cell1.CellStyle = CountCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                        Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["CostWithTransport"]));
                        Cell1.CellStyle = PriceBelCS;
                        //Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                        //Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["PaymentRate"]));
                        //Cell1.CellStyle = CurrencyCS;
                    }
                    else
                    {
                        Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                        Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["PriceWithTransport"]));
                        Cell1.CellStyle = PriceForeignCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                        if (ProfilReportTable.Rows[i]["NonStandardMargin"] != DBNull.Value)
                        {
                            Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["NonStandardMargin"]));
                        }
                        Cell1.CellStyle = CountCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                        Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["CostWithTransport"]));
                        Cell1.CellStyle = PriceForeignCS;
                        //Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                        //Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["PaymentRate"]));
                        //Cell1.CellStyle = CurrencyCS;
                    }
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["Weight"]));
                    Cell1.CellStyle = WeightCS;
                    if (ProfilReportTable.Rows[i]["PackageCount"] != DBNull.Value)
                    {
                        Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                        Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["PackageCount"]));
                        Cell1.CellStyle = SimpleCS;
                    }
                    pos++;
                }

                if (CurrencyTypeID == 0)
                {
                    Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                    Cell1.SetCellValue("Заказ, " + Currency + ":");
                    Cell1.CellStyle = SummaryWithoutBorderBelCS;
                    Cell1 = sheet1.CreateRow(pos++).CreateCell(6);
                    Cell1.SetCellValue(Convert.ToDouble(TotalProfil));
                    Cell1.CellStyle = SummaryWithoutBorderBelCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                    Cell1.SetCellValue("Транспорт, прочее, " + Currency + ":");
                    Cell1.CellStyle = SummaryWithoutBorderBelCS;
                    Cell1 = sheet1.CreateRow(pos++).CreateCell(6);
                    Cell1.SetCellValue(Convert.ToDouble(TransportAndOtherProfil));
                    Cell1.CellStyle = SummaryWithoutBorderBelCS;
                    Cell1 = sheet1.CreateRow(pos - 2).CreateCell(7);
                    Cell1.SetCellValue(Convert.ToDouble(WeightProfil));
                    Cell1.CellStyle = SummaryWeightCS;
                }
                else
                {
                    Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                    Cell1.SetCellValue("Заказ, " + Currency + ":");
                    Cell1.CellStyle = SummaryWithoutBorderForeignCS;
                    Cell1 = sheet1.CreateRow(pos++).CreateCell(6);
                    Cell1.SetCellValue(Convert.ToDouble(TotalProfil));
                    Cell1.CellStyle = SummaryWithoutBorderForeignCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                    Cell1.SetCellValue("Транспорт, прочее, " + Currency + ":");
                    Cell1.CellStyle = SummaryWithoutBorderForeignCS;
                    Cell1 = sheet1.CreateRow(pos++).CreateCell(6);
                    Cell1.SetCellValue(Convert.ToDouble(TransportAndOtherProfil));
                    Cell1.CellStyle = SummaryWithoutBorderForeignCS;
                    Cell1 = sheet1.CreateRow(pos - 2).CreateCell(7);
                    Cell1.SetCellValue(Convert.ToDouble(WeightProfil));
                    Cell1.CellStyle = SummaryWeightCS;
                }
            }

            if (TPSReportTable.Rows.Count > 0)
            {
                //ТПС
                pos++;

                Cell1 = sheet1.CreateRow(pos++).CreateCell(0);
                Cell1.SetCellValue("ЗОВ-ТПС:");
                Cell1.CellStyle = SummaryWithoutBorderBelCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                Cell1.SetCellValue("УНН");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(1);
                Cell1.SetCellValue("Инв.№");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(2);
                Cell1.SetCellValue("Наименование");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(3);
                Cell1.SetCellValue("Кол-во");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(4);
                Cell1.SetCellValue("Цена с транс, " + Currency);
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(5);
                Cell1.SetCellValue("Нестандарт, %");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(6);
                Cell1.SetCellValue("Стоимость с транс, " + Currency);
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(7);
                Cell1.SetCellValue("Курс, " + Currency);
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(8);
                Cell1.SetCellValue("Вес, кг. ");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(9);
                Cell1.SetCellValue("Кол-во уп. ");
                Cell1.CellStyle = SimpleHeaderCS;
                pos++;

                for (int i = 0; i < TPSReportTable.Rows.Count; i++)
                {
                    Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                    Cell1.SetCellValue(TPSReportTable.Rows[i]["UNN"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(1);
                    Cell1.SetCellValue(TPSReportTable.Rows[i]["InvNumber"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(2);
                    Cell1.SetCellValue(TPSReportTable.Rows[i]["AccountingName"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(3);
                    Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["Count"]));
                    Cell1.CellStyle = CountCS;
                    if (CurrencyTypeID == 0)
                    {
                        Cell1 = sheet1.CreateRow(pos).CreateCell(4);
                        Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["PriceWithTransport"]));
                        Cell1.CellStyle = PriceBelCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(5);
                        if (TPSReportTable.Rows[i]["NonStandardMargin"] != DBNull.Value)
                        {
                            Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["NonStandardMargin"]));
                        }
                        Cell1.CellStyle = CountCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(6);
                        Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["Cost"]));
                        Cell1.CellStyle = PriceBelCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(7);
                        Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["PaymentRate"]));
                        Cell1.CellStyle = CurrencyCS;
                    }
                    else
                    {
                        Cell1 = sheet1.CreateRow(pos).CreateCell(4);
                        Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["PriceWithTransport"]));
                        Cell1.CellStyle = PriceForeignCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(5);
                        if (TPSReportTable.Rows[i]["NonStandardMargin"] != DBNull.Value)
                        {
                            Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["NonStandardMargin"]));
                        }
                        Cell1.CellStyle = CountCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(6);
                        Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["CostWithTransport"]));
                        Cell1.CellStyle = PriceForeignCS;
                        Cell1 = sheet1.CreateRow(pos).CreateCell(7);
                        Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["PaymentRate"]));
                        Cell1.CellStyle = CurrencyCS;
                    }
                    Cell1 = sheet1.CreateRow(pos).CreateCell(8);
                    Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["Weight"]));
                    Cell1.CellStyle = WeightCS;
                    if (TPSReportTable.Rows[i]["PackageCount"] != DBNull.Value)
                    {
                        Cell1 = sheet1.CreateRow(pos).CreateCell(9);
                        Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["PackageCount"]));
                        Cell1.CellStyle = SimpleCS;
                    }
                    pos++;
                }

                if (CurrencyTypeID == 0)
                {
                    Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                    Cell1.SetCellValue("Заказ, " + Currency + ":");
                    Cell1.CellStyle = SummaryWithoutBorderBelCS;
                    Cell1 = sheet1.CreateRow(pos++).CreateCell(6);
                    Cell1.SetCellValue(Convert.ToDouble(TotalTPS));
                    Cell1.CellStyle = SummaryWithoutBorderBelCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                    Cell1.SetCellValue("Транспорт, прочее, " + Currency + ":");
                    Cell1.CellStyle = SummaryWithoutBorderBelCS;
                    Cell1 = sheet1.CreateRow(pos++).CreateCell(6);
                    Cell1.SetCellValue(Convert.ToDouble(TransportAndOtherTPS));
                    Cell1.CellStyle = SummaryWithoutBorderBelCS;
                    Cell1 = sheet1.CreateRow(pos - 2).CreateCell(7);
                    Cell1.SetCellValue(Convert.ToDouble(WeightTPS));
                    Cell1.CellStyle = SummaryWeightCS;
                }
                else
                {
                    Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                    Cell1.SetCellValue("Заказ, " + Currency + ":");
                    Cell1.CellStyle = SummaryWithoutBorderForeignCS;
                    Cell1 = sheet1.CreateRow(pos++).CreateCell(6);
                    Cell1.SetCellValue(Convert.ToDouble(TotalTPS));
                    Cell1.CellStyle = SummaryWithoutBorderForeignCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                    Cell1.SetCellValue("Транспорт, прочее, " + Currency + ":");
                    Cell1.CellStyle = SummaryWithoutBorderForeignCS;
                    Cell1 = sheet1.CreateRow(pos++).CreateCell(6);
                    Cell1.SetCellValue(Convert.ToDouble(TransportAndOtherTPS));
                    Cell1.CellStyle = SummaryWithoutBorderForeignCS;
                    Cell1 = sheet1.CreateRow(pos - 2).CreateCell(7);
                    Cell1.SetCellValue(Convert.ToDouble(WeightTPS));
                    Cell1.CellStyle = SummaryWeightCS;
                }
            }
            pos++;

            //if (PaymentRate != 1)
            //{
            //    Cell1 = sheet1.CreateRow(pos++).CreateCell(0);
            //    Cell1.SetCellValue("Курс, 1 EUR = " + PaymentRate + " " + Currency);
            //    Cell1.CellStyle = SummaryWithoutBorderBelCS;
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
            #endregion
        }

        public void SaveDBF(string FilePath, string DBFName, bool IsSample, ref string ProfilDBFName, ref string TPSDBFName)
        {
            string CurrentMonthName = DateTime.Now.ToString("dd-MM-yyyy");
            if (Directory.Exists(FilePath) == false)//not exists
            {
                Directory.CreateDirectory(FilePath);
            }
            if (ProfilReportTable.Rows.Count > 0 && TPSReportTable.Rows.Count > 0)
            {
                ProfilDBFName = DBFName + " (Профиль-ТПС)";
                if (IsSample)
                    ProfilDBFName = DBFName + " (Профиль-ТПС), обр";
                FileInfo f = new FileInfo(FilePath + @"\" + ProfilDBFName + ".DBF");
                int x = 1;
                while (f.Exists == true)
                    f = new FileInfo(FilePath + @"\" + ProfilDBFName + "(" + x++ + ").DBF");
                ProfilDBFName = f.FullName;
                TPSDBFName = f.FullName;

                DataSetIntoDBF(FilePath, @"PT", ProfilReportTable, TPSReportTable, 0);
                var sourcePath = Path.Combine(FilePath, "PT.DBF");
                var destinationPath = ProfilDBFName;
                File.Move(sourcePath, destinationPath);
            }
            if (ProfilReportTable.Rows.Count > 0 && TPSReportTable.Rows.Count == 0)
            {
                ProfilDBFName = DBFName + " (ОМЦ-ПРОФИЛЬ)";
                if (IsSample)
                    ProfilDBFName = DBFName + " (ОМЦ-ПРОФИЛЬ), обр";
                FileInfo f = new FileInfo(FilePath + @"\" + ProfilDBFName + ".DBF");
                int x = 1;
                while (f.Exists == true)
                    f = new FileInfo(FilePath + @"\" + ProfilDBFName + "(" + x++ + ").DBF");
                ProfilDBFName = f.FullName;

                DataSetIntoDBF(FilePath, @"Profil", ProfilReportTable, 1);
                var sourcePath = Path.Combine(FilePath, "Profil.DBF");
                var destinationPath = ProfilDBFName;
                File.Move(sourcePath, destinationPath);
            }
            if (TPSReportTable.Rows.Count > 0 && ProfilReportTable.Rows.Count == 0)
            {
                TPSDBFName = DBFName + " (ЗОВ-ТПС)";
                if (IsSample)
                    TPSDBFName = DBFName + " (ЗОВ-ТПС), обр";
                FileInfo f = new FileInfo(FilePath + @"\" + TPSDBFName + ".DBF");
                int x = 1;
                while (f.Exists == true)
                    f = new FileInfo(FilePath + @"\" + TPSDBFName + "(" + x++ + ").DBF");
                TPSDBFName = f.FullName;

                DataSetIntoDBF(FilePath, @"TPS", TPSReportTable, 2);
                var sourcePath = Path.Combine(FilePath, "TPS.DBF");
                var destinationPath = TPSDBFName;
                File.Move(sourcePath, destinationPath);
            }
        }

        public void SaveDBF(string FilePath, string DBFName, ref string ProfilDBFName, ref string TPSDBFName)
        {
            string CurrentMonthName = DateTime.Now.ToString("dd-MM-yyyy");
            if (Directory.Exists(FilePath) == false)//not exists
            {
                Directory.CreateDirectory(FilePath);
            }
            if (ProfilReportTable.Rows.Count > 0 && TPSReportTable.Rows.Count > 0)
            {
                ProfilDBFName = DBFName + " (Профиль-ТПС)";
                FileInfo f = new FileInfo(FilePath + @"\" + ProfilDBFName + ".DBF");
                int x = 1;
                while (f.Exists == true)
                    f = new FileInfo(FilePath + @"\" + ProfilDBFName + "(" + x++ + ").DBF");
                ProfilDBFName = f.FullName;
                TPSDBFName = f.FullName;

                DataSetIntoDBF(FilePath, @"PT", ProfilReportTable, TPSReportTable, 0);
                var sourcePath = Path.Combine(FilePath, "PT.DBF");
                var destinationPath = ProfilDBFName;
                File.Move(sourcePath, destinationPath);
            }
            if (ProfilReportTable.Rows.Count > 0 && TPSReportTable.Rows.Count == 0)
            {
                ProfilDBFName = DBFName + " (ОМЦ-ПРОФИЛЬ)";
                FileInfo f = new FileInfo(FilePath + @"\" + ProfilDBFName + ".DBF");
                int x = 1;
                while (f.Exists == true)
                    f = new FileInfo(FilePath + @"\" + ProfilDBFName + "(" + x++ + ").DBF");
                ProfilDBFName = f.FullName;

                DataSetIntoDBF(FilePath, @"Profil", ProfilReportTable, 1);
                var sourcePath = Path.Combine(FilePath, "Profil.DBF");
                var destinationPath = ProfilDBFName;
                File.Move(sourcePath, destinationPath);
            }
            if (TPSReportTable.Rows.Count > 0 && ProfilReportTable.Rows.Count == 0)
            {
                TPSDBFName = DBFName + " (ЗОВ-ТПС)";
                FileInfo f = new FileInfo(FilePath + @"\" + TPSDBFName + ".DBF");
                int x = 1;
                while (f.Exists == true)
                    f = new FileInfo(FilePath + @"\" + TPSDBFName + "(" + x++ + ").DBF");
                TPSDBFName = f.FullName;

                DataSetIntoDBF(FilePath, @"TPS", TPSReportTable, 2);
                var sourcePath = Path.Combine(FilePath, "TPS.DBF");
                var destinationPath = TPSDBFName;
                File.Move(sourcePath, destinationPath);
            }
        }

        public void ClearReport()
        {
            ProfilReportTable.Clear();
            TPSReportTable.Clear();
            FrontsReport.ClearReport();
            DecorReport.ClearReport();
        }

        public static void DataSetIntoDBF(string path, string fileName, DataTable DT1, int FactoryID)
        {
            if (File.Exists(path + "/"+ fileName + ".dbf"))
            {
                File.Delete(path + "/" + fileName + ".dbf");
            }

            string createSql = "create table " + fileName + " ([UNNP] varchar(20), [UNN] varchar(20), [CurrencyCode] varchar(20), [InvNumber] varchar(20), [Amount] Double, [Price] Double, [NDS] varchar(10), [Weight] Double, [PackageCount] Integer)";

            OleDbConnection con = new OleDbConnection(GetConnection(path));

            OleDbCommand cmd = new OleDbCommand()
            {
                Connection = con
            };
            con.Open();

            cmd.CommandText = createSql;

            cmd.ExecuteNonQuery();

            foreach (DataRow row in DT1.Rows)
            {
                string insertSql = "insert into " + fileName +
                    " (UNNP, UNN, CurrencyCode, InvNumber, Amount, Price, NDS, Weight, PackageCount) values(UNNP, UNN, CurrencyCode, InvNumber, Amount, PriceWithTransport, NDS, Weight, PackageCount)";
                decimal d = Decimal.Round(Convert.ToDecimal(row["Weight"]) / 1000, 3, MidpointRounding.AwayFromZero);
                double Amount = Convert.ToDouble(row["Count"]);
                double Price = Convert.ToDouble(row["PriceWithTransport"]);
                double Weight = Convert.ToDouble(d);
                int PackageCount = 0;
                if (row["PackageCount"] != DBNull.Value)
                    PackageCount = Convert.ToInt32(row["PackageCount"]);
                string InvNumber = row["InvNumber"].ToString();
                string CurrencyCode = row["CurrencyCode"].ToString();
                //string TPSCurCode = row["TPSCurCode"].ToString();
                string UNN = row["UNN"].ToString();
                string NDS = "0 %";
                string UNNP = "800014979";
                cmd.CommandText = insertSql;
                cmd.Parameters.Clear();

                if (Convert.ToInt32(CurrencyCode) == 974 || Convert.ToInt32(CurrencyCode) == 933)
                    NDS = "20 %";
                if (FactoryID == 2)
                    UNNP = "590618616";
                cmd.Parameters.Add("UNNP", OleDbType.VarChar).Value = UNNP;
                cmd.Parameters.Add("UNN", OleDbType.VarChar).Value = UNN;
                cmd.Parameters.Add("CurrencyCode", OleDbType.VarChar).Value = CurrencyCode;
                //cmd.Parameters.Add("TPSCurCode", OleDbType.VarChar).Value = TPSCurCode;
                cmd.Parameters.Add("InvNumber", OleDbType.VarChar).Value = InvNumber;
                cmd.Parameters.Add("Amount", OleDbType.Double).Value = Amount;
                cmd.Parameters.Add("PriceWithTransport", OleDbType.Double).Value = Price;
                cmd.Parameters.Add("NDS", OleDbType.VarChar).Value = NDS;
                cmd.Parameters.Add("Weight", OleDbType.Double).Value = Weight;
                cmd.Parameters.Add("PackageCount", OleDbType.Integer).Value = PackageCount;
                cmd.ExecuteNonQuery();
            }

            con.Close();
        }

        public static void DataSetIntoDBF(string path, string fileName, DataTable DT1, DataTable DT2, int FactoryID)
        {
            if (File.Exists(path + "/"+ fileName + ".dbf"))
            {
                File.Delete(path + "/" + fileName + ".dbf");
            }

            string createSql = "create table " + fileName + " ([UNNP] varchar(20), [UNN] varchar(20), [CurrencyCode] varchar(20), [InvNumber] varchar(20), [Amount] Double, [Price] Double, [NDS] varchar(10), [Weight] Double, [PackageCount] Integer)";

            OleDbConnection con = new OleDbConnection(GetConnection(path));

            OleDbCommand cmd = new OleDbCommand()
            {
                Connection = con
            };
            con.Open();

            cmd.CommandText = createSql;

            cmd.ExecuteNonQuery();

            foreach (DataRow row in DT1.Rows)
            {
                string insertSql = "insert into " + fileName +
                    " (UNNP, UNN, CurrencyCode, InvNumber, Amount, Price, NDS, Weight, PackageCount) values(UNNP, UNN, CurrencyCode, InvNumber, Amount, PriceWithTransport, NDS, Weight, PackageCount)";
                decimal d = Decimal.Round(Convert.ToDecimal(row["Weight"]) / 1000, 3, MidpointRounding.AwayFromZero);
                double Amount = Convert.ToDouble(row["Count"]);
                double Price = Convert.ToDouble(row["PriceWithTransport"]);
                double Weight = Convert.ToDouble(d);
                int PackageCount = 0;
                if (row["PackageCount"] != DBNull.Value)
                    PackageCount = Convert.ToInt32(row["PackageCount"]);
                string InvNumber = row["InvNumber"].ToString();
                string CurrencyCode = row["CurrencyCode"].ToString();
                //string TPSCurCode = row["TPSCurCode"].ToString();
                string UNN = row["UNN"].ToString();
                string NDS = "0 %";
                string UNNP = "800014979";
                cmd.CommandText = insertSql;
                cmd.Parameters.Clear();

                if (Convert.ToInt32(CurrencyCode) == 974 || Convert.ToInt32(CurrencyCode) == 933)
                    NDS = "20 %";
                //if (FactoryID == 2)
                //    UNNP = "590618616";
                cmd.Parameters.Add("UNNP", OleDbType.VarChar).Value = UNNP;
                cmd.Parameters.Add("UNN", OleDbType.VarChar).Value = UNN;
                cmd.Parameters.Add("CurrencyCode", OleDbType.VarChar).Value = CurrencyCode;
                //cmd.Parameters.Add("TPSCurCode", OleDbType.VarChar).Value = TPSCurCode;
                cmd.Parameters.Add("InvNumber", OleDbType.VarChar).Value = InvNumber;
                cmd.Parameters.Add("Amount", OleDbType.Double).Value = Amount;
                cmd.Parameters.Add("PriceWithTransport", OleDbType.Double).Value = Price;
                cmd.Parameters.Add("NDS", OleDbType.VarChar).Value = NDS;
                cmd.Parameters.Add("Weight", OleDbType.Double).Value = Weight;
                cmd.Parameters.Add("PackageCount", OleDbType.Integer).Value = PackageCount;
                cmd.ExecuteNonQuery();
            }
            foreach (DataRow row in DT2.Rows)
            {
                string insertSql = "insert into " + fileName +
                    " (UNNP, UNN, CurrencyCode, InvNumber, Amount, Price, NDS, Weight, PackageCount) values(UNNP, UNN, CurrencyCode, InvNumber, Amount, PriceWithTransport, NDS, Weight, PackageCount)";
                decimal d = Decimal.Round(Convert.ToDecimal(row["Weight"]) / 1000, 3, MidpointRounding.AwayFromZero);
                double Amount = Convert.ToDouble(row["Count"]);
                double Price = Convert.ToDouble(row["PriceWithTransport"]);
                double Weight = Convert.ToDouble(d);
                int PackageCount = 0;
                if (row["PackageCount"] != DBNull.Value)
                    PackageCount = Convert.ToInt32(row["PackageCount"]);
                string InvNumber = row["InvNumber"].ToString();
                string CurrencyCode = row["CurrencyCode"].ToString();
                //string TPSCurCode = row["TPSCurCode"].ToString();
                string UNN = row["UNN"].ToString();
                string NDS = "0 %";
                string UNNP = "590618616";
                cmd.CommandText = insertSql;
                cmd.Parameters.Clear();

                if (Convert.ToInt32(CurrencyCode) == 974 || Convert.ToInt32(CurrencyCode) == 933)
                    NDS = "20 %";
                //if (FactoryID == 2)
                //    UNNP = "590618616";
                cmd.Parameters.Add("UNNP", OleDbType.VarChar).Value = UNNP;
                cmd.Parameters.Add("UNN", OleDbType.VarChar).Value = UNN;
                cmd.Parameters.Add("CurrencyCode", OleDbType.VarChar).Value = CurrencyCode;
                //cmd.Parameters.Add("TPSCurCode", OleDbType.VarChar).Value = TPSCurCode;
                cmd.Parameters.Add("InvNumber", OleDbType.VarChar).Value = InvNumber;
                cmd.Parameters.Add("Amount", OleDbType.Double).Value = Amount;
                cmd.Parameters.Add("PriceWithTransport", OleDbType.Double).Value = Price;
                cmd.Parameters.Add("NDS", OleDbType.VarChar).Value = NDS;
                cmd.Parameters.Add("Weight", OleDbType.Double).Value = Weight;
                cmd.Parameters.Add("PackageCount", OleDbType.Integer).Value = PackageCount;
                cmd.ExecuteNonQuery();
            }

            con.Close();
        }

        private static string GetConnection(string path)
        {
            return "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + ";Extended Properties=dBASE IV;";
        }

        public static string ReplaceEscape(string str)
        {
            str = str.Replace("'", "''");
            return str;
        }
    }

}
