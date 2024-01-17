using System;
using System.Data;
using System.Data.SqlClient;
using System.Linq;

namespace Infinium.Modules.Marketing.NewOrders.InvoiceReportToDbf
{
    public class InvoiceDecorReportToDbf
    {
        public DataTable ProfilReportDataTable = null;

        public DataTable TPSReportDataTable = null;

        private decimal AdditionalCost = 0;

        private int ClientID = 0;
        
        private DataTable DecorOrdersDataTable = null;

        private decimal PaymentRate = 1;

        private string ProfilCurrencyCode = "0";

        private string TPSCurrencyCode = "0";

        private decimal TransportCost = 0;

        private string UNN = string.Empty;

        private DecorCatalogOrder decorCatalogOrder;

        public InvoiceDecorReportToDbf(DecorCatalogOrder DecorCatalogOrder)
        {
            decorCatalogOrder = DecorCatalogOrder;
            Create();
            CreateReportDataTable();
        }

        public void ClearReport()
        {
            DecorOrdersDataTable.Clear();

            ProfilReportDataTable.Clear();
            TPSReportDataTable.Clear();

            ProfilReportDataTable.AcceptChanges();
            TPSReportDataTable.AcceptChanges();
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
            DT.Columns.Add(new DataColumn("Count", Type.GetType("System.Decimal")));
            DT.Columns.Add(new DataColumn("TotalCount", Type.GetType("System.Int32")));
            DT.Columns.Add(new DataColumn("Cost", Type.GetType("System.Decimal")));
            DT.Columns.Add(new DataColumn("PriceWithTransport", Type.GetType("System.String")));
            DT.Columns.Add(new DataColumn("CostWithTransport", Type.GetType("System.String")));
            DT.Columns.Add(new DataColumn("Weight", Type.GetType("System.Decimal")));
            DT.Columns.Add(new DataColumn("PaymentRate", Type.GetType("System.Decimal")));
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

            DataRow[] Row = decorCatalogOrder.DecorConfigDataTable.Select("DecorConfigID = " + DecorOrderRow["DecorConfigID"].ToString());

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

        public void GetMegaOrderInfo(int MainOrderID)
        {
            string SelectCommand = "SELECT MegaOrderID, TransportCost, AdditionalCost, PaymentRate, ClientID, CurrencyTypeID FROM MegaOrders" +
                " WHERE MegaOrderID IN (SELECT MegaOrderID FROM MainOrders WHERE MainOrderID = " + MainOrderID + ")";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    if (DA.Fill(DT) > 0)
                    {
                        int CurrencyTypeID = 1;
                        if (DT.Rows[0]["TransportCost"] != DBNull.Value)
                            decimal.TryParse(DT.Rows[0]["TransportCost"].ToString(), out TransportCost);
                        if (DT.Rows[0]["AdditionalCost"] != DBNull.Value)
                            decimal.TryParse(DT.Rows[0]["AdditionalCost"].ToString(), out AdditionalCost);
                        //if (DT.Rows[0]["PaymentRate"] != DBNull.Value)
                        //    decimal.TryParse(DT.Rows[0]["PaymentRate"].ToString(), out PaymentRate);
                        if (DT.Rows[0]["ClientID"] != DBNull.Value)
                            int.TryParse(DT.Rows[0]["ClientID"].ToString(), out ClientID);
                        if (DT.Rows[0]["CurrencyTypeID"] != DBNull.Value)
                            int.TryParse(DT.Rows[0]["CurrencyTypeID"].ToString(), out CurrencyTypeID);

                        DataRow[] rows = decorCatalogOrder.CurrencyTypesDT.Select("CurrencyTypeID = " + CurrencyTypeID);
                        if (rows.Count() > 0)
                        {
                            ProfilCurrencyCode = rows[0]["CurrencyCode"].ToString();
                            TPSCurrencyCode = rows[0]["TPSCurrencyCode"].ToString();
                        }
                    }
                }
            }
            SelectCommand = "SELECT UNN FROM Clients WHERE ClientID = " + ClientID;
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
        }

        public void GetMegaOrderInfo1(int MegaOrderID)
        {
            string SelectCommand = "SELECT MegaOrderID, TransportCost, AdditionalCost, PaymentRate, ClientID, CurrencyTypeID FROM MegaOrders" +
                " WHERE MegaOrderID = " + MegaOrderID;
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    if (DA.Fill(DT) > 0)
                    {
                        int CurrencyTypeID = 1;
                        if (DT.Rows[0]["TransportCost"] != DBNull.Value)
                            decimal.TryParse(DT.Rows[0]["TransportCost"].ToString(), out TransportCost);
                        if (DT.Rows[0]["AdditionalCost"] != DBNull.Value)
                            decimal.TryParse(DT.Rows[0]["AdditionalCost"].ToString(), out AdditionalCost);
                        //if (DT.Rows[0]["PaymentRate"] != DBNull.Value)
                        //    decimal.TryParse(DT.Rows[0]["PaymentRate"].ToString(), out PaymentRate);
                        if (DT.Rows[0]["ClientID"] != DBNull.Value)
                            int.TryParse(DT.Rows[0]["ClientID"].ToString(), out ClientID);
                        if (DT.Rows[0]["CurrencyTypeID"] != DBNull.Value)
                            int.TryParse(DT.Rows[0]["CurrencyTypeID"].ToString(), out CurrencyTypeID);

                        DataRow[] rows = decorCatalogOrder.CurrencyTypesDT.Select("CurrencyTypeID = " + CurrencyTypeID);
                        if (rows.Count() > 0)
                        {
                            ProfilCurrencyCode = rows[0]["CurrencyCode"].ToString();
                            TPSCurrencyCode = rows[0]["TPSCurrencyCode"].ToString();
                        }
                    }
                }
            }
            SelectCommand = "SELECT UNN FROM Clients WHERE ClientID = " + ClientID;
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
        }

        public bool HasParameter(int ProductID, String Parameter)
        {
            DataRow[] Rows = decorCatalogOrder.DecorParametersDataTable.Select("ProductID = " + ProductID);

            return Convert.ToBoolean(Rows[0][Parameter]);
        }

        public void Report(int MegaOrderID, decimal dPaymentRate, bool IsSample)
        {
            DecorOrdersDataTable.Clear();
            ProfilReportDataTable.Clear();
            TPSReportDataTable.Clear();
            GetMegaOrderInfo1(MegaOrderID);

            string SelectCommand = "SELECT NewDecorOrders.*, infiniu2_catalog.dbo.DecorConfig.AccountingName, infiniu2_catalog.dbo.DecorConfig.InvNumber, NewMegaOrders.PaymentRate FROM NewDecorOrders" +
                " INNER JOIN NewMainOrders ON NewDecorOrders.MainOrderID = NewMainOrders.MainOrderID" +
                " INNER JOIN NewMegaOrders ON NewMainOrders.MegaOrderID = NewMegaOrders.MegaOrderID AND NewMegaOrders.MegaOrderID=" + MegaOrderID +
                " INNER JOIN infiniu2_catalog.dbo.DecorConfig ON NewDecorOrders.DecorConfigID = infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                " WHERE NeedCalcPrice=0 AND NewDecorOrders.IsSample=1 AND InvNumber IS NOT NULL ORDER BY InvNumber";
            if (!IsSample)
                SelectCommand = "SELECT NewDecorOrders.*, infiniu2_catalog.dbo.DecorConfig.AccountingName, infiniu2_catalog.dbo.DecorConfig.InvNumber, NewMegaOrders.PaymentRate FROM NewDecorOrders" +
                    " INNER JOIN NewMainOrders ON NewDecorOrders.MainOrderID = NewMainOrders.MainOrderID" +
                    " INNER JOIN NewMegaOrders ON NewMainOrders.MegaOrderID = NewMegaOrders.MegaOrderID AND NewMegaOrders.MegaOrderID=" + MegaOrderID +
                    " INNER JOIN infiniu2_catalog.dbo.DecorConfig ON NewDecorOrders.DecorConfigID = infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                    " WHERE NeedCalcPrice=0 AND NewDecorOrders.IsSample=0 AND InvNumber IS NOT NULL ORDER BY InvNumber";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(DecorOrdersDataTable);
            }

            for (int j = 0; j < DecorOrdersDataTable.Rows.Count; j++)
            {
                DecorOrdersDataTable.Rows[j]["PaymentRate"] = dPaymentRate;
            }
            Collect();
        }

        public void Report(int MegaOrderID, decimal dPaymentRate)
        {
            DecorOrdersDataTable.Clear();
            ProfilReportDataTable.Clear();
            TPSReportDataTable.Clear();
            GetMegaOrderInfo1(MegaOrderID);

            string SelectCommand = "SELECT NewDecorOrders.*, infiniu2_catalog.dbo.DecorConfig.AccountingName, infiniu2_catalog.dbo.DecorConfig.InvNumber, NewMegaOrders.PaymentRate FROM NewDecorOrders" +
                " INNER JOIN NewMainOrders ON NewDecorOrders.MainOrderID = NewMainOrders.MainOrderID" +
                " INNER JOIN NewMegaOrders ON NewMainOrders.MegaOrderID = NewMegaOrders.MegaOrderID AND NewMegaOrders.MegaOrderID=" + MegaOrderID +
                " INNER JOIN infiniu2_catalog.dbo.DecorConfig ON NewDecorOrders.DecorConfigID = infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                " WHERE InvNumber IS NOT NULL ORDER BY InvNumber";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(DecorOrdersDataTable);
            }

            for (int j = 0; j < DecorOrdersDataTable.Rows.Count; j++)
            {
                DecorOrdersDataTable.Rows[j]["PaymentRate"] = dPaymentRate;
            }
            Collect();
        }

        public void Report(int[] MainOrderIDs, bool IsSample)
        {
            DecorOrdersDataTable.Clear();
            ProfilReportDataTable.Clear();
            TPSReportDataTable.Clear();
            GetMegaOrderInfo(MainOrderIDs[0]);

            string SelectCommand = "SELECT DecorOrders.*, infiniu2_catalog.dbo.DecorConfig.AccountingName, infiniu2_catalog.dbo.DecorConfig.InvNumber, MegaOrders.PaymentRate FROM DecorOrders" +
                " INNER JOIN MainOrders ON DecorOrders.MainOrderID = MainOrders.MainOrderID" +
                " INNER JOIN MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID" +
                " INNER JOIN infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID = infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                " WHERE DecorOrders.IsSample=1 AND InvNumber IS NOT NULL AND DecorOrders.MainOrderID IN (" + string.Join(",", MainOrderIDs) + ") ORDER BY InvNumber";
            if (!IsSample)
                SelectCommand = "SELECT DecorOrders.*, infiniu2_catalog.dbo.DecorConfig.AccountingName, infiniu2_catalog.dbo.DecorConfig.InvNumber, MegaOrders.PaymentRate FROM DecorOrders" +
                    " INNER JOIN MainOrders ON DecorOrders.MainOrderID = MainOrders.MainOrderID" +
                    " INNER JOIN MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID" +
                    " INNER JOIN infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID = infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                    " WHERE DecorOrders.IsSample=0 AND InvNumber IS NOT NULL AND DecorOrders.MainOrderID IN (" + string.Join(",", MainOrderIDs) + ") ORDER BY InvNumber";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(DecorOrdersDataTable);
            }

            Collect();
        }

        public void Report(int[] MainOrderIDs)
        {
            DecorOrdersDataTable.Clear();
            ProfilReportDataTable.Clear();
            TPSReportDataTable.Clear();
            GetMegaOrderInfo(MainOrderIDs[0]);

            string SelectCommand = "SELECT DecorOrders.*, infiniu2_catalog.dbo.DecorConfig.AccountingName, infiniu2_catalog.dbo.DecorConfig.InvNumber, MegaOrders.PaymentRate FROM DecorOrders" +
                " INNER JOIN MainOrders ON DecorOrders.MainOrderID = MainOrders.MainOrderID" +
                " INNER JOIN MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID" +
                " INNER JOIN infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID = infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                " WHERE InvNumber IS NOT NULL AND DecorOrders.MainOrderID IN (" + string.Join(",", MainOrderIDs) + ") ORDER BY InvNumber";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(DecorOrdersDataTable);
            }

            Collect();
        }

        private void Collect()
        {
            DataTable DistRatesDT = new DataTable();
            DataTable Items = new DataTable();

            using (DataView DV = new DataView(DecorOrdersDataTable))
            {
                Items = DV.ToTable(true, new string[] { "InvNumber" });
            }

            //get count of different covertypes
            using (DataView DV = new DataView(DecorOrdersDataTable))
            {
                DistRatesDT = DV.ToTable(true, new string[] { "PaymentRate" });
            }

            for (int i = 0; i < Items.Rows.Count; i++)
            {
                for (int j = 0; j < DistRatesDT.Rows.Count; j++)
                {
                    PaymentRate = Convert.ToDecimal(DistRatesDT.Rows[j]["PaymentRate"]);

                    string InvNumber = Items.Rows[i]["InvNumber"].ToString();
                    DataRow[] ItemsRows = DecorOrdersDataTable.Select("PaymentRate='" + PaymentRate.ToString() + "' AND InvNumber = '" + InvNumber + "'",
                        "Price ASC");

                    if (ItemsRows.Count() == 0)
                        continue;
                    int DecorConfigID = Convert.ToInt32(ItemsRows[0]["DecorConfigID"]);
                    //м.п.
                    if (GetReportMeasureTypeID(Convert.ToInt32(ItemsRows[0]["DecorConfigID"])) == 2)
                    {
                        GroupCoverTypes(ItemsRows, 2);
                    }

                    //шт.
                    if (GetReportMeasureTypeID(Convert.ToInt32(ItemsRows[0]["DecorConfigID"])) == 3)
                    {
                        DataTable ParamTableProfil = new DataTable();
                        DataTable ParamTableTPS = new DataTable();

                        DataRow[] DCs = decorCatalogOrder.DecorConfigDataTable.Select("DecorConfigID = " +
                            ItemsRows[0]["DecorConfigID"].ToString());

                        CreateParamsTable(DCs[0]["ReportParam"].ToString(), ParamTableProfil);
                        CreateParamsTable(DCs[0]["ReportParam"].ToString(), ParamTableTPS);

                        GetParametrizedData(ItemsRows, ParamTableProfil, ParamTableTPS);

                        ParamTableProfil.Dispose();
                        ParamTableTPS.Dispose();
                    }

                    //м.кв.
                    if (GetReportMeasureTypeID(Convert.ToInt32(ItemsRows[0]["DecorConfigID"])) == 1)
                    {
                        GroupCoverTypes(ItemsRows, 1);
                    }
                }
            }

            Items.Dispose();
        }

        private void Create()
        {
            DecorOrdersDataTable = new DataTable();
        }

        private void CreateReportDataTable()
        {
            ProfilReportDataTable = new DataTable();
            ProfilReportDataTable.Columns.Add(new DataColumn("UNN", Type.GetType("System.String")));
            ProfilReportDataTable.Columns.Add(new DataColumn("CurrencyCode", Type.GetType("System.String")));
            ProfilReportDataTable.Columns.Add(new DataColumn("TPSCurCode", Type.GetType("System.String")));
            ProfilReportDataTable.Columns.Add(new DataColumn("InvNumber", Type.GetType("System.String")));
            ProfilReportDataTable.Columns.Add(new DataColumn("AccountingName", Type.GetType("System.String")));
            ProfilReportDataTable.Columns.Add(new DataColumn("Count", Type.GetType("System.Decimal")));
            ProfilReportDataTable.Columns.Add(new DataColumn("Price", Type.GetType("System.String")));
            ProfilReportDataTable.Columns.Add(new DataColumn("Cost", Type.GetType("System.String")));
            ProfilReportDataTable.Columns.Add(new DataColumn("PriceWithTransport", Type.GetType("System.String")));
            ProfilReportDataTable.Columns.Add(new DataColumn("CostWithTransport", Type.GetType("System.String")));
            ProfilReportDataTable.Columns.Add(new DataColumn("Weight", Type.GetType("System.String")));
            ProfilReportDataTable.Columns.Add(new DataColumn("PaymentRate", Type.GetType("System.Decimal")));
            TPSReportDataTable = ProfilReportDataTable.Clone();
        }

        private int GetMeasureTypeID(int DecorConfigID)
        {
            DataRow[] Row = decorCatalogOrder.DecorConfigDataTable.Select("DecorConfigID = " + DecorConfigID);

            return Convert.ToInt32(Row[0]["MeasureID"]);//1 м.кв.  2 м.п. 3 шт.
        }

        private void GetParametrizedData(DataRow[] Rows, DataTable PDT, DataTable TDT)
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
                        DataRow[] InvRows = PDT.Select("InvNumber = '" + Rows[r]["InvNumber"].ToString() + "'");
                        if (InvRows.Count() == 0)
                        {
                            DataRow NewRow = PDT.NewRow();
                            NewRow["UNN"] = UNN;
                            NewRow["PaymentRate"] = PaymentRate;
                            NewRow["AccountingName"] = Rows[r]["AccountingName"].ToString();
                            NewRow["InvNumber"] = Rows[r]["InvNumber"].ToString();
                            NewRow["CurrencyCode"] = ProfilCurrencyCode;
                            NewRow["Count"] = Convert.ToDecimal(Rows[r]["Count"]);
                            NewRow["PriceWithTransport"] = Convert.ToDecimal(Rows[r]["PriceWithTransport"]);
                            NewRow["CostWithTransport"] = Convert.ToDecimal(Rows[r]["CostWithTransport"]);
                            NewRow["Cost"] = Convert.ToDecimal(Rows[r]["Cost"]);
                            NewRow["Weight"] = GetDecorWeight(Rows[r]);
                            PDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            InvRows[0]["Count"] = Convert.ToDecimal(InvRows[0]["Count"]) +
                                Convert.ToDecimal(Rows[r]["Count"]);
                            InvRows[0]["Cost"] = Convert.ToDecimal(InvRows[0]["Cost"]) +
                                Convert.ToDecimal(Rows[r]["Cost"]);
                            InvRows[0]["CostWithTransport"] = Convert.ToDecimal(InvRows[0]["CostWithTransport"]) +
                                Convert.ToDecimal(Rows[r]["CostWithTransport"]);
                            InvRows[0]["Weight"] = Convert.ToDecimal(InvRows[0]["Weight"]) +
                                GetDecorWeight(Rows[r]);
                        }
                    }
                    else
                    {
                        DataRow[] InvRows = TDT.Select("InvNumber = '" + Rows[r]["InvNumber"].ToString() + "'");
                        if (InvRows.Count() == 0)
                        {
                            DataRow NewRow = TDT.NewRow();
                            NewRow["UNN"] = UNN;
                            NewRow["PaymentRate"] = PaymentRate;
                            NewRow["AccountingName"] = Rows[r]["AccountingName"].ToString();
                            NewRow["InvNumber"] = Rows[r]["InvNumber"].ToString();
                            NewRow["CurrencyCode"] = ProfilCurrencyCode;
                            NewRow["TPSCurCode"] = TPSCurrencyCode;
                            NewRow["Count"] = Convert.ToDecimal(Rows[r]["Count"]);
                            NewRow["PriceWithTransport"] = Convert.ToDecimal(Rows[r]["PriceWithTransport"]);
                            NewRow["CostWithTransport"] = Convert.ToDecimal(Rows[r]["CostWithTransport"]);
                            NewRow["Cost"] = Convert.ToDecimal(Rows[r]["Cost"]);
                            NewRow["Weight"] = GetDecorWeight(Rows[r]);
                            TDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            InvRows[0]["Count"] = Convert.ToDecimal(InvRows[0]["Count"]) +
                                Convert.ToDecimal(Rows[r]["Count"]);
                            InvRows[0]["Cost"] = Convert.ToDecimal(InvRows[0]["Cost"]) +
                                Convert.ToDecimal(Rows[r]["Cost"]);
                            InvRows[0]["CostWithTransport"] = Convert.ToDecimal(InvRows[0]["CostWithTransport"]) +
                                Convert.ToDecimal(Rows[r]["CostWithTransport"]);
                            InvRows[0]["Weight"] = Convert.ToDecimal(InvRows[0]["Weight"]) +
                                GetDecorWeight(Rows[r]);
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
                        DataRow[] InvRows = PDT.Select("InvNumber = '" + Rows[r]["InvNumber"].ToString() + "'");
                        if (InvRows.Count() == 0)
                        {
                            DataRow NewRow = PDT.NewRow();
                            NewRow["UNN"] = UNN;
                            NewRow["PaymentRate"] = PaymentRate;
                            NewRow["AccountingName"] = Rows[r]["AccountingName"].ToString();
                            NewRow["InvNumber"] = Rows[r]["InvNumber"].ToString();
                            NewRow["CurrencyCode"] = ProfilCurrencyCode;
                            NewRow[p1] = Convert.ToDecimal(Rows[r][p1]);
                            NewRow[p2] = Convert.ToDecimal(Rows[r][p2]);
                            NewRow["Count"] = Convert.ToDecimal(Rows[r]["Count"]);
                            NewRow["PriceWithTransport"] = Convert.ToDecimal(Rows[r]["PriceWithTransport"]);
                            NewRow["CostWithTransport"] = Convert.ToDecimal(Rows[r]["CostWithTransport"]);
                            NewRow["Cost"] = Convert.ToDecimal(Rows[r]["Cost"]);
                            NewRow["Weight"] = GetDecorWeight(Rows[r]);
                            PDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            InvRows[0]["Count"] = Convert.ToDecimal(InvRows[0]["Count"]) +
                                Convert.ToDecimal(Rows[r]["Count"]);
                            InvRows[0]["CostWithTransport"] = Convert.ToDecimal(InvRows[0]["CostWithTransport"]) +
                                Convert.ToDecimal(Rows[r]["CostWithTransport"]);
                            InvRows[0]["Cost"] = Convert.ToDecimal(InvRows[0]["Cost"]) +
                                Convert.ToDecimal(Rows[r]["Cost"]);
                            InvRows[0]["Weight"] = Convert.ToDecimal(InvRows[0]["Weight"]) +
                                GetDecorWeight(Rows[r]);
                        }
                    }
                    else
                    {
                        DataRow[] InvRows = TDT.Select("InvNumber = '" + Rows[r]["InvNumber"].ToString() + "'");
                        if (InvRows.Count() == 0)
                        {
                            DataRow NewRow = TDT.NewRow();
                            NewRow["UNN"] = UNN;
                            NewRow["PaymentRate"] = PaymentRate;
                            NewRow["AccountingName"] = Rows[r]["AccountingName"].ToString();
                            NewRow["InvNumber"] = Rows[r]["InvNumber"].ToString();
                            NewRow["CurrencyCode"] = ProfilCurrencyCode;
                            NewRow["TPSCurCode"] = TPSCurrencyCode;
                            NewRow[p1] = Convert.ToDecimal(Rows[r][p1]);
                            NewRow[p2] = Convert.ToDecimal(Rows[r][p2]);
                            NewRow["Count"] = Convert.ToDecimal(Rows[r]["Count"]);
                            NewRow["PriceWithTransport"] = Convert.ToDecimal(Rows[r]["PriceWithTransport"]);
                            NewRow["CostWithTransport"] = Convert.ToDecimal(Rows[r]["CostWithTransport"]);
                            NewRow["Cost"] = Convert.ToDecimal(Rows[r]["Cost"]);
                            NewRow["Weight"] = GetDecorWeight(Rows[r]);
                            TDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            InvRows[0]["Count"] = Convert.ToDecimal(InvRows[0]["Count"]) +
                                Convert.ToDecimal(Rows[r]["Count"]);
                            InvRows[0]["CostWithTransport"] = Convert.ToDecimal(InvRows[0]["CostWithTransport"]) +
                                Convert.ToDecimal(Rows[r]["CostWithTransport"]);
                            InvRows[0]["Cost"] = Convert.ToDecimal(InvRows[0]["Cost"]) +
                                Convert.ToDecimal(Rows[r]["Cost"]);
                            InvRows[0]["Weight"] = Convert.ToDecimal(InvRows[0]["Weight"]) +
                                GetDecorWeight(Rows[r]);
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
                        DataRow[] InvRows = PDT.Select("InvNumber = '" + Rows[r]["InvNumber"].ToString() + "'");
                        if (InvRows.Count() == 0)
                        {
                            DataRow NewRow = PDT.NewRow();
                            NewRow["UNN"] = UNN;
                            NewRow["PaymentRate"] = PaymentRate;
                            NewRow["AccountingName"] = Rows[r]["AccountingName"].ToString();
                            NewRow["InvNumber"] = Rows[r]["InvNumber"].ToString();
                            NewRow["CurrencyCode"] = ProfilCurrencyCode;
                            NewRow[p3] = Convert.ToDecimal(Rows[r][p3]);
                            NewRow["Count"] = Convert.ToDecimal(Rows[r]["Count"]);
                            NewRow["PriceWithTransport"] = Convert.ToDecimal(Rows[r]["PriceWithTransport"]);
                            NewRow["CostWithTransport"] = Convert.ToDecimal(Rows[r]["CostWithTransport"]);
                            NewRow["Cost"] = Convert.ToDecimal(Rows[r]["Cost"]);
                            NewRow["Weight"] = GetDecorWeight(Rows[r]);
                            PDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            InvRows[0]["Count"] = Convert.ToDecimal(InvRows[0]["Count"]) +
                                Convert.ToDecimal(Rows[r]["Count"]);
                            InvRows[0]["CostWithTransport"] = Convert.ToDecimal(InvRows[0]["CostWithTransport"]) +
                                Convert.ToDecimal(Rows[r]["CostWithTransport"]);
                            InvRows[0]["Cost"] = Convert.ToDecimal(InvRows[0]["Cost"]) +
                                Convert.ToDecimal(Rows[r]["Cost"]);
                            InvRows[0]["Weight"] = Convert.ToDecimal(InvRows[0]["Weight"]) +
                                GetDecorWeight(Rows[r]);
                        }
                    }
                    else
                    {
                        DataRow[] InvRows = TDT.Select("InvNumber = '" + Rows[r]["InvNumber"].ToString() + "'");
                        if (InvRows.Count() == 0)
                        {
                            DataRow NewRow = TDT.NewRow();
                            NewRow["UNN"] = UNN;
                            NewRow["PaymentRate"] = PaymentRate;
                            NewRow["AccountingName"] = Rows[r]["AccountingName"].ToString();
                            NewRow["InvNumber"] = Rows[r]["InvNumber"].ToString();
                            NewRow["CurrencyCode"] = ProfilCurrencyCode;
                            NewRow["TPSCurCode"] = TPSCurrencyCode;
                            NewRow[p3] = Convert.ToDecimal(Rows[r][p3]);
                            NewRow["Count"] = Convert.ToDecimal(Rows[r]["Count"]);
                            NewRow["PriceWithTransport"] = Convert.ToDecimal(Rows[r]["PriceWithTransport"]);
                            NewRow["CostWithTransport"] = Convert.ToDecimal(Rows[r]["CostWithTransport"]);
                            NewRow["Cost"] = Convert.ToDecimal(Rows[r]["Cost"]);
                            NewRow["Weight"] = GetDecorWeight(Rows[r]);
                            TDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            InvRows[0]["Count"] = Convert.ToDecimal(InvRows[0]["Count"]) +
                                Convert.ToDecimal(Rows[r]["Count"]);
                            InvRows[0]["Cost"] = Convert.ToDecimal(InvRows[0]["Cost"]) +
                                Convert.ToDecimal(Rows[r]["Cost"]);
                            InvRows[0]["CostWithTransport"] = Convert.ToDecimal(InvRows[0]["CostWithTransport"]) +
                                Convert.ToDecimal(Rows[r]["CostWithTransport"]);
                            InvRows[0]["Weight"] = Convert.ToDecimal(InvRows[0]["Weight"]) +
                                GetDecorWeight(Rows[r]);
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
                        NewRow["Count"] = PDT.Rows[g]["Count"];
                        NewRow["PriceWithTransport"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[g]["CostWithTransport"]) / Convert.ToDecimal(PDT.Rows[g]["Count"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["CostWithTransport"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[g]["CostWithTransport"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["Cost"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[g]["Cost"]), 2, MidpointRounding.AwayFromZero);
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
                        NewRow["Count"] = PDT.Rows[g]["Count"];
                        NewRow["PriceWithTransport"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[g]["CostWithTransport"]) / Convert.ToDecimal(PDT.Rows[g]["Count"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["CostWithTransport"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[g]["CostWithTransport"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["Cost"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[g]["Cost"]), 2, MidpointRounding.AwayFromZero);
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
                        NewRow["Count"] = PDT.Rows[g]["Count"];
                        NewRow["PriceWithTransport"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[g]["CostWithTransport"]) / Convert.ToDecimal(PDT.Rows[g]["Count"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["CostWithTransport"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[g]["CostWithTransport"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["Cost"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[g]["Cost"]), 2, MidpointRounding.AwayFromZero);
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
                        NewRow["Count"] = TDT.Rows[g]["Count"];
                        NewRow["PriceWithTransport"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[g]["CostWithTransport"]) / Convert.ToDecimal(TDT.Rows[g]["Count"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["CostWithTransport"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[g]["CostWithTransport"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["Cost"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[g]["Cost"]), 2, MidpointRounding.AwayFromZero);
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
                        NewRow["Count"] = TDT.Rows[g]["Count"];
                        NewRow["PriceWithTransport"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[g]["CostWithTransport"]) / Convert.ToDecimal(TDT.Rows[g]["Count"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["CostWithTransport"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[g]["CostWithTransport"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["Cost"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[g]["Cost"]), 2, MidpointRounding.AwayFromZero);
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
                        NewRow["Count"] = TDT.Rows[g]["Count"];
                        NewRow["PriceWithTransport"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[g]["CostWithTransport"]) / Convert.ToDecimal(TDT.Rows[g]["Count"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["CostWithTransport"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[g]["CostWithTransport"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["Cost"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[g]["Cost"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["Weight"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[g]["Weight"]), 3, MidpointRounding.AwayFromZero);

                        TPSReportDataTable.Rows.Add(NewRow);
                    }
                }
            }
        }

        private int GetReportMeasureTypeID(int DecorConfigID)
        {
            DataRow[] Row = decorCatalogOrder.DecorConfigDataTable.Select("DecorConfigID = " + DecorConfigID);
            if (Row.Length > 0)
                return Convert.ToInt32(Row[0]["ReportMeasureID"]);//1 м.кв.  2 м.п. 3 шт.
            else
            {
                System.Windows.Forms.MessageBox.Show(@$"Ошибка конфигурации: не найдена DecorConfigID={DecorConfigID}. 
                        decorCatalogOrder.DecorConfigDataTable.Count={decorCatalogOrder.DecorConfigDataTable.Rows.Count}");
                return -1;
            }
        }

        private void GroupCoverTypes(DataRow[] Rows, int MeasureTypeID)
        {
            DataTable PDT = new DataTable();
            DataTable TDT = new DataTable();

            PDT = DecorOrdersDataTable.Clone();
            TDT = DecorOrdersDataTable.Clone();

            PDT.Columns.Remove("Count");
            PDT.Columns.Add(new DataColumn("Count", Type.GetType("System.Decimal")));
            PDT.Columns.Add(new DataColumn("UNN", Type.GetType("System.String")));
            PDT.Columns.Add(new DataColumn("CurrencyCode", Type.GetType("System.String")));
            PDT.Columns.Add(new DataColumn("TPSCurCode", Type.GetType("System.String")));

            TDT.Columns.Remove("Count");
            TDT.Columns.Add(new DataColumn("Count", Type.GetType("System.Decimal")));
            TDT.Columns.Add(new DataColumn("UNN", Type.GetType("System.String")));
            TDT.Columns.Add(new DataColumn("CurrencyCode", Type.GetType("System.String")));
            TDT.Columns.Add(new DataColumn("TPSCurCode", Type.GetType("System.String")));

            for (int r = 0; r < Rows.Count(); r++)
            {
                int D = Convert.ToInt32(Rows[r]["DecorConfigID"]);
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
                        DataRow[] InvRows = PDT.Select("InvNumber = '" + Rows[r]["InvNumber"].ToString() + "'");
                        if (InvRows.Count() == 0)
                        {
                            DataRow NewRow = PDT.NewRow();
                            NewRow["UNN"] = UNN;
                            NewRow["PaymentRate"] = PaymentRate;
                            NewRow["AccountingName"] = Rows[r]["AccountingName"].ToString();
                            NewRow["InvNumber"] = Rows[r]["InvNumber"].ToString();
                            NewRow["CurrencyCode"] = ProfilCurrencyCode;
                            NewRow["Count"] = Convert.ToDecimal(Rows[r]["Count"]) * L;
                            //NewRow["Price"] = Convert.ToDecimal(Rows[r]["Price"]);
                            NewRow["Cost"] = Convert.ToDecimal(Rows[r]["Cost"]);
                            NewRow["PriceWithTransport"] = Convert.ToDecimal(Rows[r]["PriceWithTransport"]);
                            NewRow["CostWithTransport"] = Convert.ToDecimal(Rows[r]["CostWithTransport"]);
                            NewRow["Weight"] = GetDecorWeight(Rows[r]);
                            PDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            InvRows[0]["Count"] = Convert.ToDecimal(InvRows[0]["Count"]) +
                                Convert.ToDecimal(Rows[r]["Count"]) * L;
                            InvRows[0]["Cost"] = Convert.ToDecimal(InvRows[0]["Cost"]) +
                                Convert.ToDecimal(Rows[r]["Cost"]);
                            InvRows[0]["CostWithTransport"] = Convert.ToDecimal(InvRows[0]["CostWithTransport"]) +
                                Convert.ToDecimal(Rows[r]["CostWithTransport"]);
                            InvRows[0]["Weight"] = Convert.ToDecimal(InvRows[0]["Weight"]) +
                                GetDecorWeight(Rows[r]);
                        }
                    }
                    else
                    {
                        DataRow[] InvRows = TDT.Select("InvNumber = '" + Rows[r]["InvNumber"].ToString() + "'");
                        if (InvRows.Count() == 0)
                        {
                            DataRow NewRow = TDT.NewRow();
                            NewRow["UNN"] = UNN;
                            NewRow["PaymentRate"] = PaymentRate;
                            NewRow["AccountingName"] = Rows[r]["AccountingName"].ToString();
                            NewRow["InvNumber"] = Rows[r]["InvNumber"].ToString();
                            NewRow["CurrencyCode"] = ProfilCurrencyCode;
                            NewRow["TPSCurCode"] = TPSCurrencyCode;
                            NewRow["Count"] = Convert.ToDecimal(Rows[r]["Count"]) * L;
                            NewRow["PriceWithTransport"] = Convert.ToDecimal(Rows[r]["PriceWithTransport"]);
                            NewRow["CostWithTransport"] = Convert.ToDecimal(Rows[r]["CostWithTransport"]);
                            NewRow["Cost"] = Convert.ToDecimal(Rows[r]["Cost"]);
                            NewRow["Weight"] = GetDecorWeight(Rows[r]);
                            TDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            InvRows[0]["Count"] = Convert.ToDecimal(InvRows[0]["Count"]) +
                                Convert.ToDecimal(Rows[r]["Count"]) * L;
                            InvRows[0]["Cost"] = Convert.ToDecimal(InvRows[0]["Cost"]) +
                                Convert.ToDecimal(Rows[r]["Cost"]);
                            InvRows[0]["CostWithTransport"] = Convert.ToDecimal(InvRows[0]["CostWithTransport"]) +
                                Convert.ToDecimal(Rows[r]["CostWithTransport"]);
                            InvRows[0]["Weight"] = Convert.ToDecimal(InvRows[0]["Weight"]) +
                                GetDecorWeight(Rows[r]);
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
                        DataRow[] InvRows = PDT.Select("InvNumber = '" + Rows[r]["InvNumber"].ToString() + "'");
                        if (InvRows.Count() == 0)
                        {
                            DataRow NewRow = PDT.NewRow();
                            NewRow["UNN"] = UNN;
                            NewRow["PaymentRate"] = PaymentRate;
                            NewRow["AccountingName"] = Rows[r]["AccountingName"].ToString();
                            NewRow["InvNumber"] = Rows[r]["InvNumber"].ToString();
                            NewRow["CurrencyCode"] = ProfilCurrencyCode;
                            if (GetMeasureTypeID(Convert.ToInt32(Rows[r]["DecorConfigID"])) == 1)
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

                            if (GetMeasureTypeID(Convert.ToInt32(Rows[r]["DecorConfigID"])) == 2)
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

                            if (GetMeasureTypeID(Convert.ToInt32(Rows[r]["DecorConfigID"])) == 3)
                            {
                                decimal H = 0;
                                if (HasParameter(Convert.ToInt32(Rows[r]["ProductID"]), "Height") && Convert.ToDecimal(Rows[r]["Height"]) != -1)
                                    H = Convert.ToDecimal(Rows[r]["Height"]);
                                if (HasParameter(Convert.ToInt32(Rows[r]["ProductID"]), "Length") && Convert.ToDecimal(Rows[r]["Length"]) != -1)
                                    H = Convert.ToDecimal(Rows[r]["Length"]);
                                decimal d = H * Convert.ToDecimal(Rows[r]["Width"]) / 1000000;
                                d = Decimal.Round(d, 3, MidpointRounding.AwayFromZero);
                                decimal Square = d * Convert.ToDecimal(Rows[r]["Count"]);
                                NewRow["Count"] = Square;
                            }
                            NewRow["PriceWithTransport"] = Convert.ToDecimal(Rows[r]["PriceWithTransport"]);
                            NewRow["CostWithTransport"] = Convert.ToDecimal(Rows[r]["CostWithTransport"]);
                            NewRow["Cost"] = Convert.ToDecimal(Rows[r]["Cost"]);
                            NewRow["Weight"] = GetDecorWeight(Rows[r]);
                            PDT.Rows.Add(NewRow);
                            continue;
                        }
                        else
                        {
                            if (GetMeasureTypeID(Convert.ToInt32(Rows[r]["DecorConfigID"])) == 1)
                            {
                                if (Convert.ToDecimal(Rows[r]["Height"]) != -1)
                                {
                                    decimal d = Convert.ToDecimal(Rows[r]["Height"]) * Convert.ToDecimal(Rows[r]["Width"]) / 1000000;
                                    d = Decimal.Round(d, 3, MidpointRounding.AwayFromZero);
                                    decimal Square = d * Convert.ToDecimal(Rows[r]["Count"]);
                                    InvRows[0]["Count"] = Convert.ToDecimal(InvRows[0]["Count"]) + Square;
                                }
                                if (Convert.ToDecimal(Rows[r]["Length"]) != -1)
                                {
                                    decimal d = Convert.ToDecimal(Rows[r]["Length"]) * Convert.ToDecimal(Rows[r]["Width"]) / 1000000;
                                    d = Decimal.Round(d, 3, MidpointRounding.AwayFromZero);
                                    decimal Square = d * Convert.ToDecimal(Rows[r]["Count"]);
                                    InvRows[0]["Count"] = Convert.ToDecimal(InvRows[0]["Count"]) + Square;
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
                                InvRows[0]["Count"] = Convert.ToDecimal(InvRows[0]["Count"]) + Square;
                            }
                            if (GetMeasureTypeID(Convert.ToInt32(Rows[r]["DecorConfigID"])) == 3)
                            {
                                decimal H = 0;
                                if (HasParameter(Convert.ToInt32(Rows[r]["ProductID"]), "Height") && Convert.ToDecimal(Rows[r]["Height"]) != -1)
                                    H = Convert.ToDecimal(Rows[r]["Height"]);
                                if (HasParameter(Convert.ToInt32(Rows[r]["ProductID"]), "Length") && Convert.ToDecimal(Rows[r]["Length"]) != -1)
                                    H = Convert.ToDecimal(Rows[r]["Length"]);
                                decimal d = H * Convert.ToDecimal(Rows[r]["Width"]) / 1000000;
                                d = Decimal.Round(d, 3, MidpointRounding.AwayFromZero);
                                decimal Square = d * Convert.ToDecimal(Rows[r]["Count"]);
                                InvRows[0]["Count"] = Convert.ToDecimal(InvRows[0]["Count"]) + Square;
                            }
                            InvRows[0]["Cost"] = Convert.ToDecimal(InvRows[0]["Cost"]) +
                                Convert.ToDecimal(Rows[r]["Cost"]);
                            InvRows[0]["CostWithTransport"] = Convert.ToDecimal(InvRows[0]["CostWithTransport"]) +
                                Convert.ToDecimal(Rows[r]["CostWithTransport"]);
                            InvRows[0]["Weight"] = Convert.ToDecimal(InvRows[0]["Weight"]) +
                                GetDecorWeight(Rows[r]);
                        }
                    }
                    else
                    {
                        DataRow[] InvRows = TDT.Select("InvNumber = '" + Rows[r]["InvNumber"].ToString() + "'");
                        InvNumber = Rows[r]["InvNumber"].ToString();
                        if (InvRows.Count() == 0)
                        {
                            DataRow NewRow = TDT.NewRow();
                            NewRow["UNN"] = UNN;
                            NewRow["PaymentRate"] = PaymentRate;
                            NewRow["AccountingName"] = Rows[r]["AccountingName"].ToString();
                            NewRow["InvNumber"] = Rows[r]["InvNumber"].ToString();
                            NewRow["CurrencyCode"] = ProfilCurrencyCode;
                            NewRow["TPSCurCode"] = TPSCurrencyCode;
                            if (GetMeasureTypeID(Convert.ToInt32(Rows[r]["DecorConfigID"])) == 1)
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

                            if (GetMeasureTypeID(Convert.ToInt32(Rows[r]["DecorConfigID"])) == 3)
                            {
                                decimal H = 0;
                                if (HasParameter(Convert.ToInt32(Rows[r]["ProductID"]), "Height") && Convert.ToDecimal(Rows[r]["Height"]) != -1)
                                    H = Convert.ToDecimal(Rows[r]["Height"]);
                                if (HasParameter(Convert.ToInt32(Rows[r]["ProductID"]), "Length") && Convert.ToDecimal(Rows[r]["Length"]) != -1)
                                    H = Convert.ToDecimal(Rows[r]["Length"]);
                                decimal d = H * Convert.ToDecimal(Rows[r]["Width"]) / 1000000;
                                d = Decimal.Round(d, 3, MidpointRounding.AwayFromZero);
                                decimal Square = d * Convert.ToDecimal(Rows[r]["Count"]);
                                NewRow["Count"] = Square;
                            }
                            NewRow["PriceWithTransport"] = Convert.ToDecimal(Rows[r]["PriceWithTransport"]);
                            NewRow["CostWithTransport"] = Convert.ToDecimal(Rows[r]["CostWithTransport"]);
                            NewRow["Cost"] = Convert.ToDecimal(Rows[r]["Cost"]);
                            NewRow["Weight"] = GetDecorWeight(Rows[r]);
                            TDT.Rows.Add(NewRow);
                            continue;
                        }
                        else
                        {
                            if (GetMeasureTypeID(Convert.ToInt32(Rows[r]["DecorConfigID"])) == 1)
                            {
                                if (Convert.ToDecimal(Rows[r]["Height"]) != -1)
                                {
                                    decimal d = Convert.ToDecimal(Rows[r]["Height"]) * Convert.ToDecimal(Rows[r]["Width"]) / 1000000;
                                    d = Decimal.Round(d, 3, MidpointRounding.AwayFromZero);
                                    decimal Square = d * Convert.ToDecimal(Rows[r]["Count"]);
                                    InvRows[0]["Count"] = Convert.ToDecimal(InvRows[0]["Count"]) + Square;
                                }
                                if (Convert.ToDecimal(Rows[r]["Length"]) != -1)
                                {
                                    decimal d = Convert.ToDecimal(Rows[r]["Length"]) * Convert.ToDecimal(Rows[r]["Width"]) / 1000000;
                                    d = Decimal.Round(d, 3, MidpointRounding.AwayFromZero);
                                    decimal Square = d * Convert.ToDecimal(Rows[r]["Count"]);
                                    InvRows[0]["Count"] = Convert.ToDecimal(InvRows[0]["Count"]) + Square;
                                }
                            }
                            if (GetMeasureTypeID(Convert.ToInt32(Rows[r]["DecorConfigID"])) == 2)
                            {
                                decimal L = 0;
                                L = Convert.ToDecimal(Rows[r]["Length"]);
                                if (L == -1)
                                    L = Convert.ToDecimal(Rows[r]["Height"]);
                                decimal d = L * Convert.ToDecimal(Rows[r]["Width"]) / 1000000;
                                d = Decimal.Round(d, 3, MidpointRounding.AwayFromZero);
                                decimal Square = d * Convert.ToDecimal(Rows[r]["Count"]);
                                InvRows[0]["Count"] = Convert.ToDecimal(InvRows[0]["Count"]) + Square;
                            }
                            if (GetMeasureTypeID(Convert.ToInt32(Rows[r]["DecorConfigID"])) == 3)
                            {
                                decimal H = 0;
                                if (HasParameter(Convert.ToInt32(Rows[r]["ProductID"]), "Height") && Convert.ToDecimal(Rows[r]["Height"]) != -1)
                                    H = Convert.ToDecimal(Rows[r]["Height"]);
                                if (HasParameter(Convert.ToInt32(Rows[r]["ProductID"]), "Length") && Convert.ToDecimal(Rows[r]["Length"]) != -1)
                                    H = Convert.ToDecimal(Rows[r]["Length"]);
                                decimal d = H * Convert.ToDecimal(Rows[r]["Width"]) / 1000000;
                                d = Decimal.Round(d, 3, MidpointRounding.AwayFromZero);
                                decimal Square = d * Convert.ToDecimal(Rows[r]["Count"]);
                                InvRows[0]["Count"] = Convert.ToDecimal(InvRows[0]["Count"]) + Square;
                            }
                            InvRows[0]["Cost"] = Convert.ToDecimal(InvRows[0]["Cost"]) +
                                Convert.ToDecimal(Rows[r]["Cost"]);
                            InvRows[0]["CostWithTransport"] = Convert.ToDecimal(InvRows[0]["CostWithTransport"]) +
                                Convert.ToDecimal(Rows[r]["CostWithTransport"]);
                            InvRows[0]["Weight"] = Convert.ToDecimal(InvRows[0]["Weight"]) +
                                GetDecorWeight(Rows[r]);
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
                        NewRow["Count"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[i]["Count"]) / 1000, 3, MidpointRounding.AwayFromZero);
                        NewRow["PriceWithTransport"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[i]["CostWithTransport"]) / (Convert.ToDecimal(PDT.Rows[i]["Count"]) / 1000), 2, MidpointRounding.AwayFromZero);
                        NewRow["CostWithTransport"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[i]["CostWithTransport"]), 2, MidpointRounding.AwayFromZero);
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
                        NewRow["Count"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[i]["Count"]) / 1000, 3, MidpointRounding.AwayFromZero);
                        NewRow["PriceWithTransport"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[i]["CostWithTransport"]) / (Convert.ToDecimal(TDT.Rows[i]["Count"]) / 1000), 2, MidpointRounding.AwayFromZero);
                        NewRow["CostWithTransport"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[i]["CostWithTransport"]), 2, MidpointRounding.AwayFromZero);
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
                        NewRow["Count"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[i]["Count"]), 3, MidpointRounding.AwayFromZero);
                        NewRow["PriceWithTransport"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[i]["CostWithTransport"]) / Convert.ToDecimal(PDT.Rows[i]["Count"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["CostWithTransport"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[i]["CostWithTransport"]), 2, MidpointRounding.AwayFromZero);
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
                        NewRow["Count"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[i]["Count"]), 3, MidpointRounding.AwayFromZero);
                        NewRow["PriceWithTransport"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[i]["CostWithTransport"]) / Convert.ToDecimal(TDT.Rows[i]["Count"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["CostWithTransport"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[i]["CostWithTransport"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["Cost"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[i]["Cost"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["Weight"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[i]["Weight"]), 3, MidpointRounding.AwayFromZero);

                        TPSReportDataTable.Rows.Add(NewRow);
                    }
                }
            }

            PDT.Dispose();
            TDT.Dispose();
        }

        private bool IsProfil(int DecorConfigID)
        {
            DataRow[] Rows = decorCatalogOrder.DecorConfigDataTable.Select("DecorConfigID = " + DecorConfigID.ToString());

            if (Rows[0]["FactoryID"].ToString() == "1")
                return true;

            return false;
        }
    }
}