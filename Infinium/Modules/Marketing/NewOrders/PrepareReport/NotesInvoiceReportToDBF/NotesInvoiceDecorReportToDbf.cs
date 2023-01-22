
using System;
using System.Data;
using System.Data.SqlClient;
using System.Linq;


namespace Infinium.Modules.Marketing.NewOrders.PrepareReport.NotesInvoiceReportToDbf
{
    public class NotesInvoiceDecorReportToDbf
    {
        private decimal PaymentRate = 1;
        private int ClientID = 0;
        private string ProfilCurrencyCode = "0";
        private string TPSCurrencyCode = "0";
        private string UNN = string.Empty;

        private DataTable CurrencyTypesDT;
        private DataTable _profilReportDataTable = null;
        private DataTable _tPSReportDataTable = null;
        private DataTable FrameColorsDataTable = null;
        private DataTable PatinaDataTable = null;
        private DataTable PatinaRALDataTable = null;
        private DataTable DecorProductsDataTable = null;
        private DataTable DecorConfigDataTable = null;
        private DataTable MeasuresDataTable = null;
        private DataTable DecorOrdersDataTable = null;
        private DataTable DecorDataTable = null;
        private DataTable DecorParametersDataTable = null;

        public DataTable ProfilReportDataTable => _profilReportDataTable;
        public DataTable TPSReportDataTable => _tPSReportDataTable;

        public NotesInvoiceDecorReportToDbf()
        {
            Create();
            CreateReportDataTable();
        }

        private void GetColorsDT()
        {
            FrameColorsDataTable = new DataTable();
            FrameColorsDataTable.Columns.Add(new DataColumn("ColorID", Type.GetType("System.Int64")));
            FrameColorsDataTable.Columns.Add(new DataColumn("ColorName", Type.GetType("System.String")));
            FrameColorsDataTable.Columns.Add(new DataColumn("Cvet", Type.GetType("System.String")));
            string SelectCommand = @"SELECT TechStoreID, TechStoreName, Cvet FROM TechStore
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
                        NewRow["Cvet"] = "000";
                        FrameColorsDataTable.Rows.Add(NewRow);
                    }
                    {
                        DataRow NewRow = FrameColorsDataTable.NewRow();
                        NewRow["ColorID"] = 0;
                        NewRow["ColorName"] = "на выбор";
                        NewRow["Cvet"] = "0000000";
                        FrameColorsDataTable.Rows.Add(NewRow);
                    }
                    for (int i = 0; i < DT.Rows.Count; i++)
                    {
                        DataRow NewRow = FrameColorsDataTable.NewRow();
                        NewRow["ColorID"] = Convert.ToInt64(DT.Rows[i]["TechStoreID"]);
                        NewRow["ColorName"] = DT.Rows[i]["TechStoreName"].ToString();
                        NewRow["Cvet"] = DT.Rows[i]["Cvet"].ToString();
                        FrameColorsDataTable.Rows.Add(NewRow);
                    }
                }
            }
        }

        private void GetPatinaDT()
        {
            PatinaDataTable = new DataTable();
            PatinaRALDataTable = new DataTable();
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
        }

        private void Create()
        {
            string SelectCommand = "SELECT * FROM CurrencyTypes";
            CurrencyTypesDT = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(CurrencyTypesDT);
            }
            DecorOrdersDataTable = new DataTable();

            SelectCommand = @"SELECT ProductID, ProductName, MeasureID, ReportParam FROM DecorProducts" +
                " WHERE (ProductID IN (SELECT ProductID FROM DecorConfig WHERE (Enabled = 1))) ORDER BY ProductName ASC";
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

            GetColorsDT();
            GetPatinaDT();

            DecorConfigDataTable = new DataTable();
            //using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM DecorConfig",
            //    ConnectionStrings.CatalogConnectionString))
            //{
            //    DA.Fill(DecorConfigDataTable);
            //}
            DecorConfigDataTable = TablesManager.DecorConfigDataTable;

            DecorParametersDataTable = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM DecorParameters", ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(DecorParametersDataTable);
            }

        }
        public bool HasParameter(int ProductID, String Parameter)
        {
            DataRow[] Rows = DecorParametersDataTable.Select("ProductID = " + ProductID);

            return Convert.ToBoolean(Rows[0][Parameter]);
        }

        private void CreateReportDataTable()
        {
            _profilReportDataTable = new DataTable();
            _profilReportDataTable.Columns.Add(new DataColumn("UNN", Type.GetType("System.String")));
            _profilReportDataTable.Columns.Add(new DataColumn("CurrencyCode", Type.GetType("System.String")));
            _profilReportDataTable.Columns.Add(new DataColumn("TPSCurCode", Type.GetType("System.String")));
            _profilReportDataTable.Columns.Add(new DataColumn("InvNumber", Type.GetType("System.String")));
            _profilReportDataTable.Columns.Add(new DataColumn("Notes", Type.GetType("System.String")));
            _profilReportDataTable.Columns.Add(new DataColumn("Cvet", Type.GetType("System.String")));
            _profilReportDataTable.Columns.Add(new DataColumn("Patina", Type.GetType("System.String")));
            _profilReportDataTable.Columns.Add(new DataColumn("AccountingName", Type.GetType("System.String")));
            _profilReportDataTable.Columns.Add(new DataColumn("Count", Type.GetType("System.Decimal")));
            _profilReportDataTable.Columns.Add(new DataColumn("Price", Type.GetType("System.String")));
            _profilReportDataTable.Columns.Add(new DataColumn("Cost", Type.GetType("System.String")));
            _profilReportDataTable.Columns.Add(new DataColumn("PriceWithTransport", Type.GetType("System.String")));
            _profilReportDataTable.Columns.Add(new DataColumn("CostWithTransport", Type.GetType("System.String")));
            _profilReportDataTable.Columns.Add(new DataColumn("Weight", Type.GetType("System.String")));
            _profilReportDataTable.Columns.Add(new DataColumn("PaymentRate", Type.GetType("System.Decimal")));
            _tPSReportDataTable = _profilReportDataTable.Clone();
        }

        public void ClearReport()
        {
            DecorOrdersDataTable.Clear();

            _profilReportDataTable.Clear();
            _tPSReportDataTable.Clear();

            _profilReportDataTable.AcceptChanges();
            _tPSReportDataTable.AcceptChanges();
        }

        private bool IsProfil(int DecorConfigID)
        {
            DataRow[] Rows = DecorConfigDataTable.Select("DecorConfigID = " + DecorConfigID.ToString());

            if (Rows[0]["FactoryID"].ToString() == "1")
                return true;

            return false;
        }

        public void GetMegaOrderInfo1(int MegaOrderID)
        {
            string SelectCommand = "SELECT MegaOrderID, TransportCost, AdditionalCost, PaymentRate, ClientID, CurrencyTypeID FROM NewMegaOrders" +
                " WHERE MegaOrderID = " + MegaOrderID;
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    if (DA.Fill(DT) > 0)
                    {
                        int CurrencyTypeID = 1;
                        if (DT.Rows[0]["TransportCost"] != DBNull.Value)
                            decimal.TryParse(DT.Rows[0]["TransportCost"].ToString(), out _);
                        if (DT.Rows[0]["AdditionalCost"] != DBNull.Value)
                            decimal.TryParse(DT.Rows[0]["AdditionalCost"].ToString(), out _);
                        //if (DT.Rows[0]["PaymentRate"] != DBNull.Value)
                        //    decimal.TryParse(DT.Rows[0]["PaymentRate"].ToString(), out PaymentRate);
                        if (DT.Rows[0]["ClientID"] != DBNull.Value)
                            int.TryParse(DT.Rows[0]["ClientID"].ToString(), out ClientID);
                        if (DT.Rows[0]["CurrencyTypeID"] != DBNull.Value)
                            int.TryParse(DT.Rows[0]["CurrencyTypeID"].ToString(), out CurrencyTypeID);

                        DataRow[] rows = CurrencyTypesDT.Select("CurrencyTypeID = " + CurrencyTypeID);
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

        public void GetMegaOrderInfo(int MainOrderID)
        {
            string SelectCommand = "SELECT MegaOrderID, TransportCost, AdditionalCost, PaymentRate, ClientID, CurrencyTypeID FROM NewMegaOrders" +
                " WHERE MegaOrderID IN (SELECT MegaOrderID FROM NewMainOrders WHERE MainOrderID = " + MainOrderID + ")";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    if (DA.Fill(DT) > 0)
                    {
                        int CurrencyTypeID = 1;
                        if (DT.Rows[0]["TransportCost"] != DBNull.Value)
                            decimal.TryParse(DT.Rows[0]["TransportCost"].ToString(), out _);
                        if (DT.Rows[0]["AdditionalCost"] != DBNull.Value)
                            decimal.TryParse(DT.Rows[0]["AdditionalCost"].ToString(), out _);
                        //if (DT.Rows[0]["PaymentRate"] != DBNull.Value)
                        //    decimal.TryParse(DT.Rows[0]["PaymentRate"].ToString(), out PaymentRate);
                        if (DT.Rows[0]["ClientID"] != DBNull.Value)
                            int.TryParse(DT.Rows[0]["ClientID"].ToString(), out ClientID);
                        if (DT.Rows[0]["CurrencyTypeID"] != DBNull.Value)
                            int.TryParse(DT.Rows[0]["CurrencyTypeID"].ToString(), out CurrencyTypeID);

                        DataRow[] rows = CurrencyTypesDT.Select("CurrencyTypeID = " + CurrencyTypeID);
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

        private int GetReportMeasureTypeID(int DecorConfigID)
        {
            DataRow[] Row = DecorConfigDataTable.Select("DecorConfigID = " + DecorConfigID);

            return Convert.ToInt32(Row[0]["ReportMeasureID"]);//1 м.кв.  2 м.п. 3 шт.
        }

        private int GetMeasureTypeID(int DecorConfigID)
        {
            DataRow[] Row = DecorConfigDataTable.Select("DecorConfigID = " + DecorConfigID);

            return Convert.ToInt32(Row[0]["MeasureID"]);//1 м.кв.  2 м.п. 3 шт.
        }

        private string GetColorCode(int ColorID)
        {
            string code = string.Empty;
            try
            {
                DataRow[] Rows = FrameColorsDataTable.Select("ColorID = " + ColorID);
                code = Rows[0]["Cvet"].ToString();
            }
            catch
            {
                return string.Empty;
            }
            return code;
        }

        private string GetPatinaCode(int PatinaID)
        {
            string code = string.Empty;
            try
            {
                DataRow[] Rows = PatinaDataTable.Select("PatinaID = " + PatinaID);
                code = Rows[0]["Patina"].ToString();
            }
            catch
            {
                return string.Empty;
            }
            return code;
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
            DT.Columns.Add(new DataColumn("Notes", Type.GetType("System.String")));
            DT.Columns.Add(new DataColumn("Cvet", Type.GetType("System.String")));
            DT.Columns.Add(new DataColumn("Patina", Type.GetType("System.String")));
            DT.Columns.Add(new DataColumn("Count", Type.GetType("System.Decimal")));
            DT.Columns.Add(new DataColumn("TotalCount", Type.GetType("System.Int32")));
            DT.Columns.Add(new DataColumn("Cost", Type.GetType("System.Decimal")));
            DT.Columns.Add(new DataColumn("PriceWithTransport", Type.GetType("System.String")));
            DT.Columns.Add(new DataColumn("CostWithTransport", Type.GetType("System.String")));
            DT.Columns.Add(new DataColumn("Weight", Type.GetType("System.Decimal")));
            DT.Columns.Add(new DataColumn("PaymentRate", Type.GetType("System.Decimal")));
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
            PDT.Columns.Add(new DataColumn("Cvet", Type.GetType("System.String")));
            PDT.Columns.Add(new DataColumn("Patina", Type.GetType("System.String")));

            TDT.Columns.Remove("Count");
            TDT.Columns.Add(new DataColumn("Count", Type.GetType("System.Decimal")));
            TDT.Columns.Add(new DataColumn("UNN", Type.GetType("System.String")));
            TDT.Columns.Add(new DataColumn("CurrencyCode", Type.GetType("System.String")));
            TDT.Columns.Add(new DataColumn("TPSCurCode", Type.GetType("System.String")));
            TDT.Columns.Add(new DataColumn("Cvet", Type.GetType("System.String")));
            TDT.Columns.Add(new DataColumn("Patina", Type.GetType("System.String")));


            for (int r = 0; r < Rows.Count(); r++)
            {
                int ColorID = Convert.ToInt32(Rows[r]["ColorID"]);
                int PatinaID = Convert.ToInt32(Rows[r]["PatinaID"]);
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
                        DataRow[] InvRows = PDT.Select("InvNumber = '" + Rows[r]["InvNumber"].ToString() +
                                                       "' AND Notes = '" + Rows[r]["Notes"].ToString() +
                            "' AND ColorID = " + Rows[r]["ColorID"].ToString() +
                            " AND PatinaID = " + Rows[r]["PatinaID"].ToString());
                        if (InvRows.Count() == 0)
                        {
                            DataRow NewRow = PDT.NewRow();
                            NewRow["UNN"] = UNN;
                            NewRow["PaymentRate"] = PaymentRate;
                            NewRow["AccountingName"] = Rows[r]["AccountingName"].ToString();
                            NewRow["InvNumber"] = Rows[r]["InvNumber"].ToString();
                            NewRow["Notes"] = Rows[r]["Notes"].ToString();
                            NewRow["Cvet"] = GetColorCode(ColorID);
                            NewRow["Patina"] = GetPatinaCode(PatinaID);
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
                        DataRow[] InvRows = TDT.Select("InvNumber = '" + Rows[r]["InvNumber"].ToString() +
                                                       " AND Notes = '" + Rows[r]["Notes"].ToString() +
                            "' AND ColorID = " + Rows[r]["ColorID"].ToString() +
                            " AND PatinaID = " + Rows[r]["PatinaID"].ToString());
                        if (InvRows.Count() == 0)
                        {
                            DataRow NewRow = TDT.NewRow();
                            NewRow["UNN"] = UNN;
                            NewRow["PaymentRate"] = PaymentRate;
                            NewRow["AccountingName"] = Rows[r]["AccountingName"].ToString();
                            NewRow["InvNumber"] = Rows[r]["InvNumber"].ToString();
                            NewRow["Notes"] = Rows[r]["Notes"].ToString();
                            NewRow["Cvet"] = GetColorCode(ColorID);
                            NewRow["Patina"] = GetPatinaCode(PatinaID);
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
                        DataRow[] InvRows = PDT.Select("InvNumber = '" + Rows[r]["InvNumber"].ToString() +
                                                       "' AND Notes = '" + Rows[r]["Notes"].ToString() +
                            "' AND ColorID = " + Rows[r]["ColorID"].ToString() +
                            " AND PatinaID = " + Rows[r]["PatinaID"].ToString());
                        if (InvRows.Count() == 0)
                        {
                            DataRow NewRow = PDT.NewRow();
                            NewRow["UNN"] = UNN;
                            NewRow["PaymentRate"] = PaymentRate;
                            NewRow["AccountingName"] = Rows[r]["AccountingName"].ToString();
                            NewRow["InvNumber"] = Rows[r]["InvNumber"].ToString();
                            NewRow["Notes"] = Rows[r]["Notes"].ToString();
                            NewRow["Cvet"] = GetColorCode(ColorID);
                            NewRow["Patina"] = GetPatinaCode(PatinaID);
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
                        DataRow[] InvRows = TDT.Select("InvNumber = '" + Rows[r]["InvNumber"].ToString() +
                                                       "' AND Notes = '" + Rows[r]["Notes"].ToString() +
                            "' AND ColorID = " + Rows[r]["ColorID"].ToString() +
                            " AND PatinaID = " + Rows[r]["PatinaID"].ToString());
                        if (InvRows.Count() == 0)
                        {
                            DataRow NewRow = TDT.NewRow();
                            NewRow["UNN"] = UNN;
                            NewRow["PaymentRate"] = PaymentRate;
                            NewRow["AccountingName"] = Rows[r]["AccountingName"].ToString();
                            NewRow["InvNumber"] = Rows[r]["InvNumber"].ToString();
                            NewRow["Notes"] = Rows[r]["Notes"].ToString();
                            NewRow["Cvet"] = GetColorCode(ColorID);
                            NewRow["Patina"] = GetPatinaCode(PatinaID);
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
                        DataRow NewRow = _profilReportDataTable.NewRow();

                        NewRow["UNN"] = UNN;
                        NewRow["PaymentRate"] = PaymentRate;
                        NewRow["AccountingName"] = PDT.Rows[i]["AccountingName"].ToString();
                        NewRow["InvNumber"] = PDT.Rows[i]["InvNumber"].ToString();
                        NewRow["Notes"] = PDT.Rows[i]["Notes"].ToString();
                        NewRow["Cvet"] = PDT.Rows[i]["Cvet"].ToString();
                        NewRow["Patina"] = PDT.Rows[i]["Patina"].ToString();
                        NewRow["CurrencyCode"] = ProfilCurrencyCode;
                        NewRow["Count"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[i]["Count"]) / 1000, 3, MidpointRounding.AwayFromZero);
                        NewRow["PriceWithTransport"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[i]["CostWithTransport"]) / (Convert.ToDecimal(PDT.Rows[i]["Count"]) / 1000), 2, MidpointRounding.AwayFromZero);
                        NewRow["CostWithTransport"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[i]["CostWithTransport"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["Cost"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[i]["Cost"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["Weight"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[i]["Weight"]), 3, MidpointRounding.AwayFromZero);
                        _profilReportDataTable.Rows.Add(NewRow);
                    }
                }

                if (TDT.Rows.Count > 0)
                {
                    for (int i = 0; i < TDT.Rows.Count; i++)
                    {
                        DataRow NewRow = _tPSReportDataTable.NewRow();

                        NewRow["UNN"] = UNN;
                        NewRow["PaymentRate"] = PaymentRate;
                        NewRow["AccountingName"] = TDT.Rows[i]["AccountingName"].ToString();
                        NewRow["InvNumber"] = TDT.Rows[i]["InvNumber"].ToString();
                        NewRow["Notes"] = TDT.Rows[i]["Notes"].ToString();
                        NewRow["Cvet"] = TDT.Rows[i]["Cvet"].ToString();
                        NewRow["Patina"] = TDT.Rows[i]["Patina"].ToString();
                        NewRow["CurrencyCode"] = ProfilCurrencyCode;
                        NewRow["TPSCurCode"] = TPSCurrencyCode;
                        NewRow["Count"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[i]["Count"]) / 1000, 3, MidpointRounding.AwayFromZero);
                        NewRow["PriceWithTransport"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[i]["CostWithTransport"]) / (Convert.ToDecimal(TDT.Rows[i]["Count"]) / 1000), 2, MidpointRounding.AwayFromZero);
                        NewRow["CostWithTransport"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[i]["CostWithTransport"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["Cost"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[i]["Cost"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["Weight"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[i]["Weight"]), 3, MidpointRounding.AwayFromZero);

                        _tPSReportDataTable.Rows.Add(NewRow);
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
                        DataRow NewRow = _profilReportDataTable.NewRow();

                        NewRow["UNN"] = UNN;
                        NewRow["PaymentRate"] = PaymentRate;
                        NewRow["AccountingName"] = PDT.Rows[i]["AccountingName"].ToString();
                        NewRow["InvNumber"] = PDT.Rows[i]["InvNumber"].ToString();
                        NewRow["Notes"] = PDT.Rows[i]["Notes"].ToString();
                        NewRow["Cvet"] = PDT.Rows[i]["Cvet"].ToString();
                        NewRow["Patina"] = PDT.Rows[i]["Patina"].ToString();
                        NewRow["CurrencyCode"] = ProfilCurrencyCode;
                        NewRow["Count"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[i]["Count"]), 3, MidpointRounding.AwayFromZero);
                        NewRow["PriceWithTransport"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[i]["CostWithTransport"]) / Convert.ToDecimal(PDT.Rows[i]["Count"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["CostWithTransport"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[i]["CostWithTransport"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["Cost"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[i]["Cost"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["Weight"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[i]["Weight"]), 3, MidpointRounding.AwayFromZero);

                        _profilReportDataTable.Rows.Add(NewRow);
                    }
                }

                if (TDT.Rows.Count > 0)
                {
                    for (int i = 0; i < TDT.Rows.Count; i++)
                    {
                        DataRow NewRow = _tPSReportDataTable.NewRow();

                        NewRow["UNN"] = UNN;
                        NewRow["PaymentRate"] = PaymentRate;
                        NewRow["AccountingName"] = TDT.Rows[i]["AccountingName"].ToString();
                        NewRow["InvNumber"] = TDT.Rows[i]["InvNumber"].ToString();
                        NewRow["Notes"] = TDT.Rows[i]["Notes"].ToString();
                        NewRow["Cvet"] = TDT.Rows[i]["Cvet"].ToString();
                        NewRow["Patina"] = TDT.Rows[i]["Patina"].ToString();
                        NewRow["CurrencyCode"] = ProfilCurrencyCode;
                        NewRow["TPSCurCode"] = TPSCurrencyCode;
                        NewRow["Count"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[i]["Count"]), 3, MidpointRounding.AwayFromZero);
                        NewRow["PriceWithTransport"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[i]["CostWithTransport"]) / Convert.ToDecimal(TDT.Rows[i]["Count"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["CostWithTransport"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[i]["CostWithTransport"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["Cost"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[i]["Cost"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["Weight"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[i]["Weight"]), 3, MidpointRounding.AwayFromZero);

                        _tPSReportDataTable.Rows.Add(NewRow);
                    }
                }
            }

            PDT.Dispose();
            TDT.Dispose();
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
                    string Cvet = GetColorCode(Convert.ToInt32(Rows[r]["ColorID"]));
                    string Patina = GetPatinaCode(Convert.ToInt32(Rows[r]["PatinaID"]));
                    if (IsProfil(Convert.ToInt32(Rows[r]["DecorConfigID"])))
                    {
                        DataRow[] InvRows = PDT.Select($"InvNumber = '{ Rows[r]["InvNumber"].ToString() }' AND Notes = '{ Rows[r]["Notes"].ToString() }' AND Cvet = '{ Cvet }' AND Patina = '{ Patina }'");
                        if (InvRows.Count() == 0)
                        {
                            DataRow NewRow = PDT.NewRow();
                            NewRow["UNN"] = UNN;
                            NewRow["PaymentRate"] = PaymentRate;
                            NewRow["AccountingName"] = Rows[r]["AccountingName"].ToString();
                            NewRow["InvNumber"] = Rows[r]["InvNumber"].ToString();
                            NewRow["Notes"] = Rows[r]["Notes"].ToString();
                            NewRow["Cvet"] = Cvet;
                            NewRow["Patina"] = Patina;
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
                        DataRow[] InvRows = TDT.Select($"InvNumber = '{ Rows[r]["InvNumber"].ToString() }' AND Notes = '{ Rows[r]["Notes"].ToString() }' AND Cvet = '{ Cvet }' AND Patina = '{ Patina }'");
                        if (InvRows.Count() == 0)
                        {
                            DataRow NewRow = TDT.NewRow();
                            NewRow["UNN"] = UNN;
                            NewRow["PaymentRate"] = PaymentRate;
                            NewRow["AccountingName"] = Rows[r]["AccountingName"].ToString();
                            NewRow["InvNumber"] = Rows[r]["InvNumber"].ToString();
                            NewRow["Notes"] = Rows[r]["Notes"].ToString();
                            NewRow["Cvet"] = Cvet;
                            NewRow["Patina"] = Patina;
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
                    string Cvet = GetColorCode(Convert.ToInt32(Rows[r]["ColorID"]));
                    string Patina = GetPatinaCode(Convert.ToInt32(Rows[r]["PatinaID"]));
                    if (IsProfil(Convert.ToInt32(Rows[r]["DecorConfigID"])))
                    {
                        DataRow[] InvRows = PDT.Select($"InvNumber = '{ Rows[r]["InvNumber"].ToString() }' AND" +
                                                       $" Notes = '{ Rows[r]["Notes"].ToString() }' AND" +
                                                       $" Cvet = '{ Cvet }' AND Patina = '{ Patina }'");
                        if (InvRows.Count() == 0)
                        {
                            DataRow NewRow = PDT.NewRow();
                            NewRow["UNN"] = UNN;
                            NewRow["PaymentRate"] = PaymentRate;
                            NewRow["AccountingName"] = Rows[r]["AccountingName"].ToString();
                            NewRow["InvNumber"] = Rows[r]["InvNumber"].ToString();
                            NewRow["Notes"] = Rows[r]["Notes"].ToString();
                            NewRow["Cvet"] = Cvet;
                            NewRow["Patina"] = Patina;
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
                        DataRow[] InvRows = TDT.Select($"InvNumber = '{ Rows[r]["InvNumber"].ToString() }' AND " +
                                                       $" Notes = '{ Rows[r]["Notes"].ToString() }' AND Cvet = '{ Cvet }' AND Patina = '{ Patina }'");
                        if (InvRows.Count() == 0)
                        {
                            DataRow NewRow = TDT.NewRow();
                            NewRow["UNN"] = UNN;
                            NewRow["PaymentRate"] = PaymentRate;
                            NewRow["AccountingName"] = Rows[r]["AccountingName"].ToString();
                            NewRow["InvNumber"] = Rows[r]["InvNumber"].ToString();
                            NewRow["Notes"] = Rows[r]["Notes"].ToString();
                            NewRow["Cvet"] = Cvet;
                            NewRow["Patina"] = Patina;
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
                    string Cvet = GetColorCode(Convert.ToInt32(Rows[r]["ColorID"]));
                    string Patina = GetPatinaCode(Convert.ToInt32(Rows[r]["PatinaID"]));
                    if (IsProfil(Convert.ToInt32(Rows[r]["DecorConfigID"])))
                    {
                        DataRow[] InvRows = PDT.Select($"InvNumber = '{ Rows[r]["InvNumber"].ToString() }' AND " +
                                                       $" Notes = '{ Rows[r]["Notes"].ToString() }' AND Cvet = '{ Cvet }' AND Patina = '{ Patina }'");
                        if (InvRows.Count() == 0)
                        {
                            DataRow NewRow = PDT.NewRow();
                            NewRow["UNN"] = UNN;
                            NewRow["PaymentRate"] = PaymentRate;
                            NewRow["AccountingName"] = Rows[r]["AccountingName"].ToString();
                            NewRow["InvNumber"] = Rows[r]["InvNumber"].ToString();
                            NewRow["Notes"] = Rows[r]["Notes"].ToString();
                            NewRow["Cvet"] = Cvet;
                            NewRow["Patina"] = Patina;
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
                        DataRow[] InvRows = TDT.Select($"InvNumber = '{ Rows[r]["InvNumber"].ToString() }' AND " +
                                                       $" Notes = '{ Rows[r]["Notes"].ToString() }' AND Cvet = '{ Cvet }' AND Patina = '{ Patina }'");
                        if (InvRows.Count() == 0)
                        {
                            DataRow NewRow = TDT.NewRow();
                            NewRow["UNN"] = UNN;
                            NewRow["PaymentRate"] = PaymentRate;
                            NewRow["AccountingName"] = Rows[r]["AccountingName"].ToString();
                            NewRow["InvNumber"] = Rows[r]["InvNumber"].ToString();
                            NewRow["Notes"] = Rows[r]["Notes"].ToString();
                            NewRow["Cvet"] = Cvet;
                            NewRow["Patina"] = Patina;
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
                        DataRow NewRow = _profilReportDataTable.NewRow();

                        NewRow["UNN"] = UNN;
                        NewRow["PaymentRate"] = PaymentRate;
                        NewRow["AccountingName"] = PDT.Rows[g]["AccountingName"].ToString();
                        NewRow["InvNumber"] = PDT.Rows[g]["InvNumber"].ToString();
                        NewRow["Notes"] = PDT.Rows[g]["Notes"].ToString();
                        NewRow["Cvet"] = PDT.Rows[g]["Cvet"].ToString();
                        NewRow["Patina"] = PDT.Rows[g]["Patina"].ToString();
                        NewRow["CurrencyCode"] = ProfilCurrencyCode;
                        NewRow["Count"] = PDT.Rows[g]["Count"];
                        NewRow["PriceWithTransport"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[g]["CostWithTransport"]) / Convert.ToDecimal(PDT.Rows[g]["Count"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["CostWithTransport"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[g]["CostWithTransport"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["Cost"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[g]["Cost"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["Weight"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[g]["Weight"]), 3, MidpointRounding.AwayFromZero);

                        _profilReportDataTable.Rows.Add(NewRow);
                    }

                    if (p3.Length > 0)
                    {
                        DataRow NewRow = _profilReportDataTable.NewRow();

                        NewRow["UNN"] = UNN;
                        NewRow["PaymentRate"] = PaymentRate;
                        NewRow["AccountingName"] = PDT.Rows[g]["AccountingName"].ToString();
                        NewRow["InvNumber"] = PDT.Rows[g]["InvNumber"].ToString();
                        NewRow["Notes"] = PDT.Rows[g]["Notes"].ToString();
                        NewRow["Cvet"] = PDT.Rows[g]["Cvet"].ToString();
                        NewRow["Patina"] = PDT.Rows[g]["Patina"].ToString();
                        NewRow["CurrencyCode"] = ProfilCurrencyCode;
                        NewRow["Count"] = PDT.Rows[g]["Count"];
                        NewRow["PriceWithTransport"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[g]["CostWithTransport"]) / Convert.ToDecimal(PDT.Rows[g]["Count"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["CostWithTransport"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[g]["CostWithTransport"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["Cost"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[g]["Cost"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["Weight"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[g]["Weight"]), 3, MidpointRounding.AwayFromZero);

                        _profilReportDataTable.Rows.Add(NewRow);
                    }

                    if (p1.Length == 0 && p2.Length == 0 && p3.Length == 0)
                    {
                        DataRow NewRow = _profilReportDataTable.NewRow();

                        NewRow["UNN"] = UNN;
                        NewRow["PaymentRate"] = PaymentRate;
                        NewRow["AccountingName"] = PDT.Rows[g]["AccountingName"].ToString();
                        NewRow["InvNumber"] = PDT.Rows[g]["InvNumber"].ToString();
                        NewRow["Notes"] = PDT.Rows[g]["Notes"].ToString();
                        NewRow["Cvet"] = PDT.Rows[g]["Cvet"].ToString();
                        NewRow["Patina"] = PDT.Rows[g]["Patina"].ToString();
                        NewRow["CurrencyCode"] = ProfilCurrencyCode;
                        NewRow["Count"] = PDT.Rows[g]["Count"];
                        NewRow["PriceWithTransport"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[g]["CostWithTransport"]) / Convert.ToDecimal(PDT.Rows[g]["Count"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["CostWithTransport"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[g]["CostWithTransport"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["Cost"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[g]["Cost"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["Weight"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[g]["Weight"]), 3, MidpointRounding.AwayFromZero);

                        _profilReportDataTable.Rows.Add(NewRow);
                    }
                }
            }

            if (TDT.Rows.Count > 0)
            {
                for (int g = 0; g < TDT.Rows.Count; g++)
                {
                    if (p1.Length > 0 && p2.Length > 0)
                    {
                        DataRow NewRow = _tPSReportDataTable.NewRow();

                        NewRow["UNN"] = UNN;
                        NewRow["PaymentRate"] = PaymentRate;
                        NewRow["AccountingName"] = TDT.Rows[g]["AccountingName"].ToString();
                        NewRow["InvNumber"] = TDT.Rows[g]["InvNumber"].ToString();
                        NewRow["Notes"] = TDT.Rows[g]["Notes"].ToString();
                        NewRow["Cvet"] = TDT.Rows[g]["Cvet"].ToString();
                        NewRow["Patina"] = TDT.Rows[g]["Patina"].ToString();
                        NewRow["CurrencyCode"] = ProfilCurrencyCode;
                        NewRow["TPSCurCode"] = TPSCurrencyCode;
                        NewRow["Count"] = TDT.Rows[g]["Count"];
                        NewRow["PriceWithTransport"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[g]["CostWithTransport"]) / Convert.ToDecimal(TDT.Rows[g]["Count"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["CostWithTransport"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[g]["CostWithTransport"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["Cost"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[g]["Cost"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["Weight"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[g]["Weight"]), 3, MidpointRounding.AwayFromZero);

                        _tPSReportDataTable.Rows.Add(NewRow);
                    }

                    if (p3.Length > 0)
                    {
                        DataRow NewRow = _tPSReportDataTable.NewRow();

                        NewRow["UNN"] = UNN;
                        NewRow["PaymentRate"] = PaymentRate;
                        NewRow["AccountingName"] = TDT.Rows[g]["AccountingName"].ToString();
                        NewRow["InvNumber"] = TDT.Rows[g]["InvNumber"].ToString();
                        NewRow["Notes"] = TDT.Rows[g]["Notes"].ToString();
                        NewRow["Cvet"] = TDT.Rows[g]["Cvet"].ToString();
                        NewRow["Patina"] = TDT.Rows[g]["Patina"].ToString();
                        NewRow["CurrencyCode"] = ProfilCurrencyCode;
                        NewRow["TPSCurCode"] = TPSCurrencyCode;
                        NewRow["Count"] = TDT.Rows[g]["Count"];
                        NewRow["PriceWithTransport"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[g]["CostWithTransport"]) / Convert.ToDecimal(TDT.Rows[g]["Count"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["CostWithTransport"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[g]["CostWithTransport"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["Cost"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[g]["Cost"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["Weight"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[g]["Weight"]), 3, MidpointRounding.AwayFromZero);

                        _tPSReportDataTable.Rows.Add(NewRow);
                    }

                    if (p1.Length == 0 && p2.Length == 0 && p3.Length == 0)
                    {
                        DataRow NewRow = _tPSReportDataTable.NewRow();

                        NewRow["UNN"] = UNN;
                        NewRow["PaymentRate"] = PaymentRate;
                        NewRow["AccountingName"] = TDT.Rows[g]["AccountingName"].ToString();
                        NewRow["InvNumber"] = TDT.Rows[g]["InvNumber"].ToString();
                        NewRow["Notes"] = TDT.Rows[g]["Notes"].ToString();
                        NewRow["Cvet"] = TDT.Rows[g]["Cvet"].ToString();
                        NewRow["Patina"] = TDT.Rows[g]["Patina"].ToString();
                        NewRow["CurrencyCode"] = ProfilCurrencyCode;
                        NewRow["TPSCurCode"] = TPSCurrencyCode;
                        NewRow["Count"] = TDT.Rows[g]["Count"];
                        NewRow["PriceWithTransport"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[g]["CostWithTransport"]) / Convert.ToDecimal(TDT.Rows[g]["Count"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["CostWithTransport"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[g]["CostWithTransport"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["Cost"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[g]["Cost"]), 2, MidpointRounding.AwayFromZero);
                        NewRow["Weight"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[g]["Weight"]), 3, MidpointRounding.AwayFromZero);

                        _tPSReportDataTable.Rows.Add(NewRow);
                    }
                }
            }


        }

        private void Collect()
        {
            DataTable DistRatesDT = new DataTable();
            DataTable Items = new DataTable();

            using (DataView DV = new DataView(DecorOrdersDataTable))
            {
                Items = DV.ToTable(true, new string[] { "InvNumber", "Notes", "ColorID", "PatinaID" });
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

                    int ColorID = Convert.ToInt32(Items.Rows[i]["ColorID"]);
                    int PatinaID = Convert.ToInt32(Items.Rows[i]["PatinaID"]);
                    string InvNumber = Items.Rows[i]["InvNumber"].ToString();
                    string Notes = Items.Rows[i]["Notes"].ToString();
                    DataRow[] ItemsRows = DecorOrdersDataTable.Select("PaymentRate = '" + PaymentRate.ToString() + "' AND InvNumber = '" + InvNumber +
                                                                      "' AND Notes = '" + Notes +
                        "' AND ColorID=" + ColorID + " AND PatinaID=" + PatinaID,
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

                        DataRow[] DCs = DecorConfigDataTable.Select("DecorConfigID = " +
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

        public void Report(int[] MainOrderIDs, bool IsSample)
        {
            GetMegaOrderInfo(MainOrderIDs[0]);
            DecorOrdersDataTable.Clear();
            _profilReportDataTable.Clear();
            _tPSReportDataTable.Clear();
            string sWhere = "";

            for (int i = 0; i < MainOrderIDs.Count(); i++)
            {
                if (sWhere != "")
                    sWhere += " OR MainOrderID = " + MainOrderIDs[i].ToString();
                else
                    sWhere += "MainOrderID = " + MainOrderIDs[i].ToString();
            }

            string SelectCommand = $@"SELECT 
CASE WHEN NewMainOrders.Notes = '' THEN CAST(NewMegaOrders.OrderNumber AS varchar(12)) + '_' + CAST(NewDecorOrders.MainOrderID AS varchar(12)) ELSE CAST(NewMegaOrders.OrderNumber AS varchar(12)) 
                         + '_' + CAST(NewDecorOrders.MainOrderID AS varchar(12)) + '_' + NewMainOrders.Notes END AS Notes,
NewDecorOrders.*, infiniu2_catalog.dbo.DecorConfig.AccountingName, infiniu2_catalog.dbo.DecorConfig.InvNumber, NewMegaOrders.PaymentRate FROM NewDecorOrders" +
                " INNER JOIN NewMainOrders ON NewDecorOrders.MainOrderID = NewMainOrders.MainOrderID" +
                " INNER JOIN NewMegaOrders ON NewMainOrders.MegaOrderID = NewMegaOrders.MegaOrderID" +
                " INNER JOIN infiniu2_catalog.dbo.DecorConfig ON NewDecorOrders.DecorConfigID = infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                " WHERE NewDecorOrders.IsSample=1 AND InvNumber IS NOT NULL AND NewDecorOrders.MainOrderID IN (" + string.Join(",", MainOrderIDs) + ") ORDER BY InvNumber";
            if (!IsSample)
                SelectCommand = $@"SELECT 
CASE WHEN NewMainOrders.Notes = '' THEN CAST(NewMegaOrders.OrderNumber AS varchar(12)) + '_' + CAST(NewDecorOrders.MainOrderID AS varchar(12)) ELSE CAST(NewMegaOrders.OrderNumber AS varchar(12)) 
                         + '_' + CAST(NewDecorOrders.MainOrderID AS varchar(12)) + '_' + NewMainOrders.Notes END AS Notes,
NewDecorOrders.*, infiniu2_catalog.dbo.DecorConfig.AccountingName, infiniu2_catalog.dbo.DecorConfig.InvNumber, NewMegaOrders.PaymentRate FROM NewDecorOrders" +
                    " INNER JOIN NewMainOrders ON NewDecorOrders.MainOrderID = NewMainOrders.MainOrderID" +
                    " INNER JOIN NewMegaOrders ON NewMainOrders.MegaOrderID = NewMegaOrders.MegaOrderID" +
                    " INNER JOIN infiniu2_catalog.dbo.DecorConfig ON NewDecorOrders.DecorConfigID = infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                    " WHERE NewDecorOrders.IsSample=0 AND InvNumber IS NOT NULL AND NewDecorOrders.MainOrderID IN (" + string.Join(",", MainOrderIDs) + ") ORDER BY InvNumber";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(DecorOrdersDataTable);
            }

            Collect();

        }

        public void Report(int[] MainOrderIDs)
        {
            GetMegaOrderInfo(MainOrderIDs[0]);
            DecorOrdersDataTable.Clear();
            _profilReportDataTable.Clear();
            _tPSReportDataTable.Clear();

            string SelectCommand = $@"SELECT 
CASE WHEN NewMainOrders.Notes = '' THEN CAST(NewMegaOrders.OrderNumber AS varchar(12)) + '_' + CAST(NewDecorOrders.MainOrderID AS varchar(12)) ELSE CAST(NewMegaOrders.OrderNumber AS varchar(12)) 
                         + '_' + CAST(NewDecorOrders.MainOrderID AS varchar(12)) + '_' + NewMainOrders.Notes END AS Notes,
NewDecorOrders.*, infiniu2_catalog.dbo.DecorConfig.AccountingName, infiniu2_catalog.dbo.DecorConfig.InvNumber, NewMegaOrders.PaymentRate FROM NewDecorOrders" +
                " INNER JOIN NewMainOrders ON NewDecorOrders.MainOrderID = NewMainOrders.MainOrderID" +
                " INNER JOIN NewMegaOrders ON NewMainOrders.MegaOrderID = NewMegaOrders.MegaOrderID" +
                " INNER JOIN infiniu2_catalog.dbo.DecorConfig ON NewDecorOrders.DecorConfigID = infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                " WHERE InvNumber IS NOT NULL AND NewDecorOrders.MainOrderID IN (" + string.Join(",", MainOrderIDs) + ") ORDER BY InvNumber";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(DecorOrdersDataTable);
            }

            Collect();

        }
    }
}
