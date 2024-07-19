using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;

using System;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;

namespace Infinium.Modules.Marketing.NewOrders.PrepareReport.ColorInvoiceReportToDbf
{
    public class ColorInvoiceFrontsReportToDbf
    {
        private decimal TransportCost = 0;
        private decimal AdditionalCost = 0;
        private decimal PaymentRate = 1;
        private int ClientID = 0;
        //int CurrencyCode = 0;
        private string ProfilCurrencyCode = "0";
        private string TPSCurrencyCode = "0";
        private string UNN = string.Empty;

        private DataTable DecorInvNumbersDT = null;
        private DataTable CurrencyTypesDT;
        private DataTable ProfilFrontsOrdersDataTable = null;
        private DataTable TPSFrontsOrdersDataTable = null;
        private DataTable FrontsDataTable = null;
        private DataTable FrameColorsDataTable = null;
        private DataTable PatinaDataTable = null;
        private DataTable PatinaRALDataTable = null;
        private DataTable InsetTypesDataTable = null;
        private DataTable InsetColorsDataTable = null;
        private DataTable MeasuresDataTable = null;
        private DataTable FactoryDataTable = null;
        private DataTable GridSizesDataTable = null;
        private DataTable FrontsConfigDataTable = null;
        private DataTable DecorConfigDataTable = null;
        private DataTable TechStoreDataTable = null;
        private DataTable InsetPriceDataTable = null;
        private DataTable AluminiumFrontsDataTable = null;

        private DataTable _profilReportDataTable = null;
        private DataTable _tPSReportDataTable = null;

        public DataTable ProfilReportDataTable { get => _profilReportDataTable; }
        public DataTable TPSReportDataTable { get => _tPSReportDataTable; }

        public ColorInvoiceFrontsReportToDbf()
        {
            Create();
            CreateReportDataTables();
        }

        private void GetColorsDT()
        {
            FrameColorsDataTable = new DataTable();
            FrameColorsDataTable.Columns.Add(new DataColumn("ColorID", Type.GetType("System.Int64")));
            FrameColorsDataTable.Columns.Add(new DataColumn("ColorName", Type.GetType("System.String")));
            FrameColorsDataTable.Columns.Add(new DataColumn("Cvet", Type.GetType("System.String")));
            string SelectCommand = $"SELECT TechStoreID, TechStoreName, Cvet FROM TechStore " +
                $" WHERE TechStoreSubGroupID IN (SELECT TechStoreSubGroupID FROM TechStoreSubGroups WHERE (TechStoreGroupID = 11 OR TechStoreGroupID = 1))" +
                $" ORDER BY TechStoreName";
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

        private void GetInsetColorsDT()
        {
            InsetColorsDataTable = new DataTable();
            InsetColorsDataTable.Columns.Add(new DataColumn("InsetColorID", Type.GetType("System.Int64")));
            InsetColorsDataTable.Columns.Add(new DataColumn("GroupID", Type.GetType("System.Int64")));
            InsetColorsDataTable.Columns.Add(new DataColumn("InsetColorName", Type.GetType("System.String")));
            InsetColorsDataTable.Columns.Add(new DataColumn("Cvet", Type.GetType("System.String")));
            using (var DA = new SqlDataAdapter(
                       @"SELECT InsetColors.InsetColorID, InsetColors.GroupID, TechStore.TechStoreName AS InsetColorName, 
TechStore.Cvet AS Cvet FROM InsetColors 
INNER JOIN infiniu2_catalog.dbo.TechStore as TechStore 
ON InsetColors.InsetColorID = TechStore.TechStoreID ORDER BY TechStoreName",
                       ConnectionStrings.CatalogConnectionString))
            {
                using (var DT = new DataTable())
                {
                    DA.Fill(DT);
                    {
                        var NewRow = InsetColorsDataTable.NewRow();
                        NewRow["InsetColorID"] = -1;
                        NewRow["GroupID"] = -1;
                        NewRow["InsetColorName"] = "-";
                        NewRow["Cvet"] = "000";
                        InsetColorsDataTable.Rows.Add(NewRow);
                    }
                    {
                        var NewRow = InsetColorsDataTable.NewRow();
                        NewRow["InsetColorID"] = 0;
                        NewRow["GroupID"] = -1;
                        NewRow["InsetColorName"] = "на выбор";
                        NewRow["Cvet"] = "0000000";
                        InsetColorsDataTable.Rows.Add(NewRow);
                    }
                    for (var i = 0; i < DT.Rows.Count; i++)
                    {
                        var NewRow = InsetColorsDataTable.NewRow();
                        NewRow["InsetColorID"] = Convert.ToInt64(DT.Rows[i]["InsetColorID"]);
                        NewRow["GroupID"] = Convert.ToInt64(DT.Rows[i]["GroupID"]);
                        NewRow["InsetColorName"] = DT.Rows[i]["InsetColorName"].ToString();
                        NewRow["Cvet"] = DT.Rows[i]["Cvet"].ToString();
                        InsetColorsDataTable.Rows.Add(NewRow);
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
            ProfilFrontsOrdersDataTable = new DataTable();
            TPSFrontsOrdersDataTable = new DataTable();

            FrontsDataTable = new DataTable();
            InsetTypesDataTable = new DataTable();
            InsetPriceDataTable = new DataTable();
            AluminiumFrontsDataTable = new DataTable();

            string SelectCommand = "SELECT * FROM CurrencyTypes";
            CurrencyTypesDT = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(CurrencyTypesDT);
            }

            SelectCommand = @"SELECT TechStoreID AS FrontID, TechStoreName AS FrontName FROM TechStore 
                WHERE TechStoreID IN (SELECT FrontID FROM FrontsConfig WHERE Enabled = 1)
                ORDER BY TechStoreName";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(FrontsDataTable);
            }
            GetColorsDT();
            GetInsetColorsDT();
            GetPatinaDT();

            SelectCommand = @"SELECT * FROM InsetTypes";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(InsetTypesDataTable);
            }
            MeasuresDataTable = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM Measures",
                ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(MeasuresDataTable);
            }

            InsetPriceDataTable = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM InsetPrice",
                ConnectionStrings.MarketingReferenceConnectionString))
            {
                DA.Fill(InsetPriceDataTable);
            }

            AluminiumFrontsDataTable = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM AluminiumFronts",
                ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(AluminiumFrontsDataTable);
            }

            FactoryDataTable = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM Factory",
                ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(FactoryDataTable);
            }

            GridSizesDataTable = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM GridSizes",
                ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(GridSizesDataTable);
            }

            FrontsConfigDataTable = new DataTable();
            //using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM FrontsConfig",
            //    ConnectionStrings.CatalogConnectionString))
            //{
            //    DA.Fill(FrontsConfigDataTable);
            //}
            FrontsConfigDataTable = TablesManager.FrontsConfigDataTableAll;

            DecorConfigDataTable = new DataTable();
            //using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM DecorConfig",
            //    ConnectionStrings.CatalogConnectionString))
            //{
            //    DA.Fill(DecorConfigDataTable);
            //}
            DecorConfigDataTable = TablesManager.DecorConfigDataTableAll;

            TechStoreDataTable = new DataTable();
            //using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM TechStore",
            //    ConnectionStrings.CatalogConnectionString))
            //{
            //    DA.Fill(TechStoreDataTable);
            //}
            TechStoreDataTable = TablesManager.TechStoreDataTable;
        }

        public void ClearReport()
        {
            _profilReportDataTable.Clear();
            _tPSReportDataTable.Clear();
        }

        private void CreateReportDataTables()
        {
            DecorInvNumbersDT = new DataTable();
            DecorInvNumbersDT.Columns.Add(new DataColumn("NewFrontsOrdersID", Type.GetType("System.Int32")));
            DecorInvNumbersDT.Columns.Add(new DataColumn("DecorInvNumber", Type.GetType("System.String")));
            DecorInvNumbersDT.Columns.Add(new DataColumn("DecorAccountingName", Type.GetType("System.String")));
            DecorInvNumbersDT.Columns.Add(new DataColumn("Cvet", Type.GetType("System.String")));
            DecorInvNumbersDT.Columns.Add(new DataColumn("Patina", Type.GetType("System.String")));
            DecorInvNumbersDT.Columns.Add(new DataColumn("FactoryID", Type.GetType("System.Int32")));

            _profilReportDataTable = new DataTable();
            _profilReportDataTable.Columns.Add(new DataColumn("UNN", Type.GetType("System.String")));
            _profilReportDataTable.Columns.Add(new DataColumn("CurrencyCode", Type.GetType("System.String")));
            _profilReportDataTable.Columns.Add(new DataColumn("TPSCurCode", Type.GetType("System.String")));
            _profilReportDataTable.Columns.Add(new DataColumn("InvNumber", Type.GetType("System.String")));
            _profilReportDataTable.Columns.Add(new DataColumn("Cvet", Type.GetType("System.String")));
            _profilReportDataTable.Columns.Add(new DataColumn("Patina", Type.GetType("System.String")));
            _profilReportDataTable.Columns.Add(new DataColumn("AccountingName", Type.GetType("System.String")));
            _profilReportDataTable.Columns.Add(new DataColumn("Count", Type.GetType("System.Decimal")));
            _profilReportDataTable.Columns.Add(new DataColumn("Price", Type.GetType("System.Decimal")));
            _profilReportDataTable.Columns.Add(new DataColumn("Cost", Type.GetType("System.Decimal")));
            _profilReportDataTable.Columns.Add(new DataColumn("Weight", Type.GetType("System.Decimal")));
            _profilReportDataTable.Columns.Add(new DataColumn("OriginalPrice", Type.GetType("System.Decimal")));
            _profilReportDataTable.Columns.Add(new DataColumn("PriceWithTransport", Type.GetType("System.String")));
            _profilReportDataTable.Columns.Add(new DataColumn("CostWithTransport", Type.GetType("System.String")));
            _profilReportDataTable.Columns.Add(new DataColumn("PaymentRate", Type.GetType("System.Decimal")));
            _profilReportDataTable.Columns.Add(new DataColumn("IsNonStandard", Type.GetType("System.Boolean")));
            _profilReportDataTable.Columns.Add(new DataColumn("NonStandardMargin", Type.GetType("System.Decimal")));

            _tPSReportDataTable = _profilReportDataTable.Clone();
        }

        private string GetFrontName(int FrontID)
        {
            DataRow[] Row = FrontsDataTable.Select("FrontID = " + FrontID);

            return Row[0]["FrontName"].ToString();
        }

        private string GetColorCode(int ColorID)
        {
            string ColorName = string.Empty;
            try
            {
                DataRow[] Rows = FrameColorsDataTable.Select("ColorID = " + ColorID);
                ColorName = Rows[0]["Cvet"].ToString();
            }
            catch
            {
                return string.Empty;
            }
            return ColorName;
        }

        public string GetColorNameByCode(string Cvet)
        {
            string ColorName = string.Empty;
            try
            {
                var Rows = FrameColorsDataTable.Select($"Cvet = '{ Cvet } '");
                if (Rows.Any())
                    ColorName = Rows[0]["ColorName"].ToString();
                else
                {
                    Rows = InsetColorsDataTable.Select($"Cvet = '{ Cvet } '");
                    if (Rows.Any())
                        ColorName = Rows[0]["InsetColorName"].ToString();
                }
            }
            catch
            {
                return string.Empty;
            }
            return ColorName;
        }

        public string GetPatinaNameByCode(string Patina)
        {
            string PatinaName = string.Empty;
            try
            {
                DataRow[] Rows = PatinaDataTable.Select($"Patina = '{ Patina } '");
                PatinaName = Rows[0]["DisplayName"].ToString();
            }
            catch
            {
                return string.Empty;
            }
            return PatinaName;
        }

        private string GetPatinaCode(int PatinaID)
        {
            string PatinaName = string.Empty;
            try
            {
                DataRow[] Rows = PatinaDataTable.Select("PatinaID = " + PatinaID);
                PatinaName = Rows[0]["Patina"].ToString();
            }
            catch
            {
                return string.Empty;
            }
            return PatinaName;
        }

        private void SplitTables(DataTable NewFrontsOrdersDataTable, ref DataTable ProfilDT, ref DataTable TPSDT)
        {
            for (int i = 0; i < NewFrontsOrdersDataTable.Rows.Count; i++)
            {
                DataRow[] Rows = FrontsConfigDataTable.Select("FrontConfigID = " + NewFrontsOrdersDataTable.Rows[i]["FrontConfigID"].ToString());

                if (Convert.ToDateTime(NewFrontsOrdersDataTable.Rows[i]["CreateDateTime"]) < new DateTime(2019, 10, 01))
                {
                    if (Rows[0]["AreaID"].ToString() == "1")//profil
                        ProfilDT.ImportRow(NewFrontsOrdersDataTable.Rows[i]);

                    if (Rows[0]["AreaID"].ToString() == "2")//tps
                        TPSDT.ImportRow(NewFrontsOrdersDataTable.Rows[i]);
                }
                else
                {
                    if (Rows[0]["FactoryID"].ToString() == "1")//profil
                        ProfilDT.ImportRow(NewFrontsOrdersDataTable.Rows[i]);

                    if (Rows[0]["FactoryID"].ToString() == "2")//tps
                        TPSDT.ImportRow(NewFrontsOrdersDataTable.Rows[i]);
                }
            }
        }

        //ALUMINIUM
        private int IsAluminium(int FrontID)
        {
            DataRow[] Row = FrontsDataTable.Select("FrontID = " + FrontID);

            if (Row.Count() > 0 && Row[0]["FrontName"].ToString()[0] == 'Z')
            {
                return Convert.ToInt32(Row[0]["FrontID"]);
            }

            return -1;
        }

        //ALUMINIUM
        private int IsAluminium(DataRow NewFrontsOrdersRow)
        {
            string str = NewFrontsOrdersRow["FrontID"].ToString();
            DataRow[] Row = FrontsDataTable.Select("FrontID = " + NewFrontsOrdersRow["FrontID"].ToString());

            if (Row.Count() > 0 && Row[0]["FrontName"].ToString()[0] == 'Z')
            {
                str = Row[0]["FrontName"].ToString();
                return Convert.ToInt32(Row[0]["FrontID"]);
            }

            return -1;
        }

        private void GetGlassMarginAluminium(DataRow NewFrontsOrdersRow, ref int GlassMarginHeight, ref int GlassMarginWidth)
        {
            DataRow[] Rows = TechStoreDataTable.Select("TechStoreID = " + Convert.ToInt32(NewFrontsOrdersRow["FrontID"]));
            if (Rows.Count() > 0)
            {
                if (Rows[0]["InsetHeightAdmission"] != DBNull.Value)
                    GlassMarginHeight = Convert.ToInt32(Rows[0]["InsetHeightAdmission"]);
                if (Rows[0]["InsetWidthAdmission"] != DBNull.Value)
                    GlassMarginWidth = Convert.ToInt32(Rows[0]["InsetWidthAdmission"]);
            }
            //DataRow[] Rows = AluminiumFrontsDataTable.Select("FrontID = " + IsAluminium(NewFrontsOrdersRow));
            //if (Rows.Count() > 0)
            //{
            //    GlassMarginHeight = Convert.ToInt32(Rows[0]["GlassMarginHeight"]);
            //    GlassMarginWidth = Convert.ToInt32(Rows[0]["GlassMarginWidth"]);
            //}
        }

        private decimal GetJobPriceAluminium(DataRow NewFrontsOrdersRow)
        {
            DataRow[] Rows = AluminiumFrontsDataTable.Select("FrontID = " + IsAluminium(NewFrontsOrdersRow));
            if (Rows.Count() == 0)
                return 0;
            return Convert.ToDecimal(Rows[0]["JobPrice"]);
        }

        private decimal FrontsPriceAluminium(DataRow NewFrontsOrdersRow)
        {
            decimal Price = 0;

            DataRow[] Rows = AluminiumFrontsDataTable.Select("FrontID = " + IsAluminium(NewFrontsOrdersRow));
            if (Rows.Count() > 0)
            {
                Price = Convert.ToDecimal(Rows[0]["ProfilPrice"]);
            }
            NewFrontsOrdersRow["FrontPrice"] = Price;

            return Price;
        }

        private decimal InsetPriceAluminium(DataRow NewFrontsOrdersRow)
        {
            decimal Price = 0;

            DataRow[] Rows = InsetPriceDataTable.Select("InsetTypeID = " + NewFrontsOrdersRow["InsetColorID"].ToString());

            if (Rows.Count() > 0)
                Price = Convert.ToDecimal(Rows[0]["GlassZXPrice"]);
            else
                Price = 0;

            NewFrontsOrdersRow["InsetPrice"] = Price;

            return Price;
        }

        private decimal GetFrontCostAluminium(DataRow NewFrontsOrdersRow)
        {
            decimal Cost = 0;
            decimal Perimeter = 0;
            decimal GlassSquare = 0;

            int MarginHeight = 0;
            int MarginWidth = 0;

            decimal GlassPrice = InsetPriceAluminium(NewFrontsOrdersRow);
            decimal JobPrice = GetJobPriceAluminium(NewFrontsOrdersRow);
            decimal ProfilPrice = FrontsPriceAluminium(NewFrontsOrdersRow);

            GetGlassMarginAluminium(NewFrontsOrdersRow, ref MarginHeight, ref MarginWidth);

            decimal Height = Convert.ToInt32(NewFrontsOrdersRow["Height"]);
            decimal Width = Convert.ToInt32(NewFrontsOrdersRow["Width"]);
            decimal Count = Convert.ToInt32(NewFrontsOrdersRow["Count"]);


            Perimeter = Decimal.Round((Height * 2 + Width * 2) / 1000 * Count, 3, MidpointRounding.AwayFromZero);
            GlassSquare = Decimal.Round((Height - MarginHeight) * (Width - MarginWidth) / 1000000, 3, MidpointRounding.AwayFromZero);

            GlassSquare = GlassSquare * Count;
            Cost = Decimal.Round(JobPrice * Count + GlassPrice * GlassSquare + Perimeter * ProfilPrice, 3, MidpointRounding.AwayFromZero);

            Cost = Cost * 100 / 120;

            //NewFrontsOrdersRow["InsetPrice"] = 0;
            //NewFrontsOrdersRow["Cost"] = Cost;
            //NewFrontsOrdersRow["Square"] = (Height * Width * Count) / 1000000;
            //NewFrontsOrdersRow["FrontPrice"] = Cost / Convert.ToDecimal(NewFrontsOrdersRow["Square"]);

            return Cost;
        }

        private decimal GetAluminiumWeight(DataRow NewFrontsOrdersRow, bool WithGlass)
        {
            DataRow[] Row = AluminiumFrontsDataTable.Select("FrontID = " + IsAluminium(NewFrontsOrdersRow));
            if (Row.Count() == 0)
                return 0;
            decimal FrontHeight = Convert.ToDecimal(NewFrontsOrdersRow["Height"]);
            decimal FrontWidth = Convert.ToDecimal(NewFrontsOrdersRow["Width"]);
            decimal Count = Convert.ToInt32(NewFrontsOrdersRow["Count"]);

            int MarginHeight = 0;
            int MarginWidth = 0;

            GetGlassMarginAluminium(NewFrontsOrdersRow, ref MarginHeight, ref MarginWidth);

            DataRow[] FrontsConfigRow = FrontsConfigDataTable.Select("FrontConfigID = " + NewFrontsOrdersRow["FrontConfigID"].ToString());
            if (FrontsConfigRow.Count() == 0)
                return 0;
            int ProfileID = Convert.ToInt32(FrontsConfigRow[0]["ProfileID"]);
            decimal ProfileWeight = 0;
            DataRow[] Rows = TechStoreDataTable.Select("TechStoreID = " + ProfileID);
            if (Rows.Count() > 0)
            {
                if (Rows[0]["Weight"] != DBNull.Value)
                    ProfileWeight = Convert.ToDecimal(Rows[0]["Weight"]);
            }

            decimal GlassSquare = 0;

            if (NewFrontsOrdersRow["InsetColorID"].ToString() != "3946")//если не СТЕКЛО КЛИЕНТА
                GlassSquare = Decimal.Round((FrontHeight - MarginHeight) * (FrontWidth - MarginWidth) / 1000000, 3, MidpointRounding.AwayFromZero);

            decimal GlassWeight = GlassSquare * 10;

            decimal ResultProfileWeight = Decimal.Round((FrontWidth * 2 + FrontHeight * 2) / 1000, 3, MidpointRounding.AwayFromZero) * ProfileWeight;

            if (WithGlass)
                return (ResultProfileWeight + GlassWeight) * Count;
            else
                return (ResultProfileWeight) * Count;
        }

        private string GetGlassInvNumber(int FrontConfigID, ref int FactoryID, ref string AccountingName)
        {
            int FrontID = 0;
            int InsetTypeID = 0;
            int InsetColorID = 0;
            DataRow[] FRows = FrontsConfigDataTable.Select("FrontConfigID = " + FrontConfigID);
            if (FRows.Count() == 0)
                return string.Empty;

            FrontID = Convert.ToInt32(FRows[0]["FrontID"]);
            InsetTypeID = Convert.ToInt32(FRows[0]["InsetTypeID"]);
            InsetColorID = Convert.ToInt32(FRows[0]["InsetColorID"]);
            if (IsAluminium(FrontID) != -1)
                return string.Empty;
            DataRow[] DRows = DecorConfigDataTable.Select("DecorID  = " + InsetColorID);
            if (DRows.Count() == 0)
                return string.Empty;
            FactoryID = Convert.ToInt32(DRows[0]["FactoryID"]);
            AccountingName = DRows[0]["AccountingName"].ToString();
            return DRows[0]["InvNumber"].ToString();
        }

        private string GetGridInvNumber(int FrontConfigID, ref int FactoryID, ref string AccountingName)
        {
            int InsetTypeID = 0;
            int InsetColorID = 0;
            DataRow[] FRows = FrontsConfigDataTable.Select("FrontConfigID = " + FrontConfigID);
            if (FRows.Count() == 0)
                return string.Empty;
            InsetTypeID = Convert.ToInt32(FRows[0]["InsetTypeID"]);
            InsetColorID = Convert.ToInt32(FRows[0]["InsetColorID"]);
            DataRow[] DRows = DecorConfigDataTable.Select("DecorID  = " + InsetTypeID + " AND ColorID=" + InsetColorID);
            if (DRows.Count() == 0)
                return string.Empty;
            FactoryID = Convert.ToInt32(DRows[0]["FactoryID"]);
            AccountingName = DRows[0]["AccountingName"].ToString();
            return DRows[0]["InvNumber"].ToString();
        }

        private void GetMegaOrderInfo(int MainOrderID)
        {
            string SelectCommand = $"SELECT MegaOrderID, TransportCost, AdditionalCost, PaymentRate, ClientID, CurrencyTypeID FROM NewMegaOrders" +
                $" WHERE MegaOrderID IN (SELECT MegaOrderID FROM NewMainOrders WHERE MainOrderID = { MainOrderID })";
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

        private void GetMegaOrderInfo1(int MegaOrderID)
        {
            string SelectCommand = $"SELECT MegaOrderID, TransportCost, AdditionalCost, PaymentRate, ClientID, CurrencyTypeID FROM NewMegaOrders" +
                $" WHERE MegaOrderID = {MegaOrderID}";
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

        private int GetMeasureType(int FrontConfigID)
        {
            return Convert.ToInt32(FrontsConfigDataTable.Select("FrontConfigID = " + FrontConfigID.ToString())[0]["MeasureID"]);
        }

        private decimal GetInsetSquare(int FrontID, int Height, int Width)
        {
            decimal GridHeight = 0;
            decimal GridWidth = 0;
            DataRow[] Rows = TechStoreDataTable.Select("TechStoreID = " + FrontID);
            if (Rows.Count() > 0)
            {
                if (Rows[0]["InsetHeightAdmission"] != DBNull.Value)
                    GridHeight = Height - Convert.ToInt32(Rows[0]["InsetHeightAdmission"]);
                if (Rows[0]["InsetWidthAdmission"] != DBNull.Value)
                    GridWidth = Width - Convert.ToInt32(Rows[0]["InsetWidthAdmission"]);

                if (FrontID == 3729)
                {
                    return Decimal.Round(Convert.ToInt32(Rows[0]["InsetHeightAdmission"]) * GridWidth / 1000000, 3, MidpointRounding.AwayFromZero);
                }
            }
            return Decimal.Round(GridHeight * GridWidth / 1000000, 3, MidpointRounding.AwayFromZero);
        }

        private decimal GetInsetSquare(DataRow NewFrontsOrdersRow)
        {
            decimal GridHeight = 0;
            decimal GridWidth = 0;
            DataRow[] Rows = TechStoreDataTable.Select("TechStoreID = " + Convert.ToInt32(NewFrontsOrdersRow["FrontID"]));
            if (Rows.Count() > 0)
            {
                if (Rows[0]["InsetHeightAdmission"] != DBNull.Value)
                    GridHeight = Convert.ToInt32(NewFrontsOrdersRow["Height"]) - Convert.ToInt32(Rows[0]["InsetHeightAdmission"]);
                if (Rows[0]["InsetWidthAdmission"] != DBNull.Value)
                    GridWidth = Convert.ToInt32(NewFrontsOrdersRow["Width"]) - Convert.ToInt32(Rows[0]["InsetWidthAdmission"]);

                if (Convert.ToInt32(NewFrontsOrdersRow["FrontID"]) == 3729)
                {
                    return Decimal.Round(Convert.ToInt32(Rows[0]["InsetHeightAdmission"]) * GridWidth / 1000000, 3, MidpointRounding.AwayFromZero);
                }
            }
            return Decimal.Round(GridHeight * GridWidth / 1000000, 3, MidpointRounding.AwayFromZero);
        }

        private decimal GetNonStandardMargin(int FrontConfigID)
        {
            DataRow[] Rows = FrontsConfigDataTable.Select("FrontConfigID = " + FrontConfigID);

            return Convert.ToDecimal(Rows[0]["ZOVNonStandMargin"]);
        }

        private void GetSimpleFronts(DataTable OrdersDataTable, DataTable ReportDataTable, bool IsNonStandard)
        {
            string IsNonStandardFilter = "IsNonStandard=0";
            string DecorAccountingName = string.Empty;
            string DecorInvNumber = string.Empty;
            string AccountingName = OrdersDataTable.Rows[0]["AccountingName"].ToString();
            string InvNumber = OrdersDataTable.Rows[0]["InvNumber"].ToString();
            int fID = Convert.ToInt32(OrdersDataTable.Rows[0]["FactoryID"]);
            DataTable Fronts = new DataTable();
            if (IsNonStandard)
                IsNonStandardFilter = "IsNonStandard=1";
            using (DataView DV = new DataView(OrdersDataTable, IsNonStandardFilter + " and measureId=1", "", DataViewRowState.CurrentRows))
            {
                Fronts = DV.ToTable(true, new string[] { "FrontID", "ColorID", "PatinaID" });
            }

            for (int i = 0; i < Fronts.Rows.Count; i++)
            {
                decimal SolidCount = 0;
                decimal SolidCost = 0;
                decimal WithTransportSolidCost = 0;
                decimal SolidWeight = 0;

                decimal FilenkaCount = 0;
                decimal FilenkaCost = 0;
                decimal WithTransportFilenkaCost = 0;
                decimal FilenkaWeight = 0;

                decimal VitrinaCount = 0;
                decimal VitrinaCost = 0;
                decimal WithTransportVitrinaCost = 0;
                decimal VitrinaWeight = 0;

                decimal LuxMegaCount = 0;
                decimal LuxMegaCost = 0;
                decimal WithTransportLuxMegaCost = 0;
                decimal LuxMegaWeight = 0;

                object NonStandardMargin = 0;
                IsNonStandardFilter = " AND IsNonStandard=0";
                if (IsNonStandard)
                    IsNonStandardFilter = " AND IsNonStandard=1";
                //ГЛУХИЕ, БЕЗ ВСТАВКИ, РЕШЕТКА ОВАЛ
                DataRow[] rows = InsetTypesDataTable.Select("InsetTypeID=-1 OR GroupID = 3 OR GroupID = 4");
                string filter = string.Empty;
                foreach (DataRow item in rows)
                    filter += item["InsetTypeID"].ToString() + ",";
                if (filter.Length > 0)
                    filter = " AND NOT (FrontID IN (3728,3731,3732,3739,3740,3741,3744,3745,3746) OR InsetTypeID IN (28961,3653,3654,3655)) AND (FrontID = 3729 OR InsetTypeID IN (" + filter.Substring(0, filter.Length - 1) + "))";
                DataRow[] Rows = OrdersDataTable.Select("ColorID = " + Fronts.Rows[i]["ColorID"].ToString() + 
                    " AND PatinaID = " + Fronts.Rows[i]["PatinaID"].ToString() + 
                    " AND FrontID = " + Fronts.Rows[i]["FrontID"].ToString() + " AND (Width <> -1)" + filter + IsNonStandardFilter);
                for (int r = 0; r < Rows.Count(); r++)
                {
                    decimal DeductibleCost = 0;
                    decimal DeductibleCount = 0;
                    decimal DeductibleWeight = 0;
                    if (Convert.ToInt32(Fronts.Rows[i]["FrontID"]) == 3729)//РЕШЕТКА ОВАЛ
                    {
                        int FactoryID = 0;
                        DecorInvNumber = GetGridInvNumber(Convert.ToInt32(Rows[r]["FrontConfigID"]), ref FactoryID, ref DecorAccountingName);
                        if (DecorInvNumber.Length > 0)
                        {
                            DataRow NewRow = DecorInvNumbersDT.NewRow();
                            NewRow["NewFrontsOrdersID"] = Convert.ToInt32(Rows[r]["FrontsOrdersID"]);
                            NewRow["FactoryID"] = FactoryID;
                            NewRow["DecorAccountingName"] = DecorAccountingName;
                            NewRow["Cvet"] = GetColorCode(Convert.ToInt32(Fronts.Rows[i]["ColorID"]));
                            NewRow["Patina"] = GetPatinaCode(Convert.ToInt32(Fronts.Rows[i]["PatinaID"]));
                            NewRow["DecorInvNumber"] = DecorInvNumber;
                            DecorInvNumbersDT.Rows.Add(NewRow);
                            DeductibleWeight = GetInsetWeight(Rows[r]);

                            int MarginHeight = 0;
                            int MarginWidth = 0;
                            GetGlassMarginAluminium(Rows[r], ref MarginHeight, ref MarginWidth);
                            decimal InsetSquare = MarginHeight * (Convert.ToDecimal(Rows[r]["Width"]) - MarginWidth) / 1000000;
                            InsetSquare = Decimal.Round(InsetSquare, 3, MidpointRounding.AwayFromZero);
                            DeductibleCount = InsetSquare * Convert.ToDecimal(Rows[r]["Count"]);
                            DeductibleCost = Convert.ToDecimal(Rows[r]["InsetPrice"]) * InsetSquare * Convert.ToDecimal(Rows[r]["Count"]);
                            DeductibleWeight = Decimal.Round(DeductibleCount * DeductibleWeight, 3, MidpointRounding.AwayFromZero);
                        }
                    }
                    SolidCount += Convert.ToDecimal(Rows[r]["Square"]);
                    SolidCost += Convert.ToDecimal(Rows[r]["FrontPrice"]) * Convert.ToDecimal(Rows[r]["Square"]);

                    DeductibleCost = Math.Ceiling(DeductibleCost / 0.01m) * 0.01m;
                    SolidCost = SolidCost - DeductibleCost;
                    WithTransportSolidCost += Convert.ToDecimal(Rows[r]["CostWithTransport"]) - DeductibleCost;

                    decimal FrontWeight = 0;
                    decimal InsetWeight = 0;

                    GetFrontWeight(Rows[r], ref FrontWeight, ref InsetWeight);

                    SolidWeight += Convert.ToDecimal(FrontWeight + InsetWeight - DeductibleWeight);
                    NonStandardMargin = GetNonStandardMargin(Convert.ToInt32(Rows[r]["FrontConfigID"]));
                }
                //АППЛИКАЦИИ
                filter = " AND (FrontID IN (3728,3731,3732,3739,3740,3741,3744,3745,3746) OR InsetTypeID IN (28961,3653,3654,3655))";
                Rows = OrdersDataTable.Select("ColorID = " + Fronts.Rows[i]["ColorID"].ToString() + 
                    " AND PatinaID = " + Fronts.Rows[i]["PatinaID"].ToString() + 
                    " AND FrontID = " + Fronts.Rows[i]["FrontID"].ToString() + " AND (Width <> -1)" + filter + IsNonStandardFilter);
                for (int r = 0; r < Rows.Count(); r++)
                {
                    if (Convert.ToInt32(Rows[r]["FrontID"]) == 3728 || Convert.ToInt32(Rows[r]["FrontID"]) == 3731 || Convert.ToInt32(Rows[r]["FrontID"]) == 3732 ||
                        Convert.ToInt32(Rows[r]["FrontID"]) == 3739 || Convert.ToInt32(Rows[r]["FrontID"]) == 3740 || Convert.ToInt32(Rows[r]["FrontID"]) == 3741 ||
                        Convert.ToInt32(Rows[r]["FrontID"]) == 3744 || Convert.ToInt32(Rows[r]["FrontID"]) == 3745 || Convert.ToInt32(Rows[r]["FrontID"]) == 3746)
                    {
                        SolidCount += Convert.ToDecimal(Rows[r]["Square"]);
                        SolidCost += Convert.ToDecimal(Rows[r]["FrontPrice"]) * Convert.ToDecimal(Rows[r]["Square"]) + 5 * Convert.ToDecimal(Rows[r]["Count"]);

                        WithTransportSolidCost += Convert.ToDecimal(Rows[r]["CostWithTransport"]);
                        decimal FrontWeight = 0;
                        decimal InsetWeight = 0;

                        GetFrontWeight(Rows[r], ref FrontWeight, ref InsetWeight);

                        SolidWeight += Convert.ToDecimal(FrontWeight + InsetWeight);
                    }
                    else if (Convert.ToInt32(Rows[r]["FrontID"]) == 3415 || Convert.ToInt32(Rows[r]["FrontID"]) == 28922)
                    {
                        FilenkaCount += Convert.ToDecimal(Rows[r]["Square"]);

                        FilenkaCost += Convert.ToDecimal(Rows[r]["FrontPrice"]) * Convert.ToDecimal(Rows[r]["Square"]) + 5 * Convert.ToDecimal(Rows[r]["Count"]);

                        WithTransportFilenkaCost += Convert.ToDecimal(Rows[r]["CostWithTransport"]);
                        decimal FrontWeight = 0;
                        decimal InsetWeight = 0;

                        GetFrontWeight(Rows[r], ref FrontWeight, ref InsetWeight);

                        FilenkaWeight += Convert.ToDecimal(FrontWeight + InsetWeight);
                    }
                    NonStandardMargin = GetNonStandardMargin(Convert.ToInt32(Rows[r]["FrontConfigID"]));
                }
                //ФИЛЕНКА
                filter = " AND InsetTypeID IN (2069,2070,2071,2073,2075,42066,2077,2233,3644,29043,29531,41213)";
                Rows = OrdersDataTable.Select("ColorID = " + Fronts.Rows[i]["ColorID"].ToString() + 
                    " AND PatinaID = " + Fronts.Rows[i]["PatinaID"].ToString() + 
                    " AND FrontID = " + Fronts.Rows[i]["FrontID"].ToString() + " AND (Width <> -1)" + filter + IsNonStandardFilter);
                for (int r = 0; r < Rows.Count(); r++)
                {
                    FilenkaCount += Convert.ToDecimal(Rows[r]["Square"]);

                    FilenkaCost += Convert.ToDecimal(Rows[r]["FrontPrice"]) * Convert.ToDecimal(Rows[r]["Square"]);

                    WithTransportFilenkaCost += Convert.ToDecimal(Rows[r]["CostWithTransport"]);
                    decimal FrontWeight = 0;
                    decimal InsetWeight = 0;

                    GetFrontWeight(Rows[r], ref FrontWeight, ref InsetWeight);

                    FilenkaWeight += Convert.ToDecimal(FrontWeight + InsetWeight);
                    NonStandardMargin = GetNonStandardMargin(Convert.ToInt32(Rows[r]["FrontConfigID"]));
                }
                //ВИТРИНЫ, РЕШЕТКИ, СТЕКЛО
                filter = " AND InsetTypeID IN (1,2,685,686,687,688,29470,29471) AND FrontID <> 3729";
                Rows = OrdersDataTable.Select("ColorID = " + Fronts.Rows[i]["ColorID"].ToString() + 
                    " AND PatinaID = " + Fronts.Rows[i]["PatinaID"].ToString() + 
                    " AND FrontID = " + Fronts.Rows[i]["FrontID"].ToString() + " AND (Width <> -1)" + filter + IsNonStandardFilter);
                for (int r = 0; r < Rows.Count(); r++)
                {
                    decimal DeductibleCount = 0;
                    decimal DeductibleCost = 0;
                    decimal DeductibleWeight = 0;
                    //РЕШЕТКА 45,90,ПЛАСТИК
                    if (Convert.ToInt32(Rows[r]["InsetTypeID"]) == 685 || Convert.ToInt32(Rows[r]["InsetTypeID"]) == 686 ||
                        Convert.ToInt32(Rows[r]["InsetTypeID"]) == 687 || Convert.ToInt32(Rows[r]["InsetTypeID"]) == 688 ||
                        Convert.ToInt32(Rows[r]["InsetTypeID"]) == 29470 || Convert.ToInt32(Rows[r]["InsetTypeID"]) == 29471)
                    {
                        int FactoryID = 0;
                        DecorInvNumber = GetGridInvNumber(Convert.ToInt32(Rows[r]["FrontConfigID"]), ref FactoryID, ref DecorAccountingName);
                        if (DecorInvNumber.Length > 0)
                        {
                            DataRow NewRow = DecorInvNumbersDT.NewRow();
                            NewRow["NewFrontsOrdersID"] = Convert.ToInt32(Rows[r]["FrontsOrdersID"]);
                            NewRow["FactoryID"] = FactoryID;
                            NewRow["Cvet"] = GetColorCode(Convert.ToInt32(Fronts.Rows[i]["ColorID"]));
                            NewRow["Patina"] = GetPatinaCode(Convert.ToInt32(Fronts.Rows[i]["PatinaID"]));
                            NewRow["DecorAccountingName"] = DecorAccountingName;
                            NewRow["DecorInvNumber"] = DecorInvNumber;
                            DecorInvNumbersDT.Rows.Add(NewRow);

                            decimal InsetSquare = GetInsetSquare(Convert.ToInt32(Rows[r]["FrontID"]), Convert.ToInt32(Rows[r]["Height"]), Convert.ToInt32(Rows[r]["Width"]));
                            InsetSquare = Decimal.Round(InsetSquare, 3, MidpointRounding.AwayFromZero);
                            DeductibleCount = InsetSquare * Convert.ToDecimal(Rows[r]["Count"]);
                            DeductibleCost = Convert.ToDecimal(Rows[r]["InsetPrice"]) * InsetSquare * Convert.ToDecimal(Rows[r]["Count"]);
                            DeductibleWeight = Decimal.Round(DeductibleCount * Convert.ToDecimal(3.5), 3, MidpointRounding.AwayFromZero);
                        }
                    }
                    //СТЕКЛО
                    if (Convert.ToInt32(Rows[r]["InsetTypeID"]) == 2)
                    {
                        int FactoryID = 0;
                        DecorInvNumber = GetGlassInvNumber(Convert.ToInt32(Rows[r]["FrontConfigID"]), ref FactoryID, ref DecorAccountingName);
                        if (DecorInvNumber.Length > 0)
                        {
                            DataRow NewRow = DecorInvNumbersDT.NewRow();
                            NewRow["NewFrontsOrdersID"] = Convert.ToInt32(Rows[r]["FrontsOrdersID"]);
                            NewRow["FactoryID"] = FactoryID;
                            NewRow["Cvet"] = GetColorCode(Convert.ToInt32(Fronts.Rows[i]["ColorID"]));
                            NewRow["Patina"] = GetPatinaCode(Convert.ToInt32(Fronts.Rows[i]["PatinaID"]));
                            NewRow["DecorAccountingName"] = DecorAccountingName;
                            NewRow["DecorInvNumber"] = DecorInvNumber;
                            DecorInvNumbersDT.Rows.Add(NewRow);
                            DeductibleCount = Convert.ToDecimal(Rows[r]["Count"]) * GetInsetSquare(Convert.ToInt32(Rows[r]["FrontID"]), Convert.ToInt32(Rows[r]["Height"]), Convert.ToInt32(Rows[r]["Width"]));
                            DeductibleCost = Convert.ToDecimal(Rows[r]["InsetPrice"]) * DeductibleCount;
                            DeductibleWeight = Decimal.Round(DeductibleCount * 10, 3, MidpointRounding.AwayFromZero);
                        }
                    }

                    VitrinaCount += Convert.ToDecimal(Rows[r]["Square"]);

                    if (IsAluminium(Rows[r]) > -1)
                    {
                        VitrinaCost += GetFrontCostAluminium(Rows[r]);
                        VitrinaWeight += GetAluminiumWeight(Rows[r], true);
                    }
                    else
                    {
                        VitrinaCost += Convert.ToDecimal(Rows[r]["FrontPrice"]) * Convert.ToDecimal(Rows[r]["Square"]);

                        decimal FrontWeight = 0;
                        decimal InsetWeight = 0;

                        GetFrontWeight(Rows[r], ref FrontWeight, ref InsetWeight);

                        VitrinaWeight += Convert.ToDecimal(FrontWeight + InsetWeight - DeductibleWeight);
                    }
                    DeductibleCost = Math.Ceiling(DeductibleCost / 0.01m) * 0.01m;
                    WithTransportVitrinaCost += Convert.ToDecimal(Rows[r]["CostWithTransport"]) - DeductibleCost;
                    NonStandardMargin = GetNonStandardMargin(Convert.ToInt32(Rows[r]["FrontConfigID"]));
                }
                //ЛЮКС, МЕГА
                filter = " AND InsetTypeID IN (860,862,4310)";
                Rows = OrdersDataTable.Select("ColorID = " + Fronts.Rows[i]["ColorID"].ToString() + 
                    " AND PatinaID = " + Fronts.Rows[i]["PatinaID"].ToString() + 
                    " AND FrontID = " + Fronts.Rows[i]["FrontID"].ToString() + " AND (Width <> -1)" + filter + IsNonStandardFilter);
                for (int r = 0; r < Rows.Count(); r++)
                {
                    LuxMegaCount += Convert.ToDecimal(Rows[r]["Square"]);

                    LuxMegaCost += Convert.ToDecimal(Rows[r]["FrontPrice"]) * Convert.ToDecimal(Rows[r]["Square"]);

                    WithTransportLuxMegaCost += Convert.ToDecimal(Rows[r]["CostWithTransport"]);
                    decimal FrontWeight = 0;
                    decimal InsetWeight = 0;

                    GetFrontWeight(Rows[r], ref FrontWeight, ref InsetWeight);

                    LuxMegaWeight += Convert.ToDecimal(FrontWeight + InsetWeight);
                    NonStandardMargin = GetNonStandardMargin(Convert.ToInt32(Rows[r]["FrontConfigID"]));
                }

                if (!IsNonStandard)
                    NonStandardMargin = DBNull.Value;

                //SolidCost = Math.Ceiling(SolidCost / 0.01m) * 0.01m;
                //WithTransportSolidCost = Math.Ceiling(WithTransportSolidCost / 0.01m) * 0.01m;
                if (SolidCount > 0)
                {
                    DataRow Row = ReportDataTable.NewRow();
                    Row["NonStandardMargin"] = NonStandardMargin;
                    Row["IsNonStandard"] = IsNonStandard;
                    Row["PaymentRate"] = PaymentRate;
                    Row["OriginalPrice"] = SolidCost / SolidCount;
                    Row["UNN"] = UNN;
                    Row["Cvet"] = GetColorCode(Convert.ToInt32(Fronts.Rows[i]["ColorID"]));
                    Row["Patina"] = GetPatinaCode(Convert.ToInt32(Fronts.Rows[i]["PatinaID"]));
                    Row["AccountingName"] = AccountingName;
                    Row["InvNumber"] = InvNumber;
                    Row["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        Row["TPSCurCode"] = TPSCurrencyCode;
                    Row["Count"] = Decimal.Round(SolidCount, 3, MidpointRounding.AwayFromZero);
                    Row["Price"] = Decimal.Round(SolidCost / SolidCount, 2, MidpointRounding.AwayFromZero);
                    Row["Cost"] = Decimal.Round(Math.Ceiling(SolidCost / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                    Row["PriceWithTransport"] = Decimal.Round(WithTransportSolidCost / SolidCount, 2, MidpointRounding.AwayFromZero);
                    Row["CostWithTransport"] = Decimal.Round(Math.Ceiling(WithTransportSolidCost / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                    Row["Weight"] = Decimal.Round(SolidWeight, 3, MidpointRounding.AwayFromZero);
                    ReportDataTable.Rows.Add(Row);
                }

                //FilenkaCost = Math.Ceiling(FilenkaCost / 0.01m) * 0.01m;
                //WithTransportFilenkaCost = Math.Ceiling(WithTransportFilenkaCost / 0.01m) * 0.01m;
                if (FilenkaCount > 0)
                {
                    DataRow Row = ReportDataTable.NewRow();
                    Row["NonStandardMargin"] = NonStandardMargin;
                    Row["IsNonStandard"] = IsNonStandard;
                    Row["OriginalPrice"] = FilenkaCost / FilenkaCount;
                    Row["UNN"] = UNN;
                    Row["Cvet"] = GetColorCode(Convert.ToInt32(Fronts.Rows[i]["ColorID"]));
                    Row["Patina"] = GetPatinaCode(Convert.ToInt32(Fronts.Rows[i]["PatinaID"]));
                    Row["PaymentRate"] = PaymentRate;
                    Row["AccountingName"] = AccountingName;
                    Row["InvNumber"] = InvNumber;
                    Row["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        Row["TPSCurCode"] = TPSCurrencyCode;
                    Row["Count"] = Decimal.Round(FilenkaCount, 3, MidpointRounding.AwayFromZero);
                    Row["Price"] = Decimal.Round(FilenkaCost / FilenkaCount, 2, MidpointRounding.AwayFromZero);
                    Row["Cost"] = Decimal.Round(Math.Ceiling(FilenkaCost / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                    Row["PriceWithTransport"] = Decimal.Round(WithTransportFilenkaCost / FilenkaCount, 2, MidpointRounding.AwayFromZero);
                    Row["CostWithTransport"] = Decimal.Round(Math.Ceiling(WithTransportFilenkaCost / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                    Row["Weight"] = Decimal.Round(FilenkaWeight, 3, MidpointRounding.AwayFromZero);
                    ReportDataTable.Rows.Add(Row);
                }

                //VitrinaCost = Math.Ceiling(VitrinaCost / 0.01m) * 0.01m;
                //WithTransportVitrinaCost = Math.Ceiling(WithTransportVitrinaCost / 0.01m) * 0.01m;
                if (VitrinaCount > 0)
                {
                    DataRow Row = ReportDataTable.NewRow();
                    Row["NonStandardMargin"] = NonStandardMargin;
                    Row["IsNonStandard"] = IsNonStandard;
                    Row["OriginalPrice"] = VitrinaCost / VitrinaCount;
                    Row["UNN"] = UNN;
                    Row["Cvet"] = GetColorCode(Convert.ToInt32(Fronts.Rows[i]["ColorID"]));
                    Row["Patina"] = GetPatinaCode(Convert.ToInt32(Fronts.Rows[i]["PatinaID"]));
                    Row["PaymentRate"] = PaymentRate;
                    Row["AccountingName"] = AccountingName;
                    Row["InvNumber"] = InvNumber;
                    Row["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        Row["TPSCurCode"] = TPSCurrencyCode;
                    Row["Count"] = Decimal.Round(VitrinaCount, 3, MidpointRounding.AwayFromZero);
                    Row["Price"] = Decimal.Round(VitrinaCost / VitrinaCount, 2, MidpointRounding.AwayFromZero);
                    Row["Cost"] = Decimal.Round(Math.Ceiling(VitrinaCost / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                    Row["PriceWithTransport"] = Decimal.Round(WithTransportVitrinaCost / VitrinaCount, 2, MidpointRounding.AwayFromZero);
                    Row["CostWithTransport"] = Decimal.Round(Math.Ceiling(WithTransportVitrinaCost / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                    Row["Weight"] = Decimal.Round(VitrinaWeight, 3, MidpointRounding.AwayFromZero);
                    ReportDataTable.Rows.Add(Row);
                }

                //LuxMegaCost = Math.Ceiling(LuxMegaCost / 0.01m) * 0.01m;
                //WithTransportLuxMegaCost = Math.Ceiling(WithTransportLuxMegaCost / 0.01m) * 0.01m;
                if (LuxMegaCount > 0)
                {
                    DataRow Row = ReportDataTable.NewRow();
                    Row["NonStandardMargin"] = NonStandardMargin;
                    Row["IsNonStandard"] = IsNonStandard;
                    Row["OriginalPrice"] = LuxMegaCost / LuxMegaCount;
                    Row["UNN"] = UNN;
                    Row["Cvet"] = GetColorCode(Convert.ToInt32(Fronts.Rows[i]["ColorID"]));
                    Row["Patina"] = GetPatinaCode(Convert.ToInt32(Fronts.Rows[i]["PatinaID"]));
                    Row["PaymentRate"] = PaymentRate;
                    Row["AccountingName"] = AccountingName;
                    Row["InvNumber"] = InvNumber;
                    Row["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        Row["TPSCurCode"] = TPSCurrencyCode;
                    Row["Price"] = Decimal.Round(LuxMegaCost / LuxMegaCount, 2, MidpointRounding.AwayFromZero);
                    Row["Count"] = Decimal.Round(LuxMegaCount, 3, MidpointRounding.AwayFromZero);
                    Row["Cost"] = Decimal.Round(Math.Ceiling(LuxMegaCost / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                    Row["PriceWithTransport"] = Decimal.Round(WithTransportLuxMegaCost / LuxMegaCount, 2, MidpointRounding.AwayFromZero);
                    Row["CostWithTransport"] = Decimal.Round(Math.Ceiling(WithTransportLuxMegaCost / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                    Row["Weight"] = Decimal.Round(LuxMegaWeight, 3, MidpointRounding.AwayFromZero);
                    ReportDataTable.Rows.Add(Row);
                }
            }

            Fronts.Dispose();
        }

        private void GetCurvedFronts(DataTable OrdersDataTable, DataTable ReportDataTable)
        {
            string AccountingName = OrdersDataTable.Rows[0]["AccountingName"].ToString();
            string InvNumber = OrdersDataTable.Rows[0]["InvNumber"].ToString();
            int fID = Convert.ToInt32(OrdersDataTable.Rows[0]["FactoryID"]);
            DataTable Fronts = new DataTable();

            using (DataView DV = new DataView(OrdersDataTable, "measureId=3", "", DataViewRowState.CurrentRows))
            {
                Fronts = DV.ToTable(true, new string[] { "FrontID", "ColorID", "PatinaID" });
            }

            for (int i = 0; i < Fronts.Rows.Count; i++)
            {
                DataRow[] Rows = OrdersDataTable.Select("ColorID = " + Fronts.Rows[i]["ColorID"].ToString() +
                    " AND PatinaID = " + Fronts.Rows[i]["PatinaID"].ToString() + 
                    " AND FrontID = " + Fronts.Rows[i]["FrontID"].ToString());

                if (Rows.Count() == 0)
                    continue;

                decimal Solid713Count = 0;
                decimal Solid713Price = 0;
                decimal Solid713WithTransportCost = 0;
                decimal Solid713Weight = 0;

                decimal Filenka713Count = 0;
                decimal Filenka713Price = 0;
                decimal Filenka713WithTransportCost = 0;
                decimal Filenka713Weight = 0;

                decimal NoInset713Count = 0;
                decimal NoInset713Price = 0;
                decimal NoInset713WithTransportCost = 0;
                decimal NoInset713Weight = 0;

                decimal Vitrina713Count = 0;
                decimal Vitrina713Price = 0;
                decimal Vitrina713WithTransportCost = 0;
                decimal Vitrina713Weight = 0;

                decimal Solid910Count = 0;
                decimal Solid910Price = 0;
                decimal Solid910WithTransportCost = 0;
                decimal Solid910Weight = 0;

                decimal Filenka910Count = 0;
                decimal Filenka910Price = 0;
                decimal Filenka910WithTransportCost = 0;
                decimal Filenka910Weight = 0;

                decimal NoInset910Count = 0;
                decimal NoInset910Price = 0;
                decimal NoInset910WithTransportCost = 0;
                decimal NoInset910Weight = 0;

                decimal Vitrina910Count = 0;
                decimal Vitrina910Price = 0;
                decimal Vitrina910WithTransportCost = 0;
                decimal Vitrina910Weight = 0;

                decimal ScagenCount = 0;
                decimal ScagenPrice = 0;
                decimal ScagenWithTransportCost = 0;
                decimal ScagenWeight = 0;

                for (int r = 0; r < Rows.Count(); r++)
                {
                    if (Rows[r]["measureId"].ToString() == "3" && Rows[r]["Width"].ToString() != "-1")
                    {
                        DataRow[] rows = InsetTypesDataTable.Select("GroupID = 3 OR GroupID = 4");
                        foreach (DataRow item in rows)
                        {
                            if (Rows[r]["InsetTypeID"].ToString() == item["InsetTypeID"].ToString())
                            {
                                ScagenCount += Convert.ToDecimal(Rows[r]["Count"]);
                                ScagenPrice = Convert.ToDecimal(Rows[r]["FrontPrice"]);
                                ScagenWithTransportCost += Convert.ToDecimal(Rows[r]["PriceWithTransport"]) * Convert.ToDecimal(Rows[r]["Count"]);

                                decimal FrontWeight = 0;
                                decimal InsetWeight = 0;

                                GetFrontWeight(Rows[r], ref FrontWeight, ref InsetWeight);

                                ScagenWeight += Convert.ToDecimal(FrontWeight + InsetWeight);
                            }
                        }
                        rows = InsetTypesDataTable.Select("InsetTypeID IN (2079,2080,2081,2082,2085,2086,2087,2088,2212,2213,29210,29211,27831,27832,29210,29211)");
                        foreach (DataRow item in rows)
                        {
                            if (Rows[r]["InsetTypeID"].ToString() == item["InsetTypeID"].ToString())
                            {
                                ScagenCount += Convert.ToDecimal(Rows[r]["Count"]);
                                ScagenPrice = Convert.ToDecimal(Rows[r]["FrontPrice"]);
                                ScagenWithTransportCost += Convert.ToDecimal(Rows[r]["PriceWithTransport"]) * Convert.ToDecimal(Rows[r]["Count"]);

                                decimal FrontWeight = 0;
                                decimal InsetWeight = 0;

                                GetFrontWeight(Rows[r], ref FrontWeight, ref InsetWeight);

                                ScagenWeight += Convert.ToDecimal(FrontWeight + InsetWeight);
                            }
                        }
                        if (Rows[r]["InsetTypeID"].ToString() == "-1")
                        {
                            ScagenCount += Convert.ToDecimal(Rows[r]["Count"]);
                            ScagenPrice = Convert.ToDecimal(Rows[r]["FrontPrice"]);
                            ScagenWithTransportCost += Convert.ToDecimal(Rows[r]["PriceWithTransport"]) * Convert.ToDecimal(Rows[r]["Count"]);

                            decimal FrontWeight = 0;
                            decimal InsetWeight = 0;

                            GetFrontWeight(Rows[r], ref FrontWeight, ref InsetWeight);

                            ScagenWeight += Convert.ToDecimal(FrontWeight + InsetWeight);
                        }

                        if (Rows[r]["InsetTypeID"].ToString() == "1")
                        {
                            ScagenCount += Convert.ToDecimal(Rows[r]["Count"]);
                            ScagenPrice = Convert.ToDecimal(Rows[r]["FrontPrice"]);
                            ScagenWithTransportCost += Convert.ToDecimal(Rows[r]["PriceWithTransport"]) * Convert.ToDecimal(Rows[r]["Count"]);

                            decimal FrontWeight = 0;
                            decimal InsetWeight = 0;

                            GetFrontWeight(Rows[r], ref FrontWeight, ref InsetWeight);

                            ScagenWeight += Convert.ToDecimal(FrontWeight + InsetWeight);
                        }
                    }

                    if (Rows[r]["Height"].ToString() == "713")
                    {
                        DataRow[] rows = InsetTypesDataTable.Select("GroupID = 3 OR GroupID = 4");
                        foreach (DataRow item in rows)
                        {
                            if (Rows[r]["InsetTypeID"].ToString() == item["InsetTypeID"].ToString())
                            {
                                Solid713Count += Convert.ToDecimal(Rows[r]["Count"]);
                                Solid713Price = Convert.ToDecimal(Rows[r]["FrontPrice"]);
                                Solid713WithTransportCost += Convert.ToDecimal(Rows[r]["PriceWithTransport"]) * Convert.ToDecimal(Rows[r]["Count"]);

                                decimal FrontWeight = 0;
                                decimal InsetWeight = 0;

                                GetFrontWeight(Rows[r], ref FrontWeight, ref InsetWeight);

                                Solid713Weight += Convert.ToDecimal(FrontWeight + InsetWeight);
                            }
                        }
                        rows = InsetTypesDataTable.Select("InsetTypeID IN (2079,2080,2081,2082,2085,2086,2087,2088,2212,2213,29210,29211,27831,27832,29210,29211)");
                        foreach (DataRow item in rows)
                        {
                            if (Rows[r]["InsetTypeID"].ToString() == item["InsetTypeID"].ToString())
                            {
                                Filenka713Count += Convert.ToDecimal(Rows[r]["Count"]);
                                Filenka713Price = Convert.ToDecimal(Rows[r]["FrontPrice"]);
                                Filenka713WithTransportCost += Convert.ToDecimal(Rows[r]["PriceWithTransport"]) * Convert.ToDecimal(Rows[r]["Count"]);

                                decimal FrontWeight = 0;
                                decimal InsetWeight = 0;

                                GetFrontWeight(Rows[r], ref FrontWeight, ref InsetWeight);

                                Filenka713Weight += Convert.ToDecimal(FrontWeight + InsetWeight);
                            }
                        }
                        if (Rows[r]["InsetTypeID"].ToString() == "-1")
                        {
                            NoInset713Count += Convert.ToDecimal(Rows[r]["Count"]);
                            NoInset713Price = Convert.ToDecimal(Rows[r]["FrontPrice"]);
                            NoInset713WithTransportCost += Convert.ToDecimal(Rows[r]["PriceWithTransport"]) * Convert.ToDecimal(Rows[r]["Count"]);

                            decimal FrontWeight = 0;
                            decimal InsetWeight = 0;

                            GetFrontWeight(Rows[r], ref FrontWeight, ref InsetWeight);

                            NoInset713Weight += Convert.ToDecimal(FrontWeight + InsetWeight);
                        }

                        if (Rows[r]["InsetTypeID"].ToString() == "1")
                        {
                            Vitrina713Count += Convert.ToDecimal(Rows[r]["Count"]);
                            Vitrina713Price = Convert.ToDecimal(Rows[r]["FrontPrice"]);
                            Vitrina713WithTransportCost += Convert.ToDecimal(Rows[r]["PriceWithTransport"]) * Convert.ToDecimal(Rows[r]["Count"]);

                            decimal FrontWeight = 0;
                            decimal InsetWeight = 0;

                            GetFrontWeight(Rows[r], ref FrontWeight, ref InsetWeight);

                            Vitrina713Weight += Convert.ToDecimal(FrontWeight + InsetWeight);
                        }
                    }

                    if (Rows[r]["Height"].ToString() == "910")
                    {
                        DataRow[] rows = InsetTypesDataTable.Select("GroupID = 3 OR GroupID = 4");
                        foreach (DataRow item in rows)
                        {
                            if (Rows[r]["InsetTypeID"].ToString() == item["InsetTypeID"].ToString())
                            {
                                Solid910Count += Convert.ToDecimal(Rows[r]["Count"]);
                                Solid910Price = Convert.ToDecimal(Rows[r]["FrontPrice"]);
                                Solid910WithTransportCost += Convert.ToDecimal(Rows[r]["PriceWithTransport"]) * Convert.ToDecimal(Rows[r]["Count"]);

                                decimal FrontWeight = 0;
                                decimal InsetWeight = 0;

                                GetFrontWeight(Rows[r], ref FrontWeight, ref InsetWeight);

                                Solid910Weight += Convert.ToDecimal(FrontWeight + InsetWeight);
                            }
                        }
                        rows = InsetTypesDataTable.Select("InsetTypeID IN (2079,2080,2081,2082,2085,2086,2087,2088,2212,2213,29210,29211,27831,27832,29210,29211)");
                        foreach (DataRow item in rows)
                        {
                            if (Rows[r]["InsetTypeID"].ToString() == item["InsetTypeID"].ToString())
                            {
                                Filenka910Count += Convert.ToDecimal(Rows[r]["Count"]);
                                Filenka910Price = Convert.ToDecimal(Rows[r]["FrontPrice"]);
                                Filenka910WithTransportCost += Convert.ToDecimal(Rows[r]["PriceWithTransport"]) * Convert.ToDecimal(Rows[r]["Count"]);

                                decimal FrontWeight = 0;
                                decimal InsetWeight = 0;

                                GetFrontWeight(Rows[r], ref FrontWeight, ref InsetWeight);

                                Filenka910Weight += Convert.ToDecimal(FrontWeight + InsetWeight);
                            }
                        }
                        if (Rows[r]["InsetTypeID"].ToString() == "-1")
                        {
                            NoInset910Count += Convert.ToDecimal(Rows[r]["Count"]);
                            NoInset910Price = Convert.ToDecimal(Rows[r]["FrontPrice"]);
                            NoInset910WithTransportCost += Convert.ToDecimal(Rows[r]["PriceWithTransport"]) * Convert.ToDecimal(Rows[r]["Count"]);

                            decimal FrontWeight = 0;
                            decimal InsetWeight = 0;

                            GetFrontWeight(Rows[r], ref FrontWeight, ref InsetWeight);

                            NoInset910Weight += Convert.ToDecimal(FrontWeight + InsetWeight);
                        }

                        if (Rows[r]["InsetTypeID"].ToString() == "1")
                        {
                            Vitrina910Count += Convert.ToDecimal(Rows[r]["Count"]);
                            Vitrina910Price = Convert.ToDecimal(Rows[r]["FrontPrice"]);
                            Vitrina910WithTransportCost += Convert.ToDecimal(Rows[r]["PriceWithTransport"]) * Convert.ToDecimal(Rows[r]["Count"]);

                            decimal FrontWeight = 0;
                            decimal InsetWeight = 0;

                            GetFrontWeight(Rows[r], ref FrontWeight, ref InsetWeight);

                            Vitrina910Weight += Convert.ToDecimal(FrontWeight + InsetWeight);
                        }
                    }

                }

                if (ScagenCount > 0)
                {
                    DataRow Row = ReportDataTable.NewRow();
                    Row["OriginalPrice"] = ScagenPrice;
                    Row["UNN"] = UNN;
                    Row["Cvet"] = GetColorCode(Convert.ToInt32(Fronts.Rows[i]["ColorID"]));
                    Row["Patina"] = GetPatinaCode(Convert.ToInt32(Fronts.Rows[i]["PatinaID"]));
                    Row["PaymentRate"] = PaymentRate;
                    Row["AccountingName"] = AccountingName;
                    Row["InvNumber"] = InvNumber;
                    Row["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        Row["TPSCurCode"] = TPSCurrencyCode;
                    Row["Count"] = ScagenCount;
                    Row["PriceWithTransport"] = Decimal.Round(ScagenWithTransportCost / ScagenCount, 2, MidpointRounding.AwayFromZero);
                    Row["CostWithTransport"] = Decimal.Round(Math.Ceiling(ScagenWithTransportCost / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                    Row["Cost"] = Decimal.Round(Math.Ceiling(ScagenCount * ScagenPrice / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                    Row["Weight"] = Decimal.Round(ScagenWeight, 3, MidpointRounding.AwayFromZero);
                    ReportDataTable.Rows.Add(Row);
                }

                if (Solid713Count > 0)
                {
                    DataRow Row = ReportDataTable.NewRow();
                    Row["OriginalPrice"] = Solid713Price;
                    Row["UNN"] = UNN;
                    Row["Cvet"] = GetColorCode(Convert.ToInt32(Fronts.Rows[i]["ColorID"]));
                    Row["Patina"] = GetPatinaCode(Convert.ToInt32(Fronts.Rows[i]["PatinaID"]));
                    Row["PaymentRate"] = PaymentRate;
                    Row["AccountingName"] = AccountingName;
                    Row["InvNumber"] = InvNumber;
                    Row["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        Row["TPSCurCode"] = TPSCurrencyCode;
                    Row["Count"] = Solid713Count;
                    Row["PriceWithTransport"] = Decimal.Round(Solid713WithTransportCost / Solid713Count, 2, MidpointRounding.AwayFromZero);
                    Row["CostWithTransport"] = Decimal.Round(Math.Ceiling(Solid713WithTransportCost / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                    Row["Cost"] = Decimal.Round(Math.Ceiling(Solid713Count * Solid713Price / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                    Row["Weight"] = Decimal.Round(Solid713Weight, 3, MidpointRounding.AwayFromZero);
                    ReportDataTable.Rows.Add(Row);
                }

                //Cost = Math.Ceiling(Filenka713Count * Filenka713Price / 0.01m) * 0.01m;
                //Filenka713WithTransportCost = Math.Ceiling(Filenka713WithTransportCost / 0.01m) * 0.01m;
                if (Filenka713Count > 0)
                {
                    DataRow Row = ReportDataTable.NewRow();
                    Row["OriginalPrice"] = Filenka713Price;
                    Row["UNN"] = UNN;
                    Row["Cvet"] = GetColorCode(Convert.ToInt32(Fronts.Rows[i]["ColorID"]));
                    Row["Patina"] = GetPatinaCode(Convert.ToInt32(Fronts.Rows[i]["PatinaID"]));
                    Row["PaymentRate"] = PaymentRate;
                    Row["AccountingName"] = AccountingName;
                    Row["InvNumber"] = InvNumber;
                    Row["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        Row["TPSCurCode"] = TPSCurrencyCode;
                    Row["Count"] = Filenka713Count;
                    Row["PriceWithTransport"] = Decimal.Round(Filenka713WithTransportCost / Filenka713Count, 2, MidpointRounding.AwayFromZero);
                    Row["CostWithTransport"] = Decimal.Round(Math.Ceiling(Filenka713WithTransportCost / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                    Row["Cost"] = Decimal.Round(Math.Ceiling(Filenka713Count * Filenka713Price / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                    Row["Weight"] = Decimal.Round(Filenka713Weight, 3, MidpointRounding.AwayFromZero);
                    ReportDataTable.Rows.Add(Row);
                }

                //Cost = Math.Ceiling(NoInset713Count * NoInset713Price / 0.01m) * 0.01m;
                //NoInset713WithTransportCost = Math.Ceiling(NoInset713WithTransportCost / 0.01m) * 0.01m;
                if (NoInset713Count > 0)
                {
                    DataRow Row = ReportDataTable.NewRow();
                    Row["OriginalPrice"] = NoInset713Price;
                    Row["UNN"] = UNN;
                    Row["Cvet"] = GetColorCode(Convert.ToInt32(Fronts.Rows[i]["ColorID"]));
                    Row["Patina"] = GetPatinaCode(Convert.ToInt32(Fronts.Rows[i]["PatinaID"]));
                    Row["PaymentRate"] = PaymentRate;
                    Row["AccountingName"] = AccountingName;
                    Row["InvNumber"] = InvNumber;
                    Row["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        Row["TPSCurCode"] = TPSCurrencyCode;
                    Row["Count"] = NoInset713Count;
                    Row["PriceWithTransport"] = Decimal.Round(NoInset713WithTransportCost / NoInset713Count, 2, MidpointRounding.AwayFromZero);
                    Row["CostWithTransport"] = Decimal.Round(Math.Ceiling(NoInset713WithTransportCost / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                    Row["Cost"] = Decimal.Round(Math.Ceiling(NoInset713Count * NoInset713Price / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                    Row["Weight"] = Decimal.Round(NoInset713Weight, 3, MidpointRounding.AwayFromZero);
                    ReportDataTable.Rows.Add(Row);
                }

                //Cost = Math.Ceiling(Vitrina713Count * Vitrina713Price / 0.01m) * 0.01m;
                //Vitrina713WithTransportCost = Math.Ceiling(Vitrina713WithTransportCost / 0.01m) * 0.01m;
                if (Vitrina713Count > 0)
                {
                    DataRow Row = ReportDataTable.NewRow();
                    Row["OriginalPrice"] = Vitrina713Price;
                    Row["UNN"] = UNN;
                    Row["Cvet"] = GetColorCode(Convert.ToInt32(Fronts.Rows[i]["ColorID"]));
                    Row["Patina"] = GetPatinaCode(Convert.ToInt32(Fronts.Rows[i]["PatinaID"]));
                    Row["PaymentRate"] = PaymentRate;
                    Row["AccountingName"] = AccountingName;
                    Row["InvNumber"] = InvNumber;
                    Row["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        Row["TPSCurCode"] = TPSCurrencyCode;
                    Row["Count"] = Vitrina713Count;
                    Row["PriceWithTransport"] = Decimal.Round(Vitrina713WithTransportCost / Vitrina713Count, 2, MidpointRounding.AwayFromZero);
                    Row["CostWithTransport"] = Decimal.Round(Math.Ceiling(Vitrina713WithTransportCost / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                    Row["Cost"] = Decimal.Round(Math.Ceiling(Vitrina713Count * Vitrina713Price / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                    Row["Weight"] = Decimal.Round(Vitrina713Weight, 3, MidpointRounding.AwayFromZero);
                    ReportDataTable.Rows.Add(Row);
                }

                //Cost = Math.Ceiling(Solid910Count * Solid910Price / 0.01m) * 0.01m;
                //Solid910WithTransportCost = Math.Ceiling(Solid910WithTransportCost / 0.01m) * 0.01m;
                if (Solid910Count > 0)
                {
                    DataRow Row = ReportDataTable.NewRow();
                    Row["OriginalPrice"] = Solid910Price;
                    Row["UNN"] = UNN;
                    Row["Cvet"] = GetColorCode(Convert.ToInt32(Fronts.Rows[i]["ColorID"]));
                    Row["Patina"] = GetPatinaCode(Convert.ToInt32(Fronts.Rows[i]["PatinaID"]));
                    Row["PaymentRate"] = PaymentRate;
                    Row["InvNumber"] = InvNumber;
                    Row["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        Row["TPSCurCode"] = TPSCurrencyCode;
                    Row["Count"] = Solid910Count;
                    Row["PriceWithTransport"] = Decimal.Round(Solid910WithTransportCost / Solid910Count, 2, MidpointRounding.AwayFromZero);
                    Row["CostWithTransport"] = Decimal.Round(Math.Ceiling(Solid910WithTransportCost / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                    Row["Cost"] = Decimal.Round(Math.Ceiling(Solid910Count * Solid910Price / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                    Row["Weight"] = Decimal.Round(Solid910Weight, 3, MidpointRounding.AwayFromZero);
                    ReportDataTable.Rows.Add(Row);
                }

                //Cost = Math.Ceiling(Filenka910Count * Filenka910Price / 0.01m) * 0.01m;
                //Filenka910WithTransportCost = Math.Ceiling(Filenka910WithTransportCost / 0.01m) * 0.01m;
                if (Filenka910Count > 0)
                {
                    DataRow Row = ReportDataTable.NewRow();
                    Row["OriginalPrice"] = Filenka910Price;
                    Row["UNN"] = UNN;
                    Row["Cvet"] = GetColorCode(Convert.ToInt32(Fronts.Rows[i]["ColorID"]));
                    Row["Patina"] = GetPatinaCode(Convert.ToInt32(Fronts.Rows[i]["PatinaID"]));
                    Row["PaymentRate"] = PaymentRate;
                    Row["AccountingName"] = AccountingName;
                    Row["InvNumber"] = InvNumber;
                    Row["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        Row["TPSCurCode"] = TPSCurrencyCode;
                    Row["Count"] = Filenka910Count;
                    Row["PriceWithTransport"] = Decimal.Round(Filenka910WithTransportCost / Filenka910Count, 2, MidpointRounding.AwayFromZero);
                    Row["CostWithTransport"] = Decimal.Round(Math.Ceiling(Filenka910WithTransportCost / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                    Row["Cost"] = Decimal.Round(Math.Ceiling(Filenka910Count * Filenka910Price / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                    Row["Weight"] = Decimal.Round(Filenka910Weight, 3, MidpointRounding.AwayFromZero);
                    ReportDataTable.Rows.Add(Row);
                }

                //Cost = Math.Ceiling(NoInset910Count * NoInset910Price / 0.01m) * 0.01m;
                //NoInset910WithTransportCost = Math.Ceiling(NoInset910WithTransportCost / 0.01m) * 0.01m;
                if (NoInset910Count > 0)
                {
                    DataRow Row = ReportDataTable.NewRow();
                    Row["OriginalPrice"] = NoInset910Price;
                    Row["UNN"] = UNN;
                    Row["Cvet"] = GetColorCode(Convert.ToInt32(Fronts.Rows[i]["ColorID"]));
                    Row["Patina"] = GetPatinaCode(Convert.ToInt32(Fronts.Rows[i]["PatinaID"]));
                    Row["PaymentRate"] = PaymentRate;
                    Row["AccountingName"] = AccountingName;
                    Row["InvNumber"] = InvNumber;
                    Row["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        Row["TPSCurCode"] = TPSCurrencyCode;
                    Row["Count"] = NoInset910Count;
                    Row["PriceWithTransport"] = Decimal.Round(NoInset910WithTransportCost / NoInset910Count, 2, MidpointRounding.AwayFromZero);
                    Row["CostWithTransport"] = Decimal.Round(Math.Ceiling(NoInset910WithTransportCost / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                    Row["Cost"] = Decimal.Round(Math.Ceiling(NoInset910Count * NoInset910Price / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                    Row["Weight"] = Decimal.Round(NoInset910Weight, 3, MidpointRounding.AwayFromZero);
                    ReportDataTable.Rows.Add(Row);
                }

                //Cost = Math.Ceiling(Vitrina910Count * Vitrina910Price / 0.01m) * 0.01m;
                //Vitrina910WithTransportCost = Math.Ceiling(Vitrina910WithTransportCost / 0.01m) * 0.01m;
                if (Vitrina910Count > 0)
                {
                    DataRow Row = ReportDataTable.NewRow();
                    Row["OriginalPrice"] = Vitrina910Price;
                    Row["UNN"] = UNN;
                    Row["Cvet"] = GetColorCode(Convert.ToInt32(Fronts.Rows[i]["ColorID"]));
                    Row["Patina"] = GetPatinaCode(Convert.ToInt32(Fronts.Rows[i]["PatinaID"]));
                    Row["PaymentRate"] = PaymentRate;
                    Row["AccountingName"] = AccountingName;
                    Row["InvNumber"] = InvNumber;
                    Row["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        Row["TPSCurCode"] = TPSCurrencyCode;
                    Row["Count"] = Vitrina910Count;
                    Row["PriceWithTransport"] = Decimal.Round(Vitrina910WithTransportCost / Vitrina910Count, 2, MidpointRounding.AwayFromZero);
                    Row["CostWithTransport"] = Decimal.Round(Math.Ceiling(Vitrina910WithTransportCost / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                    Row["Cost"] = Decimal.Round(Math.Ceiling(Vitrina910Count * Vitrina910Price / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                    Row["Weight"] = Decimal.Round(Vitrina910Weight, 3, MidpointRounding.AwayFromZero);
                    ReportDataTable.Rows.Add(Row);
                }
            }

            Fronts.Dispose();
        }

        private void GetGrids(DataTable OrdersDataTable, DataTable ReportDataTable, DataTable ReportDataTable1, int FactoryID1)
        {
            if (OrdersDataTable.Select("InsetTypeID IN (685,686,687,688,29470,29471)").Count() == 0)
                return;

            int fID = Convert.ToInt32(OrdersDataTable.Rows[0]["FactoryID"]);

            DataTable Fronts = new DataTable();
            using (DataView DV = new DataView(OrdersDataTable))
            {
                DV.RowFilter = "InsetTypeID IN (685,686,687,688,29470,29471)";
                Fronts = DV.ToTable(true, new string[] { "FrontID", "ColorID", "PatinaID" });
            }

            for (int i = 0; i < Fronts.Rows.Count; i++)
            {
                int FrontID = Convert.ToInt32(Fronts.Rows[i]["FrontID"]);
                if (FrontID == 3729)
                    continue;

                DataRow[] Rows = OrdersDataTable.Select("ColorID = " + Fronts.Rows[i]["ColorID"].ToString() +
                    " AND PatinaID = " + Fronts.Rows[i]["PatinaID"].ToString() +
                    " AND FrontID = " + Fronts.Rows[i]["FrontID"].ToString());

                decimal CountPP = 0;
                decimal CostPP = 0;

                for (int j = 0; j < Rows.Count(); j++)
                {
                    decimal d = GetInsetSquare(Convert.ToInt32(Rows[j]["FrontID"]), Convert.ToInt32(Rows[j]["Height"]),
                        Convert.ToInt32(Rows[j]["Width"])) * Convert.ToDecimal(Rows[j]["Count"]);
                    CountPP += d;
                    CostPP += Math.Ceiling(Convert.ToDecimal(Rows[j]["InsetPrice"]) * d / 0.01m) * 0.01m;
                }

                if (CountPP > 0)
                {
                    DataTable ddt = OrdersDataTable.Select("ColorID = " + Fronts.Rows[i]["ColorID"].ToString() +
                                                           " AND PatinaID = " + Fronts.Rows[i]["PatinaID"].ToString() +
                                                           " AND FrontID = " + Fronts.Rows[i]["FrontID"].ToString()+
                                                           " and InsetTypeID IN (685,686,687,688,29470,29471)")
                        .CopyToDataTable();
                    for (int x = 0; x < ddt.Rows.Count; x++)
                    {
                        decimal d = GetInsetSquare(Convert.ToInt32(ddt.Rows[x]["FrontID"]), Convert.ToInt32(ddt.Rows[x]["Height"]),
                            Convert.ToInt32(ddt.Rows[x]["Width"])) * Convert.ToDecimal(ddt.Rows[x]["Count"]);
                        CountPP = d;
                        CostPP = Math.Ceiling(Convert.ToDecimal(ddt.Rows[x]["InsetPrice"]) * d / 0.01m) * 0.01m;

                        int FrontsOrdersID = Convert.ToInt32(ddt.Rows[x]["FrontsOrdersID"]);
                        DataRow[] rows = DecorInvNumbersDT.Select("NewFrontsOrdersID = " + FrontsOrdersID);

                        int FactoryID = 0;
                        if (rows.Count() > 0)
                            FactoryID = Convert.ToInt32(rows[0]["FactoryID"]);
                        if (FactoryID == 0)
                            FactoryID = FactoryID1;
                        if (FactoryID != FactoryID1)
                        {
                            //CostPP = Math.Ceiling(CostPP / 0.01m) * 0.01m;
                            DataRow NewRow = ReportDataTable1.NewRow();
                            NewRow["OriginalPrice"] = CostPP / CountPP;
                            NewRow["UNN"] = UNN;
                            NewRow["PaymentRate"] = PaymentRate;
                            NewRow["AccountingName"] = OrdersDataTable.Rows[0]["AccountingName"].ToString();
                            NewRow["InvNumber"] = OrdersDataTable.Rows[0]["InvNumber"].ToString();
                            NewRow["CurrencyCode"] = ProfilCurrencyCode;
                            if (fID == 2)
                                NewRow["TPSCurCode"] = TPSCurrencyCode;
                            NewRow["Count"] = Decimal.Round(CountPP, 3, MidpointRounding.AwayFromZero);
                            NewRow["Cost"] = Decimal.Round(Math.Ceiling(CostPP / 0.01m) * 0.01m, 2,
                                MidpointRounding.AwayFromZero);
                            NewRow["PriceWithTransport"] =
                                Decimal.Round(CostPP / CountPP, 2, MidpointRounding.AwayFromZero);
                            NewRow["CostWithTransport"] = Decimal.Round(Math.Ceiling(CostPP / 0.01m) * 0.01m, 2,
                                MidpointRounding.AwayFromZero);
                            NewRow["Weight"] =
                                Decimal.Round(Convert.ToDecimal(NewRow["Count"]) * Convert.ToDecimal(3.5), 3,
                                    MidpointRounding.AwayFromZero);
                            if (rows.Count() > 0)
                            {
                                NewRow["Cvet"] = rows[0]["Cvet"].ToString();
                                NewRow["Patina"] = rows[0]["Patina"].ToString();
                                NewRow["AccountingName"] = rows[0]["DecorAccountingName"].ToString();
                                NewRow["InvNumber"] = rows[0]["DecorInvNumber"].ToString();
                            }

                            ReportDataTable1.Rows.Add(NewRow);
                        }
                        else
                        {
                            //CostPP = Math.Ceiling(CostPP / 0.01m) * 0.01m;
                            DataRow NewRow = ReportDataTable.NewRow();
                            NewRow["OriginalPrice"] = CostPP / CountPP;
                            NewRow["UNN"] = UNN;
                            NewRow["PaymentRate"] = PaymentRate;
                            NewRow["AccountingName"] = OrdersDataTable.Rows[0]["AccountingName"].ToString();
                            NewRow["InvNumber"] = OrdersDataTable.Rows[0]["InvNumber"].ToString();
                            NewRow["CurrencyCode"] = ProfilCurrencyCode;
                            if (fID == 2)
                                NewRow["TPSCurCode"] = TPSCurrencyCode;
                            NewRow["Count"] = Decimal.Round(CountPP, 3, MidpointRounding.AwayFromZero);
                            NewRow["Cost"] = Decimal.Round(Math.Ceiling(CostPP / 0.01m) * 0.01m, 2,
                                MidpointRounding.AwayFromZero);
                            NewRow["PriceWithTransport"] =
                                Decimal.Round(CostPP / CountPP, 2, MidpointRounding.AwayFromZero);
                            NewRow["CostWithTransport"] = Decimal.Round(Math.Ceiling(CostPP / 0.01m) * 0.01m, 2,
                                MidpointRounding.AwayFromZero);
                            NewRow["Weight"] =
                                Decimal.Round(Convert.ToDecimal(NewRow["Count"]) * Convert.ToDecimal(3.5), 3,
                                    MidpointRounding.AwayFromZero);
                            if (rows.Count() > 0)
                            {
                                NewRow["AccountingName"] = rows[0]["DecorAccountingName"].ToString();
                                NewRow["InvNumber"] = rows[0]["DecorInvNumber"].ToString();
                                NewRow["Cvet"] = rows[0]["Cvet"].ToString();
                                NewRow["Patina"] = rows[0]["Patina"].ToString();
                            }

                            ReportDataTable.Rows.Add(NewRow);
                        }
                    }
                }
            }
        }

        private void GetGlass(DataTable OrdersDataTable, DataTable ReportDataTable, DataTable ReportDataTable1, int FactoryID1)
        {
            string AccountingName = OrdersDataTable.Rows[0]["AccountingName"].ToString();
            string InvNumber = OrdersDataTable.Rows[0]["InvNumber"].ToString();


            int fID = Convert.ToInt32(OrdersDataTable.Rows[0]["FactoryID"]);

            if (OrdersDataTable.Select("InsetTypeID = 2").Count() == 0)
                return;

            bool hasGlass = false;
            for (int i = 0; i < OrdersDataTable.Rows.Count; i++)
            {
                int frontID = Convert.ToInt32(OrdersDataTable.Rows[0]["FrontID"]);
                if (IsAluminium(OrdersDataTable.Rows[i]) > -1)
                {
                    hasGlass = true;
                    break;
                }
            }
            if (!hasGlass)
                return;

            DataTable Fronts = new DataTable();
            using (DataView DV = new DataView(OrdersDataTable))
            {
                DV.RowFilter = "InsetColorID = 3944";
                Fronts = DV.ToTable(true, new string[] { "FrontID", "ColorID", "PatinaID" });
            }

            for (int i = 0; i < Fronts.Rows.Count; i++)
            {
                DataRow[] FRows = OrdersDataTable.Select("ColorID = " + Fronts.Rows[i]["ColorID"].ToString() +
                    " AND PatinaID = " + Fronts.Rows[i]["PatinaID"].ToString() +
                    " AND FrontID = " + Fronts.Rows[i]["FrontID"].ToString());

                if (FRows.Count() > 0)
                {
                    decimal PriceFlutes = 0;
                    decimal CountFlutes = 0;
                    PriceFlutes = Convert.ToDecimal(FRows[0]["InsetPrice"]);

                    for (int j = 0; j < FRows.Count(); j++)
                    {
                        if (IsAluminium(FRows[j]) != -1)
                            continue;

                        CountFlutes += Convert.ToDecimal(FRows[j]["Count"]) *
                                    GetInsetSquare(Convert.ToInt32(FRows[j]["FrontID"]),
                                                  Convert.ToInt32(FRows[j]["Height"]),
                                                  Convert.ToInt32(FRows[j]["Width"]));
                    }

                    CountFlutes = Decimal.Round(CountFlutes, 3, MidpointRounding.AwayFromZero);

                    if (CountFlutes > 0)
                    {
                        //decimal Cost = Math.Ceiling(CountFlutes * PriceFlutes / 0.01m) * 0.01m;
                        DataRow[] rows = DecorInvNumbersDT.Select("NewFrontsOrdersID = " + Convert.ToInt32(FRows[0]["FrontsOrdersID"]));
                        int FactoryID = 0;
                        if (rows.Count() > 0)
                            FactoryID = Convert.ToInt32(rows[0]["FactoryID"]);
                        if (FactoryID == 0)
                            FactoryID = FactoryID1;
                        if (FactoryID != FactoryID1)
                        {
                            DataRow NewRow = ReportDataTable1.NewRow();
                            //NewRow["Name"] = "Стекло Флутес";
                            NewRow["OriginalPrice"] = PriceFlutes;
                            NewRow["UNN"] = UNN;
                            NewRow["PaymentRate"] = PaymentRate;
                            NewRow["AccountingName"] = AccountingName;
                            NewRow["InvNumber"] = InvNumber;
                            NewRow["CurrencyCode"] = ProfilCurrencyCode;
                            if (fID == 2)
                                NewRow["TPSCurCode"] = TPSCurrencyCode;
                            if (rows.Count() > 0)
                            {
                                NewRow["AccountingName"] = rows[0]["DecorAccountingName"].ToString();
                                NewRow["InvNumber"] = rows[0]["DecorInvNumber"].ToString();
                                NewRow["Cvet"] = rows[0]["Cvet"].ToString();
                                NewRow["Patina"] = rows[0]["Patina"].ToString();
                            }
                            NewRow["Count"] = CountFlutes;
                            NewRow["Cost"] = Decimal.Round(Math.Ceiling(CountFlutes * PriceFlutes / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                            NewRow["PriceWithTransport"] = Decimal.Round(PriceFlutes, 2, MidpointRounding.AwayFromZero);
                            NewRow["CostWithTransport"] = Decimal.Round(Math.Ceiling(CountFlutes * PriceFlutes / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                            NewRow["Weight"] = Decimal.Round(CountFlutes * 10, 3, MidpointRounding.AwayFromZero);
                            ReportDataTable1.Rows.Add(NewRow);
                        }
                        else
                        {
                            DataRow NewRow = ReportDataTable.NewRow();
                            //NewRow["Name"] = "Стекло Флутес";
                            NewRow["OriginalPrice"] = PriceFlutes;
                            NewRow["UNN"] = UNN;
                            NewRow["PaymentRate"] = PaymentRate;
                            NewRow["AccountingName"] = AccountingName;
                            NewRow["InvNumber"] = InvNumber;
                            if (rows.Count() > 0)
                            {
                                NewRow["AccountingName"] = rows[0]["DecorAccountingName"].ToString();
                                NewRow["InvNumber"] = rows[0]["DecorInvNumber"].ToString();
                                NewRow["Cvet"] = rows[0]["Cvet"].ToString();
                                NewRow["Patina"] = rows[0]["Patina"].ToString();
                            }
                            NewRow["CurrencyCode"] = ProfilCurrencyCode;
                            if (fID == 2)
                                NewRow["TPSCurCode"] = TPSCurrencyCode;
                            NewRow["Count"] = CountFlutes;
                            NewRow["Cost"] = Decimal.Round(Math.Ceiling(CountFlutes * PriceFlutes / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                            NewRow["PriceWithTransport"] = Decimal.Round(PriceFlutes, 2, MidpointRounding.AwayFromZero);
                            NewRow["CostWithTransport"] = Decimal.Round(Math.Ceiling(CountFlutes * PriceFlutes / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                            NewRow["Weight"] = Decimal.Round(CountFlutes * 10, 3, MidpointRounding.AwayFromZero);
                            ReportDataTable.Rows.Add(NewRow);
                        }
                    }
                }
            }

            Fronts.Clear();
            using (DataView DV = new DataView(OrdersDataTable))
            {
                DV.RowFilter = "InsetColorID = 3943";
                Fronts = DV.ToTable(true, new string[] { "FrontID", "ColorID", "PatinaID" });
            }

            for (int i = 0; i < Fronts.Rows.Count; i++)
            {
                DataRow[] LRows = OrdersDataTable.Select("ColorID = " + Fronts.Rows[i]["ColorID"].ToString() +
                    " AND PatinaID = " + Fronts.Rows[i]["PatinaID"].ToString() +
                    " AND FrontID = " + Fronts.Rows[i]["FrontID"].ToString());

                if (LRows.Count() > 0)
                {
                    decimal CountLacomat = 0;
                    decimal PriceLacomat = 0;
                    PriceLacomat = Convert.ToDecimal(LRows[0]["InsetPrice"]);

                    for (int j = 0; j < LRows.Count(); j++)
                    {
                        if (IsAluminium(LRows[j]) != -1)
                            continue;

                        CountLacomat += Convert.ToDecimal(LRows[j]["Count"]) *
                                    GetInsetSquare(Convert.ToInt32(LRows[j]["FrontID"]),
                                                  Convert.ToInt32(LRows[j]["Height"]),
                                                  Convert.ToInt32(LRows[j]["Width"]));

                    }

                    CountLacomat = Decimal.Round(CountLacomat, 3, MidpointRounding.AwayFromZero);

                    if (CountLacomat > 0)
                    {
                        //decimal Cost = Math.Ceiling(CountLacomat * PriceLacomat / 0.01m) * 0.01m;
                        DataRow[] rows = DecorInvNumbersDT.Select("NewFrontsOrdersID = " + Convert.ToInt32(LRows[0]["FrontsOrdersID"]));
                        int FactoryID = 0;
                        if (rows.Count() > 0)
                            FactoryID = Convert.ToInt32(rows[0]["FactoryID"]);
                        if (FactoryID == 0)
                            FactoryID = FactoryID1;
                        if (FactoryID != FactoryID1)
                        {
                            DataRow NewRow = ReportDataTable1.NewRow();
                            //NewRow["Name"] = "Стекло Лакомат";
                            NewRow["OriginalPrice"] = PriceLacomat;
                            NewRow["UNN"] = UNN;
                            NewRow["PaymentRate"] = PaymentRate;
                            NewRow["AccountingName"] = AccountingName;
                            NewRow["InvNumber"] = InvNumber;
                            if (rows.Count() > 0)
                            {
                                NewRow["AccountingName"] = rows[0]["DecorAccountingName"].ToString();
                                NewRow["InvNumber"] = rows[0]["DecorInvNumber"].ToString();
                                NewRow["Cvet"] = rows[0]["Cvet"].ToString();
                                NewRow["Patina"] = rows[0]["Patina"].ToString();
                            }
                            NewRow["CurrencyCode"] = ProfilCurrencyCode;
                            if (fID == 2)
                                NewRow["TPSCurCode"] = TPSCurrencyCode;
                            NewRow["Count"] = CountLacomat;
                            NewRow["Cost"] = Decimal.Round(Math.Ceiling(CountLacomat * PriceLacomat / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                            NewRow["PriceWithTransport"] = Decimal.Round(PriceLacomat, 2, MidpointRounding.AwayFromZero);
                            NewRow["CostWithTransport"] = Decimal.Round(Math.Ceiling(CountLacomat * PriceLacomat / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                            NewRow["Weight"] = Decimal.Round(CountLacomat * 10, 3, MidpointRounding.AwayFromZero);
                            ReportDataTable1.Rows.Add(NewRow);
                        }
                        else
                        {
                            DataRow NewRow = ReportDataTable.NewRow();
                            //NewRow["Name"] = "Стекло Лакомат";
                            NewRow["OriginalPrice"] = PriceLacomat;
                            NewRow["UNN"] = UNN;
                            NewRow["PaymentRate"] = PaymentRate;
                            NewRow["AccountingName"] = AccountingName;
                            NewRow["InvNumber"] = InvNumber;
                            if (rows.Count() > 0)
                            {
                                NewRow["AccountingName"] = rows[0]["DecorAccountingName"].ToString();
                                NewRow["InvNumber"] = rows[0]["DecorInvNumber"].ToString();
                                NewRow["Cvet"] = rows[0]["Cvet"].ToString();
                                NewRow["Patina"] = rows[0]["Patina"].ToString();
                            }
                            NewRow["CurrencyCode"] = ProfilCurrencyCode;
                            if (fID == 2)
                                NewRow["TPSCurCode"] = TPSCurrencyCode;
                            NewRow["Count"] = CountLacomat;
                            NewRow["Cost"] = Decimal.Round(Math.Ceiling(CountLacomat * PriceLacomat / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                            NewRow["PriceWithTransport"] = Decimal.Round(PriceLacomat, 2, MidpointRounding.AwayFromZero);
                            NewRow["CostWithTransport"] = Decimal.Round(Math.Ceiling(CountLacomat * PriceLacomat / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                            NewRow["Weight"] = Decimal.Round(CountLacomat * 10, 3, MidpointRounding.AwayFromZero);
                            ReportDataTable.Rows.Add(NewRow);
                        }
                    }
                }
            }
            
            Fronts.Clear();
            using (DataView DV = new DataView(OrdersDataTable))
            {
                DV.RowFilter = "InsetColorID = 3945";
                Fronts = DV.ToTable(true, new string[] { "FrontID", "ColorID", "PatinaID" });
            }

            for (int i = 0; i < Fronts.Rows.Count; i++)
            {
                DataRow[] KRows = OrdersDataTable.Select("ColorID = " + Fronts.Rows[i]["ColorID"].ToString() +
                    " AND PatinaID = " + Fronts.Rows[i]["PatinaID"].ToString() +
                    " AND FrontID = " + Fronts.Rows[i]["FrontID"].ToString());

                if (KRows.Count() > 0)
                {
                    decimal PriceKrizet = 0;
                    decimal CountKrizet = 0;
                    PriceKrizet = Convert.ToDecimal(KRows[0]["InsetPrice"]);

                    for (int j = 0; j < KRows.Count(); j++)
                    {
                        if (IsAluminium(KRows[j]) != -1)
                            continue;

                        CountKrizet += Convert.ToDecimal(KRows[j]["Count"]) *
                                    GetInsetSquare(Convert.ToInt32(KRows[j]["FrontID"]),
                                                  Convert.ToInt32(KRows[j]["Height"]),
                                                  Convert.ToInt32(KRows[j]["Width"]));
                    }

                    CountKrizet = Decimal.Round(CountKrizet, 3, MidpointRounding.AwayFromZero);

                    if (CountKrizet > 0)
                    {
                        //decimal Cost = Math.Ceiling(CountKrizet * PriceKrizet / 0.01m) * 0.01m;
                        DataRow[] rows = DecorInvNumbersDT.Select("NewFrontsOrdersID = " + Convert.ToInt32(KRows[0]["FrontsOrdersID"]));
                        int FactoryID = 0;
                        if (rows.Count() > 0)
                            FactoryID = Convert.ToInt32(rows[0]["FactoryID"]);
                        if (FactoryID == 0)
                            FactoryID = FactoryID1;
                        if (FactoryID != FactoryID1)
                        {
                            DataRow NewRow = ReportDataTable1.NewRow();
                            //NewRow["Name"] = "Стекло Кризет";
                            NewRow["OriginalPrice"] = PriceKrizet;
                            NewRow["UNN"] = UNN;
                            NewRow["PaymentRate"] = PaymentRate;
                            NewRow["AccountingName"] = AccountingName;
                            NewRow["InvNumber"] = InvNumber;
                            if (rows.Count() > 0)
                            {
                                NewRow["AccountingName"] = rows[0]["DecorAccountingName"].ToString();
                                NewRow["InvNumber"] = rows[0]["DecorInvNumber"].ToString();
                                NewRow["Cvet"] = rows[0]["Cvet"].ToString();
                                NewRow["Patina"] = rows[0]["Patina"].ToString();
                            }
                            NewRow["CurrencyCode"] = ProfilCurrencyCode;
                            if (fID == 2)
                                NewRow["TPSCurCode"] = TPSCurrencyCode;
                            NewRow["Count"] = CountKrizet;
                            NewRow["Cost"] = Decimal.Round(Math.Ceiling(CountKrizet * PriceKrizet / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                            NewRow["PriceWithTransport"] = Decimal.Round(PriceKrizet, 2, MidpointRounding.AwayFromZero);
                            NewRow["CostWithTransport"] = Decimal.Round(Math.Ceiling(CountKrizet * PriceKrizet / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                            NewRow["Weight"] = Decimal.Round(CountKrizet * 10, 3, MidpointRounding.AwayFromZero);
                            ReportDataTable1.Rows.Add(NewRow);
                        }
                        else
                        {
                            DataRow NewRow = ReportDataTable.NewRow();
                            //NewRow["Name"] = "Стекло Кризет";
                            NewRow["OriginalPrice"] = PriceKrizet;
                            NewRow["UNN"] = UNN;
                            NewRow["PaymentRate"] = PaymentRate;
                            NewRow["AccountingName"] = AccountingName;
                            NewRow["InvNumber"] = InvNumber;
                            if (rows.Count() > 0)
                            {
                                NewRow["AccountingName"] = rows[0]["DecorAccountingName"].ToString();
                                NewRow["InvNumber"] = rows[0]["DecorInvNumber"].ToString();
                                NewRow["Cvet"] = rows[0]["Cvet"].ToString();
                                NewRow["Patina"] = rows[0]["Patina"].ToString();
                            }
                            NewRow["CurrencyCode"] = ProfilCurrencyCode;
                            if (fID == 2)
                                NewRow["TPSCurCode"] = TPSCurrencyCode;
                            NewRow["Count"] = CountKrizet;
                            NewRow["Cost"] = Decimal.Round(Math.Ceiling(CountKrizet * PriceKrizet / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                            NewRow["PriceWithTransport"] = Decimal.Round(PriceKrizet, 2, MidpointRounding.AwayFromZero);
                            NewRow["CostWithTransport"] = Decimal.Round(Math.Ceiling(CountKrizet * PriceKrizet / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                            NewRow["Weight"] = Decimal.Round(CountKrizet * 10, 3, MidpointRounding.AwayFromZero);
                            ReportDataTable.Rows.Add(NewRow);
                        }
                    }
                }
            }

            Fronts.Clear();
            using (DataView DV = new DataView(OrdersDataTable))
            {
                DV.RowFilter = "InsetTypeID = 18";
                Fronts = DV.ToTable(true, new string[] { "FrontID", "ColorID", "PatinaID" });
            }

            for (int i = 0; i < Fronts.Rows.Count; i++)
            {
                DataRow[] ORows = OrdersDataTable.Select("ColorID = " + Fronts.Rows[i]["ColorID"].ToString() +
                    " AND PatinaID = " + Fronts.Rows[i]["PatinaID"].ToString() +
                    " AND FrontID = " + Fronts.Rows[i]["FrontID"].ToString());

                if (ORows.Count() > 0)
                {
                    decimal CountOther = 0;
                    for (int j = 0; j < ORows.Count(); j++)
                    {
                        if (IsAluminium(ORows[j]) != -1)
                            continue;

                        CountOther += Convert.ToDecimal(ORows[j]["Count"]) *
                                    GetInsetSquare(Convert.ToInt32(ORows[j]["FrontID"]),
                                                  Convert.ToInt32(ORows[j]["Height"]),
                                                  Convert.ToInt32(ORows[j]["Width"]));
                    }

                    CountOther = Decimal.Round(CountOther, 3, MidpointRounding.AwayFromZero);

                    if (CountOther > 0)
                    {
                        DataRow[] rows = DecorInvNumbersDT.Select("NewFrontsOrdersID = " + Convert.ToInt32(ORows[0]["FrontsOrdersID"]));
                        int FactoryID = 0;
                        if (rows.Count() > 0)
                            FactoryID = Convert.ToInt32(rows[0]["FactoryID"]);
                        if (FactoryID == 0)
                            FactoryID = FactoryID1;
                        if (FactoryID != FactoryID1)
                        {
                            DataRow NewRow = ReportDataTable1.NewRow();
                            //NewRow["Name"] = "Стекло другое";
                            NewRow["OriginalPrice"] = 0;
                            NewRow["UNN"] = UNN;
                            NewRow["PaymentRate"] = PaymentRate;
                            NewRow["AccountingName"] = AccountingName;
                            NewRow["InvNumber"] = InvNumber;
                            if (rows.Count() > 0)
                            {
                                NewRow["AccountingName"] = rows[0]["DecorAccountingName"].ToString();
                                NewRow["InvNumber"] = rows[0]["DecorInvNumber"].ToString();
                                NewRow["Cvet"] = rows[0]["Cvet"].ToString();
                                NewRow["Patina"] = rows[0]["Patina"].ToString();
                            }
                            NewRow["CurrencyCode"] = ProfilCurrencyCode;
                            if (fID == 2)
                                NewRow["TPSCurCode"] = TPSCurrencyCode;
                            NewRow["Count"] = CountOther;
                            NewRow["Cost"] = 0;
                            NewRow["PriceWithTransport"] = 0;
                            NewRow["CostWithTransport"] = 0;
                            NewRow["Weight"] = Decimal.Round(CountOther * 10, 3, MidpointRounding.AwayFromZero);
                            ReportDataTable1.Rows.Add(NewRow);
                        }
                        else
                        {
                            DataRow NewRow = ReportDataTable.NewRow();
                            //NewRow["Name"] = "Стекло другое";
                            NewRow["OriginalPrice"] = 0;
                            NewRow["UNN"] = UNN;
                            NewRow["PaymentRate"] = PaymentRate;
                            NewRow["AccountingName"] = AccountingName;
                            NewRow["InvNumber"] = InvNumber;
                            if (rows.Count() > 0)
                            {
                                NewRow["AccountingName"] = rows[0]["DecorAccountingName"].ToString();
                                NewRow["InvNumber"] = rows[0]["DecorInvNumber"].ToString();
                                NewRow["Cvet"] = rows[0]["Cvet"].ToString();
                                NewRow["Patina"] = rows[0]["Patina"].ToString();
                            }
                            NewRow["CurrencyCode"] = ProfilCurrencyCode;
                            if (fID == 2)
                                NewRow["TPSCurCode"] = TPSCurrencyCode;
                            NewRow["Count"] = CountOther;
                            NewRow["Cost"] = 0;
                            NewRow["PriceWithTransport"] = 0;
                            NewRow["CostWithTransport"] = 0;
                            NewRow["Weight"] = Decimal.Round(CountOther * 10, 3, MidpointRounding.AwayFromZero);
                            ReportDataTable.Rows.Add(NewRow);
                        }
                    }
                }
            }
        }

        private void GetInsets(DataTable OrdersDataTable, DataTable ReportDataTable)
        {
            int fID = Convert.ToInt32(OrdersDataTable.Rows[0]["FactoryID"]);

            DataTable Fronts = new DataTable();
            using (DataView DV = new DataView(OrdersDataTable))
            {
                DV.RowFilter = "FrontID IN (3729)";
                Fronts = DV.ToTable(true, new string[] { "FrontID", "ColorID", "PatinaID" });
            }

            for (int i = 0; i < Fronts.Rows.Count; i++)
            {
                //ellipse grid
                DataRow[] EGRows = OrdersDataTable.Select("ColorID = " + Fronts.Rows[i]["ColorID"].ToString() +
                    " AND PatinaID = " + Fronts.Rows[i]["PatinaID"].ToString() +
                    " AND FrontID = " + Fronts.Rows[i]["FrontID"].ToString());

                if (EGRows.Count() > 0)
                {
                    int MarginHeight = 0;
                    int MarginWidth = 0;

                    decimal CountEllipseGrid = 0;
                    decimal PriceEllipseGrid = 0;

                    GetGlassMarginAluminium(EGRows[0], ref MarginHeight, ref MarginWidth);
                    PriceEllipseGrid = Convert.ToDecimal(EGRows[0]["InsetPrice"]);

                    for (int j = 0; j < EGRows.Count(); j++)
                    {
                        decimal dd = Decimal.Round(Convert.ToDecimal(EGRows[j]["Count"]) * MarginHeight * (Convert.ToDecimal(EGRows[j]["Width"]) - MarginWidth) / 1000000, 3, MidpointRounding.AwayFromZero);
                        CountEllipseGrid += dd;
                    }

                    decimal Weight = GetInsetWeight(EGRows[0]);
                    Weight = Decimal.Round(CountEllipseGrid * Weight, 3, MidpointRounding.AwayFromZero);

                    //decimal Cost = Math.Ceiling(CountEllipseGrid * PriceEllipseGrid / 0.01m) * 0.01m;
                    DataRow NewRow = ReportDataTable.NewRow();
                    NewRow["OriginalPrice"] = PriceEllipseGrid;
                    NewRow["UNN"] = UNN;
                    NewRow["PaymentRate"] = PaymentRate;
                    NewRow["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        NewRow["TPSCurCode"] = TPSCurrencyCode;
                    NewRow["Count"] = CountEllipseGrid;
                    NewRow["Cost"] = Decimal.Round(Math.Ceiling(CountEllipseGrid * PriceEllipseGrid / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                    NewRow["PriceWithTransport"] = PriceEllipseGrid;
                    NewRow["CostWithTransport"] = Decimal.Round(Math.Ceiling(CountEllipseGrid * PriceEllipseGrid / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                    NewRow["Weight"] = 0;
                    DataRow[] rows = DecorInvNumbersDT.Select("NewFrontsOrdersID = " + Convert.ToInt32(EGRows[0]["FrontsOrdersID"]));
                    if (rows.Count() > 0)
                    {
                        NewRow["AccountingName"] = rows[0]["DecorAccountingName"].ToString();
                        NewRow["InvNumber"] = rows[0]["DecorInvNumber"].ToString();
                        NewRow["Cvet"] = rows[0]["Cvet"].ToString();
                        NewRow["Patina"] = rows[0]["Patina"].ToString();
                    }
                    ReportDataTable.Rows.Add(NewRow);

                }
            }
        }

        private decimal GetInsetWeight(DataRow NewFrontsOrdersRow)
        {
            int InsetTypeID = Convert.ToInt32(NewFrontsOrdersRow["InsetTypeID"]);
            if (InsetTypeID == 2)
                InsetTypeID = Convert.ToInt32(NewFrontsOrdersRow["InsetColorID"]);//стекло
            decimal InsetSquare = GetInsetSquare(NewFrontsOrdersRow);
            if (InsetSquare <= 0)
                return 0;
            decimal InsetWeight = 0;
            DataRow[] Rows = TechStoreDataTable.Select("TechStoreID = " + InsetTypeID);
            if (Rows.Count() > 0)
            {
                if (Rows[0]["Weight"] != DBNull.Value)
                    InsetWeight = Convert.ToDecimal(Rows[0]["Weight"]);
            }

            return InsetSquare * InsetWeight;
        }

        private decimal GetProfileWeight(DataRow NewFrontsOrdersRow)
        {
            decimal FrontHeight = Convert.ToDecimal(NewFrontsOrdersRow["Height"]);
            decimal FrontWidth = Convert.ToDecimal(NewFrontsOrdersRow["Width"]);
            decimal ProfileWeight = 0;
            decimal ProfileWidth = 0;
            DataRow[] FrontsConfigRow = FrontsConfigDataTable.Select("FrontConfigID = " + NewFrontsOrdersRow["FrontConfigID"].ToString());
            if (FrontsConfigRow.Count() > 0)
                ProfileWeight = Convert.ToDecimal(FrontsConfigRow[0]["Weight"]);

            int FrontID = Convert.ToInt32(NewFrontsOrdersRow["FrontID"]);
            if (Security.IsFrontsSquareCalc(FrontID))
            {
                return FrontWidth * FrontHeight / 1000000 * ProfileWeight;
            }
            //if (FrontID == 30504 || FrontID == 30505 || FrontID == 30506 ||
            //    FrontID == 30364 || FrontID == 30366 || FrontID == 30367 ||
            //    FrontID == 30501 || FrontID == 30502 || FrontID == 30503 ||
            //    FrontID == 16269 || FrontID == 28945 || FrontID == 41327 || FrontID == 41328 || FrontID == 41331 || 
            //    FrontID == 27914 || FrontID == 29597 || FrontID == 3727 || FrontID == 3728 || FrontID == 3729 ||
            //    FrontID == 3730 || FrontID == 3731 || FrontID == 3732 || FrontID == 3733 || FrontID == 3734 ||
            //    FrontID == 3735 || FrontID == 3736 || FrontID == 3737 || FrontID == 3739 || FrontID == 3740 ||
            //    FrontID == 3741 || FrontID == 3742 || FrontID == 3743 || FrontID == 3744 || FrontID == 3745 ||
            //    FrontID == 3746 || FrontID == 3747 || FrontID == 3748 || FrontID == 15108 || FrontID == 3662 || FrontID == 3663 || FrontID == 3664 || FrontID == 15760)
            //    return FrontWidth * FrontHeight / 1000000 * ProfileWeight;
            else
            {
                DataRow[] DecorConfigRow = TechStoreDataTable.Select("TechStoreID = " + FrontsConfigRow[0]["ProfileID"].ToString());
                if (DecorConfigRow.Count() > 0)
                {
                    ProfileWidth = Convert.ToDecimal(DecorConfigRow[0]["Width"]);
                    ProfileWeight = Convert.ToDecimal(DecorConfigRow[0]["Weight"]);
                    return (FrontWidth * 2 + (FrontHeight - ProfileWidth - ProfileWidth) * 2) / 1000 * ProfileWeight;
                }
            }
            return 0;
        }

        private void GetFrontWeight(DataRow NewFrontsOrdersRow, ref decimal outFrontWeight, ref decimal outInsetWeight)
        {
            //decimal FrontsWeight = 0;
            DataRow[] FrontsConfigRow = FrontsConfigDataTable.Select("FrontConfigID = " + NewFrontsOrdersRow["FrontConfigID"].ToString());
            decimal InsetWeight = Convert.ToDecimal(FrontsConfigRow[0]["InsetWeight"]);
            decimal NewFrontsOrdersquare = Convert.ToDecimal(NewFrontsOrdersRow["Square"]);
            decimal PackWeight = 0;
            if (NewFrontsOrdersquare > 0)
                PackWeight = NewFrontsOrdersquare * Convert.ToDecimal(0.7);
            //если гнутый то вес за штуки
            if (FrontsConfigRow[0]["Width"].ToString() == "-1")
            {
                outFrontWeight = PackWeight +
                    Convert.ToDecimal(NewFrontsOrdersRow["Count"]) * Convert.ToDecimal(FrontsConfigRow[0]["Weight"]);
                return;
            }
            //если алюминий
            if (IsAluminium(NewFrontsOrdersRow) > -1)
            {
                outFrontWeight = PackWeight +
                    GetAluminiumWeight(NewFrontsOrdersRow, true);
                return;
            }
            decimal ResultProfileWeight = GetProfileWeight(NewFrontsOrdersRow);
            decimal ResultInsetWeight = 0;
            DataRow[] rows = InsetTypesDataTable.Select("GroupID = 2 OR GroupID = 3 OR GroupID = 4 OR GroupID = 7 OR GroupID = 12 OR GroupID = 13");
            foreach (DataRow item in rows)
            {
                if (NewFrontsOrdersRow["InsetTypeID"].ToString() == item["InsetTypeID"].ToString())
                {
                    ResultInsetWeight = GetInsetWeight(NewFrontsOrdersRow);
                }
            }
            if (Convert.ToInt32(NewFrontsOrdersRow["FrontID"]) == 3729)
                ResultInsetWeight = GetInsetWeight(NewFrontsOrdersRow);

            outFrontWeight = PackWeight + ResultProfileWeight * Convert.ToDecimal(NewFrontsOrdersRow["Count"]);
            outInsetWeight = ResultInsetWeight * Convert.ToDecimal(NewFrontsOrdersRow["Count"]);
        }

        private DataTable[] GroupInvNumber(DataTable OrdersDataTable)
        {
            int InvCount = 0;

            //get count of different covertypes
            using (DataView DV = new DataView(OrdersDataTable))
            {
                InvCount = DV.ToTable(true, new string[] { "InvNumber" }).Rows.Count;
            }

            //create DataTables
            DataTable[] InvDataTables = new DataTable[InvCount];

            //split for covertypes
            int g = 0;

            for (int i = 0; i < InvCount; i++)
            {
                InvDataTables[i] = OrdersDataTable.Clone();

                int ColorID = -1;
                int PatinaID = -1;
                string InvNumber = "";

                for (int r = g; r < OrdersDataTable.Rows.Count; r++)
                {
                    if (InvNumber == "")
                    {
                        InvDataTables[i].ImportRow(OrdersDataTable.DefaultView[r].Row);
                        InvNumber = OrdersDataTable.DefaultView[r].Row["InvNumber"].ToString();
                        ColorID = Convert.ToInt32(OrdersDataTable.DefaultView[r].Row["ColorID"]);
                        PatinaID = Convert.ToInt32(OrdersDataTable.DefaultView[r].Row["PatinaID"]);
                        continue;
                    }

                    if (InvNumber == OrdersDataTable.DefaultView[r].Row["InvNumber"].ToString())
                    {
                        InvDataTables[i].ImportRow(OrdersDataTable.DefaultView[r].Row);
                    }
                    else
                    {
                        g = r;
                        break;
                    }
                }
            }

            return InvDataTables;
        }

        public void Report(int[] MainOrderIDs, bool IsSample)
        {
            ProfilFrontsOrdersDataTable.Clear();
            TPSFrontsOrdersDataTable.Clear();
            DecorInvNumbersDT.Clear();
            GetMegaOrderInfo(MainOrderIDs[0]);

            string sWhere = "";

            for (int i = 0; i < MainOrderIDs.Count(); i++)
            {
                if (sWhere != "")
                    sWhere += " OR MainOrderID = " + MainOrderIDs[i].ToString();
                else
                    sWhere += "MainOrderID = " + MainOrderIDs[i].ToString();
            }
            string SelectCommand = "SELECT NewFrontsOrders.*, infiniu2_catalog.dbo.FrontsConfig.measureId, infiniu2_catalog.dbo.FrontsConfig.AccountingName, infiniu2_catalog.dbo.FrontsConfig.InvNumber, NewMegaOrders.PaymentRate FROM NewFrontsOrders" +
                " INNER JOIN infiniu2_catalog.dbo.FrontsConfig ON NewFrontsOrders.FrontConfigID = infiniu2_catalog.dbo.FrontsConfig.FrontConfigID" +
                " INNER JOIN NewMainOrders ON NewFrontsOrders.MainOrderID = NewMainOrders.MainOrderID" +
                " INNER JOIN NewMegaOrders ON NewMainOrders.MegaOrderID = NewMegaOrders.MegaOrderID" +
                " WHERE NewFrontsOrders.IsSample=1 AND InvNumber IS NOT NULL AND NewFrontsOrders.MainOrderID IN (" + string.Join(",", MainOrderIDs) + ") ORDER BY InvNumber";
            if (!IsSample)
                SelectCommand = "SELECT NewFrontsOrders.*, infiniu2_catalog.dbo.FrontsConfig.measureId, infiniu2_catalog.dbo.FrontsConfig.AccountingName, infiniu2_catalog.dbo.FrontsConfig.InvNumber, NewMegaOrders.PaymentRate FROM NewFrontsOrders" +
                " INNER JOIN infiniu2_catalog.dbo.FrontsConfig ON NewFrontsOrders.FrontConfigID = infiniu2_catalog.dbo.FrontsConfig.FrontConfigID" +
                " INNER JOIN NewMainOrders ON NewFrontsOrders.MainOrderID = NewMainOrders.MainOrderID" +
                " INNER JOIN NewMegaOrders ON NewMainOrders.MegaOrderID = NewMegaOrders.MegaOrderID" +
                " WHERE NewFrontsOrders.IsSample=0 AND InvNumber IS NOT NULL AND NewFrontsOrders.MainOrderID IN (" + string.Join(",", MainOrderIDs) + ") ORDER BY InvNumber";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    if (DA.Fill(DT) == 0)
                        return;

                    DataTable DistRatesDT = new DataTable();
                    DataTable DT1 = DT.Clone();
                    //get count of different covertypes
                    using (DataView DV = new DataView(DT))
                    {
                        DistRatesDT = DV.ToTable(true, new string[] { "PaymentRate" });
                    }

                    ProfilFrontsOrdersDataTable = DT.Clone();
                    TPSFrontsOrdersDataTable = DT.Clone();

                    SplitTables(DT, ref ProfilFrontsOrdersDataTable, ref TPSFrontsOrdersDataTable);

                    if (ProfilFrontsOrdersDataTable.Rows.Count > 0)
                    {
                        DataTable[] DTs = GroupInvNumber(ProfilFrontsOrdersDataTable);

                        for (int i = 0; i < DTs.Count(); i++)
                        {
                            for (int j = 0; j < DistRatesDT.Rows.Count; j++)
                            {
                                DT1.Clear();
                                PaymentRate = Convert.ToDecimal(DistRatesDT.Rows[j]["PaymentRate"]);
                                DataRow[] rows = DTs[i].Select("PaymentRate='" + PaymentRate.ToString() + "'");
                                foreach (DataRow item in rows)
                                    DT1.Rows.Add(item.ItemArray);
                                if (DT1.Rows.Count == 0)
                                    continue;

                                GetSimpleFronts(DT1, _profilReportDataTable, false);
                                GetSimpleFronts(DT1, _profilReportDataTable, true);

                                GetCurvedFronts(DT1, _profilReportDataTable);

                                GetGrids(DT1, _profilReportDataTable, _tPSReportDataTable, 1);

                                GetInsets(DT1, _profilReportDataTable);

                                GetGlass(DT1, _profilReportDataTable, _tPSReportDataTable, 1);
                            }
                        }
                    }

                    DT1.Clear();
                    if (TPSFrontsOrdersDataTable.Rows.Count > 0)
                    {
                        DataTable[] DTs = GroupInvNumber(TPSFrontsOrdersDataTable);

                        for (int i = 0; i < DTs.Count(); i++)
                        {
                            for (int j = 0; j < DistRatesDT.Rows.Count; j++)
                            {
                                DT1.Clear();
                                PaymentRate = Convert.ToDecimal(DistRatesDT.Rows[j]["PaymentRate"]);
                                DataRow[] rows = DTs[i].Select("PaymentRate='" + PaymentRate.ToString() + "'");
                                foreach (DataRow item in rows)
                                    DT1.Rows.Add(item.ItemArray);
                                if (DT1.Rows.Count == 0)
                                    continue;

                                GetSimpleFronts(DT1, _tPSReportDataTable, false);
                                GetSimpleFronts(DT1, _tPSReportDataTable, true);

                                GetCurvedFronts(DT1, _tPSReportDataTable);

                                GetGrids(DT1, _tPSReportDataTable, _profilReportDataTable, 2);

                                GetInsets(DT1, _tPSReportDataTable);

                                GetGlass(DT1, _tPSReportDataTable, _profilReportDataTable, 2);
                            }
                        }
                    }

                }
            }
        }

        public void Report(int[] MainOrderIDs)
        {
            ProfilFrontsOrdersDataTable.Clear();
            TPSFrontsOrdersDataTable.Clear();
            DecorInvNumbersDT.Clear();
            GetMegaOrderInfo(MainOrderIDs[0]);

            string SelectCommand = "SELECT NewFrontsOrders.*, infiniu2_catalog.dbo.FrontsConfig.measureId, infiniu2_catalog.dbo.FrontsConfig.AccountingName, infiniu2_catalog.dbo.FrontsConfig.InvNumber, NewMegaOrders.PaymentRate FROM NewFrontsOrders" +
                " INNER JOIN infiniu2_catalog.dbo.FrontsConfig ON NewFrontsOrders.FrontConfigID = infiniu2_catalog.dbo.FrontsConfig.FrontConfigID" +
                " INNER JOIN NewMainOrders ON NewFrontsOrders.MainOrderID = NewMainOrders.MainOrderID" +
                " INNER JOIN NewMegaOrders ON NewMainOrders.MegaOrderID = NewMegaOrders.MegaOrderID" +
                " WHERE InvNumber IS NOT NULL AND NewFrontsOrders.MainOrderID IN (" + string.Join(",", MainOrderIDs) + ") ORDER BY InvNumber";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    if (DA.Fill(DT) == 0)
                        return;

                    DataTable DistRatesDT = new DataTable();
                    DataTable DT1 = DT.Clone();
                    //get count of different covertypes
                    using (DataView DV = new DataView(DT))
                    {
                        DistRatesDT = DV.ToTable(true, new string[] { "PaymentRate" });
                    }

                    ProfilFrontsOrdersDataTable = DT.Clone();
                    TPSFrontsOrdersDataTable = DT.Clone();

                    SplitTables(DT, ref ProfilFrontsOrdersDataTable, ref TPSFrontsOrdersDataTable);

                    if (ProfilFrontsOrdersDataTable.Rows.Count > 0)
                    {
                        DataTable[] DTs = GroupInvNumber(ProfilFrontsOrdersDataTable);

                        for (int i = 0; i < DTs.Count(); i++)
                        {
                            for (int j = 0; j < DistRatesDT.Rows.Count; j++)
                            {
                                DT1.Clear();
                                PaymentRate = Convert.ToDecimal(DistRatesDT.Rows[j]["PaymentRate"]);
                                DataRow[] rows = DTs[i].Select("PaymentRate='" + PaymentRate.ToString() + "'");
                                foreach (DataRow item in rows)
                                    DT1.Rows.Add(item.ItemArray);
                                if (DT1.Rows.Count == 0)
                                    continue;

                                GetSimpleFronts(DT1, _profilReportDataTable, false);
                                GetSimpleFronts(DT1, _profilReportDataTable, true);

                                GetCurvedFronts(DT1, _profilReportDataTable);

                                GetGrids(DT1, _profilReportDataTable, _tPSReportDataTable, 1);

                                GetInsets(DT1, _profilReportDataTable);

                                GetGlass(DT1, _profilReportDataTable, _tPSReportDataTable, 1);
                            }
                        }
                    }

                    DT1.Clear();
                    if (TPSFrontsOrdersDataTable.Rows.Count > 0)
                    {
                        DataTable[] DTs = GroupInvNumber(TPSFrontsOrdersDataTable);

                        for (int i = 0; i < DTs.Count(); i++)
                        {
                            for (int j = 0; j < DistRatesDT.Rows.Count; j++)
                            {
                                DT1.Clear();
                                PaymentRate = Convert.ToDecimal(DistRatesDT.Rows[j]["PaymentRate"]);
                                DataRow[] rows = DTs[i].Select("PaymentRate='" + PaymentRate.ToString() + "'");
                                foreach (DataRow item in rows)
                                    DT1.Rows.Add(item.ItemArray);
                                if (DT1.Rows.Count == 0)
                                    continue;

                                GetSimpleFronts(DT1, _tPSReportDataTable, false);
                                GetSimpleFronts(DT1, _tPSReportDataTable, true);

                                GetCurvedFronts(DT1, _tPSReportDataTable);

                                GetGrids(DT1, _tPSReportDataTable, _profilReportDataTable, 2);

                                GetInsets(DT1, _tPSReportDataTable);

                                GetGlass(DT1, _tPSReportDataTable, _profilReportDataTable, 2);
                            }
                        }
                    }

                }
            }
        }
    }
}
