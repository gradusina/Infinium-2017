
using System;
using System.Data;
using System.Data.SqlClient;
using System.Linq;

namespace Infinium.Modules.Marketing.NewOrders.PrepareReport.StandardReport
{
    public class FrontsReport
    {
        private FrontsCalculate FC = null;
        private int ClientID = 0;
        private string ProfilCurrencyCode = "0";
        private string TPSCurrencyCode = "0";
        private decimal PaymentRate = 0;
        private string UNN = string.Empty;

        private DataTable CurrencyTypesDT;
        private DataTable ProfilFrontsOrdersDataTable = null;
        private DataTable TPSFrontsOrdersDataTable = null;
        public DataTable FrontsDataTable = null;
        public DataTable FrameColorsDataTable = null;
        public DataTable PatinaDataTable = null;
        public DataTable InsetTypesDataTable = null;
        public DataTable InsetColorsDataTable = null;
        public DataTable TechnoInsetTypesDataTable = null;
        public DataTable TechnoInsetColorsDataTable = null;
        private DataTable MeasuresDataTable = null;
        private DataTable FactoryDataTable = null;
        private DataTable GridSizesDataTable = null;
        private DataTable FrontsConfigDataTable = null;
        private DataTable TechStoreDataTable = null;

        public DataTable ProfilReportDataTable = null;
        public DataTable TPSReportDataTable = null;

        public FrontsReport(ref FrontsCalculate tFC)
        {
            FC = tFC;

            Create();
            CreateReportDataTables();
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

        private void Create()
        {
            ProfilFrontsOrdersDataTable = new DataTable();
            TPSFrontsOrdersDataTable = new DataTable();

            string SelectCommand = "SELECT * FROM CurrencyTypes";
            CurrencyTypesDT = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(CurrencyTypesDT);
            }

            SelectCommand = @"SELECT TechStoreID AS FrontID, TechStoreName AS FrontName FROM TechStore 
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

            MeasuresDataTable = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM Measures",
                ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(MeasuresDataTable);
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
            FrontsConfigDataTable = TablesManager.FrontsConfigDataTable;

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
            ProfilReportDataTable.Clear();
            TPSReportDataTable.Clear();
        }

        private void CreateReportDataTables()
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
            ProfilReportDataTable.Columns.Add(new DataColumn("Price", Type.GetType("System.String")));
            ProfilReportDataTable.Columns.Add(new DataColumn("Cost", Type.GetType("System.String")));
            ProfilReportDataTable.Columns.Add(new DataColumn("Weight", Type.GetType("System.String")));
            ProfilReportDataTable.Columns.Add(new DataColumn("OriginalPrice", Type.GetType("System.String")));
            ProfilReportDataTable.Columns.Add(new DataColumn("PriceWithTransport", Type.GetType("System.String")));
            ProfilReportDataTable.Columns.Add(new DataColumn("CostWithTransport", Type.GetType("System.String")));
            ProfilReportDataTable.Columns.Add(new DataColumn("DiscountVolume", Type.GetType("System.String")));
            ProfilReportDataTable.Columns.Add(new DataColumn("TotalDiscount", Type.GetType("System.String")));
            ProfilReportDataTable.Columns.Add(new DataColumn("PaymentRate", Type.GetType("System.Decimal")));
            ProfilReportDataTable.Columns.Add(new DataColumn("IsNonStandard", Type.GetType("System.Boolean")));
            ProfilReportDataTable.Columns.Add(new DataColumn("NonStandardMargin", Type.GetType("System.Decimal")));

            TPSReportDataTable = ProfilReportDataTable.Clone();
        }

        public string GetFrontName(int FrontID)
        {
            DataRow[] Row = FrontsDataTable.Select("FrontID = " + FrontID);

            return Row[0]["FrontName"].ToString();
        }

        private void SplitTables(DataTable FrontsOrdersDataTable, ref DataTable ProfilDT, ref DataTable TPSDT)
        {
            for (int i = 0; i < FrontsOrdersDataTable.Rows.Count; i++)
            {
                DataRow[] Rows = FrontsConfigDataTable.Select("FrontConfigID = " + FrontsOrdersDataTable.Rows[i]["FrontConfigID"].ToString());

                if (Convert.ToDateTime(FrontsOrdersDataTable.Rows[i]["CreateDateTime"]) < new DateTime(2019, 10, 01))
                {
                    if (Rows[0]["AreaID"].ToString() == "1")//profil
                        ProfilDT.ImportRow(FrontsOrdersDataTable.Rows[i]);

                    if (Rows[0]["AreaID"].ToString() == "2")//tps
                        TPSDT.ImportRow(FrontsOrdersDataTable.Rows[i]);
                }
                else
                {
                    if (Rows[0]["FactoryID"].ToString() == "1")//profil
                        ProfilDT.ImportRow(FrontsOrdersDataTable.Rows[i]);

                    if (Rows[0]["FactoryID"].ToString() == "2")//tps
                        TPSDT.ImportRow(FrontsOrdersDataTable.Rows[i]);
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

        private decimal GetInsetSquare(DataRow FrontsOrdersRow)
        {
            decimal GridHeight = 0;
            decimal GridWidth = 0;
            DataRow[] Rows = TechStoreDataTable.Select("TechStoreID = " + Convert.ToInt32(FrontsOrdersRow["FrontID"]));
            if (Rows.Count() > 0)
            {
                if (Rows[0]["InsetHeightAdmission"] != DBNull.Value)
                    GridHeight = Convert.ToInt32(FrontsOrdersRow["Height"]) - Convert.ToInt32(Rows[0]["InsetHeightAdmission"]);
                if (Rows[0]["InsetWidthAdmission"] != DBNull.Value)
                    GridWidth = Convert.ToInt32(FrontsOrdersRow["Width"]) - Convert.ToInt32(Rows[0]["InsetWidthAdmission"]);
                if (Convert.ToInt32(FrontsOrdersRow["FrontID"]) == 3729)
                {
                    return Decimal.Round(Convert.ToInt32(Rows[0]["InsetHeightAdmission"]) * GridWidth / 1000000, 3, MidpointRounding.AwayFromZero);
                }
            }
            return Decimal.Round(GridHeight * GridWidth / 1000000, 3, MidpointRounding.AwayFromZero);
        }

        private void GetGlassMarginAluminium(DataRow FrontsOrdersRow, ref int GlassMarginHeight, ref int GlassMarginWidth)
        {
            DataRow[] Rows = TechStoreDataTable.Select("TechStoreID = " + Convert.ToInt32(FrontsOrdersRow["FrontID"]));
            if (Rows.Count() > 0)
            {
                if (Rows[0]["InsetHeightAdmission"] != DBNull.Value)
                    GlassMarginHeight = Convert.ToInt32(Rows[0]["InsetHeightAdmission"]);
                if (Rows[0]["InsetWidthAdmission"] != DBNull.Value)
                    GlassMarginWidth = Convert.ToInt32(Rows[0]["InsetWidthAdmission"]);
            }
        }

        public void GetMegaOrderInfo(int MainOrderID)
        {
            string SelectCommand = "SELECT MegaOrderID, TransportCost, AdditionalCost, Rate, ClientID, CurrencyTypeID FROM NewMegaOrders" +
                " WHERE MegaOrderID IN (SELECT MegaOrderID FROM NewMainOrders WHERE MainOrderID = " + MainOrderID + ")";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    if (DA.Fill(DT) > 0)
                    {
                        int CurrencyTypeID = 0;
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

        private decimal GetNonStandardMargin(int FrontConfigID)
        {
            DataRow[] Rows = FrontsConfigDataTable.Select("FrontConfigID = " + FrontConfigID);

            return Convert.ToDecimal(Rows[0]["ZOVNonStandMargin"]);
        }

        private void GetSimpleFronts(DataTable OrdersDataTable, DataTable ReportDataTable, bool IsNonStandard)
        {
            string IsNonStandardFilter = "IsNonStandard=0";
            DataTable Fronts = new DataTable();
            string AccountingName = OrdersDataTable.Rows[0]["AccountingName"].ToString();
            string InvNumber = OrdersDataTable.Rows[0]["InvNumber"].ToString();
            int fID = Convert.ToInt32(OrdersDataTable.Rows[0]["FactoryID"]);
            if (IsNonStandard)
                IsNonStandardFilter = "IsNonStandard=1";

            using (DataView DV = new DataView(OrdersDataTable))
            {
                DV.RowFilter = IsNonStandardFilter;
                Fronts = DV.ToTable(true, new string[] { "FrontID" });
            }

            for (int i = 0; i < Fronts.Rows.Count; i++)
            {
                decimal SolidCount = 0;
                decimal SolidCost = 0;
                decimal OriginalSolidCost = 0;
                decimal WithTransportSolidCost = 0;
                decimal SolidWeight = 0;

                decimal FilenkaCount = 0;
                decimal FilenkaCost = 0;
                decimal OriginalFilenkaCost = 0;
                decimal WithTransportFilenkaCost = 0;
                decimal FilenkaWeight = 0;

                decimal VitrinaCount = 0;
                decimal VitrinaCost = 0;
                decimal OriginalVitrinaCost = 0;
                decimal WithTransportVitrinaCost = 0;
                decimal VitrinaWeight = 0;

                decimal LuxCount = 0;
                decimal LuxCost = 0;
                decimal OriginalLuxCost = 0;
                decimal WithTransportLuxCost = 0;
                decimal LuxWeight = 0;

                decimal MegaCount = 0;
                decimal MegaCost = 0;
                decimal OriginalMegaCost = 0;
                decimal WithTransportMegaCost = 0;
                decimal MegaWeight = 0;

                decimal TotalDiscount = 0;

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
                DataRow[] Rows = OrdersDataTable.Select("FrontID = " + Fronts.Rows[i]["FrontID"].ToString() + " AND (Width <> -1)" + filter + IsNonStandardFilter);
                for (int r = 0; r < Rows.Count(); r++)
                {
                    decimal DeductibleCost = 0;
                    decimal DeductibleCount = 0;
                    decimal DeductibleWeight = 0;
                    if (Convert.ToInt32(Fronts.Rows[i]["FrontID"]) == 3729)//РЕШЕТКА ОВАЛ
                    {
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

                    TotalDiscount = Convert.ToDecimal(Rows[r]["TotalDiscount"]);
                    SolidCount += Convert.ToDecimal(Rows[r]["Square"]);

                    SolidCost += Convert.ToDecimal(Rows[r]["FrontPrice"]) * Convert.ToDecimal(Rows[r]["Square"]);
                    OriginalSolidCost += Convert.ToDecimal(Rows[r]["OriginalPrice"]) * Convert.ToDecimal(Rows[r]["Square"]);
                    DeductibleCost = Math.Ceiling(DeductibleCost / 0.01m) * 0.01m;
                    WithTransportSolidCost += Convert.ToDecimal(Rows[r]["CostWithTransport"]) - DeductibleCost;

                    decimal FrontWeight = 0;
                    decimal InsetWeight = 0;

                    GetFrontWeight(Rows[r], ref FrontWeight, ref InsetWeight);

                    SolidWeight += Convert.ToDecimal(FrontWeight + InsetWeight);
                    NonStandardMargin = GetNonStandardMargin(Convert.ToInt32(Rows[r]["FrontConfigID"]));
                }
                //АППЛИКАЦИИ
                filter = " AND (FrontID IN (3728,3731,3732,3739,3740,3741,3744,3745,3746) OR InsetTypeID IN (28961,3653,3654,3655))";
                Rows = OrdersDataTable.Select("FrontID = " + Fronts.Rows[i]["FrontID"].ToString() + " AND (Width <> -1)" + filter + IsNonStandardFilter);
                for (int r = 0; r < Rows.Count(); r++)
                {
                    TotalDiscount = Convert.ToDecimal(Rows[r]["TotalDiscount"]);
                    decimal DeductibleCost = Convert.ToDecimal(Rows[r]["InsetPrice"]) * Convert.ToDecimal(Rows[r]["Count"]);
                    if (Convert.ToInt32(Rows[r]["FrontID"]) == 3728 || Convert.ToInt32(Rows[r]["FrontID"]) == 3731 || Convert.ToInt32(Rows[r]["FrontID"]) == 3732 ||
                        Convert.ToInt32(Rows[r]["FrontID"]) == 3739 || Convert.ToInt32(Rows[r]["FrontID"]) == 3740 || Convert.ToInt32(Rows[r]["FrontID"]) == 3741 ||
                        Convert.ToInt32(Rows[r]["FrontID"]) == 3744 || Convert.ToInt32(Rows[r]["FrontID"]) == 3745 || Convert.ToInt32(Rows[r]["FrontID"]) == 3746)
                    {
                        SolidCount += Convert.ToDecimal(Rows[r]["Square"]);

                        SolidCost += Convert.ToDecimal(Rows[r]["FrontPrice"]) * Convert.ToDecimal(Rows[r]["Square"]) - DeductibleCost;
                        OriginalSolidCost += Convert.ToDecimal(Rows[r]["OriginalPrice"]) * Convert.ToDecimal(Rows[r]["Square"]) - DeductibleCost;

                        WithTransportSolidCost += Convert.ToDecimal(Rows[r]["CostWithTransport"]) - DeductibleCost;

                        decimal FrontWeight = 0;
                        decimal InsetWeight = 0;

                        GetFrontWeight(Rows[r], ref FrontWeight, ref InsetWeight);

                        SolidWeight += Convert.ToDecimal(FrontWeight + InsetWeight);
                    }
                    else if (Convert.ToInt32(Rows[r]["FrontID"]) == 3415 || Convert.ToInt32(Rows[r]["FrontID"]) == 28922)
                    {
                        FilenkaCount += Convert.ToDecimal(Rows[r]["Square"]);

                        FilenkaCost += Convert.ToDecimal(Rows[r]["FrontPrice"]) * Convert.ToDecimal(Rows[r]["Square"]) - DeductibleCost;
                        OriginalFilenkaCost += Convert.ToDecimal(Rows[r]["OriginalPrice"]) * Convert.ToDecimal(Rows[r]["Square"]) - DeductibleCost;

                        WithTransportFilenkaCost += Convert.ToDecimal(Rows[r]["CostWithTransport"]) - DeductibleCost;

                        decimal FrontWeight = 0;
                        decimal InsetWeight = 0;

                        GetFrontWeight(Rows[r], ref FrontWeight, ref InsetWeight);

                        FilenkaWeight += Convert.ToDecimal(FrontWeight + InsetWeight);
                    }
                    NonStandardMargin = GetNonStandardMargin(Convert.ToInt32(Rows[r]["FrontConfigID"]));
                }
                //ФИЛЕНКА
                filter = " AND InsetTypeID IN (2069,2070,2071,2073,2075,42066,2077,2233,3644,29043,29531,41213)";
                Rows = OrdersDataTable.Select("FrontID = " + Fronts.Rows[i]["FrontID"].ToString() + " AND (Width <> -1)" + filter + IsNonStandardFilter);
                for (int r = 0; r < Rows.Count(); r++)
                {
                    TotalDiscount = Convert.ToDecimal(Rows[r]["TotalDiscount"]);

                    FilenkaCount += Convert.ToDecimal(Rows[r]["Square"]);

                    FilenkaCost += Convert.ToDecimal(Rows[r]["FrontPrice"]) * Convert.ToDecimal(Rows[r]["Square"]);
                    OriginalFilenkaCost += Convert.ToDecimal(Rows[r]["OriginalPrice"]) * Convert.ToDecimal(Rows[r]["Square"]);

                    WithTransportFilenkaCost += Convert.ToDecimal(Rows[r]["CostWithTransport"]);

                    decimal FrontWeight = 0;
                    decimal InsetWeight = 0;

                    GetFrontWeight(Rows[r], ref FrontWeight, ref InsetWeight);

                    FilenkaWeight += Convert.ToDecimal(FrontWeight + InsetWeight);
                    NonStandardMargin = GetNonStandardMargin(Convert.ToInt32(Rows[r]["FrontConfigID"]));
                }
                //ВИТРИНЫ, РЕШЕТКИ, СТЕКЛО
                filter = " AND InsetTypeID IN (1,2,685,686,687,688,29470,29471) AND FrontID <> 3729";
                Rows = OrdersDataTable.Select("FrontID = " + Fronts.Rows[i]["FrontID"].ToString() + " AND (Width <> -1)" + filter + IsNonStandardFilter);
                for (int r = 0; r < Rows.Count(); r++)
                {
                    TotalDiscount = Convert.ToDecimal(Rows[r]["TotalDiscount"]);
                    decimal DeductibleCount = 0;
                    decimal DeductibleCost = 0;
                    decimal DeductibleWeight = 0;
                    //РЕШЕТКА 45,90,ПЛАСТИК
                    if (Convert.ToInt32(Rows[r]["InsetTypeID"]) == 685 || Convert.ToInt32(Rows[r]["InsetTypeID"]) == 686 ||
                        Convert.ToInt32(Rows[r]["InsetTypeID"]) == 687 || Convert.ToInt32(Rows[r]["InsetTypeID"]) == 688 ||
                        Convert.ToInt32(Rows[r]["InsetTypeID"]) == 29470 || Convert.ToInt32(Rows[r]["InsetTypeID"]) == 29471)
                    {
                        decimal InsetSquare = GetInsetSquare(Convert.ToInt32(Rows[r]["FrontID"]), Convert.ToInt32(Rows[r]["Height"]), Convert.ToInt32(Rows[r]["Width"]));
                        InsetSquare = Decimal.Round(InsetSquare, 3, MidpointRounding.AwayFromZero);
                        DeductibleCount = InsetSquare * Convert.ToDecimal(Rows[r]["Count"]);
                        DeductibleCost = Convert.ToDecimal(Rows[r]["InsetPrice"]) * InsetSquare * Convert.ToDecimal(Rows[r]["Count"]);
                        DeductibleWeight = Decimal.Round(DeductibleCount * Convert.ToDecimal(3.5), 3, MidpointRounding.AwayFromZero);
                    }
                    //СТЕКЛО
                    if (Convert.ToInt32(Rows[r]["InsetTypeID"]) == 2)
                    {
                        DeductibleCount = Convert.ToDecimal(Rows[r]["Count"]) * GetInsetSquare(Convert.ToInt32(Rows[r]["FrontID"]), Convert.ToInt32(Rows[r]["Height"]), Convert.ToInt32(Rows[r]["Width"]));
                        DeductibleCost = Convert.ToDecimal(Rows[r]["InsetPrice"]) * DeductibleCount;
                        DeductibleWeight = Decimal.Round(DeductibleCount * 10, 3, MidpointRounding.AwayFromZero);
                    }

                    VitrinaCount += Convert.ToDecimal(Rows[r]["Square"]);

                    if (FC.IsAluminium(Rows[r]) > -1)
                    {
                        VitrinaCost += FC.GetFrontCostAluminium(ClientID, Rows[r]);
                        VitrinaWeight += FC.GetAluminiumWeight(Rows[r], true);
                    }
                    else
                    {
                        VitrinaCost += Convert.ToDecimal(Rows[r]["FrontPrice"]) * Convert.ToDecimal(Rows[r]["Square"]);

                        decimal FrontWeight = 0;
                        decimal InsetWeight = 0;
                        GetFrontWeight(Rows[r], ref FrontWeight, ref InsetWeight);
                        VitrinaWeight += Convert.ToDecimal(FrontWeight + InsetWeight - DeductibleWeight);
                    }
                    VitrinaCost = Math.Ceiling(VitrinaCost / 0.01m) * 0.01m;
                    OriginalVitrinaCost += Convert.ToDecimal(Rows[r]["OriginalPrice"]) * Convert.ToDecimal(Rows[r]["Square"]);

                    DeductibleCost = Math.Ceiling(DeductibleCost / 0.01m) * 0.01m;
                    WithTransportVitrinaCost += Convert.ToDecimal(Rows[r]["CostWithTransport"]) - DeductibleCost;

                    NonStandardMargin = GetNonStandardMargin(Convert.ToInt32(Rows[r]["FrontConfigID"]));
                }
                //ЛЮКС
                filter = " AND InsetTypeID IN (860)";
                Rows = OrdersDataTable.Select("FrontID = " + Fronts.Rows[i]["FrontID"].ToString() + " AND (Width <> -1)" + filter + IsNonStandardFilter);
                for (int r = 0; r < Rows.Count(); r++)
                {
                    TotalDiscount = Convert.ToDecimal(Rows[r]["TotalDiscount"]);
                    LuxCount += Convert.ToDecimal(Rows[r]["Square"]);
                    LuxCost += Convert.ToDecimal(Rows[r]["FrontPrice"]) * Convert.ToDecimal(Rows[r]["Square"]);
                    OriginalLuxCost += Convert.ToDecimal(Rows[r]["OriginalPrice"]) * Convert.ToDecimal(Rows[r]["Square"]);

                    WithTransportLuxCost += Convert.ToDecimal(Rows[r]["CostWithTransport"]);

                    decimal FrontWeight = 0;
                    decimal InsetWeight = 0;

                    GetFrontWeight(Rows[r], ref FrontWeight, ref InsetWeight);

                    LuxWeight += Convert.ToDecimal(FrontWeight + InsetWeight);
                    NonStandardMargin = GetNonStandardMargin(Convert.ToInt32(Rows[r]["FrontConfigID"]));
                }
                //МЕГА, Планка
                filter = " AND InsetTypeID IN (862,4310)";
                Rows = OrdersDataTable.Select("FrontID = " + Fronts.Rows[i]["FrontID"].ToString() + " AND (Width <> -1)" + filter + IsNonStandardFilter);
                for (int r = 0; r < Rows.Count(); r++)
                {
                    TotalDiscount = Convert.ToDecimal(Rows[r]["TotalDiscount"]);

                    MegaCount += Convert.ToDecimal(Rows[r]["Square"]);

                    MegaCost += Convert.ToDecimal(Rows[r]["FrontPrice"]) * Convert.ToDecimal(Rows[r]["Square"]);
                    OriginalMegaCost += Convert.ToDecimal(Rows[r]["OriginalPrice"]) * Convert.ToDecimal(Rows[r]["Square"]);

                    WithTransportMegaCost += Convert.ToDecimal(Rows[r]["CostWithTransport"]);

                    decimal FrontWeight = 0;
                    decimal InsetWeight = 0;

                    GetFrontWeight(Rows[r], ref FrontWeight, ref InsetWeight);

                    MegaWeight += Convert.ToDecimal(FrontWeight + InsetWeight);
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
                    Row["UNN"] = UNN;
                    Row["PaymentRate"] = PaymentRate;
                    Row["AccountingName"] = AccountingName;
                    Row["InvNumber"] = InvNumber;
                    Row["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        Row["TPSCurCode"] = TPSCurrencyCode;
                    Row["TotalDiscount"] = TotalDiscount;
                    Row["Name"] = GetFrontName(Convert.ToInt32(Fronts.Rows[i]["FrontID"])) + " глухой";
                    Row["Count"] = Decimal.Round(SolidCount, 3, MidpointRounding.AwayFromZero);
                    Row["Measure"] = "м.кв.";
                    Row["Price"] = Decimal.Round(SolidCost / SolidCount, 2, MidpointRounding.AwayFromZero);
                    Row["OriginalPrice"] = Decimal.Round(OriginalSolidCost / SolidCount, 2, MidpointRounding.AwayFromZero);
                    Row["PriceWithTransport"] = Decimal.Round(WithTransportSolidCost / SolidCount, 2, MidpointRounding.AwayFromZero);
                    Row["CostWithTransport"] = Decimal.Round(Math.Ceiling(WithTransportSolidCost / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                    Row["Cost"] = Decimal.Round(Math.Ceiling(SolidCost / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                    Row["Weight"] = Decimal.Round(SolidWeight, 3, MidpointRounding.AwayFromZero);
                    ReportDataTable.Rows.Add(Row);
                }

                //FilenkaCount = Math.Ceiling(FilenkaCount / 0.01m) * 0.01m;
                //WithTransportFilenkaCost = Math.Ceiling(WithTransportFilenkaCost / 0.01m) * 0.01m;
                if (FilenkaCount > 0)
                {
                    DataRow Row = ReportDataTable.NewRow();
                    Row["NonStandardMargin"] = NonStandardMargin;
                    Row["IsNonStandard"] = IsNonStandard;
                    Row["UNN"] = UNN;
                    Row["PaymentRate"] = PaymentRate;
                    Row["AccountingName"] = AccountingName;
                    Row["InvNumber"] = InvNumber;
                    Row["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        Row["TPSCurCode"] = TPSCurrencyCode;
                    Row["TotalDiscount"] = TotalDiscount;
                    Row["Name"] = GetFrontName(Convert.ToInt32(Fronts.Rows[i]["FrontID"])) + " филенка";
                    Row["Count"] = Decimal.Round(FilenkaCount, 3, MidpointRounding.AwayFromZero);
                    Row["Measure"] = "м.кв.";
                    Row["Price"] = Decimal.Round(FilenkaCost / FilenkaCount, 2, MidpointRounding.AwayFromZero);
                    Row["OriginalPrice"] = Decimal.Round(OriginalFilenkaCost / FilenkaCount, 2, MidpointRounding.AwayFromZero);
                    Row["PriceWithTransport"] = Decimal.Round(WithTransportFilenkaCost / FilenkaCount, 2, MidpointRounding.AwayFromZero);
                    Row["CostWithTransport"] = Decimal.Round(Math.Ceiling(WithTransportFilenkaCost / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                    Row["Cost"] = Decimal.Round(Math.Ceiling(FilenkaCount / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                    Row["Weight"] = Decimal.Round(FilenkaWeight, 3, MidpointRounding.AwayFromZero);
                    ReportDataTable.Rows.Add(Row);
                }

                //VitrinaCost = Math.Ceiling(VitrinaCost / 0.01m) * 0.01m;
                //WithTransportVitrinaCost = Math.Ceiling(WithTransportVitrinaCost / 0.01m) * 0.01m;
                //decimal PriceWithTransport = WithTransportVitrinaCost / VitrinaCount;
                //PriceWithTransport = Decimal.Round(WithTransportVitrinaCost / VitrinaCount, 2, MidpointRounding.AwayFromZero);
                if (VitrinaCount > 0)
                {
                    DataRow Row = ReportDataTable.NewRow();
                    Row["NonStandardMargin"] = NonStandardMargin;
                    Row["IsNonStandard"] = IsNonStandard;
                    Row["UNN"] = UNN;
                    Row["PaymentRate"] = PaymentRate;
                    Row["AccountingName"] = AccountingName;
                    Row["InvNumber"] = InvNumber;
                    Row["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        Row["TPSCurCode"] = TPSCurrencyCode;
                    Row["TotalDiscount"] = TotalDiscount;
                    Row["Name"] = GetFrontName(Convert.ToInt32(Fronts.Rows[i]["FrontID"])) + " витрина";
                    Row["Count"] = Decimal.Round(VitrinaCount, 3, MidpointRounding.AwayFromZero);
                    Row["Measure"] = "м.кв.";
                    Row["Price"] = Decimal.Round(VitrinaCost / VitrinaCount, 2, MidpointRounding.AwayFromZero);
                    Row["OriginalPrice"] = Decimal.Round(OriginalVitrinaCost / VitrinaCount, 2, MidpointRounding.AwayFromZero);
                    Row["PriceWithTransport"] = Decimal.Round(WithTransportVitrinaCost / VitrinaCount, 2, MidpointRounding.AwayFromZero);
                    Row["CostWithTransport"] = Decimal.Round(Math.Ceiling(WithTransportVitrinaCost / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                    Row["Cost"] = Decimal.Round(Math.Ceiling(VitrinaCost / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                    Row["Weight"] = Decimal.Round(VitrinaWeight, 3, MidpointRounding.AwayFromZero);
                    ReportDataTable.Rows.Add(Row);
                }

                //LuxCost = Math.Ceiling(LuxCost / 0.01m) * 0.01m;
                //WithTransportLuxCost = Math.Ceiling(WithTransportLuxCost / 0.01m) * 0.01m;
                if (LuxCount > 0)
                {
                    DataRow Row = ReportDataTable.NewRow();
                    Row["NonStandardMargin"] = NonStandardMargin;
                    Row["IsNonStandard"] = IsNonStandard;
                    Row["UNN"] = UNN;
                    Row["PaymentRate"] = PaymentRate;
                    Row["AccountingName"] = AccountingName;
                    Row["InvNumber"] = InvNumber;
                    Row["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        Row["TPSCurCode"] = TPSCurrencyCode;
                    Row["TotalDiscount"] = TotalDiscount;
                    Row["Name"] = GetFrontName(Convert.ToInt32(Fronts.Rows[i]["FrontID"])) + " люкс";
                    Row["Count"] = Decimal.Round(LuxCount, 3, MidpointRounding.AwayFromZero);
                    Row["Measure"] = "м.кв.";
                    Row["Price"] = Decimal.Round(LuxCost / LuxCount, 2, MidpointRounding.AwayFromZero);
                    Row["OriginalPrice"] = Decimal.Round(OriginalLuxCost / LuxCount, 2, MidpointRounding.AwayFromZero);
                    Row["PriceWithTransport"] = Decimal.Round(WithTransportLuxCost / LuxCount, 2, MidpointRounding.AwayFromZero);
                    Row["CostWithTransport"] = Decimal.Round(Math.Ceiling(WithTransportLuxCost / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                    Row["Cost"] = Decimal.Round(Math.Ceiling(LuxCost / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                    Row["Weight"] = Decimal.Round(LuxWeight, 3, MidpointRounding.AwayFromZero);
                    ReportDataTable.Rows.Add(Row);
                }

                //MegaCost = Math.Ceiling(MegaCost / 0.01m) * 0.01m;
                //WithTransportMegaCost = Math.Ceiling(WithTransportMegaCost / 0.01m) * 0.01m;
                if (MegaCount > 0)
                {
                    DataRow Row = ReportDataTable.NewRow();
                    Row["NonStandardMargin"] = NonStandardMargin;
                    Row["IsNonStandard"] = IsNonStandard;
                    Row["UNN"] = UNN;
                    Row["PaymentRate"] = PaymentRate;
                    Row["AccountingName"] = AccountingName;
                    Row["InvNumber"] = InvNumber;
                    Row["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        Row["TPSCurCode"] = TPSCurrencyCode;
                    Row["TotalDiscount"] = TotalDiscount;
                    Row["Name"] = GetFrontName(Convert.ToInt32(Fronts.Rows[i]["FrontID"])) + " мега";
                    Row["Count"] = Decimal.Round(MegaCount, 3, MidpointRounding.AwayFromZero);
                    Row["Measure"] = "м.кв.";
                    Row["Price"] = Decimal.Round(MegaCost / MegaCount, 2, MidpointRounding.AwayFromZero);
                    Row["OriginalPrice"] = Decimal.Round(OriginalMegaCost / MegaCount, 2, MidpointRounding.AwayFromZero);
                    Row["PriceWithTransport"] = Decimal.Round(WithTransportMegaCost / MegaCount, 2, MidpointRounding.AwayFromZero);
                    Row["CostWithTransport"] = Decimal.Round(Math.Ceiling(WithTransportMegaCost / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                    Row["Cost"] = Decimal.Round(Math.Ceiling(MegaCost / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                    Row["Weight"] = Decimal.Round(MegaWeight, 3, MidpointRounding.AwayFromZero);
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

            using (DataView DV = new DataView(OrdersDataTable))
            {
                Fronts = DV.ToTable(true, new string[] { "FrontID" });
            }

            for (int i = 0; i < Fronts.Rows.Count; i++)
            {
                DataRow[] Rows = OrdersDataTable.Select("FrontID = " + Fronts.Rows[i]["FrontID"].ToString() +
                                                              " AND Width = -1");

                if (Rows.Count() == 0)
                    continue;

                decimal Solid713Count = 0;
                decimal Solid713Price = 0;
                decimal Solid713OriginalPrice = 0;
                decimal Solid713WithTransportCost = 0;
                decimal Solid713Weight = 0;

                decimal Filenka713Count = 0;
                decimal Filenka713Price = 0;
                decimal Filenka713OriginalPrice = 0;
                decimal Filenka713WithTransportCost = 0;
                decimal Filenka713Weight = 0;

                decimal NoInset713Count = 0;
                decimal NoInset713Price = 0;
                decimal NoInset713OriginalPrice = 0;
                decimal NoInset713WithTransportCost = 0;
                decimal NoInset713Weight = 0;

                decimal Vitrina713Count = 0;
                decimal Vitrina713Price = 0;
                decimal Vitrina713OriginalPrice = 0;
                decimal Vitrina713WithTransportCost = 0;
                decimal Vitrina713Weight = 0;

                decimal Solid910Count = 0;
                decimal Solid910Price = 0;
                decimal Solid910OriginalPrice = 0;
                decimal Solid910WithTransportCost = 0;
                decimal Solid910Weight = 0;

                decimal Filenka910Count = 0;
                decimal Filenka910Price = 0;
                decimal Filenka910OriginalPrice = 0;
                decimal Filenka910WithTransportCost = 0;
                decimal Filenka910Weight = 0;

                decimal NoInset910Count = 0;
                decimal NoInset910Price = 0;
                decimal NoInset910OriginalPrice = 0;
                decimal NoInset910WithTransportCost = 0;
                decimal NoInset910Weight = 0;

                decimal Vitrina910Count = 0;
                decimal Vitrina910Price = 0;
                decimal Vitrina910OriginalPrice = 0;
                decimal Vitrina910WithTransportCost = 0;
                decimal Vitrina910Weight = 0;

                decimal TotalDiscount = 0;

                for (int r = 0; r < Rows.Count(); r++)
                {
                    TotalDiscount = Convert.ToDecimal(Rows[r]["TotalDiscount"]);
                    if (Rows[r]["Height"].ToString() == "713")
                    {
                        DataRow[] rows = InsetTypesDataTable.Select("GroupID = 3 OR GroupID = 4");
                        foreach (DataRow item in rows)
                        {
                            if (Rows[r]["InsetTypeID"].ToString() == item["InsetTypeID"].ToString())
                            {
                                Solid713Count += Convert.ToDecimal(Rows[r]["Count"]);
                                Solid713Price = Convert.ToDecimal(Rows[r]["FrontPrice"]);
                                Solid713OriginalPrice = Convert.ToDecimal(Rows[r]["OriginalPrice"]);
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
                                Filenka713OriginalPrice = Convert.ToDecimal(Rows[r]["OriginalPrice"]);
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
                            NoInset713OriginalPrice = Convert.ToDecimal(Rows[r]["OriginalPrice"]);
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
                            Vitrina713OriginalPrice = Convert.ToDecimal(Rows[r]["OriginalPrice"]);
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
                                Solid910OriginalPrice = Convert.ToDecimal(Rows[r]["OriginalPrice"]);
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
                                Filenka910OriginalPrice = Convert.ToDecimal(Rows[r]["OriginalPrice"]);
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
                            NoInset910OriginalPrice = Convert.ToDecimal(Rows[r]["OriginalPrice"]);
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
                            Vitrina910OriginalPrice = Convert.ToDecimal(Rows[r]["OriginalPrice"]);
                            Vitrina910WithTransportCost += Convert.ToDecimal(Rows[r]["PriceWithTransport"]) * Convert.ToDecimal(Rows[r]["Count"]);

                            decimal FrontWeight = 0;
                            decimal InsetWeight = 0;

                            GetFrontWeight(Rows[r], ref FrontWeight, ref InsetWeight);

                            Vitrina910Weight += Convert.ToDecimal(FrontWeight + InsetWeight);
                        }
                    }

                }

                //decimal Cost = Math.Ceiling(Solid713Count * Solid713Price / 0.01m) * 0.01m;
                //Solid713WithTransportCost = Math.Ceiling(Solid713WithTransportCost / 0.01m) * 0.01m;
                if (Solid713Count > 0)
                {
                    DataRow Row = ReportDataTable.NewRow();
                    Row["UNN"] = UNN;
                    Row["PaymentRate"] = PaymentRate;
                    Row["AccountingName"] = AccountingName;
                    Row["InvNumber"] = InvNumber;
                    Row["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        Row["TPSCurCode"] = TPSCurrencyCode;
                    Row["TotalDiscount"] = TotalDiscount;
                    Row["Name"] = GetFrontName(Convert.ToInt32(Fronts.Rows[i]["FrontID"])) + " гнут. 713 глух.";
                    Row["Count"] = Solid713Count;
                    Row["Measure"] = "шт.";
                    Row["Price"] = Decimal.Round(Solid713Price, 2, MidpointRounding.AwayFromZero);
                    Row["OriginalPrice"] = Decimal.Round(Solid713OriginalPrice, 2, MidpointRounding.AwayFromZero);
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
                    Row["UNN"] = UNN;
                    Row["PaymentRate"] = PaymentRate;
                    Row["AccountingName"] = AccountingName;
                    Row["InvNumber"] = InvNumber;
                    Row["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        Row["TPSCurCode"] = TPSCurrencyCode;
                    Row["TotalDiscount"] = TotalDiscount;
                    Row["Name"] = GetFrontName(Convert.ToInt32(Fronts.Rows[i]["FrontID"])) + " гнут. 713 филенка";
                    Row["Count"] = Filenka713Count;
                    Row["Measure"] = "шт.";
                    Row["Price"] = Decimal.Round(Filenka713Price, 2, MidpointRounding.AwayFromZero);
                    Row["OriginalPrice"] = Decimal.Round(Filenka713OriginalPrice, 2, MidpointRounding.AwayFromZero);
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
                    Row["UNN"] = UNN;
                    Row["PaymentRate"] = PaymentRate;
                    Row["AccountingName"] = AccountingName;
                    Row["InvNumber"] = InvNumber;
                    Row["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        Row["TPSCurCode"] = TPSCurrencyCode;
                    Row["TotalDiscount"] = TotalDiscount;
                    Row["Name"] = GetFrontName(Convert.ToInt32(Fronts.Rows[i]["FrontID"])) + " гнут. 713 фрезер.";
                    Row["Count"] = NoInset713Count;
                    Row["Measure"] = "шт.";
                    Row["Price"] = Decimal.Round(NoInset713Price, 2, MidpointRounding.AwayFromZero);
                    Row["OriginalPrice"] = Decimal.Round(NoInset713OriginalPrice, 2, MidpointRounding.AwayFromZero);
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
                    Row["UNN"] = UNN;
                    Row["PaymentRate"] = PaymentRate;
                    Row["AccountingName"] = AccountingName;
                    Row["InvNumber"] = InvNumber;
                    Row["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        Row["TPSCurCode"] = TPSCurrencyCode;
                    Row["TotalDiscount"] = TotalDiscount;
                    Row["Name"] = GetFrontName(Convert.ToInt32(Fronts.Rows[i]["FrontID"])) + " гнут. 713 витрина";
                    Row["Count"] = Vitrina713Count;
                    Row["Measure"] = "шт.";
                    Row["Price"] = Decimal.Round(Vitrina713Price, 2, MidpointRounding.AwayFromZero);
                    Row["OriginalPrice"] = Decimal.Round(Vitrina713OriginalPrice, 2, MidpointRounding.AwayFromZero);
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
                    Row["UNN"] = UNN;
                    Row["PaymentRate"] = PaymentRate;
                    Row["AccountingName"] = AccountingName;
                    Row["InvNumber"] = InvNumber;
                    Row["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        Row["TPSCurCode"] = TPSCurrencyCode;
                    Row["TotalDiscount"] = TotalDiscount;
                    Row["Name"] = GetFrontName(Convert.ToInt32(Fronts.Rows[i]["FrontID"])) + " гнут. 910 глух.";
                    Row["Count"] = Solid910Count;
                    Row["Measure"] = "шт.";
                    Row["Price"] = Decimal.Round(Solid910Price, 2, MidpointRounding.AwayFromZero);
                    Row["OriginalPrice"] = Decimal.Round(Solid910OriginalPrice, 2, MidpointRounding.AwayFromZero);
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
                    Row["UNN"] = UNN;
                    Row["PaymentRate"] = PaymentRate;
                    Row["AccountingName"] = AccountingName;
                    Row["InvNumber"] = InvNumber;
                    Row["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        Row["TPSCurCode"] = TPSCurrencyCode;
                    Row["TotalDiscount"] = TotalDiscount;
                    Row["Name"] = GetFrontName(Convert.ToInt32(Fronts.Rows[i]["FrontID"])) + " гнут. 910 филенка";
                    Row["Count"] = Filenka910Count;
                    Row["Measure"] = "шт.";
                    Row["Price"] = Decimal.Round(Filenka910Price, 2, MidpointRounding.AwayFromZero);
                    Row["OriginalPrice"] = Decimal.Round(Filenka910OriginalPrice, 2, MidpointRounding.AwayFromZero);
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
                    Row["UNN"] = UNN;
                    Row["PaymentRate"] = PaymentRate;
                    Row["AccountingName"] = AccountingName;
                    Row["InvNumber"] = InvNumber;
                    Row["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        Row["TPSCurCode"] = TPSCurrencyCode;
                    Row["TotalDiscount"] = TotalDiscount;
                    Row["Name"] = GetFrontName(Convert.ToInt32(Fronts.Rows[i]["FrontID"])) + " гнут. 910 фрезер.";
                    Row["Count"] = NoInset910Count;
                    Row["Measure"] = "шт.";
                    Row["Price"] = Decimal.Round(NoInset910Price, 3, MidpointRounding.AwayFromZero);
                    Row["OriginalPrice"] = Decimal.Round(NoInset910OriginalPrice, 2, MidpointRounding.AwayFromZero);
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
                    Row["UNN"] = UNN;
                    Row["PaymentRate"] = PaymentRate;
                    Row["AccountingName"] = AccountingName;
                    Row["InvNumber"] = InvNumber;
                    Row["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        Row["TPSCurCode"] = TPSCurrencyCode;
                    Row["TotalDiscount"] = TotalDiscount;
                    Row["Name"] = GetFrontName(Convert.ToInt32(Fronts.Rows[i]["FrontID"])) + " гнут. 910 витрина";
                    Row["Count"] = Vitrina910Count;
                    Row["Measure"] = "шт.";
                    Row["Price"] = Decimal.Round(Vitrina910Price, 2, MidpointRounding.AwayFromZero);
                    Row["OriginalPrice"] = Decimal.Round(Vitrina910OriginalPrice, 2, MidpointRounding.AwayFromZero);
                    Row["PriceWithTransport"] = Decimal.Round(Vitrina910WithTransportCost / Vitrina910Count, 2, MidpointRounding.AwayFromZero);
                    Row["CostWithTransport"] = Decimal.Round(Math.Ceiling(Vitrina910WithTransportCost / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                    Row["Cost"] = Decimal.Round(Math.Ceiling(Vitrina910Count * Vitrina910Price / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                    Row["Weight"] = Decimal.Round(Vitrina910Weight, 3, MidpointRounding.AwayFromZero);
                    ReportDataTable.Rows.Add(Row);
                }
            }

            Fronts.Dispose();
        }

        private void GetGrids(DataTable OrdersDataTable, DataTable ReportDataTable)
        {
            DataRow[] Rows = OrdersDataTable.Select("InsetTypeID IN (685,686,687,688,29470,29471)");
            if (Rows.Count() == 0)
                return;

            int InsetTypeID = Convert.ToInt32(OrdersDataTable.Rows[0]["InsetTypeID"]);
            int fID = Convert.ToInt32(OrdersDataTable.Rows[0]["FactoryID"]);
            decimal Count = 0;
            decimal Cost = 0;
            decimal Cost1 = 0;
            decimal TotalDiscount = 0;

            for (int i = 0; i < Rows.Count(); i++)
            {
                int FrontID = Convert.ToInt32(Rows[i]["FrontID"]);
                if (FrontID == 3729)
                    continue;
                decimal d = GetInsetSquare(Convert.ToInt32(Rows[i]["FrontID"]), Convert.ToInt32(Rows[i]["Height"]), Convert.ToInt32(Rows[i]["Width"])) * Convert.ToDecimal(Rows[i]["Count"]);
                Count += d;
                Cost += Math.Ceiling(Convert.ToDecimal(Rows[i]["InsetPrice"]) * d / 0.001m) * 0.001m;
                Cost1 += Math.Ceiling(Convert.ToDecimal(Rows[i]["OriginalInsetPrice"]) * d / 0.001m) * 0.001m;
                TotalDiscount = Convert.ToDecimal(Rows[i]["TotalDiscount"]);
            }
            Cost = (Cost / 0.01m) * 0.01m;
            Cost1 = (Cost1 / 0.01m) * 0.01m;
            if (Count > 0)
            {
                Count = Decimal.Round(Count, 3, MidpointRounding.AwayFromZero);
                DataRow NewRow = ReportDataTable.NewRow();
                NewRow["UNN"] = UNN;
                NewRow["PaymentRate"] = PaymentRate;
                NewRow["AccountingName"] = OrdersDataTable.Rows[0]["AccountingName"].ToString();
                NewRow["InvNumber"] = OrdersDataTable.Rows[0]["InvNumber"].ToString();
                NewRow["CurrencyCode"] = ProfilCurrencyCode;
                if (fID == 2)
                    NewRow["TPSCurCode"] = TPSCurrencyCode;
                NewRow["Name"] = "Решетка";
                NewRow["Count"] = Decimal.Round(Count, 3, MidpointRounding.AwayFromZero);
                NewRow["Measure"] = "м.кв.";
                NewRow["TotalDiscount"] = TotalDiscount;
                NewRow["OriginalPrice"] = Decimal.Round(Cost1 / Count, 2, MidpointRounding.AwayFromZero);
                NewRow["Price"] = Decimal.Round(Cost1 / Count, 2, MidpointRounding.AwayFromZero);
                NewRow["Cost"] = Decimal.Round(Math.Ceiling(Cost1 / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                NewRow["PriceWithTransport"] = Decimal.Round(Cost / Count, 2, MidpointRounding.AwayFromZero);
                NewRow["CostWithTransport"] = Decimal.Round(Math.Ceiling(Cost / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                NewRow["Weight"] = Decimal.Round(Convert.ToDecimal(NewRow["Count"]) * Convert.ToDecimal(3.5), 3, MidpointRounding.AwayFromZero);
                ReportDataTable.Rows.Add(NewRow);
            }
        }

        private void GetGlass(DataTable OrdersDataTable, DataTable ReportDataTable)
        {
            string AccountingName = OrdersDataTable.Rows[0]["AccountingName"].ToString();
            string InvNumber = OrdersDataTable.Rows[0]["InvNumber"].ToString();
            int fID = Convert.ToInt32(OrdersDataTable.Rows[0]["FactoryID"]);
            decimal CountFlutes = 0;
            decimal CountLacomat = 0;
            //decimal CountMaster = 0;
            decimal CountKrizet = 0;
            decimal CountOther = 0;

            decimal PriceFlutes = 0;
            decimal PriceLacomat = 0;
            //decimal PriceMaster = 0;
            decimal PriceKrizet = 0;
            if (OrdersDataTable.Select("InsetTypeID = 2").Count() == 0)
                return;

            bool hasGlass = false;
            for (int i = 0; i < OrdersDataTable.Rows.Count; i++)
            {
                int frontID = Convert.ToInt32(OrdersDataTable.Rows[0]["FrontID"]);
                if (FC.IsAluminium(OrdersDataTable.Rows[i]) > -1)
                {
                    hasGlass = true;
                    break;
                }
            }
            if (!hasGlass)
                return;

            decimal TotalDiscount = 0;
            DataRow[] FRows = OrdersDataTable.Select("InsetColorID = 3944");

            if (FRows.Count() > 0)
            {
                PriceFlutes = Convert.ToDecimal(FRows[0]["InsetPrice"]);

                for (int i = 0; i < FRows.Count(); i++)
                {
                    if (FC.IsAluminium(FRows[i]) != -1)
                        continue;

                    TotalDiscount = Convert.ToDecimal(FRows[i]["TotalDiscount"]);
                    CountFlutes += Convert.ToDecimal(FRows[i]["Count"]) * GetInsetSquare(Convert.ToInt32(FRows[i]["FrontID"]), Convert.ToInt32(FRows[i]["Height"]), Convert.ToInt32(FRows[i]["Width"]));
                }

                CountFlutes = Decimal.Round(CountFlutes, 3, MidpointRounding.AwayFromZero);

                if (CountFlutes > 0)
                {
                    //decimal Cost = Math.Ceiling(CountFlutes * PriceFlutes / 0.01m) * 0.01m;
                    DataRow NewRow = ReportDataTable.NewRow();
                    NewRow["UNN"] = UNN;
                    NewRow["PaymentRate"] = PaymentRate;
                    NewRow["AccountingName"] = AccountingName;
                    NewRow["InvNumber"] = InvNumber;
                    NewRow["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        NewRow["TPSCurCode"] = TPSCurrencyCode;
                    NewRow["TotalDiscount"] = TotalDiscount;
                    NewRow["Name"] = "Стекло Флутес";
                    NewRow["Count"] = CountFlutes;
                    NewRow["Measure"] = "м.кв.";
                    NewRow["Price"] = Decimal.Round(PriceFlutes, 2, MidpointRounding.AwayFromZero);
                    NewRow["Cost"] = Decimal.Round(Math.Ceiling(CountFlutes * PriceFlutes / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                    NewRow["PriceWithTransport"] = Decimal.Round(PriceFlutes, 2, MidpointRounding.AwayFromZero);
                    NewRow["CostWithTransport"] = Decimal.Round(Math.Ceiling(CountFlutes * PriceFlutes / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                    NewRow["Weight"] = Decimal.Round(CountFlutes * 10, 3, MidpointRounding.AwayFromZero);
                    ReportDataTable.Rows.Add(NewRow);
                }
            }


            DataRow[] LRows = OrdersDataTable.Select("InsetColorID = 3943");

            if (LRows.Count() > 0)
            {
                PriceLacomat = Convert.ToDecimal(LRows[0]["InsetPrice"]);

                for (int i = 0; i < LRows.Count(); i++)
                {
                    if (FC.IsAluminium(LRows[i]) != -1)
                        continue;

                    TotalDiscount = Convert.ToDecimal(LRows[i]["TotalDiscount"]);
                    CountLacomat += Convert.ToDecimal(LRows[i]["Count"]) *
                                GetInsetSquare(Convert.ToInt32(LRows[i]["FrontID"]),
                                              Convert.ToInt32(LRows[i]["Height"]),
                                              Convert.ToInt32(LRows[i]["Width"]));

                }

                CountLacomat = Decimal.Round(CountLacomat, 3, MidpointRounding.AwayFromZero);

                if (CountLacomat > 0)
                {
                    //decimal Cost = Math.Ceiling(CountLacomat * PriceLacomat / 0.01m) * 0.01m;
                    DataRow NewRow = ReportDataTable.NewRow();
                    NewRow["UNN"] = UNN;
                    NewRow["PaymentRate"] = PaymentRate;
                    NewRow["AccountingName"] = AccountingName;
                    NewRow["InvNumber"] = InvNumber;
                    NewRow["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        NewRow["TPSCurCode"] = TPSCurrencyCode;
                    NewRow["TotalDiscount"] = TotalDiscount;
                    NewRow["Name"] = "Стекло Лакомат";
                    NewRow["Count"] = CountLacomat;
                    NewRow["Measure"] = "м.кв.";
                    NewRow["Price"] = Decimal.Round(PriceLacomat, 2, MidpointRounding.AwayFromZero);
                    NewRow["Cost"] = Decimal.Round(Math.Ceiling(CountLacomat * PriceLacomat / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                    NewRow["PriceWithTransport"] = Decimal.Round(PriceLacomat, 2, MidpointRounding.AwayFromZero);
                    NewRow["CostWithTransport"] = Decimal.Round(Math.Ceiling(CountLacomat * PriceLacomat / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                    NewRow["Weight"] = Decimal.Round(CountLacomat * 10, 3, MidpointRounding.AwayFromZero);
                    ReportDataTable.Rows.Add(NewRow);
                }
            }

            DataRow[] KRows = OrdersDataTable.Select("InsetColorID = 3945");

            if (KRows.Count() > 0)
            {
                PriceKrizet = Convert.ToDecimal(KRows[0]["InsetPrice"]);

                for (int i = 0; i < KRows.Count(); i++)
                {
                    if (FC.IsAluminium(KRows[i]) != -1)
                        continue;

                    TotalDiscount = Convert.ToDecimal(KRows[i]["TotalDiscount"]);
                    CountKrizet += Convert.ToDecimal(KRows[i]["Count"]) *
                                GetInsetSquare(Convert.ToInt32(KRows[i]["FrontID"]),
                                              Convert.ToInt32(KRows[i]["Height"]),
                                              Convert.ToInt32(KRows[i]["Width"]));
                }

                CountKrizet = Decimal.Round(CountKrizet, 3, MidpointRounding.AwayFromZero);

                if (CountKrizet > 0)
                {
                    //decimal Cost = Math.Ceiling(CountKrizet * PriceKrizet / 0.01m) * 0.01m;
                    DataRow NewRow = ReportDataTable.NewRow();
                    NewRow["UNN"] = UNN;
                    NewRow["PaymentRate"] = PaymentRate;
                    NewRow["AccountingName"] = AccountingName;
                    NewRow["InvNumber"] = InvNumber;
                    NewRow["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        NewRow["TPSCurCode"] = TPSCurrencyCode;
                    NewRow["TotalDiscount"] = TotalDiscount;
                    NewRow["Name"] = "Стекло Кризет";
                    NewRow["Count"] = CountKrizet;
                    NewRow["Measure"] = "м.кв.";
                    NewRow["Price"] = Decimal.Round(PriceKrizet, 2, MidpointRounding.AwayFromZero);
                    NewRow["Cost"] = Decimal.Round(Math.Ceiling(CountKrizet * PriceKrizet / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                    NewRow["PriceWithTransport"] = Decimal.Round(PriceKrizet, 2, MidpointRounding.AwayFromZero);
                    NewRow["CostWithTransport"] = Decimal.Round(Math.Ceiling(CountKrizet * PriceKrizet / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                    NewRow["Weight"] = Decimal.Round(CountKrizet * 10, 3, MidpointRounding.AwayFromZero);
                    ReportDataTable.Rows.Add(NewRow);
                }
            }
            DataRow[] ORows = OrdersDataTable.Select("InsetColorID = 18");

            if (ORows.Count() > 0)
            {
                for (int i = 0; i < ORows.Count(); i++)
                {
                    if (FC.IsAluminium(ORows[i]) != -1)
                        continue;

                    TotalDiscount = Convert.ToDecimal(ORows[i]["TotalDiscount"]);
                    CountOther += Convert.ToDecimal(ORows[i]["Count"]) *
                                GetInsetSquare(Convert.ToInt32(ORows[i]["FrontID"]),
                                              Convert.ToInt32(ORows[i]["Height"]),
                                              Convert.ToInt32(ORows[i]["Width"]));
                }

                CountOther = Decimal.Round(CountOther, 3, MidpointRounding.AwayFromZero);

                DataRow NewRow = ReportDataTable.NewRow();
                NewRow["UNN"] = UNN;
                NewRow["PaymentRate"] = PaymentRate;
                NewRow["AccountingName"] = AccountingName;
                NewRow["InvNumber"] = InvNumber;
                NewRow["CurrencyCode"] = ProfilCurrencyCode;
                if (fID == 2)
                    NewRow["TPSCurCode"] = TPSCurrencyCode;
                NewRow["TotalDiscount"] = TotalDiscount;
                NewRow["Name"] = "Стекло другое";
                NewRow["Count"] = CountOther;
                NewRow["Measure"] = "м.кв.";
                NewRow["Price"] = 0;
                NewRow["Cost"] = 0;
                NewRow["PriceWithTransport"] = 0;
                NewRow["CostWithTransport"] = 0;
                NewRow["Weight"] = Decimal.Round(CountOther * 10, 3, MidpointRounding.AwayFromZero);
                ReportDataTable.Rows.Add(NewRow);
            }
        }

        private void GetInsets(DataTable OrdersDataTable, DataTable ReportDataTable)
        {
            string AccountingName = OrdersDataTable.Rows[0]["AccountingName"].ToString();
            string InvNumber = OrdersDataTable.Rows[0]["InvNumber"].ToString();
            int fID = Convert.ToInt32(OrdersDataTable.Rows[0]["FactoryID"]);
            int CountApplic1R = 0;
            int CountApplic1L = 0;
            int CountApplic2 = 0;
            decimal CountEllipseGrid = 0;

            decimal PriceApplic1R = 0;
            decimal PriceApplic1L = 0;
            decimal PriceApplic2 = 0;
            decimal PriceEllipseGrid = 0;
            decimal TotalDiscount = 0;

            DataRow[] EGRows = OrdersDataTable.Select("FrontID IN (3729)");//ellipse grid

            if (EGRows.Count() > 0)
            {
                int MarginHeight = 0;
                int MarginWidth = 0;

                GetGlassMarginAluminium(EGRows[0], ref MarginHeight, ref MarginWidth);
                PriceEllipseGrid = Convert.ToDecimal(EGRows[0]["InsetPrice"]);

                for (int i = 0; i < EGRows.Count(); i++)
                {
                    TotalDiscount = Convert.ToDecimal(EGRows[i]["TotalDiscount"]);
                    decimal dd = Decimal.Round(Convert.ToDecimal(EGRows[i]["Count"]) * MarginHeight * (Convert.ToDecimal(EGRows[i]["Width"]) - MarginWidth) / 1000000, 3, MidpointRounding.AwayFromZero);
                    CountEllipseGrid += dd;
                }
                decimal Weight = GetInsetWeight(EGRows[0]);
                Weight = Decimal.Round(CountEllipseGrid * Weight, 3, MidpointRounding.AwayFromZero);

                //decimal Cost = Math.Ceiling(CountEllipseGrid * PriceEllipseGrid / 0.01m) * 0.01m;
                DataRow NewRow = ReportDataTable.NewRow();
                NewRow["UNN"] = UNN;
                NewRow["PaymentRate"] = PaymentRate;
                NewRow["AccountingName"] = AccountingName;
                NewRow["InvNumber"] = InvNumber;
                NewRow["CurrencyCode"] = ProfilCurrencyCode;
                if (fID == 2)
                    NewRow["TPSCurCode"] = TPSCurrencyCode;
                NewRow["TotalDiscount"] = TotalDiscount;
                NewRow["Name"] = "Решетка овал";
                NewRow["Count"] = CountEllipseGrid;
                NewRow["Measure"] = "м.кв.";
                NewRow["OriginalPrice"] = PriceEllipseGrid;
                NewRow["Price"] = PriceEllipseGrid;
                NewRow["Cost"] = Decimal.Round(Math.Ceiling(CountEllipseGrid * PriceEllipseGrid / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                NewRow["PriceWithTransport"] = PriceEllipseGrid;
                NewRow["CostWithTransport"] = Decimal.Round(Math.Ceiling(CountEllipseGrid * PriceEllipseGrid / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                NewRow["Weight"] = Weight;
                ReportDataTable.Rows.Add(NewRow);
            }

            DataRow[] A1RRows = OrdersDataTable.Select("InsetTypeID = 3654 OR FrontID IN (3731,3740,3746)");//applic 1 right

            if (A1RRows.Count() > 0)
            {
                PriceApplic1R = Convert.ToDecimal(A1RRows[0]["InsetPrice"]);

                for (int i = 0; i < A1RRows.Count(); i++)
                {
                    TotalDiscount = Convert.ToDecimal(A1RRows[i]["TotalDiscount"]);
                    CountApplic1R += Convert.ToInt32(A1RRows[i]["Count"]);
                }

                //decimal Cost = Math.Ceiling(CountApplic1R * PriceApplic1R / 0.01m) * 0.01m;
                DataRow NewRow = ReportDataTable.NewRow();
                NewRow["UNN"] = UNN;
                NewRow["PaymentRate"] = PaymentRate;
                NewRow["AccountingName"] = AccountingName;
                NewRow["InvNumber"] = InvNumber;
                NewRow["CurrencyCode"] = ProfilCurrencyCode;
                if (fID == 2)
                    NewRow["TPSCurCode"] = TPSCurrencyCode;
                NewRow["TotalDiscount"] = TotalDiscount;
                NewRow["Name"] = "Апплик. №1 правая";
                NewRow["Count"] = CountApplic1R;
                NewRow["Measure"] = "шт.";
                NewRow["Price"] = PriceApplic1R;
                NewRow["OriginalPrice"] = PriceApplic1R;
                NewRow["Cost"] = Decimal.Round(Math.Ceiling(CountApplic1R * PriceApplic1R / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                NewRow["PriceWithTransport"] = PriceApplic1R;
                NewRow["CostWithTransport"] = Decimal.Round(Math.Ceiling(CountApplic1R * PriceApplic1R / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                NewRow["Weight"] = 0;
                ReportDataTable.Rows.Add(NewRow);
            }

            DataRow[] A1LRows = OrdersDataTable.Select("InsetTypeID = 3653 OR FrontID IN (3728,3739,3745)");//applic 1 left

            if (A1LRows.Count() > 0)
            {
                PriceApplic1L = Convert.ToDecimal(A1LRows[0]["InsetPrice"]);

                for (int i = 0; i < A1LRows.Count(); i++)
                {
                    TotalDiscount = Convert.ToDecimal(A1LRows[i]["TotalDiscount"]);
                    CountApplic1L += Convert.ToInt32(A1LRows[i]["Count"]);
                }

                //decimal Cost = Math.Ceiling(CountApplic1L * PriceApplic1L / 0.01m) * 0.01m;
                DataRow NewRow = ReportDataTable.NewRow();
                NewRow["UNN"] = UNN;
                NewRow["PaymentRate"] = PaymentRate;
                NewRow["AccountingName"] = AccountingName;
                NewRow["InvNumber"] = InvNumber;
                NewRow["CurrencyCode"] = ProfilCurrencyCode;
                if (fID == 2)
                    NewRow["TPSCurCode"] = TPSCurrencyCode;
                NewRow["TotalDiscount"] = TotalDiscount;
                NewRow["Name"] = "Апплик. №1 левая";
                NewRow["Count"] = CountApplic1L;
                NewRow["Measure"] = "шт.";
                NewRow["Price"] = PriceApplic1L;
                NewRow["OriginalPrice"] = PriceApplic1L;
                NewRow["Cost"] = Decimal.Round(Math.Ceiling(CountApplic1L * PriceApplic1L / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                NewRow["PriceWithTransport"] = PriceApplic1L;
                NewRow["CostWithTransport"] = Decimal.Round(Math.Ceiling(CountApplic1L * PriceApplic1L / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                NewRow["Weight"] = 0;
                ReportDataTable.Rows.Add(NewRow);
            }

            DataRow[] A2Rows = OrdersDataTable.Select("InsetTypeID = 3655 OR FrontID IN (3732,3741,3744)");//applic 2 

            if (A2Rows.Count() > 0)
            {
                PriceApplic2 = Convert.ToDecimal(A2Rows[0]["InsetPrice"]);

                for (int i = 0; i < A2Rows.Count(); i++)
                {
                    TotalDiscount = Convert.ToDecimal(A2Rows[i]["TotalDiscount"]);
                    CountApplic2 += Convert.ToInt32(A2Rows[i]["Count"]);
                }

                //decimal Cost = Math.Ceiling(CountApplic2 * PriceApplic2 / 0.01m) * 0.01m;
                DataRow NewRow = ReportDataTable.NewRow();
                NewRow["UNN"] = UNN;
                NewRow["PaymentRate"] = PaymentRate;
                NewRow["AccountingName"] = AccountingName;
                NewRow["InvNumber"] = InvNumber;
                NewRow["CurrencyCode"] = ProfilCurrencyCode;
                if (fID == 2)
                    NewRow["TPSCurCode"] = TPSCurrencyCode;
                NewRow["Name"] = "Апплик. №2";
                NewRow["Count"] = CountApplic2;
                NewRow["Measure"] = "шт.";
                NewRow["Price"] = PriceApplic2;
                NewRow["OriginalPrice"] = PriceApplic2;
                NewRow["Cost"] = Decimal.Round(Math.Ceiling(CountApplic2 * PriceApplic2 / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                NewRow["PriceWithTransport"] = PriceApplic2;
                NewRow["CostWithTransport"] = Decimal.Round(Math.Ceiling(CountApplic2 * PriceApplic2 / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                NewRow["Weight"] = 0;
                ReportDataTable.Rows.Add(NewRow);
            }

            DataRow[] A3Rows = OrdersDataTable.Select("InsetTypeID = 28961");//applic 3 

            if (A3Rows.Count() > 0)
            {
                PriceApplic2 = Convert.ToDecimal(A3Rows[0]["InsetPrice"]);

                for (int i = 0; i < A3Rows.Count(); i++)
                {
                    TotalDiscount = Convert.ToDecimal(A3Rows[i]["TotalDiscount"]);
                    CountApplic2 += Convert.ToInt32(A3Rows[i]["Count"]);
                }

                //decimal Cost = Math.Ceiling(CountApplic2 * PriceApplic2 / 0.01m) * 0.01m;
                DataRow NewRow = ReportDataTable.NewRow();
                NewRow["UNN"] = UNN;
                NewRow["PaymentRate"] = PaymentRate;
                NewRow["AccountingName"] = AccountingName;
                NewRow["InvNumber"] = InvNumber;
                NewRow["CurrencyCode"] = ProfilCurrencyCode;
                if (fID == 2)
                    NewRow["TPSCurCode"] = TPSCurrencyCode;
                NewRow["Name"] = "Апплик. №3";
                NewRow["Count"] = CountApplic2;
                NewRow["Measure"] = "шт.";
                NewRow["Price"] = PriceApplic2;
                NewRow["OriginalPrice"] = PriceApplic2;
                NewRow["Cost"] = Decimal.Round(Math.Ceiling(CountApplic2 * PriceApplic2 / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                NewRow["PriceWithTransport"] = PriceApplic2;
                NewRow["CostWithTransport"] = Decimal.Round(Math.Ceiling(CountApplic2 * PriceApplic2 / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                NewRow["Weight"] = 0;
                ReportDataTable.Rows.Add(NewRow);
            }
        }

        private decimal GetInsetWeight(DataRow FrontsOrdersRow)
        {
            int InsetTypeID = Convert.ToInt32(FrontsOrdersRow["InsetTypeID"]);
            decimal InsetSquare = GetInsetSquare(FrontsOrdersRow);
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

        private decimal GetProfileWeight(DataRow FrontsOrdersRow)
        {
            decimal FrontHeight = Convert.ToDecimal(FrontsOrdersRow["Height"]);
            decimal FrontWidth = Convert.ToDecimal(FrontsOrdersRow["Width"]);
            decimal ProfileWeight = 0;
            decimal ProfileWidth = 0;
            DataRow[] FrontsConfigRow = FrontsConfigDataTable.Select("FrontConfigID = " + FrontsOrdersRow["FrontConfigID"].ToString());
            if (FrontsConfigRow.Count() > 0)
                ProfileWeight = Convert.ToDecimal(FrontsConfigRow[0]["Weight"]);

            int FrontID = Convert.ToInt32(FrontsOrdersRow["FrontID"]);
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

        public void GetFrontWeight(DataRow FrontsOrdersRow, ref decimal outFrontWeight, ref decimal outInsetWeight)
        {
            //decimal FrontsWeight = 0;
            DataRow[] FrontsConfigRow = FrontsConfigDataTable.Select("FrontConfigID = " + FrontsOrdersRow["FrontConfigID"].ToString());
            decimal InsetWeight = Convert.ToDecimal(FrontsConfigRow[0]["InsetWeight"]);
            decimal FrontsOrderSquare = Convert.ToDecimal(FrontsOrdersRow["Square"]);
            decimal PackWeight = 0;
            if (FrontsOrderSquare > 0)
                PackWeight = FrontsOrderSquare * Convert.ToDecimal(0.7);
            //если гнутый то вес за штуки
            if (FrontsConfigRow[0]["Width"].ToString() == "-1")
            {
                outFrontWeight = PackWeight +
                    Convert.ToDecimal(FrontsOrdersRow["Count"]) * Convert.ToDecimal(FrontsConfigRow[0]["Weight"]);
                return;
            }
            //если алюминий
            if (FC.IsAluminium(FrontsOrdersRow) > -1)
            {
                outFrontWeight = PackWeight +
                    FC.GetAluminiumWeight(FrontsOrdersRow, true);
                return;
            }
            decimal ResultProfileWeight = GetProfileWeight(FrontsOrdersRow);
            decimal ResultInsetWeight = 0;
            DataRow[] rows = InsetTypesDataTable.Select("GroupID = 3 OR GroupID = 4 OR GroupID = 7 OR GroupID = 12 OR GroupID = 13");
            foreach (DataRow item in rows)
            {
                if (FrontsOrdersRow["InsetTypeID"].ToString() == item["InsetTypeID"].ToString())
                {
                    ResultInsetWeight = GetInsetWeight(FrontsOrdersRow);
                }
            }

            outFrontWeight = PackWeight + ResultProfileWeight * Convert.ToDecimal(FrontsOrdersRow["Count"]);
            outInsetWeight = ResultInsetWeight * Convert.ToDecimal(FrontsOrdersRow["Count"]);
        }

        public void Report(int[] MainOrderIDs, DataTable InfoDT, bool IsSample)
        {
            GetMegaOrderInfo(MainOrderIDs[0]);
            ProfilFrontsOrdersDataTable.Clear();
            TPSFrontsOrdersDataTable.Clear();

            string SelectCommand = @"SELECT NewFrontsOrders.*, infiniu2_catalog.dbo.FrontsConfig.AccountingName, infiniu2_catalog.dbo.FrontsConfig.InvNumber, 
                (-NewMegaOrders.ComplaintProfilCost-NewMegaOrders.ComplaintTPSCost+NewMegaOrders.TransportCost+NewMegaOrders.AdditionalCost) AS TotalAdditionalCost,
                NewMegaOrders.ComplaintProfilCost, NewMegaOrders.ComplaintTPSCost, NewMegaOrders.TransportCost, NewMegaOrders.AdditionalCost, NewMegaOrders.MegaOrderID, NewMegaOrders.PaymentRate FROM NewFrontsOrders
                INNER JOIN NewMainOrders ON NewFrontsOrders.MainOrderID = NewMainOrders.MainOrderID
                INNER JOIN NewMegaOrders ON NewMainOrders.MegaOrderID = NewMegaOrders.MegaOrderID
                INNER JOIN infiniu2_catalog.dbo.FrontsConfig ON NewFrontsOrders.FrontConfigID = infiniu2_catalog.dbo.FrontsConfig.FrontConfigID WHERE NewFrontsOrders.MainOrderID IN (" + string.Join(",", MainOrderIDs) + ") AND NewFrontsOrders.IsSample = 1 ORDER BY FrontID";
            if (!IsSample)
                SelectCommand = @"SELECT NewFrontsOrders.*, infiniu2_catalog.dbo.FrontsConfig.AccountingName, infiniu2_catalog.dbo.FrontsConfig.InvNumber, 
                (-NewMegaOrders.ComplaintProfilCost-NewMegaOrders.ComplaintTPSCost+NewMegaOrders.TransportCost+NewMegaOrders.AdditionalCost) AS TotalAdditionalCost,
                NewMegaOrders.ComplaintProfilCost, NewMegaOrders.ComplaintTPSCost, NewMegaOrders.TransportCost, NewMegaOrders.AdditionalCost, NewMegaOrders.MegaOrderID, NewMegaOrders.PaymentRate FROM NewFrontsOrders
                INNER JOIN NewMainOrders ON NewFrontsOrders.MainOrderID = NewMainOrders.MainOrderID
                INNER JOIN NewMegaOrders ON NewMainOrders.MegaOrderID = NewMegaOrders.MegaOrderID
                INNER JOIN infiniu2_catalog.dbo.FrontsConfig ON NewFrontsOrders.FrontConfigID = infiniu2_catalog.dbo.FrontsConfig.FrontConfigID WHERE NewFrontsOrders.MainOrderID IN (" + string.Join(",", MainOrderIDs) + ") AND NewFrontsOrders.IsSample = 0 ORDER BY FrontID";
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
                        for (int j = 0; j < DistRatesDT.Rows.Count; j++)
                        {
                            DT1.Clear();
                            PaymentRate = Convert.ToDecimal(DistRatesDT.Rows[j]["PaymentRate"]);
                            DataRow[] rows = ProfilFrontsOrdersDataTable.Select("PaymentRate='" + PaymentRate.ToString() + "'");
                            foreach (DataRow item in rows)
                                DT1.Rows.Add(item.ItemArray);
                            if (DT1.Rows.Count == 0)
                                continue;

                            GetSimpleFronts(DT1, ProfilReportDataTable, false);
                            GetSimpleFronts(DT1, ProfilReportDataTable, true);

                            GetCurvedFronts(DT1, ProfilReportDataTable);

                            GetGrids(DT1, ProfilReportDataTable);

                            GetInsets(DT1, ProfilReportDataTable);

                            GetGlass(DT1, ProfilReportDataTable);
                        }
                    }

                    DT1.Clear();
                    if (TPSFrontsOrdersDataTable.Rows.Count > 0)
                    {
                        for (int j = 0; j < DistRatesDT.Rows.Count; j++)
                        {
                            DT1.Clear();
                            PaymentRate = Convert.ToDecimal(DistRatesDT.Rows[j]["PaymentRate"]);
                            DataRow[] rows = TPSFrontsOrdersDataTable.Select("PaymentRate='" + PaymentRate.ToString() + "'");
                            foreach (DataRow item in rows)
                                DT1.Rows.Add(item.ItemArray);
                            if (DT1.Rows.Count == 0)
                                continue;

                            GetSimpleFronts(DT1, TPSReportDataTable, false);
                            GetSimpleFronts(DT1, TPSReportDataTable, true);

                            GetCurvedFronts(DT1, TPSReportDataTable);

                            GetGrids(DT1, TPSReportDataTable);

                            GetInsets(DT1, TPSReportDataTable);

                            GetGlass(DT1, TPSReportDataTable);
                        }
                    }
                }
            }

        }

        public void Report(int[] MainOrderIDs, DataTable InfoDT)
        {
            GetMegaOrderInfo(MainOrderIDs[0]);
            ProfilFrontsOrdersDataTable.Clear();
            TPSFrontsOrdersDataTable.Clear();

            string SelectCommand = @"SELECT NewFrontsOrders.*, infiniu2_catalog.dbo.FrontsConfig.AccountingName, infiniu2_catalog.dbo.FrontsConfig.InvNumber, 
                (-NewMegaOrders.ComplaintProfilCost-NewMegaOrders.ComplaintTPSCost+NewMegaOrders.TransportCost+NewMegaOrders.AdditionalCost) AS TotalAdditionalCost,
                NewMegaOrders.ComplaintProfilCost, NewMegaOrders.ComplaintTPSCost, NewMegaOrders.TransportCost, NewMegaOrders.AdditionalCost, NewMegaOrders.MegaOrderID, NewMegaOrders.PaymentRate FROM NewFrontsOrders
                INNER JOIN NewMainOrders ON NewFrontsOrders.MainOrderID = NewMainOrders.MainOrderID
                INNER JOIN NewMegaOrders ON NewMainOrders.MegaOrderID = NewMegaOrders.MegaOrderID
                INNER JOIN infiniu2_catalog.dbo.FrontsConfig ON NewFrontsOrders.FrontConfigID = infiniu2_catalog.dbo.FrontsConfig.FrontConfigID WHERE NewFrontsOrders.MainOrderID IN (" + string.Join(",", MainOrderIDs) + ") ORDER BY FrontID";
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
                        for (int j = 0; j < DistRatesDT.Rows.Count; j++)
                        {
                            DT1.Clear();
                            PaymentRate = Convert.ToDecimal(DistRatesDT.Rows[j]["PaymentRate"]);
                            DataRow[] rows = ProfilFrontsOrdersDataTable.Select("PaymentRate='" + PaymentRate.ToString() + "'");
                            foreach (DataRow item in rows)
                                DT1.Rows.Add(item.ItemArray);
                            if (DT1.Rows.Count == 0)
                                continue;

                            GetSimpleFronts(DT1, ProfilReportDataTable, false);
                            GetSimpleFronts(DT1, ProfilReportDataTable, true);

                            GetCurvedFronts(DT1, ProfilReportDataTable);

                            GetGrids(DT1, ProfilReportDataTable);

                            GetInsets(DT1, ProfilReportDataTable);

                            GetGlass(DT1, ProfilReportDataTable);
                        }
                    }

                    DT1.Clear();
                    if (TPSFrontsOrdersDataTable.Rows.Count > 0)
                    {
                        for (int j = 0; j < DistRatesDT.Rows.Count; j++)
                        {
                            DT1.Clear();
                            PaymentRate = Convert.ToDecimal(DistRatesDT.Rows[j]["PaymentRate"]);
                            DataRow[] rows = TPSFrontsOrdersDataTable.Select("PaymentRate='" + PaymentRate.ToString() + "'");
                            foreach (DataRow item in rows)
                                DT1.Rows.Add(item.ItemArray);
                            if (DT1.Rows.Count == 0)
                                continue;

                            GetSimpleFronts(DT1, TPSReportDataTable, false);
                            GetSimpleFronts(DT1, TPSReportDataTable, true);

                            GetCurvedFronts(DT1, TPSReportDataTable);

                            GetGrids(DT1, TPSReportDataTable);

                            GetInsets(DT1, TPSReportDataTable);

                            GetGlass(DT1, TPSReportDataTable);
                        }
                    }
                }
            }

        }
    }
}
