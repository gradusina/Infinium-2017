using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;

using System;
using System.Collections;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using Infinium.Modules.Marketing.NewOrders;
using System.Data.OleDb;
using System.Text;
using NPOI.HSSF.Record.Formula.Functions;

namespace Infinium.Modules.StatisticsMarketing.Reports
{
    public class ProducedFrontsReport
    {
        //decimal TransportCost = 0;
        //decimal AdditionalCost = 0;
        //decimal Rate = 1;
        //int ClientID = 0;
        //int CurrencyCode = 0;
        private readonly string ProfilCurrencyCode = "0";

        private readonly string TPSCurrencyCode = "0";
        private readonly string UNN = string.Empty;

        private DataTable DecorInvNumbersDT = null;
        private DataTable CurrencyTypesDT;
        private DataTable ProfilFrontsOrdersDataTable = null;
        private DataTable TPSFrontsOrdersDataTable = null;
        public DataTable FrontsDataTable = null;
        public DataTable FrameColorsDataTable = null;
        public DataTable PatinaDataTable = null;
        public DataTable InsetTypesDataTable = null;
        public DataTable InsetColorsDataTable = null;
        private DataTable MeasuresDataTable = null;
        private DataTable FactoryDataTable = null;
        private DataTable GridSizesDataTable = null;
        public DataTable FrontsConfigDataTable = null;
        private DataTable DecorConfigDataTable = null;
        private DataTable TechStoreDataTable = null;
        private DataTable InsetPriceDataTable = null;
        private DataTable AluminiumFrontsDataTable = null;

        public DataTable ProfilReportDataTable = null;
        public DataTable TPSReportDataTable = null;

        private DataTable ProfilMarketingStartCountDT = null;
        private DataTable ProfilMarketingEndCountDT = null;
        private DataTable ProfilMarketingProducedDT = null;
        private DataTable ProfilMarketingDispatchDT = null;
        private DataTable ProfilZOVStartCountDT = null;
        private DataTable ProfilZOVEndCountDT = null;
        private DataTable ProfilZOVProducedDT = null;
        private DataTable ProfilZOVDispatchDT = null;
        public DataTable AllProfilDT = null;

        private DataTable TPSMarketingStartCountDT = null;
        private DataTable TPSMarketingEndCountDT = null;
        private DataTable TPSMarketingProducedDT = null;
        private DataTable TPSMarketingDispatchDT = null;
        private DataTable TPSZOVStartCountDT = null;
        private DataTable TPSZOVEndCountDT = null;
        private DataTable TPSZOVProducedDT = null;
        private DataTable TPSZOVDispatchDT = null;
        public DataTable AllTPSDT = null;

        public ProducedFrontsReport()
        {
            Create();
            CreateReportDataTables();
        }

        private void GetColorsDT()
        {
            FrameColorsDataTable.Columns.Add(new DataColumn("ColorID", Type.GetType("System.Int64")));
            FrameColorsDataTable.Columns.Add(new DataColumn("ColorName", Type.GetType("System.String")));
            FrameColorsDataTable.Columns.Add(new DataColumn("Cvet", Type.GetType("System.String")));
            string SelectCommand = @"SELECT TechStoreID, TechStoreName, Cvet FROM TechStore
                WHERE TechStoreSubGroupID IN (SELECT TechStoreSubGroupID FROM TechStoreSubGroups WHERE (TechStoreGroupID = 11 OR TechStoreGroupID = 1))
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

            FrontsDataTable = new DataTable();
            FrameColorsDataTable = new DataTable();
            PatinaDataTable = new DataTable();
            InsetTypesDataTable = new DataTable();
            InsetColorsDataTable = new DataTable();
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
            SelectCommand = @"SELECT * FROM Patina";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(PatinaDataTable);
            }
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
            ProfilReportDataTable.Clear();
            TPSReportDataTable.Clear();
        }

        private void CreateReportDataTables()
        {
            DecorInvNumbersDT = new DataTable();
            DecorInvNumbersDT.Columns.Add(new DataColumn("FrontsOrdersID", Type.GetType("System.Int32")));
            DecorInvNumbersDT.Columns.Add(new DataColumn("DecorInvNumber", Type.GetType("System.String")));
            DecorInvNumbersDT.Columns.Add(new DataColumn("DecorAccountingName", Type.GetType("System.String")));
            DecorInvNumbersDT.Columns.Add(new DataColumn("Cvet", Type.GetType("System.String")));
            DecorInvNumbersDT.Columns.Add(new DataColumn("Patina", Type.GetType("System.String")));
            DecorInvNumbersDT.Columns.Add(new DataColumn("FactoryID", Type.GetType("System.Int32")));

            ProfilReportDataTable = new DataTable();
            ProfilReportDataTable.Columns.Add(new DataColumn("UNN", Type.GetType("System.String")));
            ProfilReportDataTable.Columns.Add(new DataColumn("CurrencyCode", Type.GetType("System.String")));
            ProfilReportDataTable.Columns.Add(new DataColumn("TPSCurCode", Type.GetType("System.String")));
            ProfilReportDataTable.Columns.Add(new DataColumn("Measure", Type.GetType("System.String")));
            ProfilReportDataTable.Columns.Add(new DataColumn("InvNumber", Type.GetType("System.String")));
            ProfilReportDataTable.Columns.Add(new DataColumn("Cvet", Type.GetType("System.String")));
            ProfilReportDataTable.Columns.Add(new DataColumn("Patina", Type.GetType("System.String")));
            ProfilReportDataTable.Columns.Add(new DataColumn("AccountingName", Type.GetType("System.String")));
            ProfilReportDataTable.Columns.Add(new DataColumn("Count", Type.GetType("System.Decimal")));
            ProfilReportDataTable.Columns.Add(new DataColumn("Price", Type.GetType("System.Decimal")));
            ProfilReportDataTable.Columns.Add(new DataColumn("Cost", Type.GetType("System.Decimal")));
            ProfilReportDataTable.Columns.Add(new DataColumn("Weight", Type.GetType("System.Decimal")));
            ProfilReportDataTable.Columns.Add(new DataColumn("OriginalPrice", Type.GetType("System.Decimal")));
            ProfilReportDataTable.Columns.Add(new DataColumn("MarketingCost", Type.GetType("System.Decimal")));

            AllProfilDT = new DataTable();
            AllProfilDT.Columns.Add(new DataColumn("AccountingName", Type.GetType("System.String")));
            AllProfilDT.Columns.Add(new DataColumn("InvNumber", Type.GetType("System.String")));
            AllProfilDT.Columns.Add(new DataColumn("Measure", Type.GetType("System.String")));
            AllProfilDT.Columns.Add(new DataColumn("StartCount", Type.GetType("System.Decimal")));
            AllProfilDT.Columns.Add(new DataColumn("ProducedCount", Type.GetType("System.Decimal")));
            AllProfilDT.Columns.Add(new DataColumn("DispatchCount", Type.GetType("System.Decimal")));
            AllProfilDT.Columns.Add(new DataColumn("EndCount", Type.GetType("System.Decimal")));
            AllTPSDT = AllProfilDT.Clone();

            TPSReportDataTable = ProfilReportDataTable.Clone();

            ProfilMarketingStartCountDT = ProfilReportDataTable.Clone();
            ProfilMarketingEndCountDT = ProfilReportDataTable.Clone();
            ProfilMarketingProducedDT = ProfilReportDataTable.Clone();
            ProfilMarketingDispatchDT = ProfilReportDataTable.Clone();
            ProfilZOVStartCountDT = ProfilReportDataTable.Clone();
            ProfilZOVEndCountDT = ProfilReportDataTable.Clone();
            ProfilZOVProducedDT = ProfilReportDataTable.Clone();
            ProfilZOVDispatchDT = ProfilReportDataTable.Clone();

            TPSMarketingStartCountDT = ProfilReportDataTable.Clone();
            TPSMarketingEndCountDT = ProfilReportDataTable.Clone();
            TPSMarketingProducedDT = ProfilReportDataTable.Clone();
            TPSMarketingDispatchDT = ProfilReportDataTable.Clone();
            TPSZOVStartCountDT = ProfilReportDataTable.Clone();
            TPSZOVEndCountDT = ProfilReportDataTable.Clone();
            TPSZOVProducedDT = ProfilReportDataTable.Clone();
            TPSZOVDispatchDT = ProfilReportDataTable.Clone();
        }

        public string GetFrontName(int FrontID)
        {
            DataRow[] Row = FrontsDataTable.Select("FrontID = " + FrontID);

            return Row[0]["FrontName"].ToString();
        }

        private string GetColorCode(int ColorID)
        {
            string ColorName = "";
            try
            {
                DataRow[] Rows = FrameColorsDataTable.Select("ColorID = " + ColorID);
                ColorName = Rows[0]["Cvet"].ToString();
            }
            catch
            {
                return "";
            }
            return ColorName;
        }

        private string GetPatinaCode(int PatinaID)
        {
            string PatinaName = "";
            try
            {
                DataRow[] Rows = PatinaDataTable.Select("PatinaID = " + PatinaID);
                PatinaName = Rows[0]["Patina"].ToString();
            }
            catch
            {
                return "";
            }
            return PatinaName;
        }

        public string GetColorNameByCode(string Cvet)
        {
            string ColorName = "";
            try
            {
                DataRow[] Rows = FrameColorsDataTable.Select($"Cvet = '{ Cvet }'");
                ColorName = Rows[0]["ColorName"].ToString();
            }
            catch
            {
                return string.Empty;
            }
            return ColorName;
        }

        public string GetPatinaNameByCode(string Patina)
        {
            string PatinaName = "";
            try
            {
                DataRow[] Rows = PatinaDataTable.Select($"Patina = '{ Patina }'");
                PatinaName = Rows[0]["DisplayName"].ToString();
            }
            catch
            {
                return "";
            }
            return PatinaName;
        }

        private void SplitTables(DataTable FrontsOrdersDataTable, ref DataTable ProfilDT, ref DataTable TPSDT)
        {
            for (int i = 0; i < FrontsOrdersDataTable.Rows.Count; i++)
            {
                int FrontConfigID = Convert.ToInt32(FrontsOrdersDataTable.Rows[i]["FrontConfigID"]);
                DataRow[] Rows = FrontsConfigDataTable.Select("FrontConfigID = " + FrontsOrdersDataTable.Rows[i]["FrontConfigID"].ToString());
                if (Rows.Count() == 0)
                    continue;

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

        //ALUMINIUM
        public int IsAluminium(int FrontID)
        {
            DataRow[] Row = FrontsDataTable.Select("FrontID = " + FrontID);

            if (Row.Count() > 0 && Row[0]["FrontName"].ToString()[0] == 'Z')
            {
                return Convert.ToInt32(Row[0]["FrontID"]);
            }

            return -1;
        }

        //ALUMINIUM
        public int IsAluminium(DataRow FrontsOrdersRow)
        {
            string str = FrontsOrdersRow["FrontID"].ToString();
            DataRow[] Row = FrontsDataTable.Select("FrontID = " + FrontsOrdersRow["FrontID"].ToString());

            if (Row.Count() > 0 && Row[0]["FrontName"].ToString()[0] == 'Z')
            {
                str = Row[0]["FrontName"].ToString();
                return Convert.ToInt32(Row[0]["FrontID"]);
            }

            return -1;
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
            //DataRow[] Rows = AluminiumFrontsDataTable.Select("FrontID = " + IsAluminium(FrontsOrdersRow));
            //if (Rows.Count() > 0)
            //{
            //    GlassMarginHeight = Convert.ToInt32(Rows[0]["GlassMarginHeight"]);
            //    GlassMarginWidth = Convert.ToInt32(Rows[0]["GlassMarginWidth"]);
            //}
        }

        private decimal GetJobPriceAluminium(DataRow FrontsOrdersRow)
        {
            DataRow[] Rows = AluminiumFrontsDataTable.Select("FrontID = " + IsAluminium(FrontsOrdersRow));
            if (Rows.Count() == 0)
                return 0;
            return Convert.ToDecimal(Rows[0]["JobPrice"]);
        }

        private decimal FrontsPriceAluminium(DataRow FrontsOrdersRow)
        {
            decimal Price = 0;

            DataRow[] Rows = AluminiumFrontsDataTable.Select("FrontID = " + IsAluminium(FrontsOrdersRow));
            if (Rows.Count() > 0)
            {
                Price = Convert.ToDecimal(Rows[0]["ProfilPrice"]);
            }
            FrontsOrdersRow["FrontPrice"] = Price;

            return Price;
        }

        private decimal InsetPriceAluminium(DataRow FrontsOrdersRow)
        {
            decimal Price = 0;

            DataRow[] Rows = InsetPriceDataTable.Select("InsetTypeID = " + FrontsOrdersRow["InsetColorID"].ToString());

            if (Rows.Count() > 0)
                Price = Convert.ToDecimal(Rows[0]["GlassZXPrice"]);
            else
                Price = 0;

            FrontsOrdersRow["InsetPrice"] = Price;

            return Price;
        }

        public decimal GetFrontCostAluminium(DataRow FrontsOrdersRow)
        {
            decimal Cost = 0;
            decimal Perimeter = 0;
            decimal GlassSquare = 0;

            int MarginHeight = 0;
            int MarginWidth = 0;

            decimal GlassPrice = InsetPriceAluminium(FrontsOrdersRow);
            decimal JobPrice = GetJobPriceAluminium(FrontsOrdersRow);
            decimal ProfilPrice = FrontsPriceAluminium(FrontsOrdersRow);

            GetGlassMarginAluminium(FrontsOrdersRow, ref MarginHeight, ref MarginWidth);

            decimal Height = Convert.ToInt32(FrontsOrdersRow["Height"]);
            decimal Width = Convert.ToInt32(FrontsOrdersRow["Width"]);
            decimal Count = Convert.ToInt32(FrontsOrdersRow["Count"]);

            Perimeter = Decimal.Round((Height * 2 + Width * 2) / 1000 * Count, 3, MidpointRounding.AwayFromZero);
            GlassSquare = Decimal.Round((Height - MarginHeight) * (Width - MarginWidth) / 1000000, 3, MidpointRounding.AwayFromZero);

            GlassSquare = GlassSquare * Count;
            Cost = Decimal.Round(JobPrice * Count + GlassPrice * GlassSquare + Perimeter * ProfilPrice, 3, MidpointRounding.AwayFromZero);

            Cost = Cost * 100 / 120;

            //FrontsOrdersRow["InsetPrice"] = 0;
            //FrontsOrdersRow["Cost"] = Cost;
            //FrontsOrdersRow["Square"] = (Height * Width * Count) / 1000000;
            //FrontsOrdersRow["FrontPrice"] = Cost / Convert.ToDecimal(FrontsOrdersRow["Square"]);

            return Cost;
        }

        public decimal GetAluminiumWeight(DataRow FrontsOrdersRow, bool WithGlass)
        {
            DataRow[] Row = AluminiumFrontsDataTable.Select("FrontID = " + IsAluminium(FrontsOrdersRow));
            if (Row.Count() == 0)
                return 0;
            decimal FrontHeight = Convert.ToDecimal(FrontsOrdersRow["Height"]);
            decimal FrontWidth = Convert.ToDecimal(FrontsOrdersRow["Width"]);
            decimal Count = Convert.ToInt32(FrontsOrdersRow["Count"]);

            int MarginHeight = 0;
            int MarginWidth = 0;

            GetGlassMarginAluminium(FrontsOrdersRow, ref MarginHeight, ref MarginWidth);

            DataRow[] FrontsConfigRow = FrontsConfigDataTable.Select("FrontConfigID = " + FrontsOrdersRow["FrontConfigID"].ToString());
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

            if (FrontsOrdersRow["InsetColorID"].ToString() != "3946")//если не СТЕКЛО КЛИЕНТА
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

        private decimal GetMarketingCost(int FrontConfigID)
        {
            DataRow[] FRows = FrontsConfigDataTable.Select("FrontConfigID = " + FrontConfigID);
            if (FRows.Count() == 0)
                return 0;
            return Convert.ToDecimal(FRows[0]["MarketingCost"]);
        }

        private decimal GetGridMarketingCost(int FrontConfigID)
        {
            int InsetTypeID = 0;
            DataRow[] FRows = FrontsConfigDataTable.Select("FrontConfigID = " + FrontConfigID);
            if (FRows.Count() == 0)
                return 0;
            InsetTypeID = Convert.ToInt32(FRows[0]["InsetTypeID"]);
            DataRow[] DRows = DecorConfigDataTable.Select("DecorID  = " + InsetTypeID);
            if (DRows.Count() == 0)
                return 0;
            return Convert.ToDecimal(DRows[0]["MarketingCost"]);
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

        private void GetSimpleFronts(DataTable OrdersDataTable, DataTable ReportDataTable)
        {
            string DecorAccountingName = string.Empty;
            string DecorInvNumber = string.Empty;
            string AccountingName = OrdersDataTable.Rows[0]["AccountingName"].ToString();
            string InvNumber = OrdersDataTable.Rows[0]["InvNumber"].ToString();
            string Measure = OrdersDataTable.Rows[0]["Measure"].ToString();
            int fID = Convert.ToInt32(OrdersDataTable.Rows[0]["FactoryID"]);
            DataTable Fronts = new DataTable();

            using (DataView DV = new DataView(OrdersDataTable, "measureId=1", "", DataViewRowState.CurrentRows))
            {
                Fronts = DV.ToTable(true, new string[] { "FrontID", "ColorID", "PatinaID" });
            }

            for (int i = 0; i < Fronts.Rows.Count; i++)
            {
                decimal SolidCount = 0;
                decimal SolidCost = 0;
                decimal SolidWeight = 0;

                decimal FilenkaCount = 0;
                decimal FilenkaCost = 0;
                decimal FilenkaWeight = 0;

                decimal VitrinaCount = 0;
                decimal VitrinaCost = 0;
                decimal VitrinaWeight = 0;

                decimal LuxMegaCount = 0;
                decimal LuxMegaCost = 0;
                decimal LuxMegaWeight = 0;

                decimal MarketingCost = 0;

                //ГЛУХИЕ, БЕЗ ВСТАВКИ, РЕШЕТКА ОВАЛ
                DataRow[] rows = InsetTypesDataTable.Select("InsetTypeID=-1 OR GroupID = 3 OR GroupID = 4");
                string filter = string.Empty;
                foreach (DataRow item in rows)
                    filter += item["InsetTypeID"].ToString() + ",";
                if (filter.Length > 0)
                    filter = " AND NOT (FrontID IN (3728,3731,3732,3739,3740,3741,3744,3745,3746) OR InsetTypeID IN (28961,3653,3654,3655)) AND (FrontID = 3729 OR InsetTypeID IN (" + filter.Substring(0, filter.Length - 1) + "))";
                DataRow[] Rows = OrdersDataTable.Select("ColorID = " + Fronts.Rows[i]["ColorID"].ToString() +
                                                        " AND PatinaID = " + Fronts.Rows[i]["PatinaID"].ToString() +
                                                        " AND FrontID = " + Fronts.Rows[i]["FrontID"].ToString() + " AND (Width <> -1)" + filter);
                if (Rows.Count() > 0)
                    MarketingCost = GetMarketingCost(Convert.ToInt32(Rows[0]["FrontConfigID"]));
                for (int r = 0; r < Rows.Count(); r++)
                {
                    decimal DeductibleCount = 0;
                    decimal DeductibleWeight = 0;
                    if (Convert.ToInt32(Fronts.Rows[i]["FrontID"]) == 3729)//РЕШЕТКА ОВАЛ
                    {
                        int FactoryID = 0;
                        DecorInvNumber = GetGridInvNumber(Convert.ToInt32(Rows[r]["FrontConfigID"]), ref FactoryID, ref DecorAccountingName);
                        if (DecorInvNumber.Length > 0)
                        {
                            DataRow NewRow = DecorInvNumbersDT.NewRow();
                            NewRow["FrontsOrdersID"] = Convert.ToInt32(Rows[r]["FrontsOrdersID"]);
                            NewRow["FactoryID"] = FactoryID;
                            NewRow["DecorAccountingName"] = DecorAccountingName;
                            NewRow["DecorInvNumber"] = DecorInvNumber;
                            NewRow["Cvet"] = GetColorCode(Convert.ToInt32(Fronts.Rows[i]["ColorID"]));
                            NewRow["Patina"] = GetPatinaCode(Convert.ToInt32(Fronts.Rows[i]["PatinaID"]));
                            DecorInvNumbersDT.Rows.Add(NewRow);
                            DeductibleWeight = GetInsetWeight(Rows[r]);
                            DeductibleCount = GetInsetSquare(Convert.ToInt32(Rows[r]["FrontID"]), Convert.ToInt32(Rows[r]["Height"]),
                                                        Convert.ToInt32(Rows[r]["Width"])) * Convert.ToDecimal(Rows[r]["Count"]);
                            DeductibleWeight = Decimal.Round(DeductibleCount * DeductibleWeight, 3, MidpointRounding.AwayFromZero);
                        }
                    }
                    if (GetMeasureType(Convert.ToInt32(Rows[r]["FrontConfigID"])) == 3)
                        SolidCount += Decimal.Round(Convert.ToDecimal(Rows[r]["Height"]) * Convert.ToDecimal(Rows[r]["Width"]) / 1000000, 3, MidpointRounding.AwayFromZero) *
                                      Convert.ToDecimal(Rows[r]["Count"]);
                    else
                        SolidCount += Convert.ToDecimal(Rows[r]["Square"]);

                    if (GetMeasureType(Convert.ToInt32(Rows[r]["FrontConfigID"])) == 3)
                        SolidCost += Convert.ToDecimal(Rows[r]["Cost"]);
                    else
                        SolidCost += Convert.ToDecimal(Rows[r]["FrontPrice"]) * Convert.ToDecimal(Rows[r]["Square"]);

                    decimal FrontWeight = 0;
                    decimal InsetWeight = 0;

                    GetFrontWeight(Rows[r], ref FrontWeight, ref InsetWeight);

                    SolidWeight += Convert.ToDecimal(FrontWeight + InsetWeight - DeductibleWeight);
                    if (MarketingCost > GetMarketingCost(Convert.ToInt32(Rows[r]["FrontConfigID"])))
                        MarketingCost = GetMarketingCost(Convert.ToInt32(Rows[r]["FrontConfigID"]));
                }
                //АППЛИКАЦИИ
                filter = " AND (FrontID IN (3728,3731,3732,3739,3740,3741,3744,3745,3746) OR InsetTypeID IN (28961,3653,3654,3655))";
                Rows = OrdersDataTable.Select("ColorID = " + Fronts.Rows[i]["ColorID"].ToString() +
                                              " AND PatinaID = " + Fronts.Rows[i]["PatinaID"].ToString() +
                                              " AND FrontID = " + Fronts.Rows[i]["FrontID"].ToString() + " AND (Width <> -1)" + filter);
                if (Rows.Count() > 0)
                    MarketingCost = GetMarketingCost(Convert.ToInt32(Rows[0]["FrontConfigID"]));
                for (int r = 0; r < Rows.Count(); r++)
                {
                    if (Convert.ToInt32(Rows[r]["FrontID"]) == 3728 || Convert.ToInt32(Rows[r]["FrontID"]) == 3731 || Convert.ToInt32(Rows[r]["FrontID"]) == 3732 ||
                        Convert.ToInt32(Rows[r]["FrontID"]) == 3739 || Convert.ToInt32(Rows[r]["FrontID"]) == 3740 || Convert.ToInt32(Rows[r]["FrontID"]) == 3741 ||
                        Convert.ToInt32(Rows[r]["FrontID"]) == 3744 || Convert.ToInt32(Rows[r]["FrontID"]) == 3745 || Convert.ToInt32(Rows[r]["FrontID"]) == 3746)
                    {
                        if (GetMeasureType(Convert.ToInt32(Rows[r]["FrontConfigID"])) == 3)
                            SolidCount += Decimal.Round(Convert.ToDecimal(Rows[r]["Height"]) * Convert.ToDecimal(Rows[r]["Width"]) / 1000000, 3, MidpointRounding.AwayFromZero) *
                                          Convert.ToDecimal(Rows[r]["Count"]);
                        else
                            SolidCount += Convert.ToDecimal(Rows[r]["Square"]);

                        if (GetMeasureType(Convert.ToInt32(Rows[r]["FrontConfigID"])) == 3)
                            SolidCost += Convert.ToDecimal(Rows[r]["Cost"]) + 5;
                        else
                            SolidCost += Convert.ToDecimal(Rows[r]["FrontPrice"]) * Convert.ToDecimal(Rows[r]["Square"]) + 5 * Convert.ToDecimal(Rows[r]["Count"]);

                        decimal FrontWeight = 0;
                        decimal InsetWeight = 0;

                        GetFrontWeight(Rows[r], ref FrontWeight, ref InsetWeight);

                        SolidWeight += Convert.ToDecimal(FrontWeight + InsetWeight);
                    }
                    else if (Convert.ToInt32(Rows[r]["FrontID"]) == 3415 || Convert.ToInt32(Rows[r]["FrontID"]) == 28922)
                    {
                        if (GetMeasureType(Convert.ToInt32(Rows[r]["FrontConfigID"])) == 3)
                            FilenkaCount += Decimal.Round(Convert.ToDecimal(Rows[r]["Height"]) * Convert.ToDecimal(Rows[r]["Width"]) / 1000000, 3, MidpointRounding.AwayFromZero) *
                                          Convert.ToDecimal(Rows[r]["Count"]);
                        else
                            FilenkaCount += Convert.ToDecimal(Rows[r]["Square"]);

                        if (GetMeasureType(Convert.ToInt32(Rows[r]["FrontConfigID"])) == 3)
                            FilenkaCost += Convert.ToDecimal(Rows[r]["Cost"]) + 5;
                        else
                            FilenkaCost += Convert.ToDecimal(Rows[r]["FrontPrice"]) * Convert.ToDecimal(Rows[r]["Square"]) + 5 * Convert.ToDecimal(Rows[r]["Count"]);

                        decimal FrontWeight = 0;
                        decimal InsetWeight = 0;

                        GetFrontWeight(Rows[r], ref FrontWeight, ref InsetWeight);

                        FilenkaWeight += Convert.ToDecimal(FrontWeight + InsetWeight);
                    }
                    if (MarketingCost > GetMarketingCost(Convert.ToInt32(Rows[r]["FrontConfigID"])))
                        MarketingCost = GetMarketingCost(Convert.ToInt32(Rows[r]["FrontConfigID"]));
                }
                //ФИЛЕНКА
                filter = " AND InsetTypeID IN (2069,2070,2071,2073,2075,42066,2077,2233,3644,29043,29531,41213)";
                Rows = OrdersDataTable.Select("ColorID = " + Fronts.Rows[i]["ColorID"].ToString() +
                                              " AND PatinaID = " + Fronts.Rows[i]["PatinaID"].ToString() +
                                              " AND FrontID = " + Fronts.Rows[i]["FrontID"].ToString() + " AND (Width <> -1)" + filter);
                if (Rows.Count() > 0)
                    MarketingCost = GetMarketingCost(Convert.ToInt32(Rows[0]["FrontConfigID"]));
                for (int r = 0; r < Rows.Count(); r++)
                {
                    if (GetMeasureType(Convert.ToInt32(Rows[r]["FrontConfigID"])) == 3)
                        FilenkaCount += Decimal.Round(Convert.ToDecimal(Rows[r]["Height"]) * Convert.ToDecimal(Rows[r]["Width"]) / 1000000, 3, MidpointRounding.AwayFromZero) *
                                      Convert.ToDecimal(Rows[r]["Count"]);
                    else
                        FilenkaCount += Convert.ToDecimal(Rows[r]["Square"]);

                    if (GetMeasureType(Convert.ToInt32(Rows[r]["FrontConfigID"])) == 3)
                        FilenkaCost += Convert.ToDecimal(Rows[r]["Cost"]);
                    else
                        FilenkaCost += Convert.ToDecimal(Rows[r]["FrontPrice"]) * Convert.ToDecimal(Rows[r]["Square"]);

                    decimal FrontWeight = 0;
                    decimal InsetWeight = 0;

                    GetFrontWeight(Rows[r], ref FrontWeight, ref InsetWeight);

                    FilenkaWeight += Convert.ToDecimal(FrontWeight + InsetWeight);
                    if (MarketingCost > GetMarketingCost(Convert.ToInt32(Rows[r]["FrontConfigID"])))
                        MarketingCost = GetMarketingCost(Convert.ToInt32(Rows[r]["FrontConfigID"]));
                }
                //ВИТРИНЫ, РЕШЕТКИ, СТЕКЛО
                filter = " AND InsetTypeID IN (1,2,685,686,687,688,29470,29471) AND FrontID <> 3729";
                Rows = OrdersDataTable.Select("ColorID = " + Fronts.Rows[i]["ColorID"].ToString() +
                                              " AND PatinaID = " + Fronts.Rows[i]["PatinaID"].ToString() +
                                              " AND FrontID = " + Fronts.Rows[i]["FrontID"].ToString() + " AND (Width <> -1)" + filter);
                if (Rows.Count() > 0)
                    MarketingCost = GetMarketingCost(Convert.ToInt32(Rows[0]["FrontConfigID"]));
                for (int r = 0; r < Rows.Count(); r++)
                {
                    decimal DeductibleCount = 0;
                    decimal DeductibleWeight = 0;
                    //РЕШЕТКА 45,90,ПЛАСТИК
                    if (Convert.ToInt32(Rows[r]["InsetTypeID"]) == 685 || Convert.ToInt32(Rows[r]["InsetTypeID"]) == 686 ||
                        Convert.ToInt32(Rows[r]["InsetTypeID"]) == 687 || Convert.ToInt32(Rows[r]["InsetTypeID"]) == 688 ||
                        Convert.ToInt32(Rows[r]["InsetTypeID"]) == 29470 || Convert.ToInt32(Rows[r]["InsetTypeID"]) == 29471)
                    {
                        int FactoryID = 0;
                        int FrontConfigID = Convert.ToInt32(Rows[r]["FrontConfigID"]);
                        DecorInvNumber = GetGridInvNumber(Convert.ToInt32(Rows[r]["FrontConfigID"]), ref FactoryID, ref DecorAccountingName);
                        if (DecorInvNumber.Length > 0)
                        {
                            DataRow NewRow = DecorInvNumbersDT.NewRow();
                            NewRow["FrontsOrdersID"] = Convert.ToInt32(Rows[r]["FrontsOrdersID"]);
                            NewRow["FactoryID"] = FactoryID;
                            NewRow["DecorAccountingName"] = DecorAccountingName;
                            NewRow["DecorInvNumber"] = DecorInvNumber;
                            NewRow["Cvet"] = GetColorCode(Convert.ToInt32(Fronts.Rows[i]["ColorID"]));
                            NewRow["Patina"] = GetPatinaCode(Convert.ToInt32(Fronts.Rows[i]["PatinaID"]));
                            DecorInvNumbersDT.Rows.Add(NewRow);
                            DeductibleCount = GetInsetSquare(Convert.ToInt32(Rows[r]["FrontID"]), Convert.ToInt32(Rows[r]["Height"]),
                                                        Convert.ToInt32(Rows[r]["Width"])) * Convert.ToDecimal(Rows[r]["Count"]);
                            DeductibleWeight = Decimal.Round(DeductibleCount * Convert.ToDecimal(3.5), 3, MidpointRounding.AwayFromZero);
                        }
                        else
                        {

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
                            NewRow["FrontsOrdersID"] = Convert.ToInt32(Rows[r]["FrontsOrdersID"]);
                            NewRow["FactoryID"] = FactoryID;
                            NewRow["DecorAccountingName"] = DecorAccountingName;
                            NewRow["DecorInvNumber"] = DecorInvNumber;
                            NewRow["Cvet"] = GetColorCode(Convert.ToInt32(Fronts.Rows[i]["ColorID"]));
                            NewRow["Patina"] = GetPatinaCode(Convert.ToInt32(Fronts.Rows[i]["PatinaID"]));
                            DecorInvNumbersDT.Rows.Add(NewRow);
                            DeductibleCount = Convert.ToDecimal(Rows[r]["Count"]) * GetInsetSquare(Convert.ToInt32(Rows[r]["FrontID"]), Convert.ToInt32(Rows[r]["Height"]), Convert.ToInt32(Rows[r]["Width"]));
                            DeductibleWeight = Decimal.Round(DeductibleCount * 10, 3, MidpointRounding.AwayFromZero);
                        }
                    }

                    if (GetMeasureType(Convert.ToInt32(Rows[r]["FrontConfigID"])) == 3)
                        VitrinaCount += Decimal.Round(Convert.ToDecimal(Rows[r]["Height"]) * Convert.ToDecimal(Rows[r]["Width"]) / 1000000, 3, MidpointRounding.AwayFromZero) *
                                      Convert.ToDecimal(Rows[r]["Count"]);
                    else
                        VitrinaCount += Convert.ToDecimal(Rows[r]["Square"]);

                    if (IsAluminium(Rows[r]) > -1)
                    {
                        VitrinaCost += GetFrontCostAluminium(Rows[r]);
                        VitrinaWeight += GetAluminiumWeight(Rows[r], true);
                    }
                    else
                    {
                        if (GetMeasureType(Convert.ToInt32(Rows[r]["FrontConfigID"])) == 3)
                            VitrinaCost += Convert.ToDecimal(Rows[r]["Cost"]);
                        else
                            VitrinaCost += Convert.ToDecimal(Rows[r]["FrontPrice"]) * Convert.ToDecimal(Rows[r]["Square"]);

                        decimal FrontWeight = 0;
                        decimal InsetWeight = 0;

                        GetFrontWeight(Rows[r], ref FrontWeight, ref InsetWeight);

                        VitrinaWeight += Convert.ToDecimal(FrontWeight + InsetWeight - DeductibleWeight);
                    }
                    if (MarketingCost > GetMarketingCost(Convert.ToInt32(Rows[r]["FrontConfigID"])))
                        MarketingCost = GetMarketingCost(Convert.ToInt32(Rows[r]["FrontConfigID"]));
                }
                //ЛЮКС, МЕГА
                filter = " AND InsetTypeID IN (860,862,4310)";
                Rows = OrdersDataTable.Select("ColorID = " + Fronts.Rows[i]["ColorID"].ToString() +
                                              " AND PatinaID = " + Fronts.Rows[i]["PatinaID"].ToString() +
                                              " AND FrontID = " + Fronts.Rows[i]["FrontID"].ToString() + " AND (Width <> -1)" + filter);
                if (Rows.Count() > 0)
                    MarketingCost = GetMarketingCost(Convert.ToInt32(Rows[0]["FrontConfigID"]));
                for (int r = 0; r < Rows.Count(); r++)
                {
                    if (GetMeasureType(Convert.ToInt32(Rows[r]["FrontConfigID"])) == 3)
                        LuxMegaCount += Decimal.Round(Convert.ToDecimal(Rows[r]["Height"]) * Convert.ToDecimal(Rows[r]["Width"]) / 1000000, 3, MidpointRounding.AwayFromZero) *
                                      Convert.ToDecimal(Rows[r]["Count"]);
                    else
                        LuxMegaCount += Convert.ToDecimal(Rows[r]["Square"]);

                    if (GetMeasureType(Convert.ToInt32(Rows[r]["FrontConfigID"])) == 3)
                        LuxMegaCost += Convert.ToDecimal(Rows[r]["Cost"]);
                    else
                        LuxMegaCost += Convert.ToDecimal(Rows[r]["FrontPrice"]) * Convert.ToDecimal(Rows[r]["Square"]);

                    decimal FrontWeight = 0;
                    decimal InsetWeight = 0;

                    GetFrontWeight(Rows[r], ref FrontWeight, ref InsetWeight);

                    LuxMegaWeight += Convert.ToDecimal(FrontWeight + InsetWeight);
                    if (MarketingCost > GetMarketingCost(Convert.ToInt32(Rows[r]["FrontConfigID"])))
                        MarketingCost = GetMarketingCost(Convert.ToInt32(Rows[r]["FrontConfigID"]));
                }

                if (SolidCount > 0)
                {
                    DataRow Row = ReportDataTable.NewRow();
                    Row["OriginalPrice"] = SolidCost / SolidCount;
                    Row["UNN"] = UNN;
                    Row["AccountingName"] = AccountingName;
                    Row["InvNumber"] = InvNumber;
                    Row["Cvet"] = GetColorCode(Convert.ToInt32(Fronts.Rows[i]["ColorID"]));
                    Row["Patina"] = GetPatinaCode(Convert.ToInt32(Fronts.Rows[i]["PatinaID"]));
                    Row["MarketingCost"] = MarketingCost;
                    Row["Measure"] = Measure;
                    Row["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        Row["TPSCurCode"] = TPSCurrencyCode;
                    Row["Count"] = Decimal.Round(SolidCount, 3, MidpointRounding.AwayFromZero);
                    Row["Cost"] = Decimal.Round(SolidCost, 3, MidpointRounding.AwayFromZero);
                    Row["Weight"] = Decimal.Round(SolidWeight, 3, MidpointRounding.AwayFromZero);
                    ReportDataTable.Rows.Add(Row);
                }

                if (FilenkaCount > 0)
                {
                    DataRow Row = ReportDataTable.NewRow();
                    Row["OriginalPrice"] = FilenkaCost / FilenkaCount;
                    Row["UNN"] = UNN;
                    Row["AccountingName"] = AccountingName;
                    Row["InvNumber"] = InvNumber;
                    Row["Cvet"] = GetColorCode(Convert.ToInt32(Fronts.Rows[i]["ColorID"]));
                    Row["Patina"] = GetPatinaCode(Convert.ToInt32(Fronts.Rows[i]["PatinaID"]));
                    Row["MarketingCost"] = MarketingCost;
                    Row["Measure"] = Measure;
                    Row["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        Row["TPSCurCode"] = TPSCurrencyCode;
                    Row["Count"] = Decimal.Round(FilenkaCount, 3, MidpointRounding.AwayFromZero);
                    Row["Cost"] = Decimal.Round(FilenkaCost, 3, MidpointRounding.AwayFromZero);
                    Row["Weight"] = Decimal.Round(FilenkaWeight, 3, MidpointRounding.AwayFromZero);
                    ReportDataTable.Rows.Add(Row);
                }

                if (VitrinaCount > 0)
                {
                    DataRow Row = ReportDataTable.NewRow();
                    Row["OriginalPrice"] = VitrinaCost / VitrinaCount;
                    Row["UNN"] = UNN;
                    Row["AccountingName"] = AccountingName;
                    Row["InvNumber"] = InvNumber;
                    Row["Cvet"] = GetColorCode(Convert.ToInt32(Fronts.Rows[i]["ColorID"]));
                    Row["Patina"] = GetPatinaCode(Convert.ToInt32(Fronts.Rows[i]["PatinaID"]));
                    Row["MarketingCost"] = MarketingCost;
                    Row["Measure"] = Measure;
                    Row["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        Row["TPSCurCode"] = TPSCurrencyCode;
                    Row["Count"] = Decimal.Round(VitrinaCount, 3, MidpointRounding.AwayFromZero);
                    Row["Cost"] = Decimal.Round(VitrinaCost, 3, MidpointRounding.AwayFromZero);
                    Row["Weight"] = Decimal.Round(VitrinaWeight, 3, MidpointRounding.AwayFromZero);
                    ReportDataTable.Rows.Add(Row);
                }

                if (LuxMegaCount > 0)
                {
                    DataRow Row = ReportDataTable.NewRow();
                    Row["OriginalPrice"] = LuxMegaCost / LuxMegaCount;
                    Row["UNN"] = UNN;
                    Row["AccountingName"] = AccountingName;
                    Row["InvNumber"] = InvNumber;
                    Row["Cvet"] = GetColorCode(Convert.ToInt32(Fronts.Rows[i]["ColorID"]));
                    Row["Patina"] = GetPatinaCode(Convert.ToInt32(Fronts.Rows[i]["PatinaID"]));
                    Row["MarketingCost"] = MarketingCost;
                    Row["Measure"] = Measure;
                    Row["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        Row["TPSCurCode"] = TPSCurrencyCode;
                    Row["Count"] = Decimal.Round(LuxMegaCount, 3, MidpointRounding.AwayFromZero);
                    Row["Cost"] = Decimal.Round(LuxMegaCost, 3, MidpointRounding.AwayFromZero);
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
            string Measure = OrdersDataTable.Rows[0]["Measure"].ToString();
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
                decimal MarketingCost = 0;
                decimal Solid713Count = 0;
                decimal Solid713Price = 0;
                decimal Solid713Weight = 0;

                decimal Filenka713Count = 0;
                decimal Filenka713Price = 0;
                decimal Filenka713Weight = 0;

                decimal NoInset713Count = 0;
                decimal NoInset713Price = 0;
                decimal NoInset713Weight = 0;

                decimal Vitrina713Count = 0;
                decimal Vitrina713Price = 0;
                decimal Vitrina713Weight = 0;

                decimal Solid910Count = 0;
                decimal Solid910Price = 0;
                decimal Solid910Weight = 0;

                decimal Filenka910Count = 0;
                decimal Filenka910Price = 0;
                decimal Filenka910Weight = 0;

                decimal NoInset910Count = 0;
                decimal NoInset910Price = 0;
                decimal NoInset910Weight = 0;

                decimal Vitrina910Count = 0;
                decimal Vitrina910Price = 0;
                decimal Vitrina910Weight = 0;

                decimal ScagenCount = 0;
                decimal ScagenPrice = 0;
                decimal ScagenWeight = 0;

                if (Rows.Count() > 0)
                    MarketingCost = GetMarketingCost(Convert.ToInt32(Rows[0]["FrontConfigID"]));
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

                            decimal FrontWeight = 0;
                            decimal InsetWeight = 0;

                            GetFrontWeight(Rows[r], ref FrontWeight, ref InsetWeight);

                            ScagenWeight += Convert.ToDecimal(FrontWeight + InsetWeight);
                        }

                        if (Rows[r]["InsetTypeID"].ToString() == "1")
                        {
                            ScagenCount += Convert.ToDecimal(Rows[r]["Count"]);
                            ScagenPrice = Convert.ToDecimal(Rows[r]["FrontPrice"]);

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

                            decimal FrontWeight = 0;
                            decimal InsetWeight = 0;

                            GetFrontWeight(Rows[r], ref FrontWeight, ref InsetWeight);

                            NoInset713Weight += Convert.ToDecimal(FrontWeight + InsetWeight);
                        }

                        if (Rows[r]["InsetTypeID"].ToString() == "1")
                        {
                            Vitrina713Count += Convert.ToDecimal(Rows[r]["Count"]);
                            Vitrina713Price = Convert.ToDecimal(Rows[r]["FrontPrice"]);

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

                            decimal FrontWeight = 0;
                            decimal InsetWeight = 0;

                            GetFrontWeight(Rows[r], ref FrontWeight, ref InsetWeight);

                            NoInset910Weight += Convert.ToDecimal(FrontWeight + InsetWeight);
                        }

                        if (Rows[r]["InsetTypeID"].ToString() == "1")
                        {
                            Vitrina910Count += Convert.ToDecimal(Rows[r]["Count"]);
                            Vitrina910Price = Convert.ToDecimal(Rows[r]["FrontPrice"]);

                            decimal FrontWeight = 0;
                            decimal InsetWeight = 0;

                            GetFrontWeight(Rows[r], ref FrontWeight, ref InsetWeight);

                            Vitrina910Weight += Convert.ToDecimal(FrontWeight + InsetWeight);
                        }
                    }

                    if (MarketingCost > GetMarketingCost(Convert.ToInt32(Rows[r]["FrontConfigID"])))
                        MarketingCost = GetMarketingCost(Convert.ToInt32(Rows[r]["FrontConfigID"]));
                }

                if (ScagenCount > 0)
                {
                    DataRow Row = ReportDataTable.NewRow();
                    Row["OriginalPrice"] = ScagenPrice;
                    Row["UNN"] = UNN;
                    Row["batchId"] = Convert.ToInt32(Fronts.Rows[i]["batchId"]);
                    Row["IsComplaint"] = Convert.ToInt32(Fronts.Rows[i]["IsComplaint"]);
                    Row["Cvet"] = GetColorCode(Convert.ToInt32(Fronts.Rows[i]["ColorID"]));
                    Row["Patina"] = GetPatinaCode(Convert.ToInt32(Fronts.Rows[i]["PatinaID"]));
                    Row["AccountingName"] = AccountingName;
                    Row["InvNumber"] = InvNumber;
                    Row["MarketingCost"] = MarketingCost;
                    Row["Measure"] = Measure;
                    Row["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        Row["TPSCurCode"] = TPSCurrencyCode;
                    Row["Count"] = ScagenCount;
                    Row["Cost"] = Decimal.Round(ScagenCount * ScagenPrice, 3, MidpointRounding.AwayFromZero);
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
                    Row["AccountingName"] = AccountingName;
                    Row["InvNumber"] = InvNumber;
                    Row["MarketingCost"] = MarketingCost;
                    Row["Measure"] = Measure;
                    Row["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        Row["TPSCurCode"] = TPSCurrencyCode;
                    Row["Count"] = Solid713Count;
                    Row["Cost"] = Decimal.Round(Solid713Count * Solid713Price, 3, MidpointRounding.AwayFromZero);
                    Row["Weight"] = Decimal.Round(Solid713Weight, 3, MidpointRounding.AwayFromZero);
                    ReportDataTable.Rows.Add(Row);
                }

                if (Filenka713Count > 0)
                {
                    DataRow Row = ReportDataTable.NewRow();
                    Row["OriginalPrice"] = Filenka713Price;
                    Row["UNN"] = UNN;
                    Row["Cvet"] = GetColorCode(Convert.ToInt32(Fronts.Rows[i]["ColorID"]));
                    Row["Patina"] = GetPatinaCode(Convert.ToInt32(Fronts.Rows[i]["PatinaID"]));
                    Row["AccountingName"] = AccountingName;
                    Row["InvNumber"] = InvNumber;
                    Row["MarketingCost"] = MarketingCost;
                    Row["Measure"] = Measure;
                    Row["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        Row["TPSCurCode"] = TPSCurrencyCode;
                    Row["Count"] = Filenka713Count;
                    Row["Cost"] = Decimal.Round(Filenka713Count * Filenka713Price, 3, MidpointRounding.AwayFromZero);
                    Row["Weight"] = Decimal.Round(Filenka713Weight, 3, MidpointRounding.AwayFromZero);
                    ReportDataTable.Rows.Add(Row);
                }

                if (NoInset713Count > 0)
                {
                    DataRow Row = ReportDataTable.NewRow();
                    Row["OriginalPrice"] = NoInset713Price;
                    Row["UNN"] = UNN;
                    Row["Cvet"] = GetColorCode(Convert.ToInt32(Fronts.Rows[i]["ColorID"]));
                    Row["Patina"] = GetPatinaCode(Convert.ToInt32(Fronts.Rows[i]["PatinaID"]));
                    Row["AccountingName"] = AccountingName;
                    Row["InvNumber"] = InvNumber;
                    Row["MarketingCost"] = MarketingCost;
                    Row["Measure"] = Measure;
                    Row["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        Row["TPSCurCode"] = TPSCurrencyCode;
                    Row["Count"] = NoInset713Count;
                    Row["Cost"] = Decimal.Round(NoInset713Count * NoInset713Price, 3, MidpointRounding.AwayFromZero);
                    Row["Weight"] = Decimal.Round(NoInset713Weight, 3, MidpointRounding.AwayFromZero);
                    ReportDataTable.Rows.Add(Row);
                }

                if (Vitrina713Count > 0)
                {
                    DataRow Row = ReportDataTable.NewRow();
                    Row["OriginalPrice"] = Vitrina713Price;
                    Row["UNN"] = UNN;
                    Row["Cvet"] = GetColorCode(Convert.ToInt32(Fronts.Rows[i]["ColorID"]));
                    Row["Patina"] = GetPatinaCode(Convert.ToInt32(Fronts.Rows[i]["PatinaID"]));
                    Row["AccountingName"] = AccountingName;
                    Row["InvNumber"] = InvNumber;
                    Row["MarketingCost"] = MarketingCost;
                    Row["Measure"] = Measure;
                    Row["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        Row["TPSCurCode"] = TPSCurrencyCode;
                    Row["Count"] = Vitrina713Count;
                    Row["Cost"] = Decimal.Round(Vitrina713Count * Vitrina713Price, 3, MidpointRounding.AwayFromZero);
                    Row["Weight"] = Decimal.Round(Vitrina713Weight, 3, MidpointRounding.AwayFromZero);
                    ReportDataTable.Rows.Add(Row);
                }

                if (Solid910Count > 0)
                {
                    DataRow Row = ReportDataTable.NewRow();
                    Row["OriginalPrice"] = Solid910Price;
                    Row["UNN"] = UNN;
                    Row["Cvet"] = GetColorCode(Convert.ToInt32(Fronts.Rows[i]["ColorID"]));
                    Row["Patina"] = GetPatinaCode(Convert.ToInt32(Fronts.Rows[i]["PatinaID"]));
                    Row["AccountingName"] = AccountingName;
                    Row["InvNumber"] = InvNumber;
                    Row["MarketingCost"] = MarketingCost;
                    Row["Measure"] = Measure;
                    Row["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        Row["TPSCurCode"] = TPSCurrencyCode;
                    Row["Count"] = Solid910Count;
                    Row["Cost"] = Decimal.Round(Solid910Count * Solid910Price, 3, MidpointRounding.AwayFromZero);
                    Row["Weight"] = Decimal.Round(Solid910Weight, 3, MidpointRounding.AwayFromZero);
                    ReportDataTable.Rows.Add(Row);
                }

                if (Filenka910Count > 0)
                {
                    DataRow Row = ReportDataTable.NewRow();
                    Row["OriginalPrice"] = Filenka910Price;
                    Row["UNN"] = UNN;
                    Row["Cvet"] = GetColorCode(Convert.ToInt32(Fronts.Rows[i]["ColorID"]));
                    Row["Patina"] = GetPatinaCode(Convert.ToInt32(Fronts.Rows[i]["PatinaID"]));
                    Row["AccountingName"] = AccountingName;
                    Row["InvNumber"] = InvNumber;
                    Row["MarketingCost"] = MarketingCost;
                    Row["Measure"] = Measure;
                    Row["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        Row["TPSCurCode"] = TPSCurrencyCode;
                    Row["Count"] = Filenka910Count;
                    Row["Cost"] = Decimal.Round(Filenka910Count * Filenka910Price, 3, MidpointRounding.AwayFromZero);
                    Row["Weight"] = Decimal.Round(Filenka910Weight, 3, MidpointRounding.AwayFromZero);
                    ReportDataTable.Rows.Add(Row);
                }

                if (NoInset910Count > 0)
                {
                    DataRow Row = ReportDataTable.NewRow();
                    Row["OriginalPrice"] = NoInset910Price;
                    Row["UNN"] = UNN;
                    Row["Cvet"] = GetColorCode(Convert.ToInt32(Fronts.Rows[i]["ColorID"]));
                    Row["Patina"] = GetPatinaCode(Convert.ToInt32(Fronts.Rows[i]["PatinaID"]));
                    Row["AccountingName"] = AccountingName;
                    Row["InvNumber"] = InvNumber;
                    Row["MarketingCost"] = MarketingCost;
                    Row["Measure"] = Measure;
                    Row["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        Row["TPSCurCode"] = TPSCurrencyCode;
                    Row["Count"] = NoInset910Count;
                    Row["Cost"] = Decimal.Round(NoInset910Count * NoInset910Price, 3, MidpointRounding.AwayFromZero);
                    Row["Weight"] = Decimal.Round(NoInset910Weight, 3, MidpointRounding.AwayFromZero);
                    ReportDataTable.Rows.Add(Row);
                }

                if (Vitrina910Count > 0)
                {
                    DataRow Row = ReportDataTable.NewRow();
                    Row["OriginalPrice"] = Vitrina910Price;
                    Row["UNN"] = UNN;
                    Row["Cvet"] = GetColorCode(Convert.ToInt32(Fronts.Rows[i]["ColorID"]));
                    Row["Patina"] = GetPatinaCode(Convert.ToInt32(Fronts.Rows[i]["PatinaID"]));
                    Row["AccountingName"] = AccountingName;
                    Row["InvNumber"] = InvNumber;
                    Row["MarketingCost"] = MarketingCost;
                    Row["Measure"] = Measure;
                    Row["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        Row["TPSCurCode"] = TPSCurrencyCode;
                    Row["Count"] = Vitrina910Count;
                    Row["Cost"] = Decimal.Round(Vitrina910Count * Vitrina910Price, 3, MidpointRounding.AwayFromZero);
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

                decimal MarketingCost = 0;
                decimal CountPP = 0;
                decimal CostPP = 0;

                if (Rows.Count() > 0)
                    MarketingCost = GetGridMarketingCost(Convert.ToInt32(Rows[0]["FrontConfigID"]));

                for (int j = 0; j < Rows.Count(); j++)
                {
                    decimal d = GetInsetSquare(Convert.ToInt32(Rows[j]["FrontID"]), Convert.ToInt32(Rows[j]["Height"]),
                        Convert.ToInt32(Rows[j]["Width"])) * Convert.ToDecimal(Rows[j]["Count"]);
                    CountPP += d;
                    CostPP += Math.Ceiling(Convert.ToDecimal(Rows[j]["InsetPrice"]) * d / 0.01m) * 0.01m;
                    if (MarketingCost > GetGridMarketingCost(Convert.ToInt32(Rows[j]["FrontConfigID"])))
                        MarketingCost = GetGridMarketingCost(Convert.ToInt32(Rows[j]["FrontConfigID"]));
                }

                if (CountPP > 0)
                {
                    DataTable ddt = OrdersDataTable.Select("InsetTypeID IN (685,686,687,688,29470,29471)").CopyToDataTable();
                    for (int x = 0; x < ddt.Rows.Count; x++)
                    {
                        decimal d = GetInsetSquare(Convert.ToInt32(ddt.Rows[x]["FrontID"]), Convert.ToInt32(ddt.Rows[x]["Height"]),
                            Convert.ToInt32(ddt.Rows[x]["Width"])) * Convert.ToDecimal(ddt.Rows[x]["Count"]);
                        CountPP = d;
                        CostPP = Math.Ceiling(Convert.ToDecimal(ddt.Rows[x]["InsetPrice"]) * d / 0.01m) * 0.01m;

                        int FrontsOrdersID = Convert.ToInt32(ddt.Rows[x]["FrontsOrdersID"]);
                        DataRow[] rows = DecorInvNumbersDT.Select("FrontsOrdersID = " + FrontsOrdersID);

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
                            NewRow["AccountingName"] = OrdersDataTable.Rows[0]["AccountingName"].ToString();
                            NewRow["InvNumber"] = OrdersDataTable.Rows[0]["InvNumber"].ToString();
                            NewRow["MarketingCost"] = MarketingCost;
                            NewRow["CurrencyCode"] = ProfilCurrencyCode;
                            if (fID == 2)
                                NewRow["TPSCurCode"] = TPSCurrencyCode;
                            NewRow["Measure"] = OrdersDataTable.Rows[0]["Measure"].ToString();
                            NewRow["Count"] = Decimal.Round(CountPP, 3, MidpointRounding.AwayFromZero);
                            NewRow["Cost"] = Decimal.Round(Math.Ceiling(CostPP / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                            NewRow["Weight"] = Decimal.Round(Convert.ToDecimal(NewRow["Count"]) * Convert.ToDecimal(3.5), 3, MidpointRounding.AwayFromZero);
                            if (rows.Count() > 0)
                            {
                                NewRow["AccountingName"] = rows[0]["DecorAccountingName"].ToString();
                                NewRow["InvNumber"] = rows[0]["DecorInvNumber"].ToString();
                                NewRow["Cvet"] = rows[0]["Cvet"].ToString();
                                NewRow["Patina"] = rows[0]["Patina"].ToString();
                            }
                            ReportDataTable1.Rows.Add(NewRow);
                        }
                        else
                        {
                            //CostPP = Math.Ceiling(CostPP / 0.01m) * 0.01m;
                            DataRow NewRow = ReportDataTable.NewRow();
                            NewRow["OriginalPrice"] = CostPP / CountPP;
                            NewRow["UNN"] = UNN;
                            NewRow["MarketingCost"] = MarketingCost;
                            NewRow["AccountingName"] = OrdersDataTable.Rows[0]["AccountingName"].ToString();
                            NewRow["InvNumber"] = OrdersDataTable.Rows[0]["InvNumber"].ToString();
                            NewRow["CurrencyCode"] = ProfilCurrencyCode;
                            if (fID == 2)
                                NewRow["TPSCurCode"] = TPSCurrencyCode;
                            NewRow["Count"] = Decimal.Round(CountPP, 3, MidpointRounding.AwayFromZero);
                            NewRow["Cost"] = Decimal.Round(Math.Ceiling(CostPP / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                            NewRow["Measure"] = OrdersDataTable.Rows[0]["Measure"].ToString(); NewRow["Weight"] = Decimal.Round(Convert.ToDecimal(NewRow["Count"]) * Convert.ToDecimal(3.5), 3, MidpointRounding.AwayFromZero);
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


                    //DataRow[] rows = DecorInvNumbersDT.Select("FrontsOrdersID = " + Convert.ToInt32(OrdersDataTable.Select("InsetTypeID IN (685,686,687,688,29470,29471)")[0]["FrontsOrdersID"]));

                    //int FactoryID = 0;
                    //if (rows.Count() > 0)
                    //    FactoryID = Convert.ToInt32(rows[0]["FactoryID"]);
                    //if (FactoryID == 0)
                    //    FactoryID = FactoryID1;
                    //if (FactoryID != FactoryID1)
                    //{
                    //    //CostPP = Math.Ceiling(CostPP / 0.01m) * 0.01m;
                    //    DataRow NewRow = ReportDataTable1.NewRow();
                    //    NewRow["OriginalPrice"] = CostPP / CountPP;
                    //    NewRow["UNN"] = UNN;
                    //    NewRow["AccountingName"] = OrdersDataTable.Rows[0]["AccountingName"].ToString();
                    //    NewRow["InvNumber"] = OrdersDataTable.Rows[0]["InvNumber"].ToString();
                    //    NewRow["MarketingCost"] = MarketingCost;
                    //    NewRow["CurrencyCode"] = ProfilCurrencyCode;
                    //    if (fID == 2)
                    //        NewRow["TPSCurCode"] = TPSCurrencyCode;
                    //    NewRow["Measure"] = OrdersDataTable.Rows[0]["Measure"].ToString();
                    //    NewRow["Count"] = Decimal.Round(CountPP, 3, MidpointRounding.AwayFromZero);
                    //    NewRow["Cost"] = Decimal.Round(Math.Ceiling(CostPP / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                    //    NewRow["Weight"] = Decimal.Round(Convert.ToDecimal(NewRow["Count"]) * Convert.ToDecimal(3.5), 3, MidpointRounding.AwayFromZero);
                    //    if (rows.Count() > 0)
                    //    {
                    //        NewRow["AccountingName"] = rows[0]["DecorAccountingName"].ToString();
                    //        NewRow["InvNumber"] = rows[0]["DecorInvNumber"].ToString();
                    //        NewRow["Cvet"] = rows[0]["Cvet"].ToString();
                    //        NewRow["Patina"] = rows[0]["Patina"].ToString();
                    //    }
                    //    ReportDataTable1.Rows.Add(NewRow);
                    //}
                    //else
                    //{
                    //    //CostPP = Math.Ceiling(CostPP / 0.01m) * 0.01m;
                    //    DataRow NewRow = ReportDataTable.NewRow();
                    //    NewRow["OriginalPrice"] = CostPP / CountPP;
                    //    NewRow["UNN"] = UNN;
                    //    NewRow["MarketingCost"] = MarketingCost;
                    //    NewRow["AccountingName"] = OrdersDataTable.Rows[0]["AccountingName"].ToString();
                    //    NewRow["InvNumber"] = OrdersDataTable.Rows[0]["InvNumber"].ToString();
                    //    NewRow["CurrencyCode"] = ProfilCurrencyCode;
                    //    if (fID == 2)
                    //        NewRow["TPSCurCode"] = TPSCurrencyCode;
                    //    NewRow["Count"] = Decimal.Round(CountPP, 3, MidpointRounding.AwayFromZero);
                    //    NewRow["Cost"] = Decimal.Round(Math.Ceiling(CostPP / 0.01m) * 0.01m, 2, MidpointRounding.AwayFromZero);
                    //    NewRow["Measure"] = OrdersDataTable.Rows[0]["Measure"].ToString(); NewRow["Weight"] = Decimal.Round(Convert.ToDecimal(NewRow["Count"]) * Convert.ToDecimal(3.5), 3, MidpointRounding.AwayFromZero);
                    //    if (rows.Count() > 0)
                    //    {
                    //        NewRow["AccountingName"] = rows[0]["DecorAccountingName"].ToString();
                    //        NewRow["InvNumber"] = rows[0]["DecorInvNumber"].ToString();
                    //        NewRow["Cvet"] = rows[0]["Cvet"].ToString();
                    //        NewRow["Patina"] = rows[0]["Patina"].ToString();
                    //    }
                    //    ReportDataTable.Rows.Add(NewRow);
                    //}
                }
            }
        }

        private void GetGlass(DataTable OrdersDataTable, DataTable ReportDataTable, DataTable ReportDataTable1, int FactoryID1)
        {
            decimal MarketingCost = 0;
            string AccountingName = OrdersDataTable.Rows[0]["AccountingName"].ToString();
            string InvNumber = OrdersDataTable.Rows[0]["InvNumber"].ToString();
            string Measure = OrdersDataTable.Rows[0]["Measure"].ToString();

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
                    MarketingCost = GetMarketingCost(Convert.ToInt32(FRows[0]["FrontConfigID"]));
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
                        if (MarketingCost > GetMarketingCost(Convert.ToInt32(FRows[j]["FrontConfigID"])))
                            MarketingCost = GetMarketingCost(Convert.ToInt32(FRows[j]["FrontConfigID"]));
                    }

                    CountFlutes = Decimal.Round(CountFlutes, 3, MidpointRounding.AwayFromZero);

                    if (CountFlutes > 0)
                    {
                        DataRow[] rows =
                            DecorInvNumbersDT.Select("FrontsOrdersID = " + Convert.ToInt32(FRows[0]["FrontsOrdersID"]));
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
                            NewRow["AccountingName"] = AccountingName;
                            NewRow["InvNumber"] = InvNumber;
                            NewRow["MarketingCost"] = MarketingCost;
                            NewRow["Measure"] = Measure;
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
                            NewRow["Cost"] = Decimal.Round(CountFlutes * PriceFlutes);
                            NewRow["Weight"] = Decimal.Round(CountFlutes * 10, 3, MidpointRounding.AwayFromZero);
                            ReportDataTable1.Rows.Add(NewRow);
                        }
                        else
                        {
                            DataRow NewRow = ReportDataTable.NewRow();
                            //NewRow["Name"] = "Стекло Флутес";
                            NewRow["OriginalPrice"] = PriceFlutes;
                            NewRow["UNN"] = UNN;
                            NewRow["AccountingName"] = AccountingName;
                            NewRow["InvNumber"] = InvNumber;
                            NewRow["MarketingCost"] = MarketingCost;
                            NewRow["Measure"] = Measure;
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
                            NewRow["Cost"] = Decimal.Round(CountFlutes * PriceFlutes);
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
                    MarketingCost = GetMarketingCost(Convert.ToInt32(LRows[0]["FrontConfigID"]));
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
                        if (MarketingCost > GetMarketingCost(Convert.ToInt32(LRows[j]["FrontConfigID"])))
                            MarketingCost = GetMarketingCost(Convert.ToInt32(LRows[j]["FrontConfigID"]));
                    }

                    CountLacomat = Decimal.Round(CountLacomat, 3, MidpointRounding.AwayFromZero);

                    if (CountLacomat > 0)
                    {
                        DataRow[] rows =
                            DecorInvNumbersDT.Select("FrontsOrdersID = " + Convert.ToInt32(LRows[0]["FrontsOrdersID"]));
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
                            NewRow["AccountingName"] = AccountingName;
                            NewRow["InvNumber"] = InvNumber;
                            NewRow["MarketingCost"] = MarketingCost;
                            NewRow["Measure"] = Measure;
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
                            NewRow["Cost"] = Decimal.Round(CountLacomat * PriceLacomat);
                            NewRow["Weight"] = Decimal.Round(CountLacomat * 10, 3, MidpointRounding.AwayFromZero);
                            ReportDataTable1.Rows.Add(NewRow);
                        }
                        else
                        {
                            DataRow NewRow = ReportDataTable.NewRow();
                            //NewRow["Name"] = "Стекло Лакомат";
                            NewRow["OriginalPrice"] = PriceLacomat;
                            NewRow["UNN"] = UNN;
                            NewRow["AccountingName"] = AccountingName;
                            NewRow["InvNumber"] = InvNumber;
                            NewRow["MarketingCost"] = MarketingCost;
                            NewRow["Measure"] = Measure;
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
                            NewRow["Cost"] = Decimal.Round(CountLacomat * PriceLacomat);
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
                    MarketingCost = GetMarketingCost(Convert.ToInt32(KRows[0]["FrontConfigID"]));
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
                        if (MarketingCost > GetMarketingCost(Convert.ToInt32(KRows[j]["FrontConfigID"])))
                            MarketingCost = GetMarketingCost(Convert.ToInt32(KRows[j]["FrontConfigID"]));
                    }

                    CountKrizet = Decimal.Round(CountKrizet, 3, MidpointRounding.AwayFromZero);

                    if (CountKrizet > 0)
                    {
                        DataRow[] rows =
                            DecorInvNumbersDT.Select("FrontsOrdersID = " + Convert.ToInt32(KRows[0]["FrontsOrdersID"]));
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
                            NewRow["AccountingName"] = AccountingName;
                            NewRow["InvNumber"] = InvNumber;
                            NewRow["MarketingCost"] = MarketingCost;
                            NewRow["Measure"] = Measure;
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
                            NewRow["Cost"] = Decimal.Round(CountKrizet * PriceKrizet);
                            NewRow["Weight"] = Decimal.Round(CountKrizet * 10, 3, MidpointRounding.AwayFromZero);
                            ReportDataTable1.Rows.Add(NewRow);
                        }
                        else
                        {
                            DataRow NewRow = ReportDataTable.NewRow();
                            //NewRow["Name"] = "Стекло Кризет";
                            NewRow["OriginalPrice"] = PriceKrizet;
                            NewRow["UNN"] = UNN;
                            NewRow["AccountingName"] = AccountingName;
                            NewRow["InvNumber"] = InvNumber;
                            NewRow["MarketingCost"] = MarketingCost;
                            NewRow["Measure"] = Measure;
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
                            NewRow["Cost"] = Decimal.Round(CountKrizet * PriceKrizet);
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
                    MarketingCost = GetMarketingCost(Convert.ToInt32(ORows[0]["FrontConfigID"]));
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
                        if (MarketingCost > GetMarketingCost(Convert.ToInt32(ORows[j]["FrontConfigID"])))
                            MarketingCost = GetMarketingCost(Convert.ToInt32(ORows[j]["FrontConfigID"]));
                    }

                    CountOther = Decimal.Round(CountOther, 3, MidpointRounding.AwayFromZero);

                    if (CountOther > 0)
                    {
                        DataRow[] rows =
                            DecorInvNumbersDT.Select("FrontsOrdersID = " + Convert.ToInt32(ORows[0]["FrontsOrdersID"]));
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
                            NewRow["AccountingName"] = AccountingName;
                            NewRow["InvNumber"] = InvNumber;
                            NewRow["MarketingCost"] = MarketingCost;
                            NewRow["Measure"] = Measure;
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
                            NewRow["Weight"] = Decimal.Round(CountOther * 10, 3, MidpointRounding.AwayFromZero);
                            ReportDataTable1.Rows.Add(NewRow);
                        }
                        else
                        {
                            DataRow NewRow = ReportDataTable.NewRow();
                            //NewRow["Name"] = "Стекло другое";
                            NewRow["OriginalPrice"] = 0;
                            NewRow["UNN"] = UNN;
                            NewRow["AccountingName"] = AccountingName;
                            NewRow["InvNumber"] = InvNumber;
                            NewRow["MarketingCost"] = MarketingCost;
                            NewRow["Measure"] = Measure;
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
                            NewRow["Weight"] = Decimal.Round(CountOther * 10, 3, MidpointRounding.AwayFromZero);
                            ReportDataTable.Rows.Add(NewRow);
                        }
                    }
                }
            }
        }

        private void GetInsets(DataTable OrdersDataTable, DataTable ReportDataTable)
        {
            decimal MarketingCost = 0;
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

                    MarketingCost = GetMarketingCost(Convert.ToInt32(EGRows[0]["FrontConfigID"]));
                    for (int j = 0; j < EGRows.Count(); j++)
                    {
                        decimal dd =
                            Decimal.Round(
                                Convert.ToDecimal(EGRows[j]["Count"]) * MarginHeight *
                                (Convert.ToDecimal(EGRows[j]["Width"]) - MarginWidth) / 1000000, 3,
                                MidpointRounding.AwayFromZero);
                        CountEllipseGrid += dd;
                        if (MarketingCost > GetMarketingCost(Convert.ToInt32(EGRows[j]["FrontConfigID"])))
                            MarketingCost = GetMarketingCost(Convert.ToInt32(EGRows[j]["FrontConfigID"]));
                    }

                    decimal Weight = GetInsetWeight(EGRows[0]);
                    Weight = Decimal.Round(CountEllipseGrid * Weight, 3, MidpointRounding.AwayFromZero);

                    DataRow NewRow = ReportDataTable.NewRow();
                    NewRow["OriginalPrice"] = PriceEllipseGrid;
                    NewRow["MarketingCost"] = MarketingCost;
                    NewRow["UNN"] = UNN;
                    NewRow["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        NewRow["TPSCurCode"] = TPSCurrencyCode;
                    NewRow["Count"] = CountEllipseGrid;
                    NewRow["Cost"] = CountEllipseGrid * PriceEllipseGrid;
                    NewRow["Weight"] = 0;
                    NewRow["Measure"] = OrdersDataTable.Rows[0]["Measure"].ToString();
                    DataRow[] rows =
                        DecorInvNumbersDT.Select("FrontsOrdersID = " + Convert.ToInt32(EGRows[0]["FrontsOrdersID"]));
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

        private void GetSimpleFrontsByClient(DataTable OrdersDataTable, DataTable ReportDataTable)
        {
            string DecorAccountingName = string.Empty;
            string DecorInvNumber = string.Empty;
            string AccountingName = OrdersDataTable.Rows[0]["AccountingName"].ToString();
            string InvNumber = OrdersDataTable.Rows[0]["InvNumber"].ToString();
            string Measure = OrdersDataTable.Rows[0]["Measure"].ToString();
            int fID = Convert.ToInt32(OrdersDataTable.Rows[0]["FactoryID"]);
            DataTable Fronts = new DataTable();

            using (DataView DV = new DataView(OrdersDataTable))
            {
                Fronts = DV.ToTable(true, new string[] { "FrontID" });
            }

            for (int i = 0; i < Fronts.Rows.Count; i++)
            {
                decimal SolidCount = 0;
                decimal SolidCost = 0;
                decimal SolidWeight = 0;

                decimal FilenkaCount = 0;
                decimal FilenkaCost = 0;
                decimal FilenkaWeight = 0;

                decimal VitrinaCount = 0;
                decimal VitrinaCost = 0;
                decimal VitrinaWeight = 0;

                decimal LuxMegaCount = 0;
                decimal LuxMegaCost = 0;
                decimal LuxMegaWeight = 0;

                decimal MarketingCost = 0;

                //ГЛУХИЕ, БЕЗ ВСТАВКИ, РЕШЕТКА ОВАЛ
                DataRow[] rows = InsetTypesDataTable.Select("InsetTypeID=-1 OR GroupID = 3 OR GroupID = 4");
                string filter = string.Empty;
                foreach (DataRow item in rows)
                    filter += item["InsetTypeID"].ToString() + ",";
                if (filter.Length > 0)
                    filter = " AND NOT (FrontID IN (3728,3731,3732,3739,3740,3741,3744,3745,3746) OR InsetTypeID IN (28961,3653,3654,3655)) AND (FrontID = 3729 OR InsetTypeID IN (" + filter.Substring(0, filter.Length - 1) + "))";
                DataRow[] Rows = OrdersDataTable.Select("FrontID = " + Fronts.Rows[i]["FrontID"].ToString() + " AND (Width <> -1)" + filter);
                if (Rows.Count() > 0)
                    MarketingCost = GetMarketingCost(Convert.ToInt32(Rows[0]["FrontConfigID"]));
                for (int r = 0; r < Rows.Count(); r++)
                {
                    decimal DeductibleCount = 0;
                    decimal DeductibleWeight = 0;
                    if (Convert.ToInt32(Fronts.Rows[i]["FrontID"]) == 3729)//РЕШЕТКА ОВАЛ
                    {
                        int FactoryID = 0;
                        DecorInvNumber = GetGridInvNumber(Convert.ToInt32(Rows[r]["FrontConfigID"]), ref FactoryID, ref DecorAccountingName);
                        if (DecorInvNumber.Length > 0)
                        {
                            DataRow NewRow = DecorInvNumbersDT.NewRow();
                            NewRow["FrontsOrdersID"] = Convert.ToInt32(Rows[r]["FrontsOrdersID"]);
                            NewRow["FactoryID"] = FactoryID;
                            NewRow["DecorAccountingName"] = DecorAccountingName;
                            NewRow["DecorInvNumber"] = DecorInvNumber;
                            DecorInvNumbersDT.Rows.Add(NewRow);
                            DeductibleWeight = GetInsetWeight(Rows[r]);
                            DeductibleCount = GetInsetSquare(Convert.ToInt32(Rows[r]["FrontID"]), Convert.ToInt32(Rows[r]["Height"]),
                                                        Convert.ToInt32(Rows[r]["Width"])) * Convert.ToDecimal(Rows[r]["Count"]);
                            DeductibleWeight = Decimal.Round(DeductibleCount * DeductibleWeight, 3, MidpointRounding.AwayFromZero);
                        }
                    }
                    if (GetMeasureType(Convert.ToInt32(Rows[r]["FrontConfigID"])) == 3)
                        SolidCount += Decimal.Round(Convert.ToDecimal(Rows[r]["Height"]) * Convert.ToDecimal(Rows[r]["Width"]) / 1000000, 3, MidpointRounding.AwayFromZero) *
                                      Convert.ToDecimal(Rows[r]["Count"]);
                    else
                        SolidCount += Convert.ToDecimal(Rows[r]["Square"]);

                    if (GetMeasureType(Convert.ToInt32(Rows[r]["FrontConfigID"])) == 3)
                        SolidCost += Convert.ToDecimal(Rows[r]["Cost"]);
                    else
                        SolidCost += Convert.ToDecimal(Rows[r]["FrontPrice"]) * Convert.ToDecimal(Rows[r]["Square"]);

                    decimal FrontWeight = 0;
                    decimal InsetWeight = 0;

                    GetFrontWeight(Rows[r], ref FrontWeight, ref InsetWeight);

                    SolidWeight += Convert.ToDecimal(FrontWeight + InsetWeight - DeductibleWeight);
                    if (MarketingCost > GetMarketingCost(Convert.ToInt32(Rows[r]["FrontConfigID"])))
                        MarketingCost = GetMarketingCost(Convert.ToInt32(Rows[r]["FrontConfigID"]));
                }
                //АППЛИКАЦИИ
                filter = " AND (FrontID IN (3728,3731,3732,3739,3740,3741,3744,3745,3746) OR InsetTypeID IN (28961,3653,3654,3655))";
                Rows = OrdersDataTable.Select("FrontID = " + Fronts.Rows[i]["FrontID"].ToString() + " AND (Width <> -1)" + filter);
                if (Rows.Count() > 0)
                    MarketingCost = GetMarketingCost(Convert.ToInt32(Rows[0]["FrontConfigID"]));
                for (int r = 0; r < Rows.Count(); r++)
                {
                    if (Convert.ToInt32(Rows[r]["FrontID"]) == 3728 || Convert.ToInt32(Rows[r]["FrontID"]) == 3731 || Convert.ToInt32(Rows[r]["FrontID"]) == 3732 ||
                        Convert.ToInt32(Rows[r]["FrontID"]) == 3739 || Convert.ToInt32(Rows[r]["FrontID"]) == 3740 || Convert.ToInt32(Rows[r]["FrontID"]) == 3741 ||
                        Convert.ToInt32(Rows[r]["FrontID"]) == 3744 || Convert.ToInt32(Rows[r]["FrontID"]) == 3745 || Convert.ToInt32(Rows[r]["FrontID"]) == 3746)
                    {
                        if (GetMeasureType(Convert.ToInt32(Rows[r]["FrontConfigID"])) == 3)
                            SolidCount += Decimal.Round(Convert.ToDecimal(Rows[r]["Height"]) * Convert.ToDecimal(Rows[r]["Width"]) / 1000000, 3, MidpointRounding.AwayFromZero) *
                                          Convert.ToDecimal(Rows[r]["Count"]);
                        else
                            SolidCount += Convert.ToDecimal(Rows[r]["Square"]);

                        if (GetMeasureType(Convert.ToInt32(Rows[r]["FrontConfigID"])) == 3)
                            SolidCost += Convert.ToDecimal(Rows[r]["Cost"]) + 5;
                        else
                            SolidCost += Convert.ToDecimal(Rows[r]["FrontPrice"]) * Convert.ToDecimal(Rows[r]["Square"]) + 5 * Convert.ToDecimal(Rows[r]["Count"]);

                        decimal FrontWeight = 0;
                        decimal InsetWeight = 0;

                        GetFrontWeight(Rows[r], ref FrontWeight, ref InsetWeight);

                        SolidWeight += Convert.ToDecimal(FrontWeight + InsetWeight);
                    }
                    else if (Convert.ToInt32(Rows[r]["FrontID"]) == 3415 || Convert.ToInt32(Rows[r]["FrontID"]) == 28922)
                    {
                        if (GetMeasureType(Convert.ToInt32(Rows[r]["FrontConfigID"])) == 3)
                            FilenkaCount += Decimal.Round(Convert.ToDecimal(Rows[r]["Height"]) * Convert.ToDecimal(Rows[r]["Width"]) / 1000000, 3, MidpointRounding.AwayFromZero) *
                                          Convert.ToDecimal(Rows[r]["Count"]);
                        else
                            FilenkaCount += Convert.ToDecimal(Rows[r]["Square"]);

                        if (GetMeasureType(Convert.ToInt32(Rows[r]["FrontConfigID"])) == 3)
                            FilenkaCost += Convert.ToDecimal(Rows[r]["Cost"]) + 5;
                        else
                            FilenkaCost += Convert.ToDecimal(Rows[r]["FrontPrice"]) * Convert.ToDecimal(Rows[r]["Square"]) + 5 * Convert.ToDecimal(Rows[r]["Count"]);

                        decimal FrontWeight = 0;
                        decimal InsetWeight = 0;

                        GetFrontWeight(Rows[r], ref FrontWeight, ref InsetWeight);

                        FilenkaWeight += Convert.ToDecimal(FrontWeight + InsetWeight);
                    }
                    if (MarketingCost > GetMarketingCost(Convert.ToInt32(Rows[r]["FrontConfigID"])))
                        MarketingCost = GetMarketingCost(Convert.ToInt32(Rows[r]["FrontConfigID"]));
                }
                //ФИЛЕНКА
                filter = " AND InsetTypeID IN (2069,2070,2071,2073,2075,42066,2077,2233,3644,29043,29531,41213)";
                Rows = OrdersDataTable.Select("FrontID = " + Fronts.Rows[i]["FrontID"].ToString() + " AND (Width <> -1)" + filter);
                if (Rows.Count() > 0) MarketingCost = GetMarketingCost(Convert.ToInt32(Rows[0]["FrontConfigID"]));
                for (int r = 0; r < Rows.Count(); r++)
                {
                    if (GetMeasureType(Convert.ToInt32(Rows[r]["FrontConfigID"])) == 3)
                        FilenkaCount += Decimal.Round(Convert.ToDecimal(Rows[r]["Height"]) * Convert.ToDecimal(Rows[r]["Width"]) / 1000000, 3, MidpointRounding.AwayFromZero) *
                                      Convert.ToDecimal(Rows[r]["Count"]);
                    else
                        FilenkaCount += Convert.ToDecimal(Rows[r]["Square"]);

                    if (GetMeasureType(Convert.ToInt32(Rows[r]["FrontConfigID"])) == 3)
                        FilenkaCost += Convert.ToDecimal(Rows[r]["Cost"]);
                    else
                        FilenkaCost += Convert.ToDecimal(Rows[r]["FrontPrice"]) * Convert.ToDecimal(Rows[r]["Square"]);

                    decimal FrontWeight = 0;
                    decimal InsetWeight = 0;

                    GetFrontWeight(Rows[r], ref FrontWeight, ref InsetWeight);

                    FilenkaWeight += Convert.ToDecimal(FrontWeight + InsetWeight);
                    if (MarketingCost > GetMarketingCost(Convert.ToInt32(Rows[r]["FrontConfigID"])))
                        MarketingCost = GetMarketingCost(Convert.ToInt32(Rows[r]["FrontConfigID"]));
                }
                //ВИТРИНЫ, РЕШЕТКИ, СТЕКЛО
                filter = " AND InsetTypeID IN (1,2,685,686,687,688,29470,29471) AND FrontID <> 3729";
                Rows = OrdersDataTable.Select("FrontID = " + Fronts.Rows[i]["FrontID"].ToString() + " AND (Width <> -1)" + filter);
                if (Rows.Count() > 0)
                    MarketingCost = GetMarketingCost(Convert.ToInt32(Rows[0]["FrontConfigID"]));
                for (int r = 0; r < Rows.Count(); r++)
                {
                    decimal DeductibleCount = 0;
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
                            NewRow["FrontsOrdersID"] = Convert.ToInt32(Rows[r]["FrontsOrdersID"]);
                            NewRow["FactoryID"] = FactoryID;
                            NewRow["DecorAccountingName"] = DecorAccountingName;
                            NewRow["DecorInvNumber"] = DecorInvNumber;
                            DecorInvNumbersDT.Rows.Add(NewRow);
                            DeductibleCount = GetInsetSquare(Convert.ToInt32(Rows[r]["FrontID"]), Convert.ToInt32(Rows[r]["Height"]),
                                                        Convert.ToInt32(Rows[r]["Width"])) * Convert.ToDecimal(Rows[r]["Count"]);
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
                            NewRow["FrontsOrdersID"] = Convert.ToInt32(Rows[r]["FrontsOrdersID"]);
                            NewRow["FactoryID"] = FactoryID;
                            NewRow["DecorAccountingName"] = DecorAccountingName;
                            NewRow["DecorInvNumber"] = DecorInvNumber;
                            DecorInvNumbersDT.Rows.Add(NewRow);
                            DeductibleCount = Convert.ToDecimal(Rows[r]["Count"]) * GetInsetSquare(Convert.ToInt32(Rows[r]["FrontID"]), Convert.ToInt32(Rows[r]["Height"]), Convert.ToInt32(Rows[r]["Width"]));
                            DeductibleWeight = Decimal.Round(DeductibleCount * 10, 3, MidpointRounding.AwayFromZero);
                        }
                    }

                    if (GetMeasureType(Convert.ToInt32(Rows[r]["FrontConfigID"])) == 3)
                        VitrinaCount += Decimal.Round(Convert.ToDecimal(Rows[r]["Height"]) * Convert.ToDecimal(Rows[r]["Width"]) / 1000000, 3, MidpointRounding.AwayFromZero) *
                                      Convert.ToDecimal(Rows[r]["Count"]);
                    else
                        VitrinaCount += Convert.ToDecimal(Rows[r]["Square"]);

                    if (IsAluminium(Rows[r]) > -1)
                    {
                        VitrinaCost += GetFrontCostAluminium(Rows[r]);
                        VitrinaWeight += GetAluminiumWeight(Rows[r], true);
                    }
                    else
                    {
                        if (GetMeasureType(Convert.ToInt32(Rows[r]["FrontConfigID"])) == 3)
                            VitrinaCost += Convert.ToDecimal(Rows[r]["Cost"]);
                        else
                            VitrinaCost += Convert.ToDecimal(Rows[r]["FrontPrice"]) * Convert.ToDecimal(Rows[r]["Square"]);

                        decimal FrontWeight = 0;
                        decimal InsetWeight = 0;

                        GetFrontWeight(Rows[r], ref FrontWeight, ref InsetWeight);

                        VitrinaWeight += Convert.ToDecimal(FrontWeight + InsetWeight - DeductibleWeight);
                    }
                    if (MarketingCost > GetMarketingCost(Convert.ToInt32(Rows[r]["FrontConfigID"])))
                        MarketingCost = GetMarketingCost(Convert.ToInt32(Rows[r]["FrontConfigID"]));
                }
                //ЛЮКС, МЕГА
                filter = " AND InsetTypeID IN (860,862,4310)";
                Rows = OrdersDataTable.Select("FrontID = " + Fronts.Rows[i]["FrontID"].ToString() + " AND (Width <> -1)" + filter);
                if (Rows.Count() > 0)
                    MarketingCost = GetMarketingCost(Convert.ToInt32(Rows[0]["FrontConfigID"]));
                for (int r = 0; r < Rows.Count(); r++)
                {
                    if (GetMeasureType(Convert.ToInt32(Rows[r]["FrontConfigID"])) == 3)
                        LuxMegaCount += Decimal.Round(Convert.ToDecimal(Rows[r]["Height"]) * Convert.ToDecimal(Rows[r]["Width"]) / 1000000, 3, MidpointRounding.AwayFromZero) *
                                      Convert.ToDecimal(Rows[r]["Count"]);
                    else
                        LuxMegaCount += Convert.ToDecimal(Rows[r]["Square"]);

                    if (GetMeasureType(Convert.ToInt32(Rows[r]["FrontConfigID"])) == 3)
                        LuxMegaCost += Convert.ToDecimal(Rows[r]["Cost"]);
                    else
                        LuxMegaCost += Convert.ToDecimal(Rows[r]["FrontPrice"]) * Convert.ToDecimal(Rows[r]["Square"]);

                    decimal FrontWeight = 0;
                    decimal InsetWeight = 0;

                    GetFrontWeight(Rows[r], ref FrontWeight, ref InsetWeight);

                    LuxMegaWeight += Convert.ToDecimal(FrontWeight + InsetWeight);
                    if (MarketingCost > GetMarketingCost(Convert.ToInt32(Rows[r]["FrontConfigID"])))
                        MarketingCost = GetMarketingCost(Convert.ToInt32(Rows[r]["FrontConfigID"]));
                }

                if (SolidCount > 0)
                {
                    DataRow Row = ReportDataTable.NewRow();
                    Row["OriginalPrice"] = SolidCost / SolidCount;
                    Row["UNN"] = UNN;
                    Row["AccountingName"] = AccountingName;
                    Row["InvNumber"] = InvNumber;
                    Row["MarketingCost"] = MarketingCost;
                    Row["Measure"] = Measure;
                    Row["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        Row["TPSCurCode"] = TPSCurrencyCode;
                    Row["Count"] = Decimal.Round(SolidCount, 3, MidpointRounding.AwayFromZero);
                    Row["Cost"] = Decimal.Round(SolidCost, 3, MidpointRounding.AwayFromZero);
                    Row["Weight"] = Decimal.Round(SolidWeight, 3, MidpointRounding.AwayFromZero);
                    ReportDataTable.Rows.Add(Row);
                }

                if (FilenkaCount > 0)
                {
                    DataRow Row = ReportDataTable.NewRow();
                    Row["OriginalPrice"] = FilenkaCost / FilenkaCount;
                    Row["UNN"] = UNN;
                    Row["AccountingName"] = AccountingName;
                    Row["InvNumber"] = InvNumber;
                    Row["MarketingCost"] = MarketingCost;
                    Row["Measure"] = Measure;
                    Row["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        Row["TPSCurCode"] = TPSCurrencyCode;
                    Row["Count"] = Decimal.Round(FilenkaCount, 3, MidpointRounding.AwayFromZero);
                    Row["Cost"] = Decimal.Round(FilenkaCost, 3, MidpointRounding.AwayFromZero);
                    Row["Weight"] = Decimal.Round(FilenkaWeight, 3, MidpointRounding.AwayFromZero);
                    ReportDataTable.Rows.Add(Row);
                }

                if (VitrinaCount > 0)
                {
                    DataRow Row = ReportDataTable.NewRow();
                    Row["OriginalPrice"] = VitrinaCost / VitrinaCount;
                    Row["UNN"] = UNN;
                    Row["AccountingName"] = AccountingName;
                    Row["InvNumber"] = InvNumber;
                    Row["MarketingCost"] = MarketingCost;
                    Row["Measure"] = Measure;
                    Row["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        Row["TPSCurCode"] = TPSCurrencyCode;
                    Row["Count"] = Decimal.Round(VitrinaCount, 3, MidpointRounding.AwayFromZero);
                    Row["Cost"] = Decimal.Round(VitrinaCost, 3, MidpointRounding.AwayFromZero);
                    Row["Weight"] = Decimal.Round(VitrinaWeight, 3, MidpointRounding.AwayFromZero);
                    ReportDataTable.Rows.Add(Row);
                }

                if (LuxMegaCount > 0)
                {
                    DataRow Row = ReportDataTable.NewRow();
                    Row["OriginalPrice"] = LuxMegaCost / LuxMegaCount;
                    Row["UNN"] = UNN;
                    Row["AccountingName"] = AccountingName;
                    Row["InvNumber"] = InvNumber;
                    Row["MarketingCost"] = MarketingCost;
                    Row["Measure"] = Measure;
                    Row["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        Row["TPSCurCode"] = TPSCurrencyCode;
                    Row["Count"] = Decimal.Round(LuxMegaCount, 3, MidpointRounding.AwayFromZero);
                    Row["Cost"] = Decimal.Round(LuxMegaCost, 3, MidpointRounding.AwayFromZero);
                    Row["Weight"] = Decimal.Round(LuxMegaWeight, 3, MidpointRounding.AwayFromZero);
                    ReportDataTable.Rows.Add(Row);
                }
            }

            Fronts.Dispose();
        }

        private void GetCurvedFrontsByClient(DataTable OrdersDataTable, DataTable ReportDataTable)
        {
            string AccountingName = OrdersDataTable.Rows[0]["AccountingName"].ToString();
            string InvNumber = OrdersDataTable.Rows[0]["InvNumber"].ToString();
            string Measure = OrdersDataTable.Rows[0]["Measure"].ToString();
            int fID = Convert.ToInt32(OrdersDataTable.Rows[0]["FactoryID"]);
            DataTable Fronts = new DataTable();

            using (DataView DV = new DataView(OrdersDataTable))
            {
                Fronts = DV.ToTable(true, new string[] { "FrontID" });
            }

            for (int i = 0; i < Fronts.Rows.Count; i++)
            {
                DataRow[] Rows = OrdersDataTable.Select("FrontID = " + Fronts.Rows[i]["FrontID"].ToString());

                if (Rows.Count() == 0)
                    continue;
                decimal MarketingCost = 0;
                decimal Solid713Count = 0;
                decimal Solid713Price = 0;
                decimal Solid713Weight = 0;

                decimal Filenka713Count = 0;
                decimal Filenka713Price = 0;
                decimal Filenka713Weight = 0;

                decimal NoInset713Count = 0;
                decimal NoInset713Price = 0;
                decimal NoInset713Weight = 0;

                decimal Vitrina713Count = 0;
                decimal Vitrina713Price = 0;
                decimal Vitrina713Weight = 0;

                decimal Solid910Count = 0;
                decimal Solid910Price = 0;
                decimal Solid910Weight = 0;

                decimal Filenka910Count = 0;
                decimal Filenka910Price = 0;
                decimal Filenka910Weight = 0;

                decimal NoInset910Count = 0;
                decimal NoInset910Price = 0;
                decimal NoInset910Weight = 0;

                decimal Vitrina910Count = 0;
                decimal Vitrina910Price = 0;
                decimal Vitrina910Weight = 0;

                decimal ScagenCount = 0;
                decimal ScagenPrice = 0;
                decimal ScagenWeight = 0;

                if (Rows.Count() > 0)
                    MarketingCost = GetMarketingCost(Convert.ToInt32(Rows[0]["FrontConfigID"]));
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

                            decimal FrontWeight = 0;
                            decimal InsetWeight = 0;

                            GetFrontWeight(Rows[r], ref FrontWeight, ref InsetWeight);

                            ScagenWeight += Convert.ToDecimal(FrontWeight + InsetWeight);
                        }

                        if (Rows[r]["InsetTypeID"].ToString() == "1")
                        {
                            ScagenCount += Convert.ToDecimal(Rows[r]["Count"]);
                            ScagenPrice = Convert.ToDecimal(Rows[r]["FrontPrice"]);

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

                            decimal FrontWeight = 0;
                            decimal InsetWeight = 0;

                            GetFrontWeight(Rows[r], ref FrontWeight, ref InsetWeight);

                            NoInset713Weight += Convert.ToDecimal(FrontWeight + InsetWeight);
                        }

                        if (Rows[r]["InsetTypeID"].ToString() == "1")
                        {
                            Vitrina713Count += Convert.ToDecimal(Rows[r]["Count"]);
                            Vitrina713Price = Convert.ToDecimal(Rows[r]["FrontPrice"]);

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

                            decimal FrontWeight = 0;
                            decimal InsetWeight = 0;

                            GetFrontWeight(Rows[r], ref FrontWeight, ref InsetWeight);

                            NoInset910Weight += Convert.ToDecimal(FrontWeight + InsetWeight);
                        }

                        if (Rows[r]["InsetTypeID"].ToString() == "1")
                        {
                            Vitrina910Count += Convert.ToDecimal(Rows[r]["Count"]);
                            Vitrina910Price = Convert.ToDecimal(Rows[r]["FrontPrice"]);

                            decimal FrontWeight = 0;
                            decimal InsetWeight = 0;

                            GetFrontWeight(Rows[r], ref FrontWeight, ref InsetWeight);

                            Vitrina910Weight += Convert.ToDecimal(FrontWeight + InsetWeight);
                        }
                    }
                    if (MarketingCost > GetMarketingCost(Convert.ToInt32(Rows[r]["FrontConfigID"])))
                        MarketingCost = GetMarketingCost(Convert.ToInt32(Rows[r]["FrontConfigID"]));
                }

                if (ScagenCount > 0)
                {
                    DataRow Row = ReportDataTable.NewRow();
                    Row["OriginalPrice"] = ScagenPrice;
                    Row["UNN"] = UNN;
                    Row["batchId"] = Convert.ToInt32(Fronts.Rows[i]["batchId"]);
                    Row["IsComplaint"] = Convert.ToInt32(Fronts.Rows[i]["IsComplaint"]);
                    Row["Cvet"] = GetColorCode(Convert.ToInt32(Fronts.Rows[i]["ColorID"]));
                    Row["Patina"] = GetPatinaCode(Convert.ToInt32(Fronts.Rows[i]["PatinaID"]));
                    Row["AccountingName"] = AccountingName;
                    Row["InvNumber"] = InvNumber;
                    Row["MarketingCost"] = MarketingCost;
                    Row["Measure"] = Measure;
                    Row["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        Row["TPSCurCode"] = TPSCurrencyCode;
                    Row["Count"] = ScagenCount;
                    Row["Cost"] = Decimal.Round(ScagenCount * ScagenPrice, 3, MidpointRounding.AwayFromZero);
                    Row["Weight"] = Decimal.Round(ScagenWeight, 3, MidpointRounding.AwayFromZero);
                    ReportDataTable.Rows.Add(Row);
                }

                if (Solid713Count > 0)
                {
                    DataRow Row = ReportDataTable.NewRow();
                    Row["OriginalPrice"] = Solid713Price;
                    Row["UNN"] = UNN;
                    Row["AccountingName"] = AccountingName;
                    Row["InvNumber"] = InvNumber;
                    Row["MarketingCost"] = MarketingCost;
                    Row["Measure"] = Measure;
                    Row["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        Row["TPSCurCode"] = TPSCurrencyCode;
                    Row["Count"] = Solid713Count;
                    Row["Cost"] = Decimal.Round(Solid713Count * Solid713Price, 3, MidpointRounding.AwayFromZero);
                    Row["Weight"] = Decimal.Round(Solid713Weight, 3, MidpointRounding.AwayFromZero);
                    ReportDataTable.Rows.Add(Row);
                }

                if (Filenka713Count > 0)
                {
                    DataRow Row = ReportDataTable.NewRow();
                    Row["OriginalPrice"] = Filenka713Price;
                    Row["UNN"] = UNN;
                    Row["AccountingName"] = AccountingName;
                    Row["InvNumber"] = InvNumber;
                    Row["MarketingCost"] = MarketingCost;
                    Row["Measure"] = Measure;
                    Row["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        Row["TPSCurCode"] = TPSCurrencyCode;
                    Row["Count"] = Filenka713Count;
                    Row["Cost"] = Decimal.Round(Filenka713Count * Filenka713Price, 3, MidpointRounding.AwayFromZero);
                    Row["Weight"] = Decimal.Round(Filenka713Weight, 3, MidpointRounding.AwayFromZero);
                    ReportDataTable.Rows.Add(Row);
                }

                if (NoInset713Count > 0)
                {
                    DataRow Row = ReportDataTable.NewRow();
                    Row["OriginalPrice"] = NoInset713Price;
                    Row["UNN"] = UNN;
                    Row["AccountingName"] = AccountingName;
                    Row["InvNumber"] = InvNumber;
                    Row["MarketingCost"] = MarketingCost;
                    Row["Measure"] = Measure;
                    Row["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        Row["TPSCurCode"] = TPSCurrencyCode;
                    Row["Count"] = NoInset713Count;
                    Row["Cost"] = Decimal.Round(NoInset713Count * NoInset713Price, 3, MidpointRounding.AwayFromZero);
                    Row["Weight"] = Decimal.Round(NoInset713Weight, 3, MidpointRounding.AwayFromZero);
                    ReportDataTable.Rows.Add(Row);
                }

                if (Vitrina713Count > 0)
                {
                    DataRow Row = ReportDataTable.NewRow();
                    Row["OriginalPrice"] = Vitrina713Price;
                    Row["UNN"] = UNN;
                    Row["AccountingName"] = AccountingName;
                    Row["InvNumber"] = InvNumber;
                    Row["MarketingCost"] = MarketingCost;
                    Row["Measure"] = Measure;
                    Row["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        Row["TPSCurCode"] = TPSCurrencyCode;
                    Row["Count"] = Vitrina713Count;
                    Row["Cost"] = Decimal.Round(Vitrina713Count * Vitrina713Price, 3, MidpointRounding.AwayFromZero);
                    Row["Weight"] = Decimal.Round(Vitrina713Weight, 3, MidpointRounding.AwayFromZero);
                    ReportDataTable.Rows.Add(Row);
                }

                if (Solid910Count > 0)
                {
                    DataRow Row = ReportDataTable.NewRow();
                    Row["OriginalPrice"] = Solid910Price;
                    Row["UNN"] = UNN;
                    Row["AccountingName"] = AccountingName;
                    Row["InvNumber"] = InvNumber;
                    Row["MarketingCost"] = MarketingCost;
                    Row["Measure"] = Measure;
                    Row["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        Row["TPSCurCode"] = TPSCurrencyCode;
                    Row["Count"] = Solid910Count;
                    Row["Cost"] = Decimal.Round(Solid910Count * Solid910Price, 3, MidpointRounding.AwayFromZero);
                    Row["Weight"] = Decimal.Round(Solid910Weight, 3, MidpointRounding.AwayFromZero);
                    ReportDataTable.Rows.Add(Row);
                }

                if (Filenka910Count > 0)
                {
                    DataRow Row = ReportDataTable.NewRow();
                    Row["OriginalPrice"] = Filenka910Price;
                    Row["UNN"] = UNN;
                    Row["AccountingName"] = AccountingName;
                    Row["InvNumber"] = InvNumber;
                    Row["MarketingCost"] = MarketingCost;
                    Row["Measure"] = Measure;
                    Row["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        Row["TPSCurCode"] = TPSCurrencyCode;
                    Row["Count"] = Filenka910Count;
                    Row["Cost"] = Decimal.Round(Filenka910Count * Filenka910Price, 3, MidpointRounding.AwayFromZero);
                    Row["Weight"] = Decimal.Round(Filenka910Weight, 3, MidpointRounding.AwayFromZero);
                    ReportDataTable.Rows.Add(Row);
                }

                if (NoInset910Count > 0)
                {
                    DataRow Row = ReportDataTable.NewRow();
                    Row["OriginalPrice"] = NoInset910Price;
                    Row["UNN"] = UNN;
                    Row["AccountingName"] = AccountingName;
                    Row["InvNumber"] = InvNumber;
                    Row["MarketingCost"] = MarketingCost;
                    Row["Measure"] = Measure;
                    Row["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        Row["TPSCurCode"] = TPSCurrencyCode;
                    Row["Count"] = NoInset910Count;
                    Row["Cost"] = Decimal.Round(NoInset910Count * NoInset910Price, 3, MidpointRounding.AwayFromZero);
                    Row["Weight"] = Decimal.Round(NoInset910Weight, 3, MidpointRounding.AwayFromZero);
                    ReportDataTable.Rows.Add(Row);
                }

                if (Vitrina910Count > 0)
                {
                    DataRow Row = ReportDataTable.NewRow();
                    Row["OriginalPrice"] = Vitrina910Price;
                    Row["UNN"] = UNN;
                    Row["AccountingName"] = AccountingName;
                    Row["InvNumber"] = InvNumber;
                    Row["MarketingCost"] = MarketingCost;
                    Row["Measure"] = Measure;
                    Row["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        Row["TPSCurCode"] = TPSCurrencyCode;
                    Row["Count"] = Vitrina910Count;
                    Row["Cost"] = Decimal.Round(Vitrina910Count * Vitrina910Price, 3, MidpointRounding.AwayFromZero);
                    Row["Weight"] = Decimal.Round(Vitrina910Weight, 3, MidpointRounding.AwayFromZero);
                    ReportDataTable.Rows.Add(Row);
                }
            }

            Fronts.Dispose();
        }

        private void GetGridsByClient(DataTable OrdersDataTable, DataTable ReportDataTable, DataTable ReportDataTable1, int FactoryID1)
        {
            DataRow[] Rows = OrdersDataTable.Select("InsetTypeID IN (685,686,687,688,29470,29471) AND FrontID <> 3729");

            if (Rows.Count() == 0)
                return;

            decimal MarketingCost = 0;
            decimal CountPP = 0;
            decimal CostPP = 0;

            int fID = Convert.ToInt32(OrdersDataTable.Rows[0]["FactoryID"]);
            if (Rows.Count() > 0)
                MarketingCost = GetGridMarketingCost(Convert.ToInt32(Rows[0]["FrontConfigID"]));

            for (int i = 0; i < Rows.Count(); i++)
            {
                decimal d = GetInsetSquare(Convert.ToInt32(Rows[i]["FrontID"]), Convert.ToInt32(Rows[i]["Height"]),
                                            Convert.ToInt32(Rows[i]["Width"])) * Convert.ToDecimal(Rows[i]["Count"]);
                CountPP += d;
                CostPP += Convert.ToDecimal(Rows[i]["InsetPrice"]) * d;
                if (MarketingCost > GetGridMarketingCost(Convert.ToInt32(Rows[i]["FrontConfigID"])))
                    MarketingCost = GetGridMarketingCost(Convert.ToInt32(Rows[i]["FrontConfigID"]));
            }

            if (CountPP > 0)
            {
                DataRow[] rows = DecorInvNumbersDT.Select("FrontsOrdersID = " + Convert.ToInt32(OrdersDataTable.Select("InsetTypeID IN (685,686,687,688,29470,29471)")[0]["FrontsOrdersID"]));
                int FactoryID = 0;
                if (rows.Count() > 0)
                    FactoryID = Convert.ToInt32(rows[0]["FactoryID"]);
                if (FactoryID == 0)
                    FactoryID = FactoryID1;
                if (FactoryID != FactoryID1)
                {
                    DataRow NewRow = ReportDataTable1.NewRow();
                    NewRow["OriginalPrice"] = CostPP / CountPP;
                    NewRow["UNN"] = UNN;
                    NewRow["AccountingName"] = OrdersDataTable.Rows[0]["AccountingName"].ToString();
                    NewRow["InvNumber"] = OrdersDataTable.Rows[0]["InvNumber"].ToString();
                    NewRow["MarketingCost"] = MarketingCost;
                    NewRow["Measure"] = OrdersDataTable.Rows[0]["Measure"].ToString();
                    NewRow["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        NewRow["TPSCurCode"] = TPSCurrencyCode;
                    NewRow["Count"] = Decimal.Round(CountPP, 3, MidpointRounding.AwayFromZero);
                    NewRow["Cost"] = Decimal.Round(CostPP, 3, MidpointRounding.AwayFromZero);
                    NewRow["Weight"] = Decimal.Round(Convert.ToDecimal(NewRow["Count"]) * Convert.ToDecimal(3.5), 3, MidpointRounding.AwayFromZero);
                    if (rows.Count() > 0)
                    {
                        NewRow["AccountingName"] = rows[0]["DecorAccountingName"].ToString();
                        NewRow["InvNumber"] = rows[0]["DecorInvNumber"].ToString();
                    }
                    ReportDataTable1.Rows.Add(NewRow);
                }
                else
                {
                    DataRow NewRow = ReportDataTable.NewRow();
                    NewRow["OriginalPrice"] = CostPP / CountPP;
                    NewRow["UNN"] = UNN;
                    NewRow["AccountingName"] = OrdersDataTable.Rows[0]["AccountingName"].ToString();
                    NewRow["InvNumber"] = OrdersDataTable.Rows[0]["InvNumber"].ToString();
                    NewRow["MarketingCost"] = MarketingCost;
                    NewRow["Measure"] = OrdersDataTable.Rows[0]["Measure"].ToString();
                    NewRow["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        NewRow["TPSCurCode"] = TPSCurrencyCode;
                    NewRow["Count"] = Decimal.Round(CountPP, 3, MidpointRounding.AwayFromZero);
                    NewRow["Cost"] = Decimal.Round(CostPP, 3, MidpointRounding.AwayFromZero);
                    NewRow["Weight"] = Decimal.Round(Convert.ToDecimal(NewRow["Count"]) * Convert.ToDecimal(3.5), 3, MidpointRounding.AwayFromZero);
                    if (rows.Count() > 0)
                    {
                        NewRow["AccountingName"] = rows[0]["DecorAccountingName"].ToString();
                        NewRow["InvNumber"] = rows[0]["DecorInvNumber"].ToString();
                    }
                    ReportDataTable.Rows.Add(NewRow);
                }
            }
        }

        private void GetGlassByClient(DataTable OrdersDataTable, DataTable ReportDataTable, DataTable ReportDataTable1, int FactoryID1)
        {
            decimal MarketingCost = 0;
            string AccountingName = OrdersDataTable.Rows[0]["AccountingName"].ToString();
            string InvNumber = OrdersDataTable.Rows[0]["InvNumber"].ToString();
            string Measure = OrdersDataTable.Rows[0]["Measure"].ToString();
            decimal CountFlutes = 0;
            decimal CountLacomat = 0;
            //decimal CountMaster = 0;
            decimal CountKrizet = 0;
            decimal CountOther = 0;

            decimal PriceFlutes = 0;
            decimal PriceLacomat = 0;
            //decimal PriceMaster = 0;
            decimal PriceKrizet = 0;
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

            DataRow[] FRows = OrdersDataTable.Select("InsetColorID = 3944");

            if (FRows.Count() > 0)
                MarketingCost = GetMarketingCost(Convert.ToInt32(FRows[0]["FrontConfigID"]));
            if (FRows.Count() > 0)
            {
                PriceFlutes = Convert.ToDecimal(FRows[0]["InsetPrice"]);

                for (int i = 0; i < FRows.Count(); i++)
                {
                    if (IsAluminium(FRows[i]) != -1)
                        continue;

                    CountFlutes += Convert.ToDecimal(FRows[i]["Count"]) *
                                GetInsetSquare(Convert.ToInt32(FRows[i]["FrontID"]),
                                              Convert.ToInt32(FRows[i]["Height"]),
                                              Convert.ToInt32(FRows[i]["Width"]));
                    if (MarketingCost > GetMarketingCost(Convert.ToInt32(FRows[i]["FrontConfigID"])))
                        MarketingCost = GetMarketingCost(Convert.ToInt32(FRows[i]["FrontConfigID"]));
                }

                CountFlutes = Decimal.Round(CountFlutes, 3, MidpointRounding.AwayFromZero);

                if (CountFlutes > 0)
                {
                    DataRow[] rows = DecorInvNumbersDT.Select("FrontsOrdersID = " + Convert.ToInt32(FRows[0]["FrontsOrdersID"]));
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
                        NewRow["AccountingName"] = AccountingName;
                        NewRow["InvNumber"] = InvNumber;
                        NewRow["MarketingCost"] = MarketingCost;
                        NewRow["Measure"] = Measure;
                        if (rows.Count() > 0)
                        {
                            NewRow["AccountingName"] = rows[0]["DecorAccountingName"].ToString();
                            NewRow["InvNumber"] = rows[0]["DecorInvNumber"].ToString();
                        }
                        NewRow["CurrencyCode"] = ProfilCurrencyCode;
                        if (fID == 2)
                            NewRow["TPSCurCode"] = TPSCurrencyCode;
                        NewRow["Count"] = CountFlutes;
                        NewRow["Cost"] = Decimal.Round(CountFlutes * PriceFlutes);
                        NewRow["Weight"] = Decimal.Round(CountFlutes * 10, 3, MidpointRounding.AwayFromZero);
                        ReportDataTable1.Rows.Add(NewRow);
                    }
                    else
                    {
                        DataRow NewRow = ReportDataTable.NewRow();
                        //NewRow["Name"] = "Стекло Флутес";
                        NewRow["OriginalPrice"] = PriceFlutes;
                        NewRow["UNN"] = UNN;
                        NewRow["AccountingName"] = AccountingName;
                        NewRow["InvNumber"] = InvNumber;
                        NewRow["MarketingCost"] = MarketingCost;
                        NewRow["Measure"] = Measure;
                        if (rows.Count() > 0)
                        {
                            NewRow["AccountingName"] = rows[0]["DecorAccountingName"].ToString();
                            NewRow["InvNumber"] = rows[0]["DecorInvNumber"].ToString();
                        }
                        NewRow["CurrencyCode"] = ProfilCurrencyCode;
                        if (fID == 2)
                            NewRow["TPSCurCode"] = TPSCurrencyCode;
                        NewRow["Count"] = CountFlutes;
                        NewRow["Cost"] = Decimal.Round(CountFlutes * PriceFlutes);
                        NewRow["Weight"] = Decimal.Round(CountFlutes * 10, 3, MidpointRounding.AwayFromZero);
                        ReportDataTable.Rows.Add(NewRow);
                    }
                }
            }

            DataRow[] LRows = OrdersDataTable.Select("InsetColorID = 3943");

            if (LRows.Count() > 0)
                MarketingCost = GetMarketingCost(Convert.ToInt32(LRows[0]["FrontConfigID"]));
            if (LRows.Count() > 0)
            {
                PriceLacomat = Convert.ToDecimal(LRows[0]["InsetPrice"]);

                for (int i = 0; i < LRows.Count(); i++)
                {
                    if (IsAluminium(LRows[i]) != -1)
                        continue;

                    CountLacomat += Convert.ToDecimal(LRows[i]["Count"]) *
                                GetInsetSquare(Convert.ToInt32(LRows[i]["FrontID"]),
                                              Convert.ToInt32(LRows[i]["Height"]),
                                              Convert.ToInt32(LRows[i]["Width"]));
                    if (MarketingCost > GetMarketingCost(Convert.ToInt32(LRows[i]["FrontConfigID"])))
                        MarketingCost = GetMarketingCost(Convert.ToInt32(LRows[i]["FrontConfigID"]));
                }

                CountLacomat = Decimal.Round(CountLacomat, 3, MidpointRounding.AwayFromZero);

                if (CountLacomat > 0)
                {
                    DataRow[] rows = DecorInvNumbersDT.Select("FrontsOrdersID = " + Convert.ToInt32(LRows[0]["FrontsOrdersID"]));
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
                        NewRow["AccountingName"] = AccountingName;
                        NewRow["InvNumber"] = InvNumber;
                        NewRow["MarketingCost"] = MarketingCost;
                        NewRow["Measure"] = Measure;
                        if (rows.Count() > 0)
                        {
                            NewRow["AccountingName"] = rows[0]["DecorAccountingName"].ToString();
                            NewRow["InvNumber"] = rows[0]["DecorInvNumber"].ToString();
                        }
                        NewRow["CurrencyCode"] = ProfilCurrencyCode;
                        if (fID == 2)
                            NewRow["TPSCurCode"] = TPSCurrencyCode;
                        NewRow["Count"] = CountLacomat;
                        NewRow["Cost"] = Decimal.Round(CountLacomat * PriceLacomat);
                        NewRow["Weight"] = Decimal.Round(CountLacomat * 10, 3, MidpointRounding.AwayFromZero);
                        ReportDataTable1.Rows.Add(NewRow);
                    }
                    else
                    {
                        DataRow NewRow = ReportDataTable.NewRow();
                        //NewRow["Name"] = "Стекло Лакомат";
                        NewRow["OriginalPrice"] = PriceLacomat;
                        NewRow["UNN"] = UNN;
                        NewRow["AccountingName"] = AccountingName;
                        NewRow["InvNumber"] = InvNumber;
                        NewRow["MarketingCost"] = MarketingCost;
                        NewRow["Measure"] = Measure;
                        if (rows.Count() > 0)
                        {
                            NewRow["AccountingName"] = rows[0]["DecorAccountingName"].ToString();
                            NewRow["InvNumber"] = rows[0]["DecorInvNumber"].ToString();
                        }
                        NewRow["CurrencyCode"] = ProfilCurrencyCode;
                        if (fID == 2)
                            NewRow["TPSCurCode"] = TPSCurrencyCode;
                        NewRow["Count"] = CountLacomat;
                        NewRow["Cost"] = Decimal.Round(CountLacomat * PriceLacomat);
                        NewRow["Weight"] = Decimal.Round(CountLacomat * 10, 3, MidpointRounding.AwayFromZero);
                        ReportDataTable.Rows.Add(NewRow);
                    }
                }
            }

            DataRow[] KRows = OrdersDataTable.Select("InsetColorID = 3945");

            if (KRows.Count() > 0)
                MarketingCost = GetMarketingCost(Convert.ToInt32(KRows[0]["FrontConfigID"]));
            if (KRows.Count() > 0)
            {
                PriceKrizet = Convert.ToDecimal(KRows[0]["InsetPrice"]);

                for (int i = 0; i < KRows.Count(); i++)
                {
                    if (IsAluminium(KRows[i]) != -1)
                        continue;

                    CountKrizet += Convert.ToDecimal(KRows[i]["Count"]) *
                                GetInsetSquare(Convert.ToInt32(KRows[i]["FrontID"]),
                                              Convert.ToInt32(KRows[i]["Height"]),
                                              Convert.ToInt32(KRows[i]["Width"]));
                    if (MarketingCost > GetMarketingCost(Convert.ToInt32(KRows[i]["FrontConfigID"])))
                        MarketingCost = GetMarketingCost(Convert.ToInt32(KRows[i]["FrontConfigID"]));
                }

                CountKrizet = Decimal.Round(CountKrizet, 3, MidpointRounding.AwayFromZero);

                if (CountKrizet > 0)
                {
                    DataRow[] rows = DecorInvNumbersDT.Select("FrontsOrdersID = " + Convert.ToInt32(KRows[0]["FrontsOrdersID"]));
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
                        NewRow["AccountingName"] = AccountingName;
                        NewRow["InvNumber"] = InvNumber;
                        NewRow["MarketingCost"] = MarketingCost;
                        NewRow["Measure"] = Measure;
                        if (rows.Count() > 0)
                        {
                            NewRow["AccountingName"] = rows[0]["DecorAccountingName"].ToString();
                            NewRow["InvNumber"] = rows[0]["DecorInvNumber"].ToString();
                        }
                        NewRow["CurrencyCode"] = ProfilCurrencyCode;
                        if (fID == 2)
                            NewRow["TPSCurCode"] = TPSCurrencyCode;
                        NewRow["Count"] = CountKrizet;
                        NewRow["Cost"] = Decimal.Round(CountKrizet * PriceKrizet);
                        NewRow["Weight"] = Decimal.Round(CountKrizet * 10, 3, MidpointRounding.AwayFromZero);
                        ReportDataTable1.Rows.Add(NewRow);
                    }
                    else
                    {
                        DataRow NewRow = ReportDataTable.NewRow();
                        //NewRow["Name"] = "Стекло Кризет";
                        NewRow["OriginalPrice"] = PriceKrizet;
                        NewRow["UNN"] = UNN;
                        NewRow["AccountingName"] = AccountingName;
                        NewRow["InvNumber"] = InvNumber;
                        NewRow["MarketingCost"] = MarketingCost;
                        NewRow["Measure"] = Measure;
                        if (rows.Count() > 0)
                        {
                            NewRow["AccountingName"] = rows[0]["DecorAccountingName"].ToString();
                            NewRow["InvNumber"] = rows[0]["DecorInvNumber"].ToString();
                        }
                        NewRow["CurrencyCode"] = ProfilCurrencyCode;
                        if (fID == 2)
                            NewRow["TPSCurCode"] = TPSCurrencyCode;
                        NewRow["Count"] = CountKrizet;
                        NewRow["Cost"] = Decimal.Round(CountKrizet * PriceKrizet);
                        NewRow["Weight"] = Decimal.Round(CountKrizet * 10, 3, MidpointRounding.AwayFromZero);
                        ReportDataTable.Rows.Add(NewRow);
                    }
                }
            }

            DataRow[] ORows = OrdersDataTable.Select("InsetTypeID = 18");

            if (ORows.Count() > 0)
                MarketingCost = GetMarketingCost(Convert.ToInt32(ORows[0]["FrontConfigID"]));
            if (ORows.Count() > 0)
            {
                for (int i = 0; i < ORows.Count(); i++)
                {
                    if (IsAluminium(ORows[i]) != -1)
                        continue;

                    CountOther += Convert.ToDecimal(ORows[i]["Count"]) *
                                GetInsetSquare(Convert.ToInt32(ORows[i]["FrontID"]),
                                              Convert.ToInt32(ORows[i]["Height"]),
                                              Convert.ToInt32(ORows[i]["Width"]));
                    if (MarketingCost > GetMarketingCost(Convert.ToInt32(ORows[i]["FrontConfigID"])))
                        MarketingCost = GetMarketingCost(Convert.ToInt32(ORows[i]["FrontConfigID"]));
                }

                CountOther = Decimal.Round(CountOther, 3, MidpointRounding.AwayFromZero);

                if (CountOther > 0)
                {
                    DataRow[] rows = DecorInvNumbersDT.Select("FrontsOrdersID = " + Convert.ToInt32(ORows[0]["FrontsOrdersID"]));
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
                        NewRow["AccountingName"] = AccountingName;
                        NewRow["InvNumber"] = InvNumber;
                        NewRow["MarketingCost"] = MarketingCost;
                        NewRow["Measure"] = Measure;
                        if (rows.Count() > 0)
                        {
                            NewRow["AccountingName"] = rows[0]["DecorAccountingName"].ToString();
                            NewRow["InvNumber"] = rows[0]["DecorInvNumber"].ToString();
                        }
                        NewRow["CurrencyCode"] = ProfilCurrencyCode;
                        if (fID == 2)
                            NewRow["TPSCurCode"] = TPSCurrencyCode;
                        NewRow["Count"] = CountOther;
                        NewRow["Cost"] = 0;
                        NewRow["Weight"] = Decimal.Round(CountOther * 10, 3, MidpointRounding.AwayFromZero);
                        ReportDataTable1.Rows.Add(NewRow);
                    }
                    else
                    {
                        DataRow NewRow = ReportDataTable.NewRow();
                        //NewRow["Name"] = "Стекло другое";
                        NewRow["OriginalPrice"] = 0;
                        NewRow["UNN"] = UNN;
                        NewRow["AccountingName"] = AccountingName;
                        NewRow["InvNumber"] = InvNumber;
                        NewRow["MarketingCost"] = MarketingCost;
                        NewRow["Measure"] = Measure;
                        if (rows.Count() > 0)
                        {
                            NewRow["AccountingName"] = rows[0]["DecorAccountingName"].ToString();
                            NewRow["InvNumber"] = rows[0]["DecorInvNumber"].ToString();
                        }
                        NewRow["CurrencyCode"] = ProfilCurrencyCode;
                        if (fID == 2)
                            NewRow["TPSCurCode"] = TPSCurrencyCode;
                        NewRow["Count"] = CountOther;
                        NewRow["Cost"] = 0;
                        NewRow["Weight"] = Decimal.Round(CountOther * 10, 3, MidpointRounding.AwayFromZero);
                        ReportDataTable.Rows.Add(NewRow);
                    }
                }
            }
        }

        private void GetInsetsByClient(DataTable OrdersDataTable, DataTable ReportDataTable)
        {
            decimal MarketingCost = 0;
            decimal CountEllipseGrid = 0;
            decimal PriceEllipseGrid = 0;
            int fID = Convert.ToInt32(OrdersDataTable.Rows[0]["FactoryID"]);

            DataRow[] EGRows = OrdersDataTable.Select("FrontID IN (3729)");//ellipse grid

            if (EGRows.Count() > 0)
            {
                {
                    int MarginHeight = 0;
                    int MarginWidth = 0;

                    GetGlassMarginAluminium(EGRows[0], ref MarginHeight, ref MarginWidth);

                    PriceEllipseGrid = Convert.ToDecimal(EGRows[0]["InsetPrice"]);

                    MarketingCost = GetMarketingCost(Convert.ToInt32(EGRows[0]["FrontConfigID"]));
                    for (int i = 0; i < EGRows.Count(); i++)
                    {
                        decimal dd = Decimal.Round(Convert.ToDecimal(EGRows[i]["Count"]) * MarginHeight * (Convert.ToDecimal(EGRows[i]["Width"]) - MarginWidth) / 1000000, 3, MidpointRounding.AwayFromZero);
                        CountEllipseGrid += dd;
                        if (MarketingCost > GetMarketingCost(Convert.ToInt32(EGRows[i]["FrontConfigID"])))
                            MarketingCost = GetMarketingCost(Convert.ToInt32(EGRows[i]["FrontConfigID"]));
                    }

                    decimal Weight = GetInsetWeight(EGRows[0]);
                    Weight = Decimal.Round(CountEllipseGrid * Weight, 3, MidpointRounding.AwayFromZero);

                    DataRow NewRow = ReportDataTable.NewRow();
                    NewRow["OriginalPrice"] = PriceEllipseGrid;
                    NewRow["MarketingCost"] = MarketingCost;
                    NewRow["UNN"] = UNN;
                    NewRow["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        NewRow["TPSCurCode"] = TPSCurrencyCode;
                    NewRow["Count"] = CountEllipseGrid;
                    NewRow["Cost"] = CountEllipseGrid * PriceEllipseGrid;
                    NewRow["Weight"] = 0;
                    NewRow["Measure"] = OrdersDataTable.Rows[0]["Measure"].ToString();
                    DataRow[] rows = DecorInvNumbersDT.Select("FrontsOrdersID = " + Convert.ToInt32(EGRows[0]["FrontsOrdersID"]));
                    if (rows.Count() > 0)
                    {
                        NewRow["AccountingName"] = rows[0]["DecorAccountingName"].ToString();
                        NewRow["InvNumber"] = rows[0]["DecorInvNumber"].ToString();
                    }
                    ReportDataTable.Rows.Add(NewRow);
                }
                DataTable DT = OrdersDataTable.Clone();
                foreach (DataRow item in EGRows)
                    DT.Rows.Add(item.ItemArray);
                DataTable ClientDT = new DataTable();
                using (DataView DV = new DataView(DT))
                {
                    ClientDT = DV.ToTable(true, new string[] { "ClientID" });
                }
                for (int x = 0; x < ClientDT.Rows.Count; x++)
                {
                    MarketingCost = 0;
                    CountEllipseGrid = 0;
                    PriceEllipseGrid = 0;
                    int ClientID = Convert.ToInt32(ClientDT.Rows[x]["FactoryID"]);
                    string ClientName = ClientDT.Rows[x]["ClientName"].ToString();
                    EGRows = OrdersDataTable.Select("FrontID IN (3729) AND ClientID=" + ClientID);//ellipse grid
                    int MarginHeight = 0;
                    int MarginWidth = 0;

                    GetGlassMarginAluminium(EGRows[0], ref MarginHeight, ref MarginWidth);
                    PriceEllipseGrid = Convert.ToDecimal(EGRows[0]["InsetPrice"]);
                    MarketingCost = GetMarketingCost(Convert.ToInt32(EGRows[0]["FrontConfigID"]));
                    for (int i = 0; i < EGRows.Count(); i++)
                    {
                        decimal dd = Decimal.Round(Convert.ToDecimal(EGRows[i]["Count"]) * MarginHeight * (Convert.ToDecimal(EGRows[i]["Width"]) - MarginWidth) / 1000000, 3, MidpointRounding.AwayFromZero);
                        CountEllipseGrid += dd;
                        if (MarketingCost > GetMarketingCost(Convert.ToInt32(EGRows[i]["FrontConfigID"])))
                            MarketingCost = GetMarketingCost(Convert.ToInt32(EGRows[i]["FrontConfigID"]));
                    }

                    decimal Weight = GetInsetWeight(EGRows[0]);
                    Weight = Decimal.Round(CountEllipseGrid * Weight, 3, MidpointRounding.AwayFromZero);

                    DataRow NewRow = ReportDataTable.NewRow();
                    NewRow["OriginalPrice"] = PriceEllipseGrid;
                    NewRow["MarketingCost"] = MarketingCost;
                    NewRow["UNN"] = UNN;
                    NewRow["CurrencyCode"] = ProfilCurrencyCode;
                    if (fID == 2)
                        NewRow["TPSCurCode"] = TPSCurrencyCode;
                    NewRow["Count"] = CountEllipseGrid;
                    NewRow["Cost"] = CountEllipseGrid * PriceEllipseGrid;
                    NewRow["Weight"] = 0;
                    NewRow["ClientName"] = ClientName;
                    NewRow["Measure"] = OrdersDataTable.Rows[0]["Measure"].ToString();
                    DataRow[] rows = DecorInvNumbersDT.Select("FrontsOrdersID = " + Convert.ToInt32(EGRows[0]["FrontsOrdersID"]));
                    if (rows.Count() > 0)
                    {
                        NewRow["AccountingName"] = rows[0]["DecorAccountingName"].ToString();
                        NewRow["InvNumber"] = rows[0]["DecorInvNumber"].ToString();
                    }
                    ReportDataTable.Rows.Add(NewRow);
                }
            }
        }

        private decimal GetInsetWeight(DataRow FrontsOrdersRow)
        {
            int InsetTypeID = Convert.ToInt32(FrontsOrdersRow["InsetTypeID"]);
            if (InsetTypeID == 2)
                InsetTypeID = Convert.ToInt32(FrontsOrdersRow["InsetColorID"]);//стекло
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
            if (FrontsConfigRow[0]["measureId"].ToString() == "3")
            {
                outFrontWeight = PackWeight +
                    Convert.ToDecimal(FrontsOrdersRow["Count"]) * Convert.ToDecimal(FrontsConfigRow[0]["Weight"]);
                return;
            }
            //если алюминий
            if (IsAluminium(FrontsOrdersRow) > -1)
            {
                outFrontWeight = PackWeight +
                    GetAluminiumWeight(FrontsOrdersRow, true);
                return;
            }
            decimal ResultProfileWeight = GetProfileWeight(FrontsOrdersRow);
            decimal ResultInsetWeight = 0;
            DataRow[] rows = InsetTypesDataTable.Select("GroupID = 2 OR GroupID = 3 OR GroupID = 4 OR GroupID = 7 OR GroupID = 12 OR GroupID = 13");
            foreach (DataRow item in rows)
            {
                if (FrontsOrdersRow["InsetTypeID"].ToString() == item["InsetTypeID"].ToString())
                {
                    ResultInsetWeight = GetInsetWeight(FrontsOrdersRow);
                }
            }
            if (Convert.ToInt32(FrontsOrdersRow["FrontID"]) == 3729)
                ResultInsetWeight = GetInsetWeight(FrontsOrdersRow);

            outFrontWeight = PackWeight + ResultProfileWeight * Convert.ToDecimal(FrontsOrdersRow["Count"]);
            outInsetWeight = ResultInsetWeight * Convert.ToDecimal(FrontsOrdersRow["Count"]);
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
            //int g = 0;

            //for (int i = 0; i < InvCount; i++)
            //{
            //    InvDataTables[i] = OrdersDataTable.Clone();

            //    string InvNumber = "";

            //    for (int r = g; r < OrdersDataTable.Rows.Count; r++)
            //    {
            //        if (InvNumber == "")
            //        {
            //            InvDataTables[i].ImportRow(OrdersDataTable.DefaultView[r].Row);
            //            InvNumber = OrdersDataTable.DefaultView[r].Row["InvNumber"].ToString();
            //            continue;
            //        }

            //        if (InvNumber == OrdersDataTable.DefaultView[r].Row["InvNumber"].ToString())
            //        {
            //            InvDataTables[i].ImportRow(OrdersDataTable.DefaultView[r].Row);
            //        }
            //        else
            //        {
            //            g = r;
            //            break;
            //        }
            //    }
            //}

            DataTable Table = new DataTable();
            using (DataView DV = new DataView(OrdersDataTable))
            {
                Table = DV.ToTable(true, new string[] { "InvNumber" });
            }

            for (int i = 0; i < Table.Rows.Count; i++)
            {
                InvDataTables[i] = OrdersDataTable.Clone();
                string rr = Table.Rows[i]["InvNumber"].ToString();
                DataRow[] ItemsRows = OrdersDataTable.Select("InvNumber = '" + rr + "'");
                foreach (DataRow item in ItemsRows)
                {
                    InvDataTables[i].ImportRow(item);
                }
            }

            return InvDataTables;
        }

        public void ExpReport(DateTime date1, DateTime date2, bool ZOV, bool IsSample, bool IsNotSample, ArrayList MClients, ArrayList MClientGroups)
        {
            string MClientFilter = string.Empty;
            if (MClients.Count > 0)
            {
                if (MClientFilter.Length > 0)
                    MClientFilter += " AND ClientID IN (" + string.Join(",", MClients.OfType<Int32>().ToArray()) + ")";
                else
                    MClientFilter = " WHERE ClientID IN (" + string.Join(",", MClients.OfType<Int32>().ToArray()) + ")";
            }
            if (MClientGroups.Count > 0)
            {
                if (MClientFilter.Length > 0)
                    MClientFilter += " AND ClientID IN" +
                        " (SELECT ClientID FROM infiniu2_marketingreference.dbo.Clients" +
                        " WHERE ClientGroupID IN (" + string.Join(",", MClientGroups.OfType<Int32>().ToArray()) + "))";
                else
                    MClientFilter = " WHERE ClientID IN" +
                        " (SELECT ClientID FROM infiniu2_marketingreference.dbo.Clients" +
                        " WHERE ClientGroupID IN (" + string.Join(",", MClientGroups.OfType<Int32>().ToArray()) + "))";
            }
            if (MClients.Count < 1 && MClientGroups.Count < 1)
                MClientFilter = " WHERE ClientID = -1";

            ProfilFrontsOrdersDataTable.Clear();
            TPSFrontsOrdersDataTable.Clear();
            DecorInvNumbersDT.Clear();
            string ConnectionString = ConnectionStrings.MarketingOrdersConnectionString;
            if (ZOV)
                ConnectionString = ConnectionStrings.ZOVOrdersConnectionString;

            string Filter = " (ExpeditionDateTime >= '" + date1.ToString("yyyy-MM-dd") +
                " 00:00' AND ExpeditionDateTime <= '" + date2.ToString("yyyy-MM-dd") + " 23:59')";

            string MFSampleFilter = string.Empty;
            string MDSampleFilter = string.Empty;
            if (!IsSample || !IsNotSample)
            {
                if (IsSample)
                    MFSampleFilter = " AND FrontsOrders.IsSample = 1";
                if (IsNotSample)
                    MFSampleFilter = " AND FrontsOrders.IsSample = 0";
                if (IsSample)
                    MDSampleFilter = " AND (DecorOrders.IsSample = 1 OR (DecorOrders.ProductID=42 AND DecorOrders.IsSample = 0)) ";
                if (IsNotSample)
                    MDSampleFilter = " AND DecorOrders.IsSample = 0";
            }

            string SelectCommand = @"SELECT dbo.FrontsOrders.FrontsOrdersID, dbo.FrontsOrders.FrontID, dbo.FrontsOrders.ColorID, dbo.FrontsOrders.PatinaID, dbo.FrontsOrders.InsetTypeID, dbo.FrontsOrders.InsetColorID,
                dbo.FrontsOrders.TechnoColorID, dbo.FrontsOrders.TechnoInsetTypeID, dbo.FrontsOrders.TechnoInsetColorID, dbo.FrontsOrders.Height, dbo.FrontsOrders.Width, PackageDetails.Count,
                dbo.FrontsOrders.FrontConfigID, infiniu2_catalog.dbo.FrontsConfig.measureId, dbo.FrontsOrders.FactoryID, dbo.FrontsOrders.FrontPrice, dbo.FrontsOrders.InsetPrice,
                (FrontsOrders.Square * PackageDetails.Count / FrontsOrders.Count) AS Square, dbo.FrontsOrders.Cost, infiniu2_catalog.dbo.TechStore.Cvet, infiniu2_catalog.dbo.Patina.Patina, infiniu2_catalog.dbo.FrontsConfig.AccountingName, infiniu2_catalog.dbo.FrontsConfig.InvNumber, infiniu2_catalog.dbo.Measures.Measure, dbo.FrontsOrders.CreateDateTime FROM PackageDetails
                INNER JOIN FrontsOrders ON PackageDetails.OrderID = FrontsOrders.FrontsOrdersID" + MFSampleFilter + @" AND FrontsOrders.MainOrderID IN (SELECT MainOrderID FROM MainOrders WHERE MegaOrderID IN (SELECT MegaOrderID FROM MegaOrders " + MClientFilter + @"))
                INNER JOIN infiniu2_catalog.dbo.FrontsConfig ON FrontsOrders.FrontConfigID = infiniu2_catalog.dbo.FrontsConfig.FrontConfigID  LEFT JOIN
                infiniu2_catalog.dbo.TechStore ON infiniu2_catalog.dbo.FrontsConfig.ColorID = infiniu2_catalog.dbo.TechStore.TechStoreID LEFT JOIN
                infiniu2_catalog.dbo.Patina ON infiniu2_catalog.dbo.FrontsConfig.PatinaID = infiniu2_catalog.dbo.Patina.PatinaID
                INNER JOIN infiniu2_catalog.dbo.Measures ON infiniu2_catalog.dbo.FrontsConfig.MeasureID=infiniu2_catalog.dbo.Measures.MeasureID
                WHERE InvNumber IS NOT NULL AND PackageID IN (SELECT PackageID FROM Packages WHERE ProductType = 0 AND " + Filter + ") ORDER BY InvNumber";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    if (DA.Fill(DT) == 0)
                        return;

                    ProfilFrontsOrdersDataTable = DT.Clone();
                    TPSFrontsOrdersDataTable = DT.Clone();

                    SplitTables(DT, ref ProfilFrontsOrdersDataTable, ref TPSFrontsOrdersDataTable);

                    if (ProfilFrontsOrdersDataTable.Rows.Count > 0)
                    {
                        DataTable[] DTs = GroupInvNumber(ProfilFrontsOrdersDataTable);

                        for (int i = 0; i < DTs.Count(); i++)
                        {
                            GetSimpleFronts(DTs[i], ProfilReportDataTable);

                            GetCurvedFronts(DTs[i], ProfilReportDataTable);

                            GetGrids(DTs[i], ProfilReportDataTable, TPSReportDataTable, 1);

                            GetInsets(DTs[i], ProfilReportDataTable);

                            GetGlass(DTs[i], ProfilReportDataTable, TPSReportDataTable, 1);
                        }
                    }

                    if (TPSFrontsOrdersDataTable.Rows.Count > 0)
                    {
                        DataTable[] DTs = GroupInvNumber(TPSFrontsOrdersDataTable);

                        for (int i = 0; i < DTs.Count(); i++)
                        {
                            GetSimpleFronts(DTs[i], TPSReportDataTable);

                            GetCurvedFronts(DTs[i], TPSReportDataTable);

                            GetGrids(DTs[i], TPSReportDataTable, ProfilReportDataTable, 2);

                            GetInsets(DTs[i], TPSReportDataTable);

                            GetGlass(DTs[i], TPSReportDataTable, ProfilReportDataTable, 2);
                        }
                    }
                }
            }
        }
        
        public void DispReport(DateTime date1, DateTime date2, bool ZOV, bool IsSample, bool IsNotSample, ArrayList MClients, ArrayList MClientGroups)
        {
            string MClientFilter = string.Empty;
            if (MClients.Count > 0)
            {
                if (MClientFilter.Length > 0)
                    MClientFilter += " AND ClientID IN (" + string.Join(",", MClients.OfType<Int32>().ToArray()) + ")";
                else
                    MClientFilter = " WHERE ClientID IN (" + string.Join(",", MClients.OfType<Int32>().ToArray()) + ")";
            }
            if (MClientGroups.Count > 0)
            {
                if (MClientFilter.Length > 0)
                    MClientFilter += " AND ClientID IN" +
                        " (SELECT ClientID FROM infiniu2_marketingreference.dbo.Clients" +
                        " WHERE ClientGroupID IN (" + string.Join(",", MClientGroups.OfType<Int32>().ToArray()) + "))";
                else
                    MClientFilter = " WHERE ClientID IN" +
                        " (SELECT ClientID FROM infiniu2_marketingreference.dbo.Clients" +
                        " WHERE ClientGroupID IN (" + string.Join(",", MClientGroups.OfType<Int32>().ToArray()) + "))";
            }
            if (MClients.Count < 1 && MClientGroups.Count < 1)
                MClientFilter = " WHERE ClientID = -1";

            ProfilFrontsOrdersDataTable.Clear();
            TPSFrontsOrdersDataTable.Clear();
            DecorInvNumbersDT.Clear();
            string ConnectionString = ConnectionStrings.MarketingOrdersConnectionString;
            if (ZOV)
                ConnectionString = ConnectionStrings.ZOVOrdersConnectionString;

            string Filter = " (DispatchDateTime >= '" + date1.ToString("yyyy-MM-dd") +
                " 00:00' AND DispatchDateTime <= '" + date2.ToString("yyyy-MM-dd") + " 23:59')";

            string MFSampleFilter = string.Empty;
            string MDSampleFilter = string.Empty;
            if (!IsSample || !IsNotSample)
            {
                if (IsSample)
                    MFSampleFilter = " AND FrontsOrders.IsSample = 1";
                if (IsNotSample)
                    MFSampleFilter = " AND FrontsOrders.IsSample = 0";
                if (IsSample)
                    MDSampleFilter = " AND (DecorOrders.IsSample = 1 OR (DecorOrders.ProductID=42 AND DecorOrders.IsSample = 0)) ";
                if (IsNotSample)
                    MDSampleFilter = " AND DecorOrders.IsSample = 0";
            }

            string SelectCommand = @"SELECT dbo.FrontsOrders.FrontsOrdersID, dbo.FrontsOrders.FrontID, dbo.FrontsOrders.ColorID, dbo.FrontsOrders.PatinaID, dbo.FrontsOrders.InsetTypeID, dbo.FrontsOrders.InsetColorID,
                dbo.FrontsOrders.TechnoColorID, dbo.FrontsOrders.TechnoInsetTypeID, dbo.FrontsOrders.TechnoInsetColorID, dbo.FrontsOrders.Height, dbo.FrontsOrders.Width, PackageDetails.Count,
                dbo.FrontsOrders.FrontConfigID, infiniu2_catalog.dbo.FrontsConfig.measureId, dbo.FrontsOrders.FactoryID, dbo.FrontsOrders.FrontPrice, dbo.FrontsOrders.InsetPrice,
                (FrontsOrders.Square * PackageDetails.Count / FrontsOrders.Count) AS Square, dbo.FrontsOrders.Cost, infiniu2_catalog.dbo.TechStore.Cvet, infiniu2_catalog.dbo.Patina.Patina, infiniu2_catalog.dbo.FrontsConfig.AccountingName, infiniu2_catalog.dbo.FrontsConfig.InvNumber, infiniu2_catalog.dbo.Measures.Measure, dbo.FrontsOrders.CreateDateTime FROM PackageDetails
                INNER JOIN FrontsOrders ON PackageDetails.OrderID = FrontsOrders.FrontsOrdersID" + MFSampleFilter + @" AND FrontsOrders.MainOrderID IN (SELECT MainOrderID FROM MainOrders WHERE MegaOrderID IN (SELECT MegaOrderID FROM MegaOrders " + MClientFilter + @"))
                INNER JOIN infiniu2_catalog.dbo.FrontsConfig ON FrontsOrders.FrontConfigID = infiniu2_catalog.dbo.FrontsConfig.FrontConfigID LEFT JOIN
                infiniu2_catalog.dbo.TechStore ON infiniu2_catalog.dbo.FrontsConfig.ColorID = infiniu2_catalog.dbo.TechStore.TechStoreID LEFT JOIN
                infiniu2_catalog.dbo.Patina ON infiniu2_catalog.dbo.FrontsConfig.PatinaID = infiniu2_catalog.dbo.Patina.PatinaID
                INNER JOIN infiniu2_catalog.dbo.Measures ON infiniu2_catalog.dbo.FrontsConfig.MeasureID=infiniu2_catalog.dbo.Measures.MeasureID
                WHERE InvNumber IS NOT NULL AND PackageID IN (SELECT PackageID FROM Packages WHERE ProductType = 0 AND " + Filter + ") ORDER BY InvNumber";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    if (DA.Fill(DT) == 0)
                        return;

                    ProfilFrontsOrdersDataTable = DT.Clone();
                    TPSFrontsOrdersDataTable = DT.Clone();

                    SplitTables(DT, ref ProfilFrontsOrdersDataTable, ref TPSFrontsOrdersDataTable);

                    if (ProfilFrontsOrdersDataTable.Rows.Count > 0)
                    {
                        DataTable[] DTs = GroupInvNumber(ProfilFrontsOrdersDataTable);

                        for (int i = 0; i < DTs.Count(); i++)
                        {
                            GetSimpleFronts(DTs[i], ProfilReportDataTable);

                            GetCurvedFronts(DTs[i], ProfilReportDataTable);

                            GetGrids(DTs[i], ProfilReportDataTable, TPSReportDataTable, 1);

                            GetInsets(DTs[i], ProfilReportDataTable);

                            GetGlass(DTs[i], ProfilReportDataTable, TPSReportDataTable, 1);
                        }
                    }

                    if (TPSFrontsOrdersDataTable.Rows.Count > 0)
                    {
                        DataTable[] DTs = GroupInvNumber(TPSFrontsOrdersDataTable);

                        for (int i = 0; i < DTs.Count(); i++)
                        {
                            GetSimpleFronts(DTs[i], TPSReportDataTable);

                            GetCurvedFronts(DTs[i], TPSReportDataTable);

                            GetGrids(DTs[i], TPSReportDataTable, ProfilReportDataTable, 2);

                            GetInsets(DTs[i], TPSReportDataTable);

                            GetGlass(DTs[i], TPSReportDataTable, ProfilReportDataTable, 2);
                        }
                    }
                }
            }
        }

        public void StoreReport(DateTime date1, DateTime date2, bool ZOV, bool IsSample, bool IsNotSample, ArrayList MClients, ArrayList MClientGroups)
        {
            string MClientFilter = string.Empty;
            if (MClients.Count > 0)
            {
                if (MClientFilter.Length > 0)
                    MClientFilter += " AND ClientID IN (" + string.Join(",", MClients.OfType<Int32>().ToArray()) + ")";
                else
                    MClientFilter = " WHERE ClientID IN (" + string.Join(",", MClients.OfType<Int32>().ToArray()) + ")";
            }
            if (MClientGroups.Count > 0)
            {
                if (MClientFilter.Length > 0)
                    MClientFilter += " AND ClientID IN" +
                        " (SELECT ClientID FROM infiniu2_marketingreference.dbo.Clients" +
                        " WHERE ClientGroupID IN (" + string.Join(",", MClientGroups.OfType<Int32>().ToArray()) + "))";
                else
                    MClientFilter = " WHERE ClientID IN" +
                        " (SELECT ClientID FROM infiniu2_marketingreference.dbo.Clients" +
                        " WHERE ClientGroupID IN (" + string.Join(",", MClientGroups.OfType<Int32>().ToArray()) + "))";
            }
            if (MClients.Count < 1 && MClientGroups.Count < 1)
                MClientFilter = " WHERE ClientID = -1";

            ProfilFrontsOrdersDataTable.Clear();
            TPSFrontsOrdersDataTable.Clear();
            DecorInvNumbersDT.Clear();
            string ConnectionString = ConnectionStrings.MarketingOrdersConnectionString;
            if (ZOV)
                ConnectionString = ConnectionStrings.ZOVOrdersConnectionString;

            string Filter = " (StorageDateTime >= '" + date1.ToString("yyyy-MM-dd") +
                " 00:00' AND StorageDateTime <= '" + date2.ToString("yyyy-MM-dd") + " 23:59')";

            string MFSampleFilter = string.Empty;
            string MDSampleFilter = string.Empty;
            if (!IsSample || !IsNotSample)
            {
                if (IsSample)
                    MFSampleFilter = " AND FrontsOrders.IsSample = 1";
                if (IsNotSample)
                    MFSampleFilter = " AND FrontsOrders.IsSample = 0";
                if (IsSample)
                    MDSampleFilter = " AND (DecorOrders.IsSample = 1 OR (DecorOrders.ProductID=42 AND DecorOrders.IsSample = 0)) ";
                if (IsNotSample)
                    MDSampleFilter = " AND DecorOrders.IsSample = 0";
            }

            string SelectCommand = @"SELECT dbo.FrontsOrders.FrontsOrdersID, dbo.FrontsOrders.FrontID, dbo.FrontsOrders.ColorID, dbo.FrontsOrders.PatinaID, dbo.FrontsOrders.InsetTypeID, dbo.FrontsOrders.InsetColorID,
                dbo.FrontsOrders.TechnoColorID, dbo.FrontsOrders.TechnoInsetTypeID, dbo.FrontsOrders.TechnoInsetColorID, dbo.FrontsOrders.Height, dbo.FrontsOrders.Width, PackageDetails.Count,
                dbo.FrontsOrders.FrontConfigID, infiniu2_catalog.dbo.FrontsConfig.measureId, dbo.FrontsOrders.FactoryID, dbo.FrontsOrders.FrontPrice, dbo.FrontsOrders.InsetPrice,
                (FrontsOrders.Square * PackageDetails.Count / FrontsOrders.Count) AS Square, dbo.FrontsOrders.Cost, infiniu2_catalog.dbo.TechStore.Cvet, infiniu2_catalog.dbo.Patina.Patina, infiniu2_catalog.dbo.FrontsConfig.AccountingName, infiniu2_catalog.dbo.FrontsConfig.InvNumber, infiniu2_catalog.dbo.Measures.Measure, dbo.FrontsOrders.CreateDateTime FROM PackageDetails
                INNER JOIN FrontsOrders ON PackageDetails.OrderID = FrontsOrders.FrontsOrdersID" + MFSampleFilter + @" AND FrontsOrders.MainOrderID IN (SELECT MainOrderID FROM MainOrders WHERE MegaOrderID IN (SELECT MegaOrderID FROM MegaOrders " + MClientFilter + @"))
                INNER JOIN infiniu2_catalog.dbo.FrontsConfig ON FrontsOrders.FrontConfigID = infiniu2_catalog.dbo.FrontsConfig.FrontConfigID LEFT JOIN
                infiniu2_catalog.dbo.TechStore ON infiniu2_catalog.dbo.FrontsConfig.ColorID = infiniu2_catalog.dbo.TechStore.TechStoreID LEFT JOIN
                infiniu2_catalog.dbo.Patina ON infiniu2_catalog.dbo.FrontsConfig.PatinaID = infiniu2_catalog.dbo.Patina.PatinaID
                INNER JOIN infiniu2_catalog.dbo.Measures ON infiniu2_catalog.dbo.FrontsConfig.MeasureID=infiniu2_catalog.dbo.Measures.MeasureID
                WHERE InvNumber IS NOT NULL AND PackageID IN (SELECT PackageID FROM Packages WHERE ProductType = 0 AND " + Filter + ") ORDER BY InvNumber";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    if (DA.Fill(DT) == 0)
                        return;

                    ProfilFrontsOrdersDataTable = DT.Clone();
                    TPSFrontsOrdersDataTable = DT.Clone();

                    SplitTables(DT, ref ProfilFrontsOrdersDataTable, ref TPSFrontsOrdersDataTable);

                    if (ProfilFrontsOrdersDataTable.Rows.Count > 0)
                    {
                        DataTable[] DTs = GroupInvNumber(ProfilFrontsOrdersDataTable);

                        for (int i = 0; i < DTs.Count(); i++)
                        {
                            GetSimpleFronts(DTs[i], ProfilReportDataTable);

                            GetCurvedFronts(DTs[i], ProfilReportDataTable);

                            GetGrids(DTs[i], ProfilReportDataTable, TPSReportDataTable, 1);

                            GetInsets(DTs[i], ProfilReportDataTable);

                            GetGlass(DTs[i], ProfilReportDataTable, TPSReportDataTable, 1);
                        }
                    }

                    if (TPSFrontsOrdersDataTable.Rows.Count > 0)
                    {
                        DataTable[] DTs = GroupInvNumber(TPSFrontsOrdersDataTable);

                        for (int i = 0; i < DTs.Count(); i++)
                        {
                            GetSimpleFronts(DTs[i], TPSReportDataTable);

                            GetCurvedFronts(DTs[i], TPSReportDataTable);

                            GetGrids(DTs[i], TPSReportDataTable, ProfilReportDataTable, 2);

                            GetInsets(DTs[i], TPSReportDataTable);

                            GetGlass(DTs[i], TPSReportDataTable, ProfilReportDataTable, 2);
                        }
                    }
                }
            }
        }

        public void PackReport(DateTime date1, DateTime date2, bool ZOV, bool IsSample, bool IsNotSample, ArrayList MClients, ArrayList MClientGroups)
        {
            string MClientFilter = string.Empty;
            if (MClients.Count > 0)
            {
                if (MClientFilter.Length > 0)
                    MClientFilter += " AND ClientID IN (" + string.Join(",", MClients.OfType<Int32>().ToArray()) + ")";
                else
                    MClientFilter = " WHERE ClientID IN (" + string.Join(",", MClients.OfType<Int32>().ToArray()) + ")";
            }
            if (MClientGroups.Count > 0)
            {
                if (MClientFilter.Length > 0)
                    MClientFilter += " AND ClientID IN" +
                        " (SELECT ClientID FROM infiniu2_marketingreference.dbo.Clients" +
                        " WHERE ClientGroupID IN (" + string.Join(",", MClientGroups.OfType<Int32>().ToArray()) + "))";
                else
                    MClientFilter = " WHERE ClientID IN" +
                        " (SELECT ClientID FROM infiniu2_marketingreference.dbo.Clients" +
                        " WHERE ClientGroupID IN (" + string.Join(",", MClientGroups.OfType<Int32>().ToArray()) + "))";
            }
            if (MClients.Count < 1 && MClientGroups.Count < 1)
                MClientFilter = " WHERE ClientID = -1";

            ProfilFrontsOrdersDataTable.Clear();
            TPSFrontsOrdersDataTable.Clear();
            DecorInvNumbersDT.Clear();
            string ConnectionString = ConnectionStrings.MarketingOrdersConnectionString;
            if (ZOV)
                ConnectionString = ConnectionStrings.ZOVOrdersConnectionString;

            string Filter = " (PackingDateTime >= '" + date1.ToString("yyyy-MM-dd") +
                " 00:00' AND PackingDateTime <= '" + date2.ToString("yyyy-MM-dd") + " 23:59')";

            string MFSampleFilter = string.Empty;
            string MDSampleFilter = string.Empty;
            if (!IsSample || !IsNotSample)
            {
                if (IsSample)
                    MFSampleFilter = " AND FrontsOrders.IsSample = 1";
                if (IsNotSample)
                    MFSampleFilter = " AND FrontsOrders.IsSample = 0";
                if (IsSample)
                    MDSampleFilter = " AND (DecorOrders.IsSample = 1 OR (DecorOrders.ProductID=42 AND DecorOrders.IsSample = 0)) ";
                if (IsNotSample)
                    MDSampleFilter = " AND DecorOrders.IsSample = 0";
            }

            string SelectCommand = @"SELECT dbo.FrontsOrders.FrontsOrdersID, dbo.FrontsOrders.FrontID, dbo.FrontsOrders.ColorID, dbo.FrontsOrders.PatinaID, dbo.FrontsOrders.InsetTypeID, dbo.FrontsOrders.InsetColorID,
                dbo.FrontsOrders.TechnoColorID, dbo.FrontsOrders.TechnoInsetTypeID, dbo.FrontsOrders.TechnoInsetColorID, dbo.FrontsOrders.Height, dbo.FrontsOrders.Width, PackageDetails.Count,
                dbo.FrontsOrders.FrontConfigID, infiniu2_catalog.dbo.FrontsConfig.measureId, dbo.FrontsOrders.FactoryID, dbo.FrontsOrders.FrontPrice, dbo.FrontsOrders.InsetPrice,
                (FrontsOrders.Square * PackageDetails.Count / FrontsOrders.Count) AS Square, dbo.FrontsOrders.Cost, infiniu2_catalog.dbo.TechStore.Cvet, infiniu2_catalog.dbo.Patina.Patina, infiniu2_catalog.dbo.FrontsConfig.AccountingName, infiniu2_catalog.dbo.FrontsConfig.InvNumber, infiniu2_catalog.dbo.Measures.Measure, dbo.FrontsOrders.CreateDateTime FROM PackageDetails
                INNER JOIN FrontsOrders ON PackageDetails.OrderID = FrontsOrders.FrontsOrdersID" + MFSampleFilter + @" AND FrontsOrders.MainOrderID IN (SELECT MainOrderID FROM MainOrders WHERE MegaOrderID IN (SELECT MegaOrderID FROM MegaOrders " + MClientFilter + @"))
                INNER JOIN infiniu2_catalog.dbo.FrontsConfig ON FrontsOrders.FrontConfigID = infiniu2_catalog.dbo.FrontsConfig.FrontConfigID LEFT JOIN
                infiniu2_catalog.dbo.TechStore ON infiniu2_catalog.dbo.FrontsConfig.ColorID = infiniu2_catalog.dbo.TechStore.TechStoreID LEFT JOIN
                infiniu2_catalog.dbo.Patina ON infiniu2_catalog.dbo.FrontsConfig.PatinaID = infiniu2_catalog.dbo.Patina.PatinaID
                INNER JOIN infiniu2_catalog.dbo.Measures ON infiniu2_catalog.dbo.FrontsConfig.MeasureID=infiniu2_catalog.dbo.Measures.MeasureID
                WHERE InvNumber IS NOT NULL AND PackageID IN (SELECT PackageID FROM Packages WHERE ProductType = 0 AND " + Filter + ") ORDER BY InvNumber";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    if (DA.Fill(DT) == 0)
                        return;

                    ProfilFrontsOrdersDataTable = DT.Clone();
                    TPSFrontsOrdersDataTable = DT.Clone();

                    SplitTables(DT, ref ProfilFrontsOrdersDataTable, ref TPSFrontsOrdersDataTable);

                    if (ProfilFrontsOrdersDataTable.Rows.Count > 0)
                    {
                        DataTable[] DTs = GroupInvNumber(ProfilFrontsOrdersDataTable);

                        for (int i = 0; i < DTs.Count(); i++)
                        {
                            GetSimpleFronts(DTs[i], ProfilReportDataTable);

                            GetCurvedFronts(DTs[i], ProfilReportDataTable);

                            GetGrids(DTs[i], ProfilReportDataTable, TPSReportDataTable, 1);

                            GetInsets(DTs[i], ProfilReportDataTable);

                            GetGlass(DTs[i], ProfilReportDataTable, TPSReportDataTable, 1);
                        }
                    }

                    if (TPSFrontsOrdersDataTable.Rows.Count > 0)
                    {
                        DataTable[] DTs = GroupInvNumber(TPSFrontsOrdersDataTable);

                        for (int i = 0; i < DTs.Count(); i++)
                        {
                            GetSimpleFronts(DTs[i], TPSReportDataTable);

                            GetCurvedFronts(DTs[i], TPSReportDataTable);

                            GetGrids(DTs[i], TPSReportDataTable, ProfilReportDataTable, 2);

                            GetInsets(DTs[i], TPSReportDataTable);

                            GetGlass(DTs[i], TPSReportDataTable, ProfilReportDataTable, 2);
                        }
                    }
                }
            }
        }
        
        public void OnProdReport(DateTime date1, DateTime date2, bool ZOV, bool IsSample, bool IsNotSample, ArrayList MClients, ArrayList MClientGroups)
        {
            string MClientFilter = string.Empty;

            string MMainOrdersFilter = " WHERE (CAST(ProfilOnProductionDate AS Date) >= '" + date1.ToString("yyyy-MM-dd") + "' AND CAST(ProfilOnProductionDate AS Date) <= '" + date2.ToString("yyyy-MM-dd") + "')";

            if (MClients.Count > 0)
            {
                if (MClientFilter.Length > 0)
                    MClientFilter += " AND ClientID IN (" + string.Join(",", MClients.OfType<Int32>().ToArray()) + ")";
                else
                    MClientFilter = " WHERE ClientID IN (" + string.Join(",", MClients.OfType<Int32>().ToArray()) + ")";
            }
            if (MClientGroups.Count > 0)
            {
                if (MClientFilter.Length > 0)
                    MClientFilter += " AND ClientID IN" +
                        " (SELECT ClientID FROM infiniu2_marketingreference.dbo.Clients" +
                        " WHERE ClientGroupID IN (" + string.Join(",", MClientGroups.OfType<Int32>().ToArray()) + "))";
                else
                    MClientFilter = " WHERE ClientID IN" +
                        " (SELECT ClientID FROM infiniu2_marketingreference.dbo.Clients" +
                        " WHERE ClientGroupID IN (" + string.Join(",", MClientGroups.OfType<Int32>().ToArray()) + "))";
            }
            if (MClients.Count < 1 && MClientGroups.Count < 1)
                MClientFilter = " WHERE ClientID = -1";

            ProfilFrontsOrdersDataTable.Clear();
            TPSFrontsOrdersDataTable.Clear();
            DecorInvNumbersDT.Clear();
            string ConnectionString = ConnectionStrings.MarketingOrdersConnectionString;
            if (ZOV)
                ConnectionString = ConnectionStrings.ZOVOrdersConnectionString;

            string Filter = " (PackingDateTime >= '" + date1.ToString("yyyy-MM-dd") +
                " 00:00' AND PackingDateTime <= '" + date2.ToString("yyyy-MM-dd") + " 23:59')";

            string MFSampleFilter = string.Empty;
            if (!IsSample || !IsNotSample)
            {
                if (IsSample)
                    MFSampleFilter = " AND FrontsOrders.IsSample = 1";
                if (IsNotSample)
                    MFSampleFilter = " AND FrontsOrders.IsSample = 0";
            }

            string SelectCommand = @"SELECT dbo.FrontsOrders.FrontsOrdersID, dbo.FrontsOrders.FrontID, dbo.FrontsOrders.ColorID, dbo.FrontsOrders.PatinaID, dbo.FrontsOrders.InsetTypeID, dbo.FrontsOrders.InsetColorID,
                dbo.FrontsOrders.TechnoColorID, dbo.FrontsOrders.TechnoInsetTypeID, dbo.FrontsOrders.TechnoInsetColorID, dbo.FrontsOrders.Height, dbo.FrontsOrders.Width, dbo.FrontsOrders.Count,
                dbo.FrontsOrders.FrontConfigID, infiniu2_catalog.dbo.FrontsConfig.measureId, dbo.FrontsOrders.FactoryID, dbo.FrontsOrders.FrontPrice, dbo.FrontsOrders.InsetPrice,
                FrontsOrders.Square, dbo.FrontsOrders.Cost, infiniu2_catalog.dbo.TechStore.Cvet, infiniu2_catalog.dbo.Patina.Patina, infiniu2_catalog.dbo.FrontsConfig.AccountingName, infiniu2_catalog.dbo.FrontsConfig.InvNumber, infiniu2_catalog.dbo.Measures.Measure, dbo.FrontsOrders.CreateDateTime FROM FrontsOrders
                INNER JOIN infiniu2_catalog.dbo.FrontsConfig ON FrontsOrders.FrontConfigID = infiniu2_catalog.dbo.FrontsConfig.FrontConfigID LEFT JOIN
                infiniu2_catalog.dbo.TechStore ON infiniu2_catalog.dbo.FrontsConfig.ColorID = infiniu2_catalog.dbo.TechStore.TechStoreID LEFT JOIN
                infiniu2_catalog.dbo.Patina ON infiniu2_catalog.dbo.FrontsConfig.PatinaID = infiniu2_catalog.dbo.Patina.PatinaID
                INNER JOIN infiniu2_catalog.dbo.Measures ON infiniu2_catalog.dbo.FrontsConfig.MeasureID=infiniu2_catalog.dbo.Measures.MeasureID
                WHERE FrontsOrders.MainOrderID IN (SELECT MainOrderID FROM MainOrders " + MMainOrdersFilter + @" AND MegaOrderID IN (SELECT MegaOrderID FROM MegaOrders " + MClientFilter + @")) AND InvNumber IS NOT NULL " + MFSampleFilter + " ORDER BY InvNumber";

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    if (DA.Fill(DT) == 0)
                        return;

                    ProfilFrontsOrdersDataTable = DT.Clone();
                    TPSFrontsOrdersDataTable = DT.Clone();

                    SplitTables(DT, ref ProfilFrontsOrdersDataTable, ref TPSFrontsOrdersDataTable);

                    if (ProfilFrontsOrdersDataTable.Rows.Count > 0)
                    {
                        DataTable[] DTs = GroupInvNumber(ProfilFrontsOrdersDataTable);

                        for (int i = 0; i < DTs.Count(); i++)
                        {
                            GetSimpleFronts(DTs[i], ProfilReportDataTable);

                            GetCurvedFronts(DTs[i], ProfilReportDataTable);

                            GetGrids(DTs[i], ProfilReportDataTable, TPSReportDataTable, 1);

                            GetInsets(DTs[i], ProfilReportDataTable);

                            GetGlass(DTs[i], ProfilReportDataTable, TPSReportDataTable, 1);
                        }
                    }

                    if (TPSFrontsOrdersDataTable.Rows.Count > 0)
                    {
                        DataTable[] DTs = GroupInvNumber(TPSFrontsOrdersDataTable);

                        for (int i = 0; i < DTs.Count(); i++)
                        {
                            GetSimpleFronts(DTs[i], TPSReportDataTable);

                            GetCurvedFronts(DTs[i], TPSReportDataTable);

                            GetGrids(DTs[i], TPSReportDataTable, ProfilReportDataTable, 2);

                            GetInsets(DTs[i], TPSReportDataTable);

                            GetGlass(DTs[i], TPSReportDataTable, ProfilReportDataTable, 2);
                        }
                    }
                }
            }
        }
                
        public void PlanDispReport(DateTime date1, DateTime date2, bool ZOV, bool IsSample, bool IsNotSample, ArrayList MClients, ArrayList MClientGroups)
        {
            string MClientFilter = " WHERE (CAST(ProfilDispatchDate AS Date) >= '" + date1.ToString("yyyy-MM-dd") + "' AND CAST(ProfilDispatchDate AS Date) <= '" + date2.ToString("yyyy-MM-dd") + "')";

            if (MClients.Count > 0)
            {
                if (MClientFilter.Length > 0)
                    MClientFilter += " AND ClientID IN (" + string.Join(",", MClients.OfType<Int32>().ToArray()) + ")";
                else
                    MClientFilter = " WHERE ClientID IN (" + string.Join(",", MClients.OfType<Int32>().ToArray()) + ")";
            }
            if (MClientGroups.Count > 0)
            {
                if (MClientFilter.Length > 0)
                    MClientFilter += " AND ClientID IN" +
                        " (SELECT ClientID FROM infiniu2_marketingreference.dbo.Clients" +
                        " WHERE ClientGroupID IN (" + string.Join(",", MClientGroups.OfType<Int32>().ToArray()) + "))";
                else
                    MClientFilter = " WHERE ClientID IN" +
                        " (SELECT ClientID FROM infiniu2_marketingreference.dbo.Clients" +
                        " WHERE ClientGroupID IN (" + string.Join(",", MClientGroups.OfType<Int32>().ToArray()) + "))";
            }
            if (MClients.Count < 1 && MClientGroups.Count < 1)
                MClientFilter = " WHERE ClientID = -1";

            ProfilFrontsOrdersDataTable.Clear();
            TPSFrontsOrdersDataTable.Clear();
            DecorInvNumbersDT.Clear();
            string ConnectionString = ConnectionStrings.MarketingOrdersConnectionString;
            if (ZOV)
                ConnectionString = ConnectionStrings.ZOVOrdersConnectionString;

            string Filter = " (PackingDateTime >= '" + date1.ToString("yyyy-MM-dd") +
                " 00:00' AND PackingDateTime <= '" + date2.ToString("yyyy-MM-dd") + " 23:59')";

            string MFSampleFilter = string.Empty;
            if (!IsSample || !IsNotSample)
            {
                if (IsSample)
                    MFSampleFilter = " AND FrontsOrders.IsSample = 1";
                if (IsNotSample)
                    MFSampleFilter = " AND FrontsOrders.IsSample = 0";
            }

            string SelectCommand = @"SELECT dbo.FrontsOrders.FrontsOrdersID, dbo.FrontsOrders.FrontID, dbo.FrontsOrders.ColorID, dbo.FrontsOrders.PatinaID, dbo.FrontsOrders.InsetTypeID, dbo.FrontsOrders.InsetColorID,
                dbo.FrontsOrders.TechnoColorID, dbo.FrontsOrders.TechnoInsetTypeID, dbo.FrontsOrders.TechnoInsetColorID, dbo.FrontsOrders.Height, dbo.FrontsOrders.Width, dbo.FrontsOrders.Count,
                dbo.FrontsOrders.FrontConfigID, infiniu2_catalog.dbo.FrontsConfig.measureId, dbo.FrontsOrders.FactoryID, dbo.FrontsOrders.FrontPrice, dbo.FrontsOrders.InsetPrice,
                FrontsOrders.Square, dbo.FrontsOrders.Cost, infiniu2_catalog.dbo.TechStore.Cvet, infiniu2_catalog.dbo.Patina.Patina, infiniu2_catalog.dbo.FrontsConfig.AccountingName, infiniu2_catalog.dbo.FrontsConfig.InvNumber, infiniu2_catalog.dbo.Measures.Measure, dbo.FrontsOrders.CreateDateTime FROM FrontsOrders
                INNER JOIN infiniu2_catalog.dbo.FrontsConfig ON FrontsOrders.FrontConfigID = infiniu2_catalog.dbo.FrontsConfig.FrontConfigID LEFT JOIN
                infiniu2_catalog.dbo.TechStore ON infiniu2_catalog.dbo.FrontsConfig.ColorID = infiniu2_catalog.dbo.TechStore.TechStoreID LEFT JOIN
                infiniu2_catalog.dbo.Patina ON infiniu2_catalog.dbo.FrontsConfig.PatinaID = infiniu2_catalog.dbo.Patina.PatinaID
                INNER JOIN infiniu2_catalog.dbo.Measures ON infiniu2_catalog.dbo.FrontsConfig.MeasureID=infiniu2_catalog.dbo.Measures.MeasureID
                WHERE FrontsOrders.MainOrderID IN (SELECT MainOrderID FROM MainOrders WHERE MegaOrderID IN (SELECT MegaOrderID FROM MegaOrders " + MClientFilter + @")) AND InvNumber IS NOT NULL " + MFSampleFilter + " ORDER BY InvNumber";

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    if (DA.Fill(DT) == 0)
                        return;

                    ProfilFrontsOrdersDataTable = DT.Clone();
                    TPSFrontsOrdersDataTable = DT.Clone();

                    SplitTables(DT, ref ProfilFrontsOrdersDataTable, ref TPSFrontsOrdersDataTable);

                    if (ProfilFrontsOrdersDataTable.Rows.Count > 0)
                    {
                        DataTable[] DTs = GroupInvNumber(ProfilFrontsOrdersDataTable);

                        for (int i = 0; i < DTs.Count(); i++)
                        {
                            GetSimpleFronts(DTs[i], ProfilReportDataTable);

                            GetCurvedFronts(DTs[i], ProfilReportDataTable);

                            GetGrids(DTs[i], ProfilReportDataTable, TPSReportDataTable, 1);

                            GetInsets(DTs[i], ProfilReportDataTable);

                            GetGlass(DTs[i], ProfilReportDataTable, TPSReportDataTable, 1);
                        }
                    }

                    if (TPSFrontsOrdersDataTable.Rows.Count > 0)
                    {
                        DataTable[] DTs = GroupInvNumber(TPSFrontsOrdersDataTable);

                        for (int i = 0; i < DTs.Count(); i++)
                        {
                            GetSimpleFronts(DTs[i], TPSReportDataTable);

                            GetCurvedFronts(DTs[i], TPSReportDataTable);

                            GetGrids(DTs[i], TPSReportDataTable, ProfilReportDataTable, 2);

                            GetInsets(DTs[i], TPSReportDataTable);

                            GetGlass(DTs[i], TPSReportDataTable, ProfilReportDataTable, 2);
                        }
                    }
                }
            }
        }
        
        public void MarketingProduced(DateTime date1, DateTime date2, bool IsSample, bool IsNotSample, ArrayList MClients, ArrayList MClientGroups)
        {
            string MClientFilter = string.Empty;
            if (MClients.Count > 0)
            {
                MClientFilter = " WHERE ClientID IN (" + string.Join(",", MClients.OfType<Int32>().ToArray()) + ")";
            }
            if (MClientGroups.Count > 0)
            {
                MClientFilter = " WHERE ClientID IN" +
                    " (SELECT ClientID FROM infiniu2_marketingreference.dbo.Clients" +
                    " WHERE ClientGroupID IN (" + string.Join(",", MClientGroups.OfType<Int32>().ToArray()) + "))";
            }
            if (MClients.Count < 1 && MClientGroups.Count < 1)
                MClientFilter = " WHERE ClientID = -1";

            ProfilMarketingProducedDT.Clear();
            TPSMarketingProducedDT.Clear();
            ProfilFrontsOrdersDataTable.Clear();
            TPSFrontsOrdersDataTable.Clear();
            DecorInvNumbersDT.Clear();

            string Filter = " (PackingDateTime >= '" + date1.ToString("yyyy-MM-dd") +
                " 00:00' AND PackingDateTime <= '" + date2.ToString("yyyy-MM-dd") + " 23:59')";

            string MFSampleFilter = string.Empty;
            string MDSampleFilter = string.Empty;
            if (!IsSample || !IsNotSample)
            {
                if (IsSample)
                    MFSampleFilter = " AND FrontsOrders.IsSample = 1";
                if (IsNotSample)
                    MFSampleFilter = " AND FrontsOrders.IsSample = 0";
                if (IsSample)
                    MDSampleFilter = " AND (DecorOrders.IsSample = 1 OR (DecorOrders.ProductID=42 AND DecorOrders.IsSample = 0)) ";
                if (IsNotSample)
                    MDSampleFilter = " AND DecorOrders.IsSample = 0";
            }

            string SelectCommand1 = @"SELECT dbo.FrontsOrders.FrontsOrdersID, dbo.FrontsOrders.FrontID, dbo.FrontsOrders.ColorID, dbo.FrontsOrders.PatinaID, dbo.FrontsOrders.InsetTypeID, dbo.FrontsOrders.InsetColorID,
                dbo.FrontsOrders.TechnoColorID, dbo.FrontsOrders.TechnoInsetTypeID, dbo.FrontsOrders.TechnoInsetColorID, dbo.FrontsOrders.Height, dbo.FrontsOrders.Width, PackageDetails.Count,
                dbo.FrontsOrders.FrontConfigID, infiniu2_catalog.dbo.FrontsConfig.measureId, dbo.FrontsOrders.FactoryID, dbo.FrontsOrders.FrontPrice, dbo.FrontsOrders.InsetPrice,
                (FrontsOrders.Square * PackageDetails.Count / FrontsOrders.Count) AS Square, dbo.FrontsOrders.Cost, infiniu2_catalog.dbo.FrontsConfig.AccountingName, infiniu2_catalog.dbo.FrontsConfig.InvNumber, infiniu2_catalog.dbo.Measures.Measure FROM PackageDetails
                INNER JOIN FrontsOrders ON PackageDetails.OrderID = FrontsOrders.FrontsOrdersID" + MFSampleFilter + @" AND FrontsOrders.MainOrderID IN (SELECT MainOrderID FROM MainOrders WHERE MegaOrderID IN (SELECT MegaOrderID FROM MegaOrders " + MClientFilter + @"))
                INNER JOIN infiniu2_catalog.dbo.FrontsConfig ON FrontsOrders.FrontConfigID = infiniu2_catalog.dbo.FrontsConfig.FrontConfigID
                INNER JOIN infiniu2_catalog.dbo.Measures ON infiniu2_catalog.dbo.FrontsConfig.MeasureID=infiniu2_catalog.dbo.Measures.MeasureID
                WHERE InvNumber IS NOT NULL AND PackageID IN (SELECT PackageID FROM Packages WHERE FactoryID=1 AND ProductType = 0 AND " + Filter + ") ORDER BY InvNumber";
            string SelectCommand2 = @"SELECT dbo.FrontsOrders.FrontsOrdersID, dbo.FrontsOrders.FrontID, dbo.FrontsOrders.ColorID, dbo.FrontsOrders.PatinaID, dbo.FrontsOrders.InsetTypeID, dbo.FrontsOrders.InsetColorID,
                dbo.FrontsOrders.TechnoColorID, dbo.FrontsOrders.TechnoInsetTypeID, dbo.FrontsOrders.TechnoInsetColorID, dbo.FrontsOrders.Height, dbo.FrontsOrders.Width, PackageDetails.Count,
                dbo.FrontsOrders.FrontConfigID, dbo.FrontsOrders.FactoryID, dbo.FrontsOrders.FrontPrice, dbo.FrontsOrders.InsetPrice,
                (FrontsOrders.Square * PackageDetails.Count / FrontsOrders.Count) AS Square, dbo.FrontsOrders.Cost, infiniu2_catalog.dbo.FrontsConfig.AccountingName, infiniu2_catalog.dbo.FrontsConfig.InvNumber, infiniu2_catalog.dbo.Measures.Measure FROM PackageDetails
                INNER JOIN FrontsOrders ON PackageDetails.OrderID = FrontsOrders.FrontsOrdersID" + MFSampleFilter + @" AND FrontsOrders.MainOrderID IN (SELECT MainOrderID FROM MainOrders WHERE MegaOrderID IN (SELECT MegaOrderID FROM MegaOrders " + MClientFilter + @"))
                INNER JOIN infiniu2_catalog.dbo.FrontsConfig ON FrontsOrders.FrontConfigID = infiniu2_catalog.dbo.FrontsConfig.FrontConfigID
                INNER JOIN infiniu2_catalog.dbo.Measures ON infiniu2_catalog.dbo.FrontsConfig.MeasureID=infiniu2_catalog.dbo.Measures.MeasureID
                WHERE InvNumber IS NOT NULL AND PackageID IN (SELECT PackageID FROM Packages WHERE FactoryID=2 AND ProductType = 0 AND " + Filter + ") ORDER BY FrontsOrdersID";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand1, ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(ProfilFrontsOrdersDataTable);
            }
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand2, ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(TPSFrontsOrdersDataTable);
            }

            if (ProfilFrontsOrdersDataTable.Rows.Count > 0)
            {
                DataTable[] DTs = GroupInvNumber(ProfilFrontsOrdersDataTable);

                for (int i = 0; i < DTs.Count(); i++)
                {
                    if (DTs[i].Rows.Count == 0)
                        continue;
                    GetSimpleFronts(DTs[i], ProfilMarketingProducedDT);

                    GetCurvedFronts(DTs[i], ProfilMarketingProducedDT);

                    GetGrids(DTs[i], ProfilMarketingProducedDT, TPSMarketingProducedDT, 1);

                    GetInsets(DTs[i], ProfilMarketingProducedDT);

                    GetGlass(DTs[i], ProfilMarketingProducedDT, TPSMarketingProducedDT, 1);
                }
            }

            if (TPSFrontsOrdersDataTable.Rows.Count > 0)
            {
                DataTable[] DTs = GroupInvNumber(TPSFrontsOrdersDataTable);

                for (int i = 0; i < DTs.Count(); i++)
                {
                    if (DTs[i].Rows.Count == 0)
                        continue;
                    GetSimpleFronts(DTs[i], TPSMarketingProducedDT);

                    GetCurvedFronts(DTs[i], TPSMarketingProducedDT);

                    GetGrids(DTs[i], TPSMarketingProducedDT, ProfilMarketingProducedDT, 2);

                    GetInsets(DTs[i], TPSMarketingProducedDT);

                    GetGlass(DTs[i], TPSMarketingProducedDT, ProfilMarketingProducedDT, 2);
                }
            }
        }

        public void ZOVProduced(DateTime date1, DateTime date2, bool IsSample, bool IsNotSample)
        {
            ProfilZOVProducedDT.Clear();
            TPSZOVProducedDT.Clear();
            ProfilFrontsOrdersDataTable.Clear();
            TPSFrontsOrdersDataTable.Clear();
            DecorInvNumbersDT.Clear();

            string Filter = " (PackingDateTime >= '" + date1.ToString("yyyy-MM-dd") +
                " 00:00' AND PackingDateTime <= '" + date2.ToString("yyyy-MM-dd") + " 23:59')";

            string MFSampleFilter = string.Empty;
            string MDSampleFilter = string.Empty;
            if (!IsSample || !IsNotSample)
            {
                if (IsSample)
                    MFSampleFilter = " AND FrontsOrders.IsSample = 1";
                if (IsNotSample)
                    MFSampleFilter = " AND FrontsOrders.IsSample = 0";
                if (IsSample)
                    MDSampleFilter = " AND (DecorOrders.IsSample = 1 OR (DecorOrders.ProductID=42 AND DecorOrders.IsSample = 0)) ";
                if (IsNotSample)
                    MDSampleFilter = " AND DecorOrders.IsSample = 0";
            }

            string SelectCommand = @"SELECT dbo.FrontsOrders.FrontsOrdersID, dbo.FrontsOrders.FrontID, dbo.FrontsOrders.ColorID, dbo.FrontsOrders.PatinaID, dbo.FrontsOrders.InsetTypeID, dbo.FrontsOrders.InsetColorID,
                dbo.FrontsOrders.TechnoColorID, dbo.FrontsOrders.TechnoInsetTypeID, dbo.FrontsOrders.TechnoInsetColorID, dbo.FrontsOrders.Height, dbo.FrontsOrders.Width, PackageDetails.Count,
                dbo.FrontsOrders.FrontConfigID, infiniu2_catalog.dbo.FrontsConfig.measureId, dbo.FrontsOrders.FactoryID, dbo.FrontsOrders.FrontPrice, dbo.FrontsOrders.InsetPrice,
                (FrontsOrders.Square * PackageDetails.Count / FrontsOrders.Count) AS Square, dbo.FrontsOrders.Cost, infiniu2_catalog.dbo.FrontsConfig.AccountingName, infiniu2_catalog.dbo.FrontsConfig.InvNumber, infiniu2_catalog.dbo.Measures.Measure, dbo.FrontsOrders.CreateDateTime FROM PackageDetails
                INNER JOIN FrontsOrders ON PackageDetails.OrderID = FrontsOrders.FrontsOrdersID" + MFSampleFilter + @" AND FrontsOrders.MainOrderID IN (SELECT MainOrderID FROM MainOrders WHERE MegaOrderID IN (SELECT MegaOrderID FROM MegaOrders))
                INNER JOIN infiniu2_catalog.dbo.FrontsConfig ON FrontsOrders.FrontConfigID = infiniu2_catalog.dbo.FrontsConfig.FrontConfigID
                INNER JOIN infiniu2_catalog.dbo.Measures ON infiniu2_catalog.dbo.FrontsConfig.MeasureID=infiniu2_catalog.dbo.Measures.MeasureID
                WHERE InvNumber IS NOT NULL AND PackageID IN (SELECT PackageID FROM Packages WHERE ProductType = 0 AND " + Filter + ") ORDER BY InvNumber";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.ZOVOrdersConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    if (DA.Fill(DT) == 0)
                        return;

                    ProfilFrontsOrdersDataTable = DT.Clone();
                    TPSFrontsOrdersDataTable = DT.Clone();

                    SplitTables(DT, ref ProfilFrontsOrdersDataTable, ref TPSFrontsOrdersDataTable);

                    if (ProfilFrontsOrdersDataTable.Rows.Count > 0)
                    {
                        DataTable[] DTs = GroupInvNumber(ProfilFrontsOrdersDataTable);

                        for (int i = 0; i < DTs.Count(); i++)
                        {
                            if (DTs[i].Rows.Count == 0)
                                continue;
                            GetSimpleFronts(DTs[i], ProfilZOVProducedDT);

                            GetCurvedFronts(DTs[i], ProfilZOVProducedDT);

                            GetGrids(DTs[i], ProfilZOVProducedDT, TPSZOVProducedDT, 1);

                            GetInsets(DTs[i], ProfilZOVProducedDT);

                            GetGlass(DTs[i], ProfilZOVProducedDT, TPSZOVProducedDT, 1);
                        }
                    }

                    if (TPSFrontsOrdersDataTable.Rows.Count > 0)
                    {
                        DataTable[] DTs = GroupInvNumber(TPSFrontsOrdersDataTable);

                        for (int i = 0; i < DTs.Count(); i++)
                        {
                            if (DTs[i].Rows.Count == 0)
                                continue;
                            GetSimpleFronts(DTs[i], TPSZOVProducedDT);

                            GetCurvedFronts(DTs[i], TPSZOVProducedDT);

                            GetGrids(DTs[i], TPSZOVProducedDT, ProfilZOVProducedDT, 2);

                            GetInsets(DTs[i], TPSZOVProducedDT);

                            GetGlass(DTs[i], TPSZOVProducedDT, ProfilZOVProducedDT, 2);
                        }
                    }
                }
            }
        }

        public void MarketingDispatch(DateTime date1, DateTime date2, bool IsSample, bool IsNotSample, ArrayList MClients, ArrayList MClientGroups)
        {
            string MClientFilter = string.Empty;
            if (MClients.Count > 0)
            {
                MClientFilter = " WHERE ClientID IN (" + string.Join(",", MClients.OfType<Int32>().ToArray()) + ")";
            }
            if (MClientGroups.Count > 0)
            {
                MClientFilter = " WHERE ClientID IN" +
                    " (SELECT ClientID FROM infiniu2_marketingreference.dbo.Clients" +
                    " WHERE ClientGroupID IN (" + string.Join(",", MClientGroups.OfType<Int32>().ToArray()) + "))";
            }
            if (MClients.Count < 1 && MClientGroups.Count < 1)
                MClientFilter = " WHERE ClientID = -1";

            ProfilMarketingDispatchDT.Clear();
            TPSMarketingDispatchDT.Clear();
            ProfilFrontsOrdersDataTable.Clear();
            TPSFrontsOrdersDataTable.Clear();
            DecorInvNumbersDT.Clear();

            string Filter = " (DispatchDateTime >= '" + date1.ToString("yyyy-MM-dd") +
                " 00:00' AND DispatchDateTime <= '" + date2.ToString("yyyy-MM-dd") + " 23:59')";

            string MFSampleFilter = string.Empty;
            string MDSampleFilter = string.Empty;
            if (!IsSample || !IsNotSample)
            {
                if (IsSample)
                    MFSampleFilter = " AND FrontsOrders.IsSample = 1";
                if (IsNotSample)
                    MFSampleFilter = " AND FrontsOrders.IsSample = 0";
                if (IsSample)
                    MDSampleFilter = " AND (DecorOrders.IsSample = 1 OR (DecorOrders.ProductID=42 AND DecorOrders.IsSample = 0)) ";
                if (IsNotSample)
                    MDSampleFilter = " AND DecorOrders.IsSample = 0";
            }

            string SelectCommand = @"SELECT dbo.FrontsOrders.FrontsOrdersID, dbo.FrontsOrders.FrontID, dbo.FrontsOrders.ColorID, dbo.FrontsOrders.PatinaID, dbo.FrontsOrders.InsetTypeID, dbo.FrontsOrders.InsetColorID,
                dbo.FrontsOrders.TechnoColorID, dbo.FrontsOrders.TechnoInsetTypeID, dbo.FrontsOrders.TechnoInsetColorID, dbo.FrontsOrders.Height, dbo.FrontsOrders.Width, PackageDetails.Count,
                dbo.FrontsOrders.FrontConfigID, infiniu2_catalog.dbo.FrontsConfig.measureId, dbo.FrontsOrders.FactoryID, dbo.FrontsOrders.FrontPrice, dbo.FrontsOrders.InsetPrice,
                (FrontsOrders.Square * PackageDetails.Count / FrontsOrders.Count) AS Square, dbo.FrontsOrders.Cost, infiniu2_catalog.dbo.FrontsConfig.AccountingName, infiniu2_catalog.dbo.FrontsConfig.InvNumber, infiniu2_catalog.dbo.Measures.Measure, dbo.FrontsOrders.CreateDateTime FROM PackageDetails
                INNER JOIN FrontsOrders ON PackageDetails.OrderID = FrontsOrders.FrontsOrdersID" + MFSampleFilter + @" AND FrontsOrders.MainOrderID IN (SELECT MainOrderID FROM MainOrders WHERE MegaOrderID IN (SELECT MegaOrderID FROM MegaOrders " + MClientFilter + @"))
                INNER JOIN infiniu2_catalog.dbo.FrontsConfig ON FrontsOrders.FrontConfigID = infiniu2_catalog.dbo.FrontsConfig.FrontConfigID
                INNER JOIN infiniu2_catalog.dbo.Measures ON infiniu2_catalog.dbo.FrontsConfig.MeasureID=infiniu2_catalog.dbo.Measures.MeasureID
                WHERE InvNumber IS NOT NULL AND PackageID IN (SELECT PackageID FROM Packages WHERE ProductType = 0 AND " + Filter + ") ORDER BY InvNumber";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    if (DA.Fill(DT) == 0)
                        return;

                    ProfilFrontsOrdersDataTable = DT.Clone();
                    TPSFrontsOrdersDataTable = DT.Clone();

                    SplitTables(DT, ref ProfilFrontsOrdersDataTable, ref TPSFrontsOrdersDataTable);

                    if (ProfilFrontsOrdersDataTable.Rows.Count > 0)
                    {
                        DataTable[] DTs = GroupInvNumber(ProfilFrontsOrdersDataTable);

                        for (int i = 0; i < DTs.Count(); i++)
                        {
                            if (DTs[i].Rows.Count == 0)
                                continue;
                            GetSimpleFronts(DTs[i], ProfilMarketingDispatchDT);

                            GetCurvedFronts(DTs[i], ProfilMarketingDispatchDT);

                            GetGrids(DTs[i], ProfilMarketingDispatchDT, TPSMarketingDispatchDT, 1);

                            GetInsets(DTs[i], ProfilMarketingDispatchDT);

                            GetGlass(DTs[i], ProfilMarketingDispatchDT, TPSMarketingDispatchDT, 1);
                        }
                    }

                    if (TPSFrontsOrdersDataTable.Rows.Count > 0)
                    {
                        DataTable[] DTs = GroupInvNumber(TPSFrontsOrdersDataTable);

                        for (int i = 0; i < DTs.Count(); i++)
                        {
                            if (DTs[i].Rows.Count == 0)
                                continue;
                            GetSimpleFronts(DTs[i], TPSMarketingDispatchDT);

                            GetCurvedFronts(DTs[i], TPSMarketingDispatchDT);

                            GetGrids(DTs[i], TPSMarketingDispatchDT, ProfilMarketingDispatchDT, 2);

                            GetInsets(DTs[i], TPSMarketingDispatchDT);

                            GetGlass(DTs[i], TPSMarketingDispatchDT, ProfilMarketingDispatchDT, 2);
                        }
                    }
                }
            }
        }

        public void ZOVDispatch(DateTime date1, DateTime date2, bool IsSample, bool IsNotSample)
        {
            ProfilZOVDispatchDT.Clear();
            TPSZOVDispatchDT.Clear();
            ProfilFrontsOrdersDataTable.Clear();
            TPSFrontsOrdersDataTable.Clear();
            DecorInvNumbersDT.Clear();

            string Filter = " (DispatchDateTime >= '" + date1.ToString("yyyy-MM-dd") +
                " 00:00' AND DispatchDateTime <= '" + date2.ToString("yyyy-MM-dd") + " 23:59')";

            string MFSampleFilter = string.Empty;
            string MDSampleFilter = string.Empty;
            if (!IsSample || !IsNotSample)
            {
                if (IsSample)
                    MFSampleFilter = " AND FrontsOrders.IsSample = 1";
                if (IsNotSample)
                    MFSampleFilter = " AND FrontsOrders.IsSample = 0";
                if (IsSample)
                    MDSampleFilter = " AND (DecorOrders.IsSample = 1 OR (DecorOrders.ProductID=42 AND DecorOrders.IsSample = 0)) ";
                if (IsNotSample)
                    MDSampleFilter = " AND DecorOrders.IsSample = 0";
            }

            string SelectCommand = @"SELECT dbo.FrontsOrders.FrontsOrdersID, dbo.FrontsOrders.FrontID, dbo.FrontsOrders.ColorID, dbo.FrontsOrders.PatinaID, dbo.FrontsOrders.InsetTypeID, dbo.FrontsOrders.InsetColorID,
                dbo.FrontsOrders.TechnoColorID, dbo.FrontsOrders.TechnoInsetTypeID, dbo.FrontsOrders.TechnoInsetColorID, dbo.FrontsOrders.Height, dbo.FrontsOrders.Width, PackageDetails.Count,
                dbo.FrontsOrders.FrontConfigID, infiniu2_catalog.dbo.FrontsConfig.measureId, dbo.FrontsOrders.FactoryID, dbo.FrontsOrders.FrontPrice, dbo.FrontsOrders.InsetPrice,
                (FrontsOrders.Square * PackageDetails.Count / FrontsOrders.Count) AS Square, dbo.FrontsOrders.Cost, infiniu2_catalog.dbo.FrontsConfig.AccountingName, infiniu2_catalog.dbo.FrontsConfig.InvNumber, infiniu2_catalog.dbo.Measures.Measure, dbo.FrontsOrders.CreateDateTime FROM PackageDetails
                INNER JOIN FrontsOrders ON PackageDetails.OrderID = FrontsOrders.FrontsOrdersID" + MFSampleFilter + @" AND FrontsOrders.MainOrderID IN (SELECT MainOrderID FROM MainOrders WHERE MegaOrderID IN (SELECT MegaOrderID FROM MegaOrders))
                INNER JOIN infiniu2_catalog.dbo.FrontsConfig ON FrontsOrders.FrontConfigID = infiniu2_catalog.dbo.FrontsConfig.FrontConfigID
                INNER JOIN infiniu2_catalog.dbo.Measures ON infiniu2_catalog.dbo.FrontsConfig.MeasureID=infiniu2_catalog.dbo.Measures.MeasureID
                WHERE InvNumber IS NOT NULL AND PackageID IN (SELECT PackageID FROM Packages WHERE ProductType = 0 AND " + Filter + ") ORDER BY InvNumber";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.ZOVOrdersConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    if (DA.Fill(DT) == 0)
                        return;

                    ProfilFrontsOrdersDataTable = DT.Clone();
                    TPSFrontsOrdersDataTable = DT.Clone();

                    SplitTables(DT, ref ProfilFrontsOrdersDataTable, ref TPSFrontsOrdersDataTable);

                    if (ProfilFrontsOrdersDataTable.Rows.Count > 0)
                    {
                        DataTable[] DTs = GroupInvNumber(ProfilFrontsOrdersDataTable);

                        for (int i = 0; i < DTs.Count(); i++)
                        {
                            if (DTs[i].Rows.Count == 0)
                                continue;
                            GetSimpleFronts(DTs[i], ProfilZOVDispatchDT);

                            GetCurvedFronts(DTs[i], ProfilZOVProducedDT);

                            GetGrids(DTs[i], ProfilZOVProducedDT, TPSZOVProducedDT, 1);

                            GetInsets(DTs[i], ProfilZOVProducedDT);

                            GetGlass(DTs[i], ProfilZOVProducedDT, TPSZOVProducedDT, 1);
                        }
                    }

                    if (TPSFrontsOrdersDataTable.Rows.Count > 0)
                    {
                        DataTable[] DTs = GroupInvNumber(TPSFrontsOrdersDataTable);

                        for (int i = 0; i < DTs.Count(); i++)
                        {
                            if (DTs[i].Rows.Count == 0)
                                continue;
                            GetSimpleFronts(DTs[i], TPSZOVDispatchDT);

                            GetCurvedFronts(DTs[i], TPSZOVProducedDT);

                            GetGrids(DTs[i], TPSZOVProducedDT, ProfilZOVProducedDT, 2);

                            GetInsets(DTs[i], TPSZOVProducedDT);

                            GetGlass(DTs[i], TPSZOVProducedDT, ProfilZOVProducedDT, 2);
                        }
                    }
                }
            }
        }

        public void MarketingStartCount(DateTime date1, bool IsSample, bool IsNotSample, ArrayList MClients, ArrayList MClientGroups)
        {
            string MClientFilter = string.Empty;
            if (MClients.Count > 0)
            {
                MClientFilter += " AND ClientID IN (" + string.Join(",", MClients.OfType<Int32>().ToArray()) + ")";
            }
            if (MClientGroups.Count > 0)
            {
                MClientFilter += " AND ClientID IN" +
                    " (SELECT ClientID FROM infiniu2_marketingreference.dbo.Clients" +
                    " WHERE ClientGroupID IN (" + string.Join(",", MClientGroups.OfType<Int32>().ToArray()) + "))";
            }
            if (MClients.Count < 1 && MClientGroups.Count < 1)
                MClientFilter = " AND ClientID = -1";

            ProfilMarketingStartCountDT.Clear();
            TPSMarketingStartCountDT.Clear();
            ProfilFrontsOrdersDataTable.Clear();
            TPSFrontsOrdersDataTable.Clear();
            DecorInvNumbersDT.Clear();

            string MFSampleFilter = string.Empty;
            string MDSampleFilter = string.Empty;
            if (!IsSample || !IsNotSample)
            {
                if (IsSample)
                    MFSampleFilter = " AND FrontsOrders.IsSample = 1";
                if (IsNotSample)
                    MFSampleFilter = " AND FrontsOrders.IsSample = 0";
                if (IsSample)
                    MDSampleFilter = " AND (DecorOrders.IsSample = 1 OR (DecorOrders.ProductID=42 AND DecorOrders.IsSample = 0)) ";
                if (IsNotSample)
                    MDSampleFilter = " AND DecorOrders.IsSample = 0";
            }

            string SelectCommand = @"SELECT dbo.FrontsOrders.FrontsOrdersID, dbo.FrontsOrders.FrontID, dbo.FrontsOrders.ColorID, dbo.FrontsOrders.PatinaID, dbo.FrontsOrders.InsetTypeID, dbo.FrontsOrders.InsetColorID,
                dbo.FrontsOrders.TechnoColorID, dbo.FrontsOrders.TechnoInsetTypeID, dbo.FrontsOrders.TechnoInsetColorID, dbo.FrontsOrders.Height, dbo.FrontsOrders.Width, PackageDetails.Count,
                dbo.FrontsOrders.FrontConfigID, infiniu2_catalog.dbo.FrontsConfig.measureId, dbo.FrontsOrders.FactoryID, dbo.FrontsOrders.FrontPrice, dbo.FrontsOrders.InsetPrice,
                (FrontsOrders.Square * PackageDetails.Count / FrontsOrders.Count) AS Square, dbo.FrontsOrders.Cost, infiniu2_catalog.dbo.FrontsConfig.AccountingName, infiniu2_catalog.dbo.FrontsConfig.InvNumber, infiniu2_catalog.dbo.Measures.Measure, dbo.FrontsOrders.CreateDateTime, Fronts.TechStoreName, Colors.TechStoreName AS Expr35, InsetTypes.TechStoreName AS Expr36, InsetColors.TechStoreName AS Expr37,
                         TechnoInsetTypes.TechStoreName AS Expr38, TechnoInsetColors.TechStoreName AS Expr39, dbo.PackageDetails.PackageDetailID, infiniu2_marketingreference.dbo.Clients.ClientName, dbo.MegaOrders.OrderNumber
FROM            dbo.PackageDetails INNER JOIN
                         dbo.FrontsOrders ON dbo.PackageDetails.OrderID = dbo.FrontsOrders.FrontsOrdersID" + MFSampleFilter + @" LEFT OUTER JOIN
                         infiniu2_catalog.dbo.TechStore AS Fronts ON dbo.FrontsOrders.FrontID = Fronts.TechStoreID LEFT OUTER JOIN
                         infiniu2_catalog.dbo.TechStore AS Colors ON dbo.FrontsOrders.ColorID = Colors.TechStoreID LEFT OUTER JOIN
                         infiniu2_catalog.dbo.TechStore AS InsetTypes ON dbo.FrontsOrders.InsetTypeID = InsetTypes.TechStoreID LEFT OUTER JOIN
                         infiniu2_catalog.dbo.TechStore AS InsetColors ON dbo.FrontsOrders.InsetColorID = InsetColors.TechStoreID LEFT OUTER JOIN
                         infiniu2_catalog.dbo.TechStore AS TechnoInsetTypes ON dbo.FrontsOrders.TechnoInsetTypeID = TechnoInsetTypes.TechStoreID LEFT OUTER JOIN
                         infiniu2_catalog.dbo.TechStore AS TechnoInsetColors ON dbo.FrontsOrders.TechnoInsetColorID = TechnoInsetColors.TechStoreID INNER JOIN
                         infiniu2_catalog.dbo.FrontsConfig ON dbo.FrontsOrders.FrontConfigID = infiniu2_catalog.dbo.FrontsConfig.FrontConfigID
                INNER JOIN infiniu2_catalog.dbo.Measures ON infiniu2_catalog.dbo.FrontsConfig.MeasureID=infiniu2_catalog.dbo.Measures.MeasureID INNER JOIN
                         dbo.MainOrders ON dbo.FrontsOrders.MainOrderID = dbo.MainOrders.MainOrderID INNER JOIN
                         dbo.MegaOrders ON dbo.MainOrders.MegaOrderID = dbo.MegaOrders.MegaOrderID " + MClientFilter + @" INNER JOIN
                         infiniu2_marketingreference.dbo.Clients ON dbo.MegaOrders.ClientID = infiniu2_marketingreference.dbo.Clients.ClientID
WHERE dbo.PackageDetails.PackageID IN (SELECT TOP (100) PERCENT PackageID FROM dbo.Packages WHERE PackingDateTime < '" + date1.ToString("yyyy-MM-dd") + @" 00:00:00' AND ProductType = 0 AND (DispatchDateTime IS NULL OR DispatchDateTime >= '" + date1.ToString("yyyy-MM-dd") + @" 00:00:00')) ORDER BY infiniu2_marketingreference.dbo.Clients.ClientName, dbo.MegaOrders.OrderNumber";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    if (DA.Fill(DT) == 0)
                        return;

                    ProfilFrontsOrdersDataTable = DT.Clone();
                    TPSFrontsOrdersDataTable = DT.Clone();

                    SplitTables(DT, ref ProfilFrontsOrdersDataTable, ref TPSFrontsOrdersDataTable);

                    if (ProfilFrontsOrdersDataTable.Rows.Count > 0)
                    {
                        DataTable[] DTs = GroupInvNumber(ProfilFrontsOrdersDataTable);

                        for (int i = 0; i < DTs.Count(); i++)
                        {
                            if (DTs[i].Rows.Count == 0)
                                continue;
                            GetSimpleFronts(DTs[i], ProfilMarketingStartCountDT);

                            GetCurvedFronts(DTs[i], ProfilMarketingStartCountDT);

                            GetGrids(DTs[i], ProfilMarketingStartCountDT, TPSMarketingStartCountDT, 1);

                            GetInsets(DTs[i], ProfilMarketingStartCountDT);

                            GetGlass(DTs[i], ProfilMarketingStartCountDT, TPSMarketingStartCountDT, 1);
                        }
                    }
                    if (TPSFrontsOrdersDataTable.Rows.Count > 0)
                    {
                        DataTable[] DTs = GroupInvNumber(TPSFrontsOrdersDataTable);

                        for (int i = 0; i < DTs.Count(); i++)
                        {
                            if (DTs[i].Rows.Count == 0)
                                continue;
                            GetSimpleFronts(DTs[i], TPSMarketingStartCountDT);

                            GetCurvedFronts(DTs[i], TPSMarketingStartCountDT);

                            GetGrids(DTs[i], TPSMarketingStartCountDT, ProfilMarketingStartCountDT, 2);

                            GetInsets(DTs[i], TPSMarketingStartCountDT);

                            GetGlass(DTs[i], TPSMarketingStartCountDT, ProfilMarketingStartCountDT, 2);
                        }
                    }
                }
            }
        }

        public void MarketingEndCount(DateTime date1, bool IsSample, bool IsNotSample, ArrayList MClients, ArrayList MClientGroups)
        {
            string MClientFilter = string.Empty;
            if (MClients.Count > 0)
            {
                MClientFilter += " AND ClientID IN (" + string.Join(",", MClients.OfType<Int32>().ToArray()) + ")";
            }
            if (MClientGroups.Count > 0)
            {
                MClientFilter += " AND ClientID IN" +
                    " (SELECT ClientID FROM infiniu2_marketingreference.dbo.Clients" +
                    " WHERE ClientGroupID IN (" + string.Join(",", MClientGroups.OfType<Int32>().ToArray()) + "))";
            }
            if (MClients.Count < 1 && MClientGroups.Count < 1)
                MClientFilter = " AND ClientID = -1";

            ProfilMarketingEndCountDT.Clear();
            TPSMarketingEndCountDT.Clear();
            ProfilFrontsOrdersDataTable.Clear();
            TPSFrontsOrdersDataTable.Clear();
            DecorInvNumbersDT.Clear();

            string MFSampleFilter = string.Empty;
            string MDSampleFilter = string.Empty;
            if (!IsSample || !IsNotSample)
            {
                if (IsSample)
                    MFSampleFilter = " AND FrontsOrders.IsSample = 1";
                if (IsNotSample)
                    MFSampleFilter = " AND FrontsOrders.IsSample = 0";
                if (IsSample)
                    MDSampleFilter = " AND (DecorOrders.IsSample = 1 OR (DecorOrders.ProductID=42 AND DecorOrders.IsSample = 0)) ";
                if (IsNotSample)
                    MDSampleFilter = " AND DecorOrders.IsSample = 0";
            }

            string SelectCommand = @"SELECT dbo.FrontsOrders.FrontsOrdersID, dbo.FrontsOrders.FrontID, dbo.FrontsOrders.ColorID, dbo.FrontsOrders.PatinaID, dbo.FrontsOrders.InsetTypeID, dbo.FrontsOrders.InsetColorID,
                dbo.FrontsOrders.TechnoColorID, dbo.FrontsOrders.TechnoInsetTypeID, dbo.FrontsOrders.TechnoInsetColorID, dbo.FrontsOrders.Height, dbo.FrontsOrders.Width, PackageDetails.Count,
                dbo.FrontsOrders.FrontConfigID, infiniu2_catalog.dbo.FrontsConfig.measureId, dbo.FrontsOrders.FactoryID, dbo.FrontsOrders.FrontPrice, dbo.FrontsOrders.InsetPrice,
                (FrontsOrders.Square * PackageDetails.Count / FrontsOrders.Count) AS Square, dbo.FrontsOrders.Cost, infiniu2_catalog.dbo.FrontsConfig.AccountingName, infiniu2_catalog.dbo.FrontsConfig.InvNumber, infiniu2_catalog.dbo.Measures.Measure, dbo.FrontsOrders.CreateDateTime, Fronts.TechStoreName, Colors.TechStoreName AS Expr35, InsetTypes.TechStoreName AS Expr36, InsetColors.TechStoreName AS Expr37,
                         TechnoInsetTypes.TechStoreName AS Expr38, TechnoInsetColors.TechStoreName AS Expr39, dbo.PackageDetails.PackageDetailID, infiniu2_marketingreference.dbo.Clients.ClientName, dbo.MegaOrders.OrderNumber
FROM            dbo.PackageDetails INNER JOIN
                         dbo.FrontsOrders ON dbo.PackageDetails.OrderID = dbo.FrontsOrders.FrontsOrdersID" + MFSampleFilter + @" LEFT OUTER JOIN
                         infiniu2_catalog.dbo.TechStore AS Fronts ON dbo.FrontsOrders.FrontID = Fronts.TechStoreID LEFT OUTER JOIN
                         infiniu2_catalog.dbo.TechStore AS Colors ON dbo.FrontsOrders.ColorID = Colors.TechStoreID LEFT OUTER JOIN
                         infiniu2_catalog.dbo.TechStore AS InsetTypes ON dbo.FrontsOrders.InsetTypeID = InsetTypes.TechStoreID LEFT OUTER JOIN
                         infiniu2_catalog.dbo.TechStore AS InsetColors ON dbo.FrontsOrders.InsetColorID = InsetColors.TechStoreID LEFT OUTER JOIN
                         infiniu2_catalog.dbo.TechStore AS TechnoInsetTypes ON dbo.FrontsOrders.TechnoInsetTypeID = TechnoInsetTypes.TechStoreID LEFT OUTER JOIN
                         infiniu2_catalog.dbo.TechStore AS TechnoInsetColors ON dbo.FrontsOrders.TechnoInsetColorID = TechnoInsetColors.TechStoreID INNER JOIN
                         infiniu2_catalog.dbo.FrontsConfig ON dbo.FrontsOrders.FrontConfigID = infiniu2_catalog.dbo.FrontsConfig.FrontConfigID
                INNER JOIN infiniu2_catalog.dbo.Measures ON infiniu2_catalog.dbo.FrontsConfig.MeasureID=infiniu2_catalog.dbo.Measures.MeasureID INNER JOIN
                         dbo.MainOrders ON dbo.FrontsOrders.MainOrderID = dbo.MainOrders.MainOrderID INNER JOIN
                         dbo.MegaOrders ON dbo.MainOrders.MegaOrderID = dbo.MegaOrders.MegaOrderID " + MClientFilter + @" INNER JOIN
                         infiniu2_marketingreference.dbo.Clients ON dbo.MegaOrders.ClientID = infiniu2_marketingreference.dbo.Clients.ClientID
WHERE dbo.PackageDetails.PackageID IN (SELECT TOP (100) PERCENT PackageID FROM dbo.Packages WHERE PackingDateTime < '" + date1.ToString("yyyy-MM-dd") + @" 23:59:59' AND ProductType = 0 AND (DispatchDateTime IS NULL OR DispatchDateTime >= '" + date1.ToString("yyyy-MM-dd") + @" 23:59:59')) ORDER BY infiniu2_marketingreference.dbo.Clients.ClientName, dbo.MegaOrders.OrderNumber";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    if (DA.Fill(DT) == 0)
                        return;

                    ProfilFrontsOrdersDataTable = DT.Clone();
                    TPSFrontsOrdersDataTable = DT.Clone();

                    SplitTables(DT, ref ProfilFrontsOrdersDataTable, ref TPSFrontsOrdersDataTable);

                    if (ProfilFrontsOrdersDataTable.Rows.Count > 0)
                    {
                        DataTable[] DTs = GroupInvNumber(ProfilFrontsOrdersDataTable);

                        for (int i = 0; i < DTs.Count(); i++)
                        {
                            if (DTs[i].Rows.Count == 0)
                                continue;
                            GetSimpleFronts(DTs[i], ProfilMarketingEndCountDT);

                            GetCurvedFronts(DTs[i], ProfilMarketingEndCountDT);

                            GetGrids(DTs[i], ProfilMarketingEndCountDT, TPSMarketingEndCountDT, 1);

                            GetInsets(DTs[i], ProfilMarketingEndCountDT);

                            GetGlass(DTs[i], ProfilMarketingEndCountDT, TPSMarketingEndCountDT, 1);
                        }
                    }
                    if (TPSFrontsOrdersDataTable.Rows.Count > 0)
                    {
                        DataTable[] DTs = GroupInvNumber(TPSFrontsOrdersDataTable);

                        for (int i = 0; i < DTs.Count(); i++)
                        {
                            if (DTs[i].Rows.Count == 0)
                                continue;
                            GetSimpleFronts(DTs[i], TPSMarketingEndCountDT);

                            GetCurvedFronts(DTs[i], TPSMarketingEndCountDT);

                            GetGrids(DTs[i], TPSMarketingEndCountDT, ProfilMarketingEndCountDT, 2);

                            GetInsets(DTs[i], TPSMarketingEndCountDT);

                            GetGlass(DTs[i], TPSMarketingEndCountDT, ProfilMarketingEndCountDT, 2);
                        }
                    }
                }
            }
        }

        public void ZOVStartCount(DateTime date1, bool IsSample, bool IsNotSample)
        {
            ProfilZOVStartCountDT.Clear();
            TPSZOVStartCountDT.Clear();
            ProfilFrontsOrdersDataTable.Clear();
            TPSFrontsOrdersDataTable.Clear();
            DecorInvNumbersDT.Clear();

            string MFSampleFilter = string.Empty;
            string MDSampleFilter = string.Empty;
            if (!IsSample || !IsNotSample)
            {
                if (IsSample)
                    MFSampleFilter = " AND FrontsOrders.IsSample = 1";
                if (IsNotSample)
                    MFSampleFilter = " AND FrontsOrders.IsSample = 0";
                if (IsSample)
                    MDSampleFilter = " AND (DecorOrders.IsSample = 1 OR (DecorOrders.ProductID=42 AND DecorOrders.IsSample = 0)) ";
                if (IsNotSample)
                    MDSampleFilter = " AND DecorOrders.IsSample = 0";
            }

            string SelectCommand = @"SELECT dbo.FrontsOrders.FrontsOrdersID, dbo.FrontsOrders.FrontID, dbo.FrontsOrders.ColorID, dbo.FrontsOrders.PatinaID, dbo.FrontsOrders.InsetTypeID, dbo.FrontsOrders.InsetColorID,
                dbo.FrontsOrders.TechnoColorID, dbo.FrontsOrders.TechnoInsetTypeID, dbo.FrontsOrders.TechnoInsetColorID, dbo.FrontsOrders.Height, dbo.FrontsOrders.Width, PackageDetails.Count,
                dbo.FrontsOrders.FrontConfigID, infiniu2_catalog.dbo.FrontsConfig.measureId, dbo.FrontsOrders.FactoryID, dbo.FrontsOrders.FrontPrice, dbo.FrontsOrders.InsetPrice,
                (FrontsOrders.Square * PackageDetails.Count / FrontsOrders.Count) AS Square, dbo.FrontsOrders.Cost, infiniu2_catalog.dbo.FrontsConfig.AccountingName, infiniu2_catalog.dbo.FrontsConfig.InvNumber, infiniu2_catalog.dbo.Measures.Measure, dbo.FrontsOrders.CreateDateTime, Fronts.TechStoreName, Colors.TechStoreName AS Expr35, InsetTypes.TechStoreName AS Expr36, InsetColors.TechStoreName AS Expr37,
                         TechnoInsetTypes.TechStoreName AS Expr38, TechnoInsetColors.TechStoreName AS Expr39,
                         PackageDetails.PackageDetailID, infiniu2_zovreference.dbo.Clients.ClientName, MainOrders.DocNumber
FROM            dbo.PackageDetails INNER JOIN
                         FrontsOrders ON PackageDetails.OrderID = FrontsOrders.FrontsOrdersID" + MFSampleFilter + @" LEFT OUTER JOIN
                         infiniu2_catalog.dbo.TechStore AS Fronts ON FrontsOrders.FrontID = Fronts.TechStoreID LEFT OUTER JOIN
                         infiniu2_catalog.dbo.TechStore AS Colors ON FrontsOrders.ColorID = Colors.TechStoreID LEFT OUTER JOIN
                         infiniu2_catalog.dbo.TechStore AS InsetTypes ON FrontsOrders.InsetTypeID = InsetTypes.TechStoreID LEFT OUTER JOIN
                         infiniu2_catalog.dbo.TechStore AS InsetColors ON FrontsOrders.InsetColorID = InsetColors.TechStoreID LEFT OUTER JOIN
                         infiniu2_catalog.dbo.TechStore AS TechnoInsetTypes ON FrontsOrders.TechnoInsetTypeID = TechnoInsetTypes.TechStoreID LEFT OUTER JOIN
                         infiniu2_catalog.dbo.TechStore AS TechnoInsetColors ON FrontsOrders.TechnoInsetColorID = TechnoInsetColors.TechStoreID INNER JOIN
                         infiniu2_catalog.dbo.FrontsConfig ON FrontsOrders.FrontConfigID = infiniu2_catalog.dbo.FrontsConfig.FrontConfigID
                INNER JOIN infiniu2_catalog.dbo.Measures ON infiniu2_catalog.dbo.FrontsConfig.MeasureID=infiniu2_catalog.dbo.Measures.MeasureID INNER JOIN
                         MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID INNER JOIN
                         infiniu2_zovreference.dbo.Clients ON MainOrders.ClientID = infiniu2_zovreference.dbo.Clients.ClientID
WHERE        PackageDetails.PackageID IN
                             (SELECT        TOP (100) PERCENT PackageID
                               FROM            Packages
                               WHERE        PackingDateTime < '" + date1.ToString("yyyy-MM-dd") + @" 00:00:00' AND ProductType = 0 AND (DispatchDateTime IS NULL OR DispatchDateTime >= '" + date1.ToString("yyyy-MM-dd") + @" 00:00:00'))
ORDER BY infiniu2_zovreference.dbo.Clients.ClientName, MainOrders.DocNumber";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.ZOVOrdersConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    if (DA.Fill(DT) == 0)
                        return;

                    ProfilFrontsOrdersDataTable = DT.Clone();
                    TPSFrontsOrdersDataTable = DT.Clone();

                    SplitTables(DT, ref ProfilFrontsOrdersDataTable, ref TPSFrontsOrdersDataTable);

                    if (ProfilFrontsOrdersDataTable.Rows.Count > 0)
                    {
                        DataTable[] DTs = GroupInvNumber(ProfilFrontsOrdersDataTable);

                        for (int i = 0; i < DTs.Count(); i++)
                        {
                            if (DTs[i].Rows.Count == 0)
                                continue;
                            GetSimpleFronts(DTs[i], ProfilZOVStartCountDT);

                            GetCurvedFronts(DTs[i], ProfilZOVStartCountDT);

                            GetGrids(DTs[i], ProfilZOVStartCountDT, TPSZOVStartCountDT, 1);

                            GetInsets(DTs[i], ProfilZOVStartCountDT);

                            GetGlass(DTs[i], ProfilZOVStartCountDT, TPSZOVStartCountDT, 1);
                        }
                    }
                    if (TPSFrontsOrdersDataTable.Rows.Count > 0)
                    {
                        DataTable[] DTs = GroupInvNumber(TPSFrontsOrdersDataTable);

                        for (int i = 0; i < DTs.Count(); i++)
                        {
                            if (DTs[i].Rows.Count == 0)
                                continue;
                            GetSimpleFronts(DTs[i], TPSZOVStartCountDT);

                            GetCurvedFronts(DTs[i], TPSZOVStartCountDT);

                            GetGrids(DTs[i], TPSZOVStartCountDT, ProfilZOVStartCountDT, 2);

                            GetInsets(DTs[i], TPSZOVStartCountDT);

                            GetGlass(DTs[i], TPSZOVStartCountDT, ProfilZOVStartCountDT, 2);
                        }
                    }
                }
            }
        }

        public void ZOVEndCount(DateTime date1, bool IsSample, bool IsNotSample)
        {
            ProfilZOVEndCountDT.Clear();
            TPSZOVEndCountDT.Clear();
            ProfilFrontsOrdersDataTable.Clear();
            TPSFrontsOrdersDataTable.Clear();
            DecorInvNumbersDT.Clear();

            string MFSampleFilter = string.Empty;
            string MDSampleFilter = string.Empty;
            if (!IsSample || !IsNotSample)
            {
                if (IsSample)
                    MFSampleFilter = " AND FrontsOrders.IsSample = 1";
                if (IsNotSample)
                    MFSampleFilter = " AND FrontsOrders.IsSample = 0";
                if (IsSample)
                    MDSampleFilter = " AND (DecorOrders.IsSample = 1 OR (DecorOrders.ProductID=42 AND DecorOrders.IsSample = 0)) ";
                if (IsNotSample)
                    MDSampleFilter = " AND DecorOrders.IsSample = 0";
            }

            string SelectCommand = @"SELECT dbo.FrontsOrders.FrontsOrdersID, dbo.FrontsOrders.FrontID, dbo.FrontsOrders.ColorID, dbo.FrontsOrders.PatinaID, dbo.FrontsOrders.InsetTypeID, dbo.FrontsOrders.InsetColorID,
                dbo.FrontsOrders.TechnoColorID, dbo.FrontsOrders.TechnoInsetTypeID, dbo.FrontsOrders.TechnoInsetColorID, dbo.FrontsOrders.Height, dbo.FrontsOrders.Width, PackageDetails.Count,
                dbo.FrontsOrders.FrontConfigID, infiniu2_catalog.dbo.FrontsConfig.measureId, dbo.FrontsOrders.FactoryID, dbo.FrontsOrders.FrontPrice, dbo.FrontsOrders.InsetPrice,
                (FrontsOrders.Square * PackageDetails.Count / FrontsOrders.Count) AS Square, dbo.FrontsOrders.Cost, infiniu2_catalog.dbo.FrontsConfig.AccountingName, infiniu2_catalog.dbo.FrontsConfig.InvNumber, infiniu2_catalog.dbo.Measures.Measure, dbo.FrontsOrders.CreateDateTime, Fronts.TechStoreName, Colors.TechStoreName AS Expr35, InsetTypes.TechStoreName AS Expr36, InsetColors.TechStoreName AS Expr37,
                         TechnoInsetTypes.TechStoreName AS Expr38, TechnoInsetColors.TechStoreName AS Expr39,
                         PackageDetails.PackageDetailID, infiniu2_zovreference.dbo.Clients.ClientName, MainOrders.DocNumber
FROM            PackageDetails INNER JOIN
                         FrontsOrders ON PackageDetails.OrderID = FrontsOrders.FrontsOrdersID" + MFSampleFilter + @" LEFT OUTER JOIN
                         infiniu2_catalog.dbo.TechStore AS Fronts ON FrontsOrders.FrontID = Fronts.TechStoreID LEFT OUTER JOIN
                         infiniu2_catalog.dbo.TechStore AS Colors ON FrontsOrders.ColorID = Colors.TechStoreID LEFT OUTER JOIN
                         infiniu2_catalog.dbo.TechStore AS InsetTypes ON FrontsOrders.InsetTypeID = InsetTypes.TechStoreID LEFT OUTER JOIN
                         infiniu2_catalog.dbo.TechStore AS InsetColors ON FrontsOrders.InsetColorID = InsetColors.TechStoreID LEFT OUTER JOIN
                         infiniu2_catalog.dbo.TechStore AS TechnoInsetTypes ON FrontsOrders.TechnoInsetTypeID = TechnoInsetTypes.TechStoreID LEFT OUTER JOIN
                         infiniu2_catalog.dbo.TechStore AS TechnoInsetColors ON FrontsOrders.TechnoInsetColorID = TechnoInsetColors.TechStoreID INNER JOIN
                         infiniu2_catalog.dbo.FrontsConfig ON FrontsOrders.FrontConfigID = infiniu2_catalog.dbo.FrontsConfig.FrontConfigID
                INNER JOIN infiniu2_catalog.dbo.Measures ON infiniu2_catalog.dbo.FrontsConfig.MeasureID=infiniu2_catalog.dbo.Measures.MeasureID INNER JOIN
                         MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID INNER JOIN
                         infiniu2_zovreference.dbo.Clients ON MainOrders.ClientID = infiniu2_zovreference.dbo.Clients.ClientID
WHERE        PackageDetails.PackageID IN
                             (SELECT        TOP (100) PERCENT PackageID
                               FROM            Packages
                               WHERE        PackingDateTime < '" + date1.ToString("yyyy-MM-dd") + @" 23:59:59' AND ProductType = 0 AND (DispatchDateTime IS NULL OR DispatchDateTime >= '" + date1.ToString("yyyy-MM-dd") + @" 23:59:59'))
ORDER BY infiniu2_zovreference.dbo.Clients.ClientName, MainOrders.DocNumber";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.ZOVOrdersConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    if (DA.Fill(DT) == 0)
                        return;

                    ProfilFrontsOrdersDataTable = DT.Clone();
                    TPSFrontsOrdersDataTable = DT.Clone();

                    SplitTables(DT, ref ProfilFrontsOrdersDataTable, ref TPSFrontsOrdersDataTable);

                    if (ProfilFrontsOrdersDataTable.Rows.Count > 0)
                    {
                        DataTable[] DTs = GroupInvNumber(ProfilFrontsOrdersDataTable);

                        for (int i = 0; i < DTs.Count(); i++)
                        {
                            if (DTs[i].Rows.Count == 0)
                                continue;
                            GetSimpleFronts(DTs[i], ProfilZOVEndCountDT);

                            GetCurvedFronts(DTs[i], ProfilZOVEndCountDT);

                            GetGrids(DTs[i], ProfilZOVEndCountDT, TPSZOVEndCountDT, 1);

                            GetInsets(DTs[i], ProfilZOVEndCountDT);

                            GetGlass(DTs[i], ProfilZOVEndCountDT, TPSZOVEndCountDT, 1);
                        }
                    }
                    if (TPSFrontsOrdersDataTable.Rows.Count > 0)
                    {
                        DataTable[] DTs = GroupInvNumber(TPSFrontsOrdersDataTable);

                        for (int i = 0; i < DTs.Count(); i++)
                        {
                            if (DTs[i].Rows.Count == 0)
                                continue;
                            GetSimpleFronts(DTs[i], TPSZOVEndCountDT);

                            GetCurvedFronts(DTs[i], TPSZOVEndCountDT);

                            GetGrids(DTs[i], TPSZOVEndCountDT, ProfilZOVEndCountDT, 2);

                            GetInsets(DTs[i], TPSZOVEndCountDT);

                            GetGlass(DTs[i], TPSZOVEndCountDT, ProfilZOVEndCountDT, 2);
                        }
                    }
                }
            }
        }

        public void StartEndProducedDispatchReport(DateTime date1, DateTime date2, bool IsSample, bool IsNotSample, ArrayList MClients, ArrayList MClientGroups)
        {
            MarketingProduced(date1, date2, IsSample, IsNotSample, MClients, MClientGroups);
            if (date2 < new DateTime(2016, 07, 01))
                ZOVProduced(date1, date2, IsSample, IsNotSample);
            MarketingDispatch(date1, date2, IsSample, IsNotSample, MClients, MClientGroups);
            if (date2 < new DateTime(2016, 07, 01))
                ZOVDispatch(date1, date2, IsSample, IsNotSample);
            MarketingStartCount(date1, IsSample, IsNotSample, MClients, MClientGroups);
            MarketingEndCount(date2, IsSample, IsNotSample, MClients, MClientGroups);
            if (date2 < new DateTime(2016, 07, 01))
                ZOVStartCount(date1, IsSample, IsNotSample);
            if (date2 < new DateTime(2016, 07, 01))
                ZOVEndCount(date2, IsSample, IsNotSample);
            CollectProfilOrders();
            CollectTPSOrders();
        }

        private void CollectProfilOrders()
        {
            AllProfilDT.Clear();

            for (int i = 0; i < ProfilMarketingStartCountDT.Rows.Count; i++)
            {
                string rr = ProfilMarketingStartCountDT.Rows[i]["InvNumber"].ToString();
                DataRow[] ItemsRows = AllProfilDT.Select("InvNumber = '" + rr + "'");
                if (ItemsRows.Count() > 0)
                {
                    foreach (DataRow item in ItemsRows)
                    {
                        if (item["StartCount"] == DBNull.Value)
                            item["StartCount"] = ProfilMarketingStartCountDT.Rows[i]["Count"];
                        else
                            item["StartCount"] = Convert.ToDecimal(item["StartCount"]) + Convert.ToDecimal(ProfilMarketingStartCountDT.Rows[i]["Count"]);
                    }
                }
                else
                {
                    DataRow NewRow = AllProfilDT.NewRow();
                    NewRow["AccountingName"] = ProfilMarketingStartCountDT.Rows[i]["AccountingName"];
                    NewRow["Measure"] = ProfilMarketingStartCountDT.Rows[i]["Measure"];
                    NewRow["InvNumber"] = ProfilMarketingStartCountDT.Rows[i]["InvNumber"];
                    NewRow["StartCount"] = ProfilMarketingStartCountDT.Rows[i]["Count"];
                    AllProfilDT.Rows.Add(NewRow);
                }
            }
            for (int i = 0; i < ProfilMarketingEndCountDT.Rows.Count; i++)
            {
                string rr = ProfilMarketingEndCountDT.Rows[i]["InvNumber"].ToString();
                DataRow[] ItemsRows = AllProfilDT.Select("InvNumber = '" + rr + "'");
                if (ItemsRows.Count() > 0)
                {
                    foreach (DataRow item in ItemsRows)
                    {
                        if (item["EndCount"] == DBNull.Value)
                            item["EndCount"] = ProfilMarketingEndCountDT.Rows[i]["Count"];
                        else
                            item["EndCount"] = Convert.ToDecimal(item["EndCount"]) + Convert.ToDecimal(ProfilMarketingEndCountDT.Rows[i]["Count"]);
                    }
                }
                else
                {
                    DataRow NewRow = AllProfilDT.NewRow();
                    NewRow["AccountingName"] = ProfilMarketingEndCountDT.Rows[i]["AccountingName"];
                    NewRow["Measure"] = ProfilMarketingEndCountDT.Rows[i]["Measure"];
                    NewRow["InvNumber"] = ProfilMarketingEndCountDT.Rows[i]["InvNumber"];
                    NewRow["EndCount"] = ProfilMarketingEndCountDT.Rows[i]["Count"];
                    AllProfilDT.Rows.Add(NewRow);
                }
            }
            for (int i = 0; i < ProfilMarketingProducedDT.Rows.Count; i++)
            {
                string rr = ProfilMarketingProducedDT.Rows[i]["InvNumber"].ToString();
                DataRow[] ItemsRows = AllProfilDT.Select("InvNumber = '" + rr + "'");
                if (ItemsRows.Count() > 0)
                {
                    foreach (DataRow item in ItemsRows)
                    {
                        if (item["ProducedCount"] == DBNull.Value)
                            item["ProducedCount"] = ProfilMarketingProducedDT.Rows[i]["Count"];
                        else
                            item["ProducedCount"] = Convert.ToDecimal(item["ProducedCount"]) + Convert.ToDecimal(ProfilMarketingProducedDT.Rows[i]["Count"]);
                    }
                }
                else
                {
                    DataRow NewRow = AllProfilDT.NewRow();
                    NewRow["AccountingName"] = ProfilMarketingProducedDT.Rows[i]["AccountingName"];
                    NewRow["Measure"] = ProfilMarketingProducedDT.Rows[i]["Measure"];
                    NewRow["InvNumber"] = ProfilMarketingProducedDT.Rows[i]["InvNumber"];
                    NewRow["ProducedCount"] = ProfilMarketingProducedDT.Rows[i]["Count"];
                    AllProfilDT.Rows.Add(NewRow);
                }
            }
            for (int i = 0; i < ProfilMarketingDispatchDT.Rows.Count; i++)
            {
                string rr = ProfilMarketingDispatchDT.Rows[i]["InvNumber"].ToString();
                DataRow[] ItemsRows = AllProfilDT.Select("InvNumber = '" + rr + "'");
                if (ItemsRows.Count() > 0)
                {
                    foreach (DataRow item in ItemsRows)
                    {
                        if (item["DispatchCount"] == DBNull.Value)
                            item["DispatchCount"] = ProfilMarketingDispatchDT.Rows[i]["Count"];
                        else
                            item["DispatchCount"] = Convert.ToDecimal(item["DispatchCount"]) + Convert.ToDecimal(ProfilMarketingDispatchDT.Rows[i]["Count"]);
                    }
                }
                else
                {
                    DataRow NewRow = AllProfilDT.NewRow();
                    NewRow["AccountingName"] = ProfilMarketingDispatchDT.Rows[i]["AccountingName"];
                    NewRow["Measure"] = ProfilMarketingDispatchDT.Rows[i]["Measure"];
                    NewRow["InvNumber"] = ProfilMarketingDispatchDT.Rows[i]["InvNumber"];
                    NewRow["DispatchCount"] = ProfilMarketingDispatchDT.Rows[i]["Count"];
                    AllProfilDT.Rows.Add(NewRow);
                }
            }

            for (int i = 0; i < ProfilZOVStartCountDT.Rows.Count; i++)
            {
                string rr = ProfilZOVStartCountDT.Rows[i]["InvNumber"].ToString();
                DataRow[] ItemsRows = AllProfilDT.Select("InvNumber = '" + rr + "'");
                if (ItemsRows.Count() > 0)
                {
                    foreach (DataRow item in ItemsRows)
                    {
                        if (item["StartCount"] == DBNull.Value)
                            item["StartCount"] = ProfilZOVStartCountDT.Rows[i]["Count"];
                        else
                            item["StartCount"] = Convert.ToDecimal(item["StartCount"]) + Convert.ToDecimal(ProfilZOVStartCountDT.Rows[i]["Count"]);
                    }
                }
                else
                {
                    DataRow NewRow = AllProfilDT.NewRow();
                    NewRow["AccountingName"] = ProfilZOVStartCountDT.Rows[i]["AccountingName"];
                    NewRow["Measure"] = ProfilZOVStartCountDT.Rows[i]["Measure"];
                    NewRow["InvNumber"] = ProfilZOVStartCountDT.Rows[i]["InvNumber"];
                    NewRow["StartCount"] = ProfilZOVStartCountDT.Rows[i]["Count"];
                    AllProfilDT.Rows.Add(NewRow);
                }
            }
            for (int i = 0; i < ProfilZOVEndCountDT.Rows.Count; i++)
            {
                string rr = ProfilZOVEndCountDT.Rows[i]["InvNumber"].ToString();
                DataRow[] ItemsRows = AllProfilDT.Select("InvNumber = '" + rr + "'");
                if (ItemsRows.Count() > 0)
                {
                    foreach (DataRow item in ItemsRows)
                    {
                        if (item["EndCount"] == DBNull.Value)
                            item["EndCount"] = ProfilZOVEndCountDT.Rows[i]["Count"];
                        else
                            item["EndCount"] = Convert.ToDecimal(item["EndCount"]) + Convert.ToDecimal(ProfilZOVEndCountDT.Rows[i]["Count"]);
                    }
                }
                else
                {
                    DataRow NewRow = AllProfilDT.NewRow();
                    NewRow["AccountingName"] = ProfilZOVEndCountDT.Rows[i]["AccountingName"];
                    NewRow["Measure"] = ProfilZOVEndCountDT.Rows[i]["Measure"];
                    NewRow["InvNumber"] = ProfilZOVEndCountDT.Rows[i]["InvNumber"];
                    NewRow["EndCount"] = ProfilZOVEndCountDT.Rows[i]["Count"];
                    AllProfilDT.Rows.Add(NewRow);
                }
            }
            for (int i = 0; i < ProfilZOVProducedDT.Rows.Count; i++)
            {
                string rr = ProfilZOVProducedDT.Rows[i]["InvNumber"].ToString();
                DataRow[] ItemsRows = AllProfilDT.Select("InvNumber = '" + rr + "'");
                if (ItemsRows.Count() > 0)
                {
                    foreach (DataRow item in ItemsRows)
                    {
                        if (item["ProducedCount"] == DBNull.Value)
                            item["ProducedCount"] = ProfilZOVProducedDT.Rows[i]["Count"];
                        else
                            item["ProducedCount"] = Convert.ToDecimal(item["ProducedCount"]) + Convert.ToDecimal(ProfilZOVProducedDT.Rows[i]["Count"]);
                    }
                }
                else
                {
                    DataRow NewRow = AllProfilDT.NewRow();
                    NewRow["AccountingName"] = ProfilZOVProducedDT.Rows[i]["AccountingName"];
                    NewRow["Measure"] = ProfilZOVProducedDT.Rows[i]["Measure"];
                    NewRow["InvNumber"] = ProfilZOVProducedDT.Rows[i]["InvNumber"];
                    NewRow["ProducedCount"] = ProfilZOVProducedDT.Rows[i]["Count"];
                    AllProfilDT.Rows.Add(NewRow);
                }
            }
            for (int i = 0; i < ProfilZOVDispatchDT.Rows.Count; i++)
            {
                string rr = ProfilZOVDispatchDT.Rows[i]["InvNumber"].ToString();
                DataRow[] ItemsRows = AllProfilDT.Select("InvNumber = '" + rr + "'");
                if (ItemsRows.Count() > 0)
                {
                    foreach (DataRow item in ItemsRows)
                    {
                        if (item["DispatchCount"] == DBNull.Value)
                            item["DispatchCount"] = ProfilZOVDispatchDT.Rows[i]["Count"];
                        else
                            item["DispatchCount"] = Convert.ToDecimal(item["DispatchCount"]) + Convert.ToDecimal(ProfilZOVDispatchDT.Rows[i]["Count"]);
                    }
                }
                else
                {
                    DataRow NewRow = AllProfilDT.NewRow();
                    NewRow["AccountingName"] = ProfilZOVDispatchDT.Rows[i]["AccountingName"];
                    NewRow["Measure"] = ProfilZOVDispatchDT.Rows[i]["Measure"];
                    NewRow["InvNumber"] = ProfilZOVDispatchDT.Rows[i]["InvNumber"];
                    NewRow["DispatchCount"] = ProfilZOVDispatchDT.Rows[i]["Count"];
                    AllProfilDT.Rows.Add(NewRow);
                }
            }

            using (DataView DV = new DataView(AllProfilDT.Copy()))
            {
                DV.Sort = "AccountingName";
                AllProfilDT.Clear();
                AllProfilDT = DV.ToTable();
            }
            foreach (DataRow item in AllProfilDT.Rows)
            {
                if (item["StartCount"] == DBNull.Value)
                    item["StartCount"] = 0;
                if (item["EndCount"] == DBNull.Value)
                    item["EndCount"] = 0;
                if (item["ProducedCount"] == DBNull.Value)
                    item["ProducedCount"] = 0;
                if (item["DispatchCount"] == DBNull.Value)
                    item["DispatchCount"] = 0;
            }
        }

        public void CollectTPSOrders()
        {
            AllTPSDT.Clear();

            for (int i = 0; i < TPSMarketingStartCountDT.Rows.Count; i++)
            {
                string rr = TPSMarketingStartCountDT.Rows[i]["InvNumber"].ToString();
                DataRow[] ItemsRows = AllTPSDT.Select("InvNumber = '" + rr + "'");
                if (ItemsRows.Count() > 0)
                {
                    foreach (DataRow item in ItemsRows)
                    {
                        if (item["StartCount"] == DBNull.Value)
                            item["StartCount"] = TPSMarketingStartCountDT.Rows[i]["Count"];
                        else
                            item["StartCount"] = Convert.ToDecimal(item["StartCount"]) + Convert.ToDecimal(TPSMarketingStartCountDT.Rows[i]["Count"]);
                    }
                }
                else
                {
                    DataRow NewRow = AllTPSDT.NewRow();
                    NewRow["AccountingName"] = TPSMarketingStartCountDT.Rows[i]["AccountingName"];
                    NewRow["Measure"] = TPSMarketingStartCountDT.Rows[i]["Measure"];
                    NewRow["InvNumber"] = TPSMarketingStartCountDT.Rows[i]["InvNumber"];
                    NewRow["StartCount"] = TPSMarketingStartCountDT.Rows[i]["Count"];
                    AllTPSDT.Rows.Add(NewRow);
                }
            }
            for (int i = 0; i < TPSMarketingEndCountDT.Rows.Count; i++)
            {
                string rr = TPSMarketingEndCountDT.Rows[i]["InvNumber"].ToString();
                DataRow[] ItemsRows = AllTPSDT.Select("InvNumber = '" + rr + "'");
                if (ItemsRows.Count() > 0)
                {
                    foreach (DataRow item in ItemsRows)
                    {
                        if (item["EndCount"] == DBNull.Value)
                            item["EndCount"] = TPSMarketingEndCountDT.Rows[i]["Count"];
                        else
                            item["EndCount"] = Convert.ToDecimal(item["EndCount"]) + Convert.ToDecimal(TPSMarketingEndCountDT.Rows[i]["Count"]);
                    }
                }
                else
                {
                    DataRow NewRow = AllTPSDT.NewRow();
                    NewRow["AccountingName"] = TPSMarketingEndCountDT.Rows[i]["AccountingName"];
                    NewRow["Measure"] = TPSMarketingEndCountDT.Rows[i]["Measure"];
                    NewRow["InvNumber"] = TPSMarketingEndCountDT.Rows[i]["InvNumber"];
                    NewRow["EndCount"] = TPSMarketingEndCountDT.Rows[i]["Count"];
                    AllTPSDT.Rows.Add(NewRow);
                }
            }
            for (int i = 0; i < TPSMarketingProducedDT.Rows.Count; i++)
            {
                string rr = TPSMarketingProducedDT.Rows[i]["InvNumber"].ToString();
                DataRow[] ItemsRows = AllTPSDT.Select("InvNumber = '" + rr + "'");
                if (ItemsRows.Count() > 0)
                {
                    foreach (DataRow item in ItemsRows)
                    {
                        if (item["ProducedCount"] == DBNull.Value)
                            item["ProducedCount"] = TPSMarketingProducedDT.Rows[i]["Count"];
                        else
                            item["ProducedCount"] = Convert.ToDecimal(item["ProducedCount"]) + Convert.ToDecimal(TPSMarketingProducedDT.Rows[i]["Count"]);
                    }
                }
                else
                {
                    DataRow NewRow = AllTPSDT.NewRow();
                    NewRow["AccountingName"] = TPSMarketingProducedDT.Rows[i]["AccountingName"];
                    NewRow["Measure"] = TPSMarketingProducedDT.Rows[i]["Measure"];
                    NewRow["InvNumber"] = TPSMarketingProducedDT.Rows[i]["InvNumber"];
                    NewRow["ProducedCount"] = TPSMarketingProducedDT.Rows[i]["Count"];
                    AllTPSDT.Rows.Add(NewRow);
                }
            }
            for (int i = 0; i < TPSMarketingDispatchDT.Rows.Count; i++)
            {
                string rr = TPSMarketingDispatchDT.Rows[i]["InvNumber"].ToString();
                DataRow[] ItemsRows = AllTPSDT.Select("InvNumber = '" + rr + "'");
                if (ItemsRows.Count() > 0)
                {
                    foreach (DataRow item in ItemsRows)
                    {
                        if (item["DispatchCount"] == DBNull.Value)
                            item["DispatchCount"] = TPSMarketingDispatchDT.Rows[i]["Count"];
                        else
                            item["DispatchCount"] = Convert.ToDecimal(item["DispatchCount"]) + Convert.ToDecimal(TPSMarketingDispatchDT.Rows[i]["Count"]);
                    }
                }
                else
                {
                    DataRow NewRow = AllTPSDT.NewRow();
                    NewRow["AccountingName"] = TPSMarketingDispatchDT.Rows[i]["AccountingName"];
                    NewRow["Measure"] = TPSMarketingDispatchDT.Rows[i]["Measure"];
                    NewRow["InvNumber"] = TPSMarketingDispatchDT.Rows[i]["InvNumber"];
                    NewRow["DispatchCount"] = TPSMarketingDispatchDT.Rows[i]["Count"];
                    AllTPSDT.Rows.Add(NewRow);
                }
            }

            for (int i = 0; i < TPSZOVStartCountDT.Rows.Count; i++)
            {
                string rr = TPSZOVStartCountDT.Rows[i]["InvNumber"].ToString();
                DataRow[] ItemsRows = AllTPSDT.Select("InvNumber = '" + rr + "'");
                if (ItemsRows.Count() > 0)
                {
                    foreach (DataRow item in ItemsRows)
                    {
                        if (item["StartCount"] == DBNull.Value)
                            item["StartCount"] = TPSZOVStartCountDT.Rows[i]["Count"];
                        else
                            item["StartCount"] = Convert.ToDecimal(item["StartCount"]) + Convert.ToDecimal(TPSZOVStartCountDT.Rows[i]["Count"]);
                    }
                }
                else
                {
                    DataRow NewRow = AllTPSDT.NewRow();
                    NewRow["AccountingName"] = TPSZOVStartCountDT.Rows[i]["AccountingName"];
                    NewRow["Measure"] = ProfilMarketingStartCountDT.Rows[i]["Measure"];
                    NewRow["InvNumber"] = TPSZOVStartCountDT.Rows[i]["InvNumber"];
                    NewRow["StartCount"] = TPSZOVStartCountDT.Rows[i]["Count"];
                    AllTPSDT.Rows.Add(NewRow);
                }
            }
            for (int i = 0; i < TPSZOVEndCountDT.Rows.Count; i++)
            {
                string rr = TPSZOVEndCountDT.Rows[i]["InvNumber"].ToString();
                DataRow[] ItemsRows = AllTPSDT.Select("InvNumber = '" + rr + "'");
                if (ItemsRows.Count() > 0)
                {
                    foreach (DataRow item in ItemsRows)
                    {
                        if (item["EndCount"] == DBNull.Value)
                            item["EndCount"] = TPSZOVEndCountDT.Rows[i]["Count"];
                        else
                            item["EndCount"] = Convert.ToDecimal(item["EndCount"]) + Convert.ToDecimal(TPSZOVEndCountDT.Rows[i]["Count"]);
                    }
                }
                else
                {
                    DataRow NewRow = AllTPSDT.NewRow();
                    NewRow["AccountingName"] = TPSZOVEndCountDT.Rows[i]["AccountingName"];
                    NewRow["Measure"] = TPSZOVEndCountDT.Rows[i]["Measure"];
                    NewRow["InvNumber"] = TPSZOVEndCountDT.Rows[i]["InvNumber"];
                    NewRow["EndCount"] = TPSZOVEndCountDT.Rows[i]["Count"];
                    AllTPSDT.Rows.Add(NewRow);
                }
            }
            for (int i = 0; i < TPSZOVProducedDT.Rows.Count; i++)
            {
                string rr = TPSZOVProducedDT.Rows[i]["InvNumber"].ToString();
                DataRow[] ItemsRows = AllTPSDT.Select("InvNumber = '" + rr + "'");
                if (ItemsRows.Count() > 0)
                {
                    foreach (DataRow item in ItemsRows)
                    {
                        if (item["ProducedCount"] == DBNull.Value)
                            item["ProducedCount"] = TPSZOVProducedDT.Rows[i]["Count"];
                        else
                            item["ProducedCount"] = Convert.ToDecimal(item["ProducedCount"]) + Convert.ToDecimal(TPSZOVProducedDT.Rows[i]["Count"]);
                    }
                }
                else
                {
                    DataRow NewRow = AllTPSDT.NewRow();
                    NewRow["AccountingName"] = TPSZOVProducedDT.Rows[i]["AccountingName"];
                    NewRow["Measure"] = TPSZOVProducedDT.Rows[i]["Measure"];
                    NewRow["InvNumber"] = TPSZOVProducedDT.Rows[i]["InvNumber"];
                    NewRow["ProducedCount"] = TPSZOVProducedDT.Rows[i]["Count"];
                    AllTPSDT.Rows.Add(NewRow);
                }
            }
            for (int i = 0; i < TPSZOVDispatchDT.Rows.Count; i++)
            {
                string rr = TPSZOVDispatchDT.Rows[i]["InvNumber"].ToString();
                DataRow[] ItemsRows = AllTPSDT.Select("InvNumber = '" + rr + "'");
                if (ItemsRows.Count() > 0)
                {
                    foreach (DataRow item in ItemsRows)
                    {
                        if (item["DispatchCount"] == DBNull.Value)
                            item["DispatchCount"] = TPSZOVDispatchDT.Rows[i]["Count"];
                        else
                            item["DispatchCount"] = Convert.ToDecimal(item["DispatchCount"]) + Convert.ToDecimal(TPSZOVDispatchDT.Rows[i]["Count"]);
                    }
                }
                else
                {
                    DataRow NewRow = AllTPSDT.NewRow();
                    NewRow["AccountingName"] = TPSZOVDispatchDT.Rows[i]["AccountingName"];
                    NewRow["Measure"] = TPSZOVDispatchDT.Rows[i]["Measure"];
                    NewRow["InvNumber"] = TPSZOVDispatchDT.Rows[i]["InvNumber"];
                    NewRow["DispatchCount"] = TPSZOVDispatchDT.Rows[i]["Count"];
                    AllTPSDT.Rows.Add(NewRow);
                }
            }

            using (DataView DV = new DataView(AllTPSDT.Copy()))
            {
                DV.Sort = "AccountingName";
                AllTPSDT.Clear();
                AllTPSDT = DV.ToTable();
            }
            foreach (DataRow item in AllTPSDT.Rows)
            {
                if (item["StartCount"] == DBNull.Value)
                    item["StartCount"] = 0;
                if (item["EndCount"] == DBNull.Value)
                    item["EndCount"] = 0;
                if (item["ProducedCount"] == DBNull.Value)
                    item["ProducedCount"] = 0;
                if (item["DispatchCount"] == DBNull.Value)
                    item["DispatchCount"] = 0;
            }
        }
    }

    public class ProducedDecorReport
    {
        //decimal TransportCost = 0;
        //decimal AdditionalCost = 0;
        //decimal Rate = 1;
        //int ClientID = 0;
        private readonly string ProfilCurrencyCode = "0";

        private readonly string TPSCurrencyCode = "0";
        private readonly string UNN = string.Empty;

        private readonly Infinium.Modules.ZOV.DecorCatalogOrder DecorCatalogOrder = null;

        private DataTable CurrencyTypesDT;
        public DataTable ProfilReportDataTable = null;
        public DataTable TPSReportDataTable = null;
        private DataTable DecorProductsDataTable = null;
        private DataTable DecorConfigDataTable = null;
        private DataTable MeasuresDataTable = null;
        private DataTable DecorOrdersDataTable = null;
        private DataTable DecorDataTable = null;

        private DataTable FrameColorsDataTable = null;
        private DataTable PatinaDataTable = null;
        private DataTable PatinaRALDataTable = null;

        private DataTable ProfilMarketingStartCountDT = null;
        private DataTable ProfilMarketingEndCountDT = null;
        private DataTable ProfilMarketingProducedDT = null;
        private DataTable ProfilMarketingDispatchDT = null;
        private DataTable ProfilZOVStartCountDT = null;
        private DataTable ProfilZOVEndCountDT = null;
        private DataTable ProfilZOVProducedDT = null;
        private DataTable ProfilZOVDispatchDT = null;
        public DataTable AllProfilDT = null;

        private DataTable TPSMarketingStartCountDT = null;
        private DataTable TPSMarketingEndCountDT = null;
        private DataTable TPSMarketingProducedDT = null;
        private DataTable TPSMarketingDispatchDT = null;
        private DataTable TPSZOVStartCountDT = null;
        private DataTable TPSZOVEndCountDT = null;
        private DataTable TPSZOVProducedDT = null;
        private DataTable TPSZOVDispatchDT = null;
        public DataTable AllTPSDT = null;

        public ProducedDecorReport(ref Infinium.Modules.ZOV.DecorCatalogOrder tDecorCatalogOrder)
        {
            DecorCatalogOrder = tDecorCatalogOrder;

            Create();
            CreateReportDataTable();
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

            DecorConfigDataTable = new DataTable();

            GetColorsDT();
            GetPatinaDT();
            DecorConfigDataTable = TablesManager.DecorConfigDataTableAll;
        }

        private string GetColorCode(int ColorID)
        {
            string code = "";
            try
            {
                DataRow[] Rows = FrameColorsDataTable.Select("ColorID = " + ColorID);
                code = Rows[0]["Cvet"].ToString();
            }
            catch
            {
                return "";
            }
            return code;
        }

        private void GetColorsDT()
        {
            FrameColorsDataTable = new DataTable();
            FrameColorsDataTable.Columns.Add(new DataColumn("ColorID", Type.GetType("System.Int64")));
            FrameColorsDataTable.Columns.Add(new DataColumn("ColorName", Type.GetType("System.String")));
            FrameColorsDataTable.Columns.Add(new DataColumn("Cvet", Type.GetType("System.String")));
            string SelectCommand = @"SELECT TechStoreID, TechStoreName, Cvet FROM TechStore
                WHERE TechStoreSubGroupID IN (SELECT TechStoreSubGroupID FROM TechStoreSubGroups WHERE (TechStoreGroupID = 11 OR TechStoreGroupID = 1))
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

        private void CreateReportDataTable()
        {
            ProfilReportDataTable = new DataTable();
            ProfilReportDataTable.Columns.Add(new DataColumn("UNN", Type.GetType("System.String")));
            ProfilReportDataTable.Columns.Add(new DataColumn("CurrencyCode", Type.GetType("System.String")));
            ProfilReportDataTable.Columns.Add(new DataColumn("TPSCurCode", Type.GetType("System.String")));
            ProfilReportDataTable.Columns.Add(new DataColumn("Measure", Type.GetType("System.String")));
            ProfilReportDataTable.Columns.Add(new DataColumn("InvNumber", Type.GetType("System.String")));
            ProfilReportDataTable.Columns.Add(new DataColumn("Cvet", Type.GetType("System.String")));
            ProfilReportDataTable.Columns.Add(new DataColumn("Patina", Type.GetType("System.String")));
            ProfilReportDataTable.Columns.Add(new DataColumn("AccountingName", Type.GetType("System.String")));
            ProfilReportDataTable.Columns.Add(new DataColumn("Count", Type.GetType("System.Decimal")));
            ProfilReportDataTable.Columns.Add(new DataColumn("Price", Type.GetType("System.String")));
            ProfilReportDataTable.Columns.Add(new DataColumn("Cost", Type.GetType("System.String")));
            ProfilReportDataTable.Columns.Add(new DataColumn("Weight", Type.GetType("System.String")));
            ProfilReportDataTable.Columns.Add(new DataColumn("MarketingCost", Type.GetType("System.String")));

            AllProfilDT = new DataTable();
            AllProfilDT.Columns.Add(new DataColumn("AccountingName", Type.GetType("System.String")));
            AllProfilDT.Columns.Add(new DataColumn("InvNumber", Type.GetType("System.String")));
            AllProfilDT.Columns.Add(new DataColumn("Measure", Type.GetType("System.String")));
            AllProfilDT.Columns.Add(new DataColumn("StartCount", Type.GetType("System.Decimal")));
            AllProfilDT.Columns.Add(new DataColumn("ProducedCount", Type.GetType("System.Decimal")));
            AllProfilDT.Columns.Add(new DataColumn("DispatchCount", Type.GetType("System.Decimal")));
            AllProfilDT.Columns.Add(new DataColumn("EndCount", Type.GetType("System.Decimal")));
            AllTPSDT = AllProfilDT.Clone();

            TPSReportDataTable = ProfilReportDataTable.Clone();

            ProfilMarketingStartCountDT = ProfilReportDataTable.Clone();
            ProfilMarketingEndCountDT = ProfilReportDataTable.Clone();
            ProfilMarketingProducedDT = ProfilReportDataTable.Clone();
            ProfilMarketingDispatchDT = ProfilReportDataTable.Clone();
            ProfilZOVStartCountDT = ProfilReportDataTable.Clone();
            ProfilZOVEndCountDT = ProfilReportDataTable.Clone();
            ProfilZOVProducedDT = ProfilReportDataTable.Clone();
            ProfilZOVDispatchDT = ProfilReportDataTable.Clone();

            TPSMarketingStartCountDT = ProfilReportDataTable.Clone();
            TPSMarketingEndCountDT = ProfilReportDataTable.Clone();
            TPSMarketingProducedDT = ProfilReportDataTable.Clone();
            TPSMarketingDispatchDT = ProfilReportDataTable.Clone();
            TPSZOVStartCountDT = ProfilReportDataTable.Clone();
            TPSZOVEndCountDT = ProfilReportDataTable.Clone();
            TPSZOVProducedDT = ProfilReportDataTable.Clone();
            TPSZOVDispatchDT = ProfilReportDataTable.Clone();
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

        private decimal GetMarketingCost(int DecorConfigID)
        {
            DataRow[] DRows = DecorConfigDataTable.Select("DecorConfigID = " + DecorConfigID);
            if (DRows.Count() == 0)
                return 0;
            return Convert.ToDecimal(DRows[0]["MarketingCost"]);
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
            DT.Columns.Add(new DataColumn("Measure", Type.GetType("System.String")));
            DT.Columns.Add(new DataColumn("InvNumber", Type.GetType("System.String")));
            DT.Columns.Add(new DataColumn("AccountingName", Type.GetType("System.String")));
            DT.Columns.Add(new DataColumn("Cvet", Type.GetType("System.String")));
            DT.Columns.Add(new DataColumn("Patina", Type.GetType("System.String")));
            DT.Columns.Add(new DataColumn("Count", Type.GetType("System.Decimal")));
            DT.Columns.Add(new DataColumn("TotalCount", Type.GetType("System.Int32")));
            DT.Columns.Add(new DataColumn("Cost", Type.GetType("System.Decimal")));
            DT.Columns.Add(new DataColumn("Weight", Type.GetType("System.Decimal")));
            DT.Columns.Add(new DataColumn("MarketingCost", Type.GetType("System.Decimal")));
        }

        private string GetPatinaCode(int PatinaID)
        {
            string code = "";
            try
            {
                DataRow[] Rows = PatinaDataTable.Select("PatinaID = " + PatinaID);
                code = Rows[0]["Patina"].ToString();
            }
            catch
            {
                return "";
            }
            return code;
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
            PDT.Columns.Add(new DataColumn("MarketingCost", Type.GetType("System.Decimal")));
            PDT.Columns.Add(new DataColumn("Cvet", Type.GetType("System.String")));
            PDT.Columns.Add(new DataColumn("Patina", Type.GetType("System.String")));

            TDT.Columns.Remove("Count");
            TDT.Columns.Add(new DataColumn("Count", Type.GetType("System.Decimal")));
            TDT.Columns.Add(new DataColumn("UNN", Type.GetType("System.String")));
            TDT.Columns.Add(new DataColumn("CurrencyCode", Type.GetType("System.String")));
            TDT.Columns.Add(new DataColumn("TPSCurCode", Type.GetType("System.String")));
            TDT.Columns.Add(new DataColumn("MarketingCost", Type.GetType("System.Decimal")));
            TDT.Columns.Add(new DataColumn("Cvet", Type.GetType("System.String")));
            TDT.Columns.Add(new DataColumn("Patina", Type.GetType("System.String")));
            //013011
            decimal MarketingCost = 0;
            if (Rows.Count() > 0)
                MarketingCost = GetMarketingCost(Convert.ToInt32(Rows[0]["DecorConfigID"]));
            for (int r = 0; r < Rows.Count(); r++)
            {
                int ColorID = Convert.ToInt32(Rows[r]["ColorID"]);
                int PatinaID = Convert.ToInt32(Rows[r]["PatinaID"]);
                string InvNumber = Rows[r]["InvNumber"].ToString();
                if (MarketingCost > GetMarketingCost(Convert.ToInt32(Rows[r]["DecorConfigID"])))
                    MarketingCost = GetMarketingCost(Convert.ToInt32(Rows[r]["DecorConfigID"]));
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
                                                       "' AND ColorID = " + Rows[r]["ColorID"].ToString() +
                                                       " AND PatinaID = " + Rows[r]["PatinaID"].ToString());
                        if (InvRows.Count() == 0)
                        {
                            DataRow NewRow = PDT.NewRow();
                            NewRow["UNN"] = UNN;
                            NewRow["AccountingName"] = Rows[r]["AccountingName"].ToString();
                            NewRow["InvNumber"] = Rows[r]["InvNumber"].ToString();
                            NewRow["Cvet"] = GetColorCode(ColorID);
                            NewRow["Patina"] = GetPatinaCode(PatinaID);
                            NewRow["MarketingCost"] = MarketingCost;
                            NewRow["CurrencyCode"] = ProfilCurrencyCode;
                            NewRow["Measure"] = "м.п.";
                            NewRow["Count"] = Convert.ToDecimal(Rows[r]["Count"]) * L;
                            //NewRow["Price"] = Convert.ToDecimal(Rows[r]["Price"]);
                            NewRow["Cost"] = Convert.ToDecimal(Rows[r]["Cost"]);
                            NewRow["Weight"] = GetDecorWeight(Rows[r]);
                            PDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            InvRows[0]["Count"] = Convert.ToDecimal(InvRows[0]["Count"]) +
                                Convert.ToDecimal(Rows[r]["Count"]) * L;
                            //InvRows[0]["Price"] = Convert.ToDecimal(Rows[r]["Price"]);
                            InvRows[0]["Cost"] = Convert.ToDecimal(InvRows[0]["Cost"]) +
                                Convert.ToDecimal(Rows[r]["Cost"]);
                            InvRows[0]["Weight"] = Convert.ToDecimal(InvRows[0]["Weight"]) +
                                GetDecorWeight(Rows[r]);
                        }
                    }
                    else
                    {
                        DataRow[] InvRows = TDT.Select("InvNumber = '" + Rows[r]["InvNumber"].ToString() +
                                                       "' AND ColorID = " + Rows[r]["ColorID"].ToString() +
                                                       " AND PatinaID = " + Rows[r]["PatinaID"].ToString());
                        if (InvRows.Count() == 0)
                        {
                            DataRow NewRow = TDT.NewRow();
                            NewRow["UNN"] = UNN;
                            NewRow["AccountingName"] = Rows[r]["AccountingName"].ToString();
                            NewRow["InvNumber"] = Rows[r]["InvNumber"].ToString();
                            NewRow["Cvet"] = GetColorCode(ColorID);
                            NewRow["Patina"] = GetPatinaCode(PatinaID);
                            NewRow["MarketingCost"] = MarketingCost;
                            NewRow["CurrencyCode"] = ProfilCurrencyCode;
                            NewRow["TPSCurCode"] = TPSCurrencyCode;
                            NewRow["Measure"] = "м.п.";
                            NewRow["Count"] = Convert.ToDecimal(Rows[r]["Count"]) * L;
                            //NewRow["Price"] = Convert.ToDecimal(Rows[r]["Price"]);
                            NewRow["Cost"] = Convert.ToDecimal(Rows[r]["Cost"]);
                            NewRow["Weight"] = GetDecorWeight(Rows[r]);
                            TDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            InvRows[0]["Count"] = Convert.ToDecimal(InvRows[0]["Count"]) +
                                Convert.ToDecimal(Rows[r]["Count"]) * L;
                            //InvRows[0]["Price"] = Convert.ToDecimal(Rows[r]["Price"]);
                            InvRows[0]["Cost"] = Convert.ToDecimal(InvRows[0]["Cost"]) +
                                Convert.ToDecimal(Rows[r]["Cost"]);
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
                                                       "' AND ColorID = " + Rows[r]["ColorID"].ToString() +
                                                       " AND PatinaID = " + Rows[r]["PatinaID"].ToString());
                        if (InvRows.Count() == 0)
                        {
                            DataRow NewRow = PDT.NewRow();
                            NewRow["UNN"] = UNN;
                            NewRow["AccountingName"] = Rows[r]["AccountingName"].ToString();
                            NewRow["InvNumber"] = Rows[r]["InvNumber"].ToString();
                            NewRow["Cvet"] = GetColorCode(ColorID);
                            NewRow["Patina"] = GetPatinaCode(PatinaID);
                            NewRow["MarketingCost"] = MarketingCost;
                            NewRow["CurrencyCode"] = ProfilCurrencyCode;
                            NewRow["Measure"] = "м.кв.";
                            if (GetMeasureTypeID(Convert.ToInt32(Rows[r]["DecorConfigID"])) == 1)
                            {
                                if (Convert.ToDecimal(Rows[r]["Height"]) != -1)
                                    NewRow["Count"] = Decimal.Round(Convert.ToDecimal(Rows[r]["Height"]) * Convert.ToDecimal(Rows[r]["Width"]) / 1000000, 3, MidpointRounding.AwayFromZero)
                                    * Convert.ToDecimal(Rows[r]["Count"]);
                                if (Convert.ToDecimal(Rows[r]["Length"]) != -1)
                                    NewRow["Count"] = Decimal.Round(Convert.ToDecimal(Rows[r]["Length"]) * Convert.ToDecimal(Rows[r]["Width"]) / 1000000, 3, MidpointRounding.AwayFromZero)
                                    * Convert.ToDecimal(Rows[r]["Count"]);
                            }

                            if (GetMeasureTypeID(Convert.ToInt32(Rows[r]["DecorConfigID"])) == 2)
                            {
                                decimal L = 0;
                                L = Convert.ToDecimal(Rows[r]["Length"]);
                                if (L == -1)
                                    L = Convert.ToDecimal(Rows[r]["Height"]);
                                NewRow["Count"] = Convert.ToDecimal(Rows[r]["Count"]) *
                                    Decimal.Round(L * Convert.ToDecimal(Rows[r]["Width"]) / 1000000, 3, MidpointRounding.AwayFromZero);
                            }

                            if (GetMeasureTypeID(Convert.ToInt32(Rows[r]["DecorConfigID"])) == 3)
                            {
                                decimal H = 0;
                                if (DecorCatalogOrder.HasParameter(Convert.ToInt32(Rows[r]["ProductID"]), "Height") && Convert.ToDecimal(Rows[r]["Height"]) != -1)
                                    H = Convert.ToDecimal(Rows[r]["Height"]);
                                if (DecorCatalogOrder.HasParameter(Convert.ToInt32(Rows[r]["ProductID"]), "Length") && Convert.ToDecimal(Rows[r]["Length"]) != -1)
                                    H = Convert.ToDecimal(Rows[r]["Length"]);
                                NewRow["Count"] = Convert.ToDecimal(Rows[r]["Count"]) *
                                    Decimal.Round(H * Convert.ToDecimal(Rows[r]["Width"]) / 1000000, 3, MidpointRounding.AwayFromZero);
                            }
                            //NewRow["Price"] = Convert.ToDecimal(Rows[r]["Price"]);
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
                                    InvRows[0]["Count"] = Convert.ToDecimal(InvRows[0]["Count"]) +
                                        Decimal.Round(Convert.ToDecimal(Rows[r]["Height"]) * Convert.ToDecimal(Rows[r]["Width"]) / 1000000, 3, MidpointRounding.AwayFromZero)
                                    * Convert.ToDecimal(Rows[r]["Count"]);
                                if (Convert.ToDecimal(Rows[r]["Length"]) != -1)
                                    InvRows[0]["Count"] = Convert.ToDecimal(InvRows[0]["Count"]) +
                                        Decimal.Round(Convert.ToDecimal(Rows[r]["Length"]) * Convert.ToDecimal(Rows[r]["Width"]) / 1000000, 3, MidpointRounding.AwayFromZero)
                                    * Convert.ToDecimal(Rows[r]["Count"]);
                            }
                            if (GetMeasureTypeID(Convert.ToInt32(Rows[0]["DecorConfigID"])) == 2)
                            {
                                decimal L = 0;
                                L = Convert.ToDecimal(Rows[r]["Length"]);
                                if (L == -1)
                                    L = Convert.ToDecimal(Rows[r]["Height"]);
                                InvRows[0]["Count"] = Convert.ToDecimal(InvRows[0]["Count"]) +
                                    Convert.ToDecimal(Rows[r]["Count"]) *
                                    Decimal.Round(L * Convert.ToDecimal(Rows[r]["Width"]) / 1000000, 3, MidpointRounding.AwayFromZero);
                            }
                            if (GetMeasureTypeID(Convert.ToInt32(Rows[r]["DecorConfigID"])) == 3)
                            {
                                decimal H = 0;
                                if (DecorCatalogOrder.HasParameter(Convert.ToInt32(Rows[r]["ProductID"]), "Height") && Convert.ToDecimal(Rows[r]["Height"]) != -1)
                                    H = Convert.ToDecimal(Rows[r]["Height"]);
                                if (DecorCatalogOrder.HasParameter(Convert.ToInt32(Rows[r]["ProductID"]), "Length") && Convert.ToDecimal(Rows[r]["Length"]) != -1)
                                    H = Convert.ToDecimal(Rows[r]["Length"]);
                                InvRows[0]["Count"] = Convert.ToDecimal(InvRows[0]["Count"]) +
                                    Convert.ToDecimal(Rows[r]["Count"]) *
                                    Decimal.Round(H * Convert.ToDecimal(Rows[r]["Width"]) / 1000000, 3, MidpointRounding.AwayFromZero);
                            }
                            //InvRows[0]["Price"] = Convert.ToDecimal(Rows[r]["Price"]);
                            InvRows[0]["Cost"] = Convert.ToDecimal(InvRows[0]["Cost"]) +
                                Convert.ToDecimal(Rows[r]["Cost"]);
                            InvRows[0]["Weight"] = Convert.ToDecimal(InvRows[0]["Weight"]) +
                                GetDecorWeight(Rows[r]);
                        }
                    }
                    else
                    {
                        DataRow[] InvRows = TDT.Select("InvNumber = '" + Rows[r]["InvNumber"].ToString() +
                                                       "' AND ColorID = " + Rows[r]["ColorID"].ToString() +
                                                       " AND PatinaID = " + Rows[r]["PatinaID"].ToString());
                        if (InvRows.Count() == 0)
                        {
                            DataRow NewRow = TDT.NewRow();
                            NewRow["UNN"] = UNN;
                            NewRow["AccountingName"] = Rows[r]["AccountingName"].ToString();
                            NewRow["InvNumber"] = Rows[r]["InvNumber"].ToString();
                            NewRow["MarketingCost"] = MarketingCost;
                            NewRow["Cvet"] = GetColorCode(ColorID);
                            NewRow["Patina"] = GetPatinaCode(PatinaID);
                            NewRow["CurrencyCode"] = ProfilCurrencyCode;
                            NewRow["TPSCurCode"] = TPSCurrencyCode;
                            NewRow["Measure"] = "м.кв.";
                            if (GetMeasureTypeID(Convert.ToInt32(Rows[r]["DecorConfigID"])) == 1)
                            {
                                if (Convert.ToDecimal(Rows[r]["Height"]) != -1)
                                    NewRow["Count"] = Decimal.Round(Convert.ToDecimal(Rows[r]["Height"]) * Convert.ToDecimal(Rows[r]["Width"]) / 1000000, 3, MidpointRounding.AwayFromZero)
                                    * Convert.ToDecimal(Rows[r]["Count"]);
                                if (Convert.ToDecimal(Rows[r]["Length"]) != -1)
                                    NewRow["Count"] = Decimal.Round(Convert.ToDecimal(Rows[r]["Length"]) * Convert.ToDecimal(Rows[r]["Width"]) / 1000000, 3, MidpointRounding.AwayFromZero)
                                    * Convert.ToDecimal(Rows[r]["Count"]);
                            }

                            if (GetMeasureTypeID(Convert.ToInt32(Rows[0]["DecorConfigID"])) == 2)
                            {
                                decimal L = 0;
                                L = Convert.ToDecimal(Rows[r]["Length"]);
                                if (L == -1)
                                    L = Convert.ToDecimal(Rows[r]["Height"]);
                                NewRow["Count"] = Convert.ToDecimal(Rows[r]["Count"]) *
                                    Decimal.Round(L * Convert.ToDecimal(Rows[r]["Width"]) / 1000000, 3, MidpointRounding.AwayFromZero);
                            }

                            if (GetMeasureTypeID(Convert.ToInt32(Rows[r]["DecorConfigID"])) == 3)
                            {
                                decimal H = 0;
                                if (DecorCatalogOrder.HasParameter(Convert.ToInt32(Rows[r]["ProductID"]), "Height") && Convert.ToDecimal(Rows[r]["Height"]) != -1)
                                    H = Convert.ToDecimal(Rows[r]["Height"]);
                                if (DecorCatalogOrder.HasParameter(Convert.ToInt32(Rows[r]["ProductID"]), "Length") && Convert.ToDecimal(Rows[r]["Length"]) != -1)
                                    H = Convert.ToDecimal(Rows[r]["Length"]);
                                NewRow["Count"] = Convert.ToDecimal(Rows[r]["Count"]) *
                                    Decimal.Round(H * Convert.ToDecimal(Rows[r]["Width"]) / 1000000, 3, MidpointRounding.AwayFromZero);
                            }
                            //NewRow["Price"] = Convert.ToDecimal(Rows[r]["Price"]);
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
                                    InvRows[0]["Count"] = Convert.ToDecimal(InvRows[0]["Count"]) +
                                        Decimal.Round(Convert.ToDecimal(Rows[r]["Height"]) * Convert.ToDecimal(Rows[r]["Width"]) / 1000000, 3, MidpointRounding.AwayFromZero)
                                    * Convert.ToDecimal(Rows[r]["Count"]);
                                if (Convert.ToDecimal(Rows[r]["Length"]) != -1)
                                    InvRows[0]["Count"] = Convert.ToDecimal(InvRows[0]["Count"]) +
                                        Decimal.Round(Convert.ToDecimal(Rows[r]["Length"]) * Convert.ToDecimal(Rows[r]["Width"]) / 1000000, 3, MidpointRounding.AwayFromZero)
                                    * Convert.ToDecimal(Rows[r]["Count"]);
                            }
                            if (GetMeasureTypeID(Convert.ToInt32(Rows[r]["DecorConfigID"])) == 2)
                            {
                                decimal L = 0;
                                L = Convert.ToDecimal(Rows[r]["Length"]);
                                if (L == -1)
                                    L = Convert.ToDecimal(Rows[r]["Height"]);
                                InvRows[0]["Count"] = Convert.ToDecimal(InvRows[0]["Count"]) +
                                    Convert.ToDecimal(Rows[r]["Count"]) *
                                    Decimal.Round(L * Convert.ToDecimal(Rows[r]["Width"]) / 1000000, 3, MidpointRounding.AwayFromZero);
                            }
                            if (GetMeasureTypeID(Convert.ToInt32(Rows[r]["DecorConfigID"])) == 3)
                            {
                                decimal H = 0;
                                if (DecorCatalogOrder.HasParameter(Convert.ToInt32(Rows[r]["ProductID"]), "Height") && Convert.ToDecimal(Rows[r]["Height"]) != -1)
                                    H = Convert.ToDecimal(Rows[r]["Height"]);
                                if (DecorCatalogOrder.HasParameter(Convert.ToInt32(Rows[r]["ProductID"]), "Length") && Convert.ToDecimal(Rows[r]["Length"]) != -1)
                                    H = Convert.ToDecimal(Rows[r]["Length"]);
                                InvRows[0]["Count"] = Convert.ToDecimal(InvRows[0]["Count"]) +
                                    Convert.ToDecimal(Rows[r]["Count"]) *
                                    Decimal.Round(H * Convert.ToDecimal(Rows[r]["Width"]) / 1000000, 3, MidpointRounding.AwayFromZero);
                            }
                            //InvRows[0]["Price"] = Convert.ToDecimal(Rows[r]["Price"]);
                            InvRows[0]["Cost"] = Convert.ToDecimal(InvRows[0]["Cost"]) +
                                Convert.ToDecimal(Rows[r]["Cost"]);
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
                        NewRow["AccountingName"] = PDT.Rows[i]["AccountingName"].ToString();
                        NewRow["InvNumber"] = PDT.Rows[i]["InvNumber"].ToString();
                        NewRow["Cvet"] = PDT.Rows[i]["Cvet"].ToString();
                        NewRow["Patina"] = PDT.Rows[i]["Patina"].ToString();
                        NewRow["MarketingCost"] = Convert.ToDecimal(PDT.Rows[i]["MarketingCost"]);
                        NewRow["CurrencyCode"] = ProfilCurrencyCode;
                        NewRow["Measure"] = "м.п.";
                        NewRow["Count"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[i]["Count"]) / 1000, 3, MidpointRounding.AwayFromZero);
                        //NewRow["Price"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[i]["Cost"]) / (Convert.ToDecimal(PDT.Rows[i]["Count"]) / 1000), 3, MidpointRounding.AwayFromZero);
                        NewRow["Cost"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[i]["Cost"]), 3, MidpointRounding.AwayFromZero);
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
                        NewRow["AccountingName"] = TDT.Rows[i]["AccountingName"].ToString();
                        NewRow["InvNumber"] = TDT.Rows[i]["InvNumber"].ToString();
                        NewRow["MarketingCost"] = Convert.ToDecimal(TDT.Rows[i]["MarketingCost"]);
                        NewRow["Cvet"] = TDT.Rows[i]["Cvet"].ToString();
                        NewRow["Patina"] = TDT.Rows[i]["Patina"].ToString();
                        NewRow["CurrencyCode"] = ProfilCurrencyCode;
                        NewRow["TPSCurCode"] = TPSCurrencyCode;
                        NewRow["Measure"] = "м.п.";
                        NewRow["Count"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[i]["Count"]) / 1000, 3, MidpointRounding.AwayFromZero);
                        //NewRow["Price"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[i]["Cost"]) / (Convert.ToDecimal(TDT.Rows[i]["Count"]) / 1000), 3, MidpointRounding.AwayFromZero);
                        NewRow["Cost"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[i]["Cost"]), 3, MidpointRounding.AwayFromZero);
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
                        NewRow["AccountingName"] = PDT.Rows[i]["AccountingName"].ToString();
                        NewRow["InvNumber"] = PDT.Rows[i]["InvNumber"].ToString();
                        NewRow["MarketingCost"] = Convert.ToDecimal(PDT.Rows[i]["MarketingCost"]);
                        NewRow["Cvet"] = PDT.Rows[i]["Cvet"].ToString();
                        NewRow["Patina"] = PDT.Rows[i]["Patina"].ToString();
                        NewRow["CurrencyCode"] = ProfilCurrencyCode;
                        NewRow["Measure"] = "м.кв.";
                        NewRow["Count"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[i]["Count"]), 3, MidpointRounding.AwayFromZero);
                        //NewRow["Price"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[i]["Cost"]) / Convert.ToDecimal(PDT.Rows[i]["Count"]), 3, MidpointRounding.AwayFromZero);
                        NewRow["Cost"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[i]["Cost"]), 3, MidpointRounding.AwayFromZero);
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
                        NewRow["AccountingName"] = TDT.Rows[i]["AccountingName"].ToString();
                        NewRow["InvNumber"] = TDT.Rows[i]["InvNumber"].ToString();
                        NewRow["MarketingCost"] = Convert.ToDecimal(TDT.Rows[i]["MarketingCost"]);
                        NewRow["Cvet"] = TDT.Rows[i]["Cvet"].ToString();
                        NewRow["Patina"] = TDT.Rows[i]["Patina"].ToString();
                        NewRow["CurrencyCode"] = ProfilCurrencyCode;
                        NewRow["TPSCurCode"] = TPSCurrencyCode;
                        NewRow["Measure"] = "м.кв.";
                        NewRow["Count"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[i]["Count"]), 3, MidpointRounding.AwayFromZero);
                        //NewRow["Price"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[i]["Cost"]) / Convert.ToDecimal(TDT.Rows[i]["Count"]), 3, MidpointRounding.AwayFromZero);
                        NewRow["Cost"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[i]["Cost"]), 3, MidpointRounding.AwayFromZero);
                        NewRow["Weight"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[i]["Weight"]), 3, MidpointRounding.AwayFromZero);

                        TPSReportDataTable.Rows.Add(NewRow);
                    }
                }
            }

            PDT.Dispose();
            TDT.Dispose();
        }

        private void GroupCoverTypes(DataRow[] Rows, int MeasureTypeID, DataTable PRepDT, DataTable TRepDT)
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
            PDT.Columns.Add(new DataColumn("MarketingCost", Type.GetType("System.Decimal")));
            PDT.Columns.Add(new DataColumn("Cvet", Type.GetType("System.String")));
            PDT.Columns.Add(new DataColumn("Patina", Type.GetType("System.String")));

            TDT.Columns.Remove("Count");
            TDT.Columns.Add(new DataColumn("Count", Type.GetType("System.Decimal")));
            TDT.Columns.Add(new DataColumn("UNN", Type.GetType("System.String")));
            TDT.Columns.Add(new DataColumn("CurrencyCode", Type.GetType("System.String")));
            TDT.Columns.Add(new DataColumn("TPSCurCode", Type.GetType("System.String")));
            TDT.Columns.Add(new DataColumn("MarketingCost", Type.GetType("System.Decimal")));
            TDT.Columns.Add(new DataColumn("Cvet", Type.GetType("System.String")));
            TDT.Columns.Add(new DataColumn("Patina", Type.GetType("System.String")));
            //013011
            decimal MarketingCost = 0;
            if (Rows.Count() > 0)
                MarketingCost = GetMarketingCost(Convert.ToInt32(Rows[0]["DecorConfigID"]));
            for (int r = 0; r < Rows.Count(); r++)
            {
                int ColorID = Convert.ToInt32(Rows[r]["ColorID"]);
                int PatinaID = Convert.ToInt32(Rows[r]["PatinaID"]);
                string InvNumber = Rows[r]["InvNumber"].ToString();
                if (MarketingCost > GetMarketingCost(Convert.ToInt32(Rows[r]["DecorConfigID"])))
                    MarketingCost = GetMarketingCost(Convert.ToInt32(Rows[r]["DecorConfigID"]));
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
                                                       "' AND ColorID = " + Rows[r]["ColorID"].ToString() +
                                                       " AND PatinaID = " + Rows[r]["PatinaID"].ToString());
                        if (InvRows.Count() == 0)
                        {
                            DataRow NewRow = PDT.NewRow();
                            NewRow["UNN"] = UNN;
                            NewRow["AccountingName"] = Rows[r]["AccountingName"].ToString();
                            NewRow["InvNumber"] = Rows[r]["InvNumber"].ToString();
                            NewRow["MarketingCost"] = MarketingCost;
                            NewRow["Cvet"] = GetColorCode(ColorID);
                            NewRow["Patina"] = GetPatinaCode(PatinaID);
                            NewRow["CurrencyCode"] = ProfilCurrencyCode;
                            NewRow["Measure"] = "м.п.";
                            NewRow["Count"] = Convert.ToDecimal(Rows[r]["Count"]) * L;
                            //NewRow["Price"] = Convert.ToDecimal(Rows[r]["Price"]);
                            NewRow["Cost"] = Convert.ToDecimal(Rows[r]["Cost"]);
                            NewRow["Weight"] = GetDecorWeight(Rows[r]);
                            PDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            InvRows[0]["Count"] = Convert.ToDecimal(InvRows[0]["Count"]) +
                                Convert.ToDecimal(Rows[r]["Count"]) * L;
                            //InvRows[0]["Price"] = Convert.ToDecimal(Rows[r]["Price"]);
                            InvRows[0]["Cost"] = Convert.ToDecimal(InvRows[0]["Cost"]) +
                                Convert.ToDecimal(Rows[r]["Cost"]);
                            InvRows[0]["Weight"] = Convert.ToDecimal(InvRows[0]["Weight"]) +
                                GetDecorWeight(Rows[r]);
                        }
                    }
                    else
                    {
                        DataRow[] InvRows = TDT.Select("InvNumber = '" + Rows[r]["InvNumber"].ToString() +
                                                       "' AND ColorID = " + Rows[r]["ColorID"].ToString() +
                                                       " AND PatinaID = " + Rows[r]["PatinaID"].ToString());
                        if (InvRows.Count() == 0)
                        {
                            DataRow NewRow = TDT.NewRow();
                            NewRow["UNN"] = UNN;
                            NewRow["AccountingName"] = Rows[r]["AccountingName"].ToString();
                            NewRow["InvNumber"] = Rows[r]["InvNumber"].ToString();
                            NewRow["MarketingCost"] = MarketingCost;
                            NewRow["CurrencyCode"] = ProfilCurrencyCode;
                            NewRow["Cvet"] = GetColorCode(ColorID);
                            NewRow["Patina"] = GetPatinaCode(PatinaID);
                            NewRow["TPSCurCode"] = TPSCurrencyCode;
                            NewRow["Measure"] = "м.п.";
                            NewRow["Count"] = Convert.ToDecimal(Rows[r]["Count"]) * L;
                            //NewRow["Price"] = Convert.ToDecimal(Rows[r]["Price"]);
                            NewRow["Cost"] = Convert.ToDecimal(Rows[r]["Cost"]);
                            NewRow["Weight"] = GetDecorWeight(Rows[r]);
                            TDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            InvRows[0]["Count"] = Convert.ToDecimal(InvRows[0]["Count"]) +
                                Convert.ToDecimal(Rows[r]["Count"]) * L;
                            //InvRows[0]["Price"] = Convert.ToDecimal(Rows[r]["Price"]);
                            InvRows[0]["Cost"] = Convert.ToDecimal(InvRows[0]["Cost"]) +
                                Convert.ToDecimal(Rows[r]["Cost"]);
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
                                                       "' AND ColorID = " + Rows[r]["ColorID"].ToString() +
                                                       " AND PatinaID = " + Rows[r]["PatinaID"].ToString());
                        if (InvRows.Count() == 0)
                        {
                            DataRow NewRow = PDT.NewRow();
                            NewRow["UNN"] = UNN;
                            NewRow["AccountingName"] = Rows[r]["AccountingName"].ToString();
                            NewRow["InvNumber"] = Rows[r]["InvNumber"].ToString();
                            NewRow["MarketingCost"] = MarketingCost;
                            NewRow["CurrencyCode"] = ProfilCurrencyCode;
                            NewRow["Cvet"] = GetColorCode(ColorID);
                            NewRow["Patina"] = GetPatinaCode(PatinaID);
                            NewRow["Measure"] = "м.кв.";
                            if (GetMeasureTypeID(Convert.ToInt32(Rows[r]["DecorConfigID"])) == 1)
                            {
                                if (Convert.ToDecimal(Rows[r]["Height"]) != -1)
                                    NewRow["Count"] = Decimal.Round(Convert.ToDecimal(Rows[r]["Height"]) * Convert.ToDecimal(Rows[r]["Width"]) / 1000000, 3, MidpointRounding.AwayFromZero)
                                    * Convert.ToDecimal(Rows[r]["Count"]);
                                if (Convert.ToDecimal(Rows[r]["Length"]) != -1)
                                    NewRow["Count"] = Decimal.Round(Convert.ToDecimal(Rows[r]["Length"]) * Convert.ToDecimal(Rows[r]["Width"]) / 1000000, 3, MidpointRounding.AwayFromZero)
                                    * Convert.ToDecimal(Rows[r]["Count"]);
                            }

                            if (GetMeasureTypeID(Convert.ToInt32(Rows[r]["DecorConfigID"])) == 2)
                            {
                                decimal L = 0;
                                L = Convert.ToDecimal(Rows[r]["Length"]);
                                if (L == -1)
                                    L = Convert.ToDecimal(Rows[r]["Height"]);
                                NewRow["Count"] = Convert.ToDecimal(Rows[r]["Count"]) *
                                    Decimal.Round(L * Convert.ToDecimal(Rows[r]["Width"]) / 1000000, 3, MidpointRounding.AwayFromZero);
                            }

                            if (GetMeasureTypeID(Convert.ToInt32(Rows[r]["DecorConfigID"])) == 3)
                            {
                                decimal H = 0;
                                if (DecorCatalogOrder.HasParameter(Convert.ToInt32(Rows[r]["ProductID"]), "Height") && Convert.ToDecimal(Rows[r]["Height"]) != -1)
                                    H = Convert.ToDecimal(Rows[r]["Height"]);
                                if (DecorCatalogOrder.HasParameter(Convert.ToInt32(Rows[r]["ProductID"]), "Length") && Convert.ToDecimal(Rows[r]["Length"]) != -1)
                                    H = Convert.ToDecimal(Rows[r]["Length"]);
                                NewRow["Count"] = Convert.ToDecimal(Rows[r]["Count"]) *
                                    Decimal.Round(H * Convert.ToDecimal(Rows[r]["Width"]) / 1000000, 3, MidpointRounding.AwayFromZero);
                            }
                            //NewRow["Price"] = Convert.ToDecimal(Rows[r]["Price"]);
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
                                    InvRows[0]["Count"] = Convert.ToDecimal(InvRows[0]["Count"]) +
                                    Decimal.Round(Convert.ToDecimal(Rows[r]["Height"]) * Convert.ToDecimal(Rows[r]["Width"]) / 1000000, 3, MidpointRounding.AwayFromZero)
                                    * Convert.ToDecimal(Rows[r]["Count"]);
                                if (Convert.ToDecimal(Rows[r]["Length"]) != -1)
                                    InvRows[0]["Count"] = Convert.ToDecimal(InvRows[0]["Count"]) +
                                        Decimal.Round(Convert.ToDecimal(Rows[r]["Length"]) * Convert.ToDecimal(Rows[r]["Width"]) / 1000000, 3, MidpointRounding.AwayFromZero)
                                    * Convert.ToDecimal(Rows[r]["Count"]);
                            }
                            if (GetMeasureTypeID(Convert.ToInt32(Rows[0]["DecorConfigID"])) == 2)
                            {
                                decimal L = 0;
                                L = Convert.ToDecimal(Rows[r]["Length"]);
                                if (L == -1)
                                    L = Convert.ToDecimal(Rows[r]["Height"]);
                                InvRows[0]["Count"] = Convert.ToDecimal(InvRows[0]["Count"]) +
                                    Convert.ToDecimal(Rows[r]["Count"]) *
                                    Decimal.Round(L * Convert.ToDecimal(Rows[r]["Width"]) / 1000000, 3, MidpointRounding.AwayFromZero);
                            }
                            if (GetMeasureTypeID(Convert.ToInt32(Rows[r]["DecorConfigID"])) == 3)
                            {
                                decimal H = 0;
                                if (DecorCatalogOrder.HasParameter(Convert.ToInt32(Rows[r]["ProductID"]), "Height") && Convert.ToDecimal(Rows[r]["Height"]) != -1)
                                    H = Convert.ToDecimal(Rows[r]["Height"]);
                                if (DecorCatalogOrder.HasParameter(Convert.ToInt32(Rows[r]["ProductID"]), "Length") && Convert.ToDecimal(Rows[r]["Length"]) != -1)
                                    H = Convert.ToDecimal(Rows[r]["Length"]);
                                InvRows[0]["Count"] = Convert.ToDecimal(InvRows[0]["Count"]) +
                                    Convert.ToDecimal(Rows[r]["Count"]) *
                                    Decimal.Round(H * Convert.ToDecimal(Rows[r]["Width"]) / 1000000, 3, MidpointRounding.AwayFromZero);
                            }
                            //InvRows[0]["Price"] = Convert.ToDecimal(Rows[r]["Price"]);
                            InvRows[0]["Cost"] = Convert.ToDecimal(InvRows[0]["Cost"]) +
                                Convert.ToDecimal(Rows[r]["Cost"]);
                            InvRows[0]["Weight"] = Convert.ToDecimal(InvRows[0]["Weight"]) +
                                GetDecorWeight(Rows[r]);
                        }
                    }
                    else
                    {
                        DataRow[] InvRows = TDT.Select("InvNumber = '" + Rows[r]["InvNumber"].ToString() +
                                                       "' AND ColorID = " + Rows[r]["ColorID"].ToString() +
                                                       " AND PatinaID = " + Rows[r]["PatinaID"].ToString());
                        if (InvRows.Count() == 0)
                        {
                            DataRow NewRow = TDT.NewRow();
                            NewRow["UNN"] = UNN;
                            NewRow["AccountingName"] = Rows[r]["AccountingName"].ToString();
                            NewRow["InvNumber"] = Rows[r]["InvNumber"].ToString();
                            NewRow["MarketingCost"] = MarketingCost;
                            NewRow["CurrencyCode"] = ProfilCurrencyCode;
                            NewRow["Cvet"] = GetColorCode(ColorID);
                            NewRow["Patina"] = GetPatinaCode(PatinaID);
                            NewRow["TPSCurCode"] = TPSCurrencyCode;
                            NewRow["Measure"] = "м.кв.";
                            if (GetMeasureTypeID(Convert.ToInt32(Rows[r]["DecorConfigID"])) == 1)
                            {
                                if (Convert.ToDecimal(Rows[r]["Height"]) != -1)
                                    NewRow["Count"] = Decimal.Round(Convert.ToDecimal(Rows[r]["Height"]) * Convert.ToDecimal(Rows[r]["Width"]) / 1000000, 3, MidpointRounding.AwayFromZero)
                                    * Convert.ToDecimal(Rows[r]["Count"]);
                                if (Convert.ToDecimal(Rows[r]["Length"]) != -1)
                                    NewRow["Count"] = Decimal.Round(Convert.ToDecimal(Rows[r]["Length"]) * Convert.ToDecimal(Rows[r]["Width"]) / 1000000, 3, MidpointRounding.AwayFromZero)
                                    * Convert.ToDecimal(Rows[r]["Count"]);
                            }

                            if (GetMeasureTypeID(Convert.ToInt32(Rows[0]["DecorConfigID"])) == 2)
                            {
                                decimal L = 0;
                                L = Convert.ToDecimal(Rows[r]["Length"]);
                                if (L == -1)
                                    L = Convert.ToDecimal(Rows[r]["Height"]);
                                NewRow["Count"] = Convert.ToDecimal(Rows[r]["Count"]) *
                                    Decimal.Round(L * Convert.ToDecimal(Rows[r]["Width"]) / 1000000, 3, MidpointRounding.AwayFromZero);
                            }

                            if (GetMeasureTypeID(Convert.ToInt32(Rows[r]["DecorConfigID"])) == 3)
                            {
                                decimal H = 0;
                                if (DecorCatalogOrder.HasParameter(Convert.ToInt32(Rows[r]["ProductID"]), "Height") && Convert.ToDecimal(Rows[r]["Height"]) != -1)
                                    H = Convert.ToDecimal(Rows[r]["Height"]);
                                if (DecorCatalogOrder.HasParameter(Convert.ToInt32(Rows[r]["ProductID"]), "Length") && Convert.ToDecimal(Rows[r]["Length"]) != -1)
                                    H = Convert.ToDecimal(Rows[r]["Length"]);
                                NewRow["Count"] = Convert.ToDecimal(Rows[r]["Count"]) *
                                    Decimal.Round(H * Convert.ToDecimal(Rows[r]["Width"]) / 1000000, 3, MidpointRounding.AwayFromZero);
                            }
                            //NewRow["Price"] = Convert.ToDecimal(Rows[r]["Price"]);
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
                                    InvRows[0]["Count"] = Convert.ToDecimal(InvRows[0]["Count"]) +
                                    Decimal.Round(Convert.ToDecimal(Rows[r]["Height"]) * Convert.ToDecimal(Rows[r]["Width"]) / 1000000, 3, MidpointRounding.AwayFromZero)
                                    * Convert.ToDecimal(Rows[r]["Count"]);
                                if (Convert.ToDecimal(Rows[r]["Length"]) != -1)
                                    InvRows[0]["Count"] = Convert.ToDecimal(InvRows[0]["Count"]) +
                                        Decimal.Round(Convert.ToDecimal(Rows[r]["Length"]) * Convert.ToDecimal(Rows[r]["Width"]) / 1000000, 3, MidpointRounding.AwayFromZero)
                                    * Convert.ToDecimal(Rows[r]["Count"]);
                            }
                            if (GetMeasureTypeID(Convert.ToInt32(Rows[r]["DecorConfigID"])) == 2)
                            {
                                decimal L = 0;
                                L = Convert.ToDecimal(Rows[r]["Length"]);
                                if (L == -1)
                                    L = Convert.ToDecimal(Rows[r]["Height"]);
                                InvRows[0]["Count"] = Convert.ToDecimal(InvRows[0]["Count"]) +
                                    Convert.ToDecimal(Rows[r]["Count"]) *
                                    Decimal.Round(L * Convert.ToDecimal(Rows[r]["Width"]) / 1000000, 3, MidpointRounding.AwayFromZero); ;
                            }
                            if (GetMeasureTypeID(Convert.ToInt32(Rows[r]["DecorConfigID"])) == 3)
                            {
                                decimal H = 0;
                                if (DecorCatalogOrder.HasParameter(Convert.ToInt32(Rows[r]["ProductID"]), "Height") && Convert.ToDecimal(Rows[r]["Height"]) != -1)
                                    H = Convert.ToDecimal(Rows[r]["Height"]);
                                if (DecorCatalogOrder.HasParameter(Convert.ToInt32(Rows[r]["ProductID"]), "Length") && Convert.ToDecimal(Rows[r]["Length"]) != -1)
                                    H = Convert.ToDecimal(Rows[r]["Length"]);
                                InvRows[0]["Count"] = Convert.ToDecimal(InvRows[0]["Count"]) +
                                    Convert.ToDecimal(Rows[r]["Count"]) *
                                    Decimal.Round(H * Convert.ToDecimal(Rows[r]["Width"]) / 1000000, 3, MidpointRounding.AwayFromZero); ;
                            }
                            //InvRows[0]["Price"] = Convert.ToDecimal(Rows[r]["Price"]);
                            InvRows[0]["Cost"] = Convert.ToDecimal(InvRows[0]["Cost"]) +
                                Convert.ToDecimal(Rows[r]["Cost"]);
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
                        DataRow NewRow = PRepDT.NewRow();

                        NewRow["UNN"] = UNN;
                        NewRow["AccountingName"] = PDT.Rows[i]["AccountingName"].ToString();
                        NewRow["InvNumber"] = PDT.Rows[i]["InvNumber"].ToString();
                        NewRow["MarketingCost"] = Convert.ToDecimal(PDT.Rows[i]["MarketingCost"]);
                        NewRow["Cvet"] = PDT.Rows[i]["Cvet"].ToString();
                        NewRow["Patina"] = PDT.Rows[i]["Patina"].ToString();
                        NewRow["CurrencyCode"] = ProfilCurrencyCode;
                        NewRow["Measure"] = "м.п.";
                        NewRow["Count"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[i]["Count"]) / 1000, 3, MidpointRounding.AwayFromZero);
                        //NewRow["Price"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[i]["Cost"]) / (Convert.ToDecimal(PDT.Rows[i]["Count"]) / 1000), 3, MidpointRounding.AwayFromZero);
                        NewRow["Cost"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[i]["Cost"]), 3, MidpointRounding.AwayFromZero);
                        NewRow["Weight"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[i]["Weight"]), 3, MidpointRounding.AwayFromZero);
                        PRepDT.Rows.Add(NewRow);
                    }
                }

                if (TDT.Rows.Count > 0)
                {
                    for (int i = 0; i < TDT.Rows.Count; i++)
                    {
                        DataRow NewRow = TRepDT.NewRow();

                        NewRow["UNN"] = UNN;
                        NewRow["AccountingName"] = TDT.Rows[i]["AccountingName"].ToString();
                        NewRow["InvNumber"] = TDT.Rows[i]["InvNumber"].ToString();
                        NewRow["MarketingCost"] = Convert.ToDecimal(TDT.Rows[i]["MarketingCost"]);
                        NewRow["Cvet"] = TDT.Rows[i]["Cvet"].ToString();
                        NewRow["Patina"] = TDT.Rows[i]["Patina"].ToString();
                        NewRow["CurrencyCode"] = ProfilCurrencyCode;
                        NewRow["TPSCurCode"] = TPSCurrencyCode;
                        NewRow["Measure"] = "м.п.";
                        NewRow["Count"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[i]["Count"]) / 1000, 3, MidpointRounding.AwayFromZero);
                        //NewRow["Price"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[i]["Cost"]) / (Convert.ToDecimal(TDT.Rows[i]["Count"]) / 1000), 3, MidpointRounding.AwayFromZero);
                        NewRow["Cost"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[i]["Cost"]), 3, MidpointRounding.AwayFromZero);
                        NewRow["Weight"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[i]["Weight"]), 3, MidpointRounding.AwayFromZero);

                        TRepDT.Rows.Add(NewRow);
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
                        DataRow NewRow = PRepDT.NewRow();

                        NewRow["UNN"] = UNN;
                        NewRow["AccountingName"] = PDT.Rows[i]["AccountingName"].ToString();
                        NewRow["InvNumber"] = PDT.Rows[i]["InvNumber"].ToString();
                        NewRow["MarketingCost"] = Convert.ToDecimal(PDT.Rows[i]["MarketingCost"]);
                        NewRow["Cvet"] = PDT.Rows[i]["Cvet"].ToString();
                        NewRow["Patina"] = PDT.Rows[i]["Patina"].ToString();
                        NewRow["CurrencyCode"] = ProfilCurrencyCode;
                        NewRow["Measure"] = "м.кв.";
                        NewRow["Count"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[i]["Count"]), 3, MidpointRounding.AwayFromZero);
                        //NewRow["Price"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[i]["Cost"]) / Convert.ToDecimal(PDT.Rows[i]["Count"]), 3, MidpointRounding.AwayFromZero);
                        NewRow["Cost"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[i]["Cost"]), 3, MidpointRounding.AwayFromZero);
                        NewRow["Weight"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[i]["Weight"]), 3, MidpointRounding.AwayFromZero);

                        PRepDT.Rows.Add(NewRow);
                    }
                }

                if (TDT.Rows.Count > 0)
                {
                    for (int i = 0; i < TDT.Rows.Count; i++)
                    {
                        DataRow NewRow = TRepDT.NewRow();

                        NewRow["UNN"] = UNN;
                        NewRow["AccountingName"] = TDT.Rows[i]["AccountingName"].ToString();
                        NewRow["InvNumber"] = TDT.Rows[i]["InvNumber"].ToString();
                        NewRow["MarketingCost"] = Convert.ToDecimal(TDT.Rows[i]["MarketingCost"]);
                        NewRow["Cvet"] = TDT.Rows[i]["Cvet"].ToString();
                        NewRow["Patina"] = TDT.Rows[i]["Patina"].ToString();
                        NewRow["CurrencyCode"] = ProfilCurrencyCode;
                        NewRow["TPSCurCode"] = TPSCurrencyCode;
                        NewRow["Measure"] = "м.кв.";
                        NewRow["Count"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[i]["Count"]), 3, MidpointRounding.AwayFromZero);
                        //NewRow["Price"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[i]["Cost"]) / Convert.ToDecimal(TDT.Rows[i]["Count"]), 3, MidpointRounding.AwayFromZero);
                        NewRow["Cost"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[i]["Cost"]), 3, MidpointRounding.AwayFromZero);
                        NewRow["Weight"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[i]["Weight"]), 3, MidpointRounding.AwayFromZero);

                        TRepDT.Rows.Add(NewRow);
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

            decimal MarketingCost = 0;
            if (p1.Length == 0 && p2.Length == 0 && p3.Length == 0)
            {
                if (Rows.Count() > 0)
                    MarketingCost = GetMarketingCost(Convert.ToInt32(Rows[0]["DecorConfigID"]));
                for (int r = 0; r < Rows.Count(); r++)
                {
                    string Cvet = GetColorCode(Convert.ToInt32(Rows[r]["ColorID"]));
                    string Patina = GetPatinaCode(Convert.ToInt32(Rows[r]["PatinaID"]));
                    string InvNumber = Rows[r]["InvNumber"].ToString();
                    if (MarketingCost > GetMarketingCost(Convert.ToInt32(Rows[r]["DecorConfigID"])))
                        MarketingCost = GetMarketingCost(Convert.ToInt32(Rows[r]["DecorConfigID"]));
                    if (IsProfil(Convert.ToInt32(Rows[r]["DecorConfigID"])))
                    {
                        DataRow[] InvRows = PDT.Select($"InvNumber = '{ Rows[r]["InvNumber"].ToString() }' AND Cvet = '{ Cvet }' AND Patina = '{ Patina }'");
                        if (InvRows.Count() == 0)
                        {
                            DataRow NewRow = PDT.NewRow();
                            NewRow["UNN"] = UNN;
                            NewRow["AccountingName"] = Rows[r]["AccountingName"].ToString();
                            NewRow["InvNumber"] = Rows[r]["InvNumber"].ToString();
                            NewRow["MarketingCost"] = MarketingCost;
                            NewRow["Cvet"] = Cvet;
                            NewRow["Patina"] = Patina;
                            NewRow["CurrencyCode"] = ProfilCurrencyCode;
                            NewRow["Measure"] = "шт.";
                            NewRow["Count"] = Convert.ToDecimal(Rows[r]["Count"]);
                            //NewRow["Price"] = Convert.ToDecimal(Rows[r]["Price"]);
                            NewRow["Cost"] = Convert.ToDecimal(Rows[r]["Cost"]);
                            NewRow["Weight"] = GetDecorWeight(Rows[r]);
                            PDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            InvRows[0]["Count"] = Convert.ToDecimal(InvRows[0]["Count"]) +
                                Convert.ToDecimal(Rows[r]["Count"]);
                            //InvRows[0]["Price"] = Convert.ToDecimal(Rows[r]["Price"]);
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
                            NewRow["AccountingName"] = Rows[r]["AccountingName"].ToString();
                            NewRow["InvNumber"] = Rows[r]["InvNumber"].ToString();
                            NewRow["MarketingCost"] = MarketingCost;
                            NewRow["CurrencyCode"] = ProfilCurrencyCode;
                            NewRow["Cvet"] = Cvet;
                            NewRow["Patina"] = Patina;
                            NewRow["TPSCurCode"] = TPSCurrencyCode;
                            NewRow["Measure"] = "шт.";
                            NewRow["Count"] = Convert.ToDecimal(Rows[r]["Count"]);
                            //NewRow["Price"] = Convert.ToDecimal(Rows[r]["Price"]);
                            NewRow["Cost"] = Convert.ToDecimal(Rows[r]["Cost"]);
                            NewRow["Weight"] = GetDecorWeight(Rows[r]);
                            TDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            InvRows[0]["Count"] = Convert.ToDecimal(InvRows[0]["Count"]) +
                                Convert.ToDecimal(Rows[r]["Count"]);
                            //InvRows[0]["Price"] = Convert.ToDecimal(Rows[r]["Price"]);
                            InvRows[0]["Cost"] = Convert.ToDecimal(InvRows[0]["Cost"]) +
                                Convert.ToDecimal(Rows[r]["Cost"]);
                            InvRows[0]["Weight"] = Convert.ToDecimal(InvRows[0]["Weight"]) +
                                GetDecorWeight(Rows[r]);
                        }
                    }
                }
            }

            if (p1.Length > 0 && p2.Length > 0)
            {
                if (Rows.Count() > 0)
                    MarketingCost = GetMarketingCost(Convert.ToInt32(Rows[0]["DecorConfigID"]));
                for (int r = 0; r < Rows.Count(); r++)
                {
                    string Cvet = GetColorCode(Convert.ToInt32(Rows[r]["ColorID"]));
                    string Patina = GetPatinaCode(Convert.ToInt32(Rows[r]["PatinaID"]));
                    if (MarketingCost > GetMarketingCost(Convert.ToInt32(Rows[r]["DecorConfigID"])))
                        MarketingCost = GetMarketingCost(Convert.ToInt32(Rows[r]["DecorConfigID"]));
                    if (IsProfil(Convert.ToInt32(Rows[r]["DecorConfigID"])))
                    {
                        DataRow[] InvRows = PDT.Select($"InvNumber = '{ Rows[r]["InvNumber"].ToString() }' AND Cvet = '{ Cvet }' AND Patina = '{ Patina }'");
                        if (InvRows.Count() == 0)
                        {
                            DataRow NewRow = PDT.NewRow();
                            NewRow["UNN"] = UNN;
                            NewRow["AccountingName"] = Rows[r]["AccountingName"].ToString();
                            NewRow["InvNumber"] = Rows[r]["InvNumber"].ToString();
                            NewRow["MarketingCost"] = MarketingCost;
                            NewRow["Cvet"] = Cvet;
                            NewRow["Patina"] = Patina;
                            NewRow["CurrencyCode"] = ProfilCurrencyCode;
                            NewRow["Measure"] = "шт.";
                            NewRow[p1] = Convert.ToDecimal(Rows[r][p1]);
                            NewRow[p2] = Convert.ToDecimal(Rows[r][p2]);
                            NewRow["Count"] = Convert.ToDecimal(Rows[r]["Count"]);
                            //NewRow["Price"] = Convert.ToDecimal(Rows[r]["Price"]);
                            NewRow["Cost"] = Convert.ToDecimal(Rows[r]["Cost"]);
                            NewRow["Weight"] = GetDecorWeight(Rows[r]);
                            PDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            InvRows[0]["Count"] = Convert.ToDecimal(InvRows[0]["Count"]) +
                                Convert.ToDecimal(Rows[r]["Count"]);
                            //InvRows[0]["Price"] = Convert.ToDecimal(Rows[r]["Price"]);
                            InvRows[0]["Cost"] = Convert.ToDecimal(InvRows[0]["Cost"]) +
                                Convert.ToDecimal(Rows[r]["Cost"]);
                            InvRows[0]["Weight"] = Convert.ToDecimal(InvRows[0]["Weight"]) +
                                GetDecorWeight(Rows[r]);
                        }
                    }
                    else
                    {
                        DataRow[] InvRows = TDT.Select($"InvNumber = '{ Rows[r]["InvNumber"].ToString() }' AND Cvet = '{ Cvet }' AND Patina = '{ Patina }'");
                        if (InvRows.Count() == 0)
                        {
                            DataRow NewRow = TDT.NewRow();
                            NewRow["UNN"] = UNN;
                            NewRow["AccountingName"] = Rows[r]["AccountingName"].ToString();
                            NewRow["InvNumber"] = Rows[r]["InvNumber"].ToString();
                            NewRow["MarketingCost"] = MarketingCost;
                            NewRow["CurrencyCode"] = ProfilCurrencyCode;
                            NewRow["Cvet"] = Cvet;
                            NewRow["Patina"] = Patina;
                            NewRow["Measure"] = "шт.";
                            NewRow["TPSCurCode"] = TPSCurrencyCode;
                            NewRow[p1] = Convert.ToDecimal(Rows[r][p1]);
                            NewRow[p2] = Convert.ToDecimal(Rows[r][p2]);
                            NewRow["Count"] = Convert.ToDecimal(Rows[r]["Count"]);
                            //NewRow["Price"] = Convert.ToDecimal(Rows[r]["Price"]);
                            NewRow["Cost"] = Convert.ToDecimal(Rows[r]["Cost"]);
                            NewRow["Weight"] = GetDecorWeight(Rows[r]);
                            TDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            InvRows[0]["Count"] = Convert.ToDecimal(InvRows[0]["Count"]) +
                                Convert.ToDecimal(Rows[r]["Count"]);
                            //InvRows[0]["Price"] = Convert.ToDecimal(Rows[r]["Price"]);
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
                if (Rows.Count() > 0)
                    MarketingCost = GetMarketingCost(Convert.ToInt32(Rows[0]["DecorConfigID"]));
                for (int r = 0; r < Rows.Count(); r++)
                {
                    string Cvet = GetColorCode(Convert.ToInt32(Rows[r]["ColorID"]));
                    string Patina = GetPatinaCode(Convert.ToInt32(Rows[r]["PatinaID"]));
                    if (MarketingCost > GetMarketingCost(Convert.ToInt32(Rows[r]["DecorConfigID"])))
                        MarketingCost = GetMarketingCost(Convert.ToInt32(Rows[r]["DecorConfigID"]));
                    if (IsProfil(Convert.ToInt32(Rows[r]["DecorConfigID"])))
                    {
                        DataRow[] InvRows = PDT.Select($"InvNumber = '{ Rows[r]["InvNumber"].ToString() }' AND Cvet = '{ Cvet }' AND Patina = '{ Patina }'");
                        if (InvRows.Count() == 0)
                        {
                            DataRow NewRow = PDT.NewRow();
                            NewRow["UNN"] = UNN;
                            NewRow["AccountingName"] = Rows[r]["AccountingName"].ToString();
                            NewRow["InvNumber"] = Rows[r]["InvNumber"].ToString();
                            NewRow["MarketingCost"] = MarketingCost;
                            NewRow["CurrencyCode"] = ProfilCurrencyCode;
                            NewRow["Cvet"] = Cvet;
                            NewRow["Patina"] = Patina;
                            NewRow["Measure"] = "шт.";
                            NewRow[p3] = Convert.ToDecimal(Rows[r][p3]);
                            NewRow["Count"] = Convert.ToDecimal(Rows[r]["Count"]);
                            //NewRow["Price"] = Convert.ToDecimal(Rows[r]["Price"]);
                            NewRow["Cost"] = Convert.ToDecimal(Rows[r]["Cost"]);
                            NewRow["Weight"] = GetDecorWeight(Rows[r]);
                            PDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            InvRows[0]["Count"] = Convert.ToDecimal(InvRows[0]["Count"]) +
                                Convert.ToDecimal(Rows[r]["Count"]);
                            //InvRows[0]["Price"] = Convert.ToDecimal(Rows[r]["Price"]);
                            InvRows[0]["Cost"] = Convert.ToDecimal(InvRows[0]["Cost"]) +
                                Convert.ToDecimal(Rows[r]["Cost"]);
                            InvRows[0]["Weight"] = Convert.ToDecimal(InvRows[0]["Weight"]) +
                                GetDecorWeight(Rows[r]);
                        }
                    }
                    else
                    {
                        DataRow[] InvRows = TDT.Select($"InvNumber = '{ Rows[r]["InvNumber"].ToString() }' AND Cvet = '{ Cvet }' AND Patina = '{ Patina }'");
                        if (InvRows.Count() == 0)
                        {
                            DataRow NewRow = TDT.NewRow();
                            NewRow["UNN"] = UNN;
                            NewRow["AccountingName"] = Rows[r]["AccountingName"].ToString();
                            NewRow["InvNumber"] = Rows[r]["InvNumber"].ToString();
                            NewRow["MarketingCost"] = MarketingCost;
                            NewRow["CurrencyCode"] = ProfilCurrencyCode;
                            NewRow["Cvet"] = Cvet;
                            NewRow["Patina"] = Patina;
                            NewRow["Measure"] = "шт.";
                            NewRow["TPSCurCode"] = TPSCurrencyCode;
                            NewRow[p3] = Convert.ToDecimal(Rows[r][p3]);
                            NewRow["Count"] = Convert.ToDecimal(Rows[r]["Count"]);
                            //NewRow["Price"] = Convert.ToDecimal(Rows[r]["Price"]);
                            NewRow["Cost"] = Convert.ToDecimal(Rows[r]["Cost"]);
                            NewRow["Weight"] = GetDecorWeight(Rows[r]);
                            TDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            InvRows[0]["Count"] = Convert.ToDecimal(InvRows[0]["Count"]) +
                                Convert.ToDecimal(Rows[r]["Count"]);
                            //InvRows[0]["Price"] = Convert.ToDecimal(Rows[r]["Price"]);
                            InvRows[0]["Cost"] = Convert.ToDecimal(InvRows[0]["Cost"]) +
                                Convert.ToDecimal(Rows[r]["Cost"]);
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
                        NewRow["AccountingName"] = PDT.Rows[g]["AccountingName"].ToString();
                        NewRow["InvNumber"] = PDT.Rows[g]["InvNumber"].ToString();
                        NewRow["MarketingCost"] = Convert.ToDecimal(PDT.Rows[g]["MarketingCost"]);
                        NewRow["CurrencyCode"] = ProfilCurrencyCode;
                        NewRow["Cvet"] = PDT.Rows[g]["Cvet"].ToString();
                        NewRow["Patina"] = PDT.Rows[g]["Patina"].ToString();
                        NewRow["Measure"] = "шт.";
                        NewRow["Count"] = PDT.Rows[g]["Count"];
                        //NewRow["Price"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[g]["Price"]), 3, MidpointRounding.AwayFromZero);
                        NewRow["Cost"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[g]["Cost"]), 3, MidpointRounding.AwayFromZero);
                        NewRow["Weight"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[g]["Weight"]), 3, MidpointRounding.AwayFromZero);

                        ProfilReportDataTable.Rows.Add(NewRow);
                    }

                    if (p3.Length > 0)
                    {
                        DataRow NewRow = ProfilReportDataTable.NewRow();

                        NewRow["UNN"] = UNN;
                        NewRow["AccountingName"] = PDT.Rows[g]["AccountingName"].ToString();
                        NewRow["InvNumber"] = PDT.Rows[g]["InvNumber"].ToString();
                        NewRow["MarketingCost"] = Convert.ToDecimal(PDT.Rows[g]["MarketingCost"]);
                        NewRow["Cvet"] = PDT.Rows[g]["Cvet"].ToString();
                        NewRow["Patina"] = PDT.Rows[g]["Patina"].ToString();
                        NewRow["CurrencyCode"] = ProfilCurrencyCode;
                        NewRow["Measure"] = "шт.";
                        NewRow["Count"] = PDT.Rows[g]["Count"];
                        //NewRow["Price"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[g]["Price"]), 3, MidpointRounding.AwayFromZero);
                        NewRow["Cost"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[g]["Cost"]), 3, MidpointRounding.AwayFromZero);
                        NewRow["Weight"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[g]["Weight"]), 3, MidpointRounding.AwayFromZero);

                        ProfilReportDataTable.Rows.Add(NewRow);
                    }

                    if (p1.Length == 0 && p2.Length == 0 && p3.Length == 0)
                    {
                        DataRow NewRow = ProfilReportDataTable.NewRow();

                        NewRow["UNN"] = UNN;
                        NewRow["AccountingName"] = PDT.Rows[g]["AccountingName"].ToString();
                        NewRow["InvNumber"] = PDT.Rows[g]["InvNumber"].ToString();
                        NewRow["MarketingCost"] = Convert.ToDecimal(PDT.Rows[g]["MarketingCost"]);
                        NewRow["Cvet"] = PDT.Rows[g]["Cvet"].ToString();
                        NewRow["Patina"] = PDT.Rows[g]["Patina"].ToString();
                        NewRow["CurrencyCode"] = ProfilCurrencyCode;
                        NewRow["Measure"] = "шт.";
                        NewRow["Count"] = PDT.Rows[g]["Count"];
                        //NewRow["Price"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[g]["Price"]), 3, MidpointRounding.AwayFromZero);
                        NewRow["Cost"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[g]["Cost"]), 3, MidpointRounding.AwayFromZero);
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
                        NewRow["AccountingName"] = TDT.Rows[g]["AccountingName"].ToString();
                        NewRow["InvNumber"] = TDT.Rows[g]["InvNumber"].ToString();
                        NewRow["MarketingCost"] = Convert.ToDecimal(TDT.Rows[g]["MarketingCost"]);
                        NewRow["CurrencyCode"] = ProfilCurrencyCode;
                        NewRow["Cvet"] = TDT.Rows[g]["Cvet"].ToString();
                        NewRow["Patina"] = TDT.Rows[g]["Patina"].ToString();
                        NewRow["Measure"] = "шт.";
                        NewRow["TPSCurCode"] = TPSCurrencyCode;
                        NewRow["Count"] = TDT.Rows[g]["Count"];
                        //NewRow["Price"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[g]["Price"]), 3, MidpointRounding.AwayFromZero);
                        NewRow["Cost"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[g]["Cost"]), 3, MidpointRounding.AwayFromZero);
                        NewRow["Weight"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[g]["Weight"]), 3, MidpointRounding.AwayFromZero);

                        TPSReportDataTable.Rows.Add(NewRow);
                    }

                    if (p3.Length > 0)
                    {
                        DataRow NewRow = TPSReportDataTable.NewRow();

                        NewRow["UNN"] = UNN;
                        NewRow["AccountingName"] = TDT.Rows[g]["AccountingName"].ToString();
                        NewRow["InvNumber"] = TDT.Rows[g]["InvNumber"].ToString();
                        NewRow["MarketingCost"] = Convert.ToDecimal(TDT.Rows[g]["MarketingCost"]);
                        NewRow["CurrencyCode"] = ProfilCurrencyCode;
                        NewRow["Cvet"] = TDT.Rows[g]["Cvet"].ToString();
                        NewRow["Patina"] = TDT.Rows[g]["Patina"].ToString();
                        NewRow["Measure"] = "шт.";
                        NewRow["TPSCurCode"] = TPSCurrencyCode;
                        NewRow["Count"] = TDT.Rows[g]["Count"];
                        //NewRow["Price"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[g]["Price"]), 3, MidpointRounding.AwayFromZero);
                        NewRow["Cost"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[g]["Cost"]), 3, MidpointRounding.AwayFromZero);
                        NewRow["Weight"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[g]["Weight"]), 3, MidpointRounding.AwayFromZero);

                        TPSReportDataTable.Rows.Add(NewRow);
                    }

                    if (p1.Length == 0 && p2.Length == 0 && p3.Length == 0)
                    {
                        DataRow NewRow = TPSReportDataTable.NewRow();

                        NewRow["UNN"] = UNN;
                        NewRow["AccountingName"] = TDT.Rows[g]["AccountingName"].ToString();
                        NewRow["InvNumber"] = TDT.Rows[g]["InvNumber"].ToString();
                        NewRow["MarketingCost"] = Convert.ToDecimal(TDT.Rows[g]["MarketingCost"]);
                        NewRow["CurrencyCode"] = ProfilCurrencyCode;
                        NewRow["Measure"] = "шт.";
                        NewRow["Cvet"] = TDT.Rows[g]["Cvet"].ToString();
                        NewRow["Patina"] = TDT.Rows[g]["Patina"].ToString();
                        NewRow["TPSCurCode"] = TPSCurrencyCode;
                        NewRow["Count"] = TDT.Rows[g]["Count"];
                        //NewRow["Price"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[g]["Price"]), 3, MidpointRounding.AwayFromZero);
                        NewRow["Cost"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[g]["Cost"]), 3, MidpointRounding.AwayFromZero);
                        NewRow["Weight"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[g]["Weight"]), 3, MidpointRounding.AwayFromZero);

                        TPSReportDataTable.Rows.Add(NewRow);
                    }
                }
            }
        }

        private void GetParametrizedData(DataRow[] Rows, DataTable PDT, DataTable TDT, DataTable PRepDT, DataTable TRepDT)
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

            decimal MarketingCost = 0;
            if (p1.Length == 0 && p2.Length == 0 && p3.Length == 0)
            {
                if (Rows.Count() > 0)
                    MarketingCost = GetMarketingCost(Convert.ToInt32(Rows[0]["DecorConfigID"]));
                for (int r = 0; r < Rows.Count(); r++)
                {
                    string Cvet = GetColorCode(Convert.ToInt32(Rows[r]["ColorID"]));
                    string Patina = GetPatinaCode(Convert.ToInt32(Rows[r]["PatinaID"]));
                    string InvNumber = Rows[r]["InvNumber"].ToString();
                    if (MarketingCost > GetMarketingCost(Convert.ToInt32(Rows[r]["DecorConfigID"])))
                        MarketingCost = GetMarketingCost(Convert.ToInt32(Rows[r]["DecorConfigID"]));
                    if (IsProfil(Convert.ToInt32(Rows[r]["DecorConfigID"])))
                    {
                        DataRow[] InvRows = PDT.Select($"InvNumber = '{ Rows[r]["InvNumber"].ToString() }' AND Cvet = '{ Cvet }' AND Patina = '{ Patina }'");
                        if (InvRows.Count() == 0)
                        {
                            DataRow NewRow = PDT.NewRow();
                            NewRow["UNN"] = UNN;
                            NewRow["AccountingName"] = Rows[r]["AccountingName"].ToString();
                            NewRow["InvNumber"] = Rows[r]["InvNumber"].ToString();
                            NewRow["MarketingCost"] = MarketingCost;
                            NewRow["Cvet"] = Cvet;
                            NewRow["Patina"] = Patina;
                            NewRow["CurrencyCode"] = ProfilCurrencyCode;
                            NewRow["Measure"] = "шт.";
                            NewRow["Count"] = Convert.ToDecimal(Rows[r]["Count"]);
                            //NewRow["Price"] = Convert.ToDecimal(Rows[r]["Price"]);
                            NewRow["Cost"] = Convert.ToDecimal(Rows[r]["Cost"]);
                            NewRow["Weight"] = GetDecorWeight(Rows[r]);
                            PDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            InvRows[0]["Count"] = Convert.ToDecimal(InvRows[0]["Count"]) +
                                Convert.ToDecimal(Rows[r]["Count"]);
                            //InvRows[0]["Price"] = Convert.ToDecimal(Rows[r]["Price"]);
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
                            NewRow["AccountingName"] = Rows[r]["AccountingName"].ToString();
                            NewRow["InvNumber"] = Rows[r]["InvNumber"].ToString();
                            NewRow["MarketingCost"] = MarketingCost;
                            NewRow["CurrencyCode"] = ProfilCurrencyCode;
                            NewRow["Cvet"] = Cvet;
                            NewRow["Patina"] = Patina;
                            NewRow["TPSCurCode"] = TPSCurrencyCode;
                            NewRow["Measure"] = "шт.";
                            NewRow["Count"] = Convert.ToDecimal(Rows[r]["Count"]);
                            //NewRow["Price"] = Convert.ToDecimal(Rows[r]["Price"]);
                            NewRow["Cost"] = Convert.ToDecimal(Rows[r]["Cost"]);
                            NewRow["Weight"] = GetDecorWeight(Rows[r]);
                            TDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            InvRows[0]["Count"] = Convert.ToDecimal(InvRows[0]["Count"]) +
                                Convert.ToDecimal(Rows[r]["Count"]);
                            //InvRows[0]["Price"] = Convert.ToDecimal(Rows[r]["Price"]);
                            InvRows[0]["Cost"] = Convert.ToDecimal(InvRows[0]["Cost"]) +
                                Convert.ToDecimal(Rows[r]["Cost"]);
                            InvRows[0]["Weight"] = Convert.ToDecimal(InvRows[0]["Weight"]) +
                                GetDecorWeight(Rows[r]);
                        }
                    }
                }
            }

            if (p1.Length > 0 && p2.Length > 0)
            {
                if (Rows.Count() > 0)
                    MarketingCost = GetMarketingCost(Convert.ToInt32(Rows[0]["DecorConfigID"]));
                for (int r = 0; r < Rows.Count(); r++)
                {
                    string Cvet = GetColorCode(Convert.ToInt32(Rows[r]["ColorID"]));
                    string Patina = GetPatinaCode(Convert.ToInt32(Rows[r]["PatinaID"]));
                    if (MarketingCost > GetMarketingCost(Convert.ToInt32(Rows[r]["DecorConfigID"])))
                        MarketingCost = GetMarketingCost(Convert.ToInt32(Rows[r]["DecorConfigID"]));
                    if (IsProfil(Convert.ToInt32(Rows[r]["DecorConfigID"])))
                    {
                        DataRow[] InvRows = PDT.Select($"InvNumber = '{ Rows[r]["InvNumber"].ToString() }' AND Cvet = '{ Cvet }' AND Patina = '{ Patina }'");
                        if (InvRows.Count() == 0)
                        {
                            DataRow NewRow = PDT.NewRow();
                            NewRow["UNN"] = UNN;
                            NewRow["AccountingName"] = Rows[r]["AccountingName"].ToString();
                            NewRow["InvNumber"] = Rows[r]["InvNumber"].ToString();
                            NewRow["MarketingCost"] = MarketingCost;
                            NewRow["CurrencyCode"] = ProfilCurrencyCode;
                            NewRow["Cvet"] = Cvet;
                            NewRow["Patina"] = Patina;
                            NewRow["Measure"] = "шт.";
                            NewRow[p1] = Convert.ToDecimal(Rows[r][p1]);
                            NewRow[p2] = Convert.ToDecimal(Rows[r][p2]);
                            NewRow["Count"] = Convert.ToDecimal(Rows[r]["Count"]);
                            //NewRow["Price"] = Convert.ToDecimal(Rows[r]["Price"]);
                            NewRow["Cost"] = Convert.ToDecimal(Rows[r]["Cost"]);
                            NewRow["Weight"] = GetDecorWeight(Rows[r]);
                            PDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            InvRows[0]["Count"] = Convert.ToDecimal(InvRows[0]["Count"]) +
                                Convert.ToDecimal(Rows[r]["Count"]);
                            //InvRows[0]["Price"] = Convert.ToDecimal(Rows[r]["Price"]);
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
                            NewRow["AccountingName"] = Rows[r]["AccountingName"].ToString();
                            NewRow["InvNumber"] = Rows[r]["InvNumber"].ToString();
                            NewRow["MarketingCost"] = MarketingCost;
                            NewRow["CurrencyCode"] = ProfilCurrencyCode;
                            NewRow["Cvet"] = Cvet;
                            NewRow["Patina"] = Patina;
                            NewRow["Measure"] = "шт.";
                            NewRow["TPSCurCode"] = TPSCurrencyCode;
                            NewRow[p1] = Convert.ToDecimal(Rows[r][p1]);
                            NewRow[p2] = Convert.ToDecimal(Rows[r][p2]);
                            NewRow["Count"] = Convert.ToDecimal(Rows[r]["Count"]);
                            //NewRow["Price"] = Convert.ToDecimal(Rows[r]["Price"]);
                            NewRow["Cost"] = Convert.ToDecimal(Rows[r]["Cost"]);
                            NewRow["Weight"] = GetDecorWeight(Rows[r]);
                            TDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            InvRows[0]["Count"] = Convert.ToDecimal(InvRows[0]["Count"]) +
                                Convert.ToDecimal(Rows[r]["Count"]);
                            //InvRows[0]["Price"] = Convert.ToDecimal(Rows[r]["Price"]);
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
                if (Rows.Count() > 0)
                    MarketingCost = GetMarketingCost(Convert.ToInt32(Rows[0]["DecorConfigID"]));
                for (int r = 0; r < Rows.Count(); r++)
                {
                    string Cvet = GetColorCode(Convert.ToInt32(Rows[r]["ColorID"]));
                    string Patina = GetPatinaCode(Convert.ToInt32(Rows[r]["PatinaID"]));
                    if (MarketingCost > GetMarketingCost(Convert.ToInt32(Rows[r]["DecorConfigID"])))
                        MarketingCost = GetMarketingCost(Convert.ToInt32(Rows[r]["DecorConfigID"]));
                    if (IsProfil(Convert.ToInt32(Rows[r]["DecorConfigID"])))
                    {
                        DataRow[] InvRows = PDT.Select($"InvNumber = '{ Rows[r]["InvNumber"].ToString() }' AND Cvet = '{ Cvet }' AND Patina = '{ Patina }'");
                        if (InvRows.Count() == 0)
                        {
                            DataRow NewRow = PDT.NewRow();
                            NewRow["UNN"] = UNN;
                            NewRow["AccountingName"] = Rows[r]["AccountingName"].ToString();
                            NewRow["InvNumber"] = Rows[r]["InvNumber"].ToString();
                            NewRow["MarketingCost"] = MarketingCost;
                            NewRow["CurrencyCode"] = ProfilCurrencyCode;
                            NewRow["Cvet"] = Cvet;
                            NewRow["Patina"] = Patina;
                            NewRow["Measure"] = "шт.";
                            NewRow[p3] = Convert.ToDecimal(Rows[r][p3]);
                            NewRow["Count"] = Convert.ToDecimal(Rows[r]["Count"]);
                            //NewRow["Price"] = Convert.ToDecimal(Rows[r]["Price"]);
                            NewRow["Cost"] = Convert.ToDecimal(Rows[r]["Cost"]);
                            NewRow["Weight"] = GetDecorWeight(Rows[r]);
                            PDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            InvRows[0]["Count"] = Convert.ToDecimal(InvRows[0]["Count"]) +
                                Convert.ToDecimal(Rows[r]["Count"]);
                            //InvRows[0]["Price"] = Convert.ToDecimal(Rows[r]["Price"]);
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
                            NewRow["AccountingName"] = Rows[r]["AccountingName"].ToString();
                            NewRow["InvNumber"] = Rows[r]["InvNumber"].ToString();
                            NewRow["MarketingCost"] = MarketingCost;
                            NewRow["CurrencyCode"] = ProfilCurrencyCode;
                            NewRow["Cvet"] = Cvet;
                            NewRow["Patina"] = Patina;
                            NewRow["Measure"] = "шт.";
                            NewRow["TPSCurCode"] = TPSCurrencyCode;
                            NewRow[p3] = Convert.ToDecimal(Rows[r][p3]);
                            NewRow["Count"] = Convert.ToDecimal(Rows[r]["Count"]);
                            //NewRow["Price"] = Convert.ToDecimal(Rows[r]["Price"]);
                            NewRow["Cost"] = Convert.ToDecimal(Rows[r]["Cost"]);
                            NewRow["Weight"] = GetDecorWeight(Rows[r]);
                            TDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            InvRows[0]["Count"] = Convert.ToDecimal(InvRows[0]["Count"]) +
                                Convert.ToDecimal(Rows[r]["Count"]);
                            //InvRows[0]["Price"] = Convert.ToDecimal(Rows[r]["Price"]);
                            InvRows[0]["Cost"] = Convert.ToDecimal(InvRows[0]["Cost"]) +
                                Convert.ToDecimal(Rows[r]["Cost"]);
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
                        DataRow NewRow = PRepDT.NewRow();

                        NewRow["UNN"] = UNN;
                        NewRow["AccountingName"] = PDT.Rows[g]["AccountingName"].ToString();
                        NewRow["InvNumber"] = PDT.Rows[g]["InvNumber"].ToString();
                        NewRow["MarketingCost"] = Convert.ToDecimal(PDT.Rows[g]["MarketingCost"]);
                        NewRow["Cvet"] = PDT.Rows[g]["Cvet"].ToString();
                        NewRow["Patina"] = PDT.Rows[g]["Patina"].ToString();
                        NewRow["CurrencyCode"] = ProfilCurrencyCode;
                        NewRow["Measure"] = "шт.";
                        NewRow["Count"] = PDT.Rows[g]["Count"];
                        //NewRow["Price"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[g]["Price"]), 3, MidpointRounding.AwayFromZero);
                        NewRow["Cost"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[g]["Cost"]), 3, MidpointRounding.AwayFromZero);
                        NewRow["Weight"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[g]["Weight"]), 3, MidpointRounding.AwayFromZero);

                        PRepDT.Rows.Add(NewRow);
                    }

                    if (p3.Length > 0)
                    {
                        DataRow NewRow = PRepDT.NewRow();

                        NewRow["UNN"] = UNN;
                        NewRow["AccountingName"] = PDT.Rows[g]["AccountingName"].ToString();
                        NewRow["InvNumber"] = PDT.Rows[g]["InvNumber"].ToString();
                        NewRow["MarketingCost"] = Convert.ToDecimal(PDT.Rows[g]["MarketingCost"]);
                        NewRow["Cvet"] = PDT.Rows[g]["Cvet"].ToString();
                        NewRow["Patina"] = PDT.Rows[g]["Patina"].ToString();
                        NewRow["CurrencyCode"] = ProfilCurrencyCode;
                        NewRow["Measure"] = "шт.";
                        NewRow["Count"] = PDT.Rows[g]["Count"];
                        //NewRow["Price"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[g]["Price"]), 3, MidpointRounding.AwayFromZero);
                        NewRow["Cost"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[g]["Cost"]), 3, MidpointRounding.AwayFromZero);
                        NewRow["Weight"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[g]["Weight"]), 3, MidpointRounding.AwayFromZero);

                        PRepDT.Rows.Add(NewRow);
                    }

                    if (p1.Length == 0 && p2.Length == 0 && p3.Length == 0)
                    {
                        DataRow NewRow = PRepDT.NewRow();

                        NewRow["UNN"] = UNN;
                        NewRow["AccountingName"] = PDT.Rows[g]["AccountingName"].ToString();
                        NewRow["InvNumber"] = PDT.Rows[g]["InvNumber"].ToString();
                        NewRow["MarketingCost"] = Convert.ToDecimal(PDT.Rows[g]["MarketingCost"]);
                        NewRow["Cvet"] = PDT.Rows[g]["Cvet"].ToString();
                        NewRow["Patina"] = PDT.Rows[g]["Patina"].ToString();
                        NewRow["CurrencyCode"] = ProfilCurrencyCode;
                        NewRow["Measure"] = "шт.";
                        NewRow["Count"] = PDT.Rows[g]["Count"];
                        //NewRow["Price"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[g]["Price"]), 3, MidpointRounding.AwayFromZero);
                        NewRow["Cost"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[g]["Cost"]), 3, MidpointRounding.AwayFromZero);
                        NewRow["Weight"] = Decimal.Round(Convert.ToDecimal(PDT.Rows[g]["Weight"]), 3, MidpointRounding.AwayFromZero);

                        PRepDT.Rows.Add(NewRow);
                    }
                }
            }

            if (TDT.Rows.Count > 0)
            {
                for (int g = 0; g < TDT.Rows.Count; g++)
                {
                    if (p1.Length > 0 && p2.Length > 0)
                    {
                        DataRow NewRow = TRepDT.NewRow();

                        NewRow["UNN"] = UNN;
                        NewRow["AccountingName"] = TDT.Rows[g]["AccountingName"].ToString();
                        NewRow["InvNumber"] = TDT.Rows[g]["InvNumber"].ToString();
                        NewRow["MarketingCost"] = Convert.ToDecimal(TDT.Rows[g]["MarketingCost"]);
                        NewRow["Cvet"] = TDT.Rows[g]["Cvet"].ToString();
                        NewRow["Patina"] = TDT.Rows[g]["Patina"].ToString();
                        NewRow["CurrencyCode"] = ProfilCurrencyCode;
                        NewRow["Measure"] = "шт.";
                        NewRow["TPSCurCode"] = TPSCurrencyCode;
                        NewRow["Count"] = TDT.Rows[g]["Count"];
                        //NewRow["Price"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[g]["Price"]), 3, MidpointRounding.AwayFromZero);
                        NewRow["Cost"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[g]["Cost"]), 3, MidpointRounding.AwayFromZero);
                        NewRow["Weight"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[g]["Weight"]), 3, MidpointRounding.AwayFromZero);

                        TRepDT.Rows.Add(NewRow);
                    }

                    if (p3.Length > 0)
                    {
                        DataRow NewRow = TRepDT.NewRow();

                        NewRow["UNN"] = UNN;
                        NewRow["AccountingName"] = TDT.Rows[g]["AccountingName"].ToString();
                        NewRow["InvNumber"] = TDT.Rows[g]["InvNumber"].ToString();
                        NewRow["MarketingCost"] = Convert.ToDecimal(TDT.Rows[g]["MarketingCost"]);
                        NewRow["Cvet"] = TDT.Rows[g]["Cvet"].ToString();
                        NewRow["Patina"] = TDT.Rows[g]["Patina"].ToString();
                        NewRow["CurrencyCode"] = ProfilCurrencyCode;
                        NewRow["Measure"] = "шт.";
                        NewRow["TPSCurCode"] = TPSCurrencyCode;
                        NewRow["Count"] = TDT.Rows[g]["Count"];
                        //NewRow["Price"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[g]["Price"]), 3, MidpointRounding.AwayFromZero);
                        NewRow["Cost"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[g]["Cost"]), 3, MidpointRounding.AwayFromZero);
                        NewRow["Weight"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[g]["Weight"]), 3, MidpointRounding.AwayFromZero);

                        TRepDT.Rows.Add(NewRow);
                    }

                    if (p1.Length == 0 && p2.Length == 0 && p3.Length == 0)
                    {
                        DataRow NewRow = TRepDT.NewRow();

                        NewRow["UNN"] = UNN;
                        NewRow["AccountingName"] = TDT.Rows[g]["AccountingName"].ToString();
                        NewRow["InvNumber"] = TDT.Rows[g]["InvNumber"].ToString();
                        NewRow["MarketingCost"] = Convert.ToDecimal(TDT.Rows[g]["MarketingCost"]);
                        NewRow["Cvet"] = TDT.Rows[g]["Cvet"].ToString();
                        NewRow["Patina"] = TDT.Rows[g]["Patina"].ToString();
                        NewRow["CurrencyCode"] = ProfilCurrencyCode;
                        NewRow["Measure"] = "шт.";
                        NewRow["TPSCurCode"] = TPSCurrencyCode;
                        NewRow["Count"] = TDT.Rows[g]["Count"];
                        //NewRow["Price"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[g]["Price"]), 3, MidpointRounding.AwayFromZero);
                        NewRow["Cost"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[g]["Cost"]), 3, MidpointRounding.AwayFromZero);
                        NewRow["Weight"] = Decimal.Round(Convert.ToDecimal(TDT.Rows[g]["Weight"]), 3, MidpointRounding.AwayFromZero);

                        TRepDT.Rows.Add(NewRow);
                    }
                }
            }
        }

        private void Collect()
        {
            DataTable Items = new DataTable();

            using (DataView DV = new DataView(DecorOrdersDataTable))
            {
                Items = DV.ToTable(true, new string[] { "InvNumber", "ColorID", "PatinaID" });
            }

            for (int i = 0; i < Items.Rows.Count; i++)
            {
                int ColorID = Convert.ToInt32(Items.Rows[i]["ColorID"]);
                int PatinaID = Convert.ToInt32(Items.Rows[i]["PatinaID"]);
                string InvNumber = Items.Rows[i]["InvNumber"].ToString();
                DataRow[] ItemsRows = DecorOrdersDataTable.Select("InvNumber = '" + InvNumber +
                                                                  "' AND ColorID=" + ColorID + " AND PatinaID=" + PatinaID,
                                                                      "Price ASC");

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

            Items.Dispose();
        }

        private void Collect(DataTable PRepDT, DataTable TRepDT)
        {
            string InvNumber = string.Empty;

            DataTable Items = new DataTable();

            using (DataView DV = new DataView(DecorOrdersDataTable))
            {
                Items = DV.ToTable(true, new string[] { "DecorID" });
            }

            for (int i = 0; i < Items.Rows.Count; i++)
            {
                int rr = Convert.ToInt32(Items.Rows[i]["DecorID"]);
                DataRow[] ItemsRows = DecorOrdersDataTable.Select("DecorID = " + Items.Rows[i]["DecorID"].ToString(),
                                                                      "Price ASC");

                int DecorConfigID = Convert.ToInt32(ItemsRows[0]["DecorConfigID"]);
                //м.п.
                if (GetReportMeasureTypeID(Convert.ToInt32(ItemsRows[0]["DecorConfigID"])) == 2)
                {
                    GroupCoverTypes(ItemsRows, 2, PRepDT, TRepDT);
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

                    GetParametrizedData(ItemsRows, ParamTableProfil, ParamTableTPS, PRepDT, TRepDT);

                    ParamTableProfil.Dispose();
                    ParamTableTPS.Dispose();
                }

                //м.кв.
                if (GetReportMeasureTypeID(Convert.ToInt32(ItemsRows[0]["DecorConfigID"])) == 1)
                {
                    GroupCoverTypes(ItemsRows, 1, PRepDT, TRepDT);
                }
            }

            Items.Dispose();
        }

        public void ExpReport(DateTime date1, DateTime date2, bool ZOV, bool IsSample, bool IsNotSample, ArrayList MClients, ArrayList MClientGroups)
        {
            string MClientFilter = string.Empty;
            if (MClients.Count > 0)
            {
                MClientFilter = " WHERE ClientID IN (" + string.Join(",", MClients.OfType<Int32>().ToArray()) + ")";
            }
            if (MClientGroups.Count > 0)
            {
                MClientFilter = " WHERE ClientID IN" +
                    " (SELECT ClientID FROM infiniu2_marketingreference.dbo.Clients" +
                    " WHERE ClientGroupID IN (" + string.Join(",", MClientGroups.OfType<Int32>().ToArray()) + "))";
            }
            if (MClients.Count < 1 && MClientGroups.Count < 1)
                MClientFilter = " WHERE ClientID = -1";

            string ConnectionString = ConnectionStrings.MarketingOrdersConnectionString;
            if (ZOV)
                ConnectionString = ConnectionStrings.ZOVOrdersConnectionString;
            string Filter = " (ExpeditionDateTime >= '" + date1.ToString("yyyy-MM-dd") +
                " 00:00' AND ExpeditionDateTime <= '" + date2.ToString("yyyy-MM-dd") + " 23:59')";

            string MFSampleFilter = string.Empty;
            string MDSampleFilter = string.Empty;
            if (!IsSample || !IsNotSample)
            {
                if (IsSample)
                    MFSampleFilter = " AND FrontsOrders.IsSample = 1";
                if (IsNotSample)
                    MFSampleFilter = " AND FrontsOrders.IsSample = 0";
                if (IsSample)
                    MDSampleFilter = " AND (DecorOrders.IsSample = 1 OR (DecorOrders.ProductID=42 AND DecorOrders.IsSample = 0)) ";
                if (IsNotSample)
                    MDSampleFilter = " AND DecorOrders.IsSample = 0";
            }

            string SelectCommand = @"SELECT dbo.DecorOrders.DecorOrderID, dbo.DecorOrders.ProductID, dbo.DecorOrders.DecorID, dbo.DecorOrders.ColorID, dbo.DecorOrders.PatinaID, dbo.DecorOrders.Length,
                dbo.DecorOrders.Height, dbo.DecorOrders.Width, dbo.PackageDetails.Count, dbo.DecorOrders.DecorConfigID, dbo.DecorOrders.FactoryID, dbo.DecorOrders.Price, dbo.DecorOrders.Cost,
                dbo.DecorOrders.ItemWeight, dbo.DecorOrders.Weight, infiniu2_catalog.dbo.DecorConfig.AccountingName, infiniu2_catalog.dbo.DecorConfig.InvNumber, infiniu2_catalog.dbo.Measures.Measure FROM PackageDetails
                INNER JOIN DecorOrders ON PackageDetails.OrderID = DecorOrders.DecorOrderID" + MDSampleFilter + @" AND DecorOrders.MainOrderID IN (SELECT MainOrderID FROM MainOrders WHERE MegaOrderID IN (SELECT MegaOrderID FROM MegaOrders " + MClientFilter + @"))
                INNER JOIN infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID = infiniu2_catalog.dbo.DecorConfig.DecorConfigID
                INNER JOIN infiniu2_catalog.dbo.Measures ON infiniu2_catalog.dbo.DecorConfig.MeasureID=infiniu2_catalog.dbo.Measures.MeasureID
                WHERE InvNumber IS NOT NULL AND PackageID IN (SELECT PackageID FROM Packages WHERE ProductType = 1 AND " + Filter + ") ORDER BY InvNumber";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionString))
            {
                DA.Fill(DecorOrdersDataTable);
            }

            Collect();
        }
        
        public void StoreReport(DateTime date1, DateTime date2, bool ZOV, bool IsSample, bool IsNotSample, ArrayList MClients, ArrayList MClientGroups)
        {
            string MClientFilter = string.Empty;
            if (MClients.Count > 0)
            {
                MClientFilter = " WHERE ClientID IN (" + string.Join(",", MClients.OfType<Int32>().ToArray()) + ")";
            }
            if (MClientGroups.Count > 0)
            {
                MClientFilter = " WHERE ClientID IN" +
                    " (SELECT ClientID FROM infiniu2_marketingreference.dbo.Clients" +
                    " WHERE ClientGroupID IN (" + string.Join(",", MClientGroups.OfType<Int32>().ToArray()) + "))";
            }
            if (MClients.Count < 1 && MClientGroups.Count < 1)
                MClientFilter = " WHERE ClientID = -1";

            string ConnectionString = ConnectionStrings.MarketingOrdersConnectionString;
            if (ZOV)
                ConnectionString = ConnectionStrings.ZOVOrdersConnectionString;
            string Filter = " (StorageDateTime >= '" + date1.ToString("yyyy-MM-dd") +
                " 00:00' AND StorageDateTime <= '" + date2.ToString("yyyy-MM-dd") + " 23:59')";

            string MFSampleFilter = string.Empty;
            string MDSampleFilter = string.Empty;
            if (!IsSample || !IsNotSample)
            {
                if (IsSample)
                    MFSampleFilter = " AND FrontsOrders.IsSample = 1";
                if (IsNotSample)
                    MFSampleFilter = " AND FrontsOrders.IsSample = 0";
                if (IsSample)
                    MDSampleFilter = " AND (DecorOrders.IsSample = 1 OR (DecorOrders.ProductID=42 AND DecorOrders.IsSample = 0)) ";
                if (IsNotSample)
                    MDSampleFilter = " AND DecorOrders.IsSample = 0";
            }

            string SelectCommand = @"SELECT dbo.DecorOrders.DecorOrderID, dbo.DecorOrders.ProductID, dbo.DecorOrders.DecorID, dbo.DecorOrders.ColorID, dbo.DecorOrders.PatinaID, dbo.DecorOrders.Length,
                dbo.DecorOrders.Height, dbo.DecorOrders.Width, dbo.PackageDetails.Count, dbo.DecorOrders.DecorConfigID, dbo.DecorOrders.FactoryID, dbo.DecorOrders.Price, dbo.DecorOrders.Cost,
                dbo.DecorOrders.ItemWeight, dbo.DecorOrders.Weight, infiniu2_catalog.dbo.DecorConfig.AccountingName, infiniu2_catalog.dbo.DecorConfig.InvNumber, infiniu2_catalog.dbo.Measures.Measure FROM PackageDetails
                INNER JOIN DecorOrders ON PackageDetails.OrderID = DecorOrders.DecorOrderID" + MDSampleFilter + @" AND DecorOrders.MainOrderID IN (SELECT MainOrderID FROM MainOrders WHERE MegaOrderID IN (SELECT MegaOrderID FROM MegaOrders " + MClientFilter + @"))
                INNER JOIN infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID = infiniu2_catalog.dbo.DecorConfig.DecorConfigID
                INNER JOIN infiniu2_catalog.dbo.Measures ON infiniu2_catalog.dbo.DecorConfig.MeasureID=infiniu2_catalog.dbo.Measures.MeasureID
                WHERE InvNumber IS NOT NULL AND PackageID IN (SELECT PackageID FROM Packages WHERE ProductType = 1 AND " + Filter + ") ORDER BY InvNumber";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionString))
            {
                DA.Fill(DecorOrdersDataTable);
            }

            Collect();
        }
        
        public void DispReport(DateTime date1, DateTime date2, bool ZOV, bool IsSample, bool IsNotSample, ArrayList MClients, ArrayList MClientGroups)
        {
            string MClientFilter = string.Empty;
            if (MClients.Count > 0)
            {
                MClientFilter = " WHERE ClientID IN (" + string.Join(",", MClients.OfType<Int32>().ToArray()) + ")";
            }
            if (MClientGroups.Count > 0)
            {
                MClientFilter = " WHERE ClientID IN" +
                    " (SELECT ClientID FROM infiniu2_marketingreference.dbo.Clients" +
                    " WHERE ClientGroupID IN (" + string.Join(",", MClientGroups.OfType<Int32>().ToArray()) + "))";
            }
            if (MClients.Count < 1 && MClientGroups.Count < 1)
                MClientFilter = " WHERE ClientID = -1";

            string ConnectionString = ConnectionStrings.MarketingOrdersConnectionString;
            if (ZOV)
                ConnectionString = ConnectionStrings.ZOVOrdersConnectionString;
            string Filter = " (DispatchDateTime >= '" + date1.ToString("yyyy-MM-dd") +
                " 00:00' AND DispatchDateTime <= '" + date2.ToString("yyyy-MM-dd") + " 23:59')";

            string MFSampleFilter = string.Empty;
            string MDSampleFilter = string.Empty;
            if (!IsSample || !IsNotSample)
            {
                if (IsSample)
                    MFSampleFilter = " AND FrontsOrders.IsSample = 1";
                if (IsNotSample)
                    MFSampleFilter = " AND FrontsOrders.IsSample = 0";
                if (IsSample)
                    MDSampleFilter = " AND (DecorOrders.IsSample = 1 OR (DecorOrders.ProductID=42 AND DecorOrders.IsSample = 0)) ";
                if (IsNotSample)
                    MDSampleFilter = " AND DecorOrders.IsSample = 0";
            }

            string SelectCommand = @"SELECT dbo.DecorOrders.DecorOrderID, dbo.DecorOrders.ProductID, dbo.DecorOrders.DecorID, dbo.DecorOrders.ColorID, dbo.DecorOrders.PatinaID, dbo.DecorOrders.Length,
                dbo.DecorOrders.Height, dbo.DecorOrders.Width, dbo.PackageDetails.Count, dbo.DecorOrders.DecorConfigID, dbo.DecorOrders.FactoryID, dbo.DecorOrders.Price, dbo.DecorOrders.Cost,
                dbo.DecorOrders.ItemWeight, dbo.DecorOrders.Weight, infiniu2_catalog.dbo.DecorConfig.AccountingName, infiniu2_catalog.dbo.DecorConfig.InvNumber, infiniu2_catalog.dbo.Measures.Measure FROM PackageDetails
                INNER JOIN DecorOrders ON PackageDetails.OrderID = DecorOrders.DecorOrderID" + MDSampleFilter + @" AND DecorOrders.MainOrderID IN (SELECT MainOrderID FROM MainOrders WHERE MegaOrderID IN (SELECT MegaOrderID FROM MegaOrders " + MClientFilter + @"))
                INNER JOIN infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID = infiniu2_catalog.dbo.DecorConfig.DecorConfigID
                INNER JOIN infiniu2_catalog.dbo.Measures ON infiniu2_catalog.dbo.DecorConfig.MeasureID=infiniu2_catalog.dbo.Measures.MeasureID
                WHERE InvNumber IS NOT NULL AND PackageID IN (SELECT PackageID FROM Packages WHERE ProductType = 1 AND " + Filter + ") ORDER BY InvNumber";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionString))
            {
                DA.Fill(DecorOrdersDataTable);
            }

            Collect();
        }

        public void OnProdReport(DateTime date1, DateTime date2, bool ZOV, bool IsSample, bool IsNotSample, ArrayList MClients, ArrayList MClientGroups)
        {
            string MClientFilter = string.Empty;

            string MMainOrdersFilter = " WHERE (CAST(ProfilOnProductionDate AS Date) >= '" + date1.ToString("yyyy-MM-dd") + "' AND CAST(ProfilOnProductionDate AS Date) <= '" + date2.ToString("yyyy-MM-dd") + "')";

            if (MClients.Count > 0)
            {
                if (MClientFilter.Length > 0)
                    MClientFilter += " AND ClientID IN (" + string.Join(",", MClients.OfType<Int32>().ToArray()) + ")";
                else
                    MClientFilter = " WHERE ClientID IN (" + string.Join(",", MClients.OfType<Int32>().ToArray()) + ")";
            }
            if (MClientGroups.Count > 0)
            {
                if (MClientFilter.Length > 0)
                    MClientFilter += " AND ClientID IN" +
                                     " (SELECT ClientID FROM infiniu2_marketingreference.dbo.Clients" +
                                     " WHERE ClientGroupID IN (" + string.Join(",", MClientGroups.OfType<Int32>().ToArray()) + "))";
                else
                    MClientFilter = " WHERE ClientID IN" +
                                    " (SELECT ClientID FROM infiniu2_marketingreference.dbo.Clients" +
                                    " WHERE ClientGroupID IN (" + string.Join(",", MClientGroups.OfType<Int32>().ToArray()) + "))";
            }
            if (MClients.Count < 1 && MClientGroups.Count < 1)
                MClientFilter = " WHERE ClientID = -1";

            string ConnectionString = ConnectionStrings.MarketingOrdersConnectionString;
            if (ZOV)
                ConnectionString = ConnectionStrings.ZOVOrdersConnectionString;
            
            string MDSampleFilter = string.Empty;
            if (!IsSample || !IsNotSample)
            {
                if (IsSample)
                    MDSampleFilter = " AND (DecorOrders.IsSample = 1 OR (DecorOrders.ProductID=42 AND DecorOrders.IsSample = 0)) ";
                if (IsNotSample)
                    MDSampleFilter = " AND DecorOrders.IsSample = 0";
            }

            string SelectCommand = @"SELECT dbo.DecorOrders.DecorOrderID, dbo.DecorOrders.ProductID, dbo.DecorOrders.DecorID, dbo.DecorOrders.ColorID, dbo.DecorOrders.PatinaID, dbo.DecorOrders.Length,
                dbo.DecorOrders.Height, dbo.DecorOrders.Width, dbo.DecorOrders.Count, dbo.DecorOrders.DecorConfigID, dbo.DecorOrders.FactoryID, dbo.DecorOrders.Price, dbo.DecorOrders.Cost,
                dbo.DecorOrders.ItemWeight, dbo.DecorOrders.Weight, infiniu2_catalog.dbo.DecorConfig.AccountingName, infiniu2_catalog.dbo.DecorConfig.InvNumber, infiniu2_catalog.dbo.Measures.Measure FROM DecorOrders
                INNER JOIN infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID = infiniu2_catalog.dbo.DecorConfig.DecorConfigID
                INNER JOIN infiniu2_catalog.dbo.Measures ON infiniu2_catalog.dbo.DecorConfig.MeasureID=infiniu2_catalog.dbo.Measures.MeasureID
                WHERE DecorOrders.MainOrderID IN (SELECT MainOrderID FROM MainOrders " + MMainOrdersFilter + @" AND MegaOrderID IN (SELECT MegaOrderID FROM MegaOrders " + MClientFilter + @")) AND InvNumber IS NOT NULL " + MDSampleFilter + " ORDER BY InvNumber";

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionString))
            {
                DA.Fill(DecorOrdersDataTable);
            }

            Collect();
        }
        
        public void PlanDispReport(DateTime date1, DateTime date2, bool ZOV, bool IsSample, bool IsNotSample, ArrayList MClients, ArrayList MClientGroups)
        {
            string MClientFilter = " WHERE (CAST(ProfilDispatchDate AS Date) >= '" + date1.ToString("yyyy-MM-dd") + "' AND CAST(ProfilDispatchDate AS Date) <= '" + date2.ToString("yyyy-MM-dd") + "')";

            if (MClients.Count > 0)
            {
                if (MClientFilter.Length > 0)
                    MClientFilter += " AND ClientID IN (" + string.Join(",", MClients.OfType<Int32>().ToArray()) + ")";
                else
                    MClientFilter = " WHERE ClientID IN (" + string.Join(",", MClients.OfType<Int32>().ToArray()) + ")";
            }
            if (MClientGroups.Count > 0)
            {
                if (MClientFilter.Length > 0)
                    MClientFilter += " AND ClientID IN" +
                                     " (SELECT ClientID FROM infiniu2_marketingreference.dbo.Clients" +
                                     " WHERE ClientGroupID IN (" + string.Join(",", MClientGroups.OfType<Int32>().ToArray()) + "))";
                else
                    MClientFilter = " WHERE ClientID IN" +
                                    " (SELECT ClientID FROM infiniu2_marketingreference.dbo.Clients" +
                                    " WHERE ClientGroupID IN (" + string.Join(",", MClientGroups.OfType<Int32>().ToArray()) + "))";
            }
            if (MClients.Count < 1 && MClientGroups.Count < 1)
                MClientFilter = " WHERE ClientID = -1";

            string ConnectionString = ConnectionStrings.MarketingOrdersConnectionString;
            if (ZOV)
                ConnectionString = ConnectionStrings.ZOVOrdersConnectionString;
            
            string MDSampleFilter = string.Empty;
            if (!IsSample || !IsNotSample)
            {
                if (IsSample)
                    MDSampleFilter = " AND (DecorOrders.IsSample = 1 OR (DecorOrders.ProductID=42 AND DecorOrders.IsSample = 0)) ";
                if (IsNotSample)
                    MDSampleFilter = " AND DecorOrders.IsSample = 0";
            }

            string SelectCommand = @"SELECT dbo.DecorOrders.DecorOrderID, dbo.DecorOrders.ProductID, dbo.DecorOrders.DecorID, dbo.DecorOrders.ColorID, dbo.DecorOrders.PatinaID, dbo.DecorOrders.Length,
                dbo.DecorOrders.Height, dbo.DecorOrders.Width, dbo.DecorOrders.Count, dbo.DecorOrders.DecorConfigID, dbo.DecorOrders.FactoryID, dbo.DecorOrders.Price, dbo.DecorOrders.Cost,
                dbo.DecorOrders.ItemWeight, dbo.DecorOrders.Weight, infiniu2_catalog.dbo.DecorConfig.AccountingName, infiniu2_catalog.dbo.DecorConfig.InvNumber, infiniu2_catalog.dbo.Measures.Measure FROM DecorOrders
                INNER JOIN infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID = infiniu2_catalog.dbo.DecorConfig.DecorConfigID
                INNER JOIN infiniu2_catalog.dbo.Measures ON infiniu2_catalog.dbo.DecorConfig.MeasureID=infiniu2_catalog.dbo.Measures.MeasureID
                WHERE DecorOrders.MainOrderID IN (SELECT MainOrderID FROM MainOrders WHERE MegaOrderID IN (SELECT MegaOrderID FROM MegaOrders " + MClientFilter + @")) AND InvNumber IS NOT NULL " + MDSampleFilter + " ORDER BY InvNumber";

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionString))
            {
                DA.Fill(DecorOrdersDataTable);
            }

            Collect();
        }
        
        public void PackReport(DateTime date1, DateTime date2, bool ZOV, bool IsSample, bool IsNotSample, ArrayList MClients, ArrayList MClientGroups)
        {
            string MClientFilter = string.Empty;
            if (MClients.Count > 0)
            {
                MClientFilter = " WHERE ClientID IN (" + string.Join(",", MClients.OfType<Int32>().ToArray()) + ")";
            }
            if (MClientGroups.Count > 0)
            {
                MClientFilter = " WHERE ClientID IN" +
                    " (SELECT ClientID FROM infiniu2_marketingreference.dbo.Clients" +
                    " WHERE ClientGroupID IN (" + string.Join(",", MClientGroups.OfType<Int32>().ToArray()) + "))";
            }
            if (MClients.Count < 1 && MClientGroups.Count < 1)
                MClientFilter = " WHERE ClientID = -1";

            string ConnectionString = ConnectionStrings.MarketingOrdersConnectionString;
            if (ZOV)
                ConnectionString = ConnectionStrings.ZOVOrdersConnectionString;
            string Filter = " (PackingDateTime >= '" + date1.ToString("yyyy-MM-dd") +
                " 00:00' AND PackingDateTime <= '" + date2.ToString("yyyy-MM-dd") + " 23:59')";

            string MFSampleFilter = string.Empty;
            string MDSampleFilter = string.Empty;
            if (!IsSample || !IsNotSample)
            {
                if (IsSample)
                    MFSampleFilter = " AND FrontsOrders.IsSample = 1";
                if (IsNotSample)
                    MFSampleFilter = " AND FrontsOrders.IsSample = 0";
                if (IsSample)
                    MDSampleFilter = " AND (DecorOrders.IsSample = 1 OR (DecorOrders.ProductID=42 AND DecorOrders.IsSample = 0)) ";
                if (IsNotSample)
                    MDSampleFilter = " AND DecorOrders.IsSample = 0";
            }

            string SelectCommand = @"SELECT dbo.DecorOrders.DecorOrderID, dbo.DecorOrders.ProductID, dbo.DecorOrders.DecorID, dbo.DecorOrders.ColorID, dbo.DecorOrders.PatinaID, dbo.DecorOrders.Length,
                dbo.DecorOrders.Height, dbo.DecorOrders.Width, dbo.PackageDetails.Count, dbo.DecorOrders.DecorConfigID, dbo.DecorOrders.FactoryID, dbo.DecorOrders.Price, dbo.DecorOrders.Cost,
                dbo.DecorOrders.ItemWeight, dbo.DecorOrders.Weight, infiniu2_catalog.dbo.DecorConfig.AccountingName, infiniu2_catalog.dbo.DecorConfig.InvNumber, infiniu2_catalog.dbo.Measures.Measure FROM PackageDetails
                INNER JOIN DecorOrders ON PackageDetails.OrderID = DecorOrders.DecorOrderID" + MDSampleFilter + @" AND DecorOrders.MainOrderID IN (SELECT MainOrderID FROM MainOrders WHERE MegaOrderID IN (SELECT MegaOrderID FROM MegaOrders " + MClientFilter + @"))
                INNER JOIN infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID = infiniu2_catalog.dbo.DecorConfig.DecorConfigID
                INNER JOIN infiniu2_catalog.dbo.Measures ON infiniu2_catalog.dbo.DecorConfig.MeasureID=infiniu2_catalog.dbo.Measures.MeasureID
                WHERE InvNumber IS NOT NULL AND PackageID IN (SELECT PackageID FROM Packages WHERE ProductType = 1 AND " + Filter + ") ORDER BY InvNumber";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionString))
            {
                DA.Fill(DecorOrdersDataTable);
            }

            Collect();
        }

        public void MarketingProduced(DateTime date1, DateTime date2, bool IsSample, bool IsNotSample, ArrayList MClients, ArrayList MClientGroups)
        {
            string MClientFilter = string.Empty;
            if (MClients.Count > 0)
            {
                MClientFilter = " WHERE ClientID IN (" + string.Join(",", MClients.OfType<Int32>().ToArray()) + ")";
            }
            if (MClientGroups.Count > 0)
            {
                MClientFilter = " WHERE ClientID IN" +
                    " (SELECT ClientID FROM infiniu2_marketingreference.dbo.Clients" +
                    " WHERE ClientGroupID IN (" + string.Join(",", MClientGroups.OfType<Int32>().ToArray()) + "))";
            }
            if (MClients.Count < 1 && MClientGroups.Count < 1)
                MClientFilter = " WHERE ClientID = -1";

            ProfilMarketingProducedDT.Clear();
            TPSMarketingProducedDT.Clear();
            DecorOrdersDataTable.Clear();

            string Filter = " (PackingDateTime >= '" + date1.ToString("yyyy-MM-dd") +
                " 00:00' AND PackingDateTime <= '" + date2.ToString("yyyy-MM-dd") + " 23:59')";

            string MFSampleFilter = string.Empty;
            string MDSampleFilter = string.Empty;
            if (!IsSample || !IsNotSample)
            {
                if (IsSample)
                    MFSampleFilter = " AND FrontsOrders.IsSample = 1";
                if (IsNotSample)
                    MFSampleFilter = " AND FrontsOrders.IsSample = 0";
                if (IsSample)
                    MDSampleFilter = " AND (DecorOrders.IsSample = 1 OR (DecorOrders.ProductID=42 AND DecorOrders.IsSample = 0)) ";
                if (IsNotSample)
                    MDSampleFilter = " AND DecorOrders.IsSample = 0";
            }

            string SelectCommand = @"SELECT dbo.DecorOrders.DecorOrderID, dbo.DecorOrders.ProductID, dbo.DecorOrders.DecorID, dbo.DecorOrders.ColorID, dbo.DecorOrders.PatinaID, dbo.DecorOrders.Length,
                dbo.DecorOrders.Height, dbo.DecorOrders.Width, dbo.PackageDetails.Count, dbo.DecorOrders.DecorConfigID, dbo.DecorOrders.FactoryID, dbo.DecorOrders.Price, dbo.DecorOrders.Cost,
                dbo.DecorOrders.ItemWeight, dbo.DecorOrders.Weight, infiniu2_catalog.dbo.DecorConfig.AccountingName, infiniu2_catalog.dbo.DecorConfig.InvNumber, infiniu2_catalog.dbo.Measures.Measure FROM PackageDetails
                INNER JOIN DecorOrders ON PackageDetails.OrderID = DecorOrders.DecorOrderID" + MDSampleFilter + @" AND DecorOrders.MainOrderID IN (SELECT MainOrderID FROM MainOrders WHERE MegaOrderID IN (SELECT MegaOrderID FROM MegaOrders " + MClientFilter + @"))
                INNER JOIN infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID = infiniu2_catalog.dbo.DecorConfig.DecorConfigID
                INNER JOIN infiniu2_catalog.dbo.Measures ON infiniu2_catalog.dbo.DecorConfig.MeasureID=infiniu2_catalog.dbo.Measures.MeasureID
                WHERE InvNumber IS NOT NULL AND PackageID IN (SELECT PackageID FROM Packages WHERE ProductType = 1 AND " + Filter + ") ORDER BY InvNumber";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(DecorOrdersDataTable);
            }

            Collect(ProfilMarketingProducedDT, TPSMarketingProducedDT);
        }

        public void ZOVProduced(DateTime date1, DateTime date2, bool IsSample, bool IsNotSample)
        {
            ProfilZOVProducedDT.Clear();
            TPSZOVProducedDT.Clear();
            DecorOrdersDataTable.Clear();

            string Filter = " (PackingDateTime >= '" + date1.ToString("yyyy-MM-dd") +
                " 00:00' AND PackingDateTime <= '" + date2.ToString("yyyy-MM-dd") + " 23:59')";

            string MFSampleFilter = string.Empty;
            string MDSampleFilter = string.Empty;
            if (!IsSample || !IsNotSample)
            {
                if (IsSample)
                    MFSampleFilter = " AND FrontsOrders.IsSample = 1";
                if (IsNotSample)
                    MFSampleFilter = " AND FrontsOrders.IsSample = 0";
                if (IsSample)
                    MDSampleFilter = " AND (DecorOrders.IsSample = 1 OR (DecorOrders.ProductID=42 AND DecorOrders.IsSample = 0)) ";
                if (IsNotSample)
                    MDSampleFilter = " AND DecorOrders.IsSample = 0";
            }

            string SelectCommand = @"SELECT dbo.DecorOrders.DecorOrderID, dbo.DecorOrders.ProductID, dbo.DecorOrders.DecorID, dbo.DecorOrders.ColorID, dbo.DecorOrders.PatinaID, dbo.DecorOrders.Length,
                dbo.DecorOrders.Height, dbo.DecorOrders.Width, dbo.PackageDetails.Count, dbo.DecorOrders.DecorConfigID, dbo.DecorOrders.FactoryID, dbo.DecorOrders.Price, dbo.DecorOrders.Cost,
                dbo.DecorOrders.ItemWeight, dbo.DecorOrders.Weight, infiniu2_catalog.dbo.DecorConfig.AccountingName, infiniu2_catalog.dbo.DecorConfig.InvNumber, infiniu2_catalog.dbo.Measures.Measure FROM PackageDetails
                INNER JOIN DecorOrders ON PackageDetails.OrderID = DecorOrders.DecorOrderID" + MDSampleFilter + @" AND DecorOrders.MainOrderID IN (SELECT MainOrderID FROM MainOrders WHERE MegaOrderID IN (SELECT MegaOrderID FROM MegaOrders))
                INNER JOIN infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID = infiniu2_catalog.dbo.DecorConfig.DecorConfigID
                INNER JOIN infiniu2_catalog.dbo.Measures ON infiniu2_catalog.dbo.DecorConfig.MeasureID=infiniu2_catalog.dbo.Measures.MeasureID
                WHERE InvNumber IS NOT NULL AND PackageID IN (SELECT PackageID FROM Packages WHERE ProductType = 1 AND " + Filter + ") ORDER BY InvNumber";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.ZOVOrdersConnectionString))
            {
                DA.Fill(DecorOrdersDataTable);
            }

            Collect(ProfilZOVProducedDT, TPSZOVProducedDT);
        }

        public void MarketingDispatch(DateTime date1, DateTime date2, bool IsSample, bool IsNotSample, ArrayList MClients, ArrayList MClientGroups)
        {
            string MClientFilter = string.Empty;
            if (MClients.Count > 0)
            {
                MClientFilter = " WHERE ClientID IN (" + string.Join(",", MClients.OfType<Int32>().ToArray()) + ")";
            }
            if (MClientGroups.Count > 0)
            {
                MClientFilter = " WHERE ClientID IN" +
                    " (SELECT ClientID FROM infiniu2_marketingreference.dbo.Clients" +
                    " WHERE ClientGroupID IN (" + string.Join(",", MClientGroups.OfType<Int32>().ToArray()) + "))";
            }
            if (MClients.Count < 1 && MClientGroups.Count < 1)
                MClientFilter = " WHERE ClientID = -1";

            ProfilMarketingDispatchDT.Clear();
            TPSMarketingDispatchDT.Clear();
            DecorOrdersDataTable.Clear();

            string Filter = " (DispatchDateTime >= '" + date1.ToString("yyyy-MM-dd") +
                " 00:00' AND DispatchDateTime <= '" + date2.ToString("yyyy-MM-dd") + " 23:59')";

            string MFSampleFilter = string.Empty;
            string MDSampleFilter = string.Empty;
            if (!IsSample || !IsNotSample)
            {
                if (IsSample)
                    MFSampleFilter = " AND FrontsOrders.IsSample = 1";
                if (IsNotSample)
                    MFSampleFilter = " AND FrontsOrders.IsSample = 0";
                if (IsSample)
                    MDSampleFilter = " AND (DecorOrders.IsSample = 1 OR (DecorOrders.ProductID=42 AND DecorOrders.IsSample = 0)) ";
                if (IsNotSample)
                    MDSampleFilter = " AND DecorOrders.IsSample = 0";
            }

            string SelectCommand = @"SELECT dbo.DecorOrders.DecorOrderID, dbo.DecorOrders.ProductID, dbo.DecorOrders.DecorID, dbo.DecorOrders.ColorID, dbo.DecorOrders.PatinaID, dbo.DecorOrders.Length,
                dbo.DecorOrders.Height, dbo.DecorOrders.Width, dbo.PackageDetails.Count, dbo.DecorOrders.DecorConfigID, dbo.DecorOrders.FactoryID, dbo.DecorOrders.Price, dbo.DecorOrders.Cost,
                dbo.DecorOrders.ItemWeight, dbo.DecorOrders.Weight, infiniu2_catalog.dbo.DecorConfig.AccountingName, infiniu2_catalog.dbo.DecorConfig.InvNumber, infiniu2_catalog.dbo.Measures.Measure FROM PackageDetails
                INNER JOIN DecorOrders ON PackageDetails.OrderID = DecorOrders.DecorOrderID" + MDSampleFilter + @" AND DecorOrders.MainOrderID IN (SELECT MainOrderID FROM MainOrders WHERE MegaOrderID IN (SELECT MegaOrderID FROM MegaOrders " + MClientFilter + @"))
                INNER JOIN infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID = infiniu2_catalog.dbo.DecorConfig.DecorConfigID
                INNER JOIN infiniu2_catalog.dbo.Measures ON infiniu2_catalog.dbo.DecorConfig.MeasureID=infiniu2_catalog.dbo.Measures.MeasureID
                WHERE InvNumber IS NOT NULL AND PackageID IN (SELECT PackageID FROM Packages WHERE ProductType = 1 AND " + Filter + ") ORDER BY InvNumber";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(DecorOrdersDataTable);
            }

            Collect(ProfilMarketingDispatchDT, TPSMarketingDispatchDT);
        }

        public void ZOVDispatch(DateTime date1, DateTime date2, bool IsSample, bool IsNotSample)
        {
            ProfilZOVDispatchDT.Clear();
            TPSZOVDispatchDT.Clear();
            DecorOrdersDataTable.Clear();

            string Filter = " (DispatchDateTime >= '" + date1.ToString("yyyy-MM-dd") +
                " 00:00' AND DispatchDateTime <= '" + date2.ToString("yyyy-MM-dd") + " 23:59')";

            string MFSampleFilter = string.Empty;
            string MDSampleFilter = string.Empty;
            if (!IsSample || !IsNotSample)
            {
                if (IsSample)
                    MFSampleFilter = " AND FrontsOrders.IsSample = 1";
                if (IsNotSample)
                    MFSampleFilter = " AND FrontsOrders.IsSample = 0";
                if (IsSample)
                    MDSampleFilter = " AND (DecorOrders.IsSample = 1 OR (DecorOrders.ProductID=42 AND DecorOrders.IsSample = 0)) ";
                if (IsNotSample)
                    MDSampleFilter = " AND DecorOrders.IsSample = 0";
            }

            string SelectCommand = @"SELECT dbo.DecorOrders.DecorOrderID, dbo.DecorOrders.ProductID, dbo.DecorOrders.DecorID, dbo.DecorOrders.ColorID, dbo.DecorOrders.PatinaID, dbo.DecorOrders.Length,
                dbo.DecorOrders.Height, dbo.DecorOrders.Width, dbo.PackageDetails.Count, dbo.DecorOrders.DecorConfigID, dbo.DecorOrders.FactoryID, dbo.DecorOrders.Price, dbo.DecorOrders.Cost,
                dbo.DecorOrders.ItemWeight, dbo.DecorOrders.Weight, infiniu2_catalog.dbo.DecorConfig.AccountingName, infiniu2_catalog.dbo.DecorConfig.InvNumber, infiniu2_catalog.dbo.Measures.Measure FROM PackageDetails
                INNER JOIN DecorOrders ON PackageDetails.OrderID = DecorOrders.DecorOrderID" + MDSampleFilter + @" AND DecorOrders.MainOrderID IN (SELECT MainOrderID FROM MainOrders WHERE MegaOrderID IN (SELECT MegaOrderID FROM MegaOrders))
                INNER JOIN infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID = infiniu2_catalog.dbo.DecorConfig.DecorConfigID
                INNER JOIN infiniu2_catalog.dbo.Measures ON infiniu2_catalog.dbo.DecorConfig.MeasureID=infiniu2_catalog.dbo.Measures.MeasureID
                WHERE InvNumber IS NOT NULL AND PackageID IN (SELECT PackageID FROM Packages WHERE ProductType = 1 AND " + Filter + ") ORDER BY InvNumber";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.ZOVOrdersConnectionString))
            {
                DA.Fill(DecorOrdersDataTable);
            }

            Collect(ProfilZOVDispatchDT, TPSZOVDispatchDT);
        }

        public void MarketingStartCount(DateTime date1, bool IsSample, bool IsNotSample, ArrayList MClients, ArrayList MClientGroups)
        {
            string MClientFilter = string.Empty;
            if (MClients.Count > 0)
            {
                MClientFilter += " AND ClientID IN (" + string.Join(",", MClients.OfType<Int32>().ToArray()) + ")";
            }
            if (MClientGroups.Count > 0)
            {
                MClientFilter += " AND ClientID IN" +
                    " (SELECT ClientID FROM infiniu2_marketingreference.dbo.Clients" +
                    " WHERE ClientGroupID IN (" + string.Join(",", MClientGroups.OfType<Int32>().ToArray()) + "))";
            }
            if (MClients.Count < 1 && MClientGroups.Count < 1)
                MClientFilter = " AND ClientID = -1";

            ProfilMarketingStartCountDT.Clear();
            TPSMarketingStartCountDT.Clear();
            DecorOrdersDataTable.Clear();

            string MFSampleFilter = string.Empty;
            string MDSampleFilter = string.Empty;
            if (!IsSample || !IsNotSample)
            {
                if (IsSample)
                    MFSampleFilter = " AND FrontsOrders.IsSample = 1";
                if (IsNotSample)
                    MFSampleFilter = " AND FrontsOrders.IsSample = 0";
                if (IsSample)
                    MDSampleFilter = " AND (DecorOrders.IsSample = 1 OR (DecorOrders.ProductID=42 AND DecorOrders.IsSample = 0)) ";
                if (IsNotSample)
                    MDSampleFilter = " AND DecorOrders.IsSample = 0";
            }

            string SelectCommand = @"SELECT        DecorOrders.DecorOrderID, PackageDetails.PackageDetailID, DecorOrders.ProductID, DecorOrders.DecorID, DecorOrders.ColorID, DecorOrders.PatinaID, DecorOrders.Length, DecorOrders.Height,
                         DecorOrders.Width, PackageDetails.Count, DecorOrders.DecorConfigID, DecorOrders.FactoryID, DecorOrders.Price, DecorOrders.Cost, DecorOrders.ItemWeight, DecorOrders.Weight,
                         infiniu2_catalog.dbo.DecorConfig.AccountingName, infiniu2_catalog.dbo.DecorConfig.InvNumber, infiniu2_catalog.dbo.Measures.Measure, PackageDetails.OrderID, infiniu2_marketingreference.dbo.Clients.ClientName,
                         MegaOrders.OrderNumber
FROM            PackageDetails INNER JOIN
                         DecorOrders ON PackageDetails.OrderID = DecorOrders.DecorOrderID" + MDSampleFilter + @" LEFT OUTER JOIN
                         infiniu2_catalog.dbo.TechStore AS Fronts ON DecorOrders.DecorID = Fronts.TechStoreID LEFT OUTER JOIN
                         infiniu2_catalog.dbo.TechStore AS Colors ON DecorOrders.ColorID = Colors.TechStoreID INNER JOIN
                         infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID = infiniu2_catalog.dbo.DecorConfig.DecorConfigID INNER JOIN
                         MainOrders ON DecorOrders.MainOrderID = MainOrders.MainOrderID INNER JOIN
                         MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID " + MClientFilter + @" INNER JOIN
                         infiniu2_marketingreference.dbo.Clients ON MegaOrders.ClientID = infiniu2_marketingreference.dbo.Clients.ClientID INNER JOIN
                         infiniu2_catalog.dbo.Measures ON infiniu2_catalog.dbo.DecorConfig.MeasureID = infiniu2_catalog.dbo.Measures.MeasureID
WHERE dbo.PackageDetails.PackageID IN (SELECT TOP (100) PERCENT PackageID FROM dbo.Packages WHERE
PackingDateTime < '" + date1.ToString("yyyy-MM-dd") + @" 00:00:00' AND ProductType = 1 AND (DispatchDateTime IS NULL OR DispatchDateTime >= '" + date1.ToString("yyyy-MM-dd") + @" 00:00:00'))
ORDER BY infiniu2_marketingreference.dbo.Clients.ClientName, dbo.MegaOrders.OrderNumber";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(DecorOrdersDataTable);
            }

            Collect(ProfilMarketingStartCountDT, TPSMarketingStartCountDT);
        }

        public void MarketingEndCount(DateTime date1, bool IsSample, bool IsNotSample, ArrayList MClients, ArrayList MClientGroups)
        {
            string MClientFilter = string.Empty;
            if (MClients.Count > 0)
            {
                MClientFilter += " AND ClientID IN (" + string.Join(",", MClients.OfType<Int32>().ToArray()) + ")";
            }
            if (MClientGroups.Count > 0)
            {
                MClientFilter += " AND ClientID IN" +
                    " (SELECT ClientID FROM infiniu2_marketingreference.dbo.Clients" +
                    " WHERE ClientGroupID IN (" + string.Join(",", MClientGroups.OfType<Int32>().ToArray()) + "))";
            }
            if (MClients.Count < 1 && MClientGroups.Count < 1)
                MClientFilter = " AND ClientID = -1";

            ProfilMarketingEndCountDT.Clear();
            TPSMarketingEndCountDT.Clear();
            DecorOrdersDataTable.Clear();

            string MFSampleFilter = string.Empty;
            string MDSampleFilter = string.Empty;
            if (!IsSample || !IsNotSample)
            {
                if (IsSample)
                    MFSampleFilter = " AND FrontsOrders.IsSample = 1";
                if (IsNotSample)
                    MFSampleFilter = " AND FrontsOrders.IsSample = 0";
                if (IsSample)
                    MDSampleFilter = " AND (DecorOrders.IsSample = 1 OR (DecorOrders.ProductID=42 AND DecorOrders.IsSample = 0)) ";
                if (IsNotSample)
                    MDSampleFilter = " AND DecorOrders.IsSample = 0";
            }

            string SelectCommand = @"SELECT        DecorOrders.DecorOrderID, PackageDetails.PackageDetailID, DecorOrders.ProductID, DecorOrders.DecorID, DecorOrders.ColorID, DecorOrders.PatinaID, DecorOrders.Length, DecorOrders.Height,
                         DecorOrders.Width, PackageDetails.Count, DecorOrders.DecorConfigID, DecorOrders.FactoryID, DecorOrders.Price, DecorOrders.Cost, DecorOrders.ItemWeight, DecorOrders.Weight,
                         infiniu2_catalog.dbo.DecorConfig.AccountingName, infiniu2_catalog.dbo.DecorConfig.InvNumber, infiniu2_catalog.dbo.Measures.Measure, PackageDetails.OrderID, infiniu2_marketingreference.dbo.Clients.ClientName,
                         MegaOrders.OrderNumber
FROM            PackageDetails INNER JOIN
                         DecorOrders ON PackageDetails.OrderID = DecorOrders.DecorOrderID" + MDSampleFilter + @" LEFT OUTER JOIN
                         infiniu2_catalog.dbo.TechStore AS Fronts ON DecorOrders.DecorID = Fronts.TechStoreID LEFT OUTER JOIN
                         infiniu2_catalog.dbo.TechStore AS Colors ON DecorOrders.ColorID = Colors.TechStoreID INNER JOIN
                         infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID = infiniu2_catalog.dbo.DecorConfig.DecorConfigID INNER JOIN
                         MainOrders ON DecorOrders.MainOrderID = MainOrders.MainOrderID INNER JOIN
                         MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID " + MClientFilter + @" INNER JOIN
                         infiniu2_marketingreference.dbo.Clients ON MegaOrders.ClientID = infiniu2_marketingreference.dbo.Clients.ClientID INNER JOIN
                         infiniu2_catalog.dbo.Measures ON infiniu2_catalog.dbo.DecorConfig.MeasureID = infiniu2_catalog.dbo.Measures.MeasureID
WHERE dbo.PackageDetails.PackageID IN (SELECT TOP (100) PERCENT PackageID FROM dbo.Packages WHERE PackingDateTime < '" +
date1.ToString("yyyy-MM-dd") + @" 23:59:59' AND ProductType = 1 AND (DispatchDateTime IS NULL OR DispatchDateTime > '" + date1.ToString("yyyy-MM-dd") + @" 23:59:59'))
ORDER BY infiniu2_marketingreference.dbo.Clients.ClientName, dbo.MegaOrders.OrderNumber";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(DecorOrdersDataTable);
            }

            Collect(ProfilMarketingEndCountDT, TPSMarketingEndCountDT);
        }

        public void ZOVStartCount(DateTime date1, bool IsSample, bool IsNotSample)
        {
            ProfilZOVStartCountDT.Clear();
            TPSZOVStartCountDT.Clear();
            DecorOrdersDataTable.Clear();

            string MFSampleFilter = string.Empty;
            string MDSampleFilter = string.Empty;
            if (!IsSample || !IsNotSample)
            {
                if (IsSample)
                    MFSampleFilter = " AND FrontsOrders.IsSample = 1";
                if (IsNotSample)
                    MFSampleFilter = " AND FrontsOrders.IsSample = 0";
                if (IsSample)
                    MDSampleFilter = " AND (DecorOrders.IsSample = 1 OR (DecorOrders.ProductID=42 AND DecorOrders.IsSample = 0)) ";
                if (IsNotSample)
                    MDSampleFilter = " AND DecorOrders.IsSample = 0";
            }

            string SelectCommand = @"SELECT        DecorOrders.DecorOrderID, PackageDetails.PackageDetailID, DecorOrders.ProductID, DecorOrders.DecorID, DecorOrders.ColorID, DecorOrders.PatinaID, DecorOrders.Length, DecorOrders.Height,
                         DecorOrders.Width, PackageDetails.Count, DecorOrders.DecorConfigID, DecorOrders.FactoryID, DecorOrders.Price, DecorOrders.Cost, DecorOrders.ItemWeight, DecorOrders.Weight,
                         infiniu2_catalog.dbo.DecorConfig.AccountingName, infiniu2_catalog.dbo.DecorConfig.InvNumber, infiniu2_catalog.dbo.Measures.Measure, PackageDetails.OrderID, infiniu2_zovreference.dbo.Clients.ClientName, MainOrders.DocNumber
FROM            PackageDetails INNER JOIN
                         DecorOrders ON PackageDetails.OrderID = DecorOrders.DecorOrderID" + MDSampleFilter + @" LEFT OUTER JOIN
                         infiniu2_catalog.dbo.TechStore AS Fronts ON DecorOrders.DecorID = Fronts.TechStoreID LEFT OUTER JOIN
                         infiniu2_catalog.dbo.TechStore AS Colors ON DecorOrders.ColorID = Colors.TechStoreID INNER JOIN
                         infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID = infiniu2_catalog.dbo.DecorConfig.DecorConfigID INNER JOIN
                         MainOrders ON DecorOrders.MainOrderID = MainOrders.MainOrderID INNER JOIN
                         infiniu2_zovreference.dbo.Clients ON MainOrders.ClientID = infiniu2_zovreference.dbo.Clients.ClientID INNER JOIN
                         infiniu2_catalog.dbo.Measures ON infiniu2_catalog.dbo.DecorConfig.MeasureID = infiniu2_catalog.dbo.Measures.MeasureID
WHERE        PackageDetails.PackageID IN
                             (SELECT        TOP (100) PERCENT PackageID
                               FROM            Packages
                               WHERE       PackingDateTime < '" + date1.ToString("yyyy-MM-dd") + @" 00:00:00' AND ProductType = 1 AND (DispatchDateTime IS NULL OR DispatchDateTime >= '" + date1.ToString("yyyy-MM-dd") + @" 00:00:00'))
ORDER BY infiniu2_zovreference.dbo.Clients.ClientName, MainOrders.DocNumber";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.ZOVOrdersConnectionString))
            {
                DA.Fill(DecorOrdersDataTable);
            }

            Collect(ProfilZOVStartCountDT, TPSZOVStartCountDT);
        }

        public void ZOVEndCount(DateTime date1, bool IsSample, bool IsNotSample)
        {
            ProfilZOVEndCountDT.Clear();
            TPSZOVEndCountDT.Clear();
            DecorOrdersDataTable.Clear();

            string MFSampleFilter = string.Empty;
            string MDSampleFilter = string.Empty;
            if (!IsSample || !IsNotSample)
            {
                if (IsSample)
                    MFSampleFilter = " AND FrontsOrders.IsSample = 1";
                if (IsNotSample)
                    MFSampleFilter = " AND FrontsOrders.IsSample = 0";
                if (IsSample)
                    MDSampleFilter = " AND (DecorOrders.IsSample = 1 OR (DecorOrders.ProductID=42 AND DecorOrders.IsSample = 0)) ";
                if (IsNotSample)
                    MDSampleFilter = " AND DecorOrders.IsSample = 0";
            }

            string SelectCommand = @"SELECT        DecorOrders.DecorOrderID, PackageDetails.PackageDetailID, DecorOrders.ProductID, DecorOrders.DecorID, DecorOrders.ColorID, DecorOrders.PatinaID, DecorOrders.Length, DecorOrders.Height,
                         DecorOrders.Width, PackageDetails.Count, DecorOrders.DecorConfigID, DecorOrders.FactoryID, DecorOrders.Price, DecorOrders.Cost, DecorOrders.ItemWeight, DecorOrders.Weight,
                         infiniu2_catalog.dbo.DecorConfig.AccountingName, infiniu2_catalog.dbo.DecorConfig.InvNumber, infiniu2_catalog.dbo.Measures.Measure, PackageDetails.OrderID, infiniu2_zovreference.dbo.Clients.ClientName, MainOrders.DocNumber
FROM            PackageDetails INNER JOIN
                         DecorOrders ON PackageDetails.OrderID = DecorOrders.DecorOrderID" + MDSampleFilter + @" LEFT OUTER JOIN
                         infiniu2_catalog.dbo.TechStore AS Fronts ON DecorOrders.DecorID = Fronts.TechStoreID LEFT OUTER JOIN
                         infiniu2_catalog.dbo.TechStore AS Colors ON DecorOrders.ColorID = Colors.TechStoreID INNER JOIN
                         infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID = infiniu2_catalog.dbo.DecorConfig.DecorConfigID INNER JOIN
                         MainOrders ON DecorOrders.MainOrderID = MainOrders.MainOrderID INNER JOIN
                         infiniu2_zovreference.dbo.Clients ON MainOrders.ClientID = infiniu2_zovreference.dbo.Clients.ClientID INNER JOIN
                         infiniu2_catalog.dbo.Measures ON infiniu2_catalog.dbo.DecorConfig.MeasureID = infiniu2_catalog.dbo.Measures.MeasureID
WHERE        PackageDetails.PackageID IN
                             (SELECT        TOP (100) PERCENT PackageID
                               FROM            Packages
                               WHERE      PackingDateTime < '" + date1.ToString("yyyy-MM-dd") + @" 23:59:59' AND ProductType = 1 AND
                                                         (DispatchDateTime IS NULL OR DispatchDateTime >= '" + date1.ToString("yyyy-MM-dd") + @" 23:59:59'))
ORDER BY infiniu2_zovreference.dbo.Clients.ClientName, MainOrders.DocNumber";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.ZOVOrdersConnectionString))
            {
                DA.Fill(DecorOrdersDataTable);
            }

            Collect(ProfilZOVEndCountDT, TPSZOVEndCountDT);
        }

        public void StartEndProducedDispatchReport(DateTime date1, DateTime date2, bool IsSample, bool IsNotSample, ArrayList MClients, ArrayList MClientGroups)
        {
            MarketingProduced(date1, date2, IsSample, IsNotSample, MClients, MClientGroups);
            if (date2 < new DateTime(2016, 07, 01))
                ZOVProduced(date1, date2, IsSample, IsNotSample);
            MarketingDispatch(date1, date2, IsSample, IsNotSample, MClients, MClientGroups);
            if (date2 < new DateTime(2016, 07, 01))
                ZOVDispatch(date1, date2, IsSample, IsNotSample);
            MarketingStartCount(date1, IsSample, IsNotSample, MClients, MClientGroups);
            MarketingEndCount(date2, IsSample, IsNotSample, MClients, MClientGroups);
            if (date2 < new DateTime(2016, 07, 01))
                ZOVStartCount(date1, IsSample, IsNotSample);
            if (date2 < new DateTime(2016, 07, 01))
                ZOVEndCount(date2, IsSample, IsNotSample);

            CollectProfilOrders();
            CollectTPSOrders();
        }

        private void CollectProfilOrders()
        {
            AllProfilDT.Clear();

            for (int i = 0; i < ProfilMarketingStartCountDT.Rows.Count; i++)
            {
                string rr = ProfilMarketingStartCountDT.Rows[i]["InvNumber"].ToString();
                DataRow[] ItemsRows = AllProfilDT.Select("InvNumber = '" + rr + "'");
                if (ItemsRows.Count() > 0)
                {
                    foreach (DataRow item in ItemsRows)
                    {
                        if (item["StartCount"] == DBNull.Value)
                            item["StartCount"] = ProfilMarketingStartCountDT.Rows[i]["Count"];
                        else
                            item["StartCount"] = Convert.ToDecimal(item["StartCount"]) + Convert.ToDecimal(ProfilMarketingStartCountDT.Rows[i]["Count"]);
                    }
                }
                else
                {
                    DataRow NewRow = AllProfilDT.NewRow();
                    NewRow["AccountingName"] = ProfilMarketingStartCountDT.Rows[i]["AccountingName"];
                    NewRow["Measure"] = ProfilMarketingStartCountDT.Rows[i]["Measure"];
                    NewRow["InvNumber"] = ProfilMarketingStartCountDT.Rows[i]["InvNumber"];
                    NewRow["StartCount"] = ProfilMarketingStartCountDT.Rows[i]["Count"];
                    AllProfilDT.Rows.Add(NewRow);
                }
            }
            for (int i = 0; i < ProfilMarketingEndCountDT.Rows.Count; i++)
            {
                string rr = ProfilMarketingEndCountDT.Rows[i]["InvNumber"].ToString();
                DataRow[] ItemsRows = AllProfilDT.Select("InvNumber = '" + rr + "'");
                if (ItemsRows.Count() > 0)
                {
                    foreach (DataRow item in ItemsRows)
                    {
                        if (item["EndCount"] == DBNull.Value)
                            item["EndCount"] = ProfilMarketingEndCountDT.Rows[i]["Count"];
                        else
                            item["EndCount"] = Convert.ToDecimal(item["EndCount"]) + Convert.ToDecimal(ProfilMarketingEndCountDT.Rows[i]["Count"]);
                    }
                }
                else
                {
                    DataRow NewRow = AllProfilDT.NewRow();
                    NewRow["AccountingName"] = ProfilMarketingEndCountDT.Rows[i]["AccountingName"];
                    NewRow["Measure"] = ProfilMarketingEndCountDT.Rows[i]["Measure"];
                    NewRow["InvNumber"] = ProfilMarketingEndCountDT.Rows[i]["InvNumber"];
                    NewRow["EndCount"] = ProfilMarketingEndCountDT.Rows[i]["Count"];
                    AllProfilDT.Rows.Add(NewRow);
                }
            }
            for (int i = 0; i < ProfilMarketingProducedDT.Rows.Count; i++)
            {
                string rr = ProfilMarketingProducedDT.Rows[i]["InvNumber"].ToString();
                DataRow[] ItemsRows = AllProfilDT.Select("InvNumber = '" + rr + "'");
                if (ItemsRows.Count() > 0)
                {
                    foreach (DataRow item in ItemsRows)
                    {
                        if (item["ProducedCount"] == DBNull.Value)
                            item["ProducedCount"] = ProfilMarketingProducedDT.Rows[i]["Count"];
                        else
                            item["ProducedCount"] = Convert.ToDecimal(item["ProducedCount"]) + Convert.ToDecimal(ProfilMarketingProducedDT.Rows[i]["Count"]);
                    }
                }
                else
                {
                    DataRow NewRow = AllProfilDT.NewRow();
                    NewRow["AccountingName"] = ProfilMarketingProducedDT.Rows[i]["AccountingName"];
                    NewRow["Measure"] = ProfilMarketingProducedDT.Rows[i]["Measure"];
                    NewRow["InvNumber"] = ProfilMarketingProducedDT.Rows[i]["InvNumber"];
                    NewRow["ProducedCount"] = ProfilMarketingProducedDT.Rows[i]["Count"];
                    AllProfilDT.Rows.Add(NewRow);
                }
            }
            for (int i = 0; i < ProfilMarketingDispatchDT.Rows.Count; i++)
            {
                string rr = ProfilMarketingDispatchDT.Rows[i]["InvNumber"].ToString();
                DataRow[] ItemsRows = AllProfilDT.Select("InvNumber = '" + rr + "'");
                if (ItemsRows.Count() > 0)
                {
                    foreach (DataRow item in ItemsRows)
                    {
                        if (item["DispatchCount"] == DBNull.Value)
                            item["DispatchCount"] = ProfilMarketingDispatchDT.Rows[i]["Count"];
                        else
                            item["DispatchCount"] = Convert.ToDecimal(item["DispatchCount"]) + Convert.ToDecimal(ProfilMarketingDispatchDT.Rows[i]["Count"]);
                    }
                }
                else
                {
                    DataRow NewRow = AllProfilDT.NewRow();
                    NewRow["AccountingName"] = ProfilMarketingDispatchDT.Rows[i]["AccountingName"];
                    NewRow["Measure"] = ProfilMarketingDispatchDT.Rows[i]["Measure"];
                    NewRow["InvNumber"] = ProfilMarketingDispatchDT.Rows[i]["InvNumber"];
                    NewRow["DispatchCount"] = ProfilMarketingDispatchDT.Rows[i]["Count"];
                    AllProfilDT.Rows.Add(NewRow);
                }
            }

            for (int i = 0; i < ProfilZOVStartCountDT.Rows.Count; i++)
            {
                string rr = ProfilZOVStartCountDT.Rows[i]["InvNumber"].ToString();
                DataRow[] ItemsRows = AllProfilDT.Select("InvNumber = '" + rr + "'");
                if (ItemsRows.Count() > 0)
                {
                    foreach (DataRow item in ItemsRows)
                    {
                        if (item["StartCount"] == DBNull.Value)
                            item["StartCount"] = ProfilZOVStartCountDT.Rows[i]["Count"];
                        else
                            item["StartCount"] = Convert.ToDecimal(item["StartCount"]) + Convert.ToDecimal(ProfilZOVStartCountDT.Rows[i]["Count"]);
                    }
                }
                else
                {
                    DataRow NewRow = AllProfilDT.NewRow();
                    NewRow["AccountingName"] = ProfilZOVStartCountDT.Rows[i]["AccountingName"];
                    NewRow["Measure"] = ProfilZOVStartCountDT.Rows[i]["Measure"];
                    NewRow["InvNumber"] = ProfilZOVStartCountDT.Rows[i]["InvNumber"];
                    NewRow["StartCount"] = ProfilZOVStartCountDT.Rows[i]["Count"];
                    AllProfilDT.Rows.Add(NewRow);
                }
            }
            for (int i = 0; i < ProfilZOVEndCountDT.Rows.Count; i++)
            {
                string rr = ProfilZOVEndCountDT.Rows[i]["InvNumber"].ToString();
                DataRow[] ItemsRows = AllProfilDT.Select("InvNumber = '" + rr + "'");
                if (ItemsRows.Count() > 0)
                {
                    foreach (DataRow item in ItemsRows)
                    {
                        if (item["EndCount"] == DBNull.Value)
                            item["EndCount"] = ProfilZOVEndCountDT.Rows[i]["Count"];
                        else
                            item["EndCount"] = Convert.ToDecimal(item["EndCount"]) + Convert.ToDecimal(ProfilZOVEndCountDT.Rows[i]["Count"]);
                    }
                }
                else
                {
                    DataRow NewRow = AllProfilDT.NewRow();
                    NewRow["AccountingName"] = ProfilZOVEndCountDT.Rows[i]["AccountingName"];
                    NewRow["Measure"] = ProfilZOVEndCountDT.Rows[i]["Measure"];
                    NewRow["InvNumber"] = ProfilZOVEndCountDT.Rows[i]["InvNumber"];
                    NewRow["EndCount"] = ProfilZOVEndCountDT.Rows[i]["Count"];
                    AllProfilDT.Rows.Add(NewRow);
                }
            }
            for (int i = 0; i < ProfilZOVProducedDT.Rows.Count; i++)
            {
                string rr = ProfilZOVProducedDT.Rows[i]["InvNumber"].ToString();
                DataRow[] ItemsRows = AllProfilDT.Select("InvNumber = '" + rr + "'");
                if (ItemsRows.Count() > 0)
                {
                    foreach (DataRow item in ItemsRows)
                    {
                        if (item["ProducedCount"] == DBNull.Value)
                            item["ProducedCount"] = ProfilZOVProducedDT.Rows[i]["Count"];
                        else
                            item["ProducedCount"] = Convert.ToDecimal(item["ProducedCount"]) + Convert.ToDecimal(ProfilZOVProducedDT.Rows[i]["Count"]);
                    }
                }
                else
                {
                    DataRow NewRow = AllProfilDT.NewRow();
                    NewRow["AccountingName"] = ProfilZOVProducedDT.Rows[i]["AccountingName"];
                    NewRow["Measure"] = ProfilZOVProducedDT.Rows[i]["Measure"];
                    NewRow["InvNumber"] = ProfilZOVProducedDT.Rows[i]["InvNumber"];
                    NewRow["ProducedCount"] = ProfilZOVProducedDT.Rows[i]["Count"];
                    AllProfilDT.Rows.Add(NewRow);
                }
            }
            for (int i = 0; i < ProfilZOVDispatchDT.Rows.Count; i++)
            {
                string rr = ProfilZOVDispatchDT.Rows[i]["InvNumber"].ToString();
                DataRow[] ItemsRows = AllProfilDT.Select("InvNumber = '" + rr + "'");
                if (ItemsRows.Count() > 0)
                {
                    foreach (DataRow item in ItemsRows)
                    {
                        if (item["DispatchCount"] == DBNull.Value)
                            item["DispatchCount"] = ProfilZOVDispatchDT.Rows[i]["Count"];
                        else
                            item["DispatchCount"] = Convert.ToDecimal(item["DispatchCount"]) + Convert.ToDecimal(ProfilZOVDispatchDT.Rows[i]["Count"]);
                    }
                }
                else
                {
                    DataRow NewRow = AllProfilDT.NewRow();
                    NewRow["AccountingName"] = ProfilZOVDispatchDT.Rows[i]["AccountingName"];
                    NewRow["Measure"] = ProfilZOVDispatchDT.Rows[i]["Measure"];
                    NewRow["InvNumber"] = ProfilZOVDispatchDT.Rows[i]["InvNumber"];
                    NewRow["DispatchCount"] = ProfilZOVDispatchDT.Rows[i]["Count"];
                    AllProfilDT.Rows.Add(NewRow);
                }
            }

            using (DataView DV = new DataView(AllProfilDT.Copy()))
            {
                DV.Sort = "AccountingName";
                AllProfilDT.Clear();
                AllProfilDT = DV.ToTable();
            }
            foreach (DataRow item in AllProfilDT.Rows)
            {
                if (item["StartCount"] == DBNull.Value)
                    item["StartCount"] = 0;
                if (item["EndCount"] == DBNull.Value)
                    item["EndCount"] = 0;
                if (item["ProducedCount"] == DBNull.Value)
                    item["ProducedCount"] = 0;
                if (item["DispatchCount"] == DBNull.Value)
                    item["DispatchCount"] = 0;
            }
        }

        public void CollectTPSOrders()
        {
            AllTPSDT.Clear();

            for (int i = 0; i < TPSMarketingStartCountDT.Rows.Count; i++)
            {
                string rr = TPSMarketingStartCountDT.Rows[i]["InvNumber"].ToString();
                DataRow[] ItemsRows = AllTPSDT.Select("InvNumber = '" + rr + "'");
                if (ItemsRows.Count() > 0)
                {
                    foreach (DataRow item in ItemsRows)
                    {
                        if (item["StartCount"] == DBNull.Value)
                            item["StartCount"] = TPSMarketingStartCountDT.Rows[i]["Count"];
                        else
                            item["StartCount"] = Convert.ToDecimal(item["StartCount"]) + Convert.ToDecimal(TPSMarketingStartCountDT.Rows[i]["Count"]);
                    }
                }
                else
                {
                    DataRow NewRow = AllTPSDT.NewRow();
                    NewRow["AccountingName"] = TPSMarketingStartCountDT.Rows[i]["AccountingName"];
                    NewRow["Measure"] = TPSMarketingStartCountDT.Rows[i]["Measure"];
                    NewRow["InvNumber"] = TPSMarketingStartCountDT.Rows[i]["InvNumber"];
                    NewRow["StartCount"] = TPSMarketingStartCountDT.Rows[i]["Count"];
                    AllTPSDT.Rows.Add(NewRow);
                }
            }
            for (int i = 0; i < TPSMarketingEndCountDT.Rows.Count; i++)
            {
                string rr = TPSMarketingEndCountDT.Rows[i]["InvNumber"].ToString();
                DataRow[] ItemsRows = AllTPSDT.Select("InvNumber = '" + rr + "'");
                if (ItemsRows.Count() > 0)
                {
                    foreach (DataRow item in ItemsRows)
                    {
                        if (item["EndCount"] == DBNull.Value)
                            item["EndCount"] = TPSMarketingEndCountDT.Rows[i]["Count"];
                        else
                            item["EndCount"] = Convert.ToDecimal(item["EndCount"]) + Convert.ToDecimal(TPSMarketingEndCountDT.Rows[i]["Count"]);
                    }
                }
                else
                {
                    DataRow NewRow = AllTPSDT.NewRow();
                    NewRow["AccountingName"] = TPSMarketingEndCountDT.Rows[i]["AccountingName"];
                    NewRow["Measure"] = TPSMarketingEndCountDT.Rows[i]["Measure"];
                    NewRow["InvNumber"] = TPSMarketingEndCountDT.Rows[i]["InvNumber"];
                    NewRow["EndCount"] = TPSMarketingEndCountDT.Rows[i]["Count"];
                    AllTPSDT.Rows.Add(NewRow);
                }
            }
            for (int i = 0; i < TPSMarketingProducedDT.Rows.Count; i++)
            {
                string rr = TPSMarketingProducedDT.Rows[i]["InvNumber"].ToString();
                DataRow[] ItemsRows = AllTPSDT.Select("InvNumber = '" + rr + "'");
                if (ItemsRows.Count() > 0)
                {
                    foreach (DataRow item in ItemsRows)
                    {
                        if (item["ProducedCount"] == DBNull.Value)
                            item["ProducedCount"] = TPSMarketingProducedDT.Rows[i]["Count"];
                        else
                            item["ProducedCount"] = Convert.ToDecimal(item["ProducedCount"]) + Convert.ToDecimal(TPSMarketingProducedDT.Rows[i]["Count"]);
                    }
                }
                else
                {
                    DataRow NewRow = AllTPSDT.NewRow();
                    NewRow["AccountingName"] = TPSMarketingProducedDT.Rows[i]["AccountingName"];
                    NewRow["Measure"] = TPSMarketingProducedDT.Rows[i]["Measure"];
                    NewRow["InvNumber"] = TPSMarketingProducedDT.Rows[i]["InvNumber"];
                    NewRow["ProducedCount"] = TPSMarketingProducedDT.Rows[i]["Count"];
                    AllTPSDT.Rows.Add(NewRow);
                }
            }
            for (int i = 0; i < TPSMarketingDispatchDT.Rows.Count; i++)
            {
                string rr = TPSMarketingDispatchDT.Rows[i]["InvNumber"].ToString();
                DataRow[] ItemsRows = AllTPSDT.Select("InvNumber = '" + rr + "'");
                if (ItemsRows.Count() > 0)
                {
                    foreach (DataRow item in ItemsRows)
                    {
                        if (item["DispatchCount"] == DBNull.Value)
                            item["DispatchCount"] = TPSMarketingDispatchDT.Rows[i]["Count"];
                        else
                            item["DispatchCount"] = Convert.ToDecimal(item["DispatchCount"]) + Convert.ToDecimal(TPSMarketingDispatchDT.Rows[i]["Count"]);
                    }
                }
                else
                {
                    DataRow NewRow = AllTPSDT.NewRow();
                    NewRow["AccountingName"] = TPSMarketingDispatchDT.Rows[i]["AccountingName"];
                    NewRow["Measure"] = TPSMarketingDispatchDT.Rows[i]["Measure"];
                    NewRow["InvNumber"] = TPSMarketingDispatchDT.Rows[i]["InvNumber"];
                    NewRow["DispatchCount"] = TPSMarketingDispatchDT.Rows[i]["Count"];
                    AllTPSDT.Rows.Add(NewRow);
                }
            }

            for (int i = 0; i < TPSZOVStartCountDT.Rows.Count; i++)
            {
                string rr = TPSZOVStartCountDT.Rows[i]["InvNumber"].ToString();
                DataRow[] ItemsRows = AllTPSDT.Select("InvNumber = '" + rr + "'");
                if (ItemsRows.Count() > 0)
                {
                    foreach (DataRow item in ItemsRows)
                    {
                        if (item["StartCount"] == DBNull.Value)
                            item["StartCount"] = TPSZOVStartCountDT.Rows[i]["Count"];
                        else
                            item["StartCount"] = Convert.ToDecimal(item["StartCount"]) + Convert.ToDecimal(TPSZOVStartCountDT.Rows[i]["Count"]);
                    }
                }
                else
                {
                    DataRow NewRow = AllTPSDT.NewRow();
                    NewRow["AccountingName"] = TPSZOVStartCountDT.Rows[i]["AccountingName"];
                    NewRow["Measure"] = ProfilMarketingStartCountDT.Rows[i]["Measure"];
                    NewRow["InvNumber"] = TPSZOVStartCountDT.Rows[i]["InvNumber"];
                    NewRow["StartCount"] = TPSZOVStartCountDT.Rows[i]["Count"];
                    AllTPSDT.Rows.Add(NewRow);
                }
            }
            for (int i = 0; i < TPSZOVEndCountDT.Rows.Count; i++)
            {
                string rr = TPSZOVEndCountDT.Rows[i]["InvNumber"].ToString();
                DataRow[] ItemsRows = AllTPSDT.Select("InvNumber = '" + rr + "'");
                if (ItemsRows.Count() > 0)
                {
                    foreach (DataRow item in ItemsRows)
                    {
                        if (item["EndCount"] == DBNull.Value)
                            item["EndCount"] = TPSZOVEndCountDT.Rows[i]["Count"];
                        else
                            item["EndCount"] = Convert.ToDecimal(item["EndCount"]) + Convert.ToDecimal(TPSZOVEndCountDT.Rows[i]["Count"]);
                    }
                }
                else
                {
                    DataRow NewRow = AllTPSDT.NewRow();
                    NewRow["AccountingName"] = TPSZOVEndCountDT.Rows[i]["AccountingName"];
                    NewRow["Measure"] = TPSZOVEndCountDT.Rows[i]["Measure"];
                    NewRow["InvNumber"] = TPSZOVEndCountDT.Rows[i]["InvNumber"];
                    NewRow["EndCount"] = TPSZOVEndCountDT.Rows[i]["Count"];
                    AllTPSDT.Rows.Add(NewRow);
                }
            }
            for (int i = 0; i < TPSZOVProducedDT.Rows.Count; i++)
            {
                string rr = TPSZOVProducedDT.Rows[i]["InvNumber"].ToString();
                DataRow[] ItemsRows = AllTPSDT.Select("InvNumber = '" + rr + "'");
                if (ItemsRows.Count() > 0)
                {
                    foreach (DataRow item in ItemsRows)
                    {
                        if (item["ProducedCount"] == DBNull.Value)
                            item["ProducedCount"] = TPSZOVProducedDT.Rows[i]["Count"];
                        else
                            item["ProducedCount"] = Convert.ToDecimal(item["ProducedCount"]) + Convert.ToDecimal(TPSZOVProducedDT.Rows[i]["Count"]);
                    }
                }
                else
                {
                    DataRow NewRow = AllTPSDT.NewRow();
                    NewRow["AccountingName"] = TPSZOVProducedDT.Rows[i]["AccountingName"];
                    NewRow["Measure"] = TPSZOVProducedDT.Rows[i]["Measure"];
                    NewRow["InvNumber"] = TPSZOVProducedDT.Rows[i]["InvNumber"];
                    NewRow["ProducedCount"] = TPSZOVProducedDT.Rows[i]["Count"];
                    AllTPSDT.Rows.Add(NewRow);
                }
            }
            for (int i = 0; i < TPSZOVDispatchDT.Rows.Count; i++)
            {
                string rr = TPSZOVDispatchDT.Rows[i]["InvNumber"].ToString();
                DataRow[] ItemsRows = AllTPSDT.Select("InvNumber = '" + rr + "'");
                if (ItemsRows.Count() > 0)
                {
                    foreach (DataRow item in ItemsRows)
                    {
                        if (item["DispatchCount"] == DBNull.Value)
                            item["DispatchCount"] = TPSZOVDispatchDT.Rows[i]["Count"];
                        else
                            item["DispatchCount"] = Convert.ToDecimal(item["DispatchCount"]) + Convert.ToDecimal(TPSZOVDispatchDT.Rows[i]["Count"]);
                    }
                }
                else
                {
                    DataRow NewRow = AllTPSDT.NewRow();
                    NewRow["AccountingName"] = TPSZOVDispatchDT.Rows[i]["AccountingName"];
                    NewRow["Measure"] = TPSZOVDispatchDT.Rows[i]["Measure"];
                    NewRow["InvNumber"] = TPSZOVDispatchDT.Rows[i]["InvNumber"];
                    NewRow["DispatchCount"] = TPSZOVDispatchDT.Rows[i]["Count"];
                    AllTPSDT.Rows.Add(NewRow);
                }
            }

            using (DataView DV = new DataView(AllTPSDT.Copy()))
            {
                DV.Sort = "AccountingName";
                AllTPSDT.Clear();
                AllTPSDT = DV.ToTable();
            }
            foreach (DataRow item in AllTPSDT.Rows)
            {
                if (item["StartCount"] == DBNull.Value)
                    item["StartCount"] = 0;
                if (item["EndCount"] == DBNull.Value)
                    item["EndCount"] = 0;
                if (item["ProducedCount"] == DBNull.Value)
                    item["ProducedCount"] = 0;
                if (item["DispatchCount"] == DBNull.Value)
                    item["DispatchCount"] = 0;
            }
        }
    }

    public class ProducedReport
    {
        private readonly decimal VAT = 1.0m;
        public ProducedFrontsReport FrontsReport;
        public ProducedDecorReport DecorReport = null;

        public DataTable CurrencyTypesDataTable = null;
        public DataTable ProfilReportTable = null;
        public DataTable TPSReportTable = null;

        public ProducedReport(ref Infinium.Modules.ZOV.DecorCatalogOrder DecorCatalogOrder)
        {
            FrontsReport = new ProducedFrontsReport();
            DecorReport = new ProducedDecorReport(ref DecorCatalogOrder);

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
            ProfilReportTable.Columns.Add(new DataColumn("Measure", Type.GetType("System.String")));
            ProfilReportTable.Columns.Add(new DataColumn("InvNumber", Type.GetType("System.String")));
            ProfilReportTable.Columns.Add(new DataColumn("Cvet", Type.GetType("System.String")));
            ProfilReportTable.Columns.Add(new DataColumn("Patina", Type.GetType("System.String")));
            ProfilReportTable.Columns.Add(new DataColumn("AccountingName", Type.GetType("System.String")));
            ProfilReportTable.Columns.Add(new DataColumn("Count", Type.GetType("System.Decimal")));
            ProfilReportTable.Columns.Add(new DataColumn("Price", Type.GetType("System.Decimal")));
            ProfilReportTable.Columns.Add(new DataColumn("Cost", Type.GetType("System.Decimal")));
            ProfilReportTable.Columns.Add(new DataColumn("Weight", Type.GetType("System.Decimal")));
            ProfilReportTable.Columns.Add(new DataColumn("PackageCount", Type.GetType("System.Int32")));
            ProfilReportTable.Columns.Add(new DataColumn("OriginalPrice", Type.GetType("System.Decimal")));
            ProfilReportTable.Columns.Add(new DataColumn("MarketingCost", Type.GetType("System.Decimal")));

            TPSReportTable = ProfilReportTable.Clone();
        }

        public void Collect(decimal Rate)
        {
            DataTable DistinctInvNumbersDT;
            using (DataView DV = new DataView(ProfilReportTable))
            {
                DistinctInvNumbersDT = DV.ToTable(true, new string[] { "AccountingName", "InvNumber", "Cvet", "Patina", "Measure" });
            }

            for (int i = 0; i < DistinctInvNumbersDT.Rows.Count; i++)
            {
                decimal Cost = 0;
                decimal Count = 0;
                decimal Price = 0;
                decimal Weight = 0;
                int DecCount = 2;
                string UNN = ProfilReportTable.Rows[0]["UNN"].ToString();
                string AccountingName = DistinctInvNumbersDT.Rows[i]["AccountingName"].ToString();
                string InvNumber = DistinctInvNumbersDT.Rows[i]["InvNumber"].ToString();
                string Cvet = DistinctInvNumbersDT.Rows[i]["Cvet"].ToString();
                string Patina = DistinctInvNumbersDT.Rows[i]["Patina"].ToString();
                string Measure = DistinctInvNumbersDT.Rows[i]["Measure"].ToString();
                string CurrencyCode = ProfilReportTable.Rows[0]["CurrencyCode"].ToString();
                
                //if (Convert.ToInt32(CurrencyCode) == 974)
                //    DecCount = 0;
                string filter = "Cvet = '" + Cvet + "' AND Patina = '" + Patina + "' AND InvNumber = '" + InvNumber +
                                "' AND AccountingName = '" + AccountingName + "'";
                DataRow[] InvRows = ProfilReportTable.Select(filter);


                var filteredRows = ProfilReportTable
                    .AsEnumerable()
                    .Where(row => row.Field<string>("InvNumber") == InvNumber
                                  && row.Field<string>("AccountingName") == AccountingName
                                  && row.Field<string>("Cvet") == Cvet
                                  && row.Field<string>("Patina") == Patina);
                DataRow[] rows2 = filteredRows.ToArray();

                if (InvRows.Count() == 1)
                {
                    Count = Convert.ToDecimal(InvRows[0]["Count"]);
                    Cost = Convert.ToDecimal(InvRows[0]["Cost"]) * Rate / VAT;
                    Price = Convert.ToDecimal(InvRows[0]["MarketingCost"]) * Rate;
                    //if (Convert.ToInt32(CurrencyCode) == 974)
                    //{
                    //    DecCount = 0;
                    //    Price = Math.Ceiling(Price / 100.0m) * 100.0m;
                    //    Cost = Price * Count;
                    //}
                    InvRows[0]["Price"] = Decimal.Round(Price, DecCount, MidpointRounding.AwayFromZero);
                    InvRows[0]["Cost"] = Decimal.Round(Cost, DecCount, MidpointRounding.AwayFromZero);
                    continue;
                }
                else
                {
                    for (int j = 0; j < InvRows.Count(); j++)
                    {
                        Cost += Convert.ToDecimal(InvRows[j]["Cost"]);
                        Count += Convert.ToDecimal(InvRows[j]["Count"]);
                        Weight += Convert.ToDecimal(InvRows[j]["Weight"]);
                    }
                    decimal MarketingCost = Convert.ToDecimal(InvRows[0]["MarketingCost"]);
                    Price = Convert.ToDecimal(InvRows[0]["MarketingCost"]) * Rate;
                    for (int j = 0; j < InvRows.Count(); j++)
                    {
                        InvRows[j].Delete();
                    }
                    ProfilReportTable.AcceptChanges();
                    Cost = Cost * Rate / VAT;
                    //if (Convert.ToInt32(CurrencyCode) == 974)
                    //{
                    //    DecCount = 0;
                    //    Price = Math.Ceiling(Price / 100.0m) * 100.0m;
                    //    Cost = Price * Count;
                    //}
                    DataRow NewRow = ProfilReportTable.NewRow();
                    NewRow["UNN"] = UNN;
                    NewRow["MarketingCost"] = MarketingCost;
                    NewRow["AccountingName"] = AccountingName;
                    NewRow["InvNumber"] = InvNumber;
                    NewRow["Cvet"] = Cvet;
                    NewRow["Patina"] = Patina;
                    NewRow["Measure"] = Measure;
                    NewRow["CurrencyCode"] = CurrencyCode;
                    NewRow["Count"] = Count;
                    NewRow["Price"] = Decimal.Round(Price, DecCount, MidpointRounding.AwayFromZero);
                    NewRow["Cost"] = Decimal.Round(Cost, DecCount, MidpointRounding.AwayFromZero);
                    NewRow["Weight"] = Decimal.Round(Weight, 3, MidpointRounding.AwayFromZero);
                    ProfilReportTable.Rows.Add(NewRow);
                }
            }

            DistinctInvNumbersDT.Dispose();
            using (DataView DV = new DataView(TPSReportTable))
            {
                DistinctInvNumbersDT = DV.ToTable(true, new string[] { "AccountingName", "InvNumber", "Cvet", "Patina", "Measure" });
            }

            for (int i = 0; i < DistinctInvNumbersDT.Rows.Count; i++)
            {
                decimal Cost = 0;
                decimal Count = 0;
                decimal Price = 0;
                decimal Weight = 0;
                int DecCount = 2;
                string UNN = TPSReportTable.Rows[0]["UNN"].ToString();
                string AccountingName = DistinctInvNumbersDT.Rows[i]["AccountingName"].ToString();
                string InvNumber = DistinctInvNumbersDT.Rows[i]["InvNumber"].ToString();
                string Cvet = DistinctInvNumbersDT.Rows[i]["Cvet"].ToString();
                string Patina = DistinctInvNumbersDT.Rows[i]["Patina"].ToString();
                string Measure = DistinctInvNumbersDT.Rows[i]["Measure"].ToString();
                string CurrencyCode = TPSReportTable.Rows[0]["CurrencyCode"].ToString();
                string TPSCurCode = TPSReportTable.Rows[0]["TPSCurCode"].ToString();

                //if (TPSCurCode.Equals("974"))
                //    DecCount = 0;

                DataRow[] InvRows = TPSReportTable.Select("Cvet = '" + Cvet + "' AND Patina = '" + Patina + "' AND InvNumber = '" + InvNumber + "' AND AccountingName = '" + AccountingName + "'");
                if (InvRows.Count() == 1)
                {
                    Count = Convert.ToDecimal(InvRows[0]["Count"]);
                    Cost = Convert.ToDecimal(InvRows[0]["Cost"]) * Rate / VAT;
                    Price = Convert.ToDecimal(InvRows[0]["MarketingCost"]) * Rate;
                    //if (TPSCurCode.Equals("974"))
                    //{
                    //    DecCount = 0;
                    //    Price = Math.Ceiling(Price / 100.0m) * 100.0m;
                    //    Cost = Price * Count;
                    //}
                    InvRows[0]["Price"] = Decimal.Round(Price, DecCount, MidpointRounding.AwayFromZero);
                    InvRows[0]["Cost"] = Decimal.Round(Cost, DecCount, MidpointRounding.AwayFromZero);
                    continue;
                }
                else
                {
                    for (int j = 0; j < InvRows.Count(); j++)
                    {
                        Cost += Convert.ToDecimal(InvRows[j]["Cost"]);
                        Count += Convert.ToDecimal(InvRows[j]["Count"]);
                        Weight += Convert.ToDecimal(InvRows[j]["Weight"]);
                    }
                    decimal MarketingCost = Convert.ToDecimal(InvRows[0]["MarketingCost"]);
                    Price = Convert.ToDecimal(InvRows[0]["MarketingCost"]) * Rate;
                    for (int j = 0; j < InvRows.Count(); j++)
                    {
                        InvRows[j].Delete();
                    }
                    TPSReportTable.AcceptChanges();
                    Cost = Cost * Rate / VAT;
                    //if (Convert.ToInt32(CurrencyCode) == 974)
                    //{
                    //    DecCount = 0;
                    //    Price = Math.Ceiling(Price / 100.0m) * 100.0m;
                    //    Cost = Price * Count;
                    //}
                    DataRow NewRow = TPSReportTable.NewRow();
                    NewRow["UNN"] = UNN;
                    NewRow["MarketingCost"] = MarketingCost;
                    NewRow["AccountingName"] = AccountingName;
                    NewRow["Measure"] = Measure;
                    NewRow["Cvet"] = Cvet;
                    NewRow["Patina"] = Patina;
                    NewRow["InvNumber"] = InvNumber;
                    NewRow["CurrencyCode"] = CurrencyCode;
                    NewRow["TPSCurCode"] = TPSCurCode;
                    NewRow["Count"] = Count;
                    NewRow["Price"] = Decimal.Round(Price, DecCount, MidpointRounding.AwayFromZero);
                    NewRow["Cost"] = Decimal.Round(Cost, DecCount, MidpointRounding.AwayFromZero);
                    NewRow["Weight"] = Decimal.Round(Weight, 3, MidpointRounding.AwayFromZero);
                    TPSReportTable.Rows.Add(NewRow);
                }
            }
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
        }

        public void GetDateRates(DateTime date, ref decimal EURBYRCurrency)
        {
            using (DataTable DT = new DataTable())
            {
                using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT * FROM DateRates WHERE CAST(Date AS Date) =
                    '" + date.ToString("yyyy-MM-dd") + "'",
                    ConnectionStrings.MarketingReferenceConnectionString))
                {
                    DA.Fill(DT);
                    if (DT.Rows.Count > 0)
                    {
                        EURBYRCurrency = Convert.ToDecimal(DT.Rows[0]["BYN"]);
                    }
                    else
                    {
                        decimal EURRUBCurrency = 1000000;
                        decimal USDRUBCurrency = 1000000;
                        decimal EURUSDCurrency = 1000000;

                        OrdersManager.CBRDailyRates(date, ref EURRUBCurrency, ref USDRUBCurrency);
                        EURBYRCurrency = CurrencyConverter.NbrbDailyRates(DateTime.Now);
                        //OrdersManager.NBRBDailyRates(date, ref EURBYRCurrency);
                        
                        if (USDRUBCurrency != 0)
                            EURUSDCurrency = Decimal.Round(EURRUBCurrency / USDRUBCurrency, 4, MidpointRounding.AwayFromZero);
                        OrdersManager.SaveDateRates(date, EURUSDCurrency, EURRUBCurrency, EURBYRCurrency, USDRUBCurrency);
                    }
                }
            }
            return;
        }

        public void CreateMarketingDispReport(DateTime date1, DateTime date2, bool IsSample, bool IsNotSample, string FileName, decimal Rate, ArrayList MClients, ArrayList MClientGroups)
        {
            ClearReport();

            string MainOrdersList = string.Empty;

            int pos = 0;

            FrontsReport.DispReport(date1, date2, false, IsSample, IsNotSample, MClients, MClientGroups);
            DecorReport.DispReport(date1, date2, false, IsSample, IsNotSample, MClients, MClientGroups);

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

            Collect(Rate);
            //Export to excel
            HSSFWorkbook hssfworkbook = new HSSFWorkbook();

            ////create a entry of DocumentSummaryInformation
            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = "NPOI Team";
            hssfworkbook.DocumentSummaryInformation = dsi;

            ////create a entry of SummaryInformation
            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
            si.Subject = "NPOI SDK Example";
            hssfworkbook.SummaryInformation = si;

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
            PriceBelCS.DataFormat = hssfworkbook.CreateDataFormat().GetFormat("### ### ##0.00");
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

            #endregion Create fonts and styles

            HSSFSheet sheet1 = hssfworkbook.CreateSheet("Отчет Маркетинг");
            sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

            sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
            sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
            sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
            sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

            int displayIndex = 0;

            sheet1.SetColumnWidth(displayIndex++, 12 * 256);
            sheet1.SetColumnWidth(displayIndex++, 61 * 256);
            sheet1.SetColumnWidth(displayIndex++, 15 * 256);
            sheet1.SetColumnWidth(displayIndex++, 12 * 256);
            sheet1.SetColumnWidth(displayIndex++, 8 * 256);
            sheet1.SetColumnWidth(displayIndex++, 8 * 256);
            sheet1.SetColumnWidth(displayIndex, 14 * 256);

            HSSFCell Cell1;

            if (ProfilReportTable.Rows.Count > 0)
            {
                displayIndex = 0;

                Cell1 = sheet1.CreateRow(pos++).CreateCell(0);
                Cell1.SetCellValue("ОМЦ-ПРОФИЛЬ:");
                Cell1.CellStyle = SummaryWithoutBorderBelCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Инв.№");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Наименование");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Цвет");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Патина");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Ед.изм.");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Кол-во");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Себестоимость");
                Cell1.CellStyle = SimpleHeaderCS;
                pos++;

                for (int i = 0; i < ProfilReportTable.Rows.Count; i++)
                {
                    displayIndex = 0;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(ProfilReportTable.Rows[i]["InvNumber"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(ProfilReportTable.Rows[i]["AccountingName"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(FrontsReport.GetColorNameByCode(ProfilReportTable.Rows[i]["Cvet"].ToString()));
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(FrontsReport.GetPatinaNameByCode(ProfilReportTable.Rows[i]["Patina"].ToString()));
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(ProfilReportTable.Rows[i]["Measure"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["Count"]));
                    Cell1.CellStyle = CountCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["Price"]));
                    Cell1.CellStyle = PriceBelCS;
                    pos++;
                }
            }

            if (TPSReportTable.Rows.Count > 0)
            {
                //ТПС
                pos++;
                displayIndex = 0;

                Cell1 = sheet1.CreateRow(pos++).CreateCell(0);
                Cell1.SetCellValue("ЗОВ-ТПС:");
                Cell1.CellStyle = SummaryWithoutBorderBelCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Инв.№");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Наименование");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Цвет");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Патина");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Ед.изм.");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Кол-во");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Себестоимость");
                Cell1.CellStyle = SimpleHeaderCS;
                pos++;

                for (int i = 0; i < TPSReportTable.Rows.Count; i++)
                {
                    displayIndex = 0;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(TPSReportTable.Rows[i]["InvNumber"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(TPSReportTable.Rows[i]["AccountingName"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(FrontsReport.GetColorNameByCode(TPSReportTable.Rows[i]["Cvet"].ToString()));
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(FrontsReport.GetPatinaNameByCode(TPSReportTable.Rows[i]["Patina"].ToString()));
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(TPSReportTable.Rows[i]["Measure"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["Count"]));
                    Cell1.CellStyle = CountCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["Price"]));
                    Cell1.CellStyle = PriceBelCS;
                    pos++;
                }
            }
            pos++;
            FileName += " Маркетинг";
            string tempFolder = System.Environment.GetEnvironmentVariable("TEMP");
            FileInfo file = new FileInfo(tempFolder + @"\" + FileName + ".xls");
            int j = 1;
            while (file.Exists == true)
            {
                file = new FileInfo(tempFolder + @"\" + FileName + "(" + j++ + ").xls");
            }

            FileStream NewFile = new FileStream(file.FullName, FileMode.Create);
            hssfworkbook.Write(NewFile);
            NewFile.Close();

            System.Diagnostics.Process.Start(file.FullName);
            ClearReport();
        }

        public void CreateMarketingOnProdReport(DateTime date1, DateTime date2, bool IsSample, bool IsNotSample, string FileName, decimal Rate, ArrayList MClients, ArrayList MClientGroups)
        {
            ClearReport();

            string MainOrdersList = string.Empty;

            int pos = 0;

            FrontsReport.OnProdReport(date1, date2, false, IsSample, IsNotSample, MClients, MClientGroups);
            DecorReport.OnProdReport(date1, date2, false, IsSample, IsNotSample, MClients, MClientGroups);

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

            Collect(Rate);
            //Export to excel
            HSSFWorkbook hssfworkbook = new HSSFWorkbook();

            ////create a entry of DocumentSummaryInformation
            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = "NPOI Team";
            hssfworkbook.DocumentSummaryInformation = dsi;

            ////create a entry of SummaryInformation
            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
            si.Subject = "NPOI SDK Example";
            hssfworkbook.SummaryInformation = si;

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
            PriceBelCS.DataFormat = hssfworkbook.CreateDataFormat().GetFormat("### ### ##0.00");
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

            #endregion Create fonts and styles

            HSSFSheet sheet1 = hssfworkbook.CreateSheet("Отчет Маркетинг");
            sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

            sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
            sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
            sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
            sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

            int displayIndex = 0;

            sheet1.SetColumnWidth(displayIndex++, 12 * 256);
            sheet1.SetColumnWidth(displayIndex++, 61 * 256);
            sheet1.SetColumnWidth(displayIndex++, 15 * 256);
            sheet1.SetColumnWidth(displayIndex++, 12 * 256);
            sheet1.SetColumnWidth(displayIndex++, 8 * 256);
            sheet1.SetColumnWidth(displayIndex++, 8 * 256);
            sheet1.SetColumnWidth(displayIndex, 14 * 256);

            HSSFCell Cell1;

            if (ProfilReportTable.Rows.Count > 0)
            {
                displayIndex = 0;

                Cell1 = sheet1.CreateRow(pos++).CreateCell(0);
                Cell1.SetCellValue("ОМЦ-ПРОФИЛЬ:");
                Cell1.CellStyle = SummaryWithoutBorderBelCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Инв.№");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Наименование");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Цвет");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Патина");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Ед.изм.");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Кол-во");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Себестоимость");
                Cell1.CellStyle = SimpleHeaderCS;
                pos++;

                for (int i = 0; i < ProfilReportTable.Rows.Count; i++)
                {
                    displayIndex = 0;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(ProfilReportTable.Rows[i]["InvNumber"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(ProfilReportTable.Rows[i]["AccountingName"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(FrontsReport.GetColorNameByCode(ProfilReportTable.Rows[i]["Cvet"].ToString()));
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(FrontsReport.GetPatinaNameByCode(ProfilReportTable.Rows[i]["Patina"].ToString()));
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(ProfilReportTable.Rows[i]["Measure"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["Count"]));
                    Cell1.CellStyle = CountCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["Price"]));
                    Cell1.CellStyle = PriceBelCS;
                    pos++;
                }
            }

            if (TPSReportTable.Rows.Count > 0)
            {
                //ТПС
                pos++;
                displayIndex = 0;

                Cell1 = sheet1.CreateRow(pos++).CreateCell(0);
                Cell1.SetCellValue("ЗОВ-ТПС:");
                Cell1.CellStyle = SummaryWithoutBorderBelCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Инв.№");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Наименование");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Цвет");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Патина");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Ед.изм.");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Кол-во");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Себестоимость");
                Cell1.CellStyle = SimpleHeaderCS;
                pos++;

                for (int i = 0; i < TPSReportTable.Rows.Count; i++)
                {
                    displayIndex = 0;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(TPSReportTable.Rows[i]["InvNumber"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(TPSReportTable.Rows[i]["AccountingName"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(FrontsReport.GetColorNameByCode(TPSReportTable.Rows[i]["Cvet"].ToString()));
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(FrontsReport.GetPatinaNameByCode(TPSReportTable.Rows[i]["Patina"].ToString()));
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(TPSReportTable.Rows[i]["Measure"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["Count"]));
                    Cell1.CellStyle = CountCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["Price"]));
                    Cell1.CellStyle = PriceBelCS;
                    pos++;
                }
            }
            pos++;
            FileName += " Маркетинг";
            string tempFolder = System.Environment.GetEnvironmentVariable("TEMP");
            FileInfo file = new FileInfo(tempFolder + @"\" + FileName + ".xls");
            int j = 1;
            while (file.Exists == true)
            {
                file = new FileInfo(tempFolder + @"\" + FileName + "(" + j++ + ").xls");
            }

            FileStream NewFile = new FileStream(file.FullName, FileMode.Create);
            hssfworkbook.Write(NewFile);
            NewFile.Close();

            System.Diagnostics.Process.Start(file.FullName);
            ClearReport();
        }
        
        public void CreateMarketingPlanDispReport(DateTime date1, DateTime date2, bool IsSample, bool IsNotSample, string FileName, decimal Rate, ArrayList MClients, ArrayList MClientGroups)
        {
            ClearReport();

            string MainOrdersList = string.Empty;

            int pos = 0;

            FrontsReport.PlanDispReport(date1, date2, false, IsSample, IsNotSample, MClients, MClientGroups);
            DecorReport.PlanDispReport(date1, date2, false, IsSample, IsNotSample, MClients, MClientGroups);

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

            Collect(Rate);
            //Export to excel
            HSSFWorkbook hssfworkbook = new HSSFWorkbook();

            ////create a entry of DocumentSummaryInformation
            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = "NPOI Team";
            hssfworkbook.DocumentSummaryInformation = dsi;

            ////create a entry of SummaryInformation
            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
            si.Subject = "NPOI SDK Example";
            hssfworkbook.SummaryInformation = si;

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
            PriceBelCS.DataFormat = hssfworkbook.CreateDataFormat().GetFormat("### ### ##0.00");
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

            #endregion Create fonts and styles

            HSSFSheet sheet1 = hssfworkbook.CreateSheet("Отчет Маркетинг");
            sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

            sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
            sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
            sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
            sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

            int displayIndex = 0;

            sheet1.SetColumnWidth(displayIndex++, 12 * 256);
            sheet1.SetColumnWidth(displayIndex++, 61 * 256);
            sheet1.SetColumnWidth(displayIndex++, 15 * 256);
            sheet1.SetColumnWidth(displayIndex++, 12 * 256);
            sheet1.SetColumnWidth(displayIndex++, 8 * 256);
            sheet1.SetColumnWidth(displayIndex++, 8 * 256);
            sheet1.SetColumnWidth(displayIndex, 14 * 256);

            HSSFCell Cell1;

            if (ProfilReportTable.Rows.Count > 0)
            {
                displayIndex = 0;

                Cell1 = sheet1.CreateRow(pos++).CreateCell(0);
                Cell1.SetCellValue("ОМЦ-ПРОФИЛЬ:");
                Cell1.CellStyle = SummaryWithoutBorderBelCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Инв.№");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Наименование");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Цвет");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Патина");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Ед.изм.");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Кол-во");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Себестоимость");
                Cell1.CellStyle = SimpleHeaderCS;
                pos++;

                for (int i = 0; i < ProfilReportTable.Rows.Count; i++)
                {
                    displayIndex = 0;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(ProfilReportTable.Rows[i]["InvNumber"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(ProfilReportTable.Rows[i]["AccountingName"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(FrontsReport.GetColorNameByCode(ProfilReportTable.Rows[i]["Cvet"].ToString()));
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(FrontsReport.GetPatinaNameByCode(ProfilReportTable.Rows[i]["Patina"].ToString()));
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(ProfilReportTable.Rows[i]["Measure"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["Count"]));
                    Cell1.CellStyle = CountCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["Price"]));
                    Cell1.CellStyle = PriceBelCS;
                    pos++;
                }
            }

            if (TPSReportTable.Rows.Count > 0)
            {
                //ТПС
                pos++;
                displayIndex = 0;

                Cell1 = sheet1.CreateRow(pos++).CreateCell(0);
                Cell1.SetCellValue("ЗОВ-ТПС:");
                Cell1.CellStyle = SummaryWithoutBorderBelCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Инв.№");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Наименование");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Цвет");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Патина");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Ед.изм.");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Кол-во");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Себестоимость");
                Cell1.CellStyle = SimpleHeaderCS;
                pos++;

                for (int i = 0; i < TPSReportTable.Rows.Count; i++)
                {
                    displayIndex = 0;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(TPSReportTable.Rows[i]["InvNumber"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(TPSReportTable.Rows[i]["AccountingName"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(FrontsReport.GetColorNameByCode(TPSReportTable.Rows[i]["Cvet"].ToString()));
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(FrontsReport.GetPatinaNameByCode(TPSReportTable.Rows[i]["Patina"].ToString()));
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(TPSReportTable.Rows[i]["Measure"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["Count"]));
                    Cell1.CellStyle = CountCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["Price"]));
                    Cell1.CellStyle = PriceBelCS;
                    pos++;
                }
            }
            pos++;
            FileName += " Маркетинг";
            string tempFolder = System.Environment.GetEnvironmentVariable("TEMP");
            FileInfo file = new FileInfo(tempFolder + @"\" + FileName + ".xls");
            int j = 1;
            while (file.Exists == true)
            {
                file = new FileInfo(tempFolder + @"\" + FileName + "(" + j++ + ").xls");
            }

            FileStream NewFile = new FileStream(file.FullName, FileMode.Create);
            hssfworkbook.Write(NewFile);
            NewFile.Close();

            System.Diagnostics.Process.Start(file.FullName);
            ClearReport();
        }
        
        public void CreateMarketingExpReport(DateTime date1, DateTime date2, bool IsSample, bool IsNotSample, string FileName, decimal Rate, ArrayList MClients, ArrayList MClientGroups)
        {
            ClearReport();

            string MainOrdersList = string.Empty;

            int pos = 0;

            FrontsReport.ExpReport(date1, date2, false, IsSample, IsNotSample, MClients, MClientGroups);
            DecorReport.ExpReport(date1, date2, false, IsSample, IsNotSample, MClients, MClientGroups);

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

            Collect(Rate);
            //Export to excel
            HSSFWorkbook hssfworkbook = new HSSFWorkbook();

            ////create a entry of DocumentSummaryInformation
            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = "NPOI Team";
            hssfworkbook.DocumentSummaryInformation = dsi;

            ////create a entry of SummaryInformation
            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
            si.Subject = "NPOI SDK Example";
            hssfworkbook.SummaryInformation = si;

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
            PriceBelCS.DataFormat = hssfworkbook.CreateDataFormat().GetFormat("### ### ##0.00");
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

            #endregion Create fonts and styles

            HSSFSheet sheet1 = hssfworkbook.CreateSheet("Отчет Маркетинг");
            sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

            sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
            sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
            sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
            sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

            int displayIndex = 0;

            sheet1.SetColumnWidth(displayIndex++, 12 * 256);
            sheet1.SetColumnWidth(displayIndex++, 61 * 256);
            sheet1.SetColumnWidth(displayIndex++, 15 * 256);
            sheet1.SetColumnWidth(displayIndex++, 12 * 256);
            sheet1.SetColumnWidth(displayIndex++, 8 * 256);
            sheet1.SetColumnWidth(displayIndex++, 8 * 256);
            sheet1.SetColumnWidth(displayIndex, 14 * 256);

            HSSFCell Cell1;

            if (ProfilReportTable.Rows.Count > 0)
            {
                displayIndex = 0;

                Cell1 = sheet1.CreateRow(pos++).CreateCell(0);
                Cell1.SetCellValue("ОМЦ-ПРОФИЛЬ:");
                Cell1.CellStyle = SummaryWithoutBorderBelCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Инв.№");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Наименование");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Цвет");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Патина");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Ед.изм.");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Кол-во");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Себестоимость");
                Cell1.CellStyle = SimpleHeaderCS;
                pos++;

                for (int i = 0; i < ProfilReportTable.Rows.Count; i++)
                {
                    displayIndex = 0;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(ProfilReportTable.Rows[i]["InvNumber"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(ProfilReportTable.Rows[i]["AccountingName"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(FrontsReport.GetColorNameByCode(ProfilReportTable.Rows[i]["Cvet"].ToString()));
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(FrontsReport.GetPatinaNameByCode(ProfilReportTable.Rows[i]["Patina"].ToString()));
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(ProfilReportTable.Rows[i]["Measure"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["Count"]));
                    Cell1.CellStyle = CountCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["Price"]));
                    Cell1.CellStyle = PriceBelCS;
                    pos++;
                }
            }

            if (TPSReportTable.Rows.Count > 0)
            {
                //ТПС
                pos++;
                displayIndex = 0;

                Cell1 = sheet1.CreateRow(pos++).CreateCell(0);
                Cell1.SetCellValue("ЗОВ-ТПС:");
                Cell1.CellStyle = SummaryWithoutBorderBelCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Инв.№");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Наименование");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Цвет");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Патина");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Ед.изм.");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Кол-во");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Себестоимость");
                Cell1.CellStyle = SimpleHeaderCS;
                pos++;

                for (int i = 0; i < TPSReportTable.Rows.Count; i++)
                {
                    displayIndex = 0;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(TPSReportTable.Rows[i]["InvNumber"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(TPSReportTable.Rows[i]["AccountingName"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(FrontsReport.GetColorNameByCode(TPSReportTable.Rows[i]["Cvet"].ToString()));
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(FrontsReport.GetPatinaNameByCode(TPSReportTable.Rows[i]["Patina"].ToString()));
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(TPSReportTable.Rows[i]["Measure"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["Count"]));
                    Cell1.CellStyle = CountCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["Price"]));
                    Cell1.CellStyle = PriceBelCS;
                    pos++;
                }
            }
            pos++;
            FileName += " Маркетинг";
            string tempFolder = System.Environment.GetEnvironmentVariable("TEMP");
            FileInfo file = new FileInfo(tempFolder + @"\" + FileName + ".xls");
            int j = 1;
            while (file.Exists == true)
            {
                file = new FileInfo(tempFolder + @"\" + FileName + "(" + j++ + ").xls");
            }

            FileStream NewFile = new FileStream(file.FullName, FileMode.Create);
            hssfworkbook.Write(NewFile);
            NewFile.Close();

            System.Diagnostics.Process.Start(file.FullName);
            ClearReport();
        }

        public void CreateMarketingStoreReport(DateTime date1, DateTime date2, bool IsSample, bool IsNotSample, string FileName, decimal Rate, ArrayList MClients, ArrayList MClientGroups)
        {
            ClearReport();

            string MainOrdersList = string.Empty;

            int pos = 0;

            FrontsReport.StoreReport(date1, date2, false, IsSample, IsNotSample, MClients, MClientGroups);
            DecorReport.StoreReport(date1, date2, false, IsSample, IsNotSample, MClients, MClientGroups);

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

            Collect(Rate);
            //Export to excel
            HSSFWorkbook hssfworkbook = new HSSFWorkbook();

            ////create a entry of DocumentSummaryInformation
            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = "NPOI Team";
            hssfworkbook.DocumentSummaryInformation = dsi;

            ////create a entry of SummaryInformation
            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
            si.Subject = "NPOI SDK Example";
            hssfworkbook.SummaryInformation = si;

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
            PriceBelCS.DataFormat = hssfworkbook.CreateDataFormat().GetFormat("### ### ##0.00");
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

            #endregion Create fonts and styles

            HSSFSheet sheet1 = hssfworkbook.CreateSheet("Отчет Маркетинг");
            sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

            sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
            sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
            sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
            sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

            int displayIndex = 0;

            sheet1.SetColumnWidth(displayIndex++, 12 * 256);
            sheet1.SetColumnWidth(displayIndex++, 61 * 256);
            sheet1.SetColumnWidth(displayIndex++, 15 * 256);
            sheet1.SetColumnWidth(displayIndex++, 12 * 256);
            sheet1.SetColumnWidth(displayIndex++, 8 * 256);
            sheet1.SetColumnWidth(displayIndex++, 8 * 256);
            sheet1.SetColumnWidth(displayIndex, 14 * 256);

            HSSFCell Cell1;

            if (ProfilReportTable.Rows.Count > 0)
            {
                displayIndex = 0;

                Cell1 = sheet1.CreateRow(pos++).CreateCell(0);
                Cell1.SetCellValue("ОМЦ-ПРОФИЛЬ:");
                Cell1.CellStyle = SummaryWithoutBorderBelCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Инв.№");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Наименование");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Цвет");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Патина");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Ед.изм.");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Кол-во");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Себестоимость");
                Cell1.CellStyle = SimpleHeaderCS;
                pos++;

                for (int i = 0; i < ProfilReportTable.Rows.Count; i++)
                {
                    displayIndex = 0;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(ProfilReportTable.Rows[i]["InvNumber"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(ProfilReportTable.Rows[i]["AccountingName"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(FrontsReport.GetColorNameByCode(ProfilReportTable.Rows[i]["Cvet"].ToString()));
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(FrontsReport.GetPatinaNameByCode(ProfilReportTable.Rows[i]["Patina"].ToString()));
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(ProfilReportTable.Rows[i]["Measure"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["Count"]));
                    Cell1.CellStyle = CountCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["Price"]));
                    Cell1.CellStyle = PriceBelCS;
                    pos++;
                }
            }

            if (TPSReportTable.Rows.Count > 0)
            {
                //ТПС
                pos++;
                displayIndex = 0;

                Cell1 = sheet1.CreateRow(pos++).CreateCell(0);
                Cell1.SetCellValue("ЗОВ-ТПС:");
                Cell1.CellStyle = SummaryWithoutBorderBelCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Инв.№");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Наименование");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Цвет");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Патина");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Ед.изм.");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Кол-во");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Себестоимость");
                Cell1.CellStyle = SimpleHeaderCS;
                pos++;

                for (int i = 0; i < TPSReportTable.Rows.Count; i++)
                {
                    displayIndex = 0;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(TPSReportTable.Rows[i]["InvNumber"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(TPSReportTable.Rows[i]["AccountingName"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(FrontsReport.GetColorNameByCode(TPSReportTable.Rows[i]["Cvet"].ToString()));
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(FrontsReport.GetPatinaNameByCode(TPSReportTable.Rows[i]["Patina"].ToString()));
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(TPSReportTable.Rows[i]["Measure"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["Count"]));
                    Cell1.CellStyle = CountCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["Price"]));
                    Cell1.CellStyle = PriceBelCS;
                    pos++;
                }
            }
            pos++;
            FileName += " Маркетинг";
            string tempFolder = System.Environment.GetEnvironmentVariable("TEMP");
            FileInfo file = new FileInfo(tempFolder + @"\" + FileName + ".xls");
            int j = 1;
            while (file.Exists == true)
            {
                file = new FileInfo(tempFolder + @"\" + FileName + "(" + j++ + ").xls");
            }

            FileStream NewFile = new FileStream(file.FullName, FileMode.Create);
            hssfworkbook.Write(NewFile);
            NewFile.Close();

            System.Diagnostics.Process.Start(file.FullName);
            ClearReport();
        }

        public void CreateMarketingPackReport(DateTime date1, DateTime date2, bool IsSample, bool IsNotSample, string FileName, 
            decimal Rate, ArrayList MClients, ArrayList MClientGroups)
        {
            ClearReport();

            string MainOrdersList = string.Empty;

            int pos = 0;

            FrontsReport.PackReport(date1, date2, false, IsSample, IsNotSample, MClients, MClientGroups);
            DecorReport.PackReport(date1, date2, false, IsSample, IsNotSample, MClients, MClientGroups);

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

            Collect(Rate);
            //Export to excel
            HSSFWorkbook hssfworkbook = new HSSFWorkbook();

            ////create a entry of DocumentSummaryInformation
            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = "NPOI Team";
            hssfworkbook.DocumentSummaryInformation = dsi;

            ////create a entry of SummaryInformation
            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
            si.Subject = "NPOI SDK Example";
            hssfworkbook.SummaryInformation = si;

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
            PriceBelCS.DataFormat = hssfworkbook.CreateDataFormat().GetFormat("### ### ##0.00");
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

            #endregion Create fonts and styles

            string sheetName = "Отчет Маркетинг";

            HSSFSheet sheet1 = hssfworkbook.CreateSheet(sheetName);
            sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

            sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
            sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
            sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
            sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

            int displayIndex = 0;

            sheet1.SetColumnWidth(displayIndex++, 12 * 256);
            sheet1.SetColumnWidth(displayIndex++, 61 * 256);
            sheet1.SetColumnWidth(displayIndex++, 15 * 256);
            sheet1.SetColumnWidth(displayIndex++, 12 * 256);
            sheet1.SetColumnWidth(displayIndex++, 8 * 256);
            sheet1.SetColumnWidth(displayIndex++, 8 * 256);
            sheet1.SetColumnWidth(displayIndex, 14 * 256);

            HSSFCell Cell1;

            if (ProfilReportTable.Rows.Count > 0)
            {
                displayIndex = 0;

                Cell1 = sheet1.CreateRow(pos++).CreateCell(0);
                Cell1.SetCellValue("ОМЦ-ПРОФИЛЬ:");
                Cell1.CellStyle = SummaryWithoutBorderBelCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Инв.№");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Наименование");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Цвет");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Патина");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Ед.изм.");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Кол-во");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Себестоимость");
                Cell1.CellStyle = SimpleHeaderCS;
                pos++;

                for (int i = 0; i < ProfilReportTable.Rows.Count; i++)
                {
                    displayIndex = 0;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(ProfilReportTable.Rows[i]["InvNumber"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(ProfilReportTable.Rows[i]["AccountingName"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(FrontsReport.GetColorNameByCode(ProfilReportTable.Rows[i]["Cvet"].ToString()));
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(FrontsReport.GetPatinaNameByCode(ProfilReportTable.Rows[i]["Patina"].ToString()));
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(ProfilReportTable.Rows[i]["Measure"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["Count"]));
                    Cell1.CellStyle = CountCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["Price"]));
                    Cell1.CellStyle = PriceBelCS;
                    pos++;
                }
            }

            if (TPSReportTable.Rows.Count > 0)
            {
                //ТПС
                pos++;
                displayIndex = 0;

                Cell1 = sheet1.CreateRow(pos++).CreateCell(0);
                Cell1.SetCellValue("ЗОВ-ТПС:");
                Cell1.CellStyle = SummaryWithoutBorderBelCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Инв.№");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Наименование");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Цвет");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Патина");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Ед.изм.");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Кол-во");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Себестоимость");
                Cell1.CellStyle = SimpleHeaderCS;
                pos++;

                for (int i = 0; i < TPSReportTable.Rows.Count; i++)
                {
                    displayIndex = 0;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(TPSReportTable.Rows[i]["InvNumber"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(TPSReportTable.Rows[i]["AccountingName"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(FrontsReport.GetColorNameByCode(TPSReportTable.Rows[i]["Cvet"].ToString()));
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(FrontsReport.GetPatinaNameByCode(TPSReportTable.Rows[i]["Patina"].ToString()));
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(TPSReportTable.Rows[i]["Measure"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["Count"]));
                    Cell1.CellStyle = CountCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["Price"]));
                    Cell1.CellStyle = PriceBelCS;
                    pos++;
                }
            }


            sheetName = "Отчет Маркетинг (с кодами)";

            HSSFSheet sheet2 = hssfworkbook.CreateSheet(sheetName);
            sheet2.PrintSetup.PaperSize = (short)PaperSizeType.A4;

            sheet2.SetMargin(HSSFSheet.LeftMargin, (double).12);
            sheet2.SetMargin(HSSFSheet.RightMargin, (double).07);
            sheet2.SetMargin(HSSFSheet.TopMargin, (double).20);
            sheet2.SetMargin(HSSFSheet.BottomMargin, (double).20);

            displayIndex = 0;
            pos = 0;

            sheet2.SetColumnWidth(displayIndex++, 12 * 256);
            sheet2.SetColumnWidth(displayIndex++, 61 * 256);
            sheet2.SetColumnWidth(displayIndex++, 15 * 256);
            sheet2.SetColumnWidth(displayIndex++, 12 * 256);
            sheet2.SetColumnWidth(displayIndex++, 8 * 256);
            sheet2.SetColumnWidth(displayIndex++, 8 * 256);
            sheet2.SetColumnWidth(displayIndex, 14 * 256);
            
            if (ProfilReportTable.Rows.Count > 0)
            {
                displayIndex = 0;

                Cell1 = sheet2.CreateRow(pos++).CreateCell(0);
                Cell1.SetCellValue("ОМЦ-ПРОФИЛЬ:");
                Cell1.CellStyle = SummaryWithoutBorderBelCS;

                Cell1 = sheet2.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Инв.№");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet2.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Наименование");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet2.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Цвет");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet2.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Патина");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet2.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Ед.изм.");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet2.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Кол-во");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet2.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Себестоимость");
                Cell1.CellStyle = SimpleHeaderCS;
                pos++;

                for (int i = 0; i < ProfilReportTable.Rows.Count; i++)
                {
                    displayIndex = 0;
                    Cell1 = sheet2.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(ProfilReportTable.Rows[i]["InvNumber"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet2.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(ProfilReportTable.Rows[i]["AccountingName"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet2.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(ProfilReportTable.Rows[i]["Cvet"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet2.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(ProfilReportTable.Rows[i]["Patina"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet2.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(ProfilReportTable.Rows[i]["Measure"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet2.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["Count"]));
                    Cell1.CellStyle = CountCS;
                    Cell1 = sheet2.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable.Rows[i]["Price"]));
                    Cell1.CellStyle = PriceBelCS;
                    pos++;
                }
            }

            if (TPSReportTable.Rows.Count > 0)
            {
                //ТПС
                pos++;
                displayIndex = 0;

                Cell1 = sheet2.CreateRow(pos++).CreateCell(0);
                Cell1.SetCellValue("ЗОВ-ТПС:");
                Cell1.CellStyle = SummaryWithoutBorderBelCS;

                Cell1 = sheet2.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Инв.№");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet2.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Наименование");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet2.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Цвет");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet2.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Патина");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet2.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Ед.изм.");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet2.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Кол-во");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet2.CreateRow(pos).CreateCell(displayIndex++);
                Cell1.SetCellValue("Себестоимость");
                Cell1.CellStyle = SimpleHeaderCS;
                pos++;

                for (int i = 0; i < TPSReportTable.Rows.Count; i++)
                {
                    displayIndex = 0;
                    Cell1 = sheet2.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(TPSReportTable.Rows[i]["InvNumber"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet2.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(TPSReportTable.Rows[i]["AccountingName"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet2.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(TPSReportTable.Rows[i]["Cvet"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet2.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(TPSReportTable.Rows[i]["Patina"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet2.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(TPSReportTable.Rows[i]["Measure"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet2.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["Count"]));
                    Cell1.CellStyle = CountCS;
                    Cell1 = sheet2.CreateRow(pos).CreateCell(displayIndex++);
                    Cell1.SetCellValue(Convert.ToDouble(TPSReportTable.Rows[i]["Price"]));
                    Cell1.CellStyle = PriceBelCS;
                    pos++;
                }
            }

            FileName += " Маркетинг";
            string tempFolder = System.Environment.GetEnvironmentVariable("TEMP");
            FileInfo file = new FileInfo(tempFolder + @"\" + FileName + ".xls");
            int j = 1;
            while (file.Exists == true)
            {
                file = new FileInfo(tempFolder + @"\" + FileName + "(" + j++ + ").xls");
            }

            FileStream NewFile = new FileStream(file.FullName, FileMode.Create);
            hssfworkbook.Write(NewFile);
            NewFile.Close();

            System.Diagnostics.Process.Start(file.FullName);
            ClearReport();
        }

        public void StartEndProducedReport(DateTime date1, DateTime date2, bool IsSample, bool IsNotSample, string FileName, ArrayList MClients, ArrayList MClientGroups)
        {
            ClearReport();

            string MainOrdersList = string.Empty;

            int pos = 0;
            DataTable ProfilReportTable1 = FrontsReport.AllProfilDT.Clone();
            DataTable TPSReportTable1 = FrontsReport.AllTPSDT.Clone();

            FrontsReport.StartEndProducedDispatchReport(date1, date2, IsSample, IsNotSample, MClients, MClientGroups);
            DecorReport.StartEndProducedDispatchReport(date1, date2, IsSample, IsNotSample, MClients, MClientGroups);
            //PROFIL
            if (FrontsReport.AllProfilDT.Rows.Count > 0)
            {
                for (int i = 0; i < FrontsReport.AllProfilDT.Rows.Count; i++)
                {
                    ProfilReportTable1.ImportRow(FrontsReport.AllProfilDT.Rows[i]);
                }
            }

            if (DecorReport.AllProfilDT.Rows.Count > 0)
            {
                for (int i = 0; i < DecorReport.AllProfilDT.Rows.Count; i++)
                {
                    ProfilReportTable1.ImportRow(DecorReport.AllProfilDT.Rows[i]);
                }
            }

            //TPS
            if (FrontsReport.AllTPSDT.Rows.Count > 0)
            {
                for (int i = 0; i < FrontsReport.AllTPSDT.Rows.Count; i++)
                {
                    TPSReportTable1.ImportRow(FrontsReport.AllTPSDT.Rows[i]);
                }
            }

            if (DecorReport.AllTPSDT.Rows.Count > 0)
            {
                for (int i = 0; i < DecorReport.AllTPSDT.Rows.Count; i++)
                {
                    TPSReportTable1.ImportRow(DecorReport.AllTPSDT.Rows[i]);
                }
            }

            //Export to excel
            HSSFWorkbook hssfworkbook = new HSSFWorkbook();

            ////create a entry of DocumentSummaryInformation
            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = "NPOI Team";
            hssfworkbook.DocumentSummaryInformation = dsi;

            ////create a entry of SummaryInformation
            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
            si.Subject = "NPOI SDK Example";
            hssfworkbook.SummaryInformation = si;

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

            #endregion Create fonts and styles

            HSSFSheet sheet1 = hssfworkbook.CreateSheet("ОМЦ-ПРОФИЛЬ");
            sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

            sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
            sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
            sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
            sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

            sheet1.SetColumnWidth(0, 12 * 256);
            sheet1.SetColumnWidth(1, 71 * 256);
            sheet1.SetColumnWidth(2, 8 * 256);
            sheet1.SetColumnWidth(3, 8 * 256);
            sheet1.SetColumnWidth(4, 14 * 256);

            HSSFSheet sheet2 = hssfworkbook.CreateSheet("ЗОВ-ТПС");
            sheet2.PrintSetup.PaperSize = (short)PaperSizeType.A4;

            sheet2.SetMargin(HSSFSheet.LeftMargin, (double).12);
            sheet2.SetMargin(HSSFSheet.RightMargin, (double).07);
            sheet2.SetMargin(HSSFSheet.TopMargin, (double).20);
            sheet2.SetMargin(HSSFSheet.BottomMargin, (double).20);

            sheet2.SetColumnWidth(0, 12 * 256);
            sheet2.SetColumnWidth(1, 71 * 256);
            sheet2.SetColumnWidth(2, 8 * 256);
            sheet2.SetColumnWidth(3, 8 * 256);
            sheet2.SetColumnWidth(4, 14 * 256);

            HSSFCell Cell1;

            if (ProfilReportTable1.Rows.Count > 0)
            {
                Cell1 = sheet1.CreateRow(pos++).CreateCell(0);
                Cell1.SetCellValue("ОМЦ-ПРОФИЛЬ:");
                Cell1.CellStyle = SummaryWithoutBorderBelCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                Cell1.SetCellValue("Инв.№");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(1);
                Cell1.SetCellValue("Наименование");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(2);
                Cell1.SetCellValue("Ед.изм.");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(3);
                Cell1.SetCellValue("ОстН");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(4);
                Cell1.SetCellValue("Приход");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(5);
                Cell1.SetCellValue("Расход");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet1.CreateRow(pos).CreateCell(6);
                Cell1.SetCellValue("ОстК");
                Cell1.CellStyle = SimpleHeaderCS;

                pos++;

                for (int i = 0; i < ProfilReportTable1.Rows.Count; i++)
                {
                    Cell1 = sheet1.CreateRow(pos).CreateCell(0);
                    Cell1.SetCellValue(ProfilReportTable1.Rows[i]["InvNumber"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(1);
                    Cell1.SetCellValue(ProfilReportTable1.Rows[i]["AccountingName"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(2);
                    Cell1.SetCellValue(ProfilReportTable1.Rows[i]["Measure"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(3);
                    Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable1.Rows[i]["StartCount"]));
                    Cell1.CellStyle = CountCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(4);
                    Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable1.Rows[i]["ProducedCount"]));
                    Cell1.CellStyle = CountCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(5);
                    Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable1.Rows[i]["DispatchCount"]));
                    Cell1.CellStyle = CountCS;
                    Cell1 = sheet1.CreateRow(pos).CreateCell(6);
                    Cell1.SetCellValue(Convert.ToDouble(ProfilReportTable1.Rows[i]["EndCount"]));
                    Cell1.CellStyle = CountCS;
                    pos++;
                }
            }
            pos = 0;
            if (TPSReportTable1.Rows.Count > 0)
            {
                //ТПС
                pos++;

                Cell1 = sheet2.CreateRow(pos++).CreateCell(0);
                Cell1.SetCellValue("ЗОВ-ТПС:");
                Cell1.CellStyle = SummaryWithoutBorderBelCS;

                Cell1 = sheet2.CreateRow(pos).CreateCell(0);
                Cell1.SetCellValue("Инв.№");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet2.CreateRow(pos).CreateCell(1);
                Cell1.SetCellValue("Наименование");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet2.CreateRow(pos).CreateCell(2);
                Cell1.SetCellValue("Ед.изм.");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet2.CreateRow(pos).CreateCell(3);
                Cell1.SetCellValue("ОстН");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet2.CreateRow(pos).CreateCell(4);
                Cell1.SetCellValue("Приход");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet2.CreateRow(pos).CreateCell(5);
                Cell1.SetCellValue("Расход");
                Cell1.CellStyle = SimpleHeaderCS;

                Cell1 = sheet2.CreateRow(pos).CreateCell(6);
                Cell1.SetCellValue("ОстК");
                Cell1.CellStyle = SimpleHeaderCS;
                pos++;

                for (int i = 0; i < TPSReportTable1.Rows.Count; i++)
                {
                    Cell1 = sheet2.CreateRow(pos).CreateCell(0);
                    Cell1.SetCellValue(TPSReportTable1.Rows[i]["InvNumber"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet2.CreateRow(pos).CreateCell(1);
                    Cell1.SetCellValue(TPSReportTable1.Rows[i]["AccountingName"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet2.CreateRow(pos).CreateCell(2);
                    Cell1.SetCellValue(TPSReportTable1.Rows[i]["Measure"].ToString());
                    Cell1.CellStyle = SimpleCS;
                    Cell1 = sheet2.CreateRow(pos).CreateCell(3);
                    Cell1.SetCellValue(Convert.ToDouble(TPSReportTable1.Rows[i]["StartCount"]));
                    Cell1.CellStyle = CountCS;
                    Cell1 = sheet2.CreateRow(pos).CreateCell(4);
                    Cell1.SetCellValue(Convert.ToDouble(TPSReportTable1.Rows[i]["ProducedCount"]));
                    Cell1.CellStyle = CountCS;
                    Cell1 = sheet2.CreateRow(pos).CreateCell(5);
                    Cell1.SetCellValue(Convert.ToDouble(TPSReportTable1.Rows[i]["DispatchCount"]));
                    Cell1.CellStyle = CountCS;
                    Cell1 = sheet2.CreateRow(pos).CreateCell(6);
                    Cell1.SetCellValue(Convert.ToDouble(TPSReportTable1.Rows[i]["EndCount"]));
                    Cell1.CellStyle = CountCS;
                    pos++;
                }
            }
            pos++;
            FileName += " ЗОВ";
            string tempFolder = System.Environment.GetEnvironmentVariable("TEMP");
            FileInfo file = new FileInfo(tempFolder + @"\" + FileName + ".xls");
            int j = 1;
            while (file.Exists == true)
            {
                file = new FileInfo(tempFolder + @"\" + FileName + "(" + j++ + ").xls");
            }

            FileStream NewFile = new FileStream(file.FullName, FileMode.Create);
            hssfworkbook.Write(NewFile);
            NewFile.Close();

            System.Diagnostics.Process.Start(file.FullName);
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