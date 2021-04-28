using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;

namespace Infinium.Modules.StatisticsMarketing
{
    public class MarketingExpeditionStatistics
    {
        private DataTable MFSummaryDT = null;
        private DataTable MDSummaryDT = null;

        private DataTable MarketingClientsDT = null;

        private DataTable MExpFrontsCostDT = null;
        private DataTable MExpDecorCostDT = null;
        private DataTable MAllFrontsCostDT = null;
        private DataTable MAllDecorCostDT = null;

        private DataTable FrontsOrdersDT = null;
        private DataTable DecorOrdersDT = null;

        private DataTable TempFrontsDT = null;
        private DataTable TempDecorDT = null;

        private DataTable FrontsSummaryDT = null;
        private DataTable DecorProductsSummaryDT = null;
        private DataTable DecorItemsSummaryDT = null;
        private DataTable DecorConfigDT = null;

        private DataTable FrontsDT = null;
        private DataTable DecorProductsDT = null;
        private DataTable DecorItemsDT = null;

        public BindingSource MFSummaryBS = null;
        public BindingSource MDSummaryBS = null;

        public BindingSource FrontsSummaryBS = null;
        public BindingSource DecorProductsSummaryBS = null;
        public BindingSource DecorItemsSummaryBS = null;

        private PercentageDataGrid MFSummaryDG = null;
        private PercentageDataGrid MDSummaryDG = null;
        private PercentageDataGrid FrontsDG = null;
        private PercentageDataGrid DecorProductsDG = null;
        private PercentageDataGrid DecorItemsDG = null;

        public MarketingExpeditionStatistics(
            ref PercentageDataGrid tMFSummaryDG,
            ref PercentageDataGrid tMDSummaryDG,
            ref PercentageDataGrid tFrontsDG,
            ref PercentageDataGrid tDecorProductsDG,
            ref PercentageDataGrid tDecorItemsDG)
        {
            MFSummaryDG = tMFSummaryDG;
            MDSummaryDG = tMDSummaryDG;
            FrontsDG = tFrontsDG;
            DecorProductsDG = tDecorProductsDG;
            DecorItemsDG = tDecorItemsDG;

            Initialize();
        }

        private void Create()
        {
            MarketingClientsDT = new DataTable();

            MExpFrontsCostDT = new DataTable();
            MExpDecorCostDT = new DataTable();
            MAllFrontsCostDT = new DataTable();
            MAllDecorCostDT = new DataTable();

            TempFrontsDT = new DataTable();
            TempDecorDT = new DataTable();

            DecorItemsDT = new DataTable();
            FrontsDT = new DataTable();
            DecorProductsDT = new DataTable();
            DecorConfigDT = new DataTable();

            FrontsOrdersDT = new DataTable();
            DecorOrdersDT = new DataTable();

            MFSummaryDT = new DataTable();
            MFSummaryDT.Columns.Add(new DataColumn(("ClientID"), System.Type.GetType("System.Int32")));
            MFSummaryDT.Columns.Add(new DataColumn(("ClientName"), System.Type.GetType("System.String")));
            MFSummaryDT.Columns.Add(new DataColumn(("MegaOrderID"), System.Type.GetType("System.Int32")));
            MFSummaryDT.Columns.Add(new DataColumn(("OrderNumber"), System.Type.GetType("System.Int32")));
            MFSummaryDT.Columns.Add(new DataColumn(("DocDateTime"), System.Type.GetType("System.DateTime")));
            MFSummaryDT.Columns.Add(new DataColumn(("ExpPercTPS"), System.Type.GetType("System.Decimal")));
            MFSummaryDT.Columns.Add(new DataColumn(("ExpDateTPS"), System.Type.GetType("System.DateTime")));
            MFSummaryDT.Columns.Add(new DataColumn(("ExpPercProfil"), System.Type.GetType("System.Decimal")));
            MFSummaryDT.Columns.Add(new DataColumn(("ExpDateProfil"), System.Type.GetType("System.DateTime")));
            MFSummaryDT.Columns.Add(new DataColumn(("CostTPS"), System.Type.GetType("System.Decimal")));
            MFSummaryDT.Columns.Add(new DataColumn(("CostProfil"), System.Type.GetType("System.Decimal")));
            MFSummaryDT.Columns.Add(new DataColumn(("AllCostTPS"), System.Type.GetType("System.Decimal")));
            MFSummaryDT.Columns.Add(new DataColumn(("AllCostProfil"), System.Type.GetType("System.Decimal")));

            MDSummaryDT = new DataTable();
            MDSummaryDT.Columns.Add(new DataColumn(("ClientID"), System.Type.GetType("System.Int32")));
            MDSummaryDT.Columns.Add(new DataColumn(("ClientName"), System.Type.GetType("System.String")));
            MDSummaryDT.Columns.Add(new DataColumn(("MegaOrderID"), System.Type.GetType("System.Int32")));
            MDSummaryDT.Columns.Add(new DataColumn(("OrderNumber"), System.Type.GetType("System.Int32")));
            MDSummaryDT.Columns.Add(new DataColumn(("DocDateTime"), System.Type.GetType("System.DateTime")));
            MDSummaryDT.Columns.Add(new DataColumn(("ExpPercTPS"), System.Type.GetType("System.Decimal")));
            MDSummaryDT.Columns.Add(new DataColumn(("ExpDateTPS"), System.Type.GetType("System.DateTime")));
            MDSummaryDT.Columns.Add(new DataColumn(("ExpPercProfil"), System.Type.GetType("System.Decimal")));
            MDSummaryDT.Columns.Add(new DataColumn(("ExpDateProfil"), System.Type.GetType("System.DateTime")));
            MDSummaryDT.Columns.Add(new DataColumn(("CostTPS"), System.Type.GetType("System.Decimal")));
            MDSummaryDT.Columns.Add(new DataColumn(("CostProfil"), System.Type.GetType("System.Decimal")));
            MDSummaryDT.Columns.Add(new DataColumn(("AllCostTPS"), System.Type.GetType("System.Decimal")));
            MDSummaryDT.Columns.Add(new DataColumn(("AllCostProfil"), System.Type.GetType("System.Decimal")));

            FrontsSummaryDT = new DataTable();
            FrontsSummaryDT.Columns.Add(new DataColumn(("Front"), System.Type.GetType("System.String")));
            FrontsSummaryDT.Columns.Add(new DataColumn(("FrontID"), System.Type.GetType("System.Int32")));
            FrontsSummaryDT.Columns.Add(new DataColumn(("Width"), System.Type.GetType("System.Int32")));
            FrontsSummaryDT.Columns.Add(new DataColumn(("Square"), System.Type.GetType("System.Decimal")));
            FrontsSummaryDT.Columns.Add(new DataColumn(("Cost"), Type.GetType("System.Decimal")));
            FrontsSummaryDT.Columns.Add(new DataColumn(("Count"), System.Type.GetType("System.Int32")));

            DecorProductsSummaryDT = new DataTable();
            DecorProductsSummaryDT.Columns.Add(new DataColumn(("DecorProduct"), System.Type.GetType("System.String")));
            DecorProductsSummaryDT.Columns.Add(new DataColumn(("ProductID"), System.Type.GetType("System.Int32")));
            DecorProductsSummaryDT.Columns.Add(new DataColumn(("MeasureID"), System.Type.GetType("System.Int32")));
            DecorProductsSummaryDT.Columns.Add(new DataColumn(("Cost"), Type.GetType("System.Decimal")));
            DecorProductsSummaryDT.Columns.Add(new DataColumn(("Count"), System.Type.GetType("System.Decimal")));
            DecorProductsSummaryDT.Columns.Add(new DataColumn(("Measure"), System.Type.GetType("System.String")));

            DecorItemsSummaryDT = new DataTable();
            DecorItemsSummaryDT.Columns.Add(new DataColumn(("DecorItem"), System.Type.GetType("System.String")));
            DecorItemsSummaryDT.Columns.Add(new DataColumn(("ProductID"), System.Type.GetType("System.Int32")));
            DecorItemsSummaryDT.Columns.Add(new DataColumn(("MeasureID"), System.Type.GetType("System.Int32")));
            DecorItemsSummaryDT.Columns.Add(new DataColumn(("DecorID"), System.Type.GetType("System.Int32")));
            DecorItemsSummaryDT.Columns.Add(new DataColumn(("Cost"), Type.GetType("System.Decimal")));
            DecorItemsSummaryDT.Columns.Add(new DataColumn(("Count"), System.Type.GetType("System.Decimal")));
            DecorItemsSummaryDT.Columns.Add(new DataColumn(("Measure"), System.Type.GetType("System.String")));
        }

        private void Fill()
        {
            using (SqlDataAdapter DA = new SqlDataAdapter(
            @"SELECT MegaOrders.ClientID, infiniu2_marketingreference.dbo.Clients.ClientName, MegaOrders.OrderNumber, MegaOrders.MegaOrderID
            FROM MegaOrders
            INNER JOIN infiniu2_marketingreference.dbo.Clients ON MegaOrders.ClientID = infiniu2_marketingreference.dbo.Clients.ClientID
            WHERE MegaOrders.MegaOrderID IN (SELECT MegaOrderID FROM MainOrders WHERE MainOrderID IN (SELECT MainOrderID FROM Packages WHERE PackageStatusID IN (4)))
            ORDER BY infiniu2_marketingreference.dbo.Clients.ClientName, MegaOrders.OrderNumber, MegaOrders.MegaOrderID",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(MarketingClientsDT);
            }

            string SelectCommand = @"SELECT TechStoreID AS FrontID, TechStoreName AS FrontName FROM TechStore
                WHERE TechStoreID IN (SELECT FrontID FROM FrontsConfig WHERE Enabled = 1)
                ORDER BY TechStoreName";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(FrontsDT);
            }

            SelectCommand = @"SELECT ProductID, ProductName FROM DecorProducts" +
                " WHERE ProductID IN (SELECT ProductID FROM DecorConfig WHERE Enabled = 1) ORDER BY ProductName ASC";
            DecorProductsDT = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(DecorProductsDT);
            }
            DecorItemsDT = new DataTable();
            SelectCommand = @"SELECT DISTINCT TechStore.TechStoreID AS DecorID, TechStore.TechStoreName AS Name, DecorConfig.ProductID FROM TechStore
                INNER JOIN DecorConfig ON TechStore.TechStoreID = DecorConfig.DecorID AND Enabled = 1 ORDER BY TechStoreName";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(DecorItemsDT);
            }

            //using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM DecorConfig", ConnectionStrings.CatalogConnectionString))
            //{
            //    DA.Fill(DecorConfigDT);
            //}
            DecorConfigDT = TablesManager.DecorConfigDataTableAll;

            System.Diagnostics.Stopwatch sw = new System.Diagnostics.Stopwatch();
            sw.Start();

            FillMarketingTables();

            sw.Stop();
            double G = sw.Elapsed.TotalMilliseconds;

            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT TOP 0 FrontID, PatinaID, ColorID, InsetTypeID," +
                " InsetColorID, Height, Width, Count, Square, Cost FROM FrontsOrders",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                FrontsOrdersDT.Clear();
                DA.Fill(FrontsOrdersDT);
            }

            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT TOP 0 FrontID, PatinaID, ColorID, InsetTypeID," +
                " InsetColorID, Height, Width, Count, Square, Cost FROM FrontsOrders",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                TempFrontsDT.Clear();
                DA.Fill(TempFrontsDT);
            }

            //decor
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT TOP 0 DecorOrders.ProductID, DecorOrders.DecorID, DecorOrders.ColorID,DecorOrders.PatinaID," +
                " DecorOrders.Height, DecorOrders.Length, DecorOrders.Width, DecorOrders.Count, DecorOrders.Cost, DecorOrders.DecorConfigID, " +
                " MeasureID FROM DecorOrders" +
                " INNER JOIN infiniu2_catalog.dbo.DecorConfig" +
                " ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID", ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(TempDecorDT);
            }

            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT TOP 0 DecorOrders.ProductID, DecorOrders.DecorID, DecorOrders.ColorID,DecorOrders.PatinaID," +
                " DecorOrders.Height, DecorOrders.Length, DecorOrders.Width, DecorOrders.Count, DecorOrders.Cost, DecorOrders.DecorConfigID, " +
                " MeasureID FROM DecorOrders" +
                " INNER JOIN infiniu2_catalog.dbo.DecorConfig" +
                " ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID", ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(DecorOrdersDT);
            }
        }

        private void FillMarketingTables()
        {
            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT MegaOrders.ClientID, MainOrders.MegaOrderID, FrontsOrders.FactoryID, SUM(FrontsOrders.Cost * PackageDetails.Count / FrontsOrders.Count) AS FrontsCost
            FROM PackageDetails INNER JOIN
            FrontsOrders ON PackageDetails.OrderID = FrontsOrders.FrontsOrdersID INNER JOIN
            MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID INNER JOIN
            MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID
            WHERE PackageDetails.PackageID IN (SELECT PackageID FROM Packages WHERE ProductType = 0 AND PackageStatusID IN (4))
            GROUP BY MegaOrders.ClientID, MainOrders.MegaOrderID, FrontsOrders.FactoryID
            ORDER BY MegaOrders.ClientID, MainOrders.MegaOrderID, FrontsOrders.FactoryID",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                MExpFrontsCostDT.Clear();
                DA.Fill(MExpFrontsCostDT);
            }

            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT MegaOrders.ClientID, MainOrders.MegaOrderID, DecorOrders.FactoryID, SUM(DecorOrders.Cost * PackageDetails.Count / DecorOrders.Count) AS DecorCost
            FROM PackageDetails INNER JOIN
            DecorOrders ON PackageDetails.OrderID = DecorOrders.DecorOrderID INNER JOIN
            MainOrders ON DecorOrders.MainOrderID = MainOrders.MainOrderID INNER JOIN
            MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID
            WHERE PackageDetails.PackageID IN (SELECT PackageID FROM Packages WHERE ProductType = 1 AND PackageStatusID IN (4))
            GROUP BY MegaOrders.ClientID, MainOrders.MegaOrderID, DecorOrders.FactoryID
            ORDER BY MegaOrders.ClientID, MainOrders.MegaOrderID, DecorOrders.FactoryID",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                MExpDecorCostDT.Clear();
                DA.Fill(MExpDecorCostDT);
            }

            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT MegaOrders.ClientID, MainOrders.MegaOrderID, FrontsOrders.FactoryID, SUM(FrontsOrders.Cost) AS FrontsCost
            FROM FrontsOrders INNER JOIN
            MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID INNER JOIN
            MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID
            WHERE FrontsOrders.FrontsOrdersID IN
            (SELECT OrderID FROM PackageDetails WHERE PackageID IN (SELECT PackageID FROM Packages WHERE ProductType = 0))
            GROUP BY MegaOrders.ClientID, MainOrders.MegaOrderID, FrontsOrders.FactoryID
            ORDER BY MegaOrders.ClientID, MainOrders.MegaOrderID, FrontsOrders.FactoryID",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                MAllFrontsCostDT.Clear();
                DA.Fill(MAllFrontsCostDT);
            }

            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT MegaOrders.ClientID, MainOrders.MegaOrderID, DecorOrders.FactoryID, SUM(DecorOrders.Cost) AS DecorCost
            FROM DecorOrders INNER JOIN
            MainOrders ON DecorOrders.MainOrderID = MainOrders.MainOrderID INNER JOIN
            MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID
            WHERE DecorOrders.DecorOrderID IN
            (SELECT OrderID FROM PackageDetails WHERE PackageID IN (SELECT PackageID FROM Packages WHERE ProductType = 1))
            GROUP BY MegaOrders.ClientID, MainOrders.MegaOrderID, DecorOrders.FactoryID
            ORDER BY MegaOrders.ClientID, MainOrders.MegaOrderID, DecorOrders.FactoryID",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                MAllDecorCostDT.Clear();
                DA.Fill(MAllDecorCostDT);
            }
        }

        private void Binding()
        {
            MFSummaryBS = new BindingSource()
            {
                DataSource = MFSummaryDT
            };
            MFSummaryDG.DataSource = MFSummaryBS;

            MDSummaryBS = new BindingSource()
            {
                DataSource = MDSummaryDT
            };
            MDSummaryDG.DataSource = MDSummaryBS;

            FrontsSummaryBS = new BindingSource()
            {
                DataSource = FrontsSummaryDT
            };
            FrontsDG.DataSource = FrontsSummaryBS;

            DecorProductsSummaryBS = new BindingSource()
            {
                DataSource = DecorProductsSummaryDT
            };
            DecorProductsDG.DataSource = DecorProductsSummaryBS;

            DecorItemsSummaryBS = new BindingSource()
            {
                DataSource = DecorItemsSummaryDT
            };
            DecorItemsDG.DataSource = DecorItemsSummaryBS;
        }

        public void Initialize()
        {
            Create();
            Fill();
            Binding();
            SetProductsGrids();
            MarketSummaryGridSettings();
        }

        public void ShowColumns(ref PercentageDataGrid FrontsGrid, ref PercentageDataGrid DecorGrid, bool Profil, bool TPS, bool bClientSummary)
        {
            if (Profil && TPS)
            {
                FrontsGrid.Columns["CostProfil"].Visible = true;
                FrontsGrid.Columns["CostTPS"].Visible = true;
                FrontsGrid.Columns["ExpPercProfil"].Visible = true;
                FrontsGrid.Columns["ExpPercTPS"].Visible = true;
                if (FrontsGrid.Columns.Contains("ExpDateTPS"))
                    FrontsGrid.Columns["ExpDateTPS"].Visible = true;
                if (FrontsGrid.Columns.Contains("ExpDateProfil"))
                    FrontsGrid.Columns["ExpDateProfil"].Visible = true;

                DecorGrid.Columns["CostProfil"].Visible = true;
                DecorGrid.Columns["CostTPS"].Visible = true;
                DecorGrid.Columns["ExpPercProfil"].Visible = true;
                DecorGrid.Columns["ExpPercTPS"].Visible = true;
                if (DecorGrid.Columns.Contains("ExpDateTPS"))
                    DecorGrid.Columns["ExpDateTPS"].Visible = true;
                if (DecorGrid.Columns.Contains("ExpDateProfil"))
                    DecorGrid.Columns["ExpDateProfil"].Visible = true;
            }
            if (Profil && !TPS)
            {
                FrontsGrid.Columns["CostProfil"].Visible = true;
                FrontsGrid.Columns["CostTPS"].Visible = false;
                FrontsGrid.Columns["ExpPercProfil"].Visible = true;
                FrontsGrid.Columns["ExpPercTPS"].Visible = false;
                if (FrontsGrid.Columns.Contains("ExpDateTPS"))
                    FrontsGrid.Columns["ExpDateTPS"].Visible = false;
                if (FrontsGrid.Columns.Contains("ExpDateProfil"))
                    FrontsGrid.Columns["ExpDateProfil"].Visible = true;

                DecorGrid.Columns["CostProfil"].Visible = true;
                DecorGrid.Columns["CostTPS"].Visible = false;
                DecorGrid.Columns["ExpPercProfil"].Visible = true;
                DecorGrid.Columns["ExpPercTPS"].Visible = false;
                if (DecorGrid.Columns.Contains("ExpDateTPS"))
                    DecorGrid.Columns["ExpDateTPS"].Visible = false;
                if (DecorGrid.Columns.Contains("ExpDateProfil"))
                    DecorGrid.Columns["ExpDateProfil"].Visible = true;
            }
            if (!Profil && TPS)
            {
                FrontsGrid.Columns["CostProfil"].Visible = false;
                FrontsGrid.Columns["CostTPS"].Visible = true;
                FrontsGrid.Columns["ExpPercProfil"].Visible = false;
                FrontsGrid.Columns["ExpPercTPS"].Visible = true;
                if (FrontsGrid.Columns.Contains("ExpDateTPS"))
                    FrontsGrid.Columns["ExpDateTPS"].Visible = true;
                if (FrontsGrid.Columns.Contains("ExpDateProfil"))
                    FrontsGrid.Columns["ExpDateProfil"].Visible = false;

                DecorGrid.Columns["CostProfil"].Visible = false;
                DecorGrid.Columns["CostTPS"].Visible = true;
                DecorGrid.Columns["ExpPercProfil"].Visible = false;
                DecorGrid.Columns["ExpPercTPS"].Visible = true;
                if (DecorGrid.Columns.Contains("ExpDateTPS"))
                    DecorGrid.Columns["ExpDateTPS"].Visible = true;
                if (DecorGrid.Columns.Contains("ExpDateProfil"))
                    DecorGrid.Columns["ExpDateProfil"].Visible = false;
            }
            if (bClientSummary)
            {
                if (FrontsGrid.Columns.Contains("ExpDateTPS"))
                    FrontsGrid.Columns["ExpDateTPS"].Visible = false;
                if (FrontsGrid.Columns.Contains("ExpDateProfil"))
                    FrontsGrid.Columns["ExpDateProfil"].Visible = false;
                if (DecorGrid.Columns.Contains("ExpDateTPS"))
                    DecorGrid.Columns["ExpDateTPS"].Visible = false;
                if (DecorGrid.Columns.Contains("ExpDateProfil"))
                    DecorGrid.Columns["ExpDateProfil"].Visible = false;
            }
        }

        private void MarketSummaryGridSettings()
        {
            MFSummaryDG.Columns["MegaOrderID"].Visible = false;
            MFSummaryDG.Columns["ClientID"].Visible = false;
            MDSummaryDG.Columns["MegaOrderID"].Visible = false;
            MDSummaryDG.Columns["ClientID"].Visible = false;

            MFSummaryDG.Columns["AllCostProfil"].Visible = false;
            MDSummaryDG.Columns["AllCostProfil"].Visible = false;
            MFSummaryDG.Columns["AllCostTPS"].Visible = false;
            MDSummaryDG.Columns["AllCostTPS"].Visible = false;

            if (!Security.PriceAccess)
            {
                MFSummaryDG.Columns["CostProfil"].Visible = false;
                MFSummaryDG.Columns["CostTPS"].Visible = false;
                MDSummaryDG.Columns["CostProfil"].Visible = false;
                MDSummaryDG.Columns["CostTPS"].Visible = false;
            }

            NumberFormatInfo nfi1 = new NumberFormatInfo()
            {
                CurrencyGroupSeparator = " ",
                CurrencySymbol = string.Empty,
                CurrencyDecimalDigits = 2,
                CurrencyDecimalSeparator = ".",

                NumberGroupSeparator = " ",
                NumberDecimalDigits = 2,
                NumberDecimalSeparator = ","
            };
            foreach (DataGridViewColumn Column in MFSummaryDG.Columns)
            {
                Column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
            foreach (DataGridViewColumn Column in MDSummaryDG.Columns)
            {
                Column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }

            MFSummaryDG.Columns["DocDateTime"].DefaultCellStyle.Format = "dd.MM.yyyy";
            MFSummaryDG.Columns["ExpDateTPS"].DefaultCellStyle.Format = "dd.MM.yyyy HH:mm";
            MFSummaryDG.Columns["ExpDateProfil"].DefaultCellStyle.Format = "dd.MM.yyyy HH:mm";
            MFSummaryDG.Columns["DocDateTime"].HeaderText = "Дата создания";
            MFSummaryDG.Columns["ClientName"].HeaderText = "Клиент";
            MFSummaryDG.Columns["OrderNumber"].HeaderText = "№ заказа";
            MFSummaryDG.Columns["ExpPercProfil"].HeaderText = "Экспедиция\n\r Профиль, %";
            MFSummaryDG.Columns["ExpPercTPS"].HeaderText = "Экспедиция\n\r      ТПС, %";
            MFSummaryDG.Columns["CostProfil"].HeaderText = "Профиль, €";
            MFSummaryDG.Columns["CostTPS"].HeaderText = "ТПС, €";
            MFSummaryDG.Columns["ExpDateProfil"].HeaderText = "Профиль, вход\n\r   на эксп-цию";
            MFSummaryDG.Columns["ExpDateTPS"].HeaderText = "  ТПС, вход\n\rна эксп-цию";
            MFSummaryDG.Columns["DocDateTime"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            MFSummaryDG.Columns["DocDateTime"].MinimumWidth = 110;
            MFSummaryDG.Columns["ClientName"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            MFSummaryDG.Columns["ClientName"].MinimumWidth = 110;
            MFSummaryDG.Columns["OrderNumber"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MFSummaryDG.Columns["OrderNumber"].Width = 85;
            MFSummaryDG.Columns["ExpPercProfil"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            MFSummaryDG.Columns["ExpPercProfil"].MinimumWidth = 135;
            MFSummaryDG.Columns["ExpPercTPS"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            MFSummaryDG.Columns["ExpPercTPS"].MinimumWidth = 135;
            MFSummaryDG.Columns["CostProfil"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MFSummaryDG.Columns["CostProfil"].Width = 110;
            MFSummaryDG.Columns["CostTPS"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MFSummaryDG.Columns["CostTPS"].Width = 110;
            MFSummaryDG.Columns["ExpDateTPS"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            MFSummaryDG.Columns["ExpDateTPS"].MinimumWidth = 110;
            MFSummaryDG.Columns["ExpDateProfil"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            MFSummaryDG.Columns["ExpDateProfil"].MinimumWidth = 110;

            //MFSummaryDG.Columns["AllCostProfil"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            //MFSummaryDG.Columns["AllCostProfil"].Width = 110;
            //MFSummaryDG.Columns["AllCostTPS"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            //MFSummaryDG.Columns["AllCostTPS"].Width = 110;

            MFSummaryDG.Columns["CostProfil"].DefaultCellStyle.Format = "N";
            MFSummaryDG.Columns["CostProfil"].DefaultCellStyle.FormatProvider = nfi1;
            MFSummaryDG.Columns["CostTPS"].DefaultCellStyle.Format = "N";
            MFSummaryDG.Columns["CostTPS"].DefaultCellStyle.FormatProvider = nfi1;

            MFSummaryDG.Columns["ExpPercProfil"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            MFSummaryDG.AddPercentageColumn("ExpPercProfil");
            MFSummaryDG.Columns["ExpPercTPS"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            MFSummaryDG.AddPercentageColumn("ExpPercTPS");

            MDSummaryDG.Columns["DocDateTime"].DefaultCellStyle.Format = "dd.MM.yyyy";
            MDSummaryDG.Columns["ExpDateTPS"].DefaultCellStyle.Format = "dd.MM.yyyy HH:mm";
            MDSummaryDG.Columns["ExpDateProfil"].DefaultCellStyle.Format = "dd.MM.yyyy HH:mm";
            MDSummaryDG.Columns["DocDateTime"].HeaderText = "Дата создания";
            MDSummaryDG.Columns["ClientName"].HeaderText = "Клиент";
            MDSummaryDG.Columns["OrderNumber"].HeaderText = "№ заказа";
            MDSummaryDG.Columns["ExpPercProfil"].HeaderText = "Экспедиция\n\r Профиль, %";
            MDSummaryDG.Columns["ExpPercTPS"].HeaderText = "Экспедиция\n\r     ТПС, %";
            MDSummaryDG.Columns["CostProfil"].HeaderText = "Профиль, €";
            MDSummaryDG.Columns["CostTPS"].HeaderText = "ТПС, €";
            MDSummaryDG.Columns["ExpDateProfil"].HeaderText = "Профиль, вход\n\r   на эксп-цию";
            MDSummaryDG.Columns["ExpDateTPS"].HeaderText = "  ТПС, вход\n\rна эксп-цию";
            MDSummaryDG.Columns["DocDateTime"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            MDSummaryDG.Columns["DocDateTime"].MinimumWidth = 110;
            MDSummaryDG.Columns["ClientName"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            MDSummaryDG.Columns["ClientName"].MinimumWidth = 110;
            MDSummaryDG.Columns["OrderNumber"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MDSummaryDG.Columns["OrderNumber"].Width = 85;
            MDSummaryDG.Columns["ExpPercProfil"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            MDSummaryDG.Columns["ExpPercProfil"].MinimumWidth = 135;
            MDSummaryDG.Columns["ExpPercTPS"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            MDSummaryDG.Columns["ExpPercTPS"].MinimumWidth = 135;
            MDSummaryDG.Columns["CostProfil"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MDSummaryDG.Columns["CostProfil"].Width = 110;
            MDSummaryDG.Columns["CostTPS"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MDSummaryDG.Columns["CostTPS"].Width = 110;
            MDSummaryDG.Columns["ExpDateTPS"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            MDSummaryDG.Columns["ExpDateTPS"].MinimumWidth = 110;
            MDSummaryDG.Columns["ExpDateProfil"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            MDSummaryDG.Columns["ExpDateProfil"].MinimumWidth = 110;

            //MDSummaryDG.Columns["AllCostProfil"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            //MDSummaryDG.Columns["AllCostProfil"].Width = 110;
            //MDSummaryDG.Columns["AllCostTPS"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            //MDSummaryDG.Columns["AllCostTPS"].Width = 110;

            MDSummaryDG.Columns["CostProfil"].DefaultCellStyle.Format = "N";
            MDSummaryDG.Columns["CostProfil"].DefaultCellStyle.FormatProvider = nfi1;
            MDSummaryDG.Columns["CostTPS"].DefaultCellStyle.Format = "N";
            MDSummaryDG.Columns["CostTPS"].DefaultCellStyle.FormatProvider = nfi1;

            MDSummaryDG.Columns["ExpPercProfil"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            MDSummaryDG.AddPercentageColumn("ExpPercProfil");
            MDSummaryDG.Columns["ExpPercTPS"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            MDSummaryDG.AddPercentageColumn("ExpPercTPS");
        }

        private void SetProductsGrids()
        {
            if (!Security.PriceAccess)
            {
                FrontsDG.Columns["Cost"].Visible = false;
                DecorProductsDG.Columns["Cost"].Visible = false;
                DecorItemsDG.Columns["Cost"].Visible = false;
            }

            NumberFormatInfo nfi1 = new NumberFormatInfo()
            {
                CurrencyGroupSeparator = " ",
                CurrencySymbol = string.Empty,
                CurrencyDecimalDigits = 2,
                CurrencyDecimalSeparator = ".",

                NumberGroupSeparator = " ",
                NumberDecimalDigits = 2,
                NumberDecimalSeparator = ","
            };
            FrontsDG.ColumnHeadersHeight = 38;
            DecorProductsDG.ColumnHeadersHeight = 38;
            DecorItemsDG.ColumnHeadersHeight = 38;

            foreach (DataGridViewColumn Column in FrontsDG.Columns)
            {
                Column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
            foreach (DataGridViewColumn Column in DecorProductsDG.Columns)
            {
                Column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
            foreach (DataGridViewColumn Column in DecorItemsDG.Columns)
            {
                Column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }

            FrontsDG.Columns["FrontID"].Visible = false;
            FrontsDG.Columns["Width"].Visible = false;

            FrontsDG.Columns["Front"].HeaderText = "Фасад";
            FrontsDG.Columns["Cost"].HeaderText = " € ";
            FrontsDG.Columns["Square"].HeaderText = "м.кв.";
            FrontsDG.Columns["Count"].HeaderText = "шт.";

            FrontsDG.Columns["Front"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            FrontsDG.Columns["Front"].MinimumWidth = 110;
            FrontsDG.Columns["Square"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            FrontsDG.Columns["Count"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            FrontsDG.Columns["Cost"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            FrontsDG.Columns["Square"].Width = 100;
            FrontsDG.Columns["Cost"].Width = 100;
            FrontsDG.Columns["Count"].Width = 90;

            FrontsDG.Columns["Cost"].DefaultCellStyle.Format = "N";
            FrontsDG.Columns["Cost"].DefaultCellStyle.FormatProvider = nfi1;
            FrontsDG.Columns["Square"].DefaultCellStyle.Format = "N";
            FrontsDG.Columns["Square"].DefaultCellStyle.FormatProvider = nfi1;

            DecorProductsDG.Columns["ProductID"].Visible = false;
            DecorProductsDG.Columns["MeasureID"].Visible = false;

            DecorItemsDG.Columns["ProductID"].Visible = false;
            DecorItemsDG.Columns["DecorID"].Visible = false;
            DecorItemsDG.Columns["MeasureID"].Visible = false;

            DecorProductsDG.Columns["DecorProduct"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            DecorProductsDG.Columns["DecorProduct"].MinimumWidth = 100;
            DecorProductsDG.Columns["Cost"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            DecorProductsDG.Columns["Cost"].Width = 100;
            DecorProductsDG.Columns["Count"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            DecorProductsDG.Columns["Count"].Width = 100;
            DecorProductsDG.Columns["Measure"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            DecorProductsDG.Columns["Measure"].Width = 90;

            DecorProductsDG.Columns["DecorProduct"].HeaderText = "Продукт";
            DecorProductsDG.Columns["Cost"].HeaderText = " € ";
            DecorProductsDG.Columns["Count"].HeaderText = "Кол-во";
            DecorProductsDG.Columns["Measure"].HeaderText = "Ед.изм.";

            DecorProductsDG.Columns["Cost"].DefaultCellStyle.Format = "N";
            DecorProductsDG.Columns["Cost"].DefaultCellStyle.FormatProvider = nfi1;
            DecorProductsDG.Columns["Count"].DefaultCellStyle.Format = "N";
            DecorProductsDG.Columns["Count"].DefaultCellStyle.FormatProvider = nfi1;

            DecorItemsDG.Columns["DecorID"].Visible = false;

            DecorItemsDG.Columns["DecorItem"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            DecorItemsDG.Columns["DecorItem"].MinimumWidth = 100;
            DecorItemsDG.Columns["Cost"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            DecorItemsDG.Columns["Cost"].Width = 100;
            DecorItemsDG.Columns["Count"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            DecorItemsDG.Columns["Count"].Width = 100;
            DecorItemsDG.Columns["Measure"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            DecorItemsDG.Columns["Measure"].Width = 90;

            DecorItemsDG.Columns["DecorItem"].HeaderText = "Наименование";
            DecorItemsDG.Columns["Cost"].HeaderText = " € ";
            DecorItemsDG.Columns["Count"].HeaderText = "Кол-во";
            DecorItemsDG.Columns["Measure"].HeaderText = "Ед.изм.";

            DecorItemsDG.Columns["Cost"].DefaultCellStyle.Format = "N";
            DecorItemsDG.Columns["Cost"].DefaultCellStyle.FormatProvider = nfi1;
            DecorItemsDG.Columns["Count"].DefaultCellStyle.Format = "N";
            DecorItemsDG.Columns["Count"].DefaultCellStyle.FormatProvider = nfi1;

            FrontsDG.Columns["Square"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            FrontsDG.Columns["Cost"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            FrontsDG.Columns["Count"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            DecorProductsDG.Columns["Cost"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            DecorProductsDG.Columns["Count"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            DecorItemsDG.Columns["Cost"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            DecorItemsDG.Columns["Count"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        }

        public void ClientSummary(int FactoryID)
        {
            int ClientID = 0;
            int MegaOrderID = 0;

            decimal ExpPercProfil = 0;
            decimal ExpPercTPS = 0;

            decimal ExpCostProfil = 0;
            decimal ExpCostTPS = 0;

            decimal AllCostProfil = 0;
            decimal AllCostTPS = 0;

            decimal Percentage = 0;
            decimal d1 = 0;
            decimal d2 = 0;

            string ClientName = string.Empty;

            MFSummaryDT.Clear();
            MDSummaryDT.Clear();

            if (FactoryID == -1)
                return;

            DataTable Table = new DataTable();

            using (DataView DV = new DataView(MarketingClientsDT))
            {
                Table = DV.ToTable(true, new string[] { "ClientID" });
            }

            for (int i = 0; i < MarketingClientsDT.Rows.Count; i++)
            {
                ClientID = Convert.ToInt32(MarketingClientsDT.Rows[i]["ClientID"]);
                MegaOrderID = Convert.ToInt32(MarketingClientsDT.Rows[i]["MegaOrderID"]);

                ClientName = MarketingClientsDT.Rows[i]["ClientName"].ToString();

                if (FactoryID == 0)
                {
                    ExpCostProfil = 0;
                    ExpCostTPS = 0;
                    ExpPercProfil = MarketExpFrontsCost(MegaOrderID, 1, ref ExpCostProfil, ref AllCostProfil);
                    ExpPercTPS = MarketExpFrontsCost(MegaOrderID, 2, ref ExpCostTPS, ref AllCostTPS);

                    d1 = 0;
                    d2 = 0;

                    if (ExpCostProfil != 0 || ExpCostTPS != 0)
                    {
                        DataRow[] Rows = MFSummaryDT.Select("ClientName='" + ClientName + "'");
                        if (Rows.Count() == 0)
                        {
                            DataRow NewRow = MFSummaryDT.NewRow();
                            NewRow["ClientID"] = ClientID;
                            NewRow["ClientName"] = ClientName;
                            if (ExpCostProfil > 0)
                            {
                                NewRow["ExpPercProfil"] = ExpPercProfil;
                                NewRow["CostProfil"] = ExpCostProfil;
                                NewRow["AllCostProfil"] = AllCostProfil;
                            }
                            if (ExpCostTPS > 0)
                            {
                                NewRow["ExpPercTPS"] = ExpPercTPS;
                                NewRow["CostTPS"] = ExpCostTPS;
                                NewRow["AllCostTPS"] = AllCostTPS;
                            }
                            MFSummaryDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            if (ExpCostProfil > 0)
                            {
                                if (Rows[0]["CostProfil"] != DBNull.Value)
                                    d1 = Convert.ToDecimal(Rows[0]["CostProfil"]);
                                if (Rows[0]["AllCostProfil"] != DBNull.Value)
                                    d2 = Convert.ToDecimal(Rows[0]["AllCostProfil"]);
                                Percentage = (d1 + ExpCostProfil) / (d2 + AllCostProfil);
                                Percentage = Decimal.Round(Percentage * 100, 2, MidpointRounding.AwayFromZero);
                                Percentage = (d1 + ExpCostProfil) / (d2 + AllCostProfil);
                                Percentage = Decimal.Round(Percentage * 100, 2, MidpointRounding.AwayFromZero);
                                Rows[0]["ExpPercProfil"] = Percentage;
                                Rows[0]["AllCostProfil"] = d2 + AllCostProfil;
                                Rows[0]["CostProfil"] = d1 + ExpCostProfil;
                            }
                            if (ExpCostTPS > 0)
                            {
                                if (Rows[0]["CostTPS"] != DBNull.Value)
                                    d1 = Convert.ToDecimal(Rows[0]["CostTPS"]);
                                if (Rows[0]["AllCostTPS"] != DBNull.Value)
                                    d2 = Convert.ToDecimal(Rows[0]["AllCostTPS"]);
                                Percentage = (d1 + ExpCostTPS) / (d2 + AllCostTPS);
                                Percentage = Decimal.Round(Percentage * 100, 2, MidpointRounding.AwayFromZero);
                                Percentage = (d1 + ExpCostTPS) / (d2 + AllCostTPS);
                                Percentage = Decimal.Round(Percentage * 100, 2, MidpointRounding.AwayFromZero);
                                Rows[0]["ExpPercTPS"] = Percentage;
                                Rows[0]["AllCostTPS"] = d2 + AllCostTPS;
                                Rows[0]["CostTPS"] = d1 + ExpCostTPS;
                            }
                        }
                    }

                    ExpCostProfil = 0;
                    ExpCostTPS = 0;
                    ExpPercProfil = MarketExpDecorCost(MegaOrderID, 1, ref ExpCostProfil, ref AllCostProfil);
                    ExpPercTPS = MarketExpDecorCost(MegaOrderID, 2, ref ExpCostTPS, ref AllCostTPS);

                    if (ExpCostProfil != 0 || ExpCostTPS != 0)
                    {
                        DataRow[] Rows = MDSummaryDT.Select("ClientName='" + ClientName + "'");
                        if (Rows.Count() == 0)
                        {
                            DataRow NewRow = MDSummaryDT.NewRow();
                            NewRow["ClientID"] = ClientID;
                            NewRow["ClientName"] = ClientName;
                            if (ExpCostProfil > 0)
                            {
                                NewRow["ExpPercProfil"] = ExpPercProfil;
                                NewRow["CostProfil"] = ExpCostProfil;
                                NewRow["AllCostProfil"] = AllCostProfil;
                            }
                            if (ExpCostTPS > 0)
                            {
                                NewRow["ExpPercTPS"] = ExpPercTPS;
                                NewRow["CostTPS"] = ExpCostTPS;
                                NewRow["AllCostTPS"] = AllCostTPS;
                            }
                            MDSummaryDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            if (ExpCostProfil > 0)
                            {
                                if (Rows[0]["CostProfil"] != DBNull.Value)
                                    d1 = Convert.ToDecimal(Rows[0]["CostProfil"]);
                                if (Rows[0]["AllCostProfil"] != DBNull.Value)
                                    d2 = Convert.ToDecimal(Rows[0]["AllCostProfil"]);
                                Percentage = (d1 + ExpCostProfil) / (d2 + AllCostProfil);
                                Percentage = Decimal.Round(Percentage * 100, 2, MidpointRounding.AwayFromZero);
                                Percentage = (d1 + ExpCostProfil) / (d2 + AllCostProfil);
                                Percentage = Decimal.Round(Percentage * 100, 2, MidpointRounding.AwayFromZero);
                                Rows[0]["ExpPercProfil"] = Percentage;
                                Rows[0]["AllCostProfil"] = d2 + AllCostProfil;
                                Rows[0]["CostProfil"] = d1 + ExpCostProfil;
                            }
                            if (ExpCostTPS > 0)
                            {
                                if (Rows[0]["CostTPS"] != DBNull.Value)
                                    d1 = Convert.ToDecimal(Rows[0]["CostTPS"]);
                                if (Rows[0]["AllCostTPS"] != DBNull.Value)
                                    d2 = Convert.ToDecimal(Rows[0]["AllCostTPS"]);
                                Percentage = (d1 + ExpCostTPS) / (d2 + AllCostTPS);
                                Percentage = Decimal.Round(Percentage * 100, 2, MidpointRounding.AwayFromZero);
                                Percentage = (d1 + ExpCostTPS) / (d2 + AllCostTPS);
                                Percentage = Decimal.Round(Percentage * 100, 2, MidpointRounding.AwayFromZero);
                                Rows[0]["ExpPercTPS"] = Percentage;
                                Rows[0]["AllCostTPS"] = d2 + AllCostTPS;
                                Rows[0]["CostTPS"] = d1 + ExpCostTPS;
                            }
                        }
                    }
                }

                if (FactoryID == 1)
                {
                    ExpCostProfil = 0;
                    ExpPercProfil = MarketExpFrontsCost(MegaOrderID, 1, ref ExpCostProfil, ref AllCostProfil);

                    if (ExpCostProfil > 0)
                    {
                        DataRow[] Rows = MFSummaryDT.Select("ClientName='" + ClientName + "'");
                        if (Rows.Count() == 0)
                        {
                            DataRow NewRow = MFSummaryDT.NewRow();
                            NewRow["ClientID"] = ClientID;
                            NewRow["ClientName"] = ClientName;
                            NewRow["ExpPercProfil"] = ExpPercProfil;
                            NewRow["CostProfil"] = ExpCostProfil;
                            NewRow["AllCostProfil"] = AllCostProfil;
                            MFSummaryDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            if (Rows[0]["CostProfil"] != DBNull.Value)
                                d1 = Convert.ToDecimal(Rows[0]["CostProfil"]);
                            if (Rows[0]["AllCostProfil"] != DBNull.Value)
                                d2 = Convert.ToDecimal(Rows[0]["AllCostProfil"]);
                            Percentage = (d1 + ExpCostProfil) / (d2 + AllCostProfil);
                            Percentage = Decimal.Round(Percentage * 100, 2, MidpointRounding.AwayFromZero);
                            Percentage = (d1 + ExpCostProfil) / (d2 + AllCostProfil);
                            Percentage = Decimal.Round(Percentage * 100, 2, MidpointRounding.AwayFromZero);
                            Rows[0]["ExpPercProfil"] = Percentage;
                            Rows[0]["AllCostProfil"] = d2 + AllCostProfil;
                            Rows[0]["CostProfil"] = d1 + ExpCostProfil;
                        }
                    }

                    ExpCostProfil = 0;
                    ExpCostProfil = MarketExpDecorCost(MegaOrderID, 1, ref ExpCostProfil, ref AllCostProfil);

                    if (ExpCostProfil > 0)
                    {
                        DataRow[] Rows = MDSummaryDT.Select("ClientName='" + ClientName + "'");
                        if (Rows.Count() == 0)
                        {
                            DataRow NewRow = MDSummaryDT.NewRow();
                            NewRow["ClientID"] = ClientID;
                            NewRow["ClientName"] = ClientName;
                            NewRow["ExpPercProfil"] = ExpPercProfil;
                            NewRow["CostProfil"] = ExpCostProfil;
                            NewRow["AllCostProfil"] = AllCostProfil;
                            MDSummaryDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            if (Rows[0]["CostProfil"] != DBNull.Value)
                                d1 = Convert.ToDecimal(Rows[0]["CostProfil"]);
                            if (Rows[0]["AllCostProfil"] != DBNull.Value)
                                d2 = Convert.ToDecimal(Rows[0]["AllCostProfil"]);
                            Percentage = (d1 + ExpCostProfil) / (d2 + AllCostProfil);
                            Percentage = Decimal.Round(Percentage * 100, 2, MidpointRounding.AwayFromZero);
                            Percentage = (d1 + ExpCostProfil) / (d2 + AllCostProfil);
                            Percentage = Decimal.Round(Percentage * 100, 2, MidpointRounding.AwayFromZero);
                            Rows[0]["ExpPercProfil"] = Percentage;
                            Rows[0]["AllCostProfil"] = d2 + AllCostProfil;
                            Rows[0]["CostProfil"] = d1 + ExpCostProfil;
                        }
                    }
                }

                if (FactoryID == 2)
                {
                    ExpCostTPS = 0;
                    ExpPercTPS = MarketExpFrontsCost(MegaOrderID, 2, ref ExpCostTPS, ref AllCostTPS);

                    if (ExpCostTPS > 0)
                    {
                        DataRow[] Rows = MFSummaryDT.Select("ClientName='" + ClientName + "'");
                        if (Rows.Count() == 0)
                        {
                            DataRow NewRow = MFSummaryDT.NewRow();
                            NewRow["ClientID"] = ClientID;
                            NewRow["ClientName"] = ClientName;
                            NewRow["ExpPercTPS"] = ExpPercTPS;
                            NewRow["CostTPS"] = ExpCostTPS;
                            NewRow["AllCostTPS"] = AllCostTPS;
                            MFSummaryDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            if (Rows[0]["CostTPS"] != DBNull.Value)
                                d1 = Convert.ToDecimal(Rows[0]["CostTPS"]);
                            if (Rows[0]["AllCostTPS"] != DBNull.Value)
                                d2 = Convert.ToDecimal(Rows[0]["AllCostTPS"]);
                            Percentage = (d1 + ExpCostTPS) / (d2 + AllCostTPS);
                            Percentage = Decimal.Round(Percentage * 100, 2, MidpointRounding.AwayFromZero);
                            Percentage = (d1 + ExpCostTPS) / (d2 + AllCostTPS);
                            Percentage = Decimal.Round(Percentage * 100, 2, MidpointRounding.AwayFromZero);
                            Rows[0]["ExpPercTPS"] = Percentage;
                            Rows[0]["AllCostTPS"] = d2 + AllCostTPS;
                            Rows[0]["CostTPS"] = d1 + ExpCostTPS;
                        }
                    }

                    ExpCostTPS = 0;
                    ExpPercTPS = MarketExpDecorCost(MegaOrderID, 2, ref ExpCostTPS, ref AllCostTPS);

                    if (ExpCostTPS > 0)
                    {
                        DataRow[] Rows = MDSummaryDT.Select("ClientName='" + ClientName + "'");
                        if (Rows.Count() == 0)
                        {
                            DataRow NewRow = MDSummaryDT.NewRow();
                            NewRow["ClientID"] = ClientID;
                            NewRow["ClientName"] = ClientName;
                            NewRow["ExpPercTPS"] = ExpPercTPS;
                            NewRow["CostTPS"] = ExpCostTPS;
                            NewRow["AllCostTPS"] = AllCostTPS;
                            MDSummaryDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            if (Rows[0]["CostTPS"] != DBNull.Value)
                                d1 = Convert.ToDecimal(Rows[0]["CostTPS"]);
                            if (Rows[0]["AllCostTPS"] != DBNull.Value)
                                d2 = Convert.ToDecimal(Rows[0]["AllCostTPS"]);
                            Percentage = (d1 + ExpCostTPS) / (d2 + AllCostTPS);
                            Percentage = Decimal.Round(Percentage * 100, 2, MidpointRounding.AwayFromZero);
                            Percentage = (d1 + ExpCostTPS) / (d2 + AllCostTPS);
                            Percentage = Decimal.Round(Percentage * 100, 2, MidpointRounding.AwayFromZero);
                            Rows[0]["ExpPercTPS"] = Percentage;
                            Rows[0]["AllCostTPS"] = d2 + AllCostTPS;
                            Rows[0]["CostTPS"] = d1 + ExpCostTPS;
                        }
                    }
                }
            }

            MFSummaryDT.DefaultView.Sort = "ClientName";
            MDSummaryDT.DefaultView.Sort = "ClientName";
            MFSummaryBS.MoveFirst();
            MDSummaryBS.MoveFirst();
        }

        public void MarketingSummary(int FactoryID)
        {
            int ClientID = 0;
            int MegaOrderID = 0;
            int OrderNumber = 0;

            decimal ExpPercProfil = 0;
            decimal ExpPercTPS = 0;

            decimal ExpCostProfil = 0;
            decimal ExpCostTPS = 0;

            decimal AllCostProfil = 0;
            decimal AllCostTPS = 0;

            string ClientName = string.Empty;

            object DocDateTime = null;
            object ExpDateTPS = null;
            object ExpDateProfil = null;

            MFSummaryDT.Clear();
            MDSummaryDT.Clear();

            if (FactoryID == -1)
                return;

            for (int i = 0; i < MarketingClientsDT.Rows.Count; i++)
            {
                ClientID = Convert.ToInt32(MarketingClientsDT.Rows[i]["ClientID"]);
                MegaOrderID = Convert.ToInt32(MarketingClientsDT.Rows[i]["MegaOrderID"]);
                OrderNumber = Convert.ToInt32(MarketingClientsDT.Rows[i]["OrderNumber"]);
                ClientName = MarketingClientsDT.Rows[i]["ClientName"].ToString();
                DocDateTime = GetCreationDate(MegaOrderID);

                if (FactoryID == 0)
                {
                    ExpCostProfil = 0;
                    ExpCostTPS = 0;
                    ExpPercProfil = MarketExpFrontsCost(MegaOrderID, 1, ref ExpCostProfil, ref AllCostProfil);
                    ExpPercTPS = MarketExpFrontsCost(MegaOrderID, 2, ref ExpCostTPS, ref AllCostTPS);

                    if (ExpCostProfil != 0 || ExpCostTPS != 0)
                    {
                        DataRow NewRow = MFSummaryDT.NewRow();
                        NewRow["DocDateTime"] = DocDateTime;
                        NewRow["ClientID"] = ClientID;
                        NewRow["ClientName"] = ClientName;
                        NewRow["OrderNumber"] = OrderNumber;
                        NewRow["MegaOrderID"] = MegaOrderID;
                        if (ExpCostProfil > 0)
                        {
                            NewRow["ExpPercProfil"] = ExpPercProfil;
                            NewRow["CostProfil"] = ExpCostProfil;
                            ExpDateProfil = StorageAdoptionDate(MegaOrderID, 1, 0);
                            if (ExpDateProfil != null)
                                NewRow["ExpDateProfil"] = ExpDateProfil;
                        }
                        if (ExpCostTPS > 0)
                        {
                            NewRow["ExpPercTPS"] = ExpPercTPS;
                            NewRow["CostTPS"] = ExpCostTPS;
                            ExpDateTPS = StorageAdoptionDate(MegaOrderID, 2, 0);
                            if (ExpDateTPS != null)
                                NewRow["ExpDateTPS"] = ExpDateTPS;
                        }
                        MFSummaryDT.Rows.Add(NewRow);
                    }

                    ExpCostProfil = 0;
                    ExpCostTPS = 0;
                    ExpPercProfil = MarketExpDecorCost(MegaOrderID, 1, ref ExpCostProfil, ref AllCostProfil);
                    ExpPercTPS = MarketExpDecorCost(MegaOrderID, 2, ref ExpCostTPS, ref AllCostTPS);

                    if (ExpCostProfil != 0 || ExpCostTPS != 0)
                    {
                        DataRow NewRow = MDSummaryDT.NewRow();
                        NewRow["DocDateTime"] = DocDateTime;
                        NewRow["ClientID"] = ClientID;
                        NewRow["ClientName"] = ClientName;
                        NewRow["OrderNumber"] = OrderNumber;
                        NewRow["MegaOrderID"] = MegaOrderID;
                        if (ExpCostProfil > 0)
                        {
                            NewRow["ExpPercProfil"] = ExpPercProfil;
                            NewRow["CostProfil"] = ExpCostProfil;
                            ExpDateProfil = StorageAdoptionDate(MegaOrderID, 1, 1);
                            if (ExpDateProfil != null)
                                NewRow["ExpDateProfil"] = ExpDateProfil;
                        }
                        if (ExpCostTPS > 0)
                        {
                            NewRow["ExpPercTPS"] = ExpPercTPS;
                            NewRow["CostTPS"] = ExpCostTPS;
                            ExpDateTPS = StorageAdoptionDate(MegaOrderID, 2, 1);
                            if (ExpDateTPS != null)
                                NewRow["ExpDateTPS"] = ExpDateTPS;
                        }
                        MDSummaryDT.Rows.Add(NewRow);
                    }
                }

                if (FactoryID == 1)
                {
                    ExpCostProfil = 0;
                    ExpPercProfil = MarketExpFrontsCost(MegaOrderID, 1, ref ExpCostProfil, ref AllCostProfil);

                    if (ExpCostProfil > 0)
                    {
                        DataRow NewRow = MFSummaryDT.NewRow();
                        NewRow["DocDateTime"] = DocDateTime;
                        NewRow["ClientID"] = ClientID;
                        NewRow["ClientName"] = ClientName;
                        NewRow["OrderNumber"] = OrderNumber;
                        NewRow["MegaOrderID"] = MegaOrderID;
                        NewRow["ExpPercProfil"] = ExpPercProfil;
                        NewRow["CostProfil"] = ExpCostProfil;
                        ExpDateProfil = StorageAdoptionDate(MegaOrderID, 1, 0);
                        if (ExpDateProfil != null)
                            NewRow["ExpDateProfil"] = ExpDateProfil;
                        MFSummaryDT.Rows.Add(NewRow);
                    }

                    ExpCostProfil = 0;
                    ExpCostProfil = MarketExpDecorCost(MegaOrderID, 1, ref ExpCostProfil, ref AllCostProfil);

                    if (ExpCostProfil > 0)
                    {
                        DataRow NewRow = MDSummaryDT.NewRow();
                        NewRow["DocDateTime"] = DocDateTime;
                        NewRow["ClientName"] = ClientID;
                        NewRow["OrderNumber"] = OrderNumber;
                        NewRow["MegaOrderID"] = MegaOrderID;
                        NewRow["ExpPercProfil"] = ExpPercProfil;
                        NewRow["CostProfil"] = ExpCostProfil;
                        ExpDateProfil = StorageAdoptionDate(MegaOrderID, 1, 1);
                        if (ExpDateProfil != null)
                            NewRow["ExpDateProfil"] = ExpDateProfil;
                        MDSummaryDT.Rows.Add(NewRow);
                    }
                }

                if (FactoryID == 2)
                {
                    ExpCostTPS = 0;
                    ExpPercTPS = MarketExpFrontsCost(MegaOrderID, 2, ref ExpCostTPS, ref AllCostTPS);

                    if (ExpCostTPS > 0)
                    {
                        DataRow NewRow = MFSummaryDT.NewRow();
                        NewRow["DocDateTime"] = DocDateTime;
                        NewRow["ClientID"] = ClientID;
                        NewRow["ClientName"] = ClientName;
                        NewRow["OrderNumber"] = OrderNumber;
                        NewRow["MegaOrderID"] = MegaOrderID;
                        NewRow["ExpPercTPS"] = ExpPercTPS;
                        NewRow["CostTPS"] = ExpCostTPS;
                        ExpDateTPS = StorageAdoptionDate(MegaOrderID, 2, 0);
                        if (ExpDateTPS != null)
                            NewRow["ExpDateTPS"] = ExpDateTPS;
                        MFSummaryDT.Rows.Add(NewRow);
                    }

                    ExpCostTPS = 0;
                    ExpPercTPS = MarketExpDecorCost(MegaOrderID, 2, ref ExpCostTPS, ref AllCostTPS);

                    if (ExpCostTPS > 0)
                    {
                        DataRow NewRow = MDSummaryDT.NewRow();
                        NewRow["DocDateTime"] = DocDateTime;
                        NewRow["ClientID"] = ClientID;
                        NewRow["ClientName"] = ClientName;
                        NewRow["OrderNumber"] = OrderNumber;
                        NewRow["MegaOrderID"] = MegaOrderID;
                        NewRow["ExpPercTPS"] = ExpPercTPS;
                        NewRow["CostTPS"] = ExpCostTPS;
                        ExpDateTPS = StorageAdoptionDate(MegaOrderID, 2, 1);
                        if (ExpDateTPS != null)
                            NewRow["ExpDateTPS"] = ExpDateTPS;
                        MDSummaryDT.Rows.Add(NewRow);
                    }
                }
            }

            MFSummaryDT.DefaultView.Sort = "ClientName";
            MDSummaryDT.DefaultView.Sort = "ClientName";
            MFSummaryBS.MoveFirst();
            MDSummaryBS.MoveFirst();
        }

        public void FMarketingOrders1(int FactoryID, int MegaOrderID)
        {
            string MarketingSelectCommand = string.Empty;

            string MFrontsPackageFilter = string.Empty;

            string MegaOrderFilter = string.Empty;
            string PackageFactoryFilter = string.Empty;

            if (MegaOrderID != -1)
                MegaOrderFilter = " AND MegaOrders.MegaOrderID = " + MegaOrderID;
            if (FactoryID != 0)
                PackageFactoryFilter = " AND FactoryID = " + FactoryID;

            MFrontsPackageFilter = " PackageID IN" +
                " (SELECT PackageID FROM Packages WHERE PackageStatusID IN (4) AND ProductType = 0 " + PackageFactoryFilter + ")";

            MarketingSelectCommand = "SELECT FrontsOrders.FrontID, FrontsOrders.PatinaID, FrontsOrders.ColorID, FrontsOrders.InsetTypeID," +
                " FrontsOrders.InsetColorID, FrontsOrders.TechnoColorID, FrontsOrders.TechnoInsetTypeID, FrontsOrders.TechnoInsetColorID, FrontsOrders.Height, FrontsOrders.Width, PackageDetails.Count, (FrontsOrders.Square * PackageDetails.Count / FrontsOrders.Count) AS Square, MeasureID, (FrontsOrders.Cost * PackageDetails.Count / FrontsOrders.Count) AS Cost, ClientID FROM PackageDetails" +
                " INNER JOIN FrontsOrders ON PackageDetails.OrderID = FrontsOrders.FrontsOrdersID" +
                " INNER JOIN MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID" +
                " INNER JOIN MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID" + MegaOrderFilter +
                " INNER JOIN infiniu2_catalog.dbo.FrontsConfig" +
                " ON FrontsOrders.FrontConfigID=infiniu2_catalog.dbo.FrontsConfig.FrontConfigID" +
                " WHERE " + MFrontsPackageFilter;

            using (SqlDataAdapter DA = new SqlDataAdapter(MarketingSelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                FrontsOrdersDT.Clear();
                DA.Fill(FrontsOrdersDT);
            }

            GetFronts();
        }

        public void FMarketingOrders(int FactoryID, int ClientID)
        {
            string MarketingSelectCommand = string.Empty;

            string MFrontsPackageFilter = string.Empty;

            string ClientFilter = string.Empty;
            string PackageFactoryFilter = string.Empty;

            if (ClientID != -1)
                ClientFilter = " AND ClientID = " + ClientID;
            if (FactoryID != 0)
                PackageFactoryFilter = " AND FactoryID = " + FactoryID;

            MFrontsPackageFilter = " PackageID IN" +
                " (SELECT PackageID FROM Packages WHERE PackageStatusID IN (4) AND ProductType = 0 " + PackageFactoryFilter + ")";

            MarketingSelectCommand = "SELECT FrontsOrders.FrontID, FrontsOrders.PatinaID, FrontsOrders.ColorID, FrontsOrders.InsetTypeID," +
                " FrontsOrders.InsetColorID, FrontsOrders.TechnoColorID, FrontsOrders.TechnoInsetTypeID, FrontsOrders.TechnoInsetColorID, FrontsOrders.Height, FrontsOrders.Width, PackageDetails.Count, (FrontsOrders.Square * PackageDetails.Count / FrontsOrders.Count) AS Square, MeasureID, (FrontsOrders.Cost * PackageDetails.Count / FrontsOrders.Count) AS Cost, ClientID FROM PackageDetails" +
                " INNER JOIN FrontsOrders ON PackageDetails.OrderID = FrontsOrders.FrontsOrdersID" +
                " INNER JOIN MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID" +
                " INNER JOIN MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID" + ClientFilter +
                " INNER JOIN infiniu2_catalog.dbo.FrontsConfig" +
                " ON FrontsOrders.FrontConfigID=infiniu2_catalog.dbo.FrontsConfig.FrontConfigID" +
                " WHERE " + MFrontsPackageFilter;

            using (SqlDataAdapter DA = new SqlDataAdapter(MarketingSelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                FrontsOrdersDT.Clear();
                DA.Fill(FrontsOrdersDT);
            }

            GetFronts();
        }

        public void DMarketingOrders1(int FactoryID, int MegaOrderID)
        {
            string MarketingSelectCommand = string.Empty;

            string MDecorPackageFilter = string.Empty;

            string MegaOrderFilter = string.Empty;
            string PackageFactoryFilter = string.Empty;

            if (MegaOrderID != -1)
                MegaOrderFilter = " AND MegaOrders.MegaOrderID = " + MegaOrderID;
            if (FactoryID != 0)
                PackageFactoryFilter = " AND FactoryID = " + FactoryID;

            MDecorPackageFilter = " PackageID IN" +
                " (SELECT PackageID FROM Packages WHERE PackageStatusID IN (4) AND ProductType = 1 " + PackageFactoryFilter + ")";

            //decor
            MarketingSelectCommand = "SELECT DecorOrders.ProductID, DecorOrders.DecorID, DecorOrders.ColorID,DecorOrders.PatinaID," +
                " DecorOrders.Height, DecorOrders.Length, DecorOrders.Width, PackageDetails.Count, (DecorOrders.Cost * PackageDetails.Count / DecorOrders.Count) AS Cost, DecorOrders.DecorConfigID, " +
                " MeasureID, ClientID FROM PackageDetails" +
                " INNER JOIN DecorOrders ON PackageDetails.OrderID = DecorOrders.DecorOrderID" +
                " INNER JOIN MainOrders ON DecorOrders.MainOrderID = MainOrders.MainOrderID" +
                " INNER JOIN MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID" + MegaOrderFilter +
                " INNER JOIN infiniu2_catalog.dbo.DecorConfig" +
                " ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                " WHERE " + MDecorPackageFilter;

            using (SqlDataAdapter DA = new SqlDataAdapter(MarketingSelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DecorOrdersDT.Clear();
                DA.Fill(DecorOrdersDT);
            }

            GetDecorProducts();
            GetDecorItems();
        }

        public void DMarketingOrders(int FactoryID, int ClientID)
        {
            string MarketingSelectCommand = string.Empty;

            string MDecorPackageFilter = string.Empty;

            string ClientFilter = string.Empty;
            string PackageFactoryFilter = string.Empty;

            if (ClientID != -1)
                ClientFilter = " AND ClientID = " + ClientID;
            if (FactoryID != 0)
                PackageFactoryFilter = " AND FactoryID = " + FactoryID;

            MDecorPackageFilter = " PackageID IN" +
                " (SELECT PackageID FROM Packages WHERE PackageStatusID IN (4) AND ProductType = 1 " + PackageFactoryFilter + ")";

            //decor
            MarketingSelectCommand = "SELECT DecorOrders.ProductID, DecorOrders.DecorID, DecorOrders.ColorID,DecorOrders.PatinaID," +
                " DecorOrders.Height, DecorOrders.Length, DecorOrders.Width, PackageDetails.Count, (DecorOrders.Cost * PackageDetails.Count / DecorOrders.Count) AS Cost, DecorOrders.DecorConfigID, " +
                " MeasureID, ClientID FROM PackageDetails" +
                " INNER JOIN DecorOrders ON PackageDetails.OrderID = DecorOrders.DecorOrderID" +
                " INNER JOIN MainOrders ON DecorOrders.MainOrderID = MainOrders.MainOrderID" +
                " INNER JOIN MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID" + ClientFilter +
                " INNER JOIN infiniu2_catalog.dbo.DecorConfig" +
                " ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                " WHERE " + MDecorPackageFilter;

            using (SqlDataAdapter DA = new SqlDataAdapter(MarketingSelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DecorOrdersDT.Clear();
                DA.Fill(DecorOrdersDT);
            }

            GetDecorProducts();
            GetDecorItems();
        }

        public bool HasFronts
        {
            get { return FrontsOrdersDT.Rows.Count > 0; }
        }

        public bool HasDecor
        {
            get { return DecorOrdersDT.Rows.Count > 0; }
        }

        public void FilterDecorProducts(int ProductID, int MeasureID)
        {
            DecorItemsSummaryBS.Filter = "ProductID=" + ProductID + " AND MeasureID=" + MeasureID;
            DecorItemsSummaryBS.MoveFirst();
        }

        private void GetFronts()
        {
            decimal FrontCost = 0;
            decimal FrontSquare = 0;
            int FrontCount = 0;

            FrontsSummaryDT.Clear();

            DataTable Table = new DataTable();

            using (DataView DV = new DataView(FrontsOrdersDT))
            {
                Table = DV.ToTable(true, new string[] { "FrontID" });
            }

            for (int i = 0; i < Table.Rows.Count; i++)
            {
                DataRow[] Rows = FrontsOrdersDT.Select("FrontID=" + Convert.ToInt32(Table.Rows[i]["FrontID"]) +
                    " AND Width<>-1");
                if (Rows.Count() != 0)
                {
                    foreach (DataRow row in Rows)
                    {
                        FrontCost += Convert.ToDecimal(row["Cost"]);
                        FrontSquare += Convert.ToDecimal(row["Square"]);
                        FrontCount += Convert.ToInt32(row["Count"]);
                    }

                    DataRow NewRow = FrontsSummaryDT.NewRow();
                    NewRow["Front"] = GetFrontName(Convert.ToInt32(Table.Rows[i]["FrontID"]));
                    NewRow["FrontID"] = Convert.ToInt32(Table.Rows[i]["FrontID"]);
                    NewRow["Square"] = Decimal.Round(FrontSquare, 2, MidpointRounding.AwayFromZero);
                    NewRow["Cost"] = Decimal.Round(FrontCost, 2, MidpointRounding.AwayFromZero);
                    NewRow["Count"] = FrontCount;
                    NewRow["Width"] = 0;
                    FrontsSummaryDT.Rows.Add(NewRow);

                    FrontCost = 0;
                    FrontSquare = 0;
                    FrontCount = 0;
                }
                DataRow[] CurvedRows = FrontsOrdersDT.Select("FrontID=" + Convert.ToInt32(Table.Rows[i]["FrontID"]) +
                    " AND Width=-1");
                if (CurvedRows.Count() != 0)
                {
                    foreach (DataRow row in CurvedRows)
                    {
                        FrontCost += Convert.ToDecimal(row["Cost"]);
                        FrontCount += Convert.ToInt32(row["Count"]);
                    }

                    DataRow CurvedNewRow = FrontsSummaryDT.NewRow();
                    CurvedNewRow["Front"] = GetFrontName(Convert.ToInt32(Table.Rows[i]["FrontID"])) + " гнутый";
                    CurvedNewRow["FrontID"] = Convert.ToInt32(Table.Rows[i]["FrontID"]);
                    CurvedNewRow["Width"] = -1;
                    CurvedNewRow["Cost"] = Decimal.Round(FrontCost, 2, MidpointRounding.AwayFromZero);
                    CurvedNewRow["Count"] = FrontCount;
                    FrontsSummaryDT.Rows.Add(CurvedNewRow);

                    FrontCost = 0;
                    FrontCount = 0;
                }
            }

            Table.Dispose();
            FrontsSummaryDT.DefaultView.Sort = "Front, Square DESC";
            FrontsSummaryBS.MoveFirst();
        }

        private void GetDecorProducts()
        {
            decimal DecorProductCost = 0;
            decimal DecorProductCount = 0;
            int decimals = 2;
            string Measure = string.Empty;
            DecorProductsSummaryDT.Clear();

            DataTable Table = new DataTable();

            using (DataView DV = new DataView(DecorConfigDT))
            {
                Table = DV.ToTable(true, new string[] { "ProductID", "MeasureID" });
            }

            for (int i = 0; i < Table.Rows.Count; i++)
            {
                DataRow[] Rows = DecorOrdersDT.Select("ProductID=" + Convert.ToInt32(Table.Rows[i]["ProductID"]) +
                    " AND MeasureID=" + Convert.ToInt32(Table.Rows[i]["MeasureID"]));
                if (Rows.Count() == 0)
                    continue;

                foreach (DataRow row in Rows)
                {
                    if (Convert.ToInt32(row["ProductID"]) == 2)
                    {
                        DecorProductCost += Convert.ToDecimal(row["Cost"]);
                        if (row["Height"].ToString() == "-1")
                            DecorProductCount += Convert.ToDecimal(row["Length"]) * Convert.ToDecimal(row["Count"]) / 1000;
                        else
                            DecorProductCount += Convert.ToDecimal(row["Height"]) * Convert.ToDecimal(row["Count"]) / 1000;
                        Measure = "м.п.";
                        continue;
                    }

                    if (Convert.ToInt32(row["MeasureID"]) == 1)
                    {
                        DecorProductCost += Convert.ToDecimal(row["Cost"]);
                        DecorProductCount += Convert.ToInt32(row["Count"]);
                        Measure = "шт.";
                    }

                    if (Convert.ToInt32(row["MeasureID"]) == 3)
                    {
                        DecorProductCost += Convert.ToDecimal(row["Cost"]);
                        DecorProductCount += Convert.ToInt32(row["Count"]);
                        Measure = "шт.";
                    }

                    if (Convert.ToInt32(row["MeasureID"]) == 2)
                    {
                        //нет параметра "высота"
                        if (row["Height"].ToString() == "-1")
                        {
                            DecorProductCount += Convert.ToDecimal(row["Length"]) * Convert.ToDecimal(row["Count"]) / 1000;
                        }
                        else
                        {
                            DecorProductCount += Convert.ToDecimal(row["Height"]) * Convert.ToDecimal(row["Count"]) / 1000;
                        }

                        DecorProductCost += Convert.ToDecimal(row["Cost"]);
                        Measure = "м.п.";
                    }
                }

                DataRow NewRow = DecorProductsSummaryDT.NewRow();
                NewRow["ProductID"] = Convert.ToInt32(Table.Rows[i]["ProductID"]);
                NewRow["MeasureID"] = Convert.ToInt32(Table.Rows[i]["MeasureID"]);
                NewRow["DecorProduct"] = GetProductName(Convert.ToInt32(Table.Rows[i]["ProductID"]));
                //if (DecorProductCount < 3)
                //    decimals = 1;
                NewRow["Cost"] = Decimal.Round(DecorProductCost, decimals, MidpointRounding.AwayFromZero);
                NewRow["Count"] = Decimal.Round(DecorProductCount, decimals, MidpointRounding.AwayFromZero);
                NewRow["Measure"] = Measure;
                DecorProductsSummaryDT.Rows.Add(NewRow);

                Measure = string.Empty;
                DecorProductCost = 0;
                DecorProductCount = 0;
            }
            DecorProductsSummaryDT.DefaultView.Sort = "DecorProduct, Measure ASC, Count DESC";
            DecorProductsSummaryBS.MoveFirst();
        }

        private void GetDecorItems()
        {
            decimal DecorItemCost = 0;
            decimal DecorItemCount = 0;
            int decimals = 2;
            string Measure = string.Empty;
            DecorItemsSummaryDT.Clear();

            DataTable Table = new DataTable();

            using (DataView DV = new DataView(DecorOrdersDT))
            {
                Table = DV.ToTable(true, new string[] { "ProductID", "DecorID", "MeasureID" });
            }

            for (int i = 0; i < Table.Rows.Count; i++)
            {
                DataRow[] Rows = DecorOrdersDT.Select("ProductID=" + Convert.ToInt32(Table.Rows[i]["ProductID"]) +
                    " AND DecorID=" + Convert.ToInt32(Table.Rows[i]["DecorID"]) + " AND MeasureID=" + Convert.ToInt32(Table.Rows[i]["MeasureID"]));
                if (Rows.Count() == 0)
                    continue;

                foreach (DataRow row in Rows)
                {
                    if (Convert.ToInt32(row["ProductID"]) == 2)
                    {
                        DecorItemCost += Convert.ToDecimal(row["Cost"]);
                        if (row["Height"].ToString() == "-1")
                            DecorItemCount += Convert.ToDecimal(row["Length"]) * Convert.ToDecimal(row["Count"]) / 1000;
                        else
                            DecorItemCount += Convert.ToDecimal(row["Height"]) * Convert.ToDecimal(row["Count"]) / 1000;
                        Measure = "м.п.";
                        continue;
                    }

                    if (Convert.ToInt32(row["MeasureID"]) == 1)
                    {
                        DecorItemCost += Convert.ToDecimal(row["Cost"]);
                        DecorItemCount += Convert.ToInt32(row["Count"]);
                        Measure = "шт.";
                    }

                    if (Convert.ToInt32(row["MeasureID"]) == 3)
                    {
                        DecorItemCost += Convert.ToDecimal(row["Cost"]);
                        DecorItemCount += Convert.ToInt32(row["Count"]);
                        Measure = "шт.";
                    }

                    if (Convert.ToInt32(row["MeasureID"]) == 2)
                    {
                        //нет параметра "высота"
                        if (row["Height"].ToString() == "-1")
                            DecorItemCount += Convert.ToDecimal(row["Length"]) * Convert.ToDecimal(row["Count"]) / 1000;
                        else
                            DecorItemCount += Convert.ToDecimal(row["Height"]) * Convert.ToDecimal(row["Count"]) / 1000;

                        DecorItemCost += Convert.ToDecimal(row["Cost"]);
                        Measure = "м.п.";
                    }
                }

                DataRow NewRow = DecorItemsSummaryDT.NewRow();
                NewRow["ProductID"] = Convert.ToInt32(Table.Rows[i]["ProductID"]);
                NewRow["DecorID"] = Convert.ToInt32(Table.Rows[i]["DecorID"]);
                NewRow["MeasureID"] = Convert.ToInt32(Table.Rows[i]["MeasureID"]);
                NewRow["DecorItem"] = GetDecorName(Convert.ToInt32(Table.Rows[i]["DecorID"]));
                if (DecorItemCount < 3)
                    decimals = 1;
                NewRow["Cost"] = Decimal.Round(DecorItemCost, decimals, MidpointRounding.AwayFromZero);
                NewRow["Count"] = Decimal.Round(DecorItemCount, decimals, MidpointRounding.AwayFromZero);
                NewRow["Measure"] = Measure;
                DecorItemsSummaryDT.Rows.Add(NewRow);

                Measure = string.Empty;
                DecorItemCost = 0;
                DecorItemCount = 0;
            }
            Table.Dispose();
            DecorItemsSummaryDT.DefaultView.Sort = "DecorItem, Count DESC";
            DecorItemsSummaryBS.MoveFirst();
        }

        public void GetFrontsInfo(ref decimal Square, ref decimal Cost, ref int Count, ref int CurvedCount)
        {
            for (int i = 0; i < FrontsSummaryDT.Rows.Count; i++)
            {
                if (Convert.ToInt32(FrontsSummaryDT.Rows[i]["Width"]) == -1)
                    CurvedCount += Convert.ToInt32(FrontsSummaryDT.Rows[i]["Count"]);
                else
                {
                    Square += Convert.ToDecimal(FrontsSummaryDT.Rows[i]["Square"]);
                    Count += Convert.ToInt32(FrontsSummaryDT.Rows[i]["Count"]);
                }
                Cost += Convert.ToDecimal(FrontsSummaryDT.Rows[i]["Cost"]);
            }

            Cost = Decimal.Round(Cost, 2, MidpointRounding.AwayFromZero);
            Square = Decimal.Round(Square, 2, MidpointRounding.AwayFromZero);
        }

        public void GetDecorInfo(ref decimal Pogon, ref decimal Cost, ref int Count)
        {
            for (int i = 0; i < DecorProductsSummaryDT.Rows.Count; i++)
            {
                if (Convert.ToInt32(DecorProductsSummaryDT.Rows[i]["MeasureID"]) != 2)
                    Count += Convert.ToInt32(DecorProductsSummaryDT.Rows[i]["Count"]);
                else
                {
                    Pogon += Convert.ToDecimal(DecorProductsSummaryDT.Rows[i]["Count"]);
                }
                Cost += Convert.ToDecimal(DecorProductsSummaryDT.Rows[i]["Cost"]);
            }

            Cost = Decimal.Round(Cost, 2, MidpointRounding.AwayFromZero);
            Pogon = Decimal.Round(Pogon, 2, MidpointRounding.AwayFromZero);
        }

        public string GetFrontName(int FrontID)
        {
            string FrontName = string.Empty;
            DataRow[] Rows = FrontsDT.Select("FrontID = " + FrontID);
            if (Rows.Count() > 0)
                FrontName = Rows[0]["FrontName"].ToString();
            return FrontName;
        }

        /// <summary>
        /// Возвращает название продукта
        /// </summary>
        /// <param name="ProductID"></param>
        /// <returns></returns>
        private string GetProductName(int ProductID)
        {
            string ProductName = string.Empty;
            try
            {
                DataRow[] Rows = DecorProductsDT.Select("ProductID = " + ProductID);
                ProductName = Rows[0]["ProductName"].ToString();
            }
            catch
            {
                return string.Empty;
            }
            return ProductName;
        }

        /// <summary>
        /// Возвращает название наименования
        /// </summary>
        /// <param name="DecorID"></param>
        /// <returns></returns>
        private string GetDecorName(int DecorID)
        {
            string DecorName = string.Empty;
            try
            {
                DataRow[] Rows = DecorItemsDT.Select("DecorID = " + DecorID);
                DecorName = Rows[0]["Name"].ToString();
            }
            catch
            {
                return string.Empty;
            }
            return DecorName;
        }

        private DateTime GetCurrentDateTime()
        {
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT GETDATE()", ConnectionStrings.LightConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    DA.Fill(DT);

                    return Convert.ToDateTime(DT.Rows[0][0]);
                }
            }
        }

        public void ClearFrontsOrders()
        {
            MFSummaryDT.Clear();
        }

        public void ClearDecorOrders()
        {
            MDSummaryDT.Clear();
        }

        private decimal MarketExpFrontsCost(int MegaOrderID, int FactoryID, ref decimal ExpCost, ref decimal AllCost)
        {
            decimal Percentage = 0;
            decimal ExpFrontsCost = 0;
            decimal AllFrontsCost = 0;

            DataRow[] RFRows = MExpFrontsCostDT.Select("MegaOrderID = " + MegaOrderID + " AND FactoryID = " + FactoryID);
            foreach (DataRow Row in RFRows)
                ExpFrontsCost += Convert.ToDecimal(Row["FrontsCost"]);

            DataRow[] AFRows = MAllFrontsCostDT.Select("MegaOrderID = " + MegaOrderID + " AND FactoryID = " + FactoryID);
            foreach (DataRow Row in AFRows)
                AllFrontsCost += Convert.ToDecimal(Row["FrontsCost"]);

            ExpCost = ExpFrontsCost;
            AllCost = AllFrontsCost;

            if (AllFrontsCost > 0)
                Percentage = ExpFrontsCost / AllFrontsCost;
            Percentage = Decimal.Round(Percentage * 100, 2, MidpointRounding.AwayFromZero);

            return Percentage;
        }

        private decimal MarketExpDecorCost(int MegaOrderID, int FactoryID, ref decimal ExpCost, ref decimal AllCost)
        {
            decimal Percentage = 0;
            decimal ExpDecorCost = 0;
            decimal AllDecorCost = 0;

            DataRow[] RDRows = MExpDecorCostDT.Select("MegaOrderID = " + MegaOrderID + " AND FactoryID = " + FactoryID);
            foreach (DataRow Row in RDRows)
                ExpDecorCost += Convert.ToDecimal(Row["DecorCost"]);

            DataRow[] ADRows = MAllDecorCostDT.Select("MegaOrderID = " + MegaOrderID + " AND FactoryID = " + FactoryID);
            foreach (DataRow Row in ADRows)
                AllDecorCost += Convert.ToDecimal(Row["DecorCost"]);

            ExpCost = ExpDecorCost;
            AllCost = AllDecorCost;

            if (AllDecorCost > 0)
                Percentage = ExpDecorCost / AllDecorCost;
            Percentage = Decimal.Round(Percentage * 100, 2, MidpointRounding.AwayFromZero);

            return Percentage;
        }

        private object GetCreationDate(int MegaOrderID)
        {
            string ConnectionString = ConnectionStrings.MarketingOrdersConnectionString;

            object DocDateTime = null;

            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT OrderDate FROM MegaOrders" +
                " WHERE MegaOrderID = " + MegaOrderID,
                ConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    DA.Fill(DT);

                    if (DT.Rows.Count > 0 && DT.Rows[0]["OrderDate"] != DBNull.Value)
                        DocDateTime = Convert.ToDateTime(DT.Rows[0]["OrderDate"]);
                }
            }
            return DocDateTime;
        }

        private object StorageAdoptionDate(int MegaOrderID, int FactoryID, int ProductType)
        {
            string ConnectionString = ConnectionStrings.MarketingOrdersConnectionString;

            object StorageDateTime = null;

            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT MIN(StorageDateTime) AS StorageDate FROM Packages" +
                " WHERE ProductType = " + ProductType + " AND FactoryID = " + FactoryID + " AND MainOrderID IN (SELECT MainOrderID FROM MainOrders" +
                " WHERE MegaOrderID = " + MegaOrderID + ")",
                ConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    DA.Fill(DT);

                    if (DT.Rows.Count > 0 && DT.Rows[0]["StorageDate"] != DBNull.Value)
                        StorageDateTime = Convert.ToDateTime(DT.Rows[0]["StorageDate"]);
                }
            }
            return StorageDateTime;
        }
    }
}