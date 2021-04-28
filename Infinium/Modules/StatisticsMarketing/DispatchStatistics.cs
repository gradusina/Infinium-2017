using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;

namespace Infinium.Modules.StatisticsMarketing
{
    public class DispatchStatistics
    {
        private DataTable MFSummaryDT = null;
        private DataTable MDSummaryDT = null;
        private DataTable ZFSummaryDT = null;
        private DataTable ZDSummaryDT = null;

        private DataTable MarketingClientsDT = null;
        private DataTable ZOVDispatchDT = null;

        private DataTable MDispFrontsCostDT = null;
        private DataTable MDispDecorCostDT = null;
        private DataTable MAllFrontsCostDT = null;
        private DataTable MAllDecorCostDT = null;

        private DataTable ZDispFrontsCostDT = null;
        private DataTable ZDispDecorCostDT = null;
        private DataTable ZAllFrontsCostDT = null;
        private DataTable ZAllDecorCostDT = null;

        private DataTable FrontsOrdersDT = null;
        private DataTable DecorOrdersDT = null;

        private DataTable FrontsSummaryDT = null;
        private DataTable DecorProductsSummaryDT = null;
        private DataTable DecorItemsSummaryDT = null;
        private DataTable DecorConfigDT = null;

        private DataTable FrontsDT = null;
        private DataTable DecorProductsDT = null;
        private DataTable DecorItemsDT = null;

        public BindingSource MFSummaryBS = null;
        public BindingSource MDSummaryBS = null;
        public BindingSource ZFSummaryBS = null;
        public BindingSource ZDSummaryBS = null;

        public BindingSource FrontsSummaryBS = null;
        public BindingSource DecorProductsSummaryBS = null;
        public BindingSource DecorItemsSummaryBS = null;

        private PercentageDataGrid MFSummaryDG = null;
        private PercentageDataGrid MDSummaryDG = null;
        private PercentageDataGrid ZFSummaryDG = null;
        private PercentageDataGrid ZDSummaryDG = null;
        private PercentageDataGrid FrontsDG = null;
        private PercentageDataGrid DecorProductsDG = null;
        private PercentageDataGrid DecorItemsDG = null;

        public DispatchStatistics(
            ref PercentageDataGrid tMFSummaryDG,
            ref PercentageDataGrid tMDSummaryDG,
            ref PercentageDataGrid tZFSummaryDG,
            ref PercentageDataGrid tZDSummaryDG,
            ref PercentageDataGrid tFrontsDG,
            ref PercentageDataGrid tDecorProductsDG,
            ref PercentageDataGrid tDecorItemsDG)
        {
            MFSummaryDG = tMFSummaryDG;
            MDSummaryDG = tMDSummaryDG;
            ZFSummaryDG = tZFSummaryDG;
            ZDSummaryDG = tZDSummaryDG;
            FrontsDG = tFrontsDG;
            DecorProductsDG = tDecorProductsDG;
            DecorItemsDG = tDecorItemsDG;

            Initialize();
        }

        private void Create()
        {
            MarketingClientsDT = new DataTable();
            ZOVDispatchDT = new DataTable();

            MDispFrontsCostDT = new DataTable();
            MDispDecorCostDT = new DataTable();
            MAllFrontsCostDT = new DataTable();
            MAllDecorCostDT = new DataTable();

            ZDispFrontsCostDT = new DataTable();
            ZDispDecorCostDT = new DataTable();
            ZAllFrontsCostDT = new DataTable();
            ZAllDecorCostDT = new DataTable();

            DecorItemsDT = new DataTable();
            FrontsDT = new DataTable();
            DecorProductsDT = new DataTable();
            DecorConfigDT = new DataTable();

            FrontsOrdersDT = new DataTable();
            DecorOrdersDT = new DataTable();

            MFSummaryDT = new DataTable();
            MFSummaryDT.Columns.Add(new DataColumn(("ClientID"), System.Type.GetType("System.Int32")));
            MFSummaryDT.Columns.Add(new DataColumn(("ClientName"), System.Type.GetType("System.String")));
            MFSummaryDT.Columns.Add(new DataColumn(("OrderNumber"), System.Type.GetType("System.Int32")));
            MFSummaryDT.Columns.Add(new DataColumn(("DispPercTPS"), System.Type.GetType("System.Decimal")));
            MFSummaryDT.Columns.Add(new DataColumn(("DispPercProfil"), System.Type.GetType("System.Decimal")));
            MFSummaryDT.Columns.Add(new DataColumn(("CostTPS"), System.Type.GetType("System.Decimal")));
            MFSummaryDT.Columns.Add(new DataColumn(("CostProfil"), System.Type.GetType("System.Decimal")));
            MFSummaryDT.Columns.Add(new DataColumn(("AllCostTPS"), System.Type.GetType("System.Decimal")));
            MFSummaryDT.Columns.Add(new DataColumn(("AllCostProfil"), System.Type.GetType("System.Decimal")));

            MDSummaryDT = new DataTable();
            MDSummaryDT.Columns.Add(new DataColumn(("ClientID"), System.Type.GetType("System.Int32")));
            MDSummaryDT.Columns.Add(new DataColumn(("ClientName"), System.Type.GetType("System.String")));
            MDSummaryDT.Columns.Add(new DataColumn(("OrderNumber"), System.Type.GetType("System.Int32")));
            MDSummaryDT.Columns.Add(new DataColumn(("DispPercTPS"), System.Type.GetType("System.Decimal")));
            MDSummaryDT.Columns.Add(new DataColumn(("DispPercProfil"), System.Type.GetType("System.Decimal")));
            MDSummaryDT.Columns.Add(new DataColumn(("CostTPS"), System.Type.GetType("System.Decimal")));
            MDSummaryDT.Columns.Add(new DataColumn(("CostProfil"), System.Type.GetType("System.Decimal")));
            MDSummaryDT.Columns.Add(new DataColumn(("AllCostTPS"), System.Type.GetType("System.Decimal")));
            MDSummaryDT.Columns.Add(new DataColumn(("AllCostProfil"), System.Type.GetType("System.Decimal")));

            ZFSummaryDT = new DataTable();
            ZFSummaryDT.Columns.Add(new DataColumn(("DispatchDateTime"), System.Type.GetType("System.DateTime")));
            ZFSummaryDT.Columns.Add(new DataColumn(("DispatchDate"), System.Type.GetType("System.DateTime")));
            ZFSummaryDT.Columns.Add(new DataColumn(("DispPercTPS"), System.Type.GetType("System.Decimal")));
            ZFSummaryDT.Columns.Add(new DataColumn(("DispPercProfil"), System.Type.GetType("System.Decimal")));
            ZFSummaryDT.Columns.Add(new DataColumn(("CostTPS"), System.Type.GetType("System.Decimal")));
            ZFSummaryDT.Columns.Add(new DataColumn(("CostProfil"), System.Type.GetType("System.Decimal")));

            ZDSummaryDT = new DataTable();
            ZDSummaryDT.Columns.Add(new DataColumn(("DispatchDateTime"), System.Type.GetType("System.DateTime")));
            ZDSummaryDT.Columns.Add(new DataColumn(("DispatchDate"), System.Type.GetType("System.DateTime")));
            ZDSummaryDT.Columns.Add(new DataColumn(("DispPercTPS"), System.Type.GetType("System.Decimal")));
            ZDSummaryDT.Columns.Add(new DataColumn(("DispPercProfil"), System.Type.GetType("System.Decimal")));
            ZDSummaryDT.Columns.Add(new DataColumn(("CostTPS"), System.Type.GetType("System.Decimal")));
            ZDSummaryDT.Columns.Add(new DataColumn(("CostProfil"), System.Type.GetType("System.Decimal")));

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

            //System.Diagnostics.Stopwatch sw = new System.Diagnostics.Stopwatch();
            //sw.Start();

            //FillMarketingTables(DateTime.Now.AddDays(-3), DateTime.Now);
            //FillZOVTables(DateTime.Now.AddDays(-3), DateTime.Now);

            //sw.Stop();
            //double G = sw.Elapsed.TotalMilliseconds;

            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT TOP 0 FrontID, PatinaID, ColorID, InsetTypeID," +
                " InsetColorID, Height, Width, Count, Square, Cost FROM FrontsOrders",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                FrontsOrdersDT.Clear();
                DA.Fill(FrontsOrdersDT);
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

        public void FillZOVTables(DateTime FirstDate, DateTime SecondDate)
        {
            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT CONVERT(varchar(10), MegaOrders.DispatchDate, 121) AS DispatchDate, FrontsOrders.FactoryID, SUM(FrontsOrders.Cost) AS FrontsCost
            FROM PackageDetails INNER JOIN
            FrontsOrders ON PackageDetails.OrderID = FrontsOrders.FrontsOrdersID INNER JOIN
            Packages ON PackageDetails.PackageID = Packages.PackageID INNER JOIN
            MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID INNER JOIN
            MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID AND MegaOrders.MegaOrderID <> 0
            WHERE PackageDetails.PackageID IN (SELECT PackageID FROM Packages WHERE ProductType = 0)
            GROUP BY CONVERT(varchar(10), MegaOrders.DispatchDate, 121), FrontsOrders.FactoryID
            ORDER BY DispatchDate",
                ConnectionStrings.ZOVOrdersConnectionString))
            {
                ZAllFrontsCostDT.Clear();
                DA.Fill(ZAllFrontsCostDT);
            }

            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT CONVERT(varchar(10), MegaOrders.DispatchDate, 121) AS DispatchDate, DecorOrders.FactoryID, SUM(DecorOrders.Cost) AS DecorCost
            FROM PackageDetails INNER JOIN
            DecorOrders ON PackageDetails.OrderID = DecorOrders.DecorOrderID INNER JOIN
            Packages ON PackageDetails.PackageID = Packages.PackageID INNER JOIN
            MainOrders ON DecorOrders.MainOrderID = MainOrders.MainOrderID INNER JOIN
            MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID AND MegaOrders.MegaOrderID <> 0
            WHERE PackageDetails.PackageID IN (SELECT PackageID FROM Packages WHERE ProductType = 1)
            GROUP BY CONVERT(varchar(10), MegaOrders.DispatchDate, 121), DecorOrders.FactoryID
            ORDER BY DispatchDate",
                ConnectionStrings.ZOVOrdersConnectionString))
            {
                ZAllDecorCostDT.Clear();
                DA.Fill(ZAllDecorCostDT);
            }

            using (SqlDataAdapter DA = new SqlDataAdapter(
            @"SELECT CONVERT(varchar(10), Packages.DispatchDateTime, 121) AS DispatchDateTime, CONVERT(varchar(10), MegaOrders.DispatchDate, 121) AS DispatchDate
            FROM MainOrders
            INNER JOIN Packages ON MainOrders.MainOrderID = Packages.MainOrderID
            INNER JOIN MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID AND MainOrders.MegaOrderID <> 0
            WHERE MainOrders.MainOrderID IN (SELECT MainOrderID FROM Packages WHERE PackageStatusID = 3)
            GROUP BY CONVERT(varchar(10), Packages.DispatchDateTime, 121), CONVERT(varchar(10), MegaOrders.DispatchDate, 121)
            ORDER BY DispatchDateTime, DispatchDate",
                ConnectionStrings.ZOVOrdersConnectionString))
            {
                ZOVDispatchDT.Clear();
                DA.Fill(ZOVDispatchDT);
            }

            string DispDateFilter = " AND CAST(DispatchDateTime AS date) >= '" + FirstDate.ToString("yyyy-MM-dd") +
                "' AND CAST(DispatchDateTime AS date) <='" + SecondDate.ToString("yyyy-MM-dd") + "'";

            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT CONVERT(varchar(10), Packages.DispatchDateTime, 121) AS DispatchDateTime, CONVERT(varchar(10), MegaOrders.DispatchDate, 121) AS DispatchDate, FrontsOrders.FactoryID, SUM(FrontsOrders.Cost * PackageDetails.Count / FrontsOrders.Count) AS FrontsCost
            FROM PackageDetails INNER JOIN
            FrontsOrders ON PackageDetails.OrderID = FrontsOrders.FrontsOrdersID INNER JOIN
            Packages ON PackageDetails.PackageID = Packages.PackageID INNER JOIN
            MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID INNER JOIN
            MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID AND MegaOrders.MegaOrderID <> 0
            WHERE PackageDetails.PackageID IN (SELECT PackageID FROM Packages WHERE ProductType = 0 AND PackageStatusID = 3 " + DispDateFilter + @")
            GROUP BY CONVERT(varchar(10), Packages.DispatchDateTime, 121), CONVERT(varchar(10), MegaOrders.DispatchDate, 121), FrontsOrders.FactoryID
            ORDER BY DispatchDateTime, DispatchDate",
                ConnectionStrings.ZOVOrdersConnectionString))
            {
                ZDispFrontsCostDT.Clear();
                DA.Fill(ZDispFrontsCostDT);
            }

            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT CONVERT(varchar(10), Packages.DispatchDateTime, 121) AS DispatchDateTime, CONVERT(varchar(10), MegaOrders.DispatchDate, 121) AS DispatchDate, DecorOrders.FactoryID, SUM(DecorOrders.Cost * PackageDetails.Count / DecorOrders.Count) AS DecorCost
            FROM PackageDetails INNER JOIN
            DecorOrders ON PackageDetails.OrderID = DecorOrders.DecorOrderID INNER JOIN
            Packages ON PackageDetails.PackageID = Packages.PackageID INNER JOIN
            MainOrders ON DecorOrders.MainOrderID = MainOrders.MainOrderID INNER JOIN
            MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID AND MegaOrders.MegaOrderID <> 0
            WHERE PackageDetails.PackageID IN (SELECT PackageID FROM Packages WHERE ProductType = 1 AND PackageStatusID = 3 " + DispDateFilter + @")
            GROUP BY CONVERT(varchar(10), Packages.DispatchDateTime, 121), CONVERT(varchar(10), MegaOrders.DispatchDate, 121), DecorOrders.FactoryID
            ORDER BY DispatchDateTime, DispatchDate",
                ConnectionStrings.ZOVOrdersConnectionString))
            {
                ZDispDecorCostDT.Clear();
                DA.Fill(ZDispDecorCostDT);
            }
        }

        public void FillMarketingTables(DateTime FirstDate, DateTime SecondDate)
        {
            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT MainOrders.MegaOrderID, FrontsOrders.FactoryID, SUM(FrontsOrders.Cost) AS FrontsCost
            FROM FrontsOrders INNER JOIN
            MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID INNER JOIN
            MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID
            WHERE FrontsOrders.FrontsOrdersID IN
            (SELECT OrderID FROM PackageDetails WHERE PackageID IN (SELECT PackageID FROM Packages WHERE ProductType = 0))
            GROUP BY MainOrders.MegaOrderID, FrontsOrders.FactoryID
            ORDER BY MainOrders.MegaOrderID, FrontsOrders.FactoryID",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                MAllFrontsCostDT.Clear();
                DA.Fill(MAllFrontsCostDT);
            }

            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT MainOrders.MegaOrderID, DecorOrders.FactoryID, SUM(DecorOrders.Cost) AS DecorCost
            FROM DecorOrders INNER JOIN
            MainOrders ON DecorOrders.MainOrderID = MainOrders.MainOrderID INNER JOIN
            MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID
            WHERE DecorOrders.DecorOrderID IN
            (SELECT OrderID FROM PackageDetails WHERE PackageID IN (SELECT PackageID FROM Packages WHERE ProductType = 1))
            GROUP BY MainOrders.MegaOrderID, DecorOrders.FactoryID
            ORDER BY MainOrders.MegaOrderID, DecorOrders.FactoryID",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                MAllDecorCostDT.Clear();
                DA.Fill(MAllDecorCostDT);
            }

            using (SqlDataAdapter DA = new SqlDataAdapter(
            @"SELECT MegaOrders.ClientID, infiniu2_marketingreference.dbo.Clients.ClientName, MegaOrders.OrderNumber, MegaOrders.MegaOrderID
            FROM MegaOrders
            INNER JOIN infiniu2_marketingreference.dbo.Clients ON MegaOrders.ClientID = infiniu2_marketingreference.dbo.Clients.ClientID
            WHERE MegaOrders.MegaOrderID IN (SELECT MegaOrderID FROM MainOrders WHERE MainOrderID IN (SELECT MainOrderID FROM Packages WHERE PackageStatusID = 3))
            ORDER BY infiniu2_marketingreference.dbo.Clients.ClientName, MegaOrders.OrderNumber, MegaOrders.MegaOrderID",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                MarketingClientsDT.Clear();
                DA.Fill(MarketingClientsDT);
            }

            string DispDateFilter = " AND CAST(DispatchDateTime AS date) >= '" + FirstDate.ToString("yyyy-MM-dd") +
                "' AND CAST(DispatchDateTime AS date) <='" + SecondDate.ToString("yyyy-MM-dd") + "'";

            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT MainOrders.MegaOrderID, FrontsOrders.FactoryID, SUM(FrontsOrders.Cost * PackageDetails.Count / FrontsOrders.Count) AS FrontsCost
            FROM PackageDetails INNER JOIN
            FrontsOrders ON PackageDetails.OrderID = FrontsOrders.FrontsOrdersID INNER JOIN
            MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID
            WHERE PackageDetails.PackageID IN (SELECT PackageID FROM Packages WHERE ProductType = 0 AND PackageStatusID = 3 " + DispDateFilter + @")
            GROUP BY MainOrders.MegaOrderID, FrontsOrders.FactoryID
            ORDER BY MainOrders.MegaOrderID, FrontsOrders.FactoryID",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                MDispFrontsCostDT.Clear();
                DA.Fill(MDispFrontsCostDT);
            }

            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT MainOrders.MegaOrderID, DecorOrders.FactoryID, SUM(DecorOrders.Cost * PackageDetails.Count / DecorOrders.Count) AS DecorCost
            FROM PackageDetails INNER JOIN
            DecorOrders ON PackageDetails.OrderID = DecorOrders.DecorOrderID INNER JOIN
            MainOrders ON DecorOrders.MainOrderID = MainOrders.MainOrderID
            WHERE PackageDetails.PackageID IN (SELECT PackageID FROM Packages WHERE ProductType = 1 AND PackageStatusID = 3 " + DispDateFilter + @")
            GROUP BY MainOrders.MegaOrderID, DecorOrders.FactoryID
            ORDER BY MainOrders.MegaOrderID, DecorOrders.FactoryID",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                MDispDecorCostDT.Clear();
                DA.Fill(MDispDecorCostDT);
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

            ZFSummaryBS = new BindingSource()
            {
                DataSource = ZFSummaryDT
            };
            ZFSummaryDG.DataSource = ZFSummaryBS;

            ZDSummaryBS = new BindingSource()
            {
                DataSource = ZDSummaryDT
            };
            ZDSummaryDG.DataSource = ZDSummaryBS;

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
            ZOVSummaryGridSettings();
        }

        public void ShowColumns(ref PercentageDataGrid FrontsGrid, ref PercentageDataGrid DecorGrid, bool Profil, bool TPS)
        {
            if (Profil && TPS)
            {
                FrontsGrid.Columns["CostProfil"].Visible = true;
                FrontsGrid.Columns["CostTPS"].Visible = true;
                FrontsGrid.Columns["DispPercProfil"].Visible = true;
                FrontsGrid.Columns["DispPercTPS"].Visible = true;

                DecorGrid.Columns["CostProfil"].Visible = true;
                DecorGrid.Columns["CostTPS"].Visible = true;
                DecorGrid.Columns["DispPercProfil"].Visible = true;
                DecorGrid.Columns["DispPercTPS"].Visible = true;
            }
            if (Profil && !TPS)
            {
                FrontsGrid.Columns["CostProfil"].Visible = true;
                FrontsGrid.Columns["CostTPS"].Visible = false;
                FrontsGrid.Columns["DispPercProfil"].Visible = true;
                FrontsGrid.Columns["DispPercTPS"].Visible = false;

                DecorGrid.Columns["CostProfil"].Visible = true;
                DecorGrid.Columns["CostTPS"].Visible = false;
                DecorGrid.Columns["DispPercProfil"].Visible = true;
                DecorGrid.Columns["DispPercTPS"].Visible = false;
            }
            if (!Profil && TPS)
            {
                FrontsGrid.Columns["CostProfil"].Visible = false;
                FrontsGrid.Columns["CostTPS"].Visible = true;
                FrontsGrid.Columns["DispPercProfil"].Visible = false;
                FrontsGrid.Columns["DispPercTPS"].Visible = true;

                DecorGrid.Columns["CostProfil"].Visible = false;
                DecorGrid.Columns["CostTPS"].Visible = true;
                DecorGrid.Columns["DispPercProfil"].Visible = false;
                DecorGrid.Columns["DispPercTPS"].Visible = true;
            }
        }

        private void MarketSummaryGridSettings()
        {
            MFSummaryDG.Columns["ClientID"].Visible = false;
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

            MFSummaryDG.Columns["ClientName"].HeaderText = "Клиент";
            MFSummaryDG.Columns["OrderNumber"].HeaderText = "№ заказа";
            MFSummaryDG.Columns["DispPercProfil"].HeaderText = "Отгружено\n\r Профиль, %";
            MFSummaryDG.Columns["DispPercTPS"].HeaderText = "Отгружено\n\r      ТПС, %";
            MFSummaryDG.Columns["CostProfil"].HeaderText = "Профиль, €";
            MFSummaryDG.Columns["CostTPS"].HeaderText = "ТПС, €";
            MFSummaryDG.Columns["ClientName"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            MFSummaryDG.Columns["ClientName"].MinimumWidth = 110;
            MFSummaryDG.Columns["OrderNumber"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MFSummaryDG.Columns["OrderNumber"].Width = 85;
            MFSummaryDG.Columns["DispPercProfil"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            MFSummaryDG.Columns["DispPercProfil"].MinimumWidth = 135;
            MFSummaryDG.Columns["DispPercTPS"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            MFSummaryDG.Columns["DispPercTPS"].MinimumWidth = 135;
            MFSummaryDG.Columns["CostProfil"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MFSummaryDG.Columns["CostProfil"].Width = 110;
            MFSummaryDG.Columns["CostTPS"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MFSummaryDG.Columns["CostTPS"].Width = 110;
            MFSummaryDG.Columns["CostProfil"].DefaultCellStyle.Format = "N";
            MFSummaryDG.Columns["CostProfil"].DefaultCellStyle.FormatProvider = nfi1;
            MFSummaryDG.Columns["CostTPS"].DefaultCellStyle.Format = "N";
            MFSummaryDG.Columns["CostTPS"].DefaultCellStyle.FormatProvider = nfi1;

            MFSummaryDG.Columns["DispPercProfil"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            MFSummaryDG.AddPercentageColumn("DispPercProfil");
            MFSummaryDG.Columns["DispPercTPS"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            MFSummaryDG.AddPercentageColumn("DispPercTPS");

            MDSummaryDG.Columns["ClientName"].HeaderText = "Клиент";
            MDSummaryDG.Columns["OrderNumber"].HeaderText = "№ заказа";
            MDSummaryDG.Columns["DispPercProfil"].HeaderText = "Отгружено\n\r Профиль, %";
            MDSummaryDG.Columns["DispPercTPS"].HeaderText = "Отгружено\n\r     ТПС, %";
            MDSummaryDG.Columns["CostProfil"].HeaderText = "Профиль, €";
            MDSummaryDG.Columns["CostTPS"].HeaderText = "ТПС, €";
            MDSummaryDG.Columns["ClientName"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            MDSummaryDG.Columns["ClientName"].MinimumWidth = 110;
            MDSummaryDG.Columns["OrderNumber"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MDSummaryDG.Columns["OrderNumber"].Width = 85;
            MDSummaryDG.Columns["DispPercProfil"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            MDSummaryDG.Columns["DispPercProfil"].MinimumWidth = 135;
            MDSummaryDG.Columns["DispPercTPS"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            MDSummaryDG.Columns["DispPercTPS"].MinimumWidth = 135;
            MDSummaryDG.Columns["CostProfil"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MDSummaryDG.Columns["CostProfil"].Width = 110;
            MDSummaryDG.Columns["CostTPS"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MDSummaryDG.Columns["CostTPS"].Width = 110;
            MDSummaryDG.Columns["CostProfil"].DefaultCellStyle.Format = "N";
            MDSummaryDG.Columns["CostProfil"].DefaultCellStyle.FormatProvider = nfi1;
            MDSummaryDG.Columns["CostTPS"].DefaultCellStyle.Format = "N";
            MDSummaryDG.Columns["CostTPS"].DefaultCellStyle.FormatProvider = nfi1;

            MDSummaryDG.Columns["DispPercProfil"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            MDSummaryDG.AddPercentageColumn("DispPercProfil");
            MDSummaryDG.Columns["DispPercTPS"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            MDSummaryDG.AddPercentageColumn("DispPercTPS");
        }

        private void ZOVSummaryGridSettings()
        {
            if (!Security.PriceAccess)
            {
                ZFSummaryDG.Columns["CostProfil"].Visible = false;
                ZFSummaryDG.Columns["CostTPS"].Visible = false;
                ZDSummaryDG.Columns["CostProfil"].Visible = false;
                ZDSummaryDG.Columns["CostTPS"].Visible = false;
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
            foreach (DataGridViewColumn Column in ZFSummaryDG.Columns)
            {
                Column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
            foreach (DataGridViewColumn Column in ZDSummaryDG.Columns)
            {
                Column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }

            ZFSummaryDG.Columns["DispatchDate"].DefaultCellStyle.Format = "dd.MM.yyyy";
            ZFSummaryDG.Columns["DispatchDateTime"].DefaultCellStyle.Format = "dd.MM.yyyy";
            //ZFSummaryDG.Columns["DispatchDate"].DefaultCellStyle.Format = "dd.MM.yyyy";
            //ZFSummaryDG.Columns["DocDateTime"].DefaultCellStyle.Format = "dd.MM.yyyy";
            ZFSummaryDG.Columns["DispatchDate"].HeaderText = " Ожидаемая\n\rдата отгрузки";
            ZFSummaryDG.Columns["DispatchDateTime"].HeaderText = "Фактическая\n\rдата отгрузки";
            ZFSummaryDG.Columns["DispPercProfil"].HeaderText = "Отгружено\n\r Профиль, %";
            ZFSummaryDG.Columns["DispPercTPS"].HeaderText = "Отгружено\n\r      ТПС, %";
            ZFSummaryDG.Columns["CostProfil"].HeaderText = "Профиль, €";
            ZFSummaryDG.Columns["CostTPS"].HeaderText = "ТПС, €";
            ZFSummaryDG.Columns["DispatchDate"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            ZFSummaryDG.Columns["DispatchDate"].MinimumWidth = 110;
            ZFSummaryDG.Columns["DispatchDateTime"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            ZFSummaryDG.Columns["DispatchDateTime"].MinimumWidth = 110;
            ZFSummaryDG.Columns["DispPercProfil"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            ZFSummaryDG.Columns["DispPercProfil"].MinimumWidth = 135;
            ZFSummaryDG.Columns["DispPercTPS"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            ZFSummaryDG.Columns["DispPercTPS"].MinimumWidth = 135;
            ZFSummaryDG.Columns["CostProfil"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            ZFSummaryDG.Columns["CostProfil"].Width = 110;
            ZFSummaryDG.Columns["CostTPS"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            ZFSummaryDG.Columns["CostTPS"].Width = 110;
            ZFSummaryDG.Columns["CostProfil"].DefaultCellStyle.Format = "N";
            ZFSummaryDG.Columns["CostProfil"].DefaultCellStyle.FormatProvider = nfi1;
            ZFSummaryDG.Columns["CostTPS"].DefaultCellStyle.Format = "N";
            ZFSummaryDG.Columns["CostTPS"].DefaultCellStyle.FormatProvider = nfi1;

            ZFSummaryDG.Columns["DispPercProfil"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            ZFSummaryDG.AddPercentageColumn("DispPercProfil");
            ZFSummaryDG.Columns["DispPercTPS"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            ZFSummaryDG.AddPercentageColumn("DispPercTPS");

            ZDSummaryDG.Columns["DispatchDate"].DefaultCellStyle.Format = "dd.MM.yyyy";
            ZDSummaryDG.Columns["DispatchDateTime"].DefaultCellStyle.Format = "dd.MM.yyyy";
            ZDSummaryDG.Columns["DispatchDate"].HeaderText = " Ожидаемая\n\rдата отгрузки";
            ZDSummaryDG.Columns["DispatchDateTime"].HeaderText = "Фактическая\n\rдата отгрузки";
            ZDSummaryDG.Columns["DispPercProfil"].HeaderText = "Отгружено\n\r Профиль, %";
            ZDSummaryDG.Columns["DispPercTPS"].HeaderText = "Отгружено\n\r     ТПС, %";
            ZDSummaryDG.Columns["CostProfil"].HeaderText = "Профиль, €";
            ZDSummaryDG.Columns["CostTPS"].HeaderText = "ТПС, €";
            ZDSummaryDG.Columns["DispatchDate"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            ZDSummaryDG.Columns["DispatchDate"].MinimumWidth = 110;
            ZDSummaryDG.Columns["DispatchDateTime"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            ZDSummaryDG.Columns["DispatchDateTime"].MinimumWidth = 110;
            ZDSummaryDG.Columns["DispPercProfil"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            ZDSummaryDG.Columns["DispPercProfil"].MinimumWidth = 135;
            ZDSummaryDG.Columns["DispPercTPS"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            ZDSummaryDG.Columns["DispPercTPS"].MinimumWidth = 135;
            ZDSummaryDG.Columns["CostProfil"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            ZDSummaryDG.Columns["CostProfil"].Width = 110;
            ZDSummaryDG.Columns["CostTPS"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            ZDSummaryDG.Columns["CostTPS"].Width = 110;
            ZDSummaryDG.Columns["CostProfil"].DefaultCellStyle.Format = "N";
            ZDSummaryDG.Columns["CostProfil"].DefaultCellStyle.FormatProvider = nfi1;
            ZDSummaryDG.Columns["CostTPS"].DefaultCellStyle.Format = "N";
            ZDSummaryDG.Columns["CostTPS"].DefaultCellStyle.FormatProvider = nfi1;

            ZDSummaryDG.Columns["DispPercProfil"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            ZDSummaryDG.AddPercentageColumn("DispPercProfil");
            ZDSummaryDG.Columns["DispPercTPS"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            ZDSummaryDG.AddPercentageColumn("DispPercTPS");
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

            decimal DispPercProfil = 0;
            decimal DispPercTPS = 0;

            decimal DispCostProfil = 0;
            decimal DispCostTPS = 0;

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

                d1 = 0;
                d2 = 0;

                if (FactoryID == 0)
                {
                    DispCostProfil = 0;
                    DispCostTPS = 0;
                    DispPercProfil = MarketDispFrontsCost(true, MegaOrderID, 1, ref DispCostProfil, ref AllCostProfil);
                    DispPercTPS = MarketDispFrontsCost(true, MegaOrderID, 2, ref DispCostTPS, ref AllCostTPS);

                    if (DispPercProfil != 0 || DispPercTPS != 0)
                    {
                        DataRow[] Rows = MFSummaryDT.Select("ClientName='" + ClientName + "'");
                        if (Rows.Count() == 0)
                        {
                            DataRow NewRow = MFSummaryDT.NewRow();
                            NewRow["ClientID"] = ClientID;
                            NewRow["ClientName"] = ClientName;
                            if (DispPercProfil > 0)
                            {
                                NewRow["DispPercProfil"] = DispPercProfil;
                                NewRow["CostProfil"] = DispCostProfil;
                                NewRow["AllCostProfil"] = AllCostProfil;
                            }
                            if (DispPercTPS > 0)
                            {
                                NewRow["DispPercTPS"] = DispPercTPS;
                                NewRow["CostTPS"] = DispCostTPS;
                                NewRow["AllCostTPS"] = AllCostTPS;
                            }
                            MFSummaryDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            if (DispPercProfil > 0)
                            {
                                if (Rows[0]["CostProfil"] != DBNull.Value)
                                    d1 = Convert.ToDecimal(Rows[0]["CostProfil"]);
                                if (Rows[0]["AllCostProfil"] != DBNull.Value)
                                    d2 = Convert.ToDecimal(Rows[0]["AllCostProfil"]);
                                Percentage = (d1 + DispCostProfil) / (d2 + AllCostProfil);
                                Percentage = Decimal.Round(Percentage * 100, 2, MidpointRounding.AwayFromZero);
                                Rows[0]["DispPercProfil"] = Percentage;
                                Rows[0]["AllCostProfil"] = d2 + AllCostProfil;
                                Rows[0]["CostProfil"] = d1 + DispCostProfil;
                            }
                            if (DispPercTPS > 0)
                            {
                                if (Rows[0]["CostTPS"] != DBNull.Value)
                                    d1 = Convert.ToDecimal(Rows[0]["CostTPS"]);
                                if (Rows[0]["AllCostTPS"] != DBNull.Value)
                                    d2 = Convert.ToDecimal(Rows[0]["AllCostTPS"]);
                                Percentage = (d1 + DispCostTPS) / (d2 + AllCostTPS);
                                Percentage = Decimal.Round(Percentage * 100, 2, MidpointRounding.AwayFromZero);
                                Rows[0]["DispPercTPS"] = Percentage;
                                Rows[0]["AllCostTPS"] = d2 + AllCostTPS;
                                Rows[0]["CostTPS"] = d1 + DispCostTPS;
                            }
                        }
                    }

                    DispCostProfil = 0;
                    DispCostTPS = 0;
                    DispPercProfil = MarketDispDecorCost(true, MegaOrderID, 1, ref DispCostProfil, ref AllCostProfil);
                    DispPercTPS = MarketDispDecorCost(true, MegaOrderID, 2, ref DispCostTPS, ref AllCostTPS);

                    if (DispCostProfil != 0 || DispCostTPS != 0)
                    {
                        DataRow[] Rows = MDSummaryDT.Select("ClientName='" + ClientName + "'");
                        if (Rows.Count() == 0)
                        {
                            DataRow NewRow = MDSummaryDT.NewRow();
                            NewRow["ClientID"] = ClientID;
                            NewRow["ClientName"] = ClientName;
                            if (DispCostProfil > 0)
                            {
                                NewRow["DispPercProfil"] = DispPercProfil;
                                NewRow["CostProfil"] = DispCostProfil;
                                NewRow["AllCostProfil"] = AllCostProfil;
                            }
                            if (DispCostTPS > 0)
                            {
                                NewRow["DispPercTPS"] = DispPercTPS;
                                NewRow["CostTPS"] = DispCostTPS;
                                NewRow["AllCostTPS"] = AllCostTPS;
                            }
                            MDSummaryDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            if (DispPercProfil > 0)
                            {
                                if (Rows[0]["CostProfil"] != DBNull.Value)
                                    d1 = Convert.ToDecimal(Rows[0]["CostProfil"]);
                                if (Rows[0]["AllCostProfil"] != DBNull.Value)
                                    d2 = Convert.ToDecimal(Rows[0]["AllCostProfil"]);
                                Percentage = (d1 + DispCostProfil) / (d2 + AllCostProfil);
                                Percentage = Decimal.Round(Percentage * 100, 2, MidpointRounding.AwayFromZero);
                                Rows[0]["DispPercProfil"] = Percentage;
                                Rows[0]["AllCostProfil"] = d2 + AllCostProfil;
                                Rows[0]["CostProfil"] = d1 + DispCostProfil;
                            }
                            if (DispPercTPS > 0)
                            {
                                if (Rows[0]["CostTPS"] != DBNull.Value)
                                    d1 = Convert.ToDecimal(Rows[0]["CostTPS"]);
                                if (Rows[0]["AllCostTPS"] != DBNull.Value)
                                    d2 = Convert.ToDecimal(Rows[0]["AllCostTPS"]);
                                Percentage = (d1 + DispCostTPS) / (d2 + AllCostTPS);
                                Percentage = Decimal.Round(Percentage * 100, 2, MidpointRounding.AwayFromZero);
                                Rows[0]["DispPercTPS"] = Percentage;
                                Rows[0]["AllCostTPS"] = d2 + AllCostTPS;
                                Rows[0]["CostTPS"] = d1 + DispCostTPS;
                            }
                        }
                    }
                }

                if (FactoryID == 1)
                {
                    DispCostProfil = 0;
                    DispPercProfil = MarketDispFrontsCost(true, MegaOrderID, 1, ref DispCostProfil, ref AllCostProfil);

                    if (DispPercProfil > 0)
                    {
                        DataRow[] Rows = MFSummaryDT.Select("ClientName='" + ClientName + "'");
                        if (Rows.Count() == 0)
                        {
                            DataRow NewRow = MFSummaryDT.NewRow();
                            NewRow["ClientID"] = ClientID;
                            NewRow["ClientName"] = ClientName;
                            NewRow["DispPercProfil"] = DispPercProfil;
                            NewRow["CostProfil"] = DispCostProfil;
                            NewRow["AllCostProfil"] = AllCostProfil;
                            MFSummaryDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            if (Rows[0]["CostProfil"] != DBNull.Value)
                                d1 = Convert.ToDecimal(Rows[0]["CostProfil"]);
                            if (Rows[0]["AllCostProfil"] != DBNull.Value)
                                d2 = Convert.ToDecimal(Rows[0]["AllCostProfil"]);
                            Percentage = (d1 + DispCostProfil) / (d2 + AllCostProfil);
                            Percentage = Decimal.Round(Percentage * 100, 2, MidpointRounding.AwayFromZero);
                            Rows[0]["DispPercProfil"] = Percentage;
                            Rows[0]["AllCostProfil"] = d2 + AllCostProfil;
                            Rows[0]["CostProfil"] = d1 + DispCostProfil;
                        }
                    }

                    DispCostProfil = 0;
                    DispPercProfil = MarketDispDecorCost(true, MegaOrderID, 1, ref DispCostProfil, ref AllCostProfil);

                    if (DispCostProfil > 0)
                    {
                        DataRow[] Rows = MDSummaryDT.Select("ClientName='" + ClientName + "'");
                        if (Rows.Count() == 0)
                        {
                            DataRow NewRow = MDSummaryDT.NewRow();
                            NewRow["ClientID"] = ClientID;
                            NewRow["ClientName"] = ClientName;
                            NewRow["DispPercProfil"] = DispPercProfil;
                            NewRow["CostProfil"] = DispCostProfil;
                            NewRow["AllCostProfil"] = AllCostProfil;
                            MDSummaryDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            if (Rows[0]["CostProfil"] != DBNull.Value)
                                d1 = Convert.ToDecimal(Rows[0]["CostProfil"]);
                            if (Rows[0]["AllCostProfil"] != DBNull.Value)
                                d2 = Convert.ToDecimal(Rows[0]["AllCostProfil"]);
                            Percentage = (d1 + DispCostProfil) / (d2 + AllCostProfil);
                            Percentage = Decimal.Round(Percentage * 100, 2, MidpointRounding.AwayFromZero);
                            Rows[0]["DispPercProfil"] = Percentage;
                            Rows[0]["AllCostProfil"] = d2 + AllCostProfil;
                            Rows[0]["CostProfil"] = d1 + DispCostProfil;
                        }
                    }
                }

                if (FactoryID == 2)
                {
                    DispCostTPS = 0;
                    DispPercTPS = MarketDispFrontsCost(true, MegaOrderID, 2, ref DispCostTPS, ref AllCostTPS);

                    if (DispPercTPS > 0)
                    {
                        DataRow[] Rows = MFSummaryDT.Select("ClientName='" + ClientName + "'");
                        if (Rows.Count() == 0)
                        {
                            DataRow NewRow = MFSummaryDT.NewRow();
                            NewRow["ClientID"] = ClientID;
                            NewRow["ClientName"] = ClientName;
                            NewRow["DispPercTPS"] = DispPercTPS;
                            NewRow["CostTPS"] = DispCostTPS;
                            NewRow["AllCostTPS"] = AllCostTPS;
                            MFSummaryDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            if (Rows[0]["CostTPS"] != DBNull.Value)
                                d1 = Convert.ToDecimal(Rows[0]["CostTPS"]);
                            if (Rows[0]["AllCostTPS"] != DBNull.Value)
                                d2 = Convert.ToDecimal(Rows[0]["AllCostTPS"]);
                            Percentage = (d1 + DispCostTPS) / (d2 + AllCostTPS);
                            Percentage = Decimal.Round(Percentage * 100, 2, MidpointRounding.AwayFromZero);
                            Rows[0]["DispPercTPS"] = Percentage;
                            Rows[0]["AllCostTPS"] = d2 + AllCostTPS;
                            Rows[0]["CostTPS"] = d1 + DispCostTPS;
                        }
                    }

                    DispCostTPS = 0;
                    DispPercTPS = MarketDispDecorCost(true, MegaOrderID, 2, ref DispCostTPS, ref AllCostTPS);

                    if (DispCostTPS > 0)
                    {
                        DataRow[] Rows = MDSummaryDT.Select("ClientName='" + ClientName + "'");
                        if (Rows.Count() == 0)
                        {
                            DataRow NewRow = MDSummaryDT.NewRow();
                            NewRow["ClientID"] = ClientID;
                            NewRow["ClientName"] = ClientName;
                            NewRow["DispPercTPS"] = DispPercTPS;
                            NewRow["CostTPS"] = DispCostTPS;
                            NewRow["AllCostTPS"] = AllCostTPS;
                            MDSummaryDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            if (Rows[0]["CostTPS"] != DBNull.Value)
                                d1 = Convert.ToDecimal(Rows[0]["CostTPS"]);
                            if (Rows[0]["AllCostTPS"] != DBNull.Value)
                                d2 = Convert.ToDecimal(Rows[0]["AllCostTPS"]);
                            Percentage = (d1 + DispCostTPS) / (d2 + AllCostTPS);
                            Percentage = Decimal.Round(Percentage * 100, 2, MidpointRounding.AwayFromZero);
                            Rows[0]["DispPercTPS"] = Percentage;
                            Rows[0]["AllCostTPS"] = d2 + AllCostTPS;
                            Rows[0]["CostTPS"] = d1 + DispCostTPS;
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

            decimal DispPercProfil = 0;
            decimal DispPercTPS = 0;

            decimal DispCostProfil = 0;
            decimal DispCostTPS = 0;

            decimal AllCostProfil = 0;
            decimal AllCostTPS = 0;

            string ClientName = string.Empty;

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

                if (FactoryID == 0)
                {
                    DispCostProfil = 0;
                    DispCostTPS = 0;
                    DispPercProfil = MarketDispFrontsCost(true, MegaOrderID, 1, ref DispCostProfil, ref AllCostProfil);
                    DispPercTPS = MarketDispFrontsCost(true, MegaOrderID, 2, ref DispCostTPS, ref AllCostTPS);

                    if (DispPercProfil != 0 || DispPercTPS != 0)
                    {
                        DataRow NewRow = MFSummaryDT.NewRow();
                        NewRow["ClientID"] = ClientID;
                        NewRow["ClientName"] = ClientName;
                        NewRow["OrderNumber"] = OrderNumber;
                        if (DispPercProfil > 0)
                        {
                            NewRow["DispPercProfil"] = DispPercProfil;
                            NewRow["CostProfil"] = DispCostProfil;
                        }
                        if (DispPercTPS > 0)
                        {
                            NewRow["DispPercTPS"] = DispPercTPS;
                            NewRow["CostTPS"] = DispCostTPS;
                        }
                        MFSummaryDT.Rows.Add(NewRow);
                    }

                    DispCostProfil = 0;
                    DispCostTPS = 0;
                    DispPercProfil = MarketDispDecorCost(true, MegaOrderID, 1, ref DispCostProfil, ref AllCostProfil);
                    DispPercTPS = MarketDispDecorCost(true, MegaOrderID, 2, ref DispCostTPS, ref AllCostTPS);

                    if (DispCostProfil != 0 || DispCostTPS != 0)
                    {
                        DataRow NewRow = MDSummaryDT.NewRow();
                        NewRow["ClientID"] = ClientID;
                        NewRow["ClientName"] = ClientName;
                        NewRow["OrderNumber"] = OrderNumber;
                        if (DispCostProfil > 0)
                        {
                            NewRow["DispPercProfil"] = DispPercProfil;
                            NewRow["CostProfil"] = DispCostProfil;
                        }
                        if (DispCostTPS > 0)
                        {
                            NewRow["DispPercTPS"] = DispPercTPS;
                            NewRow["CostTPS"] = DispCostTPS;
                        }
                        MDSummaryDT.Rows.Add(NewRow);
                    }
                }

                if (FactoryID == 1)
                {
                    DispCostProfil = 0;
                    DispPercProfil = MarketDispFrontsCost(true, MegaOrderID, 1, ref DispCostProfil, ref AllCostProfil);

                    if (DispPercProfil > 0)
                    {
                        DataRow NewRow = MFSummaryDT.NewRow();
                        NewRow["ClientID"] = ClientID;
                        NewRow["ClientName"] = ClientName;
                        NewRow["OrderNumber"] = OrderNumber;
                        NewRow["DispPercProfil"] = DispPercProfil;
                        NewRow["CostProfil"] = DispCostProfil;
                        MFSummaryDT.Rows.Add(NewRow);
                    }

                    DispCostProfil = 0;
                    DispPercProfil = MarketDispDecorCost(true, MegaOrderID, 1, ref DispCostProfil, ref AllCostProfil);

                    if (DispCostProfil > 0)
                    {
                        DataRow NewRow = MDSummaryDT.NewRow();
                        NewRow["ClientID"] = ClientID;
                        NewRow["ClientName"] = ClientName;
                        NewRow["OrderNumber"] = OrderNumber;
                        NewRow["DispPercProfil"] = DispPercProfil;
                        NewRow["CostProfil"] = DispCostProfil;
                        MDSummaryDT.Rows.Add(NewRow);
                    }
                }

                if (FactoryID == 2)
                {
                    DispCostTPS = 0;
                    DispPercTPS = MarketDispFrontsCost(true, MegaOrderID, 2, ref DispCostTPS, ref AllCostTPS);

                    if (DispPercTPS > 0)
                    {
                        DataRow NewRow = MFSummaryDT.NewRow();
                        NewRow["ClientID"] = ClientID;
                        NewRow["ClientName"] = ClientName;
                        NewRow["OrderNumber"] = OrderNumber;
                        NewRow["DispPercTPS"] = DispPercTPS;
                        NewRow["CostTPS"] = DispCostTPS;
                        MFSummaryDT.Rows.Add(NewRow);
                    }

                    DispCostTPS = 0;
                    DispPercTPS = MarketDispDecorCost(true, MegaOrderID, 2, ref DispCostTPS, ref AllCostTPS);

                    if (DispCostTPS > 0)
                    {
                        DataRow NewRow = MDSummaryDT.NewRow();
                        NewRow["ClientID"] = ClientID;
                        NewRow["ClientName"] = ClientName;
                        NewRow["OrderNumber"] = OrderNumber;
                        NewRow["DispPercTPS"] = DispPercTPS;
                        NewRow["CostTPS"] = DispCostTPS;
                        MDSummaryDT.Rows.Add(NewRow);
                    }
                }
            }

            MFSummaryDT.DefaultView.Sort = "ClientName";
            MDSummaryDT.DefaultView.Sort = "ClientName";
            MFSummaryBS.MoveFirst();
            MDSummaryBS.MoveFirst();
        }

        public void ZOVSummary(int FactoryID)
        {
            decimal DispPercProfil = 0;
            decimal DispPercTPS = 0;

            decimal DispCostProfil = 0;
            decimal DispCostTPS = 0;
            decimal AllCostProfil = 0;
            decimal AllCostTPS = 0;

            string DispatchDateTime = string.Empty;
            string DispatchDate = string.Empty;

            ZFSummaryDT.Clear();
            ZDSummaryDT.Clear();

            if (FactoryID == -1)
                return;

            for (int i = 0; i < ZOVDispatchDT.Rows.Count; i++)
            {
                DispatchDate = ZOVDispatchDT.Rows[i]["DispatchDate"].ToString();
                DispatchDateTime = ZOVDispatchDT.Rows[i]["DispatchDateTime"].ToString();

                if (FactoryID == 0)
                {
                    DispCostProfil = 0;
                    DispCostTPS = 0;
                    DispPercProfil = ZOVDispFrontsCost(DispatchDate, DispatchDateTime, 1, ref DispCostProfil, ref AllCostProfil);
                    DispPercTPS = ZOVDispFrontsCost(DispatchDate, DispatchDateTime, 2, ref DispCostTPS, ref AllCostTPS);

                    if (DispPercProfil != 0 || DispPercTPS != 0)
                    {
                        DataRow NewRow = ZFSummaryDT.NewRow();
                        NewRow["DispatchDate"] = DispatchDate;
                        NewRow["DispatchDateTime"] = DispatchDateTime;
                        if (DispPercProfil > 0)
                        {
                            NewRow["DispPercProfil"] = DispPercProfil;
                            NewRow["CostProfil"] = DispCostProfil;
                        }
                        if (DispPercTPS > 0)
                        {
                            NewRow["DispPercTPS"] = DispPercTPS;
                            NewRow["CostTPS"] = DispCostTPS;
                        }
                        ZFSummaryDT.Rows.Add(NewRow);
                    }

                    DispCostProfil = 0;
                    DispCostTPS = 0;
                    DispPercProfil = ZOVDispDecorCost(DispatchDate, DispatchDateTime, 1, ref DispCostProfil, ref AllCostProfil);
                    DispPercTPS = ZOVDispDecorCost(DispatchDate, DispatchDateTime, 2, ref DispCostTPS, ref AllCostTPS);

                    if (DispCostProfil != 0 || DispCostTPS != 0)
                    {
                        DataRow NewRow = ZDSummaryDT.NewRow();
                        NewRow["DispatchDate"] = DispatchDate;
                        NewRow["DispatchDateTime"] = DispatchDateTime;
                        if (DispCostProfil > 0)
                        {
                            NewRow["DispPercProfil"] = DispPercProfil;
                            NewRow["CostProfil"] = DispCostProfil;
                        }
                        if (DispCostTPS > 0)
                        {
                            NewRow["DispPercTPS"] = DispPercTPS;
                            NewRow["CostTPS"] = DispCostTPS;
                        }
                        ZDSummaryDT.Rows.Add(NewRow);
                    }
                }
                if (FactoryID == 1)
                {
                    DispCostProfil = 0;
                    DispPercProfil = ZOVDispFrontsCost(DispatchDate, DispatchDateTime, 1, ref DispCostProfil, ref AllCostProfil);

                    if (DispCostProfil > 0)
                    {
                        DataRow NewRow = ZFSummaryDT.NewRow();
                        NewRow["DispatchDate"] = DispatchDate;
                        NewRow["DispatchDateTime"] = DispatchDateTime;
                        NewRow["DispPercProfil"] = DispPercProfil;
                        NewRow["CostProfil"] = DispCostProfil;
                        ZFSummaryDT.Rows.Add(NewRow);
                    }

                    DispCostProfil = 0;
                    DispPercProfil = ZOVDispDecorCost(DispatchDate, DispatchDateTime, 1, ref DispCostProfil, ref AllCostProfil);

                    if (DispCostProfil > 0)
                    {
                        DataRow NewRow = ZDSummaryDT.NewRow();
                        NewRow["DispatchDate"] = DispatchDate;
                        NewRow["DispatchDateTime"] = DispatchDateTime;
                        NewRow["DispPercProfil"] = DispPercProfil;
                        NewRow["CostProfil"] = DispCostProfil;
                        ZDSummaryDT.Rows.Add(NewRow);
                    }
                }
                if (FactoryID == 2)
                {
                    DispCostTPS = 0;
                    DispPercTPS = ZOVDispFrontsCost(DispatchDate, DispatchDateTime, 2, ref DispCostTPS, ref AllCostTPS);

                    if (DispCostTPS > 0)
                    {
                        DataRow NewRow = ZFSummaryDT.NewRow();
                        NewRow["DispatchDate"] = DispatchDate;
                        NewRow["DispatchDateTime"] = DispatchDateTime;
                        NewRow["DispPercTPS"] = DispPercTPS;
                        NewRow["CostTPS"] = DispCostTPS;
                        ZFSummaryDT.Rows.Add(NewRow);
                    }

                    DispCostTPS = 0;
                    DispPercTPS = ZOVDispDecorCost(DispatchDate, DispatchDateTime, 2, ref DispCostTPS, ref AllCostTPS);

                    if (DispCostTPS > 0)
                    {
                        DataRow NewRow = ZDSummaryDT.NewRow();
                        NewRow["DispatchDate"] = DispatchDate;
                        NewRow["DispatchDateTime"] = DispatchDateTime;
                        NewRow["DispPercTPS"] = DispPercTPS;
                        NewRow["CostTPS"] = DispCostTPS;
                        ZDSummaryDT.Rows.Add(NewRow);
                    }
                }
            }

            //ZFSummaryDT.DefaultView.Sort = "DispPercTPS desc, DispPercProfil desc";
            //ZDSummaryDT.DefaultView.Sort = "DispPercTPS desc, DispPercProfil desc";
            ZFSummaryBS.MoveFirst();
            ZDSummaryBS.MoveFirst();
        }

        public void FMarketingOrders(DateTime FirstDate, DateTime SecondDate, int FactoryID, int ClientID)
        {
            string MarketingSelectCommand = string.Empty;

            string MFrontsPackageFilter = string.Empty;

            string ClientFilter = string.Empty;
            string DispDateFilter = string.Empty;
            string PackageFactoryFilter = string.Empty;

            if (ClientID != -1)
                ClientFilter = " AND ClientID = " + ClientID;
            if (FactoryID != 0)
                PackageFactoryFilter = " AND FactoryID = " + FactoryID;

            DispDateFilter = " AND CAST(DispatchDateTime AS date) >= '" + FirstDate.ToString("yyyy-MM-dd") +
                "' AND CAST(DispatchDateTime AS date) <='" + SecondDate.ToString("yyyy-MM-dd") + "'";

            MFrontsPackageFilter = " PackageID IN" +
                " (SELECT PackageID FROM Packages WHERE PackageStatusID = 3 AND ProductType = 0 " + PackageFactoryFilter + DispDateFilter + ")";

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

        public void DMarketingOrders(DateTime FirstDate, DateTime SecondDate, int FactoryID, int ClientID)
        {
            string MarketingSelectCommand = string.Empty;

            string MDecorPackageFilter = string.Empty;

            string ClientFilter = string.Empty;
            string DispDateFilter = string.Empty;
            string PackageFactoryFilter = string.Empty;

            if (ClientID != -1)
                ClientFilter = " AND ClientID = " + ClientID;
            if (FactoryID != 0)
                PackageFactoryFilter = " AND FactoryID = " + FactoryID;

            DispDateFilter = " AND CAST(DispatchDateTime AS date) >= '" + FirstDate.ToString("yyyy-MM-dd") +
                "' AND CAST(DispatchDateTime AS date) <='" + SecondDate.ToString("yyyy-MM-dd") + "'";

            MDecorPackageFilter = " PackageID IN" +
                " (SELECT PackageID FROM Packages WHERE PackageStatusID = 3 AND ProductType = 1 " + PackageFactoryFilter + DispDateFilter + ")";

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

        public void ZOVOrders(DateTime FirstDate, DateTime SecondDate, int FactoryID)
        {
            string ZOVSelectCommand = string.Empty;

            string ZFrontsPackageFilter = string.Empty;
            string ZDecorPackageFilter = string.Empty;

            string DispDateFilter = string.Empty;
            string PackageFactoryFilter = string.Empty;

            if (FactoryID != 0)
                PackageFactoryFilter = " AND FactoryID = " + FactoryID;

            DispDateFilter = " AND CAST(DispatchDateTime AS date) >= '" + FirstDate.ToString("yyyy-MM-dd") +
                "' AND CAST(DispatchDateTime AS date) <='" + SecondDate.ToString("yyyy-MM-dd") + "'";

            ZFrontsPackageFilter = " PackageID IN" +
                " (SELECT PackageID FROM Packages WHERE PackageStatusID = 3 AND ProductType = 0 " + PackageFactoryFilter + DispDateFilter + ")";

            ZDecorPackageFilter = " PackageID IN" +
                " (SELECT PackageID FROM Packages WHERE PackageStatusID = 3 AND ProductType = 1 " + PackageFactoryFilter + DispDateFilter + ")";

            ZOVSelectCommand = "SELECT FrontsOrders.FrontID, FrontsOrders.PatinaID," +
                " FrontsOrders.ColorID, FrontsOrders.InsetTypeID," +
                " FrontsOrders.InsetColorID, FrontsOrders.TechnoColorID, FrontsOrders.TechnoInsetTypeID, FrontsOrders.TechnoInsetColorID, FrontsOrders.Height, FrontsOrders.Width," +
                " PackageDetails.Count, (FrontsOrders.Square * PackageDetails.Count / FrontsOrders.Count) AS Square, (FrontsOrders.Cost * PackageDetails.Count / FrontsOrders.Count) AS Cost, MeasureID FROM PackageDetails" +
                " INNER JOIN FrontsOrders ON PackageDetails.OrderID = FrontsOrders.FrontsOrdersID" +
                " INNER JOIN infiniu2_catalog.dbo.FrontsConfig" +
                " ON FrontsOrders.FrontConfigID=infiniu2_catalog.dbo.FrontsConfig.FrontConfigID" +
                " WHERE MainOrderID IN (SELECT MainOrderID FROM MainOrders WHERE MegaOrderID <> 0) AND " + ZFrontsPackageFilter;

            using (SqlDataAdapter DA = new SqlDataAdapter(ZOVSelectCommand,
                ConnectionStrings.ZOVOrdersConnectionString))
            {
                FrontsOrdersDT.Clear();
                DA.Fill(FrontsOrdersDT);
            }

            GetFronts();

            //decor
            ZOVSelectCommand = "SELECT DecorOrders.ProductID, DecorOrders.DecorID, DecorOrders.ColorID,DecorOrders.PatinaID," +
                " DecorOrders.Height, DecorOrders.Length, DecorOrders.Width, PackageDetails.Count, (DecorOrders.Cost * PackageDetails.Count / DecorOrders.Count) AS Cost, DecorOrders.DecorConfigID, " +
                " MeasureID FROM PackageDetails" +
                " INNER JOIN DecorOrders ON PackageDetails.OrderID = DecorOrders.DecorOrderID" +
                " INNER JOIN infiniu2_catalog.dbo.DecorConfig" +
                " ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                " WHERE MainOrderID IN (SELECT MainOrderID FROM MainOrders WHERE MegaOrderID <> 0) AND " + ZDecorPackageFilter;

            using (SqlDataAdapter DA = new SqlDataAdapter(ZOVSelectCommand,
                ConnectionStrings.ZOVOrdersConnectionString))
            {
                DecorOrdersDT.Clear();
                DA.Fill(DecorOrdersDT);
            }

            GetDecorProducts();
            GetDecorItems();

            ZOVSummary(FactoryID);
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

        public void ClearFrontsOrders(int TypeOrders)
        {
            if (TypeOrders == 1)
            {
                MFSummaryDT.Clear();
            }
            if (TypeOrders == 3)
            {
                ZFSummaryDT.Clear();
            }
        }

        public void ClearDecorOrders(int TypeOrders)
        {
            if (TypeOrders == 1)
            {
                MDSummaryDT.Clear();
            }
            if (TypeOrders == 3)
            {
                ZDSummaryDT.Clear();
            }
        }

        private decimal MarketDispFrontsCost(bool IsMarket, int MegaOrderID, int FactoryID, ref decimal DispCost, ref decimal AllCost)
        {
            decimal Percentage = 0;
            decimal DispFrontsCost = 0;
            decimal AllFrontsCost = 0;

            if (IsMarket)
            {
                DataRow[] RFRows = MDispFrontsCostDT.Select("MegaOrderID = " + MegaOrderID + " AND FactoryID = " + FactoryID);
                foreach (DataRow Row in RFRows)
                    DispFrontsCost += Convert.ToDecimal(Row["FrontsCost"]);

                DataRow[] AFRows = MAllFrontsCostDT.Select("MegaOrderID = " + MegaOrderID + " AND FactoryID = " + FactoryID);
                foreach (DataRow Row in AFRows)
                    AllFrontsCost += Convert.ToDecimal(Row["FrontsCost"]);
            }
            else
            {
                DataRow[] RFRows = ZDispFrontsCostDT.Select("MegaOrderID = " + MegaOrderID + " AND FactoryID = " + FactoryID);
                foreach (DataRow Row in RFRows)
                    DispFrontsCost += Convert.ToDecimal(Row["FrontsCost"]);

                DataRow[] AFRows = ZAllFrontsCostDT.Select("MegaOrderID = " + MegaOrderID + " AND FactoryID = " + FactoryID);
                foreach (DataRow Row in AFRows)
                    AllFrontsCost += Convert.ToDecimal(Row["FrontsCost"]);
            }

            DispCost = DispFrontsCost;
            AllCost = AllFrontsCost;

            if (AllFrontsCost > 0)
                Percentage = DispFrontsCost / AllFrontsCost;
            Percentage = Decimal.Round(Percentage * 100, 2, MidpointRounding.AwayFromZero);

            return Percentage;
        }

        private decimal MarketDispDecorCost(bool IsMarket, int MegaOrderID, int FactoryID, ref decimal DispCost, ref decimal AllCost)
        {
            decimal Percentage = 0;
            decimal DispDecorCost = 0;
            decimal AllDecorCost = 0;

            if (IsMarket)
            {
                DataRow[] RDRows = MDispDecorCostDT.Select("MegaOrderID = " + MegaOrderID + " AND FactoryID = " + FactoryID);
                foreach (DataRow Row in RDRows)
                    DispDecorCost += Convert.ToDecimal(Row["DecorCost"]);

                DataRow[] ADRows = MAllDecorCostDT.Select("MegaOrderID = " + MegaOrderID + " AND FactoryID = " + FactoryID);
                foreach (DataRow Row in ADRows)
                    AllDecorCost += Convert.ToDecimal(Row["DecorCost"]);
            }
            else
            {
                DataRow[] RDRows = ZDispDecorCostDT.Select("MegaOrderID = " + MegaOrderID + " AND FactoryID = " + FactoryID);
                foreach (DataRow Row in RDRows)
                    DispDecorCost += Convert.ToDecimal(Row["DecorCost"]);

                DataRow[] ADRows = ZAllDecorCostDT.Select("MegaOrderID = " + MegaOrderID + " AND FactoryID = " + FactoryID);
                foreach (DataRow Row in ADRows)
                    AllDecorCost += Convert.ToDecimal(Row["DecorCost"]);
            }

            DispCost = DispDecorCost;
            AllCost = AllDecorCost;

            if (AllDecorCost > 0)
                Percentage = DispDecorCost / AllDecorCost;
            Percentage = Decimal.Round(Percentage * 100, 2, MidpointRounding.AwayFromZero);

            return Percentage;
        }

        private decimal ZOVDispFrontsCost(string DispatchDate, string DispatchDateTime, int FactoryID, ref decimal DispCost, ref decimal AllCost)
        {
            decimal Percentage = 0;
            decimal DispFrontsCost = 0;
            decimal AllFrontsCost = 0;

            DataRow[] RFRows = ZDispFrontsCostDT.Select("DispatchDate = '" + DispatchDate + "' AND DispatchDateTime = '" + DispatchDateTime + "' AND FactoryID = " + FactoryID);
            foreach (DataRow Row in RFRows)
                DispFrontsCost += Convert.ToDecimal(Row["FrontsCost"]);

            DataRow[] AFRows = ZAllFrontsCostDT.Select("DispatchDate = '" + DispatchDate + "' AND FactoryID = " + FactoryID);
            foreach (DataRow Row in AFRows)
                AllFrontsCost += Convert.ToDecimal(Row["FrontsCost"]);

            DispCost = DispFrontsCost;
            AllCost = AllFrontsCost;

            if (AllFrontsCost > 0)
                Percentage = DispFrontsCost / AllFrontsCost;
            Percentage = Decimal.Round(Percentage * 100, 2, MidpointRounding.AwayFromZero);

            return Percentage;
        }

        private decimal ZOVDispDecorCost(string DispatchDate, string DispatchDateTime, int FactoryID, ref decimal DispCost, ref decimal AllCost)
        {
            decimal Percentage = 0;
            decimal DispDecorCost = 0;
            decimal AllDecorCost = 0;

            DataRow[] RDRows = ZDispDecorCostDT.Select("DispatchDate = '" + DispatchDate + "' AND DispatchDateTime = '" + DispatchDateTime + "' AND FactoryID = " + FactoryID);
            foreach (DataRow Row in RDRows)
                DispDecorCost += Convert.ToDecimal(Row["DecorCost"]);

            DataRow[] ADRows = ZAllDecorCostDT.Select("DispatchDate = '" + DispatchDate + "' AND FactoryID = " + FactoryID);
            foreach (DataRow Row in ADRows)
                AllDecorCost += Convert.ToDecimal(Row["DecorCost"]);

            DispCost = DispDecorCost;
            AllCost = AllDecorCost;

            if (AllDecorCost > 0)
                Percentage = DispDecorCost / AllDecorCost;
            Percentage = Decimal.Round(Percentage * 100, 2, MidpointRounding.AwayFromZero);

            return Percentage;
        }
    }
}