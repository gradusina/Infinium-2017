using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;

namespace Infinium.Modules.StatisticsMarketing
{
    public class StorageStatistics
    {
        private DataTable MCurvedFSummaryDT = null;
        private DataTable MFSummaryDT = null;
        private DataTable MDSummaryDT = null;
        private DataTable PrepareCurvedFSummaryDT = null;
        private DataTable PrepareFSummaryDT = null;
        private DataTable PrepareDSummaryDT = null;

        private DataTable ClientGroupsDataTable = null;
        private DataTable MarketingClientsDT = null;
        private DataTable ZOVPrepareDT = null;

        private DataTable MReadyCurvedFrontsCostDT = null;
        private DataTable MReadyFrontsCostDT = null;
        private DataTable MReadyDecorCostDT = null;
        private DataTable MAllCurvedFrontsCostDT = null;
        private DataTable MAllFrontsCostDT = null;
        private DataTable MAllDecorCostDT = null;

        private DataTable PrepareReadyCurvedFrontsCostDT = null;
        private DataTable PrepareReadyFrontsCostDT = null;
        private DataTable PrepareReadyDecorCostDT = null;
        private DataTable PrepareAllCurvedFrontsCostDT = null;
        private DataTable PrepareAllFrontsCostDT = null;
        private DataTable PrepareAllDecorCostDT = null;

        private DataTable ZReadyCurvedFrontsCostDT = null;
        private DataTable ZReadyFrontsCostDT = null;
        private DataTable ZReadyDecorCostDT = null;
        private DataTable ZAllCurvedFrontsCostDT = null;
        private DataTable ZAllFrontsCostDT = null;
        private DataTable ZAllDecorCostDT = null;

        private DataTable CurvedFrontsOrdersDT = null;
        private DataTable FrontsOrdersDT = null;
        private DataTable DecorOrdersDT = null;

        private DataTable CurvedFrontsSummaryDT = null;
        private DataTable FrontsSummaryDT = null;
        private DataTable DecorProductsSummaryDT = null;
        private DataTable DecorItemsSummaryDT = null;
        private DataTable DecorConfigDT = null;

        private DataTable FrontsDT = null;
        private DataTable DecorProductsDT = null;
        private DataTable DecorItemsDT = null;

        public BindingSource ClientGroupsBS = null;
        public BindingSource MCurvedFSummaryBS = null;
        public BindingSource MFSummaryBS = null;
        public BindingSource MDSummaryBS = null;
        public BindingSource PrepareCurvedFSummaryBS = null;
        public BindingSource PrepareFSummaryBS = null;
        public BindingSource PrepareDSummaryBS = null;

        public BindingSource CurvedFrontsSummaryBS = null;
        public BindingSource FrontsSummaryBS = null;
        public BindingSource DecorProductsSummaryBS = null;
        public BindingSource DecorItemsSummaryBS = null;

        private PercentageDataGrid MCurvedFSummaryDG = null;
        private PercentageDataGrid MFSummaryDG = null;
        private PercentageDataGrid MDSummaryDG = null;
        private PercentageDataGrid PrepareCurvedFSummaryDG = null;
        private PercentageDataGrid PrepareFSummaryDG = null;
        private PercentageDataGrid PrepareDSummaryDG = null;
        private PercentageDataGrid CurvedFrontsDG = null;
        private PercentageDataGrid FrontsDG = null;
        private PercentageDataGrid DecorProductsDG = null;
        private PercentageDataGrid DecorItemsDG = null;

        public StorageStatistics(
            ref PercentageDataGrid tMFSummaryDG,
            ref PercentageDataGrid tMCurvedFSummaryDG,
            ref PercentageDataGrid tMDSummaryDG,
            ref PercentageDataGrid tPrepareFSummaryDG,
            ref PercentageDataGrid tPrepareCurvedFSummaryDG,
            ref PercentageDataGrid tPrepareDSummaryDG,
            ref PercentageDataGrid tFrontsDG,
            ref PercentageDataGrid tCurvedFrontsDG,
            ref PercentageDataGrid tDecorProductsDG,
            ref PercentageDataGrid tDecorItemsDG)
        {
            MCurvedFSummaryDG = tMCurvedFSummaryDG;
            MFSummaryDG = tMFSummaryDG;
            MDSummaryDG = tMDSummaryDG;
            PrepareCurvedFSummaryDG = tPrepareCurvedFSummaryDG;
            PrepareFSummaryDG = tPrepareFSummaryDG;
            PrepareDSummaryDG = tPrepareDSummaryDG;
            CurvedFrontsDG = tCurvedFrontsDG;
            FrontsDG = tFrontsDG;
            DecorProductsDG = tDecorProductsDG;
            DecorItemsDG = tDecorItemsDG;

            Initialize();
        }

        private void Create()
        {
            MarketingClientsDT = new DataTable();
            ZOVPrepareDT = new DataTable();

            MReadyCurvedFrontsCostDT = new DataTable();
            MReadyFrontsCostDT = new DataTable();
            MReadyDecorCostDT = new DataTable();
            MAllCurvedFrontsCostDT = new DataTable();
            MAllFrontsCostDT = new DataTable();
            MAllDecorCostDT = new DataTable();

            ZReadyCurvedFrontsCostDT = new DataTable();
            ZReadyFrontsCostDT = new DataTable();
            ZReadyDecorCostDT = new DataTable();
            ZAllCurvedFrontsCostDT = new DataTable();
            ZAllFrontsCostDT = new DataTable();
            ZAllDecorCostDT = new DataTable();

            PrepareReadyCurvedFrontsCostDT = new DataTable();
            PrepareReadyFrontsCostDT = new DataTable();
            PrepareReadyDecorCostDT = new DataTable();
            PrepareAllCurvedFrontsCostDT = new DataTable();
            PrepareAllFrontsCostDT = new DataTable();
            PrepareAllDecorCostDT = new DataTable();

            DecorItemsDT = new DataTable();
            FrontsDT = new DataTable();
            DecorProductsDT = new DataTable();
            DecorConfigDT = new DataTable();

            CurvedFrontsOrdersDT = new DataTable();
            FrontsOrdersDT = new DataTable();
            DecorOrdersDT = new DataTable();

            MFSummaryDT = new DataTable();
            MFSummaryDT.Columns.Add(new DataColumn(("ClientID"), System.Type.GetType("System.Int32")));
            MFSummaryDT.Columns.Add(new DataColumn(("ClientName"), System.Type.GetType("System.String")));
            MFSummaryDT.Columns.Add(new DataColumn(("MegaOrderID"), System.Type.GetType("System.Int32")));
            MFSummaryDT.Columns.Add(new DataColumn(("OrderNumber"), System.Type.GetType("System.Int32")));
            MFSummaryDT.Columns.Add(new DataColumn(("DocDateTime"), System.Type.GetType("System.DateTime")));
            MFSummaryDT.Columns.Add(new DataColumn(("ReadyPercTPS"), System.Type.GetType("System.Decimal")));
            MFSummaryDT.Columns.Add(new DataColumn(("ReadyPercProfil"), System.Type.GetType("System.Decimal")));
            MFSummaryDT.Columns.Add(new DataColumn(("CostTPS"), System.Type.GetType("System.Decimal")));
            MFSummaryDT.Columns.Add(new DataColumn(("CostProfil"), System.Type.GetType("System.Decimal")));
            MFSummaryDT.Columns.Add(new DataColumn(("AllCostTPS"), System.Type.GetType("System.Decimal")));
            MFSummaryDT.Columns.Add(new DataColumn(("AllCostProfil"), System.Type.GetType("System.Decimal")));
            MCurvedFSummaryDT = MFSummaryDT.Clone();

            MDSummaryDT = new DataTable();
            MDSummaryDT.Columns.Add(new DataColumn(("ClientID"), System.Type.GetType("System.Int32")));
            MDSummaryDT.Columns.Add(new DataColumn(("ClientName"), System.Type.GetType("System.String")));
            MDSummaryDT.Columns.Add(new DataColumn(("MegaOrderID"), System.Type.GetType("System.Int32")));
            MDSummaryDT.Columns.Add(new DataColumn(("OrderNumber"), System.Type.GetType("System.Int32")));
            MDSummaryDT.Columns.Add(new DataColumn(("DocDateTime"), System.Type.GetType("System.DateTime")));
            MDSummaryDT.Columns.Add(new DataColumn(("ReadyPercTPS"), System.Type.GetType("System.Decimal")));
            MDSummaryDT.Columns.Add(new DataColumn(("ReadyPercProfil"), System.Type.GetType("System.Decimal")));
            MDSummaryDT.Columns.Add(new DataColumn(("CostTPS"), System.Type.GetType("System.Decimal")));
            MDSummaryDT.Columns.Add(new DataColumn(("CostProfil"), System.Type.GetType("System.Decimal")));
            MDSummaryDT.Columns.Add(new DataColumn(("AllCostTPS"), System.Type.GetType("System.Decimal")));
            MDSummaryDT.Columns.Add(new DataColumn(("AllCostProfil"), System.Type.GetType("System.Decimal")));

            PrepareFSummaryDT = new DataTable();
            PrepareFSummaryDT.Columns.Add(new DataColumn(("DocDateTime"), System.Type.GetType("System.DateTime")));
            PrepareFSummaryDT.Columns.Add(new DataColumn(("ReadyPercTPS"), System.Type.GetType("System.Decimal")));
            PrepareFSummaryDT.Columns.Add(new DataColumn(("ReadyPercProfil"), System.Type.GetType("System.Decimal")));
            PrepareFSummaryDT.Columns.Add(new DataColumn(("CostTPS"), System.Type.GetType("System.Decimal")));
            PrepareFSummaryDT.Columns.Add(new DataColumn(("CostProfil"), System.Type.GetType("System.Decimal")));
            PrepareCurvedFSummaryDT = PrepareFSummaryDT.Clone();

            PrepareDSummaryDT = new DataTable();
            PrepareDSummaryDT.Columns.Add(new DataColumn(("DocDateTime"), System.Type.GetType("System.DateTime")));
            PrepareDSummaryDT.Columns.Add(new DataColumn(("ReadyPercTPS"), System.Type.GetType("System.Decimal")));
            PrepareDSummaryDT.Columns.Add(new DataColumn(("ReadyPercProfil"), System.Type.GetType("System.Decimal")));
            PrepareDSummaryDT.Columns.Add(new DataColumn(("CostTPS"), System.Type.GetType("System.Decimal")));
            PrepareDSummaryDT.Columns.Add(new DataColumn(("CostProfil"), System.Type.GetType("System.Decimal")));

            FrontsSummaryDT = new DataTable();
            FrontsSummaryDT.Columns.Add(new DataColumn(("Front"), System.Type.GetType("System.String")));
            FrontsSummaryDT.Columns.Add(new DataColumn(("FrontID"), System.Type.GetType("System.Int32")));
            FrontsSummaryDT.Columns.Add(new DataColumn(("Width"), System.Type.GetType("System.Int32")));
            FrontsSummaryDT.Columns.Add(new DataColumn(("Square"), System.Type.GetType("System.Decimal")));
            FrontsSummaryDT.Columns.Add(new DataColumn(("Cost"), Type.GetType("System.Decimal")));
            FrontsSummaryDT.Columns.Add(new DataColumn(("Count"), System.Type.GetType("System.Int32")));

            CurvedFrontsSummaryDT = new DataTable();
            CurvedFrontsSummaryDT.Columns.Add(new DataColumn(("Front"), System.Type.GetType("System.String")));
            CurvedFrontsSummaryDT.Columns.Add(new DataColumn(("FrontID"), System.Type.GetType("System.Int32")));
            CurvedFrontsSummaryDT.Columns.Add(new DataColumn(("Width"), System.Type.GetType("System.Int32")));
            CurvedFrontsSummaryDT.Columns.Add(new DataColumn(("Cost"), Type.GetType("System.Decimal")));
            CurvedFrontsSummaryDT.Columns.Add(new DataColumn(("Count"), System.Type.GetType("System.Int32")));

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
            ClientGroupsDataTable = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM ClientGroups ",
                ConnectionStrings.MarketingReferenceConnectionString))
            {
                DA.Fill(ClientGroupsDataTable);
            }
            ClientGroupsDataTable.Columns.Add(new DataColumn("Check", Type.GetType("System.Boolean")));
            for (int i = 0; i < ClientGroupsDataTable.Rows.Count; i++)
            {
                ClientGroupsDataTable.Rows[i]["Check"] = true;
            }
            string SelectCommand = @"SELECT ProductID, ProductName FROM DecorProducts" +
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

            SelectCommand = @"SELECT TechStoreID AS FrontID, TechStoreName AS FrontName FROM TechStore
                WHERE TechStoreID IN (SELECT FrontID FROM FrontsConfig WHERE Enabled = 1)
                ORDER BY TechStoreName";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(FrontsDT);
            }

            //using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM DecorConfig", ConnectionStrings.CatalogConnectionString))
            //{
            //    DA.Fill(DecorConfigDT);
            //}
            DecorConfigDT = TablesManager.DecorConfigDataTable;

            System.Diagnostics.Stopwatch sw = new System.Diagnostics.Stopwatch();
            sw.Start();

            //FillZOVPrepareTables();
            //FillMarketingTables();
            //FillZOVTables();

            sw.Stop();
            double G = sw.Elapsed.TotalMilliseconds;

            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT TOP 0 FrontID, PatinaID, ColorID, InsetTypeID," +
                " InsetColorID, Height, Width, Count, Square, Cost FROM FrontsOrders",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                FrontsOrdersDT.Clear();
                DA.Fill(FrontsOrdersDT);
                CurvedFrontsOrdersDT = FrontsOrdersDT.Clone();
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

        public void FillZOVPrepareTables()
        {
            using (SqlDataAdapter DA = new SqlDataAdapter(
            @"SELECT CONVERT(varchar(10), DocDateTime, 121) AS DocDateTime
            FROM MainOrders WHERE MainOrderID IN (SELECT MainOrderID FROM Packages WHERE PackageStatusID IN (1, 2)) AND MegaOrderID = 0
            GROUP BY CONVERT(varchar(10), DocDateTime, 121)
            ORDER BY DocDateTime",
                ConnectionStrings.ZOVOrdersConnectionString))
            {
                ZOVPrepareDT.Clear();
                DA.Fill(ZOVPrepareDT);
            }

            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT CONVERT(varchar(10), MainOrders.DocDateTime, 121) AS DocDateTime, FrontsOrders.FactoryID, SUM(FrontsOrders.Cost * PackageDetails.Count / FrontsOrders.Count) AS FrontsCost
            FROM PackageDetails INNER JOIN
            FrontsOrders ON PackageDetails.OrderID = FrontsOrders.FrontsOrdersID INNER JOIN
            MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID INNER JOIN
            MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID AND MegaOrders.MegaOrderID = 0
            WHERE PackageDetails.PackageID IN (SELECT PackageID FROM Packages WHERE ProductType = 0 AND PackageStatusID IN (1, 2)) AND FrontsOrders.Width=-1
            GROUP BY CONVERT(varchar(10), MainOrders.DocDateTime, 121), FrontsOrders.FactoryID
            ORDER BY CONVERT(varchar(10), MainOrders.DocDateTime, 121), FrontsOrders.FactoryID",
                ConnectionStrings.ZOVOrdersConnectionString))
            {
                PrepareReadyCurvedFrontsCostDT.Clear();
                DA.Fill(PrepareReadyCurvedFrontsCostDT);
            }

            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT CONVERT(varchar(10), MainOrders.DocDateTime, 121) AS DocDateTime, FrontsOrders.FactoryID, SUM(FrontsOrders.Cost * PackageDetails.Count / FrontsOrders.Count) AS FrontsCost
            FROM PackageDetails INNER JOIN
            FrontsOrders ON PackageDetails.OrderID = FrontsOrders.FrontsOrdersID INNER JOIN
            MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID INNER JOIN
            MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID AND MegaOrders.MegaOrderID = 0
            WHERE PackageDetails.PackageID IN (SELECT PackageID FROM Packages WHERE ProductType = 0 AND PackageStatusID IN (1, 2)) AND FrontsOrders.Width<>-1
            GROUP BY CONVERT(varchar(10), MainOrders.DocDateTime, 121), FrontsOrders.FactoryID
            ORDER BY CONVERT(varchar(10), MainOrders.DocDateTime, 121), FrontsOrders.FactoryID",
                ConnectionStrings.ZOVOrdersConnectionString))
            {
                PrepareReadyFrontsCostDT.Clear();
                DA.Fill(PrepareReadyFrontsCostDT);
            }

            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT CONVERT(varchar(10), MainOrders.DocDateTime, 121) AS DocDateTime, DecorOrders.FactoryID, SUM(DecorOrders.Cost * PackageDetails.Count / DecorOrders.Count) AS DecorCost
            FROM PackageDetails INNER JOIN
            DecorOrders ON PackageDetails.OrderID = DecorOrders.DecorOrderID INNER JOIN
            MainOrders ON DecorOrders.MainOrderID = MainOrders.MainOrderID INNER JOIN
            MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID AND MegaOrders.MegaOrderID = 0
            WHERE PackageDetails.PackageID IN (SELECT PackageID FROM Packages WHERE ProductType = 1 AND PackageStatusID IN (1, 2))
            GROUP BY CONVERT(varchar(10), MainOrders.DocDateTime, 121), DecorOrders.FactoryID
            ORDER BY CONVERT(varchar(10), MainOrders.DocDateTime, 121), DecorOrders.FactoryID",
                ConnectionStrings.ZOVOrdersConnectionString))
            {
                PrepareReadyDecorCostDT.Clear();
                DA.Fill(PrepareReadyDecorCostDT);
            }

            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT CONVERT(varchar(10), MainOrders.DocDateTime, 121) AS DocDateTime, FrontsOrders.FactoryID, SUM(FrontsOrders.Cost) AS FrontsCost
            FROM FrontsOrders INNER JOIN
            MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID INNER JOIN
            MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID AND MegaOrders.MegaOrderID = 0
            WHERE FrontsOrders.FrontsOrdersID IN
            (SELECT OrderID FROM PackageDetails WHERE PackageID IN (SELECT PackageID FROM Packages WHERE ProductType = 0)) AND FrontsOrders.Width=-1
            GROUP BY CONVERT(varchar(10), MainOrders.DocDateTime, 121), FrontsOrders.FactoryID
            ORDER BY CONVERT(varchar(10), MainOrders.DocDateTime, 121), FrontsOrders.FactoryID",
                ConnectionStrings.ZOVOrdersConnectionString))
            {
                PrepareAllCurvedFrontsCostDT.Clear();
                DA.Fill(PrepareAllCurvedFrontsCostDT);
            }

            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT CONVERT(varchar(10), MainOrders.DocDateTime, 121) AS DocDateTime, FrontsOrders.FactoryID, SUM(FrontsOrders.Cost) AS FrontsCost
            FROM FrontsOrders INNER JOIN
            MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID INNER JOIN
            MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID AND MegaOrders.MegaOrderID = 0
            WHERE FrontsOrders.FrontsOrdersID IN
            (SELECT OrderID FROM PackageDetails WHERE PackageID IN (SELECT PackageID FROM Packages WHERE ProductType = 0)) AND FrontsOrders.Width<>-1
            GROUP BY CONVERT(varchar(10), MainOrders.DocDateTime, 121), FrontsOrders.FactoryID
            ORDER BY CONVERT(varchar(10), MainOrders.DocDateTime, 121), FrontsOrders.FactoryID",
                ConnectionStrings.ZOVOrdersConnectionString))
            {
                PrepareAllFrontsCostDT.Clear();
                DA.Fill(PrepareAllFrontsCostDT);
            }

            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT CONVERT(varchar(10), MainOrders.DocDateTime, 121) AS DocDateTime, DecorOrders.FactoryID, SUM(DecorOrders.Cost) AS DecorCost
            FROM DecorOrders INNER JOIN
            MainOrders ON DecorOrders.MainOrderID = MainOrders.MainOrderID INNER JOIN
            MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID AND MegaOrders.MegaOrderID = 0
            WHERE DecorOrders.DecorOrderID IN
            (SELECT OrderID FROM PackageDetails WHERE PackageID IN (SELECT PackageID FROM Packages WHERE ProductType = 1))
            GROUP BY CONVERT(varchar(10), MainOrders.DocDateTime, 121), DecorOrders.FactoryID
            ORDER BY CONVERT(varchar(10), MainOrders.DocDateTime, 121), DecorOrders.FactoryID",
                ConnectionStrings.ZOVOrdersConnectionString))
            {
                PrepareAllDecorCostDT.Clear();
                DA.Fill(PrepareAllDecorCostDT);
            }
        }

        public void FillMarketingTables()
        {
            using (SqlDataAdapter DA = new SqlDataAdapter(
            @"SELECT MegaOrders.ClientID, ClientGroupID, infiniu2_marketingreference.dbo.Clients.ClientName, MegaOrders.OrderNumber, MegaOrders.MegaOrderID
            FROM MegaOrders
            INNER JOIN infiniu2_marketingreference.dbo.Clients ON MegaOrders.ClientID = infiniu2_marketingreference.dbo.Clients.ClientID
            WHERE MegaOrders.MegaOrderID IN (SELECT MegaOrderID FROM MainOrders WHERE MainOrderID IN (SELECT MainOrderID FROM Packages WHERE PackageStatusID IN (1, 2)))
            ORDER BY infiniu2_marketingreference.dbo.Clients.ClientName, MegaOrders.OrderNumber, MegaOrders.MegaOrderID",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                MarketingClientsDT.Clear();
                DA.Fill(MarketingClientsDT);
            }

            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT MegaOrders.ClientID, MainOrders.MegaOrderID, FrontsOrders.FactoryID, SUM(FrontsOrders.Cost * PackageDetails.Count / FrontsOrders.Count) AS FrontsCost
            FROM PackageDetails INNER JOIN
            FrontsOrders ON PackageDetails.OrderID = FrontsOrders.FrontsOrdersID AND FrontsOrders.Width=-1 INNER JOIN
            MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID INNER JOIN
            MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID
            WHERE PackageDetails.PackageID IN (SELECT PackageID FROM Packages WHERE ProductType = 0 AND PackageStatusID IN (1, 2))
            GROUP BY MegaOrders.ClientID, MainOrders.MegaOrderID, FrontsOrders.FactoryID
            ORDER BY MegaOrders.ClientID, MainOrders.MegaOrderID, FrontsOrders.FactoryID",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                MReadyCurvedFrontsCostDT.Clear();
                DA.Fill(MReadyCurvedFrontsCostDT);
            }

            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT MegaOrders.ClientID, MainOrders.MegaOrderID, FrontsOrders.FactoryID, SUM(FrontsOrders.Cost * PackageDetails.Count / FrontsOrders.Count) AS FrontsCost
            FROM PackageDetails INNER JOIN
            FrontsOrders ON PackageDetails.OrderID = FrontsOrders.FrontsOrdersID AND FrontsOrders.Width<>-1 INNER JOIN
            MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID INNER JOIN
            MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID
            WHERE PackageDetails.PackageID IN (SELECT PackageID FROM Packages WHERE ProductType = 0 AND PackageStatusID IN (1, 2))
            GROUP BY MegaOrders.ClientID, MainOrders.MegaOrderID, FrontsOrders.FactoryID
            ORDER BY MegaOrders.ClientID, MainOrders.MegaOrderID, FrontsOrders.FactoryID",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                MReadyFrontsCostDT.Clear();
                DA.Fill(MReadyFrontsCostDT);
            }

            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT MegaOrders.ClientID, MainOrders.MegaOrderID, DecorOrders.FactoryID, SUM(DecorOrders.Cost * PackageDetails.Count / DecorOrders.Count) AS DecorCost
            FROM PackageDetails INNER JOIN
            DecorOrders ON PackageDetails.OrderID = DecorOrders.DecorOrderID INNER JOIN
            MainOrders ON DecorOrders.MainOrderID = MainOrders.MainOrderID INNER JOIN
            MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID
            WHERE PackageDetails.PackageID IN (SELECT PackageID FROM Packages WHERE ProductType = 1 AND PackageStatusID IN (1, 2))
            GROUP BY MegaOrders.ClientID, MainOrders.MegaOrderID, DecorOrders.FactoryID
            ORDER BY MegaOrders.ClientID, MainOrders.MegaOrderID, DecorOrders.FactoryID",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                MReadyDecorCostDT.Clear();
                DA.Fill(MReadyDecorCostDT);
            }

            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT MegaOrders.ClientID, MainOrders.MegaOrderID, FrontsOrders.FactoryID, SUM(FrontsOrders.Cost) AS FrontsCost
            FROM FrontsOrders INNER JOIN
            MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID INNER JOIN
            MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID
            WHERE FrontsOrders.FrontsOrdersID IN
            (SELECT OrderID FROM PackageDetails WHERE PackageID IN (SELECT PackageID FROM Packages WHERE ProductType = 0)) AND FrontsOrders.Width=-1
            GROUP BY MegaOrders.ClientID, MainOrders.MegaOrderID, FrontsOrders.FactoryID
            ORDER BY MegaOrders.ClientID, MainOrders.MegaOrderID, FrontsOrders.FactoryID",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                MAllCurvedFrontsCostDT.Clear();
                DA.Fill(MAllCurvedFrontsCostDT);
            }

            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT MegaOrders.ClientID, MainOrders.MegaOrderID, FrontsOrders.FactoryID, SUM(FrontsOrders.Cost) AS FrontsCost
            FROM FrontsOrders INNER JOIN
            MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID INNER JOIN
            MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID
            WHERE FrontsOrders.FrontsOrdersID IN
            (SELECT OrderID FROM PackageDetails WHERE PackageID IN (SELECT PackageID FROM Packages WHERE ProductType = 0)) AND FrontsOrders.Width<>-1
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
            ClientGroupsBS = new BindingSource()
            {
                DataSource = ClientGroupsDataTable
            };
            MCurvedFSummaryBS = new BindingSource()
            {
                DataSource = MCurvedFSummaryDT
            };
            MCurvedFSummaryDG.DataSource = MCurvedFSummaryBS;

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

            PrepareCurvedFSummaryBS = new BindingSource()
            {
                DataSource = PrepareCurvedFSummaryDT
            };
            PrepareCurvedFSummaryDG.DataSource = PrepareCurvedFSummaryBS;

            PrepareFSummaryBS = new BindingSource()
            {
                DataSource = PrepareFSummaryDT
            };
            PrepareFSummaryDG.DataSource = PrepareFSummaryBS;

            PrepareDSummaryBS = new BindingSource()
            {
                DataSource = PrepareDSummaryDT
            };
            PrepareDSummaryDG.DataSource = PrepareDSummaryBS;

            CurvedFrontsSummaryBS = new BindingSource()
            {
                DataSource = CurvedFrontsSummaryDT
            };
            CurvedFrontsDG.DataSource = CurvedFrontsSummaryBS;

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
            PrepareSummaryGridSettings();
        }

        public void ShowColumns(ref PercentageDataGrid FrontsGrid, ref PercentageDataGrid CurvedFrontsGrid, ref PercentageDataGrid DecorGrid, bool Profil, bool TPS, bool bClientSummary)
        {
            if (Profil && TPS)
            {
                FrontsGrid.Columns["CostProfil"].Visible = true;
                FrontsGrid.Columns["CostTPS"].Visible = true;
                FrontsGrid.Columns["ReadyPercProfil"].Visible = true;
                FrontsGrid.Columns["ReadyPercTPS"].Visible = true;

                FrontsGrid.Columns["CostProfil"].Visible = true;
                FrontsGrid.Columns["CostTPS"].Visible = true;
                FrontsGrid.Columns["ReadyPercProfil"].Visible = true;
                FrontsGrid.Columns["ReadyPercTPS"].Visible = true;

                DecorGrid.Columns["CostProfil"].Visible = true;
                DecorGrid.Columns["CostTPS"].Visible = true;
                DecorGrid.Columns["ReadyPercProfil"].Visible = true;
                DecorGrid.Columns["ReadyPercTPS"].Visible = true;
            }
            if (Profil && !TPS)
            {
                FrontsGrid.Columns["CostProfil"].Visible = true;
                FrontsGrid.Columns["CostTPS"].Visible = false;
                FrontsGrid.Columns["ReadyPercProfil"].Visible = true;
                FrontsGrid.Columns["ReadyPercTPS"].Visible = false;

                FrontsGrid.Columns["CostProfil"].Visible = true;
                FrontsGrid.Columns["CostTPS"].Visible = false;
                FrontsGrid.Columns["ReadyPercProfil"].Visible = true;

                DecorGrid.Columns["CostProfil"].Visible = true;
                DecorGrid.Columns["CostTPS"].Visible = false;
                DecorGrid.Columns["ReadyPercProfil"].Visible = true;
                DecorGrid.Columns["ReadyPercTPS"].Visible = false;
            }
            if (!Profil && TPS)
            {
                FrontsGrid.Columns["CostProfil"].Visible = false;
                FrontsGrid.Columns["CostTPS"].Visible = true;
                FrontsGrid.Columns["ReadyPercProfil"].Visible = false;
                FrontsGrid.Columns["ReadyPercTPS"].Visible = true;

                FrontsGrid.Columns["CostProfil"].Visible = false;
                FrontsGrid.Columns["CostTPS"].Visible = true;
                FrontsGrid.Columns["ReadyPercProfil"].Visible = false;
                FrontsGrid.Columns["ReadyPercTPS"].Visible = true;
                DecorGrid.Columns["CostProfil"].Visible = false;
                DecorGrid.Columns["CostTPS"].Visible = true;
                DecorGrid.Columns["ReadyPercProfil"].Visible = false;
                DecorGrid.Columns["ReadyPercTPS"].Visible = true;
            }
        }

        private void MarketSummaryGridSettings()
        {
            MFSummaryDG.Columns["MegaOrderID"].Visible = false;
            MFSummaryDG.Columns["ClientID"].Visible = false;
            MCurvedFSummaryDG.Columns["MegaOrderID"].Visible = false;
            MCurvedFSummaryDG.Columns["ClientID"].Visible = false;
            MDSummaryDG.Columns["MegaOrderID"].Visible = false;
            MDSummaryDG.Columns["ClientID"].Visible = false;

            MFSummaryDG.Columns["AllCostProfil"].Visible = false;
            MCurvedFSummaryDG.Columns["AllCostProfil"].Visible = false;
            MDSummaryDG.Columns["AllCostProfil"].Visible = false;
            MFSummaryDG.Columns["AllCostTPS"].Visible = false;
            MCurvedFSummaryDG.Columns["AllCostTPS"].Visible = false;
            MDSummaryDG.Columns["AllCostTPS"].Visible = false;

            if (!Security.PriceAccess)
            {
                MFSummaryDG.Columns["CostProfil"].Visible = false;
                MCurvedFSummaryDG.Columns["CostProfil"].Visible = false;
                MFSummaryDG.Columns["CostTPS"].Visible = false;
                MCurvedFSummaryDG.Columns["CostTPS"].Visible = false;
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
            foreach (DataGridViewColumn Column in MCurvedFSummaryDG.Columns)
            {
                Column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
            foreach (DataGridViewColumn Column in MFSummaryDG.Columns)
            {
                Column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
            foreach (DataGridViewColumn Column in MDSummaryDG.Columns)
            {
                Column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }

            MCurvedFSummaryDG.Columns["DocDateTime"].DefaultCellStyle.Format = "dd.MM.yyyy";
            MCurvedFSummaryDG.Columns["DocDateTime"].HeaderText = "Дата создания";
            MCurvedFSummaryDG.Columns["ClientName"].HeaderText = "Клиент";
            MCurvedFSummaryDG.Columns["OrderNumber"].HeaderText = "№ заказа";
            MCurvedFSummaryDG.Columns["ReadyPercProfil"].HeaderText = "Произведено\n\r  Профиль, %";
            MCurvedFSummaryDG.Columns["ReadyPercTPS"].HeaderText = "Произведено\n\r      ТПС, %";
            MCurvedFSummaryDG.Columns["CostProfil"].HeaderText = "Профиль, €";
            MCurvedFSummaryDG.Columns["CostTPS"].HeaderText = "ТПС, €";
            MCurvedFSummaryDG.Columns["DocDateTime"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            MCurvedFSummaryDG.Columns["DocDateTime"].MinimumWidth = 110;
            MCurvedFSummaryDG.Columns["ClientName"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            MCurvedFSummaryDG.Columns["ClientName"].MinimumWidth = 110;
            MCurvedFSummaryDG.Columns["OrderNumber"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MCurvedFSummaryDG.Columns["OrderNumber"].Width = 85;
            MCurvedFSummaryDG.Columns["ReadyPercProfil"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            MCurvedFSummaryDG.Columns["ReadyPercProfil"].MinimumWidth = 135;
            MCurvedFSummaryDG.Columns["ReadyPercTPS"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            MCurvedFSummaryDG.Columns["ReadyPercTPS"].MinimumWidth = 135;
            MCurvedFSummaryDG.Columns["CostProfil"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MCurvedFSummaryDG.Columns["CostProfil"].Width = 110;
            MCurvedFSummaryDG.Columns["CostTPS"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MCurvedFSummaryDG.Columns["CostTPS"].Width = 110;

            MCurvedFSummaryDG.Columns["CostProfil"].DefaultCellStyle.Format = "N";
            MCurvedFSummaryDG.Columns["CostProfil"].DefaultCellStyle.FormatProvider = nfi1;
            MCurvedFSummaryDG.Columns["CostTPS"].DefaultCellStyle.Format = "N";
            MCurvedFSummaryDG.Columns["CostTPS"].DefaultCellStyle.FormatProvider = nfi1;

            MCurvedFSummaryDG.Columns["ReadyPercProfil"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            MCurvedFSummaryDG.AddPercentageColumn("ReadyPercProfil");
            MCurvedFSummaryDG.Columns["ReadyPercTPS"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            MCurvedFSummaryDG.AddPercentageColumn("ReadyPercTPS");

            MFSummaryDG.Columns["DocDateTime"].DefaultCellStyle.Format = "dd.MM.yyyy";
            MFSummaryDG.Columns["DocDateTime"].HeaderText = "Дата создания";
            MFSummaryDG.Columns["ClientName"].HeaderText = "Клиент";
            MFSummaryDG.Columns["OrderNumber"].HeaderText = "№ заказа";
            MFSummaryDG.Columns["ReadyPercProfil"].HeaderText = "Произведено\n\r  Профиль, %";
            MFSummaryDG.Columns["ReadyPercTPS"].HeaderText = "Произведено\n\r      ТПС, %";
            MFSummaryDG.Columns["CostProfil"].HeaderText = "Профиль, €";
            MFSummaryDG.Columns["CostTPS"].HeaderText = "ТПС, €";
            MFSummaryDG.Columns["DocDateTime"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            MFSummaryDG.Columns["DocDateTime"].MinimumWidth = 110;
            MFSummaryDG.Columns["ClientName"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            MFSummaryDG.Columns["ClientName"].MinimumWidth = 110;
            MFSummaryDG.Columns["OrderNumber"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MFSummaryDG.Columns["OrderNumber"].Width = 85;
            MFSummaryDG.Columns["ReadyPercProfil"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            MFSummaryDG.Columns["ReadyPercProfil"].MinimumWidth = 135;
            MFSummaryDG.Columns["ReadyPercTPS"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            MFSummaryDG.Columns["ReadyPercTPS"].MinimumWidth = 135;
            MFSummaryDG.Columns["CostProfil"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MFSummaryDG.Columns["CostProfil"].Width = 110;
            MFSummaryDG.Columns["CostTPS"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MFSummaryDG.Columns["CostTPS"].Width = 110;

            MFSummaryDG.Columns["CostProfil"].DefaultCellStyle.Format = "N";
            MFSummaryDG.Columns["CostProfil"].DefaultCellStyle.FormatProvider = nfi1;
            MFSummaryDG.Columns["CostTPS"].DefaultCellStyle.Format = "N";
            MFSummaryDG.Columns["CostTPS"].DefaultCellStyle.FormatProvider = nfi1;

            MFSummaryDG.Columns["ReadyPercProfil"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            MFSummaryDG.AddPercentageColumn("ReadyPercProfil");
            MFSummaryDG.Columns["ReadyPercTPS"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            MFSummaryDG.AddPercentageColumn("ReadyPercTPS");

            MDSummaryDG.Columns["DocDateTime"].DefaultCellStyle.Format = "dd.MM.yyyy";
            MDSummaryDG.Columns["DocDateTime"].HeaderText = "Дата создания";
            MDSummaryDG.Columns["ClientName"].HeaderText = "Клиент";
            MDSummaryDG.Columns["OrderNumber"].HeaderText = "№ заказа";
            MDSummaryDG.Columns["ReadyPercProfil"].HeaderText = "Произведено\n\r  Профиль, %";
            MDSummaryDG.Columns["ReadyPercTPS"].HeaderText = "Произведено\n\r      ТПС, %";
            MDSummaryDG.Columns["CostProfil"].HeaderText = "Профиль, €";
            MDSummaryDG.Columns["CostTPS"].HeaderText = "ТПС, €";
            MDSummaryDG.Columns["DocDateTime"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            MDSummaryDG.Columns["DocDateTime"].MinimumWidth = 110;
            MDSummaryDG.Columns["ClientName"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            MDSummaryDG.Columns["ClientName"].MinimumWidth = 110;
            MDSummaryDG.Columns["OrderNumber"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MDSummaryDG.Columns["OrderNumber"].Width = 85;
            MDSummaryDG.Columns["ReadyPercProfil"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            MDSummaryDG.Columns["ReadyPercProfil"].MinimumWidth = 135;
            MDSummaryDG.Columns["ReadyPercTPS"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            MDSummaryDG.Columns["ReadyPercTPS"].MinimumWidth = 135;
            MDSummaryDG.Columns["CostProfil"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MDSummaryDG.Columns["CostProfil"].Width = 110;
            MDSummaryDG.Columns["CostTPS"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MDSummaryDG.Columns["CostTPS"].Width = 110;

            MDSummaryDG.Columns["CostProfil"].DefaultCellStyle.Format = "N";
            MDSummaryDG.Columns["CostProfil"].DefaultCellStyle.FormatProvider = nfi1;
            MDSummaryDG.Columns["CostTPS"].DefaultCellStyle.Format = "N";
            MDSummaryDG.Columns["CostTPS"].DefaultCellStyle.FormatProvider = nfi1;

            MDSummaryDG.Columns["ReadyPercProfil"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            MDSummaryDG.AddPercentageColumn("ReadyPercProfil");
            MDSummaryDG.Columns["ReadyPercTPS"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            MDSummaryDG.AddPercentageColumn("ReadyPercTPS");
        }

        private void PrepareSummaryGridSettings()
        {
            if (!Security.PriceAccess)
            {
                PrepareCurvedFSummaryDG.Columns["CostProfil"].Visible = false;
                PrepareCurvedFSummaryDG.Columns["CostTPS"].Visible = false;
                PrepareFSummaryDG.Columns["CostProfil"].Visible = false;
                PrepareFSummaryDG.Columns["CostTPS"].Visible = false;
                PrepareDSummaryDG.Columns["CostProfil"].Visible = false;
                PrepareDSummaryDG.Columns["CostTPS"].Visible = false;
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
            foreach (DataGridViewColumn Column in PrepareCurvedFSummaryDG.Columns)
            {
                Column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
            foreach (DataGridViewColumn Column in PrepareFSummaryDG.Columns)
            {
                Column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
            foreach (DataGridViewColumn Column in PrepareDSummaryDG.Columns)
            {
                Column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }

            PrepareCurvedFSummaryDG.Columns["DocDateTime"].DefaultCellStyle.Format = "dd.MM.yyyy";
            PrepareCurvedFSummaryDG.Columns["DocDateTime"].HeaderText = "Дата создания";
            PrepareCurvedFSummaryDG.Columns["ReadyPercProfil"].HeaderText = "Произведено\n\r  Профиль, %";
            PrepareCurvedFSummaryDG.Columns["ReadyPercTPS"].HeaderText = "Произведено\n\r      ТПС, %";
            PrepareCurvedFSummaryDG.Columns["CostProfil"].HeaderText = "Профиль, €";
            PrepareCurvedFSummaryDG.Columns["CostTPS"].HeaderText = "ТПС, €";
            PrepareCurvedFSummaryDG.Columns["DocDateTime"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            PrepareCurvedFSummaryDG.Columns["DocDateTime"].MinimumWidth = 110;
            PrepareCurvedFSummaryDG.Columns["ReadyPercProfil"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            PrepareCurvedFSummaryDG.Columns["ReadyPercProfil"].MinimumWidth = 135;
            PrepareCurvedFSummaryDG.Columns["ReadyPercTPS"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            PrepareCurvedFSummaryDG.Columns["ReadyPercTPS"].MinimumWidth = 135;
            PrepareCurvedFSummaryDG.Columns["CostProfil"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            PrepareCurvedFSummaryDG.Columns["CostProfil"].Width = 110;
            PrepareCurvedFSummaryDG.Columns["CostTPS"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            PrepareCurvedFSummaryDG.Columns["CostTPS"].Width = 110;
            PrepareCurvedFSummaryDG.Columns["CostProfil"].DefaultCellStyle.Format = "N";
            PrepareCurvedFSummaryDG.Columns["CostProfil"].DefaultCellStyle.FormatProvider = nfi1;
            PrepareCurvedFSummaryDG.Columns["CostTPS"].DefaultCellStyle.Format = "N";
            PrepareCurvedFSummaryDG.Columns["CostTPS"].DefaultCellStyle.FormatProvider = nfi1;

            PrepareCurvedFSummaryDG.Columns["ReadyPercProfil"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            PrepareCurvedFSummaryDG.AddPercentageColumn("ReadyPercProfil");
            PrepareCurvedFSummaryDG.Columns["ReadyPercTPS"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            PrepareCurvedFSummaryDG.AddPercentageColumn("ReadyPercTPS");

            PrepareFSummaryDG.Columns["DocDateTime"].DefaultCellStyle.Format = "dd.MM.yyyy";
            PrepareFSummaryDG.Columns["DocDateTime"].HeaderText = "Дата создания";
            PrepareFSummaryDG.Columns["ReadyPercProfil"].HeaderText = "Произведено\n\r  Профиль, %";
            PrepareFSummaryDG.Columns["ReadyPercTPS"].HeaderText = "Произведено\n\r      ТПС, %";
            PrepareFSummaryDG.Columns["CostProfil"].HeaderText = "Профиль, €";
            PrepareFSummaryDG.Columns["CostTPS"].HeaderText = "ТПС, €";
            PrepareFSummaryDG.Columns["DocDateTime"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            PrepareFSummaryDG.Columns["DocDateTime"].MinimumWidth = 110;
            PrepareFSummaryDG.Columns["ReadyPercProfil"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            PrepareFSummaryDG.Columns["ReadyPercProfil"].MinimumWidth = 135;
            PrepareFSummaryDG.Columns["ReadyPercTPS"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            PrepareFSummaryDG.Columns["ReadyPercTPS"].MinimumWidth = 135;
            PrepareFSummaryDG.Columns["CostProfil"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            PrepareFSummaryDG.Columns["CostProfil"].Width = 110;
            PrepareFSummaryDG.Columns["CostTPS"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            PrepareFSummaryDG.Columns["CostTPS"].Width = 110;
            PrepareFSummaryDG.Columns["CostProfil"].DefaultCellStyle.Format = "N";
            PrepareFSummaryDG.Columns["CostProfil"].DefaultCellStyle.FormatProvider = nfi1;
            PrepareFSummaryDG.Columns["CostTPS"].DefaultCellStyle.Format = "N";
            PrepareFSummaryDG.Columns["CostTPS"].DefaultCellStyle.FormatProvider = nfi1;

            PrepareFSummaryDG.Columns["ReadyPercProfil"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            PrepareFSummaryDG.AddPercentageColumn("ReadyPercProfil");
            PrepareFSummaryDG.Columns["ReadyPercTPS"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            PrepareFSummaryDG.AddPercentageColumn("ReadyPercTPS");

            PrepareDSummaryDG.Columns["DocDateTime"].DefaultCellStyle.Format = "dd.MM.yyyy";
            PrepareDSummaryDG.Columns["DocDateTime"].HeaderText = "Дата создания";
            PrepareDSummaryDG.Columns["ReadyPercProfil"].HeaderText = "Произведено\n\r  Профиль, %";
            PrepareDSummaryDG.Columns["ReadyPercTPS"].HeaderText = "Произведено\n\r      ТПС, %";
            PrepareDSummaryDG.Columns["CostProfil"].HeaderText = "Профиль, €";
            PrepareDSummaryDG.Columns["CostTPS"].HeaderText = "ТПС, €";
            PrepareDSummaryDG.Columns["DocDateTime"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            PrepareDSummaryDG.Columns["DocDateTime"].MinimumWidth = 110;
            PrepareDSummaryDG.Columns["ReadyPercProfil"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            PrepareDSummaryDG.Columns["ReadyPercProfil"].MinimumWidth = 135;
            PrepareDSummaryDG.Columns["ReadyPercTPS"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            PrepareDSummaryDG.Columns["ReadyPercTPS"].MinimumWidth = 135;
            PrepareDSummaryDG.Columns["CostProfil"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            PrepareDSummaryDG.Columns["CostProfil"].Width = 110;
            PrepareDSummaryDG.Columns["CostTPS"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            PrepareDSummaryDG.Columns["CostTPS"].Width = 110;
            PrepareDSummaryDG.Columns["CostProfil"].DefaultCellStyle.Format = "N";
            PrepareDSummaryDG.Columns["CostProfil"].DefaultCellStyle.FormatProvider = nfi1;
            PrepareDSummaryDG.Columns["CostTPS"].DefaultCellStyle.Format = "N";
            PrepareDSummaryDG.Columns["CostTPS"].DefaultCellStyle.FormatProvider = nfi1;

            PrepareDSummaryDG.Columns["ReadyPercProfil"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            PrepareDSummaryDG.AddPercentageColumn("ReadyPercProfil");
            PrepareDSummaryDG.Columns["ReadyPercTPS"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            PrepareDSummaryDG.AddPercentageColumn("ReadyPercTPS");
        }

        private void SetProductsGrids()
        {
            if (!Security.PriceAccess)
            {
                CurvedFrontsDG.Columns["Cost"].Visible = false;
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
            CurvedFrontsDG.ColumnHeadersHeight = 38;
            FrontsDG.ColumnHeadersHeight = 38;
            DecorProductsDG.ColumnHeadersHeight = 38;
            DecorItemsDG.ColumnHeadersHeight = 38;

            foreach (DataGridViewColumn Column in CurvedFrontsDG.Columns)
            {
                Column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
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

            CurvedFrontsDG.Columns["FrontID"].Visible = false;
            CurvedFrontsDG.Columns["Width"].Visible = false;

            CurvedFrontsDG.Columns["Front"].HeaderText = "Фасад";
            CurvedFrontsDG.Columns["Cost"].HeaderText = " € ";
            CurvedFrontsDG.Columns["Count"].HeaderText = "шт.";

            CurvedFrontsDG.Columns["Front"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            CurvedFrontsDG.Columns["Front"].MinimumWidth = 110;
            CurvedFrontsDG.Columns["Count"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            CurvedFrontsDG.Columns["Cost"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            CurvedFrontsDG.Columns["Cost"].Width = 100;
            CurvedFrontsDG.Columns["Count"].Width = 90;

            CurvedFrontsDG.Columns["Cost"].DefaultCellStyle.Format = "N";
            CurvedFrontsDG.Columns["Cost"].DefaultCellStyle.FormatProvider = nfi1;

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

            CurvedFrontsDG.Columns["Cost"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            CurvedFrontsDG.Columns["Count"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

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

            decimal ReadyPercProfil = 0;
            decimal ReadyPercTPS = 0;

            decimal ReadyCostProfil = 0;
            decimal ReadyCostTPS = 0;

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

            string filter = string.Empty;
            for (int i = 0; i < ClientGroupsDataTable.Rows.Count; i++)
            {
                if (!Convert.ToBoolean(ClientGroupsDataTable.Rows[i]["Check"]))
                    continue;

                filter += ClientGroupsDataTable.Rows[i]["ClientGroupID"].ToString() + ",";
            }
            if (filter.Length > 0)
                filter = "ClientGroupID IN (" + filter + ")";
            using (DataView DV = new DataView(MarketingClientsDT, filter, string.Empty, DataViewRowState.CurrentRows))
            {
                Table = DV.ToTable(true, new string[] { "ClientID" });
            }

            for (int i = 0; i < MarketingClientsDT.Rows.Count; i++)
            {
                ClientID = Convert.ToInt32(MarketingClientsDT.Rows[i]["ClientID"]);
                DataRow[] erows = Table.Select("ClientID=" + ClientID);
                if (erows.Count() == 0)
                    continue;
                MegaOrderID = Convert.ToInt32(MarketingClientsDT.Rows[i]["MegaOrderID"]);

                ClientName = MarketingClientsDT.Rows[i]["ClientName"].ToString();

                if (FactoryID == 0)
                {
                    ReadyCostProfil = 0;
                    ReadyCostTPS = 0;
                    ReadyPercProfil = MarketReadyCurvedFrontsCost(true, MegaOrderID, 1, ref ReadyCostProfil, ref AllCostProfil);
                    ReadyPercTPS = MarketReadyCurvedFrontsCost(true, MegaOrderID, 2, ref ReadyCostTPS, ref AllCostTPS);

                    d1 = 0;
                    d2 = 0;

                    if (ReadyPercProfil != 0 || ReadyPercTPS != 0)
                    {
                        DataRow[] Rows = MCurvedFSummaryDT.Select("ClientName='" + ClientName + "'");
                        if (Rows.Count() == 0)
                        {
                            DataRow NewRow = MCurvedFSummaryDT.NewRow();
                            NewRow["ClientID"] = ClientID;
                            NewRow["ClientName"] = ClientName;
                            if (ReadyPercProfil > 0)
                            {
                                NewRow["ReadyPercProfil"] = ReadyPercProfil;
                                NewRow["CostProfil"] = ReadyCostProfil;
                                NewRow["AllCostProfil"] = AllCostProfil;
                            }
                            if (ReadyPercTPS > 0)
                            {
                                NewRow["ReadyPercTPS"] = ReadyPercTPS;
                                NewRow["CostTPS"] = ReadyCostTPS;
                                NewRow["AllCostTPS"] = AllCostTPS;
                            }
                            MCurvedFSummaryDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            if (ReadyPercProfil > 0)
                            {
                                if (Rows[0]["CostProfil"] != DBNull.Value)
                                    d1 = Convert.ToDecimal(Rows[0]["CostProfil"]);
                                if (Rows[0]["AllCostProfil"] != DBNull.Value)
                                    d2 = Convert.ToDecimal(Rows[0]["AllCostProfil"]);
                                Percentage = (d1 + ReadyCostProfil) / (d2 + AllCostProfil);
                                Percentage = Decimal.Round(Percentage * 100, 2, MidpointRounding.AwayFromZero);
                                Rows[0]["ReadyPercProfil"] = Percentage;
                                Rows[0]["AllCostProfil"] = d2 + AllCostProfil;
                                Rows[0]["CostProfil"] = d1 + ReadyCostProfil;
                            }
                            if (ReadyPercTPS > 0)
                            {
                                if (Rows[0]["CostTPS"] != DBNull.Value)
                                    d1 = Convert.ToDecimal(Rows[0]["CostTPS"]);
                                if (Rows[0]["AllCostTPS"] != DBNull.Value)
                                    d2 = Convert.ToDecimal(Rows[0]["AllCostTPS"]);
                                Percentage = (d1 + ReadyCostTPS) / (d2 + AllCostTPS);
                                Percentage = Decimal.Round(Percentage * 100, 2, MidpointRounding.AwayFromZero);
                                Rows[0]["ReadyPercTPS"] = Percentage;
                                Rows[0]["AllCostTPS"] = d2 + AllCostTPS;
                                Rows[0]["CostTPS"] = d1 + ReadyCostTPS;
                            }
                        }
                    }

                    ReadyCostProfil = 0;
                    ReadyCostTPS = 0;
                    ReadyPercProfil = MarketReadyFrontsCost(true, MegaOrderID, 1, ref ReadyCostProfil, ref AllCostProfil);
                    ReadyPercTPS = MarketReadyFrontsCost(true, MegaOrderID, 2, ref ReadyCostTPS, ref AllCostTPS);

                    d1 = 0;
                    d2 = 0;

                    if (ReadyPercProfil != 0 || ReadyPercTPS != 0)
                    {
                        DataRow[] Rows = MFSummaryDT.Select("ClientName='" + ClientName + "'");
                        if (Rows.Count() == 0)
                        {
                            DataRow NewRow = MFSummaryDT.NewRow();
                            NewRow["ClientID"] = ClientID;
                            NewRow["ClientName"] = ClientName;
                            if (ReadyPercProfil > 0)
                            {
                                NewRow["ReadyPercProfil"] = ReadyPercProfil;
                                NewRow["CostProfil"] = ReadyCostProfil;
                                NewRow["AllCostProfil"] = AllCostProfil;
                            }
                            if (ReadyPercTPS > 0)
                            {
                                NewRow["ReadyPercTPS"] = ReadyPercTPS;
                                NewRow["CostTPS"] = ReadyCostTPS;
                                NewRow["AllCostTPS"] = AllCostTPS;
                            }
                            MFSummaryDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            if (ReadyPercProfil > 0)
                            {
                                if (Rows[0]["CostProfil"] != DBNull.Value)
                                    d1 = Convert.ToDecimal(Rows[0]["CostProfil"]);
                                if (Rows[0]["AllCostProfil"] != DBNull.Value)
                                    d2 = Convert.ToDecimal(Rows[0]["AllCostProfil"]);
                                Percentage = (d1 + ReadyCostProfil) / (d2 + AllCostProfil);
                                Percentage = Decimal.Round(Percentage * 100, 2, MidpointRounding.AwayFromZero);
                                Rows[0]["ReadyPercProfil"] = Percentage;
                                Rows[0]["AllCostProfil"] = d2 + AllCostProfil;
                                Rows[0]["CostProfil"] = d1 + ReadyCostProfil;
                            }
                            if (ReadyPercTPS > 0)
                            {
                                if (Rows[0]["CostTPS"] != DBNull.Value)
                                    d1 = Convert.ToDecimal(Rows[0]["CostTPS"]);
                                if (Rows[0]["AllCostTPS"] != DBNull.Value)
                                    d2 = Convert.ToDecimal(Rows[0]["AllCostTPS"]);
                                Percentage = (d1 + ReadyCostTPS) / (d2 + AllCostTPS);
                                Percentage = Decimal.Round(Percentage * 100, 2, MidpointRounding.AwayFromZero);
                                Rows[0]["ReadyPercTPS"] = Percentage;
                                Rows[0]["AllCostTPS"] = d2 + AllCostTPS;
                                Rows[0]["CostTPS"] = d1 + ReadyCostTPS;
                            }
                        }
                    }

                    ReadyCostProfil = 0;
                    ReadyCostTPS = 0;
                    ReadyPercProfil = MarketReadyDecorCost(true, MegaOrderID, 1, ref ReadyCostProfil, ref AllCostProfil);
                    ReadyPercTPS = MarketReadyDecorCost(true, MegaOrderID, 2, ref ReadyCostTPS, ref AllCostTPS);

                    if (ReadyCostProfil != 0 || ReadyCostTPS != 0)
                    {
                        DataRow[] Rows = MDSummaryDT.Select("ClientName='" + ClientName + "'");
                        if (Rows.Count() == 0)
                        {
                            DataRow NewRow = MDSummaryDT.NewRow();
                            NewRow["ClientID"] = ClientID;
                            NewRow["ClientName"] = ClientName;
                            if (ReadyCostProfil > 0)
                            {
                                NewRow["ReadyPercProfil"] = ReadyPercProfil;
                                NewRow["CostProfil"] = ReadyCostProfil;
                                NewRow["AllCostProfil"] = AllCostProfil;
                            }
                            if (ReadyCostTPS > 0)
                            {
                                NewRow["ReadyPercTPS"] = ReadyPercTPS;
                                NewRow["CostTPS"] = ReadyCostTPS;
                                NewRow["AllCostTPS"] = AllCostTPS;
                            }
                            MDSummaryDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            if (ReadyPercProfil > 0)
                            {
                                if (Rows[0]["CostProfil"] != DBNull.Value)
                                    d1 = Convert.ToDecimal(Rows[0]["CostProfil"]);
                                if (Rows[0]["AllCostProfil"] != DBNull.Value)
                                    d2 = Convert.ToDecimal(Rows[0]["AllCostProfil"]);
                                Percentage = (d1 + ReadyCostProfil) / (d2 + AllCostProfil);
                                Percentage = Decimal.Round(Percentage * 100, 2, MidpointRounding.AwayFromZero);
                                Rows[0]["ReadyPercProfil"] = Percentage;
                                Rows[0]["AllCostProfil"] = d2 + AllCostProfil;
                                Rows[0]["CostProfil"] = d1 + ReadyCostProfil;
                            }
                            if (ReadyPercTPS > 0)
                            {
                                if (Rows[0]["CostTPS"] != DBNull.Value)
                                    d1 = Convert.ToDecimal(Rows[0]["CostTPS"]);
                                if (Rows[0]["AllCostTPS"] != DBNull.Value)
                                    d2 = Convert.ToDecimal(Rows[0]["AllCostTPS"]);
                                Percentage = (d1 + ReadyCostTPS) / (d2 + AllCostTPS);
                                Percentage = Decimal.Round(Percentage * 100, 2, MidpointRounding.AwayFromZero);
                                Rows[0]["ReadyPercTPS"] = Percentage;
                                Rows[0]["AllCostTPS"] = d2 + AllCostTPS;
                                Rows[0]["CostTPS"] = d1 + ReadyCostTPS;
                            }
                        }
                    }
                }

                if (FactoryID == 1)
                {
                    ReadyCostProfil = 0;
                    ReadyPercProfil = MarketReadyCurvedFrontsCost(true, MegaOrderID, 1, ref ReadyCostProfil, ref AllCostProfil);

                    if (ReadyPercProfil > 0)
                    {
                        DataRow[] Rows = MCurvedFSummaryDT.Select("ClientName='" + ClientName + "'");
                        if (Rows.Count() == 0)
                        {
                            DataRow NewRow = MCurvedFSummaryDT.NewRow();
                            NewRow["ClientID"] = ClientID;
                            NewRow["ClientName"] = ClientName;
                            NewRow["ReadyPercProfil"] = ReadyPercProfil;
                            NewRow["CostProfil"] = ReadyCostProfil;
                            NewRow["AllCostProfil"] = AllCostProfil;
                            MCurvedFSummaryDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            if (Rows[0]["CostProfil"] != DBNull.Value)
                                d1 = Convert.ToDecimal(Rows[0]["CostProfil"]);
                            if (Rows[0]["AllCostProfil"] != DBNull.Value)
                                d2 = Convert.ToDecimal(Rows[0]["AllCostProfil"]);
                            Percentage = (d1 + ReadyCostProfil) / (d2 + AllCostProfil);
                            Percentage = Decimal.Round(Percentage * 100, 2, MidpointRounding.AwayFromZero);
                            Rows[0]["ReadyPercProfil"] = Percentage;
                            Rows[0]["AllCostProfil"] = d2 + AllCostProfil;
                            Rows[0]["CostProfil"] = d1 + ReadyCostProfil;
                        }
                    }

                    ReadyCostProfil = 0;
                    ReadyPercProfil = MarketReadyFrontsCost(true, MegaOrderID, 1, ref ReadyCostProfil, ref AllCostProfil);

                    if (ReadyPercProfil > 0)
                    {
                        DataRow[] Rows = MFSummaryDT.Select("ClientName='" + ClientName + "'");
                        if (Rows.Count() == 0)
                        {
                            DataRow NewRow = MFSummaryDT.NewRow();
                            NewRow["ClientID"] = ClientID;
                            NewRow["ClientName"] = ClientName;
                            NewRow["ReadyPercProfil"] = ReadyPercProfil;
                            NewRow["CostProfil"] = ReadyCostProfil;
                            NewRow["AllCostProfil"] = AllCostProfil;
                            MFSummaryDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            if (Rows[0]["CostProfil"] != DBNull.Value)
                                d1 = Convert.ToDecimal(Rows[0]["CostProfil"]);
                            if (Rows[0]["AllCostProfil"] != DBNull.Value)
                                d2 = Convert.ToDecimal(Rows[0]["AllCostProfil"]);
                            Percentage = (d1 + ReadyCostProfil) / (d2 + AllCostProfil);
                            Percentage = Decimal.Round(Percentage * 100, 2, MidpointRounding.AwayFromZero);
                            Rows[0]["ReadyPercProfil"] = Percentage;
                            Rows[0]["AllCostProfil"] = d2 + AllCostProfil;
                            Rows[0]["CostProfil"] = d1 + ReadyCostProfil;
                        }
                    }

                    ReadyCostProfil = 0;
                    ReadyPercProfil = MarketReadyDecorCost(true, MegaOrderID, 1, ref ReadyCostProfil, ref AllCostProfil);

                    if (ReadyCostProfil > 0)
                    {
                        DataRow[] Rows = MDSummaryDT.Select("ClientName='" + ClientName + "'");
                        if (Rows.Count() == 0)
                        {
                            DataRow NewRow = MDSummaryDT.NewRow();
                            NewRow["ClientID"] = ClientID;
                            NewRow["ClientName"] = ClientName;
                            NewRow["ReadyPercProfil"] = ReadyPercProfil;
                            NewRow["CostProfil"] = ReadyCostProfil;
                            NewRow["AllCostProfil"] = AllCostProfil;
                            MDSummaryDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            if (Rows[0]["CostProfil"] != DBNull.Value)
                                d1 = Convert.ToDecimal(Rows[0]["CostProfil"]);
                            if (Rows[0]["AllCostProfil"] != DBNull.Value)
                                d2 = Convert.ToDecimal(Rows[0]["AllCostProfil"]);
                            Percentage = (d1 + ReadyCostProfil) / (d2 + AllCostProfil);
                            Percentage = Decimal.Round(Percentage * 100, 2, MidpointRounding.AwayFromZero);
                            Rows[0]["ReadyPercProfil"] = Percentage;
                            Rows[0]["AllCostProfil"] = d2 + AllCostProfil;
                            Rows[0]["CostProfil"] = d1 + ReadyCostProfil;
                        }
                    }
                }

                if (FactoryID == 2)
                {
                    ReadyCostTPS = 0;
                    ReadyPercTPS = MarketReadyCurvedFrontsCost(true, MegaOrderID, 2, ref ReadyCostTPS, ref AllCostTPS);

                    if (ReadyPercTPS > 0)
                    {
                        DataRow[] Rows = MCurvedFSummaryDT.Select("ClientName='" + ClientName + "'");
                        if (Rows.Count() == 0)
                        {
                            DataRow NewRow = MCurvedFSummaryDT.NewRow();
                            NewRow["ClientID"] = ClientID;
                            NewRow["ClientName"] = ClientName;
                            NewRow["ReadyPercTPS"] = ReadyPercTPS;
                            NewRow["CostTPS"] = ReadyCostTPS;
                            NewRow["AllCostTPS"] = AllCostTPS;
                            MCurvedFSummaryDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            if (Rows[0]["CostTPS"] != DBNull.Value)
                                d1 = Convert.ToDecimal(Rows[0]["CostTPS"]);
                            if (Rows[0]["AllCostTPS"] != DBNull.Value)
                                d2 = Convert.ToDecimal(Rows[0]["AllCostTPS"]);
                            Percentage = (d1 + ReadyCostTPS) / (d2 + AllCostTPS);
                            Percentage = Decimal.Round(Percentage * 100, 2, MidpointRounding.AwayFromZero);
                            Rows[0]["ReadyPercTPS"] = Percentage;
                            Percentage = Decimal.Round(Percentage * 100, 2, MidpointRounding.AwayFromZero);
                            Rows[0]["ExpPercTPS"] = Percentage;
                            Rows[0]["AllCostTPS"] = d2 + AllCostTPS;
                            Rows[0]["CostTPS"] = d1 + ReadyCostTPS;
                        }
                    }

                    ReadyCostTPS = 0;
                    ReadyPercTPS = MarketReadyFrontsCost(true, MegaOrderID, 2, ref ReadyCostTPS, ref AllCostTPS);

                    if (ReadyPercTPS > 0)
                    {
                        DataRow[] Rows = MFSummaryDT.Select("ClientName='" + ClientName + "'");
                        if (Rows.Count() == 0)
                        {
                            DataRow NewRow = MFSummaryDT.NewRow();
                            NewRow["ClientID"] = ClientID;
                            NewRow["ClientName"] = ClientName;
                            NewRow["ReadyPercTPS"] = ReadyPercTPS;
                            NewRow["CostTPS"] = ReadyCostTPS;
                            NewRow["AllCostTPS"] = AllCostTPS;
                            MFSummaryDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            if (Rows[0]["CostTPS"] != DBNull.Value)
                                d1 = Convert.ToDecimal(Rows[0]["CostTPS"]);
                            if (Rows[0]["AllCostTPS"] != DBNull.Value)
                                d2 = Convert.ToDecimal(Rows[0]["AllCostTPS"]);
                            Percentage = (d1 + ReadyCostTPS) / (d2 + AllCostTPS);
                            Percentage = Decimal.Round(Percentage * 100, 2, MidpointRounding.AwayFromZero);
                            Rows[0]["ReadyPercTPS"] = Percentage;
                            Percentage = Decimal.Round(Percentage * 100, 2, MidpointRounding.AwayFromZero);
                            Rows[0]["ExpPercTPS"] = Percentage;
                            Rows[0]["AllCostTPS"] = d2 + AllCostTPS;
                            Rows[0]["CostTPS"] = d1 + ReadyCostTPS;
                        }
                    }

                    ReadyCostTPS = 0;
                    ReadyPercTPS = MarketReadyDecorCost(true, MegaOrderID, 2, ref ReadyCostTPS, ref AllCostTPS);

                    if (ReadyCostTPS > 0)
                    {
                        DataRow[] Rows = MDSummaryDT.Select("ClientName='" + ClientName + "'");
                        if (Rows.Count() == 0)
                        {
                            DataRow NewRow = MDSummaryDT.NewRow();
                            NewRow["ClientID"] = ClientID;
                            NewRow["ClientName"] = ClientName;
                            NewRow["ReadyPercTPS"] = ReadyPercTPS;
                            NewRow["CostTPS"] = ReadyCostTPS;
                            NewRow["AllCostTPS"] = AllCostTPS;
                            MDSummaryDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            if (Rows[0]["CostTPS"] != DBNull.Value)
                                d1 = Convert.ToDecimal(Rows[0]["CostTPS"]);
                            if (Rows[0]["AllCostTPS"] != DBNull.Value)
                                d2 = Convert.ToDecimal(Rows[0]["AllCostTPS"]);
                            Percentage = (d1 + ReadyCostTPS) / (d2 + AllCostTPS);
                            Percentage = Decimal.Round(Percentage * 100, 2, MidpointRounding.AwayFromZero);
                            Rows[0]["ReadyPercTPS"] = Percentage;
                            Rows[0]["AllCostTPS"] = d2 + AllCostTPS;
                            Rows[0]["CostTPS"] = d1 + ReadyCostTPS;
                        }
                    }
                }
            }

            MFSummaryDT.DefaultView.Sort = "ClientName";
            MCurvedFSummaryDT.DefaultView.Sort = "ClientName";
            MDSummaryDT.DefaultView.Sort = "ClientName";
            MFSummaryBS.MoveFirst();
            MCurvedFSummaryBS.MoveFirst();
            MDSummaryBS.MoveFirst();
        }

        public void MarketingSummary(int FactoryID)
        {
            int ClientID = 0;
            int MegaOrderID = 0;
            int OrderNumber = 0;

            decimal ReadyPercProfil = 0;
            decimal ReadyPercTPS = 0;

            decimal ReadyCostProfil = 0;
            decimal ReadyCostTPS = 0;

            decimal AllCostProfil = 0;
            decimal AllCostTPS = 0;

            string ClientName = string.Empty;

            object DocDateTime = null;

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
                    ReadyCostProfil = 0;
                    ReadyCostTPS = 0;
                    ReadyPercProfil = MarketReadyCurvedFrontsCost(true, MegaOrderID, 1, ref ReadyCostProfil, ref AllCostProfil);
                    ReadyPercTPS = MarketReadyCurvedFrontsCost(true, MegaOrderID, 2, ref ReadyCostTPS, ref AllCostTPS);

                    if (ReadyPercProfil != 0 || ReadyPercTPS != 0)
                    {
                        DataRow NewRow = MCurvedFSummaryDT.NewRow();
                        NewRow["DocDateTime"] = DocDateTime;
                        NewRow["ClientID"] = ClientID;
                        NewRow["ClientName"] = ClientName;
                        NewRow["OrderNumber"] = OrderNumber;
                        NewRow["MegaOrderID"] = MegaOrderID;
                        if (ReadyPercProfil > 0)
                        {
                            NewRow["ReadyPercProfil"] = ReadyPercProfil;
                            NewRow["CostProfil"] = ReadyCostProfil;
                        }
                        if (ReadyPercTPS > 0)
                        {
                            NewRow["ReadyPercTPS"] = ReadyPercTPS;
                            NewRow["CostTPS"] = ReadyCostTPS;
                        }
                        MCurvedFSummaryDT.Rows.Add(NewRow);
                    }

                    ReadyCostProfil = 0;
                    ReadyCostTPS = 0;
                    ReadyPercProfil = MarketReadyFrontsCost(true, MegaOrderID, 1, ref ReadyCostProfil, ref AllCostProfil);
                    ReadyPercTPS = MarketReadyFrontsCost(true, MegaOrderID, 2, ref ReadyCostTPS, ref AllCostTPS);

                    if (ReadyPercProfil != 0 || ReadyPercTPS != 0)
                    {
                        DataRow NewRow = MFSummaryDT.NewRow();
                        NewRow["DocDateTime"] = DocDateTime;
                        NewRow["ClientID"] = ClientID;
                        NewRow["ClientName"] = ClientName;
                        NewRow["OrderNumber"] = OrderNumber;
                        NewRow["MegaOrderID"] = MegaOrderID;
                        if (ReadyPercProfil > 0)
                        {
                            NewRow["ReadyPercProfil"] = ReadyPercProfil;
                            NewRow["CostProfil"] = ReadyCostProfil;
                        }
                        if (ReadyPercTPS > 0)
                        {
                            NewRow["ReadyPercTPS"] = ReadyPercTPS;
                            NewRow["CostTPS"] = ReadyCostTPS;
                        }
                        MFSummaryDT.Rows.Add(NewRow);
                    }

                    ReadyCostProfil = 0;
                    ReadyCostTPS = 0;
                    ReadyPercProfil = MarketReadyDecorCost(true, MegaOrderID, 1, ref ReadyCostProfil, ref AllCostProfil);
                    ReadyPercTPS = MarketReadyDecorCost(true, MegaOrderID, 2, ref ReadyCostTPS, ref AllCostTPS);

                    if (ReadyCostProfil != 0 || ReadyCostTPS != 0)
                    {
                        DataRow NewRow = MDSummaryDT.NewRow();
                        NewRow["DocDateTime"] = DocDateTime;
                        NewRow["ClientID"] = ClientID;
                        NewRow["ClientName"] = ClientName;
                        NewRow["OrderNumber"] = OrderNumber;
                        NewRow["MegaOrderID"] = MegaOrderID;
                        if (ReadyCostProfil > 0)
                        {
                            NewRow["ReadyPercProfil"] = ReadyPercProfil;
                            NewRow["CostProfil"] = ReadyCostProfil;
                        }
                        if (ReadyCostTPS > 0)
                        {
                            NewRow["ReadyPercTPS"] = ReadyPercTPS;
                            NewRow["CostTPS"] = ReadyCostTPS;
                        }
                        MDSummaryDT.Rows.Add(NewRow);
                    }
                }

                if (FactoryID == 1)
                {
                    ReadyCostProfil = 0;
                    ReadyPercProfil = MarketReadyCurvedFrontsCost(true, MegaOrderID, 1, ref ReadyCostProfil, ref AllCostProfil);

                    if (ReadyPercProfil > 0)
                    {
                        DataRow NewRow = MCurvedFSummaryDT.NewRow();
                        NewRow["DocDateTime"] = DocDateTime;
                        NewRow["ClientID"] = ClientID;
                        NewRow["ClientName"] = ClientName;
                        NewRow["OrderNumber"] = OrderNumber;
                        NewRow["MegaOrderID"] = MegaOrderID;
                        NewRow["ReadyPercProfil"] = ReadyPercProfil;
                        NewRow["CostProfil"] = ReadyCostProfil;
                        MCurvedFSummaryDT.Rows.Add(NewRow);
                    }

                    ReadyCostProfil = 0;
                    ReadyPercProfil = MarketReadyFrontsCost(true, MegaOrderID, 1, ref ReadyCostProfil, ref AllCostProfil);

                    if (ReadyPercProfil > 0)
                    {
                        DataRow NewRow = MFSummaryDT.NewRow();
                        NewRow["DocDateTime"] = DocDateTime;
                        NewRow["ClientID"] = ClientID;
                        NewRow["ClientName"] = ClientName;
                        NewRow["OrderNumber"] = OrderNumber;
                        NewRow["MegaOrderID"] = MegaOrderID;
                        NewRow["ReadyPercProfil"] = ReadyPercProfil;
                        NewRow["CostProfil"] = ReadyCostProfil;
                        MFSummaryDT.Rows.Add(NewRow);
                    }

                    ReadyCostProfil = 0;
                    ReadyPercProfil = MarketReadyDecorCost(true, MegaOrderID, 1, ref ReadyCostProfil, ref AllCostProfil);

                    if (ReadyCostProfil > 0)
                    {
                        DataRow NewRow = MDSummaryDT.NewRow();
                        NewRow["DocDateTime"] = DocDateTime;
                        NewRow["ClientName"] = ClientID;
                        NewRow["OrderNumber"] = OrderNumber;
                        NewRow["MegaOrderID"] = MegaOrderID;
                        NewRow["ReadyPercProfil"] = ReadyPercProfil;
                        NewRow["CostProfil"] = ReadyCostProfil;
                        MDSummaryDT.Rows.Add(NewRow);
                    }
                }

                if (FactoryID == 2)
                {
                    ReadyCostTPS = 0;
                    ReadyPercTPS = MarketReadyCurvedFrontsCost(true, MegaOrderID, 2, ref ReadyCostTPS, ref AllCostTPS);

                    if (ReadyPercTPS > 0)
                    {
                        DataRow NewRow = MCurvedFSummaryDT.NewRow();
                        NewRow["DocDateTime"] = DocDateTime;
                        NewRow["ClientID"] = ClientID;
                        NewRow["ClientName"] = ClientName;
                        NewRow["OrderNumber"] = OrderNumber;
                        NewRow["MegaOrderID"] = MegaOrderID;
                        NewRow["ReadyPercTPS"] = ReadyPercTPS;
                        NewRow["CostTPS"] = ReadyCostTPS;
                        MCurvedFSummaryDT.Rows.Add(NewRow);
                    }

                    ReadyCostTPS = 0;
                    ReadyPercTPS = MarketReadyFrontsCost(true, MegaOrderID, 2, ref ReadyCostTPS, ref AllCostTPS);

                    if (ReadyPercTPS > 0)
                    {
                        DataRow NewRow = MFSummaryDT.NewRow();
                        NewRow["DocDateTime"] = DocDateTime;
                        NewRow["ClientID"] = ClientID;
                        NewRow["ClientName"] = ClientName;
                        NewRow["OrderNumber"] = OrderNumber;
                        NewRow["MegaOrderID"] = MegaOrderID;
                        NewRow["ReadyPercTPS"] = ReadyPercTPS;
                        NewRow["CostTPS"] = ReadyCostTPS;
                        MFSummaryDT.Rows.Add(NewRow);
                    }

                    ReadyCostTPS = 0;
                    ReadyPercTPS = MarketReadyDecorCost(true, MegaOrderID, 2, ref ReadyCostTPS, ref AllCostTPS);

                    if (ReadyCostTPS > 0)
                    {
                        DataRow NewRow = MDSummaryDT.NewRow();
                        NewRow["DocDateTime"] = DocDateTime;
                        NewRow["ClientID"] = ClientID;
                        NewRow["ClientName"] = ClientName;
                        NewRow["OrderNumber"] = OrderNumber;
                        NewRow["MegaOrderID"] = MegaOrderID;
                        NewRow["ReadyPercTPS"] = ReadyPercTPS;
                        NewRow["CostTPS"] = ReadyCostTPS;
                        MDSummaryDT.Rows.Add(NewRow);
                    }
                }
            }

            MFSummaryDT.DefaultView.Sort = "ClientName";
            MCurvedFSummaryDT.DefaultView.Sort = "ClientName";
            MDSummaryDT.DefaultView.Sort = "ClientName";
            MFSummaryBS.MoveFirst();
            MCurvedFSummaryBS.MoveFirst();
            MDSummaryBS.MoveFirst();
        }

        public void PrepareSummary(int FactoryID)
        {
            decimal ReadyPercProfil = 0;
            decimal ReadyPercTPS = 0;

            decimal ReadyCostProfil = 0;
            decimal ReadyCostTPS = 0;
            decimal AllCostProfil = 0;
            decimal AllCostTPS = 0;

            string DocDateTime = string.Empty;

            PrepareFSummaryDT.Clear();
            PrepareDSummaryDT.Clear();

            if (FactoryID == -1)
                return;

            for (int i = 0; i < ZOVPrepareDT.Rows.Count; i++)
            {
                DocDateTime = ZOVPrepareDT.Rows[i]["DocDateTime"].ToString();

                if (FactoryID == 0)
                {
                    ReadyCostProfil = 0;
                    ReadyCostTPS = 0;
                    ReadyPercProfil = PrepareReadyCurvedFrontsCost(DocDateTime, 1, ref ReadyCostProfil, ref AllCostProfil);
                    ReadyPercTPS = PrepareReadyCurvedFrontsCost(DocDateTime, 2, ref ReadyCostTPS, ref AllCostTPS);

                    if (ReadyPercProfil != 0 || ReadyPercTPS != 0)
                    {
                        DataRow NewRow = PrepareCurvedFSummaryDT.NewRow();
                        NewRow["DocDateTime"] = DocDateTime;
                        if (ReadyPercProfil > 0)
                        {
                            NewRow["ReadyPercProfil"] = ReadyPercProfil;
                            NewRow["CostProfil"] = ReadyCostProfil;
                        }
                        if (ReadyPercTPS > 0)
                        {
                            NewRow["ReadyPercTPS"] = ReadyPercTPS;
                            NewRow["CostTPS"] = ReadyCostTPS;
                        }
                        PrepareCurvedFSummaryDT.Rows.Add(NewRow);
                    }

                    ReadyCostProfil = 0;
                    ReadyCostTPS = 0;
                    ReadyPercProfil = PrepareReadyFrontsCost(DocDateTime, 1, ref ReadyCostProfil, ref AllCostProfil);
                    ReadyPercTPS = PrepareReadyFrontsCost(DocDateTime, 2, ref ReadyCostTPS, ref AllCostTPS);

                    if (ReadyPercProfil != 0 || ReadyPercTPS != 0)
                    {
                        DataRow NewRow = PrepareFSummaryDT.NewRow();
                        NewRow["DocDateTime"] = DocDateTime;
                        if (ReadyPercProfil > 0)
                        {
                            NewRow["ReadyPercProfil"] = ReadyPercProfil;
                            NewRow["CostProfil"] = ReadyCostProfil;
                        }
                        if (ReadyPercTPS > 0)
                        {
                            NewRow["ReadyPercTPS"] = ReadyPercTPS;
                            NewRow["CostTPS"] = ReadyCostTPS;
                        }
                        PrepareFSummaryDT.Rows.Add(NewRow);
                    }

                    ReadyCostProfil = 0;
                    ReadyCostTPS = 0;
                    ReadyPercProfil = PrepareReadyDecorCost(DocDateTime, 1, ref ReadyCostProfil, ref AllCostProfil);
                    ReadyPercTPS = PrepareReadyDecorCost(DocDateTime, 2, ref ReadyCostTPS, ref AllCostTPS);

                    if (ReadyCostProfil != 0 || ReadyCostTPS != 0)
                    {
                        DataRow NewRow = PrepareDSummaryDT.NewRow();
                        NewRow["DocDateTime"] = DocDateTime;
                        if (ReadyCostProfil > 0)
                        {
                            NewRow["ReadyPercProfil"] = ReadyPercProfil;
                            NewRow["CostProfil"] = ReadyCostProfil;
                        }
                        if (ReadyCostTPS > 0)
                        {
                            NewRow["ReadyPercTPS"] = ReadyPercTPS;
                            NewRow["CostTPS"] = ReadyCostTPS;
                        }
                        PrepareDSummaryDT.Rows.Add(NewRow);
                    }
                }
                if (FactoryID == 1)
                {
                    ReadyCostProfil = 0;
                    ReadyPercProfil = PrepareReadyCurvedFrontsCost(DocDateTime, 1, ref ReadyCostProfil, ref AllCostProfil);

                    if (ReadyPercProfil > 0)
                    {
                        DataRow NewRow = PrepareCurvedFSummaryDT.NewRow();
                        NewRow["DocDateTime"] = DocDateTime;
                        NewRow["ReadyPercProfil"] = ReadyPercProfil;
                        NewRow["CostProfil"] = ReadyCostProfil;
                        PrepareCurvedFSummaryDT.Rows.Add(NewRow);
                    }

                    ReadyCostProfil = 0;
                    ReadyPercProfil = PrepareReadyFrontsCost(DocDateTime, 1, ref ReadyCostProfil, ref AllCostProfil);

                    if (ReadyPercProfil > 0)
                    {
                        DataRow NewRow = PrepareFSummaryDT.NewRow();
                        NewRow["DocDateTime"] = DocDateTime;
                        NewRow["ReadyPercProfil"] = ReadyPercProfil;
                        NewRow["CostProfil"] = ReadyCostProfil;
                        PrepareFSummaryDT.Rows.Add(NewRow);
                    }

                    ReadyCostProfil = 0;
                    ReadyPercProfil = PrepareReadyDecorCost(DocDateTime, 1, ref ReadyCostProfil, ref AllCostProfil);

                    if (ReadyPercProfil > 0)
                    {
                        DataRow NewRow = PrepareDSummaryDT.NewRow();
                        NewRow["DocDateTime"] = DocDateTime;
                        NewRow["ReadyPercProfil"] = ReadyPercProfil;
                        NewRow["CostProfil"] = ReadyCostProfil;
                        PrepareDSummaryDT.Rows.Add(NewRow);
                    }
                }
                if (FactoryID == 2)
                {
                    ReadyCostTPS = 0;
                    ReadyPercTPS = PrepareReadyCurvedFrontsCost(DocDateTime, 2, ref ReadyCostTPS, ref AllCostTPS);

                    if (ReadyCostTPS > 0)
                    {
                        DataRow NewRow = PrepareCurvedFSummaryDT.NewRow();
                        NewRow["DocDateTime"] = DocDateTime;
                        NewRow["ReadyPercTPS"] = ReadyPercTPS;
                        NewRow["CostTPS"] = ReadyCostTPS;
                        PrepareCurvedFSummaryDT.Rows.Add(NewRow);
                    }

                    ReadyCostTPS = 0;
                    ReadyPercTPS = PrepareReadyFrontsCost(DocDateTime, 2, ref ReadyCostTPS, ref AllCostTPS);

                    if (ReadyCostTPS > 0)
                    {
                        DataRow NewRow = PrepareFSummaryDT.NewRow();
                        NewRow["DocDateTime"] = DocDateTime;
                        NewRow["ReadyPercTPS"] = ReadyPercTPS;
                        NewRow["CostTPS"] = ReadyCostTPS;
                        PrepareFSummaryDT.Rows.Add(NewRow);
                    }

                    ReadyCostTPS = 0;
                    ReadyPercTPS = PrepareReadyDecorCost(DocDateTime, 2, ref ReadyCostTPS, ref AllCostTPS);

                    if (ReadyCostTPS > 0)
                    {
                        DataRow NewRow = PrepareDSummaryDT.NewRow();
                        NewRow["DocDateTime"] = DocDateTime;
                        NewRow["ReadyPercTPS"] = ReadyPercTPS;
                        NewRow["CostTPS"] = ReadyCostTPS;
                        PrepareDSummaryDT.Rows.Add(NewRow);
                    }
                }
            }

            PrepareFSummaryBS.MoveFirst();
            PrepareFSummaryBS.MoveFirst();
            PrepareDSummaryBS.MoveFirst();
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
                " (SELECT PackageID FROM Packages WHERE PackageStatusID IN (1, 2) AND ProductType = 0 " + PackageFactoryFilter + ")";

            MarketingSelectCommand = "SELECT FrontsOrders.FrontID, FrontsOrders.PatinaID, FrontsOrders.ColorID, FrontsOrders.InsetTypeID," +
                " FrontsOrders.InsetColorID, FrontsOrders.TechnoColorID, FrontsOrders.TechnoInsetTypeID, FrontsOrders.TechnoInsetColorID, FrontsOrders.Height, FrontsOrders.Width, PackageDetails.Count, (FrontsOrders.Square * PackageDetails.Count / FrontsOrders.Count) AS Square, MeasureID, (FrontsOrders.Cost * PackageDetails.Count / FrontsOrders.Count) AS Cost, ClientID FROM PackageDetails" +
                " INNER JOIN FrontsOrders ON PackageDetails.OrderID = FrontsOrders.FrontsOrdersID AND FrontsOrders.Width<>-1" +
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
                " (SELECT PackageID FROM Packages WHERE PackageStatusID IN (1, 2) AND ProductType = 0 " + PackageFactoryFilter + ")";

            MarketingSelectCommand = "SELECT FrontsOrders.FrontID, FrontsOrders.PatinaID, FrontsOrders.ColorID, FrontsOrders.InsetTypeID," +
                " FrontsOrders.InsetColorID, FrontsOrders.TechnoColorID, FrontsOrders.TechnoInsetTypeID, FrontsOrders.TechnoInsetColorID, FrontsOrders.Height, FrontsOrders.Width, PackageDetails.Count, (FrontsOrders.Square * PackageDetails.Count / FrontsOrders.Count) AS Square, MeasureID, (FrontsOrders.Cost * PackageDetails.Count / FrontsOrders.Count) AS Cost, ClientID FROM PackageDetails" +
                " INNER JOIN FrontsOrders ON PackageDetails.OrderID = FrontsOrders.FrontsOrdersID AND FrontsOrders.Width<>-1" +
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

        public void CurvedFMarketingOrders1(int FactoryID, int MegaOrderID)
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
                " (SELECT PackageID FROM Packages WHERE PackageStatusID IN (1, 2) AND ProductType = 0 " + PackageFactoryFilter + ")";

            MarketingSelectCommand = "SELECT FrontsOrders.FrontID, FrontsOrders.PatinaID, FrontsOrders.ColorID, FrontsOrders.InsetTypeID," +
                " FrontsOrders.InsetColorID, FrontsOrders.TechnoColorID, FrontsOrders.TechnoInsetTypeID, FrontsOrders.TechnoInsetColorID, FrontsOrders.Height, FrontsOrders.Width, PackageDetails.Count, (FrontsOrders.Square * PackageDetails.Count / FrontsOrders.Count) AS Square, MeasureID, (FrontsOrders.Cost * PackageDetails.Count / FrontsOrders.Count) AS Cost, ClientID FROM PackageDetails" +
                " INNER JOIN FrontsOrders ON PackageDetails.OrderID = FrontsOrders.FrontsOrdersID AND FrontsOrders.Width=-1" +
                " INNER JOIN MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID" +
                " INNER JOIN MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID" + MegaOrderFilter +
                " INNER JOIN infiniu2_catalog.dbo.FrontsConfig" +
                " ON FrontsOrders.FrontConfigID=infiniu2_catalog.dbo.FrontsConfig.FrontConfigID" +
                " WHERE " + MFrontsPackageFilter;

            using (SqlDataAdapter DA = new SqlDataAdapter(MarketingSelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                CurvedFrontsOrdersDT.Clear();
                DA.Fill(CurvedFrontsOrdersDT);
            }

            GetCurvedFronts();
        }

        public void CurvedFMarketingOrders(int FactoryID, int ClientID)
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
                " (SELECT PackageID FROM Packages WHERE PackageStatusID IN (1, 2) AND ProductType = 0 " + PackageFactoryFilter + ")";

            MarketingSelectCommand = "SELECT FrontsOrders.FrontID, FrontsOrders.PatinaID, FrontsOrders.ColorID, FrontsOrders.InsetTypeID," +
                " FrontsOrders.InsetColorID, FrontsOrders.TechnoColorID, FrontsOrders.TechnoInsetTypeID, FrontsOrders.TechnoInsetColorID, FrontsOrders.Height, FrontsOrders.Width, PackageDetails.Count, (FrontsOrders.Square * PackageDetails.Count / FrontsOrders.Count) AS Square, MeasureID, (FrontsOrders.Cost * PackageDetails.Count / FrontsOrders.Count) AS Cost, ClientID FROM PackageDetails" +
                " INNER JOIN FrontsOrders ON PackageDetails.OrderID = FrontsOrders.FrontsOrdersID AND FrontsOrders.Width=-1" +
                " INNER JOIN MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID" +
                " INNER JOIN MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID" + ClientFilter +
                " INNER JOIN infiniu2_catalog.dbo.FrontsConfig" +
                " ON FrontsOrders.FrontConfigID=infiniu2_catalog.dbo.FrontsConfig.FrontConfigID" +
                " WHERE " + MFrontsPackageFilter;

            using (SqlDataAdapter DA = new SqlDataAdapter(MarketingSelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                CurvedFrontsOrdersDT.Clear();
                DA.Fill(CurvedFrontsOrdersDT);
            }

            GetCurvedFronts();
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
                " (SELECT PackageID FROM Packages WHERE PackageStatusID IN (1, 2) AND ProductType = 1 " + PackageFactoryFilter + ")";

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
                " (SELECT PackageID FROM Packages WHERE PackageStatusID IN (1, 2) AND ProductType = 1 " + PackageFactoryFilter + ")";

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

        public void PrepareOrders(int FactoryID)
        {
            string ZOVSelectCommand = string.Empty;

            string ZFrontsPackageFilter = string.Empty;
            string ZDecorPackageFilter = string.Empty;

            string PackageFactoryFilter = string.Empty;

            if (FactoryID != 0)
                PackageFactoryFilter = " AND FactoryID = " + FactoryID;

            ZFrontsPackageFilter = " PackageID IN" +
                " (SELECT PackageID FROM Packages WHERE PackageStatusID IN (1, 2) AND ProductType = 0 " + PackageFactoryFilter + ")";

            ZDecorPackageFilter = " PackageID IN" +
                " (SELECT PackageID FROM Packages WHERE PackageStatusID IN (1, 2) AND ProductType = 1 " + PackageFactoryFilter + ")";

            ZOVSelectCommand = "SELECT FrontsOrders.FrontID, FrontsOrders.PatinaID," +
                " FrontsOrders.ColorID, FrontsOrders.InsetTypeID," +
                " FrontsOrders.InsetColorID, FrontsOrders.TechnoColorID, FrontsOrders.TechnoInsetTypeID, FrontsOrders.TechnoInsetColorID, FrontsOrders.Height, FrontsOrders.Width," +
                " PackageDetails.Count, (FrontsOrders.Square * PackageDetails.Count / FrontsOrders.Count) AS Square, (FrontsOrders.Cost * PackageDetails.Count / FrontsOrders.Count) AS Cost, MeasureID FROM PackageDetails" +
                " INNER JOIN FrontsOrders ON PackageDetails.OrderID = FrontsOrders.FrontsOrdersID AND FrontsOrders.Width<>-1" +
                " INNER JOIN infiniu2_catalog.dbo.FrontsConfig" +
                " ON FrontsOrders.FrontConfigID=infiniu2_catalog.dbo.FrontsConfig.FrontConfigID" +
                " WHERE MainOrderID IN (SELECT MainOrderID FROM MainOrders WHERE MegaOrderID = 0) AND " + ZFrontsPackageFilter;

            using (SqlDataAdapter DA = new SqlDataAdapter(ZOVSelectCommand,
                ConnectionStrings.ZOVOrdersConnectionString))
            {
                FrontsOrdersDT.Clear();
                DA.Fill(FrontsOrdersDT);
            }

            GetFronts();

            ZOVSelectCommand = "SELECT FrontsOrders.FrontID, FrontsOrders.PatinaID," +
                " FrontsOrders.ColorID, FrontsOrders.InsetTypeID," +
                " FrontsOrders.InsetColorID, FrontsOrders.TechnoColorID, FrontsOrders.TechnoInsetTypeID, FrontsOrders.TechnoInsetColorID, FrontsOrders.Height, FrontsOrders.Width," +
                " PackageDetails.Count, (FrontsOrders.Square * PackageDetails.Count / FrontsOrders.Count) AS Square, (FrontsOrders.Cost * PackageDetails.Count / FrontsOrders.Count) AS Cost, MeasureID FROM PackageDetails" +
                " INNER JOIN FrontsOrders ON PackageDetails.OrderID = FrontsOrders.FrontsOrdersID AND FrontsOrders.Width=-1" +
                " INNER JOIN infiniu2_catalog.dbo.FrontsConfig" +
                " ON FrontsOrders.FrontConfigID=infiniu2_catalog.dbo.FrontsConfig.FrontConfigID" +
                " WHERE MainOrderID IN (SELECT MainOrderID FROM MainOrders WHERE MegaOrderID = 0) AND " + ZFrontsPackageFilter;

            using (SqlDataAdapter DA = new SqlDataAdapter(ZOVSelectCommand,
                ConnectionStrings.ZOVOrdersConnectionString))
            {
                CurvedFrontsOrdersDT.Clear();
                DA.Fill(CurvedFrontsOrdersDT);
            }

            GetCurvedFronts();

            //decor
            ZOVSelectCommand = "SELECT DecorOrders.ProductID, DecorOrders.DecorID, DecorOrders.ColorID,DecorOrders.PatinaID," +
                " DecorOrders.Height, DecorOrders.Length, DecorOrders.Width, PackageDetails.Count," +
                " (DecorOrders.Cost * PackageDetails.Count / DecorOrders.Count) AS Cost, DecorOrders.DecorConfigID, " +
                " MeasureID FROM PackageDetails" +
                " INNER JOIN DecorOrders ON PackageDetails.OrderID = DecorOrders.DecorOrderID" +
                " INNER JOIN infiniu2_catalog.dbo.DecorConfig" +
                " ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                " WHERE MainOrderID IN (SELECT MainOrderID FROM MainOrders WHERE MegaOrderID = 0) AND " + ZDecorPackageFilter;

            using (SqlDataAdapter DA = new SqlDataAdapter(ZOVSelectCommand,
                ConnectionStrings.ZOVOrdersConnectionString))
            {
                DecorOrdersDT.Clear();
                DA.Fill(DecorOrdersDT);
            }

            GetDecorProducts();
            GetDecorItems();

            PrepareSummary(FactoryID);
        }

        public bool HasCurvedFronts
        {
            get { return CurvedFrontsOrdersDT.Rows.Count > 0; }
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
                DataRow[] Rows = FrontsOrdersDT.Select("FrontID=" + Convert.ToInt32(Table.Rows[i]["FrontID"]));
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
                    FrontsSummaryDT.Rows.Add(NewRow);

                    FrontCost = 0;
                    FrontSquare = 0;
                    FrontCount = 0;
                }
            }

            Table.Dispose();
            FrontsSummaryDT.DefaultView.Sort = "Front, Square DESC";
            FrontsSummaryBS.MoveFirst();
        }

        private void GetCurvedFronts()
        {
            decimal FrontCost = 0;
            int FrontCount = 0;

            CurvedFrontsSummaryDT.Clear();

            DataTable Table = new DataTable();

            using (DataView DV = new DataView(CurvedFrontsOrdersDT))
            {
                Table = DV.ToTable(true, new string[] { "FrontID" });
            }

            for (int i = 0; i < Table.Rows.Count; i++)
            {
                DataRow[] Rows = CurvedFrontsOrdersDT.Select("FrontID=" + Convert.ToInt32(Table.Rows[i]["FrontID"]));
                if (Rows.Count() != 0)
                {
                    foreach (DataRow row in Rows)
                    {
                        FrontCost += Convert.ToDecimal(row["Cost"]);
                        FrontCount += Convert.ToInt32(row["Count"]);
                    }

                    DataRow NewRow = CurvedFrontsSummaryDT.NewRow();
                    NewRow["Front"] = GetFrontName(Convert.ToInt32(Table.Rows[i]["FrontID"]));
                    NewRow["FrontID"] = Convert.ToInt32(Table.Rows[i]["FrontID"]);
                    NewRow["Cost"] = Decimal.Round(FrontCost, 2, MidpointRounding.AwayFromZero);
                    NewRow["Count"] = FrontCount;
                    NewRow["Width"] = 0;
                    CurvedFrontsSummaryDT.Rows.Add(NewRow);

                    FrontCost = 0;
                    FrontCount = 0;
                }
            }

            Table.Dispose();
            CurvedFrontsSummaryDT.DefaultView.Sort = "Front";
            CurvedFrontsSummaryBS.MoveFirst();
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

        public void GetFrontsInfo(ref decimal Square, ref decimal Cost, ref int Count)
        {
            for (int i = 0; i < FrontsSummaryDT.Rows.Count; i++)
            {
                Square += Convert.ToDecimal(FrontsSummaryDT.Rows[i]["Square"]);
                Count += Convert.ToInt32(FrontsSummaryDT.Rows[i]["Count"]);
                Cost += Convert.ToDecimal(FrontsSummaryDT.Rows[i]["Cost"]);
            }

            Cost = Decimal.Round(Cost, 2, MidpointRounding.AwayFromZero);
            Square = Decimal.Round(Square, 2, MidpointRounding.AwayFromZero);
        }

        public void GetCurvedFrontsInfo(ref decimal Cost, ref int CurvedCount)
        {
            for (int i = 0; i < CurvedFrontsSummaryDT.Rows.Count; i++)
            {
                CurvedCount += Convert.ToInt32(CurvedFrontsSummaryDT.Rows[i]["Count"]);
                Cost += Convert.ToDecimal(CurvedFrontsSummaryDT.Rows[i]["Cost"]);
            }

            Cost = Decimal.Round(Cost, 2, MidpointRounding.AwayFromZero);
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

        public void ClearCurvedFrontsOrders(int TypeOrders)
        {
            if (TypeOrders == 1)
            {
                MCurvedFSummaryDT.Clear();
            }
            if (TypeOrders == 2)
            {
                PrepareCurvedFSummaryDT.Clear();
            }
        }

        public void ClearFrontsOrders(int TypeOrders)
        {
            if (TypeOrders == 1)
            {
                MFSummaryDT.Clear();
            }
            if (TypeOrders == 2)
            {
                PrepareFSummaryDT.Clear();
            }
        }

        public void ClearDecorOrders(int TypeOrders)
        {
            if (TypeOrders == 1)
            {
                MDSummaryDT.Clear();
            }
            if (TypeOrders == 2)
            {
                PrepareDSummaryDT.Clear();
            }
        }

        private decimal MarketReadyCurvedFrontsCost(bool IsMarket, int MegaOrderID, int FactoryID, ref decimal ReadyCost, ref decimal AllCost)
        {
            decimal Percentage = 0;
            decimal ReadyFrontsCost = 0;
            decimal AllFrontsCost = 0;

            if (IsMarket)
            {
                DataRow[] RFRows = MReadyCurvedFrontsCostDT.Select("MegaOrderID = " + MegaOrderID + " AND FactoryID = " + FactoryID);
                foreach (DataRow Row in RFRows)
                    ReadyFrontsCost += Convert.ToDecimal(Row["FrontsCost"]);

                DataRow[] AFRows = MAllCurvedFrontsCostDT.Select("MegaOrderID = " + MegaOrderID + " AND FactoryID = " + FactoryID);
                foreach (DataRow Row in AFRows)
                    AllFrontsCost += Convert.ToDecimal(Row["FrontsCost"]);
            }
            else
            {
                DataRow[] RFRows = ZReadyCurvedFrontsCostDT.Select("MegaOrderID = " + MegaOrderID + " AND FactoryID = " + FactoryID);
                foreach (DataRow Row in RFRows)
                    ReadyFrontsCost += Convert.ToDecimal(Row["FrontsCost"]);

                DataRow[] AFRows = ZAllCurvedFrontsCostDT.Select("MegaOrderID = " + MegaOrderID + " AND FactoryID = " + FactoryID);
                foreach (DataRow Row in AFRows)
                    AllFrontsCost += Convert.ToDecimal(Row["FrontsCost"]);
            }

            ReadyCost = ReadyFrontsCost;
            AllCost = AllFrontsCost;

            if (AllFrontsCost > 0)
                Percentage = ReadyFrontsCost / AllFrontsCost;
            Percentage = Decimal.Round(Percentage * 100, 2, MidpointRounding.AwayFromZero);

            return Percentage;
        }

        private decimal MarketReadyFrontsCost(bool IsMarket, int MegaOrderID, int FactoryID, ref decimal ReadyCost, ref decimal AllCost)
        {
            decimal Percentage = 0;
            decimal ReadyFrontsCost = 0;
            decimal AllFrontsCost = 0;

            if (IsMarket)
            {
                DataRow[] RFRows = MReadyFrontsCostDT.Select("MegaOrderID = " + MegaOrderID + " AND FactoryID = " + FactoryID);
                foreach (DataRow Row in RFRows)
                    ReadyFrontsCost += Convert.ToDecimal(Row["FrontsCost"]);

                DataRow[] AFRows = MAllFrontsCostDT.Select("MegaOrderID = " + MegaOrderID + " AND FactoryID = " + FactoryID);
                foreach (DataRow Row in AFRows)
                    AllFrontsCost += Convert.ToDecimal(Row["FrontsCost"]);
            }
            else
            {
                DataRow[] RFRows = ZReadyCurvedFrontsCostDT.Select("MegaOrderID = " + MegaOrderID + " AND FactoryID = " + FactoryID);
                foreach (DataRow Row in RFRows)
                    ReadyFrontsCost += Convert.ToDecimal(Row["FrontsCost"]);

                DataRow[] AFRows = ZAllCurvedFrontsCostDT.Select("MegaOrderID = " + MegaOrderID + " AND FactoryID = " + FactoryID);
                foreach (DataRow Row in AFRows)
                    AllFrontsCost += Convert.ToDecimal(Row["FrontsCost"]);
            }

            ReadyCost = ReadyFrontsCost;
            AllCost = AllFrontsCost;

            if (AllFrontsCost > 0)
                Percentage = ReadyFrontsCost / AllFrontsCost;
            Percentage = Decimal.Round(Percentage * 100, 2, MidpointRounding.AwayFromZero);

            return Percentage;
        }

        private decimal MarketReadyDecorCost(bool IsMarket, int MegaOrderID, int FactoryID, ref decimal ReadyCost, ref decimal AllCost)
        {
            decimal Percentage = 0;
            decimal ReadyDecorCost = 0;
            decimal AllDecorCost = 0;

            if (IsMarket)
            {
                DataRow[] RDRows = MReadyDecorCostDT.Select("MegaOrderID = " + MegaOrderID + " AND FactoryID = " + FactoryID);
                foreach (DataRow Row in RDRows)
                    ReadyDecorCost += Convert.ToDecimal(Row["DecorCost"]);

                DataRow[] ADRows = MAllDecorCostDT.Select("MegaOrderID = " + MegaOrderID + " AND FactoryID = " + FactoryID);
                foreach (DataRow Row in ADRows)
                    AllDecorCost += Convert.ToDecimal(Row["DecorCost"]);
            }
            else
            {
                DataRow[] RDRows = ZReadyDecorCostDT.Select("MegaOrderID = " + MegaOrderID + " AND FactoryID = " + FactoryID);
                foreach (DataRow Row in RDRows)
                    ReadyDecorCost += Convert.ToDecimal(Row["DecorCost"]);

                DataRow[] ADRows = ZAllDecorCostDT.Select("MegaOrderID = " + MegaOrderID + " AND FactoryID = " + FactoryID);
                foreach (DataRow Row in ADRows)
                    AllDecorCost += Convert.ToDecimal(Row["DecorCost"]);
            }

            ReadyCost = ReadyDecorCost;
            AllCost = AllDecorCost;

            if (AllDecorCost > 0)
                Percentage = ReadyDecorCost / AllDecorCost;
            Percentage = Decimal.Round(Percentage * 100, 2, MidpointRounding.AwayFromZero);

            return Percentage;
        }

        private decimal PrepareReadyCurvedFrontsCost(string DocDateTime, int FactoryID, ref decimal ReadyCost, ref decimal AllCost)
        {
            decimal Percentage = 0;
            decimal ReadyFrontsCost = 0;
            decimal AllFrontsCost = 0;

            DataRow[] RFRows = PrepareReadyCurvedFrontsCostDT.Select("DocDateTime = '" + DocDateTime + "' AND FactoryID = " + FactoryID);
            foreach (DataRow Row in RFRows)
                ReadyFrontsCost += Convert.ToDecimal(Row["FrontsCost"]);

            DataRow[] AFRows = PrepareAllCurvedFrontsCostDT.Select("DocDateTime = '" + DocDateTime + "' AND FactoryID = " + FactoryID);
            foreach (DataRow Row in AFRows)
                AllFrontsCost += Convert.ToDecimal(Row["FrontsCost"]);

            ReadyCost = ReadyFrontsCost;
            AllCost = AllFrontsCost;

            if (AllFrontsCost > 0)
                Percentage = ReadyFrontsCost / AllFrontsCost;
            Percentage = Decimal.Round(Percentage * 100, 2, MidpointRounding.AwayFromZero);

            return Percentage;
        }

        private decimal PrepareReadyFrontsCost(string DocDateTime, int FactoryID, ref decimal ReadyCost, ref decimal AllCost)
        {
            decimal Percentage = 0;
            decimal ReadyFrontsCost = 0;
            decimal AllFrontsCost = 0;

            DataRow[] RFRows = PrepareReadyFrontsCostDT.Select("DocDateTime = '" + DocDateTime + "' AND FactoryID = " + FactoryID);
            foreach (DataRow Row in RFRows)
                ReadyFrontsCost += Convert.ToDecimal(Row["FrontsCost"]);

            DataRow[] AFRows = PrepareAllFrontsCostDT.Select("DocDateTime = '" + DocDateTime + "' AND FactoryID = " + FactoryID);
            foreach (DataRow Row in AFRows)
                AllFrontsCost += Convert.ToDecimal(Row["FrontsCost"]);

            ReadyCost = ReadyFrontsCost;
            AllCost = AllFrontsCost;

            if (AllFrontsCost > 0)
                Percentage = ReadyFrontsCost / AllFrontsCost;
            Percentage = Decimal.Round(Percentage * 100, 2, MidpointRounding.AwayFromZero);

            return Percentage;
        }

        private decimal PrepareReadyDecorCost(string DocDateTime, int FactoryID, ref decimal ReadyCost, ref decimal AllCost)
        {
            decimal Percentage = 0;
            decimal ReadyDecorCost = 0;
            decimal AllDecorCost = 0;

            DataRow[] RDRows = PrepareReadyDecorCostDT.Select("DocDateTime = '" + DocDateTime + "' AND FactoryID = " + FactoryID);
            foreach (DataRow Row in RDRows)
                ReadyDecorCost += Convert.ToDecimal(Row["DecorCost"]);

            DataRow[] ADRows = PrepareAllDecorCostDT.Select("DocDateTime = '" + DocDateTime + "' AND FactoryID = " + FactoryID);
            foreach (DataRow Row in ADRows)
                AllDecorCost += Convert.ToDecimal(Row["DecorCost"]);

            ReadyCost = ReadyDecorCost;
            AllCost = AllDecorCost;

            if (AllDecorCost > 0)
                Percentage = ReadyDecorCost / AllDecorCost;
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
    }
}