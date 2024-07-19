using System;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Windows.Forms;

namespace Infinium.Modules.StatisticsMarketing
{
    public class MarketingBatchStatistics
    {
        private DataTable FrontsDT = null;
        private DataTable DecorProductsDT = null;
        private DataTable DecorConfigDT = null;

        private DataTable MegaBatchDT = null;
        private DataTable FrontsOrdersDT = null;
        private DataTable DecorOrdersDT = null;
        private DataTable SimpleFrontsSummaryDT = null;
        private DataTable CurvedFrontsSummaryDT = null;
        private DataTable DecorProductsSummaryDT = null;

        private DataTable MarkProfilReadyFrontsDT = null;
        private DataTable MarkProfilReadyDecorDT = null;
        private DataTable MarkTPSReadyFrontsDT = null;
        private DataTable MarkTPSReadyDecorDT = null;
        private DataTable MarkAllFrontsDT = null;
        private DataTable MarkAllDecorDT = null;
        private DataTable MarkProfilOnProdFrontsDT = null;
        private DataTable MarkTPSOnProdFrontsDT = null;
        private DataTable MarkProfilOnProdDecorDT = null;
        private DataTable MarkTPSOnProdDecorDT = null;

        public BindingSource MegaBatchBS = null;
        public BindingSource SimpleFrontsSummaryBS = null;
        public BindingSource CurvedFrontsSummaryBS = null;
        public BindingSource DecorProductsSummaryBS = null;

        public MarketingBatchStatistics()
        {
            Create();
            Fill();
            Binding();
        }

        private void Create()
        {
            FrontsDT = new DataTable();
            DecorProductsDT = new DataTable();
            DecorConfigDT = new DataTable();
            MegaBatchDT = new DataTable();
            FrontsOrdersDT = new DataTable();
            DecorOrdersDT = new DataTable();

            MarkProfilReadyFrontsDT = new DataTable();
            MarkProfilReadyDecorDT = new DataTable();
            MarkTPSReadyFrontsDT = new DataTable();
            MarkTPSReadyDecorDT = new DataTable();

            MarkAllFrontsDT = new DataTable();
            MarkAllDecorDT = new DataTable();

            MarkProfilOnProdFrontsDT = new DataTable();
            MarkProfilOnProdDecorDT = new DataTable();
            MarkTPSOnProdFrontsDT = new DataTable();
            MarkTPSOnProdDecorDT = new DataTable();

            MegaBatchBS = new BindingSource();
            SimpleFrontsSummaryBS = new BindingSource();
            CurvedFrontsSummaryBS = new BindingSource();
            DecorProductsSummaryBS = new BindingSource();

            SimpleFrontsSummaryDT = new DataTable();
            SimpleFrontsSummaryDT.Columns.Add(new DataColumn("Front", System.Type.GetType("System.String")));
            SimpleFrontsSummaryDT.Columns.Add(new DataColumn("FrontID", System.Type.GetType("System.Int32")));
            SimpleFrontsSummaryDT.Columns.Add(new DataColumn("Width", System.Type.GetType("System.Int32")));
            SimpleFrontsSummaryDT.Columns.Add(new DataColumn("AllSquare", System.Type.GetType("System.String")));
            SimpleFrontsSummaryDT.Columns.Add(new DataColumn("AllCount", System.Type.GetType("System.Int32")));
            SimpleFrontsSummaryDT.Columns.Add(new DataColumn("OnProdSquare", System.Type.GetType("System.String")));
            SimpleFrontsSummaryDT.Columns.Add(new DataColumn("OnProdCount", System.Type.GetType("System.Int32")));
            SimpleFrontsSummaryDT.Columns.Add(new DataColumn("InProdSquare", System.Type.GetType("System.String")));
            SimpleFrontsSummaryDT.Columns.Add(new DataColumn("InProdCount", System.Type.GetType("System.Int32")));
            SimpleFrontsSummaryDT.Columns.Add(new DataColumn("ReadySquare", System.Type.GetType("System.String")));
            SimpleFrontsSummaryDT.Columns.Add(new DataColumn("ReadyCount", System.Type.GetType("System.Int32")));
            SimpleFrontsSummaryDT.Columns.Add(new DataColumn("Ready", System.Type.GetType("System.Boolean")));

            CurvedFrontsSummaryDT = new DataTable();
            CurvedFrontsSummaryDT.Columns.Add(new DataColumn("Front", System.Type.GetType("System.String")));
            CurvedFrontsSummaryDT.Columns.Add(new DataColumn("FrontID", System.Type.GetType("System.Int32")));
            CurvedFrontsSummaryDT.Columns.Add(new DataColumn("Width", System.Type.GetType("System.Int32")));
            CurvedFrontsSummaryDT.Columns.Add(new DataColumn("AllCount", System.Type.GetType("System.Int32")));
            CurvedFrontsSummaryDT.Columns.Add(new DataColumn("OnProdCount", System.Type.GetType("System.Int32")));
            CurvedFrontsSummaryDT.Columns.Add(new DataColumn("InProdCount", System.Type.GetType("System.Int32")));
            CurvedFrontsSummaryDT.Columns.Add(new DataColumn("ReadyCount", System.Type.GetType("System.Int32")));
            CurvedFrontsSummaryDT.Columns.Add(new DataColumn("Ready", System.Type.GetType("System.Boolean")));

            DecorProductsSummaryDT = new DataTable();
            DecorProductsSummaryDT.Columns.Add(new DataColumn("DecorProduct", System.Type.GetType("System.String")));
            DecorProductsSummaryDT.Columns.Add(new DataColumn("ProductID", System.Type.GetType("System.Int32")));
            DecorProductsSummaryDT.Columns.Add(new DataColumn("MeasureID", System.Type.GetType("System.Int32")));
            DecorProductsSummaryDT.Columns.Add(new DataColumn("AllCount", System.Type.GetType("System.Decimal")));
            DecorProductsSummaryDT.Columns.Add(new DataColumn("OnProdCount", System.Type.GetType("System.Decimal")));
            DecorProductsSummaryDT.Columns.Add(new DataColumn("InProdCount", System.Type.GetType("System.Decimal")));
            DecorProductsSummaryDT.Columns.Add(new DataColumn("ReadyCount", System.Type.GetType("System.Decimal")));
            DecorProductsSummaryDT.Columns.Add(new DataColumn("Ready", System.Type.GetType("System.Boolean")));
            DecorProductsSummaryDT.Columns.Add(new DataColumn("Measure", System.Type.GetType("System.String")));
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

            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT ProductID, ProductName, MeasureID, ReportParam FROM DecorProducts", ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(DecorProductsDT);
            }

            //using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM DecorConfig", ConnectionStrings.CatalogConnectionString))
            //{
            //    DA.Fill(DecorConfigDT);
            //}
            DecorConfigDT = TablesManager.DecorConfigDataTableAll;

            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT TOP 0 * FROM MegaBatch",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(MegaBatchDT);
            }
            MegaBatchDT.Columns.Add(new DataColumn("Firm", Type.GetType("System.String")));
            MegaBatchDT.Columns.Add(new DataColumn("GroupType", Type.GetType("System.Int32")));
            MegaBatchDT.Columns.Add(new DataColumn("ProfilOnProductionPerc", Type.GetType("System.Int32")));
            MegaBatchDT.Columns.Add(new DataColumn("ProfilInProductionPerc", Type.GetType("System.Int32")));
            MegaBatchDT.Columns.Add(new DataColumn("ProfilReadyPerc", Type.GetType("System.Int32")));
            MegaBatchDT.Columns.Add(new DataColumn("ProfilReady", Type.GetType("System.Boolean")));
            MegaBatchDT.Columns.Add(new DataColumn("ProfilPackingDate", Type.GetType("System.String")));
            MegaBatchDT.Columns.Add(new DataColumn("TPSOnProductionPerc", Type.GetType("System.Int32")));
            MegaBatchDT.Columns.Add(new DataColumn("TPSInProductionPerc", Type.GetType("System.Int32")));
            MegaBatchDT.Columns.Add(new DataColumn("FilenkaPerc", Type.GetType("System.Int32")));
            MegaBatchDT.Columns.Add(new DataColumn("TrimmingPerc", Type.GetType("System.Int32")));
            MegaBatchDT.Columns.Add(new DataColumn("AssemblyPerc", Type.GetType("System.Int32")));
            MegaBatchDT.Columns.Add(new DataColumn("DeyingPerc", Type.GetType("System.Int32")));
            MegaBatchDT.Columns.Add(new DataColumn("TPSReadyPerc", Type.GetType("System.Int32")));
            MegaBatchDT.Columns.Add(new DataColumn("TPSReady", Type.GetType("System.Boolean")));
            MegaBatchDT.Columns.Add(new DataColumn("TPSPackingDate", Type.GetType("System.String")));

            SelectCommand = "SELECT TOP 0 Batch.MegaBatchID, FrontsOrders.FrontID, FrontsOrders.Width, " +
                " SUM(FrontsOrders.Square) AS AllSquare, SUM(FrontsOrders.Count) AS AllCount" +
                " FROM FrontsOrders" +
                " INNER JOIN BatchDetails ON FrontsOrders.MainOrderID = BatchDetails.MainOrderID" +
                " INNER JOIN Batch ON BatchDetails.BatchID = Batch.BatchID" +
                " INNER JOIN MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID" +
                " WHERE (FrontsOrders.FactoryID = 1) AND MainOrders.ProfilProductionStatusID = 3" +
                " GROUP BY Batch.MegaBatchID, FrontsOrders.FrontID, FrontsOrders.Width" +
                " ORDER BY Batch.MegaBatchID, FrontsOrders.FrontID, FrontsOrders.Width";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(MarkProfilOnProdFrontsDT);
                MarkProfilOnProdFrontsDT.Columns.Add(new DataColumn(("GroupType"), System.Type.GetType("System.Int32")));
            }

            SelectCommand = "SELECT TOP 0 Batch.MegaBatchID, DecorOrders.ProductID," +
                " SUM(DecorOrders.Count) AS AllCount" +
                " FROM DecorOrders" +
                " INNER JOIN BatchDetails ON DecorOrders.MainOrderID = BatchDetails.MainOrderID" +
                " INNER JOIN Batch ON BatchDetails.BatchID = Batch.BatchID" +
                " INNER JOIN MainOrders ON DecorOrders.MainOrderID = MainOrders.MainOrderID" +
                " WHERE (DecorOrders.FactoryID = 1) AND MainOrders.ProfilProductionStatusID = 3" +
                " GROUP BY Batch.MegaBatchID, DecorOrders.ProductID" +
                " ORDER BY Batch.MegaBatchID, DecorOrders.ProductID";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(MarkProfilOnProdDecorDT);
                MarkProfilOnProdDecorDT.Columns.Add(new DataColumn(("GroupType"), System.Type.GetType("System.Int32")));
            }

            SelectCommand = "SELECT TOP 0 Batch.MegaBatchID, FrontsOrders.FrontID, FrontsOrders.Width," +
                " SUM(FrontsOrders.Square) AS AllSquare, SUM(FrontsOrders.Count) AS AllCount" +
                " FROM FrontsOrders" +
                " INNER JOIN BatchDetails ON FrontsOrders.MainOrderID = BatchDetails.MainOrderID" +
                " INNER JOIN Batch ON BatchDetails.BatchID = Batch.BatchID" +
                " INNER JOIN MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID" +
                " WHERE (FrontsOrders.FactoryID = 2) AND MainOrders.TPSProductionStatusID = 3" +
                " GROUP BY Batch.MegaBatchID, FrontsOrders.FrontID, FrontsOrders.Width" +
                " ORDER BY Batch.MegaBatchID, FrontsOrders.FrontID, FrontsOrders.Width";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(MarkTPSOnProdFrontsDT);
                MarkTPSOnProdFrontsDT.Columns.Add(new DataColumn(("GroupType"), System.Type.GetType("System.Int32")));
            }

            SelectCommand = "SELECT TOP 0 Batch.MegaBatchID, DecorOrders.ProductID," +
                " SUM(DecorOrders.Count) AS AllCount" +
                " FROM DecorOrders" +
                " INNER JOIN BatchDetails ON DecorOrders.MainOrderID = BatchDetails.MainOrderID" +
                " INNER JOIN Batch ON BatchDetails.BatchID = Batch.BatchID" +
                " INNER JOIN MainOrders ON DecorOrders.MainOrderID = MainOrders.MainOrderID" +
                " WHERE (DecorOrders.FactoryID = 2) AND MainOrders.TPSProductionStatusID = 3" +
                " GROUP BY Batch.MegaBatchID, DecorOrders.ProductID" +
                " ORDER BY Batch.MegaBatchID, DecorOrders.ProductID";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(MarkTPSOnProdDecorDT);
                MarkTPSOnProdDecorDT.Columns.Add(new DataColumn(("GroupType"), System.Type.GetType("System.Int32")));
            }

            SelectCommand = @"SELECT TOP 0 Batch.MegaBatchID, FrontsOrders.FrontID, FrontsOrders.Width, Packages.PackageStatusID, SUM(PackageDetails.Count) AS AllCount,
                      SUM(FrontsOrders.Square / FrontsOrders.Count * PackageDetails.Count) AS AllSquare
            FROM Packages INNER JOIN
                      PackageDetails ON PackageDetails.PackageID = Packages.PackageID INNER JOIN
                      FrontsOrders ON PackageDetails.OrderID = FrontsOrders.FrontsOrdersID INNER JOIN
                      BatchDetails ON Packages.MainOrderID = BatchDetails.MainOrderID AND BatchDetails.FactoryID = 1 INNER JOIN
                      Batch ON BatchDetails.BatchID = Batch.BatchID INNER JOIN
                      MegaBatch ON Batch.MegaBatchID = MegaBatch.MegaBatchID
            WHERE (Packages.ProductType = 0) AND (Packages.FactoryID = 1) AND PackageStatusID > 0
            GROUP BY Batch.MegaBatchID, FrontsOrders.FrontID, FrontsOrders.Width, Packages.PackageStatusID
            ORDER BY Batch.MegaBatchID, FrontsOrders.FrontID, FrontsOrders.Width, Packages.PackageStatusID";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(MarkProfilReadyFrontsDT);
                MarkProfilReadyFrontsDT.Columns.Add(new DataColumn(("GroupType"), System.Type.GetType("System.Int32")));
            }

            SelectCommand = @"SELECT TOP 0 Batch.MegaBatchID, DecorOrders.ProductID, Packages.PackageStatusID, SUM(PackageDetails.Count) AS AllCount
            FROM Packages INNER JOIN
                      PackageDetails ON PackageDetails.PackageID = Packages.PackageID INNER JOIN
                      DecorOrders ON PackageDetails.OrderID = DecorOrders.DecorOrderID INNER JOIN
                      BatchDetails ON Packages.MainOrderID = BatchDetails.MainOrderID AND BatchDetails.FactoryID = 1 INNER JOIN
                      Batch ON BatchDetails.BatchID = Batch.BatchID INNER JOIN
                      MegaBatch ON Batch.MegaBatchID = MegaBatch.MegaBatchID
            WHERE (Packages.ProductType = 1) AND (Packages.FactoryID = 1) AND PackageStatusID > 0
            GROUP BY Batch.MegaBatchID, DecorOrders.ProductID, Packages.PackageStatusID
            ORDER BY Batch.MegaBatchID, DecorOrders.ProductID, Packages.PackageStatusID";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(MarkProfilReadyDecorDT);
                MarkProfilReadyDecorDT.Columns.Add(new DataColumn(("GroupType"), System.Type.GetType("System.Int32")));
            }

            SelectCommand = @"SELECT TOP 0 Batch.MegaBatchID, FrontsOrders.FrontID, FrontsOrders.Width, Packages.PackageStatusID, SUM(PackageDetails.Count) AS AllCount,
                      SUM(FrontsOrders.Square / FrontsOrders.Count * PackageDetails.Count) AS AllSquare
            FROM Packages INNER JOIN
                      PackageDetails ON PackageDetails.PackageID = Packages.PackageID INNER JOIN
                      FrontsOrders ON PackageDetails.OrderID = FrontsOrders.FrontsOrdersID INNER JOIN
                      BatchDetails ON Packages.MainOrderID = BatchDetails.MainOrderID AND BatchDetails.FactoryID = 2 INNER JOIN
                      Batch ON BatchDetails.BatchID = Batch.BatchID INNER JOIN
                      MegaBatch ON Batch.MegaBatchID = MegaBatch.MegaBatchID
            WHERE (Packages.ProductType = 0) AND (Packages.FactoryID = 2) AND PackageStatusID > 0
            GROUP BY Batch.MegaBatchID, FrontsOrders.FrontID, FrontsOrders.Width, Packages.PackageStatusID
            ORDER BY Batch.MegaBatchID, FrontsOrders.FrontID, FrontsOrders.Width, Packages.PackageStatusID";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(MarkTPSReadyFrontsDT);
                MarkTPSReadyFrontsDT.Columns.Add(new DataColumn(("GroupType"), System.Type.GetType("System.Int32")));
            }

            SelectCommand = @"SELECT TOP 0 Batch.MegaBatchID, DecorOrders.ProductID, Packages.PackageStatusID, SUM(PackageDetails.Count) AS AllCount
            FROM Packages INNER JOIN
                      PackageDetails ON PackageDetails.PackageID = Packages.PackageID INNER JOIN
                      DecorOrders ON PackageDetails.OrderID = DecorOrders.DecorOrderID INNER JOIN
                      BatchDetails ON Packages.MainOrderID = BatchDetails.MainOrderID AND BatchDetails.FactoryID = 2 INNER JOIN
                      Batch ON BatchDetails.BatchID = Batch.BatchID INNER JOIN
                      MegaBatch ON Batch.MegaBatchID = MegaBatch.MegaBatchID
            WHERE (Packages.ProductType = 1) AND (Packages.FactoryID = 2) AND PackageStatusID > 0
            GROUP BY Batch.MegaBatchID, DecorOrders.ProductID, Packages.PackageStatusID
            ORDER BY Batch.MegaBatchID, DecorOrders.ProductID, Packages.PackageStatusID";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(MarkTPSReadyDecorDT);
                MarkTPSReadyDecorDT.Columns.Add(new DataColumn(("GroupType"), System.Type.GetType("System.Int32")));
            }

            SelectCommand = @"SELECT TOP 0 Batch.MegaBatchID, FrontsOrders.FrontID, FrontsOrders.Width, FrontsOrders.FactoryID AS FFactoryID, BatchDetails.FactoryID AS BFactoryID, SUM(FrontsOrders.Count) AS AllCount, SUM(FrontsOrders.Square) AS AllSquare
            FROM FrontsOrders INNER JOIN
                      BatchDetails ON FrontsOrders.MainOrderID = BatchDetails.MainOrderID INNER JOIN
                      Batch ON BatchDetails.BatchID = Batch.BatchID INNER JOIN
                      MegaBatch ON Batch.MegaBatchID = MegaBatch.MegaBatchID
            GROUP BY Batch.MegaBatchID, FrontsOrders.FrontID, FrontsOrders.Width, FrontsOrders.FactoryID, BatchDetails.FactoryID
            ORDER BY Batch.MegaBatchID, FrontsOrders.FrontID, FrontsOrders.Width, FrontsOrders.FactoryID, BatchDetails.FactoryID";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(MarkAllFrontsDT);
                MarkAllFrontsDT.Columns.Add(new DataColumn(("GroupType"), System.Type.GetType("System.Int32")));
            }

            SelectCommand = @"SELECT TOP 0 Batch.MegaBatchID, DecorOrders.ProductID, DecorOrders.FactoryID AS FFactoryID, BatchDetails.FactoryID AS BFactoryID, SUM(DecorOrders.Count) AS AllCount
            FROM DecorOrders INNER JOIN
                      BatchDetails ON DecorOrders.MainOrderID = BatchDetails.MainOrderID INNER JOIN
                      Batch ON BatchDetails.BatchID = Batch.BatchID INNER JOIN
                      MegaBatch ON Batch.MegaBatchID = MegaBatch.MegaBatchID
            GROUP BY Batch.MegaBatchID, DecorOrders.ProductID, DecorOrders.FactoryID, BatchDetails.FactoryID
            ORDER BY Batch.MegaBatchID, DecorOrders.ProductID, DecorOrders.FactoryID, BatchDetails.FactoryID";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(MarkAllDecorDT);
                MarkAllDecorDT.Columns.Add(new DataColumn(("GroupType"), System.Type.GetType("System.Int32")));
            }

            SelectCommand = "SELECT TOP 0 * FROM FrontsOrders" +
                " WHERE MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch))";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(FrontsOrdersDT);
            }
        }

        private void Binding()
        {
            MegaBatchBS.DataSource = MegaBatchDT;
            SimpleFrontsSummaryBS.DataSource = SimpleFrontsSummaryDT;
            CurvedFrontsSummaryBS.DataSource = CurvedFrontsSummaryDT;
            DecorProductsSummaryBS.DataSource = DecorProductsSummaryDT;
        }

        public void FilterBatch(bool Marketing, bool ZOV, bool Profil, bool TPS, bool DoNotShowReady, bool SimpleFronts, bool CurvedFronts, bool Decor)
        {
            string BatchFilter = string.Empty;
            string FactoryFilter = string.Empty;
            string Filter = string.Empty;
            int FactoryID = 0;
            if (Profil && !TPS)
                FactoryFilter = " AND FFactoryID = 1";
            if (!Profil && TPS)
                FactoryFilter = " AND FFactoryID = 2";
            if (Marketing && ZOV)
            {
                if (SimpleFronts)
                {
                    foreach (DataRow item in MegaBatchDT.Rows)
                    {
                        DataRow[] rows = MarkAllFrontsDT.Select("Width<>-1 AND MegaBatchID=" + Convert.ToInt32(item["MegaBatchID"]) + FactoryFilter);
                        if (rows.Count() == 0)
                            BatchFilter += item["MegaBatchID"].ToString() + ",";
                    }
                }
                if (CurvedFronts)
                {
                    foreach (DataRow item in MegaBatchDT.Rows)
                    {
                        DataRow[] rows = MarkAllFrontsDT.Select("Width=-1 AND MegaBatchID=" + Convert.ToInt32(item["MegaBatchID"]) + FactoryFilter);
                        if (rows.Count() == 0)
                            BatchFilter += item["MegaBatchID"].ToString() + ",";
                    }
                }
                if (Decor)
                {
                    foreach (DataRow item in MegaBatchDT.Rows)
                    {
                        DataRow[] rows = MarkAllDecorDT.Select("MegaBatchID=" + Convert.ToInt32(item["MegaBatchID"]) + FactoryFilter);
                        if (rows.Count() == 0)
                            BatchFilter += item["MegaBatchID"].ToString() + ",";
                    }
                }
            }
            if (Marketing && !ZOV)
            {
                Filter = "GroupType = 1";
                if (SimpleFronts)
                {
                    foreach (DataRow item in MegaBatchDT.Rows)
                    {
                        DataRow[] rows = MarkAllFrontsDT.Select("GroupType = 1 AND Width<>-1 AND MegaBatchID=" + Convert.ToInt32(item["MegaBatchID"]) + FactoryFilter);
                        if (rows.Count() == 0)
                            BatchFilter += item["MegaBatchID"].ToString() + ",";
                    }
                }
                if (CurvedFronts)
                {
                    foreach (DataRow item in MegaBatchDT.Rows)
                    {
                        DataRow[] rows = MarkAllFrontsDT.Select("GroupType = 1 AND Width=-1 AND MegaBatchID=" + Convert.ToInt32(item["MegaBatchID"]) + FactoryFilter);
                        if (rows.Count() == 0)
                            BatchFilter += item["MegaBatchID"].ToString() + ",";
                    }
                }
                if (Decor)
                {
                    foreach (DataRow item in MegaBatchDT.Rows)
                    {
                        DataRow[] rows = MarkAllDecorDT.Select("GroupType = 1 AND MegaBatchID=" + Convert.ToInt32(item["MegaBatchID"]) + FactoryFilter);
                        if (rows.Count() == 0)
                            BatchFilter += item["MegaBatchID"].ToString() + ",";
                    }
                }
            }
            if (!Marketing && ZOV)
            {
                Filter = "GroupType = 0";
                if (SimpleFronts)
                {
                    foreach (DataRow item in MegaBatchDT.Rows)
                    {
                        DataRow[] rows = MarkAllFrontsDT.Select("GroupType = 0 AND Width<>-1 AND MegaBatchID=" + Convert.ToInt32(item["MegaBatchID"]) + FactoryFilter);
                        if (rows.Count() == 0)
                            BatchFilter += item["MegaBatchID"].ToString() + ",";
                    }
                }
                if (CurvedFronts)
                {
                    foreach (DataRow item in MegaBatchDT.Rows)
                    {
                        DataRow[] rows = MarkAllFrontsDT.Select("GroupType = 0 AND Width=-1 AND MegaBatchID=" + Convert.ToInt32(item["MegaBatchID"]) + FactoryFilter);
                        if (rows.Count() == 0)
                            BatchFilter += item["MegaBatchID"].ToString() + ",";
                    }
                }
                if (Decor)
                {
                    foreach (DataRow item in MegaBatchDT.Rows)
                    {
                        DataRow[] rows = MarkAllDecorDT.Select("GroupType = 0 AND MegaBatchID=" + Convert.ToInt32(item["MegaBatchID"]) + FactoryFilter);
                        if (rows.Count() == 0)
                            BatchFilter += item["MegaBatchID"].ToString() + ",";
                    }
                }
            }
            if (!Marketing && !ZOV)
                Filter = "GroupType = -1";

            if (DoNotShowReady)
            {
                if (Profil && !TPS)
                    FactoryID = 1;
                if (!Profil && TPS)
                    FactoryID = 2;
                if (!Profil && !TPS)
                    FactoryID = -1;

                if (Filter.Length > 0)
                {
                    if (FactoryID == 0)
                        Filter += " AND (ProfilReady = false OR TPSReady = false)";
                    if (FactoryID == 1)
                        Filter += " AND (ProfilReady = false)";
                    if (FactoryID == 2)
                        Filter += " AND (TPSReady = false)";
                    if (FactoryID == -1)
                        Filter += " AND (MegaBatchID = 0)";
                }
                else
                {
                    if (FactoryID == 0)
                        Filter = "(ProfilReady = false OR TPSReady = false)";
                    if (FactoryID == 1)
                        Filter = "(ProfilReady = false)";
                    if (FactoryID == 2)
                        Filter = "(TPSReady = false)";
                    if (FactoryID == -1)
                        Filter += "(MegaBatchID = 0)";
                }
            }
            if (BatchFilter.Length > 0 && Filter.Length == 0)
                BatchFilter = "MegaBatchID NOT IN (" + BatchFilter.Substring(0, BatchFilter.Length - 1) + ")";
            if (BatchFilter.Length > 0 && Filter.Length > 0)
                BatchFilter = " AND MegaBatchID NOT IN (" + BatchFilter.Substring(0, BatchFilter.Length - 1) + ")";
            MegaBatchBS.Filter = Filter + BatchFilter;
            MegaBatchBS.MoveFirst();
        }

        public void UpdateMegaBatch(bool Marketing, bool ZOV, DateTime FirstDate, DateTime SecondDate)
        {
            string ConnectionString = ConnectionStrings.MarketingOrdersConnectionString;
            if (ZOV)
                ConnectionString = ConnectionStrings.ZOVOrdersConnectionString;
            string SelectCommand = "SELECT MegaBatchID, CreateDateTime, ProfilEntryDateTime, TPSEntryDateTime, Notes FROM MegaBatch" +
                " WHERE CAST(CreateDateTime AS Date) >= '" + FirstDate.ToString("yyyy-MM-dd") +
                "' AND CAST(CreateDateTime AS Date) <= '" + SecondDate.ToString("yyyy-MM-dd") + "' ORDER BY MegaBatchID DESC";
            MegaBatchDT.Clear();
            if (Marketing)
            {
                using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DA.Fill(MegaBatchDT);
                }
                for (int i = 0; i < MegaBatchDT.Rows.Count; i++)
                {
                    MegaBatchDT.Rows[i]["Firm"] = "Маркетинг";
                    MegaBatchDT.Rows[i]["GroupType"] = 1;
                }
            }
            if (ZOV)
            {
                DataTable DT = MegaBatchDT.Clone();
                using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.ZOVOrdersConnectionString))
                {
                    DA.Fill(DT);
                }
                for (int i = 0; i < DT.Rows.Count; i++)
                {
                    DT.Rows[i]["Firm"] = "ЗОВ";
                    DT.Rows[i]["GroupType"] = 0;
                }
                foreach (DataRow item in DT.Rows)
                    MegaBatchDT.ImportRow(item);
                DT.Dispose();
            }
        }

        public void UpdateFrontsOrders(bool Marketing, bool ZOV, DateTime FirstDate, DateTime SecondDate)
        {
            string SelectCommand = "SELECT Batch.MegaBatchID, FrontsOrders.FrontID, FrontsOrders.Width, " +
                " SUM(FrontsOrders.Square) AS AllSquare, SUM(FrontsOrders.Count) AS AllCount" +
                " FROM FrontsOrders" +
                " INNER JOIN BatchDetails ON FrontsOrders.MainOrderID = BatchDetails.MainOrderID" +
                " INNER JOIN Batch ON BatchDetails.BatchID = Batch.BatchID INNER JOIN" +
                     " MegaBatch ON Batch.MegaBatchID = MegaBatch.MegaBatchID AND CAST(MegaBatch.CreateDateTime AS Date) >= '" + FirstDate.ToString("yyyy-MM-dd") +
                "' AND CAST(MegaBatch.CreateDateTime AS Date) <= '" + SecondDate.ToString("yyyy-MM-dd") + @"'
                INNER JOIN MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID" +
                " WHERE (FrontsOrders.FactoryID = 1) AND MainOrders.ProfilProductionStatusID = 3" +
                " GROUP BY Batch.MegaBatchID, FrontsOrders.FrontID, FrontsOrders.Width" +
                " ORDER BY Batch.MegaBatchID, FrontsOrders.FrontID, FrontsOrders.Width";
            DataTable DT = MarkProfilOnProdFrontsDT.Clone();
            MarkProfilOnProdFrontsDT.Clear();
            if (Marketing)
            {
                using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DA.Fill(MarkProfilOnProdFrontsDT);
                }
                for (int i = 0; i < MarkProfilOnProdFrontsDT.Rows.Count; i++)
                    MarkProfilOnProdFrontsDT.Rows[i]["GroupType"] = 1;
            }
            if (ZOV)
            {
                using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.ZOVOrdersConnectionString))
                {
                    DA.Fill(DT);
                }
                for (int i = 0; i < DT.Rows.Count; i++)
                    DT.Rows[i]["GroupType"] = 0;
                foreach (DataRow item in DT.Rows)
                    MarkProfilOnProdFrontsDT.ImportRow(item);
            }

            SelectCommand = "SELECT Batch.MegaBatchID, FrontsOrders.FrontID, FrontsOrders.Width," +
                " SUM(FrontsOrders.Square) AS AllSquare, SUM(FrontsOrders.Count) AS AllCount" +
                " FROM FrontsOrders" +
                " INNER JOIN BatchDetails ON FrontsOrders.MainOrderID = BatchDetails.MainOrderID" +
                " INNER JOIN Batch ON BatchDetails.BatchID = Batch.BatchID INNER JOIN" +
                     " MegaBatch ON Batch.MegaBatchID = MegaBatch.MegaBatchID AND CAST(MegaBatch.CreateDateTime AS Date) >= '" + FirstDate.ToString("yyyy-MM-dd") +
                "' AND CAST(MegaBatch.CreateDateTime AS Date) <= '" + SecondDate.ToString("yyyy-MM-dd") + @"'
                 INNER JOIN MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID" +
                " WHERE (FrontsOrders.FactoryID = 2) AND MainOrders.TPSProductionStatusID = 3" +
                " GROUP BY Batch.MegaBatchID, FrontsOrders.FrontID, FrontsOrders.Width" +
                " ORDER BY Batch.MegaBatchID, FrontsOrders.FrontID, FrontsOrders.Width";
            MarkTPSOnProdFrontsDT.Clear();
            if (Marketing)
            {
                using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DA.Fill(MarkTPSOnProdFrontsDT);
                }
                for (int i = 0; i < MarkTPSOnProdFrontsDT.Rows.Count; i++)
                    MarkTPSOnProdFrontsDT.Rows[i]["GroupType"] = 1;
            }
            if (ZOV)
            {
                using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.ZOVOrdersConnectionString))
                {
                    DA.Fill(DT);
                }
                for (int i = 0; i < DT.Rows.Count; i++)
                    DT.Rows[i]["GroupType"] = 0;
                foreach (DataRow item in DT.Rows)
                    MarkTPSOnProdFrontsDT.ImportRow(item);
            }

            SelectCommand = @"SELECT Batch.MegaBatchID, FrontsOrders.FrontID, FrontsOrders.Width, Packages.PackageStatusID, SUM(PackageDetails.Count) AS AllCount,
                      SUM(FrontsOrders.Square / FrontsOrders.Count * PackageDetails.Count) AS AllSquare
            FROM Packages INNER JOIN
                      PackageDetails ON PackageDetails.PackageID = Packages.PackageID INNER JOIN
                      FrontsOrders ON PackageDetails.OrderID = FrontsOrders.FrontsOrdersID INNER JOIN
                      BatchDetails ON Packages.MainOrderID = BatchDetails.MainOrderID AND BatchDetails.FactoryID = 1 INNER JOIN
                      Batch ON BatchDetails.BatchID = Batch.BatchID INNER JOIN
                      MegaBatch ON Batch.MegaBatchID = MegaBatch.MegaBatchID AND CAST(MegaBatch.CreateDateTime AS Date) >= '" + FirstDate.ToString("yyyy-MM-dd") +
                "' AND CAST(MegaBatch.CreateDateTime AS Date) <= '" + SecondDate.ToString("yyyy-MM-dd") + @"'
            WHERE (Packages.ProductType = 0) AND (Packages.FactoryID = 1) AND PackageStatusID > 0
            GROUP BY Batch.MegaBatchID, FrontsOrders.FrontID, FrontsOrders.Width, Packages.PackageStatusID
            ORDER BY Batch.MegaBatchID, FrontsOrders.FrontID, FrontsOrders.Width, Packages.PackageStatusID";
            MarkProfilReadyFrontsDT.Clear();
            if (Marketing)
            {
                using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DA.Fill(MarkProfilReadyFrontsDT);
                }
                for (int i = 0; i < MarkProfilReadyFrontsDT.Rows.Count; i++)
                    MarkProfilReadyFrontsDT.Rows[i]["GroupType"] = 1;
            }
            {
                SelectCommand = "SELECT Batch.MegaBatchID, FrontsOrders.FrontID, FrontsOrders.Width, " +
                                " SUM(FrontsOrders.Square) AS AllSquare, SUM(FrontsOrders.Count) AS AllCount" +
                                " FROM FrontsOrders" +
                                " INNER JOIN BatchDetails ON FrontsOrders.MainOrderID = BatchDetails.MainOrderID" +
                                " INNER JOIN Batch ON BatchDetails.BatchID = Batch.BatchID INNER JOIN" +
                                " MegaBatch ON Batch.MegaBatchID = MegaBatch.MegaBatchID AND CAST(MegaBatch.CreateDateTime AS Date) >= '" + FirstDate.ToString("yyyy-MM-dd") +
                                "' AND CAST(MegaBatch.CreateDateTime AS Date) <= '" + SecondDate.ToString("yyyy-MM-dd") + @"'
                INNER JOIN MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID" +
                                " WHERE FrontsOrders.MainOrderID not in (select mainorderId from packages) and (FrontsOrders.FactoryID = 1)" +
                                " AND (MainOrders.ProfilProductionStatusID = 1 " +
                                " and MainOrders.ProfilStorageStatusID = 1 and MainOrders.ProfilExpeditionStatusID = 1" +
                                " and MainOrders.ProfilDispatchStatusID = 2)" +
                                " GROUP BY Batch.MegaBatchID, FrontsOrders.FrontID, FrontsOrders.Width" +
                                " ORDER BY Batch.MegaBatchID, FrontsOrders.FrontID, FrontsOrders.Width";
                using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DA.Fill(DT);
                }
                for (int i = 0; i < DT.Rows.Count; i++)
                    DT.Rows[i]["GroupType"] = 1;
                foreach (DataRow item in DT.Rows)
                    MarkProfilReadyFrontsDT.ImportRow(item);
            }

            SelectCommand = @"SELECT Batch.MegaBatchID, FrontsOrders.FrontID, FrontsOrders.Width, Packages.PackageStatusID, SUM(PackageDetails.Count) AS AllCount,
                      SUM(FrontsOrders.Square / FrontsOrders.Count * PackageDetails.Count) AS AllSquare
            FROM Packages INNER JOIN
                      PackageDetails ON PackageDetails.PackageID = Packages.PackageID INNER JOIN
                      FrontsOrders ON PackageDetails.OrderID = FrontsOrders.FrontsOrdersID INNER JOIN
                      BatchDetails ON Packages.MainOrderID = BatchDetails.MainOrderID AND BatchDetails.FactoryID = 2 INNER JOIN
                      Batch ON BatchDetails.BatchID = Batch.BatchID INNER JOIN
                      MegaBatch ON Batch.MegaBatchID = MegaBatch.MegaBatchID AND CAST(MegaBatch.CreateDateTime AS Date) >= '" + FirstDate.ToString("yyyy-MM-dd") +
                "' AND CAST(MegaBatch.CreateDateTime AS Date) <= '" + SecondDate.ToString("yyyy-MM-dd") + @"'
            WHERE (Packages.ProductType = 0) AND (Packages.FactoryID = 2) AND PackageStatusID > 0
            GROUP BY Batch.MegaBatchID, FrontsOrders.FrontID, FrontsOrders.Width, Packages.PackageStatusID
            ORDER BY Batch.MegaBatchID, FrontsOrders.FrontID, FrontsOrders.Width, Packages.PackageStatusID";
            MarkTPSReadyFrontsDT.Clear();
            if (Marketing)
            {
                using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DA.Fill(MarkTPSReadyFrontsDT);
                }
                for (int i = 0; i < MarkTPSReadyFrontsDT.Rows.Count; i++)
                    MarkTPSReadyFrontsDT.Rows[i]["GroupType"] = 1;
            }
            if (ZOV)
            {
                using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.ZOVOrdersConnectionString))
                {
                    DA.Fill(DT);
                }
                for (int i = 0; i < DT.Rows.Count; i++)
                    DT.Rows[i]["GroupType"] = 0;
                foreach (DataRow item in DT.Rows)
                    MarkTPSReadyFrontsDT.ImportRow(item);
            }

            SelectCommand = @"SELECT Batch.MegaBatchID, FrontsOrders.FrontID, FrontsOrders.Width, FrontsOrders.FactoryID AS FFactoryID, BatchDetails.FactoryID AS BFactoryID, SUM(FrontsOrders.Count) AS AllCount, SUM(FrontsOrders.Square) AS AllSquare
            FROM FrontsOrders INNER JOIN
                      BatchDetails ON FrontsOrders.MainOrderID = BatchDetails.MainOrderID INNER JOIN
                      Batch ON BatchDetails.BatchID = Batch.BatchID INNER JOIN
                      MegaBatch ON Batch.MegaBatchID = MegaBatch.MegaBatchID AND CAST(MegaBatch.CreateDateTime AS Date) >= '" + FirstDate.ToString("yyyy-MM-dd") +
                "' AND CAST(MegaBatch.CreateDateTime AS Date) <= '" + SecondDate.ToString("yyyy-MM-dd") + @"'
            GROUP BY Batch.MegaBatchID, FrontsOrders.FrontID, FrontsOrders.Width, FrontsOrders.FactoryID, BatchDetails.FactoryID
            ORDER BY Batch.MegaBatchID, FrontsOrders.FrontID, FrontsOrders.Width, FrontsOrders.FactoryID, BatchDetails.FactoryID";
            MarkAllFrontsDT.Clear();
            if (Marketing)
            {
                using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DA.Fill(MarkAllFrontsDT);
                }
                for (int i = 0; i < MarkAllFrontsDT.Rows.Count; i++)
                    MarkAllFrontsDT.Rows[i]["GroupType"] = 1;
            }
            if (ZOV)
            {
                using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.ZOVOrdersConnectionString))
                {
                    DA.Fill(DT);
                }
                for (int i = 0; i < DT.Rows.Count; i++)
                    DT.Rows[i]["GroupType"] = 0;
                foreach (DataRow item in DT.Rows)
                    MarkAllFrontsDT.ImportRow(item);
            }
            DT.Dispose();
        }

        public void UpdateDecorOrders(bool Marketing, bool ZOV, DateTime FirstDate, DateTime SecondDate)
        {
            string SelectCommand = "SELECT Batch.MegaBatchID, DecorOrders.ProductID," +
                " SUM(DecorOrders.Count) AS AllCount" +
                " FROM DecorOrders" +
                " INNER JOIN BatchDetails ON DecorOrders.MainOrderID = BatchDetails.MainOrderID" +
                " INNER JOIN Batch ON BatchDetails.BatchID = Batch.BatchID INNER JOIN" +
                     " MegaBatch ON Batch.MegaBatchID = MegaBatch.MegaBatchID AND CAST(MegaBatch.CreateDateTime AS Date) >= '" + FirstDate.ToString("yyyy-MM-dd") +
                "' AND CAST(MegaBatch.CreateDateTime AS Date) <= '" + SecondDate.ToString("yyyy-MM-dd") + @"'
                  INNER JOIN MainOrders ON DecorOrders.MainOrderID = MainOrders.MainOrderID" +
                " WHERE (DecorOrders.FactoryID = 1) AND MainOrders.ProfilProductionStatusID = 3" +
                " GROUP BY Batch.MegaBatchID, DecorOrders.ProductID" +
                " ORDER BY Batch.MegaBatchID, DecorOrders.ProductID";
            DataTable DT = MarkProfilOnProdDecorDT.Clone();
            MarkProfilOnProdDecorDT.Clear();
            if (Marketing)
            {
                using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DA.Fill(MarkProfilOnProdDecorDT);
                }
                for (int i = 0; i < MarkProfilOnProdDecorDT.Rows.Count; i++)
                    MarkProfilOnProdDecorDT.Rows[i]["GroupType"] = 1;
            }
            if (ZOV)
            {
                using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.ZOVOrdersConnectionString))
                {
                    DA.Fill(DT);
                }
                for (int i = 0; i < DT.Rows.Count; i++)
                    DT.Rows[i]["GroupType"] = 0;
                foreach (DataRow item in DT.Rows)
                    MarkProfilOnProdDecorDT.ImportRow(item);
            }

            SelectCommand = "SELECT Batch.MegaBatchID, DecorOrders.ProductID," +
                " SUM(DecorOrders.Count) AS AllCount" +
                " FROM DecorOrders" +
                " INNER JOIN BatchDetails ON DecorOrders.MainOrderID = BatchDetails.MainOrderID" +
                " INNER JOIN Batch ON BatchDetails.BatchID = Batch.BatchID INNER JOIN" +
                     " MegaBatch ON Batch.MegaBatchID = MegaBatch.MegaBatchID AND CAST(MegaBatch.CreateDateTime AS Date) >= '" + FirstDate.ToString("yyyy-MM-dd") +
                "' AND CAST(MegaBatch.CreateDateTime AS Date) <= '" + SecondDate.ToString("yyyy-MM-dd") + @"'
                INNER JOIN MainOrders ON DecorOrders.MainOrderID = MainOrders.MainOrderID" +
                " WHERE (DecorOrders.FactoryID = 2) AND MainOrders.TPSProductionStatusID = 3" +
                " GROUP BY Batch.MegaBatchID, DecorOrders.ProductID" +
                " ORDER BY Batch.MegaBatchID, DecorOrders.ProductID";
            MarkTPSOnProdDecorDT.Clear();
            if (Marketing)
            {
                using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DA.Fill(MarkTPSOnProdDecorDT);
                }
                for (int i = 0; i < MarkTPSOnProdDecorDT.Rows.Count; i++)
                    MarkTPSOnProdDecorDT.Rows[i]["GroupType"] = 1;
            }
            if (ZOV)
            {
                using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.ZOVOrdersConnectionString))
                {
                    DA.Fill(DT);
                }
                for (int i = 0; i < DT.Rows.Count; i++)
                    DT.Rows[i]["GroupType"] = 0;
                foreach (DataRow item in DT.Rows)
                    MarkTPSOnProdDecorDT.ImportRow(item);
            }

            SelectCommand = @"SELECT Batch.MegaBatchID, DecorOrders.ProductID, Packages.PackageStatusID, SUM(PackageDetails.Count) AS AllCount
            FROM Packages INNER JOIN
                      PackageDetails ON PackageDetails.PackageID = Packages.PackageID INNER JOIN
                      DecorOrders ON PackageDetails.OrderID = DecorOrders.DecorOrderID INNER JOIN
                      BatchDetails ON Packages.MainOrderID = BatchDetails.MainOrderID AND BatchDetails.FactoryID = 1 INNER JOIN
                      Batch ON BatchDetails.BatchID = Batch.BatchID INNER JOIN
                      MegaBatch ON Batch.MegaBatchID = MegaBatch.MegaBatchID AND CAST(MegaBatch.CreateDateTime AS Date) >= '" + FirstDate.ToString("yyyy-MM-dd") +
                "' AND CAST(MegaBatch.CreateDateTime AS Date) <= '" + SecondDate.ToString("yyyy-MM-dd") + @"'
            WHERE (Packages.ProductType = 1) AND (Packages.FactoryID = 1) AND PackageStatusID > 0
            GROUP BY Batch.MegaBatchID, DecorOrders.ProductID, Packages.PackageStatusID
            ORDER BY Batch.MegaBatchID, DecorOrders.ProductID, Packages.PackageStatusID";
            MarkProfilReadyDecorDT.Clear();
            if (Marketing)
            {
                using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DA.Fill(MarkProfilReadyDecorDT);
                }
                for (int i = 0; i < MarkProfilReadyDecorDT.Rows.Count; i++)
                    MarkProfilReadyDecorDT.Rows[i]["GroupType"] = 1;
            }
            {
                SelectCommand = "SELECT Batch.MegaBatchID, DecorOrders.ProductID," +
                                " SUM(DecorOrders.Count) AS AllCount" +
                                " FROM DecorOrders" +
                                " INNER JOIN BatchDetails ON DecorOrders.MainOrderID = BatchDetails.MainOrderID" +
                                " INNER JOIN Batch ON BatchDetails.BatchID = Batch.BatchID INNER JOIN" +
                                " MegaBatch ON Batch.MegaBatchID = MegaBatch.MegaBatchID AND CAST(MegaBatch.CreateDateTime AS Date) >= '" + FirstDate.ToString("yyyy-MM-dd") +
                                "' AND CAST(MegaBatch.CreateDateTime AS Date) <= '" + SecondDate.ToString("yyyy-MM-dd") + @"'
                  INNER JOIN MainOrders ON DecorOrders.MainOrderID = MainOrders.MainOrderID" +
                                " WHERE DecorOrders.MainOrderID not in (select mainorderId from packages) and (DecorOrders.FactoryID = 1) AND MainOrders.ProfilDispatchStatusID = 2" +
                                " GROUP BY Batch.MegaBatchID, DecorOrders.ProductID" +
                                " ORDER BY Batch.MegaBatchID, DecorOrders.ProductID";
                using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DA.Fill(DT);
                }
                for (int i = 0; i < DT.Rows.Count; i++)
                    DT.Rows[i]["GroupType"] = 1;
                foreach (DataRow item in DT.Rows)
                    MarkProfilReadyDecorDT.ImportRow(item);
            }

            SelectCommand = @"SELECT Batch.MegaBatchID, DecorOrders.ProductID, Packages.PackageStatusID, SUM(PackageDetails.Count) AS AllCount
            FROM Packages INNER JOIN
                      PackageDetails ON PackageDetails.PackageID = Packages.PackageID INNER JOIN
                      DecorOrders ON PackageDetails.OrderID = DecorOrders.DecorOrderID INNER JOIN
                      BatchDetails ON Packages.MainOrderID = BatchDetails.MainOrderID AND BatchDetails.FactoryID = 2 INNER JOIN
                      Batch ON BatchDetails.BatchID = Batch.BatchID INNER JOIN
                      MegaBatch ON Batch.MegaBatchID = MegaBatch.MegaBatchID AND CAST(MegaBatch.CreateDateTime AS Date) >= '" + FirstDate.ToString("yyyy-MM-dd") +
                "' AND CAST(MegaBatch.CreateDateTime AS Date) <= '" + SecondDate.ToString("yyyy-MM-dd") + @"'
            WHERE (Packages.ProductType = 1) AND (Packages.FactoryID = 2) AND PackageStatusID > 0
            GROUP BY Batch.MegaBatchID, DecorOrders.ProductID, Packages.PackageStatusID
            ORDER BY Batch.MegaBatchID, DecorOrders.ProductID, Packages.PackageStatusID";
            MarkTPSReadyDecorDT.Clear();
            if (Marketing)
            {
                using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DA.Fill(MarkTPSReadyDecorDT);
                }
                for (int i = 0; i < MarkTPSReadyDecorDT.Rows.Count; i++)
                    MarkTPSReadyDecorDT.Rows[i]["GroupType"] = 1;
            }
            if (ZOV)
            {
                using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.ZOVOrdersConnectionString))
                {
                    DA.Fill(DT);
                }
                for (int i = 0; i < DT.Rows.Count; i++)
                    DT.Rows[i]["GroupType"] = 0;
                foreach (DataRow item in DT.Rows)
                    MarkTPSReadyDecorDT.ImportRow(item);
            }

            SelectCommand = @"SELECT Batch.MegaBatchID, DecorOrders.ProductID, DecorOrders.FactoryID AS FFactoryID, BatchDetails.FactoryID AS BFactoryID, SUM(DecorOrders.Count) AS AllCount
            FROM DecorOrders INNER JOIN
                      BatchDetails ON DecorOrders.MainOrderID = BatchDetails.MainOrderID INNER JOIN
                      Batch ON BatchDetails.BatchID = Batch.BatchID INNER JOIN
                      MegaBatch ON Batch.MegaBatchID = MegaBatch.MegaBatchID AND CAST(MegaBatch.CreateDateTime AS Date) >= '" + FirstDate.ToString("yyyy-MM-dd") +
                "' AND CAST(MegaBatch.CreateDateTime AS Date) <= '" + SecondDate.ToString("yyyy-MM-dd") + @"'
            GROUP BY Batch.MegaBatchID, DecorOrders.ProductID, DecorOrders.FactoryID, BatchDetails.FactoryID
            ORDER BY Batch.MegaBatchID, DecorOrders.ProductID, DecorOrders.FactoryID, BatchDetails.FactoryID";
            MarkAllDecorDT.Clear();
            if (Marketing)
            {
                using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DA.Fill(MarkAllDecorDT);
                }
                for (int i = 0; i < MarkAllDecorDT.Rows.Count; i++)
                    MarkAllDecorDT.Rows[i]["GroupType"] = 1;
            }
            if (ZOV)
            {
                using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.ZOVOrdersConnectionString))
                {
                    DA.Fill(DT);
                }
                for (int i = 0; i < DT.Rows.Count; i++)
                    DT.Rows[i]["GroupType"] = 0;
                foreach (DataRow item in DT.Rows)
                    MarkAllDecorDT.ImportRow(item);
            }
            DT.Dispose();
        }

        public void SimpleFrontsGeneralSummary()
        {
            for (int i = 0; i < MegaBatchDT.Rows.Count; i++)
            {
                MegaBatchDT.Rows[i]["ProfilOnProductionPerc"] = 0;
                MegaBatchDT.Rows[i]["ProfilInProductionPerc"] = 0;
                MegaBatchDT.Rows[i]["ProfilReadyPerc"] = 0;
                MegaBatchDT.Rows[i]["ProfilPackingDate"] = string.Empty;
                MegaBatchDT.Rows[i]["ProfilReady"] = false;

                MegaBatchDT.Rows[i]["TPSOnProductionPerc"] = 0;
                MegaBatchDT.Rows[i]["TPSInProductionPerc"] = 0;
                MegaBatchDT.Rows[i]["FilenkaPerc"] = 0;
                MegaBatchDT.Rows[i]["TrimmingPerc"] = 0;
                MegaBatchDT.Rows[i]["AssemblyPerc"] = 0;
                MegaBatchDT.Rows[i]["DeyingPerc"] = 0;
                MegaBatchDT.Rows[i]["TPSReadyPerc"] = 0;
                MegaBatchDT.Rows[i]["TPSPackingDate"] = string.Empty;
                MegaBatchDT.Rows[i]["TPSReady"] = false;
            }

            int ProfilOnProductionCount = 0;
            int ProfilInProductionCount = 0;
            int ProfilReadyCount = 0;
            int ProfilAllCount = 0;

            int ProfilOnProductionPerc = 0;
            int ProfilInProductionPerc = 0;
            int ProfilReadyPerc = 0;

            decimal ProfilOnProductionProgressVal = 0;
            decimal ProfilInProductionProgressVal = 0;
            decimal ProfilReadyProgressVal = 0;

            int TPSOnProductionCount = 0;
            int TPSInProductionCount = 0;
            int TPSReadyCount = 0;
            int TPSAllCount = 0;

            int TPSOnProductionPerc = 0;
            int TPSInProductionPerc = 0;
            int TPSReadyPerc = 0;

            decimal TPSOnProductionProgressVal = 0;
            decimal TPSInProductionProgressVal = 0;
            decimal TPSReadyProgressVal = 0;

            decimal d1 = 0;
            decimal d2 = 0;
            decimal d3 = 0;
            decimal d4 = 0;
            decimal d5 = 0;
            decimal d6 = 0;

            decimal Square = 0;

            for (int i = 0; i < MegaBatchDT.Rows.Count; i++)
            {
                int MegaBatchID = Convert.ToInt32(MegaBatchDT.Rows[i]["MegaBatchID"]);
                int GroupType = Convert.ToInt32(MegaBatchDT.Rows[i]["GroupType"]);
                decimal FilenkaPerc = 0;
                decimal TrimmingPerc = 0;
                decimal AssemblyPerc = 0;
                decimal DeyingPerc = 0;

                ProfilReadyCount = MarketProfilFrontsInfo(-1, MegaBatchID, 0, 0, ref Square);
                ProfilAllCount = AllFrontsInfo(-1, MegaBatchID, 0, 0, 1, ref Square);
                ProfilOnProductionCount = ProfilOnProdFrontsInfo(-1, MegaBatchID, 0, 0, ref Square);
                ProfilInProductionCount = ProfilAllCount - ProfilReadyCount - ProfilOnProductionCount;

                TPSReadyCount = MarketTPSFrontsInfo(-1, MegaBatchID, 0, 0, ref Square);
                TPSAllCount = AllFrontsInfo(-1, MegaBatchID, 0, 0, 2, ref Square);
                TPSOnProductionCount = TPSOnProdFrontsInfo(-1, MegaBatchID, 0, 0, ref Square);
                TPSInProductionCount = TPSAllCount - TPSReadyCount - TPSOnProductionCount;

                if (TPSAllCount > 0)
                    SectorsInfo(GroupType, MegaBatchID, ref FilenkaPerc, ref TrimmingPerc, ref AssemblyPerc, ref DeyingPerc);

                ProfilOnProductionProgressVal = 0;
                ProfilInProductionProgressVal = 0;
                ProfilReadyProgressVal = 0;

                TPSOnProductionProgressVal = 0;
                TPSInProductionProgressVal = 0;
                TPSReadyProgressVal = 0;

                if (ProfilAllCount > 0)
                    ProfilOnProductionProgressVal = Convert.ToDecimal(Convert.ToDecimal(ProfilOnProductionCount) / Convert.ToDecimal(ProfilAllCount));

                if (ProfilAllCount > 0)
                    ProfilInProductionProgressVal = Convert.ToDecimal(Convert.ToDecimal(ProfilInProductionCount) / Convert.ToDecimal(ProfilAllCount));

                if (ProfilAllCount > 0)
                    ProfilReadyProgressVal = Convert.ToDecimal(Convert.ToDecimal(ProfilReadyCount) / Convert.ToDecimal(ProfilAllCount));

                if (TPSAllCount > 0)
                    TPSOnProductionProgressVal = Convert.ToDecimal(Convert.ToDecimal(TPSOnProductionCount) / Convert.ToDecimal(TPSAllCount));

                if (TPSAllCount > 0)
                    TPSInProductionProgressVal = Convert.ToDecimal(Convert.ToDecimal(TPSInProductionCount) / Convert.ToDecimal(TPSAllCount));

                if (TPSAllCount > 0)
                    TPSReadyProgressVal = Convert.ToDecimal(Convert.ToDecimal(TPSReadyCount) / Convert.ToDecimal(TPSAllCount));

                d1 = ProfilOnProductionProgressVal * 100;
                d2 = ProfilInProductionProgressVal * 100;
                d3 = ProfilReadyProgressVal * 100;
                d4 = TPSOnProductionProgressVal * 100;
                d5 = TPSInProductionProgressVal * 100;
                d6 = TPSReadyProgressVal * 100;

                if (d1 > 0 && d1 < 0.5M)
                    d1 = 1;
                if (d2 > 0 && d2 < 0.5M)
                    d2 = 1;
                if (d3 > 0 && d3 < 0.5M)
                    d3 = 1;
                if (d4 > 0 && d4 < 0.5M)
                    d4 = 1;
                if (d5 > 0 && d5 < 0.5M)
                    d5 = 1;
                if (d6 > 0 && d6 < 0.5M)
                    d6 = 1;

                if (d1 > 99.5M && d1 < 100)
                    d1 = 99;
                if (d2 > 99.5M && d2 < 100)
                    d2 = 99;
                if (d3 > 99.5M && d3 < 100)
                    d3 = 99;
                if (d4 > 99.5M && d4 < 100)
                    d4 = 99;
                if (d5 > 99.5M && d5 < 100)
                    d5 = 99;
                if (d6 > 99.5M && d6 < 100)
                    d6 = 99;

                ProfilOnProductionPerc = Convert.ToInt32(Math.Round(d1, 1, MidpointRounding.AwayFromZero));
                ProfilInProductionPerc = Convert.ToInt32(Math.Round(d2, 1, MidpointRounding.AwayFromZero));
                ProfilReadyPerc = Convert.ToInt32(Math.Round(d3, 1, MidpointRounding.AwayFromZero));

                if ((ProfilOnProductionPerc + ProfilInProductionPerc + ProfilReadyPerc) > 100)
                {
                    if (ProfilOnProductionPerc > ProfilInProductionPerc)
                        if (ProfilOnProductionPerc > ProfilReadyPerc)
                            ProfilOnProductionPerc--;
                        else
                            ProfilReadyPerc--;
                    else
                        if (ProfilInProductionPerc > ProfilReadyPerc)
                        ProfilInProductionPerc--;
                    else
                        ProfilReadyPerc--;
                }

                TPSOnProductionPerc = Convert.ToInt32(Math.Round(d4, 1, MidpointRounding.AwayFromZero));
                TPSInProductionPerc = Convert.ToInt32(Math.Round(d5, 1, MidpointRounding.AwayFromZero));
                FilenkaPerc = Convert.ToInt32(Math.Round(FilenkaPerc * 100, 1, MidpointRounding.AwayFromZero));
                TrimmingPerc = Convert.ToInt32(Math.Round(TrimmingPerc * 100, 1, MidpointRounding.AwayFromZero));
                AssemblyPerc = Convert.ToInt32(Math.Round(AssemblyPerc * 100, 1, MidpointRounding.AwayFromZero));
                DeyingPerc = Convert.ToInt32(Math.Round(DeyingPerc * 100, 1, MidpointRounding.AwayFromZero));
                TPSReadyPerc = Convert.ToInt32(Math.Round(d6, 1, MidpointRounding.AwayFromZero));

                if ((TPSOnProductionPerc + TPSInProductionPerc + TPSReadyPerc) > 100)
                {
                    if (TPSOnProductionPerc > TPSInProductionPerc)
                        if (TPSOnProductionPerc > TPSReadyPerc)
                            TPSOnProductionPerc--;
                        else
                            TPSReadyPerc--;
                    else
                        if (TPSInProductionPerc > TPSReadyPerc)
                        TPSInProductionPerc--;
                    else
                        TPSReadyPerc--;
                }

                string ProductionDate = EnterInProdDate(GroupType, MegaBatchID, 1);
                if (ProductionDate.Length > 0)
                    MegaBatchDT.Rows[i]["ProfilEntryDateTime"] = ProductionDate;

                ProductionDate = EnterInProdDate(GroupType, MegaBatchID, 2);
                if (ProductionDate.Length > 0)
                    MegaBatchDT.Rows[i]["TPSEntryDateTime"] = ProductionDate;

                if (ProfilOnProductionPerc > 0)
                    MegaBatchDT.Rows[i]["ProfilOnProductionPerc"] = ProfilOnProductionPerc;
                if (ProfilInProductionPerc > 0)
                    MegaBatchDT.Rows[i]["ProfilInProductionPerc"] = ProfilInProductionPerc;
                if (ProfilReadyPerc > 0)
                    MegaBatchDT.Rows[i]["ProfilReadyPerc"] = ProfilReadyPerc;
                if (ProfilReadyPerc == 100)
                {
                    MegaBatchDT.Rows[i]["ProfilReady"] = true;
                    MegaBatchDT.Rows[i]["ProfilPackingDate"] = ReadyDate(GroupType, MegaBatchID, 1);
                }
                else
                    MegaBatchDT.Rows[i]["ProfilReady"] = false;
                //if (ProfilReadyPerc == 0)
                //{
                //    MegaBatchDT.Rows[i]["ProfilReady"] = true;
                //}

                if (TPSOnProductionPerc > 0)
                    MegaBatchDT.Rows[i]["TPSOnProductionPerc"] = TPSOnProductionPerc;
                if (TPSInProductionPerc > 0)
                    MegaBatchDT.Rows[i]["TPSInProductionPerc"] = TPSInProductionPerc;
                if (FilenkaPerc > 0)
                    MegaBatchDT.Rows[i]["FilenkaPerc"] = FilenkaPerc;
                if (TrimmingPerc > 0)
                    MegaBatchDT.Rows[i]["TrimmingPerc"] = TrimmingPerc;
                if (AssemblyPerc > 0)
                    MegaBatchDT.Rows[i]["AssemblyPerc"] = AssemblyPerc;
                if (DeyingPerc > 0)
                    MegaBatchDT.Rows[i]["DeyingPerc"] = DeyingPerc;
                if (TPSReadyPerc > 0)
                    MegaBatchDT.Rows[i]["TPSReadyPerc"] = TPSReadyPerc;
                if (TPSReadyPerc == 100)
                {
                    MegaBatchDT.Rows[i]["TPSReady"] = true;
                    MegaBatchDT.Rows[i]["TPSPackingDate"] = ReadyDate(GroupType, MegaBatchID, 2);
                }
                else
                    MegaBatchDT.Rows[i]["TPSReady"] = false;
                //if (TPSReadyPerc == 0)
                //{
                //    MegaBatchDT.Rows[i]["TPSReady"] = true;
                //}
            }
            MegaBatchBS.MoveFirst();
        }

        public void CurvedFrontsGeneralSummary()
        {
            for (int i = 0; i < MegaBatchDT.Rows.Count; i++)
            {
                MegaBatchDT.Rows[i]["ProfilOnProductionPerc"] = 0;
                MegaBatchDT.Rows[i]["ProfilInProductionPerc"] = 0;
                MegaBatchDT.Rows[i]["ProfilReadyPerc"] = 0;
                MegaBatchDT.Rows[i]["ProfilPackingDate"] = string.Empty;
                MegaBatchDT.Rows[i]["ProfilReady"] = false;

                MegaBatchDT.Rows[i]["TPSOnProductionPerc"] = 0;
                MegaBatchDT.Rows[i]["TPSInProductionPerc"] = 0;
                MegaBatchDT.Rows[i]["TPSReadyPerc"] = 0;
                MegaBatchDT.Rows[i]["TPSPackingDate"] = string.Empty;
                MegaBatchDT.Rows[i]["TPSReady"] = false;
            }

            int ProfilOnProductionCount = 0;
            int ProfilInProductionCount = 0;
            int ProfilReadyCount = 0;
            int ProfilAllCount = 0;

            int ProfilOnProductionPerc = 0;
            int ProfilInProductionPerc = 0;
            int ProfilReadyPerc = 0;

            decimal ProfilOnProductionProgressVal = 0;
            decimal ProfilInProductionProgressVal = 0;
            decimal ProfilReadyProgressVal = 0;

            int TPSOnProductionCount = 0;
            int TPSInProductionCount = 0;
            int TPSReadyCount = 0;
            int TPSAllCount = 0;

            int TPSOnProductionPerc = 0;
            int TPSInProductionPerc = 0;
            int TPSReadyPerc = 0;

            decimal TPSOnProductionProgressVal = 0;
            decimal TPSInProductionProgressVal = 0;
            decimal TPSReadyProgressVal = 0;

            decimal d1 = 0;
            decimal d2 = 0;
            decimal d3 = 0;
            decimal d4 = 0;
            decimal d5 = 0;
            decimal d6 = 0;

            decimal Square = 0;

            for (int i = 0; i < MegaBatchDT.Rows.Count; i++)
            {
                int MegaBatchID = Convert.ToInt32(MegaBatchDT.Rows[i]["MegaBatchID"]);
                int GroupType = Convert.ToInt32(MegaBatchDT.Rows[i]["GroupType"]);

                ProfilReadyCount = MarketProfilFrontsInfo(-1, MegaBatchID, 0, -1, ref Square);
                ProfilAllCount = AllFrontsInfo(-1, MegaBatchID, 0, -1, 1, ref Square);
                ProfilOnProductionCount = ProfilOnProdFrontsInfo(-1, MegaBatchID, 0, -1, ref Square);
                ProfilInProductionCount = ProfilAllCount - ProfilReadyCount - ProfilOnProductionCount;

                TPSReadyCount = MarketTPSFrontsInfo(-1, MegaBatchID, 0, -1, ref Square);
                TPSAllCount = AllFrontsInfo(-1, MegaBatchID, 0, -1, 2, ref Square);
                TPSOnProductionCount = TPSOnProdFrontsInfo(-1, MegaBatchID, 0, -1, ref Square);
                TPSInProductionCount = TPSAllCount - TPSReadyCount - TPSOnProductionCount;

                ProfilOnProductionProgressVal = 0;
                ProfilInProductionProgressVal = 0;
                ProfilReadyProgressVal = 0;

                TPSOnProductionProgressVal = 0;
                TPSInProductionProgressVal = 0;
                TPSReadyProgressVal = 0;

                if (ProfilAllCount > 0)
                    ProfilOnProductionProgressVal = Convert.ToDecimal(Convert.ToDecimal(ProfilOnProductionCount) / Convert.ToDecimal(ProfilAllCount));

                if (ProfilAllCount > 0)
                    ProfilInProductionProgressVal = Convert.ToDecimal(Convert.ToDecimal(ProfilInProductionCount) / Convert.ToDecimal(ProfilAllCount));

                if (ProfilAllCount > 0)
                    ProfilReadyProgressVal = Convert.ToDecimal(Convert.ToDecimal(ProfilReadyCount) / Convert.ToDecimal(ProfilAllCount));

                if (TPSAllCount > 0)
                    TPSOnProductionProgressVal = Convert.ToDecimal(Convert.ToDecimal(TPSOnProductionCount) / Convert.ToDecimal(TPSAllCount));

                if (TPSAllCount > 0)
                    TPSInProductionProgressVal = Convert.ToDecimal(Convert.ToDecimal(TPSInProductionCount) / Convert.ToDecimal(TPSAllCount));

                if (TPSAllCount > 0)
                    TPSReadyProgressVal = Convert.ToDecimal(Convert.ToDecimal(TPSReadyCount) / Convert.ToDecimal(TPSAllCount));

                d1 = ProfilOnProductionProgressVal * 100;
                d2 = ProfilInProductionProgressVal * 100;
                d3 = ProfilReadyProgressVal * 100;
                d4 = TPSOnProductionProgressVal * 100;
                d5 = TPSInProductionProgressVal * 100;
                d6 = TPSReadyProgressVal * 100;

                if (d1 > 0 && d1 < 0.5M)
                    d1 = 1;
                if (d2 > 0 && d2 < 0.5M)
                    d2 = 1;
                if (d3 > 0 && d3 < 0.5M)
                    d3 = 1;
                if (d4 > 0 && d4 < 0.5M)
                    d4 = 1;
                if (d5 > 0 && d5 < 0.5M)
                    d5 = 1;
                if (d6 > 0 && d6 < 0.5M)
                    d6 = 1;

                if (d1 > 99.5M && d1 < 100)
                    d1 = 99;
                if (d2 > 99.5M && d2 < 100)
                    d2 = 99;
                if (d3 > 99.5M && d3 < 100)
                    d3 = 99;
                if (d4 > 99.5M && d4 < 100)
                    d4 = 99;
                if (d5 > 99.5M && d5 < 100)
                    d5 = 99;
                if (d6 > 99.5M && d6 < 100)
                    d6 = 99;

                ProfilOnProductionPerc = Convert.ToInt32(Math.Round(d1, 1, MidpointRounding.AwayFromZero));
                ProfilInProductionPerc = Convert.ToInt32(Math.Round(d2, 1, MidpointRounding.AwayFromZero));
                ProfilReadyPerc = Convert.ToInt32(Math.Round(d3, 1, MidpointRounding.AwayFromZero));

                if ((ProfilOnProductionPerc + ProfilInProductionPerc + ProfilReadyPerc) > 100)
                {
                    if (ProfilOnProductionPerc > ProfilInProductionPerc)
                        if (ProfilOnProductionPerc > ProfilReadyPerc)
                            ProfilOnProductionPerc--;
                        else
                            ProfilReadyPerc--;
                    else
                        if (ProfilInProductionPerc > ProfilReadyPerc)
                        ProfilInProductionPerc--;
                    else
                        ProfilReadyPerc--;
                }

                TPSOnProductionPerc = Convert.ToInt32(Math.Round(d4, 1, MidpointRounding.AwayFromZero));
                TPSInProductionPerc = Convert.ToInt32(Math.Round(d5, 1, MidpointRounding.AwayFromZero));
                TPSReadyPerc = Convert.ToInt32(Math.Round(d6, 1, MidpointRounding.AwayFromZero));

                if ((TPSOnProductionPerc + TPSInProductionPerc + TPSReadyPerc) > 100)
                {
                    if (TPSOnProductionPerc > TPSInProductionPerc)
                        if (TPSOnProductionPerc > TPSReadyPerc)
                            TPSOnProductionPerc--;
                        else
                            TPSReadyPerc--;
                    else
                        if (TPSInProductionPerc > TPSReadyPerc)
                        TPSInProductionPerc--;
                    else
                        TPSReadyPerc--;
                }

                string ProductionDate = EnterInProdDate(GroupType, MegaBatchID, 1);
                if (ProductionDate.Length > 0)
                    MegaBatchDT.Rows[i]["ProfilEntryDateTime"] = ProductionDate;

                ProductionDate = EnterInProdDate(GroupType, MegaBatchID, 2);
                if (ProductionDate.Length > 0)
                    MegaBatchDT.Rows[i]["TPSEntryDateTime"] = ProductionDate;

                if (ProfilOnProductionPerc > 0)
                    MegaBatchDT.Rows[i]["ProfilOnProductionPerc"] = ProfilOnProductionPerc;
                if (ProfilInProductionPerc > 0)
                    MegaBatchDT.Rows[i]["ProfilInProductionPerc"] = ProfilInProductionPerc;
                if (ProfilReadyPerc > 0)
                    MegaBatchDT.Rows[i]["ProfilReadyPerc"] = ProfilReadyPerc;
                if (ProfilReadyPerc == 100)
                {
                    MegaBatchDT.Rows[i]["ProfilReady"] = true;
                    MegaBatchDT.Rows[i]["ProfilPackingDate"] = ReadyDate(GroupType, MegaBatchID, 1);
                }
                else
                    MegaBatchDT.Rows[i]["ProfilReady"] = false;
                //if (ProfilReadyPerc == 0)
                //{
                //    MegaBatchDT.Rows[i]["ProfilReady"] = true;
                //}

                if (TPSOnProductionPerc > 0)
                    MegaBatchDT.Rows[i]["TPSOnProductionPerc"] = TPSOnProductionPerc;
                if (TPSInProductionPerc > 0)
                    MegaBatchDT.Rows[i]["TPSInProductionPerc"] = TPSInProductionPerc;
                if (TPSReadyPerc > 0)
                    MegaBatchDT.Rows[i]["TPSReadyPerc"] = TPSReadyPerc;
                if (TPSReadyPerc == 100)
                {
                    MegaBatchDT.Rows[i]["TPSReady"] = true;
                    MegaBatchDT.Rows[i]["TPSPackingDate"] = ReadyDate(GroupType, MegaBatchID, 2);
                }
                else
                    MegaBatchDT.Rows[i]["TPSReady"] = false;
                //if (TPSReadyPerc == 0)
                //{
                //    MegaBatchDT.Rows[i]["TPSReady"] = true;
                //}
            }
            MegaBatchBS.MoveFirst();
        }

        public void DecorGeneralSummary()
        {
            for (int i = 0; i < MegaBatchDT.Rows.Count; i++)
            {
                MegaBatchDT.Rows[i]["ProfilOnProductionPerc"] = 0;
                MegaBatchDT.Rows[i]["ProfilInProductionPerc"] = 0;
                MegaBatchDT.Rows[i]["ProfilReadyPerc"] = 0;
                MegaBatchDT.Rows[i]["ProfilPackingDate"] = string.Empty;
                MegaBatchDT.Rows[i]["ProfilReady"] = false;

                MegaBatchDT.Rows[i]["TPSOnProductionPerc"] = 0;
                MegaBatchDT.Rows[i]["TPSInProductionPerc"] = 0;
                MegaBatchDT.Rows[i]["TPSReadyPerc"] = 0;
                MegaBatchDT.Rows[i]["TPSPackingDate"] = string.Empty;
                MegaBatchDT.Rows[i]["TPSReady"] = false;
            }

            int ProfilOnProductionCount = 0;
            int ProfilInProductionCount = 0;
            int ProfilReadyCount = 0;
            int ProfilAllCount = 0;

            int ProfilOnProductionPerc = 0;
            int ProfilInProductionPerc = 0;
            int ProfilReadyPerc = 0;

            decimal ProfilOnProductionProgressVal = 0;
            decimal ProfilInProductionProgressVal = 0;
            decimal ProfilReadyProgressVal = 0;

            int TPSOnProductionCount = 0;
            int TPSInProductionCount = 0;
            int TPSReadyCount = 0;
            int TPSAllCount = 0;

            int TPSOnProductionPerc = 0;
            int TPSInProductionPerc = 0;
            int TPSReadyPerc = 0;

            decimal TPSOnProductionProgressVal = 0;
            decimal TPSInProductionProgressVal = 0;
            decimal TPSReadyProgressVal = 0;

            decimal d1 = 0;
            decimal d2 = 0;
            decimal d3 = 0;
            decimal d4 = 0;
            decimal d5 = 0;
            decimal d6 = 0;

            for (int i = 0; i < MegaBatchDT.Rows.Count; i++)
            {
                int MegaBatchID = Convert.ToInt32(MegaBatchDT.Rows[i]["MegaBatchID"]);
                int GroupType = Convert.ToInt32(MegaBatchDT.Rows[i]["GroupType"]);

                ProfilReadyCount = MarketProfilDecorInfo(-1, MegaBatchID, 0);
                ProfilAllCount = AllDecorInfo(-1, MegaBatchID, 0, 1);
                ProfilOnProductionCount = ProfilOnProdDecorInfo(-1, MegaBatchID, 0);
                ProfilInProductionCount = ProfilAllCount - ProfilReadyCount - ProfilOnProductionCount;

                TPSReadyCount = MarketTPSDecorInfo(-1, MegaBatchID, 0);
                TPSAllCount = AllDecorInfo(-1, MegaBatchID, 0, 2);
                TPSOnProductionCount = TPSOnProdDecorInfo(-1, MegaBatchID, 0);
                TPSInProductionCount = TPSAllCount - TPSReadyCount - TPSOnProductionCount;

                ProfilOnProductionProgressVal = 0;
                ProfilInProductionProgressVal = 0;
                ProfilReadyProgressVal = 0;

                TPSOnProductionProgressVal = 0;
                TPSInProductionProgressVal = 0;
                TPSReadyProgressVal = 0;

                if (ProfilAllCount > 0)
                    ProfilOnProductionProgressVal = Convert.ToDecimal(Convert.ToDecimal(ProfilOnProductionCount) / Convert.ToDecimal(ProfilAllCount));

                if (ProfilAllCount > 0)
                    ProfilInProductionProgressVal = Convert.ToDecimal(Convert.ToDecimal(ProfilInProductionCount) / Convert.ToDecimal(ProfilAllCount));

                if (ProfilAllCount > 0)
                    ProfilReadyProgressVal = Convert.ToDecimal(Convert.ToDecimal(ProfilReadyCount) / Convert.ToDecimal(ProfilAllCount));

                if (TPSAllCount > 0)
                    TPSOnProductionProgressVal = Convert.ToDecimal(Convert.ToDecimal(TPSOnProductionCount) / Convert.ToDecimal(TPSAllCount));

                if (TPSAllCount > 0)
                    TPSInProductionProgressVal = Convert.ToDecimal(Convert.ToDecimal(TPSInProductionCount) / Convert.ToDecimal(TPSAllCount));

                if (TPSAllCount > 0)
                    TPSReadyProgressVal = Convert.ToDecimal(Convert.ToDecimal(TPSReadyCount) / Convert.ToDecimal(TPSAllCount));

                d1 = ProfilOnProductionProgressVal * 100;
                d2 = ProfilInProductionProgressVal * 100;
                d3 = ProfilReadyProgressVal * 100;
                d4 = TPSOnProductionProgressVal * 100;
                d5 = TPSInProductionProgressVal * 100;
                d6 = TPSReadyProgressVal * 100;

                if (d1 > 0 && d1 < 0.5M)
                    d1 = 1;
                if (d2 > 0 && d2 < 0.5M)
                    d2 = 1;
                if (d3 > 0 && d3 < 0.5M)
                    d3 = 1;
                if (d4 > 0 && d4 < 0.5M)
                    d4 = 1;
                if (d5 > 0 && d5 < 0.5M)
                    d5 = 1;
                if (d6 > 0 && d6 < 0.5M)
                    d6 = 1;

                if (d1 > 99.5M && d1 < 100)
                    d1 = 99;
                if (d2 > 99.5M && d2 < 100)
                    d2 = 99;
                if (d3 > 99.5M && d3 < 100)
                    d3 = 99;
                if (d4 > 99.5M && d4 < 100)
                    d4 = 99;
                if (d5 > 99.5M && d5 < 100)
                    d5 = 99;
                if (d6 > 99.5M && d6 < 100)
                    d6 = 99;

                ProfilOnProductionPerc = Convert.ToInt32(Math.Round(d1, 1, MidpointRounding.AwayFromZero));
                ProfilInProductionPerc = Convert.ToInt32(Math.Round(d2, 1, MidpointRounding.AwayFromZero));
                ProfilReadyPerc = Convert.ToInt32(Math.Round(d3, 1, MidpointRounding.AwayFromZero));

                if ((ProfilOnProductionPerc + ProfilInProductionPerc + ProfilReadyPerc) > 100)
                {
                    if (ProfilOnProductionPerc > ProfilInProductionPerc)
                        if (ProfilOnProductionPerc > ProfilReadyPerc)
                            ProfilOnProductionPerc--;
                        else
                            ProfilReadyPerc--;
                    else
                        if (ProfilInProductionPerc > ProfilReadyPerc)
                        ProfilInProductionPerc--;
                    else
                        ProfilReadyPerc--;
                }

                TPSOnProductionPerc = Convert.ToInt32(Math.Round(d4, 1, MidpointRounding.AwayFromZero));
                TPSInProductionPerc = Convert.ToInt32(Math.Round(d5, 1, MidpointRounding.AwayFromZero));
                TPSReadyPerc = Convert.ToInt32(Math.Round(d6, 1, MidpointRounding.AwayFromZero));

                if ((TPSOnProductionPerc + TPSInProductionPerc + TPSReadyPerc) > 100)
                {
                    if (TPSOnProductionPerc > TPSInProductionPerc)
                        if (TPSOnProductionPerc > TPSReadyPerc)
                            TPSOnProductionPerc--;
                        else
                            TPSReadyPerc--;
                    else
                        if (TPSInProductionPerc > TPSReadyPerc)
                        TPSInProductionPerc--;
                    else
                        TPSReadyPerc--;
                }

                string ProductionDate = EnterInProdDate(GroupType, MegaBatchID, 1);
                if (ProductionDate.Length > 0)
                    MegaBatchDT.Rows[i]["ProfilEntryDateTime"] = ProductionDate;

                ProductionDate = EnterInProdDate(GroupType, MegaBatchID, 2);
                if (ProductionDate.Length > 0)
                    MegaBatchDT.Rows[i]["TPSEntryDateTime"] = ProductionDate;

                if (ProfilOnProductionPerc > 0)
                    MegaBatchDT.Rows[i]["ProfilOnProductionPerc"] = ProfilOnProductionPerc;
                if (ProfilInProductionPerc > 0)
                    MegaBatchDT.Rows[i]["ProfilInProductionPerc"] = ProfilInProductionPerc;
                if (ProfilReadyPerc > 0)
                    MegaBatchDT.Rows[i]["ProfilReadyPerc"] = ProfilReadyPerc;
                if (ProfilReadyPerc == 100)
                {
                    MegaBatchDT.Rows[i]["ProfilReady"] = true;
                    MegaBatchDT.Rows[i]["ProfilPackingDate"] = ReadyDate(GroupType, MegaBatchID, 1);
                }
                else
                    MegaBatchDT.Rows[i]["ProfilReady"] = false;
                //if (ProfilReadyPerc == 0)
                //{
                //    MegaBatchDT.Rows[i]["ProfilReady"] = true;
                //}

                if (TPSOnProductionPerc > 0)
                    MegaBatchDT.Rows[i]["TPSOnProductionPerc"] = TPSOnProductionPerc;
                if (TPSInProductionPerc > 0)
                    MegaBatchDT.Rows[i]["TPSInProductionPerc"] = TPSInProductionPerc;
                if (TPSReadyPerc > 0)
                    MegaBatchDT.Rows[i]["TPSReadyPerc"] = TPSReadyPerc;
                if (TPSReadyPerc == 100)
                {
                    MegaBatchDT.Rows[i]["TPSReady"] = true;
                    MegaBatchDT.Rows[i]["TPSPackingDate"] = ReadyDate(GroupType, MegaBatchID, 2);
                }
                else
                    MegaBatchDT.Rows[i]["TPSReady"] = false;
                //if (TPSReadyPerc == 0)
                //{
                //    MegaBatchDT.Rows[i]["TPSReady"] = true;
                //}
            }
            MegaBatchBS.MoveFirst();
        }

        public void FilterSimpleFrontsOrders(int GroupType, int MegaBatchID, int FactoryID)
        {
            string BatchFactoryFilter = string.Empty;
            string FrontsFactoryFilter = string.Empty;
            string ConnectionString = ConnectionStrings.MarketingOrdersConnectionString;
            if (GroupType == 0)
                ConnectionString = ConnectionStrings.ZOVOrdersConnectionString;

            if (FactoryID == 1)
            {
                BatchFactoryFilter = " AND BatchDetails.FactoryID = 1";
                FrontsFactoryFilter = " AND FrontsOrders.FactoryID = 1";
            }
            if (FactoryID == 2)
            {
                BatchFactoryFilter = " AND BatchDetails.FactoryID = 2";
                FrontsFactoryFilter = " AND FrontsOrders.FactoryID = 2";
            }
            if (FactoryID == -1)
            {
                FrontsFactoryFilter = " AND FrontsOrders.FactoryID = -1";
            }

            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM FrontsOrders" +
                " WHERE Width<> -1 AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails" +
                " WHERE BatchID IN (SELECT BatchID FROM Batch WHERE MegaBatchID = " + MegaBatchID + ")" + BatchFactoryFilter + ")" + FrontsFactoryFilter,
                ConnectionString))
            {
                FrontsOrdersDT.Clear();
                DA.Fill(FrontsOrdersDT);
            }
            GetSimpleFronts();
            SimpleFrontsCount(GroupType, MegaBatchID);
        }

        public void FilterCurvedFrontsOrders(int GroupType, int MegaBatchID, int FactoryID)
        {
            string BatchFactoryFilter = string.Empty;
            string FrontsFactoryFilter = string.Empty;
            string ConnectionString = ConnectionStrings.MarketingOrdersConnectionString;
            if (GroupType == 0)
                ConnectionString = ConnectionStrings.ZOVOrdersConnectionString;

            if (FactoryID == 1)
            {
                BatchFactoryFilter = " AND BatchDetails.FactoryID = 1";
                FrontsFactoryFilter = " AND FrontsOrders.FactoryID = 1";
            }
            if (FactoryID == 2)
            {
                BatchFactoryFilter = " AND BatchDetails.FactoryID = 2";
                FrontsFactoryFilter = " AND FrontsOrders.FactoryID = 2";
            }
            if (FactoryID == -1)
            {
                FrontsFactoryFilter = " AND FrontsOrders.FactoryID = -1";
            }

            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM FrontsOrders" +
                " WHERE Width=-1 AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails" +
                " WHERE BatchID IN (SELECT BatchID FROM Batch WHERE MegaBatchID = " + MegaBatchID + ")" + BatchFactoryFilter + ")" + FrontsFactoryFilter,
                ConnectionString))
            {
                FrontsOrdersDT.Clear();
                DA.Fill(FrontsOrdersDT);
            }
            GetCurvedFronts();
            CurvedFrontsCount(GroupType, MegaBatchID);
        }

        public void FilterDecorOrders(int GroupType, int MegaBatchID, int FactoryID)
        {
            string BatchFactoryFilter = string.Empty;
            string DecorFactoryFilter = string.Empty;
            string ConnectionString = ConnectionStrings.MarketingOrdersConnectionString;
            if (GroupType == 0)
                ConnectionString = ConnectionStrings.ZOVOrdersConnectionString;

            if (FactoryID == 1)
            {
                BatchFactoryFilter = " AND BatchDetails.FactoryID = 1";
                DecorFactoryFilter = " AND DecorOrders.FactoryID = 1";
            }
            if (FactoryID == 2)
            {
                BatchFactoryFilter = " AND BatchDetails.FactoryID = 2";
                DecorFactoryFilter = " AND DecorOrders.FactoryID = 2";
            }
            if (FactoryID == -1)
            {
                DecorFactoryFilter = " AND DecorOrders.FactoryID = -1";
            }

            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT DecorOrders.ProductID, DecorOrders.DecorID, DecorOrders.ColorID,  DecorOrders.PatinaID, DecorOrders.FactoryID," +
                " DecorOrders.Height, DecorOrders.Length, DecorOrders.Width, DecorOrders.Count, DecorOrders.Cost, DecorOrders.DecorConfigID, " +
                " MeasureID, ClientID FROM DecorOrders" +
                " INNER JOIN MainOrders ON DecorOrders.MainOrderID = MainOrders.MainOrderID" +
                " INNER JOIN MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID" +
                " INNER JOIN infiniu2_catalog.dbo.DecorConfig" +
                " ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                " WHERE DecorOrders.MainOrderID IN (SELECT MainOrderID FROM BatchDetails" +
                " WHERE BatchID IN (SELECT BatchID FROM Batch WHERE MegaBatchID = " + MegaBatchID + ")" + BatchFactoryFilter + ")" + DecorFactoryFilter,
                ConnectionString))
            {
                DecorOrdersDT.Clear();
                DA.Fill(DecorOrdersDT);
            }
            GetDecorProducts(GroupType, MegaBatchID);
        }

        private void GetCurvedFronts()
        {
            CurvedFrontsSummaryDT.Clear();
            DataTable Table = new DataTable();
            using (DataView DV = new DataView(FrontsOrdersDT))
            {
                Table = DV.ToTable(true, new string[] { "FrontID" });
            }
            for (int i = 0; i < Table.Rows.Count; i++)
            {
                DataRow[] Rows = FrontsOrdersDT.Select("FrontID=" + Convert.ToInt32(Table.Rows[i]["FrontID"]) +
                    " AND Width=-1");
                if (Rows.Count() != 0)
                {
                    DataRow NewRow = CurvedFrontsSummaryDT.NewRow();
                    NewRow["Front"] = GetFrontName(Convert.ToInt32(Table.Rows[i]["FrontID"]));
                    NewRow["FrontID"] = Convert.ToInt32(Table.Rows[i]["FrontID"]);
                    NewRow["Width"] = -1;
                    CurvedFrontsSummaryDT.Rows.Add(NewRow);
                }
            }
            Table.Dispose();
            CurvedFrontsSummaryDT.DefaultView.Sort = "Front";
            CurvedFrontsSummaryBS.MoveFirst();
        }

        private void GetSimpleFronts()
        {
            SimpleFrontsSummaryDT.Clear();
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
                    DataRow NewRow = SimpleFrontsSummaryDT.NewRow();
                    NewRow["Front"] = GetFrontName(Convert.ToInt32(Table.Rows[i]["FrontID"]));
                    NewRow["FrontID"] = Convert.ToInt32(Table.Rows[i]["FrontID"]);
                    NewRow["Width"] = 0;
                    SimpleFrontsSummaryDT.Rows.Add(NewRow);
                }
            }
            Table.Dispose();
            SimpleFrontsSummaryDT.DefaultView.Sort = "Front";
            SimpleFrontsSummaryBS.MoveFirst();
        }

        private void GetDecorProducts(int GroupType, int MegaBatchID)
        {
            int AllCount = 0;
            int OnProdCount = 0;
            int InProdCount = 0;
            int ReadyCount = 0;

            decimal DecorProductCost = 0;
            decimal AllDecorProductCount = 0;
            decimal OnProdDecorProductCount = 0;
            decimal InProdDecorProductCount = 0;
            decimal ReadyDecorProductCount = 0;

            int decimals = 2;
            string Measure = string.Empty;
            string SelectCommand = string.Empty;
            DecorProductsSummaryDT.Clear();

            DataTable Table = new DataTable();

            using (DataView DV = new DataView(DecorConfigDT))
            {
                Table = DV.ToTable(true, new string[] { "ProductID", "MeasureID" });
            }

            for (int i = 0; i < Table.Rows.Count; i++)
            {
                for (int j = 1; j <= 2; j++)
                {
                    SelectCommand = "ProductID=" + Convert.ToInt32(Table.Rows[i]["ProductID"]) +
                        " AND MeasureID=" + Convert.ToInt32(Table.Rows[i]["MeasureID"]) + " AND FactoryID = " + j;

                    DataRow[] Rows = DecorOrdersDT.Select(SelectCommand);
                    if (Rows.Count() == 0)
                        continue;

                    foreach (DataRow row in Rows)
                    {
                        AllDecorProductCount = 0;
                        OnProdDecorProductCount = 0;
                        InProdDecorProductCount = 0;
                        ReadyDecorProductCount = 0;

                        int ProductID = Convert.ToInt32(row["ProductID"]);
                        int FactoryID = Convert.ToInt32(row["FactoryID"]);

                        if (FactoryID == 1)
                        {
                            AllCount = AllDecorInfo(GroupType, MegaBatchID, ProductID, 1);
                            ReadyCount = MarketProfilDecorInfo(GroupType, MegaBatchID, ProductID);
                            OnProdCount = ProfilOnProdDecorInfo(GroupType, MegaBatchID, ProductID);
                            InProdCount = AllCount - ReadyCount - OnProdCount;
                        }
                        else
                        {
                            AllCount = AllDecorInfo(GroupType, MegaBatchID, ProductID, 2);
                            ReadyCount = MarketTPSDecorInfo(GroupType, MegaBatchID, ProductID);
                            OnProdCount = TPSOnProdDecorInfo(GroupType, MegaBatchID, ProductID);
                            InProdCount = AllCount - ReadyCount - OnProdCount;
                        }

                        if (Convert.ToInt32(row["MeasureID"]) == 1)
                        {
                            AllDecorProductCount += AllCount;
                            OnProdDecorProductCount += OnProdCount;
                            ReadyDecorProductCount += ReadyCount;
                            InProdDecorProductCount += (AllCount - ReadyCount - OnProdCount);
                            Measure = "шт.";
                        }

                        if (Convert.ToInt32(row["MeasureID"]) == 3)
                        {
                            AllDecorProductCount += AllCount;
                            OnProdDecorProductCount += OnProdCount;
                            ReadyDecorProductCount += ReadyCount;
                            InProdDecorProductCount += (AllCount - ReadyCount - OnProdCount);
                            Measure = "шт.";
                        }

                        if (Convert.ToInt32(row["MeasureID"]) == 2)
                        {
                            //нет параметра "высота"
                            if (row["Height"].ToString() == "-1")
                            {
                                AllDecorProductCount += Convert.ToDecimal(row["Length"]) * AllCount / 1000;
                                OnProdDecorProductCount += Convert.ToDecimal(row["Length"]) * OnProdCount / 1000;
                                ReadyDecorProductCount += Convert.ToDecimal(row["Length"]) * ReadyCount / 1000;
                                InProdDecorProductCount += (AllCount - ReadyCount - OnProdCount);
                            }
                            else
                            {
                                AllDecorProductCount += Convert.ToDecimal(row["Height"]) * AllCount / 1000;
                                OnProdDecorProductCount += Convert.ToDecimal(row["Height"]) * OnProdCount / 1000;
                                ReadyDecorProductCount += Convert.ToDecimal(row["Height"]) * ReadyCount / 1000;
                                InProdDecorProductCount += (AllCount - ReadyCount - OnProdCount);
                            }

                            DecorProductCost += Convert.ToDecimal(row["Cost"]);
                            Measure = "м.п.";
                        }
                    }

                    InProdDecorProductCount = (AllDecorProductCount - ReadyDecorProductCount - OnProdDecorProductCount);
                    if (InProdDecorProductCount < 0)
                        InProdDecorProductCount = 0;
                    DataRow NewRow = DecorProductsSummaryDT.NewRow();
                    NewRow["ProductID"] = Convert.ToInt32(Table.Rows[i]["ProductID"]);
                    NewRow["MeasureID"] = Convert.ToInt32(Table.Rows[i]["MeasureID"]);
                    NewRow["DecorProduct"] = GetProductName(Convert.ToInt32(Table.Rows[i]["ProductID"]));
                    if (AllDecorProductCount < 3)
                        decimals = 1;
                    NewRow["AllCount"] = Decimal.Round(AllDecorProductCount, decimals, MidpointRounding.AwayFromZero);
                    NewRow["OnProdCount"] = Decimal.Round(OnProdDecorProductCount, decimals, MidpointRounding.AwayFromZero);
                    NewRow["InProdCount"] = Decimal.Round(InProdDecorProductCount, decimals, MidpointRounding.AwayFromZero);
                    NewRow["ReadyCount"] = Decimal.Round(ReadyDecorProductCount, decimals, MidpointRounding.AwayFromZero);
                    if (AllDecorProductCount == ReadyDecorProductCount)
                        NewRow["Ready"] = true;
                    else
                        NewRow["Ready"] = false;

                    NewRow["Measure"] = Measure;
                    DecorProductsSummaryDT.Rows.Add(NewRow);

                    Measure = string.Empty;
                    DecorProductCost = 0;
                    AllDecorProductCount = 0;
                }
            }
            DecorProductsSummaryDT.DefaultView.Sort = "DecorProduct";
            DecorProductsSummaryBS.MoveFirst();
        }

        private void SimpleFrontsCount(int GroupType, int MegaBatchID)
        {
            int AllCount = 0;
            int OnProdCount = 0;
            int InProdCount = 0;
            int ReadyCount = 0;

            decimal AllSquare = 0;
            decimal OnProdSquare = 0;
            decimal InProdSquare = 0;
            decimal ReadySquare = 0;

            for (int i = 0; i < SimpleFrontsSummaryDT.Rows.Count; i++)
            {
                int FrontID = Convert.ToInt32(SimpleFrontsSummaryDT.Rows[i]["FrontID"]);
                AllSquare = 0;
                OnProdSquare = 0;
                InProdSquare = 0;
                ReadySquare = 0;

                AllCount = AllFrontsInfo(GroupType, MegaBatchID, Convert.ToInt32(SimpleFrontsSummaryDT.Rows[i]["FrontID"]), Convert.ToInt32(SimpleFrontsSummaryDT.Rows[i]["Width"]), 1, ref AllSquare);
                if (AllCount > 0)
                {
                    ReadyCount = MarketProfilFrontsInfo(GroupType, MegaBatchID, Convert.ToInt32(SimpleFrontsSummaryDT.Rows[i]["FrontID"]), Convert.ToInt32(SimpleFrontsSummaryDT.Rows[i]["Width"]), ref ReadySquare);
                    OnProdCount = ProfilOnProdFrontsInfo(GroupType, MegaBatchID, Convert.ToInt32(SimpleFrontsSummaryDT.Rows[i]["FrontID"]), Convert.ToInt32(SimpleFrontsSummaryDT.Rows[i]["Width"]), ref OnProdSquare);
                    InProdCount = AllCount - ReadyCount - OnProdCount;
                }
                else
                {
                    AllCount = AllFrontsInfo(GroupType, MegaBatchID, Convert.ToInt32(SimpleFrontsSummaryDT.Rows[i]["FrontID"]), Convert.ToInt32(SimpleFrontsSummaryDT.Rows[i]["Width"]), 2, ref AllSquare);
                    ReadyCount = MarketTPSFrontsInfo(GroupType, MegaBatchID, Convert.ToInt32(SimpleFrontsSummaryDT.Rows[i]["FrontID"]), Convert.ToInt32(SimpleFrontsSummaryDT.Rows[i]["Width"]), ref ReadySquare);
                    OnProdCount = TPSOnProdFrontsInfo(GroupType, MegaBatchID, Convert.ToInt32(SimpleFrontsSummaryDT.Rows[i]["FrontID"]), Convert.ToInt32(SimpleFrontsSummaryDT.Rows[i]["Width"]), ref OnProdSquare);
                    InProdCount = AllCount - ReadyCount - OnProdCount;
                }

                AllSquare = Decimal.Round(AllSquare, 2, MidpointRounding.AwayFromZero);
                OnProdSquare = Decimal.Round(OnProdSquare, 2, MidpointRounding.AwayFromZero);
                ReadySquare = Decimal.Round(ReadySquare, 2, MidpointRounding.AwayFromZero);
                InProdSquare = AllSquare - ReadySquare - OnProdSquare;
                InProdSquare = Decimal.Round(InProdSquare, 2, MidpointRounding.AwayFromZero);
                if (InProdSquare < 0)
                    InProdSquare = 0;

                SimpleFrontsSummaryDT.Rows[i]["AllCount"] = AllCount;
                SimpleFrontsSummaryDT.Rows[i]["OnProdCount"] = OnProdCount;
                SimpleFrontsSummaryDT.Rows[i]["InProdCount"] = InProdCount;
                SimpleFrontsSummaryDT.Rows[i]["ReadyCount"] = ReadyCount;
                if (Convert.ToInt32(SimpleFrontsSummaryDT.Rows[i]["Width"]) != -1)
                {
                    SimpleFrontsSummaryDT.Rows[i]["AllSquare"] = AllSquare;
                    SimpleFrontsSummaryDT.Rows[i]["OnProdSquare"] = OnProdSquare;
                    SimpleFrontsSummaryDT.Rows[i]["InProdSquare"] = InProdSquare;
                    SimpleFrontsSummaryDT.Rows[i]["ReadySquare"] = ReadySquare;
                }

                if (AllCount == ReadyCount)
                    SimpleFrontsSummaryDT.Rows[i]["Ready"] = true;
                else
                    SimpleFrontsSummaryDT.Rows[i]["Ready"] = false;
            }
        }

        private void CurvedFrontsCount(int GroupType, int MegaBatchID)
        {
            int AllCount = 0;
            int OnProdCount = 0;
            int InProdCount = 0;
            int ReadyCount = 0;

            decimal AllSquare = 0;
            decimal OnProdSquare = 0;
            //decimal InProdSquare = 0;
            decimal ReadySquare = 0;

            for (int i = 0; i < CurvedFrontsSummaryDT.Rows.Count; i++)
            {
                AllCount = AllFrontsInfo(GroupType, MegaBatchID, Convert.ToInt32(CurvedFrontsSummaryDT.Rows[i]["FrontID"]), Convert.ToInt32(CurvedFrontsSummaryDT.Rows[i]["Width"]), 1, ref AllSquare);
                if (AllCount > 0)
                {
                    ReadyCount = MarketProfilFrontsInfo(GroupType, MegaBatchID, Convert.ToInt32(CurvedFrontsSummaryDT.Rows[i]["FrontID"]), Convert.ToInt32(CurvedFrontsSummaryDT.Rows[i]["Width"]), ref ReadySquare);
                    OnProdCount = ProfilOnProdFrontsInfo(GroupType, MegaBatchID, Convert.ToInt32(CurvedFrontsSummaryDT.Rows[i]["FrontID"]), Convert.ToInt32(CurvedFrontsSummaryDT.Rows[i]["Width"]), ref OnProdSquare);
                    InProdCount = AllCount - ReadyCount - OnProdCount;
                }
                else
                {
                    AllCount = AllFrontsInfo(GroupType, MegaBatchID, Convert.ToInt32(CurvedFrontsSummaryDT.Rows[i]["FrontID"]), Convert.ToInt32(CurvedFrontsSummaryDT.Rows[i]["Width"]), 2, ref AllSquare);
                    ReadyCount = MarketTPSFrontsInfo(GroupType, MegaBatchID, Convert.ToInt32(CurvedFrontsSummaryDT.Rows[i]["FrontID"]), Convert.ToInt32(CurvedFrontsSummaryDT.Rows[i]["Width"]), ref ReadySquare);
                    OnProdCount = TPSOnProdFrontsInfo(GroupType, MegaBatchID, Convert.ToInt32(CurvedFrontsSummaryDT.Rows[i]["FrontID"]), Convert.ToInt32(CurvedFrontsSummaryDT.Rows[i]["Width"]), ref OnProdSquare);
                    InProdCount = AllCount - ReadyCount - OnProdCount;
                }

                CurvedFrontsSummaryDT.Rows[i]["AllCount"] = AllCount;
                CurvedFrontsSummaryDT.Rows[i]["OnProdCount"] = OnProdCount;
                CurvedFrontsSummaryDT.Rows[i]["InProdCount"] = InProdCount;
                CurvedFrontsSummaryDT.Rows[i]["ReadyCount"] = ReadyCount;

                if (AllCount == ReadyCount)
                    CurvedFrontsSummaryDT.Rows[i]["Ready"] = true;
                else
                    CurvedFrontsSummaryDT.Rows[i]["Ready"] = false;
            }
        }

        public void GetSimpleFrontsInfo(
            ref decimal AllSquare,
            ref decimal OnProdSquare,
            ref decimal InProdSquare,
            ref decimal ReadySquare,
            ref int AllCount,
            ref int OnProdCount,
            ref int InProdCount,
            ref int ReadyCount)
        {
            for (int i = 0; i < SimpleFrontsSummaryDT.Rows.Count; i++)
            {
                if (Convert.ToInt32(SimpleFrontsSummaryDT.Rows[i]["Width"]) != -1)
                {
                    if (SimpleFrontsSummaryDT.Rows[i]["AllSquare"] != DBNull.Value)
                        AllSquare += Convert.ToDecimal(SimpleFrontsSummaryDT.Rows[i]["AllSquare"]);
                    if (SimpleFrontsSummaryDT.Rows[i]["OnProdSquare"] != DBNull.Value)
                        OnProdSquare += Convert.ToDecimal(SimpleFrontsSummaryDT.Rows[i]["OnProdSquare"]);
                    if (SimpleFrontsSummaryDT.Rows[i]["InProdSquare"] != DBNull.Value)
                        InProdSquare += Convert.ToDecimal(SimpleFrontsSummaryDT.Rows[i]["InProdSquare"]);
                    if (SimpleFrontsSummaryDT.Rows[i]["ReadySquare"] != DBNull.Value)
                        ReadySquare += Convert.ToDecimal(SimpleFrontsSummaryDT.Rows[i]["ReadySquare"]);
                    if (SimpleFrontsSummaryDT.Rows[i]["AllCount"] != DBNull.Value)
                        AllCount += Convert.ToInt32(SimpleFrontsSummaryDT.Rows[i]["AllCount"]);
                    if (SimpleFrontsSummaryDT.Rows[i]["OnProdCount"] != DBNull.Value)
                        OnProdCount += Convert.ToInt32(SimpleFrontsSummaryDT.Rows[i]["OnProdCount"]);
                    if (SimpleFrontsSummaryDT.Rows[i]["InProdCount"] != DBNull.Value)
                        InProdCount += Convert.ToInt32(SimpleFrontsSummaryDT.Rows[i]["InProdCount"]);
                    if (SimpleFrontsSummaryDT.Rows[i]["ReadyCount"] != DBNull.Value)
                        ReadyCount += Convert.ToInt32(SimpleFrontsSummaryDT.Rows[i]["ReadyCount"]);
                }
            }

            AllSquare = Decimal.Round(AllSquare, 2, MidpointRounding.AwayFromZero);
            OnProdSquare = Decimal.Round(OnProdSquare, 2, MidpointRounding.AwayFromZero);
            InProdSquare = Decimal.Round(InProdSquare, 2, MidpointRounding.AwayFromZero);
            ReadySquare = Decimal.Round(ReadySquare, 2, MidpointRounding.AwayFromZero);
        }

        public void GetCurvedFrontsInfo(
            ref int AllCurvedCount,
            ref int OnProdCurvedCount,
            ref int InProdCurvedCount,
            ref int ReadyCurvedCount)
        {
            for (int i = 0; i < CurvedFrontsSummaryDT.Rows.Count; i++)
            {
                if (Convert.ToInt32(CurvedFrontsSummaryDT.Rows[i]["Width"]) == -1)
                {
                    if (CurvedFrontsSummaryDT.Rows[i]["AllCount"] != DBNull.Value)
                        AllCurvedCount += Convert.ToInt32(CurvedFrontsSummaryDT.Rows[i]["AllCount"]);
                    if (CurvedFrontsSummaryDT.Rows[i]["OnProdCount"] != DBNull.Value)
                        OnProdCurvedCount += Convert.ToInt32(CurvedFrontsSummaryDT.Rows[i]["OnProdCount"]);
                    if (CurvedFrontsSummaryDT.Rows[i]["InProdCount"] != DBNull.Value)
                        InProdCurvedCount += Convert.ToInt32(CurvedFrontsSummaryDT.Rows[i]["InProdCount"]);
                    if (CurvedFrontsSummaryDT.Rows[i]["ReadyCount"] != DBNull.Value)
                        ReadyCurvedCount += Convert.ToInt32(CurvedFrontsSummaryDT.Rows[i]["ReadyCount"]);
                }
            }
        }

        public void GetDecorInfo(
            ref decimal AllPogon,
            ref decimal OnProdPogon,
            ref decimal InProdPogon,
            ref decimal ReadyPogon,
            ref int AllCount,
            ref int OnProdCount,
            ref int InProdCount,
            ref int ReadyCount)
        {
            for (int i = 0; i < DecorProductsSummaryDT.Rows.Count; i++)
            {
                if (Convert.ToInt32(DecorProductsSummaryDT.Rows[i]["MeasureID"]) != 2)
                {
                    AllCount += Convert.ToInt32(DecorProductsSummaryDT.Rows[i]["AllCount"]);
                    OnProdCount += Convert.ToInt32(DecorProductsSummaryDT.Rows[i]["OnProdCount"]);
                    InProdCount += Convert.ToInt32(DecorProductsSummaryDT.Rows[i]["InProdCount"]);
                    ReadyCount += Convert.ToInt32(DecorProductsSummaryDT.Rows[i]["ReadyCount"]);
                }
                else
                {
                    AllPogon += Convert.ToDecimal(DecorProductsSummaryDT.Rows[i]["AllCount"]);
                    OnProdPogon += Convert.ToDecimal(DecorProductsSummaryDT.Rows[i]["OnProdCount"]);
                    InProdPogon += Convert.ToDecimal(DecorProductsSummaryDT.Rows[i]["InProdCount"]);
                    ReadyPogon += Convert.ToDecimal(DecorProductsSummaryDT.Rows[i]["ReadyCount"]);
                }
            }

            AllPogon = Decimal.Round(AllPogon, 2, MidpointRounding.AwayFromZero);
            OnProdPogon = Decimal.Round(OnProdPogon, 2, MidpointRounding.AwayFromZero);
            InProdPogon = Decimal.Round(InProdPogon, 2, MidpointRounding.AwayFromZero);
            ReadyPogon = Decimal.Round(ReadyPogon, 2, MidpointRounding.AwayFromZero);
        }

        public string GetFrontName(int FrontID)
        {
            string FrontName = string.Empty;
            DataRow[] Rows = FrontsDT.Select("FrontID = " + FrontID);
            if (Rows.Count() > 0)
                FrontName = Rows[0]["FrontName"].ToString();
            return FrontName;
        }

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

        public void ClearProductTables()
        {
            SimpleFrontsSummaryDT.Clear();
            CurvedFrontsSummaryDT.Clear();
            DecorProductsSummaryDT.Clear();
        }

        #region ОБЩЕЕ КОЛИЧЕСТВО

        private int AllFrontsInfo(int GroupType, int MegaBatchID, int FrontID, int Width, int FactoryID, ref decimal Square)
        {
            Square = 0;
            int AllCount = 0;

            string GroupTypeFilter = string.Empty;
            string FrontFilter = string.Empty;
            string FrontTypeFilter = string.Empty;
            string BatchFactoryFilter = string.Empty;
            string FactoryFilter = string.Empty;

            if (GroupType != -1)
                GroupTypeFilter = " AND GroupType = " + GroupType;
            if (FrontID != 0)
                FrontFilter = " AND FrontID = " + FrontID;
            if (Width != -1)
                FrontTypeFilter = " AND Width <> -1";
            if (Width == -1)
                FrontTypeFilter = " AND Width = -1";

            if (FactoryID != 0)
            {
                BatchFactoryFilter = " AND BFactoryID = " + FactoryID;
                FactoryFilter = " AND FFactoryID = " + FactoryID;
            }

            DataRow[] Rows = MarkAllFrontsDT.Select("MegaBatchID = " + MegaBatchID + FrontFilter + FrontTypeFilter + FactoryFilter + BatchFactoryFilter + GroupTypeFilter);
            foreach (DataRow Row in Rows)
            {
                AllCount += Convert.ToInt32(Row["AllCount"]);
                Square += Convert.ToDecimal(Row["AllSquare"]);
            }

            return AllCount;
        }

        private int AllDecorInfo(int GroupType, int MegaBatchID, int ProductID, int FactoryID)
        {
            int AllCount = 0;

            string GroupTypeFilter = string.Empty;
            string ProductFilter = string.Empty;
            string BatchFactoryFilter = string.Empty;
            string FactoryFilter = string.Empty;

            if (GroupType != -1)
                GroupTypeFilter = " AND GroupType = " + GroupType;
            if (ProductID != 0)
                ProductFilter = " AND ProductID = " + ProductID;
            if (FactoryID != 0)
            {
                BatchFactoryFilter = " AND BFactoryID = " + FactoryID;
                FactoryFilter = " AND FFactoryID = " + FactoryID;
            }

            DataRow[] Rows = MarkAllDecorDT.Select("MegaBatchID = " + MegaBatchID + ProductFilter + FactoryFilter + BatchFactoryFilter + GroupTypeFilter);
            foreach (DataRow Row in Rows)
            {
                AllCount += Convert.ToInt32(Row["AllCount"]);
            }

            return AllCount;
        }

        #endregion ОБЩЕЕ КОЛИЧЕСТВО

        #region НА ПРОИЗВОДСТВЕ

        private int ProfilOnProdFrontsInfo(int GroupType, int MegaBatchID, int FrontID, int Width, ref decimal Square)
        {
            Square = 0;
            int AllCount = 0;
            string GroupTypeFilter = string.Empty;
            string FrontFilter = string.Empty;
            string FrontTypeFilter = string.Empty;

            if (GroupType != -1)
                GroupTypeFilter = " AND GroupType = " + GroupType;
            if (FrontID != 0)
                FrontFilter = " AND FrontID = " + FrontID;
            if (Width != -1)
                FrontTypeFilter = " AND Width <> -1";
            if (Width == -1)
                FrontTypeFilter = " AND Width = -1";

            DataRow[] Rows = MarkProfilOnProdFrontsDT.Select("MegaBatchID = " + MegaBatchID + FrontFilter + FrontTypeFilter + GroupTypeFilter);
            foreach (DataRow Row in Rows)
            {
                AllCount += Convert.ToInt32(Row["AllCount"]);
                Square += Convert.ToDecimal(Row["AllSquare"]);
            }

            return AllCount;
        }

        private int TPSOnProdFrontsInfo(int GroupType, int MegaBatchID, int FrontID, int Width, ref decimal Square)
        {
            Square = 0;
            int AllCount = 0;
            string GroupTypeFilter = string.Empty;
            string FrontFilter = string.Empty;
            string FrontTypeFilter = string.Empty;

            if (GroupType != -1)
                GroupTypeFilter = " AND GroupType = " + GroupType;
            if (FrontID != 0)
                FrontFilter = " AND FrontID = " + FrontID;
            if (Width != -1)
                FrontTypeFilter = " AND Width <> -1";
            if (Width == -1)
                FrontTypeFilter = " AND Width = -1";

            DataRow[] Rows = MarkTPSOnProdFrontsDT.Select("MegaBatchID = " + MegaBatchID + FrontFilter + FrontTypeFilter + GroupTypeFilter);
            foreach (DataRow Row in Rows)
            {
                AllCount += Convert.ToInt32(Row["AllCount"]);
                Square += Convert.ToDecimal(Row["AllSquare"]);
            }

            return AllCount;
        }

        private int ProfilOnProdDecorInfo(int GroupType, int MegaBatchID, int ProductID)
        {
            int AllCount = 0;
            string GroupTypeFilter = string.Empty;
            string ProductFilter = string.Empty;

            if (GroupType != -1)
                GroupTypeFilter = " AND GroupType = " + GroupType;
            if (ProductID != 0)
                ProductFilter = " AND ProductID = " + ProductID;

            DataRow[] Rows = MarkProfilOnProdDecorDT.Select("MegaBatchID = " + MegaBatchID + ProductFilter + GroupTypeFilter);
            foreach (DataRow Row in Rows)
            {
                AllCount += Convert.ToInt32(Row["AllCount"]);
            }

            return AllCount;
        }

        private int TPSOnProdDecorInfo(int GroupType, int MegaBatchID, int ProductID)
        {
            int AllCount = 0;
            string GroupTypeFilter = string.Empty;
            string ProductFilter = string.Empty;

            if (GroupType != -1)
                GroupTypeFilter = " AND GroupType = " + GroupType;
            if (ProductID != 0)
                ProductFilter = " AND ProductID = " + ProductID;

            DataRow[] Rows = MarkTPSOnProdDecorDT.Select("MegaBatchID = " + MegaBatchID + ProductFilter + GroupTypeFilter);
            foreach (DataRow Row in Rows)
            {
                AllCount += Convert.ToInt32(Row["AllCount"]);
            }

            return AllCount;
        }

        #endregion НА ПРОИЗВОДСТВЕ

        #region ГОТОВО

        private int MarketProfilFrontsInfo(int GroupType, int MegaBatchID, int FrontID, int Width, ref decimal Square)
        {
            Square = 0;
            int AllCount = 0;

            string GroupTypeFilter = string.Empty;
            string FrontFilter = string.Empty;
            string FrontTypeFilter = string.Empty;
            string StatusFilter = string.Empty;

            if (GroupType != -1)
                GroupTypeFilter = " AND GroupType = " + GroupType;
            if (FrontID != 0)
                FrontFilter = " AND FrontID = " + FrontID;
            if (Width != -1)
                FrontTypeFilter = " AND Width <> -1";
            if (Width == -1)
                FrontTypeFilter = " AND Width = -1";

            //if (Status == ProductionStatus.InProduction)
            //    StatusFilter = " AND PackageStatusID = 0";
            //if (Status == ProductionStatus.Ready)
            //    StatusFilter = " AND (PackageStatusID = 1 OR PackageStatusID = 2 OR PackageStatusID = 3)";

            DataRow[] Rows = MarkProfilReadyFrontsDT.Select("MegaBatchID = " + MegaBatchID + FrontFilter + FrontTypeFilter + StatusFilter + GroupTypeFilter);
            foreach (DataRow Row in Rows)
            {
                AllCount += Convert.ToInt32(Row["AllCount"]);
                Square += Convert.ToDecimal(Row["AllSquare"]);
            }

            return AllCount;
        }

        private int MarketProfilDecorInfo(int GroupType, int MegaBatchID, int ProductID)
        {
            int AllCount = 0;

            string GroupTypeFilter = string.Empty;
            string ProductFilter = string.Empty;
            string StatusFilter = string.Empty;

            if (GroupType != -1)
                GroupTypeFilter = " AND GroupType = " + GroupType;
            if (ProductID != 0)
                ProductFilter = " AND ProductID = " + ProductID;

            //if (Status == ProductionStatus.InProduction)
            //    StatusFilter = " AND PackageStatusID = 0";
            //if (Status == ProductionStatus.Ready)
            //    StatusFilter = " AND (PackageStatusID = 1 OR PackageStatusID = 2 OR PackageStatusID = 3)";

            DataRow[] Rows = MarkProfilReadyDecorDT.Select("MegaBatchID = " + MegaBatchID + ProductFilter + StatusFilter + GroupTypeFilter);
            foreach (DataRow Row in Rows)
            {
                AllCount += Convert.ToInt32(Row["AllCount"]);
            }

            return AllCount;
        }

        private int MarketTPSFrontsInfo(int GroupType, int MegaBatchID, int FrontID, int Width, ref decimal Square)
        {
            Square = 0;
            int AllCount = 0;

            string GroupTypeFilter = string.Empty;
            string FrontFilter = string.Empty;
            string FrontTypeFilter = string.Empty;
            string StatusFilter = string.Empty;

            if (GroupType != -1)
                GroupTypeFilter = " AND GroupType = " + GroupType;
            if (FrontID != 0)
                FrontFilter = " AND FrontID = " + FrontID;
            if (Width != -1)
                FrontTypeFilter = " AND Width <> -1";
            if (Width == -1)
                FrontTypeFilter = " AND Width = -1";

            DataRow[] Rows = MarkTPSReadyFrontsDT.Select("MegaBatchID = " + MegaBatchID + FrontFilter + FrontTypeFilter + StatusFilter + GroupTypeFilter);
            foreach (DataRow Row in Rows)
            {
                AllCount += Convert.ToInt32(Row["AllCount"]);
                Square += Convert.ToDecimal(Row["AllSquare"]);
            }

            return AllCount;
        }

        private int MarketTPSDecorInfo(int GroupType, int MegaBatchID, int ProductID)
        {
            int AllCount = 0;

            string GroupTypeFilter = string.Empty;
            string ProductFilter = string.Empty;
            string StatusFilter = string.Empty;

            if (GroupType != -1)
                GroupTypeFilter = " AND GroupType = " + GroupType;
            if (ProductID != 0)
                ProductFilter = " AND ProductID = " + ProductID;

            DataRow[] Rows = MarkTPSReadyDecorDT.Select("MegaBatchID = " + MegaBatchID + ProductFilter + StatusFilter + GroupTypeFilter);
            foreach (DataRow Row in Rows)
            {
                AllCount += Convert.ToInt32(Row["AllCount"]);
            }

            return AllCount;
        }

        #endregion ГОТОВО

        #region участки ТПС

        public void SectorsInfo(int GroupType, int MegaBatchID, ref decimal FilenkaPerc, ref decimal TrimmingPerc, ref decimal AssemblyPerc, ref decimal DeyingPerc)
        {
            DataTable DT = new DataTable();
            if (GroupType == 0)
            {
                string SelectCommand = @"SELECT WorkAssignmentID, Square, (SELECT SUM(Square) AS Expr1 FROM infiniu2_marketingorders.dbo.WorkAssignments
                WHERE WorkAssignmentID IN (SELECT TPSWorkAssignmentID FROM infiniu2_zovorders.dbo.Batch
                WHERE MegaBatchID = " + MegaBatchID + @")) AS AllSquare, FilenkaDateTime, TrimmingDateTime, AssemblyDateTime, DeyingDateTime FROM infiniu2_marketingorders.dbo.WorkAssignments
                WHERE WorkAssignmentID IN (SELECT TPSWorkAssignmentID FROM infiniu2_zovorders.dbo.Batch
                WHERE MegaBatchID = " + MegaBatchID + ")";
                using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.ZOVOrdersConnectionString))
                {
                    if (DA.Fill(DT) > 0)
                    {
                        decimal Square1 = 0;
                        decimal Square2 = 0;
                        decimal Square3 = 0;
                        decimal Square4 = 0;
                        for (int i = 0; i < DT.Rows.Count; i++)
                        {
                            if (DT.Rows[i]["FilenkaDateTime"] != DBNull.Value && DT.Rows[i]["Square"] != DBNull.Value)
                                Square1 += Convert.ToDecimal(DT.Rows[i]["Square"]);
                            if (DT.Rows[i]["TrimmingDateTime"] != DBNull.Value && DT.Rows[i]["Square"] != DBNull.Value)
                                Square2 += Convert.ToDecimal(DT.Rows[i]["Square"]);
                            if (DT.Rows[i]["AssemblyDateTime"] != DBNull.Value && DT.Rows[i]["Square"] != DBNull.Value)
                                Square3 += Convert.ToDecimal(DT.Rows[i]["Square"]);
                            if (DT.Rows[i]["DeyingDateTime"] != DBNull.Value && DT.Rows[i]["Square"] != DBNull.Value)
                                Square4 += Convert.ToDecimal(DT.Rows[i]["Square"]);
                        }
                        decimal AllSquare = -1;
                        if (DT.Rows[0]["AllSquare"] != DBNull.Value)
                            AllSquare = Convert.ToDecimal(DT.Rows[0]["AllSquare"]);
                        FilenkaPerc = Square1 / AllSquare;
                        TrimmingPerc = Square2 / AllSquare;
                        AssemblyPerc = Square3 / AllSquare;
                        DeyingPerc = Square4 / AllSquare;
                    }
                }
            }
            if (GroupType == 1)
            {
                string SelectCommand = @"SELECT WorkAssignmentID, Square, (SELECT SUM(Square) AS Expr1 FROM WorkAssignments
                WHERE WorkAssignmentID IN (SELECT TPSWorkAssignmentID FROM Batch
                WHERE MegaBatchID = " + MegaBatchID + @")) AS AllSquare, FilenkaDateTime, TrimmingDateTime, AssemblyDateTime, DeyingDateTime FROM WorkAssignments
                WHERE WorkAssignmentID IN (SELECT TPSWorkAssignmentID FROM Batch
                WHERE MegaBatchID = " + MegaBatchID + ")";
                using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
                {
                    if (DA.Fill(DT) > 0)
                    {
                        decimal Square1 = 0;
                        decimal Square2 = 0;
                        decimal Square3 = 0;
                        decimal Square4 = 0;
                        for (int i = 0; i < DT.Rows.Count; i++)
                        {
                            if (DT.Rows[i]["FilenkaDateTime"] != DBNull.Value && DT.Rows[i]["Square"] != DBNull.Value)
                                Square1 += Convert.ToDecimal(DT.Rows[i]["Square"]);
                            if (DT.Rows[i]["TrimmingDateTime"] != DBNull.Value && DT.Rows[i]["Square"] != DBNull.Value)
                                Square2 += Convert.ToDecimal(DT.Rows[i]["Square"]);
                            if (DT.Rows[i]["AssemblyDateTime"] != DBNull.Value && DT.Rows[i]["Square"] != DBNull.Value)
                                Square3 += Convert.ToDecimal(DT.Rows[i]["Square"]);
                            if (DT.Rows[i]["DeyingDateTime"] != DBNull.Value && DT.Rows[i]["Square"] != DBNull.Value)
                                Square4 += Convert.ToDecimal(DT.Rows[i]["Square"]);
                        }
                        decimal AllSquare = -1;
                        if (DT.Rows[0]["AllSquare"] != DBNull.Value)
                            AllSquare = Convert.ToDecimal(DT.Rows[0]["AllSquare"]);
                        FilenkaPerc = Square1 / AllSquare;
                        TrimmingPerc = Square2 / AllSquare;
                        AssemblyPerc = Square3 / AllSquare;
                        DeyingPerc = Square4 / AllSquare;
                    }
                }
            }
        }

        #endregion участки ТПС

        private string EnterInProdDate(int GroupType, int MegaBatchID, int FactoryID)
        {
            string ConnectionString = ConnectionStrings.MarketingOrdersConnectionString;
            if (GroupType == 0)
                ConnectionString = ConnectionStrings.ZOVOrdersConnectionString;

            string ProductionDate = string.Empty;

            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT MIN(ProfilProductionDate), MIN(TPSProductionDate) FROM MainOrders" +
                " WHERE MainOrderID IN (SELECT MainOrderID FROM BatchDetails" +
                " WHERE BatchID IN (SELECT BatchID FROM Batch" +
                " WHERE MegaBatchID = " + MegaBatchID + "))",
                ConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    DA.Fill(DT);
                    if (FactoryID == 1)
                    {
                        if (DT.Rows.Count > 0 && DT.Rows[0]["Column1"] != DBNull.Value)
                            ProductionDate = Convert.ToDateTime(DT.Rows[0]["Column1"]).ToString("dd.MM.yyyy HH:mm");
                    }
                    if (FactoryID == 2)
                    {
                        if (DT.Rows.Count > 0 && DT.Rows[0]["Column2"] != DBNull.Value)
                            ProductionDate = Convert.ToDateTime(DT.Rows[0]["Column2"]).ToString("dd.MM.yyyy HH:mm");
                    }
                }
            }
            return ProductionDate;
        }

        private string ReadyDate(int GroupType, int MegaBatchID, int FactoryID)
        {
            string ConnectionString = ConnectionStrings.MarketingOrdersConnectionString;
            if (GroupType == 0)
                ConnectionString = ConnectionStrings.ZOVOrdersConnectionString;

            string PackingDate = string.Empty;

            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT MAX(PackingDateTime) AS PackingDate FROM Packages" +
                " WHERE FactoryID = " + FactoryID + " AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails" +
                " WHERE BatchID IN (SELECT BatchID FROM Batch" +
                " WHERE MegaBatchID = " + MegaBatchID + "))",
                ConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    DA.Fill(DT);

                    if (DT.Rows.Count > 0 && DT.Rows[0]["PackingDate"] != DBNull.Value)
                    {
                        PackingDate = Convert.ToDateTime(DT.Rows[0]["PackingDate"]).ToString("dd.MM.yyyy HH:mm");
                    }
                }
            }
            return PackingDate;
        }
    }
}