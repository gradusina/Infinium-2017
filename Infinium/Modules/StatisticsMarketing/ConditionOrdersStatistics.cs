using System;
using System.Collections;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Windows.Forms;

namespace Infinium.Modules.StatisticsMarketing
{
    public class ConditionOrdersStatistics
    {
        private DataTable ClientGroupsDT = null;
        private DataTable MarketingClientsDT = null;
        private DataTable ZOVDispatchDT = null;

        private DataTable FrontsOrdersDT = null;
        private DataTable DecorOrdersDT = null;

        public DataTable FrontsSummaryDT = null;
        public DataTable DecorProductsSummaryDT = null;
        public DataTable DecorItemsSummaryDT = null;
        private DataTable DecorConfigDT = null;

        private DataTable FrontsDT = null;
        private DataTable DecorProductsDT = null;
        private DataTable DecorItemsDT = null;

        public BindingSource ClientGroupsBS = null;

        public BindingSource FrontsSummaryBS = null;
        public BindingSource DecorProductsSummaryBS = null;
        public BindingSource DecorItemsSummaryBS = null;

        public ConditionOrdersStatistics()
        {
            Initialize();
        }

        private void Create()
        {
            ClientGroupsDT = new DataTable();
            MarketingClientsDT = new DataTable();
            ZOVDispatchDT = new DataTable();

            DecorItemsDT = new DataTable();
            FrontsDT = new DataTable();
            DecorProductsDT = new DataTable();
            DecorConfigDT = new DataTable();

            FrontsOrdersDT = new DataTable();
            DecorOrdersDT = new DataTable();

            FrontsSummaryDT = new DataTable();
            FrontsSummaryDT.Columns.Add(new DataColumn(("Front"), System.Type.GetType("System.String")));
            FrontsSummaryDT.Columns.Add(new DataColumn(("FrontID"), System.Type.GetType("System.Int32")));
            FrontsSummaryDT.Columns.Add(new DataColumn(("Square"), System.Type.GetType("System.Decimal")));
            FrontsSummaryDT.Columns.Add(new DataColumn(("Cost"), Type.GetType("System.Decimal")));
            FrontsSummaryDT.Columns.Add(new DataColumn(("Width"), System.Type.GetType("System.Int32")));
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
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM ClientGroups",
                ConnectionStrings.MarketingReferenceConnectionString))
            {
                DA.Fill(ClientGroupsDT);
            }
            ClientGroupsDT.Columns.Add(new DataColumn("Check", Type.GetType("System.Boolean")));
            for (int i = 0; i < ClientGroupsDT.Rows.Count; i++)
            {
                if (i == 1)
                    ClientGroupsDT.Rows[i]["Check"] = false;
                else
                    ClientGroupsDT.Rows[i]["Check"] = true;
            }

            using (SqlDataAdapter DA = new SqlDataAdapter(
            @"SELECT MegaOrders.ClientID, infiniu2_marketingreference.dbo.Clients.ClientName, MegaOrders.OrderNumber, MegaOrders.MegaOrderID
            FROM MegaOrders
            INNER JOIN infiniu2_marketingreference.dbo.Clients ON MegaOrders.ClientID = infiniu2_marketingreference.dbo.Clients.ClientID
            WHERE MegaOrders.MegaOrderID IN (SELECT MegaOrderID FROM MainOrders WHERE MainOrderID IN (SELECT MainOrderID FROM Packages WHERE PackageStatusID = 3))
            ORDER BY infiniu2_marketingreference.dbo.Clients.ClientName, MegaOrders.OrderNumber, MegaOrders.MegaOrderID",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(MarketingClientsDT);
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
                DA.Fill(ZOVDispatchDT);
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

            sw.Stop();
            double G = sw.Elapsed.TotalMilliseconds;

            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT TOP 0 FrontID, PatinaID, ColorID, InsetTypeID," +
                " InsetColorID, Height, Width, Count, Square, Cost FROM FrontsOrders",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
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

        private void Binding()
        {
            ClientGroupsBS = new BindingSource()
            {
                DataSource = ClientGroupsDT
            };
            FrontsSummaryBS = new BindingSource()
            {
                DataSource = FrontsSummaryDT
            };
            DecorProductsSummaryBS = new BindingSource()
            {
                DataSource = DecorProductsSummaryDT
            };
            DecorItemsSummaryBS = new BindingSource()
            {
                DataSource = DecorItemsSummaryDT
            };
        }

        public void Initialize()
        {
            Create();
            Fill();
            Binding();
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

        public void GetOnAgreementOrders(DateTime FixingDate, int FactoryID)
        {
            ArrayList ClientGroups = IsMarketingSelect();

            DataTable DT = new DataTable();

            if (ClientGroups.Count > 0)
            {
                DT.Clear();
                DT = MarketingOnAgreementFronts(ClientGroups, FixingDate, FactoryID);
                foreach (DataRow item in DT.Rows)
                    FrontsOrdersDT.ImportRow(item);

                DT.Clear();
                DT = MarketingOnAgreementDecor(ClientGroups, FixingDate, FactoryID);
                foreach (DataRow item in DT.Rows)
                    DecorOrdersDT.ImportRow(item);
            }

            if (IsZOVSelect())
            {
                DT.Clear();
                DT = ZOVOnAgreementFronts(FixingDate, FactoryID);
                foreach (DataRow item in DT.Rows)
                    FrontsOrdersDT.ImportRow(item);

                DT.Clear();
                DT = ZOVOnAgreementDecor(FixingDate, FactoryID);
                foreach (DataRow item in DT.Rows)
                    DecorOrdersDT.ImportRow(item);
            }

            DT.Dispose();

            GetFronts();
            GetDecorProducts();
            GetDecorItems();
        }

        public void GetAgreedOrders(DateTime FixingDate, int FactoryID)
        {
            ArrayList ClientGroups = IsMarketingSelect();

            DataTable DT = new DataTable();

            if (ClientGroups.Count > 0)
            {
                DT.Clear();
                DT = MarketingAgreedFronts(ClientGroups, FixingDate, FactoryID);
                foreach (DataRow item in DT.Rows)
                    FrontsOrdersDT.ImportRow(item);

                DT.Clear();
                DT = MarketingAgreedDecor(ClientGroups, FixingDate, FactoryID);
                foreach (DataRow item in DT.Rows)
                    DecorOrdersDT.ImportRow(item);
            }

            if (IsZOVSelect())
            {
                DT.Clear();
                DT = ZOVAgreedFronts(FixingDate, FactoryID);
                foreach (DataRow item in DT.Rows)
                    FrontsOrdersDT.ImportRow(item);

                DT.Clear();
                DT = ZOVAgreedDecor(FixingDate, FactoryID);
                foreach (DataRow item in DT.Rows)
                    DecorOrdersDT.ImportRow(item);
            }

            DT.Dispose();

            GetFronts();
            GetDecorProducts();
            GetDecorItems();
        }

        public void GetOnProductionOrders(DateTime FixingDate, int FactoryID)
        {
            ArrayList ClientGroups = IsMarketingSelect();

            DataTable DT = new DataTable();

            if (ClientGroups.Count > 0)
            {
                DT.Clear();
                DT = MarketingOnProductionFronts(ClientGroups, FixingDate, FactoryID);
                foreach (DataRow item in DT.Rows)
                    FrontsOrdersDT.ImportRow(item);

                DT.Clear();
                DT = MarketingOnProductionDecor(ClientGroups, FixingDate, FactoryID);
                foreach (DataRow item in DT.Rows)
                    DecorOrdersDT.ImportRow(item);
            }

            if (IsZOVSelect())
            {
                DT.Clear();
                DT = ZOVOnProductionFronts(FixingDate, FactoryID);
                foreach (DataRow item in DT.Rows)
                    FrontsOrdersDT.ImportRow(item);

                DT.Clear();
                DT = ZOVOnProductionDecor(FixingDate, FactoryID);
                foreach (DataRow item in DT.Rows)
                    DecorOrdersDT.ImportRow(item);
            }

            DT.Dispose();

            GetFronts();
            GetDecorProducts();
            GetDecorItems();
        }

        public void GetOutProductionOrders(DateTime FixingDate, int FactoryID)
        {
            ArrayList ClientGroups = IsMarketingSelect();

            DataTable DT = new DataTable();

            if (ClientGroups.Count > 0)
            {
                DT.Clear();
                DT = MarketingOutProductionFronts(ClientGroups, FixingDate, FactoryID);
                foreach (DataRow item in DT.Rows)
                    FrontsOrdersDT.ImportRow(item);

                DT.Clear();
                DT = MarketingOutProductionDecor(ClientGroups, FixingDate, FactoryID);
                foreach (DataRow item in DT.Rows)
                    DecorOrdersDT.ImportRow(item);
            }
            if (IsZOVSelect())
            {
                DT.Clear();
                DT = ZOVOutProductionFronts(FixingDate, FactoryID);
                foreach (DataRow item in DT.Rows)
                    FrontsOrdersDT.ImportRow(item);

                DT.Clear();
                DT = ZOVOutProductionDecor(FixingDate, FactoryID);
                foreach (DataRow item in DT.Rows)
                    DecorOrdersDT.ImportRow(item);
            }

            DT.Dispose();

            GetFronts();
            GetDecorProducts();
            GetDecorItems();
        }
        
        public void GetInProductionOrders(DateTime FixingDate, int FactoryID)
        {
            ArrayList ClientGroups = IsMarketingSelect();

            DataTable DT = new DataTable();

            if (ClientGroups.Count > 0)
            {
                DT.Clear();
                DT = MarketingInProductionFronts(ClientGroups, FixingDate, FactoryID);
                foreach (DataRow item in DT.Rows)
                    FrontsOrdersDT.ImportRow(item);

                DT.Clear();
                DT = MarketingInProductionDecor(ClientGroups, FixingDate, FactoryID);
                foreach (DataRow item in DT.Rows)
                    DecorOrdersDT.ImportRow(item);
            }
            if (IsZOVSelect())
            {
                DT.Clear();
                DT = ZOVInProductionFronts(FixingDate, FactoryID);
                foreach (DataRow item in DT.Rows)
                    FrontsOrdersDT.ImportRow(item);

                DT.Clear();
                DT = ZOVInProductionDecor(FixingDate, FactoryID);
                foreach (DataRow item in DT.Rows)
                    DecorOrdersDT.ImportRow(item);
            }

            DT.Dispose();

            GetFronts();
            GetDecorProducts();
            GetDecorItems();
        }

        #region Marketing

        /// <summary>
        /// фасады на согласовании
        /// </summary>
        /// <param name="FactoryID"></param>
        /// <returns></returns>
        public DataTable MarketingOnAgreementFronts(ArrayList ClientGroups, DateTime FixingDate, int FactoryID)
        {
            string AgreementFilter = " AND AgreementStatusID <> 2";
            string ClientFilter = string.Empty;

            string MainOrdersFilter = string.Empty;
            string FrontsFactoryFilter = string.Empty;

            string SelectCommand = string.Empty;

            DataTable DT = new DataTable();

            if (FactoryID == -1)
            {
                FrontsFactoryFilter = " NewFrontsOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = @" WHERE MainOrderID = -1";
            }
            if (FactoryID == 0)
            {
                MainOrdersFilter = " WHERE (DocDateTime < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") +
                      "' AND ( ( (ProfilOnProductionDate IS NULL AND ProfilProductionStatusID = 1 AND ProfilStorageStatusID = 1 AND ProfilDispatchStatusID = 1) OR (ProfilOnProductionDate > '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "'))" +
                    " OR " + " ( (TPSOnProductionDate IS NULL AND TPSProductionStatusID = 1 AND TPSStorageStatusID = 1 AND TPSDispatchStatusID = 1) OR (TPSOnProductionDate > '" +
                    FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "') ) ))";
            }
            if (FactoryID == 1)
            {
                FrontsFactoryFilter = " NewFrontsOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = " WHERE (NewMainOrders.FactoryID <> 2 AND DocDateTime < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") +
                    "' AND ((ProfilOnProductionDate IS NULL AND ProfilProductionStatusID = 1 AND ProfilStorageStatusID = 1 AND ProfilDispatchStatusID = 1) OR (ProfilOnProductionDate > '" +
                    FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "')))";
            }
            if (FactoryID == 2)
            {
                FrontsFactoryFilter = " NewFrontsOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = " WHERE (NewMainOrders.FactoryID <> 1 AND DocDateTime < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") +
                    "' AND ((TPSOnProductionDate IS NULL AND TPSProductionStatusID = 1 AND TPSStorageStatusID = 1 AND TPSDispatchStatusID = 1) OR (TPSOnProductionDate > '" +
                    FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "')))";
            }

            if (ClientGroups.Count > 0)
            {
                ClientFilter = " WHERE ClientID IN" +
                    " (SELECT ClientID FROM infiniu2_marketingreference.dbo.Clients" +
                    " WHERE ClientGroupID IN (" + string.Join(",", ClientGroups.OfType<Int32>().ToArray()) + "))";
            }
            else
                ClientFilter = " WHERE ClientID = -1";

            SelectCommand = "SELECT NewFrontsOrders.FrontID, NewFrontsOrders.PatinaID, NewFrontsOrders.ColorID, NewFrontsOrders.InsetTypeID," +
                " NewFrontsOrders.InsetColorID, NewFrontsOrders.TechnoColorID, NewFrontsOrders.TechnoInsetTypeID, NewFrontsOrders.TechnoInsetColorID, NewFrontsOrders.Height, NewFrontsOrders.Width, NewFrontsOrders.Count, NewFrontsOrders.Square, NewFrontsOrders.Cost, MeasureID, ClientID FROM NewFrontsOrders" +
                " INNER JOIN NewMainOrders ON NewFrontsOrders.MainOrderID = NewMainOrders.MainOrderID" +
                " INNER JOIN NewMegaOrders ON NewMainOrders.MegaOrderID = NewMegaOrders.MegaOrderID" +
                " INNER JOIN infiniu2_catalog.dbo.FrontsConfig" +
                " ON NewFrontsOrders.FrontConfigID=infiniu2_catalog.dbo.FrontsConfig.FrontConfigID" +
                " WHERE " + FrontsFactoryFilter + " NewFrontsOrders.MainOrderID IN (SELECT MainOrderID FROM NewMainOrders " + MainOrdersFilter +
                " AND MegaOrderID IN" +
                " (SELECT MegaOrderID FROM NewMegaOrders " + ClientFilter + AgreementFilter + "))";

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(DT);
            }

            return DT;
        }

        /// <summary>
        /// фасады на согласовании
        /// </summary>
        /// <param name="FactoryID"></param>
        /// <returns></returns>
        public DataTable MarketingOnAgreementDecor(ArrayList ClientGroups, DateTime FixingDate, int FactoryID)
        {
            string AgreementFilter = " AND AgreementStatusID <> 2";
            string ClientFilter = string.Empty;

            string MainOrdersFilter = string.Empty;
            string DecorFactoryFilter = string.Empty;

            string SelectCommand = string.Empty;

            DataTable DT = new DataTable();

            if (FactoryID == -1)
            {
                DecorFactoryFilter = " NewDecorOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = " WHERE MainOrderID = -1";
            }
            if (FactoryID == 0)
            {
                MainOrdersFilter = " WHERE (DocDateTime < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") +
                      "' AND ( ( (ProfilOnProductionDate IS NULL AND ProfilProductionStatusID = 1 AND ProfilStorageStatusID = 1 AND ProfilDispatchStatusID = 1) OR (ProfilOnProductionDate > '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "'))" +
                    " OR " + " ( (TPSOnProductionDate IS NULL AND TPSProductionStatusID = 1 AND TPSStorageStatusID = 1 AND TPSDispatchStatusID = 1) OR (TPSOnProductionDate > '" +
                    FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "') ) ))";
            }
            if (FactoryID == 1)
            {
                DecorFactoryFilter = " NewDecorOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = " WHERE (NewMainOrders.FactoryID <> 2 AND DocDateTime < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") +
                    "' AND ((ProfilOnProductionDate IS NULL AND ProfilProductionStatusID = 1 AND ProfilStorageStatusID = 1 AND ProfilDispatchStatusID = 1) OR (ProfilOnProductionDate > '" +
                    FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "')))";
            }
            if (FactoryID == 2)
            {
                DecorFactoryFilter = " NewDecorOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = " WHERE (NewMainOrders.FactoryID <> 1 AND DocDateTime < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") +
                    "' AND ((TPSOnProductionDate IS NULL AND TPSProductionStatusID = 1 AND TPSStorageStatusID = 1 AND TPSDispatchStatusID = 1) OR (TPSOnProductionDate > '" +
                    FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "')))";
            }

            if (ClientGroups.Count > 0)
            {
                ClientFilter = " WHERE ClientID IN" +
                    " (SELECT ClientID FROM infiniu2_marketingreference.dbo.Clients" +
                    " WHERE ClientGroupID IN (" + string.Join(",", ClientGroups.OfType<Int32>().ToArray()) + "))";
            }
            else
                ClientFilter = " WHERE ClientID = -1";

            //decor
            SelectCommand = "SELECT NewDecorOrders.ProductID, NewDecorOrders.DecorID, NewDecorOrders.ColorID,NewDecorOrders.PatinaID," +
                " NewDecorOrders.Height, NewDecorOrders.Length, NewDecorOrders.Width, NewDecorOrders.Count, NewDecorOrders.Cost, NewDecorOrders.DecorConfigID, " +
                " MeasureID, ClientID FROM NewDecorOrders" +
                " INNER JOIN NewMainOrders ON NewDecorOrders.MainOrderID = NewMainOrders.MainOrderID" +
                " INNER JOIN NewMegaOrders ON NewMainOrders.MegaOrderID = NewMegaOrders.MegaOrderID" +
                " INNER JOIN infiniu2_catalog.dbo.DecorConfig" +
                " ON NewDecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                " WHERE " + DecorFactoryFilter + " NewDecorOrders.MainOrderID IN (SELECT MainOrderID FROM NewMainOrders " + MainOrdersFilter +
                " AND MegaOrderID IN" +
                " (SELECT MegaOrderID FROM NewMegaOrders " + ClientFilter + AgreementFilter + "))";

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(DT);
            }

            return DT;
        }

        /// <summary>
        /// фасады на согласовании
        /// </summary>
        /// <param name="FactoryID"></param>
        /// <returns></returns>
        public DataTable MarketingAgreedFronts(ArrayList ClientGroups, DateTime FixingDate, int FactoryID)
        {
            string AgreementFilter = " AND AgreementStatusID = 2";
            string ClientFilter = string.Empty;

            string MainOrdersFilter = string.Empty;
            string FrontsFactoryFilter = string.Empty;

            string SelectCommand = string.Empty;

            DataTable DT = new DataTable();

            if (FactoryID == -1)
            {
                FrontsFactoryFilter = " FrontsOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = @" WHERE MainOrderID = -1";
            }
            if (FactoryID == 0)
            {
                MainOrdersFilter = " WHERE (DocDateTime < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") +
                      "' AND ( ( (ProfilOnProductionDate IS NULL AND ProfilProductionStatusID = 1 AND ProfilStorageStatusID = 1 AND ProfilDispatchStatusID = 1) OR (ProfilOnProductionDate > '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "'))" +
                    " OR " + " ( (TPSOnProductionDate IS NULL AND TPSProductionStatusID = 1 AND TPSStorageStatusID = 1 AND TPSDispatchStatusID = 1) OR (TPSOnProductionDate > '" +
                    FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "') ) ))";
            }
            if (FactoryID == 1)
            {
                FrontsFactoryFilter = " FrontsOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = " WHERE (MainOrders.FactoryID <> 2 AND DocDateTime < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") +
                    "' AND ((ProfilOnProductionDate IS NULL AND ProfilProductionStatusID = 1 AND ProfilStorageStatusID = 1 AND ProfilDispatchStatusID = 1) OR (ProfilOnProductionDate > '" +
                    FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "')))";
            }
            if (FactoryID == 2)
            {
                FrontsFactoryFilter = " FrontsOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = " WHERE (MainOrders.FactoryID <> 1 AND DocDateTime < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") +
                    "' AND ((TPSOnProductionDate IS NULL AND TPSProductionStatusID = 1 AND TPSStorageStatusID = 1 AND TPSDispatchStatusID = 1) OR (TPSOnProductionDate > '" +
                    FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "')))";
            }

            if (ClientGroups.Count > 0)
            {
                ClientFilter = " WHERE ClientID IN" +
                    " (SELECT ClientID FROM infiniu2_marketingreference.dbo.Clients" +
                    " WHERE ClientGroupID IN (" + string.Join(",", ClientGroups.OfType<Int32>().ToArray()) + "))";
            }
            else
                ClientFilter = " WHERE ClientID = -1";

            SelectCommand = "SELECT FrontsOrders.FrontID, FrontsOrders.PatinaID, FrontsOrders.ColorID, FrontsOrders.InsetTypeID," +
                " FrontsOrders.InsetColorID, FrontsOrders.TechnoColorID, FrontsOrders.TechnoInsetTypeID, FrontsOrders.TechnoInsetColorID, FrontsOrders.Height, FrontsOrders.Width, FrontsOrders.Count, FrontsOrders.Square, FrontsOrders.Cost, MeasureID, ClientID FROM FrontsOrders" +
                " INNER JOIN MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID" +
                " INNER JOIN MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID" +
                " INNER JOIN infiniu2_catalog.dbo.FrontsConfig" +
                " ON FrontsOrders.FrontConfigID=infiniu2_catalog.dbo.FrontsConfig.FrontConfigID" +
                " WHERE " + FrontsFactoryFilter + " FrontsOrders.MainOrderID IN (SELECT MainOrderID FROM MainOrders " + MainOrdersFilter +
                " AND MegaOrderID IN" +
                " (SELECT MegaOrderID FROM MegaOrders " + ClientFilter + AgreementFilter + "))";

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(DT);
            }

            return DT;
        }

        /// <summary>
        /// фасады на согласовании
        /// </summary>
        /// <param name="FactoryID"></param>
        /// <returns></returns>
        public DataTable MarketingAgreedDecor(ArrayList ClientGroups, DateTime FixingDate, int FactoryID)
        {
            string AgreementFilter = " AND AgreementStatusID = 2";
            string ClientFilter = string.Empty;

            string MainOrdersFilter = string.Empty;
            string DecorFactoryFilter = string.Empty;

            string SelectCommand = string.Empty;

            DataTable DT = new DataTable();

            if (FactoryID == -1)
            {
                DecorFactoryFilter = " DecorOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = " WHERE MainOrderID = -1";
            }
            if (FactoryID == 0)
            {
                MainOrdersFilter = " WHERE (DocDateTime < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") +
                      "' AND ( ( (ProfilOnProductionDate IS NULL AND ProfilProductionStatusID = 1 AND ProfilStorageStatusID = 1 AND ProfilDispatchStatusID = 1) OR (ProfilOnProductionDate > '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "'))" +
                    " OR " + " ( (TPSOnProductionDate IS NULL AND TPSProductionStatusID = 1 AND TPSStorageStatusID = 1 AND TPSDispatchStatusID = 1) OR (TPSOnProductionDate > '" +
                    FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "') ) ))";
            }
            if (FactoryID == 1)
            {
                DecorFactoryFilter = " DecorOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = " WHERE (MainOrders.FactoryID <> 2 AND DocDateTime < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") +
                    "' AND ((ProfilOnProductionDate IS NULL AND ProfilProductionStatusID = 1 AND ProfilStorageStatusID = 1 AND ProfilDispatchStatusID = 1) OR (ProfilOnProductionDate > '" +
                    FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "')))";
            }
            if (FactoryID == 2)
            {
                DecorFactoryFilter = " DecorOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = " WHERE (MainOrders.FactoryID <> 1 AND DocDateTime < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") +
                    "' AND ((TPSOnProductionDate IS NULL AND TPSProductionStatusID = 1 AND TPSStorageStatusID = 1 AND TPSDispatchStatusID = 1) OR (TPSOnProductionDate > '" +
                    FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "')))";
            }

            if (ClientGroups.Count > 0)
            {
                ClientFilter = " WHERE ClientID IN" +
                    " (SELECT ClientID FROM infiniu2_marketingreference.dbo.Clients" +
                    " WHERE ClientGroupID IN (" + string.Join(",", ClientGroups.OfType<Int32>().ToArray()) + "))";
            }
            else
                ClientFilter = " WHERE ClientID = -1";

            //decor
            SelectCommand = "SELECT DecorOrders.ProductID, DecorOrders.DecorID, DecorOrders.ColorID,DecorOrders.PatinaID," +
                " DecorOrders.Height, DecorOrders.Length, DecorOrders.Width, DecorOrders.Count, DecorOrders.Cost, DecorOrders.DecorConfigID, " +
                " MeasureID, ClientID FROM DecorOrders" +
                " INNER JOIN MainOrders ON DecorOrders.MainOrderID = MainOrders.MainOrderID" +
                " INNER JOIN MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID" +
                " INNER JOIN infiniu2_catalog.dbo.DecorConfig" +
                " ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                " WHERE " + DecorFactoryFilter + " DecorOrders.MainOrderID IN (SELECT MainOrderID FROM MainOrders " + MainOrdersFilter +
                " AND MegaOrderID IN" +
                " (SELECT MegaOrderID FROM MegaOrders " + ClientFilter + AgreementFilter + "))";

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(DT);
            }

            return DT;
        }

        /// <summary>
        /// фасады на производстве
        /// </summary>
        /// <param name="FactoryID"></param>
        /// <returns></returns>
        public DataTable MarketingOnProductionFronts(ArrayList ClientGroups, DateTime FixingDate, int FactoryID)
        {
            string ClientFilter = string.Empty;

            string MainOrdersFilter = string.Empty;
            string FrontsFactoryFilter = string.Empty;

            string SelectCommand = string.Empty;

            DataTable DT = new DataTable();

            if (FactoryID == -1)
            {
                FrontsFactoryFilter = " FrontsOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = " WHERE MainOrderID = -1";
            }
            if (FactoryID == 0)
            {
                MainOrdersFilter = " WHERE ((ProfilOnProductionDate < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "' AND ProfilProductionDate > '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") +
                    "') OR (TPSOnProductionDate < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "' AND TPSProductionDate > '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "'))";
            }
            if (FactoryID == 1)
            {
                FrontsFactoryFilter = " FrontsOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = " WHERE (ProfilOnProductionDate < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "' AND ProfilProductionDate > '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "')";
            }
            if (FactoryID == 2)
            {
                FrontsFactoryFilter = " FrontsOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = " WHERE (TPSOnProductionDate < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "' AND TPSProductionDate > '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "')";
            }

            if (ClientGroups.Count > 0)
            {
                ClientFilter = " WHERE ClientID IN" +
                    " (SELECT ClientID FROM infiniu2_marketingreference.dbo.Clients" +
                    " WHERE ClientGroupID IN (" + string.Join(",", ClientGroups.OfType<Int32>().ToArray()) + "))";
            }
            else
                ClientFilter = " WHERE ClientID = -1";

            SelectCommand = "SELECT FrontsOrders.FrontID, FrontsOrders.PatinaID, FrontsOrders.ColorID, FrontsOrders.InsetTypeID," +
                " FrontsOrders.InsetColorID, FrontsOrders.TechnoColorID, FrontsOrders.TechnoInsetTypeID, FrontsOrders.TechnoInsetColorID, FrontsOrders.Height, FrontsOrders.Width, FrontsOrders.Count, FrontsOrders.Square, FrontsOrders.Cost, MeasureID, ClientID FROM FrontsOrders" +
                " INNER JOIN MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID" +
                " INNER JOIN MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID" +
                " INNER JOIN infiniu2_catalog.dbo.FrontsConfig" +
                " ON FrontsOrders.FrontConfigID=infiniu2_catalog.dbo.FrontsConfig.FrontConfigID" +
                " WHERE " + FrontsFactoryFilter + " FrontsOrders.MainOrderID IN (SELECT MainOrderID FROM MainOrders " + MainOrdersFilter +
                " AND MegaOrderID IN" +
                " (SELECT MegaOrderID FROM MegaOrders " + ClientFilter + "))";

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(DT);
            }

            return DT;
        }

        /// <summary>
        /// декор на производстве
        /// </summary>
        /// <param name="FactoryID"></param>
        /// <returns></returns>
        public DataTable MarketingOnProductionDecor(ArrayList ClientGroups, DateTime FixingDate, int FactoryID)
        {
            string ClientFilter = string.Empty;

            string MainOrdersFilter = string.Empty;
            string DecorFactoryFilter = string.Empty;

            string SelectCommand = string.Empty;

            DataTable DT = new DataTable();

            if (FactoryID == -1)
            {
                DecorFactoryFilter = " DecorOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = " WHERE MainOrderID = -1";
            }
            if (FactoryID == 0)
            {
                MainOrdersFilter = " WHERE ((ProfilOnProductionDate < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "' AND ProfilProductionDate > '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") +
                    "') OR (TPSOnProductionDate < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "' AND TPSProductionDate > '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "'))";
            }
            if (FactoryID == 1)
            {
                DecorFactoryFilter = " DecorOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = " WHERE (ProfilOnProductionDate < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "' AND ProfilProductionDate > '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "')";
            }
            if (FactoryID == 2)
            {
                DecorFactoryFilter = " DecorOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = " WHERE (TPSOnProductionDate < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "' AND TPSProductionDate > '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "')";
            }

            if (ClientGroups.Count > 0)
            {
                ClientFilter = " WHERE ClientID IN" +
                    " (SELECT ClientID FROM infiniu2_marketingreference.dbo.Clients" +
                    " WHERE ClientGroupID IN (" + string.Join(",", ClientGroups.OfType<Int32>().ToArray()) + "))";
            }
            else
                ClientFilter = " WHERE ClientID = -1";

            //decor
            SelectCommand = "SELECT DecorOrders.ProductID, DecorOrders.DecorID, DecorOrders.ColorID,DecorOrders.PatinaID," +
                " DecorOrders.Height, DecorOrders.Length, DecorOrders.Width, DecorOrders.Count, DecorOrders.Cost, DecorOrders.DecorConfigID, " +
                " MeasureID, ClientID FROM DecorOrders" +
                " INNER JOIN MainOrders ON DecorOrders.MainOrderID = MainOrders.MainOrderID" +
                " INNER JOIN MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID" +
                " INNER JOIN infiniu2_catalog.dbo.DecorConfig" +
                " ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                " WHERE " + DecorFactoryFilter + " DecorOrders.MainOrderID IN (SELECT MainOrderID FROM MainOrders " + MainOrdersFilter +
                " AND MegaOrderID IN" +
                " (SELECT MegaOrderID FROM MegaOrders " + ClientFilter + "))";

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(DT);
            }

            return DT;
        }

        /// <summary>
        /// фасады в производстве
        /// </summary>
        /// <param name="FactoryID"></param>
        /// <returns></returns>
        public DataTable MarketingInProductionFronts(ArrayList ClientGroups, DateTime FixingDate, int FactoryID)
        {
            string ClientFilter = string.Empty;

            string MainOrdersFilter = string.Empty;
            string FrontsFactoryFilter = string.Empty;
            string PackagesFilter = string.Empty;

            string SelectCommand = string.Empty;

            DataTable DT = new DataTable();

            if (FactoryID == -1)
            {
                FrontsFactoryFilter = " FrontsOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = " WHERE MainOrderID = -1";
                PackagesFilter = " WHERE MainOrderID = -1";
            }
            if (FactoryID == 0)
            {
                MainOrdersFilter = " WHERE ((ProfilProductionDate < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") +
                    "') OR (TPSProductionDate < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "'))";
                PackagesFilter = " AND FrontsOrdersID IN (SELECT DISTINCT OrderID FROM PackageDetails WHERE PackageID IN (SELECT PackageID FROM Packages WHERE ProductType = 0 AND (PackingDateTime > '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "' OR PackageStatusID = 0)))";
            }
            if (FactoryID == 1)
            {
                FrontsFactoryFilter = " FrontsOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = " WHERE (ProfilProductionDate < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "')";
                PackagesFilter = " AND FrontsOrdersID IN (SELECT DISTINCT OrderID FROM PackageDetails WHERE PackageID IN (SELECT PackageID FROM Packages WHERE FactoryID = " + FactoryID + " AND ProductType = 0 AND (PackingDateTime > '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "' OR PackageStatusID = 0)))";
            }
            if (FactoryID == 2)
            {
                FrontsFactoryFilter = " FrontsOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = " WHERE (TPSProductionDate < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "')";
                PackagesFilter = " AND FrontsOrdersID IN (SELECT DISTINCT OrderID FROM PackageDetails WHERE PackageID IN (SELECT PackageID FROM Packages WHERE FactoryID = " + FactoryID + " AND ProductType = 0 AND (PackingDateTime > '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "' OR PackageStatusID = 0)))";
            }

            if (ClientGroups.Count > 0)
            {
                ClientFilter = " WHERE ClientID IN" +
                    " (SELECT ClientID FROM infiniu2_marketingreference.dbo.Clients" +
                    " WHERE ClientGroupID IN (" + string.Join(",", ClientGroups.OfType<Int32>().ToArray()) + "))";
            }
            else
                ClientFilter = " WHERE ClientID = -1";

            SelectCommand = "SELECT FrontsOrders.FrontID, FrontsOrders.PatinaID, FrontsOrders.ColorID, FrontsOrders.InsetTypeID," +
                " FrontsOrders.InsetColorID, FrontsOrders.TechnoColorID, FrontsOrders.TechnoInsetTypeID, FrontsOrders.TechnoInsetColorID, FrontsOrders.Height, FrontsOrders.Width, FrontsOrders.Count, FrontsOrders.Square, FrontsOrders.Cost, MeasureID, ClientID FROM FrontsOrders" +
                " INNER JOIN MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID" +
                " INNER JOIN MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID" +
                " INNER JOIN infiniu2_catalog.dbo.FrontsConfig" +
                " ON FrontsOrders.FrontConfigID=infiniu2_catalog.dbo.FrontsConfig.FrontConfigID" +
                " WHERE " + FrontsFactoryFilter + " FrontsOrders.MainOrderID IN (SELECT MainOrderID FROM MainOrders " + MainOrdersFilter +
                " AND MegaOrderID IN" +
                " (SELECT MegaOrderID FROM MegaOrders " + ClientFilter + "))" + PackagesFilter;

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(DT);
            }

            return DT;
        }

        /// <summary>
        /// декор в производстве
        /// </summary>
        /// <param name="FactoryID"></param>
        /// <returns></returns>
        public DataTable MarketingInProductionDecor(ArrayList ClientGroups, DateTime FixingDate, int FactoryID)
        {
            string ClientFilter = string.Empty;

            string MainOrdersFilter = string.Empty;
            string DecorFactoryFilter = string.Empty;
            string PackagesFilter = string.Empty;

            string SelectCommand = string.Empty;

            DataTable DT = new DataTable();

            if (FactoryID == -1)
            {
                DecorFactoryFilter = " DecorOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = " WHERE MainOrderID = -1";
                PackagesFilter = " WHERE MainOrderID = -1";
            }
            if (FactoryID == 0)
            {
                MainOrdersFilter = " WHERE ((ProfilProductionDate < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") +
                    "') OR (TPSProductionDate < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "'))";
                PackagesFilter = " AND DecorOrderID IN (SELECT DISTINCT OrderID FROM PackageDetails WHERE PackageID IN (SELECT PackageID FROM Packages WHERE ProductType = 1 AND (PackingDateTime > '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "' OR PackageStatusID = 0)))";
            }
            if (FactoryID == 1)
            {
                DecorFactoryFilter = " DecorOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = " WHERE (ProfilProductionDate < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "')";
                PackagesFilter = " AND DecorOrderID IN (SELECT DISTINCT OrderID FROM PackageDetails WHERE PackageID IN (SELECT PackageID FROM Packages WHERE FactoryID = " + FactoryID + " AND ProductType = 1 AND (PackingDateTime > '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "' OR PackageStatusID = 0)))";
            }
            if (FactoryID == 2)
            {
                DecorFactoryFilter = " DecorOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = " WHERE (TPSProductionDate < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "')";
                PackagesFilter = " AND DecorOrderID IN (SELECT DISTINCT OrderID FROM PackageDetails WHERE PackageID IN (SELECT PackageID FROM Packages WHERE FactoryID = " + FactoryID + " AND ProductType = 1 AND (PackingDateTime > '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "' OR PackageStatusID = 0)))";
            }

            if (ClientGroups.Count > 0)
            {
                ClientFilter = " WHERE ClientID IN" +
                    " (SELECT ClientID FROM infiniu2_marketingreference.dbo.Clients" +
                    " WHERE ClientGroupID IN (" + string.Join(",", ClientGroups.OfType<Int32>().ToArray()) + "))";
            }
            else
                ClientFilter = " WHERE ClientID = -1";

            //decor
            SelectCommand = "SELECT DecorOrders.ProductID, DecorOrders.DecorID, DecorOrders.ColorID,DecorOrders.PatinaID," +
                " DecorOrders.Height, DecorOrders.Length, DecorOrders.Width, DecorOrders.Count, DecorOrders.Cost, DecorOrders.DecorConfigID, " +
                " MeasureID, ClientID FROM DecorOrders" +
                " INNER JOIN MainOrders ON DecorOrders.MainOrderID = MainOrders.MainOrderID" +
                " INNER JOIN MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID" +
                " INNER JOIN infiniu2_catalog.dbo.DecorConfig" +
                " ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                " WHERE " + DecorFactoryFilter + " DecorOrders.MainOrderID IN (SELECT MainOrderID FROM MainOrders " + MainOrdersFilter +
                " AND MegaOrderID IN" +
                " (SELECT MegaOrderID FROM MegaOrders " + ClientFilter + "))" + PackagesFilter;

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(DT);
            }

            return DT;
        }

        /// <summary>
        /// фасады из производства
        /// </summary>
        /// <param name="FactoryID"></param>
        /// <returns></returns>
        public DataTable MarketingOutProductionFronts(ArrayList ClientGroups, DateTime FixingDate, int FactoryID)
        {
            string ClientFilter = string.Empty;

            string MainOrdersFilter = string.Empty;
            string FrontsFactoryFilter = string.Empty;
            string PackagesFilter = string.Empty;

            string SelectCommand = string.Empty;

            DataTable DT = new DataTable();

            if (FactoryID == -1)
            {
                FrontsFactoryFilter = " FrontsOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = " WHERE MainOrderID = -1";
                PackagesFilter = " WHERE MainOrderID = -1";
            }
            if (FactoryID == 0)
            {
                MainOrdersFilter = " WHERE ((ProfilProductionDate < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") +
                    "') OR (TPSProductionDate < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "'))";
                PackagesFilter = " AND FrontsOrdersID IN (SELECT DISTINCT OrderID FROM PackageDetails WHERE PackageID IN (SELECT PackageID FROM Packages WHERE ProductType = 0 AND (PackingDateTime > '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "' OR PackageStatusID = 1)))";
            }
            if (FactoryID == 1)
            {
                FrontsFactoryFilter = " FrontsOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = " WHERE (ProfilProductionDate < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "')";
                PackagesFilter = " AND FrontsOrdersID IN (SELECT DISTINCT OrderID FROM PackageDetails WHERE PackageID IN (SELECT PackageID FROM Packages WHERE FactoryID = " + FactoryID + " AND ProductType = 0 AND (PackingDateTime > '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "' OR PackageStatusID = 1)))";
            }
            if (FactoryID == 2)
            {
                FrontsFactoryFilter = " FrontsOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = " WHERE (TPSProductionDate < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "')";
                PackagesFilter = " AND FrontsOrdersID IN (SELECT DISTINCT OrderID FROM PackageDetails WHERE PackageID IN (SELECT PackageID FROM Packages WHERE FactoryID = " + FactoryID + " AND ProductType = 0 AND (PackingDateTime > '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "' OR PackageStatusID = 1)))";
            }

            if (ClientGroups.Count > 0)
            {
                ClientFilter = " WHERE ClientID IN" +
                    " (SELECT ClientID FROM infiniu2_marketingreference.dbo.Clients" +
                    " WHERE ClientGroupID IN (" + string.Join(",", ClientGroups.OfType<Int32>().ToArray()) + "))";
            }
            else
                ClientFilter = " WHERE ClientID = -1";

            SelectCommand = "SELECT FrontsOrders.FrontID, FrontsOrders.PatinaID, FrontsOrders.ColorID, FrontsOrders.InsetTypeID," +
                " FrontsOrders.InsetColorID, FrontsOrders.TechnoColorID, FrontsOrders.TechnoInsetTypeID, FrontsOrders.TechnoInsetColorID, FrontsOrders.Height, FrontsOrders.Width, FrontsOrders.Count, FrontsOrders.Square, FrontsOrders.Cost, MeasureID, ClientID FROM FrontsOrders" +
                " INNER JOIN MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID" +
                " INNER JOIN MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID" +
                " INNER JOIN infiniu2_catalog.dbo.FrontsConfig" +
                " ON FrontsOrders.FrontConfigID=infiniu2_catalog.dbo.FrontsConfig.FrontConfigID" +
                " WHERE " + FrontsFactoryFilter + " FrontsOrders.MainOrderID IN (SELECT MainOrderID FROM MainOrders " + MainOrdersFilter +
                " AND MegaOrderID IN" +
                " (SELECT MegaOrderID FROM MegaOrders " + ClientFilter + "))" + PackagesFilter;

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(DT);
            }

            return DT;
        }

        /// <summary>
        /// декор из производства
        /// </summary>
        /// <param name="FactoryID"></param>
        /// <returns></returns>
        public DataTable MarketingOutProductionDecor(ArrayList ClientGroups, DateTime FixingDate, int FactoryID)
        {
            string ClientFilter = string.Empty;

            string MainOrdersFilter = string.Empty;
            string DecorFactoryFilter = string.Empty;
            string PackagesFilter = string.Empty;

            string SelectCommand = string.Empty;

            DataTable DT = new DataTable();

            if (FactoryID == -1)
            {
                DecorFactoryFilter = " DecorOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = " WHERE MainOrderID = -1";
                PackagesFilter = " WHERE MainOrderID = -1";
            }
            if (FactoryID == 0)
            {
                MainOrdersFilter = " WHERE ((ProfilProductionDate < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") +
                    "') OR (TPSProductionDate < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "'))";
                PackagesFilter = " AND DecorOrderID IN (SELECT DISTINCT OrderID FROM PackageDetails WHERE PackageID IN (SELECT PackageID FROM Packages WHERE ProductType = 1 AND (PackingDateTime > '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "' OR PackageStatusID = 1)))";
            }
            if (FactoryID == 1)
            {
                DecorFactoryFilter = " DecorOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = " WHERE (ProfilProductionDate < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "')";
                PackagesFilter = " AND DecorOrderID IN (SELECT DISTINCT OrderID FROM PackageDetails WHERE PackageID IN (SELECT PackageID FROM Packages WHERE FactoryID = " + FactoryID + " AND ProductType = 1 AND (PackingDateTime > '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "' OR PackageStatusID = 1)))";
            }
            if (FactoryID == 2)
            {
                DecorFactoryFilter = " DecorOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = " WHERE (TPSProductionDate < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "')";
                PackagesFilter = " AND DecorOrderID IN (SELECT DISTINCT OrderID FROM PackageDetails WHERE PackageID IN (SELECT PackageID FROM Packages WHERE FactoryID = " + FactoryID + " AND ProductType = 1 AND (PackingDateTime > '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "' OR PackageStatusID = 1)))";
            }

            if (ClientGroups.Count > 0)
            {
                ClientFilter = " WHERE ClientID IN" +
                    " (SELECT ClientID FROM infiniu2_marketingreference.dbo.Clients" +
                    " WHERE ClientGroupID IN (" + string.Join(",", ClientGroups.OfType<Int32>().ToArray()) + "))";
            }
            else
                ClientFilter = " WHERE ClientID = -1";

            //decor
            SelectCommand = "SELECT DecorOrders.ProductID, DecorOrders.DecorID, DecorOrders.ColorID,DecorOrders.PatinaID," +
                " DecorOrders.Height, DecorOrders.Length, DecorOrders.Width, DecorOrders.Count, DecorOrders.Cost, DecorOrders.DecorConfigID, " +
                " MeasureID, ClientID FROM DecorOrders" +
                " INNER JOIN MainOrders ON DecorOrders.MainOrderID = MainOrders.MainOrderID" +
                " INNER JOIN MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID" +
                " INNER JOIN infiniu2_catalog.dbo.DecorConfig" +
                " ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                " WHERE " + DecorFactoryFilter + " DecorOrders.MainOrderID IN (SELECT MainOrderID FROM MainOrders " + MainOrdersFilter +
                " AND MegaOrderID IN" +
                " (SELECT MegaOrderID FROM MegaOrders " + ClientFilter + "))" + PackagesFilter;

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(DT);
            }

            return DT;
        }

        #endregion Marketing

        #region ZOV

        /// <summary>
        /// ЗОВ фасады на согласовании
        /// </summary>
        /// <param name="FactoryID"></param>
        /// <returns></returns>
        public DataTable ZOVOnAgreementFronts(DateTime FixingDate, int FactoryID)
        {
            string AgreementFilter = " AND MegaOrderID = 0";
            string MainOrdersFilter = string.Empty;
            string FrontsFactoryFilter = string.Empty;

            string SelectCommand = string.Empty;

            DataTable DT = new DataTable();

            if (FactoryID == -1)
            {
                FrontsFactoryFilter = " FrontsOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = @" WHERE MainOrderID = -1";
            }
            if (FactoryID == 0)
            {
                MainOrdersFilter = " WHERE (DocDateTime < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") +
                      "' AND ( ( (ProfilOnProductionDate IS NULL AND ProfilProductionStatusID = 1 AND ProfilStorageStatusID = 1 AND ProfilDispatchStatusID = 1) OR (ProfilOnProductionDate > '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "'))" +
                    " OR " + " ( (TPSOnProductionDate IS NULL AND TPSProductionStatusID = 1 AND TPSStorageStatusID = 1 AND TPSDispatchStatusID = 1) OR (TPSOnProductionDate > '" +
                    FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "') ) ))" + AgreementFilter;
            }
            if (FactoryID == 1)
            {
                FrontsFactoryFilter = " FrontsOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = " WHERE (MainOrders.FactoryID <> 2 AND DocDateTime < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") +
                    "' AND ((ProfilOnProductionDate IS NULL AND ProfilProductionStatusID = 1 AND ProfilStorageStatusID = 1 AND ProfilDispatchStatusID = 1) OR (ProfilOnProductionDate > '" +
                    FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "')))" + AgreementFilter;
            }
            if (FactoryID == 2)
            {
                FrontsFactoryFilter = " FrontsOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = " WHERE (MainOrders.FactoryID <> 1 AND DocDateTime < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") +
                    "' AND ((TPSOnProductionDate IS NULL AND TPSProductionStatusID = 1 AND TPSStorageStatusID = 1 AND TPSDispatchStatusID = 1) OR (TPSOnProductionDate > '" +
                    FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "')))" + AgreementFilter;
            }

            SelectCommand = "SELECT FrontsOrders.FrontID, FrontsOrders.PatinaID, FrontsOrders.ColorID, FrontsOrders.InsetTypeID," +
                " FrontsOrders.InsetColorID, FrontsOrders.TechnoColorID, FrontsOrders.TechnoInsetTypeID, FrontsOrders.TechnoInsetColorID, FrontsOrders.Height, FrontsOrders.Width, FrontsOrders.Count, FrontsOrders.Square, FrontsOrders.Cost, MeasureID FROM FrontsOrders" +
                " INNER JOIN infiniu2_catalog.dbo.FrontsConfig" +
                " ON FrontsOrders.FrontConfigID=infiniu2_catalog.dbo.FrontsConfig.FrontConfigID" +
                " WHERE " + FrontsFactoryFilter + " FrontsOrders.MainOrderID IN (SELECT MainOrderID FROM MainOrders " + MainOrdersFilter + ")";

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.ZOVOrdersConnectionString))
            {
                DA.Fill(DT);
            }

            return DT;
        }

        /// <summary>
        /// ЗОВ фасады на согласовании
        /// </summary>
        /// <param name="FactoryID"></param>
        /// <returns></returns>
        public DataTable ZOVOnAgreementDecor(DateTime FixingDate, int FactoryID)
        {
            string AgreementFilter = " AND MegaOrderID = 0";
            string MainOrdersFilter = string.Empty;
            string DecorFactoryFilter = string.Empty;

            string SelectCommand = string.Empty;

            DataTable DT = new DataTable();

            if (FactoryID == -1)
            {
                DecorFactoryFilter = " DecorOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = " WHERE MainOrderID = -1";
            }
            if (FactoryID == 0)
            {
                MainOrdersFilter = " WHERE (DocDateTime < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") +
                      "' AND ( ( (ProfilOnProductionDate IS NULL AND ProfilProductionStatusID = 1 AND ProfilStorageStatusID = 1 AND ProfilDispatchStatusID = 1) OR (ProfilOnProductionDate > '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "'))" +
                    " OR " + " ( (TPSOnProductionDate IS NULL AND TPSProductionStatusID = 1 AND TPSStorageStatusID = 1 AND TPSDispatchStatusID = 1) OR (TPSOnProductionDate > '" +
                    FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "') ) ))" + AgreementFilter;
            }
            if (FactoryID == 1)
            {
                DecorFactoryFilter = " DecorOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = " WHERE (MainOrders.FactoryID <> 2 AND DocDateTime < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") +
                    "' AND ((ProfilOnProductionDate IS NULL AND ProfilProductionStatusID = 1 AND ProfilStorageStatusID = 1 AND ProfilDispatchStatusID = 1) OR (ProfilOnProductionDate > '" +
                    FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "')))" + AgreementFilter;
            }
            if (FactoryID == 2)
            {
                DecorFactoryFilter = " DecorOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = " WHERE (MainOrders.FactoryID <> 1 AND DocDateTime < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") +
                    "' AND ((TPSOnProductionDate IS NULL AND TPSProductionStatusID = 1 AND TPSStorageStatusID = 1 AND TPSDispatchStatusID = 1) OR (TPSOnProductionDate > '" +
                    FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "')))" + AgreementFilter;
            }

            //decor
            SelectCommand = "SELECT DecorOrders.ProductID, DecorOrders.DecorID, DecorOrders.ColorID,DecorOrders.PatinaID," +
                " DecorOrders.Height, DecorOrders.Length, DecorOrders.Width, DecorOrders.Count, DecorOrders.Cost, DecorOrders.DecorConfigID, " +
                " MeasureID FROM DecorOrders" +
                " INNER JOIN infiniu2_catalog.dbo.DecorConfig" +
                " ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                " WHERE " + DecorFactoryFilter + " DecorOrders.MainOrderID IN (SELECT MainOrderID FROM MainOrders " + MainOrdersFilter + ")";

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.ZOVOrdersConnectionString))
            {
                DA.Fill(DT);
            }

            return DT;
        }

        /// <summary>
        /// ЗОВ фасады на согласовании
        /// </summary>
        /// <param name="FactoryID"></param>
        /// <returns></returns>
        public DataTable ZOVAgreedFronts(DateTime FixingDate, int FactoryID)
        {
            string AgreementFilter = " AND MegaOrderID <> 0";
            string MainOrdersFilter = string.Empty;
            string FrontsFactoryFilter = string.Empty;

            string SelectCommand = string.Empty;

            DataTable DT = new DataTable();

            if (FactoryID == -1)
            {
                FrontsFactoryFilter = " FrontsOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = @" WHERE MainOrderID = -1";
            }
            if (FactoryID == 0)
            {
                MainOrdersFilter = " WHERE (DocDateTime < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") +
                      "' AND ( ( (ProfilOnProductionDate IS NULL AND ProfilProductionStatusID = 1 AND ProfilStorageStatusID = 1 AND ProfilDispatchStatusID = 1) OR (ProfilOnProductionDate > '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "'))" +
                    " OR " + " ( (TPSOnProductionDate IS NULL AND TPSProductionStatusID = 1 AND TPSStorageStatusID = 1 AND TPSDispatchStatusID = 1) OR (TPSOnProductionDate > '" +
                    FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "') ) ))" + AgreementFilter;
            }
            if (FactoryID == 1)
            {
                FrontsFactoryFilter = " FrontsOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = " WHERE (MainOrders.FactoryID <> 2 AND DocDateTime < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") +
                    "' AND ((ProfilOnProductionDate IS NULL AND ProfilProductionStatusID = 1 AND ProfilStorageStatusID = 1 AND ProfilDispatchStatusID = 1) OR (ProfilOnProductionDate > '" +
                    FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "')))" + AgreementFilter;
            }
            if (FactoryID == 2)
            {
                FrontsFactoryFilter = " FrontsOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = " WHERE (MainOrders.FactoryID <> 1 AND DocDateTime < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") +
                    "' AND ((TPSOnProductionDate IS NULL AND TPSProductionStatusID = 1 AND TPSStorageStatusID = 1 AND TPSDispatchStatusID = 1) OR (TPSOnProductionDate > '" +
                    FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "')))" + AgreementFilter;
            }

            SelectCommand = "SELECT FrontsOrders.FrontID, FrontsOrders.PatinaID, FrontsOrders.ColorID, FrontsOrders.InsetTypeID," +
                " FrontsOrders.InsetColorID, FrontsOrders.TechnoColorID, FrontsOrders.TechnoInsetTypeID, FrontsOrders.TechnoInsetColorID, FrontsOrders.Height, FrontsOrders.Width, FrontsOrders.Count, FrontsOrders.Square, FrontsOrders.Cost, MeasureID FROM FrontsOrders" +
                " INNER JOIN infiniu2_catalog.dbo.FrontsConfig" +
                " ON FrontsOrders.FrontConfigID=infiniu2_catalog.dbo.FrontsConfig.FrontConfigID" +
                " WHERE " + FrontsFactoryFilter + " FrontsOrders.MainOrderID IN (SELECT MainOrderID FROM MainOrders " + MainOrdersFilter + ")";

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.ZOVOrdersConnectionString))
            {
                DA.Fill(DT);
            }

            return DT;
        }

        /// <summary>
        /// ЗОВ фасады на согласовании
        /// </summary>
        /// <param name="FactoryID"></param>
        /// <returns></returns>
        public DataTable ZOVAgreedDecor(DateTime FixingDate, int FactoryID)
        {
            string AgreementFilter = " AND MegaOrderID <> 0";
            string MainOrdersFilter = string.Empty;
            string DecorFactoryFilter = string.Empty;

            string SelectCommand = string.Empty;

            DataTable DT = new DataTable();

            if (FactoryID == -1)
            {
                DecorFactoryFilter = " DecorOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = " WHERE MainOrderID = -1";
            }
            if (FactoryID == 0)
            {
                MainOrdersFilter = " WHERE (DocDateTime < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") +
                      "' AND ( ( (ProfilOnProductionDate IS NULL AND ProfilProductionStatusID = 1 AND ProfilStorageStatusID = 1 AND ProfilDispatchStatusID = 1) OR (ProfilOnProductionDate > '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "'))" +
                    " OR " + " ( (TPSOnProductionDate IS NULL AND TPSProductionStatusID = 1 AND TPSStorageStatusID = 1 AND TPSDispatchStatusID = 1) OR (TPSOnProductionDate > '" +
                    FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "') ) ))" + AgreementFilter;
            }
            if (FactoryID == 1)
            {
                DecorFactoryFilter = " DecorOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = " WHERE (MainOrders.FactoryID <> 2 AND DocDateTime < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") +
                    "' AND ((ProfilOnProductionDate IS NULL AND ProfilProductionStatusID = 1 AND ProfilStorageStatusID = 1 AND ProfilDispatchStatusID = 1) OR (ProfilOnProductionDate > '" +
                    FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "')))" + AgreementFilter;
            }
            if (FactoryID == 2)
            {
                DecorFactoryFilter = " DecorOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = " WHERE (MainOrders.FactoryID <> 1 AND DocDateTime < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") +
                    "' AND ((TPSOnProductionDate IS NULL AND TPSProductionStatusID = 1 AND TPSStorageStatusID = 1 AND TPSDispatchStatusID = 1) OR (TPSOnProductionDate > '" +
                    FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "')))" + AgreementFilter;
            }

            //decor
            SelectCommand = "SELECT DecorOrders.ProductID, DecorOrders.DecorID, DecorOrders.ColorID,DecorOrders.PatinaID," +
                " DecorOrders.Height, DecorOrders.Length, DecorOrders.Width, DecorOrders.Count, DecorOrders.Cost, DecorOrders.DecorConfigID, " +
                " MeasureID FROM DecorOrders" +
                " INNER JOIN infiniu2_catalog.dbo.DecorConfig" +
                " ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                " WHERE " + DecorFactoryFilter + " DecorOrders.MainOrderID IN (SELECT MainOrderID FROM MainOrders " + MainOrdersFilter + ")";

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.ZOVOrdersConnectionString))
            {
                DA.Fill(DT);
            }

            return DT;
        }

        /// <summary>
        /// ЗОВ фасады на производстве
        /// </summary>
        /// <param name="FactoryID"></param>
        /// <returns></returns>
        public DataTable ZOVOnProductionFronts(DateTime FixingDate, int FactoryID)
        {
            string ClientFilter = string.Empty;

            string MainOrdersFilter = string.Empty;
            string FrontsFactoryFilter = string.Empty;

            string SelectCommand = string.Empty;

            DataTable DT = new DataTable();

            if (FactoryID == -1)
            {
                FrontsFactoryFilter = " FrontsOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = " WHERE MainOrderID = -1";
            }
            if (FactoryID == 0)
            {
                MainOrdersFilter = " WHERE ((ProfilOnProductionDate < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "' AND ProfilProductionDate > '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") +
                    "') OR (TPSOnProductionDate < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "' AND TPSProductionDate > '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "'))";
            }
            if (FactoryID == 1)
            {
                FrontsFactoryFilter = " FrontsOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = " WHERE (ProfilOnProductionDate < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "' AND ProfilProductionDate > '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "')";
            }
            if (FactoryID == 2)
            {
                FrontsFactoryFilter = " FrontsOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = " WHERE (TPSOnProductionDate < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "' AND TPSProductionDate > '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "')";
            }

            SelectCommand = "SELECT FrontsOrders.FrontID, FrontsOrders.PatinaID, FrontsOrders.ColorID, FrontsOrders.InsetTypeID," +
                " FrontsOrders.InsetColorID, FrontsOrders.TechnoColorID, FrontsOrders.TechnoInsetTypeID, FrontsOrders.TechnoInsetColorID, FrontsOrders.Height, FrontsOrders.Width, FrontsOrders.Count, FrontsOrders.Square, FrontsOrders.Cost, MeasureID FROM FrontsOrders" +
                " INNER JOIN infiniu2_catalog.dbo.FrontsConfig" +
                " ON FrontsOrders.FrontConfigID=infiniu2_catalog.dbo.FrontsConfig.FrontConfigID" +
                " WHERE " + FrontsFactoryFilter + " FrontsOrders.MainOrderID IN (SELECT MainOrderID FROM MainOrders " + MainOrdersFilter + ")";

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.ZOVOrdersConnectionString))
            {
                DA.Fill(DT);
            }

            return DT;
        }

        /// <summary>
        /// ЗОВ декор на производстве
        /// </summary>
        /// <param name="FactoryID"></param>
        /// <returns></returns>
        public DataTable ZOVOnProductionDecor(DateTime FixingDate, int FactoryID)
        {
            string ClientFilter = string.Empty;

            string MainOrdersFilter = string.Empty;
            string DecorFactoryFilter = string.Empty;

            string SelectCommand = string.Empty;

            DataTable DT = new DataTable();

            if (FactoryID == -1)
            {
                DecorFactoryFilter = " DecorOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = " WHERE MainOrderID = -1";
            }
            if (FactoryID == 0)
            {
                MainOrdersFilter = " WHERE ((ProfilOnProductionDate < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "' AND ProfilProductionDate > '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") +
                    "') OR (TPSOnProductionDate < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "' AND TPSProductionDate > '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "'))";
            }
            if (FactoryID == 1)
            {
                DecorFactoryFilter = " DecorOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = " WHERE (ProfilOnProductionDate < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "' AND ProfilProductionDate > '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "')";
            }
            if (FactoryID == 2)
            {
                DecorFactoryFilter = " DecorOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = " WHERE (TPSOnProductionDate < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "' AND TPSProductionDate > '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "')";
            }

            //decor
            SelectCommand = "SELECT DecorOrders.ProductID, DecorOrders.DecorID, DecorOrders.ColorID,DecorOrders.PatinaID," +
                " DecorOrders.Height, DecorOrders.Length, DecorOrders.Width, DecorOrders.Count, DecorOrders.Cost, DecorOrders.DecorConfigID, " +
                " MeasureID FROM DecorOrders" +
                " INNER JOIN infiniu2_catalog.dbo.DecorConfig" +
                " ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                " WHERE " + DecorFactoryFilter + " DecorOrders.MainOrderID IN (SELECT MainOrderID FROM MainOrders " + MainOrdersFilter +
                " AND MegaOrderID IN" +
                " (SELECT MegaOrderID FROM MegaOrders " + ClientFilter + "))";

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.ZOVOrdersConnectionString))
            {
                DA.Fill(DT);
            }

            return DT;
        }

        /// <summary>
        /// ЗОВ фасады в производстве
        /// </summary>
        /// <param name="FactoryID"></param>
        /// <returns></returns>
        public DataTable ZOVInProductionFronts(DateTime FixingDate, int FactoryID)
        {
            string ClientFilter = string.Empty;

            string MainOrdersFilter = string.Empty;
            string FrontsFactoryFilter = string.Empty;
            string PackagesFilter = string.Empty;

            string SelectCommand = string.Empty;

            DataTable DT = new DataTable();

            if (FactoryID == -1)
            {
                FrontsFactoryFilter = " FrontsOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = " WHERE MainOrderID = -1";
                PackagesFilter = " WHERE MainOrderID = -1";
            }
            if (FactoryID == 0)
            {
                MainOrdersFilter = " WHERE ((ProfilProductionDate < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") +
                    "') OR (TPSProductionDate < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "'))";
                PackagesFilter = " AND FrontsOrdersID IN (SELECT DISTINCT OrderID FROM PackageDetails WHERE PackageID IN (SELECT PackageID FROM Packages WHERE ProductType = 0 AND (PackingDateTime > '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "' OR PackageStatusID = 0)))";
            }
            if (FactoryID == 1)
            {
                FrontsFactoryFilter = " FrontsOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = " WHERE (ProfilProductionDate < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "')";
                PackagesFilter = " AND FrontsOrdersID IN (SELECT DISTINCT OrderID FROM PackageDetails WHERE PackageID IN (SELECT PackageID FROM Packages WHERE FactoryID = " + FactoryID + " AND ProductType = 0 AND (PackingDateTime > '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "' OR PackageStatusID = 0)))";
            }
            if (FactoryID == 2)
            {
                FrontsFactoryFilter = " FrontsOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = " WHERE (TPSProductionDate < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "')";
                PackagesFilter = " AND FrontsOrdersID IN (SELECT DISTINCT OrderID FROM PackageDetails WHERE PackageID IN (SELECT PackageID FROM Packages WHERE FactoryID = " + FactoryID + " AND ProductType = 0 AND (PackingDateTime > '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "' OR PackageStatusID = 0)))";
            }

            SelectCommand = "SELECT FrontsOrders.FrontID, FrontsOrders.PatinaID, FrontsOrders.ColorID, FrontsOrders.InsetTypeID," +
                " FrontsOrders.InsetColorID, FrontsOrders.TechnoColorID, FrontsOrders.TechnoInsetTypeID, FrontsOrders.TechnoInsetColorID, FrontsOrders.Height, FrontsOrders.Width, FrontsOrders.Count, FrontsOrders.Square, FrontsOrders.Cost, MeasureID FROM FrontsOrders" +
                " INNER JOIN infiniu2_catalog.dbo.FrontsConfig" +
                " ON FrontsOrders.FrontConfigID=infiniu2_catalog.dbo.FrontsConfig.FrontConfigID" +
                " WHERE " + FrontsFactoryFilter + " FrontsOrders.MainOrderID IN (SELECT MainOrderID FROM MainOrders " + MainOrdersFilter + ")" + PackagesFilter;

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.ZOVOrdersConnectionString))
            {
                DA.Fill(DT);
            }

            return DT;
        }

        /// <summary>
        /// ЗОВ декор в производстве
        /// </summary>
        /// <param name="FactoryID"></param>
        /// <returns></returns>
        public DataTable ZOVInProductionDecor(DateTime FixingDate, int FactoryID)
        {
            string MainOrdersFilter = string.Empty;
            string DecorFactoryFilter = string.Empty;
            string PackagesFilter = string.Empty;

            string SelectCommand = string.Empty;

            DataTable DT = new DataTable();

            if (FactoryID == -1)
            {
                DecorFactoryFilter = " DecorOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = " WHERE MainOrderID = -1";
                PackagesFilter = " WHERE MainOrderID = -1";
            }
            if (FactoryID == 0)
            {
                MainOrdersFilter = " WHERE ((ProfilProductionDate < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") +
                    "') OR (TPSProductionDate < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "'))";
                PackagesFilter = " AND DecorOrderID IN (SELECT DISTINCT OrderID FROM PackageDetails WHERE PackageID IN (SELECT PackageID FROM Packages WHERE ProductType = 1 AND (PackingDateTime > '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "' OR PackageStatusID = 0)))";
            }
            if (FactoryID == 1)
            {
                DecorFactoryFilter = " DecorOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = " WHERE (ProfilProductionDate < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "')";
                PackagesFilter = " AND DecorOrderID IN (SELECT DISTINCT OrderID FROM PackageDetails WHERE PackageID IN (SELECT PackageID FROM Packages WHERE FactoryID = " + FactoryID + " AND ProductType = 1 AND (PackingDateTime > '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "' OR PackageStatusID = 0)))";
            }
            if (FactoryID == 2)
            {
                DecorFactoryFilter = " DecorOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = " WHERE (TPSProductionDate < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "')";
                PackagesFilter = " AND DecorOrderID IN (SELECT DISTINCT OrderID FROM PackageDetails WHERE PackageID IN (SELECT PackageID FROM Packages WHERE FactoryID = " + FactoryID + " AND ProductType = 1 AND (PackingDateTime > '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "' OR PackageStatusID = 0)))";
            }

            //decor
            SelectCommand = "SELECT DecorOrders.ProductID, DecorOrders.DecorID, DecorOrders.ColorID,DecorOrders.PatinaID," +
                " DecorOrders.Height, DecorOrders.Length, DecorOrders.Width, DecorOrders.Count, DecorOrders.Cost, DecorOrders.DecorConfigID, " +
                " MeasureID FROM DecorOrders" +
                " INNER JOIN infiniu2_catalog.dbo.DecorConfig" +
                " ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                " WHERE " + DecorFactoryFilter + " DecorOrders.MainOrderID IN (SELECT MainOrderID FROM MainOrders " + MainOrdersFilter + ")" + PackagesFilter;

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.ZOVOrdersConnectionString))
            {
                DA.Fill(DT);
            }

            return DT;
        }

        /// <summary>
        /// ЗОВ фасады из производства
        /// </summary>
        /// <param name="FactoryID"></param>
        /// <returns></returns>
        public DataTable ZOVOutProductionFronts(DateTime FixingDate, int FactoryID)
        {
            string ClientFilter = string.Empty;

            string MainOrdersFilter = string.Empty;
            string FrontsFactoryFilter = string.Empty;
            string PackagesFilter = string.Empty;

            string SelectCommand = string.Empty;

            DataTable DT = new DataTable();

            if (FactoryID == -1)
            {
                FrontsFactoryFilter = " FrontsOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = " WHERE MainOrderID = -1";
                PackagesFilter = " WHERE MainOrderID = -1";
            }
            if (FactoryID == 0)
            {
                MainOrdersFilter = " WHERE ((ProfilProductionDate < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") +
                    "') OR (TPSProductionDate < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "'))";
                PackagesFilter = " AND FrontsOrdersID IN (SELECT DISTINCT OrderID FROM PackageDetails WHERE PackageID IN (SELECT PackageID FROM Packages WHERE ProductType = 0 AND (PackingDateTime > '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "' OR PackageStatusID = 1)))";
            }
            if (FactoryID == 1)
            {
                FrontsFactoryFilter = " FrontsOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = " WHERE (ProfilProductionDate < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "')";
                PackagesFilter = " AND FrontsOrdersID IN (SELECT DISTINCT OrderID FROM PackageDetails WHERE PackageID IN (SELECT PackageID FROM Packages WHERE FactoryID = " + FactoryID + " AND ProductType = 0 AND (PackingDateTime > '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "' OR PackageStatusID = 1)))";
            }
            if (FactoryID == 2)
            {
                FrontsFactoryFilter = " FrontsOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = " WHERE (TPSProductionDate < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "')";
                PackagesFilter = " AND FrontsOrdersID IN (SELECT DISTINCT OrderID FROM PackageDetails WHERE PackageID IN (SELECT PackageID FROM Packages WHERE FactoryID = " + FactoryID + " AND ProductType = 0 AND (PackingDateTime > '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "' OR PackageStatusID = 1)))";
            }

            SelectCommand = "SELECT FrontsOrders.FrontID, FrontsOrders.PatinaID, FrontsOrders.ColorID, FrontsOrders.InsetTypeID," +
                " FrontsOrders.InsetColorID, FrontsOrders.TechnoColorID, FrontsOrders.TechnoInsetTypeID, FrontsOrders.TechnoInsetColorID, FrontsOrders.Height, FrontsOrders.Width, FrontsOrders.Count, FrontsOrders.Square, FrontsOrders.Cost, MeasureID FROM FrontsOrders" +
                " INNER JOIN infiniu2_catalog.dbo.FrontsConfig" +
                " ON FrontsOrders.FrontConfigID=infiniu2_catalog.dbo.FrontsConfig.FrontConfigID" +
                " WHERE " + FrontsFactoryFilter + " FrontsOrders.MainOrderID IN (SELECT MainOrderID FROM MainOrders " + MainOrdersFilter + ")" + PackagesFilter;

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.ZOVOrdersConnectionString))
            {
                DA.Fill(DT);
            }

            return DT;
        }

        /// <summary>
        /// ЗОВ декор из производства
        /// </summary>
        /// <param name="FactoryID"></param>
        /// <returns></returns>
        public DataTable ZOVOutProductionDecor(DateTime FixingDate, int FactoryID)
        {
            string MainOrdersFilter = string.Empty;
            string DecorFactoryFilter = string.Empty;
            string PackagesFilter = string.Empty;

            string SelectCommand = string.Empty;

            DataTable DT = new DataTable();

            if (FactoryID == -1)
            {
                DecorFactoryFilter = " DecorOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = " WHERE MainOrderID = -1";
                PackagesFilter = " WHERE MainOrderID = -1";
            }
            if (FactoryID == 0)
            {
                MainOrdersFilter = " WHERE ((ProfilProductionDate < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") +
                    "') OR (TPSProductionDate < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "'))";
                PackagesFilter = " AND DecorOrderID IN (SELECT DISTINCT OrderID FROM PackageDetails WHERE PackageID IN (SELECT PackageID FROM Packages WHERE ProductType = 1 AND (PackingDateTime > '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "' OR PackageStatusID = 1)))";
            }
            if (FactoryID == 1)
            {
                DecorFactoryFilter = " DecorOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = " WHERE (ProfilProductionDate < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "')";
                PackagesFilter = " AND DecorOrderID IN (SELECT DISTINCT OrderID FROM PackageDetails WHERE PackageID IN (SELECT PackageID FROM Packages WHERE FactoryID = " + FactoryID + " AND ProductType = 1 AND (PackingDateTime > '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "' OR PackageStatusID = 1)))";
            }
            if (FactoryID == 2)
            {
                DecorFactoryFilter = " DecorOrders.FactoryID = " + FactoryID + " AND ";
                MainOrdersFilter = " WHERE (TPSProductionDate < '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "')";
                PackagesFilter = " AND DecorOrderID IN (SELECT DISTINCT OrderID FROM PackageDetails WHERE PackageID IN (SELECT PackageID FROM Packages WHERE FactoryID = " + FactoryID + " AND ProductType = 1 AND (PackingDateTime > '" + FixingDate.ToString("yyyy-MM-dd HH:mm:ss") + "' OR PackageStatusID = 1)))";
            }

            //decor
            SelectCommand = "SELECT DecorOrders.ProductID, DecorOrders.DecorID, DecorOrders.ColorID,DecorOrders.PatinaID," +
                " DecorOrders.Height, DecorOrders.Length, DecorOrders.Width, DecorOrders.Count, DecorOrders.Cost, DecorOrders.DecorConfigID, " +
                " MeasureID FROM DecorOrders" +
                " INNER JOIN infiniu2_catalog.dbo.DecorConfig" +
                " ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                " WHERE " + DecorFactoryFilter + " DecorOrders.MainOrderID IN (SELECT MainOrderID FROM MainOrders " + MainOrdersFilter + ")" + PackagesFilter;

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.ZOVOrdersConnectionString))
            {
                DA.Fill(DT);
            }

            return DT;
        }

        #endregion ZOV

        public void ClearOrders()
        {
            FrontsOrdersDT.Clear();
            DecorOrdersDT.Clear();
            FrontsSummaryDT.Clear();
            DecorProductsSummaryDT.Clear();
            DecorItemsSummaryDT.Clear();
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
                    NewRow["Width"] = 0;
                    NewRow["Count"] = FrontCount;
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

        public bool IsZOVSelect()
        {
            for (int i = 0; i < ClientGroupsDT.Rows.Count; i++)
            {
                if (Convert.ToBoolean(ClientGroupsDT.Rows[i]["Check"]))
                    return true;
            }

            return false;
        }

        public ArrayList IsMarketingSelect()
        {
            ArrayList ClientGroupIDs = new ArrayList();

            for (int i = 0; i < ClientGroupsDT.Rows.Count; i++)
            {
                if (!Convert.ToBoolean(ClientGroupsDT.Rows[i]["Check"]))
                    continue;

                ClientGroupIDs.Add(Convert.ToInt32(ClientGroupsDT.Rows[i]["ClientGroupID"]));
            }

            return ClientGroupIDs;
        }
    }
}