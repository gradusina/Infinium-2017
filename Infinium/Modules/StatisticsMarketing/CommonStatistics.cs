using System;
using System.Collections;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace Infinium.Modules.StatisticsMarketing
{
    public class CommonStatistics : IAllFrontParameterName
    {
        private int CurrentManagerID = -1;

        public DataTable FrontsOrdersDataTable = null;
        public DataTable DecorOrdersDataTable = null;

        private DataTable TempFrontsDataTable = null;
        private DataTable TempDecorDataTable = null;
        private DataTable ClientsDataTable = null;
        private DataTable ClientGroupsDataTable = null;
        private DataTable ManagersDataTable = null;

        private DataTable FrontsSummaryDataTable = null;
        private DataTable FrameColorsSummaryDataTable = null;
        private DataTable TechnoColorsSummaryDataTable = null;
        private DataTable InsetTypesSummaryDataTable = null;
        private DataTable InsetColorsSummaryDataTable = null;
        private DataTable TechnoInsetTypesSummaryDataTable = null;
        private DataTable TechnoInsetColorsSummaryDataTable = null;
        private DataTable SizesSummaryDataTable = null;
        private DataTable DecorProductsSummaryDataTable = null;
        private DataTable DecorItemsSummaryDataTable = null;
        private DataTable DecorColorsSummaryDataTable = null;
        private DataTable DecorSizesSummaryDataTable = null;
        private DataTable DecorConfigDataTable = null;

        public DataTable FrontsDataTable = null;
        public DataTable FrameColorsDataTable = null;
        public DataTable PatinaDataTable = null;
        private DataTable PatinaRALDataTable = null;
        public DataTable InsetTypesDataTable = null;
        public DataTable InsetColorsDataTable = null;
        private DataTable DecorProductsDataTable = null;
        private DataTable DecorDataTable = null;
        private DataTable TechStoreDataTable = null;

        public BindingSource ClientsBindingSource = null;
        public BindingSource ClientGroupsBindingSource = null;
        public BindingSource ManagersBindingSource = null;
        public BindingSource FrontsSummaryBindingSource = null;
        public BindingSource FrameColorsSummaryBindingSource = null;
        public BindingSource TechnoColorsSummaryBindingSource = null;
        public BindingSource InsetTypesSummaryBindingSource = null;
        public BindingSource InsetColorsSummaryBindingSource = null;
        public BindingSource TechnoInsetTypesSummaryBindingSource = null;
        public BindingSource TechnoInsetColorsSummaryBindingSource = null;
        public BindingSource SizesSummaryBindingSource = null;
        public BindingSource DecorProductsSummaryBindingSource = null;
        public BindingSource DecorItemsSummaryBindingSource = null;
        public BindingSource DecorColorsSummaryBindingSource = null;
        public BindingSource DecorSizesSummaryBindingSource = null;

        public CommonStatistics()
        {
            Initialize();
        }

        private void Create()
        {
            TempFrontsDataTable = new DataTable();
            TempDecorDataTable = new DataTable();

            ClientsDataTable = new DataTable();
            ClientGroupsDataTable = new DataTable();
            ManagersDataTable = new DataTable();

            DecorDataTable = new DataTable();
            DecorProductsDataTable = new DataTable();
            DecorConfigDataTable = new DataTable();
            FrontsDataTable = new DataTable();
            FrameColorsDataTable = new DataTable();
            PatinaDataTable = new DataTable();
            InsetTypesDataTable = new DataTable();
            InsetColorsDataTable = new DataTable();

            FrontsOrdersDataTable = new DataTable();

            DecorOrdersDataTable = new DataTable();

            FrontsSummaryDataTable = new DataTable();
            FrontsSummaryDataTable.Columns.Add(new DataColumn(("Front"), System.Type.GetType("System.String")));
            FrontsSummaryDataTable.Columns.Add(new DataColumn(("FrontID"), System.Type.GetType("System.Int32")));
            FrontsSummaryDataTable.Columns.Add(new DataColumn(("Width"), System.Type.GetType("System.Int32")));
            FrontsSummaryDataTable.Columns.Add(new DataColumn(("Square"), System.Type.GetType("System.Decimal")));
            FrontsSummaryDataTable.Columns.Add(new DataColumn("Cost", Type.GetType("System.Decimal")));
            FrontsSummaryDataTable.Columns.Add(new DataColumn(("Count"), System.Type.GetType("System.Int32")));

            FrameColorsSummaryDataTable = new DataTable();
            FrameColorsSummaryDataTable.Columns.Add(new DataColumn(("FrameColor"), System.Type.GetType("System.String")));
            FrameColorsSummaryDataTable.Columns.Add(new DataColumn(("FrontID"), System.Type.GetType("System.Int32")));
            FrameColorsSummaryDataTable.Columns.Add(new DataColumn(("ColorID"), System.Type.GetType("System.Int32")));
            FrameColorsSummaryDataTable.Columns.Add(new DataColumn(("PatinaID"), System.Type.GetType("System.Int32")));
            FrameColorsSummaryDataTable.Columns.Add(new DataColumn(("Width"), System.Type.GetType("System.Int32")));
            FrameColorsSummaryDataTable.Columns.Add(new DataColumn(("Square"), System.Type.GetType("System.Decimal")));
            FrameColorsSummaryDataTable.Columns.Add(new DataColumn("Cost", Type.GetType("System.Decimal")));
            FrameColorsSummaryDataTable.Columns.Add(new DataColumn(("Count"), System.Type.GetType("System.Int32")));

            TechnoColorsSummaryDataTable = new DataTable();
            TechnoColorsSummaryDataTable.Columns.Add(new DataColumn(("TechnoColor"), System.Type.GetType("System.String")));
            TechnoColorsSummaryDataTable.Columns.Add(new DataColumn(("FrontID"), System.Type.GetType("System.Int32")));
            TechnoColorsSummaryDataTable.Columns.Add(new DataColumn(("ColorID"), System.Type.GetType("System.Int32")));
            TechnoColorsSummaryDataTable.Columns.Add(new DataColumn(("TechnoColorID"), System.Type.GetType("System.Int32")));
            TechnoColorsSummaryDataTable.Columns.Add(new DataColumn(("PatinaID"), System.Type.GetType("System.Int32")));
            TechnoColorsSummaryDataTable.Columns.Add(new DataColumn(("Width"), System.Type.GetType("System.Int32")));
            TechnoColorsSummaryDataTable.Columns.Add(new DataColumn(("Square"), System.Type.GetType("System.Decimal")));
            TechnoColorsSummaryDataTable.Columns.Add(new DataColumn(("Count"), System.Type.GetType("System.Int32")));

            InsetTypesSummaryDataTable = new DataTable();
            InsetTypesSummaryDataTable.Columns.Add(new DataColumn(("InsetType"), System.Type.GetType("System.String")));
            InsetTypesSummaryDataTable.Columns.Add(new DataColumn(("FrontID"), System.Type.GetType("System.Int32")));
            InsetTypesSummaryDataTable.Columns.Add(new DataColumn(("ColorID"), System.Type.GetType("System.Int32")));
            InsetTypesSummaryDataTable.Columns.Add(new DataColumn(("TechnoColorID"), System.Type.GetType("System.Int32")));
            InsetTypesSummaryDataTable.Columns.Add(new DataColumn(("PatinaID"), System.Type.GetType("System.Int32")));
            InsetTypesSummaryDataTable.Columns.Add(new DataColumn(("InsetTypeID"), System.Type.GetType("System.Int32")));
            InsetTypesSummaryDataTable.Columns.Add(new DataColumn(("Width"), System.Type.GetType("System.Int32")));
            InsetTypesSummaryDataTable.Columns.Add(new DataColumn(("Square"), System.Type.GetType("System.Decimal")));
            InsetTypesSummaryDataTable.Columns.Add(new DataColumn(("Count"), System.Type.GetType("System.Int32")));

            InsetColorsSummaryDataTable = new DataTable();
            InsetColorsSummaryDataTable.Columns.Add(new DataColumn(("InsetColor"), System.Type.GetType("System.String")));
            InsetColorsSummaryDataTable.Columns.Add(new DataColumn(("FrontID"), System.Type.GetType("System.Int32")));
            InsetColorsSummaryDataTable.Columns.Add(new DataColumn(("ColorID"), System.Type.GetType("System.Int32")));
            InsetColorsSummaryDataTable.Columns.Add(new DataColumn(("TechnoColorID"), System.Type.GetType("System.Int32")));
            InsetColorsSummaryDataTable.Columns.Add(new DataColumn(("PatinaID"), System.Type.GetType("System.Int32")));
            InsetColorsSummaryDataTable.Columns.Add(new DataColumn(("InsetColorID"), System.Type.GetType("System.Int32")));
            InsetColorsSummaryDataTable.Columns.Add(new DataColumn(("InsetTypeID"), System.Type.GetType("System.Int32")));
            InsetColorsSummaryDataTable.Columns.Add(new DataColumn(("Width"), System.Type.GetType("System.Int32")));
            InsetColorsSummaryDataTable.Columns.Add(new DataColumn(("Square"), System.Type.GetType("System.Decimal")));
            InsetColorsSummaryDataTable.Columns.Add(new DataColumn(("Count"), System.Type.GetType("System.Int32")));

            TechnoInsetTypesSummaryDataTable = new DataTable();
            TechnoInsetTypesSummaryDataTable.Columns.Add(new DataColumn(("TechnoInsetType"), System.Type.GetType("System.String")));
            TechnoInsetTypesSummaryDataTable.Columns.Add(new DataColumn(("FrontID"), System.Type.GetType("System.Int32")));
            TechnoInsetTypesSummaryDataTable.Columns.Add(new DataColumn(("ColorID"), System.Type.GetType("System.Int32")));
            TechnoInsetTypesSummaryDataTable.Columns.Add(new DataColumn(("TechnoColorID"), System.Type.GetType("System.Int32")));
            TechnoInsetTypesSummaryDataTable.Columns.Add(new DataColumn(("PatinaID"), System.Type.GetType("System.Int32")));
            TechnoInsetTypesSummaryDataTable.Columns.Add(new DataColumn(("InsetTypeID"), System.Type.GetType("System.Int32")));
            TechnoInsetTypesSummaryDataTable.Columns.Add(new DataColumn(("InsetColorID"), System.Type.GetType("System.Int32")));
            TechnoInsetTypesSummaryDataTable.Columns.Add(new DataColumn(("TechnoInsetTypeID"), System.Type.GetType("System.Int32")));
            TechnoInsetTypesSummaryDataTable.Columns.Add(new DataColumn(("Width"), System.Type.GetType("System.Int32")));
            TechnoInsetTypesSummaryDataTable.Columns.Add(new DataColumn(("Square"), System.Type.GetType("System.Decimal")));
            TechnoInsetTypesSummaryDataTable.Columns.Add(new DataColumn(("Count"), System.Type.GetType("System.Int32")));

            TechnoInsetColorsSummaryDataTable = new DataTable();
            TechnoInsetColorsSummaryDataTable.Columns.Add(new DataColumn(("TechnoInsetColor"), System.Type.GetType("System.String")));
            TechnoInsetColorsSummaryDataTable.Columns.Add(new DataColumn(("FrontID"), System.Type.GetType("System.Int32")));
            TechnoInsetColorsSummaryDataTable.Columns.Add(new DataColumn(("ColorID"), System.Type.GetType("System.Int32")));
            TechnoInsetColorsSummaryDataTable.Columns.Add(new DataColumn(("TechnoColorID"), System.Type.GetType("System.Int32")));
            TechnoInsetColorsSummaryDataTable.Columns.Add(new DataColumn(("PatinaID"), System.Type.GetType("System.Int32")));
            TechnoInsetColorsSummaryDataTable.Columns.Add(new DataColumn(("InsetColorID"), System.Type.GetType("System.Int32")));
            TechnoInsetColorsSummaryDataTable.Columns.Add(new DataColumn(("InsetTypeID"), System.Type.GetType("System.Int32")));
            TechnoInsetColorsSummaryDataTable.Columns.Add(new DataColumn(("TechnoInsetTypeID"), System.Type.GetType("System.Int32")));
            TechnoInsetColorsSummaryDataTable.Columns.Add(new DataColumn(("TechnoInsetColorID"), System.Type.GetType("System.Int32")));
            TechnoInsetColorsSummaryDataTable.Columns.Add(new DataColumn(("Width"), System.Type.GetType("System.Int32")));
            TechnoInsetColorsSummaryDataTable.Columns.Add(new DataColumn(("Square"), System.Type.GetType("System.Decimal")));
            TechnoInsetColorsSummaryDataTable.Columns.Add(new DataColumn(("Count"), System.Type.GetType("System.Int32")));

            SizesSummaryDataTable = new DataTable();
            SizesSummaryDataTable.Columns.Add(new DataColumn(("Size"), System.Type.GetType("System.String")));
            SizesSummaryDataTable.Columns.Add(new DataColumn(("FrontID"), System.Type.GetType("System.Int32")));
            SizesSummaryDataTable.Columns.Add(new DataColumn(("ColorID"), System.Type.GetType("System.Int32")));
            SizesSummaryDataTable.Columns.Add(new DataColumn(("PatinaID"), System.Type.GetType("System.Int32")));
            SizesSummaryDataTable.Columns.Add(new DataColumn(("TechnoColorID"), System.Type.GetType("System.Int32")));
            SizesSummaryDataTable.Columns.Add(new DataColumn(("InsetTypeID"), System.Type.GetType("System.Int32")));
            SizesSummaryDataTable.Columns.Add(new DataColumn(("InsetColorID"), System.Type.GetType("System.Int32")));
            SizesSummaryDataTable.Columns.Add(new DataColumn(("TechnoInsetTypeID"), System.Type.GetType("System.Int32")));
            SizesSummaryDataTable.Columns.Add(new DataColumn(("TechnoInsetColorID"), System.Type.GetType("System.Int32")));
            SizesSummaryDataTable.Columns.Add(new DataColumn(("Square"), System.Type.GetType("System.Decimal")));
            SizesSummaryDataTable.Columns.Add(new DataColumn(("Width"), System.Type.GetType("System.Int32")));
            SizesSummaryDataTable.Columns.Add(new DataColumn(("Count"), System.Type.GetType("System.Int32")));
            SizesSummaryDataTable.Columns.Add(new DataColumn(("Height"), System.Type.GetType("System.Int32")));

            DecorProductsSummaryDataTable = new DataTable();
            DecorProductsSummaryDataTable.Columns.Add(new DataColumn(("DecorProduct"), System.Type.GetType("System.String")));
            DecorProductsSummaryDataTable.Columns.Add(new DataColumn(("ProductID"), System.Type.GetType("System.Int32")));
            DecorProductsSummaryDataTable.Columns.Add(new DataColumn(("MeasureID"), System.Type.GetType("System.Int32")));
            DecorProductsSummaryDataTable.Columns.Add(new DataColumn("Cost", Type.GetType("System.Decimal")));
            DecorProductsSummaryDataTable.Columns.Add(new DataColumn(("Count"), System.Type.GetType("System.Decimal")));
            DecorProductsSummaryDataTable.Columns.Add(new DataColumn(("Measure"), System.Type.GetType("System.String")));

            DecorItemsSummaryDataTable = new DataTable();
            DecorItemsSummaryDataTable.Columns.Add(new DataColumn(("DecorItem"), System.Type.GetType("System.String")));
            DecorItemsSummaryDataTable.Columns.Add(new DataColumn(("ProductID"), System.Type.GetType("System.Int32")));
            DecorItemsSummaryDataTable.Columns.Add(new DataColumn(("MeasureID"), System.Type.GetType("System.Int32")));
            DecorItemsSummaryDataTable.Columns.Add(new DataColumn(("DecorID"), System.Type.GetType("System.Int32")));
            DecorItemsSummaryDataTable.Columns.Add(new DataColumn("Cost", Type.GetType("System.Decimal")));
            DecorItemsSummaryDataTable.Columns.Add(new DataColumn(("Count"), System.Type.GetType("System.Decimal")));
            DecorItemsSummaryDataTable.Columns.Add(new DataColumn(("Measure"), System.Type.GetType("System.String")));

            DecorColorsSummaryDataTable = new DataTable();
            DecorColorsSummaryDataTable.Columns.Add(new DataColumn(("Color"), System.Type.GetType("System.String")));
            DecorColorsSummaryDataTable.Columns.Add(new DataColumn(("ProductID"), System.Type.GetType("System.Int32")));
            DecorColorsSummaryDataTable.Columns.Add(new DataColumn(("MeasureID"), System.Type.GetType("System.Int32")));
            DecorColorsSummaryDataTable.Columns.Add(new DataColumn(("ColorID"), System.Type.GetType("System.Int32")));
            DecorColorsSummaryDataTable.Columns.Add(new DataColumn(("DecorID"), System.Type.GetType("System.Int32")));
            DecorColorsSummaryDataTable.Columns.Add(new DataColumn("Cost", Type.GetType("System.Decimal")));
            DecorColorsSummaryDataTable.Columns.Add(new DataColumn(("Count"), System.Type.GetType("System.Decimal")));
            DecorColorsSummaryDataTable.Columns.Add(new DataColumn(("Measure"), System.Type.GetType("System.String")));

            DecorSizesSummaryDataTable = new DataTable();
            DecorSizesSummaryDataTable.Columns.Add(new DataColumn(("ProductID"), System.Type.GetType("System.Int32")));
            DecorSizesSummaryDataTable.Columns.Add(new DataColumn(("DecorID"), System.Type.GetType("System.Int32")));
            DecorSizesSummaryDataTable.Columns.Add(new DataColumn(("ColorID"), System.Type.GetType("System.Int32")));
            DecorSizesSummaryDataTable.Columns.Add(new DataColumn(("MeasureID"), System.Type.GetType("System.Int32")));
            DecorSizesSummaryDataTable.Columns.Add(new DataColumn(("Size"), System.Type.GetType("System.String")));
            DecorSizesSummaryDataTable.Columns.Add(new DataColumn(("Length"), System.Type.GetType("System.Int32")));
            DecorSizesSummaryDataTable.Columns.Add(new DataColumn(("Height"), System.Type.GetType("System.Int32")));
            DecorSizesSummaryDataTable.Columns.Add(new DataColumn(("Width"), System.Type.GetType("System.Int32")));
            DecorSizesSummaryDataTable.Columns.Add(new DataColumn("Cost", Type.GetType("System.Decimal")));
            DecorSizesSummaryDataTable.Columns.Add(new DataColumn(("Count"), System.Type.GetType("System.Decimal")));
            DecorSizesSummaryDataTable.Columns.Add(new DataColumn(("Measure"), System.Type.GetType("System.String")));

            TempFrontsDataTable = new DataTable();
            TempFrontsDataTable.Columns.Add(new DataColumn("Front", Type.GetType("System.String")));
            TempFrontsDataTable.Columns.Add(new DataColumn("Count", Type.GetType("System.Decimal")));
            TempFrontsDataTable.Columns.Add(new DataColumn("Measure", Type.GetType("System.String")));
            TempFrontsDataTable.Columns.Add(new DataColumn("Empty", Type.GetType("System.String")));
            TempFrontsDataTable.Columns.Add(new DataColumn("Color", Type.GetType("System.String")));
            TempFrontsDataTable.Columns.Add(new DataColumn("ColorCount", Type.GetType("System.Decimal")));
            TempFrontsDataTable.Columns.Add(new DataColumn("ColorMeasure", Type.GetType("System.String")));

            FrontsSummaryBindingSource = new BindingSource();
            FrameColorsSummaryBindingSource = new BindingSource();
            TechnoColorsSummaryBindingSource = new BindingSource();
            InsetTypesSummaryBindingSource = new BindingSource();
            InsetColorsSummaryBindingSource = new BindingSource();
            TechnoInsetTypesSummaryBindingSource = new BindingSource();
            TechnoInsetColorsSummaryBindingSource = new BindingSource();
            SizesSummaryBindingSource = new BindingSource();

            DecorProductsSummaryBindingSource = new BindingSource();
            DecorItemsSummaryBindingSource = new BindingSource();
            DecorColorsSummaryBindingSource = new BindingSource();
            DecorSizesSummaryBindingSource = new BindingSource();
        }

        private void GetColorsDT()
        {
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

        private void Fill()
        {
            string SelectCommand = @"SELECT TechStoreID AS FrontID, TechStoreName AS FrontName FROM TechStore
                WHERE TechStoreID IN (SELECT FrontID FROM FrontsConfig WHERE Enabled = 1)
                ORDER BY TechStoreName";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(FrontsDataTable);
            }
            GetColorsDT();
            SelectCommand = @"SELECT * FROM Patina";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(PatinaDataTable);
            }
            PatinaRALDataTable = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM PatinaRAL WHERE Enabled=1",
                ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(PatinaRALDataTable);
            }
            foreach (DataRow item in PatinaRALDataTable.Rows)
            {
                DataRow NewRow = PatinaDataTable.NewRow();
                NewRow["PatinaID"] = item["PatinaRALID"];
                NewRow["PatinaName"] = item["PatinaRAL"];
                NewRow["DisplayName"] = item["DisplayName"];
                PatinaDataTable.Rows.Add(NewRow);
            }
            SelectCommand = @"SELECT * FROM InsetTypes";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(InsetTypesDataTable);
            }
            SelectCommand = @"SELECT * FROM InsetColors";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(InsetColorsDataTable);
            }
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
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT ManagerID, Name FROM ClientsManagers ORDER BY Name",
                ConnectionStrings.MarketingReferenceConnectionString))
            {
                DA.Fill(ManagersDataTable);
            }
            ManagersDataTable.Columns.Add(new DataColumn("Check", Type.GetType("System.Boolean")));
            for (int i = 0; i < ManagersDataTable.Rows.Count; i++)
            {
                ManagersDataTable.Rows[i]["Check"] = false;
            }

            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM ClientGroups ORDER BY ClientGroupName",
                ConnectionStrings.MarketingReferenceConnectionString))
            {
                DA.Fill(ClientGroupsDataTable);
            }
            ClientGroupsDataTable.Columns.Add(new DataColumn("Check", Type.GetType("System.Boolean")));
            for (int i = 0; i < ClientGroupsDataTable.Rows.Count; i++)
            {
                ClientGroupsDataTable.Rows[i]["Check"] = true;
            }

            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT ClientID, ClientName, ManagerID FROM Clients", ConnectionStrings.MarketingReferenceConnectionString))
            {
                DA.Fill(ClientsDataTable);
            }
            ClientsDataTable.Columns.Add(new DataColumn("Check", Type.GetType("System.Boolean")));
            for (int i = 0; i < ClientsDataTable.Rows.Count; i++)
            {
                ClientsDataTable.Rows[i]["Check"] = false;
            }

            //using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM DecorConfig", ConnectionStrings.CatalogConnectionString))
            //{
            //    DA.Fill(DecorConfigDataTable);
            //}
            DecorConfigDataTable = TablesManager.DecorConfigDataTableAll;

            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT TOP 0  FrontID, ColorID, PatinaID, InsetTypeID, InsetColorID, TechnoColorID, TechnoInsetTypeID, TechnoInsetColorID, Height, Width, Count, Square, Cost FROM FrontsOrders",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                TempFrontsDataTable.Clear();
                DA.Fill(TempFrontsDataTable);
            }

            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT TOP 0  FrontID, ColorID, PatinaID, InsetTypeID, InsetColorID, TechnoColorID, TechnoInsetTypeID, TechnoInsetColorID, Height, Width, Count, Square, Cost FROM FrontsOrders",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                FrontsOrdersDataTable.Clear();
                DA.Fill(FrontsOrdersDataTable);
            }

            TechStoreDataTable = new DataTable();
            //using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM TechStore",
            //    ConnectionStrings.CatalogConnectionString))
            //{
            //    DA.Fill(TechStoreDataTable);
            //}
            TechStoreDataTable = TablesManager.TechStoreDataTable;

            GetFronts();
            GetFrameColors();
            GetTechnoColors();
            GetInsetTypes();
            GetInsetColors();
            GetTechnoInsetTypes();
            GetTechnoInsetColors();
            GetSizes();

            //decor
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT TOP 0 DecorOrders.ProductID, DecorOrders.DecorID, DecorOrders.ColorID,DecorOrders.PatinaID," +
                " DecorOrders.Height, DecorOrders.Length, DecorOrders.Width, DecorOrders.Count, DecorOrders.Cost, DecorOrders.DecorConfigID, " +
                " MeasureID FROM DecorOrders" +
                " INNER JOIN infiniu2_catalog.dbo.DecorConfig" +
                " ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID", ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(TempDecorDataTable);
            }

            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT TOP 0 DecorOrders.ProductID, DecorOrders.DecorID, DecorOrders.ColorID,DecorOrders.PatinaID," +
                " DecorOrders.Height, DecorOrders.Length, DecorOrders.Width, DecorOrders.Count, DecorOrders.Cost, DecorOrders.DecorConfigID, " +
                " MeasureID FROM DecorOrders" +
                " INNER JOIN infiniu2_catalog.dbo.DecorConfig" +
                " ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID", ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(DecorOrdersDataTable);
            }

            GetDecorProducts();
            GetDecorItems();
            GetDecorColors();
            GetDecorSizes();
        }

        public void FilterByPlanDispatch(
            DateTime DateFrom, DateTime DateTo,
            int FactoryID, bool IsSample, bool IsNotSample, bool IsTransport)
        {
            string MClientFilter = string.Empty;

            string MFrontsFactoryFilter = string.Empty;
            string MDecorFactoryFilter = string.Empty;

            string MarketingSelectCommand = string.Empty;

            ArrayList MClients = SelectedMarketingClients;
            ArrayList MClientGroups = SelectedMarketingClientGroups;

            if (FactoryID != 0)
            {
                MFrontsFactoryFilter = " FrontsOrders.FactoryID = " + FactoryID + " AND ";
                MDecorFactoryFilter = " DecorOrders.FactoryID = " + FactoryID + " AND ";
            }

            if (FactoryID == 0)
                MClientFilter = " WHERE ((CAST(ProfilDispatchDate AS DATE) >= '" + DateFrom.ToString("yyyy-MM-dd") +
                    "' AND CAST(ProfilDispatchDate AS DATE) <= '" + DateTo.ToString("yyyy-MM-dd") + "')" +
                    " OR (CAST(TPSDispatchDate AS DATE) >= '" + DateFrom.ToString("yyyy-MM-dd") +
                    "' AND CAST(TPSDispatchDate AS DATE) <= '" + DateTo.ToString("yyyy-MM-dd") + "')) ";

            if (FactoryID == 1)
                MClientFilter = " WHERE CAST(ProfilDispatchDate AS DATE) >= '" + DateFrom.ToString("yyyy-MM-dd") +
                    "' AND CAST(ProfilDispatchDate AS DATE) <= '" + DateTo.ToString("yyyy-MM-dd") + "' ";

            if (FactoryID == 2)
                MClientFilter = " WHERE CAST(TPSDispatchDate AS DATE) >= '" + DateFrom.ToString("yyyy-MM-dd") +
                    "' AND CAST(TPSDispatchDate AS DATE) <= '" + DateTo.ToString("yyyy-MM-dd") + "' ";

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

            string MFSampleFilter = string.Empty;
            string MDSampleFilter = string.Empty;

            if (!IsSample || !IsNotSample)
            {
                if (IsSample)
                    MFSampleFilter = " FrontsOrders.IsSample = 1 AND";
                if (IsNotSample)
                    MFSampleFilter = " FrontsOrders.IsSample = 0 AND";
                if (IsSample)
                    MDSampleFilter = " (DecorOrders.IsSample = 1 OR (DecorOrders.ProductID=42 AND DecorOrders.IsSample = 0)) AND";
                if (IsNotSample)
                    MDSampleFilter = " DecorOrders.IsSample = 0 AND";
            }
            string TransportFilter = " ";
            if (IsTransport)
                TransportFilter = " AND MegaOrders.TransportCost<>0 ";
            MarketingSelectCommand = "SELECT MegaOrders.MegaOrderID, MegaOrders.OrderNumber, FrontsOrders.MainOrderID, FrontsOrders.FrontID, FrontsOrders.PatinaID, FrontsOrders.ColorID, FrontsOrders.InsetTypeID," +
                " FrontsOrders.InsetColorID, FrontsOrders.TechnoColorID, FrontsOrders.TechnoInsetTypeID, FrontsOrders.TechnoInsetColorID, FrontsOrders.Height, FrontsOrders.Width," +
                " FrontsOrders.Count, FrontsOrders.IsSample, FrontsOrders.Square, FrontsOrders.Cost, MeasureID, ClientID, JoinMainOrders.ZOVClientID FROM FrontsOrders" +
                " INNER JOIN MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID" +
                " INNER JOIN MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID" + TransportFilter +
                " LEFT OUTER JOIN JoinMainOrders ON FrontsOrders.MainOrderID = JoinMainOrders.MarketMainOrderID" +
                " INNER JOIN infiniu2_catalog.dbo.FrontsConfig" +
                " ON FrontsOrders.FrontConfigID=infiniu2_catalog.dbo.FrontsConfig.FrontConfigID" +
                " WHERE " + MFrontsFactoryFilter + MFSampleFilter + " FrontsOrders.MainOrderID IN (SELECT MainOrderID FROM MainOrders WHERE MegaOrderID IN" +
                " (SELECT MegaOrderID FROM MegaOrders " + MClientFilter + " ))";

            using (SqlDataAdapter DA = new SqlDataAdapter(MarketingSelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                FrontsOrdersDataTable.Clear();
                DA.Fill(FrontsOrdersDataTable);
            }

            GetFronts();
            GetFrameColors();
            GetTechnoColors();
            GetInsetTypes();
            GetInsetColors();
            GetTechnoInsetTypes();
            GetTechnoInsetColors();
            GetSizes();

            //decor
            MarketingSelectCommand = "SELECT MegaOrders.MegaOrderID, MegaOrders.OrderNumber, DecorOrders.MainOrderID, DecorOrders.ProductID, DecorOrders.DecorID, DecorOrders.ColorID,DecorOrders.PatinaID," +
                " DecorOrders.Height, DecorOrders.Length, DecorOrders.Width, DecorOrders.Count, DecorOrders.IsSample, DecorOrders.Cost, DecorOrders.DecorConfigID, " +
                " MeasureID, ClientID, JoinMainOrders.ZOVClientID FROM DecorOrders" +
                " INNER JOIN MainOrders ON DecorOrders.MainOrderID = MainOrders.MainOrderID" +
                " INNER JOIN MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID" + TransportFilter +
                " LEFT OUTER JOIN JoinMainOrders ON DecorOrders.MainOrderID = JoinMainOrders.MarketMainOrderID" +
                " INNER JOIN infiniu2_catalog.dbo.DecorConfig" +
                " ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                " WHERE " + MDecorFactoryFilter + MDSampleFilter + " DecorOrders.MainOrderID IN (SELECT MainOrderID FROM MainOrders WHERE MegaOrderID IN" +
                " (SELECT MegaOrderID FROM MegaOrders " + MClientFilter + " ))";

            using (SqlDataAdapter DA = new SqlDataAdapter(MarketingSelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DecorOrdersDataTable.Clear();
                DA.Fill(DecorOrdersDataTable);
            }

            GetDecorProducts();
            GetDecorItems();
            GetDecorColors();
            GetDecorSizes();
        }

        public void FilterByOrderDate(
            DateTime DateFrom, DateTime DateTo,
            int FactoryID, bool IsSample, bool IsNotSample, bool IsTransport)
        {
            string MClientFilter = string.Empty;

            string MFilter = string.Empty;
            string MFrontsFactoryFilter = string.Empty;
            string MDecorFactoryFilter = string.Empty;

            string MarketingSelectCommand = string.Empty;

            ArrayList MClients = SelectedMarketingClients;
            ArrayList MClientGroups = SelectedMarketingClientGroups;

            if (FactoryID != 0)
            {
                MFrontsFactoryFilter = " NewFrontsOrders.FactoryID = " + FactoryID + " AND ";
                MDecorFactoryFilter = " NewDecorOrders.FactoryID = " + FactoryID + " AND ";
            }

            MFilter = " CAST(CreateDateTime AS DATE) >= '" + DateFrom.ToString("yyyy-MM-dd") +
                "' AND CAST(CreateDateTime AS DATE) <= '" + DateTo.ToString("yyyy-MM-dd") + "' AND ";

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

            string MFSampleFilter = string.Empty;
            string MDSampleFilter = string.Empty;

            if (!IsSample || !IsNotSample)
            {
                if (IsSample)
                    MFSampleFilter = " NewFrontsOrders.IsSample = 1 AND";
                if (IsNotSample)
                    MFSampleFilter = " NewFrontsOrders.IsSample = 0 AND";
                if (IsSample)
                    MDSampleFilter = " (NewDecorOrders.IsSample = 1 OR (NewDecorOrders.ProductID=42 AND NewDecorOrders.IsSample = 0)) AND";
                if (IsNotSample)
                    MDSampleFilter = " NewDecorOrders.IsSample = 0 AND";
            }
            string TransportFilter = " ";
            if (IsTransport)
                TransportFilter = " AND NewMegaOrders.TransportCost<>0 ";
            MarketingSelectCommand = "SELECT NewMegaOrders.MegaOrderID, NewMegaOrders.OrderNumber, NewFrontsOrders.MainOrderID, NewFrontsOrders.FrontID, NewFrontsOrders.PatinaID, NewFrontsOrders.ColorID, NewFrontsOrders.InsetTypeID," +
                " NewFrontsOrders.InsetColorID, NewFrontsOrders.TechnoColorID, NewFrontsOrders.TechnoInsetTypeID, NewFrontsOrders.TechnoInsetColorID, NewFrontsOrders.Height, NewFrontsOrders.Width," +
                " NewFrontsOrders.Count, NewFrontsOrders.IsSample, NewFrontsOrders.Square, NewFrontsOrders.Cost, MeasureID, ClientID, JoinMainOrders.ZOVClientID FROM NewFrontsOrders" +
                " INNER JOIN NewMainOrders ON NewFrontsOrders.MainOrderID = NewMainOrders.MainOrderID" +
                " INNER JOIN NewMegaOrders ON NewMainOrders.MegaOrderID = NewMegaOrders.MegaOrderID" + TransportFilter +
                " LEFT OUTER JOIN JoinMainOrders ON NewFrontsOrders.MainOrderID = JoinMainOrders.MarketMainOrderID" +
                " INNER JOIN infiniu2_catalog.dbo.FrontsConfig" +
                " ON NewFrontsOrders.FrontConfigID=infiniu2_catalog.dbo.FrontsConfig.FrontConfigID" +
                " WHERE " + MFilter + MFrontsFactoryFilter + MFSampleFilter + " NewFrontsOrders.MainOrderID IN (SELECT MainOrderID FROM NewMainOrders WHERE MegaOrderID IN" +
                " (SELECT MegaOrderID FROM NewMegaOrders " + MClientFilter + " ))";

            using (SqlDataAdapter DA = new SqlDataAdapter(MarketingSelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                FrontsOrdersDataTable.Clear();
                DA.Fill(FrontsOrdersDataTable);
            }

            GetFronts();
            GetFrameColors();
            GetTechnoColors();
            GetInsetTypes();
            GetInsetColors();
            GetTechnoInsetTypes();
            GetTechnoInsetColors();
            GetSizes();

            //decor
            MarketingSelectCommand = "SELECT NewMegaOrders.MegaOrderID, NewMegaOrders.OrderNumber, NewDecorOrders.MainOrderID, NewDecorOrders.ProductID, NewDecorOrders.DecorID, NewDecorOrders.ColorID,NewDecorOrders.PatinaID," +
                " NewDecorOrders.Height, NewDecorOrders.Length, NewDecorOrders.Width, NewDecorOrders.Count, NewDecorOrders.IsSample, NewDecorOrders.Cost, NewDecorOrders.DecorConfigID, " +
                " MeasureID, ClientID, JoinMainOrders.ZOVClientID FROM NewDecorOrders" +
                " INNER JOIN NewMainOrders ON NewDecorOrders.MainOrderID = NewMainOrders.MainOrderID" +
                " INNER JOIN NewMegaOrders ON NewMainOrders.MegaOrderID = NewMegaOrders.MegaOrderID" + TransportFilter +
                " LEFT OUTER JOIN JoinMainOrders ON NewDecorOrders.MainOrderID = JoinMainOrders.MarketMainOrderID" +
                " INNER JOIN infiniu2_catalog.dbo.DecorConfig" +
                " ON NewDecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                " WHERE " + MFilter + MDecorFactoryFilter + MDSampleFilter + " NewDecorOrders.MainOrderID IN (SELECT MainOrderID FROM NewMainOrders WHERE MegaOrderID IN" +
                " (SELECT MegaOrderID FROM NewMegaOrders " + MClientFilter + " ))";

            using (SqlDataAdapter DA = new SqlDataAdapter(MarketingSelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DecorOrdersDataTable.Clear();
                DA.Fill(DecorOrdersDataTable);
            }

            GetDecorProducts();
            GetDecorItems();
            GetDecorColors();
            GetDecorSizes();
        }

        public void FilterByOnProduction(
            DateTime DateFrom, DateTime DateTo,
            int FactoryID, bool IsSample, bool IsNotSample, bool IsTransport)
        {
            string MClientFilter = string.Empty;

            string MFrontsFactoryFilter = string.Empty;
            string MDecorFactoryFilter = string.Empty;

            string MMainOrdersFilter = string.Empty;

            string MarketingSelectCommand = string.Empty;

            ArrayList MClients = SelectedMarketingClients;
            ArrayList MClientGroups = SelectedMarketingClientGroups;

            if (FactoryID != 0)
            {
                MFrontsFactoryFilter = " FrontsOrders.FactoryID = " + FactoryID + " AND ";
                MDecorFactoryFilter = " DecorOrders.FactoryID = " + FactoryID + " AND ";
            }

            if (FactoryID == -1)
            {
                MMainOrdersFilter = " WHERE MainOrderID = -1";
            }
            if (FactoryID == 0)
            {
                MMainOrdersFilter = " WHERE ((CAST(ProfilOnProductionDate AS Date) >= '" + DateFrom.ToString("yyyy-MM-dd") + "' AND CAST(ProfilOnProductionDate AS Date) <= '" + DateTo.ToString("yyyy-MM-dd") +
                    "') OR (CAST(TPSOnProductionDate AS Date) >= '" + DateFrom.ToString("yyyy-MM-dd") + "' AND CAST(TPSOnProductionDate AS Date) <= '" + DateTo.ToString("yyyy-MM-dd") + "'))";
            }
            if (FactoryID == 1)
            {
                MMainOrdersFilter = " WHERE (CAST(ProfilOnProductionDate AS Date) >= '" + DateFrom.ToString("yyyy-MM-dd") + "' AND CAST(ProfilOnProductionDate AS Date) <= '" + DateTo.ToString("yyyy-MM-dd") + "')";
            }
            if (FactoryID == 2)
            {
                MMainOrdersFilter = " WHERE (CAST(TPSOnProductionDate AS Date) >= '" + DateFrom.ToString("yyyy-MM-dd") + "' AND CAST(TPSOnProductionDate AS Date) <= '" + DateTo.ToString("yyyy-MM-dd") + "')";
            }

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

            string MFSampleFilter = string.Empty;
            string MDSampleFilter = string.Empty;

            if (!IsSample || !IsNotSample)
            {
                if (IsSample)
                    MFSampleFilter = " FrontsOrders.IsSample = 1 AND";
                if (IsNotSample)
                    MFSampleFilter = " FrontsOrders.IsSample = 0 AND";
                if (IsSample)
                    MDSampleFilter = " (DecorOrders.IsSample = 1 OR (DecorOrders.ProductID=42 AND DecorOrders.IsSample = 0)) AND";
                if (IsNotSample)
                    MDSampleFilter = " DecorOrders.IsSample = 0 AND";
            }
            string TransportFilter = " ";
            if (IsTransport)
                TransportFilter = " AND MegaOrders.TransportCost<>0 ";
            MarketingSelectCommand = "SELECT MegaOrders.MegaOrderID, MegaOrders.OrderNumber, FrontsOrders.MainOrderID, FrontsOrders.FrontID, FrontsOrders.PatinaID, FrontsOrders.ColorID, FrontsOrders.InsetTypeID," +
                " FrontsOrders.InsetColorID, FrontsOrders.TechnoColorID, FrontsOrders.TechnoInsetTypeID, FrontsOrders.TechnoInsetColorID, FrontsOrders.Height, FrontsOrders.Width, " +
                " FrontsOrders.Count, FrontsOrders.IsSample, FrontsOrders.Square, FrontsOrders.Cost, MeasureID, ClientID, JoinMainOrders.ZOVClientID FROM FrontsOrders" +
                " INNER JOIN MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID" +
                " INNER JOIN MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID" + TransportFilter +
                " LEFT OUTER JOIN JoinMainOrders ON FrontsOrders.MainOrderID = JoinMainOrders.MarketMainOrderID" +
                " INNER JOIN infiniu2_catalog.dbo.FrontsConfig" +
                " ON FrontsOrders.FrontConfigID=infiniu2_catalog.dbo.FrontsConfig.FrontConfigID" +
                " WHERE " + MFrontsFactoryFilter + MFSampleFilter + " FrontsOrders.MainOrderID IN (SELECT MainOrderID FROM MainOrders " + MMainOrdersFilter +
                " AND MegaOrderID IN" +
                " (SELECT MegaOrderID FROM MegaOrders " + MClientFilter + " ))";

            using (SqlDataAdapter DA = new SqlDataAdapter(MarketingSelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                FrontsOrdersDataTable.Clear();
                DA.Fill(FrontsOrdersDataTable);
            }

            GetFronts();
            GetFrameColors();
            GetTechnoColors();
            GetInsetTypes();
            GetInsetColors();
            GetTechnoInsetTypes();
            GetTechnoInsetColors();
            GetSizes();

            //decor
            MarketingSelectCommand = "SELECT MegaOrders.MegaOrderID, MegaOrders.OrderNumber, DecorOrders.MainOrderID, DecorOrders.ProductID, DecorOrders.DecorID, DecorOrders.ColorID,DecorOrders.PatinaID," +
                " DecorOrders.Height, DecorOrders.Length, DecorOrders.Width, DecorOrders.Count, DecorOrders.IsSample, DecorOrders.Cost, DecorOrders.DecorConfigID, " +
                " MeasureID, ClientID, JoinMainOrders.ZOVClientID FROM DecorOrders" +
                " INNER JOIN MainOrders ON DecorOrders.MainOrderID = MainOrders.MainOrderID" +
                " INNER JOIN MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID" + TransportFilter +
                " LEFT OUTER JOIN JoinMainOrders ON DecorOrders.MainOrderID = JoinMainOrders.MarketMainOrderID" +
                " INNER JOIN infiniu2_catalog.dbo.DecorConfig" +
                " ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                " WHERE " + MDecorFactoryFilter + MDSampleFilter + " DecorOrders.MainOrderID IN (SELECT MainOrderID FROM MainOrders " + MMainOrdersFilter +
                " AND MegaOrderID IN" +
                " (SELECT MegaOrderID FROM MegaOrders " + MClientFilter + " ))";

            using (SqlDataAdapter DA = new SqlDataAdapter(MarketingSelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DecorOrdersDataTable.Clear();
                DA.Fill(DecorOrdersDataTable);
            }

            GetDecorProducts();
            GetDecorItems();
            GetDecorColors();
            GetDecorSizes();
        }

        public void FilterByConfirmDate(
            DateTime DateFrom, DateTime DateTo,
            int FactoryID, bool IsSample, bool IsNotSample, bool IsTransport)
        {
            string MClientFilter = string.Empty;
            string MFilter = string.Empty;

            string MFrontsFactoryFilter = string.Empty;
            string MDecorFactoryFilter = string.Empty;

            string MarketingSelectCommand = string.Empty;

            ArrayList MClients = SelectedMarketingClients;
            ArrayList MClientGroups = SelectedMarketingClientGroups;

            if (FactoryID != 0)
            {
                MFrontsFactoryFilter = " FrontsOrders.FactoryID = " + FactoryID + " AND ";
                MDecorFactoryFilter = " DecorOrders.FactoryID = " + FactoryID + " AND ";
            }

            MFilter = " WHERE CAST(ConfirmDateTime AS DATE) >= '" + DateFrom.ToString("yyyy-MM-dd") +
                "' AND CAST(ConfirmDateTime AS DATE) <= '" + DateTo.ToString("yyyy-MM-dd") + "' ";

            if (MClients.Count > 0)
            {
                if (MClientFilter.Length > 0)
                    MClientFilter += " AND ClientID IN (" + string.Join(",", MClients.OfType<Int32>().ToArray()) + ")";
                else
                    MClientFilter = " AND ClientID IN (" + string.Join(",", MClients.OfType<Int32>().ToArray()) + ")";
            }

            if (MClientGroups.Count > 0)
            {
                if (MClientFilter.Length > 0)
                    MClientFilter += " AND ClientID IN" +
                        " (SELECT ClientID FROM infiniu2_marketingreference.dbo.Clients" +
                        " WHERE ClientGroupID IN (" + string.Join(",", MClientGroups.OfType<Int32>().ToArray()) + "))";
                else
                    MClientFilter = " AND ClientID IN" +
                        " (SELECT ClientID FROM infiniu2_marketingreference.dbo.Clients" +
                        " WHERE ClientGroupID IN (" + string.Join(",", MClientGroups.OfType<Int32>().ToArray()) + "))";
            }

            if (MClients.Count < 1 && MClientGroups.Count < 1)
                MClientFilter = " AND ClientID = -1";

            string MFSampleFilter = string.Empty;
            string MDSampleFilter = string.Empty;

            if (!IsSample || !IsNotSample)
            {
                if (IsSample)
                    MFSampleFilter = " FrontsOrders.IsSample = 1 AND";
                if (IsNotSample)
                    MFSampleFilter = " FrontsOrders.IsSample = 0 AND";
                if (IsSample)
                    MDSampleFilter = " (DecorOrders.IsSample = 1 OR (DecorOrders.ProductID=42 AND DecorOrders.IsSample = 0)) AND";
                if (IsNotSample)
                    MDSampleFilter = " DecorOrders.IsSample = 0 AND";
            }
            string TransportFilter = " ";
            if (IsTransport)
                TransportFilter = " AND MegaOrders.TransportCost<>0 ";
            MarketingSelectCommand = "SELECT MegaOrders.ConfirmDateTime, MegaOrders.MegaOrderID, MegaOrders.OrderNumber, FrontsOrders.MainOrderID, FrontsOrders.FrontID, FrontsOrders.PatinaID, FrontsOrders.ColorID, FrontsOrders.InsetTypeID," +
                " FrontsOrders.InsetColorID, FrontsOrders.TechnoColorID, FrontsOrders.TechnoInsetTypeID, FrontsOrders.TechnoInsetColorID, FrontsOrders.Height, FrontsOrders.Width," +
                " FrontsOrders.Count, FrontsOrders.IsSample, FrontsOrders.Square, FrontsOrders.Cost, MeasureID, ClientID, JoinMainOrders.ZOVClientID FROM FrontsOrders" +
                " INNER JOIN MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID" +
                " INNER JOIN MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID" + TransportFilter +
                " LEFT OUTER JOIN JoinMainOrders ON FrontsOrders.MainOrderID = JoinMainOrders.MarketMainOrderID" +
                " INNER JOIN infiniu2_catalog.dbo.FrontsConfig" +
                " ON FrontsOrders.FrontConfigID=infiniu2_catalog.dbo.FrontsConfig.FrontConfigID" +
                " WHERE " + MFrontsFactoryFilter + MFSampleFilter + " FrontsOrders.MainOrderID IN (SELECT MainOrderID FROM MainOrders" +
                " WHERE MegaOrderID IN" +
                " (SELECT MegaOrderID FROM MegaOrders " + MFilter + MClientFilter + " ))";

            using (SqlDataAdapter DA = new SqlDataAdapter(MarketingSelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                FrontsOrdersDataTable.Clear();
                DA.Fill(FrontsOrdersDataTable);
            }

            GetFronts();
            GetFrameColors();
            GetTechnoColors();
            GetInsetTypes();
            GetInsetColors();
            GetTechnoInsetTypes();
            GetTechnoInsetColors();
            GetSizes();

            //decor
            MarketingSelectCommand = "SELECT MegaOrders.ConfirmDateTime, MegaOrders.MegaOrderID, MegaOrders.OrderNumber, DecorOrders.MainOrderID, DecorOrders.ProductID, DecorOrders.DecorID, DecorOrders.ColorID,DecorOrders.PatinaID," +
                " DecorOrders.Height, DecorOrders.Length, DecorOrders.Width, DecorOrders.Count, DecorOrders.IsSample, DecorOrders.Cost, DecorOrders.DecorConfigID, " +
                " MeasureID, ClientID, JoinMainOrders.ZOVClientID FROM DecorOrders" +
                " INNER JOIN MainOrders ON DecorOrders.MainOrderID = MainOrders.MainOrderID" +
                " INNER JOIN MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID" + TransportFilter +
                " LEFT OUTER JOIN JoinMainOrders ON DecorOrders.MainOrderID = JoinMainOrders.MarketMainOrderID" +
                " INNER JOIN infiniu2_catalog.dbo.DecorConfig" +
                " ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                " WHERE " + MDecorFactoryFilter + MDSampleFilter + " DecorOrders.MainOrderID IN (SELECT MainOrderID FROM MainOrders" +
                " WHERE MegaOrderID IN" +
                " (SELECT MegaOrderID FROM MegaOrders " + MFilter + MClientFilter + " ))";

            using (SqlDataAdapter DA = new SqlDataAdapter(MarketingSelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DecorOrdersDataTable.Clear();
                DA.Fill(DecorOrdersDataTable);
            }

            GetDecorProducts();
            GetDecorItems();
            GetDecorColors();
            GetDecorSizes();
        }

        public void FilterByOnAgreement(
            DateTime DateFrom, DateTime DateTo,
            int FactoryID, bool IsSample, bool IsNotSample, bool IsTransport)
        {
            string MClientFilter = string.Empty;
            string MFilter = string.Empty;

            string MFrontsFactoryFilter = string.Empty;
            string MDecorFactoryFilter = string.Empty;

            string MarketingSelectCommand = string.Empty;

            ArrayList MClients = SelectedMarketingClients;
            ArrayList MClientGroups = SelectedMarketingClientGroups;

            if (FactoryID != 0)
            {
                MFrontsFactoryFilter = " FrontsOrders.FactoryID = " + FactoryID + " AND ";
                MDecorFactoryFilter = " DecorOrders.FactoryID = " + FactoryID + " AND ";
            }

            MFilter = " WHERE CAST(OnAgreementDateTime AS DATE) >= '" + DateFrom.ToString("yyyy-MM-dd") +
                "' AND CAST(OnAgreementDateTime AS DATE) <= '" + DateTo.ToString("yyyy-MM-dd") + "' ";

            if (MClients.Count > 0)
            {
                if (MClientFilter.Length > 0)
                    MClientFilter += " AND ClientID IN (" + string.Join(",", MClients.OfType<Int32>().ToArray()) + ")";
                else
                    MClientFilter = " AND ClientID IN (" + string.Join(",", MClients.OfType<Int32>().ToArray()) + ")";
            }

            if (MClientGroups.Count > 0)
            {
                if (MClientFilter.Length > 0)
                    MClientFilter += " AND ClientID IN" +
                        " (SELECT ClientID FROM infiniu2_marketingreference.dbo.Clients" +
                        " WHERE ClientGroupID IN (" + string.Join(",", MClientGroups.OfType<Int32>().ToArray()) + "))";
                else
                    MClientFilter = " AND ClientID IN" +
                        " (SELECT ClientID FROM infiniu2_marketingreference.dbo.Clients" +
                        " WHERE ClientGroupID IN (" + string.Join(",", MClientGroups.OfType<Int32>().ToArray()) + "))";
            }

            if (MClients.Count < 1 && MClientGroups.Count < 1)
                MClientFilter = " AND ClientID = -1";

            string MFSampleFilter = string.Empty;
            string MDSampleFilter = string.Empty;

            if (!IsSample || !IsNotSample)
            {
                if (IsSample)
                    MFSampleFilter = " FrontsOrders.IsSample = 1 AND";
                if (IsNotSample)
                    MFSampleFilter = " FrontsOrders.IsSample = 0 AND";
                if (IsSample)
                    MDSampleFilter = " (DecorOrders.IsSample = 1 OR (DecorOrders.ProductID=42 AND DecorOrders.IsSample = 0)) AND";
                if (IsNotSample)
                    MDSampleFilter = " DecorOrders.IsSample = 0 AND";
            }
            string TransportFilter = " ";
            if (IsTransport)
                TransportFilter = " AND MegaOrders.TransportCost<>0 ";
            MarketingSelectCommand = "SELECT MegaOrders.OnAgreementDateTime, MegaOrders.MegaOrderID, MegaOrders.OrderNumber, FrontsOrders.MainOrderID, FrontsOrders.FrontID, FrontsOrders.PatinaID, FrontsOrders.ColorID, FrontsOrders.InsetTypeID," +
                " FrontsOrders.InsetColorID, FrontsOrders.TechnoColorID, FrontsOrders.TechnoInsetTypeID, FrontsOrders.TechnoInsetColorID, FrontsOrders.Height, FrontsOrders.Width," +
                " FrontsOrders.Count, FrontsOrders.IsSample, FrontsOrders.Square, FrontsOrders.Cost, MeasureID, ClientID, JoinMainOrders.ZOVClientID FROM FrontsOrders" +
                " INNER JOIN MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID" +
                " INNER JOIN MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID" + TransportFilter +
                " LEFT OUTER JOIN JoinMainOrders ON FrontsOrders.MainOrderID = JoinMainOrders.MarketMainOrderID" +
                " INNER JOIN infiniu2_catalog.dbo.FrontsConfig" +
                " ON FrontsOrders.FrontConfigID=infiniu2_catalog.dbo.FrontsConfig.FrontConfigID" +
                " WHERE " + MFrontsFactoryFilter + MFSampleFilter + " FrontsOrders.MainOrderID IN (SELECT MainOrderID FROM MainOrders" +
                " WHERE MegaOrderID IN" +
                " (SELECT MegaOrderID FROM MegaOrders " + MFilter + MClientFilter + " ))";

            using (SqlDataAdapter DA = new SqlDataAdapter(MarketingSelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                FrontsOrdersDataTable.Clear();
                DA.Fill(FrontsOrdersDataTable);
            }

            GetFronts();
            GetFrameColors();
            GetTechnoColors();
            GetInsetTypes();
            GetInsetColors();
            GetTechnoInsetTypes();
            GetTechnoInsetColors();
            GetSizes();

            //decor
            MarketingSelectCommand = "SELECT MegaOrders.OnAgreementDateTime, MegaOrders.MegaOrderID, MegaOrders.OrderNumber, DecorOrders.MainOrderID, DecorOrders.ProductID, DecorOrders.DecorID, DecorOrders.ColorID,DecorOrders.PatinaID," +
                " DecorOrders.Height, DecorOrders.Length, DecorOrders.Width, DecorOrders.Count, DecorOrders.IsSample, DecorOrders.Cost, DecorOrders.DecorConfigID, " +
                " MeasureID, ClientID, JoinMainOrders.ZOVClientID FROM DecorOrders" +
                " INNER JOIN MainOrders ON DecorOrders.MainOrderID = MainOrders.MainOrderID" +
                " INNER JOIN MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID" + TransportFilter +
                " LEFT OUTER JOIN JoinMainOrders ON DecorOrders.MainOrderID = JoinMainOrders.MarketMainOrderID" +
                " INNER JOIN infiniu2_catalog.dbo.DecorConfig" +
                " ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                " WHERE " + MDecorFactoryFilter + MDSampleFilter + " DecorOrders.MainOrderID IN (SELECT MainOrderID FROM MainOrders" +
                " WHERE MegaOrderID IN" +
                " (SELECT MegaOrderID FROM MegaOrders " + MFilter + MClientFilter + " ))";

            using (SqlDataAdapter DA = new SqlDataAdapter(MarketingSelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DecorOrdersDataTable.Clear();
                DA.Fill(DecorOrdersDataTable);
            }

            GetDecorProducts();
            GetDecorItems();
            GetDecorColors();
            GetDecorSizes();
        }

        public void FilterByPackages(
            DateTime DateFrom, DateTime DateTo,
            int PackageStatusID,
            int FactoryID, bool IsSample, bool IsNotSample, bool IsTransport)
        {
            string MarketingSelectCommand = string.Empty;
            string MClientFilter = string.Empty;

            string Date = string.Empty;
            string DateFilter = string.Empty;

            string MFrontsPackageFilter = string.Empty;
            string MDecorPackageFilter = string.Empty;

            string PackageFactoryFilter = string.Empty;
            string PackageProductFilter = string.Empty;

            ArrayList MClients = SelectedMarketingClients;
            ArrayList MClientGroups = SelectedMarketingClientGroups;

            if (PackageStatusID == 1)
                Date = "PackingDateTime";
            if (PackageStatusID == 2)
                Date = "StorageDateTime";
            if (PackageStatusID == 3)
                Date = "DispatchDateTime";
            if (PackageStatusID == 4)
                Date = "ExpeditionDateTime";
            DateFilter = " (CAST(" + Date + " AS DATE) >= '" + DateFrom.ToString("yyyy-MM-dd") +
                "' AND CAST(" + Date + " AS DATE) <= '" + DateTo.ToString("yyyy-MM-dd") + "')"; ;

            if (FactoryID != 0)
                PackageFactoryFilter = " AND FactoryID = " + FactoryID;

            PackageProductFilter = " AND ProductType = 0";
            MFrontsPackageFilter = " PackageID IN" +
                " (SELECT PackageID FROM Packages WHERE " + DateFilter + PackageFactoryFilter + PackageProductFilter + ")";

            PackageProductFilter = " AND ProductType = 1";
            MDecorPackageFilter = " PackageID IN" +
                " (SELECT PackageID FROM Packages WHERE " + DateFilter + PackageFactoryFilter + PackageProductFilter + ")";

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

            string TransportFilter = " ";
            if (IsTransport)
                TransportFilter = " AND MegaOrders.TransportCost<>0 ";

            MarketingSelectCommand = "SELECT MegaOrders.MegaOrderID, MegaOrders.OrderNumber, FrontsOrdersID, FrontsOrders.MainOrderID, FrontsOrders.FrontID, FrontsOrders.PatinaID, FrontsOrders.ColorID, FrontsOrders.InsetTypeID," +
                " FrontsOrders.InsetColorID, FrontsOrders.TechnoColorID, FrontsOrders.TechnoInsetTypeID, FrontsOrders.TechnoInsetColorID, FrontsOrders.Height, FrontsOrders.Width," +
                " PackageDetails.Count, FrontsOrders.IsSample, (FrontsOrders.Square * PackageDetails.Count / FrontsOrders.Count) AS Square, MeasureID, (FrontsOrders.Cost * PackageDetails.Count / FrontsOrders.Count) AS Cost, ClientID, JoinMainOrders.ZOVClientID FROM PackageDetails" +
                " INNER JOIN FrontsOrders ON PackageDetails.OrderID = FrontsOrders.FrontsOrdersID" +
                " INNER JOIN MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID" +
                " INNER JOIN MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID" + TransportFilter +
                " LEFT OUTER JOIN JoinMainOrders ON FrontsOrders.MainOrderID = JoinMainOrders.MarketMainOrderID" +
                " INNER JOIN infiniu2_catalog.dbo.FrontsConfig" +
                " ON FrontsOrders.FrontConfigID=infiniu2_catalog.dbo.FrontsConfig.FrontConfigID" +
                " WHERE " + MFrontsPackageFilter + MFSampleFilter + " AND FrontsOrders.MainOrderID IN (SELECT MainOrderID FROM MainOrders WHERE MegaOrderID IN" +
                " (SELECT MegaOrderID FROM MegaOrders " + MClientFilter + " )) ORDER BY FrontsOrdersID";

            using (SqlDataAdapter DA = new SqlDataAdapter(MarketingSelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                FrontsOrdersDataTable.Clear();
                DA.Fill(FrontsOrdersDataTable);
            }

            GetFronts();
            GetFrameColors();
            GetTechnoColors();
            GetInsetTypes();
            GetInsetColors();
            GetTechnoInsetTypes();
            GetTechnoInsetColors();
            GetSizes();

            //decor
            MarketingSelectCommand = "SELECT MegaOrders.MegaOrderID, MegaOrders.OrderNumber, DecorOrders.MainOrderID, DecorOrders.ProductID, DecorOrders.DecorID, DecorOrders.ColorID,DecorOrders.PatinaID," +
                " DecorOrders.Height, DecorOrders.Length, DecorOrders.Width, PackageDetails.Count, DecorOrders.IsSample, (DecorOrders.Cost * PackageDetails.Count / DecorOrders.Count) AS Cost, DecorOrders.DecorConfigID, " +
                " MeasureID, ClientID, JoinMainOrders.ZOVClientID FROM PackageDetails" +
                " INNER JOIN DecorOrders ON PackageDetails.OrderID = DecorOrders.DecorOrderID" +
                " INNER JOIN MainOrders ON DecorOrders.MainOrderID = MainOrders.MainOrderID" +
                " INNER JOIN MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID" + TransportFilter +
                " LEFT OUTER JOIN JoinMainOrders ON DecorOrders.MainOrderID = JoinMainOrders.MarketMainOrderID" +
                " INNER JOIN infiniu2_catalog.dbo.DecorConfig" +
                " ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                " WHERE " + MDecorPackageFilter + MDSampleFilter + " AND DecorOrders.MainOrderID IN (SELECT MainOrderID FROM MainOrders WHERE MegaOrderID IN" +
                " (SELECT MegaOrderID FROM MegaOrders " + MClientFilter + " ))";

            using (SqlDataAdapter DA = new SqlDataAdapter(MarketingSelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DecorOrdersDataTable.Clear();
                DA.Fill(DecorOrdersDataTable);
            }

            GetDecorProducts();
            GetDecorItems();
            GetDecorColors();
            GetDecorSizes();
        }

        public void GetOutProduction(DateTime DateFrom, DateTime DateTo, int FactoryID)
        {
            ArrayList MClients = SelectedMarketingClients;
            ArrayList MClientGroups = SelectedMarketingClientGroups;
        }

        private void Binding()
        {
            ManagersBindingSource = new BindingSource()
            {
                DataSource = ManagersDataTable
            };

            ClientGroupsBindingSource = new BindingSource()
            {
                DataSource = ClientGroupsDataTable
            };

            ClientsBindingSource = new BindingSource()
            {
                DataSource = ClientsDataTable,
                Sort = "ClientName ASC"
            };

            FrontsSummaryBindingSource.DataSource = FrontsSummaryDataTable;
            FrameColorsSummaryBindingSource.DataSource = FrameColorsSummaryDataTable;
            TechnoColorsSummaryBindingSource.DataSource = TechnoColorsSummaryDataTable;
            InsetTypesSummaryBindingSource.DataSource = InsetTypesSummaryDataTable;
            InsetColorsSummaryBindingSource.DataSource = InsetColorsSummaryDataTable;
            TechnoInsetTypesSummaryBindingSource.DataSource = TechnoInsetTypesSummaryDataTable;
            TechnoInsetColorsSummaryBindingSource.DataSource = TechnoInsetColorsSummaryDataTable;
            SizesSummaryBindingSource.DataSource = SizesSummaryDataTable;
            DecorProductsSummaryBindingSource.DataSource = DecorProductsSummaryDataTable;
            DecorItemsSummaryBindingSource.DataSource = DecorItemsSummaryDataTable;
            DecorColorsSummaryBindingSource.DataSource = DecorColorsSummaryDataTable;
            DecorSizesSummaryBindingSource.DataSource = DecorSizesSummaryDataTable;
        }

        public void Initialize()
        {
            Create();
            Fill();
            Binding();
            SetGrids();
        }

        private static void FreezeBand(DataGridViewBand band)
        {
            band.Frozen = true;
            DataGridViewCellStyle style = new DataGridViewCellStyle()
            {
                BackColor = Color.WhiteSmoke
            };
            band.DefaultCellStyle = style;
        }

        public void checkbox3_CheckedChanged(bool check)
        {
            CheckAllClients(check);
            CheckAllManagers(check);
        }

        private void SetGrids()
        {
        }

        public void FilterFrameColors(int FrontID)
        {
            FrameColorsSummaryBindingSource.Filter = "FrontID=" + FrontID;
            FrameColorsSummaryBindingSource.MoveFirst();
        }

        public void FilterTechnoColors(int FrontID, int ColorID, int PatinaID)
        {
            TechnoColorsSummaryBindingSource.Filter = "FrontID=" + FrontID + " AND ColorID=" + ColorID + " AND PatinaID=" + PatinaID;
            TechnoColorsSummaryBindingSource.MoveFirst();
        }

        public void FilterInsetTypes(int FrontID, int ColorID, int PatinaID, int TechnoColorID)
        {
            InsetTypesSummaryBindingSource.Filter = "FrontID=" + FrontID + " AND PatinaID=" + PatinaID +
                " AND ColorID=" + ColorID + " AND ColorID=" + ColorID + " AND TechnoColorID=" + TechnoColorID;
            InsetTypesSummaryBindingSource.MoveFirst();
        }

        public void FilterInsetColors(int FrontID, int ColorID, int PatinaID, int TechnoColorID, int InsetTypeID)
        {
            InsetColorsSummaryBindingSource.Filter = "FrontID=" + FrontID + " AND PatinaID=" +
                PatinaID + " AND InsetTypeID=" + InsetTypeID + " AND ColorID=" + ColorID;
            InsetColorsSummaryBindingSource.MoveFirst();
        }

        public void FilterTechnoInsetTypes(int FrontID, int ColorID, int PatinaID, int TechnoColorID, int InsetTypeID, int InsetColorID)
        {
            TechnoInsetTypesSummaryBindingSource.Filter = "FrontID=" + FrontID + " AND PatinaID=" + PatinaID +
                " AND TechnoColorID=" + TechnoColorID + " AND InsetColorID=" + InsetColorID + " AND InsetTypeID=" + InsetTypeID +
                " AND ColorID=" + ColorID;
            TechnoInsetTypesSummaryBindingSource.MoveFirst();
        }

        public void FilterTechnoInsetColors(int FrontID, int ColorID, int PatinaID, int TechnoColorID, int InsetTypeID, int InsetColorID, int TechnoInsetTypeID)
        {
            TechnoInsetColorsSummaryBindingSource.Filter = "FrontID=" + FrontID + " AND PatinaID=" + PatinaID +
                " AND InsetTypeID=" + InsetTypeID +
                " AND ColorID=" + ColorID + " AND TechnoColorID=" + TechnoColorID +
                " AND InsetColorID=" + InsetColorID + " AND TechnoInsetTypeID=" + TechnoInsetTypeID;
            TechnoInsetColorsSummaryBindingSource.MoveFirst();
        }

        public void FilterSizes(int FrontID, int ColorID, int PatinaID, int TechnoColorID, int InsetTypeID, int InsetColorID, int TechnoInsetTypeID, int TechnoInsetColorID)
        {
            SizesSummaryBindingSource.Filter = "FrontID=" + FrontID + " AND PatinaID=" + PatinaID +
                " AND ColorID=" + ColorID + " AND TechnoColorID=" + TechnoColorID +
                " AND InsetTypeID=" + InsetTypeID + " AND InsetColorID=" + InsetColorID + " AND TechnoInsetTypeID=" + TechnoInsetTypeID + " AND TechnoInsetColorID=" + TechnoInsetColorID;
            SizesSummaryBindingSource.MoveFirst();
        }

        public void FilterDecorProducts(int ProductID, int MeasureID)
        {
            DecorItemsSummaryBindingSource.Filter = "ProductID=" + ProductID + " AND MeasureID=" + MeasureID;
            DecorItemsSummaryBindingSource.MoveFirst();
        }

        public void FilterDecorItems(int ProductID, int DecorID, int MeasureID)
        {
            DecorColorsSummaryBindingSource.Filter = "ProductID=" + ProductID + " AND DecorID="
                + DecorID + " AND MeasureID=" + MeasureID;
            DecorColorsSummaryBindingSource.MoveFirst();
        }

        public void FilterDecorSizes(int ProductID, int DecorID, int ColorID, int MeasureID)
        {
            DecorSizesSummaryBindingSource.Filter = "ProductID=" + ProductID +
                " AND DecorID=" + DecorID + " AND ColorID=" + ColorID + " AND MeasureID=" + MeasureID;
            DecorSizesSummaryBindingSource.MoveFirst();
        }

        private void GetFronts()
        {
            decimal FrontCost = 0;
            decimal FrontSquare = 0;
            int FrontCount = 0;

            FrontsSummaryDataTable.Clear();

            DataTable Table = new DataTable();

            using (DataView DV = new DataView(FrontsOrdersDataTable))
            {
                Table = DV.ToTable(true, new string[] { "FrontID" });
            }

            for (int i = 0; i < Table.Rows.Count; i++)
            {
                DataRow[] Rows = FrontsOrdersDataTable.Select("FrontID=" + Convert.ToInt32(Table.Rows[i]["FrontID"]) +
                    " AND Width<>-1");
                if (Rows.Count() != 0)
                {
                    foreach (DataRow row in Rows)
                    {
                        FrontCost += Convert.ToDecimal(row["Cost"]);
                        FrontSquare += Convert.ToDecimal(row["Square"]);
                        FrontCount += Convert.ToInt32(row["Count"]);
                    }

                    DataRow NewRow = FrontsSummaryDataTable.NewRow();
                    NewRow["Front"] = GetFrontName(Convert.ToInt32(Table.Rows[i]["FrontID"]));
                    NewRow["FrontID"] = Convert.ToInt32(Table.Rows[i]["FrontID"]);
                    NewRow["Square"] = Decimal.Round(FrontSquare, 3, MidpointRounding.AwayFromZero);
                    NewRow["Cost"] = Decimal.Round(FrontCost, 3, MidpointRounding.AwayFromZero);
                    NewRow["Width"] = 0;
                    NewRow["Count"] = FrontCount;
                    FrontsSummaryDataTable.Rows.Add(NewRow);

                    FrontCost = 0;
                    FrontSquare = 0;
                    FrontCount = 0;
                }
                DataRow[] CurvedRows = FrontsOrdersDataTable.Select("FrontID=" + Convert.ToInt32(Table.Rows[i]["FrontID"]) +
                    " AND Width=-1");
                if (CurvedRows.Count() != 0)
                {
                    foreach (DataRow row in CurvedRows)
                    {
                        FrontCost += Convert.ToDecimal(row["Cost"]);
                        FrontCount += Convert.ToInt32(row["Count"]);
                    }

                    DataRow CurvedNewRow = FrontsSummaryDataTable.NewRow();
                    CurvedNewRow["Front"] = GetFrontName(Convert.ToInt32(Table.Rows[i]["FrontID"])) + " гнутый";
                    CurvedNewRow["FrontID"] = Convert.ToInt32(Table.Rows[i]["FrontID"]);
                    CurvedNewRow["Width"] = -1;
                    CurvedNewRow["Cost"] = Decimal.Round(FrontCost, 3, MidpointRounding.AwayFromZero);
                    CurvedNewRow["Count"] = FrontCount;
                    FrontsSummaryDataTable.Rows.Add(CurvedNewRow);

                    FrontCost = 0;
                    FrontCount = 0;
                }
            }

            Table.Dispose();
            FrontsSummaryDataTable.DefaultView.Sort = "Front, Square DESC";
            FrontsSummaryBindingSource.MoveFirst();
        }

        private void GetFrameColors()
        {
            decimal FrameColorCost = 0;
            decimal FrameColorSquare = 0;
            int FrameColorCount = 0;
            FrameColorsSummaryDataTable.Clear();

            DataTable Table = new DataTable();

            using (DataView DV = new DataView(FrontsOrdersDataTable))
            {
                Table = DV.ToTable(true, new string[] { "FrontID", "ColorID", "PatinaID" });
            }

            for (int i = 0; i < Table.Rows.Count; i++)
            {
                DataRow[] Rows = FrontsOrdersDataTable.Select("FrontID=" + Convert.ToInt32(Table.Rows[i]["FrontID"]) +
                    " AND Width<>-1 AND ColorID=" + Convert.ToInt32(Table.Rows[i]["ColorID"]) + " AND PatinaID=" + Convert.ToInt32(Table.Rows[i]["PatinaID"]));
                if (Rows.Count() != 0)
                {
                    foreach (DataRow row in Rows)
                    {
                        FrameColorCost += Convert.ToDecimal(row["Cost"]);
                        FrameColorSquare += Convert.ToDecimal(row["Square"]);
                        FrameColorCount += Convert.ToInt32(row["Count"]);
                    }

                    DataRow NewRow = FrameColorsSummaryDataTable.NewRow();
                    if (Convert.ToInt32(Table.Rows[i]["PatinaID"]) == -1)
                        NewRow["FrameColor"] = GetColorName(Convert.ToInt32(Table.Rows[i]["ColorID"]));
                    else
                        NewRow["FrameColor"] = GetColorName(Convert.ToInt32(Table.Rows[i]["ColorID"])) + " + " + GetPatinaName(Convert.ToInt32(Table.Rows[i]["PatinaID"]));
                    NewRow["FrontID"] = Convert.ToInt32(Table.Rows[i]["FrontID"]);
                    NewRow["ColorID"] = Convert.ToInt32(Table.Rows[i]["ColorID"]);
                    NewRow["PatinaID"] = Convert.ToInt32(Table.Rows[i]["PatinaID"]);
                    NewRow["Square"] = Decimal.Round(FrameColorSquare, 2, MidpointRounding.AwayFromZero);
                    NewRow["Cost"] = Decimal.Round(FrameColorCost, 2, MidpointRounding.AwayFromZero);
                    NewRow["Width"] = 0;
                    NewRow["Count"] = FrameColorCount;
                    FrameColorsSummaryDataTable.Rows.Add(NewRow);

                    FrameColorCost = 0;
                    FrameColorSquare = 0;
                    FrameColorCount = 0;
                }

                DataRow[] CurvedRows = FrontsOrdersDataTable.Select("FrontID=" + Convert.ToInt32(Table.Rows[i]["FrontID"]) +
                    " AND Width=-1 AND ColorID=" + Convert.ToInt32(Table.Rows[i]["ColorID"]) + " AND PatinaID=" + Convert.ToInt32(Table.Rows[i]["PatinaID"]));
                if (CurvedRows.Count() != 0)
                {
                    foreach (DataRow row in CurvedRows)
                    {
                        FrameColorCost += Convert.ToDecimal(row["Cost"]);
                        FrameColorCount += Convert.ToInt32(row["Count"]);
                    }

                    DataRow CurvedNewRow = FrameColorsSummaryDataTable.NewRow();
                    if (Convert.ToInt32(Table.Rows[i]["PatinaID"]) == -1)
                        CurvedNewRow["FrameColor"] = GetColorName(Convert.ToInt32(Table.Rows[i]["ColorID"]));
                    else
                        CurvedNewRow["FrameColor"] = GetColorName(Convert.ToInt32(Table.Rows[i]["ColorID"])) + " + " + GetPatinaName(Convert.ToInt32(Table.Rows[i]["PatinaID"]));
                    CurvedNewRow["FrontID"] = Convert.ToInt32(Table.Rows[i]["FrontID"]);
                    CurvedNewRow["ColorID"] = Convert.ToInt32(Table.Rows[i]["ColorID"]);
                    CurvedNewRow["PatinaID"] = Convert.ToInt32(Table.Rows[i]["PatinaID"]);
                    CurvedNewRow["Width"] = -1;
                    CurvedNewRow["Cost"] = Decimal.Round(FrameColorCost, 2, MidpointRounding.AwayFromZero);
                    CurvedNewRow["Count"] = FrameColorCount;
                    FrameColorsSummaryDataTable.Rows.Add(CurvedNewRow);

                    FrameColorCost = 0;
                    FrameColorCount = 0;
                }
            }
            Table.Dispose();
            FrameColorsSummaryDataTable.DefaultView.Sort = "FrameColor, Square DESC";
            FrameColorsSummaryBindingSource.MoveFirst();
        }

        private void GetTechnoColors()
        {
            decimal Square = 0;
            int Count = 0;

            TechnoColorsSummaryDataTable.Clear();

            DataTable Table = new DataTable();

            using (DataView DV = new DataView(FrontsOrdersDataTable))
            {
                Table = DV.ToTable(true, new string[] { "FrontID", "ColorID", "PatinaID", "TechnoColorID" });
            }

            for (int i = 0; i < Table.Rows.Count; i++)
            {
                DataRow[] Rows = FrontsOrdersDataTable.Select("FrontID=" + Convert.ToInt32(Table.Rows[i]["FrontID"]) +
                    " AND ColorID=" + Convert.ToInt32(Table.Rows[i]["ColorID"]) + " AND PatinaID=" + Convert.ToInt32(Table.Rows[i]["PatinaID"]) + " AND TechnoColorID=" + Convert.ToInt32(Table.Rows[i]["TechnoColorID"]) +
                    " AND Width<>-1");
                if (Rows.Count() != 0)
                {
                    foreach (DataRow row in Rows)
                    {
                        Square += Convert.ToDecimal(row["Square"]);
                        Count += Convert.ToInt32(row["Count"]);
                    }

                    DataRow NewRow = TechnoColorsSummaryDataTable.NewRow();
                    NewRow["TechnoColor"] = GetColorName(Convert.ToInt32(Table.Rows[i]["TechnoColorID"]));
                    NewRow["FrontID"] = Convert.ToInt32(Table.Rows[i]["FrontID"]);
                    NewRow["ColorID"] = Convert.ToInt32(Table.Rows[i]["ColorID"]);
                    NewRow["PatinaID"] = Convert.ToInt32(Table.Rows[i]["PatinaID"]);
                    NewRow["TechnoColorID"] = Convert.ToInt32(Table.Rows[i]["TechnoColorID"]);
                    NewRow["Width"] = 0;
                    NewRow["Square"] = Decimal.Round(Square, 2, MidpointRounding.AwayFromZero);
                    NewRow["Count"] = Count;
                    TechnoColorsSummaryDataTable.Rows.Add(NewRow);

                    Square = 0;
                    Count = 0;
                }

                DataRow[] CurvedRows = FrontsOrdersDataTable.Select("FrontID=" + Convert.ToInt32(Table.Rows[i]["FrontID"]) +
                    " AND ColorID=" + Convert.ToInt32(Table.Rows[i]["ColorID"]) + " AND TechnoColorID=" + Convert.ToInt32(Table.Rows[i]["TechnoColorID"]) +
                    " AND Width=-1 AND PatinaID=" + Convert.ToInt32(Table.Rows[i]["PatinaID"]));
                if (CurvedRows.Count() != 0)
                {
                    foreach (DataRow row in CurvedRows)
                    {
                        Square += Convert.ToDecimal(row["Square"]);
                        Count += Convert.ToInt32(row["Count"]);
                    }

                    DataRow CurvedNewRow = TechnoColorsSummaryDataTable.NewRow();
                    CurvedNewRow["TechnoColor"] = GetColorName(Convert.ToInt32(Table.Rows[i]["TechnoColorID"]));
                    CurvedNewRow["FrontID"] = Convert.ToInt32(Table.Rows[i]["FrontID"]);
                    CurvedNewRow["ColorID"] = Convert.ToInt32(Table.Rows[i]["ColorID"]);
                    CurvedNewRow["PatinaID"] = Convert.ToInt32(Table.Rows[i]["PatinaID"]);
                    CurvedNewRow["TechnoColorID"] = Convert.ToInt32(Table.Rows[i]["TechnoColorID"]);
                    CurvedNewRow["Width"] = -1;
                    CurvedNewRow["Count"] = Count;
                    TechnoColorsSummaryDataTable.Rows.Add(CurvedNewRow);

                    Square = 0;
                    Count = 0;
                }
            }
            Table.Dispose();
            TechnoColorsSummaryDataTable.DefaultView.Sort = "TechnoColor, Count DESC";
            TechnoColorsSummaryBindingSource.MoveFirst();
        }

        private void GetInsetTypes()
        {
            decimal InsetTypeSquare = 0;
            int InsetTypeCount = 0;

            int MarginHeight = 0;
            int MarginWidth = 0;
            decimal GridHeight = 0;
            decimal GridWidth = 0;

            InsetTypesSummaryDataTable.Clear();

            DataTable Table = new DataTable();

            using (DataView DV = new DataView(FrontsOrdersDataTable))
            {
                Table = DV.ToTable(true, new string[] { "FrontID", "ColorID", "PatinaID", "TechnoColorID", "InsetTypeID" });
            }

            for (int i = 0; i < Table.Rows.Count; i++)
            {
                DataRow[] Rows = FrontsOrdersDataTable.Select("FrontID=" + Convert.ToInt32(Table.Rows[i]["FrontID"]) +
                    " AND ColorID=" + Convert.ToInt32(Table.Rows[i]["ColorID"]) + " AND TechnoColorID=" + Convert.ToInt32(Table.Rows[i]["TechnoColorID"]) + " AND PatinaID=" + Convert.ToInt32(Table.Rows[i]["PatinaID"]) +
                    " AND Width<>-1 AND InsetTypeID=" + Convert.ToInt32(Table.Rows[i]["InsetTypeID"]));
                if (Rows.Count() != 0)
                {
                    foreach (DataRow row in Rows)
                    {
                        InsetTypeSquare += Convert.ToDecimal(row["Square"]);
                        InsetTypeCount += Convert.ToInt32(row["Count"]);
                    }
                    //foreach (DataRow row in Rows)
                    //{
                    //    GetGridMargins(Convert.ToInt32(row["FrontID"]), ref MarginHeight, ref MarginWidth);
                    //    GridHeight = Convert.ToInt32(Convert.ToInt32(row["Height"]) - MarginHeight);
                    //    GridWidth = Convert.ToInt32(Convert.ToInt32(row["Width"]) - MarginWidth);
                    //    if (GridHeight < 0 || GridWidth < 0)
                    //    {
                    //        GridHeight = 0;
                    //        GridWidth = 0;
                    //    }
                    //    InsetTypeSquare += Decimal.Round(GridHeight * GridWidth * Convert.ToInt32(row["Count"]) / 1000000, 2, MidpointRounding.AwayFromZero);
                    //    InsetTypeCount += Convert.ToInt32(row["Count"]);
                    //}

                    DataRow NewRow = InsetTypesSummaryDataTable.NewRow();
                    NewRow["InsetType"] = GetInsetTypeName(Convert.ToInt32(Table.Rows[i]["InsetTypeID"]));
                    NewRow["FrontID"] = Convert.ToInt32(Table.Rows[i]["FrontID"]);
                    NewRow["ColorID"] = Convert.ToInt32(Table.Rows[i]["ColorID"]);
                    NewRow["TechnoColorID"] = Convert.ToInt32(Table.Rows[i]["TechnoColorID"]);
                    NewRow["PatinaID"] = Convert.ToInt32(Table.Rows[i]["PatinaID"]);
                    NewRow["InsetTypeID"] = Table.Rows[i]["InsetTypeID"];
                    //if (Convert.ToInt32(Table.Rows[i]["InsetTypeID"]) == 1)
                    //    InsetTypeSquare = 0;
                    NewRow["Width"] = 0;
                    NewRow["Square"] = Decimal.Round(InsetTypeSquare, 2, MidpointRounding.AwayFromZero);
                    NewRow["Count"] = InsetTypeCount;
                    InsetTypesSummaryDataTable.Rows.Add(NewRow);

                    InsetTypeSquare = 0;
                    InsetTypeCount = 0;
                }

                DataRow[] CurvedRows = FrontsOrdersDataTable.Select("FrontID=" + Convert.ToInt32(Table.Rows[i]["FrontID"]) +
                    " AND ColorID=" + Convert.ToInt32(Table.Rows[i]["ColorID"]) + " AND TechnoColorID=" + Convert.ToInt32(Table.Rows[i]["TechnoColorID"]) +
                    " AND Width=-1 AND InsetTypeID=" + Convert.ToInt32(Table.Rows[i]["InsetTypeID"]));
                if (CurvedRows.Count() != 0)
                {
                    foreach (DataRow row in CurvedRows)
                    {
                        GetGridMargins(Convert.ToInt32(row["FrontID"]), ref MarginHeight, ref MarginWidth);
                        GridHeight = Convert.ToInt32(Convert.ToInt32(row["Height"]) - MarginHeight);
                        GridWidth = Convert.ToInt32(Convert.ToInt32(row["Width"]) - MarginWidth);
                        if (GridHeight < 0 || GridWidth < 0)
                        {
                            GridHeight = 0;
                            GridWidth = 0;
                        }
                        InsetTypeSquare += Decimal.Round(GridHeight * GridWidth * Convert.ToInt32(row["Count"]) / 1000000, 2, MidpointRounding.AwayFromZero);
                        InsetTypeCount += Convert.ToInt32(row["Count"]);
                    }

                    DataRow CurvedNewRow = InsetTypesSummaryDataTable.NewRow();
                    CurvedNewRow["InsetType"] = GetInsetTypeName(Convert.ToInt32(Table.Rows[i]["InsetTypeID"]));
                    CurvedNewRow["FrontID"] = Convert.ToInt32(Table.Rows[i]["FrontID"]);
                    CurvedNewRow["ColorID"] = Convert.ToInt32(Table.Rows[i]["ColorID"]);
                    CurvedNewRow["TechnoColorID"] = Convert.ToInt32(Table.Rows[i]["TechnoColorID"]);
                    CurvedNewRow["PatinaID"] = Convert.ToInt32(Table.Rows[i]["PatinaID"]);
                    CurvedNewRow["InsetTypeID"] = Convert.ToInt32(Table.Rows[i]["InsetTypeID"]);
                    CurvedNewRow["Width"] = -1;
                    CurvedNewRow["Count"] = InsetTypeCount;
                    InsetTypesSummaryDataTable.Rows.Add(CurvedNewRow);

                    InsetTypeSquare = 0;
                    InsetTypeCount = 0;
                }
            }
            Table.Dispose();
            InsetTypesSummaryDataTable.DefaultView.Sort = "InsetType, Count DESC";
            InsetTypesSummaryBindingSource.MoveFirst();
        }

        private void GetInsetColors()
        {
            decimal InsetColorSquare = 0;
            int InsetColorCount = 0;

            int MarginHeight = 0;
            int MarginWidth = 0;
            decimal GridHeight = 0;
            decimal GridWidth = 0;

            InsetColorsSummaryDataTable.Clear();

            DataTable Table = new DataTable();

            using (DataView DV = new DataView(FrontsOrdersDataTable))
            {
                Table = DV.ToTable(true, new string[] { "FrontID", "ColorID", "TechnoColorID", "PatinaID", "InsetTypeID", "InsetColorID" });
            }

            for (int i = 0; i < Table.Rows.Count; i++)
            {
                DataRow[] Rows = FrontsOrdersDataTable.Select("FrontID=" + Convert.ToInt32(Table.Rows[i]["FrontID"]) +
                    " AND ColorID=" + Convert.ToInt32(Table.Rows[i]["ColorID"]) + " AND TechnoColorID=" + Convert.ToInt32(Table.Rows[i]["TechnoColorID"]) + " AND PatinaID=" + Convert.ToInt32(Table.Rows[i]["PatinaID"]) +
                    " AND Width<>-1 AND InsetTypeID=" + Convert.ToInt32(Table.Rows[i]["InsetTypeID"]) +
                    " AND InsetColorID=" + Convert.ToInt32(Table.Rows[i]["InsetColorID"]));
                if (Rows.Count() != 0)
                {
                    foreach (DataRow row in Rows)
                    {
                        GetGridMargins(Convert.ToInt32(row["FrontID"]), ref MarginHeight, ref MarginWidth);
                        GridHeight = Convert.ToInt32(Convert.ToInt32(row["Height"]) - MarginHeight);
                        GridWidth = Convert.ToInt32(Convert.ToInt32(row["Width"]) - MarginWidth);
                        if (GridHeight < 0 || GridWidth < 0)
                        {
                            GridHeight = 0;
                            GridWidth = 0;
                        }
                        InsetColorSquare += Decimal.Round(GridHeight * GridWidth * Convert.ToInt32(row["Count"]) / 1000000, 2, MidpointRounding.AwayFromZero);
                        InsetColorCount += Convert.ToInt32(row["Count"]);
                    }

                    DataRow NewRow = InsetColorsSummaryDataTable.NewRow();
                    NewRow["InsetColor"] = GetInsetColorName(Convert.ToInt32(Table.Rows[i]["InsetColorID"]));
                    NewRow["FrontID"] = Convert.ToInt32(Table.Rows[i]["FrontID"]);
                    NewRow["ColorID"] = Convert.ToInt32(Table.Rows[i]["ColorID"]);
                    NewRow["TechnoColorID"] = Convert.ToInt32(Table.Rows[i]["TechnoColorID"]);
                    NewRow["PatinaID"] = Convert.ToInt32(Table.Rows[i]["PatinaID"]);
                    NewRow["InsetTypeID"] = Convert.ToInt32(Table.Rows[i]["InsetTypeID"]);
                    NewRow["InsetColorID"] = Convert.ToInt32(Table.Rows[i]["InsetColorID"]);
                    if (Convert.ToInt32(Table.Rows[i]["InsetTypeID"]) == 1)
                        InsetColorSquare = 0;
                    NewRow["Width"] = 0;
                    NewRow["Square"] = Decimal.Round(InsetColorSquare, 2, MidpointRounding.AwayFromZero);
                    NewRow["Count"] = InsetColorCount;
                    InsetColorsSummaryDataTable.Rows.Add(NewRow);

                    InsetColorSquare = 0;
                    InsetColorCount = 0;
                }

                DataRow[] CurvedRows = FrontsOrdersDataTable.Select("FrontID=" + Convert.ToInt32(Table.Rows[i]["FrontID"]) +
                    " AND ColorID=" + Convert.ToInt32(Table.Rows[i]["ColorID"]) + " AND TechnoColorID=" + Convert.ToInt32(Table.Rows[i]["TechnoColorID"]) +
                    " AND Width=-1 AND InsetTypeID=" + Convert.ToInt32(Table.Rows[i]["InsetTypeID"]) +
                    " AND InsetColorID=" + Convert.ToInt32(Table.Rows[i]["InsetColorID"]));
                if (CurvedRows.Count() != 0)
                {
                    foreach (DataRow row in CurvedRows)
                    {
                        GetGridMargins(Convert.ToInt32(row["FrontID"]), ref MarginHeight, ref MarginWidth);
                        GridHeight = Convert.ToInt32(Convert.ToInt32(row["Height"]) - MarginHeight);
                        GridWidth = Convert.ToInt32(Convert.ToInt32(row["Width"]) - MarginWidth);
                        if (GridHeight < 0 || GridWidth < 0)
                        {
                            GridHeight = 0;
                            GridWidth = 0;
                        }
                        InsetColorSquare += Decimal.Round(GridHeight * GridWidth * Convert.ToInt32(row["Count"]) / 1000000, 2, MidpointRounding.AwayFromZero);
                        InsetColorCount += Convert.ToInt32(row["Count"]);
                    }

                    DataRow CurvedNewRow = InsetColorsSummaryDataTable.NewRow();
                    CurvedNewRow["InsetColor"] = GetInsetColorName(Convert.ToInt32(Table.Rows[i]["InsetColorID"]));
                    CurvedNewRow["FrontID"] = Convert.ToInt32(Table.Rows[i]["FrontID"]);
                    CurvedNewRow["ColorID"] = Convert.ToInt32(Table.Rows[i]["ColorID"]);
                    CurvedNewRow["TechnoColorID"] = Convert.ToInt32(Table.Rows[i]["TechnoColorID"]);
                    CurvedNewRow["PatinaID"] = Convert.ToInt32(Table.Rows[i]["PatinaID"]);
                    CurvedNewRow["InsetTypeID"] = Convert.ToInt32(Table.Rows[i]["InsetTypeID"]);
                    CurvedNewRow["InsetColorID"] = Convert.ToInt32(Table.Rows[i]["InsetColorID"]);
                    CurvedNewRow["Width"] = -1;
                    CurvedNewRow["Count"] = InsetColorCount;
                    InsetColorsSummaryDataTable.Rows.Add(CurvedNewRow);

                    InsetColorCount = 0;
                }
            }
            Table.Dispose();
            InsetColorsSummaryDataTable.DefaultView.Sort = "InsetColor, Count DESC";
            InsetColorsSummaryBindingSource.MoveFirst();
        }

        private void GetTechnoInsetTypes()
        {
            decimal InsetTypeSquare = 0;
            int InsetTypeCount = 0;

            TechnoInsetTypesSummaryDataTable.Clear();

            DataTable Table = new DataTable();

            using (DataView DV = new DataView(FrontsOrdersDataTable))
            {
                Table = DV.ToTable(true, new string[] { "FrontID", "ColorID", "PatinaID", "TechnoColorID", "InsetTypeID", "InsetColorID", "TechnoInsetTypeID" });
            }

            for (int i = 0; i < Table.Rows.Count; i++)
            {
                DataRow[] Rows = FrontsOrdersDataTable.Select("FrontID=" + Convert.ToInt32(Table.Rows[i]["FrontID"]) +
                    " AND ColorID=" + Convert.ToInt32(Table.Rows[i]["ColorID"]) +
                    " AND TechnoColorID=" + Convert.ToInt32(Table.Rows[i]["TechnoColorID"]) +
                    " AND PatinaID=" + Convert.ToInt32(Table.Rows[i]["PatinaID"]) +
                    " AND Width<>-1 AND InsetTypeID=" + Convert.ToInt32(Table.Rows[i]["InsetTypeID"]) +
                    " AND InsetColorID=" + Convert.ToInt32(Table.Rows[i]["InsetColorID"]) +
                    " AND TechnoInsetTypeID=" + Convert.ToInt32(Table.Rows[i]["TechnoInsetTypeID"]));
                if (Rows.Count() != 0)
                {
                    foreach (DataRow row in Rows)
                    {
                        InsetTypeSquare += Convert.ToDecimal(row["Square"]);
                        InsetTypeCount += Convert.ToInt32(row["Count"]);
                    }

                    DataRow NewRow = TechnoInsetTypesSummaryDataTable.NewRow();
                    NewRow["TechnoInsetType"] = GetInsetTypeName(Convert.ToInt32(Table.Rows[i]["TechnoInsetTypeID"]));
                    NewRow["FrontID"] = Convert.ToInt32(Table.Rows[i]["FrontID"]);
                    NewRow["ColorID"] = Convert.ToInt32(Table.Rows[i]["ColorID"]);
                    NewRow["TechnoColorID"] = Convert.ToInt32(Table.Rows[i]["TechnoColorID"]);
                    NewRow["PatinaID"] = Convert.ToInt32(Table.Rows[i]["PatinaID"]);
                    NewRow["InsetTypeID"] = Table.Rows[i]["InsetTypeID"];
                    NewRow["InsetColorID"] = Table.Rows[i]["InsetColorID"];
                    NewRow["TechnoInsetTypeID"] = Table.Rows[i]["TechnoInsetTypeID"];
                    if (Convert.ToInt32(Table.Rows[i]["InsetTypeID"]) == 1)
                        InsetTypeSquare = 0;
                    NewRow["Width"] = 0;
                    NewRow["Square"] = Decimal.Round(InsetTypeSquare, 2, MidpointRounding.AwayFromZero);
                    NewRow["Count"] = InsetTypeCount;
                    TechnoInsetTypesSummaryDataTable.Rows.Add(NewRow);

                    InsetTypeSquare = 0;
                    InsetTypeCount = 0;
                }

                DataRow[] CurvedRows = FrontsOrdersDataTable.Select("FrontID=" + Convert.ToInt32(Table.Rows[i]["FrontID"]) +
                    " AND ColorID=" + Convert.ToInt32(Table.Rows[i]["ColorID"]) +
                    " AND TechnoColorID=" + Convert.ToInt32(Table.Rows[i]["TechnoColorID"]) +
                    " AND Width=-1 AND InsetTypeID=" + Convert.ToInt32(Table.Rows[i]["InsetTypeID"]) +
                    " AND InsetColorID=" + Convert.ToInt32(Table.Rows[i]["InsetColorID"]) +
                    " AND TechnoInsetTypeID=" + Convert.ToInt32(Table.Rows[i]["TechnoInsetTypeID"]));
                if (CurvedRows.Count() != 0)
                {
                    foreach (DataRow row in CurvedRows)
                    {
                        InsetTypeSquare += Convert.ToDecimal(row["Square"]);
                        InsetTypeCount += Convert.ToInt32(row["Count"]);
                    }

                    DataRow CurvedNewRow = TechnoInsetTypesSummaryDataTable.NewRow();
                    CurvedNewRow["TechnoInsetType"] = GetInsetTypeName(Convert.ToInt32(Table.Rows[i]["TechnoInsetTypeID"]));
                    CurvedNewRow["FrontID"] = Convert.ToInt32(Table.Rows[i]["FrontID"]);
                    CurvedNewRow["ColorID"] = Convert.ToInt32(Table.Rows[i]["ColorID"]);
                    CurvedNewRow["TechnoColorID"] = Convert.ToInt32(Table.Rows[i]["TechnoColorID"]);
                    CurvedNewRow["PatinaID"] = Convert.ToInt32(Table.Rows[i]["PatinaID"]);
                    CurvedNewRow["InsetTypeID"] = Convert.ToInt32(Table.Rows[i]["InsetTypeID"]);
                    CurvedNewRow["InsetColorID"] = Convert.ToInt32(Table.Rows[i]["InsetColorID"]);
                    CurvedNewRow["TechnoInsetTypeID"] = Convert.ToInt32(Table.Rows[i]["TechnoInsetTypeID"]);
                    CurvedNewRow["Width"] = -1;
                    CurvedNewRow["Count"] = InsetTypeCount;
                    TechnoInsetTypesSummaryDataTable.Rows.Add(CurvedNewRow);

                    InsetTypeSquare = 0;
                    InsetTypeCount = 0;
                }
            }
            Table.Dispose();
            TechnoInsetTypesSummaryDataTable.DefaultView.Sort = "TechnoInsetType, Count DESC";
            TechnoInsetTypesSummaryBindingSource.MoveFirst();
        }

        private void GetTechnoInsetColors()
        {
            decimal InsetColorSquare = 0;
            int InsetColorCount = 0;

            TechnoInsetColorsSummaryDataTable.Clear();

            DataTable Table = new DataTable();

            using (DataView DV = new DataView(FrontsOrdersDataTable))
            {
                Table = DV.ToTable(true, new string[] { "FrontID", "ColorID", "TechnoColorID", "PatinaID", "InsetTypeID", "InsetColorID", "TechnoInsetTypeID", "TechnoInsetColorID" });
            }

            for (int i = 0; i < Table.Rows.Count; i++)
            {
                DataRow[] Rows = FrontsOrdersDataTable.Select("FrontID=" + Convert.ToInt32(Table.Rows[i]["FrontID"]) +
                    " AND ColorID=" + Convert.ToInt32(Table.Rows[i]["ColorID"]) +
                    " AND TechnoColorID=" + Convert.ToInt32(Table.Rows[i]["TechnoColorID"]) +
                    " AND PatinaID=" + Convert.ToInt32(Table.Rows[i]["PatinaID"]) +
                    " AND Width<>-1 AND InsetTypeID=" + Convert.ToInt32(Table.Rows[i]["InsetTypeID"]) +
                    " AND InsetColorID=" + Convert.ToInt32(Table.Rows[i]["InsetColorID"]) +
                    " AND TechnoInsetTypeID=" + Convert.ToInt32(Table.Rows[i]["TechnoInsetTypeID"]) +
                    " AND TechnoInsetColorID=" + Convert.ToInt32(Table.Rows[i]["TechnoInsetColorID"]));
                if (Rows.Count() != 0)
                {
                    foreach (DataRow row in Rows)
                    {
                        InsetColorSquare += Convert.ToDecimal(row["Square"]);
                        InsetColorCount += Convert.ToInt32(row["Count"]);
                    }

                    DataRow NewRow = TechnoInsetColorsSummaryDataTable.NewRow();
                    NewRow["TechnoInsetColor"] = GetInsetColorName(Convert.ToInt32(Table.Rows[i]["TechnoInsetColorID"]));
                    NewRow["FrontID"] = Convert.ToInt32(Table.Rows[i]["FrontID"]);
                    NewRow["ColorID"] = Convert.ToInt32(Table.Rows[i]["ColorID"]);
                    NewRow["TechnoColorID"] = Convert.ToInt32(Table.Rows[i]["TechnoColorID"]);
                    NewRow["PatinaID"] = Convert.ToInt32(Table.Rows[i]["PatinaID"]);
                    NewRow["InsetTypeID"] = Convert.ToInt32(Table.Rows[i]["InsetTypeID"]);
                    NewRow["InsetColorID"] = Convert.ToInt32(Table.Rows[i]["InsetColorID"]);
                    NewRow["TechnoInsetTypeID"] = Convert.ToInt32(Table.Rows[i]["TechnoInsetTypeID"]);
                    NewRow["TechnoInsetColorID"] = Convert.ToInt32(Table.Rows[i]["TechnoInsetColorID"]);
                    if (Convert.ToInt32(Table.Rows[i]["InsetTypeID"]) == 1)
                        InsetColorSquare = 0;
                    NewRow["Width"] = 0;
                    NewRow["Square"] = Decimal.Round(InsetColorSquare, 2, MidpointRounding.AwayFromZero);
                    NewRow["Count"] = InsetColorCount;
                    TechnoInsetColorsSummaryDataTable.Rows.Add(NewRow);

                    InsetColorSquare = 0;
                    InsetColorCount = 0;
                }

                DataRow[] CurvedRows = FrontsOrdersDataTable.Select("FrontID=" + Convert.ToInt32(Table.Rows[i]["FrontID"]) +
                    " AND ColorID=" + Convert.ToInt32(Table.Rows[i]["ColorID"]) + " AND TechnoColorID=" + Convert.ToInt32(Table.Rows[i]["TechnoColorID"]) +
                    " AND Width=-1 AND InsetTypeID=" + Convert.ToInt32(Table.Rows[i]["InsetTypeID"]) +
                    " AND InsetColorID=" + Convert.ToInt32(Table.Rows[i]["InsetColorID"]) +
                    " AND TechnoInsetTypeID=" + Convert.ToInt32(Table.Rows[i]["TechnoInsetTypeID"]) +
                    " AND TechnoInsetColorID=" + Convert.ToInt32(Table.Rows[i]["TechnoInsetColorID"]));
                if (CurvedRows.Count() != 0)
                {
                    foreach (DataRow row in CurvedRows)
                    {
                        InsetColorSquare += Convert.ToDecimal(row["Square"]);
                        InsetColorCount += Convert.ToInt32(row["Count"]);
                    }

                    DataRow CurvedNewRow = TechnoInsetColorsSummaryDataTable.NewRow();
                    CurvedNewRow["TechnoInsetColor"] = GetInsetColorName(Convert.ToInt32(Table.Rows[i]["TechnoInsetColorID"]));
                    CurvedNewRow["FrontID"] = Convert.ToInt32(Table.Rows[i]["FrontID"]);
                    CurvedNewRow["ColorID"] = Convert.ToInt32(Table.Rows[i]["ColorID"]);
                    CurvedNewRow["TechnoColorID"] = Convert.ToInt32(Table.Rows[i]["TechnoColorID"]);
                    CurvedNewRow["PatinaID"] = Convert.ToInt32(Table.Rows[i]["PatinaID"]);
                    CurvedNewRow["InsetTypeID"] = Convert.ToInt32(Table.Rows[i]["InsetTypeID"]);
                    CurvedNewRow["InsetColorID"] = Convert.ToInt32(Table.Rows[i]["InsetColorID"]);
                    CurvedNewRow["TechnoInsetTypeID"] = Convert.ToInt32(Table.Rows[i]["TechnoInsetTypeID"]);
                    CurvedNewRow["TechnoInsetColorID"] = Convert.ToInt32(Table.Rows[i]["TechnoInsetColorID"]);
                    CurvedNewRow["Width"] = -1;
                    CurvedNewRow["Count"] = InsetColorCount;
                    TechnoInsetColorsSummaryDataTable.Rows.Add(CurvedNewRow);

                    InsetColorCount = 0;
                }
            }
            Table.Dispose();
            TechnoInsetColorsSummaryDataTable.DefaultView.Sort = "TechnoInsetColor, Count DESC";
            TechnoInsetColorsSummaryBindingSource.MoveFirst();
        }

        private void GetSizes()
        {
            decimal SizeSquare = 0;
            int SizeCount = 0;
            SizesSummaryDataTable.Clear();

            DataTable Table = new DataTable();

            using (DataView DV = new DataView(FrontsOrdersDataTable))
            {
                Table = DV.ToTable(true, new string[] { "FrontID", "ColorID", "TechnoColorID", "PatinaID", "InsetTypeID", "InsetColorID", "TechnoInsetTypeID", "TechnoInsetColorID", "Height", "Width" });
            }

            for (int i = 0; i < Table.Rows.Count; i++)
            {
                DataRow[] Rows = FrontsOrdersDataTable.Select("FrontID=" + Convert.ToInt32(Table.Rows[i]["FrontID"]) +
                    " AND Width<>-1 AND ColorID=" + Convert.ToInt32(Table.Rows[i]["ColorID"]) +
                    " AND PatinaID=" + Convert.ToInt32(Table.Rows[i]["PatinaID"]) +
                    " AND InsetTypeID=" + Convert.ToInt32(Table.Rows[i]["InsetTypeID"]) +
                    " AND InsetColorID=" + Convert.ToInt32(Table.Rows[i]["InsetColorID"]) +
                    " AND TechnoColorID=" + Convert.ToInt32(Table.Rows[i]["TechnoColorID"]) +
                    " AND TechnoInsetTypeID=" + Convert.ToInt32(Table.Rows[i]["TechnoInsetTypeID"]) +
                    " AND TechnoInsetColorID=" + Convert.ToInt32(Table.Rows[i]["TechnoInsetColorID"]) +
                    " AND Height=" + Convert.ToInt32(Table.Rows[i]["Height"]) +
                    " AND Width=" + Convert.ToInt32(Table.Rows[i]["Width"]));
                if (Rows.Count() != 0)
                {
                    foreach (DataRow row in Rows)
                    {
                        SizeSquare += Convert.ToDecimal(row["Square"]);
                        SizeCount += Convert.ToInt32(row["Count"]);
                    }

                    DataRow NewRow = SizesSummaryDataTable.NewRow();
                    NewRow["Size"] = Convert.ToInt32(Table.Rows[i]["Height"]) + " x " + Convert.ToInt32(Table.Rows[i]["Width"]);
                    NewRow["FrontID"] = Convert.ToInt32(Table.Rows[i]["FrontID"]);
                    NewRow["ColorID"] = Convert.ToInt32(Table.Rows[i]["ColorID"]);
                    NewRow["PatinaID"] = Convert.ToInt32(Table.Rows[i]["PatinaID"]);
                    NewRow["InsetColorID"] = Convert.ToInt32(Table.Rows[i]["InsetColorID"]);
                    NewRow["InsetTypeID"] = Convert.ToInt32(Table.Rows[i]["InsetTypeID"]);
                    NewRow["TechnoColorID"] = Convert.ToInt32(Table.Rows[i]["TechnoColorID"]);
                    NewRow["TechnoInsetColorID"] = Convert.ToInt32(Table.Rows[i]["TechnoInsetColorID"]);
                    NewRow["TechnoInsetTypeID"] = Convert.ToInt32(Table.Rows[i]["TechnoInsetTypeID"]);
                    NewRow["Height"] = Convert.ToInt32(Table.Rows[i]["Height"]);
                    NewRow["Width"] = Convert.ToInt32(Table.Rows[i]["Width"]);
                    NewRow["Square"] = Decimal.Round(SizeSquare, 2, MidpointRounding.AwayFromZero);
                    NewRow["Count"] = SizeCount;
                    SizesSummaryDataTable.Rows.Add(NewRow);

                    SizeSquare = 0;
                    SizeCount = 0;
                }

                DataRow[] CurvedRows = FrontsOrdersDataTable.Select("FrontID=" + Convert.ToInt32(Table.Rows[i]["FrontID"]) +
                    " AND Width=-1 AND ColorID=" + Convert.ToInt32(Table.Rows[i]["ColorID"]) +
                    " AND InsetTypeID=" + Convert.ToInt32(Table.Rows[i]["InsetTypeID"]) +
                    " AND PatinaID=" + Convert.ToInt32(Table.Rows[i]["PatinaID"]) +
                    " AND InsetColorID=" + Convert.ToInt32(Table.Rows[i]["InsetColorID"]) +
                    " AND TechnoColorID=" + Convert.ToInt32(Table.Rows[i]["TechnoColorID"]) +
                    " AND TechnoInsetTypeID=" + Convert.ToInt32(Table.Rows[i]["TechnoInsetTypeID"]) +
                    " AND TechnoInsetColorID=" + Convert.ToInt32(Table.Rows[i]["TechnoInsetColorID"]) +
                    " AND Height=" + Convert.ToInt32(Table.Rows[i]["Height"]) +
                    " AND Width=" + Convert.ToInt32(Table.Rows[i]["Width"]));
                if (CurvedRows.Count() != 0)
                {
                    foreach (DataRow row in CurvedRows)
                    {
                        SizeSquare += Convert.ToDecimal(row["Square"]);
                        SizeCount += Convert.ToInt32(row["Count"]);
                    }

                    DataRow CurvedNewRow = SizesSummaryDataTable.NewRow();
                    CurvedNewRow["Size"] = Convert.ToInt32(Table.Rows[i]["Height"]) + " x " + Convert.ToInt32(Table.Rows[i]["Width"]);
                    CurvedNewRow["FrontID"] = Convert.ToInt32(Table.Rows[i]["FrontID"]);
                    CurvedNewRow["ColorID"] = Convert.ToInt32(Table.Rows[i]["ColorID"]);
                    CurvedNewRow["PatinaID"] = Convert.ToInt32(Table.Rows[i]["PatinaID"]);
                    CurvedNewRow["InsetColorID"] = Convert.ToInt32(Table.Rows[i]["InsetColorID"]);
                    CurvedNewRow["InsetTypeID"] = Convert.ToInt32(Table.Rows[i]["InsetTypeID"]);
                    CurvedNewRow["TechnoColorID"] = Convert.ToInt32(Table.Rows[i]["TechnoColorID"]);
                    CurvedNewRow["TechnoInsetColorID"] = Convert.ToInt32(Table.Rows[i]["TechnoInsetColorID"]);
                    CurvedNewRow["TechnoInsetTypeID"] = Convert.ToInt32(Table.Rows[i]["TechnoInsetTypeID"]);
                    CurvedNewRow["Height"] = Convert.ToInt32(Table.Rows[i]["Height"]);
                    CurvedNewRow["Width"] = Convert.ToInt32(Table.Rows[i]["Width"]);
                    CurvedNewRow["Count"] = SizeCount;
                    SizesSummaryDataTable.Rows.Add(CurvedNewRow);

                    SizeCount = 0;
                }
            }
            Table.Dispose();
            SizesSummaryDataTable.DefaultView.Sort = "Square DESC";
            SizesSummaryBindingSource.MoveFirst();
        }

        private void GetDecorProducts()
        {
            decimal DecorProductCost = 0;
            decimal DecorProductCount = 0;
            int decimals = 2;
            string Measure = string.Empty;
            DecorProductsSummaryDataTable.Clear();

            DataTable Table = new DataTable();

            using (DataView DV = new DataView(DecorConfigDataTable))
            {
                Table = DV.ToTable(true, new string[] { "ProductID", "MeasureID" });
            }

            for (int i = 0; i < Table.Rows.Count; i++)
            {
                DataRow[] Rows = DecorOrdersDataTable.Select("ProductID=" + Convert.ToInt32(Table.Rows[i]["ProductID"]) +
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

                DataRow NewRow = DecorProductsSummaryDataTable.NewRow();
                NewRow["ProductID"] = Convert.ToInt32(Table.Rows[i]["ProductID"]);
                NewRow["MeasureID"] = Convert.ToInt32(Table.Rows[i]["MeasureID"]);
                NewRow["DecorProduct"] = GetProductName(Convert.ToInt32(Table.Rows[i]["ProductID"]));
                if (DecorProductCount < 3)
                    decimals = 1;
                NewRow["Cost"] = Decimal.Round(DecorProductCost, decimals, MidpointRounding.AwayFromZero);
                NewRow["Count"] = Decimal.Round(DecorProductCount, decimals, MidpointRounding.AwayFromZero);
                NewRow["Measure"] = Measure;
                DecorProductsSummaryDataTable.Rows.Add(NewRow);

                Measure = string.Empty;
                DecorProductCost = 0;
                DecorProductCount = 0;
            }
            DecorProductsSummaryDataTable.DefaultView.Sort = "DecorProduct, Measure ASC, Count DESC";
            DecorProductsSummaryBindingSource.MoveFirst();
        }

        private void GetDecorItems()
        {
            decimal DecorItemCost = 0;
            decimal DecorItemCount = 0;
            int decimals = 2;
            string Measure = string.Empty;
            DecorItemsSummaryDataTable.Clear();

            DataTable Table = new DataTable();

            using (DataView DV = new DataView(DecorOrdersDataTable))
            {
                Table = DV.ToTable(true, new string[] { "ProductID", "DecorID", "MeasureID" });
            }

            for (int i = 0; i < Table.Rows.Count; i++)
            {
                DataRow[] Rows = DecorOrdersDataTable.Select("ProductID=" + Convert.ToInt32(Table.Rows[i]["ProductID"]) +
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

                DataRow NewRow = DecorItemsSummaryDataTable.NewRow();
                NewRow["ProductID"] = Convert.ToInt32(Table.Rows[i]["ProductID"]);
                NewRow["DecorID"] = Convert.ToInt32(Table.Rows[i]["DecorID"]);
                NewRow["MeasureID"] = Convert.ToInt32(Table.Rows[i]["MeasureID"]);
                NewRow["DecorItem"] = GetDecorName(Convert.ToInt32(Table.Rows[i]["DecorID"]));
                if (DecorItemCount < 3)
                    decimals = 1;
                NewRow["Cost"] = Decimal.Round(DecorItemCost, decimals, MidpointRounding.AwayFromZero);
                NewRow["Count"] = Decimal.Round(DecorItemCount, decimals, MidpointRounding.AwayFromZero);
                NewRow["Measure"] = Measure;
                DecorItemsSummaryDataTable.Rows.Add(NewRow);

                Measure = string.Empty;
                DecorItemCost = 0;
                DecorItemCount = 0;
            }
            Table.Dispose();
            DecorItemsSummaryDataTable.DefaultView.Sort = "DecorItem, Count DESC";
            DecorItemsSummaryBindingSource.MoveFirst();
        }

        private void GetDecorColors()
        {
            decimal DecorColorCost = 0;
            decimal DecorColorCount = 0;
            int decimals = 2;
            string Measure = string.Empty;
            DecorColorsSummaryDataTable.Clear();

            DataTable Table = new DataTable();

            using (DataView DV = new DataView(DecorOrdersDataTable))
            {
                Table = DV.ToTable(true, new string[] { "ProductID", "DecorID", "ColorID", "MeasureID" });
            }

            for (int i = 0; i < Table.Rows.Count; i++)
            {
                DataRow[] Rows = DecorOrdersDataTable.Select("ProductID=" + Convert.ToInt32(Table.Rows[i]["ProductID"]) +
                    " AND DecorID=" + Convert.ToInt32(Table.Rows[i]["DecorID"]) + " AND ColorID=" + Convert.ToInt32(Table.Rows[i]["ColorID"]) +
                    " AND MeasureID=" + Convert.ToInt32(Table.Rows[i]["MeasureID"]));
                if (Rows.Count() == 0)
                    continue;

                foreach (DataRow row in Rows)
                {
                    if (Convert.ToInt32(row["ProductID"]) == 2)
                    {
                        DecorColorCost += Convert.ToDecimal(row["Cost"]);
                        if (row["Height"].ToString() == "-1")
                            DecorColorCount += Convert.ToDecimal(row["Length"]) * Convert.ToDecimal(row["Count"]) / 1000;
                        else
                            DecorColorCount += Convert.ToDecimal(row["Height"]) * Convert.ToDecimal(row["Count"]) / 1000;
                        Measure = "м.п.";
                        continue;
                    }

                    if (Convert.ToInt32(row["MeasureID"]) == 1)
                    {
                        DecorColorCost += Convert.ToDecimal(row["Cost"]);
                        DecorColorCount += Convert.ToInt32(row["Count"]);
                        Measure = "шт.";
                    }

                    if (Convert.ToInt32(row["MeasureID"]) == 3)
                    {
                        DecorColorCost += Convert.ToDecimal(row["Cost"]);
                        DecorColorCount += Convert.ToInt32(row["Count"]);
                        Measure = "шт.";
                    }

                    if (Convert.ToInt32(row["MeasureID"]) == 2)
                    {
                        //нет параметра "высота"
                        if (row["Height"].ToString() == "-1")
                            DecorColorCount += Convert.ToDecimal(row["Length"]) * Convert.ToDecimal(row["Count"]) / 1000;
                        else
                            DecorColorCount += Convert.ToDecimal(row["Height"]) * Convert.ToDecimal(row["Count"]) / 1000;

                        DecorColorCost += Convert.ToDecimal(row["Cost"]);
                        Measure = "м.п.";
                    }
                }

                DataRow NewRow = DecorColorsSummaryDataTable.NewRow();
                NewRow["ProductID"] = Convert.ToInt32(Table.Rows[i]["ProductID"]);
                NewRow["DecorID"] = Convert.ToInt32(Table.Rows[i]["DecorID"]);
                NewRow["ColorID"] = Convert.ToInt32(Table.Rows[i]["ColorID"]);
                NewRow["MeasureID"] = Convert.ToInt32(Table.Rows[i]["MeasureID"]);
                NewRow["Color"] = GetColorName(Convert.ToInt32(Table.Rows[i]["ColorID"]));
                if (DecorColorCount < 3)
                    decimals = 1;
                NewRow["Cost"] = Decimal.Round(DecorColorCost, decimals, MidpointRounding.AwayFromZero);
                NewRow["Count"] = Decimal.Round(DecorColorCount, decimals, MidpointRounding.AwayFromZero);
                NewRow["Measure"] = Measure;
                DecorColorsSummaryDataTable.Rows.Add(NewRow);

                Measure = string.Empty;
                DecorColorCost = 0;
                DecorColorCount = 0;
            }
            Table.Dispose();
            DecorColorsSummaryDataTable.DefaultView.Sort = "Color, Count DESC";
            DecorColorsSummaryBindingSource.MoveFirst();
        }

        private void GetDecorSizes()
        {
            decimal DecorSizeCost = 0;
            decimal DecorSizeCount = 0;
            int decimals = 2;
            int Height = 0;
            int Length = 0;
            int Width = 0;
            string Measure = string.Empty;
            string Sizes = string.Empty;
            DecorSizesSummaryDataTable.Clear();

            DataTable Table = new DataTable();

            using (DataView DV = new DataView(DecorOrdersDataTable))
            {
                Table = DV.ToTable(true, new string[] { "ProductID", "DecorID", "ColorID", "MeasureID", "Length", "Height", "Width" });
            }

            for (int i = 0; i < Table.Rows.Count; i++)
            {
                DataRow[] Rows = DecorOrdersDataTable.Select("ProductID=" + Convert.ToInt32(Table.Rows[i]["ProductID"]) +
                    " AND DecorID=" + Convert.ToInt32(Table.Rows[i]["DecorID"]) +
                    " AND ColorID=" + Convert.ToInt32(Table.Rows[i]["ColorID"]) +
                    " AND MeasureID=" + Convert.ToInt32(Table.Rows[i]["MeasureID"]) +
                    " AND Length=" + Convert.ToInt32(Table.Rows[i]["Length"]) +
                    " AND Height=" + Convert.ToInt32(Table.Rows[i]["Height"]) +
                    " AND Width=" + Convert.ToInt32(Table.Rows[i]["Width"]));
                if (Rows.Count() == 0)
                    continue;

                foreach (DataRow row in Rows)
                {
                    if (Convert.ToInt32(row["ProductID"]) == 2)
                    {
                        DecorSizeCost += Convert.ToDecimal(row["Cost"]);
                        if (row["Height"].ToString() == "-1")
                            DecorSizeCount += Convert.ToDecimal(row["Length"]) * Convert.ToDecimal(row["Count"]) / 1000;
                        else
                            DecorSizeCount += Convert.ToDecimal(row["Height"]) * Convert.ToDecimal(row["Count"]) / 1000;
                        Measure = "м.п.";
                        continue;
                    }

                    if (Convert.ToInt32(row["MeasureID"]) == 1)
                    {
                        DecorSizeCost += Convert.ToDecimal(row["Cost"]);
                        DecorSizeCount += Convert.ToInt32(row["Count"]);
                        Measure = "шт.";
                    }

                    if (Convert.ToInt32(row["MeasureID"]) == 3)
                    {
                        DecorSizeCost += Convert.ToDecimal(row["Cost"]);
                        DecorSizeCount += Convert.ToInt32(row["Count"]);
                        Measure = "шт.";
                    }

                    if (Convert.ToInt32(row["MeasureID"]) == 2)
                    {
                        //нет параметра "высота"
                        if (row["Height"].ToString() == "-1")
                            DecorSizeCount += Convert.ToDecimal(row["Length"]) * Convert.ToDecimal(row["Count"]) / 1000;
                        else
                            DecorSizeCount += Convert.ToDecimal(row["Height"]) * Convert.ToDecimal(row["Count"]) / 1000;

                        DecorSizeCost += Convert.ToDecimal(row["Cost"]);
                        Measure = "м.п.";
                    }
                }

                DataRow NewRow = DecorSizesSummaryDataTable.NewRow();
                NewRow["ProductID"] = Convert.ToInt32(Table.Rows[i]["ProductID"]);
                NewRow["DecorID"] = Convert.ToInt32(Table.Rows[i]["DecorID"]);
                NewRow["ColorID"] = Convert.ToInt32(Table.Rows[i]["ColorID"]);
                NewRow["MeasureID"] = Convert.ToInt32(Table.Rows[i]["MeasureID"]);
                if (DecorSizeCount < 3)
                    decimals = 1;
                NewRow["Cost"] = Decimal.Round(DecorSizeCost, decimals, MidpointRounding.AwayFromZero);
                NewRow["Count"] = Decimal.Round(DecorSizeCount, decimals, MidpointRounding.AwayFromZero);
                NewRow["Measure"] = Measure;

                Height = Convert.ToInt32(Table.Rows[i]["Height"]);
                Length = Convert.ToInt32(Table.Rows[i]["Length"]);
                Width = Convert.ToInt32(Table.Rows[i]["Width"]);

                if (Height > -1)
                    Sizes = Height.ToString();

                if (Sizes != string.Empty)
                {
                    if (Width > -1)
                        Sizes += " x " + Width.ToString();
                }
                else
                {
                    if (Length > -1)
                    {
                        Sizes = Length.ToString();
                        if (Width > -1)
                            Sizes += " x " + Width.ToString();
                    }
                    else
                    {
                        if (Width > -1)
                            Sizes = Width.ToString();
                    }
                }

                DecorSizesSummaryDataTable.Rows.Add(NewRow);
                NewRow["Size"] = Sizes;
                Sizes = string.Empty;
                Measure = string.Empty;
                DecorSizeCost = 0;
                DecorSizeCount = 0;
            }
            Table.Dispose();
            DecorSizesSummaryDataTable.DefaultView.Sort = "Count DESC";
            DecorSizesSummaryBindingSource.MoveFirst();
        }

        public void GetFrontsInfo(ref decimal Square, ref decimal Cost, ref int Count, ref int CurvedCount)
        {
        }

        public void GetDecorInfo(ref decimal Pogon, ref decimal Cost, ref decimal Count)
        {
        }

        public string GetFrontName(int FrontID)
        {
            string FrontName = string.Empty;
            DataRow[] Rows = FrontsDataTable.Select("FrontID = " + FrontID);
            if (Rows.Count() > 0)
                FrontName = Rows[0]["FrontName"].ToString();
            return FrontName;
        }

        public string GetFront2Name(int TechnoProfileID)
        {
            string FrontName = "";
            try
            {
                DataRow[] Rows = FrontsDataTable.Select("FrontID = " + TechnoProfileID);
                FrontName = Rows[0]["FrontName"].ToString();
            }
            catch
            {
                return "";
            }
            return FrontName;
        }

        public string GetColorName(int ColorID)
        {
            string ColorName = string.Empty;
            DataRow[] Rows = FrameColorsDataTable.Select("ColorID = " + ColorID);
            if (Rows.Count() > 0)
                ColorName = Rows[0]["ColorName"].ToString();
            return ColorName;
        }

        public string GetPatinaName(int PatinaID)
        {
            string PatinaName = string.Empty;
            DataRow[] Rows = PatinaDataTable.Select("PatinaID = " + PatinaID);
            if (Rows.Count() > 0)
                PatinaName = Rows[0]["PatinaName"].ToString();
            return PatinaName;
        }

        public string GetInsetTypeName(int InsetTypeID)
        {
            string InsetType = string.Empty;
            DataRow[] Rows = InsetTypesDataTable.Select("InsetTypeID = " + InsetTypeID);
            if (Rows.Count() > 0)
                InsetType = Rows[0]["InsetTypeName"].ToString();
            return InsetType;
        }

        public string GetInsetColorName(int InsetColorID)
        {
            string ColorName = string.Empty;
            DataRow[] Rows = InsetColorsDataTable.Select("InsetColorID = " + InsetColorID);
            if (Rows.Count() > 0)
                ColorName = Rows[0]["InsetColorName"].ToString();
            return ColorName;
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
                DataRow[] Rows = DecorProductsDataTable.Select("ProductID = " + ProductID);
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
                DataRow[] Rows = DecorDataTable.Select("DecorID = " + DecorID);
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

        //private void GetGridMargins(int FrontID, ref int MarginHeight, ref int MarginWidth)
        //{
        //    DataRow[] Rows = InsetMarginsDataTable.Select("FrontID = " + FrontID);
        //    if (Rows.Count() == 0)
        //        return;
        //    MarginHeight = Convert.ToInt32(Rows[0]["InsetHeightAdmission"]);
        //    MarginWidth = Convert.ToInt32(Rows[0]["GridWidth"]);
        //}

        private void GetGridMargins(int FrontID, ref int MarginHeight, ref int MarginWidth)
        {
            DataRow[] Rows = TechStoreDataTable.Select("TechStoreID = " + FrontID);
            if (Rows.Count() > 0)
            {
                if (Rows[0]["InsetHeightAdmission"] != DBNull.Value)
                    MarginHeight = Convert.ToInt32(Rows[0]["InsetHeightAdmission"]);
                if (Rows[0]["InsetWidthAdmission"] != DBNull.Value)
                    MarginWidth = Convert.ToInt32(Rows[0]["InsetWidthAdmission"]);
            }
        }

        public ArrayList SelectedMarketingClients
        {
            get
            {
                ArrayList Clients = new ArrayList();

                for (int i = 0; i < ClientsDataTable.Rows.Count; i++)
                {
                    if (!Convert.ToBoolean(ClientsDataTable.Rows[i]["Check"]))
                        continue;

                    Clients.Add(Convert.ToInt32(ClientsDataTable.Rows[i]["ClientID"]));
                }

                return Clients;
            }
        }

        public ArrayList SelectedMarketingClientGroups
        {
            get
            {
                ArrayList ClientGroupIDs = new ArrayList();

                for (int i = 0; i < ClientGroupsDataTable.Rows.Count; i++)
                {
                    if (!Convert.ToBoolean(ClientGroupsDataTable.Rows[i]["Check"]))
                        continue;
                    //if (Convert.ToInt32(ClientGroupsDataTable.Rows[i]["ClientGroupID"]) == 1)
                    //    continue;

                    ClientGroupIDs.Add(Convert.ToInt32(ClientGroupsDataTable.Rows[i]["ClientGroupID"]));
                }

                return ClientGroupIDs;
            }
        }

        public void GetCurrentManager()
        {
            if (ManagersBindingSource.Count == 0)
            {
                CurrentManagerID = -1;
                return;
            }
            if (((DataRowView)ManagersBindingSource.Current).Row["ManagerID"] == DBNull.Value)
                return;
            else
                CurrentManagerID = Convert.ToInt32(((DataRowView)ManagersBindingSource.Current).Row["ManagerID"]);
        }

        public void SetCheckClientsByManager(bool Check)
        {
            //string filter = string.Empty;
            //ArrayList ManagerIDs = new ArrayList();

            //for (int i = 0; i < ManagersDataTable.Rows.Count; i++)
            //{
            //    if (!Convert.ToBoolean(ManagersDataTable.Rows[i]["Check"]))
            //        continue;
            //    ManagerIDs.Add(ManagersDataTable.Rows[i]["ManagerID"].ToString());
            //}

            //foreach (string item in ManagerIDs)
            //    filter += item.ToString() + ",";
            //if (filter.Length > 0)
            //    filter = "ManagerID IN (" + filter + ")";
            //else
            //    filter = "";
            //ClientsBindingSource.Filter = filter;

            CurrentManagerID = Convert.ToInt32(((DataRowView)ManagersBindingSource.Current).Row["ManagerID"]);
            for (int i = 0; i < ClientsDataTable.Rows.Count; i++)
            {
                int ManagerID = Convert.ToInt32(ClientsDataTable.Rows[i]["ManagerID"]);
                if (ManagerID == CurrentManagerID)
                    ClientsDataTable.Rows[i]["Check"] = Check;
            }
        }

        public void CheckAllManagers(bool Check)
        {
            for (int i = 0; i < ManagersDataTable.Rows.Count; i++)
            {
                ManagersDataTable.Rows[i]["Check"] = Check;
            }
        }

        public void CheckAllClients(bool Check)
        {
            for (int i = 0; i < ClientsDataTable.Rows.Count; i++)
            {
                ClientsDataTable.Rows[i]["Check"] = Check;
            }
        }

        public void CheckAllClientGroups(bool Check)
        {
            for (int i = 0; i < ClientGroupsDataTable.Rows.Count; i++)
            {
                ClientGroupsDataTable.Rows[i]["Check"] = Check;
            }
        }

        public void ShowCheckColumn(PercentageDataGrid tPercentageDataGrid, bool Shown)
        {
            tPercentageDataGrid.Columns["Check"].Visible = Shown;
        }

        // Given H,S,L in range of 0-1
        // Returns a Color (RGB struct) in range of 0-255
        public static ColorRGB HSL2RGB(double h, double sl, double l)

        {
            double v;

            double r, g, b;

            r = l;   // default to gray

            g = l;

            b = l;

            v = (l <= 0.5) ? (l * (1.0 + sl)) : (l + sl - l * sl);

            if (v > 0)

            {
                double m;

                double sv;

                int sextant;

                double fract, vsf, mid1, mid2;

                m = l + l - v;

                sv = (v - m) / v;

                h *= 6.0;

                sextant = (int)h;

                fract = h - sextant;

                vsf = v * sv * fract;

                mid1 = m + vsf;

                mid2 = v - vsf;

                switch (sextant)

                {
                    case 0:

                        r = v;

                        g = mid1;

                        b = m;

                        break;

                    case 1:

                        r = mid2;

                        g = v;

                        b = m;

                        break;

                    case 2:

                        r = m;

                        g = v;

                        b = mid1;

                        break;

                    case 3:

                        r = m;

                        g = mid2;

                        b = v;

                        break;

                    case 4:

                        r = mid1;

                        g = m;

                        b = v;

                        break;

                    case 5:

                        r = v;

                        g = m;

                        b = mid2;

                        break;
                }
            }

            ColorRGB rgb;

            rgb.R = Convert.ToByte(r * 255.0f);

            rgb.G = Convert.ToByte(g * 255.0f);

            rgb.B = Convert.ToByte(b * 255.0f);

            return rgb;
        }

        // Given a Color (RGB Struct) in range of 0-255
        // Return H,S,L in range of 0-1
        public static void RGB2HSL(ColorRGB rgb, out double h, out double s, out double l)

        {
            double r = rgb.R / 255.0;

            double g = rgb.G / 255.0;

            double b = rgb.B / 255.0;

            double v;

            double m;

            double vm;

            double r2, g2, b2;

            h = 0; // default to black

            s = 0;

            l = 0;

            v = Math.Max(r, g);

            v = Math.Max(v, b);

            m = Math.Min(r, g);

            m = Math.Min(m, b);

            l = (m + v) / 2.0;

            if (l <= 0.0)

            {
                return;
            }

            vm = v - m;

            s = vm;

            if (s > 0.0)

            {
                s /= (l <= 0.5) ? (v + m) : (2.0 - v - m);
            }
            else

            {
                return;
            }

            r2 = (v - r) / vm;

            g2 = (v - g) / vm;

            b2 = (v - b) / vm;

            if (r == v)

            {
                h = (g == m ? 5.0 + b2 : 1.0 - g2);
            }
            else if (g == v)

            {
                h = (b == m ? 1.0 + r2 : 3.0 - b2);
            }
            else

            {
                h = (r == m ? 3.0 + g2 : 5.0 - r2);
            }

            h /= 6.0;
        }
    }

    public struct ColorRGB

    {
        public byte R;

        public byte G;

        public byte B;

        public ColorRGB(Color value)

        {
            this.R = value.R;

            this.G = value.G;

            this.B = value.B;
        }

        public static implicit operator Color(ColorRGB rgb)

        {
            Color c = Color.FromArgb(rgb.R, rgb.G, rgb.B);

            return c;
        }

        public static explicit operator ColorRGB(Color c)

        {
            return new ColorRGB(c);
        }
    }
}