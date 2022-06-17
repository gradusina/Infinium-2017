using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.UserModel.Contrib;
using NPOI.HSSF.Util;

using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;

namespace Infinium.Modules.WorkAssignments
{

    public class DecorAssignments
    {
        private FileManager FM = new FileManager();
        private DateTime CurrentDate;
        private DataTable DecorDT;
        public DataTable TechStoreDataTable = null;
        public DataTable PatinaDataTable = null;
        private DataTable PatinaRALDataTable = null;
        private DataTable BagetWithAngleAssemblyDT;
        private DataTable DecorAssemblyDT;
        private DataTable FrameColorsDataTable;
        private DataTable DecorParametersDT;
        private DataTable BagetWithAngelOrdersDT;
        private DataTable NotArchDecorOrdersDT;
        private DataTable ArchDecorOrdersDT;
        private DataTable StrapsDecorOrdersDT;
        private DataTable GridsDecorOrdersDT;

        public DecorAssignments()
        {

        }

        public void Initialize()
        {
            Create();
            Fill();
        }

        private void Create()
        {
            DecorParametersDT = new DataTable();
            DecorDT = new DataTable();
            TechStoreDataTable = new DataTable();
            PatinaDataTable = new DataTable();

            BagetWithAngelOrdersDT = new DataTable();
            NotArchDecorOrdersDT = new DataTable();
            ArchDecorOrdersDT = new DataTable();
            StrapsDecorOrdersDT = new DataTable();
            GridsDecorOrdersDT = new DataTable();

            DecorAssemblyDT = new DataTable();
            DecorAssemblyDT.Columns.Add(new DataColumn("Name", Type.GetType("System.String")));
            DecorAssemblyDT.Columns.Add(new DataColumn("Color", Type.GetType("System.String")));
            DecorAssemblyDT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            DecorAssemblyDT.Columns.Add(new DataColumn("Width", Type.GetType("System.Int32")));
            DecorAssemblyDT.Columns.Add(new DataColumn("Count", Type.GetType("System.Int32")));
            DecorAssemblyDT.Columns.Add(new DataColumn("Notes", Type.GetType("System.String")));
            DecorAssemblyDT.Columns.Add(new DataColumn("DecorID", Type.GetType("System.Int32")));

            BagetWithAngleAssemblyDT = new DataTable();
            BagetWithAngleAssemblyDT.Columns.Add(new DataColumn("Name", Type.GetType("System.String")));
            BagetWithAngleAssemblyDT.Columns.Add(new DataColumn("Color", Type.GetType("System.String")));
            BagetWithAngleAssemblyDT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            BagetWithAngleAssemblyDT.Columns.Add(new DataColumn("Width", Type.GetType("System.Int32")));
            BagetWithAngleAssemblyDT.Columns.Add(new DataColumn("LeftAngle", Type.GetType("System.Decimal")));
            BagetWithAngleAssemblyDT.Columns.Add(new DataColumn("RightAngle", Type.GetType("System.Decimal")));
            BagetWithAngleAssemblyDT.Columns.Add(new DataColumn("Count", Type.GetType("System.Int32")));
            BagetWithAngleAssemblyDT.Columns.Add(new DataColumn("Notes", Type.GetType("System.String")));
            BagetWithAngleAssemblyDT.Columns.Add(new DataColumn("DecorID", Type.GetType("System.Int32")));
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
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM InsetColors WHERE GroupID IN (2,3,4,5,6,21,22)", ConnectionStrings.CatalogConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    DA.Fill(DT);

                    for (int i = 0; i < DT.Rows.Count; i++)
                    {
                        DataRow[] rows = FrameColorsDataTable.Select("ColorID=" + Convert.ToInt32(DT.Rows[i]["InsetColorID"]));
                        if (rows.Count() == 0)
                        {
                            DataRow NewRow = FrameColorsDataTable.NewRow();
                            NewRow["ColorID"] = Convert.ToInt64(DT.Rows[i]["InsetColorID"]);
                            NewRow["ColorName"] = DT.Rows[i]["InsetColorName"].ToString();
                            FrameColorsDataTable.Rows.Add(NewRow);
                        }
                    }
                }
            }
        }

        private void Fill()
        {
            DecorDT = new DataTable();
            string SelectCommand = @"SELECT DISTINCT TechStore.TechStoreID AS DecorID, TechStore.TechStoreName AS Name, DecorConfig.ProductID FROM TechStore 
                INNER JOIN DecorConfig ON TechStore.TechStoreID = DecorConfig.DecorID AND Enabled = 1 ORDER BY TechStoreName";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(DecorDT);
            }
            GetColorsDT();
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
            //using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM TechStore",
            //    ConnectionStrings.CatalogConnectionString))
            //{
            //    DA.Fill(TechStoreDataTable);
            //}
            TechStoreDataTable = TablesManager.TechStoreDataTable;

            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT * FROM DecorParameters",
                ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(DecorParametersDT);
            }

            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT TOP 0 DecorOrderID, MainOrderID, ProductID, DecorID, DecorConfigID, ColorID, PatinaID,
                Height, Length, Width, LeftAngle, RightAngle, Count, Notes FROM DecorOrders",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(BagetWithAngelOrdersDT);
                BagetWithAngelOrdersDT.Columns.Add(new DataColumn("GroupType", Type.GetType("System.Int32")));
            }
            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT TOP 0 DecorOrderID, MainOrderID, ProductID, DecorID, DecorConfigID, ColorID, PatinaID,
                Height, Length, Width, Count, Notes FROM DecorOrders",
                ConnectionStrings.ZOVOrdersConnectionString))
            {
                DA.Fill(NotArchDecorOrdersDT);
                NotArchDecorOrdersDT.Columns.Add(new DataColumn("GroupType", Type.GetType("System.Int32")));
                ArchDecorOrdersDT = NotArchDecorOrdersDT.Clone();
                StrapsDecorOrdersDT = NotArchDecorOrdersDT.Clone();
                GridsDecorOrdersDT = NotArchDecorOrdersDT.Clone();
            }
        }

        public bool HasParameter(int ProductID, string Parameter)
        {
            DataRow[] Rows = DecorParametersDT.Select("ProductID = " + ProductID);

            return Convert.ToBoolean(Rows[0][Parameter]);
        }

        public string GetMarketClientName(int MainOrderID)
        {
            string ClientName = string.Empty;

            using (DataTable DT = new DataTable())
            {
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT ClientName FROM Clients WHERE ClientID = " +
                    " (SELECT ClientID FROM infiniu2_marketingorders.dbo.MegaOrders" +
                    " WHERE MegaOrderID=(SELECT TOP 1 MegaOrderID FROM infiniu2_marketingorders.dbo.MainOrders WHERE MainOrderID = " + MainOrderID + "))",
                    ConnectionStrings.MarketingReferenceConnectionString))
                {
                    DA.Fill(DT);
                    if (DT.Rows.Count > 0)
                        ClientName = DT.Rows[0]["ClientName"].ToString();
                }
            }

            return ClientName;
        }

        public string GetZOVClientName(int MainOrderID)
        {
            string ClientName = string.Empty;

            using (DataTable DT = new DataTable())
            {
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT ClientName FROM Clients WHERE ClientID = " +
                    " (SELECT ClientID FROM infiniu2_zovorders.dbo.MainOrders WHERE MainOrderID = " + MainOrderID + ")",
                    ConnectionStrings.ZOVReferenceConnectionString))
                {
                    DA.Fill(DT);
                    if (DT.Rows.Count > 0)
                        ClientName = DT.Rows[0]["ClientName"].ToString();
                }
            }

            return ClientName;
        }

        private string GetOrderName(int MainOrderID, int GroupType)
        {
            string name = string.Empty;
            string ConnectionString = ConnectionStrings.ZOVOrdersConnectionString;
            string SelectCommand = string.Empty;
            DataTable DT = new DataTable();

            if (GroupType == 1)
                ConnectionString = ConnectionStrings.MarketingOrdersConnectionString;
            SelectCommand = @"SELECT MegaBatchID, BatchID FROM Batch WHERE BatchID IN (SELECT BatchID FROM BatchDetails WHERE MainOrderID = " + MainOrderID + ")";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionString))
            {
                if (DA.Fill(DT) > 0 && DT.Rows[0]["MegaBatchID"] != DBNull.Value && DT.Rows[0]["BatchID"] != DBNull.Value)
                    name = DT.Rows[0]["MegaBatchID"].ToString() + ", " + DT.Rows[0]["BatchID"] + ", " + MainOrderID;
            }
            return name;
        }

        private string GetDecorName(int ID)
        {
            DataRow[] rows = DecorDT.Select("DecorID=" + ID);
            if (rows.Count() > 0)
                return rows[0]["Name"].ToString();
            else
                return string.Empty;
        }

        public string GetColorName(int ColorID)
        {
            string ColorName = string.Empty;
            try
            {
                DataRow[] Rows = FrameColorsDataTable.Select("ColorID = " + ColorID);
                ColorName = Rows[0]["ColorName"].ToString();
            }
            catch
            {
                return string.Empty;
            }
            return ColorName;
        }

        public string GetPatinaName(int PatinaID)
        {
            string FrontType = string.Empty;
            try
            {
                DataRow[] Rows = PatinaDataTable.Select("PatinaID = " + PatinaID);
                FrontType = Rows[0]["PatinaName"].ToString();
            }
            catch
            {
                return string.Empty;
            }
            return FrontType;
        }

        private void GetBagetWithAngleOrders(ref DataTable DestinationDT, int WorkAssignmentID, int FactoryID)
        {
            string SelectCommand = string.Empty;
            DataTable DT = new DataTable();

            SelectCommand = @"SELECT DecorOrderID, MainOrderID, ProductID, DecorID, DecorConfigID, ColorID, PatinaID,
                Height, Length, Width, LeftAngle, RightAngle, Count, Notes FROM DecorOrders
                WHERE FactoryID=" + FactoryID + " AND ProductID IN (1) AND (LeftAngle<>0 OR RightAngle<>0) AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE ProfilWorkAssignmentID=" + WorkAssignmentID + "))";
            if (FactoryID == 2)
                SelectCommand = @"SELECT DecorOrderID, MainOrderID, ProductID, DecorID, DecorConfigID, ColorID, PatinaID,
                    Height, Length, Width, LeftAngle, RightAngle, Count, Notes FROM DecorOrders
                    WHERE FactoryID=" + FactoryID + " AND ProductID IN (1) AND (LeftAngle<>0 OR RightAngle<>0) AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE TPSWorkAssignmentID=" + WorkAssignmentID + "))";

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.ZOVOrdersConnectionString))
            {
                DA.Fill(DT);
            }
            foreach (DataRow item in DT.Rows)
            {
                DataRow NewRow = DestinationDT.NewRow();
                NewRow.ItemArray = item.ItemArray;
                NewRow["GroupType"] = 0;
                DestinationDT.Rows.Add(NewRow);
            }

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                DT.Clear();
                DA.Fill(DT);
            }
            foreach (DataRow item in DT.Rows)
            {
                DataRow NewRow = DestinationDT.NewRow();
                NewRow.ItemArray = item.ItemArray;
                NewRow["GroupType"] = 1;
                DestinationDT.Rows.Add(NewRow);
            }
        }

        private void GetNotArchDecorOrders(ref DataTable DestinationDT, int WorkAssignmentID, int FactoryID)
        {
            string SelectCommand = string.Empty;
            DataTable DT = new DataTable();

            SelectCommand = @"SELECT DecorOrderID, MainOrderID, ProductID, DecorID, DecorConfigID, ColorID, PatinaID,
                Height, Length, Width, Count, Notes FROM DecorOrders
                WHERE FactoryID=" + FactoryID + " AND NOT (ProductID IN (1) AND (LeftAngle<>0 OR RightAngle<>0)) AND ProductID NOT IN (14, 31, 4, 18, 32, 10, 11, 12) AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE ProfilWorkAssignmentID=" + WorkAssignmentID + "))";
            if (FactoryID == 2)
                SelectCommand = @"SELECT DecorOrderID, MainOrderID, ProductID, DecorID, DecorConfigID, ColorID, PatinaID,
                    Height, Length, Width, Count, Notes FROM DecorOrders
                    WHERE FactoryID=" + FactoryID + " AND NOT (ProductID IN (1) AND (LeftAngle<>0 OR RightAngle<>0)) AND ProductID NOT IN (14, 31, 4, 18, 32, 10, 11, 12) AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE TPSWorkAssignmentID=" + WorkAssignmentID + "))";

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.ZOVOrdersConnectionString))
            {
                DA.Fill(DT);
            }
            foreach (DataRow item in DT.Rows)
            {
                DataRow NewRow = DestinationDT.NewRow();
                NewRow.ItemArray = item.ItemArray;
                NewRow["GroupType"] = 0;
                DestinationDT.Rows.Add(NewRow);
            }

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                DT.Clear();
                DA.Fill(DT);
            }
            foreach (DataRow item in DT.Rows)
            {
                DataRow NewRow = DestinationDT.NewRow();
                NewRow.ItemArray = item.ItemArray;
                NewRow["GroupType"] = 1;
                DestinationDT.Rows.Add(NewRow);
            }
        }

        private void GetArchDecorOrders(ref DataTable DestinationDT, int WorkAssignmentID, int FactoryID)
        {
            string SelectCommand = string.Empty;
            DataTable DT = new DataTable();

            SelectCommand = @"SELECT DecorOrderID, MainOrderID, ProductID, DecorID, DecorConfigID, ColorID, PatinaID,
                Height, Length, Width, Count, Notes FROM DecorOrders
                WHERE FactoryID=" + FactoryID + " AND ProductID IN (31, 4, 18, 32) AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE ProfilWorkAssignmentID=" + WorkAssignmentID + "))";
            if (FactoryID == 2)
                SelectCommand = @"SELECT DecorOrderID, MainOrderID, ProductID, DecorID, DecorConfigID, ColorID, PatinaID,
                    Height, Length, Width, Count, Notes FROM DecorOrders
                    WHERE FactoryID=" + FactoryID + " AND ProductID IN (31, 4, 18, 32) AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE TPSWorkAssignmentID=" + WorkAssignmentID + "))";

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.ZOVOrdersConnectionString))
            {
                DA.Fill(DT);
            }
            foreach (DataRow item in DT.Rows)
            {
                DataRow NewRow = DestinationDT.NewRow();
                NewRow.ItemArray = item.ItemArray;
                NewRow["GroupType"] = 0;
                DestinationDT.Rows.Add(NewRow);
            }

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                DT.Clear();
                DA.Fill(DT);
            }
            foreach (DataRow item in DT.Rows)
            {
                DataRow NewRow = DestinationDT.NewRow();
                NewRow.ItemArray = item.ItemArray;
                NewRow["GroupType"] = 1;
                DestinationDT.Rows.Add(NewRow);
            }
        }

        private void GetStrapsDecorOrders(ref DataTable DestinationDT, int WorkAssignmentID, int FactoryID)
        {
            string SelectCommand = string.Empty;
            DataTable DT = new DataTable();

            SelectCommand = @"SELECT DecorOrderID, MainOrderID, ProductID, DecorID, DecorConfigID, ColorID, PatinaID,
                Height, Length, Width, Count, Notes FROM DecorOrders
                WHERE FactoryID=" + FactoryID + " AND ProductID IN (14) AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE ProfilWorkAssignmentID=" + WorkAssignmentID + "))";
            if (FactoryID == 2)
                SelectCommand = @"SELECT DecorOrderID, MainOrderID, ProductID, DecorID, DecorConfigID, ColorID, PatinaID,
                    Height, Length, Width, Count, Notes FROM DecorOrders
                    WHERE FactoryID=" + FactoryID + " AND ProductID IN (14) AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE TPSWorkAssignmentID=" + WorkAssignmentID + "))";

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.ZOVOrdersConnectionString))
            {
                DA.Fill(DT);
            }
            foreach (DataRow item in DT.Rows)
            {
                DataRow NewRow = DestinationDT.NewRow();
                NewRow.ItemArray = item.ItemArray;
                NewRow["GroupType"] = 0;
                DestinationDT.Rows.Add(NewRow);
            }

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                DT.Clear();
                DA.Fill(DT);
            }
            foreach (DataRow item in DT.Rows)
            {
                DataRow NewRow = DestinationDT.NewRow();
                NewRow.ItemArray = item.ItemArray;
                NewRow["GroupType"] = 1;
                DestinationDT.Rows.Add(NewRow);
            }
        }

        private void GetGridsDecorOrders(ref DataTable DestinationDT, int WorkAssignmentID, int FactoryID)
        {
            string SelectCommand = string.Empty;
            DataTable DT = new DataTable();

            SelectCommand = @"SELECT DecorOrderID, MainOrderID, ProductID, DecorID, DecorConfigID, ColorID, PatinaID,
                Height, Length, Width, Count, Notes FROM DecorOrders
                WHERE FactoryID=" + FactoryID + " AND ProductID IN (10, 11, 12) AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE ProfilWorkAssignmentID=" + WorkAssignmentID + "))";
            if (FactoryID == 2)
                SelectCommand = @"SELECT DecorOrderID, MainOrderID, ProductID, DecorID, DecorConfigID, ColorID, PatinaID,
                    Height, Length, Width, Count, Notes FROM DecorOrders
                    WHERE FactoryID=" + FactoryID + " AND ProductID IN (10, 11, 12) AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE TPSWorkAssignmentID=" + WorkAssignmentID + "))";

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.ZOVOrdersConnectionString))
            {
                DA.Fill(DT);
            }
            foreach (DataRow item in DT.Rows)
            {
                DataRow NewRow = DestinationDT.NewRow();
                NewRow.ItemArray = item.ItemArray;
                NewRow["GroupType"] = 0;
                DestinationDT.Rows.Add(NewRow);
            }

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                DT.Clear();
                DA.Fill(DT);
            }
            foreach (DataRow item in DT.Rows)
            {
                DataRow NewRow = DestinationDT.NewRow();
                NewRow.ItemArray = item.ItemArray;
                NewRow["GroupType"] = 1;
                DestinationDT.Rows.Add(NewRow);
            }
        }

        private DataTable DistMainOrdersTable(DataTable SourceDT, bool OrderASC)
        {
            int MainOrderID = 0;
            DataTable DT = new DataTable();
            DT.Columns.Add(new DataColumn("MainOrderID", Type.GetType("System.Int32")));
            DT.Columns.Add(new DataColumn("GroupType", Type.GetType("System.Int32")));
            foreach (DataRow Row in SourceDT.Rows)
            {
                if (int.TryParse(Row["MainOrderID"].ToString(), out MainOrderID))
                {
                    DataRow NewRow = DT.NewRow();
                    NewRow["MainOrderID"] = MainOrderID;
                    NewRow["GroupType"] = Convert.ToInt32(Row["GroupType"]);
                    DT.Rows.Add(NewRow);
                }
            }
            using (DataView DV = new DataView(DT.Copy()))
            {
                DT.Clear();
                if (OrderASC)
                    DV.Sort = "MainOrderID ASC";
                else
                    DV.Sort = "MainOrderID DESC";
                DT = DV.ToTable(true, new string[] { "MainOrderID", "GroupType" });
            }
            return DT;
        }

        private DataTable DistMainOrdersTable(DataTable SourceDT1, DataTable SourceDT2, DataTable SourceDT3, DataTable SourceDT4, DataTable SourceDT5,
            DataTable SourceDT6, DataTable SourceDT7, DataTable SourceDT8, DataTable SourceDT9, bool OrderASC)
        {
            int MainOrderID = 0;
            DataTable DT = new DataTable();
            DT.Columns.Add(new DataColumn("MainOrderID", Type.GetType("System.Int32")));
            DT.Columns.Add(new DataColumn("GroupType", Type.GetType("System.Int32")));
            foreach (DataRow Row in SourceDT1.Rows)
            {
                if (int.TryParse(Row["MainOrderID"].ToString(), out MainOrderID))
                {
                    DataRow NewRow = DT.NewRow();
                    NewRow["MainOrderID"] = MainOrderID;
                    NewRow["GroupType"] = Convert.ToInt32(Row["GroupType"]);
                    DT.Rows.Add(NewRow);
                }
            }
            foreach (DataRow Row in SourceDT2.Rows)
            {
                if (int.TryParse(Row["MainOrderID"].ToString(), out MainOrderID))
                {
                    DataRow NewRow = DT.NewRow();
                    NewRow["MainOrderID"] = MainOrderID;
                    NewRow["GroupType"] = Convert.ToInt32(Row["GroupType"]);
                    DT.Rows.Add(NewRow);
                }
            }
            foreach (DataRow Row in SourceDT3.Rows)
            {
                if (int.TryParse(Row["MainOrderID"].ToString(), out MainOrderID))
                {
                    DataRow NewRow = DT.NewRow();
                    NewRow["MainOrderID"] = MainOrderID;
                    NewRow["GroupType"] = Convert.ToInt32(Row["GroupType"]);
                    DT.Rows.Add(NewRow);
                }
            }
            foreach (DataRow Row in SourceDT4.Rows)
            {
                if (int.TryParse(Row["MainOrderID"].ToString(), out MainOrderID))
                {
                    DataRow NewRow = DT.NewRow();
                    NewRow["MainOrderID"] = MainOrderID;
                    NewRow["GroupType"] = Convert.ToInt32(Row["GroupType"]);
                    DT.Rows.Add(NewRow);
                }
            }
            foreach (DataRow Row in SourceDT5.Rows)
            {
                if (int.TryParse(Row["MainOrderID"].ToString(), out MainOrderID))
                {
                    DataRow NewRow = DT.NewRow();
                    NewRow["MainOrderID"] = MainOrderID;
                    NewRow["GroupType"] = Convert.ToInt32(Row["GroupType"]);
                    DT.Rows.Add(NewRow);
                }
            }
            foreach (DataRow Row in SourceDT6.Rows)
            {
                if (int.TryParse(Row["MainOrderID"].ToString(), out MainOrderID))
                {
                    DataRow NewRow = DT.NewRow();
                    NewRow["MainOrderID"] = MainOrderID;
                    NewRow["GroupType"] = Convert.ToInt32(Row["GroupType"]);
                    DT.Rows.Add(NewRow);
                }
            }
            foreach (DataRow Row in SourceDT7.Rows)
            {
                if (int.TryParse(Row["MainOrderID"].ToString(), out MainOrderID))
                {
                    DataRow NewRow = DT.NewRow();
                    NewRow["MainOrderID"] = MainOrderID;
                    NewRow["GroupType"] = Convert.ToInt32(Row["GroupType"]);
                    DT.Rows.Add(NewRow);
                }
            }
            foreach (DataRow Row in SourceDT8.Rows)
            {
                if (int.TryParse(Row["MainOrderID"].ToString(), out MainOrderID))
                {
                    DataRow NewRow = DT.NewRow();
                    NewRow["MainOrderID"] = MainOrderID;
                    NewRow["GroupType"] = Convert.ToInt32(Row["GroupType"]);
                    DT.Rows.Add(NewRow);
                }
            }
            foreach (DataRow Row in SourceDT9.Rows)
            {
                if (int.TryParse(Row["MainOrderID"].ToString(), out MainOrderID))
                {
                    DataRow NewRow = DT.NewRow();
                    NewRow["MainOrderID"] = MainOrderID;
                    NewRow["GroupType"] = Convert.ToInt32(Row["GroupType"]);
                    DT.Rows.Add(NewRow);
                }
            }
            using (DataView DV = new DataView(DT.Copy()))
            {
                DT.Clear();
                if (OrderASC)
                    DV.Sort = "MainOrderID ASC";
                else
                    DV.Sort = "MainOrderID DESC";
                DT = DV.ToTable(true, new string[] { "MainOrderID", "GroupType" });
            }
            return DT;
        }

        private void AssemblyBagetWithAngleCollect(DataTable SourceDT, ref DataTable DestinationDT)
        {
            DataTable DT1 = new DataTable();

            using (DataView DV = new DataView(SourceDT))
            {
                DT1 = DV.ToTable(true, new string[] { "DecorID", "ColorID", "PatinaID", "Length", "Height", "Width", "LeftAngle", "RightAngle", "Notes" });
            }
            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                int Count = 0;
                string filter = "DecorID=" + Convert.ToInt32(DT1.Rows[i]["DecorID"]) + " AND ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                    " AND Length=" + Convert.ToInt32(DT1.Rows[i]["Length"]) + " AND Height=" + Convert.ToInt32(DT1.Rows[i]["Height"]) +
                    " AND Width=" + Convert.ToInt32(DT1.Rows[i]["Width"]) +
                    " AND LeftAngle=" + Convert.ToInt32(DT1.Rows[i]["LeftAngle"]) +
                    " AND RightAngle=" + Convert.ToInt32(DT1.Rows[i]["RightAngle"]) +
                    " AND (Notes='' OR Notes IS NULL)";
                if (DT1.Rows[i]["Notes"] != DBNull.Value && DT1.Rows[i]["Notes"].ToString().Length > 0)
                {
                    filter = "DecorID=" + Convert.ToInt32(DT1.Rows[i]["DecorID"]) + " AND ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                      " AND Length=" + Convert.ToInt32(DT1.Rows[i]["Length"]) + " AND Height=" + Convert.ToInt32(DT1.Rows[i]["Height"]) +
                      " AND Width=" + Convert.ToInt32(DT1.Rows[i]["Width"]) +
                    " AND LeftAngle=" + Convert.ToInt32(DT1.Rows[i]["LeftAngle"]) +
                    " AND RightAngle=" + Convert.ToInt32(DT1.Rows[i]["RightAngle"]) +
                    " AND Notes='" + DT1.Rows[i]["Notes"] + "'";
                }
                DataRow[] rows = SourceDT.Select(filter);
                if (rows.Count() == 0)
                    continue;

                foreach (DataRow item in rows)
                    Count += Convert.ToInt32(item["Count"]);

                string Color = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                if (Convert.ToInt32(DT1.Rows[i]["PatinaID"]) != -1)
                    Color += " " + GetPatinaName(Convert.ToInt32(DT1.Rows[i]["PatinaID"]));

                DataRow NewRow = DestinationDT.NewRow();
                NewRow["DecorID"] = Convert.ToInt32(DT1.Rows[i]["DecorID"]);
                NewRow["Name"] = GetDecorName(Convert.ToInt32(DT1.Rows[i]["DecorID"]));
                if (HasParameter(Convert.ToInt32(rows[0]["ProductID"]), "ColorID"))
                    NewRow["Color"] = Color;
                //if (HasParameter(Convert.ToInt32(rows[0]["ProductID"]), "Height"))
                //    NewRow["Height"] = DT1.Rows[i]["Height"];
                //if (HasParameter(Convert.ToInt32(rows[0]["ProductID"]), "Length"))
                //    NewRow["Height"] = DT1.Rows[i]["Length"];
                //if (HasParameter(Convert.ToInt32(rows[0]["ProductID"]), "Width"))
                //    NewRow["Width"] = DT1.Rows[i]["Width"];

                if (HasParameter(Convert.ToInt32(rows[0]["ProductID"]), "Height"))
                {
                    if (Convert.ToInt32(DT1.Rows[i]["Height"]) != -1)
                        NewRow["Height"] = DT1.Rows[i]["Height"];
                    if (Convert.ToInt32(DT1.Rows[i]["Length"]) != -1)
                        NewRow["Height"] = DT1.Rows[i]["Length"];
                }
                if (HasParameter(Convert.ToInt32(rows[0]["ProductID"]), "Length"))
                {
                    if (Convert.ToInt32(DT1.Rows[i]["Length"]) != -1)
                        NewRow["Height"] = DT1.Rows[i]["Length"];
                    if (Convert.ToInt32(DT1.Rows[i]["Height"]) != -1)
                        NewRow["Height"] = DT1.Rows[i]["Height"];
                }
                if (HasParameter(Convert.ToInt32(rows[0]["ProductID"]), "Width") && Convert.ToInt32(DT1.Rows[i]["Width"]) != -1)
                    NewRow["Width"] = DT1.Rows[i]["Width"];

                NewRow["Count"] = Count;
                NewRow["LeftAngle"] = DT1.Rows[i]["LeftAngle"];
                NewRow["RightAngle"] = DT1.Rows[i]["RightAngle"];
                NewRow["Notes"] = DT1.Rows[i]["Notes"];
                DestinationDT.Rows.Add(NewRow);
            }
            using (DataView DV = new DataView(DestinationDT.Copy()))
            {
                DV.Sort = "Name, Color, Height, Width, LeftAngle, RightAngle";
                DestinationDT.Clear();
                DestinationDT = DV.ToTable();
            }
        }

        private void AssemblyDecorCollect(DataTable SourceDT, ref DataTable DestinationDT)
        {
            DataTable DT1 = new DataTable();

            using (DataView DV = new DataView(SourceDT))
            {
                DT1 = DV.ToTable(true, new string[] { "DecorID", "ColorID", "PatinaID", "Length", "Height", "Width", "Notes" });
            }
            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                int Count = 0;
                string filter = "DecorID=" + Convert.ToInt32(DT1.Rows[i]["DecorID"]) + " AND ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                    " AND Length=" + Convert.ToInt32(DT1.Rows[i]["Length"]) + " AND Height=" + Convert.ToInt32(DT1.Rows[i]["Height"]) +
                    " AND Width=" + Convert.ToInt32(DT1.Rows[i]["Width"]) + " AND (Notes='' OR Notes IS NULL)";
                if (DT1.Rows[i]["Notes"] != DBNull.Value && DT1.Rows[i]["Notes"].ToString().Length > 0)
                {
                    filter = "DecorID=" + Convert.ToInt32(DT1.Rows[i]["DecorID"]) + " AND ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                      " AND Length=" + Convert.ToInt32(DT1.Rows[i]["Length"]) + " AND Height=" + Convert.ToInt32(DT1.Rows[i]["Height"]) +
                      " AND Width=" + Convert.ToInt32(DT1.Rows[i]["Width"]) + " AND Notes='" + DT1.Rows[i]["Notes"] + "'";
                }
                DataRow[] rows = SourceDT.Select(filter);
                if (rows.Count() == 0)
                    continue;

                foreach (DataRow item in rows)
                    Count += Convert.ToInt32(item["Count"]);

                string Color = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                if (Convert.ToInt32(DT1.Rows[i]["PatinaID"]) != -1)
                    Color += " " + GetPatinaName(Convert.ToInt32(DT1.Rows[i]["PatinaID"]));

                DataRow NewRow = DestinationDT.NewRow();
                NewRow["DecorID"] = Convert.ToInt32(DT1.Rows[i]["DecorID"]);
                NewRow["Name"] = GetDecorName(Convert.ToInt32(DT1.Rows[i]["DecorID"]));
                if (HasParameter(Convert.ToInt32(rows[0]["ProductID"]), "ColorID"))
                    NewRow["Color"] = Color;
                //if (HasParameter(Convert.ToInt32(rows[0]["ProductID"]), "Height"))
                //    NewRow["Height"] = DT1.Rows[i]["Height"];
                //if (HasParameter(Convert.ToInt32(rows[0]["ProductID"]), "Length"))
                //    NewRow["Height"] = DT1.Rows[i]["Length"];
                //if (HasParameter(Convert.ToInt32(rows[0]["ProductID"]), "Width"))
                //    NewRow["Width"] = DT1.Rows[i]["Width"];

                if (HasParameter(Convert.ToInt32(rows[0]["ProductID"]), "Height"))
                {
                    if (Convert.ToInt32(DT1.Rows[i]["Height"]) != -1)
                        NewRow["Height"] = DT1.Rows[i]["Height"];
                    if (Convert.ToInt32(DT1.Rows[i]["Length"]) != -1)
                        NewRow["Height"] = DT1.Rows[i]["Length"];
                }
                if (HasParameter(Convert.ToInt32(rows[0]["ProductID"]), "Length"))
                {
                    if (Convert.ToInt32(DT1.Rows[i]["Length"]) != -1)
                        NewRow["Height"] = DT1.Rows[i]["Length"];
                    if (Convert.ToInt32(DT1.Rows[i]["Height"]) != -1)
                        NewRow["Height"] = DT1.Rows[i]["Height"];
                }
                if (HasParameter(Convert.ToInt32(rows[0]["ProductID"]), "Width") && Convert.ToInt32(DT1.Rows[i]["Width"]) != -1)
                    NewRow["Width"] = DT1.Rows[i]["Width"];

                NewRow["Count"] = Count;
                NewRow["Notes"] = DT1.Rows[i]["Notes"];
                DestinationDT.Rows.Add(NewRow);
            }
            using (DataView DV = new DataView(DestinationDT.Copy()))
            {
                DV.Sort = "Name, Color, Height, Width";
                DestinationDT.Clear();
                DestinationDT = DV.ToTable();
            }
        }

        public void ClearOrders()
        {
            BagetWithAngelOrdersDT.Clear();
            NotArchDecorOrdersDT.Clear();
            ArchDecorOrdersDT.Clear();
            StrapsDecorOrdersDT.Clear();
            GridsDecorOrdersDT.Clear();
        }

        public bool GetOrders(int WorkAssignmentID, int FactoryID)
        {
            GetNotArchDecorOrders(ref NotArchDecorOrdersDT, WorkAssignmentID, FactoryID);
            GetBagetWithAngleOrders(ref BagetWithAngelOrdersDT, WorkAssignmentID, FactoryID);
            GetArchDecorOrders(ref ArchDecorOrdersDT, WorkAssignmentID, FactoryID);
            GetStrapsDecorOrders(ref StrapsDecorOrdersDT, WorkAssignmentID, FactoryID);
            GetGridsDecorOrders(ref GridsDecorOrdersDT, WorkAssignmentID, FactoryID);

            if (BagetWithAngelOrdersDT.Rows.Count == 0 && NotArchDecorOrdersDT.Rows.Count == 0 && ArchDecorOrdersDT.Rows.Count == 0 && StrapsDecorOrdersDT.Rows.Count == 0 && GridsDecorOrdersDT.Rows.Count == 0)
                return false;
            else
                return true;

        }

        public void GetCurrentDate()
        {
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT GETDATE()", ConnectionStrings.UsersConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    DA.Fill(DT);

                    CurrentDate = Convert.ToDateTime(DT.Rows[0][0]);
                }
            }
        }

        public void CreateExcel(int WorkAssignmentID, string ClientName, string BatchName, ref string sSourceFileName)
        {
            GetCurrentDate();

            System.Diagnostics.Stopwatch sw = new System.Diagnostics.Stopwatch();
            sw.Start();
            HSSFWorkbook hssfworkbook = new HSSFWorkbook();

            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = "NPOI Team";
            hssfworkbook.DocumentSummaryInformation = dsi;

            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
            si.Subject = "NPOI SDK Example";
            hssfworkbook.SummaryInformation = si;

            #region Create fonts and styles

            HSSFFont Calibri11F = hssfworkbook.CreateFont();
            Calibri11F.FontHeightInPoints = 11;
            Calibri11F.FontName = "Calibri";

            HSSFFont CalibriBold11F = hssfworkbook.CreateFont();
            CalibriBold11F.FontHeightInPoints = 11;
            CalibriBold11F.Boldweight = HSSFFont.BOLDWEIGHT_BOLD;
            CalibriBold11F.FontName = "Calibri";

            HSSFFont Serif8F = hssfworkbook.CreateFont();
            Serif8F.FontHeightInPoints = 8;
            Serif8F.FontName = "MS Sans Serif";

            HSSFFont Serif10F = hssfworkbook.CreateFont();
            Serif10F.FontHeightInPoints = 10;
            Serif10F.FontName = "MS Sans Serif";

            HSSFFont SerifBold10F = hssfworkbook.CreateFont();
            SerifBold10F.FontHeightInPoints = 10;
            SerifBold10F.FontName = "MS Sans Serif";
            SerifBold10F.Boldweight = HSSFFont.BOLDWEIGHT_BOLD;

            HSSFCellStyle Calibri11CS = hssfworkbook.CreateCellStyle();
            Calibri11CS.SetFont(Calibri11F);

            HSSFCellStyle CalibriBold11CS = hssfworkbook.CreateCellStyle();
            CalibriBold11CS.SetFont(CalibriBold11F);

            HSSFCellStyle TableHeaderCS = hssfworkbook.CreateCellStyle();
            TableHeaderCS.Alignment = HSSFCellStyle.ALIGN_LEFT;
            TableHeaderCS.BorderBottom = HSSFCellStyle.BORDER_THIN;
            TableHeaderCS.BottomBorderColor = HSSFColor.BLACK.index;
            TableHeaderCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            TableHeaderCS.LeftBorderColor = HSSFColor.BLACK.index;
            TableHeaderCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            TableHeaderCS.RightBorderColor = HSSFColor.BLACK.index;
            TableHeaderCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            TableHeaderCS.TopBorderColor = HSSFColor.BLACK.index;
            TableHeaderCS.SetFont(Serif8F);

            HSSFCellStyle TableHeaderDecCS = hssfworkbook.CreateCellStyle();
            TableHeaderDecCS.Alignment = HSSFCellStyle.ALIGN_LEFT;
            TableHeaderDecCS.BorderBottom = HSSFCellStyle.BORDER_THIN;
            TableHeaderDecCS.BottomBorderColor = HSSFColor.BLACK.index;
            TableHeaderDecCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            TableHeaderDecCS.LeftBorderColor = HSSFColor.BLACK.index;
            TableHeaderDecCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            TableHeaderDecCS.RightBorderColor = HSSFColor.BLACK.index;
            TableHeaderDecCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            TableHeaderDecCS.TopBorderColor = HSSFColor.BLACK.index;
            TableHeaderDecCS.SetFont(Serif8F);

            HSSFCellStyle WorkerColumnCS = hssfworkbook.CreateCellStyle();
            WorkerColumnCS.Alignment = HSSFCellStyle.ALIGN_RIGHT;
            WorkerColumnCS.BorderBottom = HSSFCellStyle.BORDER_THIN;
            WorkerColumnCS.BottomBorderColor = HSSFColor.BLACK.index;
            WorkerColumnCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            WorkerColumnCS.LeftBorderColor = HSSFColor.BLACK.index;
            WorkerColumnCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            WorkerColumnCS.RightBorderColor = HSSFColor.BLACK.index;
            WorkerColumnCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            WorkerColumnCS.TopBorderColor = HSSFColor.BLACK.index;
            WorkerColumnCS.SetFont(Serif10F);

            HSSFCellStyle TableContentCS = hssfworkbook.CreateCellStyle();
            TableContentCS.BorderBottom = HSSFCellStyle.BORDER_THIN;
            TableContentCS.BottomBorderColor = HSSFColor.BLACK.index;
            TableContentCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            TableContentCS.LeftBorderColor = HSSFColor.BLACK.index;
            TableContentCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            TableContentCS.RightBorderColor = HSSFColor.BLACK.index;
            TableContentCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            TableContentCS.TopBorderColor = HSSFColor.BLACK.index;
            TableContentCS.SetFont(SerifBold10F);

            HSSFCellStyle TotamAmountCS = hssfworkbook.CreateCellStyle();
            TotamAmountCS.Alignment = HSSFCellStyle.ALIGN_RIGHT;
            TotamAmountCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            TotamAmountCS.LeftBorderColor = HSSFColor.BLACK.index;
            TotamAmountCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            TotamAmountCS.RightBorderColor = HSSFColor.BLACK.index;
            TotamAmountCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            TotamAmountCS.TopBorderColor = HSSFColor.BLACK.index;
            TotamAmountCS.SetFont(SerifBold10F);

            HSSFCellStyle TotamAmountDecCS = hssfworkbook.CreateCellStyle();
            TotamAmountDecCS.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.000");
            TotamAmountCS.Alignment = HSSFCellStyle.ALIGN_RIGHT;
            TotamAmountCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            TotamAmountCS.LeftBorderColor = HSSFColor.BLACK.index;
            TotamAmountCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            TotamAmountCS.RightBorderColor = HSSFColor.BLACK.index;
            TotamAmountCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            TotamAmountCS.TopBorderColor = HSSFColor.BLACK.index;
            TotamAmountCS.SetFont(SerifBold10F);

            #endregion

            BagetWithAngleAssemblyByMainOrderToExcel(ref hssfworkbook, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, BatchName);

            NotArchDecorAssemblyByMainOrderToExcel(ref hssfworkbook, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, BatchName);

            ArchDecorAssemblyByMainOrderToExcel(ref hssfworkbook, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, BatchName);

            StrapsAssemblyByMainOrderToExcel(ref hssfworkbook, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, BatchName);

            GridsDecorAssemblyByMainOrderToExcel(ref hssfworkbook, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, BatchName);

            string FileName = WorkAssignmentID + " " + BatchName;
            string tempFolder = @"\\192.168.1.6\Public\_ДЕЙСТВУЮЩИЕ\ПРОИЗВОДСТВО\ТПС\инфиниум\";
            //string tempFolder = System.Environment.GetEnvironmentVariable("TEMP");

            string CurrentMonthName = DateTime.Now.ToString("MMMM");
            tempFolder = Path.Combine(tempFolder, CurrentMonthName);
            if (!(Directory.Exists(tempFolder)))
            {
                Directory.CreateDirectory(tempFolder);
            }
            FileInfo file = new FileInfo(tempFolder + @"\" + FileName + ".xls");
            int j = 1;
            while (file.Exists == true)
            {
                file = new FileInfo(tempFolder + @"\" + FileName + "(" + j++ + ").xls");
            }

            FileStream NewFile = new FileStream(file.FullName, FileMode.Create);
            hssfworkbook.Write(NewFile);
            NewFile.Close();

            sw.Stop();
            System.Diagnostics.Process.Start(file.FullName);

            //string sSourceFolder = System.Environment.GetEnvironmentVariable("TEMP");
            //string sFolderPath = "Общие файлы/Производство/Задания в работу";
            //string sDestFolder = Configs.DocumentsPath + sFolderPath;
            //sSourceFileName = GetFileName(sDestFolder, BatchName);

            //FileInfo file = new FileInfo(sSourceFolder + @"\" + sSourceFileName);
            //FileStream NewFile = new FileStream(file.FullName, FileMode.Create);
            //hssfworkbook.Write(NewFile);
            //NewFile.Close();

        }

        private string GetFileName(string sDestFolder, string ExcelName)
        {
            string sExtension = ".xls";
            string sFileName = ExcelName;

            int j = 1;
            while (FM.FileExist(sDestFolder + "/" + sFileName + sExtension, Configs.FTPType))
            {
                sFileName = ExcelName + "(" + j++ + ")";
            }
            sFileName = sFileName + sExtension;
            return sFileName;
        }

        private void BagetWithAngleAssemblyByMainOrderToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            int WorkAssignmentID, string BatchName)
        {
            if (BagetWithAngelOrdersDT.Rows.Count == 0)
                return;

            DataTable DistMainOrdersDT = DistMainOrdersTable(BagetWithAngelOrdersDT, true);
            DataTable DT = BagetWithAngelOrdersDT.Clone();
            DataTable ZOVOrdersNames = new DataTable();
            DataTable MarketOrdersNames = new DataTable();
            string MainOrdersID = string.Empty;
            string ClientName = string.Empty;
            string OrderName = string.Empty;
            string Notes = string.Empty;

            if (DistMainOrdersDT.Select("GroupType=0").Count() > 0)
            {
                MainOrdersID = string.Empty;
                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=0"))
                    MainOrdersID += Convert.ToInt32(item["MainOrderID"]) + ",";
                MainOrdersID = MainOrdersID.Substring(0, MainOrdersID.Length - 1);
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT ClientName, DocNumber, MainOrderID FROM MainOrders" +
                    " INNER JOIN infiniu2_zovreference.dbo.Clients ON MainOrders.ClientID=infiniu2_zovreference.dbo.Clients.ClientID" +
                    " WHERE MainOrderID IN (" + MainOrdersID + ")", ConnectionStrings.ZOVOrdersConnectionString))
                {
                    DA.Fill(ZOVOrdersNames);
                }

                int RowIndex = 0;
                HSSFSheet sheet1 = hssfworkbook.CreateSheet("Не арки ЗОВ");
                sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

                sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
                sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
                sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
                sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

                sheet1.SetColumnWidth(0, 30 * 256);
                sheet1.SetColumnWidth(1, 20 * 256);
                sheet1.SetColumnWidth(2, 9 * 256);
                sheet1.SetColumnWidth(3, 6 * 256);
                sheet1.SetColumnWidth(4, 6 * 256);
                sheet1.SetColumnWidth(5, 6 * 256);
                sheet1.SetColumnWidth(6, 6 * 256);

                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=0"))
                {
                    BagetWithAngleAssemblyDT.Clear();
                    DT.Clear();
                    int MainOrderID = Convert.ToInt32(item["MainOrderID"]);
                    DataRow[] rows = BagetWithAngelOrdersDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item1 in rows)
                        DT.Rows.Add(item1.ItemArray);
                    AssemblyBagetWithAngleCollect(DT, ref BagetWithAngleAssemblyDT);

                    DataRow[] CRows = ZOVOrdersNames.Select("MainOrderID=" + Convert.ToInt32(item["MainOrderID"]));
                    if (CRows.Count() > 0)
                    {
                        ClientName = CRows[0]["ClientName"].ToString();
                        OrderName = CRows[0]["DocNumber"].ToString();
                    }
                    if (BagetWithAngleAssemblyDT.Rows.Count > 0)
                    {
                        BagetWithAngleAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                             Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, BagetWithAngleAssemblyDT,
                             WorkAssignmentID, BatchName, ClientName, string.Empty, OrderName, Notes, ref RowIndex);
                        RowIndex++;
                        BagetWithAngleAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                            Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, BagetWithAngleAssemblyDT,
                            WorkAssignmentID, BatchName, ClientName, "ДУБЛЬ", OrderName, Notes, ref RowIndex);
                    }
                    RowIndex++;
                }
            }
            if (DistMainOrdersDT.Select("GroupType=1").Count() > 0)
            {
                MainOrdersID = string.Empty;
                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=1"))
                    MainOrdersID += Convert.ToInt32(item["MainOrderID"]) + ",";
                MainOrdersID = MainOrdersID.Substring(0, MainOrdersID.Length - 1);
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT ClientName, OrderNumber, MainOrderID, Notes FROM MainOrders" +
                    " INNER JOIN MegaOrders ON MainOrders.MegaOrderID=MegaOrders.MegaOrderID" +
                    " INNER JOIN infiniu2_marketingreference.dbo.Clients ON MegaOrders.ClientID=infiniu2_marketingreference.dbo.Clients.ClientID" +
                    " WHERE MainOrderID IN (" + MainOrdersID + ")", ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DA.Fill(MarketOrdersNames);
                }

                int RowIndex = 0;
                HSSFSheet sheet1 = hssfworkbook.CreateSheet("Багет с запилом Маркетинг");
                sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

                sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
                sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
                sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
                sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

                sheet1.SetColumnWidth(0, 30 * 256);
                sheet1.SetColumnWidth(1, 20 * 256);
                sheet1.SetColumnWidth(2, 9 * 256);
                sheet1.SetColumnWidth(3, 6 * 256);
                sheet1.SetColumnWidth(4, 6 * 256);
                sheet1.SetColumnWidth(5, 6 * 256);
                sheet1.SetColumnWidth(6, 6 * 256);

                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=1"))
                {
                    BagetWithAngleAssemblyDT.Clear();
                    DT.Clear();
                    int MainOrderID = Convert.ToInt32(item["MainOrderID"]);
                    DataRow[] rows = BagetWithAngelOrdersDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item1 in rows)
                        DT.Rows.Add(item1.ItemArray);
                    AssemblyBagetWithAngleCollect(DT, ref BagetWithAngleAssemblyDT);

                    DataRow[] CRows = MarketOrdersNames.Select("MainOrderID=" + MainOrderID);
                    if (CRows.Count() > 0)
                    {
                        ClientName = CRows[0]["ClientName"].ToString();
                        Notes = CRows[0]["Notes"].ToString();
                        OrderName = "№" + CRows[0]["OrderNumber"].ToString() + "-" + MainOrderID;
                    }
                    if (BagetWithAngleAssemblyDT.Rows.Count > 0)
                    {
                        BagetWithAngleAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                             Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, BagetWithAngleAssemblyDT,
                             WorkAssignmentID, BatchName, ClientName, string.Empty, OrderName, Notes, ref RowIndex);
                        RowIndex++;
                        BagetWithAngleAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                            Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, BagetWithAngleAssemblyDT,
                            WorkAssignmentID, BatchName, ClientName, "ДУБЛЬ", OrderName, Notes, ref RowIndex);
                    }
                    RowIndex++;
                }
            }
        }

        private void NotArchDecorAssemblyByMainOrderToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            int WorkAssignmentID, string BatchName)
        {
            if (NotArchDecorOrdersDT.Rows.Count == 0)
                return;

            DataTable DistMainOrdersDT = DistMainOrdersTable(NotArchDecorOrdersDT, true);
            DataTable DT = NotArchDecorOrdersDT.Clone();
            DataTable ZOVOrdersNames = new DataTable();
            DataTable MarketOrdersNames = new DataTable();
            string MainOrdersID = string.Empty;
            string ClientName = string.Empty;
            string OrderName = string.Empty;
            string Notes = string.Empty;

            if (DistMainOrdersDT.Select("GroupType=0").Count() > 0)
            {
                MainOrdersID = string.Empty;
                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=0"))
                    MainOrdersID += Convert.ToInt32(item["MainOrderID"]) + ",";
                MainOrdersID = MainOrdersID.Substring(0, MainOrdersID.Length - 1);
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT ClientName, DocNumber, MainOrderID FROM MainOrders" +
                    " INNER JOIN infiniu2_zovreference.dbo.Clients ON MainOrders.ClientID=infiniu2_zovreference.dbo.Clients.ClientID" +
                    " WHERE MainOrderID IN (" + MainOrdersID + ")", ConnectionStrings.ZOVOrdersConnectionString))
                {
                    DA.Fill(ZOVOrdersNames);
                }

                int RowIndex = 0;
                HSSFSheet sheet1 = hssfworkbook.CreateSheet("Не арки ЗОВ");
                sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

                sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
                sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
                sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
                sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

                sheet1.SetColumnWidth(0, 30 * 256);
                sheet1.SetColumnWidth(1, 20 * 256);
                sheet1.SetColumnWidth(2, 9 * 256);
                sheet1.SetColumnWidth(3, 6 * 256);
                sheet1.SetColumnWidth(4, 6 * 256);

                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=0"))
                {
                    DecorAssemblyDT.Clear();
                    DT.Clear();
                    int MainOrderID = Convert.ToInt32(item["MainOrderID"]);
                    DataRow[] rows = NotArchDecorOrdersDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item1 in rows)
                        DT.Rows.Add(item1.ItemArray);
                    AssemblyDecorCollect(DT, ref DecorAssemblyDT);

                    DataRow[] CRows = ZOVOrdersNames.Select("MainOrderID=" + Convert.ToInt32(item["MainOrderID"]));
                    if (CRows.Count() > 0)
                    {
                        ClientName = CRows[0]["ClientName"].ToString();
                        OrderName = CRows[0]["DocNumber"].ToString();
                    }
                    if (DecorAssemblyDT.Rows.Count > 0)
                    {
                        NotArchDecorAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                             Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DecorAssemblyDT,
                             WorkAssignmentID, BatchName, ClientName, string.Empty, OrderName, Notes, ref RowIndex);
                        RowIndex++;
                        NotArchDecorAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                            Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DecorAssemblyDT,
                            WorkAssignmentID, BatchName, ClientName, "ДУБЛЬ", OrderName, Notes, ref RowIndex);
                    }
                    RowIndex++;
                }
            }
            if (DistMainOrdersDT.Select("GroupType=1").Count() > 0)
            {
                MainOrdersID = string.Empty;
                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=1"))
                    MainOrdersID += Convert.ToInt32(item["MainOrderID"]) + ",";
                MainOrdersID = MainOrdersID.Substring(0, MainOrdersID.Length - 1);
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT ClientName, OrderNumber, MainOrderID, Notes FROM MainOrders" +
                    " INNER JOIN MegaOrders ON MainOrders.MegaOrderID=MegaOrders.MegaOrderID" +
                    " INNER JOIN infiniu2_marketingreference.dbo.Clients ON MegaOrders.ClientID=infiniu2_marketingreference.dbo.Clients.ClientID" +
                    " WHERE MainOrderID IN (" + MainOrdersID + ")", ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DA.Fill(MarketOrdersNames);
                }

                int RowIndex = 0;
                HSSFSheet sheet1 = hssfworkbook.CreateSheet("Не арки Маркетинг");
                sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

                sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
                sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
                sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
                sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

                sheet1.SetColumnWidth(0, 30 * 256);
                sheet1.SetColumnWidth(1, 20 * 256);
                sheet1.SetColumnWidth(2, 9 * 256);
                sheet1.SetColumnWidth(3, 6 * 256);
                sheet1.SetColumnWidth(4, 6 * 256);

                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=1"))
                {
                    DecorAssemblyDT.Clear();
                    DT.Clear();
                    int MainOrderID = Convert.ToInt32(item["MainOrderID"]);
                    DataRow[] rows = NotArchDecorOrdersDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item1 in rows)
                        DT.Rows.Add(item1.ItemArray);
                    AssemblyDecorCollect(DT, ref DecorAssemblyDT);

                    DataRow[] CRows = MarketOrdersNames.Select("MainOrderID=" + MainOrderID);
                    if (CRows.Count() > 0)
                    {
                        ClientName = CRows[0]["ClientName"].ToString();
                        Notes = CRows[0]["Notes"].ToString();
                        OrderName = "№" + CRows[0]["OrderNumber"].ToString() + "-" + MainOrderID;
                    }

                    if (DecorAssemblyDT.Rows.Count > 0)
                    {
                        NotArchDecorAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                             Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DecorAssemblyDT,
                             WorkAssignmentID, BatchName, ClientName, string.Empty, OrderName, Notes, ref RowIndex);
                        RowIndex++;
                        NotArchDecorAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                            Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DecorAssemblyDT,
                            WorkAssignmentID, BatchName, ClientName, "ДУБЛЬ", OrderName, Notes, ref RowIndex);
                    }
                    RowIndex++;
                }
            }
        }

        public void BagetWithAngleAssemblyToExcelSingly(ref HSSFWorkbook hssfworkbook, ref HSSFSheet sheet1,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            DataTable DT, int WorkAssignmentID, string BatchName, string ClientName, string OperationName, string OrderName, string Notes, ref int RowIndex)
        {
            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, ClientName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Заказ:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, OrderName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 4, OperationName);
            cell.CellStyle = CalibriBold11CS;
            if (Notes.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Примечание: " + Notes);
                cell.CellStyle = CalibriBold11CS;
            }

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Наименование");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Цвет");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Длин./Выс.");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Л. угол");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 5, "П. угол");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 6, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 7, "Прим.");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            int DecorID = -1;
            int AllTotalAmount = 0;
            int TotalAmount = 0;
            int DifferentDecorCount = 0;

            using (DataView DV = new DataView(DT))
            {
                DifferentDecorCount = DV.ToTable(true, new string[] { "DecorID" }).Rows.Count;
            }
            if (DT.Rows.Count > 0)
            {
                DecorID = Convert.ToInt32(DT.Rows[0]["DecorID"]);
            }

            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value)
                {
                    AllTotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    TotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                }
                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "DecorID")
                        continue;

                    Type t = DT.Rows[x][y].GetType();

                    if (t.Name == "Decimal")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToDouble(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderDecCS;
                        continue;
                    }
                    if (t.Name == "Int32")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToInt32(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (t.Name == "String" || t.Name == "DBNull")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(DT.Rows[x][y].ToString());
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }
                }

                if (x + 1 <= DT.Rows.Count - 1)
                {
                    if (DecorID != Convert.ToInt32(DT.Rows[x + 1]["DecorID"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "DecorID")
                                continue;

                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(6);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        DecorID = Convert.ToInt32(DT.Rows[x + 1]["DecorID"]);
                        TotalAmount = 0;
                        RowIndex++;
                    }
                }

                if (x == DT.Rows.Count - 1)
                {
                    if (DifferentDecorCount > 1)
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "DecorID")
                                continue;

                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(6);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;
                    }
                    RowIndex++;

                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "DecorID")
                            continue;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(6);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;
                }
                RowIndex++;
            }
        }

        public void NotArchDecorAssemblyToExcelSingly(ref HSSFWorkbook hssfworkbook, ref HSSFSheet sheet1,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            DataTable DT, int WorkAssignmentID, string BatchName, string ClientName, string OperationName, string OrderName, string Notes, ref int RowIndex)
        {
            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, ClientName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Заказ:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, OrderName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 4, OperationName);
            cell.CellStyle = CalibriBold11CS;
            if (Notes.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Примечание: " + Notes);
                cell.CellStyle = CalibriBold11CS;
            }

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Наименование");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Цвет");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Длин./Выс.");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 5, "Прим.");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            int DecorID = -1;
            int AllTotalAmount = 0;
            int TotalAmount = 0;
            int DifferentDecorCount = 0;

            using (DataView DV = new DataView(DT))
            {
                DifferentDecorCount = DV.ToTable(true, new string[] { "DecorID" }).Rows.Count;
            }
            if (DT.Rows.Count > 0)
            {
                DecorID = Convert.ToInt32(DT.Rows[0]["DecorID"]);
            }

            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value)
                {
                    AllTotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    TotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                }
                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "DecorID")
                        continue;

                    Type t = DT.Rows[x][y].GetType();

                    if (t.Name == "Decimal")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToDouble(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderDecCS;
                        continue;
                    }
                    if (t.Name == "Int32")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToInt32(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (t.Name == "String" || t.Name == "DBNull")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(DT.Rows[x][y].ToString());
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }
                }

                if (x + 1 <= DT.Rows.Count - 1)
                {
                    if (DecorID != Convert.ToInt32(DT.Rows[x + 1]["DecorID"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "DecorID")
                                continue;

                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        DecorID = Convert.ToInt32(DT.Rows[x + 1]["DecorID"]);
                        TotalAmount = 0;
                        RowIndex++;
                    }
                }

                if (x == DT.Rows.Count - 1)
                {
                    if (DifferentDecorCount > 1)
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "DecorID")
                                continue;

                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;
                    }
                    RowIndex++;

                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "DecorID")
                            continue;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;
                }
                RowIndex++;
            }
        }

        private void ArchDecorAssemblyByMainOrderToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            int WorkAssignmentID, string BatchName)
        {
            if (ArchDecorOrdersDT.Rows.Count == 0)
                return;

            DataTable DistMainOrdersDT = DistMainOrdersTable(ArchDecorOrdersDT, true);
            DataTable DT = ArchDecorOrdersDT.Clone();
            DataTable ZOVOrdersNames = new DataTable();
            DataTable MarketOrdersNames = new DataTable();
            string MainOrdersID = string.Empty;
            string ClientName = string.Empty;
            string OrderName = string.Empty;
            string Notes = string.Empty;

            if (DistMainOrdersDT.Select("GroupType=0").Count() > 0)
            {
                MainOrdersID = string.Empty;
                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=0"))
                    MainOrdersID += Convert.ToInt32(item["MainOrderID"]) + ",";
                MainOrdersID = MainOrdersID.Substring(0, MainOrdersID.Length - 1);
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT ClientName, DocNumber, MainOrderID FROM MainOrders" +
                    " INNER JOIN infiniu2_zovreference.dbo.Clients ON MainOrders.ClientID=infiniu2_zovreference.dbo.Clients.ClientID" +
                    " WHERE MainOrderID IN (" + MainOrdersID + ")", ConnectionStrings.ZOVOrdersConnectionString))
                {
                    DA.Fill(ZOVOrdersNames);
                }

                int RowIndex = 0;
                HSSFSheet sheet1 = hssfworkbook.CreateSheet("Арки ЗОВ");
                sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

                sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
                sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
                sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
                sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

                sheet1.SetColumnWidth(0, 30 * 256);
                sheet1.SetColumnWidth(1, 20 * 256);
                sheet1.SetColumnWidth(2, 9 * 256);
                sheet1.SetColumnWidth(3, 6 * 256);
                sheet1.SetColumnWidth(4, 6 * 256);

                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=0"))
                {
                    DecorAssemblyDT.Clear();
                    DT.Clear();
                    int MainOrderID = Convert.ToInt32(item["MainOrderID"]);
                    DataRow[] rows = ArchDecorOrdersDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item1 in rows)
                        DT.Rows.Add(item1.ItemArray);
                    AssemblyDecorCollect(DT, ref DecorAssemblyDT);

                    DataRow[] CRows = ZOVOrdersNames.Select("MainOrderID=" + Convert.ToInt32(item["MainOrderID"]));
                    if (CRows.Count() > 0)
                    {
                        ClientName = CRows[0]["ClientName"].ToString();
                        OrderName = CRows[0]["DocNumber"].ToString();
                    }
                    if (DecorAssemblyDT.Rows.Count > 0)
                    {
                        ArchDecorAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                             Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DecorAssemblyDT,
                             WorkAssignmentID, BatchName, ClientName, string.Empty, OrderName, Notes, ref RowIndex);
                        RowIndex++;
                        ArchDecorAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                            Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DecorAssemblyDT,
                            WorkAssignmentID, BatchName, ClientName, "ДУБЛЬ", OrderName, Notes, ref RowIndex);
                    }
                    RowIndex++;
                }
            }
            if (DistMainOrdersDT.Select("GroupType=1").Count() > 0)
            {
                MainOrdersID = string.Empty;
                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=1"))
                    MainOrdersID += Convert.ToInt32(item["MainOrderID"]) + ",";
                MainOrdersID = MainOrdersID.Substring(0, MainOrdersID.Length - 1);
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT ClientName, OrderNumber, MainOrderID, Notes FROM MainOrders" +
                    " INNER JOIN MegaOrders ON MainOrders.MegaOrderID=MegaOrders.MegaOrderID" +
                    " INNER JOIN infiniu2_marketingreference.dbo.Clients ON MegaOrders.ClientID=infiniu2_marketingreference.dbo.Clients.ClientID" +
                    " WHERE MainOrderID IN (" + MainOrdersID + ")", ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DA.Fill(MarketOrdersNames);
                }

                int RowIndex = 0;
                HSSFSheet sheet1 = hssfworkbook.CreateSheet("Арки Маркетинг");
                sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

                sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
                sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
                sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
                sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

                sheet1.SetColumnWidth(0, 30 * 256);
                sheet1.SetColumnWidth(1, 20 * 256);
                sheet1.SetColumnWidth(2, 9 * 256);
                sheet1.SetColumnWidth(3, 6 * 256);
                sheet1.SetColumnWidth(4, 6 * 256);

                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=1"))
                {
                    DecorAssemblyDT.Clear();
                    DT.Clear();
                    int MainOrderID = Convert.ToInt32(item["MainOrderID"]);
                    DataRow[] rows = ArchDecorOrdersDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item1 in rows)
                        DT.Rows.Add(item1.ItemArray);
                    AssemblyDecorCollect(DT, ref DecorAssemblyDT);

                    DataRow[] CRows = MarketOrdersNames.Select("MainOrderID=" + MainOrderID);
                    if (CRows.Count() > 0)
                    {
                        ClientName = CRows[0]["ClientName"].ToString();
                        Notes = CRows[0]["Notes"].ToString();
                        OrderName = "№" + CRows[0]["OrderNumber"].ToString() + "-" + MainOrderID;
                    }
                    if (DecorAssemblyDT.Rows.Count > 0)
                    {
                        ArchDecorAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                             Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DecorAssemblyDT,
                             WorkAssignmentID, BatchName, ClientName, string.Empty, OrderName, Notes, ref RowIndex);
                        RowIndex++;
                        ArchDecorAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                            Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DecorAssemblyDT,
                            WorkAssignmentID, BatchName, ClientName, "ДУБЛЬ", OrderName, Notes, ref RowIndex);
                    }
                    RowIndex++;
                }
            }
        }

        public void ArchDecorAssemblyToExcelSingly(ref HSSFWorkbook hssfworkbook, ref HSSFSheet sheet1,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            DataTable DT, int WorkAssignmentID, string BatchName, string ClientName, string OperationName, string OrderName, string Notes, ref int RowIndex)
        {
            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, ClientName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Заказ:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, OrderName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 4, OperationName);
            cell.CellStyle = CalibriBold11CS;
            if (Notes.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Примечание: " + Notes);
                cell.CellStyle = CalibriBold11CS;
            }

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Наименование");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Цвет");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Длин./Выс.");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 5, "Прим.");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            int DecorID = -1;
            int AllTotalAmount = 0;
            int TotalAmount = 0;
            int DifferentDecorCount = 0;

            using (DataView DV = new DataView(DT))
            {
                DifferentDecorCount = DV.ToTable(true, new string[] { "DecorID" }).Rows.Count;
            }

            if (DT.Rows.Count > 0)
            {
                DecorID = Convert.ToInt32(DT.Rows[0]["DecorID"]);
            }

            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value)
                {
                    AllTotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    TotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                }
                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "DecorID")
                        continue;

                    Type t = DT.Rows[x][y].GetType();

                    if (t.Name == "Decimal")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToDouble(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderDecCS;
                        continue;
                    }
                    if (t.Name == "Int32")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToInt32(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (t.Name == "String" || t.Name == "DBNull")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(DT.Rows[x][y].ToString());
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }
                }

                if (x + 1 <= DT.Rows.Count - 1)
                {
                    if (DecorID != Convert.ToInt32(DT.Rows[x + 1]["DecorID"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "DecorID")
                                continue;

                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        DecorID = Convert.ToInt32(DT.Rows[x + 1]["DecorID"]);
                        TotalAmount = 0;
                        RowIndex++;
                    }
                }

                if (x == DT.Rows.Count - 1)
                {
                    if (DifferentDecorCount > 1)
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "DecorID")
                                continue;

                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;
                    }
                    RowIndex++;

                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "DecorID")
                            continue;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;
                }
                RowIndex++;
            }
        }

        private void GetPlanningTime1(DataTable dt, decimal div1, decimal div2, ref decimal time, ref decimal cost)
        {
            for (int x = 0; x < dt.Rows.Count; x++)
            {
                if (dt.Rows[x]["Count"] != DBNull.Value)
                    time += Convert.ToInt32(dt.Rows[x]["Count"]);
            }

            if (div1 != 0)
                time = decimal.Round(time / div1, 3, MidpointRounding.AwayFromZero);

            cost = decimal.Round(time * div2, 2, MidpointRounding.AwayFromZero);
        }

        private void StrapsAssemblyByMainOrderToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            int WorkAssignmentID, string BatchName)
        {
            if (StrapsDecorOrdersDT.Rows.Count == 0)
                return;

            decimal div1 = 0;
            decimal div2 = 0;
            decimal time = 0;
            decimal cost = 0;

            DataTable DistMainOrdersDT = DistMainOrdersTable(StrapsDecorOrdersDT, true);
            DataTable DT = StrapsDecorOrdersDT.Clone();
            DataTable ZOVOrdersNames = new DataTable();
            DataTable MarketOrdersNames = new DataTable();
            string MainOrdersID = string.Empty;
            string ClientName = string.Empty;
            string OrderName = string.Empty;
            string Notes = string.Empty;

            if (DistMainOrdersDT.Select("GroupType=1").Count() > 0)
            {
                MainOrdersID = string.Empty;
                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=1"))
                    MainOrdersID += Convert.ToInt32(item["MainOrderID"]) + ",";
                MainOrdersID = MainOrdersID.Substring(0, MainOrdersID.Length - 1);
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT ClientName, OrderNumber, MainOrderID, Notes FROM MainOrders" +
                    " INNER JOIN MegaOrders ON MainOrders.MegaOrderID=MegaOrders.MegaOrderID" +
                    " INNER JOIN infiniu2_marketingreference.dbo.Clients ON MegaOrders.ClientID=infiniu2_marketingreference.dbo.Clients.ClientID" +
                    " WHERE MainOrderID IN (" + MainOrdersID + ")", ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DA.Fill(MarketOrdersNames);
                }

                int RowIndex = 0;
                HSSFSheet sheet1 = hssfworkbook.CreateSheet("Накладки Маркетинг");
                sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

                sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
                sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
                sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
                sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

                sheet1.SetColumnWidth(0, 30 * 256);
                sheet1.SetColumnWidth(1, 20 * 256);
                sheet1.SetColumnWidth(2, 9 * 256);
                sheet1.SetColumnWidth(3, 6 * 256);
                sheet1.SetColumnWidth(4, 6 * 256);

                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=1"))
                {
                    DecorAssemblyDT.Clear();
                    DT.Clear();
                    int MainOrderID = Convert.ToInt32(item["MainOrderID"]);
                    DataRow[] rows = StrapsDecorOrdersDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item1 in rows)
                        DT.Rows.Add(item1.ItemArray);
                    AssemblyDecorCollect(DT, ref DecorAssemblyDT);

                    DataRow[] CRows = MarketOrdersNames.Select("MainOrderID=" + MainOrderID);
                    if (CRows.Count() > 0)
                    {
                        ClientName = CRows[0]["ClientName"].ToString();
                        Notes = CRows[0]["Notes"].ToString();
                        OrderName = "№" + CRows[0]["OrderNumber"].ToString() + "-" + MainOrderID;
                    }


                    if (DecorAssemblyDT.Rows.Count > 0)
                    {
                        div1 = 6;
                        div2 = 2.15m;
                        time = 0;
                        cost = 0;
                        GetPlanningTime1(DecorAssemblyDT, div1, div2, ref time, ref cost);

                        StrapsAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                             Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DecorAssemblyDT,
                             WorkAssignmentID, BatchName, ClientName, "Изготовление НУ", string.Empty, OrderName, Notes, time, cost, ref RowIndex);
                        RowIndex++;
                        StrapsAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                            Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DecorAssemblyDT,
                            WorkAssignmentID, BatchName, ClientName, "Изготовление НУ", "ДУБЛЬ", OrderName, Notes, time, cost, ref RowIndex);
                        RowIndex++;

                        div1 = 15;
                        div2 = 1.25m;
                        time = 0;
                        cost = 0;
                        GetPlanningTime1(DecorAssemblyDT, div1, div2, ref time, ref cost);

                        StrapsAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                             Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DecorAssemblyDT,
                             WorkAssignmentID, BatchName, ClientName, "Покраска Торцов НУ(2 слоя)", string.Empty, OrderName, Notes, time, cost, ref RowIndex);
                        RowIndex++;
                        StrapsAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                            Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DecorAssemblyDT,
                            WorkAssignmentID, BatchName, ClientName, "Покраска Торцов НУ(2 слоя)", "ДУБЛЬ", OrderName, Notes, time, cost, ref RowIndex);
                        RowIndex++;

                        div1 = 45;
                        div2 = 1.25m;
                        time = 0;
                        cost = 0;
                        GetPlanningTime1(DecorAssemblyDT, div1, div2, ref time, ref cost);

                        StrapsAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                             Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DecorAssemblyDT,
                             WorkAssignmentID, BatchName, ClientName, "Зачистка Декора", string.Empty, OrderName, Notes, time, cost, ref RowIndex);
                        RowIndex++;
                        StrapsAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                            Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DecorAssemblyDT,
                            WorkAssignmentID, BatchName, ClientName, "Зачистка Декора", "ДУБЛЬ", OrderName, Notes, time, cost, ref RowIndex);
                    }
                    RowIndex++;
                }
            }
        }

        public void StrapsAssemblyToExcelSingly(ref HSSFWorkbook hssfworkbook, ref HSSFSheet sheet1,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            DataTable DT, int WorkAssignmentID, string BatchName, string ClientName, string OperationName, string DuplicateName, string OrderName, string Notes, decimal Time, decimal Cost, ref int RowIndex)
        {
            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
            cell.CellStyle = Calibri11CS;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Плановое время выполнения:");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "ч/ч");
            cell.CellStyle = Calibri11CS;

            cell = sheet1.CreateRow(RowIndex++).CreateCell(2);
            cell.SetCellValue(Convert.ToDouble(Time));
            cell.CellStyle = Calibri11CS;


            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Планово-премиальный фонд:");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "руб.");
            cell.CellStyle = Calibri11CS;

            cell = sheet1.CreateRow(RowIndex++).CreateCell(2);
            cell.SetCellValue(Convert.ToDouble(Cost));
            cell.CellStyle = Calibri11CS;

            RowIndex++;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, OperationName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, ClientName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Заказ:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, OrderName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 4, DuplicateName);
            cell.CellStyle = CalibriBold11CS;
            if (Notes.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Примечание: " + Notes);
                cell.CellStyle = CalibriBold11CS;
            }

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Наименование");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Цвет");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Длин./Выс.");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 5, "Прим.");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            int DecorID = -1;
            int AllTotalAmount = 0;
            int TotalAmount = 0;
            int DifferentDecorCount = 0;

            using (DataView DV = new DataView(DT))
            {
                DifferentDecorCount = DV.ToTable(true, new string[] { "DecorID" }).Rows.Count;
            }

            if (DT.Rows.Count > 0)
            {
                DecorID = Convert.ToInt32(DT.Rows[0]["DecorID"]);
            }

            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value)
                {
                    AllTotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    TotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                }
                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "DecorID")
                        continue;

                    Type t = DT.Rows[x][y].GetType();

                    if (t.Name == "Decimal")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToDouble(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderDecCS;
                        continue;
                    }
                    if (t.Name == "Int32")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToInt32(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (t.Name == "String" || t.Name == "DBNull")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(DT.Rows[x][y].ToString());
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }
                }

                if (x + 1 <= DT.Rows.Count - 1)
                {
                    if (DecorID != Convert.ToInt32(DT.Rows[x + 1]["DecorID"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "DecorID")
                                continue;

                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        DecorID = Convert.ToInt32(DT.Rows[x + 1]["DecorID"]);
                        TotalAmount = 0;
                        RowIndex++;
                    }
                }

                if (x == DT.Rows.Count - 1)
                {
                    if (DifferentDecorCount > 1)
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "DecorID")
                                continue;

                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;
                    }
                    RowIndex++;

                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "DecorID")
                            continue;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;
                }
                RowIndex++;
            }
        }

        private void GridsDecorAssemblyByMainOrderToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            int WorkAssignmentID, string BatchName)
        {
            if (GridsDecorOrdersDT.Rows.Count == 0)
                return;

            DataTable DistMainOrdersDT = DistMainOrdersTable(GridsDecorOrdersDT, true);
            DataTable DT = GridsDecorOrdersDT.Clone();
            DataTable ZOVOrdersNames = new DataTable();
            DataTable MarketOrdersNames = new DataTable();
            string MainOrdersID = string.Empty;
            string ClientName = string.Empty;
            string OrderName = string.Empty;
            string Notes = string.Empty;

            if (DistMainOrdersDT.Select("GroupType=0").Count() > 0)
            {
                MainOrdersID = string.Empty;
                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=0"))
                    MainOrdersID += Convert.ToInt32(item["MainOrderID"]) + ",";
                MainOrdersID = MainOrdersID.Substring(0, MainOrdersID.Length - 1);
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT ClientName, DocNumber, MainOrderID FROM MainOrders" +
                    " INNER JOIN infiniu2_zovreference.dbo.Clients ON MainOrders.ClientID=infiniu2_zovreference.dbo.Clients.ClientID" +
                    " WHERE MainOrderID IN (" + MainOrdersID + ")", ConnectionStrings.ZOVOrdersConnectionString))
                {
                    DA.Fill(ZOVOrdersNames);
                }

                int RowIndex = 0;
                HSSFSheet sheet1 = hssfworkbook.CreateSheet("Решетки1 ЗОВ");
                sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

                sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
                sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
                sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
                sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

                sheet1.SetColumnWidth(0, 30 * 256);
                sheet1.SetColumnWidth(1, 20 * 256);
                sheet1.SetColumnWidth(2, 9 * 256);
                sheet1.SetColumnWidth(3, 6 * 256);
                sheet1.SetColumnWidth(4, 6 * 256);

                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=0"))
                {
                    DecorAssemblyDT.Clear();
                    DT.Clear();
                    int MainOrderID = Convert.ToInt32(item["MainOrderID"]);
                    DataRow[] rows = GridsDecorOrdersDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item1 in rows)
                        DT.Rows.Add(item1.ItemArray);
                    AssemblyDecorCollect(DT, ref DecorAssemblyDT);

                    DataRow[] CRows = ZOVOrdersNames.Select("MainOrderID=" + Convert.ToInt32(item["MainOrderID"]));
                    if (CRows.Count() > 0)
                    {
                        ClientName = CRows[0]["ClientName"].ToString();
                        OrderName = CRows[0]["DocNumber"].ToString();
                    }
                    if (DecorAssemblyDT.Rows.Count > 0)
                    {
                        GridsDecorAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                             Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DecorAssemblyDT,
                             WorkAssignmentID, BatchName, ClientName, string.Empty, OrderName, Notes, ref RowIndex);
                        RowIndex++;
                        GridsDecorAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                            Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DecorAssemblyDT,
                            WorkAssignmentID, BatchName, ClientName, "ДУБЛЬ", OrderName, Notes, ref RowIndex);
                    }
                    RowIndex++;
                }
            }
            if (DistMainOrdersDT.Select("GroupType=1").Count() > 0)
            {
                MainOrdersID = string.Empty;
                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=1"))
                    MainOrdersID += Convert.ToInt32(item["MainOrderID"]) + ",";
                MainOrdersID = MainOrdersID.Substring(0, MainOrdersID.Length - 1);
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT ClientName, OrderNumber, MainOrderID, Notes FROM MainOrders" +
                    " INNER JOIN MegaOrders ON MainOrders.MegaOrderID=MegaOrders.MegaOrderID" +
                    " INNER JOIN infiniu2_marketingreference.dbo.Clients ON MegaOrders.ClientID=infiniu2_marketingreference.dbo.Clients.ClientID" +
                    " WHERE MainOrderID IN (" + MainOrdersID + ")", ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DA.Fill(MarketOrdersNames);
                }

                int RowIndex = 0;
                HSSFSheet sheet1 = hssfworkbook.CreateSheet("Решетки1 Маркетинг");
                sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

                sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
                sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
                sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
                sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

                sheet1.SetColumnWidth(0, 30 * 256);
                sheet1.SetColumnWidth(1, 20 * 256);
                sheet1.SetColumnWidth(2, 9 * 256);
                sheet1.SetColumnWidth(3, 6 * 256);
                sheet1.SetColumnWidth(4, 6 * 256);

                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=1"))
                {
                    DecorAssemblyDT.Clear();
                    DT.Clear();
                    int MainOrderID = Convert.ToInt32(item["MainOrderID"]);
                    DataRow[] rows = GridsDecorOrdersDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item1 in rows)
                        DT.Rows.Add(item1.ItemArray);
                    AssemblyDecorCollect(DT, ref DecorAssemblyDT);

                    DataRow[] CRows = MarketOrdersNames.Select("MainOrderID=" + MainOrderID);
                    if (CRows.Count() > 0)
                    {
                        ClientName = CRows[0]["ClientName"].ToString();
                        Notes = CRows[0]["Notes"].ToString();
                        OrderName = "№" + CRows[0]["OrderNumber"].ToString() + "-" + MainOrderID;
                    }
                    if (DecorAssemblyDT.Rows.Count > 0)
                    {
                        GridsDecorAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                             Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DecorAssemblyDT,
                             WorkAssignmentID, BatchName, ClientName, string.Empty, OrderName, Notes, ref RowIndex);
                        RowIndex++;
                        GridsDecorAssemblyToExcelSingly(ref hssfworkbook, ref sheet1,
                            Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DecorAssemblyDT,
                            WorkAssignmentID, BatchName, ClientName, "ДУБЛЬ", OrderName, Notes, ref RowIndex);
                    }
                    RowIndex++;
                }
            }
        }

        public void GridsDecorAssemblyToExcelSingly(ref HSSFWorkbook hssfworkbook, ref HSSFSheet sheet1,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            DataTable DT, int WorkAssignmentID, string BatchName, string ClientName, string OperationName, string OrderName, string Notes, ref int RowIndex)
        {
            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, ClientName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Заказ:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, OrderName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 4, OperationName);
            cell.CellStyle = CalibriBold11CS;
            if (Notes.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Примечание: " + Notes);
                cell.CellStyle = CalibriBold11CS;
            }

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Наименование");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Цвет");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Длин./Выс.");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 5, "Прим.");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            int DecorID = -1;
            int AllTotalAmount = 0;
            int TotalAmount = 0;
            int DifferentDecorCount = 0;

            using (DataView DV = new DataView(DT))
            {
                DifferentDecorCount = DV.ToTable(true, new string[] { "DecorID" }).Rows.Count;
            }

            if (DT.Rows.Count > 0)
            {
                DecorID = Convert.ToInt32(DT.Rows[0]["DecorID"]);
            }

            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value)
                {
                    AllTotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    TotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                }
                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "DecorID")
                        continue;

                    Type t = DT.Rows[x][y].GetType();

                    if (t.Name == "Decimal")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToDouble(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderDecCS;
                        continue;
                    }
                    if (t.Name == "Int32")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToInt32(DT.Rows[x][y]));
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (t.Name == "String" || t.Name == "DBNull")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(DT.Rows[x][y].ToString());
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }
                }

                if (x + 1 <= DT.Rows.Count - 1)
                {
                    if (DecorID != Convert.ToInt32(DT.Rows[x + 1]["DecorID"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "DecorID")
                                continue;

                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        DecorID = Convert.ToInt32(DT.Rows[x + 1]["DecorID"]);
                        TotalAmount = 0;
                        RowIndex++;
                    }
                }

                if (x == DT.Rows.Count - 1)
                {
                    if (DifferentDecorCount > 1)
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "DecorID")
                                continue;

                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;
                    }
                    RowIndex++;

                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "DecorID")
                            continue;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;
                }
                RowIndex++;
            }
        }

    }

}
