
using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace Infinium.Modules.WorkAssignments
{
    public class TafelSplitStruct
    {
        public int PosCount;
        public int SourceCount;
        public int FirstCount;
        public int SecondCount;
        public bool bOk;
        public bool bEqual;

        public TafelSplitStruct()
        {
            PosCount = 0;
            FirstCount = 0;
            SecondCount = 0;
            bOk = false;
            bEqual = false;
        }
    }

    public class TafelManager
    {
        private DataTable MainOrdersDT;
        private DataTable FrontOrdersDT;
        private BindingSource MainOrdersBS;
        private BindingSource FrontOrdersBS;
        public FrontsOrdersManager FrontsOrdersManager;
        public TafelManager(FrontsOrdersManager tFrontsOrdersManager)
        {
            FrontsOrdersManager = tFrontsOrdersManager;
            MainOrdersDT = new DataTable();
            FrontOrdersDT = new DataTable();
            MainOrdersDT.Columns.Add(new DataColumn("ClientName", Type.GetType("System.String")));
            MainOrdersDT.Columns.Add(new DataColumn("OrderNumber", Type.GetType("System.String")));
            MainOrdersDT.Columns.Add(new DataColumn("MainOrderID", Type.GetType("System.Int32")));
            MainOrdersDT.Columns.Add(new DataColumn("GroupType", Type.GetType("System.Int32")));
            MainOrdersBS = new BindingSource();
            FrontOrdersBS = new BindingSource();

            MainOrdersBS.DataSource = MainOrdersDT;
        }

        public DataTable EditFrontOrdersDT => FrontOrdersDT;

        public BindingSource MainOrdersList => MainOrdersBS;

        public BindingSource FrontOrdersList => FrontOrdersBS;

        public void FillTables()
        {
            FrontOrdersDT.Clear();
            FrontOrdersDT = FrontsOrdersManager.OrdersDT.Copy();
            FrontOrdersDT.Columns.Add(new DataColumn("GroupNumber", Type.GetType("System.Int32")));
            FrontOrdersBS.DataSource = FrontOrdersDT;

            DataTable DistMainOrdersDT = DistMainOrdersTable(FrontOrdersDT, true);
            DataTable DT = new DataTable();
            string MainOrdersID = string.Empty;

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
                    DA.Fill(DT);
                    foreach (DataRow item in DT.Rows)
                    {
                        DataRow NewRow = MainOrdersDT.NewRow();
                        NewRow["ClientName"] = item["ClientName"];
                        NewRow["OrderNumber"] = item["DocNumber"];
                        NewRow["MainOrderID"] = item["MainOrderID"];
                        NewRow["GroupType"] = 0;
                        MainOrdersDT.Rows.Add(NewRow);
                    }
                }
            }
            if (DistMainOrdersDT.Select("GroupType=1").Count() > 0)
            {
                MainOrdersID = string.Empty;
                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=1"))
                    MainOrdersID += Convert.ToInt32(item["MainOrderID"]) + ",";
                MainOrdersID = MainOrdersID.Substring(0, MainOrdersID.Length - 1);
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT ClientName, OrderNumber, MainOrderID FROM MainOrders" +
                    " INNER JOIN MegaOrders ON MainOrders.MegaOrderID=MegaOrders.MegaOrderID" +
                    " INNER JOIN infiniu2_marketingreference.dbo.Clients ON MegaOrders.ClientID=infiniu2_marketingreference.dbo.Clients.ClientID" +
                    " WHERE MainOrderID IN (" + MainOrdersID + ")", ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DT.Clear();
                    DA.Fill(DT);
                    foreach (DataRow item in DT.Rows)
                    {
                        DataRow NewRow = MainOrdersDT.NewRow();
                        NewRow["ClientName"] = item["ClientName"];
                        NewRow["OrderNumber"] = item["OrderNumber"];
                        NewRow["MainOrderID"] = item["MainOrderID"];
                        NewRow["GroupType"] = 1;
                        MainOrdersDT.Rows.Add(NewRow);
                    }
                }
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

        public void FilterOrdersByMainOrder(int GroupType, int MainOrderID)
        {
            FrontOrdersBS.Filter = "GroupType=" + GroupType + " AND MainOrderID=" + MainOrderID;
        }

        public void MoveToPosition(int Position)
        {
            FrontOrdersBS.Position = Position;
        }

        public void EditCurrentFrontsCout(int NewCount, ref decimal Square)
        {
            int OldCount = Convert.ToInt32(((DataRowView)FrontOrdersBS.Current).Row["Count"]);
            Square = Convert.ToDecimal(((DataRowView)FrontOrdersBS.Current).Row["Square"]) / OldCount;
            ((DataRowView)FrontOrdersBS.Current).Row["Count"] = NewCount;
            ((DataRowView)FrontOrdersBS.Current).Row["Square"] = Square * NewCount;
        }

        public void AddNewRow(int Count, decimal Square)
        {
            DataRow NewRow = FrontOrdersDT.NewRow();
            NewRow.ItemArray = ((DataRowView)FrontOrdersBS.Current).Row.ItemArray;
            NewRow["Count"] = Count;
            NewRow["Square"] = Square * Count;
            FrontOrdersDT.Rows.Add(NewRow);
        }

        public int CurrentFrontsCount
        {
            get
            {
                if (FrontOrdersBS.Count == 0 || ((DataRowView)FrontOrdersBS.Current).Row["Count"] == DBNull.Value)
                    return 0;
                return Convert.ToInt32(((DataRowView)FrontOrdersBS.Current).Row["Count"]);
            }
        }
    }


    public class FrontsOrdersManager
    {
        private DataTable AllBatchFrontsDT;
        private DataTable AllPrintedFrontsDT;
        private DataTable BatchFrontsDT;
        private DataTable FrontsSummaryDT;
        private DataTable FrameColorsSummaryDT;
        private DataTable InsetTypesSummaryDT;
        private DataTable InsetColorsSummaryDT;
        private DataTable SizesSummaryDT;

        private DataTable FrontsDataTable = null;
        private DataTable FrameColorsDataTable = null;
        private DataTable PatinaDataTable = null;
        private DataTable PatinaRALDataTable = null;
        private DataTable InsetTypesDataTable = null;
        private DataTable InsetColorsDataTable = null;
        private DataTable TechnoProfilesDataTable = null;
        private DataTable TechStoreDataTable = null;
        private BindingSource BatchFrontsBS = null;
        private BindingSource FrontsSummaryBS = null;
        private BindingSource FrameColorsSummaryBS = null;
        private BindingSource InsetTypesSummaryBS = null;
        private BindingSource InsetColorsSummaryBS = null;
        private BindingSource SizesSummaryBS = null;

        public BindingSource BatchFrontsList => BatchFrontsBS;

        public BindingSource FrontsSummaryList => FrontsSummaryBS;

        public BindingSource FrameColorsSummaryList => FrameColorsSummaryBS;

        public BindingSource InsetTypesSummaryList => InsetTypesSummaryBS;

        public BindingSource InsetColorsSummaryList => InsetColorsSummaryBS;

        public BindingSource SizesSummaryList => SizesSummaryBS;

        public FrontsOrdersManager()
        {

        }

        private void Create()
        {
            FrontsDataTable = new DataTable();
            FrameColorsDataTable = new DataTable();
            PatinaDataTable = new DataTable();
            InsetTypesDataTable = new DataTable();
            InsetColorsDataTable = new DataTable();

            FrontsSummaryDT = new DataTable();
            FrontsSummaryDT.Columns.Add(new DataColumn(("Front"), System.Type.GetType("System.String")));
            FrontsSummaryDT.Columns.Add(new DataColumn(("FrontID"), System.Type.GetType("System.Int32")));
            FrontsSummaryDT.Columns.Add(new DataColumn(("FrontTypeID"), System.Type.GetType("System.Int32")));
            FrontsSummaryDT.Columns.Add(new DataColumn(("Square"), System.Type.GetType("System.Decimal")));
            FrontsSummaryDT.Columns.Add(new DataColumn(("Count"), System.Type.GetType("System.Int32")));
            FrontsSummaryDT.Columns.Add(new DataColumn(("PrintingStatus"), System.Type.GetType("System.Int32")));

            FrameColorsSummaryDT = new DataTable();
            FrameColorsSummaryDT.Columns.Add(new DataColumn(("FrameColor"), System.Type.GetType("System.String")));
            FrameColorsSummaryDT.Columns.Add(new DataColumn(("FrontID"), System.Type.GetType("System.Int32")));
            FrameColorsSummaryDT.Columns.Add(new DataColumn(("FrontTypeID"), System.Type.GetType("System.Int32")));
            FrameColorsSummaryDT.Columns.Add(new DataColumn(("ColorID"), System.Type.GetType("System.Int32")));
            FrameColorsSummaryDT.Columns.Add(new DataColumn(("PatinaID"), System.Type.GetType("System.Int32")));
            FrameColorsSummaryDT.Columns.Add(new DataColumn(("Square"), System.Type.GetType("System.Decimal")));
            FrameColorsSummaryDT.Columns.Add(new DataColumn(("Count"), System.Type.GetType("System.Int32")));

            SizesSummaryDT = new DataTable();
            SizesSummaryDT.Columns.Add(new DataColumn(("Size"), System.Type.GetType("System.String")));
            SizesSummaryDT.Columns.Add(new DataColumn(("FrontID"), System.Type.GetType("System.Int32")));
            SizesSummaryDT.Columns.Add(new DataColumn(("FrontTypeID"), System.Type.GetType("System.Int32")));
            SizesSummaryDT.Columns.Add(new DataColumn(("ColorID"), System.Type.GetType("System.Int32")));
            SizesSummaryDT.Columns.Add(new DataColumn(("PatinaID"), System.Type.GetType("System.Int32")));
            SizesSummaryDT.Columns.Add(new DataColumn(("InsetTypeID"), System.Type.GetType("System.Int32")));
            SizesSummaryDT.Columns.Add(new DataColumn(("InsetColorID"), System.Type.GetType("System.Int32")));
            SizesSummaryDT.Columns.Add(new DataColumn(("Square"), System.Type.GetType("System.Decimal")));
            SizesSummaryDT.Columns.Add(new DataColumn(("Width"), System.Type.GetType("System.Int32")));
            SizesSummaryDT.Columns.Add(new DataColumn(("Count"), System.Type.GetType("System.Int32")));
            SizesSummaryDT.Columns.Add(new DataColumn(("Height"), System.Type.GetType("System.Int32")));

            InsetTypesSummaryDT = new DataTable();
            InsetTypesSummaryDT.Columns.Add(new DataColumn(("InsetType"), System.Type.GetType("System.String")));
            InsetTypesSummaryDT.Columns.Add(new DataColumn(("FrontID"), System.Type.GetType("System.Int32")));
            InsetTypesSummaryDT.Columns.Add(new DataColumn(("FrontTypeID"), System.Type.GetType("System.Int32")));
            InsetTypesSummaryDT.Columns.Add(new DataColumn(("ColorID"), System.Type.GetType("System.Int32")));
            InsetTypesSummaryDT.Columns.Add(new DataColumn(("PatinaID"), System.Type.GetType("System.Int32")));
            InsetTypesSummaryDT.Columns.Add(new DataColumn(("InsetTypeID"), System.Type.GetType("System.Int32")));
            InsetTypesSummaryDT.Columns.Add(new DataColumn(("Square"), System.Type.GetType("System.Decimal")));
            InsetTypesSummaryDT.Columns.Add(new DataColumn(("Count"), System.Type.GetType("System.Int32")));

            InsetColorsSummaryDT = new DataTable();
            InsetColorsSummaryDT.Columns.Add(new DataColumn(("InsetColor"), System.Type.GetType("System.String")));
            InsetColorsSummaryDT.Columns.Add(new DataColumn(("FrontID"), System.Type.GetType("System.Int32")));
            InsetColorsSummaryDT.Columns.Add(new DataColumn(("FrontTypeID"), System.Type.GetType("System.Int32")));
            InsetColorsSummaryDT.Columns.Add(new DataColumn(("ColorID"), System.Type.GetType("System.Int32")));
            InsetColorsSummaryDT.Columns.Add(new DataColumn(("PatinaID"), System.Type.GetType("System.Int32")));
            InsetColorsSummaryDT.Columns.Add(new DataColumn(("InsetColorID"), System.Type.GetType("System.Int32")));
            InsetColorsSummaryDT.Columns.Add(new DataColumn(("InsetTypeID"), System.Type.GetType("System.Int32")));
            InsetColorsSummaryDT.Columns.Add(new DataColumn(("Square"), System.Type.GetType("System.Decimal")));
            InsetColorsSummaryDT.Columns.Add(new DataColumn(("Count"), System.Type.GetType("System.Int32")));

            AllBatchFrontsDT = new DataTable();
            AllPrintedFrontsDT = new DataTable();
            BatchFrontsDT = new DataTable();
            BatchFrontsBS = new BindingSource();
            FrontsSummaryBS = new BindingSource();
            FrameColorsSummaryBS = new BindingSource();
            InsetTypesSummaryBS = new BindingSource();
            InsetColorsSummaryBS = new BindingSource();
            SizesSummaryBS = new BindingSource();
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

        private void Fill()
        {
            string SelectCommand = @"SELECT TechStoreID AS FrontID, TechStoreName AS FrontName FROM TechStore 
                WHERE TechStoreID IN (SELECT FrontID FROM FrontsConfig WHERE Enabled = 1)
                ORDER BY TechStoreName";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(FrontsDataTable);
            }
            SelectCommand = @"SELECT DISTINCT TechStoreID AS TechnoProfileID, TechStoreName AS TechnoProfileName FROM TechStore 
                WHERE TechStoreID IN (SELECT TechnoProfileID FROM FrontsConfig WHERE Enabled = 1 AND AccountingName IS NOT NULL AND InvNumber IS NOT NULL)
                ORDER BY TechStoreName";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.CatalogConnectionString))
            {
                TechnoProfilesDataTable = new DataTable();
                DA.Fill(TechnoProfilesDataTable);

                DataRow NewRow = TechnoProfilesDataTable.NewRow();
                NewRow["TechnoProfileID"] = -1;
                NewRow["TechnoProfileName"] = "-";
                TechnoProfilesDataTable.Rows.InsertAt(NewRow, 0);
            }

            GetColorsDT();
            GetInsetColorsDT();
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM Patina",
                ConnectionStrings.CatalogConnectionString))
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
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM InsetTypes",
                ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(InsetTypesDataTable);
            }

            TechStoreDataTable = new DataTable();
            //using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM TechStore",
            //    ConnectionStrings.CatalogConnectionString))
            //{
            //    DA.Fill(TechStoreDataTable);
            //}
            TechStoreDataTable = TablesManager.TechStoreDataTable;
            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT TOP 0 FrontsOrdersID, FrontsOrders.MainOrderID, FrontID, ColorID, PatinaID, 
InsetTypeID, InsetColorID, TechnoProfileID, TechnoColorID, TechnoInsetTypeID, TechnoInsetColorID, Height, Width, Count, FrontConfigID, FrontsOrders.Square, FrontsOrders.FactoryID, IsNonStandard, FrontsOrders.Notes, ClientName FROM FrontsOrders INNER JOIN
                MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID INNER JOIN
                MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID INNER JOIN
                infiniu2_marketingreference.dbo.Clients AS Client ON MegaOrders.ClientID = Client.ClientID",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(AllBatchFrontsDT);
                DA.Fill(BatchFrontsDT);
                AllBatchFrontsDT.Columns.Add(new DataColumn(("GroupType"), System.Type.GetType("System.Int32")));
            }
            for (int i = 0; i < FrontsSummaryDT.Rows.Count; i++)
                FrontsSummaryDT.Rows[i]["PrintingStatus"] = 0;
            for (int i = 0; i < FrameColorsSummaryDT.Rows.Count; i++)
                FrameColorsSummaryDT.Rows[i]["PrintingStatus"] = 0;
        }

        private void Binding()
        {
            BatchFrontsBS.DataSource = BatchFrontsDT;
            FrontsSummaryBS.DataSource = FrontsSummaryDT;
            FrameColorsSummaryBS.DataSource = FrameColorsSummaryDT;
            InsetTypesSummaryBS.DataSource = InsetTypesSummaryDT;
            InsetColorsSummaryBS.DataSource = InsetColorsSummaryDT;
            SizesSummaryBS.DataSource = SizesSummaryDT;
        }

        public void Initialize()
        {
            Create();
            Fill();
            Binding();
        }

        public bool FilterFrontsByWorkAssignment(int WorkAssignmentID, int FactoryID, int FilterType)
        {
            string FilterString = string.Empty;

            string BatchFactoryFilter = string.Empty;
            string FactoryFilter = string.Empty;
            string SelectCommand = string.Empty;

            if (FactoryID != 0)
            {
                BatchFactoryFilter = " AND BatchDetails.FactoryID = " + FactoryID;
                FactoryFilter = " AND FrontsOrders.FactoryID = " + FactoryID;
            }
            DataTable DT = AllBatchFrontsDT.Clone();
            AllBatchFrontsDT.Clear();

            if (FilterType == 1)
                FilterString = " AND FrontsOrders.FrontID IN (1975,1976,1977,1978,15760, 3737, 30501,30502,30503,30504,30505,30506,16269,30364,30366,30367,28945,3727,3728,3729,3730,3731,3732,3733,3734,3735,3736,3737,3739,3740,3741,3742,3743,3744,3745,3746,3747,3748,15108,27914,29597)";
            if (FilterType == 2)
                FilterString = " AND (FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.KansasPat) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Kansas) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Sofia) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Lorenzo) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Elegant) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.ElegantPat) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Patricia1) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Turin1) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Turin1_1) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Dakota) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.DakotaPat) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Turin3) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.LeonTPS) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Infiniti) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.InfinitiPat) + ")";
            if (FilterType == 3)
                FilterString = " AND (FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Tafel3) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Tafel2) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Tafel3Gl) + ")";
            if (FilterType == 4)
                FilterString = " AND (FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Marsel1) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Marsel3) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Marsel4) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Jersy110) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Marsel5) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Porto) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Monte) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Techno1) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Shervud) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Techno2) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Techno4) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.pFox) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.pFlorenc) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Techno5) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.PR1) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.PR2) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.PR3) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.PRU8) + ")";
            if (FilterType == 5)
                FilterString = " AND (FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.TechnoN) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Antalia) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Nord95) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.epFox) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Fat) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Leon) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Limog) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Luk) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.ep066Marsel4) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.ep110Jersy) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.ep018Marsel1) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.ep043Shervud) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.ep112) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Urban) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Alby) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Bruno) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.epsh406Techno4) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.LukPVH) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Milano) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Praga) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Sigma) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Venecia) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Bergamo) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.ep041) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.ep071) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.ep206) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.ep216) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.ep111) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Boston) + ")";

            if (FilterType == 6)
                FilterString = " AND (FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Tafel1_19) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Tafel1Gl_19) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Tafel1R1) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Tafel1R1Gl) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Tafel1R2) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Tafel1R2Gl) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Grand) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.GrandVg) + ")";

            SelectCommand = @"SELECT FrontsOrdersID, FrontsOrders.MainOrderID, FrontID, ColorID, PatinaID, InsetTypeID, InsetColorID, TechnoProfileID, TechnoColorID, TechnoInsetTypeID, TechnoInsetColorID, 
Height, Width, Count, FrontConfigID, FrontsOrders.Square, FrontsOrders.FactoryID, IsNonStandard, FrontsOrders.Notes, ClientName FROM FrontsOrders INNER JOIN
                MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID INNER JOIN
                MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID INNER JOIN
                infiniu2_marketingreference.dbo.Clients AS Client ON MegaOrders.ClientID = Client.ClientID
                WHERE FrontsOrders.MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE ProfilWorkAssignmentID=" + WorkAssignmentID + ")" + BatchFactoryFilter + ")" + FactoryFilter + FilterString;
            if (FactoryID == 2)
                SelectCommand = @"SELECT FrontsOrdersID, FrontsOrders.MainOrderID, FrontID, ColorID, PatinaID, InsetTypeID, InsetColorID, TechnoProfileID, TechnoColorID, TechnoInsetTypeID, TechnoInsetColorID, Height, Width, Count, FrontConfigID, FrontsOrders.Square, FrontsOrders.FactoryID, IsNonStandard, FrontsOrders.Notes, ClientName FROM FrontsOrders  INNER JOIN
                    MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID INNER JOIN
                    MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID INNER JOIN
                    infiniu2_marketingreference.dbo.Clients AS Client ON MegaOrders.ClientID = Client.ClientID
                    WHERE FrontsOrders.MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE TPSWorkAssignmentID=" + WorkAssignmentID + ")" + BatchFactoryFilter + ")" + FactoryFilter + FilterString;
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                DT.Clear();
                DA.Fill(DT);
            }
            foreach (DataRow item in DT.Rows)
            {
                DataRow NewRow = AllBatchFrontsDT.NewRow();
                NewRow.ItemArray = item.ItemArray;
                NewRow["GroupType"] = 1;
                AllBatchFrontsDT.Rows.Add(NewRow);
            }

            SelectCommand = @"SELECT FrontsOrdersID, FrontsOrders.MainOrderID, FrontID, ColorID, PatinaID, InsetTypeID, InsetColorID, TechnoProfileID, TechnoColorID, TechnoInsetTypeID, TechnoInsetColorID, 
Height, Width, Count, FrontConfigID, FrontsOrders.Square, FrontsOrders.FactoryID, IsNonStandard, FrontsOrders.Notes, ClientName FROM FrontsOrders INNER JOIN
                MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID INNER JOIN
                infiniu2_zovreference.dbo.Clients AS Client ON MainOrders.ClientID = Client.ClientID
                WHERE FrontsOrders.MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE ProfilWorkAssignmentID=" + WorkAssignmentID + ")" + BatchFactoryFilter + ")" + FactoryFilter + FilterString;
            if (FactoryID == 2)
                SelectCommand = @"SELECT FrontsOrdersID, FrontsOrders.MainOrderID, FrontID, ColorID, PatinaID, InsetTypeID, InsetColorID, TechnoProfileID, TechnoColorID, TechnoInsetTypeID, TechnoInsetColorID, Height, Width, Count, FrontConfigID, FrontsOrders.Square, FrontsOrders.FactoryID, IsNonStandard, FrontsOrders.Notes, ClientName FROM FrontsOrders INNER JOIN
                    MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID INNER JOIN
                    infiniu2_zovreference.dbo.Clients AS Client ON MainOrders.ClientID = Client.ClientID
                    WHERE FrontsOrders.MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE TPSWorkAssignmentID=" + WorkAssignmentID + ")" + BatchFactoryFilter + ")" + FactoryFilter + FilterString;
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.ZOVOrdersConnectionString))
            {
                DT.Clear();
                DA.Fill(DT);
            }
            foreach (DataRow item in DT.Rows)
            {
                DataRow NewRow = AllBatchFrontsDT.NewRow();
                NewRow.ItemArray = item.ItemArray;
                NewRow["GroupType"] = 0;
                AllBatchFrontsDT.Rows.Add(NewRow);
            }
            //foreach (DataRow item in DT.Rows)
            //    AllBatchFrontsDT.Rows.Add(item.ItemArray);

            //AllBatchFrontsDT.DefaultView.Sort = "Front, FrameColor, InsetType";
            return AllBatchFrontsDT.Rows.Count > 0;
        }

        public bool FilterFrontsByBatch(bool ZOV, int BatchID, int FactoryID, int FilterType)
        {
            string FilterString = string.Empty;
            string OrdersConnectionString = string.Empty;
            if (ZOV)
                OrdersConnectionString = ConnectionStrings.ZOVOrdersConnectionString;
            else
                OrdersConnectionString = ConnectionStrings.MarketingOrdersConnectionString;

            string BatchFactoryFilter = string.Empty;
            string FactoryFilter = string.Empty;
            string SelectCommand = string.Empty;

            if (FactoryID != 0)
            {
                BatchFactoryFilter = " AND BatchDetails.FactoryID = " + FactoryID;
                FactoryFilter = " AND FrontsOrders.FactoryID = " + FactoryID;
            }

            BatchFrontsDT.Clear();

            if (FilterType == 1)
                FilterString = " AND FrontsOrders.FrontID IN (1975,1976,1977,1978,15760, 3737, 30501,30502,30503,30504,30505,30506,16269,30364,30366,30367,28945,3727,3728,3729,3730,3731,3732,3733,3734,3735,3736,3737,3739,3740,3741,3742,3743,3744,3745,3746,3747,3748,15108,27914,29597)";
            if (FilterType == 2)
                FilterString = " AND (FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.KansasPat) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Kansas) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Sofia) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Lorenzo) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Elegant) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.ElegantPat) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Patricia1) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Turin1) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Turin1_1) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Dakota) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.DakotaPat) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Turin3) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.LeonTPS) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Infiniti) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.InfinitiPat) + ")";
            if (FilterType == 3)
                FilterString = " AND (FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Tafel3) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Tafel2) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Tafel3Gl) + ")";
            if (FilterType == 4)
                FilterString = " AND (FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Marsel1) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Marsel3) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Marsel4) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Jersy110) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Marsel5) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Porto) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Monte) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Techno1) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Shervud) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Techno2) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Techno4) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.pFox) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.pFlorenc) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Techno5) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.PR1) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.PR2) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.PR3) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.PRU8) + ")";
            if (FilterType == 5)
                FilterString = " AND (FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.TechnoN) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Antalia) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Nord95) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.epFox) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Fat) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Leon) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Limog) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Luk) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.ep066Marsel4) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.ep110Jersy) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.ep018Marsel1) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.ep043Shervud) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.ep112) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Urban) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Alby) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Bruno) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.epsh406Techno4) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.LukPVH) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Milano) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Praga) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Sigma) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Venecia) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Bergamo) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.ep041) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.ep071) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.ep206) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.ep216) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.ep111) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Boston) + ")";
            if (FilterType == 6)
                FilterString = " AND (FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Tafel1_19) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Tafel1Gl_19) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Tafel1R1) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Tafel1R1Gl) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Tafel1R2) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Tafel1R2Gl) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Grand) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.GrandVg) + ")";

            if (ZOV)
            {
                SelectCommand = @"SELECT FrontsOrders.*, ClientName FROM FrontsOrders INNER JOIN
                MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID INNER JOIN
                infiniu2_zovreference.dbo.Clients AS Client ON MainOrders.ClientID = Client.ClientID
                WHERE FrontsOrders.MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID = " + BatchID + BatchFactoryFilter + ")" + FactoryFilter + FilterString;
            }
            else
            {
                SelectCommand = @"SELECT FrontsOrders.*, ClientName, FrontsOrders.Notes FROM FrontsOrders INNER JOIN
                MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID INNER JOIN
                MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID INNER JOIN
                infiniu2_marketingreference.dbo.Clients AS Client ON MegaOrders.ClientID = Client.ClientID
                WHERE FrontsOrders.MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID = " + BatchID + BatchFactoryFilter + ")" + FactoryFilter + FilterString;
            }

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, OrdersConnectionString))
            {
                DA.Fill(BatchFrontsDT);
                //BatchFrontsDT.DefaultView.Sort = "Front, FrameColor, InsetType";
            }

            return BatchFrontsDT.Rows.Count > 0;
        }

        public void FilterFrameColors(int FrontID, int FrontTypeID)
        {
            FrameColorsSummaryBS.Filter = "FrontID=" + FrontID + " AND FrontTypeID=" + FrontTypeID;
            FrameColorsSummaryBS.MoveFirst();
        }

        public void FilterInsetTypes(int FrontID, int FrontTypeID, int ColorID, int PatinaID)
        {
            InsetTypesSummaryBS.Filter = "FrontID=" + FrontID + " AND FrontTypeID=" + FrontTypeID +
                " AND ColorID=" + ColorID + " AND PatinaID=" + PatinaID;
            InsetTypesSummaryBS.MoveFirst();
        }

        public void FilterInsetColors(int FrontID, int FrontTypeID, int ColorID, int PatinaID, int InsetTypeID)
        {
            InsetColorsSummaryBS.Filter = "FrontID=" + FrontID + " AND FrontTypeID=" +
                FrontTypeID + " AND InsetTypeID=" + InsetTypeID + " AND ColorID=" + ColorID + " AND PatinaID=" + PatinaID;
            InsetColorsSummaryBS.MoveFirst();
        }

        public void FilterSizes(int FrontID, int FrontTypeID, int ColorID, int PatinaID, int InsetTypeID, int InsetColorID)
        {
            SizesSummaryBS.Filter = "FrontID=" + FrontID + " AND FrontTypeID=" +
                FrontTypeID + " AND ColorID=" + ColorID + " AND PatinaID=" + PatinaID + " AND InsetTypeID=" + InsetTypeID +
                " AND InsetColorID=" + InsetColorID;
            SizesSummaryBS.MoveFirst();
        }

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

        //private void GetGridMargins(int FrontID, ref int MarginHeight, ref int MarginWidth)
        //{
        //    DataRow[] Rows = InsetMarginsDT.Select("FrontID = " + FrontID);
        //    if (Rows.Count() == 0)
        //        return;
        //    MarginHeight = Convert.ToInt32(Rows[0]["GridHeight"]);
        //    MarginWidth = Convert.ToInt32(Rows[0]["GridWidth"]);
        //}

        public string GetFrontName(int FrontID)
        {
            string FrontName = "";
            try
            {
                DataRow[] Rows = FrontsDataTable.Select("FrontID = " + FrontID);
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
            string PatinaName = string.Empty;
            try
            {
                DataRow[] Rows = PatinaDataTable.Select("PatinaID = " + PatinaID);
                PatinaName = Rows[0]["PatinaName"].ToString();
            }
            catch
            {
                return string.Empty;
            }
            return PatinaName;
        }

        public string GetInsetTypeName(int InsetTypeID)
        {
            string InsetType = string.Empty;
            try
            {
                DataRow[] Rows = InsetTypesDataTable.Select("InsetTypeID = " + InsetTypeID);
                InsetType = Rows[0]["InsetTypeName"].ToString();
            }
            catch
            {
                return string.Empty;
            }
            return InsetType;
        }

        public string GetInsetColorName(int ColorID)
        {
            string ColorName = string.Empty;
            try
            {
                DataRow[] Rows = InsetColorsDataTable.Select("InsetColorID = " + ColorID);
                ColorName = Rows[0]["InsetColorName"].ToString();
            }
            catch
            {
                return string.Empty;
            }
            return ColorName;
        }

        public bool FrontsSummary(ref decimal TotalSquare, ref int TotalCount, ref int TotalCurvedCount)
        {
            FrontsSummaryDT.Clear();
            decimal Square = 0;
            int Count = 0;
            DataTable Table = new DataTable();

            using (DataView DV = new DataView(AllBatchFrontsDT))
            {
                Table = DV.ToTable(true, new string[] { "FrontID" });
            }

            for (int i = 0; i < Table.Rows.Count; i++)
            {
                DataRow[] Rows = AllBatchFrontsDT.Select("FrontID=" + Convert.ToInt32(Table.Rows[i]["FrontID"]) +
                    " AND Width<>-1");
                if (Rows.Count() != 0)
                {
                    foreach (DataRow row in Rows)
                    {
                        Square += Convert.ToDecimal(row["Square"]);
                        TotalSquare += Convert.ToDecimal(row["Square"]);
                        Count += Convert.ToInt32(row["Count"]);
                        TotalCount += Convert.ToInt32(row["Count"]);
                    }

                    DataRow NewRow = FrontsSummaryDT.NewRow();
                    NewRow["Front"] = GetFrontName(Convert.ToInt32(Table.Rows[i]["FrontID"]));
                    NewRow["FrontID"] = Convert.ToInt32(Table.Rows[i]["FrontID"]);
                    NewRow["FrontTypeID"] = 0;
                    NewRow["Square"] = decimal.Round(Square, 2, MidpointRounding.AwayFromZero);
                    NewRow["Count"] = Count;
                    FrontsSummaryDT.Rows.Add(NewRow);

                    Square = 0;
                    Count = 0;
                }
                DataRow[] CurvedRows = AllBatchFrontsDT.Select("FrontID=" + Convert.ToInt32(Table.Rows[i]["FrontID"]) +
                    " AND Width=-1");
                if (CurvedRows.Count() != 0)
                {
                    foreach (DataRow row in CurvedRows)
                    {
                        Count += Convert.ToInt32(row["Count"]);
                        TotalCurvedCount += Convert.ToInt32(row["Count"]);
                    }

                    DataRow CurvedNewRow = FrontsSummaryDT.NewRow();
                    CurvedNewRow["Front"] = GetFrontName(Convert.ToInt32(Table.Rows[i]["FrontID"])) + " гн.";
                    CurvedNewRow["FrontID"] = Convert.ToInt32(Table.Rows[i]["FrontID"]);
                    CurvedNewRow["FrontTypeID"] = 1;
                    CurvedNewRow["Square"] = 0;
                    CurvedNewRow["Count"] = Count;
                    FrontsSummaryDT.Rows.Add(CurvedNewRow);

                    Count = 0;
                }
            }

            Table.Dispose();
            FrontsSummaryDT.DefaultView.Sort = "Square DESC";
            FrontsSummaryBS.MoveFirst();

            return FrontsSummaryDT.Rows.Count > 0;
        }

        public bool FrameColorsSummary(ref decimal TotalSquare, ref int TotalCount, ref int TotalCurvedCount)
        {
            FrameColorsSummaryDT.Clear();
            decimal Square = 0;
            int Count = 0;
            DataTable Table = new DataTable();

            using (DataView DV = new DataView(AllBatchFrontsDT))
            {
                Table = DV.ToTable(true, new string[] { "FrontID", "ColorID", "PatinaID" });
            }

            for (int i = 0; i < Table.Rows.Count; i++)
            {
                DataRow[] Rows = AllBatchFrontsDT.Select("FrontID=" + Convert.ToInt32(Table.Rows[i]["FrontID"]) +
                    " AND Width<>-1 AND ColorID=" + Convert.ToInt32(Table.Rows[i]["ColorID"]) +
                    " AND PatinaID=" + Convert.ToInt32(Table.Rows[i]["PatinaID"]));
                if (Rows.Count() != 0)
                {
                    foreach (DataRow row in Rows)
                    {
                        Square += Convert.ToDecimal(row["Square"]);
                        TotalSquare += Convert.ToDecimal(row["Square"]);
                        Count += Convert.ToInt32(row["Count"]);
                        TotalCount += Convert.ToInt32(row["Count"]);
                    }

                    DataRow NewRow = FrameColorsSummaryDT.NewRow();
                    if (Convert.ToInt32(Table.Rows[i]["PatinaID"]) == -1)
                        NewRow["FrameColor"] = GetColorName(Convert.ToInt32(Table.Rows[i]["ColorID"]));
                    else
                        NewRow["FrameColor"] = GetColorName(Convert.ToInt32(Table.Rows[i]["ColorID"])) + " " + GetPatinaName(Convert.ToInt32(Table.Rows[i]["PatinaID"]));
                    NewRow["FrontID"] = Convert.ToInt32(Table.Rows[i]["FrontID"]);
                    NewRow["FrontTypeID"] = 0;
                    NewRow["ColorID"] = Convert.ToInt32(Table.Rows[i]["ColorID"]);
                    NewRow["PatinaID"] = Convert.ToInt32(Table.Rows[i]["PatinaID"]);
                    NewRow["Square"] = decimal.Round(Square, 2, MidpointRounding.AwayFromZero);
                    NewRow["Count"] = Count;
                    FrameColorsSummaryDT.Rows.Add(NewRow);

                    Square = 0;
                    Count = 0;
                }
                DataRow[] CurvedRows = AllBatchFrontsDT.Select("FrontID=" + Convert.ToInt32(Table.Rows[i]["FrontID"]) +
                    " AND Width=-1 AND ColorID=" + Convert.ToInt32(Table.Rows[i]["ColorID"]) +
                    " AND PatinaID=" + Convert.ToInt32(Table.Rows[i]["PatinaID"]));
                if (CurvedRows.Count() != 0)
                {
                    foreach (DataRow row in CurvedRows)
                    {
                        Count += Convert.ToInt32(row["Count"]);
                        TotalCurvedCount += Convert.ToInt32(row["Count"]);
                    }

                    DataRow CurvedNewRow = FrameColorsSummaryDT.NewRow();
                    if (Convert.ToInt32(Table.Rows[i]["PatinaID"]) == -1)
                        CurvedNewRow["FrameColor"] = GetColorName(Convert.ToInt32(Table.Rows[i]["ColorID"]));
                    else
                        CurvedNewRow["FrameColor"] = GetColorName(Convert.ToInt32(Table.Rows[i]["ColorID"])) + " " + GetPatinaName(Convert.ToInt32(Table.Rows[i]["PatinaID"]));
                    CurvedNewRow["FrontID"] = Convert.ToInt32(Table.Rows[i]["FrontID"]);
                    CurvedNewRow["FrontTypeID"] = 1;
                    CurvedNewRow["ColorID"] = Convert.ToInt32(Table.Rows[i]["ColorID"]);
                    CurvedNewRow["PatinaID"] = Convert.ToInt32(Table.Rows[i]["PatinaID"]);
                    CurvedNewRow["Square"] = 0;
                    CurvedNewRow["Count"] = Count;
                    FrameColorsSummaryDT.Rows.Add(CurvedNewRow);

                    Count = 0;
                }
            }

            Table.Dispose();
            FrameColorsSummaryDT.DefaultView.Sort = "Square DESC";
            FrameColorsSummaryBS.MoveFirst();

            return FrameColorsSummaryDT.Rows.Count > 0;
        }

        public bool InsetTypesSummary(ref decimal TotalSquare, ref int TotalCount, ref int TotalCurvedCount)
        {
            InsetTypesSummaryDT.Clear();
            int Count = 0;
            int MarginHeight = 0;
            int MarginWidth = 0;
            decimal Square = 0;
            decimal GridHeight = 0;
            decimal GridWidth = 0;
            DataTable Table = new DataTable();

            using (DataView DV = new DataView(AllBatchFrontsDT))
            {
                Table = DV.ToTable(true, new string[] { "FrontID", "ColorID", "PatinaID", "InsetTypeID" });
            }

            for (int i = 0; i < Table.Rows.Count; i++)
            {
                DataRow[] Rows = AllBatchFrontsDT.Select("FrontID=" + Convert.ToInt32(Table.Rows[i]["FrontID"]) +
                    " AND ColorID=" + Convert.ToInt32(Table.Rows[i]["ColorID"]) +
                    " AND PatinaID=" + Convert.ToInt32(Table.Rows[i]["PatinaID"]) +
                    " AND Width<>-1 AND InsetTypeID=" + Convert.ToInt32(Table.Rows[i]["InsetTypeID"]));
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
                        Square += decimal.Round(GridHeight * GridWidth * Convert.ToInt32(row["Count"]) / 1000000, 2, MidpointRounding.AwayFromZero);
                        TotalSquare += decimal.Round(GridHeight * GridWidth * Convert.ToInt32(row["Count"]) / 1000000, 2, MidpointRounding.AwayFromZero);
                        Count += Convert.ToInt32(row["Count"]);
                        TotalCount += Convert.ToInt32(row["Count"]);
                    }

                    DataRow NewRow = InsetTypesSummaryDT.NewRow();
                    NewRow["InsetType"] = GetInsetTypeName(Convert.ToInt32(Table.Rows[i]["InsetTypeID"]));
                    NewRow["FrontID"] = Convert.ToInt32(Table.Rows[i]["FrontID"]);
                    NewRow["ColorID"] = Convert.ToInt32(Table.Rows[i]["ColorID"]);
                    NewRow["PatinaID"] = Convert.ToInt32(Table.Rows[i]["PatinaID"]);
                    NewRow["FrontTypeID"] = 0;
                    NewRow["InsetTypeID"] = Table.Rows[i]["InsetTypeID"];
                    if (Convert.ToInt32(Table.Rows[i]["InsetTypeID"]) == 4)
                        Square = 0;
                    NewRow["Square"] = decimal.Round(Square, 2, MidpointRounding.AwayFromZero);
                    NewRow["Count"] = Count;
                    InsetTypesSummaryDT.Rows.Add(NewRow);

                    Square = 0;
                    Count = 0;
                }
                DataRow[] CurvedRows = AllBatchFrontsDT.Select("FrontID=" + Convert.ToInt32(Table.Rows[i]["FrontID"]) +
                    " AND ColorID=" + Convert.ToInt32(Table.Rows[i]["ColorID"]) +
                    " AND PatinaID=" + Convert.ToInt32(Table.Rows[i]["PatinaID"]) +
                    " AND Width=-1 AND InsetTypeID=" + Convert.ToInt32(Table.Rows[i]["InsetTypeID"]));
                if (CurvedRows.Count() != 0)
                {
                    foreach (DataRow row in CurvedRows)
                    {
                        Count += Convert.ToInt32(row["Count"]);
                        TotalCurvedCount += Convert.ToInt32(row["Count"]);
                    }

                    DataRow CurvedNewRow = InsetTypesSummaryDT.NewRow();
                    CurvedNewRow["InsetType"] = GetInsetTypeName(Convert.ToInt32(Table.Rows[i]["InsetTypeID"]));
                    CurvedNewRow["FrontID"] = Convert.ToInt32(Table.Rows[i]["FrontID"]);
                    CurvedNewRow["ColorID"] = Convert.ToInt32(Table.Rows[i]["ColorID"]);
                    CurvedNewRow["PatinaID"] = Convert.ToInt32(Table.Rows[i]["PatinaID"]);
                    CurvedNewRow["FrontTypeID"] = 1;
                    CurvedNewRow["InsetTypeID"] = Convert.ToInt32(Table.Rows[i]["InsetTypeID"]);
                    CurvedNewRow["Square"] = 0;
                    CurvedNewRow["Count"] = Count;
                    InsetTypesSummaryDT.Rows.Add(CurvedNewRow);

                    Count = 0;
                }
            }

            Table.Dispose();
            InsetTypesSummaryDT.DefaultView.Sort = "Square DESC";
            InsetTypesSummaryBS.MoveFirst();

            return InsetTypesSummaryDT.Rows.Count > 0;
        }

        public bool InsetColorsSummary(ref decimal TotalSquare, ref int TotalCount, ref int TotalCurvedCount)
        {
            InsetColorsSummaryDT.Clear();
            int Count = 0;
            int MarginHeight = 0;
            int MarginWidth = 0;
            decimal Square = 0;
            decimal GridHeight = 0;
            decimal GridWidth = 0;
            DataTable Table = new DataTable();

            using (DataView DV = new DataView(AllBatchFrontsDT))
            {
                Table = DV.ToTable(true, new string[] { "FrontID", "ColorID", "PatinaID", "InsetTypeID", "InsetColorID" });
            }

            for (int i = 0; i < Table.Rows.Count; i++)
            {
                DataRow[] Rows = AllBatchFrontsDT.Select("FrontID=" + Convert.ToInt32(Table.Rows[i]["FrontID"]) +
                    " AND ColorID=" + Convert.ToInt32(Table.Rows[i]["ColorID"]) +
                    " AND PatinaID=" + Convert.ToInt32(Table.Rows[i]["PatinaID"]) +
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
                        Square += decimal.Round(GridHeight * GridWidth * Convert.ToInt32(row["Count"]) / 1000000, 2, MidpointRounding.AwayFromZero);
                        TotalSquare += decimal.Round(GridHeight * GridWidth * Convert.ToInt32(row["Count"]) / 1000000, 2, MidpointRounding.AwayFromZero);
                        Count += Convert.ToInt32(row["Count"]);
                        TotalCount += Convert.ToInt32(row["Count"]);
                    }

                    DataRow NewRow = InsetColorsSummaryDT.NewRow();
                    NewRow["InsetColor"] = GetInsetColorName(Convert.ToInt32(Table.Rows[i]["InsetColorID"]));
                    NewRow["FrontID"] = Convert.ToInt32(Table.Rows[i]["FrontID"]);
                    NewRow["ColorID"] = Convert.ToInt32(Table.Rows[i]["ColorID"]);
                    NewRow["PatinaID"] = Convert.ToInt32(Table.Rows[i]["PatinaID"]);
                    NewRow["FrontTypeID"] = 0;
                    NewRow["InsetTypeID"] = Convert.ToInt32(Table.Rows[i]["InsetTypeID"]);
                    NewRow["InsetColorID"] = Convert.ToInt32(Table.Rows[i]["InsetColorID"]);
                    if (Convert.ToInt32(Table.Rows[i]["InsetTypeID"]) == 4)
                        Square = 0;
                    NewRow["Square"] = decimal.Round(Square, 2, MidpointRounding.AwayFromZero);
                    NewRow["Count"] = Count;
                    InsetColorsSummaryDT.Rows.Add(NewRow);

                    Square = 0;
                    Count = 0;
                }
                DataRow[] CurvedRows = AllBatchFrontsDT.Select("FrontID=" + Convert.ToInt32(Table.Rows[i]["FrontID"]) +
                    " AND ColorID=" + Convert.ToInt32(Table.Rows[i]["ColorID"]) +
                    " AND PatinaID=" + Convert.ToInt32(Table.Rows[i]["PatinaID"]) +
                    " AND Width=-1 AND InsetTypeID=" + Convert.ToInt32(Table.Rows[i]["InsetTypeID"]) +
                    " AND InsetColorID=" + Convert.ToInt32(Table.Rows[i]["InsetColorID"]));
                if (CurvedRows.Count() != 0)
                {
                    foreach (DataRow row in CurvedRows)
                    {
                        Count += Convert.ToInt32(row["Count"]);
                        TotalCurvedCount += Convert.ToInt32(row["Count"]);
                    }

                    DataRow CurvedNewRow = InsetColorsSummaryDT.NewRow();
                    CurvedNewRow["InsetColor"] = GetInsetColorName(Convert.ToInt32(Table.Rows[i]["InsetColorID"]));
                    CurvedNewRow["FrontID"] = Convert.ToInt32(Table.Rows[i]["FrontID"]);
                    CurvedNewRow["ColorID"] = Convert.ToInt32(Table.Rows[i]["ColorID"]);
                    CurvedNewRow["PatinaID"] = Convert.ToInt32(Table.Rows[i]["PatinaID"]);
                    CurvedNewRow["FrontTypeID"] = 1;
                    CurvedNewRow["InsetTypeID"] = Convert.ToInt32(Table.Rows[i]["InsetTypeID"]);
                    CurvedNewRow["InsetColorID"] = Convert.ToInt32(Table.Rows[i]["InsetColorID"]);
                    CurvedNewRow["Square"] = 0;
                    CurvedNewRow["Count"] = Count;
                    InsetColorsSummaryDT.Rows.Add(CurvedNewRow);

                    Count = 0;
                }
            }

            Table.Dispose();
            InsetColorsSummaryDT.DefaultView.Sort = "Square DESC";
            InsetColorsSummaryBS.MoveFirst();

            return InsetColorsSummaryDT.Rows.Count > 0;
        }

        public bool SizesSummary(ref decimal TotalSquare, ref int TotalCount, ref int TotalCurvedCount)
        {
            SizesSummaryDT.Clear();
            decimal Square = 0;
            int Count = 0;

            DataTable Table = new DataTable();

            using (DataView DV = new DataView(AllBatchFrontsDT))
            {
                Table = DV.ToTable(true, new string[] { "FrontID", "ColorID", "PatinaID", "InsetTypeID", "InsetColorID", "Height", "Width" });
            }

            for (int i = 0; i < Table.Rows.Count; i++)
            {
                DataRow[] Rows = AllBatchFrontsDT.Select("FrontID=" + Convert.ToInt32(Table.Rows[i]["FrontID"]) +
                    " AND Width<>-1 AND ColorID=" + Convert.ToInt32(Table.Rows[i]["ColorID"]) +
                    " AND PatinaID=" + Convert.ToInt32(Table.Rows[i]["PatinaID"]) +
                    " AND InsetTypeID=" + Convert.ToInt32(Table.Rows[i]["InsetTypeID"]) +
                    " AND InsetColorID=" + Convert.ToInt32(Table.Rows[i]["InsetColorID"]) +
                    " AND Height=" + Convert.ToInt32(Table.Rows[i]["Height"]) +
                    " AND Width=" + Convert.ToInt32(Table.Rows[i]["Width"]));
                if (Rows.Count() != 0)
                {
                    foreach (DataRow row in Rows)
                    {
                        Square += Convert.ToDecimal(row["Square"]);
                        TotalSquare += Convert.ToDecimal(row["Square"]);
                        Count += Convert.ToInt32(row["Count"]);
                        TotalCount += Convert.ToInt32(row["Count"]);
                    }

                    DataRow NewRow = SizesSummaryDT.NewRow();
                    NewRow["Size"] = Convert.ToInt32(Table.Rows[i]["Height"]) + " x " + Convert.ToInt32(Table.Rows[i]["Width"]);
                    NewRow["FrontID"] = Convert.ToInt32(Table.Rows[i]["FrontID"]);
                    NewRow["FrontTypeID"] = 0;
                    NewRow["ColorID"] = Convert.ToInt32(Table.Rows[i]["ColorID"]);
                    NewRow["PatinaID"] = Convert.ToInt32(Table.Rows[i]["PatinaID"]);
                    NewRow["InsetColorID"] = Convert.ToInt32(Table.Rows[i]["InsetColorID"]);
                    NewRow["InsetTypeID"] = Convert.ToInt32(Table.Rows[i]["InsetTypeID"]);
                    NewRow["Height"] = Convert.ToInt32(Table.Rows[i]["Height"]);
                    NewRow["Width"] = Convert.ToInt32(Table.Rows[i]["Width"]);
                    NewRow["Square"] = decimal.Round(Square, 2, MidpointRounding.AwayFromZero);
                    NewRow["Count"] = Count;
                    SizesSummaryDT.Rows.Add(NewRow);

                    Square = 0;
                    Count = 0;
                }
                DataRow[] CurvedRows = AllBatchFrontsDT.Select("FrontID=" + Convert.ToInt32(Table.Rows[i]["FrontID"]) +
                    " AND Width=-1 AND ColorID=" + Convert.ToInt32(Table.Rows[i]["ColorID"]) +
                    " AND PatinaID=" + Convert.ToInt32(Table.Rows[i]["PatinaID"]) +
                    " AND InsetTypeID=" + Convert.ToInt32(Table.Rows[i]["InsetTypeID"]) +
                    " AND InsetColorID=" + Convert.ToInt32(Table.Rows[i]["InsetColorID"]) +
                    " AND Height=" + Convert.ToInt32(Table.Rows[i]["Height"]) +
                    " AND Width=" + Convert.ToInt32(Table.Rows[i]["Width"]));
                if (CurvedRows.Count() != 0)
                {
                    foreach (DataRow row in CurvedRows)
                    {
                        Count += Convert.ToInt32(row["Count"]);
                        TotalCurvedCount += Convert.ToInt32(row["Count"]);
                    }

                    DataRow CurvedNewRow = SizesSummaryDT.NewRow();
                    CurvedNewRow["Size"] = Convert.ToInt32(Table.Rows[i]["Height"]) + " x " + Convert.ToInt32(Table.Rows[i]["Width"]);
                    CurvedNewRow["FrontID"] = Convert.ToInt32(Table.Rows[i]["FrontID"]);
                    CurvedNewRow["FrontTypeID"] = 1;
                    CurvedNewRow["ColorID"] = Convert.ToInt32(Table.Rows[i]["ColorID"]);
                    CurvedNewRow["PatinaID"] = Convert.ToInt32(Table.Rows[i]["PatinaID"]);
                    CurvedNewRow["InsetColorID"] = Convert.ToInt32(Table.Rows[i]["InsetColorID"]);
                    CurvedNewRow["InsetTypeID"] = Convert.ToInt32(Table.Rows[i]["InsetTypeID"]);
                    CurvedNewRow["Height"] = Convert.ToInt32(Table.Rows[i]["Height"]);
                    CurvedNewRow["Width"] = Convert.ToInt32(Table.Rows[i]["Width"]);
                    CurvedNewRow["Square"] = 0;
                    CurvedNewRow["Count"] = Count;
                    SizesSummaryDT.Rows.Add(CurvedNewRow);

                    Count = 0;
                }
            }

            Table.Dispose();
            SizesSummaryDT.DefaultView.Sort = "Square DESC";
            SizesSummaryBS.MoveFirst();

            return SizesSummaryDT.Rows.Count > 0;
        }

        public DataGridViewComboBoxColumn FrontsColumn
        {
            get
            {
                DataGridViewComboBoxColumn Column = new DataGridViewComboBoxColumn()
                {
                    Name = "FrontsColumn",
                    HeaderText = "Фасад",
                    DataPropertyName = "FrontID",
                    DataSource = new DataView(FrontsDataTable),
                    ValueMember = "FrontID",
                    DisplayMember = "FrontName",
                    DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing
                };
                return Column;
            }
        }

        public DataGridViewComboBoxColumn FrameColorsColumn
        {
            get
            {
                DataGridViewComboBoxColumn Column = new DataGridViewComboBoxColumn();
                Column = new DataGridViewComboBoxColumn()
                {
                    Name = "FrameColorsColumn",
                    HeaderText = "Цвет\r\nпрофиля",
                    DataPropertyName = "ColorID",
                    DataSource = new DataView(FrameColorsDataTable),
                    ValueMember = "ColorID",
                    DisplayMember = "ColorName",
                    DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing,
                    SortMode = DataGridViewColumnSortMode.Automatic
                };
                return Column;
            }
        }

        public DataGridViewComboBoxColumn PatinaColumn
        {
            get
            {
                DataGridViewComboBoxColumn Column = new DataGridViewComboBoxColumn();
                Column = new DataGridViewComboBoxColumn()
                {
                    Name = "PatinaColumn",
                    HeaderText = "Патина",
                    DataPropertyName = "PatinaID",
                    DataSource = new DataView(PatinaDataTable),
                    ValueMember = "PatinaID",
                    DisplayMember = "PatinaName",
                    DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing,
                    SortMode = DataGridViewColumnSortMode.Automatic
                };
                return Column;
            }
        }

        public DataGridViewComboBoxColumn InsetTypesColumn
        {
            get
            {
                DataGridViewComboBoxColumn Column = new DataGridViewComboBoxColumn();
                Column = new DataGridViewComboBoxColumn()
                {
                    Name = "InsetTypesColumn",
                    HeaderText = "Тип\r\nнаполнителя",
                    DataPropertyName = "InsetTypeID",
                    DataSource = new DataView(InsetTypesDataTable),
                    ValueMember = "InsetTypeID",
                    DisplayMember = "InsetTypeName",
                    DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing,
                    SortMode = DataGridViewColumnSortMode.Automatic
                };
                return Column;
            }
        }

        public DataGridViewComboBoxColumn InsetColorsColumn
        {
            get
            {
                DataGridViewComboBoxColumn Column = new DataGridViewComboBoxColumn();
                Column = new DataGridViewComboBoxColumn()
                {
                    Name = "InsetColorsColumn",
                    HeaderText = "Цвет\r\nнаполнителя",
                    DataPropertyName = "InsetColorID",
                    DataSource = new DataView(InsetColorsDataTable),
                    ValueMember = "InsetColorID",
                    DisplayMember = "InsetColorName",
                    DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing,
                    SortMode = DataGridViewColumnSortMode.Automatic
                };
                return Column;
            }
        }

        public DataGridViewComboBoxColumn TechnoProfilesColumn
        {
            get
            {
                DataGridViewComboBoxColumn Column = new DataGridViewComboBoxColumn();
                Column = new DataGridViewComboBoxColumn()
                {
                    Name = "TechnoProfilesColumn",
                    HeaderText = "Тип\r\nпрофиля-2",
                    DataPropertyName = "TechnoProfileID",
                    DataSource = new DataView(TechnoProfilesDataTable),
                    ValueMember = "TechnoProfileID",
                    DisplayMember = "TechnoProfileName",
                    DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing
                };
                return Column;
            }
        }
        public DataGridViewComboBoxColumn TechnoFrameColorsColumn
        {
            get
            {
                DataGridViewComboBoxColumn Column = new DataGridViewComboBoxColumn();
                Column = new DataGridViewComboBoxColumn()
                {
                    Name = "TechnoFrameColorsColumn",
                    HeaderText = "Цвет профиля-2",
                    DataPropertyName = "TechnoColorID",
                    DataSource = new DataView(FrameColorsDataTable),
                    ValueMember = "ColorID",
                    DisplayMember = "ColorName",
                    DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing,
                    SortMode = DataGridViewColumnSortMode.Automatic
                };
                return Column;
            }
        }

        public DataGridViewComboBoxColumn TechnoInsetTypesColumn
        {
            get
            {
                DataGridViewComboBoxColumn Column = new DataGridViewComboBoxColumn();
                Column = new DataGridViewComboBoxColumn()
                {
                    Name = "TechnoInsetTypesColumn",
                    HeaderText = "Тип наполнителя-2",
                    DataPropertyName = "TechnoInsetTypeID",
                    DataSource = new DataView(InsetTypesDataTable),
                    ValueMember = "InsetTypeID",
                    DisplayMember = "InsetTypeName",
                    DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing,
                    SortMode = DataGridViewColumnSortMode.Automatic
                };
                return Column;
            }
        }

        public DataGridViewComboBoxColumn TechnoInsetColorsColumn
        {
            get
            {
                DataGridViewComboBoxColumn Column = new DataGridViewComboBoxColumn();
                Column = new DataGridViewComboBoxColumn()
                {
                    Name = "TechnoInsetColorsColumn",
                    HeaderText = "Цвет наполнителя-2",
                    DataPropertyName = "TechnoInsetColorID",
                    DataSource = new DataView(InsetColorsDataTable),
                    ValueMember = "InsetColorID",
                    DisplayMember = "InsetColorName",
                    DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing,
                    SortMode = DataGridViewColumnSortMode.Automatic
                };
                return Column;
            }
        }

        public DataTable DinstinctFrontsDT(int[] FrontsID)
        {
            DataTable Table = AllBatchFrontsDT.Clone();
            for (int i = 0; i < FrontsID.Count(); i++)
            {
                DataRow[] rows2 = AllBatchFrontsDT.Select("FrontID=" + FrontsID[i]);
                for (int j = 0; j < rows2.Count(); j++)
                    Table.Rows.Add(rows2[j].ItemArray);
            }
            using (DataView DV = new DataView(Table))
            {
                Table = DV.ToTable(true, new string[] { "FrontID" });
            }
            return Table;
        }

        public DataTable OrdersDT => AllBatchFrontsDT;

        public void GetPrintedFronts(int WorkAssignmentID)
        {
            string SelectCommand = @"SELECT * FROM AssignmentsInWork WHERE WorkAssignmentID=" + WorkAssignmentID;

            AllPrintedFrontsDT.Clear();
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(AllPrintedFrontsDT);
            }
        }

        public void SetFrontPrintingStatus()
        {
            for (int i = 0; i < FrontsSummaryDT.Rows.Count; i++)
            {
                DataRow[] rows = AllPrintedFrontsDT.Select("FrontID=" + FrontsSummaryDT.Rows[i]["FrontID"]);
                if (rows.Count() > 0)
                    FrontsSummaryDT.Rows[i]["PrintingStatus"] = 2;
                else
                    FrontsSummaryDT.Rows[i]["PrintingStatus"] = 0;
            }
        }
    }


    public class DecorOrdersManager
    {
        private DataTable ColorsDataTable = null;
        private DataTable PatinaDataTable = null;
        private DataTable PatinaRALDataTable = null;
        private DataTable ProductsDataTable = null;
        private DataTable DecorDataTable = null;

        private DataTable AllBatchDecorDT = null;
        private DataTable BatchDecorDT = null;
        private DataTable DecorProductsSummaryDT = null;
        private DataTable DecorItemsSummaryDT = null;
        private DataTable DecorColorsSummaryDT = null;
        private DataTable DecorSizesSummaryDT = null;

        private BindingSource BatchDecorBS = null;
        private BindingSource DecorProductsSummaryBS = null;
        private BindingSource DecorItemsSummaryBS = null;
        private BindingSource DecorColorsSummaryBS = null;
        private BindingSource DecorSizesSummaryBS = null;

        public BindingSource BatchDecorList => BatchDecorBS;

        public BindingSource DecorProductsSummaryList => DecorProductsSummaryBS;

        public BindingSource DecorItemsSummaryList => DecorItemsSummaryBS;

        public BindingSource DecorColorsSummaryList => DecorColorsSummaryBS;

        public BindingSource DecorSizesSummaryList => DecorSizesSummaryBS;

        public DataGridViewComboBoxColumn ProductColumn
        {
            get
            {
                DataGridViewComboBoxColumn ProductColumn = new DataGridViewComboBoxColumn()
                {
                    Name = "ProductColumn",
                    HeaderText = "Продукт",
                    DataPropertyName = "ProductID",

                    DataSource = new DataView(ProductsDataTable),
                    ValueMember = "ProductID",
                    DisplayMember = "ProductName",
                    DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing
                };
                return ProductColumn;
            }
        }

        public DataGridViewComboBoxColumn ItemColumn
        {
            get
            {
                DataGridViewComboBoxColumn ItemColumn = new DataGridViewComboBoxColumn()
                {
                    Name = "ItemColumn",
                    HeaderText = "Название",
                    DataPropertyName = "DecorID",

                    DataSource = new DataView(DecorDataTable),
                    ValueMember = "DecorID",
                    DisplayMember = "Name",
                    DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing
                };
                return ItemColumn;
            }
        }

        public DataGridViewComboBoxColumn ColorColumn
        {
            get
            {
                DataGridViewComboBoxColumn ColorsColumn = new DataGridViewComboBoxColumn()
                {
                    Name = "ColorsColumn",
                    HeaderText = "Цвет",
                    DataPropertyName = "ColorID",

                    DataSource = new DataView(ColorsDataTable),
                    ValueMember = "ColorID",
                    DisplayMember = "ColorName",
                    DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing
                };
                return ColorsColumn;
            }
        }

        public DataGridViewComboBoxColumn PatinaColumn
        {
            get
            {
                DataGridViewComboBoxColumn PatinaColumn = new DataGridViewComboBoxColumn()
                {
                    Name = "PatinaColumn",
                    HeaderText = "Патина",
                    DataPropertyName = "PatinaID",

                    DataSource = new DataView(PatinaDataTable),
                    ValueMember = "PatinaID",
                    DisplayMember = "PatinaName",
                    DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing
                };
                return PatinaColumn;
            }
        }

        private void GetColorsDT()
        {
            ColorsDataTable = new DataTable();
            ColorsDataTable.Columns.Add(new DataColumn("ColorID", Type.GetType("System.Int64")));
            ColorsDataTable.Columns.Add(new DataColumn("ColorName", Type.GetType("System.String")));
            string SelectCommand = @"SELECT TechStoreID, TechStoreName FROM TechStore
                WHERE TechStoreSubGroupID IN (SELECT TechStoreSubGroupID FROM TechStoreSubGroups WHERE TechStoreGroupID = 11)
                ORDER BY TechStoreName";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.CatalogConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    DA.Fill(DT);
                    {
                        DataRow NewRow = ColorsDataTable.NewRow();
                        NewRow["ColorID"] = -1;
                        NewRow["ColorName"] = "-";
                        ColorsDataTable.Rows.Add(NewRow);
                    }
                    {
                        DataRow NewRow = ColorsDataTable.NewRow();
                        NewRow["ColorID"] = 0;
                        NewRow["ColorName"] = "на выбор";
                        ColorsDataTable.Rows.Add(NewRow);
                    }
                    for (int i = 0; i < DT.Rows.Count; i++)
                    {
                        DataRow NewRow = ColorsDataTable.NewRow();
                        NewRow["ColorID"] = Convert.ToInt64(DT.Rows[i]["TechStoreID"]);
                        NewRow["ColorName"] = DT.Rows[i]["TechStoreName"].ToString();
                        ColorsDataTable.Rows.Add(NewRow);
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
                        DataRow[] rows = ColorsDataTable.Select("ColorID=" + Convert.ToInt32(DT.Rows[i]["InsetColorID"]));
                        if (rows.Count() == 0)
                        {
                            DataRow NewRow = ColorsDataTable.NewRow();
                            NewRow["ColorID"] = Convert.ToInt64(DT.Rows[i]["InsetColorID"]);
                            NewRow["ColorName"] = DT.Rows[i]["InsetColorName"].ToString();
                            ColorsDataTable.Rows.Add(NewRow);
                        }
                    }
                }
            }
        }

        public string GetProductName(int ProductID)
        {
            string Name = string.Empty;
            DataRow[] Rows = ProductsDataTable.Select("ProductID = " + ProductID);
            if (Rows.Count() > 0)
                Name = Rows[0]["ProductName"].ToString();
            return Name;
        }

        public string GetDecorName(int DecorID)
        {
            string Name = string.Empty;
            DataRow[] Rows = DecorDataTable.Select("DecorID = " + DecorID);
            if (Rows.Count() > 0)
                Name = Rows[0]["Name"].ToString();
            return Name;
        }

        public string GetColorName(int ColorID)
        {
            string ColorName = string.Empty;
            try
            {
                DataRow[] Rows = ColorsDataTable.Select("ColorID = " + ColorID);
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
            string PatinaName = string.Empty;
            try
            {
                DataRow[] Rows = PatinaDataTable.Select("PatinaID = " + PatinaID);
                PatinaName = Rows[0]["PatinaName"].ToString();
            }
            catch
            {
                return string.Empty;
            }
            return PatinaName;
        }

        //конструктор
        public DecorOrdersManager()
        {
            Initialize();
        }

        private void Create()
        {
            DecorProductsSummaryDT = new DataTable();
            DecorProductsSummaryDT.Columns.Add(new DataColumn(("DecorProduct"), System.Type.GetType("System.String")));
            DecorProductsSummaryDT.Columns.Add(new DataColumn(("ProductID"), System.Type.GetType("System.Int32")));
            DecorProductsSummaryDT.Columns.Add(new DataColumn(("MeasureID"), System.Type.GetType("System.Int32")));
            DecorProductsSummaryDT.Columns.Add(new DataColumn(("Count"), System.Type.GetType("System.Decimal")));
            DecorProductsSummaryDT.Columns.Add(new DataColumn(("Measure"), System.Type.GetType("System.String")));

            DecorItemsSummaryDT = new DataTable();
            DecorItemsSummaryDT.Columns.Add(new DataColumn(("DecorItem"), System.Type.GetType("System.String")));
            DecorItemsSummaryDT.Columns.Add(new DataColumn(("ProductID"), System.Type.GetType("System.Int32")));
            DecorItemsSummaryDT.Columns.Add(new DataColumn(("MeasureID"), System.Type.GetType("System.Int32")));
            DecorItemsSummaryDT.Columns.Add(new DataColumn(("DecorID"), System.Type.GetType("System.Int32")));
            DecorItemsSummaryDT.Columns.Add(new DataColumn(("Count"), System.Type.GetType("System.Decimal")));
            DecorItemsSummaryDT.Columns.Add(new DataColumn(("Measure"), System.Type.GetType("System.String")));

            DecorColorsSummaryDT = new DataTable();
            DecorColorsSummaryDT.Columns.Add(new DataColumn(("Color"), System.Type.GetType("System.String")));
            DecorColorsSummaryDT.Columns.Add(new DataColumn(("ProductID"), System.Type.GetType("System.Int32")));
            DecorColorsSummaryDT.Columns.Add(new DataColumn(("MeasureID"), System.Type.GetType("System.Int32")));
            DecorColorsSummaryDT.Columns.Add(new DataColumn(("ColorID"), System.Type.GetType("System.Int32")));
            DecorColorsSummaryDT.Columns.Add(new DataColumn(("PatinaID"), System.Type.GetType("System.Int32")));
            DecorColorsSummaryDT.Columns.Add(new DataColumn(("DecorID"), System.Type.GetType("System.Int32")));
            DecorColorsSummaryDT.Columns.Add(new DataColumn(("Count"), System.Type.GetType("System.Decimal")));
            DecorColorsSummaryDT.Columns.Add(new DataColumn(("Measure"), System.Type.GetType("System.String")));

            DecorSizesSummaryDT = new DataTable();
            DecorSizesSummaryDT.Columns.Add(new DataColumn(("ProductID"), System.Type.GetType("System.Int32")));
            DecorSizesSummaryDT.Columns.Add(new DataColumn(("DecorID"), System.Type.GetType("System.Int32")));
            DecorSizesSummaryDT.Columns.Add(new DataColumn(("ColorID"), System.Type.GetType("System.Int32")));
            DecorSizesSummaryDT.Columns.Add(new DataColumn(("PatinaID"), System.Type.GetType("System.Int32")));
            DecorSizesSummaryDT.Columns.Add(new DataColumn(("MeasureID"), System.Type.GetType("System.Int32")));
            DecorSizesSummaryDT.Columns.Add(new DataColumn(("Size"), System.Type.GetType("System.String")));
            DecorSizesSummaryDT.Columns.Add(new DataColumn(("Length"), System.Type.GetType("System.Int32")));
            DecorSizesSummaryDT.Columns.Add(new DataColumn(("Height"), System.Type.GetType("System.Int32")));
            DecorSizesSummaryDT.Columns.Add(new DataColumn(("Width"), System.Type.GetType("System.Int32")));
            DecorSizesSummaryDT.Columns.Add(new DataColumn(("Count"), System.Type.GetType("System.Decimal")));
            DecorSizesSummaryDT.Columns.Add(new DataColumn(("Measure"), System.Type.GetType("System.String")));

            ColorsDataTable = new DataTable();
            ProductsDataTable = new DataTable();
            DecorDataTable = new DataTable();
            PatinaDataTable = new DataTable();

            AllBatchDecorDT = new DataTable();
            BatchDecorDT = new DataTable();

            BatchDecorBS = new BindingSource();
            DecorProductsSummaryBS = new BindingSource();
            DecorItemsSummaryBS = new BindingSource();
            DecorColorsSummaryBS = new BindingSource();
            DecorSizesSummaryBS = new BindingSource();
        }

        private void Fill()
        {
            string SelectCommand = @"SELECT ProductID, ProductName FROM DecorProducts" +
                " WHERE (ProductID IN (SELECT ProductID FROM DecorConfig)) ORDER BY ProductName ASC";
            ProductsDataTable = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(ProductsDataTable);
            }
            DecorDataTable = new DataTable();
            SelectCommand = @"SELECT DISTINCT TechStore.TechStoreID AS DecorID, TechStore.TechStoreName AS Name, DecorConfig.ProductID FROM TechStore 
                INNER JOIN DecorConfig ON TechStore.TechStoreID = DecorConfig.DecorID ORDER BY TechStoreName";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(DecorDataTable);
            }
            GetColorsDT();
            SelectCommand = @"SELECT * FROM Patina ORDER BY PatinaName";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(PatinaDataTable);
            }
            PatinaRALDataTable = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM PatinaRAL",
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
            SelectCommand = @"SELECT TOP 0 DecorOrders.DecorOrderID, DecorOrders.MainOrderID, DecorOrders.ProductID, DecorOrders.DecorID, 
DecorOrders.ColorID, DecorOrders.PatinaID, DecorOrders.Length, DecorOrders.Height, DecorOrders.Width, DecorOrders.Count, DecorOrders.DecorConfigID, DecorOrders.FactoryID, DecorOrders.Notes, DecorConfig.MeasureID
                FROM DecorOrders INNER JOIN 
                infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.ZOVOrdersConnectionString))
            {
                DA.Fill(AllBatchDecorDT);
                DA.Fill(BatchDecorDT);
            }
        }

        private void Binding()
        {
            BatchDecorBS.DataSource = BatchDecorDT;
            DecorProductsSummaryBS.DataSource = DecorProductsSummaryDT;
            DecorItemsSummaryBS.DataSource = DecorItemsSummaryDT;
            DecorColorsSummaryBS.DataSource = DecorColorsSummaryDT;
            DecorSizesSummaryBS.DataSource = DecorSizesSummaryDT;
        }

        public void Initialize()
        {
            Create();
            Fill();
            Binding();
        }

        private void GridSettings(ref PercentageDataGrid MainOrdersDecorOrdersDataGrid)
        {
            MainOrdersDecorOrdersDataGrid.Columns["Height"].HeaderText = "Высота";
            MainOrdersDecorOrdersDataGrid.Columns["Length"].HeaderText = "Длина";
            MainOrdersDecorOrdersDataGrid.Columns["Width"].HeaderText = "Ширина";
            MainOrdersDecorOrdersDataGrid.Columns["Count"].HeaderText = "Кол-во";
            MainOrdersDecorOrdersDataGrid.Columns["Notes"].HeaderText = "Примечание";

            int DisplayIndex = 0;
            MainOrdersDecorOrdersDataGrid.Columns["ProductColumn"].DisplayIndex = DisplayIndex++;
            MainOrdersDecorOrdersDataGrid.Columns["ItemColumn"].DisplayIndex = DisplayIndex++;
            MainOrdersDecorOrdersDataGrid.Columns["ColorsColumn"].DisplayIndex = DisplayIndex++;
            MainOrdersDecorOrdersDataGrid.Columns["PatinaColumn"].DisplayIndex = DisplayIndex++;
            MainOrdersDecorOrdersDataGrid.Columns["Length"].DisplayIndex = DisplayIndex++;
            MainOrdersDecorOrdersDataGrid.Columns["Height"].DisplayIndex = DisplayIndex++;
            MainOrdersDecorOrdersDataGrid.Columns["Width"].DisplayIndex = DisplayIndex++;
            MainOrdersDecorOrdersDataGrid.Columns["Count"].DisplayIndex = DisplayIndex++;
            MainOrdersDecorOrdersDataGrid.Columns["Notes"].DisplayIndex = DisplayIndex++;

            MainOrdersDecorOrdersDataGrid.Columns["ProductColumn"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            MainOrdersDecorOrdersDataGrid.Columns["ProductColumn"].MinimumWidth = 110;
            MainOrdersDecorOrdersDataGrid.Columns["ItemColumn"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            MainOrdersDecorOrdersDataGrid.Columns["ItemColumn"].MinimumWidth = 110;
            MainOrdersDecorOrdersDataGrid.Columns["ColorsColumn"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            MainOrdersDecorOrdersDataGrid.Columns["ColorsColumn"].MinimumWidth = 110;
            MainOrdersDecorOrdersDataGrid.Columns["PatinaColumn"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            MainOrdersDecorOrdersDataGrid.Columns["PatinaColumn"].MinimumWidth = 110;
            MainOrdersDecorOrdersDataGrid.Columns["Height"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MainOrdersDecorOrdersDataGrid.Columns["Height"].Width = 85;
            MainOrdersDecorOrdersDataGrid.Columns["Length"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MainOrdersDecorOrdersDataGrid.Columns["Length"].Width = 85;
            MainOrdersDecorOrdersDataGrid.Columns["Width"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MainOrdersDecorOrdersDataGrid.Columns["Width"].Width = 85;
            MainOrdersDecorOrdersDataGrid.Columns["Count"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            MainOrdersDecorOrdersDataGrid.Columns["Count"].Width = 85;

            foreach (DataGridViewColumn Column in MainOrdersDecorOrdersDataGrid.Columns)
            {
                Column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
        }

        public bool FilterDecorByBatch(bool ZOV, int BatchID, int FactoryID)
        {
            string OrdersConnectionString = string.Empty;
            if (ZOV)
                OrdersConnectionString = ConnectionStrings.ZOVOrdersConnectionString;
            else
                OrdersConnectionString = ConnectionStrings.MarketingOrdersConnectionString;

            string BatchFactoryFilter = string.Empty;
            string FactoryFilter = string.Empty;
            string SelectCommand = string.Empty;

            if (FactoryID != 0)
            {
                BatchFactoryFilter = " AND BatchDetails.FactoryID = " + FactoryID;
                FactoryFilter = " AND DecorOrders.FactoryID = " + FactoryID;
            }

            BatchDecorDT.Clear();

            SelectCommand = @"SELECT DecorOrders.DecorOrderID, DecorOrders.MainOrderID, DecorOrders.ProductID, DecorOrders.DecorID, DecorOrders.ColorID, DecorOrders.PatinaID, 
DecorOrders.Length, DecorOrders.Height, DecorOrders.Width, DecorOrders.Count, DecorOrders.DecorConfigID, DecorOrders.FactoryID, DecorOrders.Notes,  DecorConfig.MeasureID
                FROM DecorOrders INNER JOIN 
                infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID
                WHERE MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID = " + BatchID + BatchFactoryFilter + ")" + FactoryFilter;

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, OrdersConnectionString))
            {
                DA.Fill(BatchDecorDT);

                foreach (DataRow Row in BatchDecorDT.Rows)
                {
                    //if (Convert.ToInt32(Row["ColorID"]) == -1)
                    //    Row["ColorID"] = 0;
                }
                //BatchDecorDT.DefaultView.Sort = "ProductName, Name";
            }

            return BatchDecorDT.Rows.Count > 0;
        }

        public bool FilterDecorByWorkAssignment(int WorkAssignmentID, int FactoryID)
        {
            string OrdersConnectionString = string.Empty;

            string BatchFactoryFilter = string.Empty;
            string FactoryFilter = string.Empty;
            string SelectCommand = string.Empty;

            if (FactoryID != 0)
            {
                BatchFactoryFilter = " AND BatchDetails.FactoryID = " + FactoryID;
                FactoryFilter = " AND DecorOrders.FactoryID = " + FactoryID;
            }
            DataTable DT = AllBatchDecorDT.Clone();
            AllBatchDecorDT.Clear();

            SelectCommand = @"SELECT DecorOrders.DecorOrderID, DecorOrders.MainOrderID, DecorOrders.ProductID, DecorOrders.DecorID, DecorOrders.ColorID, DecorOrders.PatinaID, DecorOrders.Length, DecorOrders.Height, DecorOrders.Width, DecorOrders.Count, DecorOrders.DecorConfigID, DecorOrders.FactoryID, DecorOrders.Notes,  DecorConfig.MeasureID
                FROM DecorOrders INNER JOIN 
                infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID
                WHERE MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE ProfilWorkAssignmentID=" + WorkAssignmentID + ")" + BatchFactoryFilter + ")" + FactoryFilter;
            if (FactoryID == 2)
                SelectCommand = @"SELECT DecorOrders.DecorOrderID, DecorOrders.MainOrderID, DecorOrders.ProductID, DecorOrders.DecorID, DecorOrders.ColorID, DecorOrders.PatinaID, DecorOrders.Length, DecorOrders.Height, DecorOrders.Width, DecorOrders.Count, DecorOrders.DecorConfigID, DecorOrders.FactoryID, DecorOrders.Notes, DecorConfig.MeasureID
                    FROM DecorOrders INNER JOIN 
                    infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID
                    WHERE MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE TPSWorkAssignmentID=" + WorkAssignmentID + ")" + BatchFactoryFilter + ")" + FactoryFilter;
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                DT.Clear();
                DA.Fill(DT);

                foreach (DataRow Row in DT.Rows)
                {
                    //if (Convert.ToInt32(Row["ColorID"]) == -1)
                    //    Row["ColorID"] = 0;
                }
            }
            foreach (DataRow item in DT.Rows)
                AllBatchDecorDT.Rows.Add(item.ItemArray);

            SelectCommand = @"SELECT DecorOrders.DecorOrderID, DecorOrders.MainOrderID, DecorOrders.ProductID, DecorOrders.DecorID, DecorOrders.ColorID, DecorOrders.PatinaID, DecorOrders.Length, DecorOrders.Height, DecorOrders.Width, DecorOrders.Count, 
DecorOrders.DecorConfigID, DecorOrders.FactoryID, DecorOrders.Notes, DecorConfig.MeasureID
                FROM DecorOrders INNER JOIN 
                infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID
                WHERE MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE ProfilWorkAssignmentID=" + WorkAssignmentID + ")" + BatchFactoryFilter + ")" + FactoryFilter;
            if (FactoryID == 2)
                SelectCommand = @"SELECT DecorOrders.DecorOrderID, DecorOrders.MainOrderID, DecorOrders.ProductID, DecorOrders.DecorID, DecorOrders.ColorID, DecorOrders.PatinaID, DecorOrders.Length, DecorOrders.Height, DecorOrders.Width, DecorOrders.Count, 
DecorOrders.DecorConfigID, DecorOrders.FactoryID, DecorOrders.Notes, DecorConfig.MeasureID
                    FROM DecorOrders INNER JOIN 
                    infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID 
                    WHERE MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE TPSWorkAssignmentID=" + WorkAssignmentID + ")" + BatchFactoryFilter + ")" + FactoryFilter;
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.ZOVOrdersConnectionString))
            {
                DT.Clear();
                DA.Fill(DT);

                foreach (DataRow Row in DT.Rows)
                {
                    //if (Convert.ToInt32(Row["ColorID"]) == -1)
                    //    Row["ColorID"] = 0;
                }
            }
            foreach (DataRow item in DT.Rows)
                AllBatchDecorDT.Rows.Add(item.ItemArray);

            //AllBatchDecorDT.DefaultView.Sort = "ProductName, Name";
            return AllBatchDecorDT.Rows.Count > 0;
        }

        public bool GetDecorProducts(ref decimal TotalPogon, ref int TotalCount)
        {
            decimal DecorProductCount = 0;
            int decimals = 2;
            string Measure = string.Empty;

            DecorProductsSummaryDT.Clear();

            DataTable Table = new DataTable();

            using (DataView DV = new DataView(AllBatchDecorDT))
            {
                Table = DV.ToTable(true, new string[] { "ProductID", "MeasureID" });
            }

            for (int i = 0; i < Table.Rows.Count; i++)
            {
                DataRow[] Rows = AllBatchDecorDT.Select("ProductID=" + Convert.ToInt32(Table.Rows[i]["ProductID"]) +
                    " AND MeasureID=" + Convert.ToInt32(Table.Rows[i]["MeasureID"]));
                if (Rows.Count() == 0)
                    continue;

                foreach (DataRow row in Rows)
                {
                    if (Convert.ToInt32(row["MeasureID"]) == 1)
                    {
                        DecorProductCount += Convert.ToInt32(row["Count"]);
                        TotalCount += Convert.ToInt32(row["Count"]);
                        Measure = "шт.";
                    }

                    if (Convert.ToInt32(row["MeasureID"]) == 3)
                    {
                        DecorProductCount += Convert.ToInt32(row["Count"]);
                        TotalCount += Convert.ToInt32(row["Count"]);
                        Measure = "шт.";
                    }

                    if (Convert.ToInt32(row["MeasureID"]) == 2)
                    {
                        //нет параметра "высота"
                        if (row["Height"].ToString() == "-1")
                        {
                            DecorProductCount += Convert.ToDecimal(row["Length"]) * Convert.ToDecimal(row["Count"]) / 1000;
                            TotalPogon += Convert.ToDecimal(row["Length"]) * Convert.ToDecimal(row["Count"]) / 1000;
                        }
                        else
                        {
                            DecorProductCount += Convert.ToDecimal(row["Height"]) * Convert.ToDecimal(row["Count"]) / 1000;
                            TotalPogon += Convert.ToDecimal(row["Height"]) * Convert.ToDecimal(row["Count"]) / 1000;
                        }

                        Measure = "м.п.";
                    }
                }

                //NewRow["Product"] = GetProductName(Convert.ToInt32(Row["ProductID"])) + " " + GetDecorName(Convert.ToInt32(Row["DecorID"]));
                DataRow NewRow = DecorProductsSummaryDT.NewRow();
                NewRow["ProductID"] = Convert.ToInt32(Table.Rows[i]["ProductID"]);
                NewRow["MeasureID"] = Convert.ToInt32(Table.Rows[i]["MeasureID"]);
                NewRow["DecorProduct"] = GetProductName(Convert.ToInt32(Table.Rows[i]["ProductID"]));
                if (DecorProductCount < 3)
                    decimals = 1;
                NewRow["Count"] = decimal.Round(DecorProductCount, decimals, MidpointRounding.AwayFromZero);
                NewRow["Measure"] = Measure;
                DecorProductsSummaryDT.Rows.Add(NewRow);

                Measure = "";
                DecorProductCount = 0;
            }
            DecorProductsSummaryDT.DefaultView.Sort = "Count DESC";
            DecorProductsSummaryBS.MoveFirst();

            return DecorProductsSummaryDT.Rows.Count > 0;
        }

        public bool GetDecorItems()
        {
            decimal DecorItemCount = 0;
            int decimals = 2;
            string Measure = string.Empty;

            DecorItemsSummaryDT.Clear();

            DataTable Table = new DataTable();

            using (DataView DV = new DataView(AllBatchDecorDT))
            {
                Table = DV.ToTable(true, new string[] { "ProductID", "DecorID", "MeasureID" });
            }

            for (int i = 0; i < Table.Rows.Count; i++)
            {
                DataRow[] Rows = AllBatchDecorDT.Select("ProductID=" + Convert.ToInt32(Table.Rows[i]["ProductID"]) +
                    " AND DecorID=" + Convert.ToInt32(Table.Rows[i]["DecorID"]) + " AND MeasureID=" + Convert.ToInt32(Table.Rows[i]["MeasureID"]));
                if (Rows.Count() == 0)
                    continue;

                foreach (DataRow row in Rows)
                {
                    if (Convert.ToInt32(row["MeasureID"]) == 1)
                    {
                        DecorItemCount += Convert.ToInt32(row["Count"]);
                        Measure = "шт.";
                    }

                    if (Convert.ToInt32(row["MeasureID"]) == 3)
                    {
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
                NewRow["Count"] = decimal.Round(DecorItemCount, decimals, MidpointRounding.AwayFromZero);
                NewRow["Measure"] = Measure;
                DecorItemsSummaryDT.Rows.Add(NewRow);

                Measure = "";
                DecorItemCount = 0;
            }
            Table.Dispose();
            DecorItemsSummaryDT.DefaultView.Sort = "Count DESC";
            DecorItemsSummaryBS.MoveFirst();

            return DecorItemsSummaryDT.Rows.Count > 0;
        }

        public bool GetDecorColors()
        {
            decimal DecorColorCount = 0;
            int decimals = 2;
            string Measure = string.Empty;

            DecorColorsSummaryDT.Clear();

            DataTable Table = new DataTable();

            using (DataView DV = new DataView(AllBatchDecorDT))
            {
                Table = DV.ToTable(true, new string[] { "ProductID", "DecorID", "ColorID", "PatinaID", "MeasureID" });
            }

            for (int i = 0; i < Table.Rows.Count; i++)
            {
                DataRow[] Rows = AllBatchDecorDT.Select("ProductID=" + Convert.ToInt32(Table.Rows[i]["ProductID"]) +
                    " AND DecorID=" + Convert.ToInt32(Table.Rows[i]["DecorID"]) + " AND ColorID=" + Convert.ToInt32(Table.Rows[i]["ColorID"]) + " AND PatinaID=" + Convert.ToInt32(Table.Rows[i]["PatinaID"]) +
                    " AND MeasureID=" + Convert.ToInt32(Table.Rows[i]["MeasureID"]));
                if (Rows.Count() == 0)
                    continue;

                foreach (DataRow row in Rows)
                {
                    if (Convert.ToInt32(row["ProductID"]) == 2)
                    {
                        if (row["Height"].ToString() == "-1")
                            DecorColorCount += Convert.ToDecimal(row["Length"]) * Convert.ToDecimal(row["Count"]) / 1000;
                        else
                            DecorColorCount += Convert.ToDecimal(row["Height"]) * Convert.ToDecimal(row["Count"]) / 1000;
                        Measure = "м.п.";
                        continue;
                    }

                    if (Convert.ToInt32(row["MeasureID"]) == 1)
                    {
                        DecorColorCount += Convert.ToInt32(row["Count"]);
                        Measure = "шт.";
                    }

                    if (Convert.ToInt32(row["MeasureID"]) == 3)
                    {
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

                        Measure = "м.п.";
                    }
                }

                string Color = GetColorName(Convert.ToInt32(Table.Rows[i]["ColorID"]));
                if (Convert.ToInt32(Table.Rows[i]["PatinaID"]) != -1)
                    Color += " " + GetPatinaName(Convert.ToInt32(Table.Rows[i]["PatinaID"]));

                DataRow NewRow = DecorColorsSummaryDT.NewRow();
                NewRow["ProductID"] = Convert.ToInt32(Table.Rows[i]["ProductID"]);
                NewRow["DecorID"] = Convert.ToInt32(Table.Rows[i]["DecorID"]);
                NewRow["ColorID"] = Convert.ToInt32(Table.Rows[i]["ColorID"]);
                NewRow["PatinaID"] = Convert.ToInt32(Table.Rows[i]["PatinaID"]);
                NewRow["MeasureID"] = Convert.ToInt32(Table.Rows[i]["MeasureID"]);
                NewRow["Color"] = Color;
                if (DecorColorCount < 3)
                    decimals = 1;
                NewRow["Count"] = decimal.Round(DecorColorCount, decimals, MidpointRounding.AwayFromZero);
                NewRow["Measure"] = Measure;
                DecorColorsSummaryDT.Rows.Add(NewRow);

                Measure = string.Empty;
                DecorColorCount = 0;
            }
            Table.Dispose();
            DecorColorsSummaryDT.DefaultView.Sort = "Count DESC";
            DecorColorsSummaryBS.MoveFirst();

            return DecorItemsSummaryDT.Rows.Count > 0;
        }

        public bool GetDecorSizes()
        {
            decimal DecorSizeCount = 0;
            int decimals = 2;
            int Height = 0;
            int Length = 0;
            int Width = 0;
            string Measure = string.Empty;
            string Sizes = string.Empty;

            DecorSizesSummaryDT.Clear();

            DataTable Table = new DataTable();

            using (DataView DV = new DataView(AllBatchDecorDT))
            {
                Table = DV.ToTable(true, new string[] { "ProductID", "DecorID", "ColorID", "PatinaID", "MeasureID", "Length", "Height", "Width" });
            }

            for (int i = 0; i < Table.Rows.Count; i++)
            {
                DataRow[] Rows = AllBatchDecorDT.Select("ProductID=" + Convert.ToInt32(Table.Rows[i]["ProductID"]) +
                    " AND DecorID=" + Convert.ToInt32(Table.Rows[i]["DecorID"]) +
                    " AND ColorID=" + Convert.ToInt32(Table.Rows[i]["ColorID"]) + " AND PatinaID=" + Convert.ToInt32(Table.Rows[i]["PatinaID"]) +
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
                        if (row["Height"].ToString() == "-1")
                            DecorSizeCount += Convert.ToDecimal(row["Length"]) * Convert.ToDecimal(row["Count"]) / 1000;
                        else
                            DecorSizeCount += Convert.ToDecimal(row["Height"]) * Convert.ToDecimal(row["Count"]) / 1000;
                        Measure = "м.п.";
                        continue;
                    }

                    if (Convert.ToInt32(row["MeasureID"]) == 1)
                    {
                        DecorSizeCount += Convert.ToInt32(row["Count"]);
                        Measure = "шт.";
                    }

                    if (Convert.ToInt32(row["MeasureID"]) == 3)
                    {
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

                        Measure = "м.п.";
                    }
                }

                DataRow NewRow = DecorSizesSummaryDT.NewRow();
                NewRow["ProductID"] = Convert.ToInt32(Table.Rows[i]["ProductID"]);
                NewRow["DecorID"] = Convert.ToInt32(Table.Rows[i]["DecorID"]);
                NewRow["ColorID"] = Convert.ToInt32(Table.Rows[i]["ColorID"]);
                NewRow["PatinaID"] = Convert.ToInt32(Table.Rows[i]["PatinaID"]);
                NewRow["MeasureID"] = Convert.ToInt32(Table.Rows[i]["MeasureID"]);
                if (DecorSizeCount < 3)
                    decimals = 1;
                NewRow["Count"] = decimal.Round(DecorSizeCount, decimals, MidpointRounding.AwayFromZero);
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

                DecorSizesSummaryDT.Rows.Add(NewRow);
                NewRow["Size"] = Sizes;
                Sizes = string.Empty;
                Measure = string.Empty;
                DecorSizeCount = 0;
            }
            Table.Dispose();
            DecorSizesSummaryDT.DefaultView.Sort = "Count DESC";
            DecorSizesSummaryBS.MoveFirst();

            return DecorSizesSummaryDT.Rows.Count > 0;
        }

        public void FilterDecorItems(int ProductID, int MeasureID)
        {
            DecorItemsSummaryBS.Filter = "ProductID=" + ProductID + " AND MeasureID=" + MeasureID;
            DecorItemsSummaryBS.MoveFirst();
        }

        public void FilterDecorColors(int ProductID, int DecorID, int MeasureID)
        {
            DecorColorsSummaryBS.Filter = "ProductID=" + ProductID + " AND DecorID="
                + DecorID + " AND MeasureID=" + MeasureID;
            DecorColorsSummaryBS.MoveFirst();
        }

        public void FilterDecorSizes(int ProductID, int DecorID, int ColorID, int MeasureID)
        {
            DecorSizesSummaryBS.Filter = "ProductID=" + ProductID +
                " AND DecorID=" + DecorID + " AND ColorID=" + ColorID + " AND MeasureID=" + MeasureID;
            DecorSizesSummaryBS.MoveFirst();
        }
    }



    public class CreationAssignments
    {
        //ArrayList MarketBatchFrontsID;
        //ArrayList MarketMegaBatchFrontsID;
        //ArrayList ZOVBatchFrontsID;
        //ArrayList ZOVMegaBatchFrontsID;

        private DataTable MarketBatchesInAssignmentDT = null;
        private DataTable ZOVBatchesInAssignmentDT = null;
        private DataTable MachinesDT = null;
        private DataTable MarketBatchesDT = null;
        private DataTable MarketMegaBatchesDT = null;
        private DataTable ZOVBatchesDT = null;
        private DataTable ZOVMegaBatchesDT = null;
        private BindingSource MarketBatchesBS = null;
        private BindingSource MarketMegaBatchesBS = null;
        private BindingSource ZOVBatchesBS = null;
        private BindingSource ZOVMegaBatchesBS = null;

        public BindingSource MarketBatchesList => MarketBatchesBS;

        public BindingSource MarketMegaBatchesList => MarketMegaBatchesBS;

        public BindingSource ZOVBatchesList => ZOVBatchesBS;

        public BindingSource ZOVMegaBatchesList => ZOVMegaBatchesBS;

        public CreationAssignments()
        {

        }

        public void Initialize()
        {
            Create();
            Fill();
            Binding();
        }

        private void Create()
        {
            //MarketBatchFrontsID = new ArrayList();
            //MarketMegaBatchFrontsID = new ArrayList();
            //ZOVBatchFrontsID = new ArrayList();
            //ZOVMegaBatchFrontsID = new ArrayList();

            MarketBatchesInAssignmentDT = new DataTable();
            ZOVBatchesInAssignmentDT = new DataTable();

            MachinesDT = new DataTable();
            MachinesDT.Columns.Add(new DataColumn("ValueMember", Type.GetType("System.Int32")));
            MachinesDT.Columns.Add(new DataColumn("DisplayMember", Type.GetType("System.String")));

            MarketBatchesDT = new DataTable();
            MarketMegaBatchesDT = new DataTable();
            ZOVBatchesDT = new DataTable();
            ZOVMegaBatchesDT = new DataTable();

            MarketBatchesBS = new BindingSource();
            MarketMegaBatchesBS = new BindingSource();
            ZOVBatchesBS = new BindingSource();
            ZOVMegaBatchesBS = new BindingSource();
        }

        private void MachineNewRow(string ValueMember, string DisplayMember)
        {
            DataRow NewRow = MachinesDT.NewRow();
            NewRow["ValueMember"] = ValueMember;
            NewRow["DisplayMember"] = DisplayMember;
            MachinesDT.Rows.Add(NewRow);
        }

        private void Fill()
        {
            string SelectCommand = @"SELECT TOP 0 Batch.*, BatchDetails.FactoryID FROM Batch INNER JOIN BatchDetails ON Batch.BatchID = BatchDetails.BatchID
                INNER JOIN FrontsOrders ON BatchDetails.MainOrderID=FrontsOrders.MainOrderID";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(MarketBatchesInAssignmentDT);
                DA.Fill(ZOVBatchesInAssignmentDT);
            }
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT TOP 0 * FROM Batch",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(MarketBatchesDT);
                DA.Fill(ZOVBatchesDT);
            }
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT TOP 0 * FROM MegaBatch",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(MarketMegaBatchesDT);
                DA.Fill(ZOVMegaBatchesDT);
            }
        }

        private void Binding()
        {
            MarketBatchesBS.DataSource = MarketBatchesDT;
            MarketMegaBatchesBS.DataSource = MarketMegaBatchesDT;
            ZOVBatchesBS.DataSource = ZOVBatchesDT;
            ZOVMegaBatchesBS.DataSource = ZOVMegaBatchesDT;
        }

        public void FilterBatchesByMegaBatch(bool ZOV, int MegaBatchID, int FactoryID)
        {
            string OrdersConnectionString = string.Empty;
            if (ZOV)
                OrdersConnectionString = ConnectionStrings.ZOVOrdersConnectionString;
            else
                OrdersConnectionString = ConnectionStrings.MarketingOrdersConnectionString;
            string BatchFactoryFilter = string.Empty;
            string SelectCommand = string.Empty;

            if (FactoryID != 0)
                BatchFactoryFilter = " WHERE BatchDetails.FactoryID = " + FactoryID;

            SelectCommand = @"SELECT * FROM Batch WHERE BatchID IN
                (SELECT BatchID FROM BatchDetails" + BatchFactoryFilter + ") AND MegaBatchID = " + MegaBatchID + "  ORDER BY BatchID DESC";

            try
            {
                using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                    OrdersConnectionString))
                {
                    using (DataTable DT = new DataTable())
                    {
                        if (ZOV)
                        {
                            ZOVBatchesDT.Clear();
                            DA.Fill(ZOVBatchesDT);
                        }
                        else
                        {
                            MarketBatchesDT.Clear();
                            DA.Fill(MarketBatchesDT);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void FilterMegaBatchesByFactory(bool ZOV, int FactoryID)
        {
            string OrdersConnectionString = string.Empty;
            if (ZOV)
                OrdersConnectionString = ConnectionStrings.ZOVOrdersConnectionString;
            else
                OrdersConnectionString = ConnectionStrings.MarketingOrdersConnectionString;
            string FactoryFilter = string.Empty;
            string SelectCommand = string.Empty;

            if (FactoryID != 0)
                FactoryFilter = " WHERE BatchDetails.FactoryID = " + FactoryID;

            SelectCommand = @"SELECT * FROM MegaBatch
                WHERE MegaBatchID IN (SELECT MegaBatchID FROM Batch WHERE BatchID IN
                (SELECT BatchID FROM BatchDetails" + FactoryFilter + ")) ORDER BY MegaBatchID DESC";

            try
            {
                using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                    OrdersConnectionString))
                {
                    using (DataTable DT = new DataTable())
                    {
                        if (ZOV)
                        {
                            ZOVMegaBatchesDT.Clear();
                            DA.Fill(ZOVMegaBatchesDT);
                        }
                        else
                        {
                            MarketMegaBatchesDT.Clear();
                            DA.Fill(MarketMegaBatchesDT);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public bool IsMarketBatchInAssignment(int FactoryID, int BatchID)
        {
            DataRow[] rows = MarketBatchesInAssignmentDT.Select("BatchID=" + BatchID);
            if (rows.Count() == 0)
                return false;
            if (FactoryID == 1)
            {
                if (rows[0]["ProfilWorkAssignmentID"] != DBNull.Value)
                    return true;
                else
                    return false;
            }
            else
            {
                if (rows[0]["TPSWorkAssignmentID"] != DBNull.Value)
                    return true;
                else
                    return false;
            }
        }

        public bool IsZOVBatchInAssignment(int FactoryID, int BatchID)
        {
            DataRow[] rows = ZOVBatchesInAssignmentDT.Select("BatchID=" + BatchID);
            if (rows.Count() == 0)
                return false;
            if (FactoryID == 1)
            {
                if (rows[0]["ProfilWorkAssignmentID"] != DBNull.Value)
                    return true;
                else
                    return false;
            }
            else
            {
                if (rows[0]["TPSWorkAssignmentID"] != DBNull.Value)
                    return true;
                else
                    return false;
            }
        }

        public void GetMarketBatchesInAssignment(int FactoryID)
        {
            MarketBatchesInAssignmentDT.Clear();
            DataTable DT = MarketBatchesInAssignmentDT.Clone();
            string SelectCommand = @"SELECT DISTINCT Batch.*, BatchDetails.FactoryID FROM Batch INNER JOIN BatchDetails ON Batch.BatchID = BatchDetails.BatchID
                INNER JOIN FrontsOrders ON BatchDetails.MainOrderID=FrontsOrders.MainOrderID AND FrontsOrders.FactoryID=" + FactoryID +
                " WHERE BatchDetails.FactoryID=" + FactoryID;
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                DT.Clear();
                DA.Fill(DT);
            }
            foreach (DataRow item in DT.Rows)
                MarketBatchesInAssignmentDT.Rows.Add(item.ItemArray);
            SelectCommand = @"SELECT DISTINCT Batch.*, BatchDetails.FactoryID FROM Batch INNER JOIN BatchDetails ON Batch.BatchID = BatchDetails.BatchID
                INNER JOIN DecorOrders ON BatchDetails.MainOrderID=DecorOrders.MainOrderID AND DecorOrders.FactoryID=" + FactoryID +
                " WHERE BatchDetails.FactoryID=" + FactoryID;
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                DT.Clear();
                DA.Fill(DT);
            }
            foreach (DataRow item in DT.Rows)
                MarketBatchesInAssignmentDT.Rows.Add(item.ItemArray);
        }

        public void GetZOVBatchesInAssignment(int FactoryID)
        {
            ZOVBatchesInAssignmentDT.Clear();
            DataTable DT = ZOVBatchesInAssignmentDT.Clone();
            string SelectCommand = @"SELECT DISTINCT Batch.*, BatchDetails.FactoryID FROM Batch INNER JOIN BatchDetails ON Batch.BatchID = BatchDetails.BatchID
                INNER JOIN FrontsOrders ON BatchDetails.MainOrderID=FrontsOrders.MainOrderID AND FrontsOrders.FactoryID=" + FactoryID +
                " WHERE BatchDetails.FactoryID=" + FactoryID;
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.ZOVOrdersConnectionString))
            {
                DT.Clear();
                DA.Fill(DT);
            }
            foreach (DataRow item in DT.Rows)
                ZOVBatchesInAssignmentDT.Rows.Add(item.ItemArray);
            SelectCommand = @"SELECT DISTINCT Batch.*, BatchDetails.FactoryID FROM Batch INNER JOIN BatchDetails ON Batch.BatchID = BatchDetails.BatchID
                INNER JOIN DecorOrders ON BatchDetails.MainOrderID=DecorOrders.MainOrderID AND DecorOrders.FactoryID=" + FactoryID +
                " WHERE BatchDetails.FactoryID=" + FactoryID;
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.ZOVOrdersConnectionString))
            {
                DT.Clear();
                DA.Fill(DT);
            }
            foreach (DataRow item in DT.Rows)
                ZOVBatchesInAssignmentDT.Rows.Add(item.ItemArray);
        }

        public void IsMarketMegaBatchInAssignment(int MegaBatchID, int FactoryID, ref int Status)
        {
            int BatchAmount = 0;
            int BatchInAssignmentAmount = 0;
            DataRow[] rows = MarketBatchesInAssignmentDT.Select("MegaBatchID=" + MegaBatchID + " AND FactoryID=" + FactoryID);
            if (rows.Count() > 0)
            {
                BatchAmount = rows.Count();
                if (FactoryID == 1)
                {
                    DataRow[] rows1 = MarketBatchesInAssignmentDT.Select("ProfilWorkAssignmentID IS NOT NULL AND MegaBatchID=" + MegaBatchID + " AND FactoryID=" + FactoryID);
                    BatchInAssignmentAmount = rows1.Count();
                }
                if (FactoryID == 2)
                {
                    DataRow[] rows1 = MarketBatchesInAssignmentDT.Select("TPSWorkAssignmentID IS NOT NULL AND MegaBatchID=" + MegaBatchID + " AND FactoryID=" + FactoryID);
                    BatchInAssignmentAmount = rows1.Count();
                }
                if (BatchInAssignmentAmount == 0)
                    Status = 0;
                else
                {
                    if (BatchAmount == BatchInAssignmentAmount)
                        Status = 2;
                    else
                        Status = 1;
                }
            }
        }

        public void IsZOVMegaBatchInAssignment(int MegaBatchID, int FactoryID, ref int Status)
        {
            int BatchAmount = 0;
            int BatchInAssignmentAmount = 0;
            DataRow[] rows = ZOVBatchesInAssignmentDT.Select("MegaBatchID=" + MegaBatchID + " AND FactoryID=" + FactoryID);
            if (rows.Count() > 0)
            {
                BatchAmount = rows.Count();
                if (FactoryID == 1)
                {
                    DataRow[] rows1 = ZOVBatchesInAssignmentDT.Select("ProfilWorkAssignmentID IS NOT NULL AND MegaBatchID=" + MegaBatchID + " AND FactoryID=" + FactoryID);
                    BatchInAssignmentAmount = rows1.Count();
                }
                if (FactoryID == 2)
                {
                    DataRow[] rows1 = ZOVBatchesInAssignmentDT.Select("TPSWorkAssignmentID IS NOT NULL AND MegaBatchID=" + MegaBatchID + " AND FactoryID=" + FactoryID);
                    BatchInAssignmentAmount = rows1.Count();
                }
                if (BatchInAssignmentAmount == 0)
                    Status = 0;
                else
                {
                    if (BatchAmount == BatchInAssignmentAmount)
                        Status = 2;
                    else
                        Status = 1;
                }
            }
        }
    }



    public class ControlAssignments
    {
        private DataTable MachinesDT = null;
        private DataTable BatchesDT = null;
        private DataTable MegaBatchesDT = null;
        private DataTable UsersDT = null;
        private DataTable WorkAssignmentsDT = null;
        private BindingSource BatchesBS = null;
        private BindingSource MegaBatchesBS = null;
        private BindingSource WorkAssignmentsBS = null;

        public BindingSource BatchesList => BatchesBS;

        public BindingSource MegaBatchesList => MegaBatchesBS;

        public BindingSource WorkAssignmentsList => WorkAssignmentsBS;

        public ControlAssignments()
        {

        }

        public void Initialize()
        {
            Create();
            Fill();
            Binding();
        }

        private void Create()
        {
            MachinesDT = new DataTable();
            MachinesDT.Columns.Add(new DataColumn("ValueMember", Type.GetType("System.Int32")));
            MachinesDT.Columns.Add(new DataColumn("DisplayMember", Type.GetType("System.String")));

            BatchesDT = new DataTable();
            MegaBatchesDT = new DataTable();
            UsersDT = new DataTable();
            WorkAssignmentsDT = new DataTable();

            BatchesBS = new BindingSource();
            MegaBatchesBS = new BindingSource();
            WorkAssignmentsBS = new BindingSource();
        }

        private void MachineNewRow(string ValueMember, string DisplayMember)
        {
            DataRow NewRow = MachinesDT.NewRow();
            NewRow["ValueMember"] = ValueMember;
            NewRow["DisplayMember"] = DisplayMember;
            MachinesDT.Rows.Add(NewRow);
        }

        private void Fill()
        {
            MachineNewRow("0", "-");
            MachineNewRow("1", "Elme");
            MachineNewRow("2", "Balistrini");
            MachineNewRow("3", "Rapid");
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT TOP 0 * FROM Batch",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(BatchesDT);
            }
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT TOP 0 * FROM MegaBatch",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(MegaBatchesDT);
            }
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT UserID, Name, ShortName FROM Users",
                ConnectionStrings.UsersConnectionString))
            {
                DA.Fill(UsersDT);
            }
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT TOP 0 * FROM WorkAssignments ORDER BY WorkAssignmentID DESC",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(WorkAssignmentsDT);
            }
            MegaBatchesDT.Columns.Add(new DataColumn("Group", Type.GetType("System.String")));
            MegaBatchesDT.Columns.Add(new DataColumn("GroupType", Type.GetType("System.Int32")));
            BatchesDT.Columns.Add(new DataColumn("GroupType", Type.GetType("System.Int32")));
        }

        private void Binding()
        {
            BatchesBS.DataSource = BatchesDT;
            MegaBatchesBS.DataSource = MegaBatchesDT;
            WorkAssignmentsBS.DataSource = WorkAssignmentsDT;
        }

        public DataGridViewComboBoxColumn MachineColumn
        {
            get
            {
                DataGridViewComboBoxColumn Column = new DataGridViewComboBoxColumn()
                {
                    Name = "MachineColumn",
                    HeaderText = "Станок",
                    DataPropertyName = "Machine",
                    DataSource = new DataView(MachinesDT),
                    ValueMember = "ValueMember",
                    DisplayMember = "DisplayMember",
                    DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing,
                    SortMode = DataGridViewColumnSortMode.Automatic,
                    Width = 135,
                    AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                };
                return Column;
            }
        }

        public DataGridViewComboBoxColumn TPS45UserColumn
        {
            get
            {
                DataGridViewComboBoxColumn Column = new DataGridViewComboBoxColumn()
                {
                    Name = "TPS45UserColumn",
                    HeaderText = "Угол 45 ТПС\r\nраспечатал",
                    DataPropertyName = "TPS45UserID",
                    DataSource = new DataView(UsersDT),
                    ValueMember = "UserID",
                    DisplayMember = "ShortName",
                    DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing,
                    SortMode = DataGridViewColumnSortMode.Automatic,
                    AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells,
                    MinimumWidth = 55
                };
                return Column;
            }
        }

        public DataGridViewComboBoxColumn GenevaUserColumn
        {
            get
            {
                DataGridViewComboBoxColumn Column = new DataGridViewComboBoxColumn()
                {
                    Name = "GenevaUserColumn",
                    HeaderText = "   Женева\r\nраспечатал",
                    DataPropertyName = "GenevaUserID",
                    DataSource = new DataView(UsersDT),
                    ValueMember = "UserID",
                    DisplayMember = "ShortName",
                    DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing,
                    SortMode = DataGridViewColumnSortMode.Automatic,
                    AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells,
                    MinimumWidth = 55
                };
                return Column;
            }
        }

        public DataGridViewComboBoxColumn TafelUserColumn
        {
            get
            {
                DataGridViewComboBoxColumn Column = new DataGridViewComboBoxColumn()
                {
                    Name = "TafelUserColumn",
                    HeaderText = "  Тафель\r\nраспечатал",
                    DataPropertyName = "TafelUserID",
                    DataSource = new DataView(UsersDT),
                    ValueMember = "UserID",
                    DisplayMember = "ShortName",
                    DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing,
                    SortMode = DataGridViewColumnSortMode.Automatic,
                    AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells,
                    MinimumWidth = 55
                };
                return Column;
            }
        }

        public DataGridViewComboBoxColumn Profil90UserColumn
        {
            get
            {
                DataGridViewComboBoxColumn Column = new DataGridViewComboBoxColumn()
                {
                    Name = "Profil90UserColumn",
                    HeaderText = "   Угол 90\r\nраспечатал",
                    DataPropertyName = "Profil90UserID",
                    DataSource = new DataView(UsersDT),
                    ValueMember = "UserID",
                    DisplayMember = "ShortName",
                    DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing,
                    SortMode = DataGridViewColumnSortMode.Automatic,
                    AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells,
                    MinimumWidth = 55
                };
                return Column;
            }
        }

        public DataGridViewComboBoxColumn Profil45UserColumn
        {
            get
            {
                DataGridViewComboBoxColumn Column = new DataGridViewComboBoxColumn()
                {
                    Name = "Profil45UserColumn",
                    HeaderText = "   Угол 45\r\nраспечатал",
                    DataPropertyName = "Profil45UserID",
                    DataSource = new DataView(UsersDT),
                    ValueMember = "UserID",
                    DisplayMember = "ShortName",
                    DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing,
                    SortMode = DataGridViewColumnSortMode.Automatic,
                    AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells,
                    MinimumWidth = 55
                };
                return Column;
            }
        }

        public DataGridViewComboBoxColumn DominoUserColumn
        {
            get
            {
                DataGridViewComboBoxColumn Column = new DataGridViewComboBoxColumn()
                {
                    Name = "DominoUserColumn",
                    HeaderText = "   Домино\r\nраспечатал",
                    DataPropertyName = "DominoUserID",
                    DataSource = new DataView(UsersDT),
                    ValueMember = "UserID",
                    DisplayMember = "ShortName",
                    DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing,
                    SortMode = DataGridViewColumnSortMode.Automatic,
                    AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells,
                    MinimumWidth = 55
                };
                return Column;
            }
        }

        public DataGridViewComboBoxColumn RALUserColumn
        {
            get
            {
                DataGridViewComboBoxColumn Column = new DataGridViewComboBoxColumn()
                {
                    Name = "RALUserColumn",
                    HeaderText = "   RAL\r\nраспечатал",
                    DataPropertyName = "RALUserID",
                    DataSource = new DataView(UsersDT),
                    ValueMember = "UserID",
                    DisplayMember = "ShortName",
                    DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing,
                    SortMode = DataGridViewColumnSortMode.Automatic,
                    AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells,
                    MinimumWidth = 55
                };
                return Column;
            }
        }

        public void RefreshMegaBatches(int WorkAssignmentID, int FactoryID)
        {
            string SelectCommand = string.Empty;
            string WorkAssignment = "ProfilWorkAssignmentID=";
            DataTable DT = new DataTable();
            if (FactoryID == 2)
                WorkAssignment = "TPSWorkAssignmentID=";
            SelectCommand = @"SELECT * FROM MegaBatch WHERE MegaBatchID IN (SELECT MegaBatchID FROM Batch WHERE " + WorkAssignment + WorkAssignmentID + ")";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.ZOVOrdersConnectionString))
            {
                DA.Fill(DT);
            }
            MegaBatchesDT.Clear();
            foreach (DataRow item in DT.Rows)
            {
                DataRow NewRow = MegaBatchesDT.NewRow();
                NewRow.ItemArray = item.ItemArray;
                NewRow["Group"] = "З";
                NewRow["GroupType"] = 0;
                MegaBatchesDT.Rows.Add(NewRow);
            }

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                DT.Clear();
                DA.Fill(DT);
            }
            foreach (DataRow item in DT.Rows)
            {
                DataRow NewRow = MegaBatchesDT.NewRow();
                NewRow.ItemArray = item.ItemArray;
                NewRow["Group"] = "М";
                NewRow["GroupType"] = 1;
                MegaBatchesDT.Rows.Add(NewRow);
            }
        }

        public void RefreshBatches(int WorkAssignmentID, int FactoryID)
        {
            string SelectCommand = string.Empty;
            string WorkAssignment = "ProfilWorkAssignmentID=";
            DataTable DT = new DataTable();
            if (FactoryID == 2)
                WorkAssignment = "TPSWorkAssignmentID=";
            SelectCommand = @"SELECT * FROM Batch WHERE " + WorkAssignment + WorkAssignmentID;
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.ZOVOrdersConnectionString))
            {
                DA.Fill(DT);
            }
            BatchesDT.Clear();
            foreach (DataRow item in DT.Rows)
            {
                DataRow NewRow = BatchesDT.NewRow();
                NewRow.ItemArray = item.ItemArray;
                NewRow["GroupType"] = 0;
                BatchesDT.Rows.Add(NewRow);
            }

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                DT.Clear();
                DA.Fill(DT);
            }
            foreach (DataRow item in DT.Rows)
            {
                DataRow NewRow = BatchesDT.NewRow();
                NewRow.ItemArray = item.ItemArray;
                NewRow["GroupType"] = 1;
                BatchesDT.Rows.Add(NewRow);
            }
        }

        public void FilterBatchesByMegaBatch(int GroupType, int MegaBatchID)
        {
            BatchesBS.Filter = "GroupType=" + GroupType + " AND MegaBatchID=" + MegaBatchID;
        }

        public void SetName(int WorkAssignmentID, int FactoryID)
        {
            string SelectCommand = string.Empty;
            string WorkAssignment = "ProfilWorkAssignmentID=";
            //DataTable TempDT = new DataTable();

            if (FactoryID == 2)
                WorkAssignment = "TPSWorkAssignmentID=";
            SelectCommand = @"SELECT * FROM Batch WHERE " + WorkAssignment + WorkAssignmentID;
            string name = string.Empty;
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    DA.Fill(DT);
                    //using (DataView DV = new DataView(DT))
                    //{
                    //    TempDT = DV.ToTable(true, new string[] { "MegaBatchID" });
                    //}
                    //foreach (DataRow item in TempDT.Rows)
                    //{
                    //    name += "М(" + item["MegaBatchID"].ToString() + "-";
                    //    DataRow[] rows = DT.Select("MegaBatchID=" + item["MegaBatchID"]);
                    //    foreach (DataRow item1 in rows)
                    //        name += item1["BatchID"].ToString() + ",";
                    //    if (name.Length > 0)
                    //        name = name.Substring(0, name.Length - 1);
                    //    name += ")+";
                    //}
                    foreach (DataRow item in DT.Rows)
                        name += "М(" + item["MegaBatchID"].ToString() + "," + item["BatchID"].ToString() + ")+";
                }
            }
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.ZOVOrdersConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    DA.Fill(DT);
                    foreach (DataRow item in DT.Rows)
                        name += "З(" + item["MegaBatchID"].ToString() + "," + item["BatchID"].ToString() + ")+";
                }
            }
            if (name.Length > 0)
                name = name.Substring(0, name.Length - 1);
            DataRow[] Rows = WorkAssignmentsDT.Select("WorkAssignmentID = " + WorkAssignmentID);
            if (Rows.Count() > 0)
                Rows[0]["Name"] = name;
            //TempDT.Dispose();
        }

        public void CreateWorkAssignment(string Name, int FactoryID)
        {
            DataRow NewRow = WorkAssignmentsDT.NewRow();
            NewRow["CreationDateTime"] = Security.GetCurrentDate();
            NewRow["Name"] = Name;
            NewRow["FactoryID"] = FactoryID;
            WorkAssignmentsDT.Rows.Add(NewRow);
        }

        //public void SetPrintDateTime(int WorkAssignmentID)
        //{
        //    DataRow[] Rows = WorkAssignmentsDT.Select("WorkAssignmentID = " + WorkAssignmentID);
        //    if (Rows.Count() > 0)
        //        Rows[0]["PrintDateTime"] = Security.GetCurrentDate();
        //}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="Type"></param>
        /// <param name="WorkAssignmentID"></param>
        /// <param name="PrintingStatus">0 - not printed, 1 - partially printed, 2 - all printed</param>
        public void SetPrintDateTime(int Type, int WorkAssignmentID)
        {
            DataRow[] Rows = WorkAssignmentsDT.Select("WorkAssignmentID = " + WorkAssignmentID);
            if (Rows.Count() > 0)
            {
                if (Type == 1)
                {
                    Rows[0]["TPS45DateTime"] = Security.GetCurrentDate();
                    Rows[0]["TPS45UserID"] = Security.CurrentUserID;
                }
                if (Type == 2)
                {
                    Rows[0]["GenevaDateTime"] = Security.GetCurrentDate();
                    Rows[0]["GenevaUserID"] = Security.CurrentUserID;
                }
                if (Type == 3)
                {
                    Rows[0]["TafelDateTime"] = Security.GetCurrentDate();
                    Rows[0]["TafelUserID"] = Security.CurrentUserID;
                }
                if (Type == 4)
                {
                    Rows[0]["Profil90DateTime"] = Security.GetCurrentDate();
                    Rows[0]["Profil90UserID"] = Security.CurrentUserID;
                }
                if (Type == 5)
                {
                    Rows[0]["Profil45DateTime"] = Security.GetCurrentDate();
                    Rows[0]["Profil45UserID"] = Security.CurrentUserID;
                }
                if (Type == 6)
                {
                    Rows[0]["DominoDateTime"] = Security.GetCurrentDate();
                    Rows[0]["DominoUserID"] = Security.CurrentUserID;
                }
                if (Type == 7)
                {
                    Rows[0]["RALDateTime"] = Security.GetCurrentDate();
                    Rows[0]["RALUserID"] = Security.CurrentUserID;
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="Type"></param>
        /// <param name="WorkAssignmentID"></param>
        /// <param name="PrintingStatus">0 - not printed, 1 - partially printed, 2 - all printed</param>
        public void SetPrintingStatus(int WorkAssignmentID, int FactoryID)
        {
            int PrintingStatus = 0;
            DataRow[] Rows = WorkAssignmentsDT.Select("WorkAssignmentID = " + WorkAssignmentID);
            if (Rows.Count() > 0)
            {
                //if (FactoryID == 2)
                //{
                //    PrintingStatus = IsTPS45Printed(WorkAssignmentID);
                //    Rows[0]["TPS45PrintingStatus"] = PrintingStatus;
                //    PrintingStatus = IsGenevaPrinted(WorkAssignmentID);
                //    Rows[0]["GenevaPrintingStatus"] = PrintingStatus;
                //    PrintingStatus = IsTafelPrinted(WorkAssignmentID);
                //    Rows[0]["TafelPrintingStatus"] = PrintingStatus;
                //}
                if (FactoryID == 1)
                {
                    PrintingStatus = IsProfile90Printed(WorkAssignmentID);
                    Rows[0]["Profil90PrintingStatus"] = PrintingStatus;
                    PrintingStatus = IsProfile45Printed(WorkAssignmentID);
                    Rows[0]["Profil45PrintingStatus"] = PrintingStatus;
                    PrintingStatus = IsDominoPrinted(WorkAssignmentID);
                    Rows[0]["DominoPrintingStatus"] = PrintingStatus;
                }
            }
        }

        public bool IsBatchChanged(int GroupType, int BatchID)
        {
            DataRow[] rows = WorkAssignmentsDT.Select("GroupType = " + GroupType + " AND BatchID = " + BatchID);
            if (rows.Count() > 0 && Convert.ToBoolean(rows[0]["Changed"]))
                return true;
            return false;
        }

        public void RemoveBatchFromAssignment(bool ZOV, int BatchID, int WorkAssignmentID, int FactoryID)
        {
            string OrdersConnectionString = string.Empty;
            if (ZOV)
                OrdersConnectionString = ConnectionStrings.ZOVOrdersConnectionString;
            else
                OrdersConnectionString = ConnectionStrings.MarketingOrdersConnectionString;
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM Batch WHERE BatchID =" + BatchID, OrdersConnectionString))
            {
                using (SqlCommandBuilder CB = new SqlCommandBuilder(DA))
                {
                    using (DataTable DT = new DataTable())
                    {
                        if (DA.Fill(DT) > 0)
                        {
                            for (int i = 0; i < DT.Rows.Count; i++)
                            {
                                if (FactoryID == 1)
                                    DT.Rows[i]["ProfilWorkAssignmentID"] = DBNull.Value;
                                if (FactoryID == 2)
                                    DT.Rows[i]["TPSWorkAssignmentID"] = DBNull.Value;
                            }
                            DA.Update(DT);
                        }
                    }
                }
            }
        }

        public void RemoveWorkAssignment(int WorkAssignmentID, int FactoryID)
        {
            string SelectCommand = @"SELECT * FROM Batch WHERE ProfilWorkAssignmentID=" + WorkAssignmentID;
            if (FactoryID == 2)
                SelectCommand = @"SELECT * FROM Batch WHERE TPSWorkAssignmentID=" + WorkAssignmentID;

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.ZOVOrdersConnectionString))
            {
                using (SqlCommandBuilder CB = new SqlCommandBuilder(DA))
                {
                    using (DataTable DT = new DataTable())
                    {
                        if (DA.Fill(DT) > 0)
                        {
                            for (int i = 0; i < DT.Rows.Count; i++)
                            {
                                if (FactoryID == 1)
                                    DT.Rows[i]["ProfilWorkAssignmentID"] = DBNull.Value;
                                if (FactoryID == 2)
                                    DT.Rows[i]["TPSWorkAssignmentID"] = DBNull.Value;
                            }
                            DA.Update(DT);
                        }
                    }
                }
            }
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (SqlCommandBuilder CB = new SqlCommandBuilder(DA))
                {
                    using (DataTable DT = new DataTable())
                    {
                        if (DA.Fill(DT) > 0)
                        {
                            for (int i = 0; i < DT.Rows.Count; i++)
                            {
                                if (FactoryID == 1)
                                    DT.Rows[i]["ProfilWorkAssignmentID"] = DBNull.Value;
                                if (FactoryID == 2)
                                    DT.Rows[i]["TPSWorkAssignmentID"] = DBNull.Value;
                            }
                            DA.Update(DT);
                        }
                    }
                }
            }
            SelectCommand = @"SELECT * FROM WorkAssignments WHERE WorkAssignmentID=" + WorkAssignmentID;
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (SqlCommandBuilder CB = new SqlCommandBuilder(DA))
                {
                    using (DataTable DT = new DataTable())
                    {
                        if (DA.Fill(DT) > 0)
                        {
                            foreach (DataRow item in DT.Rows)
                                item.Delete();
                            DA.Update(DT);
                        }
                    }
                }
            }
        }

        public void SaveBatches(bool ZOV, int[] Batches, int WorkAssignmentID, int FactoryID)
        {
            string OrdersConnectionString = string.Empty;
            if (ZOV)
                OrdersConnectionString = ConnectionStrings.ZOVOrdersConnectionString;
            else
                OrdersConnectionString = ConnectionStrings.MarketingOrdersConnectionString;
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM Batch WHERE BatchID IN (" + string.Join(",", Batches) + ")", OrdersConnectionString))
            {
                using (SqlCommandBuilder CB = new SqlCommandBuilder(DA))
                {
                    using (DataTable DT = new DataTable())
                    {
                        if (DA.Fill(DT) > 0)
                        {
                            for (int i = 0; i < DT.Rows.Count; i++)
                            {
                                if (FactoryID == 1)
                                    DT.Rows[i]["ProfilWorkAssignmentID"] = WorkAssignmentID;
                                if (FactoryID == 2)
                                    DT.Rows[i]["TPSWorkAssignmentID"] = WorkAssignmentID;
                            }
                            DA.Update(DT);
                        }
                    }
                }
            }
        }

        public void SaveWorkAssignments()
        {
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM WorkAssignments",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (SqlCommandBuilder CB = new SqlCommandBuilder(DA))
                {
                    DA.Update(WorkAssignmentsDT);
                }
            }
        }

        public int UpdateWorkAssignments(int FactoryID, DateTime date1, DateTime date2)
        {
            //double t1 = 0;
            //double t2 = 0;
            //string Filter = " AND CAST(CreationDateTime AS date) >= '" + date1.ToString("yyyy-MM-dd") +
            //    " 00:00' AND CAST(CreationDateTime AS date) <= '" + date2.ToString("yyyy-MM-dd") + " 23:59'";
            //string MyQueryText = "SELECT * FROM WorkAssignments WHERE FactoryID = " + FactoryID + Filter + " ORDER BY WorkAssignmentID DESC";
            //System.Diagnostics.Stopwatch sw = new System.Diagnostics.Stopwatch();
            //sw.Start();


            //using (SqlDataAdapter da = new SqlDataAdapter(MyQueryText, ConnectionStrings.MarketingOrdersConnectionString))
            //{
            //    WorkAssignmentsDT.Clear();
            //    da.Fill(WorkAssignmentsDT);
            //}

            //sw.Stop();
            //t1 = sw.Elapsed.TotalMilliseconds;

            //WorkAssignmentsDT.Clear();
            //sw.Restart();

            //using (SqlConnection conn = new SqlConnection(ConnectionStrings.MarketingOrdersConnectionString))
            //{
            //    using (SqlCommand cmd = new SqlCommand(MyQueryText, conn))
            //    {
            //        // set CommandType, parameters and SqlDependency here if needed
            //        conn.Open();
            //        using (SqlDataReader reader = cmd.ExecuteReader())
            //        {
            //            WorkAssignmentsDT.Load(reader);
            //        }
            //    }
            //}

            //sw.Stop();
            //t2 = sw.Elapsed.TotalMilliseconds;


            string Filter = " AND CAST(CreationDateTime AS date) >= '" + date1.ToString("yyyy-MM-dd") +
                " 00:00' AND CAST(CreationDateTime AS date) <= '" + date2.ToString("yyyy-MM-dd") + " 23:59'";

            int MaxWorkAssignmentID = 0;
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM WorkAssignments WHERE FactoryID = " + FactoryID + Filter + " ORDER BY WorkAssignmentID DESC",
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                WorkAssignmentsDT.Clear();
                DA.Fill(WorkAssignmentsDT);
            }
            if (WorkAssignmentsDT.Rows.Count > 0 && WorkAssignmentsDT.Rows[0]["WorkAssignmentID"] != DBNull.Value)
                MaxWorkAssignmentID = Convert.ToInt32(WorkAssignmentsDT.Rows[0]["WorkAssignmentID"]);
            return MaxWorkAssignmentID;
        }

        public void CalculateSquare(int WorkAssignmentID, int FactoryID)
        {
            decimal Square = 0;
            string SelectCommand = @"SELECT SUM(Square) As Square FROM FrontsOrders
                WHERE FactoryID=" + FactoryID + @" AND MainOrderID IN
                (SELECT MainOrderID FROM BatchDetails WHERE FactoryID=" + FactoryID + " AND BatchID IN (SELECT BatchID FROM Batch WHERE ProfilWorkAssignmentID=" + WorkAssignmentID + "))";
            if (FactoryID == 2)
                SelectCommand = @"SELECT SUM(Square) As Square FROM FrontsOrders
                    WHERE FactoryID=" + FactoryID + @" AND MainOrderID IN
                    (SELECT MainOrderID FROM BatchDetails WHERE FactoryID=" + FactoryID + " AND BatchID IN (SELECT BatchID FROM Batch WHERE TPSWorkAssignmentID=" + WorkAssignmentID + "))";
            using (DataTable DT = new DataTable())
            {
                using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DA.Fill(DT);

                    if (DT.Rows.Count > 0 && DT.Rows[0]["Square"] != DBNull.Value)
                    {
                        Square += decimal.Round(Convert.ToDecimal(DT.Rows[0]["Square"]), 2, MidpointRounding.AwayFromZero);
                    }
                }
            }
            using (DataTable DT = new DataTable())
            {
                using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.ZOVOrdersConnectionString))
                {
                    DT.Clear();
                    DA.Fill(DT);

                    if (DT.Rows.Count > 0 && DT.Rows[0]["Square"] != DBNull.Value)
                    {
                        Square += decimal.Round(Convert.ToDecimal(DT.Rows[0]["Square"]), 2, MidpointRounding.AwayFromZero);
                    }
                }
            }
            DataRow[] Rows = WorkAssignmentsDT.Select("WorkAssignmentID = " + WorkAssignmentID);
            if (Rows.Count() > 0)
                Rows[0]["Square"] = Square;
        }

        public void MoveToWorkAssignment(int WorkAssignmentID)
        {
            WorkAssignmentsBS.Position = WorkAssignmentsBS.Find("WorkAssignmentID", WorkAssignmentID);
        }

        public void SetInProduction(int WorkAssignmentID, int FactoryID)
        {
            DateTime CurrentDate = Security.GetCurrentDate();
            string SelectCommand = @"SELECT MainOrderID, MegaOrderID, ProfilProductionDate, ProfilProductionUserID, ProfilProductionStatusID FROM MainOrders
                    WHERE MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE ProfilWorkAssignmentID=" + WorkAssignmentID + "))";

            if (FactoryID == 2)
                SelectCommand = @"SELECT MainOrderID, MegaOrderID, TPSProductionDate, TPSProductionUserID, TPSProductionStatusID FROM MainOrders
                    WHERE MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE TPSWorkAssignmentID=" + WorkAssignmentID + "))";

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (SqlCommandBuilder CB = new SqlCommandBuilder(DA))
                {
                    using (DataTable DT = new DataTable())
                    {
                        if (DA.Fill(DT) > 0)
                        {
                            int[] MegaOrders = new int[DT.Rows.Count];
                            for (int i = 0; i < DT.Rows.Count; i++)
                            {
                                MegaOrders[i] = Convert.ToInt32(DT.Rows[i]["MegaOrderID"]);
                                if (FactoryID == 1)
                                {
                                    DT.Rows[i]["ProfilProductionStatusID"] = 2;
                                    DT.Rows[i]["ProfilProductionDate"] = CurrentDate;
                                    DT.Rows[i]["ProfilProductionUserID"] = Security.CurrentUserID;
                                }

                                if (FactoryID == 2)
                                {
                                    DT.Rows[i]["TPSProductionStatusID"] = 2;
                                    DT.Rows[i]["TPSProductionDate"] = CurrentDate;
                                    DT.Rows[i]["TPSProductionUserID"] = Security.CurrentUserID;
                                }
                            }
                            DA.Update(DT);
                            MegaOrders = MegaOrders.Distinct<int>().ToArray<int>();
                            for (int i = 0; i < MegaOrders.Count(); i++)
                            {
                                CheckOrdersStatus.SetMegaOrderStatus(MegaOrders[i]);
                            }
                        }
                    }
                }
            }
            SelectCommand = @"SELECT MainOrderID, MegaOrderID, ProfilProductionDate, ProfilProductionUserID, ProfilProductionStatusID FROM NewMainOrders
                    WHERE MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE ProfilWorkAssignmentID=" + WorkAssignmentID + "))";

            if (FactoryID == 2)
                SelectCommand = @"SELECT MainOrderID, MegaOrderID, TPSProductionDate, TPSProductionUserID, TPSProductionStatusID FROM NewMainOrders
                    WHERE MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE TPSWorkAssignmentID=" + WorkAssignmentID + "))";

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (SqlCommandBuilder CB = new SqlCommandBuilder(DA))
                {
                    using (DataTable DT = new DataTable())
                    {
                        if (DA.Fill(DT) > 0)
                        {
                            int[] MegaOrders = new int[DT.Rows.Count];
                            for (int i = 0; i < DT.Rows.Count; i++)
                            {
                                MegaOrders[i] = Convert.ToInt32(DT.Rows[i]["MegaOrderID"]);
                                if (FactoryID == 1)
                                {
                                    DT.Rows[i]["ProfilProductionStatusID"] = 2;
                                    DT.Rows[i]["ProfilProductionDate"] = CurrentDate;
                                    DT.Rows[i]["ProfilProductionUserID"] = Security.CurrentUserID;
                                }

                                if (FactoryID == 2)
                                {
                                    DT.Rows[i]["TPSProductionStatusID"] = 2;
                                    DT.Rows[i]["TPSProductionDate"] = CurrentDate;
                                    DT.Rows[i]["TPSProductionUserID"] = Security.CurrentUserID;
                                }
                            }
                            DA.Update(DT);
                            MegaOrders = MegaOrders.Distinct<int>().ToArray<int>();
                            for (int i = 0; i < MegaOrders.Count(); i++)
                            {
                                CheckOrdersStatus.SetMegaOrderStatus(MegaOrders[i]);
                            }
                        }
                    }
                }
            }
            SelectCommand = @"SELECT MainOrderID, MegaOrderID, ProfilProductionDate, ProfilProductionUserID, ProfilProductionStatusID FROM MainOrders
                    WHERE MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE ProfilWorkAssignmentID=" + WorkAssignmentID + "))";

            if (FactoryID == 2)
                SelectCommand = @"SELECT MainOrderID, MegaOrderID, TPSProductionDate, TPSProductionUserID, TPSProductionStatusID FROM MainOrders
                    WHERE MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE TPSWorkAssignmentID=" + WorkAssignmentID + "))";

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.ZOVOrdersConnectionString))
            {
                using (SqlCommandBuilder CB = new SqlCommandBuilder(DA))
                {
                    using (DataTable DT = new DataTable())
                    {
                        if (DA.Fill(DT) > 0)
                        {
                            for (int i = 0; i < DT.Rows.Count; i++)
                            {
                                if (FactoryID == 1)
                                {
                                    DT.Rows[i]["ProfilProductionStatusID"] = 2;
                                    DT.Rows[i]["ProfilProductionDate"] = CurrentDate;
                                    DT.Rows[i]["ProfilProductionUserID"] = Security.CurrentUserID;
                                }

                                if (FactoryID == 2)
                                {
                                    DT.Rows[i]["TPSProductionStatusID"] = 2;
                                    DT.Rows[i]["TPSProductionDate"] = CurrentDate;
                                    DT.Rows[i]["TPSProductionUserID"] = Security.CurrentUserID;
                                }
                            }
                            DA.Update(DT);
                        }
                    }
                }
            }
        }

        public void PrintAssignment(int WorkAssignmentID, DataTable FrontsDT)
        {
            DateTime CurrentDate = Security.GetCurrentDate();
            string SelectCommand = @"SELECT * FROM AssignmentsInWork WHERE WorkAssignmentID=" + WorkAssignmentID;

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (SqlCommandBuilder CB = new SqlCommandBuilder(DA))
                {
                    using (DataTable DT = new DataTable())
                    {
                        DA.Fill(DT);
                        for (int i = 0; i < FrontsDT.Rows.Count; i++)
                        {
                            DataRow NewRow = DT.NewRow();
                            NewRow["WorkAssignmentID"] = WorkAssignmentID;
                            NewRow["PrintDateTime"] = CurrentDate;
                            NewRow["UserID"] = Security.CurrentUserID;
                            NewRow["FrontID"] = FrontsDT.Rows[i]["FrontID"];
                            DT.Rows.Add(NewRow);
                        }
                        DA.Update(DT);
                    }
                }
            }
        }

        public int IsProfile45Printed(int WorkAssignmentID)
        {
            int Count = 0;
            int PrintingStatus = 0;
            int[] FrontsID = new int[29];
            FrontsID[0] = Convert.ToInt32(Fronts.Antalia);
            FrontsID[1] = Convert.ToInt32(Fronts.Venecia);
            FrontsID[2] = Convert.ToInt32(Fronts.Leon);
            FrontsID[3] = Convert.ToInt32(Fronts.Limog);
            FrontsID[4] = Convert.ToInt32(Fronts.Luk);
            FrontsID[5] = Convert.ToInt32(Fronts.LukPVH);
            FrontsID[6] = Convert.ToInt32(Fronts.Milano);
            FrontsID[7] = Convert.ToInt32(Fronts.Praga);
            FrontsID[8] = Convert.ToInt32(Fronts.Sigma);
            FrontsID[9] = Convert.ToInt32(Fronts.Fat);
            FrontsID[10] = Convert.ToInt32(Fronts.ep216);
            FrontsID[11] = Convert.ToInt32(Fronts.TechnoN);
            FrontsID[12] = Convert.ToInt32(Fronts.ep206);
            FrontsID[13] = Convert.ToInt32(Fronts.ep066Marsel4);
            FrontsID[14] = Convert.ToInt32(Fronts.ep018Marsel1);
            FrontsID[15] = Convert.ToInt32(Fronts.ep043Shervud);
            FrontsID[16] = Convert.ToInt32(Fronts.epsh406Techno4);
            FrontsID[17] = Convert.ToInt32(Fronts.Bergamo);
            FrontsID[18] = Convert.ToInt32(Fronts.Boston);
            FrontsID[19] = Convert.ToInt32(Fronts.Nord95);
            FrontsID[20] = Convert.ToInt32(Fronts.epFox);
            FrontsID[21] = Convert.ToInt32(Fronts.ep110Jersy);
            FrontsID[22] = Convert.ToInt32(Fronts.ep041);
            FrontsID[23] = Convert.ToInt32(Fronts.ep071);
            FrontsID[24] = Convert.ToInt32(Fronts.Urban);
            FrontsID[25] = Convert.ToInt32(Fronts.Bruno);
            FrontsID[26] = Convert.ToInt32(Fronts.Alby);
            FrontsID[27] = Convert.ToInt32(Fronts.ep112);
            FrontsID[28] = Convert.ToInt32(Fronts.ep111);
            string SelectCommand = @"SELECT * FROM AssignmentsInWork WHERE WorkAssignmentID=" + WorkAssignmentID;
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            DataTable TempDT = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(DT1);
            }

            SelectCommand = @"SELECT DISTINCT FrontID FROM FrontsOrders
                INNER JOIN MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID
                INNER JOIN MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID
                INNER JOIN BatchDetails ON FrontsOrders.MainOrderID = BatchDetails.MainOrderID AND BatchDetails.FactoryID = 1
                INNER JOIN Batch ON BatchDetails.BatchID = Batch.BatchID AND Batch.ProfilWorkAssignmentID = " + WorkAssignmentID +
                @" WHERE FrontsOrders.FactoryID=1 AND FrontID IN (" + string.Join(",", FrontsID.OfType<int>().ToArray()) + ")";

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(DT2);
            }

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.ZOVOrdersConnectionString))
            {
                DA.Fill(TempDT);
            }
            foreach (DataRow item in TempDT.Rows)
                DT2.Rows.Add(item.ItemArray);
            using (DataView DV = new DataView(DT2.Copy()))
            {
                DT2.Clear();
                DT2 = DV.ToTable(true, new string[] { "FrontID" });
            }

            for (int i = 0; i < DT2.Rows.Count; i++)
            {
                DataRow[] rows = DT1.Select("FrontID=" + DT2.Rows[i]["FrontID"]);
                if (rows.Count() > 0)
                    Count++;
            }
            if (Count == 0)
                PrintingStatus = 0;
            if (Count < DT2.Rows.Count)
                PrintingStatus = 1;
            if (Count == DT2.Rows.Count)
                PrintingStatus = 2;
            if (DT2.Rows.Count == 0)
                PrintingStatus = -1;
            return PrintingStatus;
        }

        public int IsProfile90Printed(int WorkAssignmentID)
        {
            int Count = 0;
            int PrintingStatus = 0;
            int[] FrontsID = new int[17];
            FrontsID[0] = Convert.ToInt32(Fronts.Marsel1);
            FrontsID[1] = Convert.ToInt32(Fronts.Marsel3);
            FrontsID[2] = Convert.ToInt32(Fronts.Techno1);
            FrontsID[3] = Convert.ToInt32(Fronts.Techno2);
            FrontsID[4] = Convert.ToInt32(Fronts.Techno4);
            //FrontsID[5] = Convert.ToInt32(Fronts.Techno4Mega);
            FrontsID[5] = Convert.ToInt32(Fronts.Techno5);
            FrontsID[6] = Convert.ToInt32(Fronts.PR1);
            FrontsID[7] = Convert.ToInt32(Fronts.PR3);
            FrontsID[8] = Convert.ToInt32(Fronts.PRU8);
            FrontsID[9] = Convert.ToInt32(Fronts.Marsel4);
            FrontsID[10] = Convert.ToInt32(Fronts.Shervud);
            FrontsID[11] = Convert.ToInt32(Fronts.Marsel5);
            FrontsID[12] = Convert.ToInt32(Fronts.pFox);
            FrontsID[13] = Convert.ToInt32(Fronts.Jersy110);
            FrontsID[14] = Convert.ToInt32(Fronts.Porto);
            FrontsID[15] = Convert.ToInt32(Fronts.Monte);
            FrontsID[16] = Convert.ToInt32(Fronts.pFlorenc);
            string SelectCommand = @"SELECT * FROM AssignmentsInWork WHERE WorkAssignmentID=" + WorkAssignmentID;
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            DataTable TempDT = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(DT1);
            }

            SelectCommand = @"SELECT DISTINCT FrontID FROM FrontsOrders
                INNER JOIN MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID
                INNER JOIN MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID
                INNER JOIN BatchDetails ON FrontsOrders.MainOrderID = BatchDetails.MainOrderID AND BatchDetails.FactoryID = 1
                INNER JOIN Batch ON BatchDetails.BatchID = Batch.BatchID AND Batch.ProfilWorkAssignmentID = " + WorkAssignmentID +
                @" WHERE FrontsOrders.FactoryID=1 AND FrontID IN (" + string.Join(",", FrontsID.OfType<int>().ToArray()) + ")";

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(DT2);
            }

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.ZOVOrdersConnectionString))
            {
                DA.Fill(TempDT);
            }
            foreach (DataRow item in TempDT.Rows)
                DT2.Rows.Add(item.ItemArray);
            using (DataView DV = new DataView(DT2.Copy()))
            {
                DT2.Clear();
                DT2 = DV.ToTable(true, new string[] { "FrontID" });
            }

            for (int i = 0; i < DT2.Rows.Count; i++)
            {
                DataRow[] rows = DT1.Select("FrontID=" + DT2.Rows[i]["FrontID"]);
                if (rows.Count() > 0)
                    Count++;
            }
            if (Count == 0)
                PrintingStatus = 0;
            if (Count < DT2.Rows.Count)
                PrintingStatus = 1;
            if (Count == DT2.Rows.Count)
                PrintingStatus = 2;
            if (DT2.Rows.Count == 0)
                PrintingStatus = -1;
            return PrintingStatus;
        }

        public int IsDominoPrinted(int WorkAssignmentID)
        {
            int PrintingStatus = 0;
            //            int[] FrontsID = new int[1];
            //            FrontsID[0] = Convert.ToInt32(Fronts.Techno4Domino);
            //            string SelectCommand = @"SELECT * FROM AssignmentsInWork WHERE WorkAssignmentID=" + WorkAssignmentID;
            //            DataTable DT1 = new DataTable();
            //            DataTable DT2 = new DataTable();
            //            DataTable TempDT = new DataTable();
            //            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            //            {
            //                DA.Fill(DT1);
            //            }

            //            SelectCommand = @"SELECT DISTINCT FrontID FROM FrontsOrders
            //                INNER JOIN MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID
            //                INNER JOIN MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID
            //                INNER JOIN BatchDetails ON FrontsOrders.MainOrderID = BatchDetails.MainOrderID AND BatchDetails.FactoryID = 1
            //                INNER JOIN Batch ON BatchDetails.BatchID = Batch.BatchID AND Batch.ProfilWorkAssignmentID = " + WorkAssignmentID +
            //                @" WHERE FrontsOrders.FactoryID=1 AND FrontID IN (" + string.Join(",", FrontsID.OfType<int>().ToArray()) + ")";

            //            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
            //                ConnectionStrings.MarketingOrdersConnectionString))
            //            {
            //                DA.Fill(DT2);
            //            }

            //            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
            //                ConnectionStrings.ZOVOrdersConnectionString))
            //            {
            //                DA.Fill(TempDT);
            //            }
            //            foreach (DataRow item in TempDT.Rows)
            //                DT2.Rows.Add(item.ItemArray);
            //            using (DataView DV = new DataView(DT2.Copy()))
            //            {
            //                DT2.Clear();
            //                DT2 = DV.ToTable(true, new string[] { "FrontID" });
            //            }

            //            for (int i = 0; i < DT2.Rows.Count; i++)
            //            {
            //                DataRow[] rows = DT1.Select("FrontID=" + DT2.Rows[i]["FrontID"]);
            //                if (rows.Count() > 0)
            //                    Count++;
            //            }
            //            if (Count == 0)
            //                PrintingStatus = 0;
            //            if (Count < DT2.Rows.Count)
            //                PrintingStatus = 1;
            //            if (Count == DT2.Rows.Count)
            //                PrintingStatus = 2;
            //            if (DT2.Rows.Count == 0)
            //                PrintingStatus = -1;
            return PrintingStatus;
        }

        public int IsTPS45Printed(int WorkAssignmentID)
        {
            int Count = 0;
            int PrintingStatus = 0;
            int[] FrontsID = new int[15];
            FrontsID[0] = Convert.ToInt32(Fronts.KansasPat);
            FrontsID[1] = Convert.ToInt32(Fronts.Turin1);
            FrontsID[2] = Convert.ToInt32(Fronts.Turin3);
            FrontsID[3] = Convert.ToInt32(Fronts.LeonTPS);
            FrontsID[4] = Convert.ToInt32(Fronts.Sofia);
            FrontsID[5] = Convert.ToInt32(Fronts.InfinitiPat);
            FrontsID[6] = Convert.ToInt32(Fronts.Lorenzo);
            FrontsID[7] = Convert.ToInt32(Fronts.Dakota);
            FrontsID[8] = Convert.ToInt32(Fronts.DakotaPat);
            FrontsID[9] = Convert.ToInt32(Fronts.Turin1_1);
            FrontsID[10] = Convert.ToInt32(Fronts.Kansas);
            FrontsID[11] = Convert.ToInt32(Fronts.Infiniti);
            FrontsID[12] = Convert.ToInt32(Fronts.Elegant);
            FrontsID[13] = Convert.ToInt32(Fronts.ElegantPat);
            FrontsID[14] = Convert.ToInt32(Fronts.Patricia1);
            string SelectCommand = @"SELECT * FROM AssignmentsInWork WHERE WorkAssignmentID=" + WorkAssignmentID;
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            DataTable TempDT = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(DT1);
            }

            SelectCommand = @"SELECT DISTINCT FrontID FROM FrontsOrders
                INNER JOIN MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID
                INNER JOIN MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID
                INNER JOIN BatchDetails ON FrontsOrders.MainOrderID = BatchDetails.MainOrderID AND BatchDetails.FactoryID = 1
                INNER JOIN Batch ON BatchDetails.BatchID = Batch.BatchID AND Batch.ProfilWorkAssignmentID = " + WorkAssignmentID +
                @" WHERE FrontsOrders.FactoryID=1 AND FrontID IN (" + string.Join(",", FrontsID.OfType<int>().ToArray()) + ")";

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(DT2);
            }

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.ZOVOrdersConnectionString))
            {
                DA.Fill(TempDT);
            }
            foreach (DataRow item in TempDT.Rows)
                DT2.Rows.Add(item.ItemArray);
            using (DataView DV = new DataView(DT2.Copy()))
            {
                DT2.Clear();
                DT2 = DV.ToTable(true, new string[] { "FrontID" });
            }

            for (int i = 0; i < DT2.Rows.Count; i++)
            {
                DataRow[] rows = DT1.Select("FrontID=" + DT2.Rows[i]["FrontID"]);
                if (rows.Count() > 0)
                    Count++;
            }
            if (Count == 0)
                PrintingStatus = 0;
            if (Count < DT2.Rows.Count)
                PrintingStatus = 1;
            if (Count == DT2.Rows.Count)
                PrintingStatus = 2;
            if (DT2.Rows.Count == 0)
                PrintingStatus = -1;
            return PrintingStatus;
        }

        public int IsGenevaPrinted(int WorkAssignmentID)
        {
            int Count = 0;
            int PrintingStatus = 0;

            string SelectCommand = @"SELECT * FROM AssignmentsInWork WHERE WorkAssignmentID=" + WorkAssignmentID;
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            DataTable TempDT = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(DT1);
            }

            SelectCommand = @"SELECT DISTINCT FrontID FROM FrontsOrders
                INNER JOIN MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID
                INNER JOIN MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID
                INNER JOIN BatchDetails ON FrontsOrders.MainOrderID = BatchDetails.MainOrderID AND BatchDetails.FactoryID = 1
                INNER JOIN Batch ON BatchDetails.BatchID = Batch.BatchID AND Batch.ProfilWorkAssignmentID = " + WorkAssignmentID +
                @" WHERE FrontsOrders.FactoryID=1 AND FrontID IN (1975,1976,1977,1978,15760, 3737, 30501,30502,30503,30504,30505,30506,16269,30364,30366,30367,28945,3727,3728,3729,3730,3731,3732,3733,3734,3735,3736,3737,3739,3740,3741,3742,3743,3744,3745,3746,3747,3748,15108,29597,27914)";

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(DT2);
            }

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.ZOVOrdersConnectionString))
            {
                DA.Fill(TempDT);
            }
            foreach (DataRow item in TempDT.Rows)
                DT2.Rows.Add(item.ItemArray);
            using (DataView DV = new DataView(DT2.Copy()))
            {
                DT2.Clear();
                DT2 = DV.ToTable(true, new string[] { "FrontID" });
            }

            for (int i = 0; i < DT2.Rows.Count; i++)
            {
                DataRow[] rows = DT1.Select("FrontID=" + DT2.Rows[i]["FrontID"]);
                if (rows.Count() > 0)
                    Count++;
            }
            if (Count == 0)
                PrintingStatus = 0;
            if (Count < DT2.Rows.Count)
                PrintingStatus = 1;
            if (Count == DT2.Rows.Count)
                PrintingStatus = 2;
            if (DT2.Rows.Count == 0)
                PrintingStatus = -1;
            return PrintingStatus;
        }

        public int IsTafelPrinted(int WorkAssignmentID)
        {
            int Count = 0;
            int PrintingStatus = 0;
            int[] FrontsID = new int[3];
            FrontsID[0] = Convert.ToInt32(Fronts.Tafel3);
            FrontsID[1] = Convert.ToInt32(Fronts.Tafel2);
            FrontsID[2] = Convert.ToInt32(Fronts.Tafel3Gl);
            string SelectCommand = @"SELECT * FROM AssignmentsInWork WHERE WorkAssignmentID=" + WorkAssignmentID;
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            DataTable TempDT = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(DT1);
            }

            SelectCommand = @"SELECT DISTINCT FrontID FROM FrontsOrders
                INNER JOIN MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID
                INNER JOIN MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID
                INNER JOIN BatchDetails ON FrontsOrders.MainOrderID = BatchDetails.MainOrderID AND BatchDetails.FactoryID = 1
                INNER JOIN Batch ON BatchDetails.BatchID = Batch.BatchID AND Batch.ProfilWorkAssignmentID = " + WorkAssignmentID +
                @" WHERE FrontsOrders.FactoryID=1 AND FrontID IN (" + string.Join(",", FrontsID.OfType<int>().ToArray()) + ")";

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(DT2);
            }

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.ZOVOrdersConnectionString))
            {
                DA.Fill(TempDT);
            }
            foreach (DataRow item in TempDT.Rows)
                DT2.Rows.Add(item.ItemArray);
            using (DataView DV = new DataView(DT2.Copy()))
            {
                DT2.Clear();
                DT2 = DV.ToTable(true, new string[] { "FrontID" });
            }

            for (int i = 0; i < DT2.Rows.Count; i++)
            {
                DataRow[] rows = DT1.Select("FrontID=" + DT2.Rows[i]["FrontID"]);
                if (rows.Count() > 0)
                    Count++;
            }
            if (Count == 0)
                PrintingStatus = 0;
            if (Count < DT2.Rows.Count)
                PrintingStatus = 1;
            if (Count == DT2.Rows.Count)
                PrintingStatus = 2;
            if (DT2.Rows.Count == 0)
                PrintingStatus = -1;
            return PrintingStatus;
        }

        public bool IsM1(int WorkAssignmentID, int FactoryID)
        {
            string SelectCommand = @"SELECT ClientID FROM MegaOrders WHERE MegaOrderID IN (SELECT MegaOrderID FROM MainOrders
                    WHERE MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE ProfilWorkAssignmentID=" + WorkAssignmentID + ")))";

            if (FactoryID == 2)
                SelectCommand = @"SELECT ClientID FROM MegaOrders WHERE MegaOrderID IN (SELECT MegaOrderID FROM MainOrders
                    WHERE MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE TPSWorkAssignmentID=" + WorkAssignmentID + ")))";

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    if (DA.Fill(DT) > 0)
                    {
                        if (DT.Rows[0]["ClientID"] != DBNull.Value && Convert.ToInt32(DT.Rows[0]["ClientID"]) == 101)
                            return true;
                    }
                }
            }
            return false;
        }
    }


    public enum Fronts
    {
        Jersy110 = 29998,
        ep110Jersy = 29996,

        ep066Marsel4 = 28919,
        ep018Marsel1 = 28918,
        ep043Shervud = 28920,
        ep112 = 30007,
        Urban = 30008,
        Alby = 30038,
        Bruno = 30037,
        epsh406Techno4 = 28917,

        Dakota = 28921,
        DakotaPat = 28922,
        Bergamo = 28963,

        ep041 = 29719,
        ep071 = 29941,
        ep206 = 4311,
        ep216 = 15534,
        ep111 = 30215,

        Boston = 28968,
        Marsel5 = 28965,

        Porto = 30214,
        Monte = 30009,

        Antalia = 2355,
        Nord95 = 29844,
        epFox = 29845,
        Venecia = 2353,
        Leon = 2348,
        Limog = 3641,
        Luk = 2344,
        LukPVH = 3635,
        Milano = 3636,
        Praga = 2352,
        Sigma = 3637,
        Fat = 2346,
        TechnoN = 3638,

        Shervud = 28791,

        Techno1 = 3625,
        Techno2 = 3626,
        Techno4 = 3627,
        pFox = 29846,
        pFlorenc = 30448,
        Techno5 = 3628,
        Marsel1 = 3629,
        Marsel3 = 3630,
        Marsel4 = 15003,
        PR1 = 3631,
        PR2 = 3632,
        PR3 = 3623,
        PRU8 = 3624,
        Infiniti = 15032,
        InfinitiPat = 3634,

        Tafel1_19 = 16579,
        Tafel1Gl_19 = 16580,
        Tafel1_16 = 29277,
        Tafel1Gl_16 = 29278,
        Tafel1R1 = 16581,
        Tafel1R1Gl = 16582,
        Tafel1R2 = 16583,
        Tafel1R2Gl = 16584,
        Grand = 40574,
        GrandVg = 40575,

        Lorenzo = 15450,
        Elegant = 30006,
        ElegantPat = 30005,
        Patricia1 = 30379,
        Kansas = 3760,
        KansasPat = 3577,
        Sofia = 3415,
        Tafel2 = 3662,
        Tafel3 = 3663,
        Tafel3Gl = 3664,
        Turin1 = 3419,
        Turin1_1 = 4388,
        LeonTPS = 27920,
        Turin3 = 3633
    }

    public enum Machines
    {
        No_machine = 0,
        Balistrini = 1,
        ELME = 2,
        Rapid = 3
    }

    public enum FrontMargins
    {
        TechnoNWidth = 280,
        BergamoWidth = 150,
        ep041Width = 200,
        ep071Width = 170,
        ep206Width = 130,
        ep216Width = 130,
        ep111Width = 142,
        BostonWidth = 100,
        AntaliaWidth = 120,
        Nord95Width = 110,
        epFoxWidth = 110,
        VeneciaWidth = 150,
        LeonWidth = 120,
        LimogWidth = 128,
        epsh406Techno4Width = 82,
        ep066Marsel4Width = 118,
        ep110JersyWidth = 118,
        ep043ShervudWidth = 138,
        ep112Width = 98,
        UrbanWidth = 136,
        AlbyWidth = 110,
        BrunoWidth = 190,
        ep018Marsel1Width = 138,
        LukWidth = 120,
        MilanoWidth = 120,
        PragaWidth = 120,
        SigmaWidth = 120,
        FatWidth = 156,

        TechnoNMargin = 99,
        BergamoMargin = 139,
        ep041Margin = 178,
        ep071Margin = 159,
        ep206Margin = 102,
        ep216Margin = 108,
        ep111Margin = 110,
        BergamoNarrowMargin = 98,
        BostonMargin = 28,
        AntaliaMargin = 108,
        Nord95Margin = 32,
        epFoxMargin = 88,
        VeneciaMargin = 138,
        LeonMargin = 125,
        LimogMargin = 125,
        epsh406Techno4Margin = 82,
        ep066Marsel4Margin = 118,
        ep110JersyMargin = 118,
        ep043ShervudMargin = 138,
        ep112Margin = 98,
        UrbanMargin = 114,
        AlbyMargin = 89,
        BrunoMargin = 169,
        ep018Marsel1Margin = 138,
        LukMargin = 108,
        MilanoMargin = 108,
        PragaMargin = 108,
        SigmaMargin = 44,
        FatMargin = 144,

        Marsel1Height = 175,

        Marsel5Height = 180,
        PortoHeight = 159,
        MonteHeight = 119,
        Marsel3Height = 209,
        Marsel4Height = 139,
        Marsel4Height1 = 146,
        Jersy110Height = 139,
        Jersy110Height1 = 146,
        ShervudHeight = 110,
        pFoxHeight = 100,
        pFlorencHeight = 119,
        Techno1Height = 153,
        Techno2Height = 123,
        Techno4Height = 171,
        Techno4NarrowHeight = 101,
        Techno5Height = 209,

        Marsel1Width = 108,
        Marsel5Width = 148,
        PortoWidth = 130,
        MonteWidth = 88,
        Marsel3Width = 148,
        Marsel4Width = 108,
        Jersy110Width = 108,
        ShervudWidth = 100,
        pFoxWidth = 100,
        pFlorencWidth = 110,
        Techno1Width = 151,
        Techno2Width = 121,
        Techno4Width = 171,
        Techno4NarrowWidth = 101,
        Techno5Width = 201,

        Marsel1InsetHeight = 138,
        Marsel5InsetHeight = 159,
        PortoInsetHeight = 140,
        MonteInsetHeight = 98,
        Marsel3InsetHeight = 178,
        Marsel3InsetImpostHeight = 133,
        //изменил здесь, было Marsel4InsetHeight = 120
        Marsel4BoxInsetHeight = 118,
        Marsel4InsetHeight = 118,
        Marsel4InsetImpostHeight = 87,
        Jersy110BoxInsetHeight = 118,
        Jersy110InsetHeight = 118,
        ShervudInsetHeight = 88,
        pFoxInsetHeight = 88,
        pFlorencInsetHeight = 98,
        Techno1InsetHeight = 133,
        Techno2InsetHeight = 108,
        PR3InsetHeight = 106,
        Techno4InsetHeight = 152,
        Techno4NarrowInsetHeight = 84,
        Techno5InsetHeight = 182,
        LuxInsetHeight = 176,
        MegaInsetHeight = 176,

        Marsel1InsetWidth = 138,
        Marsel5InsetWidth = 159,
        PortoInsetWidth = 140,
        MonteInsetWidth = 98,
        Marsel3InsetWidth = 178,
        Marsel3InsetImpostWidth = 178,
        Marsel4InsetWidth = 118,
        Marsel4InsetImpostWidth = 118,
        Jersy110InsetWidth = 118,
        ShervudInsetWidth = 138,
        pFoxInsetWidth = 88,
        pFlorencInsetWidth = 98,
        Techno1InsetWidth = 133,
        Techno2InsetWidth = 102,
        Techno4InsetWidth = 152,
        Techno4NarrowInsetWidth = 84,
        Techno5InsetWidth = 182,
        LuxInsetWidth = 176,
        MegaInsetWidth = 176,

        LorenzoSimpleInsetHeight = 126,
        ElegantSimpleInsetHeight = 134,
        Patricia1SimpleInsetHeight = 134,
        KansasSimpleInsetHeight = 128,
        SofiaSimpleInsetHeight = 127,
        DakotaSimpleInsetHeight = 113,
        Turin1SimpleInsetHeight = 127,
        Turin3SimpleInsetHeight = 127,
        LeonSimpleInsetHeight = 111,
        InfinitiSimpleInsetHeight = 134,

        LorenzoBoxInsetHeight = 99,
        ElegantBoxInsetHeight = 99,
        Patricia1BoxInsetHeight = 99,
        KansasBoxInsetHeight = 89,
        SofiaBoxInsetHeight = 99,
        DakotaBoxInsetHeight = 99,
        Turin1BoxInsetHeight = 127,
        Turin3BoxInsetHeight = 127,
        LeonBoxInsetHeight = 111,
        InfinitiBoxInsetHeight = 98,

        KansasVitrinaInsetHeight = 128,
        SofiaVitrinaInsetHeight = 127,
        DakotaVitrinaInsetHeight = 113,
        Turin1VitrinaInsetHeight = 127,
        Turin3VitrinaInsetHeight = 127,
        LeonVitrinaInsetHeight = 111,
        InfinitiVitrinaInsetHeight = 134,

        LorenzoGridInsetHeight = 126,
        ElegantGridInsetHeight = 134,
        Patricia1GridInsetHeight = 134,
        KansasGridInsetHeight = 128,
        SofiaGridInsetHeight = 127,
        DakotaGridInsetHeight = 115,
        Turin1GridInsetHeight = 127,
        Turin3GridInsetHeight = 127,
        LeonGridInsetHeight = 111,
        InfinitiGridInsetHeight = 134,

        LorenzoSimpleInsetWidth = 126,
        ElegantSimpleInsetWidth = 134,
        Patricia1SimpleInsetWidth = 134,
        KansasSimpleInsetWidth = 128,
        SofiaSimpleInsetWidth = 127,
        DakotaSimpleInsetWidth = 113,
        Turin1SimpleInsetWidth = 127,
        Turin3SimpleInsetWidth = 127,
        LeonSimpleInsetWidth = 111,
        InfinitiSimpleInsetWidth = 134,

        LorenzoBoxInsetWidth = 99,
        ElegantBoxInsetWidth = 99,
        Patricia1BoxInsetWidth = 99,
        KansasBoxInsetWidth = 89,
        SofiaBoxInsetWidth = 99,
        DakotaBoxInsetWidth = 99,
        Turin1BoxInsetWidth = 127,
        Turin3BoxInsetWidth = 127,
        LeonBoxInsetWidth = 111,
        InfinitiBoxInsetWidth = 99,

        KansasVitrinaInsetWidth = 128,
        SofiaVitrinaInsetWidth = 127,
        DakotaVitrinaInsetWidth = 113,
        Turin1VitrinaInsetWidth = 127,
        Turin3VitrinaInsetWidth = 127,
        LeonVitrinaInsetWidth = 111,
        InfinitiVitrinaInsetWidth = 134,

        LorenzoGridInsetWidth = 126,
        ElegantGridInsetWidth = 134,
        Patricia1GridInsetWidth = 134,
        KansasGridInsetWidth = 128,
        SofiaGridInsetWidth = 127,
        DakotaGridInsetWidth = 115,
        Turin1GridInsetWidth = 127,
        Turin3GridInsetWidth = 127,
        LeonGridInsetWidth = 111,
        InfinitiGridInsetWidth = 134
    }

    public enum FrontMinSizes
    {
        //  А ещё минимальные размеры:

        //Marsel1MinHeight = 180,

        //  Marsel1MinWidth = 180,

        //  Marsel1InsetMinHeight = 22,

        //  Marsel1InsetMinWidth = 22,
        Marsel1MinHeight = 175,
        Marsel5MinHeight = 180,
        PortoMinHeight = 160,
        MonteMinHeight = 120,
        Marsel3MinHeight = 139,
        Marsel4MinHeight = 140,
        Jersy110MinHeight = 140,
        ShervudMinHeight = 152,
        pFoxMinHeight = 110,
        pFlorencMinHeight = 120,
        Techno1MinHeight = 153,
        Techno2MinHeight = 123,
        Techno4MinHeight = 171,
        Techno5MinHeight = 209,

        Marsel1MinWidth = 45,
        Marsel5MinWidth = 32,
        PortoMinWidth = 36,
        MonteMinWidth = 38,
        Marsel3MinWidth = 45,
        Marsel4MinWidth = 38,
        Jersy110MinWidth = 38,
        ShervudMinWidth = 10,
        pFoxMinWidth = 10,
        pFlorencMinWidth = 40,
        Techno1MinWidth = 10,
        Techno2MinWidth = 10,
        Techno4MinWidth = 10,
        Techno5MinWidth = 10,
        PR3MinWidth = 25,

        Marsel1InsetMinHeight = 38,
        Marsel5InsetMinHeight = 22,
        PortoInsetMinHeight = 20,
        MonteInsetMinHeight = 22,
        Marsel3InsetMinHeight = 29,
        PR1InsetMinHeight = 15,
        Marsel4InsetMinHeight = 29,
        Jersy110InsetMinHeight = 29,
        ShervudInsetMinHeight = 15,
        pFoxInsetMinHeight = 22,
        pFlorencInsetMinHeight = 22,
        Techno1InsetMinHeight = 20,
        Techno2InsetMinHeight = 15,
        Techno4InsetMinHeight = 15,
        Techno5InsetMinHeight = 26,

        Marsel1InsetMinWidth = 15,
        Marsel5InsetMinWidth = 22,
        PortoInsetMinWidth = 26,
        MonteInsetMinWidth = 28,
        Marsel3InsetMinWidth = 15,
        PR1InsetMinWidth = 29,
        Marsel4InsetMinWidth = 26,
        Jersy110InsetMinWidth = 26,
        ShervudInsetMinWidth = 22,
        pFoxInsetMinWidth = 22,
        pFlorencInsetMinWidth = 52,
        Techno1InsetMinWidth = 28,
        Techno2InsetMinWidth = 28,
        Techno4InsetMinWidth = 28,
        Techno5InsetMinWidth = 28
    }

    public class Barcode
    {
        private BarcodeLib.Barcode Barcod;
        private SolidBrush FontBrush;

        public enum BarcodeLength { Short, Medium, Long };

        public Barcode()
        {
            Barcod = new BarcodeLib.Barcode();

            FontBrush = new System.Drawing.SolidBrush(Color.Black);
        }

        public void DrawBarcodeText(BarcodeLength BarcodeLength, Graphics Graphics, string Text, int X, int Y)
        {
            int CharOffset = 0;
            int CharWidth = 0;
            float FontSize = 0;

            if (BarcodeLength == Barcode.BarcodeLength.Short)
            {
                CharWidth = 8;
                CharOffset = 7;
                FontSize = 8.0f;
            }
            if (BarcodeLength == Barcode.BarcodeLength.Medium)
            {
                CharWidth = 18;
                CharOffset = 5;
                FontSize = 12.0f;
            }
            if (BarcodeLength == Barcode.BarcodeLength.Long)
            {
                CharWidth = 26;
                CharOffset = 5;
                FontSize = 14.0f;
            }

            Font F = new Font("Arial", FontSize, FontStyle.Regular);

            for (int i = 0; i < Text.Length; i++)
            {
                Graphics.DrawString(Text[i].ToString(), F, FontBrush, i * CharWidth + CharOffset + X, Y + 2);
            }

            F.Dispose();
        }

        public Image GetBarcode(BarcodeLength BarcodeLength, int BarcodeHeight, string Text)
        {
            //set length and height
            if (BarcodeLength == Barcode.BarcodeLength.Short)
            {
                Barcod.Width = 101 + 12;
            }
            if (BarcodeLength == Barcode.BarcodeLength.Medium)
            {
                Barcod.Width = 202 + 12;
            }
            if (BarcodeLength == Barcode.BarcodeLength.Long)
            {
                Barcod.Width = 303 + 12;
            }

            Barcod.Height = BarcodeHeight;


            //create area
            Bitmap B = new Bitmap(Barcod.Width, BarcodeHeight);
            Graphics G = Graphics.FromImage(B);
            G.Clear(Color.White);


            //create barcode
            Image Bar = Barcod.Encode(BarcodeLib.TYPE.CODE128C, Text);
            G.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.NearestNeighbor;
            G.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.None;
            G.DrawImage(Bar, 0, 2);
            //DrawBarcodeText(Barcode.BarcodeLength.Short, G, Text, 0, 23);


            //create text
            G.TextRenderingHint = System.Drawing.Text.TextRenderingHint.SingleBitPerPixelGridFit;


            Bar.Dispose();
            G.Dispose();

            GC.Collect();

            return B;
        }
    }


    public class FrontsProdCapacity
    {
        private int TechStoreID, InsetTypeID, PatinaID, Height, Width = 0;
        private DataTable DataTable1;
        private DataTable ResultDT;
        private DataTable SummaryDT;
        private DataTable MaterialDT;
        private DataTable FixedMaterialDT;
        private DataTable FrontsOrdersDT;
        private DataTable FrontsConfigDT;
        private DataTable OperationsDetailDT;
        private DataTable StoreDetailDT;
        private DataTable SumOperationsDetailDT;
        private DataTable SumStoreDetailDT;
        private DataTable OperationsTermsDT;

        public BindingSource ResultBS;
        public BindingSource SummaryBS;
        public BindingSource MaterialBS;

        public FrontsProdCapacity(int iTechStoreID, int iInsetTypeID, int iPatinaID, int iHeight, int iWidth)
        {
            TechStoreID = iTechStoreID;
            InsetTypeID = iInsetTypeID;
            PatinaID = iPatinaID;
            Height = iHeight;
            Width = iWidth;
        }

        public void Initialize()
        {
            Create();

            string SelectCommand = @"SELECT TechCatalogOperationsDetail.*, TechCatalogOperationsGroups.TechStoreID, TechCatalogOperationsGroups.GroupName, 
                MachinesOperations.MachinesOperationName, MachinesOperations.Norm, MachinesOperations.PreparatoryNorm, MachinesOperations.MeasureID, 
                Machines.MachineID, Machines.MachineName, SubSectors.SubSectorID, SubSectors.SubSectorName, Sectors.SectorID, Sectors.SectorName FROM TechCatalogOperationsDetail
                INNER JOIN TechCatalogOperationsGroups ON TechCatalogOperationsDetail.TechCatalogOperationsGroupID=TechCatalogOperationsGroups.TechCatalogOperationsGroupID AND GroupNumber=1
                INNER JOIN MachinesOperations ON TechCatalogOperationsDetail.MachinesOperationID = MachinesOperations.MachinesOperationID
                INNER JOIN Machines ON MachinesOperations.MachineID = Machines.MachineID
                INNER JOIN SubSectors ON Machines.SubSectorID = SubSectors.SubSectorID
                INNER JOIN Sectors ON SubSectors.SectorID = Sectors.SectorID ORDER BY SerialNumber";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.CatalogConnectionString))
            {
                OperationsDetailDT.Clear();
                DA.Fill(OperationsDetailDT);
                SumOperationsDetailDT = OperationsDetailDT.Clone();
                SumOperationsDetailDT.Columns.Add(new DataColumn("NestedLevel", Type.GetType("System.Int32")));
                SumOperationsDetailDT.Columns.Add(new DataColumn("PrevTechCatalogOperationsDetailID", Type.GetType("System.Int32")));
                SumOperationsDetailDT.Columns.Add(new DataColumn("OperationsDetail", Type.GetType("System.String")));
            }
            SelectCommand = @"SELECT TechCatalogStoreDetail.*, TechStore.TechStoreName FROM TechCatalogStoreDetail
                INNER JOIN TechStore ON TechCatalogStoreDetail.TechStoreID = TechStore.TechStoreID ORDER BY GroupA, GroupB, GroupC";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.CatalogConnectionString))
            {
                StoreDetailDT.Clear();
                DA.Fill(StoreDetailDT);
                SumStoreDetailDT = StoreDetailDT.Clone();
                SumStoreDetailDT.Columns.Add(new DataColumn("NestedLevel", Type.GetType("System.Int32")));
            }
            SelectCommand = @"SELECT * FROM TechCatalogOperationsTerms";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.CatalogConnectionString))
            {
                OperationsTermsDT.Clear();
                DA.Fill(OperationsTermsDT);
            }
        }

        private void Create()
        {
            ResultDT = new DataTable();
            ResultDT.Columns.Add(new DataColumn("SerialNumber", Type.GetType("System.Int32")));
            ResultDT.Columns.Add(new DataColumn("GroupName", Type.GetType("System.String")));
            ResultDT.Columns.Add(new DataColumn("TechCatalogOperationsGroupID", Type.GetType("System.Int32")));
            ResultDT.Columns.Add(new DataColumn("PrevTechCatalogOperationsDetailID", Type.GetType("System.Int32")));
            ResultDT.Columns.Add(new DataColumn("TechCatalogOperationsDetailID", Type.GetType("System.Int32")));

            DataTable1 = new DataTable();
            DataTable1.Columns.Add(new DataColumn("TechCatalogOperationsGroupID", Type.GetType("System.Int32")));
            DataTable1.Columns.Add(new DataColumn("PrevTechCatalogOperationsDetailID", Type.GetType("System.Int32")));
            DataTable1.Columns.Add(new DataColumn("TechCatalogOperationsDetailID", Type.GetType("System.Int32")));
            DataTable1.Columns.Add(new DataColumn("TechStoreID", Type.GetType("System.Int32")));
            DataTable1.Columns.Add(new DataColumn("SerialNumber", Type.GetType("System.Int32")));
            DataTable1.Columns.Add(new DataColumn("GroupName", Type.GetType("System.String")));
            DataTable1.Columns.Add(new DataColumn("MachinesOperationName", Type.GetType("System.String")));
            DataTable1.Columns.Add(new DataColumn("TechStoreName", Type.GetType("System.String")));
            DataTable1.Columns.Add(new DataColumn("Cost", Type.GetType("System.Decimal")));
            DataTable1.Columns.Add(new DataColumn("Measure", Type.GetType("System.String")));
            DataTable1.Columns.Add(new DataColumn("check", Type.GetType("System.Boolean")));

            SummaryDT = new DataTable();
            SummaryDT.Columns.Add(new DataColumn("TechCatalogOperationsGroupID", Type.GetType("System.Int32")));
            SummaryDT.Columns.Add(new DataColumn("PrevTechCatalogOperationsDetailID", Type.GetType("System.Int32")));
            SummaryDT.Columns.Add(new DataColumn("TechCatalogOperationsDetailID", Type.GetType("System.Int32")));
            SummaryDT.Columns.Add(new DataColumn("TechStoreID", Type.GetType("System.Int32")));
            SummaryDT.Columns.Add(new DataColumn("SerialNumber", Type.GetType("System.Int32")));
            SummaryDT.Columns.Add(new DataColumn("GroupName", Type.GetType("System.String")));
            SummaryDT.Columns.Add(new DataColumn("MachinesOperationName", Type.GetType("System.String")));
            SummaryDT.Columns.Add(new DataColumn("TechStoreName", Type.GetType("System.String")));
            SummaryDT.Columns.Add(new DataColumn("Cost", Type.GetType("System.Decimal")));
            SummaryDT.Columns.Add(new DataColumn("Measure", Type.GetType("System.String")));
            SummaryDT.Columns.Add(new DataColumn("check", Type.GetType("System.Boolean")));

            MaterialDT = new DataTable();
            MaterialDT.Columns.Add(new DataColumn("TechCatalogStoreDetailID", Type.GetType("System.Int32")));
            MaterialDT.Columns.Add(new DataColumn("TechCatalogOperationsDetailID", Type.GetType("System.Int32")));
            MaterialDT.Columns.Add(new DataColumn("TechStoreID", Type.GetType("System.Int32")));
            MaterialDT.Columns.Add(new DataColumn("SectorName", Type.GetType("System.String")));
            MaterialDT.Columns.Add(new DataColumn("MachinesOperationName", Type.GetType("System.String")));
            MaterialDT.Columns.Add(new DataColumn("TechStoreName", Type.GetType("System.String")));
            MaterialDT.Columns.Add(new DataColumn("Cost", Type.GetType("System.Decimal")));
            MaterialDT.Columns.Add(new DataColumn("Measure", Type.GetType("System.String")));
            MaterialDT.Columns.Add(new DataColumn("PanelCounter", Type.GetType("System.Int32")));
            FixedMaterialDT = MaterialDT.Clone();

            FrontsOrdersDT = new DataTable();
            FrontsConfigDT = new DataTable();
            OperationsDetailDT = new DataTable();
            StoreDetailDT = new DataTable();
            OperationsTermsDT = new DataTable();

            ResultBS = new BindingSource()
            {
                DataSource = ResultDT
            };
            SummaryBS = new BindingSource()
            {
                DataSource = SummaryDT
            };
            MaterialBS = new BindingSource()
            {
                DataSource = MaterialDT
            };
        }

        public void DeleteMaterial(int PanelCounter)
        {
            DataRow[] rows = MaterialDT.Select("PanelCounter=" + PanelCounter);
            for (int i = rows.Count() - 1; i >= 0; i--)
                rows[i].Delete();
        }

        private int NestedLevel = 1;

        public void FF()
        {
            SumOperationsDetailDT.Clear();
            string SelectCommand = @"SELECT TechCatalogOperationsDetail.*, TechCatalogOperationsGroups.TechStoreID, TechCatalogOperationsGroups.GroupName, 
                MachinesOperations.MachinesOperationName, MachinesOperations.Norm, MachinesOperations.PreparatoryNorm, MachinesOperations.MeasureID, 
                Machines.MachineID, Machines.MachineName, SubSectors.SubSectorID, SubSectors.SubSectorName, Sectors.SectorID, Sectors.SectorName FROM TechCatalogOperationsDetail
                INNER JOIN TechCatalogOperationsGroups ON TechCatalogOperationsDetail.TechCatalogOperationsGroupID=TechCatalogOperationsGroups.TechCatalogOperationsGroupID AND GroupNumber=1
                INNER JOIN MachinesOperations ON TechCatalogOperationsDetail.MachinesOperationID = MachinesOperations.MachinesOperationID
                INNER JOIN Machines ON MachinesOperations.MachineID = Machines.MachineID
                INNER JOIN SubSectors ON Machines.SubSectorID = SubSectors.SubSectorID
                INNER JOIN Sectors ON SubSectors.SectorID = Sectors.SectorID
                WHERE TechStoreID=" + TechStoreID + " ORDER BY SerialNumber";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.CatalogConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    DA.Fill(DT);
                    for (int i = 0; i < DT.Rows.Count; i++)
                    {
                        DataRow NewRow = SumOperationsDetailDT.NewRow();
                        NewRow.ItemArray = DT.Rows[i].ItemArray;
                        NewRow["NestedLevel"] = NestedLevel;
                        SumOperationsDetailDT.Rows.Add(NewRow);
                    }
                }
            }

            for (int i = 0; i < SumOperationsDetailDT.Rows.Count; i++)
            {
                if (NestedLevel == 0)
                    break;
                //if (NestedLevel != Convert.ToInt32(SumOperationsDetailDT.Rows[i]["NestedLevel"]))
                //    continue;
                int TechCatalogOperationsDetailID = Convert.ToInt32(SumOperationsDetailDT.Rows[i]["TechCatalogOperationsDetailID"]);
                if (!CheckConditions(TechCatalogOperationsDetailID, InsetTypeID, PatinaID, Height, Width))
                    continue;
                GetStoreDetail(TechCatalogOperationsDetailID, Convert.ToInt32(SumOperationsDetailDT.Rows[i]["NestedLevel"]), SumOperationsDetailDT.Rows[i]);
            }
        }

        public bool GetOperationsDetail(int PrevTechCatalogOperationsDetailID, int NestedLevel, DataRow Row)
        {
            int TechStoreID = Convert.ToInt32(Row["TechStoreID"]);

            string TechStoreName = Row["TechStoreName"].ToString();
            bool BreakChain = Convert.ToBoolean(Row["BreakChain"]);
            if (BreakChain)
                return false;
            DataRow[] rows = OperationsDetailDT.Select("TechStoreID=" + Convert.ToInt32(Row["TechStoreID"]));
            if (rows.Count() == 0)
            {
            }
            else
            {
                foreach (DataRow item in rows)
                {
                    int TechCatalogOperationsDetailID = Convert.ToInt32(item["TechCatalogOperationsDetailID"]);
                    if (!CheckConditions(Convert.ToInt32(item["TechCatalogOperationsDetailID"]), InsetTypeID, PatinaID, Height, Width))
                        continue;
                    //if (SumOperationsDetailDT.Select("TechCatalogOperationsDetailID=" + TechCatalogOperationsDetailID).Count() > 0)
                    //    continue;
                    if (TechCatalogOperationsDetailID == PrevTechCatalogOperationsDetailID)
                        continue;
                    DataRow NewRow = SumOperationsDetailDT.NewRow();
                    NewRow.ItemArray = item.ItemArray;
                    NewRow["PrevTechCatalogOperationsDetailID"] = PrevTechCatalogOperationsDetailID;
                    NewRow["NestedLevel"] = NestedLevel;
                    SumOperationsDetailDT.Rows.Add(NewRow);
                }
            }
            NestedLevel--;

            return rows.Count() > 0;
        }

        public bool GetStoreDetail(int PrevTechCatalogOperationsDetailID, int NestedLevel, DataRow Row)
        {
            int TechCatalogOperationsDetailID = Convert.ToInt32(Row["TechCatalogOperationsDetailID"]);
            DataRow[] rows = StoreDetailDT.Select("TechCatalogOperationsDetailID=" + Convert.ToInt32(Row["TechCatalogOperationsDetailID"]));
            if (rows.Count() == 0)
            {
                //if (SumOperationsDetailDT.Select("TechCatalogOperationsDetailID=" + TechCatalogOperationsDetailID).Count() > 0)
                //    return false;
                if (TechCatalogOperationsDetailID == PrevTechCatalogOperationsDetailID)
                    return false;
                DataRow NewRow = SumOperationsDetailDT.NewRow();
                NewRow.ItemArray = Row.ItemArray;
                NewRow["PrevTechCatalogOperationsDetailID"] = PrevTechCatalogOperationsDetailID;
                NewRow["NestedLevel"] = NestedLevel;
                SumOperationsDetailDT.Rows.Add(NewRow);
            }
            else
            {
                NestedLevel++;
                int iNedtedLevel = NestedLevel;
                foreach (DataRow item in rows)
                {
                    DataRow NewRow = SumStoreDetailDT.NewRow();
                    NewRow.ItemArray = item.ItemArray;
                    NewRow["NestedLevel"] = iNedtedLevel - 1;
                    SumStoreDetailDT.Rows.Add(NewRow);
                    GetOperationsDetail(PrevTechCatalogOperationsDetailID, iNedtedLevel, item);
                }
            }
            NestedLevel--;

            return rows.Count() > 0;
        }

        private bool CheckConditions(int TechCatalogOperationsDetailID, int InsetTypeID, int PatinaID, int Height, int Width)
        {
            DataRow[] rows = OperationsTermsDT.Select("TechCatalogOperationsDetailID=" + TechCatalogOperationsDetailID);
            if (rows.Count() == 0)
            {
                return true;
            }
            foreach (DataRow row in rows)
            {
                int Term = Convert.ToInt32(row["Term"]);
                string Parameter = row["Parameter"].ToString();
                switch (Parameter)
                {
                    case "CoverID":
                        break;
                    case "InsetTypeID":
                        if (row["MathSymbol"].ToString().Equals("="))
                        {
                            if (InsetTypeID == Term)
                            { }
                            else
                                return false;
                        }
                        if (row["MathSymbol"].ToString().Equals("!="))
                        {
                            if (InsetTypeID != Term)
                            { }
                            else
                                return false;
                        }
                        if (row["MathSymbol"].ToString().Equals(">="))
                        {
                            if (InsetTypeID >= Term)
                            { }
                            else
                                return false;
                        }
                        if (row["MathSymbol"].ToString().Equals("<="))
                        {
                            if (InsetTypeID <= Term)
                            { }
                            else
                                return false;
                        }
                        break;
                    case "InsetColorID":
                        break;
                    case "ColorID":
                        break;
                    case "PatinaID":
                        if (row["MathSymbol"].ToString().Equals("="))
                        {
                            if (PatinaID == Term)
                            { }
                            else
                                return false;
                        }
                        if (row["MathSymbol"].ToString().Equals("!="))
                        {
                            if (PatinaID != Term)
                            { }
                            else
                                return false;
                        }
                        if (row["MathSymbol"].ToString().Equals(">="))
                        {
                            if (PatinaID >= Term)
                            { }
                            else
                                return false;
                        }
                        if (row["MathSymbol"].ToString().Equals("<="))
                        {
                            if (PatinaID <= Term)
                            { }
                            else
                                return false;
                        }
                        break;
                    case "Diameter":
                        break;
                    case "Thickness":
                        break;
                    case "Length":
                        break;
                    case "Height":
                        if (row["MathSymbol"].ToString().Equals("="))
                        {
                            if (Height == Term)
                            { }
                            else
                                return false;
                        }
                        if (row["MathSymbol"].ToString().Equals("!="))
                        {
                            if (Height != Term)
                            { }
                            else
                                return false;
                        }
                        if (row["MathSymbol"].ToString().Equals(">="))
                        {
                            if (Height >= Term)
                            { }
                            else
                                return false;
                        }
                        if (row["MathSymbol"].ToString().Equals("<="))
                        {
                            if (Height <= Term)
                            { }
                            else
                                return false;
                        }
                        break;
                    case "Width":
                        if (row["MathSymbol"].ToString().Equals("="))
                        {
                            if (Width == Term)
                            { }
                            else
                                return false;
                        }
                        if (row["MathSymbol"].ToString().Equals("!="))
                        {
                            if (Width != Term)
                            { }
                            else
                                return false;
                        }
                        if (row["MathSymbol"].ToString().Equals(">="))
                        {
                            if (Width >= Term)
                            { }
                            else
                                return false;
                        }
                        if (row["MathSymbol"].ToString().Equals("<="))
                        {
                            if (Width <= Term)
                            { }
                            else
                                return false;
                        }
                        break;
                    case "Admission":
                        break;
                    case "InsetHeightAdmission":
                        break;
                    case "InsetWidthAdmission":
                        break;
                    case "Capacity":
                        break;
                    case "Weight":
                        break;
                }
            }
            return true;
        }
    }
}
