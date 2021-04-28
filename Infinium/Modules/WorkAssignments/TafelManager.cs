using System;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Windows.Forms;

namespace Infinium.Modules.WorkAssignments
{
    public class TafelManager
    {
        public FrontsOrdersManager FrontsOrdersManager;
        private BindingSource FrontOrdersBS;
        private DataTable FrontOrdersDT;
        private BindingSource MainOrdersBS;

        private DataTable MainOrdersDT;

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

        public int CurrentFrontsCount
        {
            get
            {
                if (FrontOrdersBS.Count == 0 || ((DataRowView)FrontOrdersBS.Current).Row["Count"] == DBNull.Value)
                    return 0;
                return Convert.ToInt32(((DataRowView)FrontOrdersBS.Current).Row["Count"]);
            }
        }

        public DataTable EditFrontOrdersDT => FrontOrdersDT;

        public BindingSource FrontOrdersList => FrontOrdersBS;

        public BindingSource MainOrdersList => MainOrdersBS;

        public void AddNewRow(int Count, decimal Square)
        {
            DataRow NewRow = FrontOrdersDT.NewRow();
            NewRow.ItemArray = ((DataRowView)FrontOrdersBS.Current).Row.ItemArray;
            NewRow["Count"] = Count;
            NewRow["Square"] = Square * Count;
            FrontOrdersDT.Rows.Add(NewRow);
        }

        public void EditCurrentFrontsCout(int NewCount, ref decimal Square)
        {
            int OldCount = Convert.ToInt32(((DataRowView)FrontOrdersBS.Current).Row["Count"]);
            Square = Convert.ToDecimal(((DataRowView)FrontOrdersBS.Current).Row["Square"]) / OldCount;
            ((DataRowView)FrontOrdersBS.Current).Row["Count"] = NewCount;
            ((DataRowView)FrontOrdersBS.Current).Row["Square"] = Square * NewCount;
        }

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

        public void FilterOrdersByMainOrder(int GroupType, int MainOrderID)
        {
            FrontOrdersBS.Filter = "GroupType=" + GroupType + " AND MainOrderID=" + MainOrderID;
        }

        public void MoveToPosition(int Position)
        {
            FrontOrdersBS.Position = Position;
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

        public class SplitStruct
        {
            public bool bEqual;
            public bool bOk;
            public int FirstCount;
            public int PosCount;
            public int SecondCount;
            public int SourceCount;

            public SplitStruct()
            {
                PosCount = 0;
                FirstCount = 0;
                SecondCount = 0;
                bOk = false;
                bEqual = false;
            }
        }
    }
}