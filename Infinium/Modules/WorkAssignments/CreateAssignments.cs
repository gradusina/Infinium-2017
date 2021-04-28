using System;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Windows.Forms;

namespace Infinium.Modules.WorkAssignments
{
    public class CreateAssignments
    {
        private DataTable MachinesDT = null;

        private BindingSource MarketBatchesBS = null;

        private DataTable MarketBatchesDT = null;

        private DataTable MarketBatchesInAssignmentDT = null;

        private BindingSource MarketMegaBatchesBS = null;

        private DataTable MarketMegaBatchesDT = null;

        private BindingSource ZOVBatchesBS = null;

        private DataTable ZOVBatchesDT = null;

        private DataTable ZOVBatchesInAssignmentDT = null;

        private BindingSource ZOVMegaBatchesBS = null;

        private DataTable ZOVMegaBatchesDT = null;

        public CreateAssignments()
        {
        }

        public BindingSource MarketBatchesList => MarketBatchesBS;

        public BindingSource MarketMegaBatchesList => MarketMegaBatchesBS;

        public BindingSource ZOVBatchesList => ZOVBatchesBS;

        public BindingSource ZOVMegaBatchesList => ZOVMegaBatchesBS;

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

        public void Initialize()
        {
            Create();
            Fill();
            Binding();
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

        private void Binding()
        {
            MarketBatchesBS.DataSource = MarketBatchesDT;
            MarketMegaBatchesBS.DataSource = MarketMegaBatchesDT;
            ZOVBatchesBS.DataSource = ZOVBatchesDT;
            ZOVMegaBatchesBS.DataSource = ZOVMegaBatchesDT;
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

        private void MachineNewRow(string ValueMember, string DisplayMember)
        {
            DataRow NewRow = MachinesDT.NewRow();
            NewRow["ValueMember"] = ValueMember;
            NewRow["DisplayMember"] = DisplayMember;
            MachinesDT.Rows.Add(NewRow);
        }
    }
}