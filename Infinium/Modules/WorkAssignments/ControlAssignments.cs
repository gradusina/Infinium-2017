using System;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Windows.Forms;

namespace Infinium.Modules.WorkAssignments
{
    public class ControlAssignments
    {
        private BindingSource BatchesBS = null;

        private DataTable BatchesDT = null;

        private DataTable MachinesDT = null;

        private BindingSource MegaBatchesBS = null;

        private DataTable MegaBatchesDT = null;

        private DataTable UsersDT = null;

        private BindingSource WorkAssignmentsBS = null;

        private DataTable WorkAssignmentsDT = null;

        public ControlAssignments()
        {
        }

        public BindingSource BatchesList => BatchesBS;

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

        public BindingSource MegaBatchesList => MegaBatchesBS;

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

        public BindingSource WorkAssignmentsList => WorkAssignmentsBS;

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

        public void CreateWorkAssignment(string Name, int FactoryID)
        {
            DataRow NewRow = WorkAssignmentsDT.NewRow();
            NewRow["CreationDateTime"] = Security.GetCurrentDate();
            NewRow["Name"] = Name;
            NewRow["FactoryID"] = FactoryID;
            WorkAssignmentsDT.Rows.Add(NewRow);
        }

        public void FilterBatchesByMegaBatch(int GroupType, int MegaBatchID)
        {
            BatchesBS.Filter = "GroupType=" + GroupType + " AND MegaBatchID=" + MegaBatchID;
        }

        public void Initialize()
        {
            Create();
            Fill();
            Binding();
        }

        public bool IsBatchChanged(int GroupType, int BatchID)
        {
            DataRow[] rows = WorkAssignmentsDT.Select("GroupType = " + GroupType + " AND BatchID = " + BatchID);
            if (rows.Count() > 0 && Convert.ToBoolean(rows[0]["Changed"]))
                return true;
            return false;
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

        public int IsTPS45Printed(int WorkAssignmentID)
        {
            int Count = 0;
            int PrintingStatus = 0;
            int[] FrontsID = new int[18];
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
            FrontsID[15] = Convert.ToInt32(Fronts.Scandia);
            FrontsID[16] = Convert.ToInt32(Fronts.SofiaNotColored);
            FrontsID[17] = Convert.ToInt32(Fronts.Turin3NotColored);
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

        public void MoveToWorkAssignment(int WorkAssignmentID)
        {
            WorkAssignmentsBS.Position = WorkAssignmentsBS.Find("WorkAssignmentID", WorkAssignmentID);
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

        private void Binding()
        {
            BatchesBS.DataSource = BatchesDT;
            MegaBatchesBS.DataSource = MegaBatchesDT;
            WorkAssignmentsBS.DataSource = WorkAssignmentsDT;
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

        private void MachineNewRow(string ValueMember, string DisplayMember)
        {
            DataRow NewRow = MachinesDT.NewRow();
            NewRow["ValueMember"] = ValueMember;
            NewRow["DisplayMember"] = DisplayMember;
            MachinesDT.Rows.Add(NewRow);
        }
    }
}