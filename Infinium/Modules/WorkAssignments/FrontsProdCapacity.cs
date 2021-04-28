using System;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Windows.Forms;

namespace Infinium.Modules.WorkAssignments
{
    public class FrontsProdCapacity
    {
        private DataTable DataTable1;
        private DataTable FixedMaterialDT;
        private DataTable FrontsConfigDT;
        private DataTable FrontsOrdersDT;
        private BindingSource MaterialBS;

        private DataTable MaterialDT;
        private int NestedLevel = 1;
        private DataTable OperationsDetailDT;
        private DataTable OperationsTermsDT;
        private BindingSource ResultBS;

        private DataTable ResultDT;
        private DataTable StoreDetailDT;
        private BindingSource SummaryBS;
        private DataTable SummaryDT;

        private DataTable SumOperationsDetailDT;

        private DataTable SumStoreDetailDT;

        private int TechStoreID, InsetTypeID, PatinaID, Height, Width = 0;

        public FrontsProdCapacity(int iTechStoreID, int iInsetTypeID, int iPatinaID, int iHeight, int iWidth)
        {
            TechStoreID = iTechStoreID;
            InsetTypeID = iInsetTypeID;
            PatinaID = iPatinaID;
            Height = iHeight;
            Width = iWidth;

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

        public void Test()
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

        private void DeleteMaterial(int PanelCounter)
        {
            DataRow[] rows = MaterialDT.Select("PanelCounter=" + PanelCounter);
            for (int i = rows.Count() - 1; i >= 0; i--)
                rows[i].Delete();
        }

        private bool GetOperationsDetail(int PrevTechCatalogOperationsDetailID, int NestedLevel, DataRow Row)
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

        private bool GetStoreDetail(int PrevTechCatalogOperationsDetailID, int NestedLevel, DataRow Row)
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
    }
}