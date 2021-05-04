using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.UserModel.Contrib;
using NPOI.HSSF.Util;

using System;
using System.Collections;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;

namespace Infinium.Modules.WorkAssignments
{

    public class TPSAngle45Assignments : IAllFrontParameterName
    {
        private ArrayList FrontsID;
        private FileManager FM = new FileManager();
        private DateTime CurrentDate;
        private int FrontType = 0;
        private DataTable AssemblyDT;
        private DataTable DeyingDT;
        private DataTable DecorDT;
        public DataTable TechStoreDataTable = null;
        public DataTable FrontsDataTable = null;
        public DataTable FrameColorsDataTable = null;
        public DataTable PatinaDataTable = null;
        private DataTable PatinaRALDataTable = null;
        public DataTable InsetTypesDataTable = null;
        public DataTable InsetColorsDataTable = null;
        private DataTable GashDT;
        private DataTable RemovingQuarterDT;
        private DataTable RemovingBoxesDT;
        private DataTable GrooveGridsDT;
        private DataTable TrimmingDT;
        private DataTable InsetDT;
        private DataTable SummOrdersDT;
        private DataTable CurvedAssemblyDT;
        private DataTable BagetWithAngleAssemblyDT;
        private DataTable DecorAssemblyDT;
        private DataTable FilenkaBoxesDT;
        private DataTable FilenkaSimpleDT;
        private DataTable DakotaFilenkaSimpleDT;
        private DataTable ProfileNamesDT;
        private DataTable DecorParametersDT;
        private DataTable InsetTypeNamesDT;
        private DataTable LorenzoVitrinaDT;
        private DataTable LorenzoBoxesDT;
        private DataTable LorenzoGridsDT;
        private DataTable LorenzoSimpleDT;
        private DataTable LorenzoOrdersDT;
        private DataTable ElegantVitrinaDT;
        private DataTable ElegantBoxesDT;
        private DataTable ElegantGridsDT;
        private DataTable ElegantSimpleDT;
        private DataTable ElegantOrdersDT;
        private DataTable Patricia1VitrinaDT;
        private DataTable Patricia1BoxesDT;
        private DataTable Patricia1GridsDT;
        private DataTable Patricia1SimpleDT;
        private DataTable Patricia1OrdersDT;
        private DataTable KansasVitrinaDT;
        private DataTable KansasBoxesDT;
        private DataTable KansasGridsDT;
        private DataTable KansasSimpleDT;
        private DataTable KansasOrdersDT;
        private DataTable DakotaVitrinaDT;
        private DataTable DakotaBoxesDT;
        private DataTable DakotaGridsDT;
        private DataTable DakotaSimpleDT;
        private DataTable DakotaOrdersDT;
        private DataTable SofiaVitrinaDT;
        private DataTable SofiaBoxesDT;
        private DataTable SofiaGridsDT;
        private DataTable SofiaSimpleDT;
        private DataTable SofiaOrdersDT;
        private DataTable Turin1VitrinaDT;
        private DataTable Turin1BoxesDT;
        private DataTable Turin1GridsDT;
        private DataTable Turin1SimpleDT;
        private DataTable Turin1OrdersDT;
        private DataTable Turin1_1VitrinaDT;
        private DataTable Turin1_1BoxesDT;
        private DataTable Turin1_1GridsDT;
        private DataTable Turin1_1SimpleDT;
        private DataTable Turin1_1OrdersDT;
        private DataTable Turin3VitrinaDT;
        private DataTable Turin3BoxesDT;
        private DataTable Turin3GridsDT;
        private DataTable Turin3SimpleDT;
        private DataTable Turin3OrdersDT;
        private DataTable InfinitiVitrinaDT;
        private DataTable InfinitiBoxesDT;
        private DataTable InfinitiGridsDT;
        private DataTable InfinitiSimpleDT;
        private DataTable InfinitiOrdersDT;
        private DataTable LeonVitrinaDT;
        private DataTable LeonBoxesDT;
        private DataTable LeonGridsDT;
        private DataTable LeonSimpleDT;
        private DataTable LeonOrdersDT;
        private DataTable Turin1RemovingBoxesDT;
        private DataTable Turin3RemovingBoxesDT;
        private DataTable DakotaAppliqueDT;
        private DataTable SofiaAppliqueDT;
        private DataTable LorenzoCurvedOrdersDT;
        private DataTable ElegantCurvedOrdersDT;
        private DataTable Patricia1CurvedOrdersDT;
        private DataTable KansasCurvedOrdersDT;
        private DataTable SofiaCurvedOrdersDT;
        private DataTable DakotaCurvedOrdersDT;
        private DataTable Turin1CurvedOrdersDT;
        private DataTable Turin1_1CurvedOrdersDT;
        private DataTable Turin3CurvedOrdersDT;
        private DataTable InfinitiCurvedOrdersDT;
        private DataTable BagetWithAngelOrdersDT;
        private DataTable NotArchDecorOrdersDT;
        private DataTable ArchDecorOrdersDT;
        private DataTable GridsDecorOrdersDT;

        public TPSAngle45Assignments()
        {

        }

        public void Initialize()
        {
            Create();
            Fill();
        }

        private void Create()
        {
            FrontsID = new ArrayList();
            ProfileNamesDT = new DataTable();
            InsetTypeNamesDT = new DataTable();
            DecorParametersDT = new DataTable();
            DecorDT = new DataTable();
            TechStoreDataTable = new DataTable();
            FrontsDataTable = new DataTable();
            FrameColorsDataTable = new DataTable();
            PatinaDataTable = new DataTable();
            InsetTypesDataTable = new DataTable();
            InsetColorsDataTable = new DataTable();
            SummOrdersDT = new DataTable();

            LorenzoBoxesDT = new DataTable();
            LorenzoGridsDT = new DataTable();
            LorenzoSimpleDT = new DataTable();
            LorenzoOrdersDT = new DataTable();

            ElegantBoxesDT = new DataTable();
            ElegantGridsDT = new DataTable();
            ElegantSimpleDT = new DataTable();
            ElegantOrdersDT = new DataTable();

            Patricia1BoxesDT = new DataTable();
            Patricia1GridsDT = new DataTable();
            Patricia1SimpleDT = new DataTable();
            Patricia1OrdersDT = new DataTable();

            KansasBoxesDT = new DataTable();
            KansasGridsDT = new DataTable();
            KansasSimpleDT = new DataTable();
            KansasOrdersDT = new DataTable();

            DakotaBoxesDT = new DataTable();
            DakotaGridsDT = new DataTable();
            DakotaSimpleDT = new DataTable();
            DakotaOrdersDT = new DataTable();

            SofiaBoxesDT = new DataTable();
            SofiaGridsDT = new DataTable();
            DakotaAppliqueDT = new DataTable();
            SofiaAppliqueDT = new DataTable();
            SofiaSimpleDT = new DataTable();
            SofiaOrdersDT = new DataTable();

            Turin1BoxesDT = new DataTable();
            Turin1GridsDT = new DataTable();
            Turin1SimpleDT = new DataTable();
            Turin1OrdersDT = new DataTable();

            Turin1_1BoxesDT = new DataTable();
            Turin1_1GridsDT = new DataTable();
            Turin1_1SimpleDT = new DataTable();
            Turin1_1OrdersDT = new DataTable();

            Turin3BoxesDT = new DataTable();
            Turin3GridsDT = new DataTable();
            Turin3SimpleDT = new DataTable();
            Turin3OrdersDT = new DataTable();

            LeonBoxesDT = new DataTable();
            LeonGridsDT = new DataTable();
            LeonSimpleDT = new DataTable();
            LeonOrdersDT = new DataTable();

            InfinitiBoxesDT = new DataTable();
            InfinitiGridsDT = new DataTable();
            InfinitiSimpleDT = new DataTable();

            LorenzoCurvedOrdersDT = new DataTable();
            ElegantCurvedOrdersDT = new DataTable();
            Patricia1CurvedOrdersDT = new DataTable();
            KansasCurvedOrdersDT = new DataTable();
            SofiaCurvedOrdersDT = new DataTable();
            DakotaCurvedOrdersDT = new DataTable();
            Turin1CurvedOrdersDT = new DataTable();
            Turin1_1CurvedOrdersDT = new DataTable();
            Turin3CurvedOrdersDT = new DataTable();
            InfinitiCurvedOrdersDT = new DataTable();

            BagetWithAngelOrdersDT = new DataTable();
            NotArchDecorOrdersDT = new DataTable();
            ArchDecorOrdersDT = new DataTable();
            GridsDecorOrdersDT = new DataTable();

            GashDT = new DataTable();
            GashDT.Columns.Add(new DataColumn("Front", Type.GetType("System.String")));
            GashDT.Columns.Add(new DataColumn("Color", Type.GetType("System.String")));
            GashDT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            GashDT.Columns.Add(new DataColumn("Count", Type.GetType("System.Int32")));
            GashDT.Columns.Add(new DataColumn("VitrinaNotes", Type.GetType("System.String")));
            GashDT.Columns.Add(new DataColumn("GridNotes", Type.GetType("System.String")));
            GashDT.Columns.Add(new DataColumn("FrontID", Type.GetType("System.Int32")));
            GashDT.Columns.Add(new DataColumn("ColorID", Type.GetType("System.Int32")));
            GashDT.Columns.Add(new DataColumn("ProfileType", Type.GetType("System.Int32")));
            GashDT.Columns.Add(new DataColumn("VitrinaCount", Type.GetType("System.Int32")));
            GashDT.Columns.Add(new DataColumn("GridCount", Type.GetType("System.Int32")));

            RemovingQuarterDT = new DataTable();
            RemovingQuarterDT.Columns.Add(new DataColumn("Front", Type.GetType("System.String")));
            RemovingQuarterDT.Columns.Add(new DataColumn("Color", Type.GetType("System.String")));
            RemovingQuarterDT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            RemovingQuarterDT.Columns.Add(new DataColumn("Count", Type.GetType("System.Int32")));
            RemovingQuarterDT.Columns.Add(new DataColumn("FrontID", Type.GetType("System.Int32")));
            RemovingQuarterDT.Columns.Add(new DataColumn("ColorID", Type.GetType("System.Int32")));

            RemovingBoxesDT = new DataTable();
            RemovingBoxesDT.Columns.Add(new DataColumn("Front", Type.GetType("System.String")));
            RemovingBoxesDT.Columns.Add(new DataColumn("Color", Type.GetType("System.String")));
            RemovingBoxesDT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            RemovingBoxesDT.Columns.Add(new DataColumn("Count", Type.GetType("System.Int32")));
            RemovingBoxesDT.Columns.Add(new DataColumn("FrontID", Type.GetType("System.Int32")));
            RemovingBoxesDT.Columns.Add(new DataColumn("ColorID", Type.GetType("System.Int32")));

            GrooveGridsDT = new DataTable();
            GrooveGridsDT.Columns.Add(new DataColumn("Front", Type.GetType("System.String")));
            GrooveGridsDT.Columns.Add(new DataColumn("Color", Type.GetType("System.String")));
            GrooveGridsDT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            GrooveGridsDT.Columns.Add(new DataColumn("Count", Type.GetType("System.Int32")));
            GrooveGridsDT.Columns.Add(new DataColumn("FrontID", Type.GetType("System.Int32")));
            GrooveGridsDT.Columns.Add(new DataColumn("ColorID", Type.GetType("System.Int32")));

            TrimmingDT = new DataTable();
            TrimmingDT.Columns.Add(new DataColumn("Front", Type.GetType("System.String")));
            TrimmingDT.Columns.Add(new DataColumn("Color", Type.GetType("System.String")));
            TrimmingDT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            TrimmingDT.Columns.Add(new DataColumn("Count", Type.GetType("System.Int32")));
            TrimmingDT.Columns.Add(new DataColumn("VitrinaNotes", Type.GetType("System.String")));
            TrimmingDT.Columns.Add(new DataColumn("GridNotes", Type.GetType("System.String")));
            TrimmingDT.Columns.Add(new DataColumn("FrontID", Type.GetType("System.Int32")));
            TrimmingDT.Columns.Add(new DataColumn("ColorID", Type.GetType("System.Int32")));
            TrimmingDT.Columns.Add(new DataColumn("ProfileType", Type.GetType("System.Int32")));
            TrimmingDT.Columns.Add(new DataColumn("VitrinaCount", Type.GetType("System.Int32")));
            TrimmingDT.Columns.Add(new DataColumn("GridCount", Type.GetType("System.Int32")));

            InsetDT = new DataTable();
            InsetDT.Columns.Add(new DataColumn("Name", Type.GetType("System.String")));
            InsetDT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            InsetDT.Columns.Add(new DataColumn("Width", Type.GetType("System.Int32")));
            InsetDT.Columns.Add(new DataColumn("Count", Type.GetType("System.Int32")));
            InsetDT.Columns.Add(new DataColumn("ColorType", Type.GetType("System.Int32")));

            FilenkaBoxesDT = new DataTable();
            FilenkaBoxesDT.Columns.Add(new DataColumn("Name", Type.GetType("System.String")));
            FilenkaBoxesDT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            FilenkaBoxesDT.Columns.Add(new DataColumn("Width", Type.GetType("System.Int32")));
            FilenkaBoxesDT.Columns.Add(new DataColumn("Count", Type.GetType("System.Int32")));
            FilenkaBoxesDT.Columns.Add(new DataColumn("Square", Type.GetType("System.Decimal")));
            FilenkaBoxesDT.Columns.Add(new DataColumn("ColorType", Type.GetType("System.Int32")));

            FilenkaSimpleDT = new DataTable();
            FilenkaSimpleDT.Columns.Add(new DataColumn("Name", Type.GetType("System.String")));
            FilenkaSimpleDT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            FilenkaSimpleDT.Columns.Add(new DataColumn("Width", Type.GetType("System.Int32")));
            FilenkaSimpleDT.Columns.Add(new DataColumn("Count", Type.GetType("System.Int32")));
            FilenkaSimpleDT.Columns.Add(new DataColumn("Square", Type.GetType("System.Decimal")));
            FilenkaSimpleDT.Columns.Add(new DataColumn("ColorType", Type.GetType("System.Int32")));

            DakotaFilenkaSimpleDT = new DataTable();
            DakotaFilenkaSimpleDT.Columns.Add(new DataColumn("Name", Type.GetType("System.String")));
            DakotaFilenkaSimpleDT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            DakotaFilenkaSimpleDT.Columns.Add(new DataColumn("Width", Type.GetType("System.Int32")));
            DakotaFilenkaSimpleDT.Columns.Add(new DataColumn("Count", Type.GetType("System.Int32")));
            DakotaFilenkaSimpleDT.Columns.Add(new DataColumn("Square", Type.GetType("System.Decimal")));
            DakotaFilenkaSimpleDT.Columns.Add(new DataColumn("ColorType", Type.GetType("System.Int32")));

            AssemblyDT = new DataTable();
            AssemblyDT.Columns.Add(new DataColumn("Name", Type.GetType("System.String")));
            AssemblyDT.Columns.Add(new DataColumn("FrameColor", Type.GetType("System.String")));
            AssemblyDT.Columns.Add(new DataColumn("InsetColor", Type.GetType("System.String")));
            AssemblyDT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            AssemblyDT.Columns.Add(new DataColumn("Width", Type.GetType("System.Int32")));
            AssemblyDT.Columns.Add(new DataColumn("Count", Type.GetType("System.Int32")));
            AssemblyDT.Columns.Add(new DataColumn("Square", Type.GetType("System.Decimal")));
            AssemblyDT.Columns.Add(new DataColumn("Worker", Type.GetType("System.String")));
            AssemblyDT.Columns.Add(new DataColumn("FrontType", Type.GetType("System.Int32")));
            AssemblyDT.Columns.Add(new DataColumn("ColorType", Type.GetType("System.Int32")));

            CurvedAssemblyDT = new DataTable();
            CurvedAssemblyDT.Columns.Add(new DataColumn("Front", Type.GetType("System.String")));
            CurvedAssemblyDT.Columns.Add(new DataColumn("FrameColor", Type.GetType("System.String")));
            CurvedAssemblyDT.Columns.Add(new DataColumn("InsetType", Type.GetType("System.String")));
            CurvedAssemblyDT.Columns.Add(new DataColumn("InsetColor", Type.GetType("System.String")));
            CurvedAssemblyDT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            CurvedAssemblyDT.Columns.Add(new DataColumn("Count", Type.GetType("System.Int32")));
            CurvedAssemblyDT.Columns.Add(new DataColumn("FrontID", Type.GetType("System.Int32")));
            CurvedAssemblyDT.Columns.Add(new DataColumn("ColorID", Type.GetType("System.Int32")));
            CurvedAssemblyDT.Columns.Add(new DataColumn("PatinaID", Type.GetType("System.Int32")));
            CurvedAssemblyDT.Columns.Add(new DataColumn("InsetTypeID", Type.GetType("System.Int32")));
            CurvedAssemblyDT.Columns.Add(new DataColumn("InsetColorID", Type.GetType("System.Int32")));

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

            DeyingDT = new DataTable();
            DeyingDT.Columns.Add(new DataColumn("Name", Type.GetType("System.String")));
            DeyingDT.Columns.Add(new DataColumn("FrameColor", Type.GetType("System.String")));
            DeyingDT.Columns.Add(new DataColumn("InsetColor", Type.GetType("System.String")));
            DeyingDT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            DeyingDT.Columns.Add(new DataColumn("Width", Type.GetType("System.Int32")));
            DeyingDT.Columns.Add(new DataColumn("Count", Type.GetType("System.Int32")));
            DeyingDT.Columns.Add(new DataColumn("Square", Type.GetType("System.Decimal")));
            DeyingDT.Columns.Add(new DataColumn("Notes", Type.GetType("System.String")));
            DeyingDT.Columns.Add(new DataColumn("ColorType", Type.GetType("System.Int32")));
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
            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT TOP 0 infiniu2_catalog.dbo.TechStore.TechStoreName, infiniu2_catalog.dbo.FrontsConfig.FrontConfigID FROM infiniu2_catalog.dbo.TechStore
                    INNER JOIN infiniu2_catalog.dbo.FrontsConfig ON infiniu2_catalog.dbo.TechStore.TechStoreID = infiniu2_catalog.dbo.FrontsConfig.ProfileID",
                ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(ProfileNamesDT);
            }
            DecorDT = new DataTable();
            string SelectCommand = @"SELECT DISTINCT TechStore.TechStoreID AS DecorID, TechStore.TechStoreName AS Name, DecorConfig.ProductID FROM TechStore 
                INNER JOIN DecorConfig ON TechStore.TechStoreID = DecorConfig.DecorID AND Enabled = 1 ORDER BY TechStoreName";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(DecorDT);
            }

            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT * FROM DecorParameters",
                ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(DecorParametersDT);
            }
            SelectCommand = @"SELECT DISTINCT TechStoreID AS FrontID, TechStoreName AS FrontName FROM TechStore 
                WHERE TechStoreID IN (SELECT FrontID FROM FrontsConfig WHERE Enabled = 1) ORDER BY TechStoreName";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(FrontsDataTable);
            }

            GetColorsDT();
            GetInsetColorsDT();
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM InsetTypes",
                ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(InsetTypesDataTable);
            }
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
            //using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM TechStore",
            //    ConnectionStrings.CatalogConnectionString))
            //{
            //    DA.Fill(TechStoreDataTable);
            //}
            TechStoreDataTable = TablesManager.TechStoreDataTable;

            using (SqlDataAdapter DA = new SqlDataAdapter(@"SELECT TOP 0 FrontsOrdersID, TechnoProfileID, MainOrderID, FrontID,PatinaID, InsetTypeID,
                ColorID, InsetColorID, Height, Width, Count, FrontConfigID, RTRIM(lower(Notes)) AS Notes FROM FrontsOrders",
                ConnectionStrings.ZOVOrdersConnectionString))
            {
                DA.Fill(SofiaOrdersDT);
                SofiaOrdersDT.Columns.Add(new DataColumn("GroupType", Type.GetType("System.Int32")));

                LorenzoCurvedOrdersDT = SofiaOrdersDT.Clone();
                ElegantCurvedOrdersDT = SofiaOrdersDT.Clone();
                Patricia1CurvedOrdersDT = SofiaOrdersDT.Clone();
                KansasCurvedOrdersDT = SofiaOrdersDT.Clone();
                SofiaCurvedOrdersDT = SofiaOrdersDT.Clone();
                DakotaCurvedOrdersDT = SofiaOrdersDT.Clone();
                Turin1CurvedOrdersDT = SofiaOrdersDT.Clone();
                Turin1_1CurvedOrdersDT = SofiaOrdersDT.Clone();
                Turin3CurvedOrdersDT = SofiaOrdersDT.Clone();
                InfinitiCurvedOrdersDT = SofiaOrdersDT.Clone();

                SofiaBoxesDT = SofiaOrdersDT.Clone();
                SofiaGridsDT = SofiaOrdersDT.Clone();
                SofiaSimpleDT = SofiaOrdersDT.Clone();
                DakotaAppliqueDT = SofiaOrdersDT.Clone();
                SofiaAppliqueDT = SofiaOrdersDT.Clone();

                LorenzoVitrinaDT = SofiaOrdersDT.Clone();
                ElegantVitrinaDT = SofiaOrdersDT.Clone();
                Patricia1VitrinaDT = SofiaOrdersDT.Clone();
                KansasVitrinaDT = SofiaOrdersDT.Clone();
                DakotaVitrinaDT = SofiaOrdersDT.Clone();
                SofiaVitrinaDT = SofiaOrdersDT.Clone();
                Turin1VitrinaDT = SofiaOrdersDT.Clone();
                Turin1_1VitrinaDT = SofiaOrdersDT.Clone();
                InfinitiVitrinaDT = SofiaOrdersDT.Clone();

                Turin1RemovingBoxesDT = SofiaOrdersDT.Clone();
                Turin3RemovingBoxesDT = SofiaOrdersDT.Clone();

                LorenzoOrdersDT = SofiaOrdersDT.Clone();
                LorenzoBoxesDT = SofiaOrdersDT.Clone();
                LorenzoGridsDT = SofiaOrdersDT.Clone();
                LorenzoSimpleDT = SofiaOrdersDT.Clone();

                ElegantOrdersDT = SofiaOrdersDT.Clone();
                ElegantBoxesDT = SofiaOrdersDT.Clone();
                ElegantGridsDT = SofiaOrdersDT.Clone();
                ElegantSimpleDT = SofiaOrdersDT.Clone();

                Patricia1OrdersDT = SofiaOrdersDT.Clone();
                Patricia1BoxesDT = SofiaOrdersDT.Clone();
                Patricia1GridsDT = SofiaOrdersDT.Clone();
                Patricia1SimpleDT = SofiaOrdersDT.Clone();

                KansasOrdersDT = SofiaOrdersDT.Clone();
                KansasBoxesDT = SofiaOrdersDT.Clone();
                KansasGridsDT = SofiaOrdersDT.Clone();
                KansasSimpleDT = SofiaOrdersDT.Clone();

                DakotaOrdersDT = SofiaOrdersDT.Clone();
                DakotaBoxesDT = SofiaOrdersDT.Clone();
                DakotaGridsDT = SofiaOrdersDT.Clone();
                DakotaSimpleDT = SofiaOrdersDT.Clone();

                Turin1OrdersDT = SofiaOrdersDT.Clone();
                Turin1BoxesDT = SofiaOrdersDT.Clone();
                Turin1GridsDT = SofiaOrdersDT.Clone();
                Turin1SimpleDT = SofiaOrdersDT.Clone();

                Turin1_1OrdersDT = SofiaOrdersDT.Clone();
                Turin1_1BoxesDT = SofiaOrdersDT.Clone();
                Turin1_1GridsDT = SofiaOrdersDT.Clone();
                Turin1_1SimpleDT = SofiaOrdersDT.Clone();

                Turin3VitrinaDT = SofiaOrdersDT.Clone();
                Turin3OrdersDT = SofiaOrdersDT.Clone();
                Turin3BoxesDT = SofiaOrdersDT.Clone();
                Turin3GridsDT = SofiaOrdersDT.Clone();
                Turin3SimpleDT = SofiaOrdersDT.Clone();

                LeonVitrinaDT = SofiaOrdersDT.Clone();
                LeonOrdersDT = SofiaOrdersDT.Clone();
                LeonBoxesDT = SofiaOrdersDT.Clone();
                LeonGridsDT = SofiaOrdersDT.Clone();
                LeonSimpleDT = SofiaOrdersDT.Clone();

                InfinitiOrdersDT = SofiaOrdersDT.Clone();
                InfinitiBoxesDT = SofiaOrdersDT.Clone();
                InfinitiGridsDT = SofiaOrdersDT.Clone();
                InfinitiSimpleDT = SofiaOrdersDT.Clone();
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
                GridsDecorOrdersDT = NotArchDecorOrdersDT.Clone();
            }
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

        public bool HasParameter(int ProductID, string Parameter)
        {
            DataRow[] Rows = DecorParametersDT.Select("ProductID = " + ProductID);

            return Convert.ToBoolean(Rows[0][Parameter]);
        }

        private string GetDecorName(int ID)
        {
            DataRow[] rows = DecorDT.Select("DecorID=" + ID);
            if (rows.Count() > 0)
                return rows[0]["Name"].ToString();
            else
                return string.Empty;
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

        private string ProfileName(int ID)
        {
            string name = string.Empty;
            DataRow[] rows = ProfileNamesDT.Select("FrontConfigID=" + ID);
            if (rows.Count() > 0)
                name = rows[0]["TechStoreName"].ToString();
            return name;
        }

        //private string InsetTypeName(int ID)
        //{
        //    string name = string.Empty;
        //    DataRow[] rows = InsetTypeNamesDT.Select("FrontConfigID=" + ID);
        //    if (rows.Count() > 0)
        //        name = rows[0]["TechStoreName"].ToString();
        //    return name;
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

        private void GetProfileNames(ref DataTable DestinationDT, int WorkAssignmentID, int FactoryID, Fronts Front)
        {
            string SelectCommand = string.Empty;
            DataTable DT = new DataTable();

            SelectCommand = @"SELECT infiniu2_catalog.dbo.TechStore.TechStoreName, infiniu2_catalog.dbo.FrontsConfig.FrontConfigID FROM infiniu2_catalog.dbo.TechStore
                    INNER JOIN infiniu2_catalog.dbo.FrontsConfig ON infiniu2_catalog.dbo.TechStore.TechStoreID = infiniu2_catalog.dbo.FrontsConfig.ProfileID AND infiniu2_catalog.dbo.FrontsConfig.FrontConfigID IN (SELECT FrontConfigID FROM FrontsOrders
                    WHERE FrontID=" + Convert.ToInt32(Front) +
                    " AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE ProfilWorkAssignmentID=" + WorkAssignmentID + ")))";
            if (FactoryID == 2)
                SelectCommand = @"SELECT infiniu2_catalog.dbo.TechStore.TechStoreName, infiniu2_catalog.dbo.FrontsConfig.FrontConfigID FROM infiniu2_catalog.dbo.TechStore
                    INNER JOIN infiniu2_catalog.dbo.FrontsConfig ON infiniu2_catalog.dbo.TechStore.TechStoreID = infiniu2_catalog.dbo.FrontsConfig.ProfileID AND infiniu2_catalog.dbo.FrontsConfig.FrontConfigID IN (SELECT FrontConfigID FROM FrontsOrders
                    WHERE FrontID=" + Convert.ToInt32(Front) +
                        " AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE TPSWorkAssignmentID=" + WorkAssignmentID + ")))";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.ZOVOrdersConnectionString))
            {
                DA.Fill(DT);
            }
            foreach (DataRow item in DT.Rows)
            {
                DataRow NewRow = DestinationDT.NewRow();
                NewRow.ItemArray = item.ItemArray;
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
                DestinationDT.Rows.Add(NewRow);
            }
        }

        private void GetBoxFronts(DataTable SourceDT, ref DataTable DestinationDT, int Admission)
        {
            DataRow[] rows = SourceDT.Select("Height<=" + Admission + " OR Width<=" + Admission);
            foreach (DataRow dr in rows)
                DestinationDT.Rows.Add(dr.ItemArray);
        }

        private void GetGridFronts(DataTable SourceDT, ref DataTable DestinationDT)
        {
            DataRow[] rows = SourceDT.Select("InsetTypeID IN (685,686,687,688,29470,29471)");
            foreach (DataRow dr in rows)
                DestinationDT.Rows.Add(dr.ItemArray);
        }

        private void GetAppliqueFronts(DataTable SourceDT, ref DataTable DestinationDT)
        {
            DataRow[] rows = SourceDT.Select("InsetTypeID IN (28961,3653,3654,3655)");
            foreach (DataRow dr in rows)
                DestinationDT.Rows.Add(dr.ItemArray);
        }

        private void GetSimpleFronts(DataTable SourceDT, ref DataTable DestinationDT, int Admission)
        {
            DataRow[] rows = SourceDT.Select("InsetTypeID NOT IN (685,686,687,688,29470,29471,28961,3653,3654,3655) AND Height>" + Admission + " AND Width>" + Admission);
            foreach (DataRow dr in rows)
                DestinationDT.Rows.Add(dr.ItemArray);
        }

        private void GetVitrinaFronts(DataTable SourceDT, ref DataTable DestinationDT)
        {
            DataRow[] rows = SourceDT.Select("InsetTypeID=1");
            foreach (DataRow dr in rows)
                DestinationDT.Rows.Add(dr.ItemArray);
        }

        private void GetRemovingBoxesFronts(DataTable SourceDT, ref DataTable DestinationDT, int Admission)
        {
            DataRow[] rows = SourceDT.Select("Height<" + Admission + " OR Width<" + Admission);
            foreach (DataRow dr in rows)
                DestinationDT.Rows.Add(dr.ItemArray);
        }

        private void GetFrontsOrders(ref DataTable DestinationDT, int WorkAssignmentID, int FactoryID, Fronts Front)
        {
            string SelectCommand = string.Empty;
            DataTable DT = new DataTable();

            SelectCommand = @"SELECT FrontsOrdersID, TechnoProfileID, MainOrderID, FrontID, PatinaID, InsetTypeID,
                ColorID, InsetColorID, Height, Width, Count, FrontConfigID, RTRIM(lower(Notes)) AS Notes FROM FrontsOrders
                WHERE FrontID=" + Convert.ToInt32(Front) +
                " AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE FactoryID=1 AND BatchID IN (SELECT BatchID FROM Batch WHERE ProfilWorkAssignmentID=" + WorkAssignmentID + "))";
            if (FactoryID == 2)
                SelectCommand = @"SELECT FrontsOrdersID, TechnoProfileID, MainOrderID, FrontID, PatinaID, InsetTypeID,
                    ColorID, InsetColorID, Height, Width, Count, FrontConfigID, RTRIM(lower(Notes)) AS Notes FROM FrontsOrders
                    WHERE FrontID=" + Convert.ToInt32(Front) +
                    " AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE FactoryID=2 AND BatchID IN (SELECT BatchID FROM Batch WHERE TPSWorkAssignmentID=" + WorkAssignmentID + "))";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.ZOVOrdersConnectionString))
            {
                DA.Fill(DT);
            }
            foreach (DataRow item in DT.Rows)
            {
                if (item["Notes"] != DBNull.Value && item["Notes"].ToString().Length == 0)
                    item["Notes"] = DBNull.Value;
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
                if (item["Notes"] != DBNull.Value && item["Notes"].ToString().Length == 0)
                    item["Notes"] = DBNull.Value;
                DataRow NewRow = DestinationDT.NewRow();
                NewRow.ItemArray = item.ItemArray;
                NewRow["GroupType"] = 1;
                DestinationDT.Rows.Add(NewRow);
            }
        }

        private void GetCurvedFrontsOrders(ref DataTable DestinationDT, int WorkAssignmentID, int FactoryID, string filter)
        {
            string SelectCommand = string.Empty;
            DataTable DT = new DataTable();

            SelectCommand = @"SELECT FrontsOrdersID, TechnoProfileID, MainOrderID, FrontID, PatinaID, InsetTypeID,
                ColorID, InsetColorID, Height, Width, Count, FrontConfigID, Notes FROM FrontsOrders
                WHERE FrontID IN " + filter +
                " AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE ProfilWorkAssignmentID=" + WorkAssignmentID + "))";
            if (FactoryID == 2)
                SelectCommand = @"SELECT FrontsOrdersID, TechnoProfileID, MainOrderID, FrontID, PatinaID, InsetTypeID,
                    ColorID, InsetColorID, Height, Width, Count, FrontConfigID, Notes FROM FrontsOrders
                    WHERE FrontID IN " + filter +
                    " AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE TPSWorkAssignmentID=" + WorkAssignmentID + "))";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.ZOVOrdersConnectionString))
            {
                DA.Fill(DT);
            }
            foreach (DataRow item in DT.Rows)
            {
                if (item["Notes"] != DBNull.Value && item["Notes"].ToString().Length == 0)
                    item["Notes"] = DBNull.Value;
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
                if (item["Notes"] != DBNull.Value && item["Notes"].ToString().Length == 0)
                    item["Notes"] = DBNull.Value;
                DataRow NewRow = DestinationDT.NewRow();
                NewRow.ItemArray = item.ItemArray;
                NewRow["GroupType"] = 1;
                DestinationDT.Rows.Add(NewRow);
            }
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
                WHERE FactoryID=" + FactoryID + " AND NOT (ProductID IN (1) AND (LeftAngle<>0 OR RightAngle<>0)) AND ProductID NOT IN (31, 4, 18, 32, 10, 11, 12) AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE ProfilWorkAssignmentID=" + WorkAssignmentID + "))";
            if (FactoryID == 2)
                SelectCommand = @"SELECT DecorOrderID, MainOrderID, ProductID, DecorID, DecorConfigID, ColorID, PatinaID,
                    Height, Length, Width, Count, Notes FROM DecorOrders
                    WHERE FactoryID=" + FactoryID + " AND NOT (ProductID IN (1) AND (LeftAngle<>0 OR RightAngle<>0)) AND ProductID NOT IN (31, 4, 18, 32, 10, 11, 12) AND MainOrderID IN (SELECT MainOrderID FROM BatchDetails WHERE BatchID IN (SELECT BatchID FROM Batch WHERE TPSWorkAssignmentID=" + WorkAssignmentID + "))";

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

        private DataTable DistFrameColorsTable(DataTable SourceDT, bool OrderASC)
        {
            int ColorID = 0;
            DataTable DT = new DataTable();
            DT.Columns.Add(new DataColumn("ColorID", Type.GetType("System.Int32")));
            foreach (DataRow Row in SourceDT.Rows)
            {
                if (int.TryParse(Row["ColorID"].ToString(), out ColorID))
                {
                    DataRow NewRow = DT.NewRow();
                    NewRow["ColorID"] = ColorID;
                    DT.Rows.Add(NewRow);
                }
            }
            using (DataView DV = new DataView(DT.Copy()))
            {
                DT.Clear();
                if (OrderASC)
                    DV.Sort = "ColorID ASC";
                else
                    DV.Sort = "ColorID DESC";
                DT = DV.ToTable(true, new string[] { "ColorID" });
            }
            return DT;
        }

        private DataTable DistInsetColorsTable(DataTable SourceDT, bool OrderASC)
        {
            int InsetColorID = 0;
            DataTable DT = new DataTable();
            DT.Columns.Add(new DataColumn("InsetColorID", Type.GetType("System.Int32")));
            foreach (DataRow Row in SourceDT.Rows)
            {
                //if (Convert.ToInt32(Row["InsetTypeID"]) != 2 && Convert.ToInt32(Row["InsetTypeID"]) != 5 && Convert.ToInt32(Row["InsetTypeID"]) != 6
                //    && Convert.ToInt32(Row["InsetTypeID"]) != 9 && Convert.ToInt32(Row["InsetTypeID"]) != 10 && Convert.ToInt32(Row["InsetTypeID"]) != 11)
                //    continue;

                if (int.TryParse(Row["InsetColorID"].ToString(), out InsetColorID))
                {
                    DataRow NewRow = DT.NewRow();
                    NewRow["InsetColorID"] = InsetColorID;
                    DT.Rows.Add(NewRow);
                }
            }
            using (DataView DV = new DataView(DT.Copy()))
            {
                DT.Clear();
                if (OrderASC)
                    DV.Sort = "InsetColorID ASC";
                else
                    DV.Sort = "InsetColorID DESC";
                DT = DV.ToTable(true, new string[] { "InsetColorID" });
            }
            return DT;
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
            DataTable SourceDT6, DataTable SourceDT7, DataTable SourceDT8, DataTable SourceDT9, DataTable SourceDT10, DataTable SourceDT11, bool OrderASC)
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
            foreach (DataRow Row in SourceDT10.Rows)
            {
                if (int.TryParse(Row["MainOrderID"].ToString(), out MainOrderID))
                {
                    DataRow NewRow = DT.NewRow();
                    NewRow["MainOrderID"] = MainOrderID;
                    NewRow["GroupType"] = Convert.ToInt32(Row["GroupType"]);
                    DT.Rows.Add(NewRow);
                }
            }
            foreach (DataRow Row in SourceDT11.Rows)
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

        private void CurvedAssemblyCollect(DataTable SourceDT, ref DataTable DestinationDT)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            DataTable DT3 = new DataTable();
            DataTable DT4 = new DataTable();

            using (DataView DV = new DataView(SourceDT, string.Empty, "ColorID, PatinaID", DataViewRowState.CurrentRows))
            {
                DT1 = DV.ToTable(true, new string[] { "ColorID", "PatinaID" });
            }
            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                using (DataView DV = new DataView(SourceDT, "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) + " AND PatinaID=" + Convert.ToInt32(DT1.Rows[i]["PatinaID"]), "InsetTypeID", DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "InsetTypeID" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    using (DataView DV = new DataView(SourceDT, "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) + " AND PatinaID=" + Convert.ToInt32(DT1.Rows[i]["PatinaID"]) +
                        " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]), "InsetColorID", DataViewRowState.CurrentRows))
                    {
                        DT3 = DV.ToTable(true, new string[] { "InsetColorID" });
                    }
                    for (int x = 0; x < DT3.Rows.Count; x++)
                    {
                        using (DataView DV = new DataView(SourceDT, "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) + " AND PatinaID=" + Convert.ToInt32(DT1.Rows[i]["PatinaID"]) +
                            " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) +
                            " AND InsetColorID=" + Convert.ToInt32(DT3.Rows[x]["InsetColorID"]), "Height", DataViewRowState.CurrentRows))
                        {
                            DT4 = DV.ToTable(true, new string[] { "Height", "Width" });
                        }
                        for (int y = 0; y < DT4.Rows.Count; y++)
                        {
                            DataRow[] Srows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) + " AND PatinaID=" + Convert.ToInt32(DT1.Rows[i]["PatinaID"]) +
                                " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) +
                                " AND InsetColorID=" + Convert.ToInt32(DT3.Rows[x]["InsetColorID"]) +
                                " AND Height=" + Convert.ToInt32(DT4.Rows[y]["Height"]));
                            if (Srows.Count() == 0)
                                continue;

                            int Count = 0;
                            foreach (DataRow item in Srows)
                                Count += Convert.ToInt32(item["Count"]);

                            DataRow[] rows = DestinationDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) + " AND PatinaID=" + Convert.ToInt32(DT1.Rows[i]["PatinaID"]) +
                                " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) +
                                " AND InsetColorID=" + Convert.ToInt32(DT3.Rows[x]["InsetColorID"]) +
                                " AND Height=" + Convert.ToInt32(DT4.Rows[y]["Height"]));
                            if (rows.Count() == 0)
                            {
                                DataRow NewRow = DestinationDT.NewRow();
                                NewRow["Front"] = GetFrontName(Convert.ToInt32(Srows[0]["FrontID"]));
                                if (Convert.ToInt32(DT1.Rows[i]["PatinaID"]) == -1)
                                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                                else
                                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"])) + "+" + GetPatinaName(Convert.ToInt32(DT1.Rows[i]["PatinaID"]));
                                NewRow["InsetType"] = GetInsetTypeName(Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]));
                                NewRow["InsetColor"] = GetInsetColorName(Convert.ToInt32(DT3.Rows[x]["InsetColorID"]));
                                NewRow["Height"] = Convert.ToInt32(DT4.Rows[y]["Height"]);
                                NewRow["Count"] = Count;
                                NewRow["FrontID"] = Convert.ToInt32(Srows[0]["FrontID"]);
                                NewRow["ColorID"] = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                                NewRow["InsetTypeID"] = Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]);
                                NewRow["InsetColorID"] = Convert.ToInt32(DT3.Rows[x]["InsetColorID"]);
                                DestinationDT.Rows.Add(NewRow);
                            }
                            else
                            {
                                rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count;
                            }
                        }
                    }
                }
            }
            DT1.Dispose();
            DT2.Dispose();
            DT3.Dispose();
            DT4.Dispose();
        }

        private void CollectDakotaAssemblySimple(int ColorID, DataTable SourceDT, ref DataTable DestinationDT, int FrontType)
        {
            DataTable DT2 = new DataTable();

            using (DataView DV = new DataView(SourceDT, "InsetTypeID=1 AND ColorID=" + ColorID, "ColorID, PatinaID, InsetColorID, Height, Width", DataViewRowState.CurrentRows))
            {
                DT2 = DV.ToTable(true, new string[] { "TechnoProfileID", "ColorID", "PatinaID", "InsetTypeID", "InsetColorID", "Height", "Width", "Notes" });
            }
            for (int j = 0; j < DT2.Rows.Count; j++)
            {
                decimal Square = 0;
                int Count = 0;
                //витрины
                string filter1 = "ColorID=" + Convert.ToInt32(DT2.Rows[j]["ColorID"]) + " AND PatinaID=" + Convert.ToInt32(DT2.Rows[j]["PatinaID"]) + " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) +
                            " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                            " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) + " AND (Notes='' OR Notes IS NULL)";

                if (DT2.Rows[j]["Notes"] != DBNull.Value && DT2.Rows[j]["Notes"].ToString().Length > 0)
                    filter1 = "ColorID=" + Convert.ToInt32(DT2.Rows[j]["ColorID"]) + " AND PatinaID=" + Convert.ToInt32(DT2.Rows[j]["PatinaID"]) + " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) +
                    " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                    " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) + " AND Notes='" + DT2.Rows[j]["Notes"] + "'";

                DataRow[] rows = SourceDT.Select(filter1);
                foreach (DataRow item in rows)
                {
                    Count += Convert.ToInt32(item["Count"]);
                    Square += Convert.ToDecimal(item["Height"]) * Convert.ToDecimal(item["Width"]) * Convert.ToDecimal(item["Count"]) / 1000000;
                }
                Square = decimal.Round(Square, 3, MidpointRounding.AwayFromZero);

                DataRow NewRow = DestinationDT.NewRow();
                NewRow["FrontType"] = FrontType;
                NewRow["ColorType"] = ColorID;
                NewRow["Name"] = ProfileName(Convert.ToInt32(rows[0]["FrontConfigID"]));
                if (Convert.ToInt32(rows[0]["PatinaID"]) == -1)
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"]));
                else
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"])) + "+" + GetPatinaName(Convert.ToInt32(rows[0]["PatinaID"]));
                //NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"]));
                string InsetColor = "Витрина";
                if (Convert.ToInt32(DT2.Rows[j]["TechnoProfileID"]) != -1)
                    InsetColor = "Витрина Г-170";
                NewRow["InsetColor"] = InsetColor;
                NewRow["Height"] = rows[0]["Height"];
                NewRow["Width"] = rows[0]["Width"];
                NewRow["Count"] = Count;
                NewRow["Square"] = Square;
                DestinationDT.Rows.Add(NewRow);
            }

            DataRow[] irows = InsetTypesDataTable.Select("GroupID = 7");
            string filter = string.Empty;
            foreach (DataRow item in irows)
                filter += item["InsetTypeID"].ToString() + ",";
            if (filter.Length > 0)
                filter = "InsetTypeID IN (" + filter.Substring(0, filter.Length - 1) + ")";
            using (DataView DV = new DataView(SourceDT, filter + " AND ColorID=" + ColorID, "ColorID, InsetColorID, Height, Width", DataViewRowState.CurrentRows))
            {
                DT2 = DV.ToTable(true, new string[] { "ColorID", "PatinaID", "InsetTypeID", "InsetColorID", "Height", "Width", "Notes" });
            }

            //using (DataView DV = new DataView(SourceDT, "InsetTypeID IN (2069,2070,2071,2073,2075,2077,2233,3644,29043,29531) AND ColorID=" + ColorID, "ColorID, InsetColorID, Height, Width", DataViewRowState.CurrentRows))
            //{
            //    DT2 = DV.ToTable(true, new string[] { "ColorID", "PatinaID", "InsetTypeID", "InsetColorID", "Height", "Width", "Notes" });
            //}
            for (int j = 0; j < DT2.Rows.Count; j++)
            {
                decimal Square = 0;
                int Count = 0;
                //филенки
                string filter1 = "ColorID=" + Convert.ToInt32(DT2.Rows[j]["ColorID"]) + " AND PatinaID=" + Convert.ToInt32(DT2.Rows[j]["PatinaID"]) + " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) +
                            " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                            " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) + " AND (Notes='' OR Notes IS NULL)";

                if (DT2.Rows[j]["Notes"] != DBNull.Value && DT2.Rows[j]["Notes"].ToString().Length > 0)
                    filter1 = "ColorID=" + Convert.ToInt32(DT2.Rows[j]["ColorID"]) + " AND PatinaID=" + Convert.ToInt32(DT2.Rows[j]["PatinaID"]) + " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) +
                    " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                    " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) + " AND Notes='" + DT2.Rows[j]["Notes"] + "'";

                DataRow[] rows = SourceDT.Select(filter1);
                foreach (DataRow item in rows)
                {
                    Count += Convert.ToInt32(item["Count"]);
                    Square += Convert.ToDecimal(item["Height"]) * Convert.ToDecimal(item["Width"]) * Convert.ToDecimal(item["Count"]) / 1000000;
                }
                Square = decimal.Round(Square, 3, MidpointRounding.AwayFromZero);

                DataRow NewRow = DestinationDT.NewRow();
                NewRow["FrontType"] = FrontType;
                NewRow["ColorType"] = ColorID;
                NewRow["Name"] = ProfileName(Convert.ToInt32(rows[0]["FrontConfigID"]));
                if (Convert.ToInt32(rows[0]["PatinaID"]) == -1)
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"]));
                else
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"])) + "+" + GetPatinaName(Convert.ToInt32(rows[0]["PatinaID"]));
                NewRow["InsetColor"] = GetInsetTypeName(Convert.ToInt32(rows[0]["InsetTypeID"])) + " " + GetInsetColorName(Convert.ToInt32(rows[0]["InsetColorID"]));
                NewRow["Height"] = rows[0]["Height"];
                NewRow["Width"] = rows[0]["Width"];
                NewRow["Count"] = Count;
                NewRow["Square"] = Square;
                DestinationDT.Rows.Add(NewRow);
            }
            irows = InsetTypesDataTable.Select("GroupID = 3 OR GroupID = 4");
            filter = string.Empty;
            foreach (DataRow item in irows)
                filter += item["InsetTypeID"].ToString() + ",";
            if (filter.Length > 0)
                filter = "InsetTypeID IN (" + filter.Substring(0, filter.Length - 1) + ")";
            using (DataView DV = new DataView(SourceDT, filter + " AND ColorID=" + ColorID, "ColorID, InsetColorID, Height, Width", DataViewRowState.CurrentRows))
            {
                DT2 = DV.ToTable(true, new string[] { "ColorID", "PatinaID", "InsetTypeID", "InsetColorID", "Height", "Width", "Notes" });
            }
            for (int j = 0; j < DT2.Rows.Count; j++)
            {
                decimal Square = 0;
                int Count = 0;
                //глухие
                string filter1 = "ColorID=" + Convert.ToInt32(DT2.Rows[j]["ColorID"]) + " AND PatinaID=" + Convert.ToInt32(DT2.Rows[j]["PatinaID"]) + " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) +
                            " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                            " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) + " AND (Notes='' OR Notes IS NULL)";

                if (DT2.Rows[j]["Notes"] != DBNull.Value && DT2.Rows[j]["Notes"].ToString().Length > 0)
                    filter1 = "ColorID=" + Convert.ToInt32(DT2.Rows[j]["ColorID"]) + " AND PatinaID=" + Convert.ToInt32(DT2.Rows[j]["PatinaID"]) + " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) +
                    " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                    " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) + " AND Notes='" + DT2.Rows[j]["Notes"] + "'";

                DataRow[] rows = SourceDT.Select(filter1);
                foreach (DataRow item in rows)
                {
                    Count += Convert.ToInt32(item["Count"]);
                    Square += Convert.ToDecimal(item["Height"]) * Convert.ToDecimal(item["Width"]) * Convert.ToDecimal(item["Count"]) / 1000000;
                }
                Square = decimal.Round(Square, 3, MidpointRounding.AwayFromZero);

                DataRow NewRow = DestinationDT.NewRow();
                NewRow["FrontType"] = FrontType;
                NewRow["ColorType"] = ColorID;
                NewRow["Name"] = ProfileName(Convert.ToInt32(rows[0]["FrontConfigID"]));
                if (Convert.ToInt32(rows[0]["PatinaID"]) == -1)
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"]));
                else
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"])) + "+" + GetPatinaName(Convert.ToInt32(rows[0]["PatinaID"]));
                NewRow["InsetColor"] = GetInsetTypeName(Convert.ToInt32(rows[0]["InsetTypeID"])) + " " + GetInsetColorName(Convert.ToInt32(rows[0]["InsetColorID"]));
                NewRow["Height"] = rows[0]["Height"];
                NewRow["Width"] = rows[0]["Width"];
                NewRow["Count"] = Count;
                NewRow["Square"] = Square;
                DestinationDT.Rows.Add(NewRow);
            }
        }

        private void CollectAssemblySimple(int ColorID, DataTable SourceDT, ref DataTable DestinationDT, int FrontType)
        {
            DataTable DT2 = new DataTable();

            using (DataView DV = new DataView(SourceDT, "InsetTypeID=1 AND ColorID=" + ColorID, "ColorID, PatinaID, InsetColorID, Height, Width", DataViewRowState.CurrentRows))
            {
                DT2 = DV.ToTable(true, new string[] { "ColorID", "PatinaID", "InsetTypeID", "InsetColorID", "Height", "Width", "Notes" });
            }
            for (int j = 0; j < DT2.Rows.Count; j++)
            {
                decimal Square = 0;
                int Count = 0;
                //витрины
                string filter1 = "ColorID=" + Convert.ToInt32(DT2.Rows[j]["ColorID"]) + " AND PatinaID=" + Convert.ToInt32(DT2.Rows[j]["PatinaID"]) + " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) +
                            " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                            " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) + " AND (Notes='' OR Notes IS NULL)";

                if (DT2.Rows[j]["Notes"] != DBNull.Value && DT2.Rows[j]["Notes"].ToString().Length > 0)
                    filter1 = "ColorID=" + Convert.ToInt32(DT2.Rows[j]["ColorID"]) + " AND PatinaID=" + Convert.ToInt32(DT2.Rows[j]["PatinaID"]) + " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) +
                    " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                    " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) + " AND Notes='" + DT2.Rows[j]["Notes"] + "'";

                DataRow[] rows = SourceDT.Select(filter1);
                foreach (DataRow item in rows)
                {
                    Count += Convert.ToInt32(item["Count"]);
                    Square += Convert.ToDecimal(item["Height"]) * Convert.ToDecimal(item["Width"]) * Convert.ToDecimal(item["Count"]) / 1000000;
                }
                Square = decimal.Round(Square, 3, MidpointRounding.AwayFromZero);

                DataRow NewRow = DestinationDT.NewRow();
                NewRow["FrontType"] = FrontType;
                NewRow["ColorType"] = ColorID;
                NewRow["Name"] = ProfileName(Convert.ToInt32(rows[0]["FrontConfigID"]));
                if (Convert.ToInt32(rows[0]["PatinaID"]) == -1)
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"]));
                else
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"])) + "+" + GetPatinaName(Convert.ToInt32(rows[0]["PatinaID"]));
                //NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"]));
                NewRow["InsetColor"] = "Витрина";
                NewRow["Height"] = rows[0]["Height"];
                NewRow["Width"] = rows[0]["Width"];
                NewRow["Count"] = Count;
                NewRow["Square"] = Square;
                DestinationDT.Rows.Add(NewRow);
            }

            DataRow[] irows = InsetTypesDataTable.Select("GroupID = 7");
            string filter = string.Empty;
            foreach (DataRow item in irows)
                filter += item["InsetTypeID"].ToString() + ",";
            if (filter.Length > 0)
                filter = "InsetTypeID IN (" + filter.Substring(0, filter.Length - 1) + ")";
            using (DataView DV = new DataView(SourceDT, filter + " AND ColorID=" + ColorID, "ColorID, InsetColorID, Height, Width", DataViewRowState.CurrentRows))
            {
                DT2 = DV.ToTable(true, new string[] { "ColorID", "PatinaID", "InsetTypeID", "InsetColorID", "Height", "Width", "Notes" });
            }

            //using (DataView DV = new DataView(SourceDT, "InsetTypeID IN (2069,2070,2071,2073,2075,2077,2233,3644,29043,29531) AND ColorID=" + ColorID, "ColorID, InsetColorID, Height, Width", DataViewRowState.CurrentRows))
            //{
            //    DT2 = DV.ToTable(true, new string[] { "ColorID", "PatinaID", "InsetTypeID", "InsetColorID", "Height", "Width", "Notes" });
            //}
            for (int j = 0; j < DT2.Rows.Count; j++)
            {
                decimal Square = 0;
                int Count = 0;
                //филенки
                string filter1 = "ColorID=" + Convert.ToInt32(DT2.Rows[j]["ColorID"]) + " AND PatinaID=" + Convert.ToInt32(DT2.Rows[j]["PatinaID"]) + " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) +
                            " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                            " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) + " AND (Notes='' OR Notes IS NULL)";

                if (DT2.Rows[j]["Notes"] != DBNull.Value && DT2.Rows[j]["Notes"].ToString().Length > 0)
                    filter1 = "ColorID=" + Convert.ToInt32(DT2.Rows[j]["ColorID"]) + " AND PatinaID=" + Convert.ToInt32(DT2.Rows[j]["PatinaID"]) + " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) +
                    " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                    " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) + " AND Notes='" + DT2.Rows[j]["Notes"] + "'";

                DataRow[] rows = SourceDT.Select(filter1);
                foreach (DataRow item in rows)
                {
                    Count += Convert.ToInt32(item["Count"]);
                    Square += Convert.ToDecimal(item["Height"]) * Convert.ToDecimal(item["Width"]) * Convert.ToDecimal(item["Count"]) / 1000000;
                }
                Square = decimal.Round(Square, 3, MidpointRounding.AwayFromZero);

                DataRow NewRow = DestinationDT.NewRow();
                NewRow["FrontType"] = FrontType;
                NewRow["ColorType"] = ColorID;
                NewRow["Name"] = ProfileName(Convert.ToInt32(rows[0]["FrontConfigID"]));
                if (Convert.ToInt32(rows[0]["PatinaID"]) == -1)
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"]));
                else
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"])) + "+" + GetPatinaName(Convert.ToInt32(rows[0]["PatinaID"]));
                NewRow["InsetColor"] = GetInsetTypeName(Convert.ToInt32(rows[0]["InsetTypeID"])) + " " + GetInsetColorName(Convert.ToInt32(rows[0]["InsetColorID"]));
                NewRow["Height"] = rows[0]["Height"];
                NewRow["Width"] = rows[0]["Width"];
                NewRow["Count"] = Count;
                NewRow["Square"] = Square;
                DestinationDT.Rows.Add(NewRow);
            }
            irows = InsetTypesDataTable.Select("GroupID = 3 OR GroupID = 4");
            filter = string.Empty;
            foreach (DataRow item in irows)
                filter += item["InsetTypeID"].ToString() + ",";
            if (filter.Length > 0)
                filter = "InsetTypeID IN (" + filter.Substring(0, filter.Length - 1) + ")";
            using (DataView DV = new DataView(SourceDT, filter + " AND ColorID=" + ColorID, "ColorID, InsetColorID, Height, Width", DataViewRowState.CurrentRows))
            {
                DT2 = DV.ToTable(true, new string[] { "ColorID", "PatinaID", "InsetTypeID", "InsetColorID", "Height", "Width", "Notes" });
            }
            for (int j = 0; j < DT2.Rows.Count; j++)
            {
                decimal Square = 0;
                int Count = 0;
                //глухие
                string filter1 = "ColorID=" + Convert.ToInt32(DT2.Rows[j]["ColorID"]) + " AND PatinaID=" + Convert.ToInt32(DT2.Rows[j]["PatinaID"]) + " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) +
                            " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                            " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) + " AND (Notes='' OR Notes IS NULL)";

                if (DT2.Rows[j]["Notes"] != DBNull.Value && DT2.Rows[j]["Notes"].ToString().Length > 0)
                    filter1 = "ColorID=" + Convert.ToInt32(DT2.Rows[j]["ColorID"]) + " AND PatinaID=" + Convert.ToInt32(DT2.Rows[j]["PatinaID"]) + " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) +
                    " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                    " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) + " AND Notes='" + DT2.Rows[j]["Notes"] + "'";

                DataRow[] rows = SourceDT.Select(filter1);
                foreach (DataRow item in rows)
                {
                    Count += Convert.ToInt32(item["Count"]);
                    Square += Convert.ToDecimal(item["Height"]) * Convert.ToDecimal(item["Width"]) * Convert.ToDecimal(item["Count"]) / 1000000;
                }
                Square = decimal.Round(Square, 3, MidpointRounding.AwayFromZero);

                DataRow NewRow = DestinationDT.NewRow();
                NewRow["FrontType"] = FrontType;
                NewRow["ColorType"] = ColorID;
                NewRow["Name"] = ProfileName(Convert.ToInt32(rows[0]["FrontConfigID"]));
                if (Convert.ToInt32(rows[0]["PatinaID"]) == -1)
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"]));
                else
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"])) + "+" + GetPatinaName(Convert.ToInt32(rows[0]["PatinaID"]));
                NewRow["InsetColor"] = GetInsetTypeName(Convert.ToInt32(rows[0]["InsetTypeID"])) + " " + GetInsetColorName(Convert.ToInt32(rows[0]["InsetColorID"]));
                NewRow["Height"] = rows[0]["Height"];
                NewRow["Width"] = rows[0]["Width"];
                NewRow["Count"] = Count;
                NewRow["Square"] = Square;
                DestinationDT.Rows.Add(NewRow);
            }
        }

        private void CollectAssemblyBoxes(int ColorID, DataTable SourceDT, ref DataTable DestinationDT, int FrontType)
        {
            DataTable DT2 = new DataTable();

            using (DataView DV = new DataView(SourceDT, "InsetTypeID=1 AND ColorID=" + ColorID, "ColorID, InsetColorID, Height, Width", DataViewRowState.CurrentRows))
            {
                DT2 = DV.ToTable(true, new string[] { "ColorID", "PatinaID", "InsetTypeID", "InsetColorID", "Height", "Width", "Notes" });
            }
            for (int j = 0; j < DT2.Rows.Count; j++)
            {
                decimal Square = 0;
                int Count = 0;
                //витрины
                string filter1 = "ColorID=" + Convert.ToInt32(DT2.Rows[j]["ColorID"]) + " AND PatinaID=" + Convert.ToInt32(DT2.Rows[j]["PatinaID"]) + " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) +
                            " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                            " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) + " AND (Notes='' OR Notes IS NULL)";

                if (DT2.Rows[j]["Notes"] != DBNull.Value && DT2.Rows[j]["Notes"].ToString().Length > 0)
                    filter1 = "ColorID=" + Convert.ToInt32(DT2.Rows[j]["ColorID"]) + " AND PatinaID=" + Convert.ToInt32(DT2.Rows[j]["PatinaID"]) + " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) +
                    " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                    " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) + " AND Notes='" + DT2.Rows[j]["Notes"] + "'";

                DataRow[] rows = SourceDT.Select(filter1);
                foreach (DataRow item in rows)
                {
                    Count += Convert.ToInt32(item["Count"]);
                    Square += Convert.ToDecimal(item["Height"]) * Convert.ToDecimal(item["Width"]) * Convert.ToDecimal(item["Count"]) / 1000000;
                }
                Square = decimal.Round(Square, 3, MidpointRounding.AwayFromZero);

                DataRow NewRow = DestinationDT.NewRow();
                NewRow["FrontType"] = FrontType;
                NewRow["ColorType"] = ColorID;
                NewRow["Name"] = ProfileName(Convert.ToInt32(rows[0]["FrontConfigID"])) + " ШУФ";
                if (Convert.ToInt32(rows[0]["PatinaID"]) == -1)
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"]));
                else
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"])) + "+" + GetPatinaName(Convert.ToInt32(rows[0]["PatinaID"]));
                NewRow["InsetColor"] = "Витрина";
                NewRow["Height"] = rows[0]["Height"];
                NewRow["Width"] = rows[0]["Width"];
                NewRow["Count"] = Count;
                NewRow["Square"] = Square;
                DestinationDT.Rows.Add(NewRow);
            }

            DataRow[] irows = InsetTypesDataTable.Select("GroupID = 7");
            string filter = string.Empty;
            foreach (DataRow item in irows)
                filter += item["InsetTypeID"].ToString() + ",";
            if (filter.Length > 0)
                filter = "InsetTypeID IN (" + filter.Substring(0, filter.Length - 1) + ")";
            using (DataView DV = new DataView(SourceDT, filter + " AND ColorID=" + ColorID, "ColorID, InsetColorID, Height, Width", DataViewRowState.CurrentRows))
            {
                DT2 = DV.ToTable(true, new string[] { "ColorID", "PatinaID", "InsetTypeID", "InsetColorID", "Height", "Width", "Notes" });
            }

            //using (DataView DV = new DataView(SourceDT, "InsetTypeID IN (2069,2070,2071,2073,2075,2077,2233,3644,29043,29531) AND ColorID=" + ColorID, "ColorID, InsetColorID, Height, Width", DataViewRowState.CurrentRows))
            //{
            //    DT2 = DV.ToTable(true, new string[] { "ColorID", "PatinaID", "InsetTypeID", "InsetColorID", "Height", "Width", "Notes" });
            //}
            for (int j = 0; j < DT2.Rows.Count; j++)
            {
                decimal Square = 0;
                int Count = 0;
                //филенки
                string filter1 = "ColorID=" + Convert.ToInt32(DT2.Rows[j]["ColorID"]) + " AND PatinaID=" + Convert.ToInt32(DT2.Rows[j]["PatinaID"]) + " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) +
                            " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                            " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) + " AND (Notes='' OR Notes IS NULL)";

                if (DT2.Rows[j]["Notes"] != DBNull.Value && DT2.Rows[j]["Notes"].ToString().Length > 0)
                    filter1 = "ColorID=" + Convert.ToInt32(DT2.Rows[j]["ColorID"]) + " AND PatinaID=" + Convert.ToInt32(DT2.Rows[j]["PatinaID"]) + " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) +
                    " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                    " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) + " AND Notes='" + DT2.Rows[j]["Notes"] + "'";

                DataRow[] rows = SourceDT.Select(filter1);
                foreach (DataRow item in rows)
                {
                    Count += Convert.ToInt32(item["Count"]);
                    Square += Convert.ToDecimal(item["Height"]) * Convert.ToDecimal(item["Width"]) * Convert.ToDecimal(item["Count"]) / 1000000;
                }
                Square = decimal.Round(Square, 3, MidpointRounding.AwayFromZero);

                DataRow NewRow = DestinationDT.NewRow();
                NewRow["FrontType"] = FrontType;
                NewRow["ColorType"] = ColorID;
                NewRow["Name"] = ProfileName(Convert.ToInt32(rows[0]["FrontConfigID"])) + " ШУФ";
                if (Convert.ToInt32(rows[0]["PatinaID"]) == -1)
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"]));
                else
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"])) + "+" + GetPatinaName(Convert.ToInt32(rows[0]["PatinaID"]));
                NewRow["InsetColor"] = GetInsetTypeName(Convert.ToInt32(rows[0]["InsetTypeID"])) + " " + GetInsetColorName(Convert.ToInt32(rows[0]["InsetColorID"]));
                NewRow["Height"] = rows[0]["Height"];
                NewRow["Width"] = rows[0]["Width"];
                NewRow["Count"] = Count;
                NewRow["Square"] = Square;
                DestinationDT.Rows.Add(NewRow);
            }
            irows = InsetTypesDataTable.Select("GroupID = 3 OR GroupID = 4");
            filter = string.Empty;
            foreach (DataRow item in irows)
                filter += item["InsetTypeID"].ToString() + ",";
            if (filter.Length > 0)
                filter = "InsetTypeID IN (" + filter.Substring(0, filter.Length - 1) + ")";
            using (DataView DV = new DataView(SourceDT, filter + " AND ColorID=" + ColorID, "ColorID, InsetColorID, Height, Width", DataViewRowState.CurrentRows))
            {
                DT2 = DV.ToTable(true, new string[] { "ColorID", "PatinaID", "InsetTypeID", "InsetColorID", "Height", "Width", "Notes" });
            }
            for (int j = 0; j < DT2.Rows.Count; j++)
            {
                decimal Square = 0;
                int Count = 0;
                //глухие
                string filter1 = "ColorID=" + Convert.ToInt32(DT2.Rows[j]["ColorID"]) + " AND PatinaID=" + Convert.ToInt32(DT2.Rows[j]["PatinaID"]) + " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) +
                            " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                            " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) + " AND (Notes='' OR Notes IS NULL)";

                if (DT2.Rows[j]["Notes"] != DBNull.Value && DT2.Rows[j]["Notes"].ToString().Length > 0)
                    filter1 = "ColorID=" + Convert.ToInt32(DT2.Rows[j]["ColorID"]) + " AND PatinaID=" + Convert.ToInt32(DT2.Rows[j]["PatinaID"]) + " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) +
                    " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                    " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) + " AND Notes='" + DT2.Rows[j]["Notes"] + "'";

                DataRow[] rows = SourceDT.Select(filter1);
                foreach (DataRow item in rows)
                {
                    Count += Convert.ToInt32(item["Count"]);
                    Square += Convert.ToDecimal(item["Height"]) * Convert.ToDecimal(item["Width"]) * Convert.ToDecimal(item["Count"]) / 1000000;
                }
                Square = decimal.Round(Square, 3, MidpointRounding.AwayFromZero);

                DataRow NewRow = DestinationDT.NewRow();
                NewRow["FrontType"] = FrontType;
                NewRow["ColorType"] = ColorID;
                NewRow["Name"] = ProfileName(Convert.ToInt32(rows[0]["FrontConfigID"])) + " ШУФ";
                if (Convert.ToInt32(rows[0]["PatinaID"]) == -1)
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"]));
                else
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"])) + "+" + GetPatinaName(Convert.ToInt32(rows[0]["PatinaID"]));
                NewRow["InsetColor"] = GetInsetTypeName(Convert.ToInt32(rows[0]["InsetTypeID"])) + " " + GetInsetColorName(Convert.ToInt32(rows[0]["InsetColorID"]));
                NewRow["Height"] = rows[0]["Height"];
                NewRow["Width"] = rows[0]["Width"];
                NewRow["Count"] = Count;
                NewRow["Square"] = Square;
                DestinationDT.Rows.Add(NewRow);
            }
        }

        private void CollectAssemblyGrids(int ColorID, DataTable SourceDT, ref DataTable DestinationDT, int FrontType)
        {
            DataTable DT = new DataTable();

            using (DataView DV = new DataView(SourceDT, "ColorID=" + ColorID, "ColorID, InsetTypeID, InsetColorID, Height, Width", DataViewRowState.CurrentRows))
            {
                DT = DV.ToTable(true, new string[] { "ColorID", "PatinaID", "InsetTypeID", "InsetColorID", "Height", "Width", "Notes" });
            }
            for (int i = 0; i < DT.Rows.Count; i++)
            {
                decimal Square = 0;
                int Count = 0;
                //филенки
                string filter1 = "ColorID=" + Convert.ToInt32(DT.Rows[i]["ColorID"]) + " AND PatinaID=" + Convert.ToInt32(DT.Rows[i]["PatinaID"]) +
                    " AND InsetTypeID=" + Convert.ToInt32(DT.Rows[i]["InsetTypeID"]) +
                    " AND InsetColorID=" + Convert.ToInt32(DT.Rows[i]["InsetColorID"]) + " AND Height=" + Convert.ToInt32(DT.Rows[i]["Height"]) +
                    " AND Width=" + Convert.ToInt32(DT.Rows[i]["Width"]) + " AND (Notes='' OR Notes IS NULL)";

                if (DT.Rows[i]["Notes"] != DBNull.Value && DT.Rows[i]["Notes"].ToString().Length > 0)
                    filter1 = "ColorID=" + Convert.ToInt32(DT.Rows[i]["ColorID"]) + " AND PatinaID=" + Convert.ToInt32(DT.Rows[i]["PatinaID"]) + " AND InsetTypeID=" + Convert.ToInt32(DT.Rows[i]["InsetTypeID"]) +
                    " AND InsetColorID=" + Convert.ToInt32(DT.Rows[i]["InsetColorID"]) + " AND Height=" + Convert.ToInt32(DT.Rows[i]["Height"]) +
                    " AND Width=" + Convert.ToInt32(DT.Rows[i]["Width"]) + " AND Notes='" + DT.Rows[i]["Notes"] + "'";

                DataRow[] rows = SourceDT.Select(filter1);
                foreach (DataRow item in rows)
                {
                    Count += Convert.ToInt32(item["Count"]);
                    Square += Convert.ToDecimal(item["Height"]) * Convert.ToDecimal(item["Width"]) * Convert.ToDecimal(item["Count"]) / 1000000;
                }
                Square = decimal.Round(Square, 3, MidpointRounding.AwayFromZero);

                DataRow NewRow = DestinationDT.NewRow();
                NewRow["FrontType"] = FrontType;
                NewRow["ColorType"] = ColorID;
                NewRow["Name"] = ProfileName(Convert.ToInt32(rows[0]["FrontConfigID"])) + " РЕШ";
                if (Convert.ToInt32(rows[0]["PatinaID"]) == -1)
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"]));
                else
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"])) + "+" + GetPatinaName(Convert.ToInt32(rows[0]["PatinaID"]));
                NewRow["InsetColor"] = GetInsetTypeName(Convert.ToInt32(rows[0]["InsetTypeID"])) + " " + GetInsetColorName(Convert.ToInt32(rows[0]["InsetColorID"]));
                NewRow["Height"] = rows[0]["Height"];
                NewRow["Width"] = rows[0]["Width"];
                NewRow["Count"] = Count;
                NewRow["Square"] = Square;
                DestinationDT.Rows.Add(NewRow);
            }
        }

        private void CollectAssemblyApplique(int ColorID, DataTable SourceDT, ref DataTable DestinationDT, int FrontType)
        {
            DataTable DT = new DataTable();

            using (DataView DV = new DataView(SourceDT, "ColorID=" + ColorID, "ColorID, InsetTypeID, InsetColorID, Height, Width", DataViewRowState.CurrentRows))
            {
                DT = DV.ToTable(true, new string[] { "ColorID", "PatinaID", "InsetTypeID", "InsetColorID", "Height", "Width", "Notes" });
            }
            for (int i = 0; i < DT.Rows.Count; i++)
            {
                decimal Square = 0;
                int Count = 0;
                //филенки
                string filter1 = "ColorID=" + Convert.ToInt32(DT.Rows[i]["ColorID"]) + " AND PatinaID=" + Convert.ToInt32(DT.Rows[i]["PatinaID"]) +
                    " AND InsetTypeID=" + Convert.ToInt32(DT.Rows[i]["InsetTypeID"]) +
                    " AND InsetColorID=" + Convert.ToInt32(DT.Rows[i]["InsetColorID"]) + " AND Height=" + Convert.ToInt32(DT.Rows[i]["Height"]) +
                    " AND Width=" + Convert.ToInt32(DT.Rows[i]["Width"]) + " AND (Notes='' OR Notes IS NULL)";

                if (DT.Rows[i]["Notes"] != DBNull.Value && DT.Rows[i]["Notes"].ToString().Length > 0)
                    filter1 = "ColorID=" + Convert.ToInt32(DT.Rows[i]["ColorID"]) + " AND PatinaID=" + Convert.ToInt32(DT.Rows[i]["PatinaID"]) + " AND InsetTypeID=" + Convert.ToInt32(DT.Rows[i]["InsetTypeID"]) +
                    " AND InsetColorID=" + Convert.ToInt32(DT.Rows[i]["InsetColorID"]) + " AND Height=" + Convert.ToInt32(DT.Rows[i]["Height"]) +
                    " AND Width=" + Convert.ToInt32(DT.Rows[i]["Width"]) + " AND Notes='" + DT.Rows[i]["Notes"] + "'";

                DataRow[] rows = SourceDT.Select(filter1);
                foreach (DataRow item in rows)
                {
                    Count += Convert.ToInt32(item["Count"]);
                    Square += Convert.ToDecimal(item["Height"]) * Convert.ToDecimal(item["Width"]) * Convert.ToDecimal(item["Count"]) / 1000000;
                }
                Square = decimal.Round(Square, 3, MidpointRounding.AwayFromZero);

                DataRow NewRow = DestinationDT.NewRow();
                NewRow["FrontType"] = FrontType;
                NewRow["ColorType"] = ColorID;
                NewRow["Name"] = ProfileName(Convert.ToInt32(rows[0]["FrontConfigID"])) + " Аппл";
                if (Convert.ToInt32(rows[0]["PatinaID"]) == -1)
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"]));
                else
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"])) + "+" + GetPatinaName(Convert.ToInt32(rows[0]["PatinaID"]));
                NewRow["InsetColor"] = GetInsetTypeName(Convert.ToInt32(rows[0]["InsetTypeID"])) + " " + GetInsetColorName(Convert.ToInt32(rows[0]["InsetColorID"]));
                NewRow["Height"] = rows[0]["Height"];
                NewRow["Width"] = rows[0]["Width"];
                NewRow["Count"] = Count;
                NewRow["Square"] = Square;
                DestinationDT.Rows.Add(NewRow);
            }
        }

        private void CollectDakotaDeying(int ColorID, DataTable SourceDT, ref DataTable DestinationDT, string AdditionalName)
        {
            DataTable DT2 = new DataTable();
            //Витрины сначала
            using (DataView DV = new DataView(SourceDT, "InsetTypeID=1 AND ColorID=" + ColorID,
                "InsetTypeID, InsetColorID, Height, Width", DataViewRowState.CurrentRows))
            {
                DT2 = DV.ToTable(true, new string[] { "TechnoProfileID", "PatinaID", "InsetTypeID", "InsetColorID", "Height", "Width", "Notes" });
            }
            for (int j = 0; j < DT2.Rows.Count; j++)
            {
                decimal Square = 0;
                int Count = 0;
                string InsetColor = string.Empty;
                string filter1 = "ColorID=" + ColorID + " AND PatinaID=" + Convert.ToInt32(DT2.Rows[j]["PatinaID"]) + " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) +
                            " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                            " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) + " AND (Notes='' OR Notes IS NULL)";

                if (DT2.Rows[j]["Notes"] != DBNull.Value && DT2.Rows[j]["Notes"].ToString().Length > 0)
                    filter1 = "ColorID=" + ColorID + " AND PatinaID=" + Convert.ToInt32(DT2.Rows[j]["PatinaID"]) + " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) +
                    " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                    " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) + " AND Notes='" + DT2.Rows[j]["Notes"] + "'";

                DataRow[] rows = SourceDT.Select(filter1);
                foreach (DataRow item in rows)
                {
                    Count += Convert.ToInt32(item["Count"]);
                    Square += Convert.ToDecimal(item["Height"]) * Convert.ToDecimal(item["Width"]) * Convert.ToDecimal(item["Count"]) / 1000000;
                }
                Square = decimal.Round(Square, 3, MidpointRounding.AwayFromZero);

                DataRow NewRow = DestinationDT.NewRow();
                NewRow["ColorType"] = ColorID;
                NewRow["Name"] = ProfileName(Convert.ToInt32(rows[0]["FrontConfigID"])) + AdditionalName;
                if (Convert.ToInt32(rows[0]["PatinaID"]) == -1)
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"]));
                else
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"])) + "+" + GetPatinaName(Convert.ToInt32(rows[0]["PatinaID"]));
                if (Convert.ToInt32(rows[0]["InsetTypeID"]) == 1)
                {
                    InsetColor = "Витрина";
                    if (Convert.ToInt32(DT2.Rows[j]["TechnoProfileID"]) != -1)
                        InsetColor = "Витрина Г-170";
                    NewRow["InsetColor"] = InsetColor;
                }
                else
                    NewRow["InsetColor"] = GetInsetTypeName(Convert.ToInt32(rows[0]["InsetTypeID"])) + " " + GetInsetColorName(Convert.ToInt32(rows[0]["InsetColorID"]));
                NewRow["Height"] = rows[0]["Height"];
                NewRow["Width"] = rows[0]["Width"];
                NewRow["Count"] = Count;
                NewRow["Square"] = Square;
                NewRow["Notes"] = rows[0]["Notes"];
                DestinationDT.Rows.Add(NewRow);
            }

            using (DataView DV = new DataView(SourceDT, "InsetTypeID<>1 AND ColorID=" + ColorID,
                "InsetTypeID, InsetColorID, Height, Width", DataViewRowState.CurrentRows))
            {
                DT2 = DV.ToTable(true, new string[] { "PatinaID", "InsetTypeID", "InsetColorID", "Height", "Width", "Notes" });
            }
            for (int j = 0; j < DT2.Rows.Count; j++)
            {
                decimal Square = 0;
                int Count = 0;
                string InsetColor = string.Empty;
                string filter1 = "ColorID=" + ColorID + " AND PatinaID=" + Convert.ToInt32(DT2.Rows[j]["PatinaID"]) + " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) +
                            " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                            " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) + " AND (Notes='' OR Notes IS NULL)";

                if (DT2.Rows[j]["Notes"] != DBNull.Value && DT2.Rows[j]["Notes"].ToString().Length > 0)
                    filter1 = "ColorID=" + ColorID + " AND PatinaID=" + Convert.ToInt32(DT2.Rows[j]["PatinaID"]) + " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) +
                    " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                    " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) + " AND Notes='" + DT2.Rows[j]["Notes"] + "'";

                DataRow[] rows = SourceDT.Select(filter1);
                foreach (DataRow item in rows)
                {
                    Count += Convert.ToInt32(item["Count"]);
                    Square += Convert.ToDecimal(item["Height"]) * Convert.ToDecimal(item["Width"]) * Convert.ToDecimal(item["Count"]) / 1000000;
                }
                Square = decimal.Round(Square, 3, MidpointRounding.AwayFromZero);

                DataRow NewRow = DestinationDT.NewRow();
                NewRow["ColorType"] = ColorID;
                NewRow["Name"] = ProfileName(Convert.ToInt32(rows[0]["FrontConfigID"])) + AdditionalName;
                if (Convert.ToInt32(rows[0]["PatinaID"]) == -1)
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"]));
                else
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"])) + "+" + GetPatinaName(Convert.ToInt32(rows[0]["PatinaID"]));
                if (Convert.ToInt32(rows[0]["InsetTypeID"]) == 1)
                    NewRow["InsetColor"] = "Витрина";
                else
                    NewRow["InsetColor"] = GetInsetTypeName(Convert.ToInt32(rows[0]["InsetTypeID"])) + " " + GetInsetColorName(Convert.ToInt32(rows[0]["InsetColorID"]));
                NewRow["Height"] = rows[0]["Height"];
                NewRow["Width"] = rows[0]["Width"];
                NewRow["Count"] = Count;
                NewRow["Square"] = Square;
                NewRow["Notes"] = rows[0]["Notes"];
                DestinationDT.Rows.Add(NewRow);
            }
        }

        private void CollectDeying(int ColorID, DataTable SourceDT, ref DataTable DestinationDT, string AdditionalName)
        {
            DataTable DT2 = new DataTable();
            //Витрины сначала
            using (DataView DV = new DataView(SourceDT, "InsetTypeID=1 AND ColorID=" + ColorID,
                "InsetTypeID, InsetColorID, Height, Width", DataViewRowState.CurrentRows))
            {
                DT2 = DV.ToTable(true, new string[] { "PatinaID", "InsetTypeID", "InsetColorID", "Height", "Width", "Notes" });
            }
            for (int j = 0; j < DT2.Rows.Count; j++)
            {
                decimal Square = 0;
                int Count = 0;
                string InsetColor = string.Empty;
                string filter1 = "ColorID=" + ColorID + " AND PatinaID=" + Convert.ToInt32(DT2.Rows[j]["PatinaID"]) + " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) +
                            " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                            " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) + " AND (Notes='' OR Notes IS NULL)";

                if (DT2.Rows[j]["Notes"] != DBNull.Value && DT2.Rows[j]["Notes"].ToString().Length > 0)
                    filter1 = "ColorID=" + ColorID + " AND PatinaID=" + Convert.ToInt32(DT2.Rows[j]["PatinaID"]) + " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) +
                    " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                    " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) + " AND Notes='" + DT2.Rows[j]["Notes"] + "'";

                DataRow[] rows = SourceDT.Select(filter1);
                foreach (DataRow item in rows)
                {
                    Count += Convert.ToInt32(item["Count"]);
                    Square += Convert.ToDecimal(item["Height"]) * Convert.ToDecimal(item["Width"]) * Convert.ToDecimal(item["Count"]) / 1000000;
                }
                Square = decimal.Round(Square, 3, MidpointRounding.AwayFromZero);

                DataRow NewRow = DestinationDT.NewRow();
                NewRow["ColorType"] = ColorID;
                NewRow["Name"] = ProfileName(Convert.ToInt32(rows[0]["FrontConfigID"])) + AdditionalName;
                if (Convert.ToInt32(rows[0]["PatinaID"]) == -1)
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"]));
                else
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"])) + "+" + GetPatinaName(Convert.ToInt32(rows[0]["PatinaID"]));
                if (Convert.ToInt32(rows[0]["InsetTypeID"]) == 1)
                    NewRow["InsetColor"] = "Витрина";
                else
                    NewRow["InsetColor"] = GetInsetTypeName(Convert.ToInt32(rows[0]["InsetTypeID"])) + " " + GetInsetColorName(Convert.ToInt32(rows[0]["InsetColorID"]));
                NewRow["Height"] = rows[0]["Height"];
                NewRow["Width"] = rows[0]["Width"];
                NewRow["Count"] = Count;
                NewRow["Square"] = Square;
                NewRow["Notes"] = rows[0]["Notes"];
                DestinationDT.Rows.Add(NewRow);
            }

            using (DataView DV = new DataView(SourceDT, "InsetTypeID<>1 AND ColorID=" + ColorID,
                "InsetTypeID, InsetColorID, Height, Width", DataViewRowState.CurrentRows))
            {
                DT2 = DV.ToTable(true, new string[] { "PatinaID", "InsetTypeID", "InsetColorID", "Height", "Width", "Notes" });
            }
            for (int j = 0; j < DT2.Rows.Count; j++)
            {
                decimal Square = 0;
                int Count = 0;
                string InsetColor = string.Empty;
                string filter1 = "ColorID=" + ColorID + " AND PatinaID=" + Convert.ToInt32(DT2.Rows[j]["PatinaID"]) + " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) +
                            " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                            " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) + " AND (Notes='' OR Notes IS NULL)";

                if (DT2.Rows[j]["Notes"] != DBNull.Value && DT2.Rows[j]["Notes"].ToString().Length > 0)
                    filter1 = "ColorID=" + ColorID + " AND PatinaID=" + Convert.ToInt32(DT2.Rows[j]["PatinaID"]) + " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) +
                    " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) + " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]) +
                    " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Width"]) + " AND Notes='" + DT2.Rows[j]["Notes"] + "'";

                DataRow[] rows = SourceDT.Select(filter1);
                foreach (DataRow item in rows)
                {
                    Count += Convert.ToInt32(item["Count"]);
                    Square += Convert.ToDecimal(item["Height"]) * Convert.ToDecimal(item["Width"]) * Convert.ToDecimal(item["Count"]) / 1000000;
                }
                Square = decimal.Round(Square, 3, MidpointRounding.AwayFromZero);

                DataRow NewRow = DestinationDT.NewRow();
                NewRow["ColorType"] = ColorID;
                NewRow["Name"] = ProfileName(Convert.ToInt32(rows[0]["FrontConfigID"])) + AdditionalName;
                if (Convert.ToInt32(rows[0]["PatinaID"]) == -1)
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"]));
                else
                    NewRow["FrameColor"] = GetColorName(Convert.ToInt32(rows[0]["ColorID"])) + "+" + GetPatinaName(Convert.ToInt32(rows[0]["PatinaID"]));
                if (Convert.ToInt32(rows[0]["InsetTypeID"]) == 1)
                    NewRow["InsetColor"] = "Витрина";
                else
                    NewRow["InsetColor"] = GetInsetTypeName(Convert.ToInt32(rows[0]["InsetTypeID"])) + " " + GetInsetColorName(Convert.ToInt32(rows[0]["InsetColorID"]));
                NewRow["Height"] = rows[0]["Height"];
                NewRow["Width"] = rows[0]["Width"];
                NewRow["Count"] = Count;
                NewRow["Square"] = Square;
                NewRow["Notes"] = rows[0]["Notes"];
                DestinationDT.Rows.Add(NewRow);
            }
        }

        private void CollectDakotaVitrinaOrders(DataTable DistinctSizesDT, DataTable SourceDT, ref DataTable DestinationDT, int FrontType, string FrontName)
        {
            string ColName = string.Empty;
            string FrameColor = string.Empty;
            string InsetType = string.Empty;
            string InsetColor = string.Empty;

            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            DataTable DT3 = new DataTable();

            for (int y = 0; y < DistinctSizesDT.Rows.Count; y++)
            {
                using (DataView DV = new DataView(SourceDT))
                {
                    DT1 = DV.ToTable(true, new string[] { "ColorID", "PatinaID" });
                }
                for (int i = 0; i < DT1.Rows.Count; i++)
                {
                    using (DataView DV = new DataView(SourceDT, "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) + " AND PatinaID=" + Convert.ToInt32(DT1.Rows[i]["PatinaID"]), string.Empty, DataViewRowState.CurrentRows))
                    {
                        DT2 = DV.ToTable(true, new string[] { "InsetTypeID" });
                    }
                    for (int j = 0; j < DT2.Rows.Count; j++)
                    {
                        using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]), string.Empty, DataViewRowState.CurrentRows))
                        {
                            DT3 = DV.ToTable(true, new string[] { "InsetColorID" });
                            for (int x = 0; x < DT3.Rows.Count; x++)
                            {
                                DataRow[] rows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) + " AND PatinaID=" + Convert.ToInt32(DT1.Rows[i]["PatinaID"]) +
                                    " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) + " AND InsetColorID=" + Convert.ToInt32(DT3.Rows[x]["InsetColorID"]));

                                if (rows.Count() > 0)
                                {
                                    if (Convert.ToInt32(rows[0]["PatinaID"]) == -1)
                                        FrameColor = GetColorName(Convert.ToInt32(rows[0]["ColorID"]));
                                    else
                                        FrameColor = GetColorName(Convert.ToInt32(rows[0]["ColorID"])) + "+" + GetPatinaName(Convert.ToInt32(rows[0]["PatinaID"]));
                                    //FrameColor = GetColorName(Convert.ToInt32(rows[0]["ColorID"]));
                                    InsetType = GetInsetTypeName(Convert.ToInt32(rows[0]["InsetTypeID"]));
                                    if (Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) == 1)
                                    {
                                        InsetType = "Витрина";
                                        if (Convert.ToInt32(rows[0]["TechnoProfileID"]) != -1)
                                            InsetType = "Витрина Г-170";

                                        InsetColor = "Витрина";
                                    }
                                    else
                                        InsetColor = GetInsetColorName(Convert.ToInt32(rows[0]["InsetColorID"]));

                                    ColName = FrameColor + "(" + InsetType + " " + InsetColor + ")_" + FrontType;
                                    if (!DestinationDT.Columns.Contains(ColName))
                                        DestinationDT.Columns.Add(new DataColumn(ColName, Type.GetType("System.String")));

                                    DestinationDT.Rows[0][ColName] = FrontName;

                                    DataRow[] Srows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) + " AND PatinaID=" + Convert.ToInt32(DT1.Rows[i]["PatinaID"]) +
                                        " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) + " AND InsetColorID=" + Convert.ToInt32(DT3.Rows[x]["InsetColorID"]) +
                                        " AND Height=" + Convert.ToInt32(DistinctSizesDT.Rows[y]["Height"]) + " AND Width=" + Convert.ToInt32(DistinctSizesDT.Rows[y]["Width"]));
                                    if (Srows.Count() > 0)
                                    {
                                        int Count = 0;
                                        foreach (DataRow item in Srows)
                                            Count += Convert.ToInt32(item["Count"]);

                                        DataRow[] Drows = DestinationDT.Select("Sizes='" + DistinctSizesDT.Rows[y]["Height"].ToString() + " X " + DistinctSizesDT.Rows[y]["Width"].ToString() + "'");
                                        if (Drows.Count() == 0)
                                        {
                                            DataRow NewRow = DestinationDT.NewRow();
                                            NewRow["Sizes"] = DistinctSizesDT.Rows[y]["Height"].ToString() + " X " + DistinctSizesDT.Rows[y]["Width"].ToString();
                                            NewRow["Height"] = DistinctSizesDT.Rows[y]["Height"];
                                            NewRow["Width"] = DistinctSizesDT.Rows[y]["Width"];
                                            NewRow[ColName] = Count;
                                            DestinationDT.Rows.Add(NewRow);
                                        }
                                        else
                                        {
                                            Drows[0][ColName] = Count;
                                        }
                                    }
                                }
                                else
                                    continue;

                            }
                        }
                    }
                }
            }
        }

        private void CollectOrders(DataTable DistinctSizesDT, DataTable SourceDT, ref DataTable DestinationDT, int FrontType, string FrontName)
        {
            string ColName = string.Empty;
            string FrameColor = string.Empty;
            string InsetType = string.Empty;
            string InsetColor = string.Empty;

            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            DataTable DT3 = new DataTable();

            for (int y = 0; y < DistinctSizesDT.Rows.Count; y++)
            {
                using (DataView DV = new DataView(SourceDT))
                {
                    DT1 = DV.ToTable(true, new string[] { "ColorID", "PatinaID" });
                }
                for (int i = 0; i < DT1.Rows.Count; i++)
                {
                    using (DataView DV = new DataView(SourceDT, "ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) + " AND PatinaID=" + Convert.ToInt32(DT1.Rows[i]["PatinaID"]), string.Empty, DataViewRowState.CurrentRows))
                    {
                        DT2 = DV.ToTable(true, new string[] { "InsetTypeID" });
                    }
                    for (int j = 0; j < DT2.Rows.Count; j++)
                    {
                        using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]), string.Empty, DataViewRowState.CurrentRows))
                        {
                            DT3 = DV.ToTable(true, new string[] { "InsetColorID" });
                            for (int x = 0; x < DT3.Rows.Count; x++)
                            {
                                DataRow[] rows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) + " AND PatinaID=" + Convert.ToInt32(DT1.Rows[i]["PatinaID"]) +
                                    " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) + " AND InsetColorID=" + Convert.ToInt32(DT3.Rows[x]["InsetColorID"]));

                                if (rows.Count() > 0)
                                {
                                    if (Convert.ToInt32(rows[0]["PatinaID"]) == -1)
                                        FrameColor = GetColorName(Convert.ToInt32(rows[0]["ColorID"]));
                                    else
                                        FrameColor = GetColorName(Convert.ToInt32(rows[0]["ColorID"])) + "+" + GetPatinaName(Convert.ToInt32(rows[0]["PatinaID"]));
                                    //FrameColor = GetColorName(Convert.ToInt32(rows[0]["ColorID"]));
                                    InsetType = GetInsetTypeName(Convert.ToInt32(rows[0]["InsetTypeID"]));
                                    if (Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) == 1)
                                        InsetColor = "Витрина";
                                    else
                                        InsetColor = GetInsetColorName(Convert.ToInt32(rows[0]["InsetColorID"]));

                                    ColName = FrameColor + "(" + InsetType + " " + InsetColor + ")_" + FrontType;
                                    if (!DestinationDT.Columns.Contains(ColName))
                                        DestinationDT.Columns.Add(new DataColumn(ColName, Type.GetType("System.String")));

                                    DestinationDT.Rows[0][ColName] = FrontName;

                                    DataRow[] Srows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) + " AND PatinaID=" + Convert.ToInt32(DT1.Rows[i]["PatinaID"]) +
                                        " AND InsetTypeID=" + Convert.ToInt32(DT2.Rows[j]["InsetTypeID"]) + " AND InsetColorID=" + Convert.ToInt32(DT3.Rows[x]["InsetColorID"]) +
                                        " AND Height=" + Convert.ToInt32(DistinctSizesDT.Rows[y]["Height"]) + " AND Width=" + Convert.ToInt32(DistinctSizesDT.Rows[y]["Width"]));
                                    if (Srows.Count() > 0)
                                    {
                                        int Count = 0;
                                        foreach (DataRow item in Srows)
                                            Count += Convert.ToInt32(item["Count"]);

                                        DataRow[] Drows = DestinationDT.Select("Sizes='" + DistinctSizesDT.Rows[y]["Height"].ToString() + " X " + DistinctSizesDT.Rows[y]["Width"].ToString() + "'");
                                        if (Drows.Count() == 0)
                                        {
                                            DataRow NewRow = DestinationDT.NewRow();
                                            NewRow["Sizes"] = DistinctSizesDT.Rows[y]["Height"].ToString() + " X " + DistinctSizesDT.Rows[y]["Width"].ToString();
                                            NewRow["Height"] = DistinctSizesDT.Rows[y]["Height"];
                                            NewRow["Width"] = DistinctSizesDT.Rows[y]["Width"];
                                            NewRow[ColName] = Count;
                                            DestinationDT.Rows.Add(NewRow);
                                        }
                                        else
                                        {
                                            Drows[0][ColName] = Count;
                                        }
                                    }
                                }
                                else
                                    continue;

                            }
                        }
                    }
                }
            }
        }

        #region Trimming and Gash

        private DataTable TrimDistHeightTable(DataTable SourceDT, bool OrderASC)
        {
            int Height = 0;
            DataTable DT = new DataTable();
            DT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            foreach (DataRow Row in SourceDT.Rows)
            {
                if (int.TryParse(Row["Height"].ToString(), out Height))
                {
                    DataRow NewRow = DT.NewRow();
                    NewRow["Height"] = Height;
                    DT.Rows.Add(NewRow);
                }
                if (int.TryParse(Row["Width"].ToString(), out Height))
                {
                    DataRow NewRow = DT.NewRow();
                    NewRow["Height"] = Height;
                    DT.Rows.Add(NewRow);
                }
            }
            using (DataView DV = new DataView(DT.Copy()))
            {
                DT.Clear();
                if (OrderASC)
                    DV.Sort = "Height ASC";
                else
                    DV.Sort = "Height DESC";
                DT = DV.ToTable(true, new string[] { "Height" });
            }
            return DT;
        }

        private void TrimCollectSimpleFronts(ref DataTable DestinationDT, int Admission, bool HeightASC)
        {
            DataTable DT = Turin1GridsDT.Clone();
            foreach (DataRow item in SofiaSimpleDT.Rows)
                DT.Rows.Add(item.ItemArray);
            foreach (DataRow item in DakotaAppliqueDT.Rows)
                DT.Rows.Add(item.ItemArray);
            foreach (DataRow item in SofiaAppliqueDT.Rows)
                DT.Rows.Add(item.ItemArray);

            DataTable TempDT = DestinationDT.Clone();

            if (LorenzoSimpleDT.Rows.Count > 0)
                TrimmingSyngly(LorenzoSimpleDT, ref TempDT, Admission, 1, true, HeightASC, string.Empty);
            if (ElegantSimpleDT.Rows.Count > 0)
                TrimmingSyngly(ElegantSimpleDT, ref TempDT, Admission, 1, true, HeightASC, string.Empty);
            if (Patricia1SimpleDT.Rows.Count > 0)
                TrimmingSyngly(Patricia1SimpleDT, ref TempDT, Admission, 1, true, HeightASC, string.Empty);
            if (KansasSimpleDT.Rows.Count > 0)
                TrimmingSyngly(KansasSimpleDT, ref TempDT, Admission, 1, true, HeightASC, string.Empty);

            DataTable dt = DakotaSimpleDT.Clone();
            DataRow[] rows = DakotaSimpleDT.Select("InsetTypeID=29272");
            foreach (DataRow item in rows)
                dt.Rows.Add(item.ItemArray);

            if (dt.Rows.Count > 0)
                TrimmingSyngly(dt, ref TempDT, Admission, 1, true, HeightASC, " РЕШ");
            dt.Clear();
            rows = DakotaSimpleDT.Select("InsetTypeID<>29272");
            foreach (DataRow item in rows)
                dt.Rows.Add(item.ItemArray);
            if (dt.Rows.Count > 0)
                TrimmingSyngly(dt, ref TempDT, Admission, 1, true, HeightASC, string.Empty);

            if (DT.Rows.Count > 0)
                TrimmingSyngly(DT, ref TempDT, Admission, 1, true, HeightASC, string.Empty);

            DT.Clear();
            foreach (DataRow item in Turin1SimpleDT.Rows)
                DT.Rows.Add(item.ItemArray);
            foreach (DataRow item in Turin1_1SimpleDT.Rows)
                DT.Rows.Add(item.ItemArray);
            foreach (DataRow item in Turin1GridsDT.Rows)
                DT.Rows.Add(item.ItemArray);
            foreach (DataRow item in Turin1_1GridsDT.Rows)
                DT.Rows.Add(item.ItemArray);
            if (DT.Rows.Count > 0)
                TrimmingSyngly(DT, ref TempDT, Admission, 1, true, HeightASC, string.Empty);
            if (Turin3SimpleDT.Rows.Count > 0)
                TrimmingSyngly(Turin3SimpleDT, ref TempDT, Admission, 1, true, HeightASC, string.Empty);
            if (LeonSimpleDT.Rows.Count > 0)
                TrimmingSyngly(LeonSimpleDT, ref TempDT, Admission, 1, true, HeightASC, string.Empty);
            if (InfinitiSimpleDT.Rows.Count > 0)
                TrimmingSyngly(InfinitiSimpleDT, ref TempDT, Admission, 1, true, HeightASC, string.Empty);
            for (int i = 1; i < TempDT.Rows.Count; i++)
            {
                if (Convert.ToInt32(TempDT.Rows[i]["FrontID"]) == Convert.ToInt32(TempDT.Rows[i - 1]["FrontID"]) &&
                    Convert.ToInt32(TempDT.Rows[i]["ColorID"]) == Convert.ToInt32(TempDT.Rows[i - 1]["ColorID"]))
                {
                    TempDT.Rows[i]["Front"] = string.Empty;
                    TempDT.Rows[i]["Color"] = string.Empty;
                }
            }
            foreach (DataRow item in TempDT.Rows)
                DestinationDT.Rows.Add(item.ItemArray);
        }

        private void TrimCollectGridFronts(ref DataTable DestinationDT, int Admission, bool HeightASC)
        {
            DataTable TempDT = DestinationDT.Clone();

            if (LorenzoGridsDT.Rows.Count > 0)
                TrimmingSyngly(LorenzoGridsDT, ref TempDT, Admission, 3, true, HeightASC, " РЕШ");
            if (ElegantGridsDT.Rows.Count > 0)
                TrimmingSyngly(ElegantGridsDT, ref TempDT, Admission, 3, true, HeightASC, " РЕШ");
            if (Patricia1GridsDT.Rows.Count > 0)
                TrimmingSyngly(Patricia1GridsDT, ref TempDT, Admission, 3, true, HeightASC, " РЕШ");
            if (KansasGridsDT.Rows.Count > 0)
                TrimmingSyngly(KansasGridsDT, ref TempDT, Admission, 3, true, HeightASC, " РЕШ");
            if (DakotaGridsDT.Rows.Count > 0)
                TrimmingSyngly(DakotaGridsDT, ref TempDT, Admission, 3, true, HeightASC, " РЕШ");
            if (SofiaGridsDT.Rows.Count > 0)
                TrimmingSyngly(SofiaGridsDT, ref TempDT, Admission, 3, true, HeightASC, " РЕШ");
            if (Turin3GridsDT.Rows.Count > 0)
                TrimmingSyngly(Turin3GridsDT, ref TempDT, Admission, 3, true, HeightASC, " РЕШ");
            if (LeonGridsDT.Rows.Count > 0)
                TrimmingSyngly(LeonGridsDT, ref TempDT, Admission, 3, true, HeightASC, " РЕШ");
            if (InfinitiGridsDT.Rows.Count > 0)
                TrimmingSyngly(InfinitiGridsDT, ref TempDT, Admission, 3, true, HeightASC, " РЕШ");
            for (int i = 1; i < TempDT.Rows.Count; i++)
            {
                if (Convert.ToInt32(TempDT.Rows[i]["FrontID"]) == Convert.ToInt32(TempDT.Rows[i - 1]["FrontID"]) &&
                    Convert.ToInt32(TempDT.Rows[i]["ColorID"]) == Convert.ToInt32(TempDT.Rows[i - 1]["ColorID"]))
                {
                    TempDT.Rows[i]["Front"] = string.Empty;
                    TempDT.Rows[i]["Color"] = string.Empty;
                }
            }
            foreach (DataRow item in TempDT.Rows)
                DestinationDT.Rows.Add(item.ItemArray);
        }

        private void TrimCollectBoxFronts(ref DataTable DestinationDT, int Admission, bool HeightASC)
        {
            DataTable TempDT = DestinationDT.Clone();

            if (LorenzoBoxesDT.Rows.Count > 0)
                TrimmingSyngly(LorenzoBoxesDT, ref TempDT, Admission, 2, true, HeightASC, " ШУФ");
            if (ElegantBoxesDT.Rows.Count > 0)
                TrimmingSyngly(ElegantBoxesDT, ref TempDT, Admission, 2, true, HeightASC, " ШУФ");
            if (Patricia1BoxesDT.Rows.Count > 0)
                TrimmingSyngly(Patricia1BoxesDT, ref TempDT, Admission, 2, true, HeightASC, " ШУФ");
            if (KansasBoxesDT.Rows.Count > 0)
                TrimmingSyngly(KansasBoxesDT, ref TempDT, Admission, 2, true, HeightASC, " ШУФ");
            if (DakotaBoxesDT.Rows.Count > 0)
                TrimmingSyngly(DakotaBoxesDT, ref TempDT, Admission, 2, true, HeightASC, " ШУФ");
            if (SofiaBoxesDT.Rows.Count > 0)
                TrimmingSyngly(SofiaBoxesDT, ref TempDT, Admission, 2, true, HeightASC, " ШУФ");
            if (Turin1BoxesDT.Rows.Count > 0)
                TrimmingSyngly(Turin1BoxesDT, ref TempDT, Admission, 2, true, HeightASC, " ШУФ");
            if (Turin1_1BoxesDT.Rows.Count > 0)
                TrimmingSyngly(Turin1_1BoxesDT, ref TempDT, Admission, 2, true, HeightASC, " ШУФ");
            if (Turin3BoxesDT.Rows.Count > 0)
                TrimmingSyngly(Turin3BoxesDT, ref TempDT, Admission, 2, true, HeightASC, " ШУФ");
            if (LeonBoxesDT.Rows.Count > 0)
                TrimmingSyngly(LeonBoxesDT, ref TempDT, Admission, 2, true, HeightASC, " ШУФ");
            if (InfinitiBoxesDT.Rows.Count > 0)
                TrimmingSyngly(InfinitiBoxesDT, ref TempDT, Admission, 2, true, HeightASC, " ШУФ");
            for (int i = 1; i < TempDT.Rows.Count; i++)
            {
                if (Convert.ToInt32(TempDT.Rows[i]["FrontID"]) == Convert.ToInt32(TempDT.Rows[i - 1]["FrontID"]) &&
                    Convert.ToInt32(TempDT.Rows[i]["ColorID"]) == Convert.ToInt32(TempDT.Rows[i - 1]["ColorID"]))
                {
                    TempDT.Rows[i]["Front"] = string.Empty;
                    TempDT.Rows[i]["Color"] = string.Empty;
                }
            }
            foreach (DataRow item in TempDT.Rows)
                DestinationDT.Rows.Add(item.ItemArray);
        }

        private void GashCollectSimpleFronts(ref DataTable DestinationDT, int Admission, bool HeightASC)
        {
            DataTable DT = Turin1GridsDT.Clone();
            foreach (DataRow item in Turin1SimpleDT.Rows)
                DT.Rows.Add(item.ItemArray);
            foreach (DataRow item in Turin1_1SimpleDT.Rows)
                DT.Rows.Add(item.ItemArray);
            foreach (DataRow item in Turin1GridsDT.Rows)
                DT.Rows.Add(item.ItemArray);
            foreach (DataRow item in Turin1_1GridsDT.Rows)
                DT.Rows.Add(item.ItemArray);

            DataTable TempDT = DestinationDT.Clone();

            if (InfinitiSimpleDT.Rows.Count > 0)
                TrimmingSyngly(InfinitiSimpleDT, ref TempDT, Admission, 1, false, HeightASC, string.Empty);
            if (Turin3SimpleDT.Rows.Count > 0)
                TrimmingSyngly(Turin3SimpleDT, ref TempDT, Admission, 1, false, HeightASC, string.Empty);
            if (LeonSimpleDT.Rows.Count > 0)
                TrimmingSyngly(LeonSimpleDT, ref TempDT, Admission, 1, false, HeightASC, string.Empty);

            if (DT.Rows.Count > 0)
                TrimmingSyngly(DT, ref TempDT, Admission, 1, false, HeightASC, string.Empty);

            DT.Clear();
            foreach (DataRow item in SofiaSimpleDT.Rows)
                DT.Rows.Add(item.ItemArray);
            foreach (DataRow item in DakotaAppliqueDT.Rows)
                DT.Rows.Add(item.ItemArray);
            foreach (DataRow item in SofiaAppliqueDT.Rows)
                DT.Rows.Add(item.ItemArray);

            if (DT.Rows.Count > 0)
                TrimmingSyngly(DT, ref TempDT, Admission, 1, true, HeightASC, string.Empty);
            if (LorenzoSimpleDT.Rows.Count > 0)
                TrimmingSyngly(LorenzoSimpleDT, ref TempDT, Admission, 1, false, HeightASC, string.Empty);
            if (ElegantSimpleDT.Rows.Count > 0)
                TrimmingSyngly(ElegantSimpleDT, ref TempDT, Admission, 1, false, HeightASC, string.Empty);
            if (Patricia1SimpleDT.Rows.Count > 0)
                TrimmingSyngly(Patricia1SimpleDT, ref TempDT, Admission, 1, false, HeightASC, string.Empty);
            if (KansasSimpleDT.Rows.Count > 0)
                TrimmingSyngly(KansasSimpleDT, ref TempDT, Admission, 1, false, HeightASC, string.Empty);

            DataTable dt = DakotaSimpleDT.Clone();
            DataRow[] rows = DakotaSimpleDT.Select("InsetTypeID=29272");
            foreach (DataRow item in rows)
                dt.Rows.Add(item.ItemArray);

            if (dt.Rows.Count > 0)
                TrimmingSyngly(dt, ref TempDT, Admission, 4, false, HeightASC, " РЕШ");
            dt.Clear();
            rows = DakotaSimpleDT.Select("InsetTypeID<>29272");
            foreach (DataRow item in rows)
                dt.Rows.Add(item.ItemArray);
            if (dt.Rows.Count > 0)
                TrimmingSyngly(dt, ref TempDT, Admission, 4, false, HeightASC, string.Empty);


            if (TempDT.Rows.Count > 0 && Convert.ToInt32(TempDT.Rows[0]["VitrinaCount"]) > 0)
                TempDT.Rows[0]["VitrinaNotes"] = Convert.ToInt32(TempDT.Rows[0]["VitrinaCount"]) + " витр.";
            for (int i = 1; i < TempDT.Rows.Count; i++)
            {
                if (Convert.ToInt32(TempDT.Rows[i]["VitrinaCount"]) > 0)
                    TempDT.Rows[i]["VitrinaNotes"] = Convert.ToInt32(TempDT.Rows[i]["VitrinaCount"]) + " витр.";
                if (Convert.ToInt32(TempDT.Rows[i]["FrontID"]) == 28922 && Convert.ToInt32(TempDT.Rows[i]["GridCount"]) > 0)
                    TempDT.Rows[i]["GridNotes"] = Convert.ToInt32(TempDT.Rows[i]["GridCount"]) + " реш.СОТЫ";
                if (Convert.ToInt32(TempDT.Rows[i]["FrontID"]) == Convert.ToInt32(TempDT.Rows[i - 1]["FrontID"]) &&
                    Convert.ToInt32(TempDT.Rows[i]["ColorID"]) == Convert.ToInt32(TempDT.Rows[i - 1]["ColorID"]))
                {
                    TempDT.Rows[i]["Front"] = string.Empty;
                    TempDT.Rows[i]["Color"] = string.Empty;
                }
            }
            foreach (DataRow item in TempDT.Rows)
                DestinationDT.Rows.Add(item.ItemArray);
        }

        private void GashCollectGridFronts(ref DataTable DestinationDT, int Admission, bool HeightASC)
        {
            DataTable TempDT = DestinationDT.Clone();

            if (InfinitiGridsDT.Rows.Count > 0)
                TrimmingSyngly(InfinitiGridsDT, ref TempDT, Admission, 3, false, HeightASC, " РЕШ");
            if (Turin3GridsDT.Rows.Count > 0)
                TrimmingSyngly(Turin3GridsDT, ref TempDT, Admission, 3, false, HeightASC, " РЕШ");
            if (LeonGridsDT.Rows.Count > 0)
                TrimmingSyngly(LeonGridsDT, ref TempDT, Admission, 3, false, HeightASC, " РЕШ");
            if (SofiaGridsDT.Rows.Count > 0)
                TrimmingSyngly(SofiaGridsDT, ref TempDT, Admission, 3, false, HeightASC, " РЕШ");
            if (LorenzoGridsDT.Rows.Count > 0)
                TrimmingSyngly(LorenzoGridsDT, ref TempDT, Admission, 3, false, HeightASC, " РЕШ");
            if (ElegantGridsDT.Rows.Count > 0)
                TrimmingSyngly(ElegantGridsDT, ref TempDT, Admission, 3, false, HeightASC, " РЕШ");
            if (Patricia1GridsDT.Rows.Count > 0)
                TrimmingSyngly(Patricia1GridsDT, ref TempDT, Admission, 3, false, HeightASC, " РЕШ");
            if (KansasGridsDT.Rows.Count > 0)
                TrimmingSyngly(KansasGridsDT, ref TempDT, Admission, 3, false, HeightASC, " РЕШ");
            if (DakotaGridsDT.Rows.Count > 0)
                TrimmingSyngly(DakotaGridsDT, ref TempDT, Admission, 3, false, HeightASC, " РЕШ");
            for (int i = 1; i < TempDT.Rows.Count; i++)
            {
                if (Convert.ToInt32(TempDT.Rows[i]["FrontID"]) == Convert.ToInt32(TempDT.Rows[i - 1]["FrontID"]) &&
                    Convert.ToInt32(TempDT.Rows[i]["ColorID"]) == Convert.ToInt32(TempDT.Rows[i - 1]["ColorID"]))
                {
                    TempDT.Rows[i]["Front"] = string.Empty;
                    TempDT.Rows[i]["Color"] = string.Empty;
                }
            }
            foreach (DataRow item in TempDT.Rows)
                DestinationDT.Rows.Add(item.ItemArray);
        }

        private void GashCollectBoxFronts(ref DataTable DestinationDT, int Admission, bool HeightASC)
        {
            DataTable TempDT = DestinationDT.Clone();

            if (InfinitiBoxesDT.Rows.Count > 0)
                TrimmingSyngly(InfinitiBoxesDT, ref TempDT, Admission, 2, false, HeightASC, " ШУФ");
            if (Turin3BoxesDT.Rows.Count > 0)
                TrimmingSyngly(Turin3BoxesDT, ref TempDT, Admission, 2, false, HeightASC, " ШУФ");
            if (LeonBoxesDT.Rows.Count > 0)
                TrimmingSyngly(LeonBoxesDT, ref TempDT, Admission, 2, false, HeightASC, " ШУФ");
            if (Turin1BoxesDT.Rows.Count > 0)
                TrimmingSyngly(Turin1BoxesDT, ref TempDT, Admission, 2, false, HeightASC, " ШУФ");
            if (Turin1_1BoxesDT.Rows.Count > 0)
                TrimmingSyngly(Turin1_1BoxesDT, ref TempDT, Admission, 2, false, HeightASC, " ШУФ");
            if (SofiaBoxesDT.Rows.Count > 0)
                TrimmingSyngly(SofiaBoxesDT, ref TempDT, Admission, 2, false, HeightASC, " ШУФ");
            if (LorenzoBoxesDT.Rows.Count > 0)
                TrimmingSyngly(LorenzoBoxesDT, ref TempDT, Admission, 2, false, HeightASC, " ШУФ");
            if (ElegantBoxesDT.Rows.Count > 0)
                TrimmingSyngly(ElegantBoxesDT, ref TempDT, Admission, 2, false, HeightASC, " ШУФ");
            if (Patricia1BoxesDT.Rows.Count > 0)
                TrimmingSyngly(Patricia1BoxesDT, ref TempDT, Admission, 2, false, HeightASC, " ШУФ");
            if (KansasBoxesDT.Rows.Count > 0)
                TrimmingSyngly(KansasBoxesDT, ref TempDT, Admission, 2, false, HeightASC, " ШУФ");
            if (DakotaBoxesDT.Rows.Count > 0)
                TrimmingSyngly(DakotaBoxesDT, ref TempDT, Admission, 2, false, HeightASC, " ШУФ");
            if (TempDT.Rows.Count > 0 && Convert.ToInt32(TempDT.Rows[0]["VitrinaCount"]) > 0)
                TempDT.Rows[0]["VitrinaNotes"] = Convert.ToInt32(TempDT.Rows[0]["VitrinaCount"]) + " витр.";
            for (int i = 1; i < TempDT.Rows.Count; i++)
            {
                if (Convert.ToInt32(TempDT.Rows[i]["VitrinaCount"]) > 0)
                    TempDT.Rows[i]["VitrinaNotes"] = Convert.ToInt32(TempDT.Rows[i]["VitrinaCount"]) + " витр.";
                if (Convert.ToInt32(TempDT.Rows[i]["FrontID"]) == Convert.ToInt32(TempDT.Rows[i - 1]["FrontID"]) &&
                    Convert.ToInt32(TempDT.Rows[i]["ColorID"]) == Convert.ToInt32(TempDT.Rows[i - 1]["ColorID"]))
                {
                    TempDT.Rows[i]["Front"] = string.Empty;
                    TempDT.Rows[i]["Color"] = string.Empty;
                }
            }
            foreach (DataRow item in TempDT.Rows)
                DestinationDT.Rows.Add(item.ItemArray);
        }

        #endregion

        private void TrimmingSyngly(DataTable SourceDT, ref DataTable DestinationDT, int Admission, int ProfileType, bool FrameColorASC, bool HeightASC, string AdditionalName)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = TrimDistHeightTable(SourceDT, HeightASC);

            using (DataView DV = new DataView(SourceDT))
            {
                if (FrameColorASC)
                    DV.Sort = "ColorID";
                else
                    DV.Sort = "ColorID DESC";
                DT1 = DV.ToTable(true, new string[] { "ColorID" });
            }

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    DataRow[] Srows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]));
                    if (Srows.Count() > 0)
                    {
                        int Count = 0;
                        int VitrinaCount = 0;
                        int GridCount = 0;
                        int Height = Convert.ToInt32(DT2.Rows[j]["Height"]) + Admission;
                        foreach (DataRow item in Srows)
                        {
                            Count += Convert.ToInt32(item["Count"]);
                            if (Convert.ToInt32(item["InsetTypeID"]) == 1)
                                VitrinaCount += Convert.ToInt32(item["Count"]);
                            if (Convert.ToInt32(item["InsetTypeID"]) == 685 || Convert.ToInt32(item["InsetTypeID"]) == 686 || Convert.ToInt32(item["InsetTypeID"]) == 687 || Convert.ToInt32(item["InsetTypeID"]) == 688 || Convert.ToInt32(item["InsetTypeID"]) == 29272)
                                GridCount += Convert.ToInt32(item["Count"]);
                        }
                        string Name = ProfileName(Convert.ToInt32(Srows[0]["FrontConfigID"])) + AdditionalName;
                        DataRow[] rows = DestinationDT.Select("Front='" + Name + "' AND ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                            " AND Height=" + Height);
                        if (rows.Count() == 0)
                        {
                            DataRow NewRow = DestinationDT.NewRow();
                            NewRow["Front"] = Name;
                            NewRow["Color"] = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                            NewRow["Height"] = Height;
                            NewRow["Count"] = Count * 2;
                            NewRow["VitrinaCount"] = VitrinaCount * 2;
                            NewRow["GridCount"] = GridCount * 2;
                            NewRow["FrontID"] = Convert.ToInt32(Srows[0]["FrontID"]);
                            NewRow["ColorID"] = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                            NewRow["ProfileType"] = ProfileType;
                            DestinationDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count * 2;
                            rows[0]["VitrinaCount"] = Convert.ToInt32(rows[0]["VitrinaCount"]) + VitrinaCount * 2;
                            rows[0]["GridCount"] = Convert.ToInt32(rows[0]["GridCount"]) + GridCount * 2;
                        }
                    }
                    Srows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Height"]));
                    if (Srows.Count() > 0)
                    {
                        int Count = 0;
                        int VitrinaCount = 0;
                        int GridCount = 0;
                        int Height = Convert.ToInt32(DT2.Rows[j]["Height"]) + Admission;
                        foreach (DataRow item in Srows)
                        {
                            Count += Convert.ToInt32(item["Count"]);
                            if (Convert.ToInt32(item["InsetTypeID"]) == 1)
                                VitrinaCount += Convert.ToInt32(item["Count"]);
                            if (Convert.ToInt32(item["InsetTypeID"]) == 685 || Convert.ToInt32(item["InsetTypeID"]) == 686 || Convert.ToInt32(item["InsetTypeID"]) == 687 || Convert.ToInt32(item["InsetTypeID"]) == 688 || Convert.ToInt32(item["InsetTypeID"]) == 29272)
                                GridCount += Convert.ToInt32(item["Count"]);
                        }

                        string Name = ProfileName(Convert.ToInt32(Srows[0]["FrontConfigID"])) + AdditionalName;
                        DataRow[] rows = DestinationDT.Select("Front='" + Name + "' AND ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                            " AND Height=" + Height);
                        if (rows.Count() == 0)
                        {
                            DataRow NewRow = DestinationDT.NewRow();
                            NewRow["Front"] = Name;
                            NewRow["Color"] = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                            NewRow["Height"] = Height;
                            NewRow["Count"] = Count * 2;
                            NewRow["VitrinaCount"] = VitrinaCount * 2;
                            NewRow["GridCount"] = GridCount * 2;
                            NewRow["FrontID"] = Convert.ToInt32(Srows[0]["FrontID"]);
                            NewRow["ColorID"] = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                            NewRow["ProfileType"] = ProfileType;
                            DestinationDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count * 2;
                            rows[0]["VitrinaCount"] = Convert.ToInt32(rows[0]["VitrinaCount"]) + VitrinaCount * 2;
                            rows[0]["GridCount"] = Convert.ToInt32(rows[0]["GridCount"]) + GridCount * 2;
                        }
                    }
                }
            }
            DT1.Dispose();
            DT2.Dispose();
        }

        private void AdditionsSyngly(DataTable SourceDT, ref DataTable DestinationDT)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = TrimDistHeightTable(SourceDT, true);

            using (DataView DV = new DataView(SourceDT))
            {
                DT1 = DV.ToTable(true, new string[] { "ColorID" });
            }

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    DataRow[] Srows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]));
                    if (Srows.Count() > 0)
                    {
                        int Count = 0;

                        foreach (DataRow item in Srows)
                            Count += Convert.ToInt32(item["Count"]);

                        DataRow[] rows = DestinationDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                            " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]));
                        if (rows.Count() == 0)
                        {
                            DataRow NewRow = DestinationDT.NewRow();
                            NewRow["Front"] = ProfileName(Convert.ToInt32(Srows[0]["FrontConfigID"]));
                            NewRow["Color"] = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                            NewRow["Height"] = Convert.ToInt32(DT2.Rows[j]["Height"]);
                            NewRow["Count"] = Count * 2;
                            NewRow["FrontID"] = Convert.ToInt32(Srows[0]["FrontID"]);
                            NewRow["ColorID"] = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                            DestinationDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count * 2;
                        }
                    }
                    Srows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Height"]));
                    if (Srows.Count() > 0)
                    {
                        int Count = 0;
                        foreach (DataRow item in Srows)
                            Count += Convert.ToInt32(item["Count"]);

                        DataRow[] rows = DestinationDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                            " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]));
                        if (rows.Count() == 0)
                        {
                            DataRow NewRow = DestinationDT.NewRow();
                            NewRow["Front"] = ProfileName(Convert.ToInt32(Srows[0]["FrontConfigID"]));
                            NewRow["Color"] = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                            NewRow["Height"] = Convert.ToInt32(DT2.Rows[j]["Height"]);
                            NewRow["Count"] = Count * 2;
                            NewRow["FrontID"] = Convert.ToInt32(Srows[0]["FrontID"]);
                            NewRow["ColorID"] = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                            DestinationDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count * 2;
                        }
                    }
                }
            }
            DT1.Dispose();
            DT2.Dispose();
        }

        private void AdditionsSyngly(DataTable SourceDT, ref DataTable DestinationDT, int HeightMin)
        {
            DataTable DT1 = new DataTable();
            DataTable DT2 = TrimDistHeightTable(SourceDT, true);

            using (DataView DV = new DataView(SourceDT))
            {
                DT1 = DV.ToTable(true, new string[] { "ColorID" });
            }

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    if (Convert.ToInt32(DT2.Rows[j]["Height"]) < HeightMin)
                        continue;
                    DataRow[] Srows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]));
                    if (Srows.Count() > 0)
                    {
                        int Count = 0;

                        foreach (DataRow item in Srows)
                            Count += Convert.ToInt32(item["Count"]);

                        DataRow[] rows = DestinationDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                            " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]));
                        if (rows.Count() == 0)
                        {
                            DataRow NewRow = DestinationDT.NewRow();
                            NewRow["Front"] = ProfileName(Convert.ToInt32(Srows[0]["FrontConfigID"]));
                            NewRow["Color"] = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                            NewRow["Height"] = Convert.ToInt32(DT2.Rows[j]["Height"]);
                            NewRow["Count"] = Count * 2;
                            NewRow["FrontID"] = Convert.ToInt32(Srows[0]["FrontID"]);
                            NewRow["ColorID"] = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                            DestinationDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count * 2;
                        }
                    }
                    Srows = SourceDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                        " AND Width=" + Convert.ToInt32(DT2.Rows[j]["Height"]));
                    if (Srows.Count() > 0)
                    {
                        int Count = 0;
                        foreach (DataRow item in Srows)
                            Count += Convert.ToInt32(item["Count"]);

                        DataRow[] rows = DestinationDT.Select("ColorID=" + Convert.ToInt32(DT1.Rows[i]["ColorID"]) +
                            " AND Height=" + Convert.ToInt32(DT2.Rows[j]["Height"]));
                        if (rows.Count() == 0)
                        {
                            DataRow NewRow = DestinationDT.NewRow();
                            NewRow["Front"] = ProfileName(Convert.ToInt32(Srows[0]["FrontConfigID"]));
                            NewRow["Color"] = GetColorName(Convert.ToInt32(DT1.Rows[i]["ColorID"]));
                            NewRow["Height"] = Convert.ToInt32(DT2.Rows[j]["Height"]);
                            NewRow["Count"] = Count * 2;
                            NewRow["FrontID"] = Convert.ToInt32(Srows[0]["FrontID"]);
                            NewRow["ColorID"] = Convert.ToInt32(DT1.Rows[i]["ColorID"]);
                            DestinationDT.Rows.Add(NewRow);
                        }
                        else
                        {
                            rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count * 2;
                        }
                    }
                }
            }
            DT1.Dispose();
            DT2.Dispose();
        }

        private void CollectRemovingQuarter(ref DataTable DestinationDT, bool OrderASC)
        {
            if (LorenzoVitrinaDT.Rows.Count > 0)
                AdditionsSyngly(LorenzoVitrinaDT, ref DestinationDT);
            if (ElegantVitrinaDT.Rows.Count > 0)
                AdditionsSyngly(ElegantVitrinaDT, ref DestinationDT);
            if (Patricia1VitrinaDT.Rows.Count > 0)
                AdditionsSyngly(Patricia1VitrinaDT, ref DestinationDT);
            if (KansasVitrinaDT.Rows.Count > 0)
                AdditionsSyngly(KansasVitrinaDT, ref DestinationDT);
            if (DakotaVitrinaDT.Rows.Count > 0)
                AdditionsSyngly(DakotaVitrinaDT, ref DestinationDT);
            if (SofiaVitrinaDT.Rows.Count > 0)
                AdditionsSyngly(SofiaVitrinaDT, ref DestinationDT);
            if (Turin1VitrinaDT.Rows.Count > 0)
                AdditionsSyngly(Turin1VitrinaDT, ref DestinationDT);
            if (Turin1_1VitrinaDT.Rows.Count > 0)
                AdditionsSyngly(Turin1_1VitrinaDT, ref DestinationDT);
            if (Turin3VitrinaDT.Rows.Count > 0)
                AdditionsSyngly(Turin3VitrinaDT, ref DestinationDT);
            if (LeonVitrinaDT.Rows.Count > 0)
                AdditionsSyngly(LeonVitrinaDT, ref DestinationDT);
            if (InfinitiVitrinaDT.Rows.Count > 0)
                AdditionsSyngly(InfinitiVitrinaDT, ref DestinationDT);

            for (int i = 1; i < DestinationDT.Rows.Count; i++)
            {
                if (Convert.ToInt32(DestinationDT.Rows[i]["FrontID"]) == Convert.ToInt32(DestinationDT.Rows[i - 1]["FrontID"]) &&
                    Convert.ToInt32(DestinationDT.Rows[i]["ColorID"]) == Convert.ToInt32(DestinationDT.Rows[i - 1]["ColorID"]))
                {
                    DestinationDT.Rows[i]["Front"] = string.Empty;
                    DestinationDT.Rows[i]["Color"] = string.Empty;
                }
            }
        }

        private void CollectRemovingBoxes(ref DataTable DestinationDT, bool OrderASC)
        {
            if (Turin1RemovingBoxesDT.Rows.Count > 0)
                AdditionsSyngly(Turin1RemovingBoxesDT, ref DestinationDT, 138);
            if (Turin3RemovingBoxesDT.Rows.Count > 0)
                AdditionsSyngly(Turin3RemovingBoxesDT, ref DestinationDT, 138);

            for (int i = 1; i < DestinationDT.Rows.Count; i++)
            {
                if (Convert.ToInt32(DestinationDT.Rows[i]["FrontID"]) == Convert.ToInt32(DestinationDT.Rows[i - 1]["FrontID"]) &&
                    Convert.ToInt32(DestinationDT.Rows[i]["ColorID"]) == Convert.ToInt32(DestinationDT.Rows[i - 1]["ColorID"]))
                {
                    DestinationDT.Rows[i]["Front"] = string.Empty;
                    DestinationDT.Rows[i]["Color"] = string.Empty;
                }
            }
        }

        private void CollectGrooveGrids(ref DataTable DestinationDT, bool OrderASC)
        {
            if (Turin1GridsDT.Rows.Count > 0)
                AdditionsSyngly(Turin1GridsDT, ref DestinationDT);
            if (Turin1_1GridsDT.Rows.Count > 0)
                AdditionsSyngly(Turin1_1GridsDT, ref DestinationDT);

            for (int i = 1; i < DestinationDT.Rows.Count; i++)
            {
                if (Convert.ToInt32(DestinationDT.Rows[i]["FrontID"]) == Convert.ToInt32(DestinationDT.Rows[i - 1]["FrontID"]) &&
                    Convert.ToInt32(DestinationDT.Rows[i]["ColorID"]) == Convert.ToInt32(DestinationDT.Rows[i - 1]["ColorID"]))
                {
                    DestinationDT.Rows[i]["Front"] = string.Empty;
                    DestinationDT.Rows[i]["Color"] = string.Empty;
                }
            }
        }

        #region Inset

        private DataTable InsetDistSizesTable(DataTable SourceDT, bool OrderASC)
        {
            DataTable DT = new DataTable();
            DT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            DT.Columns.Add(new DataColumn("Width", Type.GetType("System.Int32")));
            foreach (DataRow Row in SourceDT.Rows)
            {
                DataRow NewRow = DT.NewRow();
                NewRow["Height"] = Row["Height"];
                NewRow["Width"] = Row["Width"];
                DT.Rows.Add(NewRow);
            }
            using (DataView DV = new DataView(DT.Copy()))
            {
                DT.Clear();
                if (OrderASC)
                    DV.Sort = "Height ASC, Width ASC";
                else
                    DV.Sort = "Height DESC, Width DESC";
                DT = DV.ToTable(true, new string[] { "Height", "Width" });
            }
            return DT;
        }

        private void SimpleInsetsOnly(DataTable SourceDT, ref DataTable DestinationDT, int HeightMargin, int WidthMargin, bool OrderASC, bool IsBox)
        {
            string SizeASC = "Height, Width, Notes";
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            DataTable DT3 = new DataTable();
            int H = 0;
            int W = 0;
            string N = string.Empty;

            if (!OrderASC)
                SizeASC = "Height DESC, Width DESC, Notes DESC";

            DataRow[] irows = InsetTypesDataTable.Select("(GroupID = 3 OR GroupID = 4) AND InsetTypeID<>29272");
            string filter = string.Empty;
            foreach (DataRow item in irows)
                filter += item["InsetTypeID"].ToString() + ",";
            if (filter.Length > 0)
                filter = "InsetTypeID IN (" + filter.Substring(0, filter.Length - 1) + ")";
            //Нужно добавлять Id Новых филенок
            using (DataView DV = new DataView(SourceDT, filter + " OR InsetTypeID IN (2069,2070,2071,2073,2075,2077,2233,3644,29043,29531) OR InsetTypeID IN (685,686,687,688,29470,29471) OR InsetTypeID IN (28961,3653,3654,3655)", string.Empty, DataViewRowState.CurrentRows))
            {
                DT1 = DV.ToTable(true, new string[] { "InsetTypeID" });
            }

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]), "InsetColorID", DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "InsetColorID" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    H = 0;
                    W = 0;
                    N = string.Empty;
                    using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                        " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]), SizeASC, DataViewRowState.CurrentRows))
                    {
                        DT3 = DV.ToTable(true, new string[] { "Height", "Width", "Notes" });
                    }

                    for (int y = 0; y < DT3.Rows.Count; y++)
                    {
                        if (Convert.ToInt32(DT3.Rows[y]["Height"]) == H && Convert.ToInt32(DT3.Rows[y]["Width"]) == W &&
                            DT3.Rows[y]["Notes"].ToString() == N)
                            continue;

                        H = Convert.ToInt32(DT3.Rows[y]["Height"]);
                        W = Convert.ToInt32(DT3.Rows[y]["Width"]);
                        N = DT3.Rows[y]["Notes"].ToString();

                        filter = "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                            " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) +
                            " AND Height=" + Convert.ToInt32(DT3.Rows[y]["Height"]) + " AND Width=" + Convert.ToInt32(DT3.Rows[y]["Width"]) + " AND (Notes='' OR Notes IS NULL)";
                        if (DT3.Rows[y]["Notes"] != DBNull.Value && DT3.Rows[y]["Notes"].ToString().Length > 0)
                            filter = "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                            " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) +
                            " AND Height=" + Convert.ToInt32(DT3.Rows[y]["Height"]) + " AND Width=" + Convert.ToInt32(DT3.Rows[y]["Width"]) +
                            " AND Notes='" + DT3.Rows[y]["Notes"].ToString() + "'";

                        DataRow[] Srows = SourceDT.Select(filter);
                        if (Srows.Count() == 0)
                            continue;

                        int Count = 0;
                        int Height = Convert.ToInt32(DT3.Rows[y]["Height"]) - HeightMargin;
                        int Width = Convert.ToInt32(DT3.Rows[y]["Width"]) - WidthMargin;
                        string Name = string.Empty;
                        string Notes = DT3.Rows[y]["Notes"].ToString();

                        if (Convert.ToInt32(DT3.Rows[y]["Height"]) <= HeightMargin || Convert.ToInt32(DT3.Rows[y]["Width"]) <= WidthMargin)
                            continue;

                        if (IsBox)
                        {
                            if (Height >= 100 && Width >= 100)
                                continue;
                        }
                        else
                        {
                            if (Height < 100 || Width < 100)
                                continue;
                        }

                        foreach (DataRow item in Srows)
                            Count += Convert.ToInt32(item["Count"]);

                        Name = GetInsetTypeName(Convert.ToInt32(Srows[0]["InsetTypeID"])) + " " + GetInsetColorName(Convert.ToInt32(DT2.Rows[j]["InsetColorID"]));
                        if (Notes.Equals("БРВ"))
                            Name += " " + Notes;
                        DataRow[] rows = DestinationDT.Select("Name = '" + Name + "' AND ColorType=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) +
                            " AND Height=" + Height + " AND Width=" + Width);
                        if (rows.Count() == 0)
                        {
                            DataRow NewRow = DestinationDT.NewRow();
                            NewRow["ColorType"] = Convert.ToInt32(DT2.Rows[j]["InsetColorID"]);
                            NewRow["Name"] = Name;
                            NewRow["Height"] = Height;
                            NewRow["Width"] = Width;
                            NewRow["Count"] = Count;
                            DestinationDT.Rows.Add(NewRow);

                        }
                        else
                        {
                            rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count;
                        }
                    }
                }
            }
        }

        private void DakotaInsetsOnly(DataTable SourceDT, ref DataTable DestinationDT, int HeightMargin, int WidthMargin, bool OrderASC, bool IsBox)
        {
            string SizeASC = "Height, Width, Notes";
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            DataTable DT3 = new DataTable();
            int H = 0;
            int W = 0;
            string N = string.Empty;

            if (!OrderASC)
                SizeASC = "Height DESC, Width DESC, Notes DESC";
            string filter = string.Empty;
            //Нужно добавлять Id Новых филенок
            using (DataView DV = new DataView(SourceDT, "InsetTypeID IN (29272,29275,29279)", string.Empty, DataViewRowState.CurrentRows))
            {
                DT1 = DV.ToTable(true, new string[] { "InsetTypeID" });
            }

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]), "InsetColorID", DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "InsetColorID" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    H = 0;
                    W = 0;
                    N = string.Empty;
                    using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                        " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]), SizeASC, DataViewRowState.CurrentRows))
                    {
                        DT3 = DV.ToTable(true, new string[] { "Height", "Width", "Notes" });
                    }

                    for (int y = 0; y < DT3.Rows.Count; y++)
                    {
                        if (Convert.ToInt32(DT3.Rows[y]["Height"]) == H && Convert.ToInt32(DT3.Rows[y]["Width"]) == W &&
                            DT3.Rows[y]["Notes"].ToString() == N)
                            continue;

                        H = Convert.ToInt32(DT3.Rows[y]["Height"]);
                        W = Convert.ToInt32(DT3.Rows[y]["Width"]);
                        N = DT3.Rows[y]["Notes"].ToString();

                        filter = "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                            " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) +
                            " AND Height=" + Convert.ToInt32(DT3.Rows[y]["Height"]) + " AND Width=" + Convert.ToInt32(DT3.Rows[y]["Width"]) + " AND (Notes='' OR Notes IS NULL)";
                        if (DT3.Rows[y]["Notes"] != DBNull.Value && DT3.Rows[y]["Notes"].ToString().Length > 0)
                            filter = "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                            " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) +
                            " AND Height=" + Convert.ToInt32(DT3.Rows[y]["Height"]) + " AND Width=" + Convert.ToInt32(DT3.Rows[y]["Width"]) +
                            " AND Notes='" + DT3.Rows[y]["Notes"].ToString() + "'";

                        DataRow[] Srows = SourceDT.Select(filter);
                        if (Srows.Count() == 0)
                            continue;

                        int Count = 0;
                        int Height = Convert.ToInt32(DT3.Rows[y]["Height"]) - HeightMargin;
                        int Width = Convert.ToInt32(DT3.Rows[y]["Width"]) - WidthMargin;
                        string Name = string.Empty;
                        string Notes = DT3.Rows[y]["Notes"].ToString();

                        if (Convert.ToInt32(DT3.Rows[y]["Height"]) <= HeightMargin || Convert.ToInt32(DT3.Rows[y]["Width"]) <= WidthMargin)
                            continue;

                        if (IsBox)
                        {
                            if (Height >= 100 && Width >= 100)
                                continue;
                        }
                        else
                        {
                            if (Height < 100 || Width < 100)
                                continue;
                        }

                        foreach (DataRow item in Srows)
                            Count += Convert.ToInt32(item["Count"]);

                        Name = GetInsetTypeName(Convert.ToInt32(Srows[0]["InsetTypeID"])) + " " + GetInsetColorName(Convert.ToInt32(DT2.Rows[j]["InsetColorID"]));
                        if (Notes.Equals("БРВ"))
                            Name += " " + Notes;
                        DataRow[] rows = DestinationDT.Select("Name = '" + Name + "' AND ColorType=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) +
                            " AND Height=" + Height + " AND Width=" + Width);
                        if (rows.Count() == 0)
                        {
                            DataRow NewRow = DestinationDT.NewRow();
                            NewRow["ColorType"] = Convert.ToInt32(DT2.Rows[j]["InsetColorID"]);
                            NewRow["Name"] = Name;
                            NewRow["Height"] = Height;
                            NewRow["Width"] = Width;
                            NewRow["Count"] = Count;
                            DestinationDT.Rows.Add(NewRow);

                        }
                        else
                        {
                            rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count;
                        }
                    }
                }
            }
        }

        private void BoxInsetsOnly(DataTable SourceDT, ref DataTable DestinationDT, int HeightMargin, int WidthMargin, bool OrderASC, bool IsBox)
        {
            string SizeASC = "Height, Width";
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            DataTable DT3 = new DataTable();

            if (!OrderASC)
                SizeASC = "Height DESC, Width DESC";

            DataRow[] irows = InsetTypesDataTable.Select("(GroupID = 3 OR GroupID = 4) AND InsetTypeID<>29272");
            string filter = string.Empty;
            foreach (DataRow item in irows)
                filter += item["InsetTypeID"].ToString() + ",";
            if (filter.Length > 0)
                filter = "InsetTypeID IN (" + filter.Substring(0, filter.Length - 1) + ")";

            using (DataView DV = new DataView(SourceDT, filter + " OR InsetTypeID IN (2069,2070,2071,2073,2075,2077,2233,3644,29043,29531) OR InsetTypeID IN (685,686,687,688,29470,29471) OR InsetTypeID IN (28961,3653,3654,3655)", string.Empty, DataViewRowState.CurrentRows))
            {
                DT1 = DV.ToTable(true, new string[] { "InsetTypeID" });
            }

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]), "InsetColorID", DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "InsetColorID" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                        " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]), SizeASC, DataViewRowState.CurrentRows))
                    {
                        DT3 = DV.ToTable(true, new string[] { "Height", "Width" });
                    }
                    for (int y = 0; y < DT3.Rows.Count; y++)
                    {
                        DataRow[] Srows = SourceDT.Select("InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                            " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) +
                            " AND Height=" + Convert.ToInt32(DT3.Rows[y]["Height"]) + " AND Width=" + Convert.ToInt32(DT3.Rows[y]["Width"]));
                        if (Srows.Count() == 0)
                            continue;

                        int Count = 0;
                        int Height = Convert.ToInt32(DT3.Rows[y]["Height"]) - HeightMargin;
                        int Width = Convert.ToInt32(DT3.Rows[y]["Width"]) - WidthMargin;
                        string Name = string.Empty;
                        if (Convert.ToInt32(DT3.Rows[y]["Height"]) <= HeightMargin || Convert.ToInt32(DT3.Rows[y]["Width"]) <= WidthMargin)
                            continue;

                        foreach (DataRow item in Srows)
                            Count += Convert.ToInt32(item["Count"]);
                        Name = GetInsetTypeName(Convert.ToInt32(Srows[0]["InsetTypeID"])) + " " + GetInsetColorName(Convert.ToInt32(DT2.Rows[j]["InsetColorID"]));
                        DataRow[] rows = DestinationDT.Select("Name = '" + Name + "' AND ColorType = " + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) +
                            " AND Height=" + Height + " AND Width=" + Width);
                        if (rows.Count() == 0)
                        {
                            DataRow NewRow = DestinationDT.NewRow();
                            NewRow["ColorType"] = Convert.ToInt32(DT2.Rows[j]["InsetColorID"]);
                            NewRow["Name"] = Name;
                            NewRow["Height"] = Height;
                            NewRow["Width"] = Width;
                            NewRow["Count"] = Count;
                            DestinationDT.Rows.Add(NewRow);

                        }
                        else
                        {
                            rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count;
                        }
                    }
                }
            }
        }

        private void GridInsetsOnly(DataTable SourceDT, ref DataTable DestinationDT, int HeightMargin, int WidthMargin, bool OrderASC, bool IsBox)
        {
            string SizeASC = "Height, Width";
            DataTable DT1 = new DataTable();
            DataTable DT2 = new DataTable();
            DataTable DT3 = new DataTable();

            if (!OrderASC)
                SizeASC = "Height DESC, Width DESC";
            //InsetTypeID IN (685,686,687,688,29470,29471) РЕШЕТКИ
            //InsetTypeID IN (28961,3653,3654,3655) АПЛИКАЦИИ
            //InsetTypeID IN (2069,2070,2071,2073,2075,2077,2233,3644,29043,29531) ФИЛЕНКИ

            using (DataView DV = new DataView(SourceDT, "InsetTypeID IN (2069,2070,2071,2073,2075,2077,2233,3644,29043,29531) OR InsetTypeID IN (685,686,687,688,29470,29471) OR InsetTypeID IN (28961,3653,3654,3655)", string.Empty, DataViewRowState.CurrentRows))
            {
                DT1 = DV.ToTable(true, new string[] { "InsetTypeID" });
            }

            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]), "InsetColorID", DataViewRowState.CurrentRows))
                {
                    DT2 = DV.ToTable(true, new string[] { "InsetColorID" });
                }
                for (int j = 0; j < DT2.Rows.Count; j++)
                {
                    using (DataView DV = new DataView(SourceDT, "InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                        " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]), SizeASC, DataViewRowState.CurrentRows))
                    {
                        DT3 = DV.ToTable(true, new string[] { "Height", "Width" });
                    }
                    for (int y = 0; y < DT3.Rows.Count; y++)
                    {
                        DataRow[] Srows = SourceDT.Select("InsetTypeID=" + Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) +
                            " AND InsetColorID=" + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) +
                            " AND Height=" + Convert.ToInt32(DT3.Rows[y]["Height"]) + " AND Width=" + Convert.ToInt32(DT3.Rows[y]["Width"]));
                        if (Srows.Count() == 0)
                            continue;

                        int Count = 0;
                        int Height = Convert.ToInt32(DT3.Rows[y]["Height"]) - HeightMargin;
                        int Width = Convert.ToInt32(DT3.Rows[y]["Width"]) - WidthMargin;
                        string Name = string.Empty;

                        if (Convert.ToInt32(DT3.Rows[y]["Height"]) <= HeightMargin || Convert.ToInt32(DT3.Rows[y]["Width"]) <= WidthMargin)
                            continue;

                        foreach (DataRow item in Srows)
                            Count += Convert.ToInt32(item["Count"]);

                        if (Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) == 685 || Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) == 688 || Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) == 29470)
                            Name = " 45";
                        if (Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) == 686 || Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) == 687 || Convert.ToInt32(DT1.Rows[i]["InsetTypeID"]) == 29471)
                            Name = " 90";
                        Name = GetInsetTypeName(Convert.ToInt32(Srows[0]["InsetTypeID"])) + " " + GetInsetColorName(Convert.ToInt32(DT2.Rows[j]["InsetColorID"])) + Name;
                        DataRow[] rows = DestinationDT.Select("Name='" + Name + "' AND ColorType = " + Convert.ToInt32(DT2.Rows[j]["InsetColorID"]) +
                            " AND Height=" + Height + " AND Width=" + Width);
                        if (rows.Count() == 0)
                        {
                            DataRow NewRow = DestinationDT.NewRow();
                            NewRow["ColorType"] = Convert.ToInt32(DT2.Rows[j]["InsetColorID"]);
                            NewRow["Name"] = Name;
                            NewRow["Height"] = Height;
                            NewRow["Width"] = Width;
                            NewRow["Count"] = Count;
                            DestinationDT.Rows.Add(NewRow);

                        }
                        else
                        {
                            rows[0]["Count"] = Convert.ToInt32(rows[0]["Count"]) + Count;
                        }
                    }
                }
            }
        }

        private void CollectSimpleInsets(ref DataTable DestinationDT, bool OrderASC, bool IsBox)
        {
            if (LorenzoSimpleDT.Rows.Count > 0)
                SimpleInsetsOnly(LorenzoSimpleDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.LorenzoSimpleInsetHeight), Convert.ToInt32(FrontMargins.LorenzoSimpleInsetWidth), OrderASC, IsBox);
            if (ElegantSimpleDT.Rows.Count > 0)
                SimpleInsetsOnly(ElegantSimpleDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.ElegantSimpleInsetHeight), Convert.ToInt32(FrontMargins.ElegantSimpleInsetWidth), OrderASC, IsBox);
            if (Patricia1SimpleDT.Rows.Count > 0)
                SimpleInsetsOnly(Patricia1SimpleDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.Patricia1SimpleInsetHeight), Convert.ToInt32(FrontMargins.Patricia1SimpleInsetWidth), OrderASC, IsBox);
            if (KansasSimpleDT.Rows.Count > 0)
                SimpleInsetsOnly(KansasSimpleDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.KansasSimpleInsetHeight), Convert.ToInt32(FrontMargins.KansasSimpleInsetWidth), OrderASC, IsBox);
            if (DakotaSimpleDT.Rows.Count > 0)
                SimpleInsetsOnly(DakotaSimpleDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.DakotaSimpleInsetHeight), Convert.ToInt32(FrontMargins.DakotaSimpleInsetWidth), OrderASC, IsBox);
            if (SofiaSimpleDT.Rows.Count > 0)
                SimpleInsetsOnly(SofiaSimpleDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.SofiaSimpleInsetHeight), Convert.ToInt32(FrontMargins.SofiaSimpleInsetWidth), OrderASC, IsBox);
            if (!IsBox && DakotaAppliqueDT.Rows.Count > 0)
                SimpleInsetsOnly(DakotaAppliqueDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.DakotaSimpleInsetHeight), Convert.ToInt32(FrontMargins.DakotaSimpleInsetWidth), OrderASC, IsBox);
            if (!IsBox && SofiaAppliqueDT.Rows.Count > 0)
                SimpleInsetsOnly(SofiaAppliqueDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.SofiaSimpleInsetHeight), Convert.ToInt32(FrontMargins.SofiaSimpleInsetWidth), OrderASC, IsBox);
            if (Turin1SimpleDT.Rows.Count > 0)
                SimpleInsetsOnly(Turin1SimpleDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.Turin1SimpleInsetHeight), Convert.ToInt32(FrontMargins.Turin1SimpleInsetWidth), OrderASC, IsBox);
            if (Turin1_1SimpleDT.Rows.Count > 0)
                SimpleInsetsOnly(Turin1_1SimpleDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.Turin1SimpleInsetHeight), Convert.ToInt32(FrontMargins.Turin1SimpleInsetWidth), OrderASC, IsBox);
            if (Turin3SimpleDT.Rows.Count > 0)
                SimpleInsetsOnly(Turin3SimpleDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.Turin3SimpleInsetHeight), Convert.ToInt32(FrontMargins.Turin3SimpleInsetWidth), OrderASC, IsBox);
            if (LeonSimpleDT.Rows.Count > 0)
                SimpleInsetsOnly(LeonSimpleDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.LeonSimpleInsetHeight), Convert.ToInt32(FrontMargins.LeonSimpleInsetWidth), OrderASC, IsBox);
            if (InfinitiSimpleDT.Rows.Count > 0)
                SimpleInsetsOnly(InfinitiSimpleDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.InfinitiSimpleInsetHeight), Convert.ToInt32(FrontMargins.InfinitiSimpleInsetWidth), OrderASC, IsBox);
        }

        private void CollectDakotaInsets(ref DataTable DestinationDT, bool OrderASC, bool IsBox)
        {
            if (DakotaSimpleDT.Rows.Count > 0)
                DakotaInsetsOnly(DakotaSimpleDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.DakotaSimpleInsetHeight), Convert.ToInt32(FrontMargins.DakotaSimpleInsetWidth), OrderASC, IsBox);
        }

        private void CollectGridInsets(ref DataTable DestinationDT, bool OrderASC, bool IsBox)
        {
            if (LorenzoGridsDT.Rows.Count > 0)
                GridInsetsOnly(LorenzoGridsDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.LorenzoGridInsetHeight), Convert.ToInt32(FrontMargins.LorenzoGridInsetWidth), OrderASC, IsBox);
            if (ElegantGridsDT.Rows.Count > 0)
                GridInsetsOnly(ElegantGridsDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.ElegantGridInsetHeight), Convert.ToInt32(FrontMargins.ElegantGridInsetWidth), OrderASC, IsBox);
            if (Patricia1GridsDT.Rows.Count > 0)
                GridInsetsOnly(Patricia1GridsDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.Patricia1GridInsetHeight), Convert.ToInt32(FrontMargins.Patricia1GridInsetWidth), OrderASC, IsBox);
            if (KansasGridsDT.Rows.Count > 0)
                GridInsetsOnly(KansasGridsDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.KansasGridInsetHeight), Convert.ToInt32(FrontMargins.KansasGridInsetWidth), OrderASC, IsBox);
            if (DakotaGridsDT.Rows.Count > 0)
                GridInsetsOnly(DakotaGridsDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.DakotaGridInsetHeight), Convert.ToInt32(FrontMargins.DakotaGridInsetWidth), OrderASC, IsBox);
            if (SofiaGridsDT.Rows.Count > 0)
                GridInsetsOnly(SofiaGridsDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.SofiaGridInsetHeight), Convert.ToInt32(FrontMargins.SofiaGridInsetWidth), OrderASC, IsBox);
            //if (!IsBox && SofiaAppliqueDT.Rows.Count > 0)
            //    GridInsetsOnly(SofiaAppliqueDT, ref DestinationDT,
            //        Convert.ToInt32(FrontMargins.SofiaGridInsetHeight), Convert.ToInt32(FrontMargins.SofiaGridInsetWidth), OrderASC, IsBox);
            if (Turin1GridsDT.Rows.Count > 0)
                GridInsetsOnly(Turin1GridsDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.Turin1GridInsetHeight), Convert.ToInt32(FrontMargins.Turin1GridInsetWidth), OrderASC, IsBox);
            if (Turin1_1GridsDT.Rows.Count > 0)
                GridInsetsOnly(Turin1_1GridsDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.Turin1GridInsetHeight), Convert.ToInt32(FrontMargins.Turin1GridInsetWidth), OrderASC, IsBox);
            if (Turin3GridsDT.Rows.Count > 0)
                GridInsetsOnly(Turin3GridsDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.Turin3GridInsetHeight), Convert.ToInt32(FrontMargins.Turin3GridInsetWidth), OrderASC, IsBox);
            if (LeonGridsDT.Rows.Count > 0)
                GridInsetsOnly(LeonGridsDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.LeonGridInsetHeight), Convert.ToInt32(FrontMargins.LeonGridInsetWidth), OrderASC, IsBox);
            if (InfinitiGridsDT.Rows.Count > 0)
                GridInsetsOnly(InfinitiGridsDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.InfinitiGridInsetHeight), Convert.ToInt32(FrontMargins.InfinitiGridInsetWidth), OrderASC, IsBox);
        }

        private void CollectBoxInsets(ref DataTable DestinationDT, bool OrderASC, bool IsBox)
        {
            if (LorenzoBoxesDT.Rows.Count > 0)
                BoxInsetsOnly(LorenzoBoxesDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.LorenzoBoxInsetHeight), Convert.ToInt32(FrontMargins.LorenzoBoxInsetWidth), OrderASC, IsBox);
            if (ElegantBoxesDT.Rows.Count > 0)
                BoxInsetsOnly(ElegantBoxesDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.ElegantBoxInsetHeight), Convert.ToInt32(FrontMargins.ElegantBoxInsetWidth), OrderASC, IsBox);
            if (Patricia1BoxesDT.Rows.Count > 0)
                BoxInsetsOnly(Patricia1BoxesDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.Patricia1BoxInsetHeight), Convert.ToInt32(FrontMargins.Patricia1BoxInsetWidth), OrderASC, IsBox);
            if (KansasBoxesDT.Rows.Count > 0)
                BoxInsetsOnly(KansasBoxesDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.KansasBoxInsetHeight), Convert.ToInt32(FrontMargins.KansasBoxInsetWidth), OrderASC, IsBox);
            if (DakotaBoxesDT.Rows.Count > 0)
                BoxInsetsOnly(DakotaBoxesDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.DakotaBoxInsetHeight), Convert.ToInt32(FrontMargins.DakotaBoxInsetWidth), OrderASC, IsBox);
            if (SofiaBoxesDT.Rows.Count > 0)
                BoxInsetsOnly(SofiaBoxesDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.SofiaBoxInsetHeight), Convert.ToInt32(FrontMargins.SofiaBoxInsetWidth), OrderASC, IsBox);
            if (Turin1BoxesDT.Rows.Count > 0)
                BoxInsetsOnly(Turin1BoxesDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.Turin1BoxInsetHeight), Convert.ToInt32(FrontMargins.Turin1BoxInsetWidth), OrderASC, IsBox);
            if (Turin3BoxesDT.Rows.Count > 0)
                BoxInsetsOnly(Turin3BoxesDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.Turin3BoxInsetHeight), Convert.ToInt32(FrontMargins.Turin3BoxInsetWidth), OrderASC, IsBox);
            if (LeonBoxesDT.Rows.Count > 0)
                BoxInsetsOnly(LeonBoxesDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.LeonBoxInsetHeight), Convert.ToInt32(FrontMargins.LeonBoxInsetWidth), OrderASC, IsBox);
            if (InfinitiBoxesDT.Rows.Count > 0)
                BoxInsetsOnly(InfinitiBoxesDT, ref DestinationDT,
                    Convert.ToInt32(FrontMargins.InfinitiBoxInsetHeight), Convert.ToInt32(FrontMargins.InfinitiBoxInsetWidth), OrderASC, IsBox);
        }

        #endregion

        public void ClearOrders()
        {
            FrontsID.Clear();
            BagetWithAngelOrdersDT.Clear();
            NotArchDecorOrdersDT.Clear();
            ArchDecorOrdersDT.Clear();
            GridsDecorOrdersDT.Clear();

            LorenzoOrdersDT.Clear();
            ElegantOrdersDT.Clear();
            Patricia1OrdersDT.Clear();
            KansasOrdersDT.Clear();
            DakotaOrdersDT.Clear();
            SofiaOrdersDT.Clear();
            Turin1OrdersDT.Clear();
            Turin1_1OrdersDT.Clear();
            Turin3OrdersDT.Clear();
            LeonOrdersDT.Clear();
            InfinitiOrdersDT.Clear();

            LorenzoCurvedOrdersDT.Clear();
            ElegantCurvedOrdersDT.Clear();
            Patricia1CurvedOrdersDT.Clear();
            KansasCurvedOrdersDT.Clear();
            SofiaCurvedOrdersDT.Clear();
            DakotaCurvedOrdersDT.Clear();
            Turin1CurvedOrdersDT.Clear();
            Turin1_1CurvedOrdersDT.Clear();
            Turin3CurvedOrdersDT.Clear();
            InfinitiCurvedOrdersDT.Clear();
        }

        public ArrayList GetFrontsID
        {
            set => FrontsID = value;
        }

        public bool GetOrders(int WorkAssignmentID, int FactoryID)
        {
            GetNotArchDecorOrders(ref NotArchDecorOrdersDT, WorkAssignmentID, FactoryID);
            GetBagetWithAngleOrders(ref BagetWithAngelOrdersDT, WorkAssignmentID, FactoryID);
            GetArchDecorOrders(ref ArchDecorOrdersDT, WorkAssignmentID, FactoryID);
            GetGridsDecorOrders(ref GridsDecorOrdersDT, WorkAssignmentID, FactoryID);

            GetCurvedFrontsOrders(ref LorenzoCurvedOrdersDT, WorkAssignmentID, FactoryID, @"(15580,15581,15582,15583,15584,15585,15586,15587)");
            GetCurvedFrontsOrders(ref ElegantCurvedOrdersDT, WorkAssignmentID, FactoryID, @"(-1)");
            GetCurvedFrontsOrders(ref Patricia1CurvedOrdersDT, WorkAssignmentID, FactoryID, @"(-1)");
            GetCurvedFrontsOrders(ref KansasCurvedOrdersDT, WorkAssignmentID, FactoryID, @"(1658,1659,1660,1661,1991,1992,1993,1994)");
            GetCurvedFrontsOrders(ref SofiaCurvedOrdersDT, WorkAssignmentID, FactoryID, @"(1654,1655,1656,1657,1987,1988,1989,1990)");
            GetCurvedFrontsOrders(ref DakotaCurvedOrdersDT, WorkAssignmentID, FactoryID, @"(29212,29214,29215,29216)");
            GetCurvedFrontsOrders(ref Turin1_1CurvedOrdersDT, WorkAssignmentID, FactoryID, @"(-1)");
            GetCurvedFrontsOrders(ref Turin1CurvedOrdersDT, WorkAssignmentID, FactoryID, @"(1646,1647,1648,1649,1979,1980,1981,1982)");
            GetCurvedFrontsOrders(ref Turin3CurvedOrdersDT, WorkAssignmentID, FactoryID, @"(1650,1651,1652,1653,1983,1984,1985,1986)");
            GetCurvedFrontsOrders(ref InfinitiCurvedOrdersDT, WorkAssignmentID, FactoryID, @"(14958,14959,14960,14961,14994,14995,14996,14997)");

            ProfileNamesDT.Clear();
            for (int i = 0; i < FrontsID.Count; i++)
            {
                if (Convert.ToInt32(Convert.ToInt32(FrontsID[i])) == Convert.ToInt32(Fronts.ElegantPat))
                {
                    GetProfileNames(ref ProfileNamesDT, WorkAssignmentID, FactoryID, Fronts.ElegantPat);
                    GetFrontsOrders(ref ElegantOrdersDT, WorkAssignmentID, FactoryID, Fronts.ElegantPat);
                }
                if (Convert.ToInt32(Convert.ToInt32(FrontsID[i])) == Convert.ToInt32(Fronts.Elegant))
                {
                    GetProfileNames(ref ProfileNamesDT, WorkAssignmentID, FactoryID, Fronts.Elegant);
                    GetFrontsOrders(ref ElegantOrdersDT, WorkAssignmentID, FactoryID, Fronts.Elegant);
                }
                if (Convert.ToInt32(Convert.ToInt32(FrontsID[i])) == Convert.ToInt32(Fronts.Patricia1))
                {
                    GetProfileNames(ref ProfileNamesDT, WorkAssignmentID, FactoryID, Fronts.Patricia1);
                    GetFrontsOrders(ref Patricia1OrdersDT, WorkAssignmentID, FactoryID, Fronts.Patricia1);
                }
                if (Convert.ToInt32(Convert.ToInt32(FrontsID[i])) == Convert.ToInt32(Fronts.Lorenzo))
                {
                    GetProfileNames(ref ProfileNamesDT, WorkAssignmentID, FactoryID, Fronts.Lorenzo);
                    GetFrontsOrders(ref LorenzoOrdersDT, WorkAssignmentID, FactoryID, Fronts.Lorenzo);
                }
                if (Convert.ToInt32(Convert.ToInt32(FrontsID[i])) == Convert.ToInt32(Fronts.Kansas))
                {
                    GetFrontsOrders(ref KansasOrdersDT, WorkAssignmentID, FactoryID, Fronts.Kansas);
                    GetProfileNames(ref ProfileNamesDT, WorkAssignmentID, FactoryID, Fronts.Kansas);
                }
                if (Convert.ToInt32(Convert.ToInt32(FrontsID[i])) == Convert.ToInt32(Fronts.KansasPat))
                {
                    GetFrontsOrders(ref KansasOrdersDT, WorkAssignmentID, FactoryID, Fronts.KansasPat);
                    GetProfileNames(ref ProfileNamesDT, WorkAssignmentID, FactoryID, Fronts.KansasPat);
                }
                if (Convert.ToInt32(Convert.ToInt32(FrontsID[i])) == Convert.ToInt32(Fronts.DakotaPat))
                {
                    GetFrontsOrders(ref DakotaOrdersDT, WorkAssignmentID, FactoryID, Fronts.DakotaPat);
                    GetProfileNames(ref ProfileNamesDT, WorkAssignmentID, FactoryID, Fronts.DakotaPat);
                }
                if (Convert.ToInt32(Convert.ToInt32(FrontsID[i])) == Convert.ToInt32(Fronts.Dakota))
                {
                    GetFrontsOrders(ref DakotaOrdersDT, WorkAssignmentID, FactoryID, Fronts.Dakota);
                    GetProfileNames(ref ProfileNamesDT, WorkAssignmentID, FactoryID, Fronts.Dakota);
                }
                if (Convert.ToInt32(Convert.ToInt32(FrontsID[i])) == Convert.ToInt32(Fronts.Sofia))
                {
                    GetFrontsOrders(ref SofiaOrdersDT, WorkAssignmentID, FactoryID, Fronts.Sofia);
                    GetProfileNames(ref ProfileNamesDT, WorkAssignmentID, FactoryID, Fronts.Sofia);
                }
                if (Convert.ToInt32(Convert.ToInt32(FrontsID[i])) == Convert.ToInt32(Fronts.Turin1))
                {
                    GetFrontsOrders(ref Turin1OrdersDT, WorkAssignmentID, FactoryID, Fronts.Turin1);
                    GetProfileNames(ref ProfileNamesDT, WorkAssignmentID, FactoryID, Fronts.Turin1);
                }
                if (Convert.ToInt32(Convert.ToInt32(FrontsID[i])) == Convert.ToInt32(Fronts.Turin1_1))
                {
                    GetFrontsOrders(ref Turin1_1OrdersDT, WorkAssignmentID, FactoryID, Fronts.Turin1_1);
                    GetProfileNames(ref ProfileNamesDT, WorkAssignmentID, FactoryID, Fronts.Turin1_1);
                }
                if (Convert.ToInt32(Convert.ToInt32(FrontsID[i])) == Convert.ToInt32(Fronts.Turin3))
                {
                    GetFrontsOrders(ref Turin3OrdersDT, WorkAssignmentID, FactoryID, Fronts.Turin3);
                    GetProfileNames(ref ProfileNamesDT, WorkAssignmentID, FactoryID, Fronts.Turin3);
                }
                if (Convert.ToInt32(Convert.ToInt32(FrontsID[i])) == Convert.ToInt32(Fronts.LeonTPS))
                {
                    GetFrontsOrders(ref LeonOrdersDT, WorkAssignmentID, FactoryID, Fronts.LeonTPS);
                    GetProfileNames(ref ProfileNamesDT, WorkAssignmentID, FactoryID, Fronts.LeonTPS);
                }
                if (Convert.ToInt32(Convert.ToInt32(FrontsID[i])) == Convert.ToInt32(Fronts.Infiniti))
                {
                    GetFrontsOrders(ref InfinitiOrdersDT, WorkAssignmentID, FactoryID, Fronts.Infiniti);
                    GetProfileNames(ref ProfileNamesDT, WorkAssignmentID, FactoryID, Fronts.Infiniti);
                }
                if (Convert.ToInt32(Convert.ToInt32(FrontsID[i])) == Convert.ToInt32(Fronts.InfinitiPat))
                {
                    GetFrontsOrders(ref InfinitiOrdersDT, WorkAssignmentID, FactoryID, Fronts.InfinitiPat);
                    GetProfileNames(ref ProfileNamesDT, WorkAssignmentID, FactoryID, Fronts.InfinitiPat);
                }
            }

            if (LorenzoCurvedOrdersDT.Rows.Count == 0 && ElegantCurvedOrdersDT.Rows.Count == 0 && Patricia1CurvedOrdersDT.Rows.Count == 0 && KansasCurvedOrdersDT.Rows.Count == 0 && SofiaCurvedOrdersDT.Rows.Count == 0 && DakotaCurvedOrdersDT.Rows.Count == 0 && Turin1CurvedOrdersDT.Rows.Count == 0 && Turin1_1CurvedOrdersDT.Rows.Count == 0 && Turin3CurvedOrdersDT.Rows.Count == 0 && InfinitiCurvedOrdersDT.Rows.Count == 0 &&
                LorenzoOrdersDT.Rows.Count == 0 && ElegantOrdersDT.Rows.Count == 0 && Patricia1OrdersDT.Rows.Count == 0 && KansasOrdersDT.Rows.Count == 0 && DakotaOrdersDT.Rows.Count == 0 && SofiaOrdersDT.Rows.Count == 0 && Turin1OrdersDT.Rows.Count == 0 && Turin1_1OrdersDT.Rows.Count == 0 && Turin3OrdersDT.Rows.Count == 0 && LeonOrdersDT.Rows.Count == 0 && InfinitiOrdersDT.Rows.Count == 0 &&
                BagetWithAngelOrdersDT.Rows.Count == 0 && NotArchDecorOrdersDT.Rows.Count == 0 && ArchDecorOrdersDT.Rows.Count == 0 && GridsDecorOrdersDT.Rows.Count == 0)
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

        public void CreateExcel(int WorkAssignmentID, Machines Machine, string ClientName, string BatchName, ref string sSourceFileName)
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

            int Admission = 0;
            string MachineName = string.Empty;

            switch (Machine)
            {
                case Machines.Balistrini:
                    Admission = 10;
                    MachineName = "Balistrini";
                    break;
                case Machines.ELME:
                    Admission = 20;
                    MachineName = "Elme";
                    break;
                case Machines.Rapid:
                    Admission = 30;
                    MachineName = "Rapid";
                    break;
                default:
                    break;
            }

            LorenzoSimpleDT.Clear();
            ElegantSimpleDT.Clear();
            Patricia1SimpleDT.Clear();
            KansasSimpleDT.Clear();
            DakotaSimpleDT.Clear();
            SofiaSimpleDT.Clear();
            Turin1SimpleDT.Clear();
            Turin1_1SimpleDT.Clear();
            Turin3SimpleDT.Clear();
            LeonSimpleDT.Clear();
            InfinitiSimpleDT.Clear();

            LorenzoVitrinaDT.Clear();
            ElegantVitrinaDT.Clear();
            Patricia1VitrinaDT.Clear();
            KansasVitrinaDT.Clear();
            DakotaVitrinaDT.Clear();
            SofiaVitrinaDT.Clear();
            Turin1VitrinaDT.Clear();
            Turin1_1VitrinaDT.Clear();
            Turin3VitrinaDT.Clear();
            LeonVitrinaDT.Clear();
            InfinitiVitrinaDT.Clear();

            Turin1RemovingBoxesDT.Clear();
            Turin3RemovingBoxesDT.Clear();

            GetVitrinaFronts(LorenzoOrdersDT, ref LorenzoVitrinaDT);
            GetVitrinaFronts(ElegantOrdersDT, ref ElegantVitrinaDT);
            GetVitrinaFronts(Patricia1OrdersDT, ref Patricia1VitrinaDT);
            GetVitrinaFronts(KansasOrdersDT, ref KansasVitrinaDT);
            GetVitrinaFronts(DakotaOrdersDT, ref DakotaVitrinaDT);
            GetVitrinaFronts(SofiaOrdersDT, ref SofiaVitrinaDT);
            GetVitrinaFronts(Turin1OrdersDT, ref Turin1VitrinaDT);
            GetVitrinaFronts(Turin1_1OrdersDT, ref Turin1_1VitrinaDT);
            GetVitrinaFronts(Turin3OrdersDT, ref Turin3VitrinaDT);
            GetVitrinaFronts(LeonOrdersDT, ref LeonVitrinaDT);
            GetVitrinaFronts(InfinitiOrdersDT, ref InfinitiVitrinaDT);

            GetRemovingBoxesFronts(Turin1OrdersDT, ref Turin1RemovingBoxesDT, 138);
            GetRemovingBoxesFronts(Turin1_1OrdersDT, ref Turin1RemovingBoxesDT, 138);
            GetRemovingBoxesFronts(Turin3OrdersDT, ref Turin3RemovingBoxesDT, 138);

            GetSimpleFronts(LorenzoOrdersDT, ref LorenzoSimpleDT, 222);
            GetSimpleFronts(ElegantOrdersDT, ref ElegantSimpleDT, 222);
            GetSimpleFronts(Patricia1OrdersDT, ref Patricia1SimpleDT, 222);
            GetSimpleFronts(KansasOrdersDT, ref KansasSimpleDT, 222);
            GetSimpleFronts(DakotaOrdersDT, ref DakotaSimpleDT, 222);
            GetSimpleFronts(SofiaOrdersDT, ref SofiaSimpleDT, 222);
            GetSimpleFronts(Turin1OrdersDT, ref Turin1SimpleDT, 175);
            GetSimpleFronts(Turin1_1OrdersDT, ref Turin1_1SimpleDT, 175);
            GetSimpleFronts(Turin3OrdersDT, ref Turin3SimpleDT, 175);
            GetSimpleFronts(LeonOrdersDT, ref LeonSimpleDT, 175);
            GetSimpleFronts(InfinitiOrdersDT, ref InfinitiSimpleDT, 222);

            LorenzoGridsDT.Clear();
            ElegantGridsDT.Clear();
            Patricia1GridsDT.Clear();
            KansasGridsDT.Clear();
            DakotaGridsDT.Clear();
            SofiaGridsDT.Clear();
            Turin1GridsDT.Clear();
            Turin1_1GridsDT.Clear();
            Turin3GridsDT.Clear();
            LeonGridsDT.Clear();
            InfinitiGridsDT.Clear();

            DakotaAppliqueDT.Clear();
            GetAppliqueFronts(DakotaOrdersDT, ref DakotaAppliqueDT);
            SofiaAppliqueDT.Clear();
            GetAppliqueFronts(SofiaOrdersDT, ref SofiaAppliqueDT);

            GetGridFronts(LorenzoOrdersDT, ref LorenzoGridsDT);
            GetGridFronts(ElegantOrdersDT, ref ElegantGridsDT);
            GetGridFronts(Patricia1OrdersDT, ref Patricia1GridsDT);
            GetGridFronts(KansasOrdersDT, ref KansasGridsDT);
            GetGridFronts(DakotaOrdersDT, ref DakotaGridsDT);
            GetGridFronts(SofiaOrdersDT, ref SofiaGridsDT);
            GetGridFronts(Turin1OrdersDT, ref Turin1GridsDT);
            GetGridFronts(Turin1_1OrdersDT, ref Turin1_1GridsDT);
            GetGridFronts(Turin3OrdersDT, ref Turin3GridsDT);
            GetGridFronts(LeonOrdersDT, ref LeonGridsDT);
            GetGridFronts(InfinitiOrdersDT, ref InfinitiGridsDT);

            LorenzoBoxesDT.Clear();
            ElegantBoxesDT.Clear();
            Patricia1BoxesDT.Clear();
            KansasBoxesDT.Clear();
            DakotaBoxesDT.Clear();
            SofiaBoxesDT.Clear();
            Turin1BoxesDT.Clear();
            Turin1_1BoxesDT.Clear();
            Turin3BoxesDT.Clear();
            LeonBoxesDT.Clear();
            InfinitiBoxesDT.Clear();

            GetBoxFronts(LorenzoOrdersDT, ref LorenzoBoxesDT, 222);
            GetBoxFronts(ElegantOrdersDT, ref ElegantBoxesDT, 222);
            GetBoxFronts(Patricia1OrdersDT, ref Patricia1BoxesDT, 222);
            GetBoxFronts(KansasOrdersDT, ref KansasBoxesDT, 222);
            GetBoxFronts(DakotaOrdersDT, ref DakotaBoxesDT, 222);
            GetBoxFronts(SofiaOrdersDT, ref SofiaBoxesDT, 222);
            GetBoxFronts(Turin1OrdersDT, ref Turin1BoxesDT, 175);
            GetBoxFronts(Turin1_1OrdersDT, ref Turin1_1BoxesDT, 175);
            GetBoxFronts(Turin3OrdersDT, ref Turin3BoxesDT, 175);
            GetBoxFronts(LeonOrdersDT, ref LeonBoxesDT, 175);
            GetBoxFronts(InfinitiOrdersDT, ref InfinitiBoxesDT, 222);

            BagetWithAngleAssemblyByMainOrderToExcel(ref hssfworkbook, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, BatchName);

            NotArchDecorAssemblyByMainOrderToExcel(ref hssfworkbook, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, BatchName);

            ArchDecorAssemblyByMainOrderToExcel(ref hssfworkbook, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, BatchName);

            CurvedAssemblyToExcel(ref hssfworkbook,
                Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, BatchName);

            GridsDecorAssemblyByMainOrderToExcel(ref hssfworkbook, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, BatchName);

            Admission = 0;

            InsetToExcel(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, 0, WorkAssignmentID, BatchName, ClientName);

            FilenkaToExcel(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, BatchName, ClientName);

            TrimmingToExcel(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, Admission, 0, WorkAssignmentID, BatchName, ClientName);

            AdditionsToExcel(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, BatchName, ClientName);

            GashToExcel(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, 0, WorkAssignmentID, BatchName, ClientName, MachineName);

            OrdersToExcel(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, BatchName, ClientName);

            DataTable DistFrameColorsDT = DistFrameColorsTable(LorenzoOrdersDT, true);
            AssemblyDT.Clear();
            FrontType = 0;
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                CollectAssemblySimple(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), LorenzoSimpleDT, ref AssemblyDT, FrontType);
                CollectAssemblyBoxes(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), LorenzoBoxesDT, ref AssemblyDT, FrontType);
                CollectAssemblyGrids(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), LorenzoGridsDT, ref AssemblyDT, FrontType);
            }
            DistFrameColorsDT.Clear();
            DistFrameColorsDT = DistFrameColorsTable(ElegantOrdersDT, true);
            FrontType = 9;
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                CollectAssemblySimple(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), ElegantSimpleDT, ref AssemblyDT, FrontType);
                CollectAssemblyBoxes(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), ElegantBoxesDT, ref AssemblyDT, FrontType);
                CollectAssemblyGrids(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), ElegantGridsDT, ref AssemblyDT, FrontType);
            }
            DistFrameColorsDT.Clear();
            DistFrameColorsDT = DistFrameColorsTable(Patricia1OrdersDT, true);
            FrontType = 9;
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                CollectAssemblySimple(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), Patricia1SimpleDT, ref AssemblyDT, FrontType);
                CollectAssemblyBoxes(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), Patricia1BoxesDT, ref AssemblyDT, FrontType);
                CollectAssemblyGrids(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), Patricia1GridsDT, ref AssemblyDT, FrontType);
            }
            DistFrameColorsDT.Clear();
            DistFrameColorsDT = DistFrameColorsTable(KansasOrdersDT, true);
            FrontType = 8;
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                CollectAssemblySimple(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), KansasSimpleDT, ref AssemblyDT, FrontType);
                CollectAssemblyBoxes(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), KansasBoxesDT, ref AssemblyDT, FrontType);
                CollectAssemblyGrids(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), KansasGridsDT, ref AssemblyDT, FrontType);
            }
            DistFrameColorsDT.Clear();
            DistFrameColorsDT = DistFrameColorsTable(DakotaOrdersDT, true);
            FrontType = 7;
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                CollectDakotaAssemblySimple(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), DakotaSimpleDT, ref AssemblyDT, FrontType);
                CollectAssemblyBoxes(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), DakotaBoxesDT, ref AssemblyDT, FrontType);
                CollectAssemblyGrids(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), DakotaGridsDT, ref AssemblyDT, FrontType);
                CollectAssemblyApplique(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), DakotaAppliqueDT, ref AssemblyDT, FrontType);
            }
            DistFrameColorsDT.Clear();
            DistFrameColorsDT = DistFrameColorsTable(Turin1_1OrdersDT, true);
            FrontType = 6;
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                CollectAssemblySimple(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), Turin1_1SimpleDT, ref AssemblyDT, FrontType);
                CollectAssemblyBoxes(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), Turin1_1BoxesDT, ref AssemblyDT, FrontType);
                CollectAssemblyGrids(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), Turin1_1GridsDT, ref AssemblyDT, FrontType);
            }
            DistFrameColorsDT.Clear();
            DistFrameColorsDT = DistFrameColorsTable(SofiaOrdersDT, true);
            FrontType = 1;
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                CollectAssemblySimple(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), SofiaSimpleDT, ref AssemblyDT, FrontType);
                CollectAssemblyBoxes(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), SofiaBoxesDT, ref AssemblyDT, FrontType);
                CollectAssemblyGrids(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), SofiaGridsDT, ref AssemblyDT, FrontType);
                CollectAssemblyApplique(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), SofiaAppliqueDT, ref AssemblyDT, FrontType);
            }
            DistFrameColorsDT.Clear();
            DistFrameColorsDT = DistFrameColorsTable(Turin1OrdersDT, true);
            FrontType = 2;
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                CollectAssemblySimple(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), Turin1SimpleDT, ref AssemblyDT, FrontType);
                CollectAssemblyBoxes(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), Turin1BoxesDT, ref AssemblyDT, FrontType);
                CollectAssemblyGrids(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), Turin1GridsDT, ref AssemblyDT, FrontType);
            }
            DistFrameColorsDT.Clear();
            DistFrameColorsDT = DistFrameColorsTable(Turin3OrdersDT, true);
            FrontType = 3;
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                CollectAssemblySimple(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), Turin3SimpleDT, ref AssemblyDT, FrontType);
                CollectAssemblyBoxes(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), Turin3BoxesDT, ref AssemblyDT, FrontType);
                CollectAssemblyGrids(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), Turin3GridsDT, ref AssemblyDT, FrontType);
            }
            DistFrameColorsDT.Clear();
            DistFrameColorsDT = DistFrameColorsTable(LeonOrdersDT, true);
            FrontType = 4;
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                CollectAssemblySimple(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), LeonSimpleDT, ref AssemblyDT, FrontType);
                CollectAssemblyBoxes(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), LeonBoxesDT, ref AssemblyDT, FrontType);
                CollectAssemblyGrids(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), LeonGridsDT, ref AssemblyDT, FrontType);
            }
            DistFrameColorsDT.Clear();
            DistFrameColorsDT = DistFrameColorsTable(InfinitiOrdersDT, true);
            FrontType = 5;
            for (int i = 0; i < DistFrameColorsDT.Rows.Count; i++)
            {
                CollectAssemblySimple(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), InfinitiSimpleDT, ref AssemblyDT, FrontType);
                CollectAssemblyBoxes(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), InfinitiBoxesDT, ref AssemblyDT, FrontType);
                CollectAssemblyGrids(Convert.ToInt32(DistFrameColorsDT.Rows[i]["ColorID"]), InfinitiGridsDT, ref AssemblyDT, FrontType);
            }

            decimal div1 = 25;
            decimal div2 = 2.14m;
            decimal time = 0;
            decimal cost = 0;
            PlanningTimebyCount(AssemblyDT, div1, div2, ref time, ref cost);

            AssemblyToExcel(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, BatchName, ClientName, time, cost);

            DeyingToExcel(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, BatchName, ClientName, Machine);

            DeyingByMainOrderToExcel(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, BatchName);

            string FileName = WorkAssignmentID + " " + BatchName;
            string tempFolder = @"\\192.168.1.6\Public\USERS_2016\_ДЕЙСТВУЮЩИЕ\ПРОИЗВОДСТВО\ТПС\инфиниум\";
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

        private void PlanningCurved(DataTable table, ref decimal time, ref decimal cost)
        {
            decimal VitrinaCount = 0;
            decimal NotVitrinaCount = 0;

            for (int x = 0; x < table.Rows.Count; x++)
            {
                if (table.Rows[x]["Count"] != DBNull.Value)
                {
                    // Витрины
                    if (Convert.ToInt32(table.Rows[x]["InsetTypeID"]) == 1)
                        VitrinaCount += Convert.ToDecimal(table.Rows[x]["Count"]);
                    // Глухие
                    if (Convert.ToInt32(table.Rows[x]["InsetTypeID"]) != 1 && Convert.ToInt32(table.Rows[x]["InsetTypeID"]) != -1)
                        NotVitrinaCount += Convert.ToDecimal(table.Rows[x]["Count"]);
                }
            }

            time = decimal.Round(VitrinaCount / 4 + NotVitrinaCount / 2, 3, MidpointRounding.AwayFromZero);
            cost = decimal.Round(time * 2.14m, 2, MidpointRounding.AwayFromZero);
        }

        private void PlanningTimebyCount(DataTable table, decimal div1, decimal div2, ref decimal time, ref decimal cost)
        {
            for (int x = 0; x < table.Rows.Count; x++)
            {
                if (table.Rows[x]["Count"] != DBNull.Value)
                    time += Convert.ToInt32(table.Rows[x]["Count"]);
            }

            if (div1 != 0)
                time = decimal.Round(time / div1, 3, MidpointRounding.AwayFromZero);

            cost = decimal.Round(time * div2, 2, MidpointRounding.AwayFromZero);
        }

        private void GetPlanningFilenka(DataTable table, decimal div1, decimal div2, decimal div3, ref decimal time, ref decimal cost)
        {
            decimal filenkaCount = 0;
            decimal allCount = 0;

            for (int x = 0; x < table.Rows.Count; x++)
            {
                if (table.Rows[x]["Square"] != DBNull.Value)
                {
                    // Фл04, Фл07
                    string name = table.Rows[x]["Name"].ToString();
                    string substr1 = "Фл04";
                    string substr2 = "Фл07";

                    if (name.Contains((substr1)) || name.Contains((substr2)))
                        filenkaCount += Convert.ToDecimal(table.Rows[x]["Square"]);
                    else
                        allCount += Convert.ToDecimal(table.Rows[x]["Square"]);
                }
            }

            time = decimal.Round(filenkaCount / div1 + allCount / div2, 3, MidpointRounding.AwayFromZero);
            cost = decimal.Round(time * div3, 2, MidpointRounding.AwayFromZero);
        }

        private void PlanningTimebySquare(DataTable table, decimal div1, decimal div2, ref decimal time, ref decimal cost)
        {
            for (int x = 0; x < table.Rows.Count; x++)
            {
                if (table.Rows[x]["Square"] != DBNull.Value)
                    time += Convert.ToDecimal(table.Rows[x]["Square"]);
            }

            if (div1 != 0)
                time = decimal.Round(time / div1, 3, MidpointRounding.AwayFromZero);

            cost = decimal.Round(time * div2, 2, MidpointRounding.AwayFromZero);
        }

        private void PlanningTimeGashRapid(DataTable table, ref decimal time, ref decimal cost)
        {
            decimal sum1 = 0;
            decimal sum2 = 0;
            for (int x = 0; x < table.Rows.Count; x++)
            {
                int Height = Convert.ToInt32(table.Rows[x]["Height"]);

                if (table.Rows[x]["Count"] != DBNull.Value)
                {
                    if (Height <= 296)
                    {
                        sum1 += Convert.ToInt32(table.Rows[x]["Count"]);
                    }
                    else
                    {
                        sum2 += Convert.ToInt32(table.Rows[x]["Count"]);
                    }
                }
            }

            time = sum1 * 2 / 200 + sum2 / 200;
            cost = decimal.Round(time * 2.14m, 2, MidpointRounding.AwayFromZero);
        }

        private void PlanningTimeDeyning(DataTable table, Machines machine, ref decimal time, ref decimal cost)
        {
            int div = 7;
            decimal Square = 0;

            for (int x = 0; x < table.Rows.Count; x++)
            {
                if (table.Rows[x]["Square"] != DBNull.Value)
                {
                    Square += Convert.ToDecimal(table.Rows[x]["Square"]);
                }
            }

            if (machine == Machines.Balistrini)
                div = 5;
            time = decimal.Round(Square / div, 3, MidpointRounding.AwayFromZero);
            cost = decimal.Round(time * 1.25m, 2, MidpointRounding.AwayFromZero);
        }

        private void PlanningTimeMarketPacking(DataTable table, ref decimal time, ref decimal cost)
        {
            decimal VitrinaCount = 0;
            decimal Square = 0;

            for (int x = 0; x < table.Rows.Count; x++)
            {
                if (table.Rows[x]["Count"] != DBNull.Value)
                {
                    // Витрины
                    if (table.Rows[x]["InsetColor"].ToString() == "Витрина")
                        VitrinaCount += Convert.ToDecimal(table.Rows[x]["Count"]);
                    // квадратура всех фасадов
                    Square += Convert.ToDecimal(table.Rows[x]["Square"]);
                }
            }

            time = decimal.Round(VitrinaCount / 26 + Square / 8, 3, MidpointRounding.AwayFromZero);
            cost = decimal.Round(time * 2.14m, 2, MidpointRounding.AwayFromZero);
        }

        public void OrdersToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            int WorkAssignmentID, string BatchName, string ClientName)
        {
            int RowIndex = 0;
            HSSFSheet sheet1 = hssfworkbook.CreateSheet("Заказы");
            sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

            sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
            sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
            sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
            sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

            if (LorenzoBoxesDT.Rows.Count > 0 || LorenzoGridsDT.Rows.Count > 0 || LorenzoSimpleDT.Rows.Count > 0)
            {
                SummingOrders(LorenzoOrdersDT, LorenzoBoxesDT, LorenzoGridsDT, LorenzoSimpleDT, LorenzoVitrinaDT, "Лоренцо ШУФ", "Лоренцо РЕШ", "Лоренцо");
                OrdersToExcelSingly(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, WorkAssignmentID, BatchName, ClientName, ref RowIndex);
                RowIndex++;
                RowIndex++;
            }
            if (ElegantBoxesDT.Rows.Count > 0 || ElegantGridsDT.Rows.Count > 0 || ElegantSimpleDT.Rows.Count > 0)
            {
                SummingOrders(ElegantOrdersDT, ElegantBoxesDT, ElegantGridsDT, ElegantSimpleDT, ElegantVitrinaDT, "Элегант ШУФ", "Элегант РЕШ", "Элегант");
                OrdersToExcelSingly(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, WorkAssignmentID, BatchName, ClientName, ref RowIndex);
                RowIndex++;
                RowIndex++;
            }
            if (Patricia1BoxesDT.Rows.Count > 0 || Patricia1GridsDT.Rows.Count > 0 || Patricia1SimpleDT.Rows.Count > 0)
            {
                SummingOrders(Patricia1OrdersDT, Patricia1BoxesDT, Patricia1GridsDT, Patricia1SimpleDT, Patricia1VitrinaDT, "Патриция-1 ШУФ", "Патриция-1 РЕШ", "Патриция-1");
                OrdersToExcelSingly(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, WorkAssignmentID, BatchName, ClientName, ref RowIndex);
                RowIndex++;
                RowIndex++;
            }
            if (KansasBoxesDT.Rows.Count > 0 || KansasGridsDT.Rows.Count > 0 || KansasSimpleDT.Rows.Count > 0)
            {
                SummingOrders(KansasOrdersDT, KansasBoxesDT, KansasGridsDT, KansasSimpleDT, KansasVitrinaDT, "Канзас ШУФ", "Канзас РЕШ", "Канзас");
                OrdersToExcelSingly(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, WorkAssignmentID, BatchName, ClientName, ref RowIndex);
                RowIndex++;
                RowIndex++;
            }
            if (DakotaBoxesDT.Rows.Count > 0 || DakotaGridsDT.Rows.Count > 0 || DakotaAppliqueDT.Rows.Count > 0 || DakotaSimpleDT.Rows.Count > 0
                || DakotaVitrinaDT.Rows.Count > 0)
            {
                SummingDakotaOrders(DakotaOrdersDT, DakotaBoxesDT, DakotaGridsDT, DakotaAppliqueDT, DakotaSimpleDT, DakotaVitrinaDT, "Дакота ШУФ", "Дакота РЕШ", "Дакота Аппл", "Дакота");
                OrdersToExcelSingly(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, WorkAssignmentID, BatchName, ClientName, ref RowIndex);
                RowIndex++;
                RowIndex++;
            }
            if (SofiaBoxesDT.Rows.Count > 0 || SofiaGridsDT.Rows.Count > 0 || SofiaAppliqueDT.Rows.Count > 0 || SofiaSimpleDT.Rows.Count > 0)
            {
                SummingSofiaOrders(SofiaOrdersDT, SofiaBoxesDT, SofiaGridsDT, SofiaAppliqueDT, SofiaSimpleDT, SofiaVitrinaDT, "София ШУФ", "София РЕШ", "София Аппл", "София");
                OrdersToExcelSingly(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, WorkAssignmentID, BatchName, ClientName, ref RowIndex);
                RowIndex++;
                RowIndex++;
            }
            if (Turin1BoxesDT.Rows.Count > 0 || Turin1GridsDT.Rows.Count > 0 || Turin1SimpleDT.Rows.Count > 0)
            {
                SummingOrders(Turin1OrdersDT, Turin1BoxesDT, Turin1GridsDT, Turin1SimpleDT, Turin1VitrinaDT, "Турин 1 ШУФ", "Турин 1 РЕШ", "Турин 1");
                OrdersToExcelSingly(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, WorkAssignmentID, BatchName, ClientName, ref RowIndex);
                RowIndex++;
                RowIndex++;
            }
            if (Turin1_1BoxesDT.Rows.Count > 0 || Turin1_1GridsDT.Rows.Count > 0 || Turin1_1SimpleDT.Rows.Count > 0)
            {
                SummingOrders(Turin1_1OrdersDT, Turin1_1BoxesDT, Turin1_1GridsDT, Turin1_1SimpleDT, Turin1_1VitrinaDT, "Турин 1 ШУФ", "Турин 1 РЕШ", "Турин 1");
                OrdersToExcelSingly(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, WorkAssignmentID, BatchName, ClientName, ref RowIndex);
                RowIndex++;
                RowIndex++;
            }
            if (Turin3BoxesDT.Rows.Count > 0 || Turin3GridsDT.Rows.Count > 0 || Turin3SimpleDT.Rows.Count > 0)
            {
                SummingOrders(Turin3OrdersDT, Turin3BoxesDT, Turin3GridsDT, Turin3SimpleDT, Turin3VitrinaDT, "Турин 3 ШУФ", "Турин 3 РЕШ", "Турин 3");
                OrdersToExcelSingly(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, WorkAssignmentID, BatchName, ClientName, ref RowIndex);
                RowIndex++;
                RowIndex++;
            }
            if (LeonBoxesDT.Rows.Count > 0 || LeonGridsDT.Rows.Count > 0 || LeonSimpleDT.Rows.Count > 0)
            {
                SummingOrders(LeonOrdersDT, LeonBoxesDT, LeonGridsDT, LeonSimpleDT, LeonVitrinaDT, "Леон ШУФ", "Леон РЕШ", "Леон");
                OrdersToExcelSingly(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, WorkAssignmentID, BatchName, ClientName, ref RowIndex);
                RowIndex++;
                RowIndex++;
            }
            if (InfinitiBoxesDT.Rows.Count > 0 || InfinitiGridsDT.Rows.Count > 0 || InfinitiSimpleDT.Rows.Count > 0)
            {
                SummingOrders(InfinitiOrdersDT, InfinitiBoxesDT, InfinitiGridsDT, InfinitiSimpleDT, InfinitiVitrinaDT, "Инфинити ШУФ", "Инфинити РЕШ", "Инфинити");
                OrdersToExcelSingly(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, WorkAssignmentID, BatchName, ClientName, ref RowIndex);
                RowIndex++;
                RowIndex++;
            }
        }

        private void DeyingToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            int WorkAssignmentID, string BatchName, string ClientName, Machines Machine)
        {
            DeyingDT.Clear();

            DataTable DT1 = new DataTable();

            using (DataView DV = new DataView(LorenzoOrdersDT, string.Empty, "ColorID", DataViewRowState.CurrentRows))
            {
                DT1.Clear();
                DT1 = DV.ToTable(true, new string[] { "ColorID" });
            }
            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                CollectDeying(Convert.ToInt32(DT1.Rows[i]["ColorID"]), LorenzoSimpleDT, ref DeyingDT, string.Empty);
                CollectDeying(Convert.ToInt32(DT1.Rows[i]["ColorID"]), LorenzoBoxesDT, ref DeyingDT, " ШУФ");
                CollectDeying(Convert.ToInt32(DT1.Rows[i]["ColorID"]), LorenzoGridsDT, ref DeyingDT, " РЕШ");
            }

            using (DataView DV = new DataView(ElegantOrdersDT, string.Empty, "ColorID", DataViewRowState.CurrentRows))
            {
                DT1.Clear();
                DT1 = DV.ToTable(true, new string[] { "ColorID" });
            }
            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                CollectDeying(Convert.ToInt32(DT1.Rows[i]["ColorID"]), ElegantSimpleDT, ref DeyingDT, string.Empty);
                CollectDeying(Convert.ToInt32(DT1.Rows[i]["ColorID"]), ElegantBoxesDT, ref DeyingDT, " ШУФ");
                CollectDeying(Convert.ToInt32(DT1.Rows[i]["ColorID"]), ElegantGridsDT, ref DeyingDT, " РЕШ");
            }

            using (DataView DV = new DataView(Patricia1OrdersDT, string.Empty, "ColorID", DataViewRowState.CurrentRows))
            {
                DT1.Clear();
                DT1 = DV.ToTable(true, new string[] { "ColorID" });
            }
            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                CollectDeying(Convert.ToInt32(DT1.Rows[i]["ColorID"]), Patricia1SimpleDT, ref DeyingDT, string.Empty);
                CollectDeying(Convert.ToInt32(DT1.Rows[i]["ColorID"]), Patricia1BoxesDT, ref DeyingDT, " ШУФ");
                CollectDeying(Convert.ToInt32(DT1.Rows[i]["ColorID"]), Patricia1GridsDT, ref DeyingDT, " РЕШ");
            }

            using (DataView DV = new DataView(KansasOrdersDT, string.Empty, "ColorID", DataViewRowState.CurrentRows))
            {
                DT1.Clear();
                DT1 = DV.ToTable(true, new string[] { "ColorID" });
            }
            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                CollectDeying(Convert.ToInt32(DT1.Rows[i]["ColorID"]), KansasSimpleDT, ref DeyingDT, string.Empty);
                CollectDeying(Convert.ToInt32(DT1.Rows[i]["ColorID"]), KansasBoxesDT, ref DeyingDT, " ШУФ");
                CollectDeying(Convert.ToInt32(DT1.Rows[i]["ColorID"]), KansasGridsDT, ref DeyingDT, " РЕШ");
            }

            using (DataView DV = new DataView(DakotaOrdersDT, string.Empty, "ColorID", DataViewRowState.CurrentRows))
            {
                DT1.Clear();
                DT1 = DV.ToTable(true, new string[] { "ColorID" });
            }
            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                CollectDakotaDeying(Convert.ToInt32(DT1.Rows[i]["ColorID"]), DakotaSimpleDT, ref DeyingDT, string.Empty);
                CollectDeying(Convert.ToInt32(DT1.Rows[i]["ColorID"]), DakotaBoxesDT, ref DeyingDT, " ШУФ");
                CollectDeying(Convert.ToInt32(DT1.Rows[i]["ColorID"]), DakotaGridsDT, ref DeyingDT, " РЕШ");
                CollectDeying(Convert.ToInt32(DT1.Rows[i]["ColorID"]), DakotaAppliqueDT, ref DeyingDT, " Аппл");
            }
            using (DataView DV = new DataView(SofiaOrdersDT, string.Empty, "ColorID", DataViewRowState.CurrentRows))
            {
                DT1.Clear();
                DT1 = DV.ToTable(true, new string[] { "ColorID" });
            }
            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                CollectDeying(Convert.ToInt32(DT1.Rows[i]["ColorID"]), SofiaSimpleDT, ref DeyingDT, string.Empty);
                CollectDeying(Convert.ToInt32(DT1.Rows[i]["ColorID"]), SofiaBoxesDT, ref DeyingDT, " ШУФ");
                CollectDeying(Convert.ToInt32(DT1.Rows[i]["ColorID"]), SofiaGridsDT, ref DeyingDT, " РЕШ");
                CollectDeying(Convert.ToInt32(DT1.Rows[i]["ColorID"]), SofiaAppliqueDT, ref DeyingDT, " Аппл");
            }
            using (DataView DV = new DataView(Turin1OrdersDT, string.Empty, "ColorID", DataViewRowState.CurrentRows))
            {
                DT1.Clear();
                DT1 = DV.ToTable(true, new string[] { "ColorID" });
            }
            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                CollectDeying(Convert.ToInt32(DT1.Rows[i]["ColorID"]), Turin1SimpleDT, ref DeyingDT, string.Empty);
                CollectDeying(Convert.ToInt32(DT1.Rows[i]["ColorID"]), Turin1BoxesDT, ref DeyingDT, " ШУФ");
                CollectDeying(Convert.ToInt32(DT1.Rows[i]["ColorID"]), Turin1GridsDT, ref DeyingDT, " РЕШ");
            }
            //using (DataView DV = new DataView(Turin1_1OrdersDT, string.Empty, "ColorID", DataViewRowState.CurrentRows))
            //{
            //    DT1.Clear();
            //    DT1 = DV.ToTable(true, new string[] { "ColorID" });
            //}
            //for (int i = 0; i < DT1.Rows.Count; i++)
            //{
            //    CollectDeying(Convert.ToInt32(DT1.Rows[i]["ColorID"]), Turin1_1SimpleDT, ref DeyingDT, string.Empty);
            //    CollectDeying(Convert.ToInt32(DT1.Rows[i]["ColorID"]), Turin1_1BoxesDT, ref DeyingDT, " ШУФ");
            //    CollectDeying(Convert.ToInt32(DT1.Rows[i]["ColorID"]), Turin1_1GridsDT, ref DeyingDT, " РЕШ");
            //}
            using (DataView DV = new DataView(Turin3OrdersDT, string.Empty, "ColorID", DataViewRowState.CurrentRows))
            {
                DT1.Clear();
                DT1 = DV.ToTable(true, new string[] { "ColorID" });
            }
            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                CollectDeying(Convert.ToInt32(DT1.Rows[i]["ColorID"]), Turin3SimpleDT, ref DeyingDT, string.Empty);
                CollectDeying(Convert.ToInt32(DT1.Rows[i]["ColorID"]), Turin3BoxesDT, ref DeyingDT, " ШУФ");
                CollectDeying(Convert.ToInt32(DT1.Rows[i]["ColorID"]), Turin3GridsDT, ref DeyingDT, " РЕШ");
            }
            using (DataView DV = new DataView(LeonOrdersDT, string.Empty, "ColorID", DataViewRowState.CurrentRows))
            {
                DT1.Clear();
                DT1 = DV.ToTable(true, new string[] { "ColorID" });
            }
            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                CollectDeying(Convert.ToInt32(DT1.Rows[i]["ColorID"]), LeonSimpleDT, ref DeyingDT, string.Empty);
                CollectDeying(Convert.ToInt32(DT1.Rows[i]["ColorID"]), LeonBoxesDT, ref DeyingDT, " ШУФ");
                CollectDeying(Convert.ToInt32(DT1.Rows[i]["ColorID"]), LeonGridsDT, ref DeyingDT, " РЕШ");
            }
            using (DataView DV = new DataView(InfinitiOrdersDT, string.Empty, "ColorID", DataViewRowState.CurrentRows))
            {
                DT1.Clear();
                DT1 = DV.ToTable(true, new string[] { "ColorID" });
            }
            for (int i = 0; i < DT1.Rows.Count; i++)
            {
                CollectDeying(Convert.ToInt32(DT1.Rows[i]["ColorID"]), InfinitiSimpleDT, ref DeyingDT, string.Empty);
                CollectDeying(Convert.ToInt32(DT1.Rows[i]["ColorID"]), InfinitiBoxesDT, ref DeyingDT, " ШУФ");
                CollectDeying(Convert.ToInt32(DT1.Rows[i]["ColorID"]), InfinitiGridsDT, ref DeyingDT, " РЕШ");
            }

            if (DeyingDT.Rows.Count > 0)
                DeyingToExcelSingly(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, WorkAssignmentID, BatchName, ClientName, "Покраска", Machine);
        }

        private void DeyingByMainOrderToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            int WorkAssignmentID, string BatchName)
        {
            DataTable DistMainOrdersDT = DistMainOrdersTable(LorenzoOrdersDT, ElegantOrdersDT, Patricia1OrdersDT, KansasOrdersDT, SofiaOrdersDT, Turin1OrdersDT, Turin1_1OrdersDT, Turin3OrdersDT, LeonOrdersDT, InfinitiOrdersDT, DakotaOrdersDT, true);
            DataTable DT = KansasOrdersDT.Clone();
            DataTable DT1 = new DataTable();
            DataTable ZOVOrdersNames = new DataTable();
            DataTable MarketOrdersNames = new DataTable();
            string MainOrdersID = string.Empty;
            string ClientName = string.Empty;
            string OrderName = string.Empty;

            if (DistMainOrdersDT.Select("GroupType=0").Count() > 0)
            {
                MainOrdersID = string.Empty;
                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=0"))
                    MainOrdersID += Convert.ToInt32(item["MainOrderID"]) + ",";
                MainOrdersID = MainOrdersID.Substring(0, MainOrdersID.Length - 1);
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT ClientName, MainOrders.DocNumber, MainOrders.MainOrderID, Batch.MegaBatchID, Batch.BatchID FROM MainOrders" +
                    " INNER JOIN BatchDetails ON MainOrders.MainOrderID=BatchDetails.MainOrderID" +
                    " INNER JOIN Batch ON BatchDetails.BatchID=Batch.BatchID" +
                    " INNER JOIN infiniu2_zovreference.dbo.Clients ON MainOrders.CLientID=infiniu2_zovreference.dbo.Clients.ClientID" +
                    " WHERE MainOrders.MainOrderID IN (" + MainOrdersID + ")", ConnectionStrings.ZOVOrdersConnectionString))
                {
                    DA.Fill(ZOVOrdersNames);
                }
            }
            if (DistMainOrdersDT.Select("GroupType=1").Count() > 0)
            {
                MainOrdersID = string.Empty;
                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=1"))
                    MainOrdersID += Convert.ToInt32(item["MainOrderID"]) + ",";
                MainOrdersID = MainOrdersID.Substring(0, MainOrdersID.Length - 1);
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT MegaOrders.ClientID, ClientName, MegaOrders.OrderNumber, MainOrders.MainOrderID, MainOrders.Notes, Batch.MegaBatchID, Batch.BatchID FROM MainOrders" +
                    " INNER JOIN BatchDetails ON MainOrders.MainOrderID=BatchDetails.MainOrderID" +
                    " INNER JOIN Batch ON BatchDetails.BatchID=Batch.BatchID" +
                    " INNER JOIN MegaOrders ON MainOrders.MegaOrderID=MegaOrders.MegaOrderID" +
                    " INNER JOIN infiniu2_marketingreference.dbo.Clients ON MegaOrders.ClientID=infiniu2_marketingreference.dbo.Clients.ClientID" +
                    " WHERE MainOrders.MainOrderID IN (" + MainOrdersID + ")", ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DA.Fill(MarketOrdersNames);
                }
            }

            if (DistMainOrdersDT.Select("GroupType=0").Count() > 0)
            {
                int RowIndex = 0;

                HSSFSheet sheet1 = hssfworkbook.CreateSheet("ЗОВ");
                sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

                sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
                sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
                sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
                sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

                sheet1.SetColumnWidth(0, 25 * 256);
                sheet1.SetColumnWidth(1, 20 * 256);
                sheet1.SetColumnWidth(2, 20 * 256);
                sheet1.SetColumnWidth(3, 6 * 256);
                sheet1.SetColumnWidth(4, 6 * 256);
                sheet1.SetColumnWidth(5, 6 * 256);

                for (int i = 0; i < DistMainOrdersDT.Rows.Count; i++)
                {
                    if (Convert.ToInt32(DistMainOrdersDT.Rows[i]["GroupType"]) == 1)
                        continue;

                    int MainOrderID = Convert.ToInt32(DistMainOrdersDT.Rows[i]["MainOrderID"]);
                    int MegaBatchID = 0;
                    int BatchID = 0;
                    DeyingDT.Clear();

                    using (DataView DV = new DataView(LorenzoOrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = LorenzoSimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = LorenzoBoxesDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " ШУФ");

                        DT.Clear();
                        rows = LorenzoGridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                    }

                    using (DataView DV = new DataView(ElegantOrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = ElegantSimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = ElegantBoxesDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " ШУФ");

                        DT.Clear();
                        rows = ElegantGridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                    }

                    using (DataView DV = new DataView(Patricia1OrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = Patricia1SimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = Patricia1BoxesDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " ШУФ");

                        DT.Clear();
                        rows = Patricia1GridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                    }

                    using (DataView DV = new DataView(KansasOrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = KansasSimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = KansasBoxesDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " ШУФ");

                        DT.Clear();
                        rows = KansasGridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                    }

                    using (DataView DV = new DataView(DakotaOrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = DakotaSimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDakotaDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = DakotaBoxesDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " ШУФ");

                        DT.Clear();
                        rows = DakotaGridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");

                        DT.Clear();
                        rows = DakotaAppliqueDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " АППЛ");
                    }

                    using (DataView DV = new DataView(SofiaOrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = SofiaSimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = SofiaBoxesDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " ШУФ");

                        DT.Clear();
                        rows = SofiaGridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");

                        DT.Clear();
                        rows = SofiaAppliqueDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " АППЛ");
                    }

                    using (DataView DV = new DataView(Turin1OrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = Turin1SimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = Turin1BoxesDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " ШУФ");

                        DT.Clear();
                        rows = Turin1GridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                    }

                    using (DataView DV = new DataView(Turin3OrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = Turin3SimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = Turin3BoxesDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " ШУФ");

                        DT.Clear();
                        rows = Turin3GridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                    }

                    using (DataView DV = new DataView(LeonOrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = LeonSimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = LeonBoxesDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " ШУФ");

                        DT.Clear();
                        rows = LeonGridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                    }

                    using (DataView DV = new DataView(InfinitiOrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = InfinitiSimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = InfinitiBoxesDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " ШУФ");

                        DT.Clear();
                        rows = InfinitiGridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                    }

                    DataRow[] CRows = ZOVOrdersNames.Select("MainOrderID=" + Convert.ToInt32(DistMainOrdersDT.Rows[i]["MainOrderID"]));
                    if (CRows.Count() > 0)
                    {
                        MegaBatchID = Convert.ToInt32(CRows[0]["MegaBatchID"]);
                        BatchID = Convert.ToInt32(CRows[0]["BatchID"]);
                        ClientName = CRows[0]["ClientName"].ToString();
                        OrderName = CRows[0]["DocNumber"].ToString();
                    }
                    if (DeyingDT.Rows.Count > 0)
                        DeyingByMainOrderToExcelSingly(ref hssfworkbook, ref sheet1,
                            Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DeyingDT, WorkAssignmentID,
                            "ЗОВ " + MegaBatchID + ", " + BatchID + ", " + MainOrderID, ClientName, OrderName, string.Empty, ref RowIndex);
                }
            }

            if (DistMainOrdersDT.Select("GroupType=1").Count() > 0)
            {
                int RowIndex = 0;

                HSSFSheet sheet1 = hssfworkbook.CreateSheet("Маркет");
                sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

                sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
                sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
                sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
                sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

                sheet1.SetColumnWidth(0, 25 * 256);
                sheet1.SetColumnWidth(1, 20 * 256);
                sheet1.SetColumnWidth(2, 20 * 256);
                sheet1.SetColumnWidth(3, 6 * 256);
                sheet1.SetColumnWidth(4, 6 * 256);
                sheet1.SetColumnWidth(5, 6 * 256);

                for (int i = 0; i < DistMainOrdersDT.Rows.Count; i++)
                {
                    if (Convert.ToInt32(DistMainOrdersDT.Rows[i]["GroupType"]) == 0)
                        continue;
                    int MainOrderID = Convert.ToInt32(DistMainOrdersDT.Rows[i]["MainOrderID"]);
                    int MegaBatchID = 0;
                    int BatchID = 0;
                    string Notes = string.Empty;

                    DeyingDT.Clear();
                    using (DataView DV = new DataView(LorenzoOrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = LorenzoSimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = LorenzoBoxesDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " ШУФ");

                        DT.Clear();
                        rows = LorenzoGridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                    }

                    using (DataView DV = new DataView(ElegantOrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = ElegantSimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = ElegantBoxesDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " ШУФ");

                        DT.Clear();
                        rows = ElegantGridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                    }

                    using (DataView DV = new DataView(Patricia1OrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = Patricia1SimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = Patricia1BoxesDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " ШУФ");

                        DT.Clear();
                        rows = Patricia1GridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                    }

                    using (DataView DV = new DataView(KansasOrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = KansasSimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = KansasBoxesDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " ШУФ");

                        DT.Clear();
                        rows = KansasGridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                    }

                    using (DataView DV = new DataView(DakotaOrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = DakotaSimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDakotaDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = DakotaBoxesDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " ШУФ");

                        DT.Clear();
                        rows = DakotaGridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");

                        DT.Clear();
                        rows = DakotaAppliqueDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " АППЛ");
                    }

                    using (DataView DV = new DataView(SofiaOrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = SofiaSimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = SofiaBoxesDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " ШУФ");

                        DT.Clear();
                        rows = SofiaGridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");

                        DT.Clear();
                        rows = SofiaAppliqueDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " АППЛ");
                    }

                    using (DataView DV = new DataView(Turin1OrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = Turin1SimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = Turin1BoxesDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " ШУФ");

                        DT.Clear();
                        rows = Turin1GridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                    }

                    using (DataView DV = new DataView(Turin3OrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = Turin3SimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = Turin3BoxesDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " ШУФ");

                        DT.Clear();
                        rows = Turin3GridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                    }

                    using (DataView DV = new DataView(LeonOrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = LeonSimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = LeonBoxesDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " ШУФ");

                        DT.Clear();
                        rows = LeonGridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                    }

                    using (DataView DV = new DataView(InfinitiOrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                    {
                        DT1.Clear();
                        DT1 = DV.ToTable(true, new string[] { "ColorID" });
                    }
                    for (int j = 0; j < DT1.Rows.Count; j++)
                    {
                        DT.Clear();
                        DataRow[] rows = InfinitiSimpleDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                        DT.Clear();
                        rows = InfinitiBoxesDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " ШУФ");

                        DT.Clear();
                        rows = InfinitiGridsDT.Select("MainOrderID=" + MainOrderID);
                        foreach (DataRow item in rows)
                            DT.Rows.Add(item.ItemArray);
                        CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                    }

                    string C = "Маркетинг ";
                    DataRow[] CRows = MarketOrdersNames.Select("MainOrderID=" + Convert.ToInt32(DistMainOrdersDT.Rows[i]["MainOrderID"]));
                    if (CRows.Count() > 0)
                    {
                        MegaBatchID = Convert.ToInt32(CRows[0]["MegaBatchID"]);
                        BatchID = Convert.ToInt32(CRows[0]["BatchID"]);
                        ClientName = CRows[0]["ClientName"].ToString();
                        Notes = CRows[0]["Notes"].ToString();
                        OrderName = "№" + CRows[0]["OrderNumber"].ToString() + "-" + MainOrderID;
                        if (Convert.ToInt32(CRows[0]["ClientID"]) == 101)
                            C = "Москва-1 ";
                    }
                    if (DeyingDT.Rows.Count > 0)
                        DeyingByMainOrderToExcelSingly(ref hssfworkbook, ref sheet1,
                            Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, DeyingDT, WorkAssignmentID,
                            C + MegaBatchID + ", " + BatchID + ", " + MainOrderID, ClientName, OrderName, Notes, ref RowIndex);
                }
            }
            for (int i = 0; i < DistMainOrdersDT.Rows.Count; i++)
            {
                int MainOrderID = Convert.ToInt32(DistMainOrdersDT.Rows[i]["MainOrderID"]);
                int MegaBatchID = 0;
                int BatchID = 0;
                string Notes = string.Empty;

                DeyingDT.Clear();
                using (DataView DV = new DataView(LorenzoOrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                {
                    DT1.Clear();
                    DT1 = DV.ToTable(true, new string[] { "ColorID" });
                }
                for (int j = 0; j < DT1.Rows.Count; j++)
                {
                    DT.Clear();
                    DataRow[] rows = LorenzoSimpleDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item in rows)
                        DT.Rows.Add(item.ItemArray);
                    CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                    DT.Clear();
                    rows = LorenzoBoxesDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item in rows)
                        DT.Rows.Add(item.ItemArray);
                    CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " ШУФ");

                    DT.Clear();
                    rows = LorenzoGridsDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item in rows)
                        DT.Rows.Add(item.ItemArray);
                    CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                }

                using (DataView DV = new DataView(ElegantOrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                {
                    DT1.Clear();
                    DT1 = DV.ToTable(true, new string[] { "ColorID" });
                }
                for (int j = 0; j < DT1.Rows.Count; j++)
                {
                    DT.Clear();
                    DataRow[] rows = ElegantSimpleDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item in rows)
                        DT.Rows.Add(item.ItemArray);
                    CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                    DT.Clear();
                    rows = ElegantBoxesDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item in rows)
                        DT.Rows.Add(item.ItemArray);
                    CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " ШУФ");

                    DT.Clear();
                    rows = ElegantGridsDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item in rows)
                        DT.Rows.Add(item.ItemArray);
                    CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                }

                using (DataView DV = new DataView(Patricia1OrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                {
                    DT1.Clear();
                    DT1 = DV.ToTable(true, new string[] { "ColorID" });
                }
                for (int j = 0; j < DT1.Rows.Count; j++)
                {
                    DT.Clear();
                    DataRow[] rows = Patricia1SimpleDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item in rows)
                        DT.Rows.Add(item.ItemArray);
                    CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                    DT.Clear();
                    rows = Patricia1BoxesDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item in rows)
                        DT.Rows.Add(item.ItemArray);
                    CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " ШУФ");

                    DT.Clear();
                    rows = Patricia1GridsDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item in rows)
                        DT.Rows.Add(item.ItemArray);
                    CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                }

                using (DataView DV = new DataView(KansasOrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                {
                    DT1.Clear();
                    DT1 = DV.ToTable(true, new string[] { "ColorID" });
                }
                for (int j = 0; j < DT1.Rows.Count; j++)
                {
                    DT.Clear();
                    DataRow[] rows = KansasSimpleDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item in rows)
                        DT.Rows.Add(item.ItemArray);
                    CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                    DT.Clear();
                    rows = KansasBoxesDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item in rows)
                        DT.Rows.Add(item.ItemArray);
                    CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " ШУФ");

                    DT.Clear();
                    rows = KansasGridsDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item in rows)
                        DT.Rows.Add(item.ItemArray);
                    CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                }

                using (DataView DV = new DataView(DakotaOrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                {
                    DT1.Clear();
                    DT1 = DV.ToTable(true, new string[] { "ColorID" });
                }
                for (int j = 0; j < DT1.Rows.Count; j++)
                {
                    DT.Clear();
                    DataRow[] rows = DakotaSimpleDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item in rows)
                        DT.Rows.Add(item.ItemArray);
                    CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                    DT.Clear();
                    rows = DakotaBoxesDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item in rows)
                        DT.Rows.Add(item.ItemArray);
                    CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " ШУФ");

                    DT.Clear();
                    rows = DakotaGridsDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item in rows)
                        DT.Rows.Add(item.ItemArray);
                    CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");

                    DT.Clear();
                    rows = DakotaAppliqueDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item in rows)
                        DT.Rows.Add(item.ItemArray);
                    CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " АППЛ");
                }

                using (DataView DV = new DataView(SofiaOrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                {
                    DT1.Clear();
                    DT1 = DV.ToTable(true, new string[] { "ColorID" });
                }
                for (int j = 0; j < DT1.Rows.Count; j++)
                {
                    DT.Clear();
                    DataRow[] rows = SofiaSimpleDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item in rows)
                        DT.Rows.Add(item.ItemArray);
                    CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                    DT.Clear();
                    rows = SofiaBoxesDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item in rows)
                        DT.Rows.Add(item.ItemArray);
                    CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " ШУФ");

                    DT.Clear();
                    rows = SofiaGridsDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item in rows)
                        DT.Rows.Add(item.ItemArray);
                    CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");

                    DT.Clear();
                    rows = SofiaAppliqueDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item in rows)
                        DT.Rows.Add(item.ItemArray);
                    CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " АППЛ");
                }

                using (DataView DV = new DataView(Turin1OrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                {
                    DT1.Clear();
                    DT1 = DV.ToTable(true, new string[] { "ColorID" });
                }
                for (int j = 0; j < DT1.Rows.Count; j++)
                {
                    DT.Clear();
                    DataRow[] rows = Turin1SimpleDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item in rows)
                        DT.Rows.Add(item.ItemArray);
                    CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                    DT.Clear();
                    rows = Turin1BoxesDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item in rows)
                        DT.Rows.Add(item.ItemArray);
                    CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " ШУФ");

                    DT.Clear();
                    rows = Turin1GridsDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item in rows)
                        DT.Rows.Add(item.ItemArray);
                    CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                }

                //using (DataView DV = new DataView(Turin1_1OrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                //{
                //    DT1.Clear();
                //    DT1 = DV.ToTable(true, new string[] { "ColorID" });
                //}
                //for (int j = 0; j < DT1.Rows.Count; j++)
                //{
                //    DT.Clear();
                //    DataRow[] rows = Turin1_1SimpleDT.Select("MainOrderID=" + MainOrderID);
                //    foreach (DataRow item in rows)
                //        DT.Rows.Add(item.ItemArray);
                //    CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                //    DT.Clear();
                //    rows = Turin1_1BoxesDT.Select("MainOrderID=" + MainOrderID);
                //    foreach (DataRow item in rows)
                //        DT.Rows.Add(item.ItemArray);
                //    CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " ШУФ");

                //    DT.Clear();
                //    rows = Turin1_1GridsDT.Select("MainOrderID=" + MainOrderID);
                //    foreach (DataRow item in rows)
                //        DT.Rows.Add(item.ItemArray);
                //    CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                //}

                using (DataView DV = new DataView(Turin3OrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                {
                    DT1.Clear();
                    DT1 = DV.ToTable(true, new string[] { "ColorID" });
                }
                for (int j = 0; j < DT1.Rows.Count; j++)
                {
                    DT.Clear();
                    DataRow[] rows = Turin3SimpleDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item in rows)
                        DT.Rows.Add(item.ItemArray);
                    CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                    DT.Clear();
                    rows = Turin3BoxesDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item in rows)
                        DT.Rows.Add(item.ItemArray);
                    CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " ШУФ");

                    DT.Clear();
                    rows = Turin3GridsDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item in rows)
                        DT.Rows.Add(item.ItemArray);
                    CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                }

                using (DataView DV = new DataView(LeonOrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                {
                    DT1.Clear();
                    DT1 = DV.ToTable(true, new string[] { "ColorID" });
                }
                for (int j = 0; j < DT1.Rows.Count; j++)
                {
                    DT.Clear();
                    DataRow[] rows = LeonSimpleDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item in rows)
                        DT.Rows.Add(item.ItemArray);
                    CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                    DT.Clear();
                    rows = LeonBoxesDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item in rows)
                        DT.Rows.Add(item.ItemArray);
                    CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " ШУФ");

                    DT.Clear();
                    rows = LeonGridsDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item in rows)
                        DT.Rows.Add(item.ItemArray);
                    CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                }

                using (DataView DV = new DataView(InfinitiOrdersDT, "MainOrderID=" + MainOrderID, "ColorID", DataViewRowState.CurrentRows))
                {
                    DT1.Clear();
                    DT1 = DV.ToTable(true, new string[] { "ColorID" });
                }
                for (int j = 0; j < DT1.Rows.Count; j++)
                {
                    DT.Clear();
                    DataRow[] rows = InfinitiSimpleDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item in rows)
                        DT.Rows.Add(item.ItemArray);
                    CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, string.Empty);

                    DT.Clear();
                    rows = InfinitiBoxesDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item in rows)
                        DT.Rows.Add(item.ItemArray);
                    CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " ШУФ");

                    DT.Clear();
                    rows = InfinitiGridsDT.Select("MainOrderID=" + MainOrderID);
                    foreach (DataRow item in rows)
                        DT.Rows.Add(item.ItemArray);
                    CollectDeying(Convert.ToInt32(DT1.Rows[j]["ColorID"]), DT, ref DeyingDT, " РЕШ");
                }

                if (Convert.ToInt32(DistMainOrdersDT.Rows[i]["GroupType"]) == 0)
                {
                    DataRow[] CRows = ZOVOrdersNames.Select("MainOrderID=" + Convert.ToInt32(DistMainOrdersDT.Rows[i]["MainOrderID"]));
                    if (CRows.Count() > 0)
                    {
                        MegaBatchID = Convert.ToInt32(CRows[0]["MegaBatchID"]);
                        BatchID = Convert.ToInt32(CRows[0]["BatchID"]);
                        ClientName = CRows[0]["ClientName"].ToString();
                        OrderName = CRows[0]["DocNumber"].ToString();
                    }
                    if (DeyingDT.Rows.Count > 0)
                        DeyingByMainOrderToExcelSingly(ref hssfworkbook,
                            Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS,
                            WorkAssignmentID, "ЗОВ " + MegaBatchID + ", " + BatchID + ", " + MainOrderID, ClientName, OrderName, OrderName.Replace("/", "-"), Notes);
                }
                if (Convert.ToInt32(DistMainOrdersDT.Rows[i]["GroupType"]) == 1)
                {
                    string C = "Маркетинг ";
                    DataRow[] CRows = MarketOrdersNames.Select("MainOrderID=" + Convert.ToInt32(DistMainOrdersDT.Rows[i]["MainOrderID"]));
                    if (CRows.Count() > 0)
                    {
                        MegaBatchID = Convert.ToInt32(CRows[0]["MegaBatchID"]);
                        BatchID = Convert.ToInt32(CRows[0]["BatchID"]);
                        ClientName = CRows[0]["ClientName"].ToString();
                        Notes = CRows[0]["Notes"].ToString();
                        OrderName = "№" + CRows[0]["OrderNumber"].ToString() + "-" + MainOrderID;
                        if (Convert.ToInt32(CRows[0]["ClientID"]) == 101)
                            C = "Москва-1 ";
                    }
                    if (DeyingDT.Rows.Count > 0)
                        DeyingByMainOrderToExcelSingly(ref hssfworkbook,
                            Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS,
                            WorkAssignmentID, C + MegaBatchID + ", " + BatchID + ", " + MainOrderID, ClientName, OrderName, OrderName.Replace("/", "-"), Notes);
                }
            }

        }

        public void DeyingByMainOrderToExcelSingly(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            int WorkAssignmentID, string BatchName, string ClientName, string OrderName, string PageName, string Notes)
        {
            int RowIndex = 0;

            int Index = hssfworkbook.GetSheetIndex(PageName);

            int j = 0;
            string PageName1 = PageName;
            while (Index != -1)
            {
                PageName1 = PageName + "(" + j++ + ")";
                Index = hssfworkbook.GetSheetIndex(PageName1);
            }
            HSSFSheet sheet1 = hssfworkbook.CreateSheet(PageName1);
            sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

            sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
            sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
            sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
            sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

            sheet1.SetColumnWidth(0, 25 * 256);
            sheet1.SetColumnWidth(1, 20 * 256);
            sheet1.SetColumnWidth(2, 20 * 256);
            sheet1.SetColumnWidth(3, 6 * 256);
            sheet1.SetColumnWidth(4, 6 * 256);
            sheet1.SetColumnWidth(5, 6 * 256);
            DataTable DT = new DataTable();
            DataColumn Col1 = new DataColumn("Col1", System.Type.GetType("System.String"));
            DataColumn Col2 = new DataColumn("Col2", System.Type.GetType("System.String"));
            DataColumn Col3 = new DataColumn("Col3", System.Type.GetType("System.String"));

            HSSFCell cell = null;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Плановое время выполнения:");
            cell.CellStyle = Calibri11CS;
            RowIndex++;
            RowIndex++;
            if (DeyingDT.Rows.Count > 0)
            {
                //DT.Dispose();
                //Col1.Dispose();
                //Col2.Dispose();
                //Col3.Dispose();
                //DT = DeyingDT.Copy();
                //Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
                //Col1.SetOrdinal(6);
                //DT.Columns["Square"].SetOrdinal(7);
                //DyeingWomen1ToExcel(ref hssfworkbook,
                //        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT, WorkAssignmentID, ClientName, BatchName,
                //    "Жен1. (" + Security.CurrentUserShortName + " от " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + ")", ref RowIndex);
                //RowIndex++;

                DT.Dispose();
                Col1.Dispose();
                Col2.Dispose();
                Col3.Dispose();
                DT = DeyingDT.Copy();
                Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
                Col1.SetOrdinal(6);
                DT.Columns["Square"].SetOrdinal(7);
                DyeingWomen2ToExcel(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT, WorkAssignmentID, BatchName, ClientName, OrderName,
                    "Жен2. (" + Security.CurrentUserShortName + " от " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + ")", Notes, ref RowIndex);
                RowIndex++;

                DT.Dispose();
                Col1.Dispose();
                Col2.Dispose();
                Col3.Dispose();
                DT = DeyingDT.Copy();
                Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
                Col2 = DT.Columns.Add("Col2", System.Type.GetType("System.String"));
                Col3 = DT.Columns.Add("Col3", System.Type.GetType("System.String"));
                Col1.SetOrdinal(6);
                Col2.SetOrdinal(7);
                Col3.SetOrdinal(8);
                DT.Columns["Square"].SetOrdinal(9);
                DyeingMen1ToExcel(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT, WorkAssignmentID, BatchName, ClientName, OrderName,
                    "Муж1. (" + Security.CurrentUserShortName + " от " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + ")", Notes, ref RowIndex);
                RowIndex++;

                DT.Dispose();
                Col1.Dispose();
                Col2.Dispose();
                Col3.Dispose();
                DT = DeyingDT.Copy();
                Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
                Col1.SetOrdinal(6);
                DT.Columns["Square"].SetOrdinal(7);
                DyeingWomen3ToExcel(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT, WorkAssignmentID, BatchName, ClientName, OrderName,
                    "Жен3. (" + Security.CurrentUserShortName + " от " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + ")", Notes, ref RowIndex);
                RowIndex++;

                DT.Dispose();
                Col1.Dispose();
                Col2.Dispose();
                Col3.Dispose();
                DT = DeyingDT.Copy();
                Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
                Col1.SetOrdinal(6);
                DT.Columns["Square"].SetOrdinal(7);
                DyeingMen2ToExcel(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT, WorkAssignmentID, BatchName, ClientName, OrderName,
                    "Муж2. (" + Security.CurrentUserShortName + " от " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + ")", Notes, ref RowIndex);
                RowIndex++;

                DT.Dispose();
                Col1.Dispose();
                Col2.Dispose();
                Col3.Dispose();
                DT = DeyingDT.Copy();
                Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
                Col2 = DT.Columns.Add("Col2", System.Type.GetType("System.String"));
                Col1.SetOrdinal(6);
                Col2.SetOrdinal(7);
                DT.Columns["Square"].SetOrdinal(8);

                decimal div1 = 165;
                decimal div2 = 2.14m;
                decimal time = 0;
                decimal cost = 0;
                PlanningTimeMarketPacking(DT, ref time, ref cost);

                DyeingPackingToExcel(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT, WorkAssignmentID, BatchName, ClientName, OrderName,
                    "Упаковка. (" + Security.CurrentUserShortName + " от " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + ")", Notes, time, cost, ref RowIndex);
                RowIndex++;

                DT.Dispose();
                Col1.Dispose();
                Col2.Dispose();
                Col3.Dispose();
                DT = DeyingDT.Copy();
                Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
                Col1.SetOrdinal(6);
                DT.Columns["Square"].SetOrdinal(7);
                DT.Columns["Notes"].SetOrdinal(8);

                div1 = 48;
                div2 = 2.14m;
                time = 0;
                cost = 0;
                PlanningTimebyCount(DT, div1, div2, ref time, ref cost);

                DyeingBoringToExcel(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT, WorkAssignmentID, BatchName, ClientName, OrderName,
                    "Сверление. (" + Security.CurrentUserShortName + " от " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + ")", Notes, time, cost, ref RowIndex);
            }

            RowIndex++;
        }

        public void DeyingByMainOrderToExcelSingly(ref HSSFWorkbook hssfworkbook, ref HSSFSheet sheet1,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            DataTable DT, int WorkAssignmentID, string BatchName, string ClientName, string OrderName, string Notes, ref int RowIndex)
        {
            DataTable TempDT = new DataTable();
            DataColumn Col1 = new DataColumn("Col1", System.Type.GetType("System.String"));
            DataColumn Col2 = new DataColumn("Col2", System.Type.GetType("System.String"));
            DataColumn Col3 = new DataColumn("Col3", System.Type.GetType("System.String"));

            if (DT.Rows.Count > 0)
            {
                TempDT.Dispose();
                Col1.Dispose();
                Col2.Dispose();
                Col3.Dispose();
                TempDT = DT.Copy();
                Col1 = TempDT.Columns.Add("Col1", System.Type.GetType("System.String"));
                Col2 = TempDT.Columns.Add("Col2", System.Type.GetType("System.String"));
                Col1.SetOrdinal(6);
                Col2.SetOrdinal(7);
                TempDT.Columns["Square"].SetOrdinal(8);

                decimal div1 = 165;
                decimal div2 = 2.14m;
                decimal time = 0;
                decimal cost = 0;
                PlanningTimeMarketPacking(DT, ref time, ref cost);

                DyeingPackingToExcel(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, TempDT, WorkAssignmentID, BatchName, ClientName, OrderName,
                    "Упаковка. (" + Security.CurrentUserShortName + " от " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + ")", Notes, time, cost, ref RowIndex);
                RowIndex++;

                TempDT.Dispose();
                Col1.Dispose();
                Col2.Dispose();
                Col3.Dispose();
                TempDT = DT.Copy();
                Col1 = TempDT.Columns.Add("Col1", System.Type.GetType("System.String"));
                Col1.SetOrdinal(6);
                TempDT.Columns["Square"].SetOrdinal(7);
                TempDT.Columns["Notes"].SetOrdinal(8);

                div1 = 48;
                div2 = 2.14m;
                time = 0;
                cost = 0;
                PlanningTimebyCount(DT, div1, div2, ref time, ref cost);

                DyeingBoringToExcel(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, TempDT, WorkAssignmentID, BatchName, ClientName, OrderName,
                    "Сверление. (" + Security.CurrentUserShortName + " от " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + ")", Notes, time, cost, ref RowIndex);
            }

            RowIndex++;
        }

        private void SummingOrders(DataTable SourceDT, DataTable BoxesDT, DataTable GridsDT, DataTable SimpleDT, DataTable VitrinaDT,
            string BoxName, string GridName, string FrontName)
        {
            SummOrdersDT.Dispose();
            SummOrdersDT = new DataTable();
            SummOrdersDT.Columns.Add(new DataColumn("Sizes", Type.GetType("System.String")));
            SummOrdersDT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            SummOrdersDT.Columns.Add(new DataColumn("Width", Type.GetType("System.Int32")));

            DataRow NewRow1 = SummOrdersDT.NewRow();
            NewRow1["Sizes"] = "Профиль";
            SummOrdersDT.Rows.Add(NewRow1);

            DataTable DistinctSizesDT = InsetDistSizesTable(SourceDT, true);

            DataTable DT = SimpleDT.Clone();
            foreach (DataRow item in SimpleDT.Select("InsetTypeID<>4"))
                DT.Rows.Add(item.ItemArray);

            CollectOrders(DistinctSizesDT, DT, ref SummOrdersDT, 2, FrontName);
            CollectOrders(DistinctSizesDT, BoxesDT, ref SummOrdersDT, 1, BoxName);
            CollectOrders(DistinctSizesDT, VitrinaDT, ref SummOrdersDT, 2, FrontName);
            CollectOrders(DistinctSizesDT, GridsDT, ref SummOrdersDT, 3, GridName);
            SummOrdersDT.Columns.Add(new DataColumn("TotalAmount", Type.GetType("System.String")));

            using (DataView DV = new DataView(SummOrdersDT.Copy()))
            {
                SummOrdersDT.Clear();
                DV.Sort = "Height, Width";
                SummOrdersDT = DV.ToTable();
            }

            DataRow NewRow2 = SummOrdersDT.NewRow();
            NewRow2["Sizes"] = "Квадратура";
            SummOrdersDT.Rows.Add(NewRow2);

            DataRow NewRow3 = SummOrdersDT.NewRow();
            NewRow3["Sizes"] = "Кол-во";
            SummOrdersDT.Rows.Add(NewRow3);

            decimal TotalSquare = 0;
            int TotalCount = 0;

            for (int y = 0; y < SummOrdersDT.Columns.Count; y++)
            {
                if (SummOrdersDT.Columns[y].ColumnName == "Sizes" || SummOrdersDT.Columns[y].ColumnName == "TotalAmount"
                    || SummOrdersDT.Columns[y].ColumnName == "Height" || SummOrdersDT.Columns[y].ColumnName == "Width")
                    continue;
                decimal Square = 0;
                int Count = 0;

                for (int x = 0; x < SummOrdersDT.Rows.Count; x++)
                {
                    if (SummOrdersDT.Rows[x]["Height"] != DBNull.Value && SummOrdersDT.Rows[x]["Width"] != DBNull.Value && SummOrdersDT.Rows[x][y] != DBNull.Value)
                    {
                        Square += Convert.ToDecimal(SummOrdersDT.Rows[x]["Height"]) * Convert.ToDecimal(SummOrdersDT.Rows[x]["Width"]) * Convert.ToDecimal(SummOrdersDT.Rows[x][y]) / 1000000;
                        Count += Convert.ToInt32(SummOrdersDT.Rows[x][y]);
                        TotalSquare += Convert.ToDecimal(SummOrdersDT.Rows[x]["Height"]) * Convert.ToDecimal(SummOrdersDT.Rows[x]["Width"]) * Convert.ToDecimal(SummOrdersDT.Rows[x][y]) / 1000000;
                        TotalCount += Convert.ToInt32(SummOrdersDT.Rows[x][y]);
                    }
                }
                Square = decimal.Round(Square, 3, MidpointRounding.AwayFromZero);
                SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 1][y] = Count;
                SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 2][y] = Square;
            }
            TotalSquare = decimal.Round(TotalSquare, 3, MidpointRounding.AwayFromZero);
            SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 1]["TotalAmount"] = TotalCount;
            SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 2]["TotalAmount"] = TotalSquare;
        }

        private void SummingDakotaOrders(DataTable SourceDT, DataTable BoxesDT, DataTable GridsDT, DataTable AppliqueDT, DataTable SimpleDT, DataTable VitrinaDT,
            string BoxName, string GridName, string AppliqueName, string FrontName)
        {
            SummOrdersDT.Dispose();
            SummOrdersDT = new DataTable();
            SummOrdersDT.Columns.Add(new DataColumn("Sizes", Type.GetType("System.String")));
            SummOrdersDT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            SummOrdersDT.Columns.Add(new DataColumn("Width", Type.GetType("System.Int32")));

            DataRow NewRow1 = SummOrdersDT.NewRow();
            NewRow1["Sizes"] = "Профиль";
            SummOrdersDT.Rows.Add(NewRow1);

            DataTable DistinctSizesDT = InsetDistSizesTable(SourceDT, true);

            DataTable DT = SimpleDT.Clone();
            foreach (DataRow item in SimpleDT.Select("InsetTypeID<>4"))
                DT.Rows.Add(item.ItemArray);

            CollectOrders(DistinctSizesDT, DT, ref SummOrdersDT, 2, FrontName);
            CollectOrders(DistinctSizesDT, BoxesDT, ref SummOrdersDT, 1, BoxName);
            CollectDakotaVitrinaOrders(DistinctSizesDT, VitrinaDT, ref SummOrdersDT, 2, FrontName);
            CollectOrders(DistinctSizesDT, GridsDT, ref SummOrdersDT, 3, GridName);
            CollectOrders(DistinctSizesDT, AppliqueDT, ref SummOrdersDT, 4, AppliqueName);
            SummOrdersDT.Columns.Add(new DataColumn("TotalAmount", Type.GetType("System.String")));

            using (DataView DV = new DataView(SummOrdersDT.Copy()))
            {
                SummOrdersDT.Clear();
                DV.Sort = "Height, Width";
                SummOrdersDT = DV.ToTable();
            }

            DataRow NewRow2 = SummOrdersDT.NewRow();
            NewRow2["Sizes"] = "Квадратура";
            SummOrdersDT.Rows.Add(NewRow2);

            DataRow NewRow3 = SummOrdersDT.NewRow();
            NewRow3["Sizes"] = "Кол-во";
            SummOrdersDT.Rows.Add(NewRow3);

            decimal TotalSquare = 0;
            int TotalCount = 0;

            for (int y = 0; y < SummOrdersDT.Columns.Count; y++)
            {
                if (SummOrdersDT.Columns[y].ColumnName == "Sizes" || SummOrdersDT.Columns[y].ColumnName == "TotalAmount"
                    || SummOrdersDT.Columns[y].ColumnName == "Height" || SummOrdersDT.Columns[y].ColumnName == "Width")
                    continue;
                decimal Square = 0;
                int Count = 0;

                for (int x = 0; x < SummOrdersDT.Rows.Count; x++)
                {
                    if (SummOrdersDT.Rows[x]["Height"] != DBNull.Value && SummOrdersDT.Rows[x]["Width"] != DBNull.Value && SummOrdersDT.Rows[x][y] != DBNull.Value)
                    {
                        Square += Convert.ToDecimal(SummOrdersDT.Rows[x]["Height"]) * Convert.ToDecimal(SummOrdersDT.Rows[x]["Width"]) * Convert.ToDecimal(SummOrdersDT.Rows[x][y]) / 1000000;
                        Count += Convert.ToInt32(SummOrdersDT.Rows[x][y]);
                        TotalSquare += Convert.ToDecimal(SummOrdersDT.Rows[x]["Height"]) * Convert.ToDecimal(SummOrdersDT.Rows[x]["Width"]) * Convert.ToDecimal(SummOrdersDT.Rows[x][y]) / 1000000;
                        TotalCount += Convert.ToInt32(SummOrdersDT.Rows[x][y]);
                    }
                }
                Square = decimal.Round(Square, 3, MidpointRounding.AwayFromZero);
                SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 1][y] = Count;
                SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 2][y] = Square;
            }
            TotalSquare = decimal.Round(TotalSquare, 3, MidpointRounding.AwayFromZero);
            SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 1]["TotalAmount"] = TotalCount;
            SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 2]["TotalAmount"] = TotalSquare;
        }

        private void SummingSofiaOrders(DataTable SourceDT, DataTable BoxesDT, DataTable GridsDT, DataTable AppliqueDT, DataTable SimpleDT, DataTable VitrinaDT,
            string BoxName, string GridName, string AppliqueName, string FrontName)
        {
            SummOrdersDT.Dispose();
            SummOrdersDT = new DataTable();
            SummOrdersDT.Columns.Add(new DataColumn("Sizes", Type.GetType("System.String")));
            SummOrdersDT.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
            SummOrdersDT.Columns.Add(new DataColumn("Width", Type.GetType("System.Int32")));

            DataRow NewRow1 = SummOrdersDT.NewRow();
            NewRow1["Sizes"] = "Профиль";
            SummOrdersDT.Rows.Add(NewRow1);

            DataTable DistinctSizesDT = InsetDistSizesTable(SourceDT, true);

            DataTable DT = SimpleDT.Clone();
            foreach (DataRow item in SimpleDT.Select("InsetTypeID<>4"))
                DT.Rows.Add(item.ItemArray);

            CollectOrders(DistinctSizesDT, DT, ref SummOrdersDT, 2, FrontName);
            CollectOrders(DistinctSizesDT, BoxesDT, ref SummOrdersDT, 1, BoxName);
            CollectOrders(DistinctSizesDT, VitrinaDT, ref SummOrdersDT, 2, FrontName);
            CollectOrders(DistinctSizesDT, GridsDT, ref SummOrdersDT, 3, GridName);
            CollectOrders(DistinctSizesDT, AppliqueDT, ref SummOrdersDT, 4, AppliqueName);
            SummOrdersDT.Columns.Add(new DataColumn("TotalAmount", Type.GetType("System.String")));

            using (DataView DV = new DataView(SummOrdersDT.Copy()))
            {
                SummOrdersDT.Clear();
                DV.Sort = "Height, Width";
                SummOrdersDT = DV.ToTable();
            }

            DataRow NewRow2 = SummOrdersDT.NewRow();
            NewRow2["Sizes"] = "Квадратура";
            SummOrdersDT.Rows.Add(NewRow2);

            DataRow NewRow3 = SummOrdersDT.NewRow();
            NewRow3["Sizes"] = "Кол-во";
            SummOrdersDT.Rows.Add(NewRow3);

            decimal TotalSquare = 0;
            int TotalCount = 0;

            for (int y = 0; y < SummOrdersDT.Columns.Count; y++)
            {
                if (SummOrdersDT.Columns[y].ColumnName == "Sizes" || SummOrdersDT.Columns[y].ColumnName == "TotalAmount"
                    || SummOrdersDT.Columns[y].ColumnName == "Height" || SummOrdersDT.Columns[y].ColumnName == "Width")
                    continue;
                decimal Square = 0;
                int Count = 0;

                for (int x = 0; x < SummOrdersDT.Rows.Count; x++)
                {
                    if (SummOrdersDT.Rows[x]["Height"] != DBNull.Value && SummOrdersDT.Rows[x]["Width"] != DBNull.Value && SummOrdersDT.Rows[x][y] != DBNull.Value)
                    {
                        Square += Convert.ToDecimal(SummOrdersDT.Rows[x]["Height"]) * Convert.ToDecimal(SummOrdersDT.Rows[x]["Width"]) * Convert.ToDecimal(SummOrdersDT.Rows[x][y]) / 1000000;
                        Count += Convert.ToInt32(SummOrdersDT.Rows[x][y]);
                        TotalSquare += Convert.ToDecimal(SummOrdersDT.Rows[x]["Height"]) * Convert.ToDecimal(SummOrdersDT.Rows[x]["Width"]) * Convert.ToDecimal(SummOrdersDT.Rows[x][y]) / 1000000;
                        TotalCount += Convert.ToInt32(SummOrdersDT.Rows[x][y]);
                    }
                }
                Square = decimal.Round(Square, 3, MidpointRounding.AwayFromZero);
                SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 1][y] = Count;
                SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 2][y] = Square;
            }
            TotalSquare = decimal.Round(TotalSquare, 3, MidpointRounding.AwayFromZero);
            SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 1]["TotalAmount"] = TotalCount;
            SummOrdersDT.Rows[SummOrdersDT.Rows.Count - 2]["TotalAmount"] = TotalSquare;
        }

        public void TrimmingToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            int Admission, int BoxAdmission, int WorkAssignmentID, string BatchName, string ClientName)
        {
            TrimmingDT.Clear();

            TrimCollectSimpleFronts(ref TrimmingDT, Admission, true);
            TrimCollectGridFronts(ref TrimmingDT, Admission, true);
            TrimCollectBoxFronts(ref TrimmingDT, BoxAdmission, true);

            decimal div1 = 370;
            decimal div2 = 2.14m;
            decimal time = 0;
            decimal cost = 0;
            PlanningTimebyCount(TrimmingDT, div1, div2, ref time, ref cost);

            if (TrimmingDT.Rows.Count == 0)
                return;

            DataTable DT = TrimmingDT.Copy();
            DataColumn Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
            Col1.SetOrdinal(4);

            HSSFSheet sheet1 = hssfworkbook.CreateSheet("Торцовка");
            sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

            sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
            sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
            sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
            sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

            sheet1.SetColumnWidth(0, 25 * 256);
            sheet1.SetColumnWidth(1, 17 * 256);
            //sheet1.SetColumnWidth(2, 13 * 256);
            //sheet1.SetColumnWidth(3, 13 * 256);
            sheet1.SetColumnWidth(4, 11 * 256);

            int RowIndex = 0;

            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
            cell.CellStyle = Calibri11CS;


            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Плановое время выполнения:");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "ч/ч");
            cell.CellStyle = Calibri11CS;

            cell = sheet1.CreateRow(RowIndex++).CreateCell(2);
            cell.SetCellValue(Convert.ToDouble(time));
            cell.CellStyle = Calibri11CS;


            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Планово-премиальный фонд:");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "руб.");
            cell.CellStyle = Calibri11CS;

            cell = sheet1.CreateRow(RowIndex++).CreateCell(2);
            cell.SetCellValue(Convert.ToDouble(cost));
            cell.CellStyle = Calibri11CS;

            RowIndex++;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, "ТСК-01");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, ClientName);
            cell.CellStyle = CalibriBold11CS;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Профиль");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Цвет профиля");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Работник");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            decimal SticksCount = 0;
            int CType = 0;
            int PType = 0;
            int AllTotalAmount = 0;
            int Count = 0;
            int Height = 0;

            if (DT.Rows.Count > 0)
            {
                CType = Convert.ToInt32(DT.Rows[0]["ColorID"]);
                PType = Convert.ToInt32(DT.Rows[0]["ProfileType"]);
            }

            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value && DT.Rows[x]["Height"] != DBNull.Value)
                {
                    Count = Convert.ToInt32(DT.Rows[x]["Count"]);
                    Height = Convert.ToInt32(DT.Rows[x]["Height"]);
                    SticksCount += (Height + 4) * Count;
                    AllTotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    Count = Convert.ToInt32(DT.Rows[x]["Count"]);
                }

                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "ColorID" || DT.Columns[y].ColumnName == "ProfileType"
                        || DT.Columns[y].ColumnName == "VitrinaNotes" || DT.Columns[y].ColumnName == "GridNotes"
                        || DT.Columns[y].ColumnName == "VitrinaCount" || DT.Columns[y].ColumnName == "GridCount")
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

                if (x + 1 <= DT.Rows.Count - 1
                    && (CType != Convert.ToInt32(DT.Rows[x + 1]["ColorID"]) || PType != Convert.ToInt32(DT.Rows[x + 1]["ProfileType"])))
                {
                    RowIndex++;
                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "ColorID" || DT.Columns[y].ColumnName == "ProfileType"
                            || DT.Columns[y].ColumnName == "VitrinaNotes" || DT.Columns[y].ColumnName == "GridNotes"
                            || DT.Columns[y].ColumnName == "VitrinaCount" || DT.Columns[y].ColumnName == "GridCount")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    SticksCount = SticksCount * 1.15m / 2800;
                    SticksCount = decimal.Round(SticksCount, 1, MidpointRounding.AwayFromZero);
                    cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                    cell.SetCellValue(SticksCount + " палок");
                    cell.CellStyle = TableHeaderCS;

                    CType = Convert.ToInt32(DT.Rows[x + 1]["ColorID"]);
                    PType = Convert.ToInt32(DT.Rows[x + 1]["ProfileType"]);
                    Count = 0;
                    Height = 0;
                    SticksCount = 0;
                }

                if (x == DT.Rows.Count - 1)
                {
                    RowIndex++;
                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "ColorID" || DT.Columns[y].ColumnName == "ProfileType"
                            || DT.Columns[y].ColumnName == "VitrinaNotes" || DT.Columns[y].ColumnName == "GridNotes"
                            || DT.Columns[y].ColumnName == "VitrinaCount" || DT.Columns[y].ColumnName == "GridCount")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    SticksCount = SticksCount * 1.15m / 2800;
                    SticksCount = decimal.Round(SticksCount, 1, MidpointRounding.AwayFromZero);
                    cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                    cell.SetCellValue(SticksCount + " палок");
                    cell.CellStyle = TableHeaderCS;

                    RowIndex++;
                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "ColorID" || DT.Columns[y].ColumnName == "ProfileType"
                            || DT.Columns[y].ColumnName == "VitrinaNotes" || DT.Columns[y].ColumnName == "GridNotes"
                            || DT.Columns[y].ColumnName == "VitrinaCount" || DT.Columns[y].ColumnName == "GridCount")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;
                }
                RowIndex++;
            }

            CalibriBold11CS = hssfworkbook.CreateCellStyle();
            CalibriBold11CS.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.BottomBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.BorderLeft = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.LeftBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.BorderRight = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.RightBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.BorderTop = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.TopBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.SetFont(CalibriBold11F);

            //RowIndex++;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "задание начали:");
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "задание закончили:");
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //sheet1.AddMergedRegion(new NPOI.HSSF.Util.Region(RowIndex, 1, RowIndex, 4));
            //RowIndex++;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "работало человек:");
            //cell.CellStyle = CalibriBold11CS;
            //RowIndex++;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "№ станка:");
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "№ операции:");
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //sheet1.AddMergedRegion(new NPOI.HSSF.Util.Region(RowIndex, 1, RowIndex, 4));
        }

        public void AdditionsToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            int WorkAssignmentID, string BatchName, string ClientName)
        {
            RemovingQuarterDT.Clear();
            CollectRemovingQuarter(ref RemovingQuarterDT, true);

            RemovingBoxesDT.Clear();
            CollectRemovingBoxes(ref RemovingBoxesDT, true);

            GrooveGridsDT.Clear();
            CollectGrooveGrids(ref GrooveGridsDT, true);

            if (RemovingQuarterDT.Rows.Count == 0 && RemovingBoxesDT.Rows.Count == 0 && GrooveGridsDT.Rows.Count == 0)
                return;

            int RowIndex = 0;

            HSSFSheet sheet1 = hssfworkbook.CreateSheet("Доп. задания");
            sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

            sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
            sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
            sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
            sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

            sheet1.SetColumnWidth(0, 25 * 256);
            sheet1.SetColumnWidth(1, 11 * 256);
            sheet1.SetColumnWidth(2, 11 * 256);
            sheet1.SetColumnWidth(3, 7 * 256);
            sheet1.SetColumnWidth(4, 23 * 256);

            DataTable DT = new DataTable();
            DataColumn Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));

            if (RemovingQuarterDT.Rows.Count > 0)
            {
                DT.Dispose();
                Col1.Dispose();
                DT = RemovingQuarterDT.Copy();
                Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
                Col1.SetOrdinal(4);

                AdditionsToExcelSingly(ref hssfworkbook,
                    Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT, WorkAssignmentID, BatchName, ClientName, "Снятие четверти", ref RowIndex);
            }

            if (RemovingBoxesDT.Rows.Count > 0)
            {
                DT.Dispose();
                Col1.Dispose();
                DT = RemovingBoxesDT.Copy();
                Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
                Col1.SetOrdinal(4);

                AdditionsToExcelSingly(ref hssfworkbook,
                    Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT, WorkAssignmentID, BatchName, ClientName, "Снятие шуфляд", ref RowIndex);
            }

            if (GrooveGridsDT.Rows.Count > 0)
            {
                DT.Dispose();
                Col1.Dispose();
                DT = GrooveGridsDT.Copy();
                Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
                Col1.SetOrdinal(4);

                AdditionsToExcelSingly(ref hssfworkbook,
                    Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT, WorkAssignmentID, BatchName, ClientName, "Пазировка под решетку", ref RowIndex);
            }
        }

        public void AdditionsToExcelSingly(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            ref HSSFSheet sheet1, DataTable DT, int WorkAssignmentID, string BatchName, string ClientName, string PageName, ref int RowIndex)
        {
            //HSSFCell cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex + 2), 1, "УТВЕРЖДАЮ_____________");
            //cell.CellStyle = Calibri11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(0), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
            //cell.CellStyle = Calibri11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex + 4), 0, "Клиент:");
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex + 4), 1, ClientName);
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex + 5), 0, "Партия:");
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex + 5), 1, BatchName);
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex + 3), 0, "Задание №" + WorkAssignmentID.ToString());
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex + 3), 1, PageName);
            //cell.CellStyle = CalibriBold11CS;
            HSSFCell cell = null;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Плановое время выполнения:");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, PageName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, ClientName);
            cell.CellStyle = CalibriBold11CS;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Профиль");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Цвет профиля");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Работник");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            decimal SticksCount = 0;
            int CType = 0;
            int AllTotalAmount = 0;
            int Count = 0;
            int Height = 0;

            if (DT.Rows.Count > 0)
            {
                CType = Convert.ToInt32(DT.Rows[0]["ColorID"]);
            }

            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value && DT.Rows[x]["Height"] != DBNull.Value)
                {
                    Count = Convert.ToInt32(DT.Rows[x]["Count"]);
                    Height = Convert.ToInt32(DT.Rows[x]["Height"]);
                    SticksCount += (Height + 4) * Count;
                    AllTotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    Count = Convert.ToInt32(DT.Rows[x]["Count"]);
                }

                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "ColorID")
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

                if (x + 1 <= DT.Rows.Count - 1 && CType != Convert.ToInt32(DT.Rows[x + 1]["ColorID"]))
                {
                    RowIndex++;
                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "ColorID")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    SticksCount = SticksCount * 1.15m / 2800;
                    SticksCount = decimal.Round(SticksCount, 1, MidpointRounding.AwayFromZero);
                    cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                    cell.SetCellValue(SticksCount + " палок");
                    cell.CellStyle = TableHeaderCS;

                    CType = Convert.ToInt32(DT.Rows[x + 1]["ColorID"]);
                    Count = 0;
                    Height = 0;
                    SticksCount = 0;
                }

                if (x == DT.Rows.Count - 1)
                {
                    RowIndex++;
                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "ColorID")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    SticksCount = SticksCount * 1.15m / 2800;
                    SticksCount = decimal.Round(SticksCount, 1, MidpointRounding.AwayFromZero);
                    cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                    cell.SetCellValue(SticksCount + " палок");
                    cell.CellStyle = TableHeaderCS;

                    RowIndex++;
                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "ColorID")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;
                }
                RowIndex++;
            }

            CalibriBold11CS = hssfworkbook.CreateCellStyle();
            CalibriBold11CS.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.BottomBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.BorderLeft = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.LeftBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.BorderRight = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.RightBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.BorderTop = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.TopBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.SetFont(CalibriBold11F);

            //RowIndex++;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "задание начали:");
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "задание закончили:");
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //sheet1.AddMergedRegion(new NPOI.HSSF.Util.Region(RowIndex, 1, RowIndex, 4));
            //RowIndex++;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "работало человек:");
            //cell.CellStyle = CalibriBold11CS;
            //RowIndex++;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "№ станка:");
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "№ операции:");
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //sheet1.AddMergedRegion(new NPOI.HSSF.Util.Region(RowIndex, 1, RowIndex, 4));
        }

        public void GashToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            int Admission, int WorkAssignmentID, string BatchName, string ClientName, string MachineName)
        {
            int RowIndex = 0;

            HSSFSheet sheet1 = hssfworkbook.CreateSheet("Запил 45");
            sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

            sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
            sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
            sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
            sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

            sheet1.SetColumnWidth(0, 25 * 256);
            sheet1.SetColumnWidth(1, 17 * 256);
            sheet1.SetColumnWidth(4, 11 * 256);

            DataTable DT = new DataTable();
            DataColumn Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));

            if (MachineName == "Balistrini" || MachineName == "Elme")
            {
                GashDT.Clear();
                GashCollectSimpleFronts(ref GashDT, 0, true);
                GashCollectGridFronts(ref GashDT, 0, true);

                DT.Dispose();
                Col1.Dispose();
                DT = GashDT.Copy();
                Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
                Col1.SetOrdinal(4);

                decimal div1 = 165;
                decimal div2 = 2.14m;
                decimal time = 0;
                decimal cost = 0;
                PlanningTimebyCount(DT, div1, div2, ref time, ref cost);

                if (DT.Rows.Count > 0)
                    GashToExcelSingly(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1,
                        DT, WorkAssignmentID, BatchName, ClientName, MachineName, time, cost, ref RowIndex);
                RowIndex++;
                RowIndex++;

                GashDT.Clear();
                GashCollectBoxFronts(ref GashDT, Admission, true);

                DT.Dispose();
                Col1.Dispose();
                DT = GashDT.Copy();
                Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
                Col1.SetOrdinal(4);

                time = 0;
                cost = 0;
                PlanningTimeGashRapid(DT, ref time, ref cost);

                if (DT.Rows.Count > 0)
                    GashToExcelSingly(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT, WorkAssignmentID, BatchName, ClientName, "Rapid",
                        time, cost, ref RowIndex);
            }
            if (MachineName == "Rapid")
            {
                GashDT.Clear();
                GashCollectSimpleFronts(ref GashDT, Admission, true);
                GashCollectGridFronts(ref GashDT, Admission, true);
                GashCollectBoxFronts(ref GashDT, Admission, true);

                DT.Dispose();
                Col1.Dispose();
                DT = GashDT.Copy();
                Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
                Col1.SetOrdinal(4);

                decimal time = 0;
                decimal cost = 0;
                PlanningTimeGashRapid(DT, ref time, ref cost);
                if (DT.Rows.Count > 0)
                    GashToExcelSingly(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT, WorkAssignmentID, BatchName, ClientName, "Rapid",
                        time, cost, ref RowIndex);
            }
            RowIndex++;
        }

        public void InsetToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            int Admission, int WorkAssignmentID, string BatchName, string ClientName)
        {
            InsetDT.Clear();
            CollectGridInsets(ref InsetDT, false, false);
            if (InsetDT.Rows.Count == 0)
                return;

            int RowIndex = 0;

            HSSFSheet sheet1 = hssfworkbook.CreateSheet("Решетки2");
            sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

            sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
            sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
            sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
            sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

            sheet1.SetColumnWidth(0, 25 * 256);
            sheet1.SetColumnWidth(1, 11 * 256);
            sheet1.SetColumnWidth(2, 11 * 256);
            sheet1.SetColumnWidth(3, 7 * 256);
            sheet1.SetColumnWidth(4, 23 * 256);

            if (InsetDT.Rows.Count > 0)
            {
                DataTable DT = InsetDT.Copy();
                DataColumn Col1 = new DataColumn();

                Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
                Col1.SetOrdinal(4);

                decimal div1 = 7;
                decimal div2 = 2.14m;
                decimal time = 0;
                decimal cost = 0;
                PlanningTimebyCount(DT, div1, div2, ref time, ref cost);

                InsetToExcelSingly(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT, WorkAssignmentID, BatchName, ClientName,
                        string.Empty, "Сборка решеток", time, cost, ref RowIndex);
                RowIndex++;
                RowIndex++;

                InsetToExcelSingly(ref hssfworkbook,
                    Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT, WorkAssignmentID, BatchName, ClientName,
                    "ДУБЛЬ", "Сборка решеток", time, cost, ref RowIndex);
                RowIndex++;
                RowIndex++;

                div1 = 40;
                div2 = 2.14m;
                time = 0;
                cost = 0;
                PlanningTimebyCount(DT, div1, div2, ref time, ref cost);

                InsetToExcelSingly(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT, WorkAssignmentID, BatchName, ClientName,
                        string.Empty, "Пила DFTP-400", time, cost, ref RowIndex);
                RowIndex++;
                RowIndex++;
            }

            RowIndex++;
        }

        public void FilenkaToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            int WorkAssignmentID, string BatchName, string ClientName)
        {
            InsetDT.Clear();
            FilenkaBoxesDT.Clear();
            CollectBoxInsets(ref InsetDT, false, false);
            CollectSimpleInsets(ref InsetDT, false, true);

            foreach (DataRow item in InsetDT.Rows)
            {
                decimal TotalSquare = 0;
                DataRow NewRow = FilenkaBoxesDT.NewRow();
                NewRow["ColorType"] = item["ColorType"];
                NewRow["Name"] = item["Name"];
                NewRow["Height"] = item["Height"];
                NewRow["Width"] = item["Width"];
                NewRow["Count"] = item["Count"];
                TotalSquare = decimal.Round(Convert.ToDecimal(item["Height"]) * Convert.ToDecimal(item["Width"]) * Convert.ToDecimal(item["Count"]) / 1000000, 3, MidpointRounding.AwayFromZero);
                NewRow["Square"] = TotalSquare;
                FilenkaBoxesDT.Rows.Add(NewRow);
            }

            InsetDT.Clear();
            FilenkaSimpleDT.Clear();
            CollectSimpleInsets(ref InsetDT, false, false);
            foreach (DataRow item in InsetDT.Rows)
            {
                decimal TotalSquare = 0;
                DataRow NewRow = FilenkaSimpleDT.NewRow();
                NewRow["ColorType"] = item["ColorType"];
                NewRow["Name"] = item["Name"];
                NewRow["Height"] = item["Height"];
                NewRow["Width"] = item["Width"];
                NewRow["Count"] = item["Count"];
                TotalSquare = decimal.Round(Convert.ToDecimal(item["Height"]) * Convert.ToDecimal(item["Width"]) * Convert.ToDecimal(item["Count"]) / 1000000, 3, MidpointRounding.AwayFromZero);
                NewRow["Square"] = TotalSquare;
                FilenkaSimpleDT.Rows.Add(NewRow);
            }

            InsetDT.Clear();
            DakotaFilenkaSimpleDT.Clear();
            CollectDakotaInsets(ref InsetDT, false, false);
            foreach (DataRow item in InsetDT.Rows)
            {
                decimal TotalSquare = 0;
                DataRow NewRow = DakotaFilenkaSimpleDT.NewRow();
                NewRow["ColorType"] = item["ColorType"];
                NewRow["Name"] = item["Name"];
                NewRow["Height"] = item["Height"];
                NewRow["Width"] = item["Width"];
                NewRow["Count"] = item["Count"];
                TotalSquare = decimal.Round(Convert.ToDecimal(item["Height"]) * Convert.ToDecimal(item["Width"]) * Convert.ToDecimal(item["Count"]) / 1000000, 3, MidpointRounding.AwayFromZero);
                NewRow["Square"] = TotalSquare;
                DakotaFilenkaSimpleDT.Rows.Add(NewRow);
            }

            if (FilenkaBoxesDT.Rows.Count == 0 && FilenkaSimpleDT.Rows.Count == 0 && DakotaFilenkaSimpleDT.Rows.Count == 0)
                return;

            int RowIndex = 0;

            HSSFSheet sheet1 = hssfworkbook.CreateSheet("Филенка");
            sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

            sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
            sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
            sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
            sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

            sheet1.SetColumnWidth(0, 25 * 256);

            DataTable DT = new DataTable();
            DataColumn Col1 = new DataColumn("Col1", System.Type.GetType("System.String"));

            decimal div1 = 54;
            decimal div2 = 2.14m;
            decimal div3 = 2.14m;
            decimal time = 0;
            decimal cost = 0;

            if (FilenkaBoxesDT.Rows.Count > 0)
            {
                DT = FilenkaBoxesDT.Copy();
                Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
                Col1.SetOrdinal(4);

                PlanningTimebyCount(DT, div1, div2, ref time, ref cost);

                FilenkaBoxesToExcel(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT,
                        WorkAssignmentID, BatchName, ClientName, time, cost, ref RowIndex);
                RowIndex++;
                RowIndex++;
            }

            if (FilenkaSimpleDT.Rows.Count > 0)
            {
                DT.Dispose();
                Col1.Dispose();
                DT = FilenkaSimpleDT.Copy();
                Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
                Col1.SetOrdinal(4);

                div1 = 54;
                div2 = 2.14m;
                time = 0;
                cost = 0;
                PlanningTimebyCount(DT, div1, div2, ref time, ref cost);

                FilenkaSimple1ToExcel(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT,
                        WorkAssignmentID, BatchName, ClientName, "Вставка", time, cost, ref RowIndex);
                RowIndex++;
                RowIndex++;
            }

            if (FilenkaSimpleDT.Rows.Count > 0)
            {
                DT.Dispose();
                Col1.Dispose();
                DT = FilenkaSimpleDT.Copy();
                Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
                Col1.SetOrdinal(4); ;

                decimal time1 = 0;
                decimal cost1 = 0;
                decimal time2 = 0;
                decimal cost2 = 0;
                decimal time3 = 0;
                decimal cost3 = 0;
                decimal time4 = 0;
                decimal cost4 = 0;

                div1 = 8.5m;
                div2 = 9.5m;
                div3 = 2.14m;
                GetPlanningFilenka(DT, div1, div2, div3, ref time1, ref cost1);
                div1 = 14;
                div2 = 2.14m;
                PlanningTimebySquare(DT, div1, div2, ref time2, ref cost2);
                div1 = 7.2m;
                div2 = 10.3m;
                div3 = 2.14m;
                GetPlanningFilenka(DT, div1, div2, div3, ref time3, ref cost3);
                div1 = 7.2m;
                div2 = 10.3m;
                div3 = 1.25m;
                GetPlanningFilenka(DT, div1, div2, div3, ref time4, ref cost4);

                FilenkaSimpleFToExcel(ref hssfworkbook,
                    Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT,
                    WorkAssignmentID, BatchName, ClientName, "Филенка", time1, cost1, ref RowIndex);
                RowIndex++;
                RowIndex++;
                FilenkaSimpleKToExcel(ref hssfworkbook,
                    Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT,
                    WorkAssignmentID, BatchName, ClientName, "Филенка", time2, cost2, ref RowIndex);
                RowIndex++;
                RowIndex++;
                FilenkaSimpleP1ToExcel(ref hssfworkbook,
                    Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT,
                    WorkAssignmentID, BatchName, ClientName, "Филенка", time3, cost3, time4, cost4, ref RowIndex);
                RowIndex++;
                RowIndex++;
                FilenkaSimpleP2ToExcel(ref hssfworkbook,
                    Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT,
                    WorkAssignmentID, BatchName, ClientName, "Филенка", ref RowIndex);
                RowIndex++;
                RowIndex++;
            }

            if (DakotaFilenkaSimpleDT.Rows.Count > 0)
            {
                DT.Dispose();
                Col1.Dispose();
                DT = DakotaFilenkaSimpleDT.Copy();
                Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
                Col1.SetOrdinal(4);

                div1 = 54;
                div2 = 2.14m;
                time = 0;
                cost = 0;
                PlanningTimebyCount(DT, div1, div2, ref time, ref cost);

                FilenkaSimple1ToExcel(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT,
                        WorkAssignmentID, BatchName, ClientName, "Вставка", time, cost, ref RowIndex);
                RowIndex++;
                RowIndex++;
            }

            if (DakotaFilenkaSimpleDT.Rows.Count > 0)
            {
                DT.Dispose();
                Col1.Dispose();
                DT = DakotaFilenkaSimpleDT.Copy();
                Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
                Col1.SetOrdinal(4);

                decimal time1 = 0;
                decimal cost1 = 0;
                decimal time2 = 0;
                decimal cost2 = 0;
                decimal time3 = 0;
                decimal cost3 = 0;
                decimal time4 = 0;
                decimal cost4 = 0;

                div1 = 8.5m;
                div2 = 9.5m;
                div3 = 2.14m;
                GetPlanningFilenka(DT, div1, div2, div3, ref time1, ref cost1);
                div1 = 14;
                div2 = 2.14m;
                PlanningTimebySquare(DT, div1, div2, ref time2, ref cost2);
                div1 = 7.2m;
                div2 = 10.3m;
                div3 = 2.14m;
                GetPlanningFilenka(DT, div1, div2, div3, ref time3, ref cost3);
                div1 = 7.2m;
                div2 = 10.3m;
                div3 = 1.25m;
                GetPlanningFilenka(DT, div1, div2, div3, ref time4, ref cost4);

                FilenkaSimpleFToExcel(ref hssfworkbook,
                    Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT,
                    WorkAssignmentID, BatchName, ClientName, "Филенка", time1, cost1, ref RowIndex);
                RowIndex++;
                RowIndex++;
                FilenkaSimpleKToExcel(ref hssfworkbook,
                    Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT,
                    WorkAssignmentID, BatchName, ClientName, "Филенка", time2, cost2, ref RowIndex);
                RowIndex++;
                RowIndex++;
                FilenkaSimpleP1ToExcel(ref hssfworkbook,
                    Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT,
                    WorkAssignmentID, BatchName, ClientName, "Филенка", time3, cost3, time4, cost4, ref RowIndex);
                RowIndex++;
                RowIndex++;
                FilenkaSimpleP2ToExcel(ref hssfworkbook,
                    Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT,
                    WorkAssignmentID, BatchName, ClientName, "Филенка", ref RowIndex);
                RowIndex++;
                RowIndex++;
            }
            RowIndex++;
        }

        public void AssemblyToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            int WorkAssignmentID, string BatchName, string ClientName, decimal Time, decimal Cost)
        {
            int RowIndex = 0;
            HSSFSheet sheet1 = hssfworkbook.CreateSheet("Сборка");
            sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

            sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
            sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
            sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
            sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

            sheet1.SetColumnWidth(0, 25 * 256);
            sheet1.SetColumnWidth(1, 20 * 256);
            sheet1.SetColumnWidth(2, 20 * 256);
            sheet1.SetColumnWidth(3, 6 * 256);
            sheet1.SetColumnWidth(4, 6 * 256);
            sheet1.SetColumnWidth(5, 6 * 256);

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
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, "Сборка");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, ClientName);
            cell.CellStyle = CalibriBold11CS;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Профиль");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Цвет профиля");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Цвет наполнителя");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 5, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 6, "Квадратура");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 7, "Работник");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            int CType = -1;
            int AllTotalAmount = 0;
            int TotalAmount = 0;
            decimal AllTotalSquare = 0;
            decimal TotalSquare = 0;

            if (AssemblyDT.Rows.Count > 0)
            {
                CType = Convert.ToInt32(AssemblyDT.Rows[0]["ColorType"]);
            }

            for (int x = 0; x < AssemblyDT.Rows.Count; x++)
            {
                if (AssemblyDT.Rows[x]["Count"] != DBNull.Value)
                {
                    AllTotalAmount += Convert.ToInt32(AssemblyDT.Rows[x]["Count"]);
                    TotalAmount += Convert.ToInt32(AssemblyDT.Rows[x]["Count"]);
                }
                if (AssemblyDT.Rows[x]["Square"] != DBNull.Value)
                {
                    AllTotalSquare += Convert.ToDecimal(AssemblyDT.Rows[x]["Square"]);
                    TotalSquare += Convert.ToDecimal(AssemblyDT.Rows[x]["Square"]);
                }

                for (int y = 0; y < AssemblyDT.Columns.Count; y++)
                {
                    if (AssemblyDT.Columns[y].ColumnName == "FrontType" || AssemblyDT.Columns[y].ColumnName == "ColorType")
                        continue;

                    Type t = AssemblyDT.Rows[x][y].GetType();

                    if (t.Name == "Decimal")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToDouble(AssemblyDT.Rows[x][y]));
                        cell.CellStyle = TableHeaderDecCS;
                        continue;
                    }
                    if (t.Name == "Int32")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(Convert.ToInt32(AssemblyDT.Rows[x][y]));
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    if (t.Name == "String" || t.Name == "DBNull")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(AssemblyDT.Rows[x][y].ToString());
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }
                }
                if (x + 1 <= AssemblyDT.Rows.Count - 1)
                {
                    if (CType != Convert.ToInt32(AssemblyDT.Rows[x + 1]["ColorType"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < AssemblyDT.Columns.Count; y++)
                        {
                            if (AssemblyDT.Columns[y].ColumnName == "ColorType" || AssemblyDT.Columns[y].ColumnName == "FrontType")
                                continue;
                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(6);
                        cell.SetCellValue(Convert.ToDouble(TotalSquare));
                        cell.CellStyle = TableHeaderCS;

                        CType = Convert.ToInt32(AssemblyDT.Rows[x + 1]["ColorType"]);
                        TotalAmount = 0;
                        TotalSquare = 0;

                        RowIndex++;
                    }
                }

                if (x == AssemblyDT.Rows.Count - 1)
                {
                    RowIndex++;
                    for (int y = 0; y < AssemblyDT.Columns.Count; y++)
                    {
                        if (AssemblyDT.Columns[y].ColumnName == "FrontType" || AssemblyDT.Columns[y].ColumnName == "ColorType")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Итого:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                    cell.SetCellValue(TotalAmount);
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(6);
                    cell.SetCellValue(Convert.ToDouble(TotalSquare));
                    cell.CellStyle = TableHeaderDecCS;

                    RowIndex++;

                    for (int y = 0; y < AssemblyDT.Columns.Count; y++)
                    {
                        if (AssemblyDT.Columns[y].ColumnName == "FrontType" || AssemblyDT.Columns[y].ColumnName == "ColorType")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(6);
                    cell.SetCellValue(Convert.ToDouble(AllTotalSquare));
                    cell.CellStyle = TableHeaderDecCS;
                }
                RowIndex++;
            }

            CalibriBold11CS = hssfworkbook.CreateCellStyle();
            CalibriBold11CS.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.BottomBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.BorderLeft = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.LeftBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.BorderRight = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.RightBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.BorderTop = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.TopBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.SetFont(CalibriBold11F);

            //RowIndex++;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "задание начали:");
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //sheet1.AddMergedRegion(new NPOI.HSSF.Util.Region(RowIndex, 0, RowIndex, 1));
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "задание закончили:");
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //sheet1.AddMergedRegion(new NPOI.HSSF.Util.Region(RowIndex, 2, RowIndex, 4));
            //RowIndex++;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "работало человек:");
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //sheet1.AddMergedRegion(new NPOI.HSSF.Util.Region(RowIndex, 0, RowIndex, 1));
            //RowIndex++;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "№ станка:");
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //sheet1.AddMergedRegion(new NPOI.HSSF.Util.Region(RowIndex, 0, RowIndex, 1));
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "№ операции:");
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //sheet1.AddMergedRegion(new NPOI.HSSF.Util.Region(RowIndex, 2, RowIndex, 4));
        }

        public void DeyingToExcelSingly(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            int WorkAssignmentID, string BatchName, string ClientName, string PageName, Machines machine)
        {
            int RowIndex = 0;

            HSSFSheet sheet1 = hssfworkbook.CreateSheet(PageName);
            sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

            sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
            sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
            sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
            sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

            sheet1.SetColumnWidth(0, 25 * 256);
            sheet1.SetColumnWidth(1, 20 * 256);
            sheet1.SetColumnWidth(2, 20 * 256);
            sheet1.SetColumnWidth(3, 6 * 256);
            sheet1.SetColumnWidth(4, 6 * 256);
            sheet1.SetColumnWidth(5, 6 * 256);

            DataTable DT = new DataTable();
            DataColumn Col1 = new DataColumn("Col1", System.Type.GetType("System.String"));
            DataColumn Col2 = new DataColumn("Col2", System.Type.GetType("System.String"));
            DataColumn Col3 = new DataColumn("Col3", System.Type.GetType("System.String"));

            if (DeyingDT.Rows.Count > 0)
            {
                DT = DeyingDT.Copy();
                Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
                Col1.SetOrdinal(6);
                DT.Columns["Square"].SetOrdinal(7);

                decimal time = 0;
                decimal cost = 0;
                PlanningTimeDeyning(DT, machine, ref time, ref cost);

                DyeingWomen1ToExcel(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT, WorkAssignmentID, BatchName, ClientName, string.Empty,
                    "Жен1. (" + Security.CurrentUserShortName + " от " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + ")", string.Empty, time, cost, ref RowIndex);
                RowIndex++;

                DT.Dispose();
                Col1.Dispose();
                Col2.Dispose();
                Col3.Dispose();
                DT = DeyingDT.Copy();
                Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
                Col1.SetOrdinal(6);
                DT.Columns["Square"].SetOrdinal(7);
                DyeingWomen2ToExcel(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT, WorkAssignmentID, BatchName, ClientName, string.Empty,
                    "Жен2. (" + Security.CurrentUserShortName + " от " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + ")", string.Empty, ref RowIndex);
                RowIndex++;

                DT.Dispose();
                Col1.Dispose();
                Col2.Dispose();
                Col3.Dispose();
                DT = DeyingDT.Copy();
                Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
                Col2 = DT.Columns.Add("Col2", System.Type.GetType("System.String"));
                Col3 = DT.Columns.Add("Col3", System.Type.GetType("System.String"));
                Col1.SetOrdinal(6);
                Col2.SetOrdinal(7);
                Col3.SetOrdinal(8);
                DT.Columns["Square"].SetOrdinal(9);
                DyeingMen1ToExcel(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT, WorkAssignmentID, BatchName, ClientName, string.Empty,
                    "Муж1. (" + Security.CurrentUserShortName + " от " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + ")", string.Empty, ref RowIndex);
                RowIndex++;

                DT.Dispose();
                Col1.Dispose();
                Col2.Dispose();
                Col3.Dispose();
                DT = DeyingDT.Copy();
                Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
                Col1.SetOrdinal(6);
                DT.Columns["Square"].SetOrdinal(7);
                DyeingWomen3ToExcel(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT, WorkAssignmentID, BatchName, ClientName, string.Empty,
                    "Жен3. (" + Security.CurrentUserShortName + " от " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + ")", string.Empty, ref RowIndex);
                RowIndex++;

                DT.Dispose();
                Col1.Dispose();
                Col2.Dispose();
                Col3.Dispose();
                DT = DeyingDT.Copy();
                Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
                Col1.SetOrdinal(6);
                DT.Columns["Square"].SetOrdinal(7);
                DyeingMen2ToExcel(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT, WorkAssignmentID, BatchName, ClientName, string.Empty,
                    "Муж2. (" + Security.CurrentUserShortName + " от " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + ")", string.Empty, ref RowIndex);
                RowIndex++;

                DT.Dispose();
                Col1.Dispose();
                Col2.Dispose();
                Col3.Dispose();
                DT = DeyingDT.Copy();
                Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
                Col2 = DT.Columns.Add("Col2", System.Type.GetType("System.String"));
                Col1.SetOrdinal(6);
                Col2.SetOrdinal(7);
                DT.Columns["Square"].SetOrdinal(8);

                time = 0;
                cost = 0;
                PlanningTimeMarketPacking(DT, ref time, ref cost);

                DyeingPackingToExcel(ref hssfworkbook,
                        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT, WorkAssignmentID, BatchName, ClientName, string.Empty,
                    "Упаковка. (" + Security.CurrentUserShortName + " от " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + ")", string.Empty, time, cost, ref RowIndex);

                //DT.Dispose();
                //Col1.Dispose();
                //Col2.Dispose();
                //Col3.Dispose();
                //DT = DeyingDT.Copy();
                //Col1 = DT.Columns.Add("Col1", System.Type.GetType("System.String"));
                //Col1.SetOrdinal(6);
                //DT.Columns["Square"].SetOrdinal(7);
                //DT.Columns["Notes"].SetOrdinal(8);
                //DyeingBoringToExcel(ref hssfworkbook,
                //        Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, ref sheet1, DT, WorkAssignmentID, ClientName, BatchName,
                //    "Сверление. (" + Security.CurrentUserShortName + " от " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + ")", ref RowIndex);
            }
        }

        public void GashToExcelSingly(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            ref HSSFSheet sheet1, DataTable DT, int WorkAssignmentID, string BatchName, string ClientName, string PageName, decimal Time, decimal Cost, ref int RowIndex)
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
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, PageName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, ClientName);
            cell.CellStyle = CalibriBold11CS;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Профиль");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Цвет профиля");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Работник");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 5, "Витрины");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 6, "Решетки");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            decimal SticksCount = 0;
            int CType = 0;
            int PType = 0;
            int AllTotalAmount = 0;
            int Count = 0;
            int Height = 0;

            if (DT.Rows.Count > 0)
            {
                CType = Convert.ToInt32(DT.Rows[0]["ColorID"]);
                PType = Convert.ToInt32(DT.Rows[0]["ProfileType"]);
            }

            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value && DT.Rows[x]["Height"] != DBNull.Value)
                {
                    Count = Convert.ToInt32(DT.Rows[x]["Count"]);
                    Height = Convert.ToInt32(DT.Rows[x]["Height"]);
                    SticksCount += (Height + 4) * Count;
                    AllTotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    Count = Convert.ToInt32(DT.Rows[x]["Count"]);
                }

                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "ColorID" || DT.Columns[y].ColumnName == "ProfileType" ||
                        DT.Columns[y].ColumnName == "VitrinaCount" || DT.Columns[y].ColumnName == "GridCount")
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

                if (x + 1 <= DT.Rows.Count - 1
                    && (CType != Convert.ToInt32(DT.Rows[x + 1]["ColorID"]) || PType != Convert.ToInt32(DT.Rows[x + 1]["ProfileType"])))
                {
                    RowIndex++;
                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "ColorID" || DT.Columns[y].ColumnName == "ProfileType" ||
                            DT.Columns[y].ColumnName == "VitrinaCount" || DT.Columns[y].ColumnName == "GridCount")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    SticksCount = SticksCount * 1.15m / 2800;
                    SticksCount = decimal.Round(SticksCount, 1, MidpointRounding.AwayFromZero);
                    cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                    cell.SetCellValue(SticksCount + " палок");
                    cell.CellStyle = TableHeaderCS;

                    CType = Convert.ToInt32(DT.Rows[x + 1]["ColorID"]);
                    PType = Convert.ToInt32(DT.Rows[x + 1]["ProfileType"]);
                    Count = 0;
                    Height = 0;
                    SticksCount = 0;
                }

                if (x == DT.Rows.Count - 1)
                {
                    RowIndex++;
                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "ColorID" || DT.Columns[y].ColumnName == "ProfileType" ||
                            DT.Columns[y].ColumnName == "VitrinaCount" || DT.Columns[y].ColumnName == "GridCount")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    SticksCount = SticksCount * 1.15m / 2800;
                    SticksCount = decimal.Round(SticksCount, 1, MidpointRounding.AwayFromZero);
                    cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                    cell.SetCellValue(SticksCount + " палок");
                    cell.CellStyle = TableHeaderCS;

                    RowIndex++;
                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "ColorID" || DT.Columns[y].ColumnName == "ProfileType" ||
                            DT.Columns[y].ColumnName == "VitrinaCount" || DT.Columns[y].ColumnName == "GridCount")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;
                }
                RowIndex++;
            }

            CalibriBold11CS = hssfworkbook.CreateCellStyle();
            CalibriBold11CS.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.BottomBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.BorderLeft = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.LeftBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.BorderRight = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.RightBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.BorderTop = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.TopBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.SetFont(CalibriBold11F);

            //RowIndex++;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "задание начали:");
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "задание закончили:");
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //sheet1.AddMergedRegion(new NPOI.HSSF.Util.Region(RowIndex, 1, RowIndex, 4));
            //RowIndex++;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "работало человек:");
            //cell.CellStyle = CalibriBold11CS;
            //RowIndex++;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "№ станка:");
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "№ операции:");
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //sheet1.AddMergedRegion(new NPOI.HSSF.Util.Region(RowIndex, 1, RowIndex, 4));
        }

        public void InsetToExcelSingly(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            ref HSSFSheet sheet1, DataTable DT, int WorkAssignmentID, string BatchName, string ClientName, string OperationName, string PageName, decimal time, decimal cost, ref int RowIndex)
        {

            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
            cell.CellStyle = Calibri11CS;


            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Плановое время выполнения:");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "ч/ч");
            cell.CellStyle = Calibri11CS;

            cell = sheet1.CreateRow(RowIndex++).CreateCell(2);
            cell.SetCellValue(Convert.ToDouble(time));
            cell.CellStyle = Calibri11CS;


            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Планово-премиальный фонд:");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "руб.");
            cell.CellStyle = Calibri11CS;

            cell = sheet1.CreateRow(RowIndex++).CreateCell(2);
            cell.SetCellValue(Convert.ToDouble(cost));
            cell.CellStyle = Calibri11CS;

            RowIndex++;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, PageName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, ClientName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 4, OperationName);
            cell.CellStyle = CalibriBold11CS;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Цвет наполнителя");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Работник");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            int CType = -1;
            int AllTotalAmount = 0;
            int TotalAmount = 0;
            decimal AllTotalSquare = 0;
            decimal TotalSquare = 0;
            string str = string.Empty;
            if (DT.Rows.Count > 0)
                CType = Convert.ToInt32(DT.Rows[0]["ColorType"]);

            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value)
                {
                    AllTotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    TotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    AllTotalSquare += Convert.ToDecimal(DT.Rows[x]["Height"]) * Convert.ToDecimal(DT.Rows[x]["Width"]) * Convert.ToDecimal(DT.Rows[x]["Count"]) / 1000000;
                    TotalSquare += Convert.ToDecimal(DT.Rows[x]["Height"]) * Convert.ToDecimal(DT.Rows[x]["Width"]) * Convert.ToDecimal(DT.Rows[x]["Count"]) / 1000000;
                }

                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "ColorType")
                        continue;
                    Type t = DT.Rows[x][y].GetType();

                    if (DT.Columns[y].ColumnName == "Name")
                        str = DT.Rows[x][y].ToString();

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
                    if (CType != Convert.ToInt32(DT.Rows[x + 1]["ColorType"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "ColorType")
                                continue;
                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        TotalSquare = decimal.Round(TotalSquare, 3, MidpointRounding.AwayFromZero);
                        cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                        cell.SetCellValue("Квадратура: " + Convert.ToDouble(TotalSquare));
                        cell.CellStyle = TableHeaderCS;

                        CType = Convert.ToInt32(DT.Rows[x + 1]["ColorType"]);
                        TotalAmount = 0;
                        TotalSquare = 0;

                        RowIndex++;
                    }
                }

                if (x == DT.Rows.Count - 1)
                {
                    RowIndex++;
                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "ColorType")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Итого:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                    cell.SetCellValue(TotalAmount);
                    cell.CellStyle = TableHeaderCS;

                    TotalSquare = decimal.Round(TotalSquare, 3, MidpointRounding.AwayFromZero);
                    cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                    cell.SetCellValue("Квадратура: " + Convert.ToDouble(TotalSquare));
                    cell.CellStyle = TableHeaderDecCS;

                    RowIndex++;

                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "ColorType")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;

                    AllTotalSquare = decimal.Round(AllTotalSquare, 3, MidpointRounding.AwayFromZero);
                    cell = sheet1.CreateRow(RowIndex).CreateCell(4);
                    cell.SetCellValue("Квадратура: " + Convert.ToDouble(AllTotalSquare));
                    cell.CellStyle = TableHeaderDecCS;
                }
                RowIndex++;
            }

            CalibriBold11CS = hssfworkbook.CreateCellStyle();
            CalibriBold11CS.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.BottomBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.BorderLeft = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.LeftBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.BorderRight = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.RightBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.BorderTop = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.TopBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.SetFont(CalibriBold11F);

            //RowIndex++;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "задание начали:");
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //sheet1.AddMergedRegion(new NPOI.HSSF.Util.Region(RowIndex, 0, RowIndex, 1));
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "задание закончили:");
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //sheet1.AddMergedRegion(new NPOI.HSSF.Util.Region(RowIndex, 2, RowIndex, 4));
            //RowIndex++;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "работало человек:");
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //sheet1.AddMergedRegion(new NPOI.HSSF.Util.Region(RowIndex, 0, RowIndex, 1));
            //RowIndex++;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "№ станка:");
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //sheet1.AddMergedRegion(new NPOI.HSSF.Util.Region(RowIndex, 0, RowIndex, 1));
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "№ операции:");
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //sheet1.AddMergedRegion(new NPOI.HSSF.Util.Region(RowIndex, 2, RowIndex, 4));
        }

        public void FilenkaSimple1ToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            ref HSSFSheet sheet1, DataTable DT, int WorkAssignmentID, string BatchName, string ClientName, string PageName, decimal time, decimal cost, ref int RowIndex)
        {
            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
            cell.CellStyle = Calibri11CS;


            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Плановое время выполнения:");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "ч/ч");
            cell.CellStyle = Calibri11CS;

            cell = sheet1.CreateRow(RowIndex++).CreateCell(2);
            cell.SetCellValue(Convert.ToDouble(time));
            cell.CellStyle = Calibri11CS;


            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Планово-премиальный фонд:");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "руб.");
            cell.CellStyle = Calibri11CS;

            cell = sheet1.CreateRow(RowIndex++).CreateCell(2);
            cell.SetCellValue(Convert.ToDouble(cost));
            cell.CellStyle = Calibri11CS;

            RowIndex++;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, PageName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, ClientName);
            cell.CellStyle = CalibriBold11CS;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Филенка");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Пила");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 5, "м.кв.");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            int CType = -1;
            int AllTotalAmount = 0;
            int TotalAmount = 0;
            decimal AllTotalSquare = 0;
            decimal TotalSquare = 0;

            if (DT.Rows.Count > 0)
                CType = Convert.ToInt32(DT.Rows[0]["ColorType"]);

            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value)
                {
                    AllTotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    TotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                }
                if (DT.Rows[x]["Square"] != DBNull.Value)
                {
                    AllTotalSquare += Convert.ToDecimal(DT.Rows[x]["Square"]);
                    TotalSquare += Convert.ToDecimal(DT.Rows[x]["Square"]);
                }

                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "ColorType")
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
                    if (CType != Convert.ToInt32(DT.Rows[x + 1]["ColorType"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "ColorType")
                                continue;
                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        TotalSquare = decimal.Round(TotalSquare, 3, MidpointRounding.AwayFromZero);
                        cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                        cell.SetCellValue(Convert.ToDouble(TotalSquare));
                        cell.CellStyle = TableHeaderCS;

                        CType = Convert.ToInt32(DT.Rows[x + 1]["ColorType"]);
                        TotalAmount = 0;
                        TotalSquare = 0;
                        RowIndex++;
                    }
                }

                if (x == DT.Rows.Count - 1)
                {
                    RowIndex++;
                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "ColorType")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Итого:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                    cell.SetCellValue(TotalAmount);
                    cell.CellStyle = TableHeaderCS;

                    TotalSquare = decimal.Round(TotalSquare, 3, MidpointRounding.AwayFromZero);
                    cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                    cell.SetCellValue(Convert.ToDouble(TotalSquare));
                    cell.CellStyle = TableHeaderDecCS;

                    RowIndex++;

                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "ColorType")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;

                    AllTotalSquare = decimal.Round(AllTotalSquare, 3, MidpointRounding.AwayFromZero);
                    cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                    cell.SetCellValue(Convert.ToDouble(AllTotalSquare));
                    cell.CellStyle = TableHeaderDecCS;
                }
                RowIndex++;
            }

            CalibriBold11CS = hssfworkbook.CreateCellStyle();
            CalibriBold11CS.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.BottomBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.BorderLeft = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.LeftBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.BorderRight = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.RightBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.BorderTop = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.TopBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.SetFont(CalibriBold11F);

        }

        public void FilenkaSimpleFToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            ref HSSFSheet sheet1, DataTable DT, int WorkAssignmentID, string BatchName, string ClientName, string PageName, decimal time1, decimal cost1, ref int RowIndex)
        {
            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
            cell.CellStyle = Calibri11CS;


            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Плановое время выполнения (фрезер):");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "ч/ч");
            cell.CellStyle = Calibri11CS;
            cell = sheet1.CreateRow(RowIndex++).CreateCell(2);
            cell.SetCellValue(Convert.ToDouble(time1));
            cell.CellStyle = Calibri11CS;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Планово-премиальный фонд (фрезер):");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "руб.");
            cell.CellStyle = Calibri11CS;
            cell = sheet1.CreateRow(RowIndex++).CreateCell(2);
            cell.SetCellValue(Convert.ToDouble(cost1));
            cell.CellStyle = Calibri11CS;

            RowIndex++;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, PageName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, ClientName);
            cell.CellStyle = CalibriBold11CS;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Филенка");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Фрезер");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 5, "м.кв.");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            int CType = -1;
            int AllTotalAmount = 0;
            int TotalAmount = 0;
            decimal AllTotalSquare = 0;
            decimal TotalSquare = 0;

            if (DT.Rows.Count > 0)
                CType = Convert.ToInt32(DT.Rows[0]["ColorType"]);

            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value)
                {
                    AllTotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    TotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                }
                if (DT.Rows[x]["Square"] != DBNull.Value)
                {
                    AllTotalSquare += Convert.ToDecimal(DT.Rows[x]["Square"]);
                    TotalSquare += Convert.ToDecimal(DT.Rows[x]["Square"]);
                }

                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "ColorType")
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
                    if (CType != Convert.ToInt32(DT.Rows[x + 1]["ColorType"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "ColorType")
                                continue;
                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        TotalSquare = decimal.Round(TotalSquare, 3, MidpointRounding.AwayFromZero);
                        cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                        cell.SetCellValue(Convert.ToDouble(TotalSquare));
                        cell.CellStyle = TableHeaderCS;

                        CType = Convert.ToInt32(DT.Rows[x + 1]["ColorType"]);
                        TotalAmount = 0;
                        TotalSquare = 0;
                        RowIndex++;
                    }
                }

                if (x == DT.Rows.Count - 1)
                {
                    RowIndex++;
                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "ColorType")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Итого:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                    cell.SetCellValue(TotalAmount);
                    cell.CellStyle = TableHeaderCS;

                    TotalSquare = decimal.Round(TotalSquare, 3, MidpointRounding.AwayFromZero);
                    cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                    cell.SetCellValue(Convert.ToDouble(TotalSquare));
                    cell.CellStyle = TableHeaderDecCS;

                    RowIndex++;

                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "ColorType")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;

                    AllTotalSquare = decimal.Round(AllTotalSquare, 3, MidpointRounding.AwayFromZero);
                    cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                    cell.SetCellValue(Convert.ToDouble(AllTotalSquare));
                    cell.CellStyle = TableHeaderDecCS;
                }
                RowIndex++;
            }

            CalibriBold11CS = hssfworkbook.CreateCellStyle();
            CalibriBold11CS.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.BottomBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.BorderLeft = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.LeftBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.BorderRight = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.RightBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.BorderTop = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.TopBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.SetFont(CalibriBold11F);

        }

        public void FilenkaSimpleKToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            ref HSSFSheet sheet1, DataTable DT, int WorkAssignmentID, string BatchName, string ClientName, string PageName, decimal time1, decimal cost1, ref int RowIndex)
        {
            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
            cell.CellStyle = Calibri11CS;


            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Плановое время выполнения (клей):");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "ч/ч");
            cell.CellStyle = Calibri11CS;
            cell = sheet1.CreateRow(RowIndex++).CreateCell(2);
            cell.SetCellValue(Convert.ToDouble(time1));
            cell.CellStyle = Calibri11CS;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Планово-премиальный фонд (клей):");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "руб.");
            cell.CellStyle = Calibri11CS;
            cell = sheet1.CreateRow(RowIndex++).CreateCell(2);
            cell.SetCellValue(Convert.ToDouble(cost1));
            cell.CellStyle = Calibri11CS;

            RowIndex++;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, PageName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, ClientName);
            cell.CellStyle = CalibriBold11CS;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Филенка");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Клей");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 5, "м.кв.");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            int CType = -1;
            int AllTotalAmount = 0;
            int TotalAmount = 0;
            decimal AllTotalSquare = 0;
            decimal TotalSquare = 0;

            if (DT.Rows.Count > 0)
                CType = Convert.ToInt32(DT.Rows[0]["ColorType"]);

            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value)
                {
                    AllTotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    TotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                }
                if (DT.Rows[x]["Square"] != DBNull.Value)
                {
                    AllTotalSquare += Convert.ToDecimal(DT.Rows[x]["Square"]);
                    TotalSquare += Convert.ToDecimal(DT.Rows[x]["Square"]);
                }

                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "ColorType")
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
                    if (CType != Convert.ToInt32(DT.Rows[x + 1]["ColorType"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "ColorType")
                                continue;
                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        TotalSquare = decimal.Round(TotalSquare, 3, MidpointRounding.AwayFromZero);
                        cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                        cell.SetCellValue(Convert.ToDouble(TotalSquare));
                        cell.CellStyle = TableHeaderCS;

                        CType = Convert.ToInt32(DT.Rows[x + 1]["ColorType"]);
                        TotalAmount = 0;
                        TotalSquare = 0;
                        RowIndex++;
                    }
                }

                if (x == DT.Rows.Count - 1)
                {
                    RowIndex++;
                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "ColorType")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Итого:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                    cell.SetCellValue(TotalAmount);
                    cell.CellStyle = TableHeaderCS;

                    TotalSquare = decimal.Round(TotalSquare, 3, MidpointRounding.AwayFromZero);
                    cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                    cell.SetCellValue(Convert.ToDouble(TotalSquare));
                    cell.CellStyle = TableHeaderDecCS;

                    RowIndex++;

                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "ColorType")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;

                    AllTotalSquare = decimal.Round(AllTotalSquare, 3, MidpointRounding.AwayFromZero);
                    cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                    cell.SetCellValue(Convert.ToDouble(AllTotalSquare));
                    cell.CellStyle = TableHeaderDecCS;
                }
                RowIndex++;
            }

            CalibriBold11CS = hssfworkbook.CreateCellStyle();
            CalibriBold11CS.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.BottomBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.BorderLeft = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.LeftBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.BorderRight = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.RightBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.BorderTop = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.TopBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.SetFont(CalibriBold11F);

        }

        public void FilenkaSimpleP1ToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            ref HSSFSheet sheet1, DataTable DT, int WorkAssignmentID, string BatchName, string ClientName, string PageName, decimal time1, decimal cost1, decimal time2, decimal cost2, ref int RowIndex)
        {
            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
            cell.CellStyle = Calibri11CS;


            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Плановое время выполнения (пресс):");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "ч/ч");
            cell.CellStyle = Calibri11CS;
            cell = sheet1.CreateRow(RowIndex++).CreateCell(2);
            cell.SetCellValue(Convert.ToDouble(time1));
            cell.CellStyle = Calibri11CS;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Плановое время выполнения (обрезка):");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "ч/ч");
            cell.CellStyle = Calibri11CS;
            cell = sheet1.CreateRow(RowIndex++).CreateCell(2);
            cell.SetCellValue(Convert.ToDouble(time2));
            cell.CellStyle = Calibri11CS;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Планово-премиальный фонд (пресс):");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "руб.");
            cell.CellStyle = Calibri11CS;
            cell = sheet1.CreateRow(RowIndex++).CreateCell(2);
            cell.SetCellValue(Convert.ToDouble(cost1));
            cell.CellStyle = Calibri11CS;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Планово-премиальный фонд (обрезка):");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "руб.");
            cell.CellStyle = Calibri11CS;
            cell = sheet1.CreateRow(RowIndex++).CreateCell(2);
            cell.SetCellValue(Convert.ToDouble(cost2));
            cell.CellStyle = Calibri11CS;

            RowIndex++;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, PageName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, ClientName);
            cell.CellStyle = CalibriBold11CS;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Филенка");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Пресс");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 5, "м.кв.");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            int CType = -1;
            int AllTotalAmount = 0;
            int TotalAmount = 0;
            decimal AllTotalSquare = 0;
            decimal TotalSquare = 0;

            if (DT.Rows.Count > 0)
                CType = Convert.ToInt32(DT.Rows[0]["ColorType"]);

            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value)
                {
                    AllTotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    TotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                }
                if (DT.Rows[x]["Square"] != DBNull.Value)
                {
                    AllTotalSquare += Convert.ToDecimal(DT.Rows[x]["Square"]);
                    TotalSquare += Convert.ToDecimal(DT.Rows[x]["Square"]);
                }

                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "ColorType")
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
                    if (CType != Convert.ToInt32(DT.Rows[x + 1]["ColorType"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "ColorType")
                                continue;
                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        TotalSquare = decimal.Round(TotalSquare, 3, MidpointRounding.AwayFromZero);
                        cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                        cell.SetCellValue(Convert.ToDouble(TotalSquare));
                        cell.CellStyle = TableHeaderCS;

                        CType = Convert.ToInt32(DT.Rows[x + 1]["ColorType"]);
                        TotalAmount = 0;
                        TotalSquare = 0;
                        RowIndex++;
                    }
                }

                if (x == DT.Rows.Count - 1)
                {
                    RowIndex++;
                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "ColorType")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Итого:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                    cell.SetCellValue(TotalAmount);
                    cell.CellStyle = TableHeaderCS;

                    TotalSquare = decimal.Round(TotalSquare, 3, MidpointRounding.AwayFromZero);
                    cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                    cell.SetCellValue(Convert.ToDouble(TotalSquare));
                    cell.CellStyle = TableHeaderDecCS;

                    RowIndex++;

                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "ColorType")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;

                    AllTotalSquare = decimal.Round(AllTotalSquare, 3, MidpointRounding.AwayFromZero);
                    cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                    cell.SetCellValue(Convert.ToDouble(AllTotalSquare));
                    cell.CellStyle = TableHeaderDecCS;
                }
                RowIndex++;
            }

            CalibriBold11CS = hssfworkbook.CreateCellStyle();
            CalibriBold11CS.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.BottomBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.BorderLeft = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.LeftBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.BorderRight = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.RightBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.BorderTop = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.TopBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.SetFont(CalibriBold11F);

        }

        public void FilenkaSimpleP2ToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            ref HSSFSheet sheet1, DataTable DT, int WorkAssignmentID, string BatchName, string ClientName, string PageName, ref int RowIndex)
        {
            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
            cell.CellStyle = Calibri11CS;

            RowIndex++;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, PageName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, ClientName);
            cell.CellStyle = CalibriBold11CS;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Филенка");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Пресс");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 5, "м.кв.");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            int CType = -1;
            int AllTotalAmount = 0;
            int TotalAmount = 0;
            decimal AllTotalSquare = 0;
            decimal TotalSquare = 0;

            if (DT.Rows.Count > 0)
                CType = Convert.ToInt32(DT.Rows[0]["ColorType"]);

            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value)
                {
                    AllTotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    TotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                }
                if (DT.Rows[x]["Square"] != DBNull.Value)
                {
                    AllTotalSquare += Convert.ToDecimal(DT.Rows[x]["Square"]);
                    TotalSquare += Convert.ToDecimal(DT.Rows[x]["Square"]);
                }

                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "ColorType")
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
                    if (CType != Convert.ToInt32(DT.Rows[x + 1]["ColorType"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "ColorType")
                                continue;
                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        TotalSquare = decimal.Round(TotalSquare, 3, MidpointRounding.AwayFromZero);
                        cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                        cell.SetCellValue(Convert.ToDouble(TotalSquare));
                        cell.CellStyle = TableHeaderCS;

                        CType = Convert.ToInt32(DT.Rows[x + 1]["ColorType"]);
                        TotalAmount = 0;
                        TotalSquare = 0;
                        RowIndex++;
                    }
                }

                if (x == DT.Rows.Count - 1)
                {
                    RowIndex++;
                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "ColorType")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Итого:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                    cell.SetCellValue(TotalAmount);
                    cell.CellStyle = TableHeaderCS;

                    TotalSquare = decimal.Round(TotalSquare, 3, MidpointRounding.AwayFromZero);
                    cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                    cell.SetCellValue(Convert.ToDouble(TotalSquare));
                    cell.CellStyle = TableHeaderDecCS;

                    RowIndex++;

                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "ColorType")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;

                    AllTotalSquare = decimal.Round(AllTotalSquare, 3, MidpointRounding.AwayFromZero);
                    cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                    cell.SetCellValue(Convert.ToDouble(AllTotalSquare));
                    cell.CellStyle = TableHeaderDecCS;
                }
                RowIndex++;
            }

            CalibriBold11CS = hssfworkbook.CreateCellStyle();
            CalibriBold11CS.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.BottomBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.BorderLeft = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.LeftBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.BorderRight = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.RightBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.BorderTop = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.TopBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.SetFont(CalibriBold11F);

        }

        public void FilenkaBoxesToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            ref HSSFSheet sheet1, DataTable DT, int WorkAssignmentID, string BatchName, string ClientName, decimal time, decimal cost, ref int RowIndex)
        {

            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
            cell.CellStyle = Calibri11CS;


            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Плановое время выполнения:");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "ч/ч");
            cell.CellStyle = Calibri11CS;

            cell = sheet1.CreateRow(RowIndex++).CreateCell(2);
            cell.SetCellValue(Convert.ToDouble(time));
            cell.CellStyle = Calibri11CS;


            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Планово-премиальный фонд:");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "руб.");
            cell.CellStyle = Calibri11CS;

            cell = sheet1.CreateRow(RowIndex++).CreateCell(2);
            cell.SetCellValue(Convert.ToDouble(cost));
            cell.CellStyle = Calibri11CS;

            RowIndex++;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, "Вставка ВП-204");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, ClientName);
            cell.CellStyle = CalibriBold11CS;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Филенка");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Пила");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 5, "м.кв.");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            int CType = -1;
            int AllTotalAmount = 0;
            int TotalAmount = 0;
            decimal AllTotalSquare = 0;
            decimal TotalSquare = 0;

            if (DT.Rows.Count > 0)
                CType = Convert.ToInt32(DT.Rows[0]["ColorType"]);

            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value)
                {
                    AllTotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    TotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                }
                if (DT.Rows[x]["Square"] != DBNull.Value)
                {
                    AllTotalSquare += Convert.ToDecimal(DT.Rows[x]["Square"]);
                    TotalSquare += Convert.ToDecimal(DT.Rows[x]["Square"]);
                }

                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "ColorType")
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
                    if (CType != Convert.ToInt32(DT.Rows[x + 1]["ColorType"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "ColorType")
                                continue;
                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        TotalSquare = decimal.Round(TotalSquare, 3, MidpointRounding.AwayFromZero);
                        cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                        cell.SetCellValue(Convert.ToDouble(TotalSquare));
                        cell.CellStyle = TableHeaderCS;

                        CType = Convert.ToInt32(DT.Rows[x + 1]["ColorType"]);
                        TotalAmount = 0;
                        TotalSquare = 0;
                        RowIndex++;
                    }
                }

                if (x == DT.Rows.Count - 1)
                {
                    RowIndex++;
                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "ColorType")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Итого:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                    cell.SetCellValue(TotalAmount);
                    cell.CellStyle = TableHeaderCS;

                    TotalSquare = decimal.Round(TotalSquare, 3, MidpointRounding.AwayFromZero);
                    cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                    cell.SetCellValue(Convert.ToDouble(TotalSquare));
                    cell.CellStyle = TableHeaderDecCS;

                    RowIndex++;

                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "ColorType")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(3);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;

                    AllTotalSquare = decimal.Round(AllTotalSquare, 3, MidpointRounding.AwayFromZero);
                    cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                    cell.SetCellValue(Convert.ToDouble(AllTotalSquare));
                    cell.CellStyle = TableHeaderDecCS;
                }
                RowIndex++;
            }

            CalibriBold11CS = hssfworkbook.CreateCellStyle();
            CalibriBold11CS.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.BottomBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.BorderLeft = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.LeftBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.BorderRight = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.RightBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.BorderTop = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.TopBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.SetFont(CalibriBold11F);

            //RowIndex++;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "задание начали:");
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //sheet1.AddMergedRegion(new NPOI.HSSF.Util.Region(RowIndex, 0, RowIndex, 1));
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "задание закончили:");
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //sheet1.AddMergedRegion(new NPOI.HSSF.Util.Region(RowIndex, 2, RowIndex, 4));
            //RowIndex++;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "работало человек:");
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //sheet1.AddMergedRegion(new NPOI.HSSF.Util.Region(RowIndex, 0, RowIndex, 1));
            //RowIndex++;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "№ станка:");
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //sheet1.AddMergedRegion(new NPOI.HSSF.Util.Region(RowIndex, 0, RowIndex, 1));
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "№ операции:");
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //sheet1.AddMergedRegion(new NPOI.HSSF.Util.Region(RowIndex, 2, RowIndex, 4));
        }

        public void OrdersToExcelSingly(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            ref HSSFSheet sheet1, int WorkAssignmentID, string BatchName, string ClientName, ref int RowIndex)
        {
            HSSFCell cell = null;

            HSSFFont Serif8F = hssfworkbook.CreateFont();
            Serif8F.FontHeightInPoints = 8;
            Serif8F.FontName = "Calibri";

            HSSFFont Serif10F = hssfworkbook.CreateFont();
            Serif10F.FontHeightInPoints = 10;
            Serif10F.FontName = "Calibri";

            HSSFCellStyle TableHeaderCS7 = hssfworkbook.CreateCellStyle();
            TableHeaderCS7.Alignment = HSSFCellStyle.ALIGN_LEFT;
            TableHeaderCS7.BorderBottom = HSSFCellStyle.BORDER_THIN;
            TableHeaderCS7.BottomBorderColor = HSSFColor.BLACK.index;
            TableHeaderCS7.BorderLeft = HSSFCellStyle.BORDER_THIN;
            TableHeaderCS7.LeftBorderColor = HSSFColor.BLACK.index;
            TableHeaderCS7.BorderRight = HSSFCellStyle.BORDER_THIN;
            TableHeaderCS7.RightBorderColor = HSSFColor.BLACK.index;
            TableHeaderCS7.BorderTop = HSSFCellStyle.BORDER_THIN;
            TableHeaderCS7.TopBorderColor = HSSFColor.BLACK.index;
            TableHeaderCS7.SetFont(Serif8F);

            HSSFCellStyle TableHeaderCS9 = hssfworkbook.CreateCellStyle();
            TableHeaderCS9.Alignment = HSSFCellStyle.ALIGN_LEFT;
            TableHeaderCS9.BorderBottom = HSSFCellStyle.BORDER_THIN;
            TableHeaderCS9.BottomBorderColor = HSSFColor.BLACK.index;
            TableHeaderCS9.BorderLeft = HSSFCellStyle.BORDER_THIN;
            TableHeaderCS9.LeftBorderColor = HSSFColor.BLACK.index;
            TableHeaderCS9.BorderRight = HSSFCellStyle.BORDER_THIN;
            TableHeaderCS9.RightBorderColor = HSSFColor.BLACK.index;
            TableHeaderCS9.BorderTop = HSSFCellStyle.BORDER_THIN;
            TableHeaderCS9.TopBorderColor = HSSFColor.BLACK.index;
            TableHeaderCS9.SetFont(Serif10F);

            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex + 2), 1, "УТВЕРЖДАЮ_____________");
            //cell.CellStyle = Calibri11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
            //cell.CellStyle = Calibri11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex + 4), 0, "Клиент:");
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex + 4), 1, ClientName);
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex + 5), 0, "Партия:");
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex + 5), 1, BatchName);
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex + 3), 0, "Задание №" + WorkAssignmentID.ToString());
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex + 3), 1, "Заказы");
            //cell.CellStyle = CalibriBold11CS;
            //RowIndex += 6;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Плановое время выполнения:");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, "Заказы");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, ClientName);
            cell.CellStyle = CalibriBold11CS;

            int ColumnIndex = -1;
            string ColumnName = string.Empty;

            for (int x = 0; x < SummOrdersDT.Columns.Count; x++)
            {
                if (SummOrdersDT.Columns[x].ColumnName == "Height" || SummOrdersDT.Columns[x].ColumnName == "Width")
                    continue;
                ColumnIndex++;
                ColumnName = SummOrdersDT.Columns[x].ColumnName;
                if (ColumnName == "Sizes")
                {
                    ColumnName = "Размер";
                    cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), ColumnIndex, ColumnName);
                    cell.CellStyle = TableHeaderCS7;
                    sheet1.SetColumnWidth(ColumnIndex, 12 * 256);
                    continue;
                }
                if (ColumnName == "TotalAmount")
                {
                    ColumnName = "Итого";
                    cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), ColumnIndex, ColumnName);
                    cell.CellStyle = TableHeaderCS7;
                    sheet1.SetColumnWidth(ColumnIndex, 8 * 256);
                    continue;
                }
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), ColumnIndex, ColumnName);
                cell.CellStyle = TableHeaderCS7;
                sheet1.SetColumnWidth(ColumnIndex, 19 * 256);
            }
            sheet1.SetColumnWidth(0, 15 * 256);
            RowIndex++;

            ColumnIndex = -1;
            for (int y = 0; y < SummOrdersDT.Columns.Count; y++)
            {
                if (SummOrdersDT.Columns[y].ColumnName == "Height" || SummOrdersDT.Columns[y].ColumnName == "Width")
                    continue;
                Type t = SummOrdersDT.Rows[0][y].GetType();

                ColumnIndex++;

                if (int.TryParse(SummOrdersDT.Rows[0][y].ToString(), out int IntValue))
                {
                    cell = sheet1.CreateRow(RowIndex).CreateCell(ColumnIndex);
                    cell.SetCellValue(IntValue);
                    cell.CellStyle = TableHeaderCS7;
                    continue;
                }

                if (t.Name == "String" || t.Name == "DBNull")
                {
                    cell = sheet1.CreateRow(RowIndex).CreateCell(ColumnIndex);
                    cell.SetCellValue(SummOrdersDT.Rows[0][y].ToString());
                    cell.CellStyle = TableHeaderCS7;
                    continue;
                }
            }
            RowIndex++;

            for (int x = 1; x < SummOrdersDT.Rows.Count; x++)
            {
                ColumnIndex = -1;
                for (int y = 0; y < SummOrdersDT.Columns.Count; y++)
                {
                    if (SummOrdersDT.Columns[y].ColumnName == "Height" || SummOrdersDT.Columns[y].ColumnName == "Width")
                        continue;
                    Type t = SummOrdersDT.Rows[x][y].GetType();

                    ColumnIndex++;

                    if (x == SummOrdersDT.Rows.Count - 1 && int.TryParse(SummOrdersDT.Rows[x][y].ToString(), out int IntValue))
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(ColumnIndex);
                        cell.SetCellValue(IntValue);
                        cell.CellStyle = TableHeaderCS9;
                        continue;
                    }

                    if (x == SummOrdersDT.Rows.Count - 2 && double.TryParse(SummOrdersDT.Rows[x][y].ToString(), out double DecValue))
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(ColumnIndex);
                        cell.SetCellValue(DecValue);
                        cell.CellStyle = TableHeaderCS9;
                        continue;
                    }

                    if (int.TryParse(SummOrdersDT.Rows[x][y].ToString(), out IntValue))
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(ColumnIndex);
                        cell.SetCellValue(IntValue);
                        cell.CellStyle = TableHeaderCS9;
                        continue;
                    }

                    if (t.Name == "String" || t.Name == "DBNull")
                    {
                        cell = sheet1.CreateRow(RowIndex).CreateCell(ColumnIndex);
                        cell.SetCellValue(SummOrdersDT.Rows[x][y].ToString());
                        cell.CellStyle = TableHeaderCS9;
                        continue;
                    }
                }
                RowIndex++;
            }

            CalibriBold11CS = hssfworkbook.CreateCellStyle();
            CalibriBold11CS.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.BottomBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.BorderLeft = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.LeftBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.BorderRight = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.RightBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.BorderTop = HSSFCellStyle.BORDER_MEDIUM;
            CalibriBold11CS.TopBorderColor = HSSFColor.BLACK.index;
            CalibriBold11CS.SetFont(CalibriBold11F);

            //RowIndex++;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "задание начали:");
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //sheet1.AddMergedRegion(new NPOI.HSSF.Util.Region(RowIndex, 0, RowIndex, 1));
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "задание закончили:");
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //sheet1.AddMergedRegion(new NPOI.HSSF.Util.Region(RowIndex, 2, RowIndex, 4));
            //RowIndex++;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "работало человек:");
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //sheet1.AddMergedRegion(new NPOI.HSSF.Util.Region(RowIndex, 0, RowIndex, 1));
            //RowIndex++;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "№ станка:");
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //sheet1.AddMergedRegion(new NPOI.HSSF.Util.Region(RowIndex, 0, RowIndex, 1));
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "№ операции:");
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, string.Empty);
            //cell.CellStyle = CalibriBold11CS;
            //sheet1.AddMergedRegion(new NPOI.HSSF.Util.Region(RowIndex, 2, RowIndex, 4));
        }

        public void DyeingMen1ToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            ref HSSFSheet sheet1, DataTable DT, int WorkAssignmentID, string BatchName, string ClientName, string OrderName, string PageName, string Notes,
            ref int RowIndex)
        {
            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, PageName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, ClientName);
            cell.CellStyle = CalibriBold11CS;
            if (OrderName.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Заказ:");
                cell.CellStyle = CalibriBold11CS;
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, OrderName);
                cell.CellStyle = CalibriBold11CS;
            }
            if (Notes.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Примечание: " + Notes);
                cell.CellStyle = CalibriBold11CS;
            }

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Профиль");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Цвет профиля");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Цвет наполнителя");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 5, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 6, "Гр.в.");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 7, "Гр.н.");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 8, "Патина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 9, "м.кв.");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            int CType = -1;
            int AllTotalAmount = 0;
            int TotalAmount = 0;
            decimal AllTotalSquare = 0;
            decimal TotalSquare = 0;
            int DifferentDecorCount = 0;

            using (DataView DV = new DataView(DT))
            {
                DifferentDecorCount = DV.ToTable(true, new string[] { "ColorType" }).Rows.Count;
            }


            if (DT.Rows.Count > 0)
                CType = Convert.ToInt32(DT.Rows[0]["ColorType"]);

            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value)
                {
                    AllTotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    TotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                }
                if (DT.Rows[x]["Square"] != DBNull.Value)
                {
                    AllTotalSquare += Convert.ToDecimal(DT.Rows[x]["Square"]);
                    TotalSquare += Convert.ToDecimal(DT.Rows[x]["Square"]);
                }

                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "ColorType" || DT.Columns[y].ColumnName == "Notes")
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
                    if (CType != Convert.ToInt32(DT.Rows[x + 1]["ColorType"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "ColorType" || DT.Columns[y].ColumnName == "Notes")
                                continue;

                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(9);
                        cell.SetCellValue(Convert.ToDouble(TotalSquare));
                        cell.CellStyle = TableHeaderCS;

                        CType = Convert.ToInt32(DT.Rows[x + 1]["ColorType"]);
                        TotalAmount = 0;
                        TotalSquare = 0;
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
                            if (DT.Columns[y].ColumnName == "ColorType" || DT.Columns[y].ColumnName == "Notes")
                                continue;

                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(7);
                        cell.SetCellValue(Convert.ToDouble(TotalSquare));
                        cell.CellStyle = TableHeaderDecCS;
                    }
                    RowIndex++;

                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "ColorType" || DT.Columns[y].ColumnName == "Notes")
                            continue;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(9);
                    cell.SetCellValue(Convert.ToDouble(AllTotalSquare));
                    cell.CellStyle = TableHeaderDecCS;
                }
                RowIndex++;
            }
        }

        public void DyeingMen2ToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            ref HSSFSheet sheet1, DataTable DT, int WorkAssignmentID, string BatchName, string ClientName, string OrderName, string PageName, string Notes,
            ref int RowIndex)
        {
            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, PageName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, ClientName);
            cell.CellStyle = CalibriBold11CS;
            if (OrderName.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Заказ:");
                cell.CellStyle = CalibriBold11CS;
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, OrderName);
                cell.CellStyle = CalibriBold11CS;
            }
            if (Notes.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Примечание: " + Notes);
                cell.CellStyle = CalibriBold11CS;
            }

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Профиль");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Цвет профиля");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Цвет наполнителя");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 5, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 6, "Лак");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 7, "м.кв.");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            int CType = -1;
            int AllTotalAmount = 0;
            int TotalAmount = 0;
            decimal AllTotalSquare = 0;
            decimal TotalSquare = 0;
            int DifferentDecorCount = 0;

            using (DataView DV = new DataView(DT))
            {
                DifferentDecorCount = DV.ToTable(true, new string[] { "ColorType" }).Rows.Count;
            }


            if (DT.Rows.Count > 0)
                CType = Convert.ToInt32(DT.Rows[0]["ColorType"]);

            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value)
                {
                    AllTotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    TotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                }
                if (DT.Rows[x]["Square"] != DBNull.Value)
                {
                    AllTotalSquare += Convert.ToDecimal(DT.Rows[x]["Square"]);
                    TotalSquare += Convert.ToDecimal(DT.Rows[x]["Square"]);
                }

                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "ColorType" || DT.Columns[y].ColumnName == "Notes")
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
                    if (CType != Convert.ToInt32(DT.Rows[x + 1]["ColorType"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "ColorType" || DT.Columns[y].ColumnName == "Notes")
                                continue;

                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(7);
                        cell.SetCellValue(Convert.ToDouble(TotalSquare));
                        cell.CellStyle = TableHeaderCS;

                        CType = Convert.ToInt32(DT.Rows[x + 1]["ColorType"]);
                        TotalAmount = 0;
                        TotalSquare = 0;
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
                            if (DT.Columns[y].ColumnName == "ColorType" || DT.Columns[y].ColumnName == "Notes")
                                continue;

                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(7);
                        cell.SetCellValue(Convert.ToDouble(TotalSquare));
                        cell.CellStyle = TableHeaderDecCS;
                    }
                    RowIndex++;

                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "ColorType" || DT.Columns[y].ColumnName == "Notes")
                            continue;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(7);
                    cell.SetCellValue(Convert.ToDouble(AllTotalSquare));
                    cell.CellStyle = TableHeaderDecCS;
                }
                RowIndex++;
            }
        }

        public void DyeingWomen1ToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            ref HSSFSheet sheet1, DataTable DT, int WorkAssignmentID, string BatchName, string ClientName, string OrderName, string PageName, string Notes,
            decimal Time, decimal Cost, ref int RowIndex)
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
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, PageName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, ClientName);
            cell.CellStyle = CalibriBold11CS;
            if (OrderName.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Заказ:");
                cell.CellStyle = CalibriBold11CS;
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, OrderName);
                cell.CellStyle = CalibriBold11CS;
            }
            if (Notes.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Примечание: " + Notes);
                cell.CellStyle = CalibriBold11CS;
            }

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Профиль");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Цвет профиля");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Цвет наполнителя");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 5, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 6, "Зачистка");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 7, "м.кв.");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            int CType = -1;
            int AllTotalAmount = 0;
            int TotalAmount = 0;
            decimal AllTotalSquare = 0;
            decimal TotalSquare = 0;
            int DifferentDecorCount = 0;

            using (DataView DV = new DataView(DT))
            {
                DifferentDecorCount = DV.ToTable(true, new string[] { "ColorType" }).Rows.Count;
            }


            if (DT.Rows.Count > 0)
                CType = Convert.ToInt32(DT.Rows[0]["ColorType"]);

            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value)
                {
                    AllTotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    TotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                }
                if (DT.Rows[x]["Square"] != DBNull.Value)
                {
                    AllTotalSquare += Convert.ToDecimal(DT.Rows[x]["Square"]);
                    TotalSquare += Convert.ToDecimal(DT.Rows[x]["Square"]);
                }

                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "ColorType" || DT.Columns[y].ColumnName == "Notes")
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
                    if (CType != Convert.ToInt32(DT.Rows[x + 1]["ColorType"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "ColorType" || DT.Columns[y].ColumnName == "Notes")
                                continue;

                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(7);
                        cell.SetCellValue(Convert.ToDouble(TotalSquare));
                        cell.CellStyle = TableHeaderCS;

                        CType = Convert.ToInt32(DT.Rows[x + 1]["ColorType"]);
                        TotalAmount = 0;
                        TotalSquare = 0;
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
                            if (DT.Columns[y].ColumnName == "ColorType" || DT.Columns[y].ColumnName == "Notes")
                                continue;

                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(7);
                        cell.SetCellValue(Convert.ToDouble(TotalSquare));
                        cell.CellStyle = TableHeaderDecCS;
                    }
                    RowIndex++;

                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "ColorType" || DT.Columns[y].ColumnName == "Notes")
                            continue;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(7);
                    cell.SetCellValue(Convert.ToDouble(AllTotalSquare));
                    cell.CellStyle = TableHeaderDecCS;
                }
                RowIndex++;
            }
        }

        public void DyeingWomen2ToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            ref HSSFSheet sheet1, DataTable DT, int WorkAssignmentID, string BatchName, string ClientName, string OrderName, string PageName, string Notes,
            ref int RowIndex)
        {
            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, PageName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, ClientName);
            cell.CellStyle = CalibriBold11CS;
            if (OrderName.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Заказ:");
                cell.CellStyle = CalibriBold11CS;
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, OrderName);
                cell.CellStyle = CalibriBold11CS;
            }
            if (Notes.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Примечание: " + Notes);
                cell.CellStyle = CalibriBold11CS;
            }

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Профиль");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Цвет профиля");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Цвет наполнителя");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 5, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 6, "Обезжиривание");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 7, "м.кв.");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            int CType = -1;
            int AllTotalAmount = 0;
            int TotalAmount = 0;
            decimal AllTotalSquare = 0;
            decimal TotalSquare = 0;
            int DifferentDecorCount = 0;

            using (DataView DV = new DataView(DT))
            {
                DifferentDecorCount = DV.ToTable(true, new string[] { "ColorType" }).Rows.Count;
            }


            if (DT.Rows.Count > 0)
                CType = Convert.ToInt32(DT.Rows[0]["ColorType"]);

            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value)
                {
                    AllTotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    TotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                }
                if (DT.Rows[x]["Square"] != DBNull.Value)
                {
                    AllTotalSquare += Convert.ToDecimal(DT.Rows[x]["Square"]);
                    TotalSquare += Convert.ToDecimal(DT.Rows[x]["Square"]);
                }

                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "ColorType" || DT.Columns[y].ColumnName == "Notes")
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
                    if (CType != Convert.ToInt32(DT.Rows[x + 1]["ColorType"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "ColorType" || DT.Columns[y].ColumnName == "Notes")
                                continue;

                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(7);
                        cell.SetCellValue(Convert.ToDouble(TotalSquare));
                        cell.CellStyle = TableHeaderCS;

                        CType = Convert.ToInt32(DT.Rows[x + 1]["ColorType"]);
                        TotalAmount = 0;
                        TotalSquare = 0;
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
                            if (DT.Columns[y].ColumnName == "ColorType" || DT.Columns[y].ColumnName == "Notes")
                                continue;

                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(7);
                        cell.SetCellValue(Convert.ToDouble(TotalSquare));
                        cell.CellStyle = TableHeaderDecCS;
                    }
                    RowIndex++;

                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "ColorType" || DT.Columns[y].ColumnName == "Notes")
                            continue;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(7);
                    cell.SetCellValue(Convert.ToDouble(AllTotalSquare));
                    cell.CellStyle = TableHeaderDecCS;
                }
                RowIndex++;
            }
        }

        public void DyeingWomen3ToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            ref HSSFSheet sheet1, DataTable DT, int WorkAssignmentID, string BatchName, string ClientName, string OrderName, string PageName, string Notes,
            ref int RowIndex)
        {
            HSSFCell cell = null;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, PageName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, ClientName);
            cell.CellStyle = CalibriBold11CS;
            if (OrderName.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Заказ:");
                cell.CellStyle = CalibriBold11CS;
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, OrderName);
                cell.CellStyle = CalibriBold11CS;
            }
            if (Notes.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Примечание: " + Notes);
                cell.CellStyle = CalibriBold11CS;
            }

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Профиль");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Цвет профиля");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Цвет наполнителя");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 5, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 6, "Протирка патины");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 7, "м.кв.");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            int CType = -1;
            int AllTotalAmount = 0;
            int TotalAmount = 0;
            decimal AllTotalSquare = 0;
            decimal TotalSquare = 0;
            int DifferentDecorCount = 0;

            using (DataView DV = new DataView(DT))
            {
                DifferentDecorCount = DV.ToTable(true, new string[] { "ColorType" }).Rows.Count;
            }


            if (DT.Rows.Count > 0)
                CType = Convert.ToInt32(DT.Rows[0]["ColorType"]);

            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value)
                {
                    AllTotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    TotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                }
                if (DT.Rows[x]["Square"] != DBNull.Value)
                {
                    AllTotalSquare += Convert.ToDecimal(DT.Rows[x]["Square"]);
                    TotalSquare += Convert.ToDecimal(DT.Rows[x]["Square"]);
                }

                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "ColorType" || DT.Columns[y].ColumnName == "Notes")
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
                    if (CType != Convert.ToInt32(DT.Rows[x + 1]["ColorType"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "ColorType" || DT.Columns[y].ColumnName == "Notes")
                                continue;

                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(7);
                        cell.SetCellValue(Convert.ToDouble(TotalSquare));
                        cell.CellStyle = TableHeaderCS;

                        CType = Convert.ToInt32(DT.Rows[x + 1]["ColorType"]);
                        TotalAmount = 0;
                        TotalSquare = 0;
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
                            if (DT.Columns[y].ColumnName == "ColorType" || DT.Columns[y].ColumnName == "Notes")
                                continue;

                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(7);
                        cell.SetCellValue(Convert.ToDouble(TotalSquare));
                        cell.CellStyle = TableHeaderDecCS;
                    }
                    RowIndex++;

                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "ColorType" || DT.Columns[y].ColumnName == "Notes")
                            continue;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(7);
                    cell.SetCellValue(Convert.ToDouble(AllTotalSquare));
                    cell.CellStyle = TableHeaderDecCS;
                }
                RowIndex++;
            }
        }

        public void DyeingPackingToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            ref HSSFSheet sheet1, DataTable DT, int WorkAssignmentID, string BatchName, string ClientName, string OrderName, string PageName, string Notes,
            decimal Time, decimal Cost, ref int RowIndex)
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
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, PageName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, ClientName);
            cell.CellStyle = CalibriBold11CS;
            if (OrderName.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Заказ:");
                cell.CellStyle = CalibriBold11CS;
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, OrderName);
                cell.CellStyle = CalibriBold11CS;
            }
            if (Notes.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Примечание: " + Notes);
                cell.CellStyle = CalibriBold11CS;
            }

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Профиль");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Цвет профиля");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Цвет наполнителя");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 5, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 6, "Пленка");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 7, "Упаковка");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 8, "м.кв.");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            int CType = -1;
            int AllTotalAmount = 0;
            int TotalAmount = 0;
            decimal AllTotalSquare = 0;
            decimal TotalSquare = 0;
            int DifferentDecorCount = 0;

            using (DataView DV = new DataView(DT))
            {
                DifferentDecorCount = DV.ToTable(true, new string[] { "ColorType" }).Rows.Count;
            }


            if (DT.Rows.Count > 0)
                CType = Convert.ToInt32(DT.Rows[0]["ColorType"]);

            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value)
                {
                    AllTotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    TotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                }
                if (DT.Rows[x]["Square"] != DBNull.Value)
                {
                    AllTotalSquare += Convert.ToDecimal(DT.Rows[x]["Square"]);
                    TotalSquare += Convert.ToDecimal(DT.Rows[x]["Square"]);
                }

                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "ColorType" || DT.Columns[y].ColumnName == "Notes")
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
                    if (CType != Convert.ToInt32(DT.Rows[x + 1]["ColorType"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "ColorType" || DT.Columns[y].ColumnName == "Notes")
                                continue;

                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(8);
                        cell.SetCellValue(Convert.ToDouble(TotalSquare));
                        cell.CellStyle = TableHeaderCS;

                        CType = Convert.ToInt32(DT.Rows[x + 1]["ColorType"]);
                        TotalAmount = 0;
                        TotalSquare = 0;
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
                            if (DT.Columns[y].ColumnName == "ColorType" || DT.Columns[y].ColumnName == "Notes")
                                continue;

                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(8);
                        cell.SetCellValue(Convert.ToDouble(TotalSquare));
                        cell.CellStyle = TableHeaderDecCS;
                    }
                    RowIndex++;

                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "ColorType" || DT.Columns[y].ColumnName == "Notes")
                            continue;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(8);
                    cell.SetCellValue(Convert.ToDouble(AllTotalSquare));
                    cell.CellStyle = TableHeaderDecCS;
                }
                RowIndex++;
            }
        }

        public void DyeingBoringToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            ref HSSFSheet sheet1, DataTable DT, int WorkAssignmentID, string BatchName, string ClientName, string OrderName, string PageName, string Notes,
            decimal Time, decimal Cost, ref int RowIndex)
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
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, PageName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, ClientName);
            cell.CellStyle = CalibriBold11CS;
            if (OrderName.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Заказ:");
                cell.CellStyle = CalibriBold11CS;
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, OrderName);
                cell.CellStyle = CalibriBold11CS;
            }
            if (Notes.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Примечание: " + Notes);
                cell.CellStyle = CalibriBold11CS;
            }

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Профиль");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Цвет профиля");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Цвет наполнителя");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Ширина");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 5, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 6, "Сверление");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 7, "м.кв.");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 8, "Прим.");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            int CType = -1;
            int AllTotalAmount = 0;
            int TotalAmount = 0;
            decimal AllTotalSquare = 0;
            decimal TotalSquare = 0;
            int DifferentDecorCount = 0;

            using (DataView DV = new DataView(DT))
            {
                DifferentDecorCount = DV.ToTable(true, new string[] { "ColorType" }).Rows.Count;
            }


            if (DT.Rows.Count > 0)
                CType = Convert.ToInt32(DT.Rows[0]["ColorType"]);

            for (int x = 0; x < DT.Rows.Count; x++)
            {
                if (DT.Rows[x]["Count"] != DBNull.Value)
                {
                    AllTotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                    TotalAmount += Convert.ToInt32(DT.Rows[x]["Count"]);
                }
                if (DT.Rows[x]["Square"] != DBNull.Value)
                {
                    AllTotalSquare += Convert.ToDecimal(DT.Rows[x]["Square"]);
                    TotalSquare += Convert.ToDecimal(DT.Rows[x]["Square"]);
                }

                for (int y = 0; y < DT.Columns.Count; y++)
                {
                    if (DT.Columns[y].ColumnName == "ColorType")
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
                    if (CType != Convert.ToInt32(DT.Rows[x + 1]["ColorType"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "ColorType")
                                continue;
                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(7);
                        cell.SetCellValue(Convert.ToDouble(TotalSquare));
                        cell.CellStyle = TableHeaderCS;

                        CType = Convert.ToInt32(DT.Rows[x + 1]["ColorType"]);
                        TotalAmount = 0;
                        TotalSquare = 0;
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
                            if (DT.Columns[y].ColumnName == "ColorType")
                                continue;
                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(7);
                        cell.SetCellValue(Convert.ToDouble(TotalSquare));
                        cell.CellStyle = TableHeaderDecCS;
                    }
                    RowIndex++;

                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "ColorType")
                            continue;
                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(7);
                    cell.SetCellValue(Convert.ToDouble(AllTotalSquare));
                    cell.CellStyle = TableHeaderDecCS;
                }
                RowIndex++;
            }
        }

        public void CurvedAssemblyToExcel(ref HSSFWorkbook hssfworkbook,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            int WorkAssignmentID, string BatchName)
        {
            if (LorenzoCurvedOrdersDT.Rows.Count == 0 && ElegantCurvedOrdersDT.Rows.Count == 0 && Patricia1CurvedOrdersDT.Rows.Count == 0 && KansasCurvedOrdersDT.Rows.Count == 0 && SofiaCurvedOrdersDT.Rows.Count == 0 && DakotaCurvedOrdersDT.Rows.Count == 0 && Turin1CurvedOrdersDT.Rows.Count == 0 && Turin1_1CurvedOrdersDT.Rows.Count == 0
                && Turin3CurvedOrdersDT.Rows.Count == 0 && InfinitiCurvedOrdersDT.Rows.Count == 0)
                return;

            DataTable DistMainOrdersDT = DistMainOrdersTable(LorenzoCurvedOrdersDT, ElegantCurvedOrdersDT, Patricia1CurvedOrdersDT, KansasCurvedOrdersDT, SofiaCurvedOrdersDT, Turin1CurvedOrdersDT, Turin1_1CurvedOrdersDT, Turin3CurvedOrdersDT, Turin3CurvedOrdersDT.Clone(), InfinitiCurvedOrdersDT, DakotaCurvedOrdersDT, true);
            DataTable DT = KansasCurvedOrdersDT.Clone();
            DataTable ZOVOrdersNames = new DataTable();
            DataTable MarketOrdersNames = new DataTable();
            string MainOrdersID = string.Empty;
            string ClientName1 = string.Empty;
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
                HSSFSheet sheet1 = hssfworkbook.CreateSheet("Гнутые ЗОВ");
                sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

                sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
                sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
                sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
                sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

                sheet1.SetColumnWidth(0, 17 * 256);
                sheet1.SetColumnWidth(1, 15 * 256);
                sheet1.SetColumnWidth(2, 10 * 256);
                sheet1.SetColumnWidth(3, 15 * 256);
                sheet1.SetColumnWidth(4, 6 * 256);
                sheet1.SetColumnWidth(5, 6 * 256);
                sheet1.SetColumnWidth(6, 6 * 256);

                CurvedAssemblyDT.Clear();
                DT.Clear();

                DataTable DT1 = new DataTable();
                foreach (DataRow item in LorenzoCurvedOrdersDT.Select("GroupType=0"))
                    DT.Rows.Add(item.ItemArray);
                foreach (DataRow item in ElegantCurvedOrdersDT.Select("GroupType=0"))
                    DT.Rows.Add(item.ItemArray);
                foreach (DataRow item in Patricia1CurvedOrdersDT.Select("GroupType=0"))
                    DT.Rows.Add(item.ItemArray);
                foreach (DataRow item in KansasCurvedOrdersDT.Select("GroupType=0"))
                    DT.Rows.Add(item.ItemArray);
                foreach (DataRow item in SofiaCurvedOrdersDT.Select("GroupType=0"))
                    DT.Rows.Add(item.ItemArray);
                foreach (DataRow item in DakotaCurvedOrdersDT.Select("GroupType=0"))
                    DT.Rows.Add(item.ItemArray);
                foreach (DataRow item in Turin1CurvedOrdersDT.Select("GroupType=0"))
                    DT.Rows.Add(item.ItemArray);
                foreach (DataRow item in Turin1_1CurvedOrdersDT.Select("GroupType=0"))
                    DT.Rows.Add(item.ItemArray);
                foreach (DataRow item in Turin3CurvedOrdersDT.Select("GroupType=0"))
                    DT.Rows.Add(item.ItemArray);
                foreach (DataRow item in InfinitiCurvedOrdersDT.Select("GroupType=0"))
                    DT.Rows.Add(item.ItemArray);

                CurvedAssemblyCollect(DT, ref CurvedAssemblyDT);

                DT1 = CurvedAssemblyDT.Copy();

                decimal div1 = 370;
                decimal div2 = 2.14m;
                decimal time = 0;
                decimal cost = 0;
                PlanningCurved(DT1, ref time, ref cost);
                CurvedAssembly1ToExcelSingly(ref hssfworkbook, ref sheet1, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS,
                    DT1, WorkAssignmentID, BatchName, "ЗОВ", string.Empty, "Гнутые фасады", time, cost, ref RowIndex);
                RowIndex++;

                CurvedAssembly1ToExcelSingly(ref hssfworkbook, ref sheet1, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS,
                    DT1, WorkAssignmentID, BatchName, "ЗОВ", "ДУБЛЬ", "Гнутые фасады", time, cost, ref RowIndex);
                RowIndex++;

                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=0"))
                {
                    CurvedAssemblyDT.Clear();
                    DT.Clear();
                    int MainOrderID = Convert.ToInt32(item["MainOrderID"]);

                    foreach (DataRow item1 in LorenzoCurvedOrdersDT.Select("MainOrderID=" + MainOrderID))
                        DT.Rows.Add(item1.ItemArray);
                    foreach (DataRow item1 in ElegantCurvedOrdersDT.Select("MainOrderID=" + MainOrderID))
                        DT.Rows.Add(item1.ItemArray);
                    foreach (DataRow item1 in Patricia1CurvedOrdersDT.Select("MainOrderID=" + MainOrderID))
                        DT.Rows.Add(item1.ItemArray);
                    foreach (DataRow item1 in KansasCurvedOrdersDT.Select("MainOrderID=" + MainOrderID))
                        DT.Rows.Add(item1.ItemArray);
                    foreach (DataRow item1 in SofiaCurvedOrdersDT.Select("MainOrderID=" + MainOrderID))
                        DT.Rows.Add(item1.ItemArray);
                    foreach (DataRow item1 in DakotaCurvedOrdersDT.Select("MainOrderID=" + MainOrderID))
                        DT.Rows.Add(item1.ItemArray);
                    foreach (DataRow item1 in Turin1CurvedOrdersDT.Select("MainOrderID=" + MainOrderID))
                        DT.Rows.Add(item1.ItemArray);
                    foreach (DataRow item1 in Turin1_1CurvedOrdersDT.Select("MainOrderID=" + MainOrderID))
                        DT.Rows.Add(item1.ItemArray);
                    foreach (DataRow item1 in Turin3CurvedOrdersDT.Select("MainOrderID=" + MainOrderID))
                        DT.Rows.Add(item1.ItemArray);
                    foreach (DataRow item1 in InfinitiCurvedOrdersDT.Select("MainOrderID=" + MainOrderID))
                        DT.Rows.Add(item1.ItemArray);

                    CurvedAssemblyCollect(DT, ref CurvedAssemblyDT);

                    DataRow[] CRows = ZOVOrdersNames.Select("MainOrderID=" + MainOrderID);
                    if (CRows.Count() > 0)
                    {
                        ClientName1 = CRows[0]["ClientName"].ToString();
                        OrderName = CRows[0]["DocNumber"].ToString();
                    }

                    if (CurvedAssemblyDT.Rows.Count > 0)
                    {
                        div1 = 10;
                        div2 = 2.14m;
                        time = 0;
                        cost = 0;
                        PlanningTimebyCount(CurvedAssemblyDT, div1, div2, ref time, ref cost);
                        CurvedAssembly2ToExcelSingly(ref hssfworkbook, ref sheet1,
                            Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, CurvedAssemblyDT,
                            WorkAssignmentID, BatchName, ClientName1, OrderName, "Гнутые фасады", Notes, time, cost, ref RowIndex);
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
                HSSFSheet sheet1 = hssfworkbook.CreateSheet("Гнутые Маркетинг");
                sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

                sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
                sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
                sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
                sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

                sheet1.SetColumnWidth(0, 17 * 256);
                sheet1.SetColumnWidth(1, 15 * 256);
                sheet1.SetColumnWidth(2, 10 * 256);
                sheet1.SetColumnWidth(3, 15 * 256);
                sheet1.SetColumnWidth(4, 6 * 256);
                sheet1.SetColumnWidth(5, 6 * 256);
                sheet1.SetColumnWidth(6, 6 * 256);

                CurvedAssemblyDT.Clear();
                DT.Clear();

                DataTable DT1 = new DataTable();
                foreach (DataRow item in LorenzoCurvedOrdersDT.Select("GroupType=1"))
                    DT.Rows.Add(item.ItemArray);
                foreach (DataRow item in ElegantCurvedOrdersDT.Select("GroupType=1"))
                    DT.Rows.Add(item.ItemArray);
                foreach (DataRow item in Patricia1CurvedOrdersDT.Select("GroupType=1"))
                    DT.Rows.Add(item.ItemArray);
                foreach (DataRow item in KansasCurvedOrdersDT.Select("GroupType=1"))
                    DT.Rows.Add(item.ItemArray);
                foreach (DataRow item in SofiaCurvedOrdersDT.Select("GroupType=1"))
                    DT.Rows.Add(item.ItemArray);
                foreach (DataRow item in DakotaCurvedOrdersDT.Select("GroupType=1"))
                    DT.Rows.Add(item.ItemArray);
                foreach (DataRow item in Turin1CurvedOrdersDT.Select("GroupType=1"))
                    DT.Rows.Add(item.ItemArray);
                foreach (DataRow item in Turin1_1CurvedOrdersDT.Select("GroupType=1"))
                    DT.Rows.Add(item.ItemArray);
                foreach (DataRow item in Turin3CurvedOrdersDT.Select("GroupType=1"))
                    DT.Rows.Add(item.ItemArray);
                foreach (DataRow item in InfinitiCurvedOrdersDT.Select("GroupType=1"))
                    DT.Rows.Add(item.ItemArray);

                CurvedAssemblyCollect(DT, ref CurvedAssemblyDT);

                DT1 = CurvedAssemblyDT.Copy();

                decimal div1 = 370;
                decimal div2 = 2.14m;
                decimal time = 0;
                decimal cost = 0;
                PlanningCurved(DT1, ref time, ref cost);
                CurvedAssembly1ToExcelSingly(ref hssfworkbook, ref sheet1, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS,
                    DT1, WorkAssignmentID, BatchName, "Маркетинг", string.Empty, "Гнутые фасады", time, cost, ref RowIndex);
                RowIndex++;

                CurvedAssembly1ToExcelSingly(ref hssfworkbook, ref sheet1, Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS,
                    DT1, WorkAssignmentID, BatchName, "Маркетинг", "ДУБЛЬ", "Гнутые фасады", time, cost, ref RowIndex);
                RowIndex++;

                foreach (DataRow item in DistMainOrdersDT.Select("GroupType=1"))
                {
                    CurvedAssemblyDT.Clear();
                    DT.Clear();
                    int MainOrderID = Convert.ToInt32(item["MainOrderID"]);

                    foreach (DataRow item1 in LorenzoCurvedOrdersDT.Select("MainOrderID=" + MainOrderID))
                        DT.Rows.Add(item1.ItemArray);
                    foreach (DataRow item1 in ElegantCurvedOrdersDT.Select("MainOrderID=" + MainOrderID))
                        DT.Rows.Add(item1.ItemArray);
                    foreach (DataRow item1 in Patricia1CurvedOrdersDT.Select("MainOrderID=" + MainOrderID))
                        DT.Rows.Add(item1.ItemArray);
                    foreach (DataRow item1 in KansasCurvedOrdersDT.Select("MainOrderID=" + MainOrderID))
                        DT.Rows.Add(item1.ItemArray);
                    foreach (DataRow item1 in SofiaCurvedOrdersDT.Select("MainOrderID=" + MainOrderID))
                        DT.Rows.Add(item1.ItemArray);
                    foreach (DataRow item1 in DakotaCurvedOrdersDT.Select("MainOrderID=" + MainOrderID))
                        DT.Rows.Add(item1.ItemArray);
                    foreach (DataRow item1 in Turin1CurvedOrdersDT.Select("MainOrderID=" + MainOrderID))
                        DT.Rows.Add(item1.ItemArray);
                    foreach (DataRow item1 in Turin1_1CurvedOrdersDT.Select("MainOrderID=" + MainOrderID))
                        DT.Rows.Add(item1.ItemArray);
                    foreach (DataRow item1 in Turin3CurvedOrdersDT.Select("MainOrderID=" + MainOrderID))
                        DT.Rows.Add(item1.ItemArray);
                    foreach (DataRow item1 in InfinitiCurvedOrdersDT.Select("MainOrderID=" + MainOrderID))
                        DT.Rows.Add(item1.ItemArray);

                    CurvedAssemblyCollect(DT, ref CurvedAssemblyDT);

                    DataRow[] CRows = MarketOrdersNames.Select("MainOrderID=" + MainOrderID);
                    if (CRows.Count() > 0)
                    {
                        ClientName1 = CRows[0]["ClientName"].ToString();
                        Notes = CRows[0]["Notes"].ToString();
                        OrderName = "№" + CRows[0]["OrderNumber"].ToString() + "-" + MainOrderID;
                    }

                    if (CurvedAssemblyDT.Rows.Count > 0)
                    {

                        div1 = 10;
                        div2 = 2.14m;
                        time = 0;
                        cost = 0;
                        PlanningTimebyCount(CurvedAssemblyDT, div1, div2, ref time, ref cost);
                        CurvedAssembly2ToExcelSingly(ref hssfworkbook, ref sheet1,
                            Calibri11CS, CalibriBold11CS, CalibriBold11F, TableHeaderCS, TableHeaderDecCS, CurvedAssemblyDT,
                            WorkAssignmentID, BatchName, ClientName1, OrderName, "Гнутые фасады", Notes, time, cost, ref RowIndex);
                    }

                    RowIndex++;
                }
            }
        }

        public void CurvedAssembly1ToExcelSingly(ref HSSFWorkbook hssfworkbook, ref HSSFSheet sheet1,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            DataTable DT, int WorkAssignmentID, string BatchName, string ClientName, string OperationName, string PageName, decimal Time, decimal Cost, ref int RowIndex)
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
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, PageName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Задание №" + WorkAssignmentID.ToString());
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, BatchName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Клиент:");
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, ClientName);
            cell.CellStyle = CalibriBold11CS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 4, OperationName);
            cell.CellStyle = CalibriBold11CS;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Фасад");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Цвет профиля");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Тип наполнителя");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Цвет наполнителя");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 5, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            int ColorID = -1;
            int AllTotalAmount = 0;
            int TotalAmount = 0;
            int DifferentDecorCount = 0;

            using (DataView DV = new DataView(DT))
            {
                DifferentDecorCount = DV.ToTable(true, new string[] { "ColorID" }).Rows.Count;
            }


            if (DT.Rows.Count > 0)
            {
                ColorID = Convert.ToInt32(DT.Rows[0]["ColorID"]);
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
                    if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "ColorID" ||
                        DT.Columns[y].ColumnName == "InsetTypeID" || DT.Columns[y].ColumnName == "InsetColorID")
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
                    if (ColorID != Convert.ToInt32(DT.Rows[x + 1]["ColorID"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "ColorID" ||
                                DT.Columns[y].ColumnName == "InsetTypeID" || DT.Columns[y].ColumnName == "InsetColorID")
                                continue;

                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        ColorID = Convert.ToInt32(DT.Rows[x + 1]["ColorID"]);
                        TotalAmount = 0;
                    }
                }

                if (x == DT.Rows.Count - 1)
                {
                    if (DifferentDecorCount > 1)
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "ColorID" ||
                                DT.Columns[y].ColumnName == "InsetTypeID" || DT.Columns[y].ColumnName == "InsetColorID")
                                continue;

                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                    }
                    RowIndex++;

                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "ColorID" ||
                            DT.Columns[y].ColumnName == "InsetTypeID" || DT.Columns[y].ColumnName == "InsetColorID")
                            continue;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;
                }
                RowIndex++;
            }
        }

        public void CurvedAssembly2ToExcelSingly(ref HSSFWorkbook hssfworkbook, ref HSSFSheet sheet1,
            HSSFCellStyle Calibri11CS, HSSFCellStyle CalibriBold11CS, HSSFFont CalibriBold11F, HSSFCellStyle TableHeaderCS, HSSFCellStyle TableHeaderDecCS,
            DataTable DT, int WorkAssignmentID, string BatchName, string ClientName, string OrderName, string PageName, string Notes, decimal Time, decimal Cost, ref int RowIndex)
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
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, PageName);
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
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, OrderName);
            cell.CellStyle = CalibriBold11CS;
            if (Notes.Length > 0)
            {
                cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Примечание: " + Notes);
                cell.CellStyle = CalibriBold11CS;
            }

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Фасад");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "Цвет профиля");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Тип наполнителя");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Цвет наполнителя");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Высота");
            cell.CellStyle = TableHeaderCS;
            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 5, "Кол-во");
            cell.CellStyle = TableHeaderCS;
            RowIndex++;

            int ColorID = -1;
            int AllTotalAmount = 0;
            int TotalAmount = 0;
            int DifferentDecorCount = 0;

            using (DataView DV = new DataView(DT))
            {
                DifferentDecorCount = DV.ToTable(true, new string[] { "ColorID" }).Rows.Count;
            }


            if (DT.Rows.Count > 0)
            {
                ColorID = Convert.ToInt32(DT.Rows[0]["ColorID"]);
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
                    if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "ColorID" ||
                        DT.Columns[y].ColumnName == "InsetTypeID" || DT.Columns[y].ColumnName == "InsetColorID")
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
                    if (ColorID != Convert.ToInt32(DT.Rows[x + 1]["ColorID"]))
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "ColorID" ||
                                DT.Columns[y].ColumnName == "InsetTypeID" || DT.Columns[y].ColumnName == "InsetColorID")
                                continue;

                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                        ColorID = Convert.ToInt32(DT.Rows[x + 1]["ColorID"]);
                        TotalAmount = 0;
                    }
                }

                if (x == DT.Rows.Count - 1)
                {
                    if (DifferentDecorCount > 1)
                    {
                        RowIndex++;
                        for (int y = 0; y < DT.Columns.Count; y++)
                        {
                            if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "ColorID" ||
                                DT.Columns[y].ColumnName == "InsetTypeID" || DT.Columns[y].ColumnName == "InsetColorID")
                                continue;

                            cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(string.Empty);
                            cell.CellStyle = TableHeaderCS;
                            continue;
                        }

                        cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                        cell.SetCellValue("Итого:");
                        cell.CellStyle = TableHeaderCS;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                        cell.SetCellValue(TotalAmount);
                        cell.CellStyle = TableHeaderCS;

                    }
                    RowIndex++;

                    for (int y = 0; y < DT.Columns.Count; y++)
                    {
                        if (DT.Columns[y].ColumnName == "FrontID" || DT.Columns[y].ColumnName == "ColorID" ||
                            DT.Columns[y].ColumnName == "InsetTypeID" || DT.Columns[y].ColumnName == "InsetColorID")
                            continue;

                        cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                        cell.SetCellValue(string.Empty);
                        cell.CellStyle = TableHeaderCS;
                        continue;
                    }

                    cell = sheet1.CreateRow(RowIndex).CreateCell(0);
                    cell.SetCellValue("Всего:");
                    cell.CellStyle = TableHeaderCS;

                    cell = sheet1.CreateRow(RowIndex).CreateCell(5);
                    cell.SetCellValue(AllTotalAmount);
                    cell.CellStyle = TableHeaderCS;
                }
                RowIndex++;
            }
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

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
            cell.CellStyle = Calibri11CS;

            RowIndex++;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
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

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
            cell.CellStyle = Calibri11CS;

            RowIndex++;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
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

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
            cell.CellStyle = Calibri11CS;

            RowIndex++;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
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

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, "Распечатал: Дата/время " + CurrentDate.ToString("dd.MM.yyyy HH:mm") + " \r\n ФИО: " + Security.CurrentUserShortName);
            cell.CellStyle = Calibri11CS;

            RowIndex++;

            cell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, "УТВЕРЖДАЮ_____________");
            cell.CellStyle = Calibri11CS;
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
