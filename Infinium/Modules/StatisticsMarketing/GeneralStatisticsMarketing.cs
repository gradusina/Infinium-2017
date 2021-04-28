using System;
using System.Collections;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;

namespace Infinium.Modules.StatisticsMarketing
{
    public class GeneralStatisticsMarketing : IAllFrontParameterName
    {
        private PercentageDataGrid NonAgreementDataGrid = null;
        private PercentageDataGrid AgreementDataGrid = null;
        private PercentageDataGrid OnProductionDataGrid = null;
        private PercentageDataGrid InProductionDataGrid = null;
        private PercentageDataGrid OnStorageDataGrid = null;
        private PercentageDataGrid OnExpeditionDataGrid = null;

        private CheckedListBox ClientGroupsList = null;

        private DataTable ClientGroupsDataTable = null;

        private DataTable NonAgreementDataTable = null;
        private DataTable AgreementDataTable = null;
        private DataTable OnProductionDataTable = null;
        private DataTable InProductionDataTable = null;
        private DataTable OnStorageDataTable = null;
        private DataTable OnExpeditionDataTable = null;

        private DataTable DecorProductsDataTable = null;
        private DataTable DecorDataTable = null;
        public DataTable FrontsDataTable = null;
        public DataTable FrameColorsDataTable = null;
        public DataTable PatinaDataTable = null;
        private DataTable PatinaRALDataTable = null;
        public DataTable InsetTypesDataTable = null;
        public DataTable InsetColorsDataTable = null;

        public BindingSource NonAgreementBindingSource = null;
        public BindingSource AgreementBindingSource = null;
        public BindingSource OnProductionBindingSource = null;
        public BindingSource InProductionBindingSource = null;
        public BindingSource OnStorageBindingSource = null;
        public BindingSource OnExpeditionBindingSource = null;

        private decimal NonAgreementProfilCost = 0;
        private decimal AgreementProfilCost = 0;
        private decimal OnProductionProfilCost = 0;
        private decimal InProductionProfilCost = 0;
        private decimal OnStorageProfilCost = 0;
        private decimal OnExpeditionProfilCost = 0;
        private decimal NonAgreementTPSCost = 0;
        private decimal AgreementTPSCost = 0;
        private decimal OnProductionTPSCost = 0;
        private decimal InProductionTPSCost = 0;
        private decimal OnStorageTPSCost = 0;
        private decimal OnExpeditionTPSCost = 0;

        private ArrayList ClientGroups;

        public GeneralStatisticsMarketing(ref CheckedListBox tClientGroupsList,
            ref PercentageDataGrid tNonAgreementDataGrid,
            ref PercentageDataGrid tAgreementDataGrid,
            ref PercentageDataGrid tOnProductionDataGrid,
            ref PercentageDataGrid tInProductionDataGrid,
            ref PercentageDataGrid tOnStorageDataGrid,
            ref PercentageDataGrid tOnExpeditionDataGrid)
        {
            ClientGroupsList = tClientGroupsList;
            NonAgreementDataGrid = tNonAgreementDataGrid;
            AgreementDataGrid = tAgreementDataGrid;
            OnProductionDataGrid = tOnProductionDataGrid;
            InProductionDataGrid = tInProductionDataGrid;
            OnStorageDataGrid = tOnStorageDataGrid;
            OnExpeditionDataGrid = tOnExpeditionDataGrid;

            Initialize();
        }

        private void Create()
        {
            NonAgreementDataTable = new DataTable();
            NonAgreementDataTable.Columns.Add(new DataColumn(("Status"), System.Type.GetType("System.String")));
            NonAgreementDataTable.Columns.Add(new DataColumn(("Count"), System.Type.GetType("System.Decimal")));
            NonAgreementDataTable.Columns.Add(new DataColumn(("Measure"), System.Type.GetType("System.String")));

            AgreementDataTable = new DataTable();
            AgreementDataTable.Columns.Add(new DataColumn(("Status"), System.Type.GetType("System.String")));
            AgreementDataTable.Columns.Add(new DataColumn(("Count"), System.Type.GetType("System.Decimal")));
            AgreementDataTable.Columns.Add(new DataColumn(("Measure"), System.Type.GetType("System.String")));

            OnProductionDataTable = new DataTable();
            OnProductionDataTable.Columns.Add(new DataColumn(("Status"), System.Type.GetType("System.String")));
            OnProductionDataTable.Columns.Add(new DataColumn(("Count"), System.Type.GetType("System.Decimal")));
            OnProductionDataTable.Columns.Add(new DataColumn(("Measure"), System.Type.GetType("System.String")));

            InProductionDataTable = new DataTable();
            InProductionDataTable.Columns.Add(new DataColumn(("Status"), System.Type.GetType("System.String")));
            InProductionDataTable.Columns.Add(new DataColumn(("Count"), System.Type.GetType("System.Decimal")));
            InProductionDataTable.Columns.Add(new DataColumn(("Measure"), System.Type.GetType("System.String")));//OnExpedition

            OnStorageDataTable = new DataTable();
            OnStorageDataTable.Columns.Add(new DataColumn(("Status"), System.Type.GetType("System.String")));
            OnStorageDataTable.Columns.Add(new DataColumn(("Count"), System.Type.GetType("System.Decimal")));
            OnStorageDataTable.Columns.Add(new DataColumn(("Measure"), System.Type.GetType("System.String")));

            OnExpeditionDataTable = new DataTable();
            OnExpeditionDataTable.Columns.Add(new DataColumn(("Status"), System.Type.GetType("System.String")));
            OnExpeditionDataTable.Columns.Add(new DataColumn(("Count"), System.Type.GetType("System.Decimal")));
            OnExpeditionDataTable.Columns.Add(new DataColumn(("Measure"), System.Type.GetType("System.String")));

            NonAgreementBindingSource = new BindingSource();
            AgreementBindingSource = new BindingSource();
            OnProductionBindingSource = new BindingSource();
            InProductionBindingSource = new BindingSource();
            OnStorageBindingSource = new BindingSource();
            OnExpeditionBindingSource = new BindingSource();

            FrontsDataTable = new DataTable();
            FrameColorsDataTable = new DataTable();
            PatinaDataTable = new DataTable();
            InsetTypesDataTable = new DataTable();
            InsetColorsDataTable = new DataTable();
            DecorProductsDataTable = new DataTable();
            DecorDataTable = new DataTable();
            ClientGroupsDataTable = new DataTable();
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
            ClientGroups = new ArrayList();

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

            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM ClientGroups",
                ConnectionStrings.MarketingReferenceConnectionString))
            {
                DA.Fill(ClientGroupsDataTable);
            }
        }

        private void Binding()
        {
            NonAgreementBindingSource.DataSource = NonAgreementDataTable;
            NonAgreementDataGrid.DataSource = NonAgreementBindingSource;

            AgreementBindingSource.DataSource = AgreementDataTable;
            AgreementDataGrid.DataSource = AgreementBindingSource;

            OnProductionBindingSource.DataSource = OnProductionDataTable;
            OnProductionDataGrid.DataSource = OnProductionBindingSource;

            InProductionBindingSource.DataSource = InProductionDataTable;
            InProductionDataGrid.DataSource = InProductionBindingSource;

            OnStorageBindingSource.DataSource = OnStorageDataTable;
            OnStorageDataGrid.DataSource = OnStorageBindingSource;

            OnExpeditionBindingSource.DataSource = OnExpeditionDataTable;
            OnExpeditionDataGrid.DataSource = OnExpeditionBindingSource;

            ClientGroupsList.DataSource = ClientGroupsDataTable;
            ClientGroupsList.DisplayMember = "ClientGroupName";
            ClientGroupsList.ValueMember = "ClientGroupID";
        }

        private void SettingGrid()
        {
            NumberFormatInfo nfi1 = new NumberFormatInfo()
            {
                CurrencyGroupSeparator = " ",
                CurrencySymbol = string.Empty,
                CurrencyDecimalDigits = 1
            };
            NonAgreementDataGrid.Columns["Status"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            NonAgreementDataGrid.Columns["Status"].MinimumWidth = 100;
            NonAgreementDataGrid.Columns["Count"].DefaultCellStyle.Format = "C";
            NonAgreementDataGrid.Columns["Count"].DefaultCellStyle.FormatProvider = nfi1;
            NonAgreementDataGrid.Columns["Count"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            NonAgreementDataGrid.Columns["Count"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            NonAgreementDataGrid.Columns["Count"].Width = 100;
            NonAgreementDataGrid.Columns["Measure"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            NonAgreementDataGrid.Columns["Measure"].Width = 65;

            AgreementDataGrid.Columns["Status"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            AgreementDataGrid.Columns["Status"].MinimumWidth = 100;
            AgreementDataGrid.Columns["Count"].DefaultCellStyle.Format = "C";
            AgreementDataGrid.Columns["Count"].DefaultCellStyle.FormatProvider = nfi1;
            AgreementDataGrid.Columns["Count"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            AgreementDataGrid.Columns["Count"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            AgreementDataGrid.Columns["Count"].Width = 100;
            AgreementDataGrid.Columns["Measure"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            AgreementDataGrid.Columns["Measure"].Width = 65;

            OnProductionDataGrid.Columns["Status"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            OnProductionDataGrid.Columns["Status"].MinimumWidth = 100;
            OnProductionDataGrid.Columns["Count"].DefaultCellStyle.Format = "C";
            OnProductionDataGrid.Columns["Count"].DefaultCellStyle.FormatProvider = nfi1;
            OnProductionDataGrid.Columns["Count"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            OnProductionDataGrid.Columns["Count"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            OnProductionDataGrid.Columns["Count"].Width = 100;
            OnProductionDataGrid.Columns["Measure"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            OnProductionDataGrid.Columns["Measure"].Width = 65;

            InProductionDataGrid.Columns["Status"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            InProductionDataGrid.Columns["Status"].MinimumWidth = 100;
            InProductionDataGrid.Columns["Count"].DefaultCellStyle.Format = "C";
            InProductionDataGrid.Columns["Count"].DefaultCellStyle.FormatProvider = nfi1;
            InProductionDataGrid.Columns["Count"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            InProductionDataGrid.Columns["Count"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            InProductionDataGrid.Columns["Count"].Width = 100;
            InProductionDataGrid.Columns["Measure"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            InProductionDataGrid.Columns["Measure"].Width = 65;

            OnStorageDataGrid.Columns["Status"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            OnStorageDataGrid.Columns["Status"].MinimumWidth = 100;
            OnStorageDataGrid.Columns["Count"].DefaultCellStyle.Format = "C";
            OnStorageDataGrid.Columns["Count"].DefaultCellStyle.FormatProvider = nfi1;
            OnStorageDataGrid.Columns["Count"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            OnStorageDataGrid.Columns["Count"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            OnStorageDataGrid.Columns["Count"].Width = 100;
            OnStorageDataGrid.Columns["Measure"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            OnStorageDataGrid.Columns["Measure"].Width = 65;

            OnExpeditionDataGrid.Columns["Status"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            OnExpeditionDataGrid.Columns["Status"].MinimumWidth = 100;
            OnExpeditionDataGrid.Columns["Count"].DefaultCellStyle.Format = "C";
            OnExpeditionDataGrid.Columns["Count"].DefaultCellStyle.FormatProvider = nfi1;
            OnExpeditionDataGrid.Columns["Count"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            OnExpeditionDataGrid.Columns["Count"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            OnExpeditionDataGrid.Columns["Count"].Width = 100;
            OnExpeditionDataGrid.Columns["Measure"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            OnExpeditionDataGrid.Columns["Measure"].Width = 65;
        }

        public void Initialize()
        {
            Create();
            Fill();
            Binding();
            SettingGrid();

            for (int i = 0; i < ClientGroupsList.Items.Count; i++)
                ClientGroupsList.SetItemChecked(i, true);
            ClientGroupsList.SetItemChecked(1, false);
        }

        private void FillSummaryTable(DataTable Table, decimal ProfilSquare, decimal ProfilPogon, decimal ProfilCount,
            decimal TPSSquare, decimal TPSPogon, decimal TPSCount, decimal CurvedCount)
        {
            DataRow NewRow1 = Table.NewRow();
            NewRow1["Status"] = "Фасады";
            NewRow1["Count"] = ProfilSquare + TPSSquare;
            NewRow1["Measure"] = "м.кв.";
            Table.Rows.Add(NewRow1);

            DataRow NewRow2 = Table.NewRow();
            NewRow2["Status"] = "   Профиль";
            NewRow2["Count"] = ProfilSquare;
            NewRow2["Measure"] = "м.кв.";
            Table.Rows.Add(NewRow2);

            DataRow NewRow3 = Table.NewRow();
            NewRow3["Status"] = "   ТПС";
            NewRow3["Count"] = TPSSquare;
            NewRow3["Measure"] = "м.кв.";
            Table.Rows.Add(NewRow3);

            DataRow NewRow4 = Table.NewRow();
            Table.Rows.Add(NewRow4);

            DataRow CurvedRow = Table.NewRow();
            CurvedRow["Status"] = "Гнутые";
            CurvedRow["Count"] = CurvedCount;
            CurvedRow["Measure"] = "шт.";
            Table.Rows.Add(CurvedRow);

            DataRow NewRow12 = Table.NewRow();
            Table.Rows.Add(NewRow12);

            DataRow NewRow5 = Table.NewRow();
            NewRow5["Status"] = "Погонаж";
            NewRow5["Count"] = ProfilPogon + TPSPogon;
            NewRow5["Measure"] = "м.п.";
            Table.Rows.Add(NewRow5);

            DataRow NewRow6 = Table.NewRow();
            NewRow6["Status"] = "   Профиль";
            NewRow6["Count"] = ProfilPogon;
            NewRow6["Measure"] = "м.п.";
            Table.Rows.Add(NewRow6);

            DataRow NewRow7 = Table.NewRow();
            NewRow7["Status"] = "   ТПС";
            NewRow7["Count"] = TPSPogon;
            NewRow7["Measure"] = "м.п.";
            Table.Rows.Add(NewRow7);

            DataRow NewRow8 = Table.NewRow();
            Table.Rows.Add(NewRow8);

            DataRow NewRow9 = Table.NewRow();
            NewRow9["Status"] = "Декор";
            NewRow9["Count"] = ProfilCount + TPSCount;
            NewRow9["Measure"] = "шт.";
            Table.Rows.Add(NewRow9);

            DataRow NewRow10 = Table.NewRow();
            NewRow10["Status"] = "   Профиль";
            NewRow10["Count"] = ProfilCount;
            NewRow10["Measure"] = "шт.";
            Table.Rows.Add(NewRow10);

            DataRow NewRow11 = Table.NewRow();
            NewRow11["Status"] = "   ТПС";
            NewRow11["Count"] = TPSCount;
            NewRow11["Measure"] = "шт.";
            Table.Rows.Add(NewRow11);
        }

        public void NonAgreementStatistics()
        {
            decimal NonAgreementProfilSquare = 0;
            decimal NonAgreementProfilPogon = 0;
            decimal NonAgreementProfilCount = 0;
            decimal NonAgreementTPSSquare = 0;
            decimal NonAgreementTPSPogon = 0;
            decimal NonAgreementTPSCount = 0;
            int NonAgreementCurvedCount = 0;

            GeneralNonAgreementOrders(ref NonAgreementProfilSquare, ref NonAgreementProfilPogon, ref NonAgreementProfilCount,
                ref NonAgreementTPSSquare, ref NonAgreementTPSPogon, ref NonAgreementTPSCount, ref NonAgreementCurvedCount);

            NonAgreementDataTable.Clear();
            FillSummaryTable(NonAgreementDataTable, NonAgreementProfilSquare, NonAgreementProfilPogon, NonAgreementProfilCount,
                NonAgreementTPSSquare, NonAgreementTPSPogon, NonAgreementTPSCount, NonAgreementCurvedCount);
        }

        public void AgreementStatistics()
        {
            decimal AgreementProfilSquare = 0;
            decimal AgreementProfilPogon = 0;
            decimal AgreementProfilCount = 0;
            decimal AgreementTPSSquare = 0;
            decimal AgreementTPSPogon = 0;
            decimal AgreementTPSCount = 0;
            decimal AgreementCurvedCount = 0;

            GeneralAgreementOrders(ref AgreementProfilSquare, ref AgreementProfilPogon, ref AgreementProfilCount,
                ref AgreementTPSSquare, ref AgreementTPSPogon, ref AgreementTPSCount, ref AgreementCurvedCount);

            AgreementDataTable.Clear();
            FillSummaryTable(AgreementDataTable, AgreementProfilSquare, AgreementProfilPogon, AgreementProfilCount,
                AgreementTPSSquare, AgreementTPSPogon, AgreementTPSCount, AgreementCurvedCount);
        }

        public void OnProductionStatistics()
        {
            decimal OnProductionProfilSquare = 0;
            decimal OnProductionProfilPogon = 0;
            decimal OnProductionProfilCount = 0;
            decimal OnProductionTPSSquare = 0;
            decimal OnProductionTPSPogon = 0;
            decimal OnProductionTPSCount = 0;
            decimal OnProductionCurvedCount = 0;

            GeneralOnProductionOrders(ref OnProductionProfilSquare, ref OnProductionProfilPogon, ref OnProductionProfilCount,
                ref OnProductionTPSSquare, ref OnProductionTPSPogon, ref OnProductionTPSCount, ref OnProductionCurvedCount);

            OnProductionDataTable.Clear();
            FillSummaryTable(OnProductionDataTable, OnProductionProfilSquare, OnProductionProfilPogon, OnProductionProfilCount,
                OnProductionTPSSquare, OnProductionTPSPogon, OnProductionTPSCount, OnProductionCurvedCount);
        }

        public void InProductionStatistics()
        {
            decimal InProductionProfilSquare = 0;
            decimal InProductionProfilPogon = 0;
            decimal InProductionProfilCount = 0;
            decimal InProductionTPSSquare = 0;
            decimal InProductionTPSPogon = 0;
            decimal InProductionTPSCount = 0;
            decimal InProductionCurvedCount = 0;

            GeneralInProductionOrders(ref InProductionProfilSquare, ref InProductionProfilPogon, ref InProductionProfilCount,
                ref InProductionTPSSquare, ref InProductionTPSPogon, ref InProductionTPSCount, ref InProductionCurvedCount);

            InProductionDataTable.Clear();
            FillSummaryTable(InProductionDataTable, InProductionProfilSquare, InProductionProfilPogon, InProductionProfilCount,
                InProductionTPSSquare, InProductionTPSPogon, InProductionTPSCount, InProductionCurvedCount);
        }

        public void OnStorageStatistics()
        {
            decimal OnStorageProfilSquare = 0;
            decimal OnStorageProfilPogon = 0;
            decimal OnStorageProfilCount = 0;
            decimal OnStorageTPSSquare = 0;
            decimal OnStorageTPSPogon = 0;
            decimal OnStorageTPSCount = 0;
            decimal OnStorageCurvedCount = 0;

            FillTables(DateTime.Now, ref OnStorageProfilSquare, ref OnStorageProfilPogon, ref OnStorageProfilCount,
                ref OnStorageTPSSquare, ref OnStorageTPSPogon, ref OnStorageTPSCount, ref OnStorageCurvedCount);

            OnStorageDataTable.Clear();
            FillSummaryTable(OnStorageDataTable, OnStorageProfilSquare, OnStorageProfilPogon, OnStorageProfilCount,
                OnStorageTPSSquare, OnStorageTPSPogon, OnStorageTPSCount, OnStorageCurvedCount);
        }

        public void OnExpeditionStatistics()
        {
            decimal OnExpeditionProfilSquare = 0;
            decimal OnExpeditionProfilPogon = 0;
            decimal OnExpeditionProfilCount = 0;
            decimal OnExpeditionTPSSquare = 0;
            decimal OnExpeditionTPSPogon = 0;
            decimal OnExpeditionTPSCount = 0;
            decimal OnExpeditionCurvedCount = 0;

            GeneralOnExpeditionOrders(ref OnExpeditionProfilSquare, ref OnExpeditionProfilPogon, ref OnExpeditionProfilCount,
                ref OnExpeditionTPSSquare, ref OnExpeditionTPSPogon, ref OnExpeditionTPSCount, ref OnExpeditionCurvedCount);

            OnExpeditionDataTable.Clear();
            FillSummaryTable(OnExpeditionDataTable, OnExpeditionProfilSquare, OnExpeditionProfilPogon, OnExpeditionProfilCount,
                OnExpeditionTPSSquare, OnExpeditionTPSPogon, OnExpeditionTPSCount, OnExpeditionCurvedCount);
        }

        public void GetNonAgreementCost(ref decimal ProfilCost, ref decimal TPSCost)
        {
            ProfilCost = NonAgreementProfilCost;
            TPSCost = NonAgreementTPSCost;
        }

        public void GetAgreementCost(ref decimal ProfilCost, ref decimal TPSCost)
        {
            ProfilCost = AgreementProfilCost;
            TPSCost = AgreementTPSCost;
        }

        public void GetOnProductionCost(ref decimal ProfilCost, ref decimal TPSCost)
        {
            ProfilCost = OnProductionProfilCost;
            TPSCost = OnProductionTPSCost;
        }

        public void GetInProductionCost(ref decimal ProfilCost, ref decimal TPSCost)
        {
            ProfilCost = InProductionProfilCost;
            TPSCost = InProductionTPSCost;
        }

        public void GetOnStorageCost(ref decimal ProfilCost, ref decimal TPSCost)
        {
            ProfilCost = OnStorageProfilCost;
            TPSCost = OnStorageTPSCost;
        }

        public void GetOnExpeditionCost(ref decimal ProfilCost, ref decimal TPSCost)
        {
            ProfilCost = OnExpeditionProfilCost;
            TPSCost = OnExpeditionTPSCost;
        }

        public void GetClientGroup()
        {
            ClientGroups.Clear();
            foreach (object itemChecked in ClientGroupsList.CheckedItems)
            {
                DataRowView castedItem = (DataRowView)itemChecked;
                ClientGroups.Add(Convert.ToInt32(castedItem["ClientGroupID"]));
            }
        }

        public void GeneralNonAgreementOrders(ref decimal ProfilSquare,
            ref decimal ProfilPogon, ref decimal ProfilCount, ref decimal TPSSquare,
            ref decimal TPSPogon, ref decimal TPSCount, ref int CurvedCount)
        {
            bool CheckZOV = GetZOV;
            string MegaOrderFilter = " AND ClientID = -1";

            NonAgreementProfilCost = 0;
            NonAgreementTPSCost = 0;

            if (ClientGroups.Count > 0)
            {
                MegaOrderFilter = " AND ClientID IN (SELECT ClientID FROM infiniu2_marketingreference.dbo.Clients" +
                    " WHERE ClientGroupID IN (" + string.Join(",", ClientGroups.OfType<int>().ToArray()) + "))";
            }

            using (DataTable DT = new DataTable())
            {
                //ФАСАДЫ НА СОГЛАСОВАНИИ
                //Профиль
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM NewFrontsOrders" +
                    " WHERE FactoryID=1 AND MainOrderID IN (SELECT MainOrderID FROM NewMainOrders" +
                    " WHERE ProfilProductionStatusID=1" +
                    " AND ProfilStorageStatusID=1" +
                    " AND ProfilExpeditionStatusID=1" +
                    " AND ProfilDispatchStatusID=1" +
                    " AND MegaOrderID IN (SELECT MegaOrderID FROM NewMegaOrders" +
                    " WHERE ((AgreementStatusID=0 AND CreatedByClient=0) OR AgreementStatusID=1)" + MegaOrderFilter + "))", ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DT.Clear();
                    DA.Fill(DT);

                    foreach (DataRow Row in DT.Rows)
                    {
                        if (Convert.ToInt32(Row["Width"]) == -1)
                        {
                            CurvedCount += Convert.ToInt32(Row["Count"]);
                        }
                        else
                        {
                            ProfilSquare += Convert.ToDecimal(Row["Square"]);
                        }
                        NonAgreementProfilCost += Convert.ToDecimal(Row["Cost"]);
                    }
                }
                if (CheckZOV)
                {
                    using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM FrontsOrders" +
                        " WHERE FactoryID=1 AND MainOrderID IN (SELECT MainOrderID FROM MainOrders" +
                        " WHERE MegaOrderID = 0)", ConnectionStrings.ZOVOrdersConnectionString))
                    {
                        DT.Clear();
                        DA.Fill(DT);

                        foreach (DataRow Row in DT.Rows)
                        {
                            if (Convert.ToInt32(Row["Width"]) == -1)
                            {
                                CurvedCount += Convert.ToInt32(Row["Count"]);
                            }
                            else
                            {
                                ProfilSquare += Convert.ToDecimal(Row["Square"]);
                            }
                            NonAgreementProfilCost += Convert.ToDecimal(Row["Cost"]);
                        }
                    }
                }
                ProfilSquare = Decimal.Round(ProfilSquare, 1, MidpointRounding.AwayFromZero);
                NonAgreementProfilCost = Decimal.Round(NonAgreementProfilCost, 1, MidpointRounding.AwayFromZero);

                //ТПС
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM NewFrontsOrders" +
                    " WHERE FactoryID=2 AND MainOrderID IN (SELECT MainOrderID FROM NewMainOrders" +
                    " WHERE TPSProductionStatusID=1" +
                    " AND TPSStorageStatusID=1" +
                    " AND TPSExpeditionStatusID=1" +
                    " AND TPSDispatchStatusID=1" +
                    " AND MegaOrderID IN (SELECT MegaOrderID FROM NewMegaOrders" +
                    " WHERE ((AgreementStatusID=0 AND CreatedByClient=0) OR AgreementStatusID=1)" + MegaOrderFilter + "))", ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DT.Clear();
                    DA.Fill(DT);

                    foreach (DataRow Row in DT.Rows)
                    {
                        if (Convert.ToInt32(Row["Width"]) == -1)
                        {
                            CurvedCount += Convert.ToInt32(Row["Count"]);
                        }
                        else
                        {
                            TPSSquare += Convert.ToDecimal(Row["Square"]);
                        }
                        NonAgreementTPSCost += Convert.ToDecimal(Row["Cost"]);
                    }
                }

                if (CheckZOV)
                {
                    using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM FrontsOrders" +
                        " WHERE FactoryID=2 AND MainOrderID IN (SELECT MainOrderID FROM MainOrders" +
                        " WHERE MegaOrderID = 0)", ConnectionStrings.ZOVOrdersConnectionString))
                    {
                        DT.Clear();
                        DA.Fill(DT);

                        foreach (DataRow Row in DT.Rows)
                        {
                            if (Convert.ToInt32(Row["Width"]) == -1)
                            {
                                CurvedCount += Convert.ToInt32(Row["Count"]);
                            }
                            else
                            {
                                TPSSquare += Convert.ToDecimal(Row["Square"]);
                            }
                            NonAgreementTPSCost += Convert.ToDecimal(Row["Cost"]);
                        }
                    }
                }
                TPSSquare = Decimal.Round(TPSSquare, 1, MidpointRounding.AwayFromZero);
                NonAgreementTPSCost = Decimal.Round(NonAgreementTPSCost, 1, MidpointRounding.AwayFromZero);

                //ДЕКОР НА СОГЛАСОВАНИИ
                //Профиль
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT NewDecorOrders.*, infiniu2_catalog.dbo.DecorConfig.MeasureID FROM NewDecorOrders" +
                    " INNER JOIN infiniu2_catalog.dbo.DecorConfig ON NewDecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                    " WHERE NewDecorOrders.FactoryID=1" +
                    " AND MainOrderID IN (SELECT MainOrderID FROM NewMainOrders" +
                    " WHERE ProfilProductionStatusID=1" +
                    " AND ProfilStorageStatusID=1" +
                    " AND ProfilExpeditionStatusID=1" +
                    " AND ProfilDispatchStatusID=1" +
                    " AND MegaOrderID IN (SELECT MegaOrderID FROM NewMegaOrders" +
                    " WHERE ((AgreementStatusID=0 AND CreatedByClient=0) OR AgreementStatusID=1)" + MegaOrderFilter + "))", ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DT.Clear();
                    DA.Fill(DT);

                    foreach (DataRow Row in DT.Rows)
                    {
                        if (Convert.ToInt32(Row["MeasureID"]) == 2)
                        {
                            //нет параметра "высота"
                            if (Row["Height"].ToString() == "-1")
                                ProfilPogon += Convert.ToDecimal(Row["Length"]) * Convert.ToDecimal(Row["Count"]) / 1000;
                            else
                                ProfilPogon += Convert.ToDecimal(Row["Height"]) * Convert.ToDecimal(Row["Count"]) / 1000;
                        }
                        else
                        {
                            ProfilCount += Convert.ToDecimal(Row["Count"]);
                        }
                        NonAgreementProfilCost += Convert.ToDecimal(Row["Cost"]);
                    }
                }

                if (CheckZOV)
                {
                    using (SqlDataAdapter DA = new SqlDataAdapter("SELECT DecorOrders.*, infiniu2_catalog.dbo.DecorConfig.MeasureID FROM DecorOrders" +
                        " INNER JOIN infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                        " WHERE DecorOrders.FactoryID=1" +
                        " AND MainOrderID IN (SELECT MainOrderID FROM MainOrders" +
                        " WHERE MegaOrderID = 0)", ConnectionStrings.ZOVOrdersConnectionString))
                    {
                        DT.Clear();
                        DA.Fill(DT);

                        foreach (DataRow Row in DT.Rows)
                        {
                            if (Convert.ToInt32(Row["MeasureID"]) == 2)
                            {
                                //нет параметра "высота"
                                if (Row["Height"].ToString() == "-1")
                                    ProfilPogon += Convert.ToDecimal(Row["Length"]) * Convert.ToDecimal(Row["Count"]) / 1000;
                                else
                                    ProfilPogon += Convert.ToDecimal(Row["Height"]) * Convert.ToDecimal(Row["Count"]) / 1000;
                            }
                            else
                            {
                                ProfilCount += Convert.ToDecimal(Row["Count"]);
                            }
                            NonAgreementProfilCost += Convert.ToDecimal(Row["Cost"]);
                        }
                    }
                }
                NonAgreementProfilCost = Decimal.Round(NonAgreementProfilCost, 1, MidpointRounding.AwayFromZero);
                ProfilPogon = Decimal.Round(ProfilPogon, 1, MidpointRounding.AwayFromZero);
                ProfilCount = Decimal.Round(ProfilCount, 1, MidpointRounding.AwayFromZero);

                //ТПС
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT NewDecorOrders.*, infiniu2_catalog.dbo.DecorConfig.MeasureID FROM NewDecorOrders" +
                    " INNER JOIN infiniu2_catalog.dbo.DecorConfig ON NewDecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                    " WHERE NewDecorOrders.FactoryID=2" +
                    " AND MainOrderID IN (SELECT MainOrderID FROM NewMainOrders" +
                    " WHERE TPSProductionStatusID=1" +
                    " AND TPSStorageStatusID=1" +
                    " AND TPSExpeditionStatusID=1" +
                    " AND TPSDispatchStatusID=1" +
                    " AND MegaOrderID IN (SELECT MegaOrderID FROM NewMegaOrders" +
                    " WHERE ((AgreementStatusID=0 AND CreatedByClient=0) OR AgreementStatusID=1)" + MegaOrderFilter + "))", ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DT.Clear();
                    DA.Fill(DT);

                    foreach (DataRow Row in DT.Rows)
                    {
                        if (Convert.ToInt32(Row["MeasureID"]) == 2)
                        {
                            //нет параметра "высота"
                            if (Row["Height"].ToString() == "-1")
                                TPSPogon += Convert.ToDecimal(Row["Length"]) * Convert.ToDecimal(Row["Count"]) / 1000;
                            else
                                TPSPogon += Convert.ToDecimal(Row["Height"]) * Convert.ToDecimal(Row["Count"]) / 1000;
                        }
                        else
                        {
                            TPSCount += Convert.ToDecimal(Row["Count"]);
                        }
                        NonAgreementTPSCost += Convert.ToDecimal(Row["Cost"]);
                    }
                }

                if (CheckZOV)
                {
                    using (SqlDataAdapter DA = new SqlDataAdapter("SELECT DecorOrders.*, infiniu2_catalog.dbo.DecorConfig.MeasureID FROM DecorOrders" +
                        " INNER JOIN infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                        " WHERE DecorOrders.FactoryID=2" +
                        " AND MainOrderID IN (SELECT MainOrderID FROM MainOrders" +
                        " WHERE MegaOrderID = 0)", ConnectionStrings.ZOVOrdersConnectionString))
                    {
                        DT.Clear();
                        DA.Fill(DT);

                        foreach (DataRow Row in DT.Rows)
                        {
                            if (Convert.ToInt32(Row["MeasureID"]) == 2)
                            {
                                //нет параметра "высота"
                                if (Row["Height"].ToString() == "-1")
                                    TPSPogon += Convert.ToDecimal(Row["Length"]) * Convert.ToDecimal(Row["Count"]) / 1000;
                                else
                                    TPSPogon += Convert.ToDecimal(Row["Height"]) * Convert.ToDecimal(Row["Count"]) / 1000;
                            }
                            else
                            {
                                TPSCount += Convert.ToDecimal(Row["Count"]);
                            }
                            NonAgreementTPSCost += Convert.ToDecimal(Row["Cost"]);
                        }
                    }
                }
                NonAgreementTPSCost = Decimal.Round(NonAgreementTPSCost, 1, MidpointRounding.AwayFromZero);
                TPSPogon = Decimal.Round(TPSPogon, 1, MidpointRounding.AwayFromZero);
                TPSCount = Decimal.Round(TPSCount, 1, MidpointRounding.AwayFromZero);
            }
        }

        public void GeneralAgreementOrders(ref decimal ProfilSquare,
            ref decimal ProfilPogon, ref decimal ProfilCount, ref decimal TPSSquare,
            ref decimal TPSPogon, ref decimal TPSCount, ref decimal CurvedCount)
        {
            bool CheckZOV = GetZOV;
            string MegaOrderFilter = " AND ClientID = -1";

            AgreementProfilCost = 0;
            AgreementTPSCost = 0;

            if (ClientGroups.Count > 0)
            {
                MegaOrderFilter = " AND ClientID IN (SELECT ClientID FROM infiniu2_marketingreference.dbo.Clients" +
                    " WHERE ClientGroupID IN (" + string.Join(",", ClientGroups.OfType<int>().ToArray()) + "))";
            }

            using (DataTable DT = new DataTable())
            {
                //ФАСАДЫ НА СОГЛАСОВАНИИ
                //Профиль
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM FrontsOrders" +
                    " WHERE FactoryID=1 AND MainOrderID IN (SELECT MainOrderID FROM MainOrders" +
                    " WHERE ProfilProductionStatusID=1" +
                    " AND ProfilStorageStatusID=1" +
                    " AND ProfilExpeditionStatusID=1" +
                    " AND ProfilDispatchStatusID=1" +
                    " AND MegaOrderID IN (SELECT MegaOrderID FROM MegaOrders" +
                    " WHERE AgreementStatusID=2" + MegaOrderFilter + "))", ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DT.Clear();
                    DA.Fill(DT);

                    foreach (DataRow Row in DT.Rows)
                    {
                        if (Convert.ToInt32(Row["Width"]) == -1)
                        {
                            CurvedCount += Convert.ToInt32(Row["Count"]);
                        }
                        else
                        {
                            ProfilSquare += Convert.ToDecimal(Row["Square"]);
                        }
                        AgreementProfilCost += Convert.ToDecimal(Row["Cost"]);
                    }
                }
                if (CheckZOV)
                {
                    using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM FrontsOrders" +
                        " WHERE FactoryID=1 AND MainOrderID IN (SELECT MainOrderID FROM MainOrders" +
                        " WHERE ProfilProductionStatusID=1" +
                        " AND ProfilStorageStatusID=1" +
                        " AND ProfilExpeditionStatusID=1" +
                        " AND ProfilDispatchStatusID=1)", ConnectionStrings.ZOVOrdersConnectionString))
                    {
                        DT.Clear();
                        DA.Fill(DT);

                        foreach (DataRow Row in DT.Rows)
                        {
                            if (Convert.ToInt32(Row["Width"]) == -1)
                            {
                                CurvedCount += Convert.ToInt32(Row["Count"]);
                            }
                            else
                            {
                                ProfilSquare += Convert.ToDecimal(Row["Square"]);
                            }
                            AgreementProfilCost += Convert.ToDecimal(Row["Cost"]);
                        }
                    }
                }
                ProfilSquare = Decimal.Round(ProfilSquare, 1, MidpointRounding.AwayFromZero);
                AgreementProfilCost = Decimal.Round(AgreementProfilCost, 1, MidpointRounding.AwayFromZero);

                //ТПС
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM FrontsOrders" +
                    " WHERE FactoryID=2 AND MainOrderID IN (SELECT MainOrderID FROM MainOrders" +
                    " WHERE TPSProductionStatusID=1" +
                    " AND TPSStorageStatusID=1" +
                    " AND TPSExpeditionStatusID=1" +
                    " AND TPSDispatchStatusID=1" +
                    " AND MegaOrderID IN (SELECT MegaOrderID FROM MegaOrders" +
                    " WHERE AgreementStatusID=2" + MegaOrderFilter + "))", ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DT.Clear();
                    DA.Fill(DT);

                    foreach (DataRow Row in DT.Rows)
                    {
                        if (Convert.ToInt32(Row["Width"]) == -1)
                        {
                            CurvedCount += Convert.ToInt32(Row["Count"]);
                        }
                        else
                        {
                            TPSSquare += Convert.ToDecimal(Row["Square"]);
                        }
                        AgreementTPSCost += Convert.ToDecimal(Row["Cost"]);
                    }
                }

                if (CheckZOV)
                {
                    using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM FrontsOrders" +
                        " WHERE FactoryID=2 AND MainOrderID IN (SELECT MainOrderID FROM MainOrders" +
                        " WHERE TPSProductionStatusID=1" +
                        " AND TPSStorageStatusID=1" +
                    " AND TPSExpeditionStatusID=1" +
                        " AND TPSDispatchStatusID=1)", ConnectionStrings.ZOVOrdersConnectionString))
                    {
                        DT.Clear();
                        DA.Fill(DT);

                        foreach (DataRow Row in DT.Rows)
                        {
                            if (Convert.ToInt32(Row["Width"]) == -1)
                            {
                                CurvedCount += Convert.ToInt32(Row["Count"]);
                            }
                            else
                            {
                                TPSSquare += Convert.ToDecimal(Row["Square"]);
                            }
                            AgreementTPSCost += Convert.ToDecimal(Row["Cost"]);
                        }
                    }
                }
                TPSSquare = Decimal.Round(TPSSquare, 1, MidpointRounding.AwayFromZero);
                AgreementTPSCost = Decimal.Round(AgreementTPSCost, 1, MidpointRounding.AwayFromZero);

                //ДЕКОР НА СОГЛАСОВАНИИ
                //Профиль
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT DecorOrders.*, infiniu2_catalog.dbo.DecorConfig.MeasureID FROM DecorOrders" +
                    " INNER JOIN infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                    " WHERE DecorOrders.FactoryID=1" +
                    " AND MainOrderID IN (SELECT MainOrderID FROM MainOrders" +
                    " WHERE ProfilProductionStatusID=1" +
                    " AND ProfilStorageStatusID=1" +
                    " AND ProfilExpeditionStatusID=1" +
                    " AND ProfilDispatchStatusID=1" +
                    " AND MegaOrderID IN (SELECT MegaOrderID FROM MegaOrders" +
                    " WHERE AgreementStatusID=2" + MegaOrderFilter + "))", ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DT.Clear();
                    DA.Fill(DT);

                    foreach (DataRow Row in DT.Rows)
                    {
                        if (Convert.ToInt32(Row["MeasureID"]) == 2)
                        {
                            //нет параметра "высота"
                            if (Row["Height"].ToString() == "-1")
                                ProfilPogon += Convert.ToDecimal(Row["Length"]) * Convert.ToDecimal(Row["Count"]) / 1000;
                            else
                                ProfilPogon += Convert.ToDecimal(Row["Height"]) * Convert.ToDecimal(Row["Count"]) / 1000;
                        }
                        else
                        {
                            ProfilCount += Convert.ToDecimal(Row["Count"]);
                        }
                        AgreementProfilCost += Convert.ToDecimal(Row["Cost"]);
                    }
                }

                if (CheckZOV)
                {
                    using (SqlDataAdapter DA = new SqlDataAdapter("SELECT DecorOrders.*, infiniu2_catalog.dbo.DecorConfig.MeasureID FROM DecorOrders" +
                        " INNER JOIN infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                        " WHERE DecorOrders.FactoryID=1" +
                        " AND MainOrderID IN (SELECT MainOrderID FROM MainOrders" +
                        " WHERE ProfilProductionStatusID=1" +
                        " AND ProfilStorageStatusID=1" +
                        " AND ProfilExpeditionStatusID=1" +
                        " AND ProfilDispatchStatusID=1)", ConnectionStrings.ZOVOrdersConnectionString))
                    {
                        DT.Clear();
                        DA.Fill(DT);

                        foreach (DataRow Row in DT.Rows)
                        {
                            if (Convert.ToInt32(Row["MeasureID"]) == 2)
                            {
                                //нет параметра "высота"
                                if (Row["Height"].ToString() == "-1")
                                    ProfilPogon += Convert.ToDecimal(Row["Length"]) * Convert.ToDecimal(Row["Count"]) / 1000;
                                else
                                    ProfilPogon += Convert.ToDecimal(Row["Height"]) * Convert.ToDecimal(Row["Count"]) / 1000;
                            }
                            else
                            {
                                ProfilCount += Convert.ToDecimal(Row["Count"]);
                            }
                            AgreementProfilCost += Convert.ToDecimal(Row["Cost"]);
                        }
                    }
                }
                AgreementProfilCost = Decimal.Round(AgreementProfilCost, 1, MidpointRounding.AwayFromZero);
                ProfilPogon = Decimal.Round(ProfilPogon, 1, MidpointRounding.AwayFromZero);
                ProfilCount = Decimal.Round(ProfilCount, 1, MidpointRounding.AwayFromZero);

                //ТПС
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT DecorOrders.*, infiniu2_catalog.dbo.DecorConfig.MeasureID FROM DecorOrders" +
                    " INNER JOIN infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                    " WHERE DecorOrders.FactoryID=2" +
                    " AND MainOrderID IN (SELECT MainOrderID FROM MainOrders" +
                    " WHERE TPSProductionStatusID=1" +
                    " AND TPSStorageStatusID=1" +
                    " AND TPSExpeditionStatusID=1" +
                    " AND TPSDispatchStatusID=1" +
                    " AND MegaOrderID IN (SELECT MegaOrderID FROM MegaOrders" +
                    " WHERE AgreementStatusID=2" + MegaOrderFilter + "))", ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DT.Clear();
                    DA.Fill(DT);

                    foreach (DataRow Row in DT.Rows)
                    {
                        if (Convert.ToInt32(Row["MeasureID"]) == 2)
                        {
                            //нет параметра "высота"
                            if (Row["Height"].ToString() == "-1")
                                TPSPogon += Convert.ToDecimal(Row["Length"]) * Convert.ToDecimal(Row["Count"]) / 1000;
                            else
                                TPSPogon += Convert.ToDecimal(Row["Height"]) * Convert.ToDecimal(Row["Count"]) / 1000;
                        }
                        else
                        {
                            TPSCount += Convert.ToDecimal(Row["Count"]);
                        }
                        AgreementTPSCost += Convert.ToDecimal(Row["Cost"]);
                    }
                }

                if (CheckZOV)
                {
                    using (SqlDataAdapter DA = new SqlDataAdapter("SELECT DecorOrders.*, infiniu2_catalog.dbo.DecorConfig.MeasureID FROM DecorOrders" +
                        " INNER JOIN infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                        " WHERE DecorOrders.FactoryID=2" +
                        " AND MainOrderID IN (SELECT MainOrderID FROM MainOrders" +
                        " WHERE TPSProductionStatusID=1" +
                        " AND TPSStorageStatusID=1" +
                    " AND TPSExpeditionStatusID=1" +
                        " AND TPSDispatchStatusID=1)", ConnectionStrings.ZOVOrdersConnectionString))
                    {
                        DT.Clear();
                        DA.Fill(DT);

                        foreach (DataRow Row in DT.Rows)
                        {
                            if (Convert.ToInt32(Row["MeasureID"]) == 2)
                            {
                                //нет параметра "высота"
                                if (Row["Height"].ToString() == "-1")
                                    TPSPogon += Convert.ToDecimal(Row["Length"]) * Convert.ToDecimal(Row["Count"]) / 1000;
                                else
                                    TPSPogon += Convert.ToDecimal(Row["Height"]) * Convert.ToDecimal(Row["Count"]) / 1000;
                            }
                            else
                            {
                                TPSCount += Convert.ToDecimal(Row["Count"]);
                            }
                            AgreementTPSCost += Convert.ToDecimal(Row["Cost"]);
                        }
                    }
                }
                AgreementTPSCost = Decimal.Round(AgreementTPSCost, 1, MidpointRounding.AwayFromZero);
                TPSPogon = Decimal.Round(TPSPogon, 1, MidpointRounding.AwayFromZero);
                TPSCount = Decimal.Round(TPSCount, 1, MidpointRounding.AwayFromZero);
            }
        }

        public void GeneralOnProductionOrders(ref decimal ProfilSquare,
            ref decimal ProfilPogon, ref decimal ProfilCount, ref decimal TPSSquare,
            ref decimal TPSPogon, ref decimal TPSCount, ref decimal CurvedCount)
        {
            bool CheckZOV = GetZOV;
            string MegaOrderFilter = " AND ClientID = -1";

            OnProductionProfilCost = 0;
            OnProductionTPSCost = 0;

            if (ClientGroups.Count > 0)
            {
                MegaOrderFilter = " AND ClientID IN (SELECT ClientID FROM infiniu2_marketingreference.dbo.Clients" +
                    " WHERE ClientGroupID IN (" + string.Join(",", ClientGroups.OfType<int>().ToArray()) + "))";
            }

            using (DataTable DT = new DataTable())
            {
                //ФАСАДЫ НА СОГЛАСОВАНИИ
                //Профиль
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT Width, Cost, Square, Count FROM FrontsOrders" +
                    " WHERE FactoryID=1 AND MainOrderID IN (SELECT MainOrderID FROM MainOrders" +
                    " WHERE ProfilProductionStatusID=3" +
                    " AND ProfilStorageStatusID=1" +
                    " AND ProfilExpeditionStatusID=1" +
                    " AND ProfilDispatchStatusID=1" +
                    " AND MegaOrderID IN (SELECT MegaOrderID FROM MegaOrders" +
                    " WHERE AgreementStatusID=2" + MegaOrderFilter + "))", ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DT.Clear();
                    DA.Fill(DT);

                    foreach (DataRow Row in DT.Rows)
                    {
                        if (Convert.ToInt32(Row["Width"]) == -1)
                        {
                            CurvedCount += Convert.ToInt32(Row["Count"]);
                        }
                        else
                        {
                            ProfilSquare += Convert.ToDecimal(Row["Square"]);
                        }
                        OnProductionProfilCost += Convert.ToDecimal(Row["Cost"]);
                    }
                }
                if (CheckZOV)
                {
                    using (SqlDataAdapter DA = new SqlDataAdapter("SELECT Width, Cost, Square, Count FROM FrontsOrders" +
                        " WHERE FactoryID=1 AND MainOrderID IN (SELECT MainOrderID FROM MainOrders" +
                        " WHERE ProfilProductionStatusID=3" +
                        " AND ProfilStorageStatusID=1" +
                        " AND ProfilExpeditionStatusID=1" +
                        " AND ProfilDispatchStatusID=1)", ConnectionStrings.ZOVOrdersConnectionString))
                    {
                        DT.Clear();
                        DA.Fill(DT);

                        foreach (DataRow Row in DT.Rows)
                        {
                            if (Convert.ToInt32(Row["Width"]) == -1)
                            {
                                CurvedCount += Convert.ToInt32(Row["Count"]);
                            }
                            else
                            {
                                ProfilSquare += Convert.ToDecimal(Row["Square"]);
                            }
                            OnProductionProfilCost += Convert.ToDecimal(Row["Cost"]);
                        }
                    }
                }
                ProfilSquare = Decimal.Round(ProfilSquare, 1, MidpointRounding.AwayFromZero);
                OnProductionProfilCost = Decimal.Round(OnProductionProfilCost, 1, MidpointRounding.AwayFromZero);

                //ТПС
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT Width, Cost, Square, Count FROM FrontsOrders" +
                    " WHERE FactoryID=2 AND MainOrderID IN (SELECT MainOrderID FROM MainOrders" +
                    " WHERE TPSProductionStatusID=3" +
                    " AND TPSStorageStatusID=1" +
                    " AND TPSExpeditionStatusID=1" +
                    " AND TPSDispatchStatusID=1" +
                    " AND MegaOrderID IN (SELECT MegaOrderID FROM MegaOrders" +
                    " WHERE AgreementStatusID=2" + MegaOrderFilter + "))", ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DT.Clear();
                    DA.Fill(DT);

                    foreach (DataRow Row in DT.Rows)
                    {
                        if (Convert.ToInt32(Row["Width"]) == -1)
                        {
                            CurvedCount += Convert.ToInt32(Row["Count"]);
                        }
                        else
                        {
                            TPSSquare += Convert.ToDecimal(Row["Square"]);
                        }
                        OnProductionTPSCost += Convert.ToDecimal(Row["Cost"]);
                    }
                }

                if (CheckZOV)
                {
                    using (SqlDataAdapter DA = new SqlDataAdapter("SELECT Width, Cost, Square, Count FROM FrontsOrders" +
                        " WHERE FactoryID=2 AND MainOrderID IN (SELECT MainOrderID FROM MainOrders" +
                        " WHERE TPSProductionStatusID=3" +
                        " AND TPSStorageStatusID=1" +
                        " AND TPSExpeditionStatusID=1" +
                        " AND TPSDispatchStatusID=1)", ConnectionStrings.ZOVOrdersConnectionString))
                    {
                        DT.Clear();
                        DA.Fill(DT);

                        foreach (DataRow Row in DT.Rows)
                        {
                            if (Convert.ToInt32(Row["Width"]) == -1)
                            {
                                CurvedCount += Convert.ToInt32(Row["Count"]);
                            }
                            else
                            {
                                TPSSquare += Convert.ToDecimal(Row["Square"]);
                            }
                            OnProductionTPSCost += Convert.ToDecimal(Row["Cost"]);
                        }
                    }
                }
                TPSSquare = Decimal.Round(TPSSquare, 1, MidpointRounding.AwayFromZero);
                OnProductionTPSCost = Decimal.Round(OnProductionTPSCost, 1, MidpointRounding.AwayFromZero);

                //ДЕКОР НА СОГЛАСОВАНИИ
                //Профиль
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT DecorOrders.*, infiniu2_catalog.dbo.DecorConfig.MeasureID FROM DecorOrders" +
                    " INNER JOIN infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                    " WHERE DecorOrders.FactoryID=1" +
                    " AND MainOrderID IN (SELECT MainOrderID FROM MainOrders" +
                    " WHERE ProfilProductionStatusID=3" +
                    " AND ProfilStorageStatusID=1" +
                    " AND ProfilExpeditionStatusID=1" +
                    " AND ProfilDispatchStatusID=1" +
                    " AND MegaOrderID IN (SELECT MegaOrderID FROM MegaOrders" +
                    " WHERE AgreementStatusID=2" + MegaOrderFilter + "))", ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DT.Clear();
                    DA.Fill(DT);

                    foreach (DataRow Row in DT.Rows)
                    {
                        if (Convert.ToInt32(Row["MeasureID"]) == 2)
                        {
                            //нет параметра "высота"
                            if (Row["Height"].ToString() == "-1")
                                ProfilPogon += Convert.ToDecimal(Row["Length"]) * Convert.ToDecimal(Row["Count"]) / 1000;
                            else
                                ProfilPogon += Convert.ToDecimal(Row["Height"]) * Convert.ToDecimal(Row["Count"]) / 1000;
                        }
                        else
                        {
                            ProfilCount += Convert.ToDecimal(Row["Count"]);
                        }
                        OnProductionProfilCost += Convert.ToDecimal(Row["Cost"]);
                    }
                }

                if (CheckZOV)
                {
                    using (SqlDataAdapter DA = new SqlDataAdapter("SELECT DecorOrders.*, infiniu2_catalog.dbo.DecorConfig.MeasureID FROM DecorOrders" +
                        " INNER JOIN infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                        " WHERE DecorOrders.FactoryID=1" +
                        " AND MainOrderID IN (SELECT MainOrderID FROM MainOrders" +
                        " WHERE ProfilProductionStatusID=3" +
                        " AND ProfilStorageStatusID=1" +
                        " AND ProfilExpeditionStatusID=1" +
                        " AND ProfilDispatchStatusID=1)", ConnectionStrings.ZOVOrdersConnectionString))
                    {
                        DT.Clear();
                        DA.Fill(DT);

                        foreach (DataRow Row in DT.Rows)
                        {
                            if (Convert.ToInt32(Row["MeasureID"]) == 2)
                            {
                                //нет параметра "высота"
                                if (Row["Height"].ToString() == "-1")
                                    ProfilPogon += Convert.ToDecimal(Row["Length"]) * Convert.ToDecimal(Row["Count"]) / 1000;
                                else
                                    ProfilPogon += Convert.ToDecimal(Row["Height"]) * Convert.ToDecimal(Row["Count"]) / 1000;
                            }
                            else
                            {
                                ProfilCount += Convert.ToDecimal(Row["Count"]);
                            }
                            OnProductionProfilCost += Convert.ToDecimal(Row["Cost"]);
                        }
                    }
                }
                OnProductionProfilCost = Decimal.Round(OnProductionProfilCost, 1, MidpointRounding.AwayFromZero);
                ProfilPogon = Decimal.Round(ProfilPogon, 1, MidpointRounding.AwayFromZero);
                ProfilCount = Decimal.Round(ProfilCount, 1, MidpointRounding.AwayFromZero);

                //ТПС
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT DecorOrders.*, infiniu2_catalog.dbo.DecorConfig.MeasureID FROM DecorOrders" +
                    " INNER JOIN infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                    " WHERE DecorOrders.FactoryID=2" +
                    " AND MainOrderID IN (SELECT MainOrderID FROM MainOrders" +
                    " WHERE TPSProductionStatusID=3" +
                    " AND TPSStorageStatusID=1" +
                    " AND TPSExpeditionStatusID=1" +
                    " AND TPSDispatchStatusID=1" +
                    " AND MegaOrderID IN (SELECT MegaOrderID FROM MegaOrders" +
                    " WHERE AgreementStatusID=2" + MegaOrderFilter + "))", ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DT.Clear();
                    DA.Fill(DT);

                    foreach (DataRow Row in DT.Rows)
                    {
                        if (Convert.ToInt32(Row["MeasureID"]) == 2)
                        {
                            //нет параметра "высота"
                            if (Row["Height"].ToString() == "-1")
                                TPSPogon += Convert.ToDecimal(Row["Length"]) * Convert.ToDecimal(Row["Count"]) / 1000;
                            else
                                TPSPogon += Convert.ToDecimal(Row["Height"]) * Convert.ToDecimal(Row["Count"]) / 1000;
                        }
                        else
                        {
                            TPSCount += Convert.ToDecimal(Row["Count"]);
                        }
                        OnProductionTPSCost += Convert.ToDecimal(Row["Cost"]);
                    }
                }

                if (CheckZOV)
                {
                    using (SqlDataAdapter DA = new SqlDataAdapter("SELECT Cost, DecorOrders.Length, DecorOrders.Height, MeasureID, Count FROM DecorOrders" +
                        " INNER JOIN infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                        " WHERE DecorOrders.FactoryID=2" +
                        " AND MainOrderID IN (SELECT MainOrderID FROM MainOrders" +
                        " WHERE TPSProductionStatusID=3" +
                        " AND TPSStorageStatusID=1" +
                        " AND TPSExpeditionStatusID=1" +
                        " AND TPSDispatchStatusID=1)", ConnectionStrings.ZOVOrdersConnectionString))
                    {
                        DT.Clear();
                        DA.Fill(DT);

                        foreach (DataRow Row in DT.Rows)
                        {
                            if (Convert.ToInt32(Row["MeasureID"]) == 2)
                            {
                                //нет параметра "высота"
                                if (Row["Height"].ToString() == "-1")
                                    TPSPogon += Convert.ToDecimal(Row["Length"]) * Convert.ToDecimal(Row["Count"]) / 1000;
                                else
                                    TPSPogon += Convert.ToDecimal(Row["Height"]) * Convert.ToDecimal(Row["Count"]) / 1000;
                            }
                            else
                            {
                                TPSCount += Convert.ToDecimal(Row["Count"]);
                            }
                            OnProductionTPSCost += Convert.ToDecimal(Row["Cost"]);
                        }
                    }
                }
                OnProductionTPSCost = Decimal.Round(OnProductionTPSCost, 1, MidpointRounding.AwayFromZero);
                TPSPogon = Decimal.Round(TPSPogon, 1, MidpointRounding.AwayFromZero);
                TPSCount = Decimal.Round(TPSCount, 1, MidpointRounding.AwayFromZero);
            }
        }

        public void GeneralInProductionOrders(ref decimal ProfilSquare,
            ref decimal ProfilPogon, ref decimal ProfilCount, ref decimal TPSSquare,
            ref decimal TPSPogon, ref decimal TPSCount, ref decimal CurvedCount)
        {
            bool CheckZOV = GetZOV;
            string MegaOrderFilter = " AND MainOrderID = -1";

            InProductionProfilCost = 0;
            InProductionTPSCost = 0;

            if (ClientGroups.Count > 0)
            {
                MegaOrderFilter = " AND MainOrderID IN (SELECT MainOrderID FROM MainOrders WHERE MegaOrderID IN (SELECT MegaOrderID FROM MegaOrders" +
                       " WHERE ClientID IN (SELECT ClientID FROM infiniu2_marketingreference.dbo.Clients" +
                       " WHERE ClientGroupID IN (" + string.Join(",", ClientGroups.OfType<int>().ToArray()) + "))))";
            }

            using (DataTable DT = new DataTable())
            {
                //ФАСАДЫ В ПРОИЗВОДСТВЕ
                //Профиль
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM FrontsOrders" +
                    " WHERE FactoryID=1 AND FrontsOrdersID IN (SELECT OrderID FROM PackageDetails WHERE PackageID IN" +
                    " (SELECT PackageID FROM Packages WHERE (PackageStatusID=0) AND FactoryID=1 AND ProductType=0))" + MegaOrderFilter +
                    " UNION" +
                    " SELECT * FROM FrontsOrders" +
                    " WHERE FactoryID=1 AND MainOrderID IN (SELECT MainOrderID FROM MainOrders" +
                    " WHERE ProfilProductionStatusID = 2 AND ProfilStorageStatusID = 1 AND ProfilExpeditionStatusID=1 AND ProfilDispatchStatusID = 1)" + MegaOrderFilter,
                    ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DT.Clear();
                    DA.Fill(DT);

                    foreach (DataRow Row in DT.Rows)
                    {
                        if (Convert.ToInt32(Row["Width"]) == -1)
                        {
                            CurvedCount += Convert.ToInt32(Row["Count"]);
                        }
                        else
                        {
                            ProfilSquare += Convert.ToDecimal(Row["Square"]);
                        }
                        InProductionProfilCost += Convert.ToDecimal(Row["Cost"]);
                    }
                }
                if (CheckZOV)
                {
                    using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM FrontsOrders" +
                        " WHERE FactoryID=1 AND FrontsOrdersID IN (SELECT OrderID FROM PackageDetails WHERE PackageID IN" +
                        " (SELECT PackageID FROM Packages WHERE (PackageStatusID=0) AND FactoryID=1 AND ProductType=0))" +
                        " UNION" +
                        " SELECT * FROM FrontsOrders" +
                        " WHERE FactoryID=1 AND MainOrderID IN (SELECT MainOrderID FROM MainOrders" +
                        " WHERE ProfilProductionStatusID = 2 AND ProfilStorageStatusID = 1 AND ProfilExpeditionStatusID=1 AND ProfilDispatchStatusID = 1)",
                        ConnectionStrings.ZOVOrdersConnectionString))
                    {
                        DT.Clear();
                        DA.Fill(DT);

                        foreach (DataRow Row in DT.Rows)
                        {
                            if (Convert.ToInt32(Row["Width"]) == -1)
                            {
                                CurvedCount += Convert.ToInt32(Row["Count"]);
                            }
                            else
                            {
                                ProfilSquare += Convert.ToDecimal(Row["Square"]);
                            }
                            InProductionProfilCost += Convert.ToDecimal(Row["Cost"]);
                        }
                    }
                }
                ProfilSquare = Decimal.Round(ProfilSquare, 1, MidpointRounding.AwayFromZero);
                InProductionProfilCost = Decimal.Round(InProductionProfilCost, 1, MidpointRounding.AwayFromZero);

                //ТПС
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM FrontsOrders" +
                    " WHERE FactoryID=2 AND FrontsOrdersID IN (SELECT OrderID FROM PackageDetails WHERE PackageID IN" +
                    " (SELECT PackageID FROM Packages WHERE (PackageStatusID=0) AND FactoryID=2 AND ProductType=0))" + MegaOrderFilter +
                    " UNION" +
                    " SELECT * FROM FrontsOrders" +
                    " WHERE FactoryID=2 AND MainOrderID IN (SELECT MainOrderID FROM MainOrders" +
                    " WHERE TPSProductionStatusID = 2 AND TPSStorageStatusID = 1 AND TPSExpeditionStatusID=1 AND TPSDispatchStatusID = 1)" + MegaOrderFilter,
                    ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DT.Clear();
                    DA.Fill(DT);

                    foreach (DataRow Row in DT.Rows)
                    {
                        if (Convert.ToInt32(Row["Width"]) == -1)
                        {
                            CurvedCount += Convert.ToInt32(Row["Count"]);
                        }
                        else
                        {
                            TPSSquare += Convert.ToDecimal(Row["Square"]);
                        }
                        InProductionTPSCost += Convert.ToDecimal(Row["Cost"]);
                    }
                }

                if (CheckZOV)
                {
                    using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM FrontsOrders" +
                        " WHERE FactoryID=2 AND FrontsOrdersID IN (SELECT OrderID FROM PackageDetails WHERE PackageID IN" +
                        " (SELECT PackageID FROM Packages WHERE (PackageStatusID=0) AND FactoryID=2 AND ProductType=0))" +
                        " UNION" +
                        " SELECT * FROM FrontsOrders" +
                        " WHERE FactoryID=2 AND MainOrderID IN (SELECT MainOrderID FROM MainOrders" +
                        " WHERE TPSProductionStatusID = 2 AND TPSStorageStatusID = 1 AND TPSExpeditionStatusID=1 AND TPSDispatchStatusID = 1)",
                        ConnectionStrings.ZOVOrdersConnectionString))
                    {
                        DT.Clear();
                        DA.Fill(DT);

                        foreach (DataRow Row in DT.Rows)
                        {
                            if (Convert.ToInt32(Row["Width"]) == -1)
                            {
                                CurvedCount += Convert.ToInt32(Row["Count"]);
                            }
                            else
                            {
                                TPSSquare += Convert.ToDecimal(Row["Square"]);
                            }
                            InProductionTPSCost += Convert.ToDecimal(Row["Cost"]);
                        }
                    }
                }
                TPSSquare = Decimal.Round(TPSSquare, 1, MidpointRounding.AwayFromZero);
                InProductionTPSCost = Decimal.Round(InProductionTPSCost, 1, MidpointRounding.AwayFromZero);

                //ДЕКОР В ПРОИЗВОДСТВЕ
                //Профиль
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT DecorOrders.*, infiniu2_catalog.dbo.DecorConfig.MeasureID FROM DecorOrders" +
                    " INNER JOIN infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                    " WHERE DecorOrders.FactoryID=1 AND DecorOrderID IN (SELECT OrderID FROM PackageDetails WHERE PackageID IN" +
                    " (SELECT PackageID FROM Packages WHERE (PackageStatusID=0) AND FactoryID=1 AND ProductType=1))" + MegaOrderFilter +
                    " UNION" +
                    " SELECT DecorOrders.*, infiniu2_catalog.dbo.DecorConfig.MeasureID FROM DecorOrders" +
                    " INNER JOIN infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                    " WHERE DecorOrders.FactoryID=1 AND MainOrderID IN (SELECT MainOrderID FROM MainOrders" +
                    " WHERE ProfilProductionStatusID = 2 AND ProfilStorageStatusID = 1 AND ProfilExpeditionStatusID=1 AND ProfilDispatchStatusID = 1)" + MegaOrderFilter,
                    ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DT.Clear();
                    DA.Fill(DT);

                    foreach (DataRow Row in DT.Rows)
                    {
                        if (Convert.ToInt32(Row["MeasureID"]) == 2)
                        {
                            //нет параметра "высота"
                            if (Row["Height"].ToString() == "-1")
                                ProfilPogon += Convert.ToDecimal(Row["Length"]) * Convert.ToDecimal(Row["Count"]) / 1000;
                            else
                                ProfilPogon += Convert.ToDecimal(Row["Height"]) * Convert.ToDecimal(Row["Count"]) / 1000;
                        }
                        else
                        {
                            ProfilCount += Convert.ToDecimal(Row["Count"]);
                        }
                        InProductionProfilCost += Convert.ToDecimal(Row["Cost"]);
                    }
                }

                if (CheckZOV)
                {
                    using (SqlDataAdapter DA = new SqlDataAdapter("SELECT DecorOrders.*, infiniu2_catalog.dbo.DecorConfig.MeasureID FROM DecorOrders" +
                        " INNER JOIN infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                        " WHERE DecorOrders.FactoryID=1 AND DecorOrderID IN (SELECT OrderID FROM PackageDetails WHERE PackageID IN" +
                        " (SELECT PackageID FROM Packages WHERE (PackageStatusID=0) AND FactoryID=1 AND ProductType=1))" +
                        " UNION" +
                        " SELECT DecorOrders.*, infiniu2_catalog.dbo.DecorConfig.MeasureID FROM DecorOrders" +
                        " INNER JOIN infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                        " WHERE DecorOrders.FactoryID=1 AND MainOrderID IN (SELECT MainOrderID FROM MainOrders" +
                        " WHERE ProfilProductionStatusID = 2 AND ProfilStorageStatusID = 1 AND ProfilExpeditionStatusID=1 AND ProfilDispatchStatusID = 1)",
                        ConnectionStrings.ZOVOrdersConnectionString))
                    {
                        DT.Clear();
                        DA.Fill(DT);

                        foreach (DataRow Row in DT.Rows)
                        {
                            if (Convert.ToInt32(Row["MeasureID"]) == 2)
                            {
                                //нет параметра "высота"
                                if (Row["Height"].ToString() == "-1")
                                    ProfilPogon += Convert.ToDecimal(Row["Length"]) * Convert.ToDecimal(Row["Count"]) / 1000;
                                else
                                    ProfilPogon += Convert.ToDecimal(Row["Height"]) * Convert.ToDecimal(Row["Count"]) / 1000;
                            }
                            else
                            {
                                ProfilCount += Convert.ToDecimal(Row["Count"]);
                            }
                            InProductionProfilCost += Convert.ToDecimal(Row["Cost"]);
                        }
                    }
                }
                InProductionProfilCost = Decimal.Round(InProductionProfilCost, 1, MidpointRounding.AwayFromZero);
                ProfilPogon = Decimal.Round(ProfilPogon, 1, MidpointRounding.AwayFromZero);
                ProfilCount = Decimal.Round(ProfilCount, 1, MidpointRounding.AwayFromZero);

                //ТПС
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT DecorOrders.*, infiniu2_catalog.dbo.DecorConfig.MeasureID FROM DecorOrders" +
                    " INNER JOIN infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                    " WHERE DecorOrders.FactoryID=2 AND DecorOrderID IN (SELECT OrderID FROM PackageDetails WHERE PackageID IN" +
                    " (SELECT PackageID FROM Packages WHERE (PackageStatusID=0) AND FactoryID=2 AND ProductType=1))" + MegaOrderFilter +
                    " UNION" +
                    " SELECT DecorOrders.*, infiniu2_catalog.dbo.DecorConfig.MeasureID FROM DecorOrders" +
                    " INNER JOIN infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                    " WHERE DecorOrders.FactoryID=2 AND MainOrderID IN (SELECT MainOrderID FROM MainOrders" +
                    " WHERE TPSProductionStatusID = 2 AND TPSStorageStatusID = 1 AND TPSExpeditionStatusID=1 AND TPSDispatchStatusID = 1)" + MegaOrderFilter,
                    ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DT.Clear();
                    DA.Fill(DT);

                    foreach (DataRow Row in DT.Rows)
                    {
                        if (Convert.ToInt32(Row["MeasureID"]) == 2)
                        {
                            //нет параметра "высота"
                            if (Row["Height"].ToString() == "-1")
                                TPSPogon += Convert.ToDecimal(Row["Length"]) * Convert.ToDecimal(Row["Count"]) / 1000;
                            else
                                TPSPogon += Convert.ToDecimal(Row["Height"]) * Convert.ToDecimal(Row["Count"]) / 1000;
                        }
                        else
                        {
                            TPSCount += Convert.ToDecimal(Row["Count"]);
                        }
                        InProductionTPSCost += Convert.ToDecimal(Row["Cost"]);
                    }
                }

                if (CheckZOV)
                {
                    using (SqlDataAdapter DA = new SqlDataAdapter("SELECT DecorOrders.*, infiniu2_catalog.dbo.DecorConfig.MeasureID FROM DecorOrders" +
                        " INNER JOIN infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                        " WHERE DecorOrders.FactoryID=2 AND DecorOrderID IN (SELECT OrderID FROM PackageDetails WHERE PackageID IN" +
                        " (SELECT PackageID FROM Packages WHERE (PackageStatusID=0) AND FactoryID=2 AND ProductType=1))" +
                        " UNION" +
                        " SELECT DecorOrders.*, infiniu2_catalog.dbo.DecorConfig.MeasureID FROM DecorOrders" +
                        " INNER JOIN infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                        " WHERE DecorOrders.FactoryID=2 AND MainOrderID IN (SELECT MainOrderID FROM MainOrders" +
                        " WHERE TPSProductionStatusID = 2 AND TPSStorageStatusID = 1 AND TPSExpeditionStatusID=1 AND TPSDispatchStatusID = 1)",
                        ConnectionStrings.ZOVOrdersConnectionString))
                    {
                        DT.Clear();
                        DA.Fill(DT);

                        foreach (DataRow Row in DT.Rows)
                        {
                            if (Convert.ToInt32(Row["MeasureID"]) == 2)
                            {
                                //нет параметра "высота"
                                if (Row["Height"].ToString() == "-1")
                                    TPSPogon += Convert.ToDecimal(Row["Length"]) * Convert.ToDecimal(Row["Count"]) / 1000;
                                else
                                    TPSPogon += Convert.ToDecimal(Row["Height"]) * Convert.ToDecimal(Row["Count"]) / 1000;
                            }
                            else
                            {
                                TPSCount += Convert.ToDecimal(Row["Count"]);
                            }
                            InProductionTPSCost += Convert.ToDecimal(Row["Cost"]);
                        }
                    }
                }
                InProductionTPSCost = Decimal.Round(InProductionTPSCost, 1, MidpointRounding.AwayFromZero);
                TPSPogon = Decimal.Round(TPSPogon, 1, MidpointRounding.AwayFromZero);
                TPSCount = Decimal.Round(TPSCount, 1, MidpointRounding.AwayFromZero);
            }
        }

        public void FillTables(DateTime date1, ref decimal ProfilSquare,
            ref decimal ProfilPogon, ref decimal ProfilCount, ref decimal TPSSquare,
            ref decimal TPSPogon, ref decimal TPSCount, ref decimal CurvedCount)
        {
            string MClientFilter = string.Empty;
            if (ClientGroups.Count > 0)
            {
                MClientFilter = " AND MegaOrders.ClientID IN" +
                    " (SELECT ClientID FROM infiniu2_marketingreference.dbo.Clients" +
                    " WHERE ClientGroupID IN (" + string.Join(",", ClientGroups.OfType<Int32>().ToArray()) + "))";
            }
            if (ClientGroups.Count < 1)
                MClientFilter = " AND MegaOrders.ClientID = -1";

            string Filter = " AND Packages.ProductType = 0 AND Packages.PackingDateTime <  '" + date1.ToString("yyyy-MM-dd") +
                " 23:59:59' AND (DispatchDateTime IS NULL OR Packages.DispatchDateTime >= '" + date1.ToString("yyyy-MM-dd") + " 23:59:59')";

            string SelectCommand = @"SELECT infiniu2_marketingreference.dbo.Clients.ClientName, dbo.MegaOrders.OrderNumber, Packages.PackingDateTime, dbo.PackageDetails.PackageID, dbo.PackageDetails.PackNumber, dbo.FrontsOrders.FrontsOrdersID,
                        dbo.PackageDetails.Count, FrontsOrders.Cost*dbo.PackageDetails.Count/FrontsOrders.Count as Cost, FrontsOrders.Square*dbo.PackageDetails.Count/FrontsOrders.Count as Square, FrontsOrders.FactoryID, FrontsOrders.Width,
                         infiniu2_catalog.dbo.FrontsConfig.AccountingName, infiniu2_catalog.dbo.FrontsConfig.InvNumber, infiniu2_catalog.dbo.Measures.Measure, Fronts.TechStoreName, Colors.TechStoreName AS Expr35,
                         InsetTypes.TechStoreName AS Expr36, InsetColors.TechStoreName AS Expr37, TechnoInsetTypes.TechStoreName AS Expr38, TechnoInsetColors.TechStoreName AS Expr1
FROM dbo.PackageDetails INNER JOIN
                         dbo.FrontsOrders ON dbo.PackageDetails.OrderID = dbo.FrontsOrders.FrontsOrdersID INNER JOIN
                         dbo.Packages ON dbo.PackageDetails.PackageID = dbo.Packages.PackageID " + Filter + @" LEFT OUTER JOIN
                         infiniu2_catalog.dbo.TechStore AS Fronts ON dbo.FrontsOrders.FrontID = Fronts.TechStoreID LEFT OUTER JOIN
                         infiniu2_catalog.dbo.TechStore AS Colors ON dbo.FrontsOrders.ColorID = Colors.TechStoreID LEFT OUTER JOIN
                         infiniu2_catalog.dbo.TechStore AS InsetTypes ON dbo.FrontsOrders.InsetTypeID = InsetTypes.TechStoreID LEFT OUTER JOIN
                         infiniu2_catalog.dbo.TechStore AS InsetColors ON dbo.FrontsOrders.InsetColorID = InsetColors.TechStoreID LEFT OUTER JOIN
                         infiniu2_catalog.dbo.TechStore AS TechnoInsetTypes ON dbo.FrontsOrders.TechnoInsetTypeID = TechnoInsetTypes.TechStoreID LEFT OUTER JOIN
                         infiniu2_catalog.dbo.TechStore AS TechnoInsetColors ON dbo.FrontsOrders.TechnoInsetColorID = TechnoInsetColors.TechStoreID INNER JOIN
                         infiniu2_catalog.dbo.FrontsConfig ON dbo.FrontsOrders.FrontConfigID = infiniu2_catalog.dbo.FrontsConfig.FrontConfigID INNER JOIN
                         infiniu2_catalog.dbo.Measures ON infiniu2_catalog.dbo.FrontsConfig.MeasureID = infiniu2_catalog.dbo.Measures.MeasureID INNER JOIN
                         dbo.MainOrders ON dbo.FrontsOrders.MainOrderID = dbo.MainOrders.MainOrderID INNER JOIN
                         dbo.MegaOrders ON dbo.MainOrders.MegaOrderID = dbo.MegaOrders.MegaOrderID " + MClientFilter + @" INNER JOIN
                         infiniu2_marketingreference.dbo.Clients ON dbo.MegaOrders.ClientID = infiniu2_marketingreference.dbo.Clients.ClientID
ORDER BY infiniu2_marketingreference.dbo.Clients.ClientName, dbo.MegaOrders.OrderNumber, dbo.PackageDetails.PackageID";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    DA.Fill(DT);

                    foreach (DataRow Row in DT.Rows)
                    {
                        if (Convert.ToInt32(Row["FactoryID"]) == 1)
                        {
                            if (Convert.ToInt32(Row["Width"]) == -1)
                            {
                                CurvedCount += Convert.ToInt32(Row["Count"]);
                            }
                            else
                            {
                                ProfilSquare += Convert.ToDecimal(Row["Square"]);
                            }
                            OnStorageProfilCost += Convert.ToDecimal(Row["Cost"]);
                        }
                        if (Convert.ToInt32(Row["FactoryID"]) == 2)
                        {
                            if (Convert.ToInt32(Row["Width"]) == -1)
                            {
                                CurvedCount += Convert.ToInt32(Row["Count"]);
                            }
                            else
                            {
                                TPSSquare += Convert.ToDecimal(Row["Square"]);
                            }
                            OnStorageTPSCost += Convert.ToDecimal(Row["Cost"]);
                        }
                    }
                    ProfilSquare = Decimal.Round(ProfilSquare, 1, MidpointRounding.AwayFromZero);
                    OnStorageProfilCost = Decimal.Round(OnStorageProfilCost, 1, MidpointRounding.AwayFromZero);
                    TPSSquare = Decimal.Round(TPSSquare, 1, MidpointRounding.AwayFromZero);
                    OnStorageTPSCost = Decimal.Round(OnStorageTPSCost, 1, MidpointRounding.AwayFromZero);
                }
            }
            Filter = " AND Packages.ProductType = 1 AND Packages.PackingDateTime <  '" + date1.ToString("yyyy-MM-dd") +
                " 23:59:59' AND (DispatchDateTime IS NULL OR Packages.DispatchDateTime >= '" + date1.ToString("yyyy-MM-dd") + " 23:59:59')";
            SelectCommand = @"SELECT infiniu2_marketingreference.dbo.Clients.ClientName, dbo.MegaOrders.OrderNumber, Packages.PackingDateTime, dbo.PackageDetails.PackageID, dbo.PackageDetails.PackNumber, dbo.DecorOrders.DecorOrderID,
                         DecorOrders.FactoryID, infiniu2_catalog.dbo.DecorConfig.MeasureID, infiniu2_catalog.dbo.DecorConfig.AccountingName, infiniu2_catalog.dbo.DecorConfig.InvNumber, infiniu2_catalog.dbo.Measures.Measure, Fronts.TechStoreName, Colors.TechStoreName AS Expr35, dbo.DecorOrders.Length, dbo.DecorOrders.Height, dbo.DecorOrders.Width,
                        dbo.PackageDetails.Count, dbo.DecorOrders.Cost*dbo.PackageDetails.Count/dbo.DecorOrders.Count AS Cost
FROM dbo.PackageDetails INNER JOIN
                         dbo.DecorOrders ON dbo.PackageDetails.OrderID = dbo.DecorOrders.DecorOrderID INNER JOIN
                         dbo.Packages ON dbo.PackageDetails.PackageID = dbo.Packages.PackageID " + Filter + @" LEFT OUTER JOIN
                         infiniu2_catalog.dbo.TechStore AS Fronts ON dbo.DecorOrders.DecorID = Fronts.TechStoreID LEFT OUTER JOIN
                         infiniu2_catalog.dbo.TechStore AS Colors ON dbo.DecorOrders.ColorID = Colors.TechStoreID INNER JOIN
                         infiniu2_catalog.dbo.DecorConfig ON dbo.DecorOrders.DecorConfigID = infiniu2_catalog.dbo.DecorConfig.DecorConfigID INNER JOIN
                        infiniu2_catalog.dbo.Measures ON infiniu2_catalog.dbo.DecorConfig.MeasureID = infiniu2_catalog.dbo.Measures.MeasureID INNER JOIN
                         dbo.MainOrders ON dbo.DecorOrders.MainOrderID = dbo.MainOrders.MainOrderID INNER JOIN
                         dbo.MegaOrders ON dbo.MainOrders.MegaOrderID = dbo.MegaOrders.MegaOrderID " + MClientFilter + @" INNER JOIN
                         infiniu2_marketingreference.dbo.Clients ON dbo.MegaOrders.ClientID = infiniu2_marketingreference.dbo.Clients.ClientID
ORDER BY infiniu2_marketingreference.dbo.Clients.ClientName, dbo.MegaOrders.OrderNumber, dbo.PackageDetails.PackageID";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    DA.Fill(DT);
                    foreach (DataRow Row in DT.Rows)
                    {
                        if (Convert.ToInt32(Row["FactoryID"]) == 1)
                        {
                            if (Convert.ToInt32(Row["MeasureID"]) == 2)
                            {
                                //нет параметра "высота"
                                if (Row["Height"].ToString() == "-1")
                                    ProfilPogon += Convert.ToDecimal(Row["Length"]) * Convert.ToDecimal(Row["Count"]) / 1000;
                                else
                                    ProfilPogon += Convert.ToDecimal(Row["Height"]) * Convert.ToDecimal(Row["Count"]) / 1000;
                            }
                            else
                            {
                                ProfilCount += Convert.ToDecimal(Row["Count"]);
                            }
                            OnStorageProfilCost += Convert.ToDecimal(Row["Cost"]);
                        }
                        if (Convert.ToInt32(Row["FactoryID"]) == 2)
                        {
                            if (Convert.ToInt32(Row["MeasureID"]) == 2)
                            {
                                //нет параметра "высота"
                                if (Row["Height"].ToString() == "-1")
                                    TPSPogon += Convert.ToDecimal(Row["Length"]) * Convert.ToDecimal(Row["Count"]) / 1000;
                                else
                                    TPSPogon += Convert.ToDecimal(Row["Height"]) * Convert.ToDecimal(Row["Count"]) / 1000;
                            }
                            else
                            {
                                TPSCount += Convert.ToDecimal(Row["Count"]);
                            }
                            OnStorageTPSCost += Convert.ToDecimal(Row["Cost"]);
                        }
                    }
                    OnStorageProfilCost = Decimal.Round(OnStorageProfilCost, 1, MidpointRounding.AwayFromZero);
                    ProfilPogon = Decimal.Round(ProfilPogon, 1, MidpointRounding.AwayFromZero);
                    ProfilCount = Decimal.Round(ProfilCount, 1, MidpointRounding.AwayFromZero);
                    OnStorageTPSCost = Decimal.Round(OnStorageTPSCost, 1, MidpointRounding.AwayFromZero);
                    TPSPogon = Decimal.Round(TPSPogon, 1, MidpointRounding.AwayFromZero);
                    TPSCount = Decimal.Round(TPSCount, 1, MidpointRounding.AwayFromZero);
                }
            }
        }

        //public void GeneralOnStorageOrders(ref decimal ProfilSquare,
        //    ref decimal ProfilPogon, ref decimal ProfilCount, ref decimal TPSSquare,
        //    ref decimal TPSPogon, ref decimal TPSCount, ref decimal CurvedCount)
        //{
        //    string MegaOrderFilter = " AND MainOrderID = -1";

        //    OnStorageProfilCost = 0;
        //    OnStorageTPSCost = 0;

        //    if (ClientGroups.Count > 0)
        //    {
        //        MegaOrderFilter = " AND MainOrderID IN (SELECT MainOrderID FROM MainOrders WHERE MegaOrderID IN (SELECT MegaOrderID FROM MegaOrders" +
        //            " WHERE ClientID IN (SELECT ClientID FROM infiniu2_marketingreference.dbo.Clients" +
        //            " WHERE ClientGroupID IN (" + string.Join(",", ClientGroups.OfType<int>().ToArray()) + "))))";
        //    }

        //    using (DataTable DT = new DataTable())
        //    {
        //        //ФАСАДЫ НА ЭКСПЕДИЦИИ
        //        //Профиль
        //        using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM FrontsOrders" +
        //            " WHERE FactoryID=1 AND FrontsOrdersID IN (SELECT OrderID FROM PackageDetails WHERE PackageID IN" +
        //            " (SELECT PackageID FROM Packages WHERE (PackageStatusID=1 OR PackageStatusID=2) AND FactoryID=1 AND ProductType=0))" + MegaOrderFilter +
        //            " UNION" +
        //            " SELECT * FROM FrontsOrders" +
        //            " WHERE FactoryID=1 AND MainOrderID IN (SELECT MainOrderID FROM MainOrders" +
        //            " WHERE ProfilProductionStatusID = 1 AND ProfilStorageStatusID = 2 AND ProfilStorageStatusID=1 AND ProfilDispatchStatusID = 1)" + MegaOrderFilter,
        //            ConnectionStrings.MarketingOrdersConnectionString))
        //        {
        //            DT.Clear();
        //            DA.Fill(DT);

        //            foreach (DataRow Row in DT.Rows)
        //            {
        //                if (Convert.ToInt32(Row["Width"]) == -1)
        //                {
        //                    CurvedCount += Convert.ToInt32(Row["Count"]);
        //                }
        //                else
        //                {
        //                    ProfilSquare += Convert.ToDecimal(Row["Square"]);
        //                }
        //                OnStorageProfilCost += Convert.ToDecimal(Row["Cost"]);
        //            }
        //        }
        //        ProfilSquare = Decimal.Round(ProfilSquare, 1, MidpointRounding.AwayFromZero);
        //        OnStorageProfilCost = Decimal.Round(OnStorageProfilCost, 1, MidpointRounding.AwayFromZero);

        //        //foreach (DataRow item in TempDecorDataTable.Rows)
        //        //{
        //        //    DecorOrdersDataTable.ImportRow(item);
        //        //}

        //        //ТПС
        //        using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM FrontsOrders" +
        //            " WHERE FactoryID=2 AND FrontsOrdersID IN (SELECT OrderID FROM PackageDetails WHERE PackageID IN" +
        //            " (SELECT PackageID FROM Packages WHERE (PackageStatusID=1 OR PackageStatusID=2) AND FactoryID=2 AND ProductType=0))" + MegaOrderFilter +
        //            " UNION" +
        //            " SELECT * FROM FrontsOrders" +
        //            " WHERE FactoryID=2 AND MainOrderID IN (SELECT MainOrderID FROM MainOrders" +
        //            " WHERE TPSProductionStatusID = 1 AND TPSStorageStatusID = 2 AND TPSStorageStatusID=1 AND TPSDispatchStatusID = 1)" + MegaOrderFilter,
        //            ConnectionStrings.MarketingOrdersConnectionString))
        //        {
        //            DT.Clear();
        //            DA.Fill(DT);

        //            foreach (DataRow Row in DT.Rows)
        //            {
        //                if (Convert.ToInt32(Row["Width"]) == -1)
        //                {
        //                    CurvedCount += Convert.ToInt32(Row["Count"]);
        //                }
        //                else
        //                {
        //                    TPSSquare += Convert.ToDecimal(Row["Square"]);
        //                }
        //                OnStorageTPSCost += Convert.ToDecimal(Row["Cost"]);
        //            }
        //        }
        //        TPSSquare = Decimal.Round(TPSSquare, 1, MidpointRounding.AwayFromZero);
        //        OnStorageTPSCost = Decimal.Round(OnStorageTPSCost, 1, MidpointRounding.AwayFromZero);

        //        //ДЕКОР НА ЭКСПЕДИЦИИ
        //        //Профиль
        //        using (SqlDataAdapter DA = new SqlDataAdapter("SELECT DecorOrders.*, infiniu2_catalog.dbo.DecorConfig.MeasureID FROM DecorOrders" +
        //            " INNER JOIN infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
        //            " WHERE DecorOrders.FactoryID=1 AND DecorOrderID IN (SELECT OrderID FROM PackageDetails WHERE PackageID IN" +
        //            " (SELECT PackageID FROM Packages WHERE (PackageStatusID=1 OR PackageStatusID=2) AND FactoryID=1 AND ProductType=1))" + MegaOrderFilter +
        //            " UNION" +
        //            " SELECT DecorOrders.*, infiniu2_catalog.dbo.DecorConfig.MeasureID FROM DecorOrders" +
        //            " INNER JOIN infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
        //            " WHERE DecorOrders.FactoryID=1 AND MainOrderID IN (SELECT MainOrderID FROM MainOrders" +
        //            " WHERE ProfilProductionStatusID = 1 AND ProfilStorageStatusID = 2 AND ProfilStorageStatusID=1 AND ProfilDispatchStatusID = 1)" + MegaOrderFilter,
        //            ConnectionStrings.MarketingOrdersConnectionString))
        //        {
        //            DT.Clear();
        //            DA.Fill(DT);

        //            foreach (DataRow Row in DT.Rows)
        //            {
        //                if (Convert.ToInt32(Row["MeasureID"]) == 2)
        //                {
        //                    //нет параметра "высота"
        //                    if (Row["Height"].ToString() == "-1")
        //                        ProfilPogon += Convert.ToDecimal(Row["Length"]) * Convert.ToDecimal(Row["Count"]) / 1000;
        //                    else
        //                        ProfilPogon += Convert.ToDecimal(Row["Height"]) * Convert.ToDecimal(Row["Count"]) / 1000;
        //                }
        //                else
        //                {
        //                    ProfilCount += Convert.ToDecimal(Row["Count"]);
        //                }
        //                OnStorageProfilCost += Convert.ToDecimal(Row["Cost"]);
        //            }
        //        }
        //        OnStorageProfilCost = Decimal.Round(OnStorageProfilCost, 1, MidpointRounding.AwayFromZero);
        //        ProfilPogon = Decimal.Round(ProfilPogon, 1, MidpointRounding.AwayFromZero);
        //        ProfilCount = Decimal.Round(ProfilCount, 1, MidpointRounding.AwayFromZero);

        //        //ТПС
        //        using (SqlDataAdapter DA = new SqlDataAdapter("SELECT DecorOrders.*, infiniu2_catalog.dbo.DecorConfig.MeasureID FROM DecorOrders" +
        //            " INNER JOIN infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
        //            " WHERE DecorOrders.FactoryID=2 AND DecorOrderID IN (SELECT OrderID FROM PackageDetails WHERE PackageID IN" +
        //            " (SELECT PackageID FROM Packages WHERE (PackageStatusID=1 OR PackageStatusID=2) AND FactoryID=2 AND ProductType=1))" + MegaOrderFilter +
        //            " UNION" +
        //            " SELECT DecorOrders.*, infiniu2_catalog.dbo.DecorConfig.MeasureID FROM DecorOrders" +
        //            " INNER JOIN infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
        //            " WHERE DecorOrders.FactoryID=2 AND MainOrderID IN (SELECT MainOrderID FROM MainOrders" +
        //            " WHERE TPSProductionStatusID = 1 AND TPSStorageStatusID = 2 AND TPSStorageStatusID=1 AND TPSDispatchStatusID = 1)" + MegaOrderFilter,
        //            ConnectionStrings.MarketingOrdersConnectionString))
        //        {
        //            DT.Clear();
        //            DA.Fill(DT);

        //            foreach (DataRow Row in DT.Rows)
        //            {
        //                if (Convert.ToInt32(Row["MeasureID"]) == 2)
        //                {
        //                    //нет параметра "высота"
        //                    if (Row["Height"].ToString() == "-1")
        //                        TPSPogon += Convert.ToDecimal(Row["Length"]) * Convert.ToDecimal(Row["Count"]) / 1000;
        //                    else
        //                        TPSPogon += Convert.ToDecimal(Row["Height"]) * Convert.ToDecimal(Row["Count"]) / 1000;
        //                }
        //                else
        //                {
        //                    TPSCount += Convert.ToDecimal(Row["Count"]);
        //                }
        //                OnStorageTPSCost += Convert.ToDecimal(Row["Cost"]);
        //            }
        //        }
        //        OnStorageTPSCost = Decimal.Round(OnStorageTPSCost, 1, MidpointRounding.AwayFromZero);
        //        TPSPogon = Decimal.Round(TPSPogon, 1, MidpointRounding.AwayFromZero);
        //        TPSCount = Decimal.Round(TPSCount, 1, MidpointRounding.AwayFromZero);
        //    }
        //}

        public void GeneralOnExpeditionOrders(ref decimal ProfilSquare,
            ref decimal ProfilPogon, ref decimal ProfilCount, ref decimal TPSSquare,
            ref decimal TPSPogon, ref decimal TPSCount, ref decimal CurvedCount)
        {
            bool CheckZOV = GetZOV;
            string MegaOrderFilter = " AND MainOrderID = -1";

            OnExpeditionProfilCost = 0;
            OnExpeditionTPSCost = 0;

            if (ClientGroups.Count > 0)
            {
                MegaOrderFilter = " AND MainOrderID IN (SELECT MainOrderID FROM MainOrders WHERE MegaOrderID IN (SELECT MegaOrderID FROM MegaOrders" +
                    " WHERE ClientID IN (SELECT ClientID FROM infiniu2_marketingreference.dbo.Clients" +
                    " WHERE ClientGroupID IN (" + string.Join(",", ClientGroups.OfType<int>().ToArray()) + "))))";
            }

            using (DataTable DT = new DataTable())
            {
                //ФАСАДЫ НА ЭКСПЕДИЦИИ
                //Профиль
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM FrontsOrders" +
                    " WHERE FactoryID=1 AND FrontsOrdersID IN (SELECT OrderID FROM PackageDetails WHERE PackageID IN" +
                    " (SELECT PackageID FROM Packages WHERE (PackageStatusID=4) AND FactoryID=1 AND ProductType=0))" + MegaOrderFilter +
                    " UNION" +
                    " SELECT * FROM FrontsOrders" +
                    " WHERE FactoryID=1 AND MainOrderID IN (SELECT MainOrderID FROM MainOrders" +
                    " WHERE ProfilProductionStatusID = 1 AND ProfilStorageStatusID = 1 AND ProfilExpeditionStatusID=2 AND ProfilDispatchStatusID = 1)" + MegaOrderFilter,
                    ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DT.Clear();
                    DA.Fill(DT);

                    foreach (DataRow Row in DT.Rows)
                    {
                        if (Convert.ToInt32(Row["Width"]) == -1)
                        {
                            CurvedCount += Convert.ToInt32(Row["Count"]);
                        }
                        else
                        {
                            ProfilSquare += Convert.ToDecimal(Row["Square"]);
                        }
                        OnExpeditionProfilCost += Convert.ToDecimal(Row["Cost"]);
                    }
                }
                if (CheckZOV)
                {
                    using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM FrontsOrders" +
                        " WHERE FactoryID=1 AND FrontsOrdersID IN (SELECT OrderID FROM PackageDetails WHERE PackageID IN" +
                        " (SELECT PackageID FROM Packages WHERE (PackageStatusID=4) AND FactoryID=1 AND ProductType=0))" +
                        " UNION" +
                        " SELECT * FROM FrontsOrders" +
                        " WHERE FactoryID=1 AND MainOrderID IN (SELECT MainOrderID FROM MainOrders" +
                        " WHERE ProfilProductionStatusID = 1 AND ProfilStorageStatusID = 1 AND ProfilExpeditionStatusID=2 AND ProfilDispatchStatusID = 1)",
                        ConnectionStrings.ZOVOrdersConnectionString))
                    {
                        DT.Clear();
                        DA.Fill(DT);

                        foreach (DataRow Row in DT.Rows)
                        {
                            if (Convert.ToInt32(Row["Width"]) == -1)
                            {
                                CurvedCount += Convert.ToInt32(Row["Count"]);
                            }
                            else
                            {
                                ProfilSquare += Convert.ToDecimal(Row["Square"]);
                            }
                            OnExpeditionProfilCost += Convert.ToDecimal(Row["Cost"]);
                        }
                    }
                }
                ProfilSquare = Decimal.Round(ProfilSquare, 1, MidpointRounding.AwayFromZero);
                OnExpeditionProfilCost = Decimal.Round(OnExpeditionProfilCost, 1, MidpointRounding.AwayFromZero);

                //foreach (DataRow item in TempDecorDataTable.Rows)
                //{
                //    DecorOrdersDataTable.ImportRow(item);
                //}

                //ТПС
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM FrontsOrders" +
                    " WHERE FactoryID=2 AND FrontsOrdersID IN (SELECT OrderID FROM PackageDetails WHERE PackageID IN" +
                    " (SELECT PackageID FROM Packages WHERE (PackageStatusID=4) AND FactoryID=2 AND ProductType=0))" + MegaOrderFilter +
                    " UNION" +
                    " SELECT * FROM FrontsOrders" +
                    " WHERE FactoryID=2 AND MainOrderID IN (SELECT MainOrderID FROM MainOrders" +
                    " WHERE TPSProductionStatusID = 1 AND TPSStorageStatusID = 1 AND TPSExpeditionStatusID=2 AND TPSDispatchStatusID = 1)" + MegaOrderFilter,
                    ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DT.Clear();
                    DA.Fill(DT);

                    foreach (DataRow Row in DT.Rows)
                    {
                        if (Convert.ToInt32(Row["Width"]) == -1)
                        {
                            CurvedCount += Convert.ToInt32(Row["Count"]);
                        }
                        else
                        {
                            TPSSquare += Convert.ToDecimal(Row["Square"]);
                        }
                        OnExpeditionTPSCost += Convert.ToDecimal(Row["Cost"]);
                    }
                }
                if (CheckZOV)
                {
                    using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM FrontsOrders" +
                        " WHERE FactoryID=2 AND FrontsOrdersID IN (SELECT OrderID FROM PackageDetails WHERE PackageID IN" +
                        " (SELECT PackageID FROM Packages WHERE (PackageStatusID=4) AND FactoryID=2 AND ProductType=0))" +
                        " UNION" +
                        " SELECT * FROM FrontsOrders" +
                        " WHERE FactoryID=2 AND MainOrderID IN (SELECT MainOrderID FROM MainOrders" +
                        " WHERE TPSProductionStatusID = 1 AND TPSStorageStatusID = 1 AND TPSExpeditionStatusID=2 AND TPSDispatchStatusID = 1)",
                        ConnectionStrings.ZOVOrdersConnectionString))
                    {
                        DT.Clear();
                        DA.Fill(DT);

                        foreach (DataRow Row in DT.Rows)
                        {
                            if (Convert.ToInt32(Row["Width"]) == -1)
                            {
                                CurvedCount += Convert.ToInt32(Row["Count"]);
                            }
                            else
                            {
                                TPSSquare += Convert.ToDecimal(Row["Square"]);
                            }
                            OnExpeditionTPSCost += Convert.ToDecimal(Row["Cost"]);
                        }
                    }
                }
                TPSSquare = Decimal.Round(TPSSquare, 1, MidpointRounding.AwayFromZero);
                OnExpeditionTPSCost = Decimal.Round(OnExpeditionTPSCost, 1, MidpointRounding.AwayFromZero);

                //ДЕКОР НА ЭКСПЕДИЦИИ
                //Профиль
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT DecorOrders.*, infiniu2_catalog.dbo.DecorConfig.MeasureID FROM DecorOrders" +
                    " INNER JOIN infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                    " WHERE DecorOrders.FactoryID=1 AND DecorOrderID IN (SELECT OrderID FROM PackageDetails WHERE PackageID IN" +
                    " (SELECT PackageID FROM Packages WHERE (PackageStatusID=4) AND FactoryID=1 AND ProductType=1))" + MegaOrderFilter +
                    " UNION" +
                    " SELECT DecorOrders.*, infiniu2_catalog.dbo.DecorConfig.MeasureID FROM DecorOrders" +
                    " INNER JOIN infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                    " WHERE DecorOrders.FactoryID=1 AND MainOrderID IN (SELECT MainOrderID FROM MainOrders" +
                    " WHERE ProfilProductionStatusID = 1 AND ProfilStorageStatusID = 1 AND ProfilExpeditionStatusID=2 AND ProfilDispatchStatusID = 1)" + MegaOrderFilter,
                    ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DT.Clear();
                    DA.Fill(DT);

                    foreach (DataRow Row in DT.Rows)
                    {
                        if (Convert.ToInt32(Row["MeasureID"]) == 2)
                        {
                            //нет параметра "высота"
                            if (Row["Height"].ToString() == "-1")
                                ProfilPogon += Convert.ToDecimal(Row["Length"]) * Convert.ToDecimal(Row["Count"]) / 1000;
                            else
                                ProfilPogon += Convert.ToDecimal(Row["Height"]) * Convert.ToDecimal(Row["Count"]) / 1000;
                        }
                        else
                        {
                            ProfilCount += Convert.ToDecimal(Row["Count"]);
                        }
                        OnExpeditionProfilCost += Convert.ToDecimal(Row["Cost"]);
                    }
                }
                if (CheckZOV)
                {
                    using (SqlDataAdapter DA = new SqlDataAdapter("SELECT DecorOrders.*, infiniu2_catalog.dbo.DecorConfig.MeasureID FROM DecorOrders" +
                        " INNER JOIN infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                        " WHERE DecorOrders.FactoryID=1 AND DecorOrderID IN (SELECT OrderID FROM PackageDetails WHERE PackageID IN" +
                        " (SELECT PackageID FROM Packages WHERE (PackageStatusID=4) AND FactoryID=1 AND ProductType=1))" +
                        " UNION" +
                        " SELECT DecorOrders.*, infiniu2_catalog.dbo.DecorConfig.MeasureID FROM DecorOrders" +
                        " INNER JOIN infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                        " WHERE DecorOrders.FactoryID=1 AND MainOrderID IN (SELECT MainOrderID FROM MainOrders" +
                        " WHERE ProfilProductionStatusID = 1 AND ProfilStorageStatusID = 1 AND ProfilExpeditionStatusID=2 AND ProfilDispatchStatusID = 1)",
                        ConnectionStrings.ZOVOrdersConnectionString))
                    {
                        DT.Clear();
                        DA.Fill(DT);

                        foreach (DataRow Row in DT.Rows)
                        {
                            if (Convert.ToInt32(Row["MeasureID"]) == 2)
                            {
                                //нет параметра "высота"
                                if (Row["Height"].ToString() == "-1")
                                    ProfilPogon += Convert.ToDecimal(Row["Length"]) * Convert.ToDecimal(Row["Count"]) / 1000;
                                else
                                    ProfilPogon += Convert.ToDecimal(Row["Height"]) * Convert.ToDecimal(Row["Count"]) / 1000;
                            }
                            else
                            {
                                ProfilCount += Convert.ToDecimal(Row["Count"]);
                            }
                            OnExpeditionProfilCost += Convert.ToDecimal(Row["Cost"]);
                        }
                    }
                }
                OnExpeditionProfilCost = Decimal.Round(OnExpeditionProfilCost, 1, MidpointRounding.AwayFromZero);
                ProfilPogon = Decimal.Round(ProfilPogon, 1, MidpointRounding.AwayFromZero);
                ProfilCount = Decimal.Round(ProfilCount, 1, MidpointRounding.AwayFromZero);

                //ТПС
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT DecorOrders.*, infiniu2_catalog.dbo.DecorConfig.MeasureID FROM DecorOrders" +
                    " INNER JOIN infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                    " WHERE DecorOrders.FactoryID=2 AND DecorOrderID IN (SELECT OrderID FROM PackageDetails WHERE PackageID IN" +
                    " (SELECT PackageID FROM Packages WHERE (PackageStatusID=4) AND FactoryID=2 AND ProductType=1))" + MegaOrderFilter +
                    " UNION" +
                    " SELECT DecorOrders.*, infiniu2_catalog.dbo.DecorConfig.MeasureID FROM DecorOrders" +
                    " INNER JOIN infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                    " WHERE DecorOrders.FactoryID=2 AND MainOrderID IN (SELECT MainOrderID FROM MainOrders" +
                    " WHERE TPSProductionStatusID = 1 AND TPSStorageStatusID = 1 AND TPSExpeditionStatusID=2 AND TPSDispatchStatusID = 1)" + MegaOrderFilter,
                    ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DT.Clear();
                    DA.Fill(DT);

                    foreach (DataRow Row in DT.Rows)
                    {
                        if (Convert.ToInt32(Row["MeasureID"]) == 2)
                        {
                            //нет параметра "высота"
                            if (Row["Height"].ToString() == "-1")
                                TPSPogon += Convert.ToDecimal(Row["Length"]) * Convert.ToDecimal(Row["Count"]) / 1000;
                            else
                                TPSPogon += Convert.ToDecimal(Row["Height"]) * Convert.ToDecimal(Row["Count"]) / 1000;
                        }
                        else
                        {
                            TPSCount += Convert.ToDecimal(Row["Count"]);
                        }
                        OnExpeditionTPSCost += Convert.ToDecimal(Row["Cost"]);
                    }
                }
                if (CheckZOV)
                {
                    using (SqlDataAdapter DA = new SqlDataAdapter("SELECT DecorOrders.*, infiniu2_catalog.dbo.DecorConfig.MeasureID FROM DecorOrders" +
                        " INNER JOIN infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                        " WHERE DecorOrders.FactoryID=2 AND DecorOrderID IN (SELECT OrderID FROM PackageDetails WHERE PackageID IN" +
                        " (SELECT PackageID FROM Packages WHERE (PackageStatusID=4) AND FactoryID=2 AND ProductType=1))" +
                        " UNION" +
                        " SELECT DecorOrders.*, infiniu2_catalog.dbo.DecorConfig.MeasureID FROM DecorOrders" +
                        " INNER JOIN infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID=infiniu2_catalog.dbo.DecorConfig.DecorConfigID" +
                        " WHERE DecorOrders.FactoryID=2 AND MainOrderID IN (SELECT MainOrderID FROM MainOrders" +
                        " WHERE TPSProductionStatusID = 1 AND TPSStorageStatusID = 1 AND TPSExpeditionStatusID=2 AND TPSDispatchStatusID = 1)",
                        ConnectionStrings.ZOVOrdersConnectionString))
                    {
                        DT.Clear();
                        DA.Fill(DT);

                        foreach (DataRow Row in DT.Rows)
                        {
                            if (Convert.ToInt32(Row["MeasureID"]) == 2)
                            {
                                //нет параметра "высота"
                                if (Row["Height"].ToString() == "-1")
                                    TPSPogon += Convert.ToDecimal(Row["Length"]) * Convert.ToDecimal(Row["Count"]) / 1000;
                                else
                                    TPSPogon += Convert.ToDecimal(Row["Height"]) * Convert.ToDecimal(Row["Count"]) / 1000;
                            }
                            else
                            {
                                TPSCount += Convert.ToDecimal(Row["Count"]);
                            }
                            OnExpeditionTPSCost += Convert.ToDecimal(Row["Cost"]);
                        }
                    }
                }
                OnExpeditionTPSCost = Decimal.Round(OnExpeditionTPSCost, 1, MidpointRounding.AwayFromZero);
                TPSPogon = Decimal.Round(TPSPogon, 1, MidpointRounding.AwayFromZero);
                TPSCount = Decimal.Round(TPSCount, 1, MidpointRounding.AwayFromZero);
            }
        }

        public void Filter()
        {
            GetClientGroup();
            NonAgreementStatistics();
            AgreementStatistics();
            OnProductionStatistics();
            InProductionStatistics();
            OnStorageStatistics();
            OnExpeditionStatistics();
        }

        public bool GetZOV
        {
            get
            {
                //foreach (object itemChecked in ClientGroupsList.CheckedItems)
                //{
                //    DataRowView castedItem = (DataRowView)itemChecked;
                //    if (Convert.ToInt32(castedItem["ClientGroupID"]) == 1)
                //        return true;
                //}

                return false;
            }
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
    }
}