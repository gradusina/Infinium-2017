using System;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Windows.Forms;

namespace Infinium.Modules.WorkAssignments
{
    public class FrontsOrdersManager
    {
        private DataTable AllBatchFrontsDT;

        private DataTable AllPrintedFrontsDT;

        private BindingSource BatchFrontsBS = null;

        private DataTable BatchFrontsDT;

        private DataTable FrameColorsDataTable = null;

        private BindingSource FrameColorsSummaryBS = null;

        private DataTable FrameColorsSummaryDT;

        private DataTable FrontsDataTable = null;

        private BindingSource FrontsSummaryBS = null;

        private DataTable FrontsSummaryDT;

        private DataTable InsetColorsDataTable = null;

        private BindingSource InsetColorsSummaryBS = null;

        private DataTable InsetColorsSummaryDT;

        private DataTable InsetTypesDataTable = null;

        private BindingSource InsetTypesSummaryBS = null;

        private DataTable InsetTypesSummaryDT;

        private DataTable PatinaDataTable = null;

        private DataTable PatinaRALDataTable = null;

        private BindingSource SizesSummaryBS = null;

        private DataTable SizesSummaryDT;

        private DataTable TechnoProfilesDataTable = null;

        private DataTable TechStoreDataTable = null;

        public FrontsOrdersManager()
        {
        }

        public BindingSource BatchFrontsList => BatchFrontsBS;

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

        public BindingSource FrameColorsSummaryList => FrameColorsSummaryBS;

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

        public BindingSource FrontsSummaryList => FrontsSummaryBS;

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

        public BindingSource InsetColorsSummaryList => InsetColorsSummaryBS;

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

        public BindingSource InsetTypesSummaryList => InsetTypesSummaryBS;

        public DataTable OrdersDT => AllBatchFrontsDT;

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

        public BindingSource SizesSummaryList => SizesSummaryBS;

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

        public void FilterFrameColors(int FrontID, int FrontTypeID)
        {
            FrameColorsSummaryBS.Filter = "FrontID=" + FrontID + " AND FrontTypeID=" + FrontTypeID;
            FrameColorsSummaryBS.MoveFirst();
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
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Sofia) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.SofiaNotColored) + 
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Lorenzo) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Elegant) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.ElegantPat) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Patricia1) +" OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Scandia) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Turin1) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Turin1_1) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Dakota) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.DakotaPat) + 
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Turin3) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Turin3NotColored) + 
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.LeonTPS) +
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
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.GrandVg) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Polo) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.PoloVg) + ")";

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
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Sofia) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.SofiaNotColored) + 
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Lorenzo) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Elegant) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.ElegantPat) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Patricia1) +" OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Scandia) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Turin1) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Turin1_1) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Dakota) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.DakotaPat) + 
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Turin3) + " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Turin3NotColored) + 
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.LeonTPS) +
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
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.GrandVg) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.Polo) +
                    " OR FrontsOrders.FrontID=" + Convert.ToInt32(Fronts.PoloVg) + ")";

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

        public void FilterInsetColors(int FrontID, int FrontTypeID, int ColorID, int PatinaID, int InsetTypeID)
        {
            InsetColorsSummaryBS.Filter = "FrontID=" + FrontID + " AND FrontTypeID=" +
                FrontTypeID + " AND InsetTypeID=" + InsetTypeID + " AND ColorID=" + ColorID + " AND PatinaID=" + PatinaID;
            InsetColorsSummaryBS.MoveFirst();
        }

        public void FilterInsetTypes(int FrontID, int FrontTypeID, int ColorID, int PatinaID)
        {
            InsetTypesSummaryBS.Filter = "FrontID=" + FrontID + " AND FrontTypeID=" + FrontTypeID +
                " AND ColorID=" + ColorID + " AND PatinaID=" + PatinaID;
            InsetTypesSummaryBS.MoveFirst();
        }

        public void FilterSizes(int FrontID, int FrontTypeID, int ColorID, int PatinaID, int InsetTypeID, int InsetColorID)
        {
            SizesSummaryBS.Filter = "FrontID=" + FrontID + " AND FrontTypeID=" +
                FrontTypeID + " AND ColorID=" + ColorID + " AND PatinaID=" + PatinaID + " AND InsetTypeID=" + InsetTypeID +
                " AND InsetColorID=" + InsetColorID;
            SizesSummaryBS.MoveFirst();
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

        //private void GetGridMargins(int FrontID, ref int MarginHeight, ref int MarginWidth)
        //{
        //    DataRow[] Rows = InsetMarginsDT.Select("FrontID = " + FrontID);
        //    if (Rows.Count() == 0)
        //        return;
        //    MarginHeight = Convert.ToInt32(Rows[0]["GridHeight"]);
        //    MarginWidth = Convert.ToInt32(Rows[0]["GridWidth"]);
        //}
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

        public void GetPrintedFronts(int WorkAssignmentID)
        {
            string SelectCommand = @"SELECT * FROM AssignmentsInWork WHERE WorkAssignmentID=" + WorkAssignmentID;

            AllPrintedFrontsDT.Clear();
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                DA.Fill(AllPrintedFrontsDT);
            }
        }

        public void Initialize()
        {
            Create();
            Fill();
            Binding();
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

        private void Binding()
        {
            BatchFrontsBS.DataSource = BatchFrontsDT;
            FrontsSummaryBS.DataSource = FrontsSummaryDT;
            FrameColorsSummaryBS.DataSource = FrameColorsSummaryDT;
            InsetTypesSummaryBS.DataSource = InsetTypesSummaryDT;
            InsetColorsSummaryBS.DataSource = InsetColorsSummaryDT;
            SizesSummaryBS.DataSource = SizesSummaryDT;
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
    }
}