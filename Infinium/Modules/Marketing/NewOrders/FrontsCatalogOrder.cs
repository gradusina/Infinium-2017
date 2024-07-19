
using System;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Windows.Forms;

namespace Infinium.Modules.Marketing.NewOrders
{
    public class FrontsCatalogOrder
    {
        private ComponentFactory.Krypton.Toolkit.KryptonComboBox HeightEdit;
        private ComponentFactory.Krypton.Toolkit.KryptonComboBox WidthEdit;

        private DataTable TempFrontsConfigDataTable = null;
        private DataTable TempFrontsDataTable = null;

        public DataTable ConstFrontsConfigDataTable = null;
        public DataTable ConstFrontsDataTable = null;
        public DataTable ConstTechnoProfilesDataTable = null;
        public DataTable ConstColorsDataTable = null;
        public DataTable ConstPatinaDataTable = null;
        public DataTable ConstInsetTypesDataTable = null;
        public DataTable ConstInsetColorsDataTable = null;

        public DataTable FrontsDataTable = null;
        public DataTable TechnoProfilesDataTable = null;
        public DataTable FrameColorsDataTable = null;
        public DataTable PatinaDataTable = null;
        public DataTable PatinaRALDataTable = null;
        public DataTable InsetTypesDataTable = null;
        public DataTable InsetColorsDataTable = null;
        public DataTable TechnoFrameColorsDataTable = null;
        public DataTable TechnoInsetTypesDataTable = null;
        public DataTable TechnoInsetColorsDataTable = null;
        public DataTable InsetPriceDataTable = null;
        public DataTable AluminiumFrontsDataTable = null;
        public DataTable GridSizesDataTable = null;
        public DataTable CurrencyTypesDT = null;

        public DataTable HeightDataTable = null;
        public DataTable WidthDataTable = null;

        public BindingSource FrontsBindingSource = null;
        public BindingSource TechnoProfilesBindingSource = null;
        public BindingSource FrameColorsBindingSource = null;
        public BindingSource PatinaBindingSource = null;
        public BindingSource InsetTypesBindingSource = null;
        public BindingSource InsetColorsBindingSource = null;
        public BindingSource TechnoFrameColorsBindingSource = null;
        public BindingSource TechnoInsetTypesBindingSource = null;
        public BindingSource TechnoInsetColorsBindingSource = null;
        public BindingSource FrontsConfigBindingSource = null;

        public String FrontsBindingSourceDisplayMember = null;
        public String TechnoProfilesBindingSourceDisplayMember = null;
        public String FrameColorsBindingSourceDisplayMember = null;
        public String PatinaBindingSourceDisplayMember = null;
        public String InsetColorsBindingSourceDisplayMember = null;
        public String InsetTypesBindingSourceDisplayMember = null;

        public String FrontsBindingSourceValueMember = null;
        public String TechnoProfilesBindingSourceValueMember = null;
        public String FrameColorsBindingSourceValueMember = null;
        public String PatinaBindingSourceValueMember = null;
        public String InsetColorsBindingSourceValueMember = null;
        public String InsetTypesBindingSourceValueMember = null;

        public FrontsCatalogOrder()
        {

        }

        public FrontsCatalogOrder(
            ref ComponentFactory.Krypton.Toolkit.KryptonComboBox tHeightEdit,
            ref ComponentFactory.Krypton.Toolkit.KryptonComboBox tWidthEdit)
        {
            HeightEdit = tHeightEdit;
            WidthEdit = tWidthEdit;
        }

        private void Create()
        {
            ConstFrontsConfigDataTable = new DataTable();
            ConstFrontsDataTable = new DataTable();
            ConstTechnoProfilesDataTable = new DataTable();
            ConstColorsDataTable = new DataTable();
            ConstPatinaDataTable = new DataTable();
            ConstInsetTypesDataTable = new DataTable();
            ConstInsetColorsDataTable = new DataTable();

            FrontsDataTable = new DataTable();
            TechnoProfilesDataTable = new DataTable();
            FrameColorsDataTable = new DataTable();
            PatinaDataTable = new DataTable();
            InsetTypesDataTable = new DataTable();
            InsetColorsDataTable = new DataTable();
            TechnoFrameColorsDataTable = new DataTable();
            TechnoInsetTypesDataTable = new DataTable();
            TechnoInsetColorsDataTable = new DataTable();
            HeightDataTable = new DataTable();
            WidthDataTable = new DataTable();

            TempFrontsConfigDataTable = new DataTable();
            TempFrontsDataTable = new DataTable();

            ConstPatinaDataTable = new DataTable();
            PatinaDataTable = new DataTable();

            HeightDataTable = new DataTable();
            WidthDataTable = new DataTable();

            FrontsConfigBindingSource = new BindingSource();
            FrontsBindingSource = new BindingSource();
            TechnoProfilesBindingSource = new BindingSource();
            FrameColorsBindingSource = new BindingSource();
            PatinaBindingSource = new BindingSource();
            InsetTypesBindingSource = new BindingSource();
            InsetColorsBindingSource = new BindingSource();
            TechnoFrameColorsBindingSource = new BindingSource();
            TechnoInsetTypesBindingSource = new BindingSource();
            TechnoInsetColorsBindingSource = new BindingSource();
        }

        private void GetColorsDT()
        {
            ConstColorsDataTable = new DataTable();
            ConstColorsDataTable.Columns.Add(new DataColumn("ColorID", Type.GetType("System.Int64")));
            ConstColorsDataTable.Columns.Add(new DataColumn("ColorName", Type.GetType("System.String")));
            ConstColorsDataTable.Columns.Add(new DataColumn("Cvet", Type.GetType("System.String")));
            string SelectCommand = @"SELECT TechStoreID, TechStoreName, Cvet FROM TechStore
                WHERE TechStoreSubGroupID IN (SELECT TechStoreSubGroupID FROM TechStoreSubGroups WHERE (TechStoreGroupID = 11 OR TechStoreGroupID = 1))
                ORDER BY TechStoreName";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.CatalogConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    DA.Fill(DT);
                    {
                        DataRow NewRow = ConstColorsDataTable.NewRow();
                        NewRow["ColorID"] = -1;
                        NewRow["ColorName"] = "-";
                        NewRow["Cvet"] = "000";
                        ConstColorsDataTable.Rows.Add(NewRow);
                    }
                    {
                        DataRow NewRow = ConstColorsDataTable.NewRow();
                        NewRow["ColorID"] = 0;
                        NewRow["ColorName"] = "на выбор";
                        NewRow["Cvet"] = "0000000";
                        ConstColorsDataTable.Rows.Add(NewRow);
                    }
                    for (int i = 0; i < DT.Rows.Count; i++)
                    {
                        DataRow NewRow = ConstColorsDataTable.NewRow();
                        NewRow["ColorID"] = Convert.ToInt64(DT.Rows[i]["TechStoreID"]);
                        NewRow["ColorName"] = DT.Rows[i]["TechStoreName"].ToString();
                        NewRow["Cvet"] = DT.Rows[i]["Cvet"].ToString();
                        ConstColorsDataTable.Rows.Add(NewRow);
                    }
                }
            }
        }

        private void GetInsetColorsDT()
        {
            ConstInsetColorsDataTable = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT InsetColors.InsetColorID, InsetColors.GroupID, infiniu2_catalog.dbo.TechStore.TechStoreName AS InsetColorName FROM InsetColors" +
                " INNER JOIN infiniu2_catalog.dbo.TechStore ON InsetColors.InsetColorID = infiniu2_catalog.dbo.TechStore.TechStoreID ORDER BY TechStoreName", ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(ConstInsetColorsDataTable);
                {
                    DataRow NewRow = ConstInsetColorsDataTable.NewRow();
                    NewRow["InsetColorID"] = -1;
                    NewRow["GroupID"] = -1;
                    NewRow["InsetColorName"] = "-";
                    ConstInsetColorsDataTable.Rows.Add(NewRow);
                }
                {
                    DataRow NewRow = ConstInsetColorsDataTable.NewRow();
                    NewRow["InsetColorID"] = 0;
                    NewRow["GroupID"] = -1;
                    NewRow["InsetColorName"] = "на выбор";
                    ConstInsetColorsDataTable.Rows.Add(NewRow);
                }

            }

        }

        private void Fill(bool isNewOrders)
        {
            string SelectCommand = @"SELECT * FROM FrontsConfig" +
                " WHERE AccountingName IS NOT NULL AND InvNumber IS NOT NULL";
            //using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.CatalogConnectionString))
            //{
            //    DA.Fill(ConstFrontsConfigDataTable);
            //    ConstFrontsConfigDataTable.Columns.Add(new DataColumn("Excluzive", Type.GetType("System.Int32")));
            //}
            if (isNewOrders)
                ConstFrontsConfigDataTable = TablesManager.FrontsConfigDataTable;
            else
                ConstFrontsConfigDataTable = TablesManager.FrontsConfigDataTableAll;
            ConstFrontsConfigDataTable.Columns.Add(new DataColumn("Excluzive", Type.GetType("System.Int32")));


            if (isNewOrders)
                SelectCommand = @"SELECT DISTINCT TechStoreID AS FrontID, TechStoreName AS FrontName FROM TechStore 
                WHERE TechStoreID IN (SELECT FrontID FROM FrontsConfig WHERE Enabled = 1 AND AccountingName IS NOT NULL AND InvNumber IS NOT NULL)
                ORDER BY TechStoreName";
            else
                SelectCommand = @"SELECT DISTINCT TechStoreID AS FrontID, TechStoreName AS FrontName FROM TechStore 
                WHERE TechStoreID IN (SELECT FrontID FROM FrontsConfig WHERE AccountingName IS NOT NULL AND InvNumber IS NOT NULL)
                ORDER BY TechStoreName";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(ConstFrontsDataTable);
                ConstFrontsDataTable.Columns.Add(new DataColumn("Excluzive", Type.GetType("System.Int32")));
            }

            if (isNewOrders)
                SelectCommand = @"SELECT DISTINCT TechStoreID AS TechnoProfileID, TechStoreName AS TechnoProfileName FROM TechStore 
                WHERE TechStoreID IN (SELECT TechnoProfileID FROM FrontsConfig WHERE Enabled = 1 AND AccountingName IS NOT NULL AND InvNumber IS NOT NULL)
                ORDER BY TechStoreName";
            else
            SelectCommand = @"SELECT DISTINCT TechStoreID AS TechnoProfileID, TechStoreName AS TechnoProfileName FROM TechStore 
                WHERE TechStoreID IN (SELECT TechnoProfileID FROM FrontsConfig WHERE AccountingName IS NOT NULL AND InvNumber IS NOT NULL)
                ORDER BY TechStoreName";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(ConstTechnoProfilesDataTable);

                DataRow NewRow = ConstTechnoProfilesDataTable.NewRow();
                NewRow["TechnoProfileID"] = -1;
                NewRow["TechnoProfileName"] = "-";
                ConstTechnoProfilesDataTable.Rows.InsertAt(NewRow, 0);
            }

            SelectCommand = "SELECT * FROM CurrencyTypes";
            CurrencyTypesDT = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(CurrencyTypesDT);
            }

            InsetPriceDataTable = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM InsetPrice",
                ConnectionStrings.MarketingReferenceConnectionString))
            {
                DA.Fill(InsetPriceDataTable);
            }

            AluminiumFrontsDataTable = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM AluminiumFronts",
                ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(AluminiumFrontsDataTable);
            }

            GridSizesDataTable = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM GridSizes",
                ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(GridSizesDataTable);
            }

            GetColorsDT();
            GetInsetColorsDT();
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM InsetTypes",
                ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(ConstInsetTypesDataTable);
            }
            InsetTypesDataTable = ConstInsetTypesDataTable.Copy();
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM Patina",
                ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(ConstPatinaDataTable);
            }
            using (DataView DV = new DataView(ConstFrontsDataTable))
            {
                FrontsDataTable = DV.ToTable(true, new string[] { "FrontName" });
            }

            TempFrontsConfigDataTable = ConstFrontsConfigDataTable.Copy();
            TempFrontsDataTable = ConstFrontsDataTable.Copy();
            //FrontsDataTable = ConstFrontsDataTable.Copy();
            FrameColorsDataTable = ConstColorsDataTable.Copy();
            TechnoProfilesDataTable = ConstTechnoProfilesDataTable.Copy();
            PatinaDataTable = ConstPatinaDataTable.Copy();
            PatinaRALDataTable = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT PatinaRAL.*, Patina.Patina FROM PatinaRAL INNER JOIN Patina ON Patina.PatinaID=PatinaRAL.PatinaID",
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
            InsetTypesDataTable = ConstInsetTypesDataTable.Copy();
            InsetColorsDataTable = ConstInsetColorsDataTable.Copy();
            TechnoFrameColorsDataTable = ConstColorsDataTable.Copy();
            TechnoInsetTypesDataTable = ConstInsetTypesDataTable.Copy();
            TechnoInsetColorsDataTable = ConstInsetColorsDataTable.Copy();
        }

        private void Binding()
        {
            FrontsBindingSource.DataSource = FrontsDataTable;
            TechnoProfilesBindingSource.DataSource = TechnoProfilesDataTable;
            FrameColorsBindingSource.DataSource = FrameColorsDataTable;
            PatinaBindingSource.DataSource = PatinaDataTable;
            InsetTypesBindingSource.DataSource = InsetTypesDataTable;
            InsetColorsBindingSource.DataSource = InsetColorsDataTable;
            TechnoFrameColorsBindingSource.DataSource = TechnoFrameColorsDataTable;
            TechnoInsetTypesBindingSource.DataSource = TechnoInsetTypesDataTable;
            TechnoInsetColorsBindingSource.DataSource = TechnoInsetColorsDataTable;
            FrontsConfigBindingSource.DataSource = ConstFrontsConfigDataTable;
            //HeightBindingSource.DataSource = HeightDataTable;
            //WidthBindingSource.DataSource = WidthDataTable;

            FrontsBindingSourceDisplayMember = "FrontName";
            TechnoProfilesBindingSourceDisplayMember = "TechnoProfileName";
            FrameColorsBindingSourceDisplayMember = "ColorName";
            PatinaBindingSourceDisplayMember = "PatinaName";
            InsetColorsBindingSourceDisplayMember = "InsetColorName";
            InsetTypesBindingSourceDisplayMember = "InsetTypeName";

            FrontsBindingSourceValueMember = "FrontName";
            TechnoProfilesBindingSourceValueMember = "TechnoProfileID";
            FrameColorsBindingSourceValueMember = "ColorID";
            PatinaBindingSourceValueMember = "PatinaID";
            InsetColorsBindingSourceValueMember = "InsetColorID";
            InsetTypesBindingSourceValueMember = "InsetTypeID";
        }

        public void GetInsetTypes()
        {
            DataTable Table = new DataTable();

            using (DataView DV = new DataView(TempFrontsConfigDataTable))
            {
                Table = DV.ToTable(true, new string[] { "InsetTypeID" });
            }

            InsetTypesDataTable.Clear();

            foreach (DataRow Row in Table.Rows)
                foreach (DataRow CRow in ConstInsetTypesDataTable.Rows)
                {
                    if (Convert.ToInt32(CRow["InsetTypeID"]) == Convert.ToInt32(Row["InsetTypeID"]))
                    {
                        DataRow NewRow = InsetTypesDataTable.NewRow();
                        NewRow["InsetTypeID"] = CRow["InsetTypeID"];
                        NewRow["GroupID"] = CRow["GroupID"];
                        NewRow["InsetTypeName"] = CRow["InsetTypeName"];
                        InsetTypesDataTable.Rows.Add(NewRow);
                        break;
                    }
                }

            Table.Dispose();

            InsetTypesDataTable.DefaultView.Sort = "InsetTypeName ASC";
            InsetTypesBindingSource.MoveFirst();
        }

        public void GetFrameColors()
        {
            DataTable Table = new DataTable();

            using (DataView DV = new DataView(TempFrontsConfigDataTable))
            {
                Table = DV.ToTable(true, new string[] { "ColorID" });
            }

            FrameColorsDataTable.Clear();

            foreach (DataRow Row in Table.Rows)
                foreach (DataRow CRow in ConstColorsDataTable.Rows)
                {
                    if (Convert.ToInt32(CRow["ColorID"]) == Convert.ToInt32(Row["ColorID"]))
                    {
                        DataRow NewRow = FrameColorsDataTable.NewRow();
                        NewRow["ColorID"] = CRow["ColorID"];
                        NewRow["ColorName"] = CRow["ColorName"];
                        FrameColorsDataTable.Rows.Add(NewRow);
                        break;
                    }
                }

            Table.Dispose();

            FrameColorsDataTable.DefaultView.Sort = "ColorName ASC";
            FrameColorsBindingSource.MoveFirst();
        }

        public void GetInsetColors()
        {
            DataTable Table = new DataTable();

            using (DataView DV = new DataView(TempFrontsConfigDataTable))
            {
                Table = DV.ToTable(true, new string[] { "InsetColorID" });
            }

            InsetColorsDataTable.Clear();

            foreach (DataRow Row in Table.Rows)
                foreach (DataRow CRow in ConstInsetColorsDataTable.Rows)
                {
                    if (Convert.ToInt32(CRow["InsetColorID"]) == Convert.ToInt32(Row["InsetColorID"]))
                    {
                        DataRow NewRow = InsetColorsDataTable.NewRow();
                        NewRow["InsetColorID"] = CRow["InsetColorID"];
                        NewRow["GroupID"] = CRow["GroupID"];
                        NewRow["InsetColorName"] = CRow["InsetColorName"];
                        InsetColorsDataTable.Rows.Add(NewRow);
                        break;
                    }
                }

            Table.Dispose();
            InsetColorsDataTable.DefaultView.Sort = "InsetColorName ASC";
            InsetColorsBindingSource.MoveFirst();
        }

        public void GetTechnoProfiles()
        {
            DataTable Table = new DataTable();

            using (DataView DV = new DataView(TempFrontsConfigDataTable))
            {
                Table = DV.ToTable(true, new string[] { "TechnoProfileID" });
            }

            TechnoProfilesDataTable.Clear();

            foreach (DataRow Row in Table.Rows)
                foreach (DataRow CRow in ConstTechnoProfilesDataTable.Rows)
                {
                    if (Convert.ToInt32(CRow["TechnoProfileID"]) == Convert.ToInt32(Row["TechnoProfileID"]))
                    {
                        DataRow NewRow = TechnoProfilesDataTable.NewRow();
                        NewRow["TechnoProfileID"] = CRow["TechnoProfileID"];
                        NewRow["TechnoProfileName"] = CRow["TechnoProfileName"];
                        TechnoProfilesDataTable.Rows.Add(NewRow);
                        break;
                    }
                }

            Table.Dispose();

            TechnoProfilesDataTable.DefaultView.Sort = "TechnoProfileName ASC";
            TechnoProfilesBindingSource.MoveFirst();
        }

        public void GetTechnoFrameColors()
        {
            DataTable Table = new DataTable();

            using (DataView DV = new DataView(TempFrontsConfigDataTable))
            {
                Table = DV.ToTable(true, new string[] { "TechnoColorID" });
            }

            TechnoFrameColorsDataTable.Clear();

            foreach (DataRow Row in Table.Rows)
                foreach (DataRow CRow in ConstColorsDataTable.Rows)
                {
                    if (Convert.ToInt32(CRow["ColorID"]) == Convert.ToInt32(Row["TechnoColorID"]))
                    {
                        DataRow NewRow = TechnoFrameColorsDataTable.NewRow();
                        NewRow["ColorID"] = CRow["ColorID"];
                        NewRow["ColorName"] = CRow["ColorName"];
                        TechnoFrameColorsDataTable.Rows.Add(NewRow);
                        break;
                    }
                }

            Table.Dispose();

            TechnoFrameColorsDataTable.DefaultView.Sort = "ColorName ASC";
            TechnoFrameColorsBindingSource.MoveFirst();
        }

        public void GetTechnoInsetTypes()
        {
            DataTable Table = new DataTable();

            using (DataView DV = new DataView(TempFrontsConfigDataTable))
            {
                Table = DV.ToTable(true, new string[] { "TechnoInsetTypeID" });
            }

            TechnoInsetTypesDataTable.Clear();

            foreach (DataRow Row in Table.Rows)
                foreach (DataRow CRow in ConstInsetTypesDataTable.Rows)
                {
                    if (Convert.ToInt32(CRow["InsetTypeID"]) == Convert.ToInt32(Row["TechnoInsetTypeID"]))
                    {
                        DataRow NewRow = TechnoInsetTypesDataTable.NewRow();
                        NewRow["InsetTypeID"] = CRow["InsetTypeID"];
                        NewRow["GroupID"] = CRow["GroupID"];
                        NewRow["InsetTypeName"] = CRow["InsetTypeName"];
                        TechnoInsetTypesDataTable.Rows.Add(NewRow);
                        break;
                    }
                }

            Table.Dispose();

            TechnoInsetTypesDataTable.DefaultView.Sort = "InsetTypeName ASC";
            TechnoInsetTypesBindingSource.MoveFirst();
        }

        public void GetTechnoInsetColors()
        {
            DataTable Table = new DataTable();

            using (DataView DV = new DataView(TempFrontsConfigDataTable))
            {
                Table = DV.ToTable(true, new string[] { "TechnoInsetColorID" });
            }

            TechnoInsetColorsDataTable.Clear();

            foreach (DataRow Row in Table.Rows)
                foreach (DataRow CRow in ConstInsetColorsDataTable.Rows)
                {
                    if (Convert.ToInt32(CRow["InsetColorID"]) == Convert.ToInt32(Row["TechnoInsetColorID"]))
                    {
                        DataRow NewRow = TechnoInsetColorsDataTable.NewRow();
                        NewRow["InsetColorID"] = CRow["InsetColorID"];
                        NewRow["GroupID"] = CRow["GroupID"];
                        NewRow["InsetColorName"] = CRow["InsetColorName"];
                        TechnoInsetColorsDataTable.Rows.Add(NewRow);
                        break;
                    }
                }

            Table.Dispose();
            //if (TechnoInsetColorsDataTable.Rows.Count == 0)
            //{
            //    DataRow NewRow = TechnoInsetColorsDataTable.NewRow();
            //    NewRow["InsetColorID"] = "-1";
            //    NewRow["InsetColorName"] = "-";
            //    TechnoInsetColorsDataTable.Rows.Add(NewRow);
            //}
            //else
            TechnoInsetColorsDataTable.DefaultView.Sort = "InsetColorName ASC";
            TechnoInsetColorsBindingSource.MoveFirst();
        }

        public void GetPatina()
        {
            DataTable Table = new DataTable();

            using (DataView DV = new DataView(TempFrontsConfigDataTable, string.Empty, "PatinaID", DataViewRowState.CurrentRows))
            {
                Table = DV.ToTable(true, new string[] { "PatinaID" });
            }

            PatinaDataTable.Clear();
            bool HasPatinaRAL = false;

            for (int i = 0; i < Table.Rows.Count; i++)
            {
                DataRow[] fRows = PatinaRALDataTable.Select("PatinaID=" + Convert.ToInt32(Table.Rows[i]["PatinaID"]));
                if (fRows.Count() > 0)
                {
                    HasPatinaRAL = true;
                    break;
                }
            }
            if (Table.Rows.Count > 0 && HasPatinaRAL)
            {
                foreach (DataRow Row in Table.Rows)
                {
                    if (Convert.ToInt32(Row["PatinaID"]) == -1)
                    {
                        DataRow NewRow = PatinaDataTable.NewRow();
                        NewRow["PatinaID"] = Row["PatinaID"];
                        NewRow["PatinaName"] = "-";
                        PatinaDataTable.Rows.Add(NewRow);
                    }
                    foreach (DataRow item in PatinaRALDataTable.Select("PatinaID=" + Convert.ToInt32(Row["PatinaID"])))
                    {
                        DataRow NewRow = PatinaDataTable.NewRow();
                        NewRow["PatinaID"] = item["PatinaRALID"];
                        NewRow["PatinaName"] = item["PatinaRAL"]; NewRow["Patina"] = item["Patina"];
                        NewRow["DisplayName"] = item["DisplayName"];
                        PatinaDataTable.Rows.Add(NewRow);
                    }
                }
            }
            else
            {
                foreach (DataRow Row in Table.Rows)
                    foreach (DataRow CRow in ConstPatinaDataTable.Rows)
                    {
                        if (Convert.ToInt32(CRow["PatinaID"]) == Convert.ToInt32(Row["PatinaID"]))
                        {
                            DataRow NewRow = PatinaDataTable.NewRow();
                            NewRow["PatinaID"] = CRow["PatinaID"];
                            NewRow["PatinaName"] = CRow["PatinaName"];
                            PatinaDataTable.Rows.Add(NewRow);
                            break;
                        }
                    }
            }
            Table.Dispose();

            PatinaBindingSource.MoveFirst();
        }

        public void GetHeight()
        {
            HeightDataTable.Clear();

            using (DataView DV = new DataView(TempFrontsConfigDataTable))
            {
                HeightDataTable = DV.ToTable(true, new string[] { "Height" });
            }

            HeightDataTable.DefaultView.Sort = "Height ASC";
        }

        public void GetWidth()
        {
            WidthDataTable.Clear();

            using (DataView DV = new DataView(TempFrontsConfigDataTable))
            {
                WidthDataTable = DV.ToTable(true, new string[] { "Width" });
            }

            WidthDataTable.DefaultView.Sort = "Width ASC";
        }


        public void Initialize(bool isNewOrders)
        {
            Create();
            Fill(isNewOrders);
            Binding();
        }

        private bool HasHeightParameter(DataRow[] Rows, int Height)
        {
            foreach (DataRow Row in Rows)
            {
                if (Row["Height"].ToString() == Height.ToString())
                    return true;
            }

            return false;
        }

        private bool HasWidthParameter(DataRow[] Rows, int Width)
        {
            foreach (DataRow Row in Rows)
            {
                if (Row["Width"].ToString() == Width.ToString())
                    return true;
            }

            return false;
        }

        public void FilterFronts(bool bExcluzive)
        {
            FrontsDataTable.Clear();
            TechnoProfilesDataTable.Clear();
            FrameColorsDataTable.Clear();
            PatinaDataTable.Clear();
            InsetTypesDataTable.Clear();
            InsetColorsDataTable.Clear();
            TechnoFrameColorsDataTable.Clear();
            TechnoInsetTypesDataTable.Clear();
            TechnoInsetColorsDataTable.Clear();
            HeightDataTable.Clear();
            WidthDataTable.Clear();
            DataTable TempItemsDataTable = ConstFrontsDataTable.Copy();
            string RowFilter = string.Empty;
            using (DataView DV = new DataView(TempItemsDataTable))
            {
                if (bExcluzive)
                    RowFilter += " (Excluzive=1)";
                else
                    RowFilter += " (Excluzive IS NULL OR Excluzive=1)";
                DV.RowFilter = RowFilter;
                TempItemsDataTable = DV.ToTable(true, new string[] { "FrontName" });
            }

            for (int d = 0; d < TempItemsDataTable.Rows.Count; d++)
            {
                DataRow NewRow = FrontsDataTable.NewRow();
                NewRow["FrontName"] = TempItemsDataTable.Rows[d]["FrontName"].ToString();
                FrontsDataTable.Rows.Add(NewRow);
            }
            FrontsDataTable.DefaultView.Sort = "FrontName ASC";
            TempItemsDataTable.Dispose();
        }

        public void FilterCatalogFrameColors(string FrontName, bool bExcluzive)
        {
            TempFrontsDataTable.Clear();
            TempFrontsDataTable = ConstFrontsDataTable;
            using (DataView DV = new DataView(TempFrontsDataTable))
            {
                DV.RowFilter = "FrontName='" + FrontName + "'";
                TempFrontsDataTable = DV.ToTable();
            }
            string filter = string.Empty;
            for (int i = 0; i < TempFrontsDataTable.Rows.Count; i++)
                filter += Convert.ToInt32(TempFrontsDataTable.Rows[i]["FrontID"]) + ",";
            if (filter.Length > 0)
            {
                filter = filter.Substring(0, filter.Length - 1);
                filter = "FrontID IN (" + filter + ")";
            }
            else
                filter = "FrontID <> - 1";

            TempFrontsConfigDataTable.Clear();
            TempFrontsConfigDataTable = ConstFrontsConfigDataTable;
            if (bExcluzive)
                filter += " AND (Excluzive=1)";
            else
                filter += " AND (Excluzive IS NULL OR Excluzive=1)";
            using (DataView DV = new DataView(TempFrontsConfigDataTable))
            {
                DV.RowFilter = filter;

                TempFrontsConfigDataTable = DV.ToTable();

                GetFrameColors();
            }
        }//фильтрует и заполняет цвета профиля по выбранному фасаду

        public void FilterCatalogTechnoProfiles(string FrontName, int ColorID, bool bExcluzive) //фильтрует и заполняет типы наполнителя по выбранному типу фасада
        {
            TempFrontsDataTable.Clear();
            TempFrontsDataTable = ConstFrontsDataTable;
            using (DataView DV = new DataView(TempFrontsDataTable))
            {
                DV.RowFilter = "FrontName='" + FrontName + "'";
                TempFrontsDataTable = DV.ToTable();
            }
            string filter = string.Empty;
            for (int i = 0; i < TempFrontsDataTable.Rows.Count; i++)
                filter += Convert.ToInt32(TempFrontsDataTable.Rows[i]["FrontID"]) + ",";
            if (filter.Length > 0)
            {
                filter = filter.Substring(0, filter.Length - 1);
                filter = "FrontID IN (" + filter + ")";
            }
            else
                filter = "FrontID <> - 1";

            TempFrontsConfigDataTable.Clear();
            TempFrontsConfigDataTable = ConstFrontsConfigDataTable;
            if (bExcluzive)
                filter += " AND (Excluzive=1)";
            else
                filter += " AND (Excluzive IS NULL OR Excluzive=1)";
            using (DataView DV = new DataView(TempFrontsConfigDataTable))
            {
                DV.RowFilter = filter;
                DV.RowFilter += " AND ColorID=" + ColorID;

                TempFrontsConfigDataTable = DV.ToTable();

                GetTechnoProfiles();
            }
        }

        public void FilterCatalogTechnoFrameColors(string FrontName, int ColorID, int TechnoProfileID, bool bExcluzive) //фильтрует и заполняет типы наполнителя по выбранному типу фасада
        {
            TempFrontsDataTable.Clear();
            TempFrontsDataTable = ConstFrontsDataTable;
            using (DataView DV = new DataView(TempFrontsDataTable))
            {
                DV.RowFilter = "FrontName='" + FrontName + "'";
                TempFrontsDataTable = DV.ToTable();
            }
            string filter = string.Empty;
            for (int i = 0; i < TempFrontsDataTable.Rows.Count; i++)
                filter += Convert.ToInt32(TempFrontsDataTable.Rows[i]["FrontID"]) + ",";
            if (filter.Length > 0)
            {
                filter = filter.Substring(0, filter.Length - 1);
                filter = "FrontID IN (" + filter + ")";
            }
            else
                filter = "FrontID <> - 1";

            TempFrontsConfigDataTable.Clear();
            TempFrontsConfigDataTable = ConstFrontsConfigDataTable;
            if (bExcluzive)
                filter += " AND (Excluzive=1)";
            else
                filter += " AND (Excluzive IS NULL OR Excluzive=1)";
            using (DataView DV = new DataView(TempFrontsConfigDataTable))
            {
                DV.RowFilter = filter;
                DV.RowFilter += " AND ColorID=" + ColorID;
                DV.RowFilter += " AND TechnoProfileID=" + TechnoProfileID;

                TempFrontsConfigDataTable = DV.ToTable();

                GetTechnoFrameColors();
            }
        }

        public void FilterCatalogPatina(string FrontName, int ColorID, int TechnoProfileID, int TechnoColorID, bool bExcluzive) //фильтрует и заполняет типы наполнителя по выбранному типу фасада
        {
            TempFrontsDataTable.Clear();
            TempFrontsDataTable = ConstFrontsDataTable;
            using (DataView DV = new DataView(TempFrontsDataTable))
            {
                DV.RowFilter = "FrontName='" + FrontName + "'";
                TempFrontsDataTable = DV.ToTable();
            }
            string filter = string.Empty;
            for (int i = 0; i < TempFrontsDataTable.Rows.Count; i++)
                filter += Convert.ToInt32(TempFrontsDataTable.Rows[i]["FrontID"]) + ",";
            if (filter.Length > 0)
            {
                filter = filter.Substring(0, filter.Length - 1);
                filter = "FrontID IN (" + filter + ")";
            }
            else
                filter = "FrontID <> - 1";

            TempFrontsConfigDataTable.Clear();
            TempFrontsConfigDataTable = ConstFrontsConfigDataTable;
            if (bExcluzive)
                filter += " AND (Excluzive=1)";
            else
                filter += " AND (Excluzive IS NULL OR Excluzive=1)";
            using (DataView DV = new DataView(TempFrontsConfigDataTable))
            {
                DV.RowFilter = filter;
                DV.RowFilter += " AND ColorID=" + ColorID;
                DV.RowFilter += " AND TechnoProfileID=" + TechnoProfileID;
                DV.RowFilter += " AND TechnoColorID=" + TechnoColorID;

                TempFrontsConfigDataTable = DV.ToTable();

                GetPatina();
            }
        }

        public void FilterCatalogInsetTypes(string FrontName, int ColorID, int TechnoProfileID, int TechnoColorID, int PatinaID, bool bExcluzive) //фильтрует и заполняет типы наполнителя по выбранному типу фасада
        {
            TempFrontsDataTable.Clear();
            TempFrontsDataTable = ConstFrontsDataTable;
            using (DataView DV = new DataView(TempFrontsDataTable))
            {
                DV.RowFilter = "FrontName='" + FrontName + "'";
                TempFrontsDataTable = DV.ToTable();
            }
            string filter = string.Empty;
            for (int i = 0; i < TempFrontsDataTable.Rows.Count; i++)
                filter += Convert.ToInt32(TempFrontsDataTable.Rows[i]["FrontID"]) + ",";
            if (filter.Length > 0)
            {
                filter = filter.Substring(0, filter.Length - 1);
                filter = "FrontID IN (" + filter + ")";
            }
            else
                filter = "FrontID <> - 1";

            TempFrontsConfigDataTable.Clear();
            TempFrontsConfigDataTable = ConstFrontsConfigDataTable;
            if (bExcluzive)
                filter += " AND (Excluzive=1)";
            else
                filter += " AND (Excluzive IS NULL OR Excluzive=1)";
            if (PatinaID > 1000)
            {
                DataRow[] fRows = PatinaRALDataTable.Select("PatinaRALID=" + PatinaID);
                if (fRows.Count() > 0)
                    PatinaID = Convert.ToInt32(fRows[0]["PatinaID"]);
            }
            using (DataView DV = new DataView(TempFrontsConfigDataTable))
            {
                DV.RowFilter = filter;
                DV.RowFilter += " AND ColorID=" + ColorID;
                DV.RowFilter += " AND TechnoProfileID=" + TechnoProfileID;
                DV.RowFilter += " AND TechnoColorID=" + TechnoColorID;
                DV.RowFilter += " AND PatinaID=" + PatinaID;

                TempFrontsConfigDataTable = DV.ToTable();

                GetInsetTypes();
            }
        }

        public void FilterCatalogInsetColors(string FrontName, int ColorID, int TechnoProfileID, int TechnoColorID, int PatinaID, int InsetTypeID, bool bExcluzive) //фильтрует и заполняет цвета наполнителя по выбранному типу наполнителя
        {
            TempFrontsDataTable.Clear();
            TempFrontsDataTable = ConstFrontsDataTable;
            using (DataView DV = new DataView(TempFrontsDataTable))
            {
                DV.RowFilter = "FrontName='" + FrontName + "'";
                TempFrontsDataTable = DV.ToTable();
            }
            string filter = string.Empty;
            for (int i = 0; i < TempFrontsDataTable.Rows.Count; i++)
                filter += Convert.ToInt32(TempFrontsDataTable.Rows[i]["FrontID"]) + ",";
            if (filter.Length > 0)
            {
                filter = filter.Substring(0, filter.Length - 1);
                filter = "FrontID IN (" + filter + ")";
            }
            else
                filter = "FrontID <> - 1";

            TempFrontsConfigDataTable.Clear();
            TempFrontsConfigDataTable = ConstFrontsConfigDataTable;
            if (bExcluzive)
                filter += " AND (Excluzive=1)";
            else
                filter += " AND (Excluzive IS NULL OR Excluzive=1)";
            if (PatinaID > 1000)
            {
                DataRow[] fRows = PatinaRALDataTable.Select("PatinaRALID=" + PatinaID);
                if (fRows.Count() > 0)
                    PatinaID = Convert.ToInt32(fRows[0]["PatinaID"]);
            }
            using (DataView DV = new DataView(TempFrontsConfigDataTable))
            {
                DV.RowFilter = filter;
                DV.RowFilter += " AND ColorID=" + ColorID;
                DV.RowFilter += " AND TechnoProfileID=" + TechnoProfileID;
                DV.RowFilter += " AND TechnoColorID=" + TechnoColorID;
                DV.RowFilter += " AND PatinaID=" + PatinaID;
                DV.RowFilter += " AND InsetTypeID=" + InsetTypeID;

                TempFrontsConfigDataTable = DV.ToTable();

                GetInsetColors();
            }
        }

        public void FilterCatalogTechnoInsetTypes(string FrontName, int ColorID, int TechnoProfileID, int TechnoColorID, int PatinaID, int InsetTypeID, int InsetColorID, bool bExcluzive) //фильтрует и заполняет типы наполнителя по выбранному типу фасада
        {
            TempFrontsDataTable.Clear();
            TempFrontsDataTable = ConstFrontsDataTable;
            using (DataView DV = new DataView(TempFrontsDataTable))
            {
                DV.RowFilter = "FrontName='" + FrontName + "'";
                TempFrontsDataTable = DV.ToTable();
            }
            string filter = string.Empty;
            for (int i = 0; i < TempFrontsDataTable.Rows.Count; i++)
                filter += Convert.ToInt32(TempFrontsDataTable.Rows[i]["FrontID"]) + ",";
            if (filter.Length > 0)
            {
                filter = filter.Substring(0, filter.Length - 1);
                filter = "FrontID IN (" + filter + ")";
            }
            else
                filter = "FrontID <> - 1";

            TempFrontsConfigDataTable.Clear();
            TempFrontsConfigDataTable = ConstFrontsConfigDataTable;
            if (bExcluzive)
                filter += " AND (Excluzive=1)";
            else
                filter += " AND (Excluzive IS NULL OR Excluzive=1)";
            if (PatinaID > 1000)
            {
                DataRow[] fRows = PatinaRALDataTable.Select("PatinaRALID=" + PatinaID);
                if (fRows.Count() > 0)
                    PatinaID = Convert.ToInt32(fRows[0]["PatinaID"]);
            }
            using (DataView DV = new DataView(TempFrontsConfigDataTable))
            {
                DV.RowFilter = filter;
                DV.RowFilter += " AND ColorID=" + ColorID;
                DV.RowFilter += " AND TechnoProfileID=" + TechnoProfileID;
                DV.RowFilter += " AND TechnoColorID=" + TechnoColorID;
                DV.RowFilter += " AND PatinaID=" + PatinaID;
                DV.RowFilter += " AND InsetTypeID=" + InsetTypeID;
                DV.RowFilter += " AND InsetColorID=" + InsetColorID;

                TempFrontsConfigDataTable = DV.ToTable();

                GetTechnoInsetTypes();
            }
        }

        public void FilterCatalogTechnoInsetColors(string FrontName, int ColorID, int TechnoProfileID, int TechnoColorID, int PatinaID, int InsetTypeID, int InsetColorID, int TechnoInsetTypeID, bool bExcluzive) //фильтрует и заполняет цвета наполнителя по выбранному типу наполнителя
        {
            TempFrontsDataTable.Clear();
            TempFrontsDataTable = ConstFrontsDataTable;
            using (DataView DV = new DataView(TempFrontsDataTable))
            {
                DV.RowFilter = "FrontName='" + FrontName + "'";
                TempFrontsDataTable = DV.ToTable();
            }
            string filter = string.Empty;
            for (int i = 0; i < TempFrontsDataTable.Rows.Count; i++)
                filter += Convert.ToInt32(TempFrontsDataTable.Rows[i]["FrontID"]) + ",";
            if (filter.Length > 0)
            {
                filter = filter.Substring(0, filter.Length - 1);
                filter = "FrontID IN (" + filter + ")";
            }
            else
                filter = "FrontID <> - 1";

            TempFrontsConfigDataTable.Clear();
            TempFrontsConfigDataTable = ConstFrontsConfigDataTable;
            if (bExcluzive)
                filter += " AND (Excluzive=1)";
            else
                filter += " AND (Excluzive IS NULL OR Excluzive=1)";
            if (PatinaID > 1000)
            {
                DataRow[] fRows = PatinaRALDataTable.Select("PatinaRALID=" + PatinaID);
                if (fRows.Count() > 0)
                    PatinaID = Convert.ToInt32(fRows[0]["PatinaID"]);
            }
            using (DataView DV = new DataView(TempFrontsConfigDataTable))
            {
                DV.RowFilter = filter;
                DV.RowFilter += " AND ColorID=" + ColorID;
                DV.RowFilter += " AND PatinaID=" + PatinaID;
                DV.RowFilter += " AND InsetTypeID=" + InsetTypeID;
                DV.RowFilter += " AND InsetColorID=" + InsetColorID;
                DV.RowFilter += " AND TechnoProfileID=" + TechnoProfileID;
                DV.RowFilter += " AND TechnoColorID=" + TechnoColorID;
                DV.RowFilter += " AND TechnoInsetTypeID=" + TechnoInsetTypeID;
                TempFrontsConfigDataTable = DV.ToTable();

                GetTechnoInsetColors();
            }
        }

        public void FilterCatalogHeight(string FrontName, int ColorID, int TechnoProfileID, int TechnoColorID, int PatinaID,
            int InsetTypeID, int InsetColorID, int TechnoInsetTypeID, int TechnoInsetColorID)
        {
            TempFrontsDataTable.Clear();
            TempFrontsDataTable = ConstFrontsDataTable;
            using (DataView DV = new DataView(TempFrontsDataTable))
            {
                DV.RowFilter = "FrontName='" + FrontName + "'";

                TempFrontsDataTable = DV.ToTable();
            }
            string filter = string.Empty;
            for (int i = 0; i < TempFrontsDataTable.Rows.Count; i++)
                filter += Convert.ToInt32(TempFrontsDataTable.Rows[i]["FrontID"]) + ",";
            if (filter.Length > 0)
            {
                filter = filter.Substring(0, filter.Length - 1);
                filter = "FrontID IN (" + filter + ")";
            }
            else
                filter = "FrontID <> - 1";

            TempFrontsConfigDataTable.Clear();
            TempFrontsConfigDataTable = ConstFrontsConfigDataTable;

            if (PatinaID > 1000)
            {
                DataRow[] fRows = PatinaRALDataTable.Select("PatinaRALID=" + PatinaID);
                if (fRows.Count() > 0)
                    PatinaID = Convert.ToInt32(fRows[0]["PatinaID"]);
            }
            using (DataView DV = new DataView(TempFrontsConfigDataTable))
            {
                DV.RowFilter = filter;
                DV.RowFilter += " AND InsetTypeID=" + InsetTypeID;
                DV.RowFilter += " AND ColorID=" + ColorID;
                DV.RowFilter += " AND PatinaID=" + PatinaID;
                DV.RowFilter += " AND InsetColorID=" + InsetColorID;
                DV.RowFilter += " AND TechnoProfileID=" + TechnoProfileID;
                DV.RowFilter += " AND TechnoColorID=" + TechnoColorID;
                DV.RowFilter += " AND TechnoInsetTypeID=" + TechnoInsetTypeID;
                DV.RowFilter += " AND TechnoInsetColorID=" + TechnoInsetColorID;

                TempFrontsConfigDataTable = DV.ToTable();
            }

            HeightEdit.Items.Clear();
            HeightEdit.Text = "";

            GetHeight();

            if (HeightDataTable.Rows.Count > 0 && HeightDataTable.Rows[0]["Height"].ToString() != "0" && HeightDataTable.Rows[0]["Height"].ToString() != "-1")
            {
                HeightEdit.DropDownStyle = ComboBoxStyle.DropDownList;

                foreach (DataRow Row in HeightDataTable.Rows)
                    HeightEdit.Items.Add(Row["Height"].ToString());

                HeightEdit.SelectedIndex = 0;
            }
            else
            {
                HeightEdit.DropDownStyle = ComboBoxStyle.DropDown;
            }
        }

        public void FilterCatalogWidth(string FrontName, int ColorID, int TechnoProfileID, int TechnoColorID, int PatinaID,
            int InsetTypeID, int InsetColorID, int TechnoInsetTypeID, int TechnoInsetColorID, int Height)
        {
            TempFrontsDataTable.Clear();
            TempFrontsDataTable = ConstFrontsDataTable;
            using (DataView DV = new DataView(TempFrontsDataTable))
            {
                DV.RowFilter = "FrontName='" + FrontName + "'";

                TempFrontsDataTable = DV.ToTable();
            }
            string filter = string.Empty;
            for (int i = 0; i < TempFrontsDataTable.Rows.Count; i++)
                filter += Convert.ToInt32(TempFrontsDataTable.Rows[i]["FrontID"]) + ",";
            if (filter.Length > 0)
            {
                filter = filter.Substring(0, filter.Length - 1);
                filter = "FrontID IN (" + filter + ")";
            }
            else
                filter = "FrontID <> - 1";

            TempFrontsConfigDataTable.Clear();
            TempFrontsConfigDataTable = ConstFrontsConfigDataTable;

            if (PatinaID > 1000)
            {
                DataRow[] fRows = PatinaRALDataTable.Select("PatinaRALID=" + PatinaID);
                if (fRows.Count() > 0)
                    PatinaID = Convert.ToInt32(fRows[0]["PatinaID"]);
            }
            using (DataView DV = new DataView(TempFrontsConfigDataTable))
            {
                DV.RowFilter = filter;
                DV.RowFilter += " AND InsetTypeID=" + InsetTypeID;
                DV.RowFilter += " AND ColorID=" + ColorID;
                DV.RowFilter += " AND PatinaID=" + PatinaID;
                DV.RowFilter += " AND InsetColorID=" + InsetColorID;
                DV.RowFilter += " AND TechnoProfileID=" + TechnoProfileID;
                DV.RowFilter += " AND TechnoColorID=" + TechnoColorID;
                DV.RowFilter += " AND TechnoInsetTypeID=" + TechnoInsetTypeID;
                DV.RowFilter += " AND TechnoInsetColorID=" + TechnoInsetColorID;
                DV.RowFilter += " AND Height=" + Height;

                TempFrontsConfigDataTable = DV.ToTable();
            }

            WidthEdit.Items.Clear();
            WidthEdit.Text = "";

            GetWidth();

            if (WidthDataTable.Rows.Count > 0 && WidthDataTable.Rows[0]["Width"].ToString() != "0" && WidthDataTable.Rows[0]["Width"].ToString() != "-1")
            {
                WidthEdit.Enabled = true;
                WidthEdit.DropDownStyle = ComboBoxStyle.DropDownList;

                foreach (DataRow Row in WidthDataTable.Rows)
                    WidthEdit.Items.Add(Row["Width"].ToString());

                WidthEdit.SelectedIndex = 0;
            }
            else
            {
                if (HeightEdit.Items.Count > 0)
                {
                    WidthEdit.DropDownStyle = ComboBoxStyle.DropDown;
                    WidthEdit.Text = "-1";
                    WidthEdit.Enabled = false;
                }
                else
                {
                    WidthEdit.DropDownStyle = ComboBoxStyle.DropDown;
                    WidthEdit.Enabled = true;
                }
            }
        }

        public decimal GetNonStandardMargin(int FrontConfigID)
        {
            DataRow[] Rows = ConstFrontsConfigDataTable.Select("FrontConfigID = " + FrontConfigID);
            if (Rows.Count() == 0)
                return 0;
            else
                return Convert.ToDecimal(Rows[0]["ZOVNonStandMargin"]);
        }

        public int GetFrontConfigID(int FrontID, int ColorID, int PatinaID,
            int InsetTypeID, int InsetColorID, int TechnoProfileID, int TechnoColorID, int TechnoInsetTypeID, int TechnoInsetColorID,
            int Height, int Width, ref int FactoryID, ref int AreaID)
        {
            string HeightFilter = null;
            string WidthFilter = null;

            if (PatinaID > 1000)
            {
                DataRow[] fRows = PatinaRALDataTable.Select("PatinaRALID=" + PatinaID);
                if (fRows.Count() > 0)
                    PatinaID = Convert.ToInt32(fRows[0]["PatinaID"]);
            }
            DataRow[] Rows = ConstFrontsConfigDataTable.Select(
                            "FrontID = " + Convert.ToInt32(FrontID) + " AND " +
                            "InsetTypeID = " + Convert.ToInt32(InsetTypeID) + " AND " +
                            "InsetColorID = " + Convert.ToInt32(InsetColorID) + " AND " +
                            "ColorID = " + Convert.ToInt32(ColorID) + " AND " +
                            "PatinaID = " + Convert.ToInt32(PatinaID) + " AND " +
                            "TechnoProfileID = " + Convert.ToInt32(TechnoProfileID) + " AND " +
                            "TechnoColorID = " + Convert.ToInt32(TechnoColorID) + " AND " +
                            "TechnoInsetTypeID = " + Convert.ToInt32(TechnoInsetTypeID) + " AND " +
                            "TechnoInsetColorID = " + Convert.ToInt32(TechnoInsetColorID));

            if (HasHeightParameter(Rows, Height))
                HeightFilter = " AND Height = " + Height;
            else
                HeightFilter = " AND Height = 0";

            if (Height == -1)
                HeightFilter = " AND Height = -1";

            if (HasWidthParameter(Rows, Width))
                WidthFilter = " AND Width = " + Width;
            else
                WidthFilter = " AND Width = 0";

            if (Width == -1)
                WidthFilter = " AND Width = -1";

            Rows = ConstFrontsConfigDataTable.Select(
                "FrontID = " + Convert.ToInt32(FrontID) + " AND " +
                "InsetTypeID = " + Convert.ToInt32(InsetTypeID) + " AND " +
                "ColorID = " + Convert.ToInt32(ColorID) + " AND " +
                "PatinaID = " + Convert.ToInt32(PatinaID) + " AND " +
                "InsetColorID = " + Convert.ToInt32(InsetColorID) + " AND " +
                "TechnoProfileID = " + Convert.ToInt32(TechnoProfileID) + " AND " +
                "TechnoColorID = " + Convert.ToInt32(TechnoColorID) + " AND " +
                "TechnoInsetTypeID = " + Convert.ToInt32(TechnoInsetTypeID) + " AND " +
                "TechnoInsetColorID = " + Convert.ToInt32(TechnoInsetColorID) +
                HeightFilter + WidthFilter);

            if (Rows.Count() < 1)
            {
                MessageBox.Show("Ошибка конфигурации фасада. Проверьте каталог " + Rows.Count().ToString());
                return -1;
            }

            FactoryID = Convert.ToInt32(Rows[0]["FactoryID"]);
            AreaID = Convert.ToInt32(Rows[0]["AreaID"]);

            if (Rows[0]["Excluzive"] != DBNull.Value && Convert.ToInt32(Rows[0]["Excluzive"]) == 0)
            {
                MessageBox.Show("Конфигурация является эксклюзивом и недоступна другим клиентам");
                return -1;
            }
            return Convert.ToInt32(Rows[0]["FrontConfigID"]);
        }

        public int GetFrontConfigID(int FrontID, int ColorID, int PatinaID,
            int InsetTypeID, int InsetColorID, int TechnoProfileID, int TechnoColorID, int TechnoInsetTypeID, int TechnoInsetColorID,
            int Height, int Width)
        {
            string HeightFilter = null;
            string WidthFilter = null;

            if (PatinaID > 1000)
            {
                DataRow[] fRows = PatinaRALDataTable.Select("PatinaRALID=" + PatinaID);
                if (fRows.Count() > 0)
                    PatinaID = Convert.ToInt32(fRows[0]["PatinaID"]);
            }
            DataRow[] Rows = ConstFrontsConfigDataTable.Select(
                            "FrontID = " + Convert.ToInt32(FrontID) + " AND " +
                            "InsetTypeID = " + Convert.ToInt32(InsetTypeID) + " AND " +
                            "InsetColorID = " + Convert.ToInt32(InsetColorID) + " AND " +
                            "ColorID = " + Convert.ToInt32(ColorID) + " AND " +
                            "PatinaID = " + Convert.ToInt32(PatinaID) + " AND " +
                            "TechnoProfileID = " + Convert.ToInt32(TechnoProfileID) + " AND " +
                            "TechnoColorID = " + Convert.ToInt32(TechnoColorID) + " AND " +
                            "TechnoInsetTypeID = " + Convert.ToInt32(TechnoInsetTypeID) + " AND " +
                            "TechnoInsetColorID = " + Convert.ToInt32(TechnoInsetColorID));

            if (HasHeightParameter(Rows, Height))
                HeightFilter = " AND Height = " + Height;
            else
                HeightFilter = " AND Height = 0";

            if (Height == -1)
                HeightFilter = " AND Height = -1";

            if (HasWidthParameter(Rows, Width))
                WidthFilter = " AND Width = " + Width;
            else
                WidthFilter = " AND Width = 0";

            if (Width == -1)
                WidthFilter = " AND Width = -1";

            Rows = ConstFrontsConfigDataTable.Select(
                "FrontID = " + Convert.ToInt32(FrontID) + " AND " +
                "InsetTypeID = " + Convert.ToInt32(InsetTypeID) + " AND " +
                "ColorID = " + Convert.ToInt32(ColorID) + " AND " +
                "PatinaID = " + Convert.ToInt32(PatinaID) + " AND " +
                "InsetColorID = " + Convert.ToInt32(InsetColorID) + " AND " +
                "TechnoProfileID = " + Convert.ToInt32(TechnoProfileID) + " AND " +
                "TechnoColorID = " + Convert.ToInt32(TechnoColorID) + " AND " +
                "TechnoInsetTypeID = " + Convert.ToInt32(TechnoInsetTypeID) + " AND " +
                "TechnoInsetColorID = " + Convert.ToInt32(TechnoInsetColorID) +
                HeightFilter + WidthFilter);

            if (Rows.Count() < 1)
            {
                MessageBox.Show("Ошибка конфигурации фасада. Проверьте каталог " + Rows.Count().ToString());
                return -1;
            }
            
            if (Rows[0]["Excluzive"] != DBNull.Value && Convert.ToInt32(Rows[0]["Excluzive"]) == 0)
            {
                MessageBox.Show("Конфигурация является эксклюзивом и недоступна другим клиентам");
                return -1;
            }
            return Convert.ToInt32(Rows[0]["FrontConfigID"]);
        }

        public int GetFrontConfigID(string FrontName, int ColorID, int PatinaID,
            int InsetTypeID, int InsetColorID, int TechnoProfileID, int TechnoColorID, int TechnoInsetTypeID, int TechnoInsetColorID,
            int Height, int Width, ref int FrontID, ref int FactoryID, ref int AreaID)
        {
            TempFrontsDataTable.Clear();
            TempFrontsDataTable = ConstFrontsDataTable;
            using (DataView DV = new DataView(TempFrontsDataTable))
            {
                DV.RowFilter = "FrontName='" + FrontName + "'";

                TempFrontsDataTable = DV.ToTable();
            }
            string filter = string.Empty;
            for (int i = 0; i < TempFrontsDataTable.Rows.Count; i++)
                filter += Convert.ToInt32(TempFrontsDataTable.Rows[i]["FrontID"]) + ",";
            if (filter.Length > 0)
            {
                filter = filter.Substring(0, filter.Length - 1);
                filter = "FrontID IN (" + filter + ")";
            }
            else
                filter = "FrontID <> - 1";

            string HeightFilter = null;
            string WidthFilter = null;

            if (PatinaID > 1000)
            {
                DataRow[] fRows = PatinaRALDataTable.Select("PatinaRALID=" + PatinaID);
                if (fRows.Count() > 0)
                    PatinaID = Convert.ToInt32(fRows[0]["PatinaID"]);
            }
            DataRow[] Rows = ConstFrontsConfigDataTable.Select(
                            filter + " AND " +
                            "InsetTypeID = " + Convert.ToInt32(InsetTypeID) + " AND " +
                            "InsetColorID = " + Convert.ToInt32(InsetColorID) + " AND " +
                            "ColorID = " + Convert.ToInt32(ColorID) + " AND " +
                            "PatinaID = " + Convert.ToInt32(PatinaID) + " AND " +
                            "TechnoProfileID = " + Convert.ToInt32(TechnoProfileID) + " AND " +
                            "TechnoColorID = " + Convert.ToInt32(TechnoColorID) + " AND " +
                            "TechnoInsetTypeID = " + Convert.ToInt32(TechnoInsetTypeID) + " AND " +
                            "TechnoInsetColorID = " + Convert.ToInt32(TechnoInsetColorID));

            if (HasHeightParameter(Rows, Height))
                HeightFilter = " AND Height = " + Height;
            else
                HeightFilter = " AND Height = 0";

            if (Height == -1)
                HeightFilter = " AND Height = -1";

            if (HasWidthParameter(Rows, Width))
                WidthFilter = " AND Width = " + Width;
            else
                WidthFilter = " AND Width = 0";

            if (Width == -1)
                WidthFilter = " AND Width = -1";

            Rows = ConstFrontsConfigDataTable.Select(
                filter + " AND " +
                "InsetTypeID = " + Convert.ToInt32(InsetTypeID) + " AND " +
                "ColorID = " + Convert.ToInt32(ColorID) + " AND " +
                "PatinaID = " + Convert.ToInt32(PatinaID) + " AND " +
                "InsetColorID = " + Convert.ToInt32(InsetColorID) + " AND " +
                "TechnoProfileID = " + Convert.ToInt32(TechnoProfileID) + " AND " +
                "TechnoColorID = " + Convert.ToInt32(TechnoColorID) + " AND " +
                "TechnoInsetTypeID = " + Convert.ToInt32(TechnoInsetTypeID) + " AND " +
                "TechnoInsetColorID = " + Convert.ToInt32(TechnoInsetColorID) +
                HeightFilter + WidthFilter);

            if (Rows.Count() < 1)
            {
                MessageBox.Show("Ошибка конфигурации фасада. Проверьте каталог " + Rows.Count().ToString());
                return -1;
            }

            if (Rows[0]["Excluzive"] != DBNull.Value && Convert.ToInt32(Rows[0]["Excluzive"]) == 0)
            {
                MessageBox.Show("Конфигурация является эксклюзивом и недоступна другим клиентам");
                return -1;
            }
            DataTable DT = new DataTable();
            using (DataView DV = new DataView(ConstFrontsConfigDataTable,
                filter + " AND(Excluzive IS NULL OR Excluzive=1) " + " AND " +
                "InsetTypeID = " + Convert.ToInt32(InsetTypeID) + " AND " +
                "ColorID = " + Convert.ToInt32(ColorID) + " AND " +
                "PatinaID = " + Convert.ToInt32(PatinaID) + " AND " +
                "InsetColorID = " + Convert.ToInt32(InsetColorID) + " AND " +
                "TechnoColorID = " + Convert.ToInt32(TechnoColorID) + " AND " +
                "TechnoProfileID = " + Convert.ToInt32(TechnoProfileID) + " AND " +
                "TechnoInsetTypeID = " + Convert.ToInt32(TechnoInsetTypeID) + " AND " +
                "TechnoInsetColorID = " + Convert.ToInt32(TechnoInsetColorID) +
                HeightFilter + WidthFilter, string.Empty, DataViewRowState.CurrentRows))
            {
                DT = DV.ToTable(true, new string[] { "FrontID" });
            }

            if (DT.Rows.Count > 1)
            {
                MessageBox.Show("Ошибка конфигурации фасада. Проверьте каталог " + Rows.Count().ToString());
                return -1;
            }

            FrontID = Convert.ToInt32(DT.Rows[0]["FrontID"]);
            FactoryID = Convert.ToInt32(Rows[0]["FactoryID"]);
            AreaID = Convert.ToInt32(Rows[0]["AreaID"]);

            return Convert.ToInt32(Rows[0]["FrontConfigID"]);
        }
    }
}
