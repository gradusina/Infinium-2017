﻿
using System;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Windows.Forms;

namespace Infinium.Modules.Marketing.NewOrders
{
    public class DecorCatalogOrder
    {
        private ComponentFactory.Krypton.Toolkit.KryptonComboBox LengthEdit;
        private ComponentFactory.Krypton.Toolkit.KryptonComboBox HeightEdit;
        private ComponentFactory.Krypton.Toolkit.KryptonComboBox WidthEdit;

        public int DecorProductsCount = 0;

        public DataTable ItemsDataTable = null;
        public DataTable ColorsDataTable = null;
        public DataTable PatinaDataTable = null;
        public DataTable PatinaRALDataTable = null;
        public DataTable InsetTypesDataTable = null;
        public DataTable InsetColorsDataTable = null;

        public DataTable DecorProductsDataTable = null;
        public DataTable DecorDataTable = null;
        public DataTable DecorConfigDataTable = null;
        public DataTable DecorParametersDataTable = null;

        public DataTable TempItemsDataTable = null;
        public DataTable ItemColorsDataTable = null;
        public DataTable ItemPatinaDataTable = null;
        public DataTable ItemLengthDataTable = null;
        public DataTable ItemHeightDataTable = null;
        public DataTable ItemWidthDataTable = null;
        public DataTable ItemInsetTypesDataTable = null;
        public DataTable ItemInsetColorsDataTable = null;
        public DataTable TechStoreDataTable = null;
        public DataTable CurrencyTypesDT = null;

        public BindingSource DecorProductsBindingSource = null;
        public BindingSource DecorBindingSource = null;
        public BindingSource ItemsBindingSource = null;
        public BindingSource ItemColorsBindingSource = null;
        public BindingSource ItemPatinaBindingSource = null;
        public BindingSource ItemLengthBindingSource = null;
        public BindingSource ItemHeightBindingSource = null;
        public BindingSource ItemWidthBindingSource = null;
        public BindingSource ColorsBindingSource = null;
        public BindingSource PatinaBindingSource = null;
        public BindingSource ItemInsetTypesBindingSource = null;
        public BindingSource ItemInsetColorsBindingSource = null;

        public String DecorProductsBindingSourceDisplayMember = null;
        public String ItemsBindingSourceDisplayMember = null;
        public String ItemColorsBindingSourceDisplayMember = null;
        public String ItemPatinaBindingSourceDisplayMember = null;
        public String ItemLengthBindingSourceDisplayMember = null;
        public String ItemHeightBindingSourceDisplayMember = null;
        public String ItemWidthBindingSourceDisplayMember = null;

        public String DecorProductsBindingSourceValueMember = null;
        public String ItemsBindingSourceValueMember = null;
        public String ItemColorsBindingSourceValueMember = null;
        public String ItemPatinaBindingSourceValueMember = null;

        public DecorCatalogOrder()
        {

        }

        private void Create()
        {
            DecorProductsDataTable = new DataTable();
            DecorParametersDataTable = new DataTable();
            DecorConfigDataTable = new DataTable();
            ItemColorsDataTable = new DataTable();
            ItemPatinaDataTable = new DataTable();
            ColorsDataTable = new DataTable();
            PatinaDataTable = new DataTable();
            ItemLengthDataTable = new DataTable();
            ItemHeightDataTable = new DataTable();
            ItemWidthDataTable = new DataTable();

            DecorProductsBindingSource = new BindingSource();
            DecorBindingSource = new BindingSource();
            ItemsBindingSource = new BindingSource();
            ItemLengthBindingSource = new BindingSource();
            ItemHeightBindingSource = new BindingSource();
            ItemWidthBindingSource = new BindingSource();
            ItemColorsBindingSource = new BindingSource();
            ItemPatinaBindingSource = new BindingSource();
            ColorsBindingSource = new BindingSource();
            PatinaBindingSource = new BindingSource();
        }

        private void GetColorsDT()
        {
            ColorsDataTable = new DataTable();
            ColorsDataTable.Columns.Add(new DataColumn("ColorID", Type.GetType("System.Int64")));
            ColorsDataTable.Columns.Add(new DataColumn("ColorName", Type.GetType("System.String")));
            ColorsDataTable.Columns.Add(new DataColumn("Cvet", Type.GetType("System.String")));
            ColorsDataTable.Columns.Add(new DataColumn("Excluzive", Type.GetType("System.Int32")));
            string SelectCommand = @"SELECT TechStoreID, TechStoreName, Cvet FROM TechStore
                WHERE TechStoreSubGroupID IN (SELECT TechStoreSubGroupID FROM TechStoreSubGroups WHERE (TechStoreGroupID = 11))
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
                        NewRow["Cvet"] = "000";
                        ColorsDataTable.Rows.Add(NewRow);
                    }
                    {
                        DataRow NewRow = ColorsDataTable.NewRow();
                        NewRow["ColorID"] = 0;
                        NewRow["ColorName"] = "на выбор";
                        NewRow["Cvet"] = "0000000";
                        ColorsDataTable.Rows.Add(NewRow);
                    }
                    for (int i = 0; i < DT.Rows.Count; i++)
                    {
                        DataRow NewRow = ColorsDataTable.NewRow();
                        NewRow["ColorID"] = Convert.ToInt64(DT.Rows[i]["TechStoreID"]);
                        NewRow["ColorName"] = DT.Rows[i]["TechStoreName"].ToString();
                        NewRow["Cvet"] = DT.Rows[i]["Cvet"].ToString();
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
                            NewRow["Cvet"] = DT.Rows[i]["InsetColorName"].ToString();
                            ColorsDataTable.Rows.Add(NewRow);
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

        private void Fill(bool isNewOrders)
        {
            string SelectCommand = string.Empty;

            if (isNewOrders)
            SelectCommand = @"SELECT ProductID, ProductName, MeasureID, ReportParam FROM DecorProducts" +
                " WHERE ProductID IN (SELECT ProductID FROM DecorConfig WHERE (Enabled = 1 AND AccountingName IS NOT NULL AND InvNumber IS NOT NULL)) ORDER BY ProductName ASC";
            else
                SelectCommand = @"SELECT ProductID, ProductName, MeasureID, ReportParam FROM DecorProducts" +
                                " WHERE ProductID IN (SELECT ProductID FROM DecorConfig WHERE (AccountingName IS NOT NULL AND InvNumber IS NOT NULL)) ORDER BY ProductName ASC";

            DecorProductsDataTable = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(DecorProductsDataTable);
                DecorProductsDataTable.Columns.Add(new DataColumn("Excluzive", Type.GetType("System.Int32")));
            }
            DecorDataTable = new DataTable();

            if (isNewOrders)
                SelectCommand = @"SELECT DISTINCT TechStore.TechStoreID AS DecorID, TechStore.TechStoreName AS Name, DecorConfig.ProductID FROM TechStore 
                INNER JOIN DecorConfig ON TechStore.TechStoreID = DecorConfig.DecorID AND Enabled = 1 AND AccountingName IS NOT NULL AND InvNumber IS NOT NULL ORDER BY TechStoreName";
            else
                SelectCommand = @"SELECT DISTINCT TechStore.TechStoreID AS DecorID, TechStore.TechStoreName AS Name, DecorConfig.ProductID FROM TechStore 
                INNER JOIN DecorConfig ON TechStore.TechStoreID = DecorConfig.DecorID AND AccountingName IS NOT NULL AND InvNumber IS NOT NULL ORDER BY TechStoreName";

            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(DecorDataTable);
                DecorDataTable.Columns.Add(new DataColumn("Excluzive", Type.GetType("System.Int32")));
            }
            GetColorsDT();
            GetInsetColorsDT();

            CurrencyTypesDT = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM CurrencyTypes", ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(CurrencyTypesDT);
            }

            InsetTypesDataTable = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM InsetTypes",
                ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(InsetTypesDataTable);
            }
            PatinaDataTable = new DataTable();
            SelectCommand = @"SELECT * FROM Patina ORDER BY PatinaName";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(PatinaDataTable);
                PatinaDataTable.Columns.Add(new DataColumn("Excluzive", Type.GetType("System.Int32")));
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
            DecorParametersDataTable = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM DecorParameters", ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(DecorParametersDataTable);
            }

            TempItemsDataTable = DecorDataTable.Clone();

            using (DataView DV = new DataView(DecorDataTable))
            {
                ItemsDataTable = DV.ToTable(true, new string[] { "Name" });
            }
            DecorProductsCount = DecorProductsDataTable.Rows.Count;

            ItemColorsDataTable = ColorsDataTable.Clone();
            ItemPatinaDataTable = PatinaDataTable.Clone();
            ItemInsetTypesDataTable = InsetTypesDataTable.Clone();
            ItemInsetColorsDataTable = InsetColorsDataTable.Clone();

            if (isNewOrders)
                DecorConfigDataTable = TablesManager.DecorConfigDataTable;
            else
                DecorConfigDataTable = TablesManager.DecorConfigDataTableAll;
            DecorConfigDataTable.Columns.Add(new DataColumn("Excluzive", Type.GetType("System.Boolean")));
            TechStoreDataTable = new DataTable();
            TechStoreDataTable = TablesManager.TechStoreDataTable;
        }

        private void Binding()
        {
            DecorProductsBindingSource.DataSource = DecorProductsDataTable;
            DecorBindingSource.DataSource = DecorDataTable;
            ItemsBindingSource.DataSource = ItemsDataTable;
            ItemLengthBindingSource.DataSource = ItemLengthDataTable;
            ItemHeightBindingSource.DataSource = ItemHeightDataTable;
            ItemWidthBindingSource.DataSource = ItemWidthDataTable;
            ItemColorsBindingSource.DataSource = ItemColorsDataTable;
            ItemPatinaBindingSource.DataSource = ItemPatinaDataTable;
            ColorsBindingSource.DataSource = ColorsDataTable;
            PatinaBindingSource.DataSource = PatinaDataTable;
            ItemInsetTypesBindingSource = new BindingSource()
            {
                DataSource = ItemInsetTypesDataTable
            };
            ItemInsetColorsBindingSource = new BindingSource()
            {
                DataSource = ItemInsetColorsDataTable
            };
            DecorProductsBindingSourceDisplayMember = "ProductName";
            ItemsBindingSourceDisplayMember = "Name";
            ItemColorsBindingSourceDisplayMember = "ColorName";
            ItemPatinaBindingSourceDisplayMember = "PatinaName";
            ItemLengthBindingSourceDisplayMember = "Length";
            ItemHeightBindingSourceDisplayMember = "Height";
            ItemWidthBindingSourceDisplayMember = "Width";

            DecorProductsBindingSourceValueMember = "ProductID";
            ItemsBindingSourceValueMember = "Name";
            ItemColorsBindingSourceValueMember = "ColorID";
            ItemPatinaBindingSourceValueMember = "PatinaID";
        }

        public void Initialize(bool isNewOrders)
        {
            Create();
            Fill(isNewOrders);
            Binding();
            CreateLengthDataTable();
            CreateHeightDataTable();
            CreateWidthDataTable();
        }

        //External
        public bool HasParameter(int ProductID, String Parameter)
        {
            DataRow[] Rows = DecorParametersDataTable.Select("ProductID = " + ProductID);

            return Convert.ToBoolean(Rows[0][Parameter]);
        }

        private bool HasColorParameter(DataRow[] Rows, int ColorID)
        {
            foreach (DataRow Row in Rows)
            {
                if (Row["ColorID"].ToString() == ColorID.ToString())
                    return true;
            }

            return false;
        }

        private bool HasPatinaParameter(DataRow[] Rows, int ColorID)
        {
            foreach (DataRow Row in Rows)
            {
                if (Row["PatinaID"].ToString() == ColorID.ToString())
                    return true;
            }

            return false;
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

        private bool HasLengthParameter(DataRow[] Rows, int Length)
        {
            foreach (DataRow Row in Rows)
            {
                if (Row["Length"].ToString() == Length.ToString())
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

        public string GetProductName(int ProductId)
        {
            return DecorProductsDataTable.Select("ProductId = " + ProductId)[0]["ProductName"].ToString();
        }

        public string GetItemName(int DecorID)
        {
            return DecorDataTable.Select("DecorID = " + DecorID)[0]["Name"].ToString();
        }

        public int GetDecorConfigID(int ProductID, int DecorID, int ColorID, int PatinaID, int InsetTypeID, int InsetColorID, int Length, int Height, int Width, ref int FactoryID, ref int AreaID)
        {
            string LengthFilter = null;
            string HeightFilter = null;
            string WidthFilter = null;
            string ColorFilter = null;
            string PatinaFilter = null;
            string InsetTypeFilter = null;
            string InsetColorFilter = null;

            if (PatinaID > 1000)
            {
                DataRow[] fRows = PatinaRALDataTable.Select("PatinaRALID=" + PatinaID);
                if (fRows.Count() > 0)
                    PatinaID = Convert.ToInt32(fRows[0]["PatinaID"]);
            }
            DataRow[] Rows = DecorConfigDataTable.Select(
                            "ProductID = " + Convert.ToInt32(ProductID) + " AND " +
                            "DecorID = " + Convert.ToInt32(DecorID));


            ColorFilter = " AND ColorID = " + ColorID;
            PatinaFilter = " AND PatinaID = " + PatinaID;
            InsetTypeFilter = " AND InsetTypeID = " + InsetTypeID;
            InsetColorFilter = " AND InsetColorID = " + InsetColorID;
            //if (HasColorParameter(Rows, ColorID))
            //    ColorFilter = " AND ColorID = " + ColorID;
            //else
            //    ColorFilter = " AND ColorID = -1";

            //if (HasPatinaParameter(Rows, PatinaID))
            //    PatinaFilter = " AND PatinaID = " + PatinaID;
            //else
            //    PatinaFilter = " AND PatinaID = -1";

            if (HasLengthParameter(Rows, Length))
                LengthFilter = " AND Length = " + Length;
            else
                LengthFilter = " AND Length = 0";

            if (Length == -1)
                LengthFilter = " AND Length = -1";

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

            Rows = DecorConfigDataTable.Select("(Excluzive IS NULL OR Excluzive=1) AND " +
                                "ProductID = " + Convert.ToInt32(ProductID) + " AND " +
                                "DecorID = " + Convert.ToInt32(DecorID) +
                                ColorFilter + PatinaFilter + InsetTypeFilter + InsetColorFilter + LengthFilter + HeightFilter + WidthFilter);

            if (Rows.Count() < 1 || Rows.Count() > 1)
            {
                MessageBox.Show("Ошибка конфигурации декора. Проверьте каталог\r\n" +
                    GetDecorName(Convert.ToInt32(DecorID)) + GetColorName(Convert.ToInt32(ColorID)));
                return -1;
            }

            FactoryID = Convert.ToInt32(Rows[0]["FactoryID"]);
            AreaID = Convert.ToInt32(Rows[0]["AreaID"]);

            if (Rows[0]["Excluzive"] != DBNull.Value && Convert.ToInt32(Rows[0]["Excluzive"]) == 0)
            {
                MessageBox.Show("Конфигурация является эксклюзивом и недоступна другим клиентам");
                return -1;
            }
            return Convert.ToInt32(Rows[0]["DecorConfigID"]);
        }

        public int GetDecorConfigID(int ProductID, string Name, int ColorID, int PatinaID, int InsetTypeID, int InsetColorID, int Length, int Height, int Width, ref int DecorID, ref int FactoryID, ref int AreaID)
        {
            TempItemsDataTable.Clear();
            TempItemsDataTable = DecorDataTable;
            using (DataView DV = new DataView(TempItemsDataTable))
            {
                DV.RowFilter = "Name='" + Name + "'";

                TempItemsDataTable = DV.ToTable();
            }
            string filter = string.Empty;
            for (int i = 0; i < TempItemsDataTable.Rows.Count; i++)
                filter += Convert.ToInt32(TempItemsDataTable.Rows[i]["DecorID"]) + ",";
            if (filter.Length > 0)
            {
                filter = filter.Substring(0, filter.Length - 1);
                filter = "DecorID IN (" + filter + ")";
            }
            else
                filter = "DecorID <> - 1";

            string LengthFilter = null;
            string HeightFilter = null;
            string WidthFilter = null;
            string ColorFilter = null;
            string PatinaFilter = null;
            string InsetTypeFilter = null;
            string InsetColorFilter = null;

            if (PatinaID > 1000)
            {
                DataRow[] fRows = PatinaRALDataTable.Select("PatinaRALID=" + PatinaID);
                if (fRows.Count() > 0)
                    PatinaID = Convert.ToInt32(fRows[0]["PatinaID"]);
            }
            DataRow[] Rows = DecorConfigDataTable.Select(
                            "ProductID = " + Convert.ToInt32(ProductID) + " AND " + filter);


            ColorFilter = " AND ColorID = " + ColorID;
            PatinaFilter = " AND PatinaID = " + PatinaID;
            InsetTypeFilter = " AND InsetTypeID = " + InsetTypeID;
            InsetColorFilter = " AND InsetColorID = " + InsetColorID;
            //if (HasColorParameter(Rows, ColorID))
            //    ColorFilter = " AND ColorID = " + ColorID;
            //else
            //    ColorFilter = " AND ColorID = -1";

            //if (HasPatinaParameter(Rows, PatinaID))
            //    PatinaFilter = " AND PatinaID = " + PatinaID;
            //else
            //    PatinaFilter = " AND PatinaID = -1";

            if (HasLengthParameter(Rows, Length))
                LengthFilter = " AND Length = " + Length;
            else
                LengthFilter = " AND Length = 0";

            if (Length == -1)
                LengthFilter = " AND Length = -1";

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


            Rows = DecorConfigDataTable.Select("(Excluzive IS NULL OR Excluzive=1) AND " +
                                "ProductID = " + Convert.ToInt32(ProductID) + " AND " + filter +
                                ColorFilter + PatinaFilter + InsetTypeFilter + InsetColorFilter + LengthFilter + HeightFilter + WidthFilter);

            if (Rows.Count() < 1 || Rows.Count() > 1)
            {
                MessageBox.Show("Ошибка конфигурации декора. Проверьте каталог\r\n" +
                    GetDecorName(Convert.ToInt32(DecorID)) + GetColorName(Convert.ToInt32(ColorID)));
                return -1;
            }

            FactoryID = Convert.ToInt32(Rows[0]["FactoryID"]);
            AreaID = Convert.ToInt32(Rows[0]["AreaID"]);
            DecorID = Convert.ToInt32(Rows[0]["DecorID"]);

            if (Rows[0]["Excluzive"] != DBNull.Value && Convert.ToInt32(Rows[0]["Excluzive"]) == 0)
            {
                MessageBox.Show("Конфигурация является эксклюзивом и недоступна другим клиентам");
                return -1;
            }
            return Convert.ToInt32(Rows[0]["DecorConfigID"]);
        }

        private string GetDecorName(int DecorID)
        {
            string Name = string.Empty;
            try
            {
                DataRow[] Rows = DecorDataTable.Select("DecorID = " + DecorID);
                Name = Rows[0]["Name"].ToString();
            }
            catch
            {
                return string.Empty;
            }
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

        public string PatinaDisplayName(int PatinaID)
        {
            DataRow[] rows = PatinaDataTable.Select("PatinaID = " + PatinaID);
            if (rows.Count() > 0)
                return rows[0]["DisplayName"].ToString();
            return string.Empty;
        }

        public string GetInsetTypeName(int InsetTypeID)
        {
            string InsetTypeName = string.Empty;
            try
            {
                DataRow[] Rows = InsetTypesDataTable.Select("InsetTypeID = " + InsetTypeID);
                InsetTypeName = Rows[0]["InsetTypeName"].ToString();
            }
            catch
            {
                return string.Empty;
            }
            return InsetTypeName;
        }

        public string GetInsetColorName(int InsetColorID)
        {
            string InsetColorName = string.Empty;
            try
            {
                DataRow[] Rows = InsetColorsDataTable.Select("InsetColorID = " + InsetColorID);
                InsetColorName = Rows[0]["InsetColorName"].ToString();
            }
            catch
            {
                return string.Empty;
            }
            return InsetColorName;
        }

        private void CreateLengthDataTable()
        {
            ItemLengthDataTable.Columns.Add(new DataColumn("Length", Type.GetType("System.Int32")));
        }

        private void CreateHeightDataTable()
        {
            ItemHeightDataTable.Columns.Add(new DataColumn("Height", Type.GetType("System.Int32")));
        }

        private void CreateWidthDataTable()
        {
            ItemWidthDataTable.Columns.Add(new DataColumn("Width", Type.GetType("System.Int32")));
        }

        public void FilterProducts(bool bExcluzive)
        {
            ItemsDataTable.Clear();
            ItemColorsDataTable.Clear();
            ItemPatinaDataTable.Clear();
            ItemInsetTypesDataTable.Clear();
            ItemInsetColorsDataTable.Clear();
            ItemLengthDataTable.Clear();
            ItemHeightDataTable.Clear();
            ItemWidthDataTable.Clear();
            string RowFilter = "Excluzive=1";
            if (!bExcluzive)
                RowFilter = "Excluzive IS NULL OR Excluzive<>0";

            DecorProductsBindingSource.Filter = RowFilter;
            DecorProductsBindingSource.MoveFirst();
        }

        public void FilterItems(int ProductID, bool bExcluzive)
        {
            TempItemsDataTable.Clear();
            TempItemsDataTable = DecorDataTable;
            string RowFilter = "ProductID=" + ProductID;
            using (DataView DV = new DataView(TempItemsDataTable))
            {
                if (bExcluzive)
                    RowFilter += " AND (Excluzive=1)";
                else
                    RowFilter += " AND (Excluzive IS NULL OR Excluzive=1)";
                DV.RowFilter = RowFilter;
                TempItemsDataTable = DV.ToTable(true, new string[] { "Name" });
            }

            ItemsDataTable.Clear();
            for (int d = 0; d < TempItemsDataTable.Rows.Count; d++)
            {
                DataRow NewRow = ItemsDataTable.NewRow();
                NewRow["Name"] = TempItemsDataTable.Rows[d]["Name"].ToString();
                ItemsDataTable.Rows.Add(NewRow);
            }
            ItemsDataTable.DefaultView.Sort = "Name ASC";
        }

        public bool FilterColors(string Name, bool bExcluzive)
        {
            TempItemsDataTable.Clear();
            TempItemsDataTable = DecorDataTable;
            using (DataView DV = new DataView(TempItemsDataTable))
            {
                DV.RowFilter = "Name='" + Name + "'";

                TempItemsDataTable = DV.ToTable();
            }
            string filter = string.Empty;
            for (int i = 0; i < TempItemsDataTable.Rows.Count; i++)
                filter += Convert.ToInt32(TempItemsDataTable.Rows[i]["DecorID"]) + ",";
            if (filter.Length > 0)
            {
                filter = filter.Substring(0, filter.Length - 1);
                filter = "DecorID IN (" + filter + ")";
            }
            else
                filter = "DecorID <> - 1";

            if (bExcluzive)
                filter += " AND (Excluzive=1)";
            else
                filter += " AND (Excluzive IS NULL OR Excluzive=1)";
            ItemColorsDataTable.Clear();
            ItemColorsDataTable.AcceptChanges();

            DataRow[] DCR = DecorConfigDataTable.Select(filter);


            for (int d = 0; d < DCR.Count(); d++)
            {
                if (d == 0)
                {
                    DataRow NewRow = ItemColorsDataTable.NewRow();
                    NewRow["ColorID"] = (DCR[d]["ColorID"]);
                    NewRow["ColorName"] = GetColorName(Convert.ToInt32(DCR[d]["ColorID"]));
                    ItemColorsDataTable.Rows.Add(NewRow);
                    continue;
                }


                for (int i = 0; i < ItemColorsDataTable.Rows.Count; i++)
                {
                    if (ItemColorsDataTable.Rows[i]["ColorID"].ToString() == DCR[d]["ColorID"].ToString())
                        break;

                    if (i == ItemColorsDataTable.Rows.Count - 1)
                    {
                        DataRow NewRow = ItemColorsDataTable.NewRow();
                        NewRow["ColorID"] = (DCR[d]["ColorID"]);
                        NewRow["ColorName"] = GetColorName(Convert.ToInt32(DCR[d]["ColorID"]));
                        ItemColorsDataTable.Rows.Add(NewRow);
                        break;
                    }
                }
            }

            ItemColorsDataTable.DefaultView.Sort = "ColorName ASC";
            ItemColorsBindingSource.MoveFirst();
            if (ItemColorsDataTable.Rows.Count == 1 && ItemColorsDataTable.Rows[0]["ColorID"].ToString() == "-1")
                return false;

            return true;

        }

        public bool FilterPatina(string Name, int ColorID, bool bExcluzive)
        {
            TempItemsDataTable.Clear();
            TempItemsDataTable = DecorDataTable;
            using (DataView DV = new DataView(TempItemsDataTable))
            {
                DV.RowFilter = "Name='" + Name + "'";

                TempItemsDataTable = DV.ToTable();
            }
            string filter = string.Empty;
            for (int i = 0; i < TempItemsDataTable.Rows.Count; i++)
                filter += Convert.ToInt32(TempItemsDataTable.Rows[i]["DecorID"]) + ",";
            if (filter.Length > 0)
            {
                filter = filter.Substring(0, filter.Length - 1);
                filter = "DecorID IN (" + filter + ")";
            }
            else
                filter = "DecorID <> - 1";

            if (bExcluzive)
                filter += " AND (Excluzive=1)";
            else
                filter += " AND (Excluzive IS NULL OR Excluzive=1)";
            ItemPatinaDataTable.Clear();
            ItemPatinaDataTable.AcceptChanges();

            DataRow[] DCR = DecorConfigDataTable.Select(filter + " AND ColorID = " + ColorID.ToString());

            for (int d = 0; d < DCR.Count(); d++)
            {
                if (d == 0)
                {
                    DataRow NewRow = ItemPatinaDataTable.NewRow();
                    NewRow["PatinaID"] = (DCR[d]["PatinaID"]);
                    NewRow["PatinaName"] = GetPatinaName(Convert.ToInt32(DCR[d]["PatinaID"]));
                    ItemPatinaDataTable.Rows.Add(NewRow);
                    continue;
                }

                for (int i = 0; i < ItemPatinaDataTable.Rows.Count; i++)
                {
                    if (ItemPatinaDataTable.Rows[i]["PatinaID"].ToString() == DCR[d]["PatinaID"].ToString())
                        break;

                    if (i == ItemPatinaDataTable.Rows.Count - 1)
                    {
                        DataRow NewRow = ItemPatinaDataTable.NewRow();
                        NewRow["PatinaID"] = (DCR[d]["PatinaID"]);
                        NewRow["PatinaName"] = GetPatinaName(Convert.ToInt32(DCR[d]["PatinaID"]));
                        ItemPatinaDataTable.Rows.Add(NewRow);
                        break;
                    }
                }
            }

            bool HasPatinaRAL = false;
            if (ItemPatinaDataTable.Rows.Count > 0)
            {
                for (int i = 0; i < ItemPatinaDataTable.Rows.Count; i++)
                {
                    DataRow[] fRows = PatinaRALDataTable.Select("PatinaID=" + Convert.ToInt32(ItemPatinaDataTable.Rows[i]["PatinaID"]));
                    if (fRows.Count() > 0)
                    {
                        HasPatinaRAL = true;
                        break;
                    }
                }
            }
            if (ItemPatinaDataTable.Rows.Count > 0 && HasPatinaRAL)
            {
                DataTable ddd = ItemPatinaDataTable.Copy();
                ItemPatinaDataTable.Clear();
                ItemPatinaDataTable.AcceptChanges();

                foreach (DataRow Row in ddd.Rows)
                {
                    if (Convert.ToInt32(Row["PatinaID"]) == -1)
                    {
                        DataRow NewRow = ItemPatinaDataTable.NewRow();
                        NewRow["PatinaID"] = Convert.ToInt32(Row["PatinaID"]);
                        NewRow["PatinaName"] = GetPatinaName(Convert.ToInt32(Row["PatinaID"]));
                        ItemPatinaDataTable.Rows.Add(NewRow);
                    }
                    foreach (DataRow item in PatinaRALDataTable.Select("PatinaID=" + Convert.ToInt32(Row["PatinaID"])))
                    {
                        DataRow NewRow = ItemPatinaDataTable.NewRow();
                        NewRow["PatinaID"] = item["PatinaRALID"];
                        NewRow["PatinaName"] = item["PatinaRAL"]; NewRow["Patina"] = item["Patina"];
                        NewRow["DisplayName"] = item["DisplayName"];
                        ItemPatinaDataTable.Rows.Add(NewRow);
                    }
                }
            }
            ItemPatinaDataTable.DefaultView.Sort = "PatinaName ASC";
            ItemPatinaBindingSource.MoveFirst();

            if (ItemPatinaDataTable.Rows.Count == 1 && ItemPatinaDataTable.Rows[0]["PatinaID"].ToString() == "-1")
                return false;

            return true;
        }

        public bool FilterInsetType(string Name, int ColorID, int PatinaID, bool bExcluzive)
        {
            TempItemsDataTable.Clear();
            TempItemsDataTable = DecorDataTable;
            using (DataView DV = new DataView(TempItemsDataTable))
            {
                DV.RowFilter = "Name='" + Name + "'";

                TempItemsDataTable = DV.ToTable();
            }
            string filter = string.Empty;
            for (int i = 0; i < TempItemsDataTable.Rows.Count; i++)
                filter += Convert.ToInt32(TempItemsDataTable.Rows[i]["DecorID"]) + ",";
            if (filter.Length > 0)
            {
                filter = filter.Substring(0, filter.Length - 1);
                filter = "DecorID IN (" + filter + ")";
            }
            else
                filter = "DecorID <> - 1";

            if (bExcluzive)
                filter += " AND (Excluzive=1)";
            else
                filter += " AND (Excluzive IS NULL OR Excluzive=1)";
            ItemInsetTypesDataTable.Clear();
            ItemInsetTypesDataTable.AcceptChanges();

            if (PatinaID > 1000)
            {
                DataRow[] fRows = PatinaRALDataTable.Select("PatinaRALID=" + PatinaID);
                if (fRows.Count() > 0)
                    PatinaID = Convert.ToInt32(fRows[0]["PatinaID"]);
            }
            DataRow[] DCR = DecorConfigDataTable.Select(filter + " AND ColorID = " + ColorID.ToString()
                + " AND PatinaID = " + PatinaID.ToString());


            for (int d = 0; d < DCR.Count(); d++)
            {
                if (d == 0)
                {
                    DataRow NewRow = ItemInsetTypesDataTable.NewRow();
                    NewRow["InsetTypeID"] = (DCR[d]["InsetTypeID"]);
                    NewRow["InsetTypeName"] = GetInsetTypeName(Convert.ToInt32(DCR[d]["InsetTypeID"]));
                    ItemInsetTypesDataTable.Rows.Add(NewRow);
                    continue;
                }


                for (int i = 0; i < ItemInsetTypesDataTable.Rows.Count; i++)
                {
                    if (ItemInsetTypesDataTable.Rows[i]["InsetTypeID"].ToString() == DCR[d]["InsetTypeID"].ToString())
                        break;

                    if (i == ItemInsetTypesDataTable.Rows.Count - 1)
                    {
                        DataRow NewRow = ItemInsetTypesDataTable.NewRow();
                        NewRow["InsetTypeID"] = (DCR[d]["InsetTypeID"]);
                        NewRow["InsetTypeName"] = GetInsetTypeName(Convert.ToInt32(DCR[d]["InsetTypeID"]));
                        ItemInsetTypesDataTable.Rows.Add(NewRow);
                        break;
                    }
                }
            }

            ItemInsetTypesDataTable.DefaultView.Sort = "InsetTypeName ASC";
            ItemInsetTypesBindingSource.MoveFirst();

            if (ItemInsetTypesDataTable.Rows.Count == 1 && ItemInsetTypesDataTable.Rows[0]["InsetTypeID"].ToString() == "-1")
                return false;

            return true;
        }

        public bool FilterInsetColor(string Name, int ColorID, int PatinaID, int InsetTypeID, bool bExcluzive)
        {
            TempItemsDataTable.Clear();
            TempItemsDataTable = DecorDataTable;
            using (DataView DV = new DataView(TempItemsDataTable))
            {
                DV.RowFilter = "Name='" + Name + "'";

                TempItemsDataTable = DV.ToTable();
            }
            string filter = string.Empty;
            for (int i = 0; i < TempItemsDataTable.Rows.Count; i++)
                filter += Convert.ToInt32(TempItemsDataTable.Rows[i]["DecorID"]) + ",";
            if (filter.Length > 0)
            {
                filter = filter.Substring(0, filter.Length - 1);
                filter = "DecorID IN (" + filter + ")";
            }
            else
                filter = "DecorID <> - 1";

            if (bExcluzive)
                filter += " AND (Excluzive=1)";
            else
                filter += " AND (Excluzive IS NULL OR Excluzive=1)";
            ItemInsetColorsDataTable.Clear();
            ItemInsetColorsDataTable.AcceptChanges();

            if (PatinaID > 1000)
            {
                DataRow[] fRows = PatinaRALDataTable.Select("PatinaRALID=" + PatinaID);
                if (fRows.Count() > 0)
                    PatinaID = Convert.ToInt32(fRows[0]["PatinaID"]);
            }
            DataRow[] DCR = DecorConfigDataTable.Select(filter + " AND ColorID = " + ColorID.ToString()
                + " AND PatinaID = " + PatinaID.ToString() + " AND InsetTypeID = " + InsetTypeID.ToString());


            for (int d = 0; d < DCR.Count(); d++)
            {
                if (d == 0)
                {
                    DataRow NewRow = ItemInsetColorsDataTable.NewRow();
                    NewRow["InsetColorID"] = (DCR[d]["InsetColorID"]);
                    NewRow["InsetColorName"] = GetInsetColorName(Convert.ToInt32(DCR[d]["InsetColorID"]));
                    ItemInsetColorsDataTable.Rows.Add(NewRow);
                    continue;
                }


                for (int i = 0; i < ItemInsetColorsDataTable.Rows.Count; i++)
                {
                    if (ItemInsetColorsDataTable.Rows[i]["InsetColorID"].ToString() == DCR[d]["InsetColorID"].ToString())
                        break;

                    if (i == ItemInsetColorsDataTable.Rows.Count - 1)
                    {
                        DataRow NewRow = ItemInsetColorsDataTable.NewRow();
                        NewRow["InsetColorID"] = (DCR[d]["InsetColorID"]);
                        NewRow["InsetColorName"] = GetInsetColorName(Convert.ToInt32(DCR[d]["InsetColorID"]));
                        ItemInsetColorsDataTable.Rows.Add(NewRow);
                        break;
                    }
                }
            }

            ItemInsetColorsDataTable.DefaultView.Sort = "InsetColorName ASC";
            ItemInsetColorsBindingSource.MoveFirst();

            if (ItemInsetColorsDataTable.Rows.Count == 1 && ItemInsetColorsDataTable.Rows[0]["InsetColorID"].ToString() == "-1")
                return false;

            return true;
        }

        public int FilterLength(string Name, int ColorID, int PatinaID, int InsetTypeID, int InsetColorID)
        {
            TempItemsDataTable.Clear();
            TempItemsDataTable = DecorDataTable;
            using (DataView DV = new DataView(TempItemsDataTable))
            {
                DV.RowFilter = "Name='" + Name + "'";

                TempItemsDataTable = DV.ToTable();
            }
            string filter = string.Empty;
            for (int i = 0; i < TempItemsDataTable.Rows.Count; i++)
                filter += Convert.ToInt32(TempItemsDataTable.Rows[i]["DecorID"]) + ",";
            if (filter.Length > 0)
            {
                filter = filter.Substring(0, filter.Length - 1);
                filter = "DecorID IN (" + filter + ")";
            }
            else
                filter = "DecorID <> - 1";

            ItemLengthDataTable.Clear();
            ItemLengthDataTable.AcceptChanges();

            if (PatinaID > 1000)
            {
                DataRow[] fRows = PatinaRALDataTable.Select("PatinaRALID=" + PatinaID);
                if (fRows.Count() > 0)
                    PatinaID = Convert.ToInt32(fRows[0]["PatinaID"]);
            }
            DataRow[] DCR = DecorConfigDataTable.Select(filter + " AND ColorID = " + ColorID.ToString()
                + " AND PatinaID = " + PatinaID.ToString() + " AND InsetTypeID = " + InsetTypeID.ToString() + " AND InsetColorID = " + InsetColorID.ToString());


            for (int d = 0; d < DCR.Count(); d++)
            {
                if (d == 0)
                {
                    int Length = Convert.ToInt32(DCR[d]["Length"]);

                    DataRow NewRow = ItemLengthDataTable.NewRow();
                    NewRow["Length"] = Length;
                    ItemLengthDataTable.Rows.Add(NewRow);
                    continue;
                }


                for (int i = 0; i < ItemLengthDataTable.Rows.Count; i++)
                {
                    if (ItemLengthDataTable.Rows[i]["Length"].ToString() == DCR[d]["Length"].ToString())
                        break;

                    if (i == ItemLengthDataTable.Rows.Count - 1)
                    {
                        int Height = Convert.ToInt32(DCR[d]["Length"]);

                        DataRow NewRow = ItemLengthDataTable.NewRow();
                        NewRow["Length"] = Height;
                        ItemLengthDataTable.Rows.Add(NewRow);
                        break;
                    }
                }
            }

            ItemLengthDataTable.DefaultView.Sort = "Length ASC";
            ItemLengthBindingSource.MoveFirst();

            if (ItemLengthDataTable.Rows.Count > 0)
            {
                if (Convert.ToInt32(ItemLengthDataTable.Rows[0]["Length"]) == 0)
                {
                    return 0;
                }

                if (Convert.ToInt32(ItemLengthDataTable.Rows[0]["Length"]) == -1)
                {
                    return -1;
                }
            }

            //LengthEdit.Text = "";
            //LengthEdit.Items.Clear();
            //if (ItemLengthDataTable.Rows.Count > 0)
            //{
            //    if (Convert.ToInt32(ItemLengthDataTable.Rows[0]["Length"]) == 0)
            //    {
            //        LengthEdit.DropDownStyle = ComboBoxStyle.DropDown;
            //        return 0;
            //    }

            //    if (Convert.ToInt32(ItemLengthDataTable.Rows[0]["Length"]) == -1)
            //    {
            //        return -1;
            //    }
            //}

            //foreach (DataRow Row in ItemLengthDataTable.Rows)
            //    LengthEdit.Items.Add(Row["Length"].ToString());

            //LengthEdit.DropDownStyle = ComboBoxStyle.DropDownList;
            //if (ItemLengthDataTable.Rows.Count > 0)
            //    LengthEdit.SelectedIndex = 0;

            return ItemLengthDataTable.Rows.Count;
        }

        public int FilterHeight(string Name, int ColorID, int PatinaID, int InsetTypeID, int InsetColorID, int Length)
        {
            TempItemsDataTable.Clear();
            TempItemsDataTable = DecorDataTable;
            using (DataView DV = new DataView(TempItemsDataTable))
            {
                DV.RowFilter = "Name='" + Name + "'";

                TempItemsDataTable = DV.ToTable();
            }
            string filter = string.Empty;
            for (int i = 0; i < TempItemsDataTable.Rows.Count; i++)
                filter += Convert.ToInt32(TempItemsDataTable.Rows[i]["DecorID"]) + ",";
            if (filter.Length > 0)
            {
                filter = filter.Substring(0, filter.Length - 1);
                filter = "DecorID IN (" + filter + ")";
            }
            else
                filter = "DecorID <> - 1";

            ItemHeightDataTable.Clear();
            ItemHeightDataTable.AcceptChanges();

            if (PatinaID > 1000)
            {
                DataRow[] fRows = PatinaRALDataTable.Select("PatinaRALID=" + PatinaID);
                if (fRows.Count() > 0)
                    PatinaID = Convert.ToInt32(fRows[0]["PatinaID"]);
            }
            DataRow[] DCR = DecorConfigDataTable.Select(filter + " AND ColorID = " + ColorID.ToString()
                + " AND PatinaID = " + PatinaID.ToString() + " AND InsetTypeID = " + InsetTypeID.ToString() + " AND InsetColorID = " + InsetColorID.ToString() + " AND Length = " + Length.ToString());


            for (int d = 0; d < DCR.Count(); d++)
            {
                if (d == 0)
                {
                    int Height = Convert.ToInt32(DCR[d]["Height"]);

                    DataRow NewRow = ItemHeightDataTable.NewRow();
                    NewRow["Height"] = Height;
                    ItemHeightDataTable.Rows.Add(NewRow);
                    continue;
                }


                for (int i = 0; i < ItemHeightDataTable.Rows.Count; i++)
                {
                    if (ItemHeightDataTable.Rows[i]["Height"].ToString() == DCR[d]["Height"].ToString())
                        break;

                    if (i == ItemHeightDataTable.Rows.Count - 1)
                    {
                        int Height = Convert.ToInt32(DCR[d]["Height"]);

                        DataRow NewRow = ItemHeightDataTable.NewRow();
                        NewRow["Height"] = Height;
                        ItemHeightDataTable.Rows.Add(NewRow);
                        break;
                    }
                }
            }

            //ItemHeightDataTable.DefaultView.Sort = "Height ASC";

            //HeightEdit.Text = "";
            //HeightEdit.Items.Clear();
            //if (ItemHeightDataTable.Rows.Count > 0)
            //{
            //    if (Convert.ToInt32(ItemHeightDataTable.Rows[0]["Height"]) == 0)
            //    {
            //        HeightEdit.DropDownStyle = ComboBoxStyle.DropDown;
            //        return 0;
            //    }

            //    if (Convert.ToInt32(ItemHeightDataTable.Rows[0]["Height"]) == -1)
            //    {
            //        return -1;
            //    }
            //}

            //foreach (DataRow Row in ItemHeightDataTable.Rows)
            //    HeightEdit.Items.Add(Row["Height"].ToString());

            //HeightEdit.DropDownStyle = ComboBoxStyle.DropDownList;
            //if (ItemHeightDataTable.Rows.Count > 0)
            //    HeightEdit.SelectedIndex = 0;

            ItemHeightDataTable.DefaultView.Sort = "Height ASC";
            ItemHeightBindingSource.MoveFirst();

            if (ItemHeightDataTable.Rows.Count > 0)
            {
                if (Convert.ToInt32(ItemHeightDataTable.Rows[0]["Height"]) == 0)
                {
                    return 0;
                }

                if (Convert.ToInt32(ItemHeightDataTable.Rows[0]["Height"]) == -1)
                {
                    return -1;
                }
            }

            return ItemHeightDataTable.Rows.Count;
        }

        public int FilterWidth(string Name, int ColorID, int PatinaID, int InsetTypeID, int InsetColorID, int Length, int Height)
        {
            TempItemsDataTable.Clear();
            TempItemsDataTable = DecorDataTable;
            using (DataView DV = new DataView(TempItemsDataTable))
            {
                DV.RowFilter = "Name='" + Name + "'";

                TempItemsDataTable = DV.ToTable();
            }
            string filter = string.Empty;
            for (int i = 0; i < TempItemsDataTable.Rows.Count; i++)
                filter += Convert.ToInt32(TempItemsDataTable.Rows[i]["DecorID"]) + ",";
            if (filter.Length > 0)
            {
                filter = filter.Substring(0, filter.Length - 1);
                filter = "DecorID IN (" + filter + ")";
            }
            else
                filter = "DecorID <> - 1";

            ItemWidthDataTable.Clear();
            ItemWidthDataTable.AcceptChanges();

            if (PatinaID > 1000)
            {
                DataRow[] fRows = PatinaRALDataTable.Select("PatinaRALID=" + PatinaID);
                if (fRows.Count() > 0)
                    PatinaID = Convert.ToInt32(fRows[0]["PatinaID"]);
            }
            DataRow[] DCR = DecorConfigDataTable.Select(filter
                + " AND ColorID = " + ColorID.ToString() + " AND PatinaID = " + PatinaID.ToString()
                + " AND InsetTypeID = " + InsetTypeID.ToString() + " AND InsetColorID = " + InsetColorID.ToString() +
                " AND Length = " + Length.ToString() + " AND Height = " + Height.ToString());


            for (int d = 0; d < DCR.Count(); d++)
            {
                if (d == 0)
                {
                    int Width = Convert.ToInt32(DCR[d]["Width"]);

                    DataRow NewRow = ItemWidthDataTable.NewRow();
                    NewRow["Width"] = Width;
                    ItemWidthDataTable.Rows.Add(NewRow);
                    continue;
                }


                for (int i = 0; i < ItemWidthDataTable.Rows.Count; i++)
                {
                    if (ItemWidthDataTable.Rows[i]["Width"].ToString() == DCR[d]["Width"].ToString())
                        break;

                    if (i == ItemWidthDataTable.Rows.Count - 1)
                    {
                        int Width = Convert.ToInt32(DCR[d]["Width"]);

                        DataRow NewRow = ItemWidthDataTable.NewRow();
                        NewRow["Width"] = Width;
                        ItemWidthDataTable.Rows.Add(NewRow);
                        break;
                    }
                }
            }

            //ItemWidthDataTable.DefaultView.Sort = "Width ASC";

            //WidthEdit.Text = "";
            //WidthEdit.Items.Clear();
            //if (ItemWidthDataTable.Rows.Count > 0)
            //{
            //    if (Convert.ToInt32(ItemWidthDataTable.Rows[0]["Width"]) == 0)
            //    {
            //        WidthEdit.DropDownStyle = ComboBoxStyle.DropDown;
            //        return 0;
            //    }

            //    if (Convert.ToInt32(ItemWidthDataTable.Rows[0]["Width"]) == -1)
            //    {

            //        return -1;
            //    }
            //}
            //foreach (DataRow Row in ItemWidthDataTable.Rows)
            //    WidthEdit.Items.Add(Row["Width"].ToString());

            //WidthEdit.DropDownStyle = ComboBoxStyle.DropDownList;
            //if (ItemWidthDataTable.Rows.Count > 0)
            //    WidthEdit.SelectedIndex = 0;

            ItemWidthDataTable.DefaultView.Sort = "Width ASC";
            ItemWidthBindingSource.MoveFirst();

            if (ItemWidthDataTable.Rows.Count > 0)
            {
                if (Convert.ToInt32(ItemWidthDataTable.Rows[0]["Width"]) == 0)
                {
                    return 0;
                }

                if (Convert.ToInt32(ItemWidthDataTable.Rows[0]["Width"]) == -1)
                {
                    return -1;
                }
            }

            return ItemWidthDataTable.Rows.Count;
        }
    }
}
