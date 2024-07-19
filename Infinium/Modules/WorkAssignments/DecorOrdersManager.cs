using System;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Windows.Forms;

namespace Infinium.Modules.WorkAssignments
{
    public class DecorOrdersManager
    {
        private DataTable AllBatchDecorDT = null;

        private BindingSource BatchDecorBS = null;

        private DataTable BatchDecorDT = null;

        private DataTable ColorsDataTable = null;

        private BindingSource DecorColorsSummaryBS = null;

        private DataTable DecorColorsSummaryDT = null;

        private DataTable DecorDataTable = null;

        private BindingSource DecorItemsSummaryBS = null;

        private DataTable DecorItemsSummaryDT = null;

        private BindingSource DecorProductsSummaryBS = null;

        private DataTable DecorProductsSummaryDT = null;

        private BindingSource DecorSizesSummaryBS = null;

        private DataTable DecorSizesSummaryDT = null;

        private DataTable PatinaDataTable = null;

        private DataTable PatinaRALDataTable = null;

        private DataTable ProductsDataTable = null;

        //конструктор
        public DecorOrdersManager()
        {
            Initialize();
        }

        public BindingSource BatchDecorList => BatchDecorBS;

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

        public BindingSource DecorColorsSummaryList => DecorColorsSummaryBS;

        public BindingSource DecorItemsSummaryList => DecorItemsSummaryBS;

        public BindingSource DecorProductsSummaryList => DecorProductsSummaryBS;

        public BindingSource DecorSizesSummaryList => DecorSizesSummaryBS;

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

        public void FilterDecorColors(int ProductID, int DecorID, int MeasureID)
        {
            DecorColorsSummaryBS.Filter = "ProductID=" + ProductID + " AND DecorID="
                + DecorID + " AND MeasureID=" + MeasureID;
            DecorColorsSummaryBS.MoveFirst();
        }

        public void FilterDecorItems(int ProductID, int MeasureID)
        {
            DecorItemsSummaryBS.Filter = "ProductID=" + ProductID + " AND MeasureID=" + MeasureID;
            DecorItemsSummaryBS.MoveFirst();
        }

        public void FilterDecorSizes(int ProductID, int DecorID, int ColorID, int MeasureID)
        {
            DecorSizesSummaryBS.Filter = "ProductID=" + ProductID +
                " AND DecorID=" + DecorID + " AND ColorID=" + ColorID + " AND MeasureID=" + MeasureID;
            DecorSizesSummaryBS.MoveFirst();
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

        public string GetDecorName(int DecorID)
        {
            string Name = string.Empty;
            DataRow[] Rows = DecorDataTable.Select("DecorID = " + DecorID);
            if (Rows.Count() > 0)
                Name = Rows[0]["Name"].ToString();
            return Name;
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

        public string GetProductName(int ProductID)
        {
            string Name = string.Empty;
            DataRow[] Rows = ProductsDataTable.Select("ProductID = " + ProductID);
            if (Rows.Count() > 0)
                Name = Rows[0]["ProductName"].ToString();
            return Name;
        }

        public void Initialize()
        {
            Create();
            Fill();
            Binding();
        }

        private void Binding()
        {
            BatchDecorBS.DataSource = BatchDecorDT;
            DecorProductsSummaryBS.DataSource = DecorProductsSummaryDT;
            DecorItemsSummaryBS.DataSource = DecorItemsSummaryDT;
            DecorColorsSummaryBS.DataSource = DecorColorsSummaryDT;
            DecorSizesSummaryBS.DataSource = DecorSizesSummaryDT;
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

        private void GetColorsDT()
        {
            ColorsDataTable = new DataTable();
            ColorsDataTable.Columns.Add(new DataColumn("ColorID", Type.GetType("System.Int64")));
            ColorsDataTable.Columns.Add(new DataColumn("ColorName", Type.GetType("System.String")));
            string SelectCommand = @"SELECT TechStoreID, TechStoreName FROM TechStore
                WHERE TechStoreSubGroupID IN (SELECT TechStoreSubGroupID FROM TechStoreSubGroups WHERE (TechStoreGroupID = 11 OR TechStoreGroupID = 1))
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
    }
}