using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.UserModel.Contrib;
using NPOI.HSSF.Util;

using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;

namespace Infinium.Modules.StatisticsMarketing
{
    public class StatisticsReportByProduction : IAllFrontParameterName
    {
        private DataTable ClientsDataTable = null;

        private DataTable FrontsResultDataTable = null;
        private DataTable[] DecorResultDataTable = null;

        public DataTable FrontsDataTable = null;
        public DataTable FrameColorsDataTable = null;
        public DataTable PatinaDataTable = null;
        private DataTable PatinaRALDataTable = null;
        public DataTable InsetTypesDataTable = null;
        public DataTable InsetColorsDataTable = null;

        private Modules.ZOV.DecorCatalogOrder DecorCatalog = null;

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

        public StatisticsReportByProduction(ref Modules.ZOV.DecorCatalogOrder tDecorCatalog)
        {
            DecorCatalog = tDecorCatalog;

            Create();

            ClientsDataTable = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM Clients",
                ConnectionStrings.MarketingReferenceConnectionString))
            {
                DA.Fill(ClientsDataTable);
            }

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

            CreateFrontsDataTable();
            CreateDecorDataTable();
        }

        private void Create()
        {
            FrontsDataTable = new DataTable();
            FrameColorsDataTable = new DataTable();
            PatinaDataTable = new DataTable();
            InsetTypesDataTable = new DataTable();
            InsetColorsDataTable = new DataTable();
            FrontsResultDataTable = new DataTable();
            DecorResultDataTable = new DataTable[DecorCatalog.DecorProductsCount];
        }

        private void CreateFrontsDataTable()
        {
            FrontsResultDataTable = new DataTable();

            FrontsResultDataTable.Columns.Add(new DataColumn(("Front"), System.Type.GetType("System.String")));
            FrontsResultDataTable.Columns.Add(new DataColumn(("FrameColor"), System.Type.GetType("System.String")));
            FrontsResultDataTable.Columns.Add(new DataColumn(("TechnoColor"), System.Type.GetType("System.String")));
            FrontsResultDataTable.Columns.Add(new DataColumn(("Patina"), System.Type.GetType("System.String")));
            FrontsResultDataTable.Columns.Add(new DataColumn(("InsetType"), System.Type.GetType("System.String")));
            FrontsResultDataTable.Columns.Add(new DataColumn(("InsetColor"), System.Type.GetType("System.String")));
            FrontsResultDataTable.Columns.Add(new DataColumn(("TechnoInsetType"), System.Type.GetType("System.String")));
            FrontsResultDataTable.Columns.Add(new DataColumn(("TechnoInsetColor"), System.Type.GetType("System.String")));
            FrontsResultDataTable.Columns.Add(new DataColumn(("Height"), System.Type.GetType("System.Int32")));
            FrontsResultDataTable.Columns.Add(new DataColumn(("Width"), System.Type.GetType("System.Int32")));
            FrontsResultDataTable.Columns.Add(new DataColumn(("Square"), System.Type.GetType("System.Decimal")));
            FrontsResultDataTable.Columns.Add(new DataColumn(("Count"), System.Type.GetType("System.Int32")));
            FrontsResultDataTable.Columns.Add(new DataColumn(("AvgPrice"), System.Type.GetType("System.Decimal")));
            FrontsResultDataTable.Columns.Add(new DataColumn(("SumCost"), System.Type.GetType("System.Decimal")));
        }

        private void CreateDecorDataTable()
        {
            for (int i = 0; i < DecorCatalog.DecorProductsCount; i++)
            {
                DecorResultDataTable[i] = new DataTable();

                DecorResultDataTable[i].Columns.Add("Product", Type.GetType("System.String"));
                DecorResultDataTable[i].Columns.Add("Color", Type.GetType("System.String"));
                DecorResultDataTable[i].Columns.Add("Height", Type.GetType("System.Int32"));
                DecorResultDataTable[i].Columns.Add("Width", Type.GetType("System.Int32"));
                DecorResultDataTable[i].Columns.Add("Count", Type.GetType("System.Int32"));
                DecorResultDataTable[i].Columns.Add(new DataColumn(("AvgPrice"), System.Type.GetType("System.Decimal")));
                DecorResultDataTable[i].Columns.Add(new DataColumn(("SumCost"), System.Type.GetType("System.Decimal")));
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

        private string GetClientName(int ClientID)
        {
            DataRow[] Rows = ClientsDataTable.Select("ClientID = " + ClientID);
            return Rows[0]["ClientName"].ToString();
        }

        public DateTime GetCurrentDate
        {
            get
            {
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT GETDATE()", ConnectionStrings.CatalogConnectionString))
                {
                    using (DataTable DT = new DataTable())
                    {
                        DA.Fill(DT);

                        return Convert.ToDateTime(DT.Rows[0][0]);
                    }
                }
            }
        }

        private void FillFronts(DataTable FrontsOrdersDataTable)
        {
            FrontsResultDataTable.Clear();

            if (FrontsOrdersDataTable.Rows.Count < 1)
                return;

            FillMFronts(FrontsOrdersDataTable);
            //FillZFronts(FrontsOrdersDataTable);

            DataView DV1 = new DataView(FrontsResultDataTable)
            {
                Sort = "Front, FrameColor, TechnoColor, Patina, InsetType, InsetColor, TechnoInsetType, TechnoInsetColor, Height, Width, Count"
            };
            FrontsResultDataTable = DV1.ToTable();
            DV1.Dispose();
        }

        private void FillMFronts(DataTable FrontsOrdersDataTable)
        {
            if (FrontsOrdersDataTable.Rows.Count < 1)
                return;

            string Front = string.Empty;
            string FrameColor = string.Empty;
            string TechnoColor = string.Empty;
            string Patina = string.Empty;
            string InsetType = string.Empty;
            string InsetColor = string.Empty;
            string TechnoInsetType = string.Empty;
            string TechnoInsetColor = string.Empty;

            string QueryString = string.Empty;

            decimal FrontCost = 0;
            decimal AvgPrice = 0;
            decimal FrontSquare = 0;
            int FrontCount = 0;

            DataTable Table = new DataTable();

            using (DataView DV = new DataView(FrontsOrdersDataTable))
            {
                Table = DV.ToTable(true, new string[] { "FrontID",
                    "ColorID", "TechnoColorID", "PatinaID","InsetTypeID", "InsetColorID", "TechnoInsetTypeID", "TechnoInsetColorID", "Height", "Width"});
            }

            for (int i = 0; i < Table.Rows.Count; i++)
            {
                QueryString = "FrontID=" + Convert.ToInt32(Table.Rows[i]["FrontID"]) +
                    " AND ColorID=" + Convert.ToInt32(Table.Rows[i]["ColorID"]) +
                    " AND TechnoColorID=" + Convert.ToInt32(Table.Rows[i]["TechnoColorID"]) +
                    " AND PatinaID=" + Convert.ToInt32(Table.Rows[i]["PatinaID"]) +
                    " AND InsetTypeID=" + Convert.ToInt32(Table.Rows[i]["InsetTypeID"]) +
                    " AND InsetColorID=" + Convert.ToInt32(Table.Rows[i]["InsetColorID"]) +
                    " AND TechnoInsetTypeID=" + Convert.ToInt32(Table.Rows[i]["TechnoInsetTypeID"]) +
                    " AND TechnoInsetColorID=" + Convert.ToInt32(Table.Rows[i]["TechnoInsetColorID"]) +
                    " AND Height = '" + Table.Rows[i]["Height"].ToString() + "'" +
                    " AND Width = '" + Table.Rows[i]["Width"].ToString() + "'";

                DataRow[] Rows = FrontsOrdersDataTable.Select(QueryString);
                if (Rows.Count() != 0)
                {
                    foreach (DataRow row in Rows)
                    {
                        FrontCost += Convert.ToDecimal(row["Cost"]);
                        FrontSquare += Convert.ToDecimal(row["Square"]);
                        FrontCount += Convert.ToInt32(row["Count"]);
                    }
                    if (Convert.ToInt32(Rows[0]["MeasureID"]) == 1)
                    {
                        if (FrontSquare != 0)
                            AvgPrice = FrontCost / FrontSquare;
                    }
                    if (Convert.ToInt32(Rows[0]["MeasureID"]) == 3)
                    {
                        if (FrontCount != 0)
                            AvgPrice = FrontCost / FrontCount;
                    }

                    Front = GetFrontName(Convert.ToInt32(Table.Rows[i]["FrontID"]));
                    FrameColor = GetColorName(Convert.ToInt32(Table.Rows[i]["ColorID"]));
                    TechnoColor = GetColorName(Convert.ToInt32(Table.Rows[i]["TechnoColorID"]));
                    Patina = GetPatinaName(Convert.ToInt32(Table.Rows[i]["PatinaID"]));
                    InsetType = GetInsetTypeName(Convert.ToInt32(Table.Rows[i]["InsetTypeID"]));
                    InsetColor = GetInsetColorName(Convert.ToInt32(Table.Rows[i]["InsetColorID"]));
                    TechnoInsetType = GetInsetTypeName(Convert.ToInt32(Table.Rows[i]["TechnoInsetTypeID"]));
                    TechnoInsetColor = GetInsetColorName(Convert.ToInt32(Table.Rows[i]["TechnoInsetColorID"]));

                    //if (FrontType == "Прямой")
                    //    FrontType = string.Empty;

                    DataRow NewRow = FrontsResultDataTable.NewRow();
                    NewRow["Front"] = Front;
                    NewRow["Patina"] = Patina;
                    NewRow["FrameColor"] = FrameColor;
                    NewRow["TechnoColor"] = TechnoColor;
                    NewRow["InsetType"] = InsetType;
                    NewRow["InsetColor"] = InsetColor;
                    NewRow["TechnoInsetType"] = TechnoInsetType;
                    NewRow["TechnoInsetColor"] = TechnoInsetColor;
                    NewRow["Height"] = Convert.ToInt32(Table.Rows[i]["Height"]);
                    NewRow["Width"] = Convert.ToInt32(Table.Rows[i]["Width"]);
                    NewRow["Square"] = Decimal.Round(FrontSquare, 3, MidpointRounding.AwayFromZero);
                    NewRow["Count"] = FrontCount;
                    NewRow["AvgPrice"] = Decimal.Round(AvgPrice, 2, MidpointRounding.AwayFromZero);
                    NewRow["SumCost"] = Decimal.Round(FrontCost, 3, MidpointRounding.AwayFromZero);
                    FrontsResultDataTable.Rows.Add(NewRow);

                    AvgPrice = 0;
                    FrontCost = 0;
                    FrontSquare = 0;
                    FrontCount = 0;
                }
            }
        }

        private void FillZFronts(DataTable FrontsOrdersDataTable)
        {
            if (FrontsOrdersDataTable.Rows.Count < 1)
                return;

            string Front = string.Empty;
            string FrameColor = string.Empty;
            string TechnoColor = string.Empty;
            string Patina = string.Empty;
            string InsetType = string.Empty;
            string InsetColor = string.Empty;
            string TechnoInsetType = string.Empty;
            string TechnoInsetColor = string.Empty;

            string QueryString = string.Empty;

            decimal FrontCost = 0;
            decimal AvgPrice = 0;
            decimal FrontSquare = 0;
            int FrontCount = 0;

            DataTable Table = new DataTable();

            using (DataView DV = new DataView(FrontsOrdersDataTable, string.Empty, string.Empty, DataViewRowState.CurrentRows))
            {
                Table = DV.ToTable(true, new string[] { "FrontID",
                    "ColorID", "TechnoColorID", "PatinaID","InsetTypeID", "InsetColorID", "TechnoInsetTypeID", "TechnoInsetColorID", "Height", "Width"});
            }

            for (int i = 0; i < Table.Rows.Count; i++)
            {
                QueryString = "FrontID=" + Convert.ToInt32(Table.Rows[i]["FrontID"]) +
                    " AND ColorID=" + Convert.ToInt32(Table.Rows[i]["ColorID"]) +
                    " AND TechnoColorID=" + Convert.ToInt32(Table.Rows[i]["TechnoColorID"]) +
                    " AND PatinaID=" + Convert.ToInt32(Table.Rows[i]["PatinaID"]) +
                    " AND InsetTypeID=" + Convert.ToInt32(Table.Rows[i]["InsetTypeID"]) +
                    " AND InsetColorID=" + Convert.ToInt32(Table.Rows[i]["InsetColorID"]) +
                    " AND TechnoInsetTypeID=" + Convert.ToInt32(Table.Rows[i]["TechnoInsetTypeID"]) +
                    " AND TechnoInsetColorID=" + Convert.ToInt32(Table.Rows[i]["TechnoInsetColorID"]) +
                    " AND Height = '" + Table.Rows[i]["Height"].ToString() + "'" +
                    " AND Width = '" + Table.Rows[i]["Width"].ToString() + "'";

                DataRow[] Rows = FrontsOrdersDataTable.Select(QueryString);
                if (Rows.Count() != 0)
                {
                    foreach (DataRow row in Rows)
                    {
                        FrontCost += Convert.ToDecimal(row["Cost"]);
                        FrontSquare += Convert.ToDecimal(row["Square"]);
                        FrontCount += Convert.ToInt32(row["Count"]);
                    }

                    FrontCost = FrontCost * 100 / 120;
                    if (Convert.ToInt32(Rows[0]["MeasureID"]) == 1)
                    {
                        if (FrontSquare != 0)
                            AvgPrice = FrontCost / FrontSquare;
                    }
                    if (Convert.ToInt32(Rows[0]["MeasureID"]) == 3)
                    {
                        if (FrontCount != 0)
                            AvgPrice = FrontCost / FrontCount;
                    }
                    Front = GetFrontName(Convert.ToInt32(Table.Rows[i]["FrontID"]));
                    FrameColor = GetColorName(Convert.ToInt32(Table.Rows[i]["ColorID"]));
                    TechnoColor = GetColorName(Convert.ToInt32(Table.Rows[i]["TechnoColorID"]));
                    Patina = GetPatinaName(Convert.ToInt32(Table.Rows[i]["PatinaID"]));
                    InsetType = GetInsetTypeName(Convert.ToInt32(Table.Rows[i]["InsetTypeID"]));
                    InsetColor = GetInsetColorName(Convert.ToInt32(Table.Rows[i]["InsetColorID"]));
                    TechnoInsetType = GetInsetTypeName(Convert.ToInt32(Table.Rows[i]["TechnoInsetTypeID"]));
                    TechnoInsetColor = GetInsetColorName(Convert.ToInt32(Table.Rows[i]["TechnoInsetColorID"]));

                    //if (FrontType == "Прямой")
                    //    FrontType = string.Empty;

                    DataRow NewRow = FrontsResultDataTable.NewRow();
                    NewRow["Front"] = Front;
                    NewRow["Patina"] = Patina;
                    NewRow["FrameColor"] = FrameColor;
                    NewRow["TechnoColor"] = TechnoColor;
                    NewRow["InsetType"] = InsetType;
                    NewRow["InsetColor"] = InsetColor;
                    NewRow["TechnoInsetType"] = TechnoInsetType;
                    NewRow["TechnoInsetColor"] = TechnoInsetColor;
                    NewRow["Height"] = Convert.ToInt32(Table.Rows[i]["Height"]);
                    NewRow["Width"] = Convert.ToInt32(Table.Rows[i]["Width"]);
                    NewRow["Square"] = Decimal.Round(FrontSquare, 3, MidpointRounding.AwayFromZero);
                    NewRow["Count"] = FrontCount;
                    NewRow["AvgPrice"] = Decimal.Round(AvgPrice, 2, MidpointRounding.AwayFromZero);
                    NewRow["SumCost"] = Decimal.Round(FrontCost, 3, MidpointRounding.AwayFromZero);
                    FrontsResultDataTable.Rows.Add(NewRow);

                    AvgPrice = 0;
                    FrontCost = 0;
                    FrontSquare = 0;
                    FrontCount = 0;
                }
            }
        }

        private void FillDecor(DataTable DecorOrdersDataTable)
        {
            for (int i = 0; i < DecorCatalog.DecorProductsCount; i++)
            {
                DecorResultDataTable[i].Clear();
                DecorResultDataTable[i].AcceptChanges();
            }

            if (DecorOrdersDataTable.Rows.Count < 1)
                return;

            FillMDecor(DecorOrdersDataTable);
            FillZDecor(DecorOrdersDataTable);

            for (int i = 0; i < DecorCatalog.DecorProductsCount; i++)
            {
                DataView DV1 = new DataView(DecorResultDataTable[i])
                {
                    Sort = "Product, Color, Height, Width, Count"
                };
                DecorResultDataTable[i] = DV1.ToTable();
                DV1.Dispose();
            }
        }

        private void FillMDecor(DataTable DecorOrdersDataTable)
        {
            if (DecorOrdersDataTable.Rows.Count < 1)
                return;

            string QueryString = string.Empty;
            string Product = string.Empty;
            string Decor = string.Empty;
            string Color = string.Empty;

            int Height = 0;
            int Width = 0;
            int Count = 0;

            decimal DecorCost = 0;
            decimal SumCost = 0;
            decimal SumCount = 0;
            decimal AvgPrice = 0;

            DataTable TempDecorProductsDT = new DataTable();
            DataTable Table = new DataTable();

            DataView DV3 = new DataView(DecorCatalog.DecorProductsDataTable)
            {
                Sort = "ProductName"
            };
            TempDecorProductsDT = DV3.ToTable();
            DV3.Dispose();

            for (int i = 0; i < DecorCatalog.DecorProductsCount; i++)
            {
                using (DataView DV = new DataView(DecorOrdersDataTable,
                    "ProductID = " + Convert.ToInt32(TempDecorProductsDT.Rows[i]["ProductID"]),
                    string.Empty, DataViewRowState.CurrentRows))
                {
                    Table = DV.ToTable(true, new string[] { "DecorID", "ColorID", "MeasureID", "Length", "Height", "Width" });
                }

                for (int j = 0; j < Table.Rows.Count; j++)
                {
                    QueryString = "ProductID=" + Convert.ToInt32(TempDecorProductsDT.Rows[i]["ProductID"]) +
                        " AND DecorID=" + Convert.ToInt32(Table.Rows[j]["DecorID"]) +
                        " AND ColorID=" + Convert.ToInt32(Table.Rows[j]["ColorID"]) +
                        " AND MeasureID=" + Convert.ToInt32(Table.Rows[j]["MeasureID"]) +
                        " AND Length=" + Convert.ToInt32(Table.Rows[j]["Length"]) +
                        " AND Height=" + Convert.ToInt32(Table.Rows[j]["Height"]) +
                        " AND Width=" + Convert.ToInt32(Table.Rows[j]["Width"]);

                    DataRow[] Rows = DecorOrdersDataTable.Select(QueryString);

                    if (Rows.Count() == 0)
                        continue;

                    foreach (DataRow Row in Rows)
                    {
                        SumCount = 0;
                        SumCost = 0;

                        foreach (DataRow row in Rows)
                        {
                            if (Convert.ToInt32(row["MeasureID"]) == 1)
                            {
                                SumCount += Convert.ToDecimal(row["Height"]) * Convert.ToDecimal(row["Width"]) * Convert.ToInt32(row["Count"]) / 1000000;
                            }

                            if (Convert.ToInt32(row["MeasureID"]) == 3)
                            {
                                SumCount += Convert.ToInt32(row["Count"]);
                            }

                            if (Convert.ToInt32(row["MeasureID"]) == 2)
                            {
                                //нет параметра "высота"
                                if (row["Height"].ToString() == "-1")
                                    SumCount += Convert.ToDecimal(row["Length"]) * Convert.ToDecimal(row["Count"]) / 1000;
                                else
                                    SumCount += Convert.ToDecimal(row["Height"]) * Convert.ToDecimal(row["Count"]) / 1000;
                            }
                            SumCost += Convert.ToDecimal(row["Cost"]);
                        }

                        Product = TempDecorProductsDT.Rows[i]["ProductName"].ToString() + " " +
                            DecorCatalog.GetItemName(Convert.ToInt32(Row["DecorID"]));
                        Count = Convert.ToInt32(Row["Count"]);
                        DecorCost = Convert.ToDecimal(Row["Cost"]);

                        if (SumCount != 0)
                            AvgPrice = SumCost / SumCount;

                        QueryString = "Product = '" + Product + "'";

                        if (DecorCatalog.HasParameter(Convert.ToInt32(Row["ProductID"]), "ColorID"))
                        {
                            Color = GetColorName(Convert.ToInt32(Row["ColorID"]));
                            if (Convert.ToInt32(Row["PatinaID"]) != -1)
                                Color += " " + GetPatinaName(Convert.ToInt32(Row["PatinaID"]));
                            QueryString += " AND Color = '" + Color.ToString() + "'";
                        }

                        if (DecorCatalog.HasParameter(Convert.ToInt32(Row["ProductID"]), "Height"))
                        {
                            Height = Convert.ToInt32(Row["Height"]);
                            QueryString += " AND Height = '" + Height.ToString() + "'";
                        }

                        if (DecorCatalog.HasParameter(Convert.ToInt32(Row["ProductID"]), "Length"))
                        {
                            Height = Convert.ToInt32(Row["Length"]);
                            QueryString += " AND Height = '" + Height.ToString() + "'";
                        }

                        if (DecorCatalog.HasParameter(Convert.ToInt32(Row["ProductID"]), "Width"))
                        {
                            Width = Convert.ToInt32(Row["Width"]);
                            QueryString += " AND Width = '" + Width.ToString() + "'";
                        }

                        DataRow[] dRow = DecorResultDataTable[i].Select(QueryString);

                        if (dRow.Count() == 0)
                        {
                            DataRow NewRow = DecorResultDataTable[i].NewRow();

                            NewRow["Product"] = Product;

                            if (DecorCatalog.HasParameter(Convert.ToInt32(Row["ProductID"]), "Height"))
                            {
                                if (Convert.ToInt32(Row["Height"]) != -1)
                                    NewRow["Height"] = Row["Height"];
                                if (Convert.ToInt32(Row["Length"]) != -1)
                                    NewRow["Height"] = Row["Length"];
                            }
                            if (DecorCatalog.HasParameter(Convert.ToInt32(Row["ProductID"]), "Length"))
                            {
                                if (Convert.ToInt32(Row["Length"]) != -1)
                                    NewRow["Height"] = Row["Length"];
                                if (Convert.ToInt32(Row["Height"]) != -1)
                                    NewRow["Height"] = Row["Height"];
                            }
                            if (DecorCatalog.HasParameter(Convert.ToInt32(Row["ProductID"]), "Width"))
                                NewRow["Width"] = Row["Width"];

                            if (DecorCatalog.HasParameter(Convert.ToInt32(Row["ProductID"]), "ColorID"))
                                NewRow["Color"] = Color;

                            NewRow["Count"] = Row["Count"];
                            NewRow["AvgPrice"] = Decimal.Round(AvgPrice, 2, MidpointRounding.AwayFromZero);
                            NewRow["SumCost"] = Decimal.Round(Convert.ToDecimal(Row["Cost"]), 3, MidpointRounding.AwayFromZero);

                            DecorResultDataTable[i].Rows.Add(NewRow);
                        }
                        else
                        {
                            dRow[0]["Count"] = Convert.ToDecimal(dRow[0]["Count"]) + Count;
                            dRow[0]["SumCost"] = Convert.ToDecimal(dRow[0]["SumCost"]) + DecorCost;
                        }
                    }
                }
            }

            TempDecorProductsDT.Dispose();
        }

        private void FillZDecor(DataTable DecorOrdersDataTable)
        {
            if (DecorOrdersDataTable.Rows.Count < 1)
                return;

            string QueryString = string.Empty;
            string Product = string.Empty;
            string Decor = string.Empty;
            string Color = string.Empty;

            int Height = 0;
            int Width = 0;
            int Count = 0;

            decimal DecorCost = 0;
            decimal SumCost = 0;
            decimal SumCount = 0;
            decimal AvgPrice = 0;

            DataTable TempDecorProductsDT = new DataTable();
            DataTable Table = new DataTable();

            DataView DV3 = new DataView(DecorCatalog.DecorProductsDataTable)
            {
                Sort = "ProductName"
            };
            TempDecorProductsDT = DV3.ToTable();
            DV3.Dispose();

            for (int i = 0; i < DecorCatalog.DecorProductsCount; i++)
            {
                using (DataView DV = new DataView(DecorOrdersDataTable,
                    "ProductID = " + Convert.ToInt32(TempDecorProductsDT.Rows[i]["ProductID"]),
                    string.Empty, DataViewRowState.CurrentRows))
                {
                    Table = DV.ToTable(true, new string[] { "DecorID", "ColorID", "MeasureID", "Length", "Height", "Width" });
                }

                for (int j = 0; j < Table.Rows.Count; j++)
                {
                    QueryString = "ProductID=" + Convert.ToInt32(TempDecorProductsDT.Rows[i]["ProductID"]) +
                        " AND DecorID=" + Convert.ToInt32(Table.Rows[j]["DecorID"]) +
                        " AND ColorID=" + Convert.ToInt32(Table.Rows[j]["ColorID"]) +
                        " AND MeasureID=" + Convert.ToInt32(Table.Rows[j]["MeasureID"]) +
                        " AND Length=" + Convert.ToInt32(Table.Rows[j]["Length"]) +
                        " AND Height=" + Convert.ToInt32(Table.Rows[j]["Height"]) +
                        " AND Width=" + Convert.ToInt32(Table.Rows[j]["Width"]);

                    DataRow[] Rows = DecorOrdersDataTable.Select(QueryString);

                    if (Rows.Count() == 0)
                        continue;

                    foreach (DataRow Row in Rows)
                    {
                        SumCount = 0;
                        SumCost = 0;

                        foreach (DataRow row in Rows)
                        {
                            if (Convert.ToInt32(row["MeasureID"]) == 1)
                            {
                                SumCount += Convert.ToDecimal(row["Height"]) * Convert.ToDecimal(row["Width"]) * Convert.ToInt32(row["Count"]) / 1000000;
                            }

                            if (Convert.ToInt32(row["MeasureID"]) == 3)
                            {
                                SumCount += Convert.ToInt32(row["Count"]);
                            }

                            if (Convert.ToInt32(row["MeasureID"]) == 2)
                            {
                                //нет параметра "высота"
                                if (row["Height"].ToString() == "-1")
                                    SumCount += Convert.ToDecimal(row["Length"]) * Convert.ToDecimal(row["Count"]) / 1000;
                                else
                                    SumCount += Convert.ToDecimal(row["Height"]) * Convert.ToDecimal(row["Count"]) / 1000;
                            }
                            SumCost += Convert.ToDecimal(row["Cost"]);
                        }

                        Product = TempDecorProductsDT.Rows[i]["ProductName"].ToString() + " " +
                            DecorCatalog.GetItemName(Convert.ToInt32(Row["DecorID"]));
                        Count = Convert.ToInt32(Row["Count"]);
                        DecorCost = Convert.ToDecimal(Row["Cost"]);

                        DecorCost = DecorCost * 100 / 120;

                        if (SumCount != 0)
                            AvgPrice = SumCost / SumCount;

                        QueryString = "Product = '" + Product + "'";

                        if (DecorCatalog.HasParameter(Convert.ToInt32(Row["ProductID"]), "ColorID"))
                        {
                            Color = GetColorName(Convert.ToInt32(Row["ColorID"]));
                            if (Convert.ToInt32(Row["PatinaID"]) != -1)
                                Color += " " + GetPatinaName(Convert.ToInt32(Row["PatinaID"]));
                            QueryString += " AND Color = '" + Color.ToString() + "'";
                        }

                        if (DecorCatalog.HasParameter(Convert.ToInt32(Row["ProductID"]), "Height"))
                        {
                            Height = Convert.ToInt32(Row["Height"]);
                            QueryString += " AND Height = '" + Height.ToString() + "'";
                        }

                        if (DecorCatalog.HasParameter(Convert.ToInt32(Row["ProductID"]), "Length"))
                        {
                            Height = Convert.ToInt32(Row["Length"]);
                            QueryString += " AND Height = '" + Height.ToString() + "'";
                        }

                        if (DecorCatalog.HasParameter(Convert.ToInt32(Row["ProductID"]), "Width"))
                        {
                            Width = Convert.ToInt32(Row["Width"]);
                            QueryString += " AND Width = '" + Width.ToString() + "'";
                        }

                        DataRow[] dRow = DecorResultDataTable[i].Select(QueryString);

                        if (dRow.Count() == 0)
                        {
                            DataRow NewRow = DecorResultDataTable[i].NewRow();

                            NewRow["Product"] = Product;

                            if (DecorCatalog.HasParameter(Convert.ToInt32(Row["ProductID"]), "Height"))
                            {
                                if (Convert.ToInt32(Row["Height"]) != -1)
                                    NewRow["Height"] = Row["Height"];
                                if (Convert.ToInt32(Row["Length"]) != -1)
                                    NewRow["Height"] = Row["Length"];
                            }
                            if (DecorCatalog.HasParameter(Convert.ToInt32(Row["ProductID"]), "Length"))
                            {
                                if (Convert.ToInt32(Row["Length"]) != -1)
                                    NewRow["Height"] = Row["Length"];
                                if (Convert.ToInt32(Row["Height"]) != -1)
                                    NewRow["Height"] = Row["Height"];
                            }
                            if (DecorCatalog.HasParameter(Convert.ToInt32(Row["ProductID"]), "Width"))
                                NewRow["Width"] = Row["Width"];

                            if (DecorCatalog.HasParameter(Convert.ToInt32(Row["ProductID"]), "ColorID"))
                                NewRow["Color"] = Color;

                            NewRow["Count"] = Row["Count"];
                            NewRow["AvgPrice"] = Decimal.Round(AvgPrice, 2, MidpointRounding.AwayFromZero);
                            NewRow["SumCost"] = Row["Cost"];

                            DecorResultDataTable[i].Rows.Add(NewRow);
                        }
                        else
                        {
                            dRow[0]["Count"] = Convert.ToDecimal(dRow[0]["Count"]) + Count;
                            dRow[0]["SumCost"] = Convert.ToDecimal(dRow[0]["SumCost"]) + DecorCost;
                        }
                    }
                }
            }

            TempDecorProductsDT.Dispose();
        }

        private int GetCount(DataTable DT, bool Curved)
        {
            int S = 0;

            foreach (DataRow Row in DT.Rows)
            {
                if (Curved)
                {
                    if (Convert.ToInt32(Row["Width"]) == -1)
                        S += Convert.ToInt32(Row["Count"]);
                }
                else
                    S += Convert.ToInt32(Row["Count"]);
            }

            return S;
        }

        private decimal GetSquare(DataTable DT)
        {
            decimal S = 0;

            foreach (DataRow Row in DT.Rows)
            {
                if (Row["Square"] != DBNull.Value)
                    S += Convert.ToDecimal(Row["Square"]);
            }

            return S;
        }

        public void CreateReport(DateTime DateFrom, DateTime DateTo, int FactoryID,
            DataTable FrontsOrdersDataTable, DataTable DecorOrdersDataTable, string FileName)
        {
            if (FrontsOrdersDataTable.Rows.Count < 1 && DecorOrdersDataTable.Rows.Count < 1)
                return;

            HSSFWorkbook hssfworkbook = new HSSFWorkbook();

            ////create a entry of DocumentSummaryInformation
            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = "NPOI Team";
            hssfworkbook.DocumentSummaryInformation = dsi;

            ////create a entry of SummaryInformation
            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
            si.Subject = "NPOI SDK Example";
            hssfworkbook.SummaryInformation = si;

            #region Create fonts and styles

            HSSFFont ClientNameFont = hssfworkbook.CreateFont();
            ClientNameFont.FontHeightInPoints = 14;
            ClientNameFont.Boldweight = HSSFFont.BOLDWEIGHT_BOLD;
            ClientNameFont.FontName = "Calibri";

            HSSFFont MainFont = hssfworkbook.CreateFont();
            MainFont.FontHeightInPoints = 13;
            MainFont.Boldweight = HSSFFont.BOLDWEIGHT_BOLD;
            MainFont.FontName = "Calibri";

            HSSFCellStyle ClientNameStyle = hssfworkbook.CreateCellStyle();
            ClientNameStyle.SetFont(ClientNameFont);

            HSSFCellStyle MainStyle = hssfworkbook.CreateCellStyle();
            MainStyle.SetFont(MainFont);

            HSSFFont HeaderFont = hssfworkbook.CreateFont();
            HeaderFont.FontHeightInPoints = 13;
            HeaderFont.Boldweight = HSSFFont.BOLDWEIGHT_BOLD;
            HeaderFont.FontName = "Calibri";

            HSSFCellStyle HeaderStyle = hssfworkbook.CreateCellStyle();
            HeaderStyle.BorderBottom = HSSFCellStyle.BORDER_THIN;
            HeaderStyle.BottomBorderColor = HSSFColor.BLACK.index;
            HeaderStyle.BorderLeft = HSSFCellStyle.BORDER_THIN;
            HeaderStyle.LeftBorderColor = HSSFColor.BLACK.index;
            HeaderStyle.BorderRight = HSSFCellStyle.BORDER_THIN;
            HeaderStyle.RightBorderColor = HSSFColor.BLACK.index;
            HeaderStyle.BorderTop = HSSFCellStyle.BORDER_THIN;
            HeaderStyle.TopBorderColor = HSSFColor.BLACK.index;
            HeaderStyle.SetFont(HeaderFont);

            HSSFFont PackNumberFont = hssfworkbook.CreateFont();
            PackNumberFont.Boldweight = 12 * 256;
            PackNumberFont.FontName = "Calibri";

            HSSFCellStyle PackNumberStyle = hssfworkbook.CreateCellStyle();
            PackNumberStyle.BorderBottom = HSSFCellStyle.BORDER_THIN;
            PackNumberStyle.BottomBorderColor = HSSFColor.BLACK.index;
            PackNumberStyle.BorderLeft = HSSFCellStyle.BORDER_THIN;
            PackNumberStyle.LeftBorderColor = HSSFColor.BLACK.index;
            PackNumberStyle.BorderRight = HSSFCellStyle.BORDER_THIN;
            PackNumberStyle.RightBorderColor = HSSFColor.BLACK.index;
            PackNumberStyle.BorderTop = HSSFCellStyle.BORDER_THIN;
            PackNumberStyle.TopBorderColor = HSSFColor.BLACK.index;
            PackNumberStyle.Alignment = HSSFCellStyle.ALIGN_CENTER;
            PackNumberStyle.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
            PackNumberStyle.SetFont(PackNumberFont);

            HSSFFont SimpleFont = hssfworkbook.CreateFont();
            SimpleFont.FontHeightInPoints = 12;
            SimpleFont.FontName = "Calibri";

            HSSFCellStyle SimpleCellStyle = hssfworkbook.CreateCellStyle();
            SimpleCellStyle.BorderBottom = HSSFCellStyle.BORDER_THIN;
            SimpleCellStyle.BottomBorderColor = HSSFColor.BLACK.index;
            SimpleCellStyle.BorderLeft = HSSFCellStyle.BORDER_THIN;
            SimpleCellStyle.LeftBorderColor = HSSFColor.BLACK.index;
            SimpleCellStyle.BorderRight = HSSFCellStyle.BORDER_THIN;
            SimpleCellStyle.RightBorderColor = HSSFColor.BLACK.index;
            SimpleCellStyle.BorderTop = HSSFCellStyle.BORDER_THIN;
            SimpleCellStyle.TopBorderColor = HSSFColor.BLACK.index;
            SimpleCellStyle.SetFont(SimpleFont);

            HSSFCellStyle cellStyle = hssfworkbook.CreateCellStyle();
            cellStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.000");
            cellStyle.BorderBottom = HSSFCellStyle.BORDER_THIN;
            cellStyle.BottomBorderColor = HSSFColor.BLACK.index;
            cellStyle.BorderLeft = HSSFCellStyle.BORDER_THIN;
            cellStyle.LeftBorderColor = HSSFColor.BLACK.index;
            cellStyle.BorderRight = HSSFCellStyle.BORDER_THIN;
            cellStyle.RightBorderColor = HSSFColor.BLACK.index;
            cellStyle.BorderTop = HSSFCellStyle.BORDER_THIN;
            cellStyle.TopBorderColor = HSSFColor.BLACK.index;
            cellStyle.SetFont(SimpleFont);

            HSSFCellStyle GreyCellStyle = hssfworkbook.CreateCellStyle();
            GreyCellStyle.BorderBottom = HSSFCellStyle.BORDER_THIN;
            GreyCellStyle.BottomBorderColor = HSSFColor.BLACK.index;
            GreyCellStyle.BorderLeft = HSSFCellStyle.BORDER_THIN;
            GreyCellStyle.LeftBorderColor = HSSFColor.BLACK.index;
            GreyCellStyle.BorderRight = HSSFCellStyle.BORDER_THIN;
            GreyCellStyle.RightBorderColor = HSSFColor.BLACK.index;
            GreyCellStyle.BorderTop = HSSFCellStyle.BORDER_THIN;
            GreyCellStyle.TopBorderColor = HSSFColor.BLACK.index;
            GreyCellStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.GREY_25_PERCENT.index;
            GreyCellStyle.FillPattern = HSSFCellStyle.SOLID_FOREGROUND;
            GreyCellStyle.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.YELLOW.index;
            GreyCellStyle.SetFont(SimpleFont);

            HSSFFont TotalFont = hssfworkbook.CreateFont();
            TotalFont.FontHeightInPoints = 12;
            TotalFont.FontName = "Calibri";

            HSSFCellStyle TotalStyle = hssfworkbook.CreateCellStyle();
            TotalStyle.BorderTop = HSSFCellStyle.BORDER_MEDIUM;
            TotalStyle.TopBorderColor = HSSFColor.BLACK.index;
            TotalStyle.SetFont(TotalFont);

            HSSFCellStyle TempStyle = hssfworkbook.CreateCellStyle();
            TempStyle.SetFont(TotalFont);

            #endregion Create fonts and styles

            #region границы между упаковками

            HSSFCellStyle BottomMediumBorderCellStyle = hssfworkbook.CreateCellStyle();
            BottomMediumBorderCellStyle.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            BottomMediumBorderCellStyle.BottomBorderColor = HSSFColor.BLACK.index;
            BottomMediumBorderCellStyle.BorderLeft = HSSFCellStyle.BORDER_THIN;
            BottomMediumBorderCellStyle.LeftBorderColor = HSSFColor.BLACK.index;
            BottomMediumBorderCellStyle.BorderRight = HSSFCellStyle.BORDER_THIN;
            BottomMediumBorderCellStyle.RightBorderColor = HSSFColor.BLACK.index;
            BottomMediumBorderCellStyle.BorderTop = HSSFCellStyle.BORDER_THIN;
            BottomMediumBorderCellStyle.TopBorderColor = HSSFColor.BLACK.index;
            BottomMediumBorderCellStyle.SetFont(SimpleFont);

            HSSFCellStyle BottomMediumLeftBorderCellStyle = hssfworkbook.CreateCellStyle();
            BottomMediumLeftBorderCellStyle.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            BottomMediumLeftBorderCellStyle.BottomBorderColor = HSSFColor.BLACK.index;
            BottomMediumLeftBorderCellStyle.BorderLeft = HSSFCellStyle.BORDER_MEDIUM;
            BottomMediumLeftBorderCellStyle.LeftBorderColor = HSSFColor.BLACK.index;
            BottomMediumLeftBorderCellStyle.BorderRight = HSSFCellStyle.BORDER_THIN;
            BottomMediumLeftBorderCellStyle.RightBorderColor = HSSFColor.BLACK.index;
            BottomMediumLeftBorderCellStyle.BorderTop = HSSFCellStyle.BORDER_THIN;
            BottomMediumLeftBorderCellStyle.TopBorderColor = HSSFColor.BLACK.index;
            BottomMediumLeftBorderCellStyle.Alignment = HSSFCellStyle.ALIGN_CENTER;
            BottomMediumLeftBorderCellStyle.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
            BottomMediumLeftBorderCellStyle.SetFont(PackNumberFont);

            HSSFCellStyle BottomMediumRightBorderCellStyle = hssfworkbook.CreateCellStyle();
            BottomMediumRightBorderCellStyle.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            BottomMediumRightBorderCellStyle.BottomBorderColor = HSSFColor.BLACK.index;
            BottomMediumRightBorderCellStyle.BorderLeft = HSSFCellStyle.BORDER_THIN;
            BottomMediumRightBorderCellStyle.LeftBorderColor = HSSFColor.BLACK.index;
            BottomMediumRightBorderCellStyle.BorderRight = HSSFCellStyle.BORDER_MEDIUM;
            BottomMediumRightBorderCellStyle.RightBorderColor = HSSFColor.BLACK.index;
            BottomMediumRightBorderCellStyle.BorderTop = HSSFCellStyle.BORDER_THIN;
            BottomMediumRightBorderCellStyle.TopBorderColor = HSSFColor.BLACK.index;
            BottomMediumRightBorderCellStyle.SetFont(SimpleFont);

            HSSFCellStyle LeftMediumBorderCellStyle = hssfworkbook.CreateCellStyle();
            LeftMediumBorderCellStyle.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            LeftMediumBorderCellStyle.BottomBorderColor = HSSFColor.BLACK.index;
            LeftMediumBorderCellStyle.BorderLeft = HSSFCellStyle.BORDER_MEDIUM;
            LeftMediumBorderCellStyle.LeftBorderColor = HSSFColor.BLACK.index;
            LeftMediumBorderCellStyle.BorderRight = HSSFCellStyle.BORDER_THIN;
            LeftMediumBorderCellStyle.RightBorderColor = HSSFColor.BLACK.index;
            LeftMediumBorderCellStyle.BorderTop = HSSFCellStyle.BORDER_THIN;
            LeftMediumBorderCellStyle.TopBorderColor = HSSFColor.BLACK.index;
            LeftMediumBorderCellStyle.Alignment = HSSFCellStyle.ALIGN_CENTER;
            LeftMediumBorderCellStyle.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
            LeftMediumBorderCellStyle.SetFont(PackNumberFont);

            HSSFCellStyle RightMediumBorderCellStyle = hssfworkbook.CreateCellStyle();
            RightMediumBorderCellStyle.BorderBottom = HSSFCellStyle.BORDER_THIN;
            RightMediumBorderCellStyle.BottomBorderColor = HSSFColor.BLACK.index;
            RightMediumBorderCellStyle.BorderLeft = HSSFCellStyle.BORDER_THIN;
            RightMediumBorderCellStyle.LeftBorderColor = HSSFColor.BLACK.index;
            RightMediumBorderCellStyle.BorderRight = HSSFCellStyle.BORDER_MEDIUM;
            RightMediumBorderCellStyle.RightBorderColor = HSSFColor.BLACK.index;
            RightMediumBorderCellStyle.BorderTop = HSSFCellStyle.BORDER_THIN;
            RightMediumBorderCellStyle.TopBorderColor = HSSFColor.BLACK.index;
            RightMediumBorderCellStyle.SetFont(SimpleFont);

            HSSFCellStyle GreyBottomMediumBorderCellStyle = hssfworkbook.CreateCellStyle();
            GreyBottomMediumBorderCellStyle.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            GreyBottomMediumBorderCellStyle.BottomBorderColor = HSSFColor.BLACK.index;
            GreyBottomMediumBorderCellStyle.BorderLeft = HSSFCellStyle.BORDER_THIN;
            GreyBottomMediumBorderCellStyle.LeftBorderColor = HSSFColor.BLACK.index;
            GreyBottomMediumBorderCellStyle.BorderRight = HSSFCellStyle.BORDER_THIN;
            GreyBottomMediumBorderCellStyle.RightBorderColor = HSSFColor.BLACK.index;
            GreyBottomMediumBorderCellStyle.BorderTop = HSSFCellStyle.BORDER_THIN;
            GreyBottomMediumBorderCellStyle.TopBorderColor = HSSFColor.BLACK.index;
            GreyBottomMediumBorderCellStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.GREY_25_PERCENT.index;
            GreyBottomMediumBorderCellStyle.FillPattern = HSSFCellStyle.SOLID_FOREGROUND;
            GreyBottomMediumBorderCellStyle.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.YELLOW.index;
            GreyBottomMediumBorderCellStyle.SetFont(SimpleFont);

            HSSFCellStyle GreyBottomMediumLeftBorderCellStyle = hssfworkbook.CreateCellStyle();
            GreyBottomMediumLeftBorderCellStyle.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            GreyBottomMediumLeftBorderCellStyle.BottomBorderColor = HSSFColor.BLACK.index;
            GreyBottomMediumLeftBorderCellStyle.BorderLeft = HSSFCellStyle.BORDER_MEDIUM;
            GreyBottomMediumLeftBorderCellStyle.LeftBorderColor = HSSFColor.BLACK.index;
            GreyBottomMediumLeftBorderCellStyle.BorderRight = HSSFCellStyle.BORDER_THIN;
            GreyBottomMediumLeftBorderCellStyle.RightBorderColor = HSSFColor.BLACK.index;
            GreyBottomMediumLeftBorderCellStyle.BorderTop = HSSFCellStyle.BORDER_THIN;
            GreyBottomMediumLeftBorderCellStyle.TopBorderColor = HSSFColor.BLACK.index;
            GreyBottomMediumLeftBorderCellStyle.Alignment = HSSFCellStyle.ALIGN_CENTER;
            GreyBottomMediumLeftBorderCellStyle.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
            GreyBottomMediumLeftBorderCellStyle.SetFont(PackNumberFont);

            HSSFCellStyle GreyBottomMediumRightBorderCellStyle = hssfworkbook.CreateCellStyle();
            GreyBottomMediumRightBorderCellStyle.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            GreyBottomMediumRightBorderCellStyle.BottomBorderColor = HSSFColor.BLACK.index;
            GreyBottomMediumRightBorderCellStyle.BorderLeft = HSSFCellStyle.BORDER_THIN;
            GreyBottomMediumRightBorderCellStyle.LeftBorderColor = HSSFColor.BLACK.index;
            GreyBottomMediumRightBorderCellStyle.BorderRight = HSSFCellStyle.BORDER_MEDIUM;
            GreyBottomMediumRightBorderCellStyle.RightBorderColor = HSSFColor.BLACK.index;
            GreyBottomMediumRightBorderCellStyle.BorderTop = HSSFCellStyle.BORDER_THIN;
            GreyBottomMediumRightBorderCellStyle.TopBorderColor = HSSFColor.BLACK.index;
            GreyBottomMediumRightBorderCellStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.GREY_25_PERCENT.index;
            GreyBottomMediumRightBorderCellStyle.FillPattern = HSSFCellStyle.SOLID_FOREGROUND;
            GreyBottomMediumRightBorderCellStyle.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.YELLOW.index;
            GreyBottomMediumRightBorderCellStyle.SetFont(SimpleFont);

            HSSFCellStyle GreyLeftMediumBorderCellStyle = hssfworkbook.CreateCellStyle();
            GreyLeftMediumBorderCellStyle.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            GreyLeftMediumBorderCellStyle.BottomBorderColor = HSSFColor.BLACK.index;
            GreyLeftMediumBorderCellStyle.BorderLeft = HSSFCellStyle.BORDER_MEDIUM;
            GreyLeftMediumBorderCellStyle.LeftBorderColor = HSSFColor.BLACK.index;
            GreyLeftMediumBorderCellStyle.BorderRight = HSSFCellStyle.BORDER_THIN;
            GreyLeftMediumBorderCellStyle.RightBorderColor = HSSFColor.BLACK.index;
            GreyLeftMediumBorderCellStyle.BorderTop = HSSFCellStyle.BORDER_THIN;
            GreyLeftMediumBorderCellStyle.TopBorderColor = HSSFColor.BLACK.index;
            GreyLeftMediumBorderCellStyle.Alignment = HSSFCellStyle.ALIGN_CENTER;
            GreyLeftMediumBorderCellStyle.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
            GreyLeftMediumBorderCellStyle.SetFont(PackNumberFont);

            HSSFCellStyle GreyRightMediumBorderCellStyle = hssfworkbook.CreateCellStyle();
            GreyRightMediumBorderCellStyle.BorderBottom = HSSFCellStyle.BORDER_THIN;
            GreyRightMediumBorderCellStyle.BottomBorderColor = HSSFColor.BLACK.index;
            GreyRightMediumBorderCellStyle.BorderLeft = HSSFCellStyle.BORDER_THIN;
            GreyRightMediumBorderCellStyle.LeftBorderColor = HSSFColor.BLACK.index;
            GreyRightMediumBorderCellStyle.BorderRight = HSSFCellStyle.BORDER_MEDIUM;
            GreyRightMediumBorderCellStyle.RightBorderColor = HSSFColor.BLACK.index;
            GreyRightMediumBorderCellStyle.BorderTop = HSSFCellStyle.BORDER_THIN;
            GreyRightMediumBorderCellStyle.TopBorderColor = HSSFColor.BLACK.index;
            GreyRightMediumBorderCellStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.GREY_25_PERCENT.index;
            GreyRightMediumBorderCellStyle.FillPattern = HSSFCellStyle.SOLID_FOREGROUND;
            GreyRightMediumBorderCellStyle.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.YELLOW.index;
            GreyRightMediumBorderCellStyle.SetFont(SimpleFont);

            #endregion границы между упаковками

            //HSSFCell ConfirmCell = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 0, s);
            //ConfirmCell.CellStyle = TempStyle;

            if (FrontsOrdersDataTable.Rows.Count > 0)
                FrontsReport(hssfworkbook, HeaderStyle, PackNumberFont, SimpleFont, SimpleCellStyle, cellStyle, FrontsOrdersDataTable);

            if (DecorOrdersDataTable.Rows.Count > 0)
                DecorReport(hssfworkbook, HeaderStyle, SimpleFont, SimpleCellStyle, cellStyle, DecorOrdersDataTable);

            string ReportFilePath = string.Empty;

            //ReportFilePath = Application.StartupPath + @"\" + "Отчеты";

            //if (!(Directory.Exists(ReportFilePath)))
            //{
            //    Directory.CreateDirectory(ReportFilePath);
            //}

            //ReportFilePath = Application.StartupPath + @"\Отчеты\" + @"Статистика\";

            //if (!(Directory.Exists(ReportFilePath)))
            //{
            //    Directory.CreateDirectory(ReportFilePath);
            //}

            //ReportFilePath = ReadReportFilePath("StatisticsReportPath.config");
            //FileInfo file = new FileInfo(ReportFilePath + FileName + ".xls");

            //int j = 1;
            //while (file.Exists == true)
            //{
            //    file = new FileInfo(ReportFilePath + FileName + "(" + j++ + ").xls");
            //}

            string tempFolder = System.Environment.GetEnvironmentVariable("TEMP");
            FileInfo file = new FileInfo(tempFolder + @"\" + FileName + ".xls");
            int j = 1;
            while (file.Exists == true)
            {
                file = new FileInfo(tempFolder + @"\" + FileName + "(" + j++ + ").xls");
            }

            FileStream NewFile = new FileStream(file.FullName, FileMode.Create);
            hssfworkbook.Write(NewFile);
            NewFile.Close();

            System.Diagnostics.Process.Start(file.FullName);
        }

        private int FrontsReport(HSSFWorkbook hssfworkbook, HSSFCellStyle HeaderStyle, HSSFFont PackNumberFont,
            HSSFFont SimpleFont, HSSFCellStyle SimpleCellStyle, HSSFCellStyle cellStyle,
            DataTable FrontsOrdersDataTable)
        {
            int RowIndex = 0;

            HSSFSheet sheet1 = hssfworkbook.CreateSheet("Фасады");
            sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

            sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
            sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
            sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
            sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

            int DisplayIndex = 0;
            sheet1.SetColumnWidth(DisplayIndex++, 32 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 28 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 18 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 14 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 17 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 18 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 15 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 16 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 6 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 6 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 8 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 6 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 16 * 256);
            sheet1.SetColumnWidth(DisplayIndex++, 14 * 256);

            decimal Square = 0;
            int FrontsCount = 0;
            int CurvedCount = 0;

            FillFronts(FrontsOrdersDataTable);

            if (FrontsResultDataTable.Rows.Count != 0)
            {
                DisplayIndex = 0;
                HSSFCell cell4 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Фасад");
                cell4.CellStyle = HeaderStyle;
                cell4 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Цвет профиля");
                cell4.CellStyle = HeaderStyle;
                cell4 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Цвет профиля-2");
                cell4.CellStyle = HeaderStyle;
                cell4 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Патина");
                cell4.CellStyle = HeaderStyle;
                cell4 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Тип наполнителя");
                cell4.CellStyle = HeaderStyle;
                cell4 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Цвет наполнителя");
                cell4.CellStyle = HeaderStyle;
                cell4 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Тип наполнителя-2");
                cell4.CellStyle = HeaderStyle;
                cell4 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Цвет наполнителя-2");
                cell4.CellStyle = HeaderStyle;
                cell4 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Выс.");
                cell4.CellStyle = HeaderStyle;
                cell4 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Шир.");
                cell4.CellStyle = HeaderStyle;
                cell4 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Площ.");
                cell4.CellStyle = HeaderStyle;
                cell4 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Кол.");
                cell4.CellStyle = HeaderStyle;
                if (Security.PriceAccess)
                {
                    HSSFCell cell15 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Ср.цена, евро");
                    cell15.CellStyle = HeaderStyle;
                    HSSFCell cell16 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), DisplayIndex++, "Сумма, евро");
                    cell16.CellStyle = HeaderStyle;
                }

                RowIndex++;

                Square = GetSquare(FrontsResultDataTable);
                FrontsCount = GetCount(FrontsResultDataTable, false);
                CurvedCount = GetCount(FrontsResultDataTable, true);

                int ColumnCount = FrontsResultDataTable.Columns.Count;
                if (!Security.PriceAccess)
                {
                    ColumnCount = FrontsResultDataTable.Columns.Count - 2;
                }

                //вывод заказов фасадов
                for (int x = 0; x < FrontsResultDataTable.Rows.Count; x++)
                {
                    if (FrontsResultDataTable.Rows.Count == 0)
                        break;

                    for (int y = 0; y < ColumnCount; y++)
                    {
                        Type t = FrontsResultDataTable.Rows[x][y].GetType();

                        if (t.Name == "Decimal")
                        {
                            HSSFCell cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(Convert.ToDouble(FrontsResultDataTable.Rows[x][y]));

                            cell.CellStyle = cellStyle;
                            continue;
                        }
                        if (t.Name == "Int32")
                        {
                            HSSFCell cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(Convert.ToInt32(FrontsResultDataTable.Rows[x][y]));
                            cell.CellStyle = SimpleCellStyle;
                            continue;
                        }

                        if (t.Name == "String" || t.Name == "DBNull")
                        {
                            HSSFCell cell = sheet1.CreateRow(RowIndex).CreateCell(y);
                            cell.SetCellValue(FrontsResultDataTable.Rows[x][y].ToString());
                            cell.CellStyle = SimpleCellStyle;
                            continue;
                        }
                    }
                    RowIndex++;
                }

                RowIndex++;

                HSSFCellStyle cellStyle1 = hssfworkbook.CreateCellStyle();
                cellStyle1.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.000");
                cellStyle1.SetFont(SimpleFont);

                HSSFCell cell17 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Итого: ");
                cell17.CellStyle = cellStyle1;
                HSSFCell cell18 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, "всего фасадов: " + FrontsCount + " шт.");
                cell18.CellStyle = cellStyle1;
                HSSFCell cell19 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex++), 1, "в том числе гнутых: " + CurvedCount + " шт.");
                cell19.CellStyle = cellStyle1;

                if (Square > 0)
                {
                    HSSFCell cell20 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1,
                        "площадь: " + Decimal.Round(Square, 3, MidpointRounding.AwayFromZero) + " м.кв.");

                    cell20.CellStyle = cellStyle1;
                }

                RowIndex++;
            }
            return RowIndex;
        }

        private int DecorReport(HSSFWorkbook hssfworkbook, HSSFCellStyle HeaderStyle,
            HSSFFont SimpleFont, HSSFCellStyle SimpleCellStyle, HSSFCellStyle cellStyle,
            DataTable DecorOrdersDataTable)
        {
            int RowIndex = 0;

            HSSFSheet sheet1 = hssfworkbook.CreateSheet("Декор");
            sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

            sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
            sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
            sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
            sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

            sheet1.SetColumnWidth(0, 20 * 256);
            sheet1.SetColumnWidth(1, 10 * 256);
            sheet1.SetColumnWidth(2, 26 * 256);
            sheet1.SetColumnWidth(3, 15 * 256);
            sheet1.SetColumnWidth(4, 15 * 256);
            sheet1.SetColumnWidth(5, 9 * 256);
            if (Security.PriceAccess)
            {
                sheet1.SetColumnWidth(6, 17 * 256);
                sheet1.SetColumnWidth(7, 15 * 256);
            }
            //декор

            FillDecor(DecorOrdersDataTable);

            for (int c = 0; c < DecorCatalog.DecorProductsCount; c++)
            {
                if (DecorResultDataTable[c].Rows.Count == 0)
                    continue;
                HSSFCell cell15 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 0, "Название");
                cell15.CellStyle = HeaderStyle;
                HSSFCell cell16 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 1, "");
                cell16.CellStyle = HeaderStyle;
                HSSFCell cell17 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 2, "Цвет");
                cell17.CellStyle = HeaderStyle;
                HSSFCell cell18 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 3, "Длин\\Выс.");
                cell18.CellStyle = HeaderStyle;
                HSSFCell cell19 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 4, "Шир.");
                cell19.CellStyle = HeaderStyle;
                HSSFCell cell20 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 5, "Кол.");
                cell20.CellStyle = HeaderStyle;
                if (Security.PriceAccess)
                {
                    HSSFCell cell22 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 6, "Ср.цена, евро");
                    cell22.CellStyle = HeaderStyle;
                    HSSFCell cell23 = HSSFCellUtil.CreateCell(sheet1.CreateRow(RowIndex), 7, "Сумма, евро");
                    cell23.CellStyle = HeaderStyle;
                }

                RowIndex++;

                int ColumnCount = DecorResultDataTable[c].Columns.Count;
                if (!Security.PriceAccess)
                {
                    ColumnCount = DecorResultDataTable[c].Columns.Count - 2;
                }

                //вывод заказов декора в excel
                for (int x = 0; x < DecorResultDataTable[c].Rows.Count; x++)
                {
                    for (int y = 0; y < ColumnCount; y++)
                    {
                        int ColumnIndex = y;

                        if (y == 0)
                        {
                            ColumnIndex = y;
                        }
                        else
                        {
                            ColumnIndex = y + 1;
                        }

                        Type t = DecorResultDataTable[c].Rows[x][y].GetType();

                        sheet1.CreateRow(RowIndex).CreateCell(1).CellStyle = SimpleCellStyle;

                        if (t.Name == "Decimal")
                        {
                            HSSFCell cell = sheet1.CreateRow(RowIndex).CreateCell(ColumnIndex);
                            cell.SetCellValue(Convert.ToDouble(DecorResultDataTable[c].Rows[x][y]));

                            cell.CellStyle = SimpleCellStyle;
                            continue;
                        }
                        if (t.Name == "Int32")
                        {
                            HSSFCell cell = sheet1.CreateRow(RowIndex).CreateCell(ColumnIndex);
                            cell.SetCellValue(Convert.ToInt32(DecorResultDataTable[c].Rows[x][y]));
                            cell.CellStyle = SimpleCellStyle;
                            continue;
                        }

                        if (t.Name == "String" || t.Name == "DBNull")
                        {
                            HSSFCell cell = sheet1.CreateRow(RowIndex).CreateCell(ColumnIndex);
                            cell.SetCellValue(DecorResultDataTable[c].Rows[x][y].ToString());
                            cell.CellStyle = SimpleCellStyle;
                            continue;
                        }
                    }

                    RowIndex++;
                }
                RowIndex++;
            }
            return RowIndex;
        }

        //private string ReadReportFilePath(string FileName)
        //{
        //    string ReportFilePath = string.Empty;

        //    using (System.IO.StreamReader sr = new System.IO.StreamReader(FileName, Encoding.Default))
        //    {
        //        ReportFilePath = sr.ReadToEnd();
        //    }
        //    return ReportFilePath;
        //}
    }
}