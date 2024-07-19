using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;

using System;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Windows.Forms;

namespace Infinium.Modules.Marketing.NewOrders.PrepareReport
{
    public class DetailsReport : IAllFrontParameterName, IIsMarsel
    {
        private StandardReport.Report ClientReport = null;

        //string ReportFilePath = null;

        public bool Save = false;
        public bool Send = false;

        private DataTable ClientsDataTable = null;

        private DataTable FrontsResultDataTable = null;
        private DataTable[] DecorResultDataTable = null;

        public DataTable[] ClientReportTables = null;

        public DataTable ReportTable = null;

        private FrontsCalculate FrontsCalculate = null;

        private DataTable FrontsOrdersDataTable = null;
        private DataTable DecorOrdersDataTable = null;

        public DataTable CurrencyTypesDataTable = null;
        public DataTable FrontsDataTable = null;
        public DataTable FrameColorsDataTable = null;
        public DataTable PatinaDataTable = null;
        private DataTable PatinaRALDataTable = null;
        public DataTable InsetTypesDataTable = null;
        public DataTable InsetColorsDataTable = null;
        public DataTable TechnoInsetTypesDataTable = null;
        public DataTable TechnoInsetColorsDataTable = null;

        private DecorCatalogOrder DecorCatalogOrder = null;

        public DetailsReport(
            ref DecorCatalogOrder tDecorCatalogOrder,
            ref FrontsCalculate tFrontsCalculate)
        {
            DecorCatalogOrder = tDecorCatalogOrder;
            FrontsCalculate = tFrontsCalculate;

            Create();
            string SelectCommand = @"SELECT TechStoreID AS FrontID, TechStoreName AS FrontName FROM TechStore
                WHERE TechStoreID IN (SELECT FrontID FROM FrontsConfig WHERE Enabled = 1)
                ORDER BY TechStoreName";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(FrontsDataTable);
            }
            GetColorsDT();
            GetInsetColorsDT();
            SelectCommand = @"SELECT * FROM Patina";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.CatalogConnectionString))
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
            SelectCommand = @"SELECT * FROM InsetTypes";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(InsetTypesDataTable);
            }
            //SelectCommand = @"SELECT * FROM InsetColors";
            //using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.CatalogConnectionString))
            //{
            //    DA.Fill(InsetColorsDataTable);
            //}
            TechnoInsetTypesDataTable = InsetTypesDataTable.Copy();
            TechnoInsetColorsDataTable = InsetColorsDataTable.Copy();

            ClientsDataTable = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM Clients",
                ConnectionStrings.MarketingReferenceConnectionString))
            {
                DA.Fill(ClientsDataTable);
            }
            CreateFrontsDataTable();
            CreateDecorDataTable();

            //ReadReportFilePath("MarketingClientReportPath.config");

            //if (!(Directory.Exists(ReportFilePath)))
            //{
            //    Directory.CreateDirectory(ReportFilePath);
            //}

            //if (!(Directory.Exists(ReportFilePath)))
            //{
            //    Directory.CreateDirectory(ReportFilePath);
            //}

            //string FileName = InventoryName + ".xls";

            //FileName = Path.Combine(ReportFilePath, FileName);

            //int DocNumber = 1;

            //while (File.Exists(FileName))
            //{
            //    FileName = InventoryName + "(" + DocNumber++ + ").xls";
            //    FileName = Path.Combine(ReportFilePath, FileName);
            //}

            //FileInfo file = new FileInfo(FileName);

            //FileStream NewFile = new FileStream(file.FullName, FileMode.Create);
            //hssfworkbook.Write(NewFile);
            //NewFile.Close();

            //System.Diagnostics.Process.Start(file.FullName);
        }

        private void Create()
        {
            FrontsOrdersDataTable = new DataTable();

            FrontsDataTable = new DataTable();
            FrameColorsDataTable = new DataTable();
            PatinaDataTable = new DataTable();
            InsetTypesDataTable = new DataTable();
            InsetColorsDataTable = new DataTable();
            TechnoInsetTypesDataTable = new DataTable();
            TechnoInsetColorsDataTable = new DataTable();

            FrontsResultDataTable = new DataTable();
            CurrencyTypesDataTable = new DataTable();
            DecorResultDataTable = new DataTable[DecorCatalogOrder.DecorProductsCount];
            DecorOrdersDataTable = new DataTable();
        }

        private void GetColorsDT()
        {
            FrameColorsDataTable.Columns.Add(new DataColumn("ColorID", Type.GetType("System.Int64")));
            FrameColorsDataTable.Columns.Add(new DataColumn("ColorName", Type.GetType("System.String")));
            string SelectCommand = @"SELECT TechStoreID, TechStoreName FROM TechStore
                WHERE TechStoreSubGroupID IN (SELECT TechStoreSubGroupID FROM TechStoreSubGroups WHERE (TechStoreGroupID = 11 OR TechStoreGroupID = 1))
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

        private void CreateFrontsDataTable()
        {
            FrontsResultDataTable = new DataTable();

            FrontsResultDataTable.Columns.Add(new DataColumn(("FrontName"), System.Type.GetType("System.String")));
            FrontsResultDataTable.Columns.Add(new DataColumn(("FrameColor1"), System.Type.GetType("System.String")));
            FrontsResultDataTable.Columns.Add(new DataColumn(("FrameColor2"), System.Type.GetType("System.String")));
            FrontsResultDataTable.Columns.Add(new DataColumn(("InsetType1"), System.Type.GetType("System.String")));
            FrontsResultDataTable.Columns.Add(new DataColumn(("InsetColor1"), System.Type.GetType("System.String")));
            FrontsResultDataTable.Columns.Add(new DataColumn(("InsetType2"), System.Type.GetType("System.String")));
            FrontsResultDataTable.Columns.Add(new DataColumn(("InsetColor2"), System.Type.GetType("System.String")));
            FrontsResultDataTable.Columns.Add(new DataColumn(("Patina"), System.Type.GetType("System.String")));
            FrontsResultDataTable.Columns.Add(new DataColumn(("Height"), System.Type.GetType("System.Int32")));
            FrontsResultDataTable.Columns.Add(new DataColumn(("Width"), System.Type.GetType("System.Int32")));
            FrontsResultDataTable.Columns.Add(new DataColumn(("Count"), System.Type.GetType("System.Int32")));
            FrontsResultDataTable.Columns.Add(new DataColumn(("Weight"), System.Type.GetType("System.Decimal")));
            FrontsResultDataTable.Columns.Add(new DataColumn(("FrontPrice"), System.Type.GetType("System.Decimal")));
            FrontsResultDataTable.Columns.Add(new DataColumn(("InsetPrice"), System.Type.GetType("System.Decimal")));
            FrontsResultDataTable.Columns.Add(new DataColumn(("Cost"), System.Type.GetType("System.Decimal")));
            FrontsResultDataTable.Columns.Add(new DataColumn(("Rate"), System.Type.GetType("System.Decimal")));
            FrontsResultDataTable.Columns.Add(new DataColumn("Notes", Type.GetType("System.String")));
        }

        private void CreateDecorDataTable()
        {
            for (int i = 0; i < DecorCatalogOrder.DecorProductsCount; i++)
            {
                DecorResultDataTable[i] = new DataTable();

                DecorResultDataTable[i].Columns.Add(new DataColumn(("Name"), Type.GetType("System.String")));
                DecorResultDataTable[i].Columns.Add(new DataColumn(("Height"), Type.GetType("System.Int32")));
                DecorResultDataTable[i].Columns.Add(new DataColumn(("Length"), Type.GetType("System.Int32")));
                DecorResultDataTable[i].Columns.Add(new DataColumn(("Width"), Type.GetType("System.Int32")));
                DecorResultDataTable[i].Columns.Add(new DataColumn(("Color"), Type.GetType("System.String")));
                DecorResultDataTable[i].Columns.Add(new DataColumn(("Patina"), Type.GetType("System.String")));

                DecorResultDataTable[i].Columns.Add(new DataColumn(("Count"), System.Type.GetType("System.Int32")));
                DecorResultDataTable[i].Columns.Add(new DataColumn(("Weight"), System.Type.GetType("System.Decimal")));
                DecorResultDataTable[i].Columns.Add(new DataColumn(("Price"), System.Type.GetType("System.Decimal")));
                DecorResultDataTable[i].Columns.Add(new DataColumn(("Cost"), System.Type.GetType("System.Decimal")));
                DecorResultDataTable[i].Columns.Add(new DataColumn(("Rate"), System.Type.GetType("System.Decimal")));
                DecorResultDataTable[i].Columns.Add(new DataColumn(("Notes"), Type.GetType("System.String")));
            }
        }

        private string GetMainOrderNotes(int MainOrderID)
        {
            string Notes = null;
            using (DataTable DT = new DataTable())
            {
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM NewMainOrders WHERE MainOrderID=" +
                    MainOrderID, ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DA.Fill(DT);
                    if (DT.Rows.Count > 0)
                        Notes = DT.Rows[0]["Notes"].ToString();
                }
            }
            return Notes;
        }

        private int GetOrderNumber(int MainOrderID)
        {
            int OrderNumber = 0;
            using (DataTable DT = new DataTable())
            {
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM NewMegaOrders WHERE MegaOrderID=" +
                    "(SELECT MegaOrderID FROM NewMainOrders WHERE MainOrderID=" + MainOrderID + ")", ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DA.Fill(DT);
                    if (DT.Rows.Count > 0)
                        OrderNumber = Convert.ToInt32(DT.Rows[0]["OrderNumber"]);
                }
            }
            return OrderNumber;
        }

        private int GetClientID(int MainOrderID)
        {
            int ClientID = -1;
            using (DataTable DT = new DataTable())
            {
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM NewMegaOrders WHERE MegaOrderID=" +
                    "(SELECT MegaOrderID FROM NewMainOrders WHERE MainOrderID=" + MainOrderID + ")", ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DA.Fill(DT);

                    try
                    {
                        if (DT.Rows.Count > 0)
                            ClientID = Convert.ToInt32(DT.Rows[0]["ClientID"]);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
            }
            return ClientID;
        }

        private string GetClientName(int MainOrderID)
        {
            string ClientName = "unnamed_client";
            try
            {
                DataRow[] Rows = ClientsDataTable.Select("ClientID = " + GetClientID(MainOrderID));
                if (Rows.Count() > 0)
                    ClientName = Rows[0]["ClientName"].ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            return ClientName;
        }

        public bool IsMarsel3(int FrontID)
        {
            return FrontID == 3630;
        }

        public bool IsMarsel4(int FrontID)
        {
            return FrontID == 15003;
        }

        public bool IsImpost(int TechnoProfileID)
        {
            return TechnoProfileID == -1;
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

        private void FillFronts(int MainOrderID)
        {
            FrontsResultDataTable.Clear();
            foreach (DataRow Row in FrontsOrdersDataTable.Rows)
            {
                DataRow NewRow = FrontsResultDataTable.NewRow();

                string FrameColor1 = GetColorName(Convert.ToInt32(Row["ColorID"]));
                string FrameColor2 = string.Empty;

                var InsetType = GetInsetTypeName(Convert.ToInt32(Row["InsetTypeID"]));
                var bMarsel3 = IsMarsel3(Convert.ToInt32(Row["FrontID"]));
                var bMarsel4 = IsMarsel4(Convert.ToInt32(Row["FrontID"]));
                if (bMarsel3 || bMarsel4)
                {
                    var bImpost = IsImpost(Convert.ToInt32(Row["TechnoProfileID"]));
                    if (bImpost)
                    {
                        string Front2 = GetFront2Name(Convert.ToInt32(Row["TechnoProfileID"]));
                        if (Front2.Length > 0)
                            InsetType = InsetType + "/" + Front2;
                    }
                }
                string InsetType2 = string.Empty;
                string InsetColor1 = GetInsetColorName(Convert.ToInt32(Row["InsetColorID"]));
                string InsetColor2 = string.Empty;
                string PatinaName = GetPatinaName(Convert.ToInt32(Row["PatinaID"]));

                if (Convert.ToInt32(Row["TechnoColorID"]) != -1)
                    FrameColor2 = GetColorName(Convert.ToInt32(Row["TechnoColorID"]));
                if (Convert.ToInt32(Row["TechnoInsetTypeID"]) != -1)
                    InsetType2 = GetInsetTypeName(Convert.ToInt32(Row["TechnoInsetTypeID"]));
                if (Convert.ToInt32(Row["TechnoInsetColorID"]) != -1)
                    InsetColor2 = GetInsetColorName(Convert.ToInt32(Row["TechnoInsetColorID"]));

                NewRow["FrontName"] = GetFrontName(Convert.ToInt32(Row["FrontID"]));
                NewRow["FrameColor1"] = FrameColor1;
                NewRow["FrameColor2"] = FrameColor2;
                NewRow["Patina"] = PatinaName;
                NewRow["InsetType1"] = InsetType;
                NewRow["InsetColor1"] = InsetColor1;
                NewRow["InsetType2"] = InsetType2;
                NewRow["InsetColor2"] = InsetColor2;
                NewRow["Weight"] = Row["Weight"];
                NewRow["Height"] = Row["Height"];
                NewRow["Width"] = Row["Width"];
                NewRow["Count"] = Convert.ToInt32(Row["Count"]);
                NewRow["FrontPrice"] = Decimal.Round(Convert.ToDecimal(Row["FrontPrice"]), 2, MidpointRounding.AwayFromZero);
                NewRow["InsetPrice"] = Decimal.Round(Convert.ToDecimal(Row["InsetPrice"]), 2, MidpointRounding.AwayFromZero);
                NewRow["Cost"] = Decimal.Round(Convert.ToDecimal(Row["Cost"]), 2, MidpointRounding.AwayFromZero);
                NewRow["Rate"] = Row["PaymentRate"];
                NewRow["Notes"] = Row["Notes"];
                FrontsResultDataTable.Rows.Add(NewRow);
            }
        }

        private void FillDecor(int MainOrderID)
        {
            for (int i = 0; i < DecorCatalogOrder.DecorProductsCount; i++)
            {
                DecorResultDataTable[i].Clear();
                DecorResultDataTable[i].AcceptChanges();

                DataRow[] Rows = DecorOrdersDataTable.Select("ProductID = " + DecorCatalogOrder.DecorProductsDataTable.Rows[i]["ProductID"]);

                if (Rows.Count() == 0)
                    continue;

                foreach (DataRow Row in Rows)
                {
                    DataRow NewRow2 = DecorResultDataTable[i].NewRow();

                    NewRow2["Name"] = DecorCatalogOrder.DecorProductsDataTable.Rows[i]["ProductName"].ToString() + " " +
                                            DecorCatalogOrder.GetItemName(Convert.ToInt32(Row["DecorID"]));

                    //if (DecorCatalogOrder.HasParameter(Convert.ToInt32(DecorCatalogOrder.DecorProductsDataTable.Rows[i]["ProductID"]), "Height"))
                    //    NewRow2["Height"] = Convert.ToInt32(Row["Height"]);
                    //else
                    if (Convert.ToInt32(Row["Height"]) != -1)
                        NewRow2["Height"] = Convert.ToInt32(Row["Height"]);

                    //if (DecorCatalogOrder.HasParameter(Convert.ToInt32(DecorCatalogOrder.DecorProductsDataTable.Rows[i]["ProductID"]), "Length"))
                    //    NewRow2["Height"] = Convert.ToInt32(Row["Length"]);
                    //else
                    if (Convert.ToInt32(Row["Length"]) != -1)
                        NewRow2["Length"] = Convert.ToInt32(Row["Length"]);

                    //if (DecorCatalogOrder.HasParameter(Convert.ToInt32(DecorCatalogOrder.DecorProductsDataTable.Rows[i]["ProductID"]), "Width"))
                    //    NewRow2["Width"] = Convert.ToInt32(Row["Width"]);
                    //else
                    if (Convert.ToInt32(Row["Width"]) != -1)
                        NewRow2["Width"] = Convert.ToInt32(Row["Width"]);

                    //if (DecorCatalogOrder.HasParameter(Convert.ToInt32(DecorCatalogOrder.DecorProductsDataTable.Rows[i]["ProductID"]), "Length") && Convert.ToInt32(Row["Length"]) != -1)
                    //    NewRow2["Height"] = Convert.ToInt32(Row["Length"]);

                    //if (DecorCatalogOrder.HasParameter(Convert.ToInt32(DecorCatalogOrder.DecorProductsDataTable.Rows[i]["ProductID"]), "Width") && Convert.ToInt32(Row["Width"]) != -1)
                    //    NewRow2["Width"] = Convert.ToInt32(Row["Width"]);

                    if (DecorCatalogOrder.HasParameter(Convert.ToInt32(DecorCatalogOrder.DecorProductsDataTable.Rows[i]["ProductID"]), "ColorID"))
                        NewRow2["Color"] = GetColorName(Convert.ToInt32(Row["ColorID"]));

                    NewRow2["Patina"] = GetPatinaName(Convert.ToInt32(Row["PatinaID"]));
                    int Count = Convert.ToInt32(Row["Count"]);
                    NewRow2["Count"] = Convert.ToInt32(Row["Count"]);
                    NewRow2["Price"] = Decimal.Round(Convert.ToDecimal(Row["Price"]), 2, MidpointRounding.AwayFromZero);
                    NewRow2["Cost"] = Decimal.Round(Convert.ToDecimal(Row["Cost"]), 2, MidpointRounding.AwayFromZero);
                    NewRow2["Rate"] = Row["PaymentRate"];
                    NewRow2["Notes"] = Row["Notes"];
                    NewRow2["Weight"] = Row["Weight"];
                    DecorResultDataTable[i].Rows.Add(NewRow2);
                }
            }
        }

        private bool FilterOrders(int MainOrderID, bool IsSample)
        {
            bool IsNotEmpty = false;

            FrontsOrdersDataTable.Clear();
            FrontsOrdersDataTable.AcceptChanges();
            string SelectCommand = "SELECT NewFrontsOrders.*, NewMegaOrders.PaymentRate FROM NewFrontsOrders INNER JOIN NewMainOrders ON NewFrontsOrders.MainOrderID = NewMainOrders.MainOrderID AND NewFrontsOrders.MainOrderID=" + MainOrderID +
                " INNER JOIN NewMegaOrders ON NewMainOrders.MegaOrderID = NewMegaOrders.MegaOrderID WHERE NewFrontsOrders.IsSample=1 ";
            if (!IsSample)
                SelectCommand = "SELECT NewFrontsOrders.*, NewMegaOrders.PaymentRate FROM NewFrontsOrders INNER JOIN NewMainOrders ON NewFrontsOrders.MainOrderID = NewMainOrders.MainOrderID AND NewFrontsOrders.MainOrderID=" + MainOrderID +
                " INNER JOIN NewMegaOrders ON NewMainOrders.MegaOrderID = NewMegaOrders.MegaOrderID WHERE NewFrontsOrders.IsSample=0 ";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                if (DA.Fill(FrontsOrdersDataTable) > 0)
                    IsNotEmpty = true;
            }

            DecorOrdersDataTable.Clear();
            DecorOrdersDataTable.AcceptChanges();
            SelectCommand = "SELECT NewDecorOrders.*, NewMegaOrders.PaymentRate FROM NewDecorOrders INNER JOIN NewMainOrders ON NewDecorOrders.MainOrderID = NewMainOrders.MainOrderID AND NewDecorOrders.MainOrderID=" + MainOrderID +
                " INNER JOIN NewMegaOrders ON NewMainOrders.MegaOrderID = NewMegaOrders.MegaOrderID WHERE NewDecorOrders.IsSample=1 ";
            if (!IsSample)
                SelectCommand = "SELECT NewDecorOrders.*, NewMegaOrders.PaymentRate FROM NewDecorOrders INNER JOIN NewMainOrders ON NewDecorOrders.MainOrderID = NewMainOrders.MainOrderID AND NewDecorOrders.MainOrderID=" + MainOrderID +
                " INNER JOIN NewMegaOrders ON NewMainOrders.MegaOrderID = NewMegaOrders.MegaOrderID WHERE NewDecorOrders.IsSample=0 ";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                if (DA.Fill(DecorOrdersDataTable) > 0)
                    IsNotEmpty = true;
            }

            return IsNotEmpty;
        }

        private bool FilterOrders(int MainOrderID)
        {
            bool IsNotEmpty = false;

            FrontsOrdersDataTable.Clear();
            FrontsOrdersDataTable.AcceptChanges();
            string SelectCommand = "SELECT NewFrontsOrders.*, NewMegaOrders.PaymentRate FROM NewFrontsOrders INNER JOIN NewMainOrders ON NewFrontsOrders.MainOrderID = NewMainOrders.MainOrderID AND NewFrontsOrders.MainOrderID=" + MainOrderID +
                " INNER JOIN NewMegaOrders ON NewMainOrders.MegaOrderID = NewMegaOrders.MegaOrderID";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                if (DA.Fill(FrontsOrdersDataTable) > 0)
                    IsNotEmpty = true;
            }

            DecorOrdersDataTable.Clear();
            DecorOrdersDataTable.AcceptChanges();
            SelectCommand = "SELECT NewDecorOrders.*, NewMegaOrders.PaymentRate FROM NewDecorOrders INNER JOIN NewMainOrders ON NewDecorOrders.MainOrderID = NewMainOrders.MainOrderID AND NewDecorOrders.MainOrderID=" + MainOrderID +
                " INNER JOIN NewMegaOrders ON NewMainOrders.MegaOrderID = NewMegaOrders.MegaOrderID";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                if (DA.Fill(DecorOrdersDataTable) > 0)
                    IsNotEmpty = true;
            }

            return IsNotEmpty;
        }

        private decimal CalculateCost()
        {
            decimal FrontsCost = 0;
            decimal DecorCost = 0;
            decimal OrderCost = 0;

            foreach (DataRow rows1 in FrontsOrdersDataTable.Rows)
                FrontsCost += Convert.ToDecimal(rows1["CurrencyCost"]);

            foreach (DataRow rows2 in DecorOrdersDataTable.Rows)
                DecorCost += Convert.ToDecimal(rows2["CurrencyCost"]);

            OrderCost = FrontsCost + DecorCost;
            return OrderCost;
        }

        private int GetMegaOrderID(int MainOrderID)
        {
            int MegaOrderID = 0;
            using (DataTable DT = new DataTable())
            {
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT MegaOrderID FROM NewMainOrders" +
                    " WHERE MainOrderID=" + MainOrderID, ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DA.Fill(DT);
                    if (DT.Rows.Count > 0)
                        MegaOrderID = Convert.ToInt32(DT.Rows[0]["MegaOrderID"]);
                }
            }
            return MegaOrderID;
        }

        private bool IsComplaint(int MegaOrderID)
        {
            bool IsComplaint = false;

            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT IsComplaint" +
                " FROM NewMegaOrders WHERE MegaOrderID = " + MegaOrderID, ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (SqlCommandBuilder CB = new SqlCommandBuilder(DA))
                {
                    using (DataTable DT = new DataTable())
                    {
                        DA.Fill(DT);

                        if (DBNull.Value != DT.Rows[0]["IsComplaint"] && Convert.ToInt32(DT.Rows[0]["IsComplaint"]) > 0)
                        {
                            IsComplaint = Convert.ToBoolean(DT.Rows[0]["IsComplaint"]);
                        }
                    }
                }
            }

            return IsComplaint;
        }
        
        public void Report(ref HSSFWorkbook hssfworkbook, int[] NewMegaOrders, int[] OrderNumbers, int[] MainOrdersIDs, int ClientID, string ClientName,
            decimal ComplaintProfilCost, decimal ComplaintTPSCost,
            decimal TransportCost, decimal AdditionalCost,
            decimal TotalCost, int CurrencyTypeID, decimal TotalWeight, bool IsSample)
        {
            ClearReport();

            ClientReport = new StandardReport.Report(ref DecorCatalogOrder, ref FrontsCalculate);

            #region Create fonts and styles

            HSSFFont HeaderF1 = hssfworkbook.CreateFont();
            HeaderF1.FontHeightInPoints = 13;
            HeaderF1.Boldweight = 13 * 256;
            HeaderF1.FontName = "Calibri";

            HSSFFont HeaderF2 = hssfworkbook.CreateFont();
            HeaderF2.FontHeightInPoints = 12;
            HeaderF2.Boldweight = 12 * 256;
            HeaderF2.FontName = "Calibri";

            HSSFFont HeaderF3 = hssfworkbook.CreateFont();
            HeaderF3.FontHeightInPoints = 9;
            HeaderF3.Boldweight = 9 * 256;
            HeaderF3.FontName = "Calibri";

            HSSFFont SimpleF = hssfworkbook.CreateFont();
            SimpleF.FontHeightInPoints = 9;
            SimpleF.FontName = "Calibri";

            HSSFCellStyle SimpleCS = hssfworkbook.CreateCellStyle();
            SimpleCS.BorderBottom = HSSFCellStyle.BORDER_THIN;
            SimpleCS.BottomBorderColor = HSSFColor.BLACK.index;
            SimpleCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            SimpleCS.LeftBorderColor = HSSFColor.BLACK.index;
            SimpleCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            SimpleCS.RightBorderColor = HSSFColor.BLACK.index;
            SimpleCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            SimpleCS.TopBorderColor = HSSFColor.BLACK.index;
            SimpleCS.SetFont(SimpleF);

            HSSFCellStyle SimpleDecCS = hssfworkbook.CreateCellStyle();
            SimpleDecCS.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.00");
            SimpleDecCS.BorderBottom = HSSFCellStyle.BORDER_THIN;
            SimpleDecCS.BottomBorderColor = HSSFColor.BLACK.index;
            SimpleDecCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            SimpleDecCS.LeftBorderColor = HSSFColor.BLACK.index;
            SimpleDecCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            SimpleDecCS.RightBorderColor = HSSFColor.BLACK.index;
            SimpleDecCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            SimpleDecCS.TopBorderColor = HSSFColor.BLACK.index;
            SimpleDecCS.SetFont(SimpleF);

            HSSFCellStyle ReportCS = hssfworkbook.CreateCellStyle();
            ReportCS.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            ReportCS.BottomBorderColor = HSSFColor.BLACK.index;
            ReportCS.SetFont(HeaderF1);

            HSSFCellStyle HeaderWithoutBorderCS = hssfworkbook.CreateCellStyle();
            HeaderWithoutBorderCS.SetFont(HeaderF2);

            HSSFCellStyle HeaderWithBorderCS = hssfworkbook.CreateCellStyle();
            HeaderWithBorderCS.BorderBottom = HSSFCellStyle.BORDER_THIN;
            HeaderWithBorderCS.BottomBorderColor = HSSFColor.BLACK.index;
            HeaderWithBorderCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            HeaderWithBorderCS.LeftBorderColor = HSSFColor.BLACK.index;
            HeaderWithBorderCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            HeaderWithBorderCS.RightBorderColor = HSSFColor.BLACK.index;
            HeaderWithBorderCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            HeaderWithBorderCS.TopBorderColor = HSSFColor.BLACK.index;
            //HeaderWithBorderCS.WrapText = true;
            HeaderWithBorderCS.SetFont(HeaderF2);

            HSSFCellStyle SimpleHeaderCS = hssfworkbook.CreateCellStyle();
            //HeaderStyle.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
            //HeaderStyle.Alignment = HSSFCellStyle.ALIGN_CENTER;
            SimpleHeaderCS.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            SimpleHeaderCS.BottomBorderColor = HSSFColor.BLACK.index;
            SimpleHeaderCS.BorderLeft = HSSFCellStyle.BORDER_MEDIUM;
            SimpleHeaderCS.LeftBorderColor = HSSFColor.BLACK.index;
            SimpleHeaderCS.BorderRight = HSSFCellStyle.BORDER_MEDIUM;
            SimpleHeaderCS.RightBorderColor = HSSFColor.BLACK.index;
            SimpleHeaderCS.BorderTop = HSSFCellStyle.BORDER_MEDIUM;
            SimpleHeaderCS.TopBorderColor = HSSFColor.BLACK.index;
            //SimpleHeaderCS.WrapText = true;
            SimpleHeaderCS.SetFont(HeaderF3);

            #endregion Create fonts and styles

            HSSFSheet sheet1 = hssfworkbook.CreateSheet("Подробный отчет");
            sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

            sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
            sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
            sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
            sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

            sheet1.SetColumnWidth(0, 25 * 256);
            sheet1.SetColumnWidth(1, 15 * 256);
            sheet1.SetColumnWidth(2, 15 * 256);
            sheet1.SetColumnWidth(3, 15 * 256);
            sheet1.SetColumnWidth(4, 15 * 256);
            sheet1.SetColumnWidth(5, 15 * 256);
            sheet1.SetColumnWidth(6, 15 * 256);
            sheet1.SetColumnWidth(7, 18 * 256);
            sheet1.SetColumnWidth(8, 13 * 256);
            sheet1.SetColumnWidth(9, 10 * 256);
            sheet1.SetColumnWidth(10, 9 * 256);
            sheet1.SetColumnWidth(11, 12 * 256);
            sheet1.SetColumnWidth(12, 12 * 256);
            sheet1.SetColumnWidth(13, 13 * 256);
            sheet1.SetColumnWidth(14, 12 * 256);
            sheet1.SetColumnWidth(15, 10 * 256);

            decimal OrderCost = 0;

            int RowIndex = 0;

            string Currency = string.Empty;

            CurrencyTypesDataTable = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM CurrencyTypes",
                ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(CurrencyTypesDataTable);
            }

            DataRow[] Row = CurrencyTypesDataTable.Select("CurrencyTypeID = " + CurrencyTypeID);

            Currency = Row[0]["CurrencyType"].ToString();

            ClientName = GetClientName(MainOrdersIDs[0]);

            HSSFCell Cell1 = sheet1.CreateRow(RowIndex).CreateCell(0);
            Cell1.SetCellValue("Клиент:");
            Cell1.CellStyle = HeaderWithoutBorderCS;
            Cell1 = sheet1.CreateRow(RowIndex++).CreateCell(1);
            Cell1.SetCellValue(ClientName);
            Cell1.CellStyle = HeaderWithoutBorderCS;
            Cell1 = sheet1.CreateRow(RowIndex).CreateCell(0);
            Cell1.SetCellValue("№ заказа:");
            Cell1.CellStyle = HeaderWithoutBorderCS;
            Cell1 = sheet1.CreateRow(RowIndex++).CreateCell(1);
            Cell1.SetCellValue(string.Join(",", OrderNumbers));
            Cell1.CellStyle = HeaderWithoutBorderCS;
            if (IsComplaint(GetMegaOrderID(MainOrdersIDs[0])))
            {
                Cell1 = sheet1.CreateRow(RowIndex++).CreateCell(0);
                Cell1.SetCellValue("РЕКЛАМАЦИЯ");
                Cell1.CellStyle = HeaderWithoutBorderCS;
            }

            RowIndex++;
            for (int i = 0; i < MainOrdersIDs.Count(); i++)
            {
                OrderCost = 0;

                if (FilterOrders(MainOrdersIDs[i], IsSample))
                {
                    FillFronts(MainOrdersIDs[i]);
                    FillDecor(MainOrdersIDs[i]);

                    RowIndex++;

                    Cell1 = sheet1.CreateRow(RowIndex).CreateCell(0);
                    Cell1.SetCellValue("№" + GetOrderNumber(MainOrdersIDs[i]) + "-" + MainOrdersIDs[i].ToString());
                    Cell1.CellStyle = HeaderWithoutBorderCS;
                    RowIndex++;
                    Cell1 = sheet1.CreateRow(RowIndex).CreateCell(0);
                    Cell1.SetCellValue("Примечание к заказу: " + GetMainOrderNotes(MainOrdersIDs[i]));
                    Cell1.CellStyle = HeaderWithoutBorderCS;

                    int DisplayIndex = 0;
                    if (FrontsResultDataTable.Rows.Count != 0)
                    {
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue("Фасад");
                        Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue("Цвет профиля-1");
                        Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue("Цвет профиля-2");
                        Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue("Тип наполнителя-1");
                        Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue("Цвет наполнителя-1");
                        Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue("Тип наполнителя-2");
                        Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue("Цвет наполнителя-2");
                        Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue("Патина");
                        Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue("Выс.");
                        Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue("Шир.");
                        Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue("Кол.");
                        Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue("Вес");
                        Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue("Цена за фасад");
                        Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue("Цена за вставку");
                        Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue("Стоимость");
                        Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue("Курс, " + Currency);
                        Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue("Примечание");
                        Cell1.CellStyle = SimpleHeaderCS;

                        RowIndex++;
                    }

                    if (FrontsResultDataTable.Rows.Count != 0)
                        RowIndex++;
                    OrderCost = CalculateCost();
                    DisplayIndex = 0;
                    //вывод заказов фасадов
                    for (int x = 0; x < FrontsResultDataTable.Rows.Count; x++)
                    {
                        DisplayIndex = 0;
                        Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(FrontsResultDataTable.Rows[x]["FrontName"].ToString());
                        Cell1.CellStyle = SimpleCS;
                        Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(FrontsResultDataTable.Rows[x]["FrameColor1"].ToString());
                        Cell1.CellStyle = SimpleCS;
                        Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(FrontsResultDataTable.Rows[x]["FrameColor2"].ToString());
                        Cell1.CellStyle = SimpleCS;
                        Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(FrontsResultDataTable.Rows[x]["InsetType1"].ToString());
                        Cell1.CellStyle = SimpleCS;
                        Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(FrontsResultDataTable.Rows[x]["InsetColor1"].ToString());
                        Cell1.CellStyle = SimpleCS;
                        Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(FrontsResultDataTable.Rows[x]["InsetType2"].ToString());
                        Cell1.CellStyle = SimpleCS;
                        Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(FrontsResultDataTable.Rows[x]["InsetColor2"].ToString());
                        Cell1.CellStyle = SimpleCS;
                        Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(FrontsResultDataTable.Rows[x]["Patina"].ToString());
                        Cell1.CellStyle = SimpleCS;
                        Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(Convert.ToInt32(FrontsResultDataTable.Rows[x]["Height"]));
                        Cell1.CellStyle = SimpleCS;
                        Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(Convert.ToInt32(FrontsResultDataTable.Rows[x]["Width"]));
                        Cell1.CellStyle = SimpleCS;
                        Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(Convert.ToInt32(FrontsResultDataTable.Rows[x]["Count"]));
                        Cell1.CellStyle = SimpleCS;
                        Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                        {
                            decimal Weight = Convert.ToDecimal(FrontsResultDataTable.Rows[x]["Weight"]);
                            //decimal Weight = 0;
                            //if (Width != -1)
                            //    Weight = Convert.ToDecimal(FrontsResultDataTable.Rows[x]["Weight"]) +
                            //        Convert.ToDecimal(FrontsResultDataTable.Rows[x]["Height"]) * Convert.ToDecimal(FrontsResultDataTable.Rows[x]["Width"]) *
                            //        Convert.ToDecimal(FrontsResultDataTable.Rows[x]["Count"]) / 1000000 * Convert.ToDecimal(0.7);
                            //else
                            //    Weight = Convert.ToDecimal(FrontsResultDataTable.Rows[x]["Weight"]);

                            Weight = Decimal.Round(Weight, 3, MidpointRounding.AwayFromZero);
                            Cell1.SetCellValue(Convert.ToDouble(Weight));
                        }
                        Cell1.CellStyle = SimpleCS;
                        Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(Convert.ToDouble(FrontsResultDataTable.Rows[x]["FrontPrice"]));
                        Cell1.CellStyle = SimpleDecCS;
                        Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(Convert.ToDouble(FrontsResultDataTable.Rows[x]["InsetPrice"]));
                        Cell1.CellStyle = SimpleDecCS;
                        Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(Convert.ToDouble(FrontsResultDataTable.Rows[x]["Cost"]));
                        Cell1.CellStyle = SimpleDecCS;
                        Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(Convert.ToDouble(FrontsResultDataTable.Rows[x]["Rate"]));
                        Cell1.CellStyle = SimpleDecCS;
                        Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(FrontsResultDataTable.Rows[x]["Notes"].ToString());
                        Cell1.CellStyle = SimpleCS;

                        RowIndex++;
                    }
                    //if (FrontsResultDataTable.Rows.Count != 0)
                    //    RowIndex++;

                    DisplayIndex = 0;
                    //декор
                    for (int c = 0; c < DecorCatalogOrder.DecorProductsCount; c++)
                    {
                        if (DecorResultDataTable[c].Rows.Count == 0)
                            continue;
                        //int ColumnCount = 0;

                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(0);
                        Cell1.SetCellValue("Наименование");
                        Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(1);
                        Cell1.SetCellValue("Цвет");
                        Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(2);
                        Cell1.SetCellValue(string.Empty);
                        Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(3);
                        Cell1.SetCellValue(string.Empty);
                        Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(4);
                        Cell1.SetCellValue(string.Empty);
                        Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(5);
                        Cell1.SetCellValue(string.Empty);
                        Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(6);
                        Cell1.SetCellValue(string.Empty);
                        Cell1.CellStyle = SimpleHeaderCS;

                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(7);
                        Cell1.SetCellValue("Патина");
                        Cell1.CellStyle = SimpleHeaderCS;

                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(8);
                        Cell1.SetCellValue("Высота");
                        Cell1.CellStyle = SimpleHeaderCS;

                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(9);
                        Cell1.SetCellValue("Длина");
                        Cell1.CellStyle = SimpleHeaderCS;

                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(10);
                        Cell1.SetCellValue("Ширина");
                        Cell1.CellStyle = SimpleHeaderCS;

                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(11);
                        Cell1.SetCellValue("Кол-во");
                        Cell1.CellStyle = SimpleHeaderCS;

                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(12);
                        Cell1.SetCellValue("Вес");
                        Cell1.CellStyle = SimpleHeaderCS;

                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(13);
                        Cell1.SetCellValue("Цена");
                        Cell1.CellStyle = SimpleHeaderCS;

                        //Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(14);
                        //Cell1.SetCellValue(string.Empty);
                        //Cell1.CellStyle = SimpleHeaderCS;

                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(14);
                        Cell1.SetCellValue("Стоимость");
                        Cell1.CellStyle = SimpleHeaderCS;

                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(15);
                        Cell1.SetCellValue("Курс, " + Currency);
                        Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(16);
                        Cell1.SetCellValue("Примечание");
                        Cell1.CellStyle = SimpleHeaderCS;
                        RowIndex++;
                        RowIndex++;

                        DisplayIndex = 0;
                        //вывод заказов декора в excel
                        for (int x = 0; x < DecorResultDataTable[c].Rows.Count; x++)
                        {
                            DisplayIndex = 0;
                            Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                            Cell1.SetCellValue(DecorResultDataTable[c].Rows[x]["Name"].ToString());
                            Cell1.CellStyle = SimpleCS;
                            Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                            if (DecorResultDataTable[c].Rows[x]["Color"] != DBNull.Value)
                                Cell1.SetCellValue(DecorResultDataTable[c].Rows[x]["Color"].ToString());
                            Cell1.CellStyle = SimpleCS;
                            Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                            Cell1.SetCellValue(string.Empty);
                            Cell1.CellStyle = SimpleCS;
                            Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                            Cell1.SetCellValue(string.Empty);
                            Cell1.CellStyle = SimpleCS;
                            Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                            Cell1.SetCellValue(string.Empty);
                            Cell1.CellStyle = SimpleCS;
                            Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                            Cell1.SetCellValue(string.Empty);
                            Cell1.CellStyle = SimpleCS;
                            Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                            Cell1.SetCellValue(string.Empty);
                            Cell1.CellStyle = SimpleCS;
                            Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                            Cell1.SetCellValue(DecorResultDataTable[c].Rows[x]["Patina"].ToString());
                            Cell1.CellStyle = SimpleCS;
                            Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                            if (DecorResultDataTable[c].Rows[x]["Height"] != DBNull.Value)
                                Cell1.SetCellValue(Convert.ToInt32(DecorResultDataTable[c].Rows[x]["Height"]));
                            Cell1.CellStyle = SimpleCS;
                            Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                            if (DecorResultDataTable[c].Rows[x]["Length"] != DBNull.Value)
                                Cell1.SetCellValue(Convert.ToInt32(DecorResultDataTable[c].Rows[x]["Length"]));
                            Cell1.CellStyle = SimpleCS;
                            Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                            if (DecorResultDataTable[c].Rows[x]["Width"] != DBNull.Value)
                                Cell1.SetCellValue(Convert.ToInt32(DecorResultDataTable[c].Rows[x]["Width"]));
                            Cell1.CellStyle = SimpleCS;
                            Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                            Cell1.SetCellValue(Convert.ToInt32(DecorResultDataTable[c].Rows[x]["Count"]));
                            Cell1.CellStyle = SimpleCS;
                            Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                            decimal Weight = Decimal.Round(Convert.ToDecimal(DecorResultDataTable[c].Rows[x]["Weight"]), 3, MidpointRounding.AwayFromZero);
                            Cell1.SetCellValue(Convert.ToDouble(Weight));
                            Cell1.CellStyle = SimpleCS;
                            Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                            Cell1.SetCellValue(Convert.ToDouble(DecorResultDataTable[c].Rows[x]["Price"]));
                            Cell1.CellStyle = SimpleDecCS;
                            //Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                            //Cell1.SetCellValue(string.Empty);
                            //Cell1.CellStyle = SimpleCS;
                            Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                            Cell1.SetCellValue(Convert.ToDouble(DecorResultDataTable[c].Rows[x]["Cost"]));
                            Cell1.CellStyle = SimpleDecCS;
                            Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                            Cell1.SetCellValue(Convert.ToDouble(DecorResultDataTable[c].Rows[x]["Rate"]));
                            Cell1.CellStyle = SimpleDecCS;
                            Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                            Cell1.SetCellValue(DecorResultDataTable[c].Rows[x]["Notes"].ToString());
                            Cell1.CellStyle = SimpleCS;
                            RowIndex++;
                        }
                        //RowIndex++;
                    }
                    Cell1 = sheet1.CreateRow(RowIndex).CreateCell(0);
                    Cell1.SetCellValue($"Итого, {Currency}: " + Decimal.Round(OrderCost, 2, MidpointRounding.AwayFromZero));
                    Cell1.CellStyle = HeaderWithoutBorderCS;
                }

                RowIndex++;
                RowIndex++;
            }
            RowIndex++;
            RowIndex++;

            ClientReport.CreateReport(
                ref hssfworkbook, ref sheet1,
                NewMegaOrders, OrderNumbers, MainOrdersIDs, ClientID, ClientName, ComplaintProfilCost, ComplaintTPSCost,
                TransportCost, AdditionalCost, TotalCost, CurrencyTypeID, TotalWeight, RowIndex, IsSample);
        }

        public void Report(ref HSSFWorkbook hssfworkbook, int[] NewMegaOrders, int[] OrderNumbers, int[] MainOrdersIDs, int ClientID, string ClientName,
            decimal ComplaintProfilCost, decimal ComplaintTPSCost,
            decimal TransportCost, decimal AdditionalCost,
            decimal TotalCost, int CurrencyTypeID, decimal TotalWeight)
        {
            ClearReport();

            ClientReport = new StandardReport.Report(ref DecorCatalogOrder, ref FrontsCalculate);

            #region Create fonts and styles

            HSSFFont HeaderF1 = hssfworkbook.CreateFont();
            HeaderF1.FontHeightInPoints = 13;
            HeaderF1.Boldweight = 13 * 256;
            HeaderF1.FontName = "Calibri";

            HSSFFont HeaderF2 = hssfworkbook.CreateFont();
            HeaderF2.FontHeightInPoints = 12;
            HeaderF2.Boldweight = 12 * 256;
            HeaderF2.FontName = "Calibri";

            HSSFFont HeaderF3 = hssfworkbook.CreateFont();
            HeaderF3.FontHeightInPoints = 9;
            HeaderF3.Boldweight = 9 * 256;
            HeaderF3.FontName = "Calibri";

            HSSFFont SimpleF = hssfworkbook.CreateFont();
            SimpleF.FontHeightInPoints = 9;
            SimpleF.FontName = "Calibri";

            HSSFCellStyle SimpleCS = hssfworkbook.CreateCellStyle();
            SimpleCS.BorderBottom = HSSFCellStyle.BORDER_THIN;
            SimpleCS.BottomBorderColor = HSSFColor.BLACK.index;
            SimpleCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            SimpleCS.LeftBorderColor = HSSFColor.BLACK.index;
            SimpleCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            SimpleCS.RightBorderColor = HSSFColor.BLACK.index;
            SimpleCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            SimpleCS.TopBorderColor = HSSFColor.BLACK.index;
            SimpleCS.SetFont(SimpleF);

            HSSFCellStyle SimpleDecCS = hssfworkbook.CreateCellStyle();
            SimpleDecCS.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.00");
            SimpleDecCS.BorderBottom = HSSFCellStyle.BORDER_THIN;
            SimpleDecCS.BottomBorderColor = HSSFColor.BLACK.index;
            SimpleDecCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            SimpleDecCS.LeftBorderColor = HSSFColor.BLACK.index;
            SimpleDecCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            SimpleDecCS.RightBorderColor = HSSFColor.BLACK.index;
            SimpleDecCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            SimpleDecCS.TopBorderColor = HSSFColor.BLACK.index;
            SimpleDecCS.SetFont(SimpleF);

            HSSFCellStyle ReportCS = hssfworkbook.CreateCellStyle();
            ReportCS.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            ReportCS.BottomBorderColor = HSSFColor.BLACK.index;
            ReportCS.SetFont(HeaderF1);

            HSSFCellStyle HeaderWithoutBorderCS = hssfworkbook.CreateCellStyle();
            HeaderWithoutBorderCS.SetFont(HeaderF2);

            HSSFCellStyle HeaderWithBorderCS = hssfworkbook.CreateCellStyle();
            HeaderWithBorderCS.BorderBottom = HSSFCellStyle.BORDER_THIN;
            HeaderWithBorderCS.BottomBorderColor = HSSFColor.BLACK.index;
            HeaderWithBorderCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            HeaderWithBorderCS.LeftBorderColor = HSSFColor.BLACK.index;
            HeaderWithBorderCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            HeaderWithBorderCS.RightBorderColor = HSSFColor.BLACK.index;
            HeaderWithBorderCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            HeaderWithBorderCS.TopBorderColor = HSSFColor.BLACK.index;
            //HeaderWithBorderCS.WrapText = true;
            HeaderWithBorderCS.SetFont(HeaderF2);

            HSSFCellStyle SimpleHeaderCS = hssfworkbook.CreateCellStyle();
            //HeaderStyle.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
            //HeaderStyle.Alignment = HSSFCellStyle.ALIGN_CENTER;
            SimpleHeaderCS.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            SimpleHeaderCS.BottomBorderColor = HSSFColor.BLACK.index;
            SimpleHeaderCS.BorderLeft = HSSFCellStyle.BORDER_MEDIUM;
            SimpleHeaderCS.LeftBorderColor = HSSFColor.BLACK.index;
            SimpleHeaderCS.BorderRight = HSSFCellStyle.BORDER_MEDIUM;
            SimpleHeaderCS.RightBorderColor = HSSFColor.BLACK.index;
            SimpleHeaderCS.BorderTop = HSSFCellStyle.BORDER_MEDIUM;
            SimpleHeaderCS.TopBorderColor = HSSFColor.BLACK.index;
            //SimpleHeaderCS.WrapText = true;
            SimpleHeaderCS.SetFont(HeaderF3);

            #endregion Create fonts and styles

            HSSFSheet sheet1 = hssfworkbook.CreateSheet("Подробный отчет");
            sheet1.PrintSetup.PaperSize = (short)PaperSizeType.A4;

            sheet1.SetMargin(HSSFSheet.LeftMargin, (double).12);
            sheet1.SetMargin(HSSFSheet.RightMargin, (double).07);
            sheet1.SetMargin(HSSFSheet.TopMargin, (double).20);
            sheet1.SetMargin(HSSFSheet.BottomMargin, (double).20);

            sheet1.SetColumnWidth(0, 25 * 256);
            sheet1.SetColumnWidth(1, 15 * 256);
            sheet1.SetColumnWidth(2, 15 * 256);
            sheet1.SetColumnWidth(3, 15 * 256);
            sheet1.SetColumnWidth(4, 15 * 256);
            sheet1.SetColumnWidth(5, 15 * 256);
            sheet1.SetColumnWidth(6, 15 * 256);
            sheet1.SetColumnWidth(7, 18 * 256);
            sheet1.SetColumnWidth(8, 13 * 256);
            sheet1.SetColumnWidth(9, 10 * 256);
            sheet1.SetColumnWidth(10, 9 * 256);
            sheet1.SetColumnWidth(11, 12 * 256);
            sheet1.SetColumnWidth(12, 12 * 256);
            sheet1.SetColumnWidth(13, 13 * 256);
            sheet1.SetColumnWidth(14, 12 * 256);
            sheet1.SetColumnWidth(15, 10 * 256);

            decimal OrderCost = 0;

            int RowIndex = 0;

            string Currency = string.Empty;

            CurrencyTypesDataTable = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM CurrencyTypes",
                ConnectionStrings.CatalogConnectionString))
            {
                DA.Fill(CurrencyTypesDataTable);
            }

            DataRow[] Row = CurrencyTypesDataTable.Select("CurrencyTypeID = " + CurrencyTypeID);

            Currency = Row[0]["CurrencyType"].ToString();

            ClientName = GetClientName(MainOrdersIDs[0]);

            HSSFCell Cell1 = sheet1.CreateRow(RowIndex).CreateCell(0);
            Cell1.SetCellValue("Клиент:");
            Cell1.CellStyle = HeaderWithoutBorderCS;
            Cell1 = sheet1.CreateRow(RowIndex++).CreateCell(1);
            Cell1.SetCellValue(ClientName);
            Cell1.CellStyle = HeaderWithoutBorderCS;
            Cell1 = sheet1.CreateRow(RowIndex).CreateCell(0);
            Cell1.SetCellValue("№ заказа:");
            Cell1.CellStyle = HeaderWithoutBorderCS;
            Cell1 = sheet1.CreateRow(RowIndex++).CreateCell(1);
            Cell1.SetCellValue(string.Join(",", OrderNumbers));
            Cell1.CellStyle = HeaderWithoutBorderCS;
            if (IsComplaint(GetMegaOrderID(MainOrdersIDs[0])))
            {
                Cell1 = sheet1.CreateRow(RowIndex++).CreateCell(0);
                Cell1.SetCellValue("РЕКЛАМАЦИЯ");
                Cell1.CellStyle = HeaderWithoutBorderCS;
            }

            RowIndex++;
            for (int i = 0; i < MainOrdersIDs.Count(); i++)
            {
                OrderCost = 0;

                if (FilterOrders(MainOrdersIDs[i]))
                {
                    FillFronts(MainOrdersIDs[i]);
                    FillDecor(MainOrdersIDs[i]);

                    RowIndex++;

                    Cell1 = sheet1.CreateRow(RowIndex).CreateCell(0);
                    Cell1.SetCellValue("№" + GetOrderNumber(MainOrdersIDs[i]) + "-" + MainOrdersIDs[i].ToString());
                    Cell1.CellStyle = HeaderWithoutBorderCS;
                    RowIndex++;
                    Cell1 = sheet1.CreateRow(RowIndex).CreateCell(0);
                    Cell1.SetCellValue("Примечание к заказу: " + GetMainOrderNotes(MainOrdersIDs[i]));
                    Cell1.CellStyle = HeaderWithoutBorderCS;

                    int DisplayIndex = 0;
                    if (FrontsResultDataTable.Rows.Count != 0)
                    {
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue("Фасад");
                        Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue("Цвет профиля-1");
                        Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue("Цвет профиля-2");
                        Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue("Тип наполнителя-1");
                        Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue("Цвет наполнителя-1");
                        Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue("Тип наполнителя-2");
                        Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue("Цвет наполнителя-2");
                        Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue("Патина");
                        Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue("Выс.");
                        Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue("Шир.");
                        Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue("Кол.");
                        Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue("Вес");
                        Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue("Цена за фасад");
                        Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue("Цена за вставку");
                        Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue("Стоимость");
                        Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue("Курс, " + Currency);
                        Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue("Примечание");
                        Cell1.CellStyle = SimpleHeaderCS;

                        RowIndex++;
                    }

                    if (FrontsResultDataTable.Rows.Count != 0)
                        RowIndex++;
                    OrderCost = CalculateCost();
                    DisplayIndex = 0;
                    //вывод заказов фасадов
                    for (int x = 0; x < FrontsResultDataTable.Rows.Count; x++)
                    {
                        DisplayIndex = 0;
                        Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(FrontsResultDataTable.Rows[x]["FrontName"].ToString());
                        Cell1.CellStyle = SimpleCS;
                        Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(FrontsResultDataTable.Rows[x]["FrameColor1"].ToString());
                        Cell1.CellStyle = SimpleCS;
                        Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(FrontsResultDataTable.Rows[x]["FrameColor2"].ToString());
                        Cell1.CellStyle = SimpleCS;
                        Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(FrontsResultDataTable.Rows[x]["InsetType1"].ToString());
                        Cell1.CellStyle = SimpleCS;
                        Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(FrontsResultDataTable.Rows[x]["InsetColor1"].ToString());
                        Cell1.CellStyle = SimpleCS;
                        Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(FrontsResultDataTable.Rows[x]["InsetType2"].ToString());
                        Cell1.CellStyle = SimpleCS;
                        Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(FrontsResultDataTable.Rows[x]["InsetColor2"].ToString());
                        Cell1.CellStyle = SimpleCS;
                        Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(FrontsResultDataTable.Rows[x]["Patina"].ToString());
                        Cell1.CellStyle = SimpleCS;
                        Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(Convert.ToInt32(FrontsResultDataTable.Rows[x]["Height"]));
                        Cell1.CellStyle = SimpleCS;
                        Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(Convert.ToInt32(FrontsResultDataTable.Rows[x]["Width"]));
                        Cell1.CellStyle = SimpleCS;
                        Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(Convert.ToInt32(FrontsResultDataTable.Rows[x]["Count"]));
                        Cell1.CellStyle = SimpleCS;
                        Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                        {
                            decimal Weight = Convert.ToDecimal(FrontsResultDataTable.Rows[x]["Weight"]);
                            //decimal Weight = 0;
                            //if (Width != -1)
                            //    Weight = Convert.ToDecimal(FrontsResultDataTable.Rows[x]["Weight"]) +
                            //        Convert.ToDecimal(FrontsResultDataTable.Rows[x]["Height"]) * Convert.ToDecimal(FrontsResultDataTable.Rows[x]["Width"]) *
                            //        Convert.ToDecimal(FrontsResultDataTable.Rows[x]["Count"]) / 1000000 * Convert.ToDecimal(0.7);
                            //else
                            //    Weight = Convert.ToDecimal(FrontsResultDataTable.Rows[x]["Weight"]);

                            Weight = Decimal.Round(Weight, 3, MidpointRounding.AwayFromZero);
                            Cell1.SetCellValue(Convert.ToDouble(Weight));
                        }
                        Cell1.CellStyle = SimpleCS;
                        Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(Convert.ToDouble(FrontsResultDataTable.Rows[x]["FrontPrice"]));
                        Cell1.CellStyle = SimpleDecCS;
                        Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(Convert.ToDouble(FrontsResultDataTable.Rows[x]["InsetPrice"]));
                        Cell1.CellStyle = SimpleDecCS;
                        Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(Convert.ToDouble(FrontsResultDataTable.Rows[x]["Cost"]));
                        Cell1.CellStyle = SimpleDecCS;
                        Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(Convert.ToDouble(FrontsResultDataTable.Rows[x]["Rate"]));
                        Cell1.CellStyle = SimpleDecCS;
                        Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(FrontsResultDataTable.Rows[x]["Notes"].ToString());
                        Cell1.CellStyle = SimpleCS;

                        RowIndex++;
                    }
                    //if (FrontsResultDataTable.Rows.Count != 0)
                    //    RowIndex++;

                    DisplayIndex = 0;
                    //декор
                    for (int c = 0; c < DecorCatalogOrder.DecorProductsCount; c++)
                    {
                        if (DecorResultDataTable[c].Rows.Count == 0)
                            continue;
                        //int ColumnCount = 0;

                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(0);
                        Cell1.SetCellValue("Наименование");
                        Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(1);
                        Cell1.SetCellValue("Цвет");
                        Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(2);
                        Cell1.SetCellValue(string.Empty);
                        Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(3);
                        Cell1.SetCellValue(string.Empty);
                        Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(4);
                        Cell1.SetCellValue(string.Empty);
                        Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(5);
                        Cell1.SetCellValue(string.Empty);
                        Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(6);
                        Cell1.SetCellValue(string.Empty);
                        Cell1.CellStyle = SimpleHeaderCS;

                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(7);
                        Cell1.SetCellValue("Патина");
                        Cell1.CellStyle = SimpleHeaderCS;

                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(8);
                        Cell1.SetCellValue("Высота");
                        Cell1.CellStyle = SimpleHeaderCS;

                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(9);
                        Cell1.SetCellValue("Длина");
                        Cell1.CellStyle = SimpleHeaderCS;

                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(10);
                        Cell1.SetCellValue("Ширина");
                        Cell1.CellStyle = SimpleHeaderCS;

                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(11);
                        Cell1.SetCellValue("Кол-во");
                        Cell1.CellStyle = SimpleHeaderCS;

                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(12);
                        Cell1.SetCellValue("Вес");
                        Cell1.CellStyle = SimpleHeaderCS;

                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(13);
                        Cell1.SetCellValue("Цена");
                        Cell1.CellStyle = SimpleHeaderCS;

                        //Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(14);
                        //Cell1.SetCellValue(string.Empty);
                        //Cell1.CellStyle = SimpleHeaderCS;

                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(14);
                        Cell1.SetCellValue("Стоимость");
                        Cell1.CellStyle = SimpleHeaderCS;

                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(15);
                        Cell1.SetCellValue("Курс, " + Currency);
                        Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(16);
                        Cell1.SetCellValue("Примечание");
                        Cell1.CellStyle = SimpleHeaderCS;
                        RowIndex++;
                        RowIndex++;

                        DisplayIndex = 0;
                        //вывод заказов декора в excel
                        for (int x = 0; x < DecorResultDataTable[c].Rows.Count; x++)
                        {
                            DisplayIndex = 0;
                            Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                            Cell1.SetCellValue(DecorResultDataTable[c].Rows[x]["Name"].ToString());
                            Cell1.CellStyle = SimpleCS;
                            Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                            if (DecorResultDataTable[c].Rows[x]["Color"] != DBNull.Value)
                                Cell1.SetCellValue(DecorResultDataTable[c].Rows[x]["Color"].ToString());
                            Cell1.CellStyle = SimpleCS;
                            Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                            Cell1.SetCellValue(string.Empty);
                            Cell1.CellStyle = SimpleCS;
                            Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                            Cell1.SetCellValue(string.Empty);
                            Cell1.CellStyle = SimpleCS;
                            Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                            Cell1.SetCellValue(string.Empty);
                            Cell1.CellStyle = SimpleCS;
                            Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                            Cell1.SetCellValue(string.Empty);
                            Cell1.CellStyle = SimpleCS;
                            Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                            Cell1.SetCellValue(string.Empty);
                            Cell1.CellStyle = SimpleCS;
                            Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                            Cell1.SetCellValue(DecorResultDataTable[c].Rows[x]["Patina"].ToString());
                            Cell1.CellStyle = SimpleCS;
                            Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                            if (DecorResultDataTable[c].Rows[x]["Height"] != DBNull.Value)
                                Cell1.SetCellValue(Convert.ToInt32(DecorResultDataTable[c].Rows[x]["Height"]));
                            Cell1.CellStyle = SimpleCS;
                            Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                            if (DecorResultDataTable[c].Rows[x]["Length"] != DBNull.Value)
                                Cell1.SetCellValue(Convert.ToInt32(DecorResultDataTable[c].Rows[x]["Length"]));
                            Cell1.CellStyle = SimpleCS;
                            Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                            if (DecorResultDataTable[c].Rows[x]["Width"] != DBNull.Value)
                                Cell1.SetCellValue(Convert.ToInt32(DecorResultDataTable[c].Rows[x]["Width"]));
                            Cell1.CellStyle = SimpleCS;
                            Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                            Cell1.SetCellValue(Convert.ToInt32(DecorResultDataTable[c].Rows[x]["Count"]));
                            Cell1.CellStyle = SimpleCS;
                            Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                            decimal Weight = Decimal.Round(Convert.ToDecimal(DecorResultDataTable[c].Rows[x]["Weight"]), 3, MidpointRounding.AwayFromZero);
                            Cell1.SetCellValue(Convert.ToDouble(Weight));
                            Cell1.CellStyle = SimpleCS;
                            Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                            Cell1.SetCellValue(Convert.ToDouble(DecorResultDataTable[c].Rows[x]["Price"]));
                            Cell1.CellStyle = SimpleDecCS;
                            //Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                            //Cell1.SetCellValue(string.Empty);
                            //Cell1.CellStyle = SimpleCS;
                            Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                            Cell1.SetCellValue(Convert.ToDouble(DecorResultDataTable[c].Rows[x]["Cost"]));
                            Cell1.CellStyle = SimpleDecCS;
                            Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                            Cell1.SetCellValue(Convert.ToDouble(DecorResultDataTable[c].Rows[x]["Rate"]));
                            Cell1.CellStyle = SimpleDecCS;
                            Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                            Cell1.SetCellValue(DecorResultDataTable[c].Rows[x]["Notes"].ToString());
                            Cell1.CellStyle = SimpleCS;
                            RowIndex++;
                        }
                        //RowIndex++;
                    }
                    Cell1 = sheet1.CreateRow(RowIndex).CreateCell(0);
                    Cell1.SetCellValue($"Итого, {Currency}: " + Decimal.Round(OrderCost, 2, MidpointRounding.AwayFromZero));
                    Cell1.CellStyle = HeaderWithoutBorderCS;
                }

                RowIndex++;
                RowIndex++;
            }
            RowIndex++;
            RowIndex++;

            ClientReport.CreateReport(
                ref hssfworkbook, ref sheet1,
                NewMegaOrders, OrderNumbers, MainOrdersIDs, ClientID, ClientName, ComplaintProfilCost, ComplaintTPSCost,
                TransportCost, AdditionalCost, TotalCost, CurrencyTypeID, TotalWeight, RowIndex);
        }

        public void ClearReport()
        {
            FrontsResultDataTable.Clear();

            for (int i = 0; i < DecorCatalogOrder.DecorProductsCount; i++)
            {
                DecorResultDataTable[i].Clear();
            }
        }
    }
}