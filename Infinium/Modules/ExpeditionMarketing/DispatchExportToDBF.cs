using Infinium.Modules.Marketing.Orders;

using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;

using System;
using System.Data;
using System.Data.SqlClient;
using System.Linq;

namespace Infinium.Modules.Marketing.Dispatch
{

    public class DetailsReport
    {
        private Infinium.Modules.ExpeditionMarketing.StandardReport.Report ClientReport = null;

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

        private FrontsCatalogOrder frontsCatalogOrder = null;
        private DecorCatalogOrder decorCatalogOrder = null;

        public DetailsReport(FrontsCatalogOrder FrontsCatalogOrder, DecorCatalogOrder DecorCatalogOrder,
            ref FrontsCalculate tFrontsCalculate)
        {
            frontsCatalogOrder = FrontsCatalogOrder;
            decorCatalogOrder = DecorCatalogOrder;
            FrontsCalculate = tFrontsCalculate;

            Create();
            ClientsDataTable = new DataTable();
            using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM Clients",
                ConnectionStrings.MarketingReferenceConnectionString))
            {
                DA.Fill(ClientsDataTable);
            }
            CreateFrontsDataTable();
            CreateDecorDataTable();
        }

        private void Create()
        {
            FrontsOrdersDataTable = new DataTable();
            FrontsResultDataTable = new DataTable();
            DecorResultDataTable = new DataTable[decorCatalogOrder.DecorProductsCount];
            DecorOrdersDataTable = new DataTable();
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
            FrontsResultDataTable.Columns.Add(new DataColumn(("FrontPrice"), System.Type.GetType("System.Decimal")));
            FrontsResultDataTable.Columns.Add(new DataColumn(("InsetPrice"), System.Type.GetType("System.Decimal")));
            FrontsResultDataTable.Columns.Add(new DataColumn(("Cost"), System.Type.GetType("System.Decimal")));
            FrontsResultDataTable.Columns.Add(new DataColumn(("Rate"), System.Type.GetType("System.Decimal")));
            FrontsResultDataTable.Columns.Add(new DataColumn(("Notes"), System.Type.GetType("System.String")));
            FrontsResultDataTable.Columns.Add(new DataColumn(("ConfirmDateTime"), System.Type.GetType("System.String")));
        }

        private void CreateDecorDataTable()
        {
            for (int i = 0; i < decorCatalogOrder.DecorProductsCount; i++)
            {
                DecorResultDataTable[i] = new DataTable();

                DecorResultDataTable[i].Columns.Add(new DataColumn(("Name"), Type.GetType("System.String")));
                DecorResultDataTable[i].Columns.Add(new DataColumn(("Height"), Type.GetType("System.Int32")));
                DecorResultDataTable[i].Columns.Add(new DataColumn(("Length"), Type.GetType("System.Int32")));
                DecorResultDataTable[i].Columns.Add(new DataColumn(("Width"), Type.GetType("System.Int32")));
                DecorResultDataTable[i].Columns.Add(new DataColumn(("Color"), Type.GetType("System.String")));
                DecorResultDataTable[i].Columns.Add(new DataColumn(("Patina"), Type.GetType("System.String")));

                DecorResultDataTable[i].Columns.Add(new DataColumn(("Count"), System.Type.GetType("System.Int32")));
                DecorResultDataTable[i].Columns.Add(new DataColumn(("Price"), System.Type.GetType("System.Decimal")));
                DecorResultDataTable[i].Columns.Add(new DataColumn(("Cost"), System.Type.GetType("System.Decimal")));
                DecorResultDataTable[i].Columns.Add(new DataColumn(("Rate"), System.Type.GetType("System.Decimal")));
                DecorResultDataTable[i].Columns.Add(new DataColumn(("Notes"), System.Type.GetType("System.String")));
                DecorResultDataTable[i].Columns.Add(new DataColumn(("ConfirmDateTime"), System.Type.GetType("System.String")));
            }
        }

        private string GetMainOrderNotes(int MainOrderID)
        {
            string Notes = null;
            using (DataTable DT = new DataTable())
            {
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM MainOrders WHERE MainOrderID=" +
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
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM MegaOrders WHERE MegaOrderID=" +
                    "(SELECT MegaOrderID FROM MainOrders WHERE MainOrderID=" + MainOrderID + ")", ConnectionStrings.MarketingOrdersConnectionString))
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
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT * FROM MegaOrders WHERE MegaOrderID=" +
                    "(SELECT MegaOrderID FROM MainOrders WHERE MainOrderID=" + MainOrderID + ")", ConnectionStrings.MarketingOrdersConnectionString))
                {
                    DA.Fill(DT);
                    if (DT.Rows.Count > 0)
                        ClientID = Convert.ToInt32(DT.Rows[0]["ClientID"]);
                }
            }
            return ClientID;
        }

        private string GetClientName(int MainOrderID)
        {
            string ClientName = string.Empty;
            DataRow[] Rows = ClientsDataTable.Select("ClientID = " + GetClientID(MainOrderID));
            if (Rows.Count() > 0)
                Rows[0]["ClientName"].ToString();
            return ClientName;
        }

        public string GetFrontName(int FrontID)
        {
            string FrontName = "";
            try
            {
                DataRow[] Rows = frontsCatalogOrder.ConstFrontsDataTable.Select("FrontID = " + FrontID);
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
                DataRow[] Rows = frontsCatalogOrder.ConstColorsDataTable.Select("ColorID = " + ColorID);
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
                DataRow[] Rows = frontsCatalogOrder.PatinaDataTable.Select("PatinaID = " + PatinaID);
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
                DataRow[] Rows = frontsCatalogOrder.ConstInsetTypesDataTable.Select("InsetTypeID = " + InsetTypeID);
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
                DataRow[] Rows = frontsCatalogOrder.ConstInsetColorsDataTable.Select("InsetColorID = " + ColorID);
                ColorName = Rows[0]["InsetColorName"].ToString();
            }
            catch
            {
                return string.Empty;
            }
            return ColorName;
        }

        private void FillFronts()
        {
            FrontsResultDataTable.Clear();
            foreach (DataRow Row in FrontsOrdersDataTable.Rows)
            {
                DataRow NewRow = FrontsResultDataTable.NewRow();

                string FrameColor1 = GetColorName(Convert.ToInt32(Row["ColorID"]));
                string FrameColor2 = string.Empty;
                string InsetType1 = GetInsetTypeName(Convert.ToInt32(Row["InsetTypeID"]));
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
                NewRow["InsetType1"] = InsetType1;
                NewRow["InsetColor1"] = InsetColor1;
                NewRow["InsetType2"] = InsetType2;
                NewRow["InsetColor2"] = InsetColor2;
                NewRow["Height"] = Row["Height"];
                NewRow["Width"] = Row["Width"];
                NewRow["Count"] = Convert.ToInt32(Row["Count"]);
                NewRow["FrontPrice"] = Decimal.Round(Convert.ToDecimal(Row["FrontPrice"]), 2, MidpointRounding.AwayFromZero);
                NewRow["InsetPrice"] = Decimal.Round(Convert.ToDecimal(Row["InsetPrice"]), 2, MidpointRounding.AwayFromZero);
                NewRow["Cost"] = Decimal.Round(Convert.ToDecimal(Row["Cost"]), 2, MidpointRounding.AwayFromZero);
                NewRow["Rate"] = Row["PaymentRate"];
                NewRow["Notes"] = Row["Notes"];
                NewRow["ConfirmDateTime"] = Row["ConfirmDateTime"];
                FrontsResultDataTable.Rows.Add(NewRow);
            }
        }

        private void FillDecor()
        {
            for (int i = 0; i < decorCatalogOrder.DecorProductsCount; i++)
            {
                DecorResultDataTable[i].Clear();
                DecorResultDataTable[i].AcceptChanges();

                DataRow[] Rows = DecorOrdersDataTable.Select("ProductID = " + decorCatalogOrder.DecorProductsDataTable.Rows[i]["ProductID"]);

                if (Rows.Count() == 0)
                    continue;

                foreach (DataRow Row in Rows)
                {
                    DataRow NewRow2 = DecorResultDataTable[i].NewRow();

                    NewRow2["Name"] = decorCatalogOrder.DecorProductsDataTable.Rows[i]["ProductName"].ToString() + " " +
                                            decorCatalogOrder.GetItemName(Convert.ToInt32(Row["DecorID"]));

                    if (Convert.ToInt32(Row["Height"]) != -1)
                        NewRow2["Height"] = Convert.ToInt32(Row["Height"]);

                    if (Convert.ToInt32(Row["Length"]) != -1)
                        NewRow2["Length"] = Convert.ToInt32(Row["Length"]);

                    if (Convert.ToInt32(Row["Width"]) != -1)
                        NewRow2["Width"] = Convert.ToInt32(Row["Width"]);

                    if (decorCatalogOrder.HasParameter(Convert.ToInt32(decorCatalogOrder.DecorProductsDataTable.Rows[i]["ProductID"]), "ColorID"))
                        NewRow2["Color"] = GetColorName(Convert.ToInt32(Row["ColorID"]));

                    //if (DecorCatalogOrder.HasParameter(Convert.ToInt32(DecorCatalogOrder.DecorProductsDataTable.Rows[i]["ProductID"]), "PatinaID"))
                    NewRow2["Patina"] = GetPatinaName(Convert.ToInt32(Row["PatinaID"]));

                    NewRow2["Count"] = Convert.ToInt32(Row["Count"]);
                    NewRow2["Price"] = Decimal.Round(Convert.ToDecimal(Row["Price"]), 2, MidpointRounding.AwayFromZero);
                    NewRow2["Cost"] = Decimal.Round(Convert.ToDecimal(Row["Cost"]), 2, MidpointRounding.AwayFromZero);
                    NewRow2["Rate"] = Row["PaymentRate"];
                    NewRow2["Notes"] = Row["Notes"];
                    NewRow2["ConfirmDateTime"] = Row["ConfirmDateTime"];
                    DecorResultDataTable[i].Rows.Add(NewRow2);
                }
            }
        }

        private bool FilterOrders(int[] DispatchID, int ClientID, int MainOrderID, bool ProfilVerify, bool TPSVerify, int DiscountPaymentConditionID)
        {
            bool IsNotEmpty = false;

            FrontsOrdersDataTable.Clear();
            FrontsOrdersDataTable.AcceptChanges();

            string SelectCommand = @"SELECT PackageDetails.Count AS Count, (FrontsOrders.Square * PackageDetails.Count / FrontsOrders.Count) AS Square, (FrontsOrders.Cost * PackageDetails.Count / FrontsOrders.Count) AS Cost, FrontsOrders.*, infiniu2_catalog.dbo.FrontsConfig.AccountingName, infiniu2_catalog.dbo.FrontsConfig.InvNumber, MegaOrders.Rate, MegaOrders.PaymentRate, MegaOrders.ConfirmDateTime FROM PackageDetails
                INNER JOIN FrontsOrders ON PackageDetails.OrderID = FrontsOrders.FrontsOrdersID
                INNER JOIN MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID
                INNER JOIN MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID
                INNER JOIN  Packages ON PackageDetails.PackageID = Packages.PackageID AND Packages.ProductType = 0 AND Packages.MainOrderID = " + MainOrderID + @" AND DispatchID IN (" + string.Join(",", DispatchID) + @")
                INNER JOIN infiniu2_catalog.dbo.FrontsConfig ON FrontsOrders.FrontConfigID = infiniu2_catalog.dbo.FrontsConfig.FrontConfigID
                WHERE InvNumber IS NOT NULL ORDER BY InvNumber";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                if (DA.Fill(FrontsOrdersDataTable) > 0)
                    IsNotEmpty = true;
            }

            DecorOrdersDataTable.Clear();
            DecorOrdersDataTable.AcceptChanges();

            SelectCommand = @"SELECT PackageDetails.Count AS Count, (DecorOrders.Cost * PackageDetails.Count / DecorOrders.Count) AS Cost, DecorOrders.*, infiniu2_catalog.dbo.DecorConfig.AccountingName, infiniu2_catalog.dbo.DecorConfig.InvNumber, MegaOrders.Rate, MegaOrders.PaymentRate, MegaOrders.ConfirmDateTime FROM PackageDetails
                INNER JOIN DecorOrders ON PackageDetails.OrderID = DecorOrders.DecorOrderID
                INNER JOIN MainOrders ON DecorOrders.MainOrderID = MainOrders.MainOrderID
                INNER JOIN MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID
                INNER JOIN  Packages ON PackageDetails.PackageID = Packages.PackageID AND Packages.ProductType = 1 AND Packages.MainOrderID = " + MainOrderID + @" AND DispatchID IN (" + string.Join(",", DispatchID) + @")
                INNER JOIN infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID = infiniu2_catalog.dbo.DecorConfig.DecorConfigID
                WHERE InvNumber IS NOT NULL ORDER BY InvNumber";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                if (DA.Fill(DecorOrdersDataTable) > 0)
                    IsNotEmpty = true;
            }

            for (int i = 0; i < FrontsOrdersDataTable.Rows.Count; i++)
            {
                //if (FrontsOrdersDataTable.Rows[i]["FactoryID"].ToString() == "1")//profil
                //    if (ClientID != 145 && DiscountPaymentConditionID != 6 && !ProfilVerify)
                //        FrontsOrdersDataTable.Rows[i]["PaymentRate"] = Convert.ToDecimal(FrontsOrdersDataTable.Rows[i]["Rate"]) * 1.05m;
                //if (FrontsOrdersDataTable.Rows[i]["FactoryID"].ToString() == "2")//tps
                //    if (ClientID != 145 && DiscountPaymentConditionID != 6 && !TPSVerify)
                //        FrontsOrdersDataTable.Rows[i]["PaymentRate"] = Convert.ToDecimal(FrontsOrdersDataTable.Rows[i]["Rate"]) * 1.05m;
            }
            for (int i = 0; i < DecorOrdersDataTable.Rows.Count; i++)
            {
                //if (DecorOrdersDataTable.Rows[i]["FactoryID"].ToString() == "1")//profil
                //    if (ClientID != 145 && DiscountPaymentConditionID != 6 && !ProfilVerify)
                //        DecorOrdersDataTable.Rows[i]["PaymentRate"] = Convert.ToDecimal(DecorOrdersDataTable.Rows[i]["Rate"]) * 1.05m;
                //if (DecorOrdersDataTable.Rows[i]["FactoryID"].ToString() == "2")//tps
                //    if (ClientID != 145 && DiscountPaymentConditionID != 6 && !TPSVerify)
                //        DecorOrdersDataTable.Rows[i]["PaymentRate"] = Convert.ToDecimal(DecorOrdersDataTable.Rows[i]["Rate"]) * 1.05m;
            }
            return IsNotEmpty;
        }

        private bool FilterOrders(int[] DispatchID, int ClientID, int MainOrderID, bool ProfilVerify, bool TPSVerify, int DiscountPaymentConditionID, bool IsSample)
        {
            bool IsNotEmpty = false;

            FrontsOrdersDataTable.Clear();
            FrontsOrdersDataTable.AcceptChanges();

            string SelectCommand = @"SELECT PackageDetails.Count AS Count, (FrontsOrders.Square * PackageDetails.Count / FrontsOrders.Count) AS Square, 
                (FrontsOrders.Cost * PackageDetails.Count / FrontsOrders.Count) AS Cost, FrontsOrders.*, infiniu2_catalog.dbo.FrontsConfig.AccountingName, 
                infiniu2_catalog.dbo.FrontsConfig.InvNumber, MegaOrders.Rate, MegaOrders.PaymentRate, MegaOrders.ConfirmDateTime FROM PackageDetails
                INNER JOIN FrontsOrders ON PackageDetails.OrderID = FrontsOrders.FrontsOrdersID
                INNER JOIN MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID
                INNER JOIN MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID
                INNER JOIN  Packages ON PackageDetails.PackageID = Packages.PackageID AND Packages.ProductType = 0 AND Packages.MainOrderID = " + MainOrderID + @" AND DispatchID IN (" + string.Join(",", DispatchID) + @")
                INNER JOIN infiniu2_catalog.dbo.FrontsConfig ON FrontsOrders.FrontConfigID = infiniu2_catalog.dbo.FrontsConfig.FrontConfigID
                WHERE FrontsOrders.IsSample=1 AND InvNumber IS NOT NULL ORDER BY InvNumber";
            if (!IsSample)
                SelectCommand = @"SELECT PackageDetails.Count AS Count, (FrontsOrders.Square * PackageDetails.Count / FrontsOrders.Count) AS Square, (FrontsOrders.Cost * PackageDetails.Count / FrontsOrders.Count) AS Cost, FrontsOrders.*, infiniu2_catalog.dbo.FrontsConfig.AccountingName, infiniu2_catalog.dbo.FrontsConfig.InvNumber, MegaOrders.Rate, MegaOrders.PaymentRate, MegaOrders.ConfirmDateTime FROM PackageDetails
                INNER JOIN FrontsOrders ON PackageDetails.OrderID = FrontsOrders.FrontsOrdersID
                INNER JOIN MainOrders ON FrontsOrders.MainOrderID = MainOrders.MainOrderID
                INNER JOIN MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID
                INNER JOIN  Packages ON PackageDetails.PackageID = Packages.PackageID AND Packages.ProductType = 0 AND Packages.MainOrderID = " + MainOrderID + @" AND DispatchID IN (" + string.Join(",", DispatchID) + @")
                INNER JOIN infiniu2_catalog.dbo.FrontsConfig ON FrontsOrders.FrontConfigID = infiniu2_catalog.dbo.FrontsConfig.FrontConfigID
                WHERE FrontsOrders.IsSample=0 AND InvNumber IS NOT NULL ORDER BY InvNumber";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                if (DA.Fill(FrontsOrdersDataTable) > 0)
                    IsNotEmpty = true;
            }

            DecorOrdersDataTable.Clear();
            DecorOrdersDataTable.AcceptChanges();

            SelectCommand = @"SELECT PackageDetails.Count AS Count, (DecorOrders.Cost * PackageDetails.Count / DecorOrders.Count) AS Cost, DecorOrders.*, infiniu2_catalog.dbo.DecorConfig.AccountingName, infiniu2_catalog.dbo.DecorConfig.InvNumber, MegaOrders.Rate, MegaOrders.PaymentRate, MegaOrders.ConfirmDateTime FROM PackageDetails
                INNER JOIN DecorOrders ON PackageDetails.OrderID = DecorOrders.DecorOrderID
                INNER JOIN MainOrders ON DecorOrders.MainOrderID = MainOrders.MainOrderID
                INNER JOIN MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID
                INNER JOIN  Packages ON PackageDetails.PackageID = Packages.PackageID AND Packages.ProductType = 1 AND Packages.MainOrderID = " + MainOrderID + @" AND DispatchID IN (" + string.Join(",", DispatchID) + @")
                INNER JOIN infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID = infiniu2_catalog.dbo.DecorConfig.DecorConfigID
                WHERE DecorOrders.IsSample=1 AND InvNumber IS NOT NULL ORDER BY InvNumber";
            if (IsSample)
                SelectCommand = @"SELECT PackageDetails.Count AS Count, (DecorOrders.Cost * PackageDetails.Count / DecorOrders.Count) AS Cost, DecorOrders.*, infiniu2_catalog.dbo.DecorConfig.AccountingName, infiniu2_catalog.dbo.DecorConfig.InvNumber, MegaOrders.Rate, MegaOrders.PaymentRate, MegaOrders.ConfirmDateTime FROM PackageDetails
                INNER JOIN DecorOrders ON PackageDetails.OrderID = DecorOrders.DecorOrderID
                INNER JOIN MainOrders ON DecorOrders.MainOrderID = MainOrders.MainOrderID
                INNER JOIN MegaOrders ON MainOrders.MegaOrderID = MegaOrders.MegaOrderID
                INNER JOIN  Packages ON PackageDetails.PackageID = Packages.PackageID AND Packages.ProductType = 1 AND Packages.MainOrderID = " + MainOrderID + @" AND DispatchID IN (" + string.Join(",", DispatchID) + @")
                INNER JOIN infiniu2_catalog.dbo.DecorConfig ON DecorOrders.DecorConfigID = infiniu2_catalog.dbo.DecorConfig.DecorConfigID
                WHERE DecorOrders.IsSample=0 AND InvNumber IS NOT NULL ORDER BY InvNumber";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand,
                ConnectionStrings.MarketingOrdersConnectionString))
            {
                if (DA.Fill(DecorOrdersDataTable) > 0)
                    IsNotEmpty = true;
            }

            for (int i = 0; i < FrontsOrdersDataTable.Rows.Count; i++)
            {
                //if (FrontsOrdersDataTable.Rows[i]["FactoryID"].ToString() == "1")//profil
                //    if (ClientID != 145 && DiscountPaymentConditionID != 6 && !ProfilVerify)
                //        FrontsOrdersDataTable.Rows[i]["PaymentRate"] = Convert.ToDecimal(FrontsOrdersDataTable.Rows[i]["Rate"]) * 1.05m;
                //if (FrontsOrdersDataTable.Rows[i]["FactoryID"].ToString() == "2")//tps
                //    if (ClientID != 145 && DiscountPaymentConditionID != 6 && !TPSVerify)
                //        FrontsOrdersDataTable.Rows[i]["PaymentRate"] = Convert.ToDecimal(FrontsOrdersDataTable.Rows[i]["Rate"]) * 1.05m;
            }
            for (int i = 0; i < DecorOrdersDataTable.Rows.Count; i++)
            {
                //if (DecorOrdersDataTable.Rows[i]["FactoryID"].ToString() == "1")//profil
                //    if (ClientID != 145 && DiscountPaymentConditionID != 6 && !ProfilVerify)
                //        DecorOrdersDataTable.Rows[i]["PaymentRate"] = Convert.ToDecimal(DecorOrdersDataTable.Rows[i]["Rate"]) * 1.05m;
                //if (DecorOrdersDataTable.Rows[i]["FactoryID"].ToString() == "2")//tps
                //    if (ClientID != 145 && DiscountPaymentConditionID != 6 && !TPSVerify)
                //        DecorOrdersDataTable.Rows[i]["PaymentRate"] = Convert.ToDecimal(DecorOrdersDataTable.Rows[i]["Rate"]) * 1.05m;
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
                using (SqlDataAdapter DA = new SqlDataAdapter("SELECT MegaOrderID FROM MainOrders" +
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
                " FROM MegaOrders WHERE MegaOrderID = " + MegaOrderID, ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (SqlCommandBuilder CB = new SqlCommandBuilder(DA))
                {
                    using (DataTable DT = new DataTable())
                    {
                        if (DA.Fill(DT) > 0)
                        {

                            if (DBNull.Value != DT.Rows[0]["IsComplaint"] && Convert.ToInt32(DT.Rows[0]["IsComplaint"]) > 0)
                            {
                                IsComplaint = Convert.ToBoolean(DT.Rows[0]["IsComplaint"]);
                            }
                        }
                    }
                }
            }

            return IsComplaint;
        }

        public void GetDispatchInfo(int[] DispatchID, ref int[] OrderNumbers, ref decimal ComplaintProfilCost, ref decimal ComplaintTPSCost, ref decimal TransportCost, ref decimal AdditionalCost, ref int CurrencyTypeID, ref decimal TotalWeight)
        {
            string SelectCommand = @"SELECT MegaOrderID, OrderNumber, TransportCost, AdditionalCost, PaymentRate, CurrencyTypeID, Weight FROM MegaOrders
                WHERE MegaOrderID IN (SELECT MegaOrderID FROM MainOrders
                WHERE MainOrderID IN (SELECT MainOrderID FROM Packages WHERE DispatchID IN (" + string.Join(",", DispatchID) + ")))";
            using (SqlDataAdapter DA = new SqlDataAdapter(SelectCommand, ConnectionStrings.MarketingOrdersConnectionString))
            {
                using (DataTable DT = new DataTable())
                {
                    if (DA.Fill(DT) > 0)
                    {
                        if (DT.Rows[0]["CurrencyTypeID"] != DBNull.Value)
                            int.TryParse(DT.Rows[0]["CurrencyTypeID"].ToString(), out CurrencyTypeID);

                        OrderNumbers = new int[DT.Rows.Count];
                        for (int i = 0; i < DT.Rows.Count; i++)
                        {
                            TotalWeight += Convert.ToDecimal(DT.Rows[i]["Weight"]);
                            TransportCost += Convert.ToDecimal(DT.Rows[i]["TransportCost"]) * Convert.ToDecimal(DT.Rows[i]["PaymentRate"]);
                            AdditionalCost += Convert.ToDecimal(DT.Rows[i]["AdditionalCost"]) * Convert.ToDecimal(DT.Rows[i]["PaymentRate"]);
                            OrderNumbers[i] = Convert.ToInt32(DT.Rows[i]["OrderNumber"]);
                        }
                    }
                }
            }
        }

        public void CreateReport(ref HSSFWorkbook hssfworkbook, int[] DispatchID, int[] OrderNumbers, int[] MainOrdersIDs, int ClientID, string ClientName, bool ProfilVerify, bool TPSVerify, int DiscountPaymentConditionID)
        {
            ClearReport();

            decimal ComplaintProfilCost = 0;
            decimal ComplaintTPSCost = 0;
            decimal TransportCost = 0;
            decimal AdditionalCost = 0;
            decimal TotalCost = 0;
            int CurrencyTypeID = 0;
            decimal TotalWeight = 0;

            GetDispatchInfo(DispatchID, ref OrderNumbers, ref ComplaintProfilCost, ref ComplaintTPSCost, ref TransportCost, ref AdditionalCost, ref CurrencyTypeID, ref TotalWeight);
            ClientReport = new ExpeditionMarketing.StandardReport.Report(ref decorCatalogOrder, ref FrontsCalculate);

            #region Create fonts and styles

            HSSFFont HeaderF1 = hssfworkbook.CreateFont();
            HeaderF1.FontHeightInPoints = 13;
            HeaderF1.Boldweight = 13 * 256;
            HeaderF1.FontName = "Calibri";

            HSSFFont HeaderItalicF2 = hssfworkbook.CreateFont();
            HeaderItalicF2.Color = HSSFColor.RED.index;
            HeaderItalicF2.FontHeightInPoints = 12;
            HeaderItalicF2.Boldweight = 12 * 256;
            HeaderItalicF2.FontName = "Calibri";
            HeaderItalicF2.IsItalic = true;

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

            HSSFCellStyle RateCS = hssfworkbook.CreateCellStyle();
            RateCS.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.0000");
            RateCS.BorderBottom = HSSFCellStyle.BORDER_THIN;
            RateCS.BottomBorderColor = HSSFColor.BLACK.index;
            RateCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            RateCS.LeftBorderColor = HSSFColor.BLACK.index;
            RateCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            RateCS.RightBorderColor = HSSFColor.BLACK.index;
            RateCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            RateCS.TopBorderColor = HSSFColor.BLACK.index;
            RateCS.SetFont(SimpleF);

            HSSFCellStyle ReportCS = hssfworkbook.CreateCellStyle();
            ReportCS.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            ReportCS.BottomBorderColor = HSSFColor.BLACK.index;
            ReportCS.SetFont(HeaderF1);

            HSSFCellStyle HeaderItalicWithoutBorderCS = hssfworkbook.CreateCellStyle();
            HeaderItalicWithoutBorderCS.SetFont(HeaderItalicF2);

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

            #endregion

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
            sheet1.SetColumnWidth(7, 14 * 256);
            sheet1.SetColumnWidth(8, 13 * 256);
            sheet1.SetColumnWidth(9, 10 * 256);
            sheet1.SetColumnWidth(10, 9 * 256);
            sheet1.SetColumnWidth(11, 12 * 256);
            sheet1.SetColumnWidth(12, 13 * 256);
            sheet1.SetColumnWidth(13, 12 * 256);
            sheet1.SetColumnWidth(14, 10 * 256);

            decimal OrderCost = 0;

            int RowIndex = 0;

            string Currency = string.Empty;
            
            DataRow[] Row = frontsCatalogOrder.CurrencyTypesDT.Select("CurrencyTypeID = " + CurrencyTypeID);

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
            bool Complaint = IsComplaint(GetMegaOrderID(MainOrdersIDs[0]));
            if (Complaint)
            {
                Cell1 = sheet1.CreateRow(RowIndex++).CreateCell(0);
                Cell1.SetCellValue("РЕКЛАМАЦИЯ");
                Cell1.CellStyle = HeaderItalicWithoutBorderCS;
            }

            RowIndex++;
            for (int i = 0; i < MainOrdersIDs.Count(); i++)
            {
                OrderCost = 0;

                int MainOrderID = MainOrdersIDs[i];
                Complaint = IsComplaint(GetMegaOrderID(MainOrdersIDs[i]));
                if (FilterOrders(DispatchID, ClientID, MainOrdersIDs[i], ProfilVerify, TPSVerify, DiscountPaymentConditionID))
                {
                    FillFronts();
                    FillDecor();

                    RowIndex++;
                    if (Complaint)
                    {
                        Cell1 = sheet1.CreateRow(RowIndex).CreateCell(0);
                        Cell1.SetCellValue("№" + GetOrderNumber(MainOrdersIDs[i]) + "-" + MainOrdersIDs[i].ToString());
                        Cell1.CellStyle = HeaderItalicWithoutBorderCS;
                        RowIndex++;
                        Cell1 = sheet1.CreateRow(RowIndex).CreateCell(0);
                        Cell1.SetCellValue("Примечание к заказу: " + GetMainOrderNotes(MainOrdersIDs[i]));
                        Cell1.CellStyle = HeaderItalicWithoutBorderCS;
                    }
                    else
                    {
                        Cell1 = sheet1.CreateRow(RowIndex).CreateCell(0);
                        Cell1.SetCellValue("№" + GetOrderNumber(MainOrdersIDs[i]) + "-" + MainOrdersIDs[i].ToString());
                        Cell1.CellStyle = HeaderWithoutBorderCS;
                        RowIndex++;
                        Cell1 = sheet1.CreateRow(RowIndex).CreateCell(0);
                        Cell1.SetCellValue("Примечание к заказу: " + GetMainOrderNotes(MainOrdersIDs[i]));
                        Cell1.CellStyle = HeaderWithoutBorderCS;
                    }
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
                        Cell1.SetCellValue("Цена за фасад");
                        Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue("Цена за вставку");
                        Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue("Стоимость");
                        Cell1.CellStyle = SimpleHeaderCS;
                        //Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex++);
                        //Cell1.SetCellValue("Курс, " + Currency);
                        //Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue("Примечание");
                        Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue("Согласовано");
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
                        Cell1.SetCellValue(Convert.ToDouble(FrontsResultDataTable.Rows[x]["FrontPrice"]));
                        Cell1.CellStyle = SimpleDecCS;
                        Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(Convert.ToDouble(FrontsResultDataTable.Rows[x]["InsetPrice"]));
                        Cell1.CellStyle = SimpleDecCS;
                        Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(Convert.ToDouble(FrontsResultDataTable.Rows[x]["Cost"]));
                        Cell1.CellStyle = SimpleDecCS;
                        //Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                        //Cell1.SetCellValue(Convert.ToDouble(FrontsResultDataTable.Rows[x]["Rate"]));
                        //Cell1.CellStyle = RateCS;
                        Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(FrontsResultDataTable.Rows[x]["Notes"].ToString());
                        Cell1.CellStyle = SimpleCS;
                        Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(FrontsResultDataTable.Rows[x]["ConfirmDateTime"].ToString());
                        Cell1.CellStyle = SimpleCS;

                        RowIndex++;
                    }
                    //if (FrontsResultDataTable.Rows.Count != 0)
                    //    RowIndex++;

                    DisplayIndex = 0;
                    //декор
                    for (int c = 0; c < decorCatalogOrder.DecorProductsCount; c++)
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
                        Cell1.SetCellValue("Цена");
                        Cell1.CellStyle = SimpleHeaderCS;

                        //Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(13);
                        //Cell1.SetCellValue(string.Empty);
                        //Cell1.CellStyle = SimpleHeaderCS;

                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(13);
                        Cell1.SetCellValue("Стоимость");
                        Cell1.CellStyle = SimpleHeaderCS;

                        //Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(14);
                        //Cell1.SetCellValue("Курс, " + Currency);
                        //Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(14);
                        Cell1.SetCellValue("Примечание");
                        Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(15);
                        Cell1.SetCellValue("Согласовано");
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
                            Cell1.SetCellValue(Convert.ToDouble(DecorResultDataTable[c].Rows[x]["Price"]));
                            Cell1.CellStyle = SimpleDecCS;
                            //Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                            //Cell1.SetCellValue(string.Empty);
                            //Cell1.CellStyle = SimpleCS;
                            Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                            Cell1.SetCellValue(Convert.ToDouble(DecorResultDataTable[c].Rows[x]["Cost"]));
                            Cell1.CellStyle = SimpleDecCS;
                            //Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                            //Cell1.SetCellValue(Convert.ToDouble(DecorResultDataTable[c].Rows[x]["Rate"]));
                            //Cell1.CellStyle = RateCS;
                            Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                            Cell1.SetCellValue(DecorResultDataTable[c].Rows[x]["Notes"].ToString());
                            Cell1.CellStyle = SimpleCS;
                            Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                            Cell1.SetCellValue(DecorResultDataTable[c].Rows[x]["ConfirmDateTime"].ToString());
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
                DispatchID, OrderNumbers, MainOrdersIDs, ClientID, ClientName, ComplaintProfilCost, ComplaintTPSCost,
                TransportCost, AdditionalCost, TotalCost, CurrencyTypeID, TotalWeight, RowIndex, ProfilVerify, TPSVerify, DiscountPaymentConditionID);
        }

        public void CreateReport(ref HSSFWorkbook hssfworkbook, int[] DispatchID, int[] OrderNumbers, int[] MainOrdersIDs,
            int ClientID, string ClientName, bool ProfilVerify, bool TPSVerify, int DiscountPaymentConditionID, bool IsSample)
        {
            ClearReport();

            decimal ComplaintProfilCost = 0;
            decimal ComplaintTPSCost = 0;
            decimal TransportCost = 0;
            decimal AdditionalCost = 0;
            decimal TotalCost = 0;
            int CurrencyTypeID = 0;
            decimal TotalWeight = 0;

            GetDispatchInfo(DispatchID, ref OrderNumbers, ref ComplaintProfilCost, ref ComplaintTPSCost, ref TransportCost, ref AdditionalCost, ref CurrencyTypeID, ref TotalWeight);
            ClientReport = new ExpeditionMarketing.StandardReport.Report(ref decorCatalogOrder, ref FrontsCalculate);

            #region Create fonts and styles

            HSSFFont HeaderF1 = hssfworkbook.CreateFont();
            HeaderF1.FontHeightInPoints = 13;
            HeaderF1.Boldweight = 13 * 256;
            HeaderF1.FontName = "Calibri";

            HSSFFont HeaderItalicF2 = hssfworkbook.CreateFont();
            HeaderItalicF2.Color = HSSFColor.RED.index;
            HeaderItalicF2.FontHeightInPoints = 12;
            HeaderItalicF2.Boldweight = 12 * 256;
            HeaderItalicF2.FontName = "Calibri";
            HeaderItalicF2.IsItalic = true;

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

            HSSFCellStyle RateCS = hssfworkbook.CreateCellStyle();
            RateCS.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.0000");
            RateCS.BorderBottom = HSSFCellStyle.BORDER_THIN;
            RateCS.BottomBorderColor = HSSFColor.BLACK.index;
            RateCS.BorderLeft = HSSFCellStyle.BORDER_THIN;
            RateCS.LeftBorderColor = HSSFColor.BLACK.index;
            RateCS.BorderRight = HSSFCellStyle.BORDER_THIN;
            RateCS.RightBorderColor = HSSFColor.BLACK.index;
            RateCS.BorderTop = HSSFCellStyle.BORDER_THIN;
            RateCS.TopBorderColor = HSSFColor.BLACK.index;
            RateCS.SetFont(SimpleF);

            HSSFCellStyle ReportCS = hssfworkbook.CreateCellStyle();
            ReportCS.BorderBottom = HSSFCellStyle.BORDER_MEDIUM;
            ReportCS.BottomBorderColor = HSSFColor.BLACK.index;
            ReportCS.SetFont(HeaderF1);

            HSSFCellStyle HeaderItalicWithoutBorderCS = hssfworkbook.CreateCellStyle();
            HeaderItalicWithoutBorderCS.SetFont(HeaderItalicF2);

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

            #endregion

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
            sheet1.SetColumnWidth(7, 14 * 256);
            sheet1.SetColumnWidth(8, 13 * 256);
            sheet1.SetColumnWidth(9, 10 * 256);
            sheet1.SetColumnWidth(10, 9 * 256);
            sheet1.SetColumnWidth(11, 12 * 256);
            sheet1.SetColumnWidth(12, 13 * 256);
            sheet1.SetColumnWidth(13, 12 * 256);
            sheet1.SetColumnWidth(14, 10 * 256);

            decimal OrderCost = 0;

            int RowIndex = 0;

            string Currency = string.Empty;
            
            DataRow[] Row = frontsCatalogOrder.CurrencyTypesDT.Select("CurrencyTypeID = " + CurrencyTypeID);

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
            bool Complaint = IsComplaint(GetMegaOrderID(MainOrdersIDs[0]));
            if (Complaint)
            {
                Cell1 = sheet1.CreateRow(RowIndex++).CreateCell(0);
                Cell1.SetCellValue("РЕКЛАМАЦИЯ");
                Cell1.CellStyle = HeaderItalicWithoutBorderCS;
            }

            RowIndex++;
            for (int i = 0; i < MainOrdersIDs.Count(); i++)
            {
                OrderCost = 0;

                int MainOrderID = MainOrdersIDs[i];
                Complaint = IsComplaint(GetMegaOrderID(MainOrdersIDs[i]));
                if (FilterOrders(DispatchID, ClientID, MainOrdersIDs[i], ProfilVerify, TPSVerify, DiscountPaymentConditionID, IsSample))
                {
                    FillFronts();
                    FillDecor();

                    RowIndex++;
                    if (Complaint)
                    {
                        Cell1 = sheet1.CreateRow(RowIndex).CreateCell(0);
                        Cell1.SetCellValue("№" + GetOrderNumber(MainOrdersIDs[i]) + "-" + MainOrdersIDs[i].ToString());
                        Cell1.CellStyle = HeaderItalicWithoutBorderCS;
                        RowIndex++;
                        Cell1 = sheet1.CreateRow(RowIndex).CreateCell(0);
                        Cell1.SetCellValue("Примечание к заказу: " + GetMainOrderNotes(MainOrdersIDs[i]));
                        Cell1.CellStyle = HeaderItalicWithoutBorderCS;
                    }
                    else
                    {
                        Cell1 = sheet1.CreateRow(RowIndex).CreateCell(0);
                        Cell1.SetCellValue("№" + GetOrderNumber(MainOrdersIDs[i]) + "-" + MainOrdersIDs[i].ToString());
                        Cell1.CellStyle = HeaderWithoutBorderCS;
                        RowIndex++;
                        Cell1 = sheet1.CreateRow(RowIndex).CreateCell(0);
                        Cell1.SetCellValue("Примечание к заказу: " + GetMainOrderNotes(MainOrdersIDs[i]));
                        Cell1.CellStyle = HeaderWithoutBorderCS;
                    }
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
                        Cell1.SetCellValue("Цена за фасад");
                        Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue("Цена за вставку");
                        Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue("Стоимость");
                        Cell1.CellStyle = SimpleHeaderCS;
                        //Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex++);
                        //Cell1.SetCellValue("Курс, " + Currency);
                        //Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue("Примечание");
                        Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue("Согласовано");
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
                        Cell1.SetCellValue(Convert.ToDouble(FrontsResultDataTable.Rows[x]["FrontPrice"]));
                        Cell1.CellStyle = SimpleDecCS;
                        Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(Convert.ToDouble(FrontsResultDataTable.Rows[x]["InsetPrice"]));
                        Cell1.CellStyle = SimpleDecCS;
                        Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(Convert.ToDouble(FrontsResultDataTable.Rows[x]["Cost"]));
                        Cell1.CellStyle = SimpleDecCS;
                        //Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                        //Cell1.SetCellValue(Convert.ToDouble(FrontsResultDataTable.Rows[x]["Rate"]));
                        //Cell1.CellStyle = RateCS;
                        Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(FrontsResultDataTable.Rows[x]["Notes"].ToString());
                        Cell1.CellStyle = SimpleCS;
                        Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                        Cell1.SetCellValue(FrontsResultDataTable.Rows[x]["ConfirmDateTime"].ToString());
                        Cell1.CellStyle = SimpleCS;

                        RowIndex++;
                    }
                    //if (FrontsResultDataTable.Rows.Count != 0)
                    //    RowIndex++;

                    DisplayIndex = 0;
                    //декор
                    for (int c = 0; c < decorCatalogOrder.DecorProductsCount; c++)
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
                        Cell1.SetCellValue("Цена");
                        Cell1.CellStyle = SimpleHeaderCS;

                        //Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(13);
                        //Cell1.SetCellValue(string.Empty);
                        //Cell1.CellStyle = SimpleHeaderCS;

                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(13);
                        Cell1.SetCellValue("Стоимость");
                        Cell1.CellStyle = SimpleHeaderCS;

                        //Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(14);
                        //Cell1.SetCellValue("Курс, " + Currency);
                        //Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(14);
                        Cell1.SetCellValue("Примечание");
                        Cell1.CellStyle = SimpleHeaderCS;
                        Cell1 = sheet1.CreateRow(RowIndex + 1).CreateCell(15);
                        Cell1.SetCellValue("Согласовано");
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
                            Cell1.SetCellValue(Convert.ToDouble(DecorResultDataTable[c].Rows[x]["Price"]));
                            Cell1.CellStyle = SimpleDecCS;
                            //Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                            //Cell1.SetCellValue(string.Empty);
                            //Cell1.CellStyle = SimpleCS;
                            Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                            Cell1.SetCellValue(Convert.ToDouble(DecorResultDataTable[c].Rows[x]["Cost"]));
                            Cell1.CellStyle = SimpleDecCS;
                            //Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                            //Cell1.SetCellValue(Convert.ToDouble(DecorResultDataTable[c].Rows[x]["Rate"]));
                            //Cell1.CellStyle = RateCS;
                            Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                            Cell1.SetCellValue(DecorResultDataTable[c].Rows[x]["Notes"].ToString());
                            Cell1.CellStyle = SimpleCS;
                            Cell1 = sheet1.CreateRow(RowIndex).CreateCell(DisplayIndex++);
                            Cell1.SetCellValue(DecorResultDataTable[c].Rows[x]["ConfirmDateTime"].ToString());
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
                DispatchID, OrderNumbers, MainOrdersIDs, ClientID, ClientName, ComplaintProfilCost, ComplaintTPSCost,
                TransportCost, AdditionalCost, TotalCost, CurrencyTypeID, TotalWeight, RowIndex, ProfilVerify, TPSVerify, DiscountPaymentConditionID, IsSample);
        }

        public void ClearReport()
        {
            FrontsResultDataTable.Clear();

            for (int i = 0; i < decorCatalogOrder.DecorProductsCount; i++)
            {
                DecorResultDataTable[i].Clear();
            }
        }
    }
}
